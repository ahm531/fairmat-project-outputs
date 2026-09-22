"""Compare the sheet's event and contribution names against the cited pages.

Step 3 of the review. ``fetch_source_titles.py`` has already pulled every
candidate title off every URL the sheet cites; this script decides, row by row,
whether one of those candidates is the same title spelt better.

It is deliberately timid. A cited page is often a programme index or an
institute front page rather than the talk itself, so a candidate is only
accepted when it is recognisably the same string:

    identical once case, punctuation and spacing are ignored
        -> a typo fix, applied, painted light yellow
    clearly the same title but materially different wording
        -> applied, painted blue, because a human should see it
    anything less similar
        -> recorded as a suggestion and NOT applied

Output: local/title_matches.json, which build_curated_sheet.py reads. Nothing
is written to the spreadsheet here, so the decision stays reviewable and the
sheet stays reproducible from its inputs.

    python scripts/match_titles.py
"""

from __future__ import annotations

import datetime
import json
import re
import sys
import unicodedata
import warnings
from difflib import SequenceMatcher
from pathlib import Path

warnings.filterwarnings('ignore', module='openpyxl')

sys.path.insert(0, str(Path(__file__).resolve().parent))
# Shared so the two scripts agree on what identifies a row; a key built here
# has to find the same row when build_curated_sheet.py applies the match.
from build_curated_sheet import DEFAULT_SOURCE, curate_all, title_key  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
LOCAL = ROOT / 'local'
TITLES = LOCAL / 'source_titles.json'
OUT = LOCAL / 'title_matches.json'

# Same string, different spelling -> a typo fix.
TYPO_FLOOR = 1.0
# Same title, materially reworded -> worth a human's eye, so blue.
APPLY_FLOOR = 0.78
# Below this a candidate is not reported at all; it is a different page.
SUGGEST_FLOOR = 0.60
# A word shorter than this carries no meaning of its own.
MIN_WORD = 3
# More added words than this, and it is a sentence about the event.
MAX_ADDED_WORDS = 5
# The head of a split page title has to be long enough to be a title.
MIN_HEAD = 9

URL_RE = re.compile(r'https?://[^\s<>"\']+')

# Site furniture that is never a talk title, however well it scores.
JUNK = re.compile(
    r'^(home|search|menu|login|log in|sign in|abstract|abstracts|presenters?|'
    r'programme?|program|schedule|timetable|overview|news|events?|contact|'
    r'quick links|services|impressum|privacy|cookie|navigation|skip to|'
    r'session|sessions|speakers?|registration|participants?|committee)\b',
    re.I,
)
# A title that is only a series or venue name tells us nothing about the talk.
TOO_GENERIC = {
    'fairmat', 'nomad', 'indico', 'youtube', 'conference', 'workshop',
    'seminar', 'colloquium', 'tutorial', 'meeting',
}


def fold(text: str) -> str:
    """Case, accents, punctuation and spacing removed -- spelling only."""
    text = unicodedata.normalize('NFKD', text)
    text = ''.join(c for c in text if not unicodedata.combining(c))
    text = text.lower().replace('&', ' and ')
    text = re.sub(r'[^a-z0-9]+', ' ', text)
    return re.sub(r'\s+', ' ', text).strip()


def words(text: str) -> set[str]:
    return {w for w in fold(text).split() if len(w) >= MIN_WORD}


def similarity(left: str, right: str) -> float:
    """Blend of sequence and word overlap, so neither alone can carry a match."""
    a, b = fold(left), fold(right)
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    sequence = SequenceMatcher(None, a, b).ratio()
    left_words, right_words = words(left), words(right)
    if not left_words or not right_words:
        return sequence
    overlap = len(left_words & right_words) / len(left_words | right_words)
    return max(sequence, overlap)


# Everything a programme page wraps around a talk title. Each is stripped
# repeatedly, because pages stack them: 'Colloquium: Claudia Draxl: <title>'.
PREFIXES = (
    # 'PAN04 - ', 'O 84: ', 'TT 2: ', 'MT03.02.06: ' -- session codes.
    re.compile(r'^[A-Z]{1,4}\s?\d{1,3}(?:\.\d+)*\s*[-:–]\s*'),
    # 'Andrea Albino: <title>' -- the speaker, not the talk.
    re.compile(r'^[A-Z][a-zÀ-ſ]+(?: [A-Z][a-zÀ-ſ]+){1,2}:\s+'),
    # The kind of session, which the sheet already records in its own column.
    re.compile(r'^(?:title|talk|lecture|webinar|colloquium|seminar|keynote|'
               r'invited talk|hands-on tutorial|tutorial|poster|session|'
               r'symposium|workshop)\s*:\s*', re.I),
    # The event, prepended to the contribution: 'FAIRmat Users Meeting | ...'.
    re.compile(r'^[^|]{4,60}(?:meeting|conference|workshop|school|seminar|'
               r'symposium|colloquium)\s*[|–-]\s+', re.I),
    # List numbering: '1 - ', '1. ', '[1] ', '(1) ', CJK '【1】'.
    re.compile(r'^(?:\d{1,3}\s*[-.–)]\s+|\[\d{1,3}\]\s*|\(\d{1,3}\)\s*|'
               r'[【［]\s*\d{1,3}\s*[】］]\s*)'),
)
SUFFIXES = (
    # ' | MTEX', ' | NFDI' -- the site name appended to its own page title.
    re.compile(r'\s*[|·]\s*[^|·]{1,30}$'),
    # Record identifiers: ' [ID:52263]'.
    re.compile(r'\s*\[(?:id|ID)\s*:?\s*\d+\]\s*$'),
    # ' - Sathya Sai Seetharaman', ' by Carina Kanitz' -- the speaker again.
    re.compile(r'\s*[–—-]\s*[A-Z][\wÀ-ſ.]+'
               r'(?: [A-Z][\wÀ-ſ.]+){1,3}\s*$'),
    re.compile(r'\s+by\s+[A-Z][\wÀ-ſ.]+(?: [A-Z][\wÀ-ſ.]+){1,3}\s*$'),
    # The APS footnote marker.
    re.compile(r'\s*\(\*\)\s*$'),
    re.compile(r'\.(?:pdf|docx?|pptx?)$', re.I),
)

# Words that, when a candidate adds them, say the page is *about* the event
# rather than being its title.
ABOUT_WORDS = {
    'proposal', 'about', 'overview', 'registration', 'programme', 'program',
    'abstract', 'abstracts', 'announcement', 'call', 'symposium', 'archive',
}
# A hyphen with a space after it inside a word: text recovered badly from a
# PDF ('data- centric'). The sheet's own spelling is better.
BROKEN_HYPHEN = re.compile(r'\w[-‐–]\s\w')


def tidy_candidate(candidate: str) -> str:
    """Peel off everything the page wrapped around the title."""
    for _ in range(4):
        before = candidate
        for pattern in PREFIXES:
            candidate = pattern.sub('', candidate, count=1).strip()
        for pattern in SUFFIXES:
            candidate = pattern.sub('', candidate, count=1).strip()
        if candidate == before:
            break
    return re.sub(r'\s+', ' ', candidate).strip()


def is_improvement(current: str, candidate: str) -> bool:
    """True only when the candidate says everything the sheet says, and more.

    The sheet's own text was typed by a curator who knew which talk this was.
    A page title that drops part of it is not a correction but a loss --
    'DPG_SKM21' -> 'SKM21', 'Gordon Research Conference Correlated Electron
    Systems' -> 'Correlated Electron Systems'. So a rewording is only accepted
    when it is a superset: every content word survives, and the additions are
    few enough that it is plainly the same title rather than a page heading
    that happens to quote it.
    """
    here, there = words(current), words(candidate)
    if not here or not there:
        return False
    if not here <= there:
        return False
    if here == there:
        # The same words in a different order ('NFDI Metadata Workshop' vs
        # 'NFDI Workshop Metadata'). Neither spelling is evidently righter.
        return False
    # Never shorter: dropping '56.' from '56. Jahrestreffen Deutscher
    # Katalytiker' loses the edition, which is the one thing that dates it.
    if len(fold(candidate)) < len(fold(current)):
        return False
    added = there - here
    if len(added) > MAX_ADDED_WORDS or added & ABOUT_WORDS:
        return False
    # A candidate several times longer is a page heading wrapped around the
    # title ('About the Conference on ... 2025 in Aachen'), not the title.
    return len(fold(candidate)) <= len(fold(current)) * 1.8


def plausible(candidate: str) -> bool:
    if JUNK.match(candidate.strip()):
        return False
    if fold(candidate) in TOO_GENERIC:
        return False
    if BROKEN_HYPHEN.search(candidate):
        return False
    # A bare host name, a file name, or a breadcrumb trail.
    return not re.fullmatch(r'[\w.-]+\.(?:de|eu|org|com|net)', candidate.strip())


def split_candidate(candidate: str) -> list[str]:
    """Page titles often append the site name; try the head as well."""
    out = [candidate]
    for separator in (' - ', ' | ', ' · ', ' — ', ' :: '):
        if separator in candidate:
            head = candidate.split(separator, 1)[0].strip()
            if len(head) >= MIN_HEAD:
                out.append(head)
    return out


def newest_sheet() -> Path:
    candidates = sorted(LOCAL.glob('Events database - CURATED_*.xlsx'))
    if not candidates:
        raise SystemExit('no curated sheet found; run build_curated_sheet.py first')
    return candidates[-1]


def cited_urls(*fields: str) -> list[str]:
    """Every distinct URL in these cells, in the order they appear."""
    out: list[str] = []
    for value in fields:
        for found in URL_RE.findall(value):
            trimmed = found.rstrip('.,);')
            if trimmed not in out:
                out.append(trimmed)
    return out


def candidates_for(urls: list[str], pages: dict) -> list[tuple[str, str]]:
    """Every string on those pages that could be a title, with its page."""
    out: list[tuple[str, str]] = []
    for url in urls:
        for candidate in pages.get(url, {}).get('titles', []):
            for variant in split_candidate(candidate):
                cleaned = tidy_candidate(variant)
                if plausible(cleaned):
                    out.append((cleaned, url))
    return out


def best_candidate(current: str, candidates: list[tuple[str, str]]) -> tuple[str, str, float]:
    best, best_url, best_score = '', '', 0.0
    for candidate, url in candidates:
        score = similarity(current, candidate)
        # Prefer the longer, more specific of two equally good candidates.
        if score > best_score or (score == best_score and len(candidate) > len(best)):
            best, best_url, best_score = candidate, url, score
    return best, best_url, best_score


def classify(current: str, best: str, score: float) -> str | None:
    """'minor', 'title', 'suggest', or None when there is nothing to say."""
    if not best or score < SUGGEST_FLOOR:
        return None
    if fold(best) == fold(current):
        return None if best == current else 'minor'
    if score >= APPLY_FLOOR and is_improvement(current, best):
        return 'title'
    return 'suggest'


# Hosts that serve an output -- a video, a paper, a deposit -- never a page
# that announces an event. A link to the recording of a talk is not the
# conference's homepage, however well its title matches.
OUTPUT_HOSTS = (
    'youtube.com', 'youtu.be', 'doi.org', 'hdl.handle.net', 'zenodo.org',
    'arxiv.org', 'koushare.com', 'slideshare.net', 'figshare.com',
    'researchgate.net', 'vimeo.com', 'dpg-verhandlungen.de',
)
# How well a page title has to match the event name before the page counts as
# that event's own. Lower than the title floors: an event page legitimately
# adds a year, a city or 'Registration' to the name.
EVENT_URL_FLOOR = 0.55
# An Indico-style URL IS an event page: the platform has nothing else at that
# path. The title still has to be in the right area, but it need not match the
# sheet's wording, which is often an abbreviation of the official name.
EVENT_PLATFORM = re.compile(r'(?:^https?://(?:events?|indico)\.|/event/\d|/e/\d)', re.I)
PLATFORM_FLOOR = 0.30
# Indico deep links: .../event/648/contributions/1824/ is one talk, not the
# event. Trim back to the event root so the field points at the announcement.
EVENT_ROOT = re.compile(r'(/event/\d+)/(?:contributions?|timetable|sessions?|'
                        r'overview|registrations?|book-of-abstracts).*$', re.I)


def event_root(url: str) -> str:
    return EVENT_ROOT.sub(r'\1/', url)


def event_url_for(event: str, urls: list[str], pages: dict) -> tuple[str, float]:
    """The URL among these whose page most looks like it announces the event."""
    best, best_score = '', 0.0
    for url in urls:
        if any(host in url.lower() for host in OUTPUT_HOSTS):
            continue
        floor = PLATFORM_FLOOR if EVENT_PLATFORM.search(url) else EVENT_URL_FLOOR
        for candidate in pages.get(url, {}).get('titles', []):
            for variant in split_candidate(candidate):
                score = similarity(event, tidy_candidate(variant))
                if score >= floor and score > best_score:
                    best, best_score = event_root(url), score
    return (best, best_score) if best else ('', 0.0)


def main() -> int:
    if not TITLES.exists():
        raise SystemExit('run fetch_source_titles.py first')
    pages = json.loads(TITLES.read_text(encoding='utf-8'))

    # The rows as build_curated_sheet.py makes them on its own, before any
    # title match is applied -- reading the finished spreadsheet instead would
    # mean reading back this script's own previous corrections.
    records, _, _ = curate_all(DEFAULT_SOURCE)

    matches: dict[str, dict] = {}
    suggestions: list[dict] = []
    # event name -> (url, score). Keyed by event, not by row, because the URL
    # describes the event; the best-scoring row's page wins.
    event_urls: dict[str, tuple[str, float]] = {}
    counts = {'minor': 0, 'title': 0, 'suggest': 0, 'rows': 0}

    for record in records:
        def field(name: str, _record: dict = record) -> str:
            return re.sub(r'\s+', ' ', str(_record.get(name) or '')).strip()

        urls = cited_urls(field('Link to relevant output 1'), field('Source URLs'))
        if not urls:
            continue

        candidates = candidates_for(urls, pages)
        if not candidates:
            continue

        counts['rows'] += 1
        key = title_key(field('ID'), field('Name'))

        event = field('Canonical event name')
        url, score = event_url_for(event, urls, pages)
        if url and score > event_urls.get(event, ('', 0.0))[1]:
            event_urls[event] = (url, score)

        for column, current in (
            ('Contribution title', field('Contribution title')),
            ('Canonical event name', field('Canonical event name')),
        ):
            if not current:
                continue
            best, best_url, score = best_candidate(current, candidates)
            verdict = classify(current, best, score)
            if verdict is None:
                continue
            counts[verdict] += 1
            entry = {'from': current, 'to': best, 'similarity': round(score, 3),
                     'level': verdict, 'url': best_url}
            if verdict == 'suggest':
                suggestions.append({'key': key, 'column': column, **entry})
            else:
                matches.setdefault(key, {})[column] = entry

    OUT.write_text(json.dumps({
        '_generated': datetime.date.today().isoformat(),
        '_floors': {'typo': TYPO_FLOOR, 'apply': APPLY_FLOOR,
                    'suggest': SUGGEST_FLOOR},
        'matches': matches,
        'suggestions': suggestions,
        'event_urls': {name: {'url': url, 'similarity': round(score, 3)}
                       for name, (url, score) in sorted(event_urls.items())},
    }, indent=1, ensure_ascii=False), encoding='utf-8')

    print(f"{counts['rows']} rows had at least one usable candidate")
    print(f"  {counts['minor']} spelling fixes      (applied, light yellow)")
    print(f"  {counts['title']} reworded titles (applied, blue)")
    print(f"  {counts['suggest']} weaker suggestions (recorded, not applied)")
    print(f'  {len(event_urls)} events got an announcing page as their event URL')
    print(f'written to {OUT}')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
