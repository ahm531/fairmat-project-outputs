#!/usr/bin/env python3
"""Generate ``FAIRmatEvent`` ``*.archive.json`` files from the events spreadsheet.

This deliberately does **not** add a parser. The plugin's own parsers and schema
are left untouched; this script runs locally and the resulting archive files are
uploaded to the oasis, where NOMAD's built-in archive parser turns each one into
an editable ``FAIRmat Event`` entry.

The event objects are built with the plugin's own schema classes, so every
controlled vocabulary (event type, series, FAIRmat role, FaBiO class,
contribution type, ...) is validated while generating. A row that cannot be
mapped is reported, never silently written.

Usage (from the distro root)::

    uv run python packages/fairmat-project-outputs/scripts/generate_event_archives.py
    uv run python .../generate_event_archives.py local/other.xlsx -o local/out

Input and output live under ``local/`` on purpose: the sheet carries the names
of real people and ``local/`` is git-ignored.

Row model (agreed with the data owner)
--------------------------------------
Sheet ``Events``, columns ``B..M``: ID, From, To, Name, Speaker/Participant,
Event type (main), Event type (sub), Event type (old), FAIRmat role,
Link to relevant output 1, FaBio Outpu1, Educational Purpose.

Column ``Name`` carries the event/contribution hierarchy: a talk row reads
``<event name> Talk title: <talk title>`` (usually with a newline before the
marker). Rows are therefore split into an event part and an optional
contribution part, then **grouped by event name** so that one event with several
talks becomes ONE entry with several ``Contribution`` sub-sections.

Three deliberate decisions worth knowing about:

* ``event_type`` is inferred from keywords in the *event name* first, and only
  then from ``Event type (sub)`` / ``Event type (main)``. The sub-type describes
  the *contribution* ('S_Invited talks'), not the event, so trusting it first
  would label "APS March Meeting" a Seminar.
* Event-level ``fairmat_contributors`` is left unset whenever the event has
  contributions: ``FAIRmatEvent.normalize`` overwrites it with the names
  harvested from those contributions anyway.
* ``contributions[].contribution_role`` is written only where the source states
  a role of its own. The sheet has no per-contribution role information (no
  event has rows that disagree on the role column), so deriving it from the
  event's ``fairmat_role`` would only restate that field one level down.

The ``ID`` column is NOT used for hierarchy. Its ``x.y`` convention is unsound
in the source data (``373.1`` belongs to a different event than ``373``, and
``374.1``-``374.6`` are empty).
"""

from __future__ import annotations

import argparse
import json
import re
import sys
import unicodedata
import warnings
from datetime import date, datetime
from pathlib import Path

warnings.filterwarnings('ignore', module='openpyxl')

import openpyxl  # noqa: E402
from nomad.datamodel import EntryArchive, EntryMetadata  # noqa: E402

from fairmat_project_outputs.schema_packages.schema_package import (  # noqa: E402
    Contribution,
    FAIRmatEvent,
)

# Filename prefix. The plugin's parsers prefix with their own class
# ('generic_event_', 'dpg_event_'), which is why every Users_Meetings.csv
# archive is misleadingly called 'dpg_event_*'. Naming the source instead keeps
# the origin of an entry visible.
FILENAME_PREFIX = 'events_excel'

REPO_ROOT = Path(__file__).resolve().parents[1]
DEFAULT_INPUT = REPO_ROOT / 'local' / 'Events database - (September 2024).xlsx'
SHEET = 'Events'

# Column indices within the sheet row tuple (col A is an unused marker column).
C_ID, C_FROM, C_TO, C_NAME, C_SPEAKER = 1, 2, 3, 4, 5
C_MAIN, C_SUB, C_OLD, C_ROLE, C_LINK, C_FABIO, C_EDU = 6, 7, 8, 9, 10, 11, 12

# ---------------------------------------------------------------------------
# Column E -> (event name, contribution title)
# ---------------------------------------------------------------------------

# Strict form actually used by the data owner.
STRICT_MARKER = re.compile(
    r'^(?P<ev>.+?)\s*(?P<kind>talk|poster)\s+title\s*:\s*(?P<title>.+)$',
    re.IGNORECASE | re.DOTALL,
)

# Tolerant form, absorbing the five typo/variant spellings present in the sheet:
# 'Talke title:', 'talk titile:', 'talk title' (no colon), bare 'Title:'.
# Verified NOT to match ordinary titles that merely contain a colon, such as
# 'prompt:UX 2024', 'FAIRmat Tutorial 1: ...' or 'FAIRmat seminar: ...'.
LOOSE_MARKER = re.compile(
    r'^(?P<ev>.+?)\s*[-–]?\s*'
    r'(?P<kind>talke?|poster|presentation|lecture)?\s*'
    r'tit[li]?[lt]?e\s*:?\s+(?P<title>.+)$',
    re.IGNORECASE | re.DOTALL,
)

# Group B: contribution markers that are not the word "title". There is no safe
# general rule for these, so each is pinned by row ID. Keys are the ID cell as a
# string. Values: (event_name, contribution_title, contribution_type_or_None).
GROUP_B_OVERRIDES: dict[str, tuple[str, str, str | None]] = {
    '38': (
        'DPG_SKM21',
        'Exotic charge density wave states of matter: correlations and topology',
        'Conference session',
    ),
    '94': (
        'Psi-K conference 2022',
        'B8 | Materials discovery by high-throughput screening and artificial '
        'intelligence',
        'Symposium',
    ),
    '95': (
        'Psi-K conference 2022',
        'B5 | Machine-learned surrogate models: the quest for ab initio accuracy '
        'at a fraction of the cost',
        'Symposium',
    ),
    '311': (
        '18th International Congress on Catalysis',
        'Data as the key resource in digital catalysis',
        'Conference session',
    ),
    '312': (
        '18th International Congress on Catalysis',
        'Round Table “Data as the key resource in digital catalysis” '
        '(moderator)',
        'Panel discussion',
    ),
    '191': (
        'CyberWeek summer school 2023',
        'FAIR Data Management for Materials Science',
        'Tutorial',
    ),
    '252': (
        '5. Lüscher-Wassermann Seminar',
        'Session 5.2.24: New Concepts of Magnetism („Altermagnetism“)',
        'Conference session',
    ),
    '373.1': ('Psi-k conference 2025', 'Information booth', 'Information booth'),
    '377': ('QBIC VII', 'Information booth', 'Information booth'),
    '406': (
        '2026 MaRDA Annual Meeting',
        'AI-Ready Data and Materials Data Repositories',
        'Panel discussion',
    ),
}

# contribution_role is only set where the source says something the event-level
# `fairmat_role` does not already say. Deriving it from `fairmat_role` was tried
# and dropped: no event in the sheet has rows that disagree on the role column
# (0 of 363), so such a value would be `fairmat_role` restated one level down,
# looking like per-contribution information while carrying none.
# Row 312 is the only row in the sheet that states its own role: '(moderator)'.
CONTRIBUTION_ROLE_OVERRIDES = {'312': 'Moderation'}

# ---------------------------------------------------------------------------
# Vocabulary mapping
# ---------------------------------------------------------------------------

# Event name keyword -> EVENT_TYPE, most specific first.
NAME_KEYWORDS: list[tuple[str, str]] = [
    ('colloquium', 'Colloquium'),
    ('symposium', 'Symposium'),
    ('hackathon', 'Hackathon'),
    ('summer school', 'School'),
    ('winterschool', 'School'),
    ('winter school', 'School'),
    ('conference', 'Conference'),
    ('congress', 'Conference'),
    ('workshop', 'Workshop'),
    ('school', 'School'),
    ('tutorial', 'Tutorial'),
    ('course', 'Course'),
    ('seminar', 'Seminar'),
    ('webinar', 'Seminar'),
    ('colloquy', 'Colloquium'),
    ('meeting', 'Meeting'),
    ('tagung', 'Conference'),
    ('demonstration', 'Demonstration'),
    ('booth', 'Public event'),
]

# 'Event type (sub)' with its S_/Ce_/CM_/E_ prefix stripped -> EVENT_TYPE.
SUBTYPE_TO_EVENT_TYPE: dict[str, str] = {
    'fair-di - fairmat colloquium series': 'Colloquium',
    'fairmat seminar series': 'Seminar',
    'nfdi physical sciences joint colloquium': 'Colloquium',
    'invited talks': 'Seminar',
    'plenary talks': 'Conference',
    'contributed talks': 'Conference',
    'contributed presentations': 'Conference',
    'public workshops': 'Workshop',
    'fairmat internal workshops': 'Workshop',
    'hackathon': 'Hackathon',
    'tech. partners workshop series': 'Workshop',
    'project meeting': 'Meeting',
    'conference session': 'Conference',
    'users meeting': 'Meeting',
    'information booths': 'Public event',
    'fairmat tutorials series': 'Tutorial',
    'external tutorials': 'Tutorial',
    'schools': 'School',
    'course instance': 'Course',
    'lecture': 'Seminar',
    'demonstrations': 'Demonstration',
    'fair-di conference': 'Conference',
}

# Stripped sub-type -> EVENT_SERIES. Note the schema spells the colloquium
# series with a slash while the sheet uses a hyphen.
SUBTYPE_TO_SERIES: dict[str, str] = {
    'fair-di - fairmat colloquium series': 'FAIR-DI / FAIRmat colloquium series',
    'fairmat seminar series': 'FAIRmat seminar series',
    'nfdi physical sciences joint colloquium': 'NFDI Physical Sciences Joint Colloquium',
    'tech. partners workshop series': 'Tech. partners workshop series',
    'fairmat tutorials series': 'FAIRmat tutorials series',
    'project meeting': 'FAIRmat project meeting',
    'users meeting': 'FAIRmat users meeting',
}

# Stripped sub-type -> CONTRIBUTION_TYPE.
SUBTYPE_TO_CONTRIBUTION: dict[str, str] = {
    'invited talks': 'Invited talk',
    'plenary talks': 'Plenary talk',
    'contributed talks': 'Contributed talk',
    'contributed presentations': 'Contributed talk',
    'conference session': 'Conference session',
    'information booths': 'Information booth',
    'fairmat tutorials series': 'Tutorial',
    'external tutorials': 'Tutorial',
    'public workshops': 'Workshop session',
    'fairmat internal workshops': 'Workshop session',
}

MAIN_TO_EVENT_TYPE: dict[str, str] = {
    'scholarly talks (s)': 'Seminar',
    'community engagement (ce)': 'Workshop',
    'conferences and meetings (cm)': 'Meeting',
    'educational events (e)': 'Tutorial',
    'public engagement (pe)': 'Public event',
}

ROLE_MAP: dict[str, str] = {
    'organizer': 'Organizer',
    'co-organizer': 'Co-organizer',
    'coorganizer': 'Co-organizer',
    'supporter': 'Supporter',
    'support': 'Supporter',
    'participant': 'Participant',
    'speaker': 'Participant',
    'host': 'Organizer',
}

FABIO_CANONICAL = [
    'fabio: Abstract', 'fabio: Announcement', 'fabio: Book chapter',
    'fabio: Computer program', 'fabio: Conference proceedings',
    'fabio: Entity metadata', 'fabio: Instructional work', 'fabio: Interview',
    'fabio: Journal article', 'fabio: Meeting report', 'fabio: Periodical issue',
    'fabio: Poster', 'fabio: Preprint', 'fabio: Presentation',
    'fabio: Scholarly work', 'fabio: Software dataset', 'fabio: Timetable',
    'fabio: Web page',
]
FABIO_BY_SUFFIX = {v.split(': ', 1)[1].lower(): v for v in FABIO_CANONICAL}

SUB_PREFIX = re.compile(r'^(?:S|Ce|CM|E)_', re.IGNORECASE)
SKIP_CONTRIBUTORS = {'various', 'n/a', 'na', 'tbd', 'tbc', '-'}
TRUTHY = {'yes', 'true', '1', 'y', 'x'}


# ---------------------------------------------------------------------------
# Cell helpers
# ---------------------------------------------------------------------------


def clean(value) -> str:
    """Collapse whitespace and normalise the empty markers to ''."""
    if value is None:
        return ''
    text = re.sub(r'\s+', ' ', str(value).replace('\r', ' ')).strip()
    return '' if text.lower() in ('nan', 'none', '') else text


def cell(row: tuple, idx: int) -> str:
    return clean(row[idx]) if idx < len(row) else ''


def as_datetime(value) -> datetime | None:
    if isinstance(value, datetime):
        return value
    text = clean(value)
    if not text:
        return None
    for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d-%b-%y', '%d-%b-%Y',
                '%d/%m/%Y', '%d.%m.%Y'):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def split_names(raw: str) -> list[str] | None:
    names = [n.strip() for n in re.split(r'[,;]', raw) if n.strip()]
    names = [n for n in names if n.lower() not in SKIP_CONTRIBUTORS]
    return names or None


def strip_sub_prefix(raw: str) -> str:
    return SUB_PREFIX.sub('', raw).strip()


def normalise_fabio(raw: str) -> str | None:
    """'[fabio:Web page]' / '[fabio: announcement]' -> canonical FABIO_TERMS."""
    if not raw:
        return None
    text = re.sub(r'^\[|\]$', '', raw.strip())
    text = re.sub(r'(?i)^fabio\s*:\s*', '', text).strip()
    return FABIO_BY_SUFFIX.get(text.lower())


def event_type_from_name(name: str) -> str | None:
    lowered = name.lower()
    for keyword, etype in NAME_KEYWORDS:
        if keyword in lowered:
            return etype
    return None


def slugify(text: str) -> str:
    ascii_text = (
        unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode()
    )
    return re.sub(r'_+', '_', re.sub(r'[^\w]+', '_', ascii_text)).strip('_')[:70]


def tidy_event_name(name: str) -> str:
    """Drop the separator left behind when the title marker is stripped.

    ``Psi-k Conference 2022 - Talk title: ...`` leaves a dangling ``-``.
    """
    return name.strip().strip('-–—:;,.').strip()


def event_key(name: str) -> str:
    """Grouping key that ignores case, quoting and punctuation.

    Merges ``CECAM Flagship Workshop on "X"`` with the unquoted spelling, and
    ``Psi-K conference 2022`` with ``Psi-k Conference 2022``.
    """
    ascii_name = (
        unicodedata.normalize('NFKD', name).encode('ascii', 'ignore').decode()
    )
    return re.sub(r'\s+', ' ', re.sub(r'[^a-z0-9 ]', '', ascii_name.lower())).strip()


# ---------------------------------------------------------------------------
# Row -> (event name, contribution title, review reason)
# ---------------------------------------------------------------------------


def split_name_cell(raw_name: str, row_id: str) -> tuple[str, str | None, str | None]:
    """Return (event_name, contribution_title, review_reason)."""
    if row_id in GROUP_B_OVERRIDES:
        event_name, title, _ = GROUP_B_OVERRIDES[row_id]
        return event_name, title, 'adjudicated: contribution marker is not the word "title"'

    text = clean(raw_name)
    strict = STRICT_MARKER.match(text)
    if strict and strict.group('ev').strip():
        return strict.group('ev').strip(), strict.group('title').strip(), None

    loose = LOOSE_MARKER.match(text)
    if loose and loose.group('ev').strip():
        return (
            loose.group('ev').strip(),
            loose.group('title').strip(),
            'adjudicated: non-standard spelling of the "Talk title:" marker',
        )

    return text, None, None


def contribution_type_for(
    row_id: str, sub_clean: str, raw_name: str
) -> str | None:
    if row_id in GROUP_B_OVERRIDES:
        return GROUP_B_OVERRIDES[row_id][2]
    if re.search(r'poster\s+title', raw_name, re.IGNORECASE):
        return 'Poster'
    return SUBTYPE_TO_CONTRIBUTION.get(sub_clean.lower())


# ---------------------------------------------------------------------------
# Grouping
# ---------------------------------------------------------------------------


def read_rows(path: Path) -> list[tuple]:
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    if SHEET not in workbook.sheetnames:
        raise SystemExit(f'sheet {SHEET!r} not found in {path} ({workbook.sheetnames})')
    return list(workbook[SHEET].iter_rows(values_only=True))[1:]


def merge_event_fields(
    info: dict,
    row: tuple,
    sub_clean: str,
    row_id: str,
    warnings_out: list[str],
) -> None:
    """Fill the event-level fields; the first non-empty value across rows wins."""
    if info['event_type'] is None:
        info['event_type'] = (
            event_type_from_name(info['event_name'])
            or SUBTYPE_TO_EVENT_TYPE.get(sub_clean.lower())
            or MAIN_TO_EVENT_TYPE.get(cell(row, C_MAIN).lower())
        )
    if info['event_series'] is None:
        info['event_series'] = SUBTYPE_TO_SERIES.get(sub_clean.lower())
    if info['fairmat_role'] is None:
        raw_role = cell(row, C_ROLE)
        if raw_role:
            mapped = ROLE_MAP.get(raw_role.lower())
            if mapped is None:
                warnings_out.append(
                    f'row {row_id}: unknown FAIRmat role {raw_role!r}, left empty'
                )
            info['fairmat_role'] = mapped
    if info['event_purpose'] is None and cell(row, C_EDU).lower() in TRUTHY:
        info['event_purpose'] = ['Educational event']


def group_events(rows: list[tuple], warnings_out: list[str]) -> tuple[dict, int]:
    """Group sheet rows into ``{normalised event name: event dict}``."""
    events: dict[str, dict] = {}
    skipped = 0

    for row in rows:
        raw_name = row[C_NAME] if len(row) > C_NAME else None
        if not clean(raw_name):
            # The 14 ID-only scaffolding rows (178-183, 374, 374.1-374.6, 381).
            if any(clean(c) for i, c in enumerate(row) if i not in (0, C_ID)):
                warnings_out.append(f'row with data but no Name dropped: {row!r}')
            skipped += 1
            continue

        raw_id = row[C_ID]
        row_id = clean(raw_id)
        if isinstance(raw_id, float) and raw_id.is_integer():
            row_id = str(int(raw_id))

        event_name, title, review = split_name_cell(raw_name, row_id)
        event_name = tidy_event_name(event_name)
        sub_clean = strip_sub_prefix(cell(row, C_SUB) or cell(row, C_OLD))
        start, end = as_datetime(row[C_FROM]), as_datetime(row[C_TO])

        key = event_key(event_name)
        info = events.setdefault(
            key,
            {
                'event_name': event_name,
                'start': None,
                'end': None,
                'event_type': None,
                'event_series': None,
                'fairmat_role': None,
                'event_purpose': None,
                'event_url': None,
                'contributions': [],
                'plain_contributors': None,
                'reviews': [],
                'row_ids': [],
            },
        )
        info['row_ids'].append(row_id)
        if review:
            info['reviews'].append(f'row {row_id}: {review}')

        # Widest date span across every row of the event.
        if start and (info['start'] is None or start < info['start']):
            info['start'] = start
        if end and (info['end'] is None or end > info['end']):
            info['end'] = end

        merge_event_fields(info, row, sub_clean, row_id, warnings_out)

        link = cell(row, C_LINK)
        if title is None:
            # A plain event row: its link describes the event itself.
            if info['event_url'] is None and link:
                info['event_url'] = link
            if info['plain_contributors'] is None:
                info['plain_contributors'] = split_names(cell(row, C_SPEAKER))
            continue

        fabio_raw = cell(row, C_FABIO)
        fabio = normalise_fabio(fabio_raw)
        if fabio_raw and fabio is None:
            warnings_out.append(
                f'row {row_id}: unrecognised FaBiO value {fabio_raw!r}, left empty'
            )

        info['contributions'].append(
            {
                'title': title,
                'start_date': start,
                'end_date': end,
                'contribution_type': contribution_type_for(
                    row_id, sub_clean, clean(raw_name)
                ),
                'contribution_role': CONTRIBUTION_ROLE_OVERRIDES.get(row_id),
                'fairmat_contributors': split_names(cell(row, C_SPEAKER)),
                'fabio_type': fabio,
                'fabio_url': link or None,
            }
        )

    return events, skipped


def build_event(info: dict) -> FAIRmatEvent:
    contributions = [Contribution(**c) for c in info['contributions']]
    event = FAIRmatEvent(
        event_name=info['event_name'],
        start_date=info['start'],
        end_date=info['end'],
        event_type=info['event_type'],
        event_series=info['event_series'],
        fairmat_role=info['fairmat_role'],
        event_purpose=info['event_purpose'],
        contributions=contributions,
    )
    if info['event_url']:
        event.event_url = info['event_url']
    # Only meaningful without contributions: FAIRmatEvent.normalize() replaces
    # this with the names collected from the contributions otherwise.
    if not contributions and info['plain_contributors']:
        event.fairmat_contributors = info['plain_contributors']
    return event


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------


def write_report(
    out_dir: Path,
    source: Path,
    stats: dict[str, int],
    flagged: list[tuple[str, list[str]]],
    problems: list[str],
) -> Path:
    """Write ``review_report.md`` and return its path."""
    report = out_dir / 'review_report.md'
    with open(report, 'w', encoding='utf-8') as handle:
        handle.write('# Review report\n\n')
        handle.write(f'Source: `{source.name}`\n\n')
        handle.write(
            f'- {stats["created"]} entries written from '
            f'{stats["usable"]} usable rows\n'
            f'- {stats["skipped"]} empty rows skipped, '
            f'{stats["errors"]} events failed\n'
            f'- {len(flagged)} entries need a human look\n\n'
        )
        handle.write('## Entries flagged for review\n\n')
        if flagged:
            handle.write(
                'These are stamped with a `NEEDS REVIEW` comment in NOMAD.\n\n'
            )
            for event_name, reasons in sorted(flagged):
                handle.write(f'### {event_name}\n\n')
                for reason in reasons:
                    handle.write(f'- {reason}\n')
                handle.write('\n')
        else:
            handle.write('None.\n\n')
        handle.write('## Mapping warnings\n\n')
        handle.write(
            ''.join(f'- {p}\n' for p in problems) if problems else 'None.\n'
        )
    return report


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.split('\n')[0])
    parser.add_argument(
        'excel_file', nargs='?', default=str(DEFAULT_INPUT),
        help=f'input spreadsheet (default: {DEFAULT_INPUT})',
    )
    parser.add_argument(
        '-o', '--output',
        default=str(REPO_ROOT / 'local' / f'archives_{date.today():%Y%m%d}'),
        help='output directory (default: local/archives_<today>)',
    )
    args = parser.parse_args()

    source = Path(args.excel_file)
    if not source.exists():
        print(f'ERROR input not found: {source}', file=sys.stderr)
        return 1

    out_dir = Path(args.output)
    out_dir.mkdir(parents=True, exist_ok=True)

    problems: list[str] = []
    rows = read_rows(source)
    events, skipped = group_events(rows, problems)

    created = errors = 0
    used: set[str] = set()
    flagged: list[tuple[str, list[str]]] = []

    for info in events.values():
        try:
            event = build_event(info)
        except Exception as exc:  # noqa: BLE001 - report and continue
            errors += 1
            print(
                f'ERROR event {info["event_name"]!r} (rows '
                f'{", ".join(info["row_ids"])}): {exc}',
                file=sys.stderr,
            )
            continue

        metadata = EntryMetadata(entry_name=info['event_name'])
        if info['reviews']:
            metadata.comment = 'NEEDS REVIEW - ' + '; '.join(info['reviews'])
            flagged.append((info['event_name'], info['reviews']))

        stem = slugify(info['event_name']) or 'event'
        name = f'{FILENAME_PREFIX}_{stem}'
        if name in used:
            name = f'{name}_{len(used)}'
        used.add(name)

        archive = EntryArchive(data=event, metadata=metadata)
        with open(out_dir / f'{name}.archive.json', 'w', encoding='utf-8') as handle:
            json.dump(archive.m_to_dict(), handle, indent=2, ensure_ascii=False)
        created += 1

    report = write_report(
        out_dir,
        source,
        {
            'created': created,
            'usable': len(rows) - skipped,
            'skipped': skipped,
            'errors': errors,
        },
        flagged,
        problems,
    )

    for problem in problems:
        print(f'WARNING {problem}', file=sys.stderr)

    print(
        f'{created} archive files written to {out_dir} '
        f'({skipped} empty rows skipped, {errors} events failed, '
        f'{len(flagged)} flagged for review, {len(problems)} warnings)\n'
        f'review report: {report}'
    )
    return 1 if errors else 0


if __name__ == '__main__':
    sys.exit(main())
