#!/usr/bin/env python3
"""Build the curated events sheet from the master spreadsheet.

One sheet, one source of truth. Every original column is carried through
untouched; every judgement this script makes lands in a **new column beside**
it, so the data owner can compare the two side by side and sign off before
anything is turned into archives.

What it adds to each row
------------------------
``Canonical event name``  the grouping key -- one event per entry in NOMAD.
                          Recurring events whose name carries no year get the
                          year appended, so the 2022 and 2023 QUANTSOL schools
                          stop collapsing into one entry that spans both.
``Contribution title``    the part after the "Talk title:" marker.
``Contribution type``     Invited talk / Contributed talk / Poster / ...
``Event type``            Conference / Workshop / School / ...
``Start date (curated)``  corrected dates, where research established a
``End date (curated)``    correction. The original From/To are left alone.
``Location name`` ``City`` ``Country`` ``Mode``
                          researched venue data (``event_venues_researched.json``).
``FAIRmat contributors``  speakers matched against the FAIRmat roster,
``Other contributors``    speakers that matched nothing -- the outside guests.
``FaBiO (canonical)``     one spelling per concept, FaBiO's own lowercase form.
``Source``                which row of which file this came from.
``Curation note``         what was changed and why.
``Needs check``           ``x`` where a human should look.

The 65 Users-Meeting contributions that exist only in ``Cordi_events.csv`` are
merged in with ``ID = new``, matching the convention already used by rows
510-536 of the master sheet.

Usage::

    uv run python packages/fairmat-project-outputs/scripts/build_curated_sheet.py
"""

from __future__ import annotations

import argparse
import csv
import datetime
import difflib
import json
import re
import sys
import unicodedata
import warnings
from collections import Counter, defaultdict
from pathlib import Path

warnings.filterwarnings('ignore', module='openpyxl')

import openpyxl  # noqa: E402
from openpyxl.styles import Alignment, Font, PatternFill  # noqa: E402
from openpyxl.utils import get_column_letter  # noqa: E402

sys.path.insert(0, str(Path(__file__).resolve().parent))
from fairmat_roster import Roster  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
LOCAL = ROOT / 'local'
DEFAULT_SOURCE = LOCAL / 'Events database - (September 2024)_full_20260904.xlsx'
ORPHAN_CSV = LOCAL / 'Cordi_events.csv'
RESEARCH = LOCAL / 'event_venues_researched.json'
DERIVED = LOCAL / 'event_venues_derived.json'
TITLE_MATCHES = LOCAL / 'title_matches.json'
SHEET = 'Events'

ORIGINAL_COLUMNS = [
    'ID', 'From', 'To', 'Name', 'Speaker/Participant',
    'Event type (main)', 'Event type (sub)', 'Event type (old)',
    'FAIRmat role', 'Link to relevant output 1', 'FaBio Outpu1',
    'Educational Purpose',
]

# The sheet is laid out for review: each original column is followed
# immediately by the curated column derived from it, so a reviewer compares a
# pair side by side instead of scrolling between two blocks. ORIGINAL is grey,
# CURATED is green. Columns with no counterpart stand alone; the curated
# columns that have no original at all (location, mode) come after the pairs,
# and the URLs backing every judgement are the very last column.
# Column layout for review. Each entry is (internal key, ORIGINAL|CURATED,
# header shown in the sheet).
#
# The header carries the pairing explicitly: the twelve source columns are
# lettered A..L in the order they appear in the master spreadsheet, and every
# curated column derived from one carries the SAME letter with a number --
# D / D1 / D2 all come from 'Name'. Columns that have no counterpart in any
# source file are prefixed '+'. Names are kept identical to the master sheet
# and to Ahmed's CSVs wherever those files already had a name for the thing
# (Location / City / Country come from Cordi_events.csv).
LAYOUT: list[tuple[str, str, str]] = [
    ('ID',                        'ORIGINAL', 'A | ID'),

    ('From',                      'ORIGINAL', 'B | From'),
    ('Start date (curated)',       'CURATED', 'B1 | From (curated)'),

    ('To',                        'ORIGINAL', 'C | To'),
    ('End date (curated)',         'CURATED', 'C1 | To (curated)'),

    ('Name',                      'ORIGINAL', 'D | Name'),
    ('Canonical event name',       'CURATED', 'D1 | Name - event'),
    ('Contribution title',         'CURATED', 'D2 | Name - contribution'),

    ('Speaker/Participant',       'ORIGINAL', 'E | Speaker/Participant'),
    ('FAIRmat contributors',       'CURATED', 'E1 | Speaker - FAIRmat'),
    ('Other contributors',         'CURATED', 'E2 | Speaker - external'),

    ('Event type (main)',         'ORIGINAL', 'F | Event type (main)'),
    ('Event type',                 'CURATED', 'F1 | Event type (curated)'),

    ('Event type (sub)',          'ORIGINAL', 'G | Event type (sub)'),
    ('Contribution type',          'CURATED', 'G1 | Contribution type'),
    ('Contribution role',          'CURATED', 'G2 | Contribution role'),

    ('Event type (old)',          'ORIGINAL', 'H | Event type (old)'),
    ('FAIRmat role',              'ORIGINAL', 'I | FAIRmat role'),
    ('Link to relevant output 1', 'ORIGINAL', 'J | Link to relevant output 1'),

    ('FaBio Outpu1',              'ORIGINAL', 'K | FaBio Outpu1'),
    ('FaBiO (canonical)',          'CURATED', 'K1 | FaBio Outpu1 (curated)'),

    ('Educational Purpose',       'ORIGINAL', 'L | Educational Purpose'),
    # Purpose is read off three columns at once -- F, G1 and L -- so it belongs
    # to no single pair, but it sits next to L because that is the column a
    # reviewer checks it against. '(row)' is what this row alone implies;
    # '(curated)' is the union over every row of the same event, which is what
    # the schema stores, since event_purpose is a property of the event.
    ('Event purpose (row)',        'CURATED', '+ | Event purpose (row)'),
    ('Event purpose',              'CURATED', '+ | Event purpose (curated)'),
    # Both are properties of the event, not of the row, so every row of an
    # event carries the same value. Series is read off column G, which names
    # it outright; the URL is the page found to announce the event.
    ('Event series',               'CURATED', '+ | Event series'),
    ('Event URL',                  'CURATED', '+ | Event URL'),

    ('Location name',              'CURATED', '+ | Location'),
    ('City',                       'CURATED', '+ | City'),
    ('Country',                    'CURATED', '+ | Country'),
    ('Mode',                       'CURATED', '+ | Mode'),
    ('Location source',            'CURATED', '+ | Location source'),
    ('Source',                     'CURATED', '+ | Row source'),
    ('Curation note',              'CURATED', '+ | Curation note'),
    ('Needs check',                'CURATED', '+ | Needs check'),
    ('Source URLs',                'CURATED', '+ | Source URLs'),
]
COLUMNS = [key for key, _, _ in LAYOUT]
HEADERS = {key: header for key, _, header in LAYOUT}
KINDS = {key: kind for key, kind, _ in LAYOUT}
CURATED_COLUMNS = [key for key, kind, _ in LAYOUT if kind == 'CURATED']

# --------------------------------------------------------------------------
# Column E -> (event name, contribution title). Same rules as the archive
# generator, kept in step with it deliberately.
# --------------------------------------------------------------------------

STRICT_MARKER = re.compile(
    r'^(?P<ev>.+?)\s*(?P<kind>talk|poster)\s+title\s*:\s*(?P<title>.+)$',
    re.IGNORECASE | re.DOTALL,
)
LOOSE_MARKER = re.compile(
    r'^(?P<ev>.+?)\s*[-–]?\s*'
    r'(?P<kind>talke?|poster|presentation|lecture|workshop)?\s*'
    r'tit[li]?[lt]?e\s*:?\s+(?P<title>.+)$',
    re.IGNORECASE | re.DOTALL,
)
# The curated CSV writes '<event> talk: <title>' instead of 'Talk title:'.
# 'workshop' is deliberately NOT in this list: 'PSinNFDI workshop:' and
# 'CECAM flagship workshop:' are event names, not contribution markers. The one
# row that really does use it is pinned in NAME_OVERRIDES.
CSV_MARKER = re.compile(
    r'^(?P<ev>.+?)\s+(?P<kind>talk|poster|tutorial|exhibition|symposium)\s*:\s*(?P<title>.+)$',
    re.IGNORECASE | re.DOTALL,
)

# Contribution markers that are not the word "title". No safe general rule
# exists, so each is pinned by row ID: (event, title, contribution type).
ID_OVERRIDES: dict[str, tuple[str, str, str | None]] = {
    '38': ('DPG_SKM21', 'Exotic charge density wave states of matter: correlations and topology', 'Conference session'),
    '94': ('Psi-K conference 2022', 'B8 | Materials discovery by high-throughput screening and artificial intelligence', 'Symposium'),
    '95': ('Psi-K conference 2022', 'B5 | Machine-learned surrogate models: the quest for ab initio accuracy at a fraction of the cost', 'Symposium'),
    '191': ('CyberWeek summer school 2023', 'FAIR Data Management for Materials Science', 'Tutorial'),
    '252': ('5. Lüscher-Wassermann Seminar', 'Session 5.2.24: New Concepts of Magnetism („Altermagnetism“)', 'Conference session'),
    '311': ('18th International Congress on Catalysis', 'Data as the key resource in digital catalysis', 'Conference session'),
    '312': ('18th International Congress on Catalysis', 'Round Table “Data as the key resource in digital catalysis” (moderator)', 'Panel discussion'),
    '373.1': ('Psi-k conference 2025', 'Information booth', 'Information booth'),
    '377': ('QBIC VII', 'Information booth', 'Information booth'),
    '406': ('2026 MaRDA Annual Meeting', 'AI-Ready Data and Materials Data Repositories', 'Panel discussion'),
}

# Rows whose event/contribution split cannot be read off the text, because the
# two sources name the same booth four different ways. Keyed by the exact name.
NAME_OVERRIDES: dict[str, tuple[str, str]] = {
    'cordi 2025 information booth': (
        'Conference on Research Data Infrastructure (CoRDI 2025)',
        'Market place of the consortia: FAIRmat information booth'),
    'conference on research data infrastructure cordi 2025 market place of the consortia fairmat information booth': (
        'Conference on Research Data Infrastructure (CoRDI 2025)',
        'Market place of the consortia: FAIRmat information booth'),
    'cordi 2023 market place of the consortia fairmat infromation': (
        'Conference on Research Data Infrastructure (CoRDI 2023)',
        'Market place of the consortia: FAIRmat information booth'),
    'conference on research data infrastructure cordi 2023 market place of the consortia fairmat infromation booth': (
        'Conference on Research Data Infrastructure (CoRDI 2023)',
        'Market place of the consortia: FAIRmat information booth'),
    'e mrs spring meeting fairmat booth': (
        'E-MRS Spring Meeting 2023', 'FAIRmat information booth'),
    'e mrs spring meeting 2023 information booth': (
        'E-MRS Spring Meeting 2023', 'FAIRmat information booth'),
    'emrs spring meeting 2025 fairmat information booth': (
        'E-MRS Spring Meeting 2025', 'FAIRmat information booth'),
    'e mrs spring meeting 2025 fairmat information booth': (
        'E-MRS Spring Meeting 2025', 'FAIRmat information booth'),
    'mc conference karlsruhe information booth': (
        'MC Conference Karlsruhe', 'Information booth'),
    'mc conference karlsruhe workshop research data management in microscopy': (
        'MC Conference Karlsruhe', 'Research data management in microscopy'),
}

# FAIRmat's role in a single contribution, where the source states one of its
# own. Only row 312 does -- 'Round Table ... (moderator)'. Everywhere else the
# event-level FAIRmat role already says it, and repeating it one level down
# would look like per-contribution information while carrying none.
CONTRIBUTION_ROLE_OVERRIDES = {'312': 'Moderation'}

# Editorial renames. Not derivable -- these are decisions, so they live here in
# plain sight rather than inside a regex.
RENAMES: list[tuple[re.Pattern, str]] = [
    (re.compile(r'^1st Conference on Research Data Infrastructure \(CoRDI 2023\)$', re.I),
     'Conference on Research Data Infrastructure (CoRDI 2023)'),
    (re.compile(r'^CoRDI 2023$', re.I), 'Conference on Research Data Infrastructure (CoRDI 2023)'),
    (re.compile(r'^CorDI Conference$', re.I), 'Conference on Research Data Infrastructure (CoRDI 2023)'),
    (re.compile(r'^CoRDI 2025( Information Booth)?$', re.I),
     'Conference on Research Data Infrastructure (CoRDI 2025)'),
    (re.compile(r'^EMRS Spring Meeting 2025$', re.I), 'E-MRS Spring Meeting 2025'),
    (re.compile(r'^E-MRS Meeting France$', re.I), 'E-MRS Spring Meeting 2023'),
    (re.compile(r'^E-MRS Spring Meeting$', re.I), 'E-MRS Spring Meeting 2023'),
    (re.compile(r'^E-MRS spring meeting( 2023)?$', re.I), 'E-MRS Spring Meeting 2023'),
    (re.compile(r'^E-MRS spring meeting FAIRmat booth$', re.I), 'E-MRS Spring Meeting 2023'),
    (re.compile(r'^APS March Meeting 2021 \(March 15-21\)$', re.I), 'APS March Meeting 2021'),
]

# Events that recur under one name. Their rows must be split per edition, or a
# single entry ends up spanning several years. Confirmed one by one against the
# rows' own links -- see event_venues_researched.json.
RECURRING = [
    'aps march meeting',
    'quantsol summer school on solar cells',
    'quantsol winterschool on characterization of pv materials',
    'unisyscat big-nse basic lecture program',
    'on-site llm hackathon for applications in materials and chemistry in berlin',
    'cecam flagship workshop: open databases integration for materials design',
    'psinnfdi workshop: unlocking the potential of data',
]

SUBTYPE_TO_CONTRIBUTION = {
    'invited talks': 'Invited talk', 'plenary talks': 'Plenary talk',
    'contributed talks': 'Contributed talk', 'contributed presentations': 'Contributed talk',
    'conference session': 'Conference session', 'information booths': 'Information booth',
    'fairmat tutorials series': 'Tutorial', 'external tutorials': 'Tutorial',
    'public workshops': 'Workshop session', 'fairmat internal workshops': 'Workshop session',
}
# Fallback when 'Event type (sub)' says nothing useful about the contribution
# -- the Users-Meeting rows are all 'CM_Users meeting', which describes the
# event, not the talk. The marker word in the name does describe the talk.
MARKER_TO_CONTRIBUTION = {
    'talk': 'Contributed talk', 'poster': 'Poster', 'tutorial': 'Tutorial',
    'workshop': 'Workshop session', 'symposium': 'Symposium',
    'exhibition': 'Information booth', 'presentation': 'Contributed talk',
    'lecture': 'Contributed talk',
}

SUBTYPE_TO_EVENT_TYPE = {
    'fair-di - fairmat colloquium series': 'Colloquium', 'fairmat seminar series': 'Seminar',
    'nfdi physical sciences joint colloquium': 'Colloquium', 'invited talks': 'Seminar',
    'plenary talks': 'Conference', 'contributed talks': 'Conference',
    'contributed presentations': 'Conference', 'public workshops': 'Workshop',
    'fairmat internal workshops': 'Workshop', 'hackathon': 'Hackathon',
    'tech. partners workshop series': 'Workshop', 'project meeting': 'Meeting',
    'conference session': 'Conference', 'users meeting': 'Meeting',
    'information booths': 'Public event', 'fairmat tutorials series': 'Tutorial',
    'external tutorials': 'Tutorial', 'schools': 'School', 'course instance': 'Course',
    'lecture': 'Seminar', 'demonstrations': 'Demonstration', 'fair-di conference': 'Conference',
}
NAME_KEYWORDS = [
    ('colloquium', 'Colloquium'), ('symposium', 'Symposium'), ('hackathon', 'Hackathon'),
    ('summer school', 'School'), ('winterschool', 'School'), ('winter school', 'School'),
    ('conference', 'Conference'), ('congress', 'Conference'), ('workshop', 'Workshop'),
    ('school', 'School'), ('tutorial', 'Tutorial'), ('course', 'Course'),
    ('seminar', 'Seminar'), ('webinar', 'Seminar'), ('meeting', 'Meeting'),
    ('tagung', 'Conference'), ('demonstration', 'Demonstration'), ('booth', 'Public event'),
]
MAIN_TO_EVENT_TYPE = {
    'scholarly talks (s)': 'Seminar', 'community engagement (ce)': 'Workshop',
    'conferences and meetings (cm)': 'Meeting', 'educational events (e)': 'Tutorial',
    'public engagement (pe)': 'Public event',
}

# Every value is spelled as its ontology spells it: prefix + the class's own
# rdfs:label, checked against the current release of http://purl.org/spar/fabio
# (254 classes). FaBiO labels are lowercase; BIBO capitalises its own.
#
# Three spellings the master sheet used are not FaBiO classes at all, so they
# could never resolve to a URI: 'poster' -> conference poster, 'software
# dataset' -> dataset, and 'interview', which FaBiO has no class for -- BIBO
# does, and is FaBiO's usual companion ontology.
FABIO_CANONICAL = {
    'abstract': 'fabio: abstract', 'announcement': 'fabio: announcement',
    'book chapter': 'fabio: book chapter', 'computer program': 'fabio: computer program',
    'conference proceedings': 'fabio: conference proceedings',
    'conference poster': 'fabio: conference poster', 'poster': 'fabio: conference poster',
    'dataset': 'fabio: dataset', 'software dataset': 'fabio: dataset',
    'entity metadata': 'fabio: entity metadata',
    'instructional work': 'fabio: instructional work',
    'interview': 'bibo: Interview',
    'journal article': 'fabio: journal article', 'meeting report': 'fabio: meeting report',
    'periodical issue': 'fabio: periodical issue', 'preprint': 'fabio: preprint',
    'presentation': 'fabio: presentation', 'scholarly work': 'fabio: scholarly work',
    'timetable': 'fabio: timetable', 'web page': 'fabio: web page',
}
# --------------------------------------------------------------------------
# Event purpose
#
# The curation owners stated four rules, reading the event category (column F)
# together with the curated contribution type (G1). Those four left 106 of the
# 597 rows with no purpose at all, and the three decisions below close the gap:
#   * a workshop session or tutorial teaches whatever category it sits in, so
#     Community engagement counts as educational alongside Conferences;
#   * 'Plenary talk' and 'Symposium' join the scholarly contributions, which
#     the stated list omitted;
#   * a row naming no contribution type falls back to its event category
#     alone. That is a weaker signal, so those cells are flagged for a human.
# --------------------------------------------------------------------------

PURPOSE_EDUCATIONAL = 'Educational event'
PURPOSE_SCHOLARLY = 'Scholarly event'
PURPOSE_INFORMATIONAL = 'Informational event'
# Fixed order, so a multi-valued cell reads the same way in every row.
PURPOSE_ORDER = (PURPOSE_EDUCATIONAL, PURPOSE_SCHOLARLY, PURPOSE_INFORMATIONAL)

MAIN_EDUCATIONAL = 'educational events (e)'
MAIN_CONFERENCES = 'conferences and meetings (cm)'
MAIN_SCHOLARLY = 'scholarly talks (s)'
MAIN_COMMUNITY = 'community engagement (ce)'
MAIN_PUBLIC = 'public engagement (pe)'

TEACHING_CONTRIBUTIONS = {'Tutorial', 'Workshop session'}
SCHOLARLY_CONTRIBUTIONS = {
    'Invited talk', 'Contributed talk', 'Plenary talk', 'Poster',
    'Panel discussion', 'Conference session', 'Symposium',
}
BOOTH_CONTRIBUTIONS = {'Information booth'}

TEACHING_CATEGORIES = {MAIN_EDUCATIONAL, MAIN_CONFERENCES, MAIN_COMMUNITY}
SCHOLARLY_CATEGORIES = {MAIN_SCHOLARLY, MAIN_CONFERENCES}
BOOTH_CATEGORIES = {MAIN_PUBLIC, MAIN_COMMUNITY, MAIN_CONFERENCES}

# Used only when the row names no contribution type at all.
MAIN_ONLY_PURPOSE = {
    MAIN_SCHOLARLY: PURPOSE_SCHOLARLY,
    MAIN_CONFERENCES: PURPOSE_SCHOLARLY,
    MAIN_COMMUNITY: PURPOSE_INFORMATIONAL,
    MAIN_PUBLIC: PURPOSE_INFORMATIONAL,
}
TRUTHY = {'yes', 'true', '1', 'y', 'x'}

# Column G names the series outright wherever a row belongs to one, so this is
# a reading of the sheet rather than a guess about it. 'NFDI Physical Sciences
# Workshop' is in the schema but appears nowhere in the data.
SUBTYPE_TO_SERIES = {
    'fair-di - fairmat colloquium series': 'FAIR-DI / FAIRmat colloquium series',
    'fairmat seminar series': 'FAIRmat seminar series',
    'nfdi physical sciences joint colloquium': 'NFDI Physical Sciences Joint Colloquium',
    'nfdi physical sciences workshop': 'NFDI Physical Sciences Workshop',
    'tech. partners workshop series': 'Tech. partners workshop series',
    'fairmat tutorials series': 'FAIRmat tutorials series',
    'project meeting': 'FAIRmat project meeting',
    'users meeting': 'FAIRmat users meeting',
}

# The master spells column I four different ways for the same four roles, and
# the schema accepts only one of each. This is the one grey column the sheet
# corrects in place, on the curation owners' instruction; every corrected cell
# is painted so the change is not silent.
ROLE_CANONICAL = {
    'organizer': 'Organizer', 'organiser': 'Organizer',
    'co-organizer': 'Co-organizer', 'co-organiser': 'Co-organizer',
    'coorganizer': 'Co-organizer', 'co organizer': 'Co-organizer',
    'supporter': 'Supporter', 'support': 'Supporter',
    'participant': 'Participant', 'speaker': 'Participant',
}

SUB_PREFIX = re.compile(r'^(?:S|Ce|CM|E)_', re.IGNORECASE)

# Two rows of a recurring series more than this many days apart are separate
# editions. 45 days keeps a week-long conference together while separating the
# April and October runs of the PSinNFDI workshop.
EDITION_GAP_DAYS = 45
MIN_VARIANTS = 2
TITLE_MATCH_CUTOFF = 0.85
EVENT_MATCH_CUTOFF = 0.92

# Review highlighting. Every curated cell that says something the source did
# not is painted by how much judgement went into it, so the reviewer sees at a
# glance which columns are mechanical and which are this script's reading.
MARK_MINOR, MARK_MAJOR, MARK_TITLE, MARK_CHECK = 'minor', 'major', 'title', 'check'
MARK_RANK = {MARK_MINOR: 1, MARK_MAJOR: 2, MARK_TITLE: 3, MARK_CHECK: 4}
MARK_COLOUR = {
    MARK_MINOR: 'FFF2CC',   # light yellow -- spelling
    MARK_MAJOR: 'FFD966',   # dark yellow  -- this script's reading
    MARK_TITLE: 'BDD7EE',   # blue         -- rewritten from the cited page
    MARK_CHECK: 'F4B183',   # orange       -- needs a human
}

# Venue evidence generalised from the neighbouring rows rather than read off a
# page. Weaker than 'web' (a link), 'event name' (the city is in the name) or
# 'institution' (the host has one seat), so it is flagged for a human.
INFERRED_EVIDENCE = {'series'}


def clean(value) -> str:
    if value is None:
        return ''
    text = re.sub(r'\s+', ' ', str(value).replace('\r', ' ')).strip()
    return '' if text.lower() in ('nan', 'none') else text


def as_date(value) -> datetime.date | None:
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    text = clean(value)
    for fmt in ('%Y-%m-%d', '%d-%b-%y', '%d-%b-%Y', '%d/%m/%Y', '%d.%m.%Y'):
        try:
            return datetime.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def row_id(value) -> str:
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, float):
        return f'{value:.10g}'
    return clean(value)


def fold(text: str) -> str:
    ascii_text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode()
    return re.sub(r'\s+', ' ', re.sub(r'[^a-z0-9 ]', ' ', ascii_text.lower())).strip()


def title_key(ident: str, raw_name: str) -> str:
    """A row's identity across rebuilds: its ID plus its untouched name cell.

    Used to line the sheet up with title_matches.json, which is produced from a
    previous build. Both halves come from the source spreadsheet and neither is
    something this script rewrites, so the key survives re-runs.
    """
    return f'{ident}|{fold(raw_name)}'


def split_name(raw: str, ident: str) -> tuple[str, str | None, str | None]:
    """Return (event name, contribution title, curation note)."""
    if ident in ID_OVERRIDES:
        event, title, _ = ID_OVERRIDES[ident]
        return event, title, 'contribution marker is not the word "title"; split by hand'
    text = clean(raw)
    pinned = NAME_OVERRIDES.get(fold(text))
    if pinned:
        return pinned[0], pinned[1], 'booth row; event and title pinned by hand'
    strict = STRICT_MARKER.match(text)
    if strict and strict.group('ev').strip():
        return strict.group('ev').strip(), strict.group('title').strip(), None
    csv_form = CSV_MARKER.match(text)
    if csv_form and csv_form.group('ev').strip():
        return csv_form.group('ev').strip(), csv_form.group('title').strip(), None
    loose = LOOSE_MARKER.match(text)
    if loose and loose.group('ev').strip():
        return (loose.group('ev').strip(), loose.group('title').strip(),
                'non-standard spelling of the "Talk title:" marker')
    return text, None, None


# The same event is written with an en dash in one row and a hyphen in the
# next, with curly quotes in one and straight quotes in another. Left alone,
# each variant becomes its own entry in NOMAD.
DASHES = str.maketrans({'–': '-', '—': '-', '‐': '-', '‑': '-'})
QUOTES = str.maketrans({'“': '"', '”': '"', '„': '"', '‘': "'", '’': "'"})


def tidy(name: str) -> str:
    name = name.translate(DASHES).translate(QUOTES)
    name = re.sub(r'\s+', ' ', name)
    return name.strip().strip('-:;,.').strip()


def canonical_event(name: str, start: datetime.date | None = None) -> tuple[str, str | None]:
    """Apply the editorial renames. Recurring series are split in a later pass."""
    for pattern, replacement in RENAMES:
        if pattern.match(name):
            return replacement, (f'renamed from {name!r}' if replacement != name else None)
    return name, None


def is_recurring(name: str) -> bool:
    folded = fold(name)
    # startswith, not equality: the PSinNFDI workshop carries a subtitle that
    # differs between rows ('... - FAIR Data Principles in NFDI').
    return any(folded.startswith(fold(base)) for base in RECURRING)


def consolidate_spellings(records: list[dict]) -> None:
    """Make one spelling win where names differ only in punctuation or case.

    'CECAM Flagship Workshop on "FAIR and TRUE ..."' and the same name without
    the quotes are one event; left alone they become two entries in NOMAD.
    """
    variants: dict[str, Counter] = defaultdict(Counter)
    for record in records:
        variants[fold(record['Canonical event name'])][record['Canonical event name']] += 1
    for record in records:
        spellings = variants[fold(record['Canonical event name'])]
        if len(spellings) < MIN_VARIANTS:
            continue
        winner = spellings.most_common(1)[0][0]
        if record['Canonical event name'] != winner:
            note(record, f"spelling unified with {winner!r}")
            record['Canonical event name'] = winner
            mark(record, 'Canonical event name', MARK_MINOR)


def split_recurring(records: list[dict]) -> None:
    """Give each edition of a recurring event its own name.

    Appending the year is not enough -- the PSinNFDI workshop ran in April and
    again in October 2024 -- so rows are clustered by date and the label falls
    back to month-and-year when one year holds more than one edition.
    """
    groups: dict[str, list[dict]] = defaultdict(list)
    for record in records:
        if is_recurring(record['Canonical event name']):
            groups[record['Canonical event name']].append(record)

    for name, rows in groups.items():
        dated = sorted((r for r in rows if r['Start date (curated)']),
                       key=lambda r: r['Start date (curated)'])
        if not dated:
            continue
        clusters: list[list[dict]] = [[dated[0]]]
        for row in dated[1:]:
            previous = clusters[-1][-1]['Start date (curated)']
            if (row['Start date (curated)'] - previous).days > EDITION_GAP_DAYS:
                clusters.append([])
            clusters[-1].append(row)
        if len(clusters) < MIN_VARIANTS:
            continue
        years = [c[0]['Start date (curated)'].year for c in clusters]
        by_month = len(set(years)) != len(years)
        for cluster in clusters:
            start = cluster[0]['Start date (curated)']
            label = f'{start:%B %Y}' if by_month else str(start.year)
            for row in cluster:
                row['Canonical event name'] = f'{name} {label}'
                note(row, f'recurring series, edition labelled {label}')
                mark(row, 'Canonical event name', MARK_MAJOR)


def remark(record: dict, text: str) -> None:
    """Add a curation note without raising the 'Needs check' flag."""
    existing = record['Curation note']
    record['Curation note'] = f'{existing}; {text}' if existing else text


def note(record: dict, text: str) -> None:
    remark(record, text)
    record['Needs check'] = 'x'


def mark(record: dict, column: str, level: str, force: bool = False) -> None:
    """Record how much judgement went into one cell. The strongest mark wins.

    ``force`` overrides that, and exists for one case: a title checked against
    the page the sheet itself cites. D1 and D2 are already dark from being
    split out of column D, so without this the reviewer could not tell a title
    that was verified against a source from one this script merely guessed at.
    """
    marks = record.setdefault('_marks', {})
    if force:
        marks[column] = level
        return
    if MARK_RANK[level] > MARK_RANK.get(marks.get(column), 0):
        marks[column] = level


def canonical_fabio(raw: str) -> tuple[str, str | None, str | None]:
    """Return (canonical value, note, mark level)."""
    if not raw:
        return '', None, None
    text = re.sub(r'^\[|\]$', '', raw.strip())
    text = re.sub(r'(?i)^(?:fabio|bibo)\s*:\s*', '', text).strip().lower()
    hit = FABIO_CANONICAL.get(text)
    if hit is None:
        return '', f'unrecognised FaBiO value {raw!r}', MARK_CHECK
    if hit == raw:
        return hit, None, None
    # Same concept spelt differently, or a different concept altogether:
    # 'Poster' is not a FaBiO class and becomes 'conference poster'.
    level = MARK_MINOR if hit.split(': ', 1)[1].lower() == text else MARK_MAJOR
    return hit, f'FaBiO spelling normalised from {raw!r}', level


class Research:
    """The venue / mode / date facts gathered from the web."""

    def __init__(self, path: Path, derived_path: Path | None = None):
        data = json.loads(path.read_text(encoding='utf-8'))
        self.venues = data['venues']
        self.urls = data.get('_sources', {})
        self.online = [fold(x) for x in data['online_events']]
        # Venues established without a lookup -- the location is in the event
        # name, or the host institution has one well-known seat. Matched on the
        # exact name, not a substring, because these are individual judgements.
        self.derived: dict[str, dict] = {}
        if derived_path and derived_path.exists():
            extra = json.loads(derived_path.read_text(encoding='utf-8'))
            self.derived = {fold(k): v for k, v in extra['venues'].items()}
            self.online += [fold(x) for x in extra.get('online', [])]
        self.dates = {}
        for entry in data['date_corrections']:
            if 'UNRESOLVED' in entry['correct'] or 'CORRECT' in entry['correct']:
                continue
            self.dates[fold(entry['event'])] = entry
        self.unresolved = {
            fold(e['event']): e for e in data['date_corrections']
            if 'UNRESOLVED' in e['correct'] or 'within' in e['correct']
        }

    @staticmethod
    def matches_broken_row(fix: dict, start, end) -> bool:
        """Only correct the row the research actually looked at.

        A correction describes one broken row. Applying it to every row of the
        event would overwrite each contribution's own dates with the event span.
        """
        parts = [p.strip() for p in fix['sheet'].split('..')]
        was_start = as_date(parts[0])
        was_end = as_date(parts[-1]) if len(parts) > 1 else None
        if was_start and start != was_start:
            return False
        if was_end and end is not None and end != was_end:
            return False
        return True

    def venue(self, name: str) -> dict | None:
        folded = fold(name)
        best = None
        for entry in self.venues:
            if fold(entry['match']) in folded:
                if best is None or len(entry['match']) > len(best['match']):
                    best = entry
        if best:
            return {**best, 'source_url': self.urls.get(best.get('src', ''), '')}
        hit = self.derived.get(folded)
        if hit:
            return {**hit, 'src': hit.get('evidence', 'derived'),
                    'source_url': hit.get('source_url', '')}
        return None

    def is_online_exact(self, name: str) -> bool:
        return fold(name) in self.online

    def is_online(self, name: str) -> bool:
        folded = fold(name)
        return any(token in folded for token in self.online)


def read_master(path: Path) -> list[dict]:
    workbook = openpyxl.load_workbook(path, data_only=True)
    sheet = workbook[SHEET]
    rows = []
    for index in range(2, sheet.max_row + 1):
        cells = [sheet.cell(index, col).value for col in range(1, 14)]
        if not any(clean(c) for c in cells):
            continue
        if not clean(cells[4]):
            continue  # ID-only scaffolding left behind when rows were moved out
        rows.append({
            'source': f'master row {index}',
            'ID': row_id(cells[1]), 'raw ID': clean(cells[1]),
            'From': cells[2], 'To': cells[3],
            'Name': cells[4], 'Speaker/Participant': clean(cells[5]),
            'Event type (main)': clean(cells[6]), 'Event type (sub)': clean(cells[7]),
            'Event type (old)': clean(cells[8]), 'FAIRmat role': clean(cells[9]),
            'Link to relevant output 1': clean(cells[10]),
            'FaBio Outpu1': clean(cells[11]), 'Educational Purpose': clean(cells[12]),
        })
    return rows


class RowIndex:
    """Recognises a row already present in the master sheet.

    The two files name the same thing differently -- the master writes
    ``<event>\nTalk title: <t>`` and ``EMRS``, the curated CSV writes
    ``<event> talk: <t>`` and ``E-MRS`` -- so raw names cannot be compared.
    A contribution is identified by its **title**, which is distinctive on its
    own; a plain event row by its canonical name. Both allow for the typos the
    sources are full of (``infromation`` for ``information``).
    """

    def __init__(self, rows: list[dict]):
        self.titles: set[str] = set()
        self.events: set[str] = set()
        for row in rows:
            event, title, _ = split_name(clean(row['Name']), row['ID'])
            name, _ = canonical_event(tidy(event), as_date(row['From']))
            if title:
                self.titles.add(fold(title))
            self.events.add(fold(name))
        self._titles = list(self.titles)
        self._events = list(self.events)

    def contains(self, raw: str, start: datetime.date | None) -> bool:
        event, title, _ = split_name(clean(raw), '')
        name, _ = canonical_event(tidy(event), start)
        if title:
            key = fold(title)
            return key in self.titles or bool(
                difflib.get_close_matches(key, self._titles, n=1, cutoff=TITLE_MATCH_CUTOFF))
        key = fold(name)
        return key in self.events or bool(
            difflib.get_close_matches(key, self._events, n=1, cutoff=EVENT_MATCH_CUTOFF))


def read_orphans(path: Path, seen: RowIndex) -> list[dict]:
    """The Users-Meeting contributions that live only in the curated CSV."""
    roles = {'organizer', 'co-organizer', 'participant', 'support'}
    rows = [r for r in list(csv.reader(open(path, encoding='cp1252')))[1:]
            if any(c.strip() for c in r)]
    out = []
    for index, row in enumerate(rows, start=2):
        # 21 rows are shifted one column left: the 'Event type (old)' header was
        # deleted but its data was not. Detect and realign.
        shifted = (row[6] or '').strip().lower() not in roles
        if shifted:
            field = dict(zip(
                ['From', 'To', 'Name', 'Speaker', 'main', 'sub', 'old', 'role',
                 'link', 'fabio', 'edu'], row))
            field.update({'Location': '', 'City': '', 'Country': ''})
        else:
            field = dict(zip(
                ['From', 'To', 'Name', 'Speaker', 'main', 'sub', 'role', 'link',
                 'fabio', 'edu', 'Location', 'City', 'Country'], row))
            field['old'] = ''
        title = re.sub(r'\s+', ' ', field['Name']).strip()
        start = as_date(field['From'])
        if seen.contains(title, start):
            continue
        out.append({
            'source': f'Cordi_events.csv row {index}',
            'ID': 'new', 'From': as_date(field['From']), 'To': as_date(field['To']),
            'Name': title, 'Speaker/Participant': field['Speaker'].strip(),
            'Event type (main)': field['main'].strip(),
            'Event type (sub)': field['sub'].strip(),
            'Event type (old)': field.get('old', '').strip(),
            'FAIRmat role': field['role'].strip(),
            'Link to relevant output 1': field['link'].strip(),
            'FaBio Outpu1': field['fabio'].strip(),
            'Educational Purpose': field['edu'].strip(),
        })
    return out


FLAG_KEYWORDS = (
    'UNRESOLVED', 'BEFORE START', 'unrecognised', 'ambiguous', 'fuzzy',
    'surname-only', 'by hand', 'non-standard', 'renamed', 'date corrected',
)


def resolve_dates(
    row: dict, event_name: str, raw_name: str, research: Research, notes: list[str],
) -> tuple[datetime.date | None, datetime.date | None, str]:
    """Apply the researched corrections, then the single-date rule."""
    start, end = as_date(row['From']), as_date(row['To'])
    date_url = ''
    fix = research.dates.get(fold(event_name)) or research.dates.get(fold(raw_name))
    if fix and research.matches_broken_row(fix, start, end):
        parts = [p.strip() for p in fix['correct'].split('..')]
        new_start = as_date(parts[0])
        new_end = as_date(parts[-1]) if len(parts) > 1 else new_start
        start, end = new_start or start, new_end or end
        date_url = research.urls.get(fix.get('src', ''), '')
        notes.append(f"date corrected ({fix['note']}; source: {date_url or fix['src']})")
    if start and not end:
        end = start
        notes.append('no end date in the source; treated as a single-day event')
    if start and end and end < start:
        notes.append('END DATE IS BEFORE START DATE in the source')
    unresolved = research.unresolved.get(fold(event_name))
    if unresolved:
        notes.append(f"UNRESOLVED: {unresolved['note']}")
    return start, end, date_url


def resolve_contribution_type(ident: str, sub: str, raw_name: str, title: str | None) -> str | None:
    if ident in ID_OVERRIDES:
        return ID_OVERRIDES[ident][2]
    if re.search(r'poster', raw_name, re.I):
        return 'Poster'
    mapped = SUBTYPE_TO_CONTRIBUTION.get(sub)
    if mapped or not title:
        return mapped
    # 'Event type (sub)' can describe the event rather than the talk (every
    # Users-Meeting row says 'CM_Users meeting'); the marker word does not.
    marker = re.search(
        r'\b(talk|poster|tutorial|workshop|symposium|exhibition|presentation|lecture)\b'
        r'\s*(?:title)?\s*:', raw_name, re.I)
    return MARKER_TO_CONTRIBUTION.get(marker.group(1).lower()) if marker else None


def resolve_event_type(event_name: str, sub: str, old: str, main: str) -> str | None:
    """Name keywords first: 'Event type (sub)' describes the contribution."""
    lowered = event_name.lower()
    for keyword, mapped in NAME_KEYWORDS:
        if keyword in lowered:
            return mapped
    return (SUBTYPE_TO_EVENT_TYPE.get(sub) or SUBTYPE_TO_EVENT_TYPE.get(old)
            or MAIN_TO_EVENT_TYPE.get(main.lower()))


def resolve_place(event_name: str, research: Research, notes: list[str]) -> tuple[dict, str]:
    venue = research.venue(event_name)
    mode = (venue or {}).get('mode', '')
    if not mode and (research.is_online_exact(event_name) or research.is_online(event_name)
                     or re.search(r'\b(online|virtual|webinar)\b', event_name, re.I)):
        mode = 'Online'
    if venue:
        where = venue.get('source_url') or venue['src']
        notes.append(f"venue: {where}")
    return venue or {}, mode


def mark_identity(record: dict, row: dict) -> None:
    if not clean(record['ID']):
        mark(record, 'ID', MARK_CHECK)   # blank in the master sheet
    elif not record['Source'].startswith('master'):
        mark(record, 'ID', MARK_MAJOR)   # whole row merged in from Cordi_events.csv
    elif row.get('raw ID', record['ID']) != record['ID']:
        # Excel stores 84.8 as 84.8000000000001; the trailing noise is dropped.
        mark(record, 'ID', MARK_MINOR)


def mark_names(record: dict, split_note: str | None, rename_note: str | None) -> None:
    raw, name = clean(record['Name']), record['Canonical event name']
    title = record['Contribution title']
    if split_note:
        # The marker was not the word 'title' -- split by hand or by a rule
        # that only fits this row.
        mark(record, 'Canonical event name', MARK_CHECK)
        if title:
            mark(record, 'Contribution title', MARK_CHECK)
    elif title:
        mark(record, 'Canonical event name', MARK_MAJOR)
        mark(record, 'Contribution title', MARK_MAJOR)
    elif name != raw:
        mark(record, 'Canonical event name',
             MARK_MINOR if fold(name) == fold(raw) else MARK_MAJOR)
    if rename_note:
        mark(record, 'Canonical event name', MARK_MAJOR)


def mark_dates(record: dict, row: dict, notes: list[str]) -> None:
    start, end = record['Start date (curated)'], record['End date (curated)']
    was_start, was_end = as_date(row['From']), as_date(row['To'])
    if start != was_start:
        mark(record, 'Start date (curated)', MARK_MAJOR)
    if end != was_end:
        single_day = was_end is None and end == start
        mark(record, 'End date (curated)', MARK_MINOR if single_day else MARK_MAJOR)
    joined = ' '.join(notes)
    if 'UNRESOLVED' in joined or 'BEFORE START' in joined:
        mark(record, 'Start date (curated)', MARK_CHECK)
        mark(record, 'End date (curated)', MARK_CHECK)


def mark_people(record: dict, flags: list[str]) -> None:
    if flags:
        # A name matched fuzzily, on the surname alone, or ambiguously.
        mark(record, 'FAIRmat contributors', MARK_CHECK)
        if record['Other contributors']:
            mark(record, 'Other contributors', MARK_CHECK)
        return
    if record['FAIRmat contributors']:
        mark(record, 'FAIRmat contributors', MARK_MINOR)
    if record['Other contributors']:
        # A claim that this person is not on the FAIRmat roster.
        mark(record, 'Other contributors', MARK_MAJOR)


def mark_derived(record: dict, fabio_level: str | None) -> None:
    for column in ('Event type', 'Contribution type', 'Contribution role'):
        if record[column]:
            mark(record, column, MARK_MAJOR)
    if fabio_level:
        mark(record, 'FaBiO (canonical)', fabio_level)


def mark_place(record: dict) -> None:
    """Run after ``relocate``, when the venue columns have settled."""
    if not record['City'] and record['Mode'] != 'Online':
        for column in ('Location name', 'City', 'Country'):
            mark(record, column, MARK_CHECK)
    elif record['Location source']:
        level = MARK_CHECK if record['Location source'] in INFERRED_EVIDENCE else MARK_MAJOR
        for column in ('Location name', 'City', 'Country', 'Mode'):
            if record[column]:
                mark(record, column, level)
    if not record['Mode']:
        mark(record, 'Mode', MARK_CHECK)


def apply_title_matches(records: list[dict], path: Path) -> Counter:
    """Replace event and contribution names with the cited page's own wording.

    match_titles.py decided which of these are safe; this only carries them
    out, so the sheet stays reproducible and the decisions stay reviewable in
    one JSON file rather than being buried in the spreadsheet.
    """
    counts: Counter = Counter()
    if not path.exists():
        return counts
    matches = json.loads(path.read_text(encoding='utf-8'))['matches']

    def put(record: dict, column: str, change: dict) -> None:
        record[column] = change['to']
        level = MARK_TITLE if change['level'] == 'title' else MARK_MINOR
        mark(record, column, level, force=True)
        counts[change['level']] += 1
        if level == MARK_TITLE:
            remark(record, f"title taken from {change['url']}")

    # A contribution title belongs to its row, so it is applied row by row.
    # An event name does not: only the rows that happen to cite a usable page
    # produce a match, and renaming just those would split one event in two --
    # which then splits its venue, its date span and its purpose union. So the
    # event renames are collected first and applied to every row of the event.
    renames: dict[str, dict] = {}
    for record in records:
        found = matches.get(title_key(record['ID'], clean(record['Name'])))
        if not found:
            continue
        change = found.get('Contribution title')
        if change:
            # The match was made against a separate pass over the same inputs.
            # If the cell has moved on since, leave it rather than overwrite.
            if record.get('Contribution title') == change['from']:
                put(record, 'Contribution title', change)
            else:
                counts['stale'] += 1
        change = found.get('Canonical event name')
        if change:
            if record.get('Canonical event name') == change['from']:
                renames.setdefault(change['from'], change)
            else:
                counts['stale'] += 1

    for record in records:
        change = renames.get(record['Canonical event name'])
        if change:
            put(record, 'Canonical event name', change)
    return counts


def normalise_role(record: dict) -> None:
    """Correct column I in place, to the four roles the schema accepts.

    The one grey column the sheet rewrites rather than mirrors. Every cell it
    touches is painted, so nothing is corrected silently.
    """
    raw = clean(record['FAIRmat role'])
    if not raw:
        mark(record, 'FAIRmat role', MARK_CHECK)
        return
    hit = ROLE_CANONICAL.get(raw.lower())
    if hit is None:
        mark(record, 'FAIRmat role', MARK_CHECK)
        remark(record, f'FAIRmat role {raw!r} is not one the schema accepts')
        return
    if hit == raw:
        return
    record['FAIRmat role'] = hit
    # 'support' -> 'Supporter' replaces the word; the rest are only case.
    mark(record, 'FAIRmat role',
         MARK_MINOR if hit.lower() == raw.lower() else MARK_MAJOR)


def assign_event_fields(records: list[dict], path: Path) -> Counter:
    """Fill the two event-level columns: series and URL.

    Both describe the event, so one row establishing a value settles it for
    every row of that event -- otherwise the same event would reach NOMAD with
    the field set on some of its contributions and blank on others.

    Neither is painted when it stays empty. Most events genuinely belong to no
    series and have no page that announces them, so a blank is the right answer
    rather than an omission, and orange is reserved for what a human must fix.
    """
    counts: Counter = Counter()
    urls: dict[str, dict] = {}
    if path.exists():
        urls = json.loads(path.read_text(encoding='utf-8')).get('event_urls', {})

    series: dict[str, str] = {}
    for record in records:
        if record.get('Event series'):
            series.setdefault(record['Canonical event name'], record['Event series'])

    for record in records:
        event = record['Canonical event name']
        found = series.get(event, '')
        record['Event series'] = found
        if found:
            mark(record, 'Event series', MARK_MAJOR)
            counts['series'] += 1
        entry = urls.get(event)
        record['Event URL'] = entry['url'] if entry else ''
        if entry:
            mark(record, 'Event URL', MARK_MAJOR)
            counts['url'] += 1
    counts['events with a series'] = len(series)
    counts['events with a URL'] = len(urls)
    return counts


def note_suggestions(records: list[dict], path: Path) -> int:
    """Record the title suggestions that were too weak to apply.

    They go in the curation note, on the row they belong to, so a reviewer sees
    the alternative wording beside the cell instead of having to open a JSON
    file. Nothing in D1 or D2 is changed.
    """
    if not path.exists():
        return 0
    suggestions = json.loads(path.read_text(encoding='utf-8')).get('suggestions', [])
    by_key: dict[str, list[dict]] = {}
    for item in suggestions:
        by_key.setdefault(item['key'], []).append(item)
    added = 0
    for record in records:
        for item in by_key.get(title_key(record['ID'], clean(record['Name'])), []):
            which = 'event' if item['column'].startswith('Canonical') else 'contribution'
            remark(record, f"source page calls this {which} {item['to']!r}")
            added += 1
    return added


def note_missing_titles(records: list[dict]) -> int:
    """Say so on the rows that will not become a contribution.

    A contribution needs a title, and these rows have none, so the archive
    generator skips them. That is a decision worth seeing in the sheet rather
    than discovering later in NOMAD.
    """
    count = 0
    for record in records:
        if record['Contribution title']:
            continue
        if not any(record[column] for column in
                   ('Contribution type', 'FaBiO (canonical)',
                    'FAIRmat contributors', 'Other contributors')):
            continue
        remark(record, 'no contribution title: this row adds no contribution, '
                       'only its event')
        count += 1
    return count


def resolve_purposes(record: dict) -> tuple[list[str], bool]:
    """Return this row's purposes, and whether they rest on the category alone."""
    main = clean(record['Event type (main)']).lower()
    contribution = clean(record['Contribution type'])
    found: set[str] = set()

    if main == MAIN_EDUCATIONAL:
        found.add(PURPOSE_EDUCATIONAL)
    if contribution in TEACHING_CONTRIBUTIONS and main in TEACHING_CATEGORIES:
        found.add(PURPOSE_EDUCATIONAL)
    if contribution in SCHOLARLY_CONTRIBUTIONS and main in SCHOLARLY_CATEGORIES:
        found.add(PURPOSE_SCHOLARLY)
    if contribution in BOOTH_CONTRIBUTIONS and main in BOOTH_CATEGORIES:
        found.add(PURPOSE_INFORMATIONAL)
    # The 'Educational Purpose' flag is the curator's own word and always adds
    # the educational purpose, whatever the other two columns say.
    if clean(record['Educational Purpose']).lower() in TRUTHY:
        found.add(PURPOSE_EDUCATIONAL)

    inferred = False
    if not found and not contribution:
        fallback = MAIN_ONLY_PURPOSE.get(main)
        if fallback:
            found.add(fallback)
            inferred = True
    return [p for p in PURPOSE_ORDER if p in found], inferred


def assign_purposes(records: list[dict]) -> None:
    """Fill both purpose columns.

    ``event_purpose`` in the schema belongs to the event, but the rules read
    the contribution type, which belongs to the row. So each row is resolved on
    its own, and the event-level column is the union over the event -- which is
    where an event with two or three purposes comes from.
    """
    per_event: dict[str, set[str]] = defaultdict(set)
    weak: set[str] = set()
    for record in records:
        purposes, inferred = resolve_purposes(record)
        record['Event purpose (row)'] = '; '.join(purposes)
        record['_purpose_inferred'] = inferred
        event = record['Canonical event name']
        per_event[event].update(purposes)
        if inferred:
            weak.add(event)
        # A remark, not a flag: 'Needs check' keeps its existing meaning, and
        # the orange cell already tells the reviewer where to look.
        if inferred:
            remark(record, 'event purpose inferred from the event category alone')
        elif not purposes:
            remark(record, 'event purpose undecidable from this row')

    for record in records:
        event = record['Canonical event name']
        union = per_event[event]
        record['Event purpose'] = '; '.join(p for p in PURPOSE_ORDER if p in union)
        for column, empty, shaky in (
            ('Event purpose (row)', not record['Event purpose (row)'],
             record['_purpose_inferred']),
            ('Event purpose', not union, event in weak),
        ):
            mark(record, column, MARK_CHECK if (empty or shaky) else MARK_MAJOR)


def curate_row(row: dict, research: Research, roster: Roster) -> dict:
    """Every original column, plus this script's reading of it."""
    notes: list[str] = []
    ident = row['ID']
    raw_name = clean(row['Name'])

    event_name, title, split_note = split_name(raw_name, ident)
    if split_note:
        notes.append(split_note)
    event_name, rename_note = canonical_event(tidy(event_name))
    if rename_note:
        notes.append(rename_note)

    start, end, date_url = resolve_dates(row, event_name, raw_name, research, notes)

    sub = SUB_PREFIX.sub('', row['Event type (sub)']).strip().lower()
    old = row['Event type (old)'].strip().lower()
    place, mode = resolve_place(event_name, research, notes)
    internal, external, flags = roster.classify(row['Speaker/Participant'])
    notes.extend(flags)
    fabio, fabio_note, fabio_level = canonical_fabio(row['FaBio Outpu1'])
    if fabio_note and 'unrecognised' in fabio_note:
        notes.append(fabio_note)

    joined = ' '.join(notes)
    urls: list[str] = []
    for candidate in (place.get('source_url', ''), date_url,
                      row.get('Link to relevant output 1', '')):
        if candidate and candidate.startswith('http') and candidate not in urls:
            urls.append(candidate)
    if flags:
        roster_url = 'https://fair-di.eu/fairmat/about-fairmat/team-fairmat'
        if roster_url not in urls:
            urls.append(roster_url)
    record = {
        **{k: row.get(k, '') for k in ORIGINAL_COLUMNS},
        'Canonical event name': event_name,
        'Contribution title': title or '',
        'Contribution type': resolve_contribution_type(ident, sub, raw_name, title) or '',
        'Contribution role': CONTRIBUTION_ROLE_OVERRIDES.get(ident, ''),
        'Event type': resolve_event_type(event_name, sub, old, row['Event type (main)']) or '',
        'Start date (curated)': start,
        'End date (curated)': end,
        'Location name': place.get('location_name', ''),
        'City': place.get('city', ''),
        'Country': place.get('country', ''),
        'Mode': mode,
        'Location source': place.get('src', ''),
        'FAIRmat contributors': '; '.join(internal),
        'Other contributors': '; '.join(external),
        'FaBiO (canonical)': fabio,
        # Column G names the series where there is one; settled per event in
        # assign_event_fields, and the URL is filled there too.
        'Event series': SUBTYPE_TO_SERIES.get(sub, ''),
        'Event URL': '',
        'Source': row['source'],
        'Curation note': '; '.join(notes),
        'Needs check': 'x' if any(k in joined for k in FLAG_KEYWORDS) else '',
        'Source URLs': '\n'.join(urls),
    }
    normalise_role(record)
    mark_identity(record, row)
    mark_names(record, split_note, rename_note)
    mark_dates(record, row, notes)
    mark_people(record, flags)
    mark_derived(record, fabio_level)
    # The venue columns are marked after relocate(), once the canonical name
    # has settled and the second lookup has had its go.
    return record


def relocate(records: list[dict], research: Research) -> None:
    """Re-resolve the venue after the canonical name has settled.

    ``consolidate_spellings`` and ``split_recurring`` rename events, and a
    venue keyed on the final name ("APS March Meeting 2020") cannot match the
    name the first pass looked up ("APS March Meeting"). Anything that only
    resolves now is filled in here.
    """
    for record in records:
        # Already placed by the first pass -- including online events, which
        # legitimately have no city but do have a Location source.
        if record['City'] or record['Location name'] or record['Location source']:
            continue
        notes: list[str] = []
        place, mode = resolve_place(record['Canonical event name'], research, notes)
        if not place and not mode:
            continue
        record['Location name'] = place.get('location_name', '') or record['Location name']
        record['City'] = place.get('city', '') or record['City']
        record['Country'] = place.get('country', '') or record['Country']
        record['Mode'] = mode or record['Mode']
        record['Location source'] = place.get('src', '') or record['Location source']
        url = place.get('source_url', '')
        if url and url not in record['Source URLs']:
            record['Source URLs'] = f"{record['Source URLs']}\n{url}".strip()
        for text in notes:
            existing = record['Curation note']
            record['Curation note'] = f'{existing}; {text}' if existing else text


def curate_all(source: Path) -> tuple[list[dict], int, int]:
    """Every row curated, up to but NOT including the title matches.

    Split out so match_titles.py can compare against the names this script
    produces on its own. If it read the finished spreadsheet instead, it would
    be reading back its own previous corrections and find nothing left to fix.
    """
    research = Research(RESEARCH, DERIVED)
    roster = Roster()
    master = read_master(source)
    orphans = read_orphans(ORPHAN_CSV, RowIndex(master))

    records = [curate_row(row, research, roster) for row in master + orphans]
    consolidate_spellings(records)
    split_recurring(records)
    relocate(records, research)
    for record in records:
        mark_place(record)
    return records, len(master), len(orphans)


def build(source: Path, out_path: Path) -> dict:
    records, n_master, n_orphans = curate_all(source)
    # After relocate, so the venue lookups still key on the name they were
    # gathered under; before assign_purposes, so the union groups final names.
    # Before the renames: match_titles.py keyed its event URLs on the names it
    # saw, which are the names this function has produced so far. Renaming
    # first would leave those keys unfindable. The renames are consistent
    # across every row of an event, so the grouping is the same either way.
    fields = assign_event_fields(records, TITLE_MATCHES)
    titles = apply_title_matches(records, TITLE_MATCHES)
    assign_purposes(records)
    suggested = note_suggestions(records, TITLE_MATCHES)
    titleless = note_missing_titles(records)
    records.sort(key=lambda r: (r['Start date (curated)'] or datetime.date(1900, 1, 1),
                                r['Canonical event name']))

    flagged = sum(1 for r in records if r['Needs check'] == 'x')
    painted = Counter(level for r in records for level in r.get('_marks', {}).values())
    painted[MARK_CHECK] += 2 * flagged  # the 'Needs check' / 'Curation note' pair
    stats = {
        'master': n_master, 'orphans': n_orphans,
        'flagged': flagged,
        'rows': len(records),
        'events': len({r['Canonical event name'] for r in records}),
        'minor': painted[MARK_MINOR], 'major': painted[MARK_MAJOR],
        'title': painted[MARK_TITLE], 'check': painted[MARK_CHECK],
        'titles_applied': titles['title'], 'titles_respelt': titles['minor'],
        'titles_stale': titles['stale'],
        'series_events': fields['events with a series'],
        'url_events': fields['events with a URL'],
        'suggestions_noted': suggested, 'rows_without_a_title': titleless,
    }
    write(records, out_path, stats)
    return stats


def write(records: list[dict], path: Path, stats: dict) -> None:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = 'Events'
    columns = COLUMNS

    header_original = PatternFill('solid', fgColor='D9D9D9')
    header_curated = PatternFill('solid', fgColor='C6E0B4')
    fills = {level: PatternFill('solid', fgColor=colour)
             for level, colour in MARK_COLOUR.items()}

    # Columns holding several lines or long prose are wrapped so the reviewer
    # can read them in place; everything else stays on one line to keep rows
    # short and the sheet scannable.
    wrapped = {'Source URLs', 'Curation note', 'Event URL'}

    for index, name in enumerate(columns, start=1):
        cell = sheet.cell(1, index, HEADERS[name])
        cell.font = Font(bold=True)
        cell.fill = header_original if KINDS[name] == 'ORIGINAL' else header_curated
        cell.alignment = Alignment(vertical='center', wrap_text=True)
    sheet.row_dimensions[1].height = 30

    for row_index, record in enumerate(records, start=2):
        marks = dict(record.get('_marks', {}))
        if record['Needs check'] == 'x':
            marks['Needs check'] = marks['Curation note'] = MARK_CHECK
        for col_index, name in enumerate(columns, start=1):
            value = record.get(name, '')
            cell = sheet.cell(row_index, col_index, value)
            if isinstance(value, datetime.date):
                cell.number_format = 'YYYY-MM-DD'
            cell.alignment = Alignment(vertical='top', wrap_text=name in wrapped)
            if name in marks:
                cell.fill = fills[marks[name]]

    # Width: never narrower than the header, and wide enough for ordinary
    # content. Long free text is capped so one column cannot fill the screen.
    explicit = {
        'Name': 58, 'Canonical event name': 44, 'Contribution title': 46,
        'Speaker/Participant': 32, 'FAIRmat contributors': 34,
        'Other contributors': 28, 'Link to relevant output 1': 30,
        'Location name': 40, 'Curation note': 62, 'Source URLs': 54,
        'Event type (main)': 30, 'Event type (sub)': 30, 'Event type (old)': 26,
        'FaBio Outpu1': 26, 'FaBiO (canonical)': 26, 'Source': 24,
        'From': 12, 'To': 12, 'Start date (curated)': 20, 'End date (curated)': 20,
        'ID': 10, 'City': 18, 'Country': 16, 'Mode': 10, 'Needs check': 12,
        'Event series': 34, 'Event URL': 46, 'FAIRmat role': 20,
        'Event purpose (row)': 30, 'Event purpose': 34,
    }
    for index, name in enumerate(columns, start=1):
        width = explicit.get(name, 16)
        header = HEADERS[name]
        sheet.column_dimensions[get_column_letter(index)].width = max(width, len(header) + 2)

    # Freeze the header and the ID column: with 31 columns the reviewer scrolls
    # sideways constantly and needs to know which row they are on.
    sheet.freeze_panes = 'B2'
    sheet.auto_filter.ref = f'A1:{get_column_letter(len(columns))}{len(records) + 1}'

    notes = workbook.create_sheet('README')
    notes.cell(1, 1, 'COLOUR KEY -- what the highlighting in the Events sheet means').font = (
        Font(bold=True))
    legend = (
        (MARK_MINOR, f"  spelling, punctuation, or a value carried straight across "
                     f"({stats['minor']} cells)"),
        (MARK_MAJOR, f"  a real change: extracted, mapped, or looked up on the web "
                     f"({stats['major']} cells)"),
        (MARK_TITLE, f"  rewritten to match the title on the page cited in the last "
                     f"column ({stats['title']} cells)"),
        (MARK_CHECK, f"  NEEDS A HUMAN: uncertain, inferred without a source, or still "
                     f"blank ({stats['check']} cells)"),
    )
    for offset, (level, text) in enumerate(legend, start=3):
        notes.cell(offset, 1).fill = fills[level]
        notes.cell(offset, 2, text)
    index = len(legend) + 4
    for line in README.strip().splitlines():
        notes.cell(index, 2, line).alignment = Alignment(vertical='top')
        index += 1
    notes.column_dimensions['A'].width = 6
    notes.column_dimensions['B'].width = 102
    workbook.save(path)


README = """
CURATED FAIRmat EVENTS SHEET
============================
Generated by scripts/build_curated_sheet.py from
  'Events database - (September 2024)_full_20260904.xlsx'   (the master)
  'Cordi_events.csv'                                        (63 rows found only there)
  'local/event_venues_researched.json' + 'event_venues_derived.json'  (venue research)
  'local/fairmat_roster.json'                               (229 FAIRmat people)
Do not hand-edit and regenerate: edit the master sheet or the inputs instead.

HOW THE COLUMNS ARE LAID OUT
  Every header starts with a code that says what it pairs with.

    A .. L   the twelve columns of the master spreadsheet, in their original
             order. Grey header. Carried through UNCHANGED.
    A1, B1,  a curated column derived from the source column with the same
    D1, D2   letter. Green header. D1 and D2 both come from D (Name).
    +        a curated column with no counterpart in any source file.

  So you read a pair straight across:

    B  | From                     ->  B1 | From (curated)
    C  | To                       ->  C1 | To (curated)
    D  | Name                     ->  D1 | Name - event
                                      D2 | Name - contribution
    E  | Speaker/Participant      ->  E1 | Speaker - FAIRmat
                                      E2 | Speaker - external
    F  | Event type (main)        ->  F1 | Event type (curated)
    G  | Event type (sub)         ->  G1 | Contribution type
                                      G2 | Contribution role
    K  | FaBio Outpu1             ->  K1 | FaBio Outpu1 (curated)

  A (ID), H (Event type (old)), J (Link to relevant output 1) and L
  (Educational Purpose) are used as they stand and have no curated twin.

  I (FAIRmat role) is the ONE grey column corrected in place. The master spells
  the same four roles several ways -- 'participant', 'organizer' and
  'Organizer', 'co-organizer', 'support' -- and the schema accepts only
  Organizer / Co-organizer / Supporter / Participant. 513 cells were recased or
  reworded and every one of them is painted, so nothing changed silently.

  The '+' columns are new information, not a re-reading of anything:
    + Event purpose (row), + Event purpose (curated)
    + Event series   - read off column G, which names it outright
    + Event URL      - the page found to announce the event
    + Location, + City, + Country   - named as in Ahmed's Cordi_events.csv
    + Mode, + Location source
    + Row source, + Curation note, + Needs check, + Source URLs

THE TWO EVENT PURPOSE COLUMNS
  Purpose is read off F (event category), G1 (contribution type) and L
  (Educational Purpose) together, by the rules the curation owners set:

    F = Educational Events                          -> Educational event
    contribution is a tutorial or workshop session  -> Educational event
      (within Conferences and Meetings, Community engagement or Educational)
    contribution is a talk, poster, panel discussion, symposium or conference
      session, within Scholarly talks or Conferences -> Scholarly event
    contribution is an information booth, within Public engagement, Community
      engagement or Conferences                      -> Informational event
    L = Yes                                          -> always adds Educational

  '(row)' is what that one row implies. '(curated)' is the union across every
  row of the same event, and is the one to load into the schema, because
  event_purpose belongs to the event -- which is how an event comes to have two
  or three purposes.

  Where a row names no contribution type, the purpose falls back to the event
  category alone. That is a weaker reading, so those cells are orange and the
  curation note says so.

WHERE THE HIGHLIGHTING COMES FROM
  Grey (original) columns are not painted -- they are the master sheet as it
  stands. The colour sits on the green (curated) columns and says how much
  judgement went into that one cell. The single exception is 'A | ID', which is
  painted where the row itself was added or where the master has no ID at all.

    light yellow   Canonical name that differs from the master only in
                   punctuation, spacing or case; a name respelt to match the
                   page cited in 'Source URLs'; an end date copied from the
                   start date because the source gave one date; a FaBiO value
                   respelt; speakers that matched the roster exactly.
    BLUE           D1 or D2 rewritten to the title the cited page actually
                   uses -- not just a respelling, so read it. The page is
                   named in the curation note and in 'Source URLs'. Only
                   applied where the page title said everything the sheet said
                   and more, so nothing was shortened away.
    dark yellow    A FAIRmat role reworded to what the schema accepts; an
                   event series read off column G; an event URL whose page was
                   confirmed to announce that event; a contribution title
                   split off the Name column; an event
                   type or contribution type inferred; a date corrected from a
                   source; a venue read off a page, off the event name, or off
                   the host institution; a speaker judged to be an outside
                   guest; a row merged in from Cordi_events.csv (ID = new).
    ORANGE         Look at this. A name matched only on the surname or only
                   fuzzily; a contribution marker that had to be split by hand;
                   a date that is unresolved or ends before it starts; a venue
                   generalised from the rest of its series; a FaBiO value that
                   is not a FaBiO class; a blank the research could not fill
                   (missing city on an onsite event, missing mode, missing ID).

HOW TO REVIEW
  Row 1 is frozen and column A is frozen, so the ID stays visible while you
  scroll sideways. Every column has an autofilter.
  * The fastest pass is: read the orange cells only. Everything else is either
    mechanical or has a source behind it in 'Source URLs'.
  * Filter 'Needs check' = x to see only the rows where curation made a
    judgement. 'Curation note' says what was done; 'Source URLs' gives the
    page it came from.
  * 'Location source' says HOW a venue was established:
      web          - looked up; the proving page is in Source URLs
      fairmat_past - the FAIRmat past-events page
      event name   - the city is written in the Name column itself
      institution  - the host has one well-known seat (HZB -> Berlin, ...)
      series       - the whole series shares a venue or ran online
      colleague    - from Ahmed's original Cordi_events.csv curation

THE EVENT-LEVEL COLUMNS
  Event series, Event URL and Event purpose (curated) describe the EVENT, so
  every row of the same event carries the same value. Series and URL are left
  blank, not orange, where there is none: most events belong to no series and
  have no page announcing them, so blank is the right answer rather than an
  omission, and orange stays reserved for what a human must fix.

  An event URL is only taken from a page that actually announces the event --
  a conference site or an Indico event page whose title matches the event name.
  Links to a recording, a DOI or a deposit are outputs, not announcements, and
  are never used. 163 of 418 events have one.

WHAT WAS DONE
  * 63 Users-Meeting contributions that existed only in Cordi_events.csv are
    merged in with ID = new (see the 'Source' column).
  * Recurring events whose name carries no year are split per edition by date
    cluster, so the 2022 and 2023 QUANTSOL schools stop merging into one entry.
    A year that holds two editions is labelled by month (PSinNFDI April 2024 /
    October 2024).
  * Dates corrected only where research established a correction, and only on
    the row that was actually wrong; the original From/To are untouched.
  * Venue, city, country and mode filled for 409 of 418 events.
  * Speakers split into FAIRmat people and outside guests against the roster.
  * FaBiO values spelt exactly as the ontology spells them, checked against
    the current release of http://purl.org/spar/fabio (254 classes). FaBiO
    labels are lowercase. Three spellings the master used are not FaBiO
    classes: 'Poster' -> 'fabio: conference poster', 'Software dataset' ->
    'fabio: dataset', and 'Interview' -> 'bibo: Interview', FaBiO having no
    class for it. The plugin schema now carries the same list.
  * Every URL the sheet cites was downloaded and its titles compared with D1
    and D2, so the names match what the programme or event page calls them.
    Page titles are only accepted when they say everything the sheet said and
    more; a page that shortens the sheet's own wording is ignored, and page
    furniture (session codes, speaker names, site names) is stripped first.

  * Rows with no contribution title add no contribution to their event, only
    the event itself. Their curation note says so. The contribution type, FaBiO
    term, speaker and output link on those rows are therefore not carried into
    the schema -- a deliberate decision, so that every contribution title in
    NOMAD is a real title.
  * Title wordings that were close but too uncertain to apply are written into
    the curation note as "source page calls this ...", for a human to accept or
    ignore. D1 and D2 are untouched by those.

KNOWN OPEN POINTS
  * 10 events have no location of any kind and are not online: Applied AI in
    materials science workshop, The 2nd MGE Workshop, 'Maschine Learning',
    CyberWeek summer school 2023, Annual Digital Catalysis Conference 2023,
    TRR277 PI meeting, Networking Data Literacy VDI/VDE, both SolMates project
    meetings, NFDI Workshop. None of them cites a page anywhere in the sheet
    and their names are too generic to identify, so they are left blank rather
    than guessed. 8 further events have no Mode.
  * contributions_overview is generated from the contributions themselves, so
    it restates the entry rather than adding to it; regenerate it freely.
  * CoRDI 2023 S. Auer talk (ID 218): dated 2023-09-20 in the source, but the
    conference ran 12-14 Sep 2023.
  * QUANTSOL 2022: HZB gives 4-11 Sep 2022, the master sheet says 3-10 Sep.
  * ID 384.4 is almost certainly a typo for 381.4. The ID is left as written.
  * The unnamed column A of the master sheet was an internal curation marker
    with no remembered meaning (Ahmed Mansour, 2026-09-04) and was dropped.
  * 'bibo: Interview' is a proposal. FaBiO has no class for an interview, and
    BIBO's is the only published one; the alternative would be 'fabio: movie',
    which names the medium and loses the meaning. No row uses it today.
  * event_interaction_approach is deliberately not filled. Column F states two
    of its five values, but the other three appear nowhere in the data, so the
    field is left to a human rather than half-guessed.
  * 3 rows have neither an event category nor a contribution type, so they get
    no event purpose at all.
"""


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument('source', nargs='?', default=str(DEFAULT_SOURCE))
    parser.add_argument('-o', '--output',
                        default=str(LOCAL / f'Events database - CURATED_{datetime.date.today():%Y%m%d}.xlsx'))
    args = parser.parse_args()

    source = Path(args.source)
    if not source.exists():
        print(f'ERROR input not found: {source}', file=sys.stderr)
        return 1

    stats = build(source, Path(args.output))
    print(
        f"{stats['rows']} rows written to {args.output}\n"
        f"  {stats['master']} from the master sheet + {stats['orphans']} merged from Cordi_events.csv\n"
        f"  {stats['events']} distinct events\n"
        f"  {stats['flagged']} rows flagged for review"
    )
    return 0


if __name__ == '__main__':
    sys.exit(main())
