#!/usr/bin/env python3
"""Turn the curated events sheet into ``*.archive.json`` files for NOMAD.

This is the second half of the pipeline. ``build_curated_sheet.py`` makes every
judgement and writes it into a reviewable column; this script only reads those
columns and assembles the archives. No parsing, no guessing, no regexes over
event names -- if an entry is wrong, the sheet is wrong, and the sheet is the
thing a human can fix.

Rows are grouped by ``Canonical event name``: one event becomes one entry with
one ``Contribution`` sub-section per contributing row.

Usage::

    uv run python packages/fairmat-project-outputs/scripts/generate_archives_from_curated.py
"""

from __future__ import annotations

import argparse
import datetime
import json
import re
import sys
import unicodedata
import warnings
from collections import defaultdict
from pathlib import Path

warnings.filterwarnings('ignore', module='openpyxl')

import openpyxl  # noqa: E402
from build_curated_sheet import HEADERS  # noqa: E402
from nomad.datamodel import EntryArchive, EntryMetadata  # noqa: E402

from fairmat_project_outputs.schema_packages.schema_package import (  # noqa: E402
    EVENT_PURPOSE,
    FAIRMAT_ROLE,
    Contribution,
    FAIRmatEvent,
)

ROOT = Path(__file__).resolve().parents[1]
LOCAL = ROOT / 'local'
FILENAME_PREFIX = 'event'
SHEET = 'Events'

# The sheet and the schema now both spell these the way the ontologies do, so
# the mapping is an identity and exists only to reject anything that is not a
# published class. See FABIO_TERMS in the schema package.
FABIO_TO_SCHEMA = {
    term: term
    for term in (
        'fabio: abstract',
        'fabio: announcement',
        'fabio: book chapter',
        'fabio: computer program',
        'fabio: conference poster',
        'fabio: conference proceedings',
        'fabio: dataset',
        'fabio: entity metadata',
        'fabio: instructional work',
        'fabio: journal article',
        'fabio: meeting report',
        'fabio: periodical issue',
        'fabio: preprint',
        'fabio: presentation',
        'fabio: scholarly work',
        'fabio: timetable',
        'fabio: web page',
    )
}
FABIO_TO_SCHEMA['bibo: interview'] = 'bibo: Interview'


def text(value) -> str:
    return '' if value is None else re.sub(r'\s+', ' ', str(value)).strip()


def as_datetime(value):
    if isinstance(value, datetime.datetime):
        return value
    if isinstance(value, datetime.date):
        return datetime.datetime(value.year, value.month, value.day)
    return None


def people(value) -> list[str] | None:
    names = [n.strip() for n in text(value).split(';') if n.strip()]
    return names or None


def slugify(value: str) -> str:
    ascii_value = unicodedata.normalize('NFKD', value).encode('ascii', 'ignore').decode()
    return re.sub(r'_+', '_', re.sub(r'[^\w]+', '_', ascii_value)).strip('_')[:70]


def read_sheet(path: Path) -> list[dict]:
    """Read the curated sheet back into its internal column keys.

    The sheet shows review-friendly headers ('D1 | Name - event'); this maps
    them back to the keys the rest of the script uses.
    """
    workbook = openpyxl.load_workbook(path, data_only=True)
    sheet = workbook[SHEET]
    by_header = {header: key for key, header in HEADERS.items()}
    header = [by_header.get(c.value, c.value) for c in sheet[1]]
    rows = []
    for cells in sheet.iter_rows(min_row=2):
        row = dict(zip(header, [c.value for c in cells]))
        if text(row.get('Canonical event name')):
            rows.append(row)
    return rows


# Notes written for whoever reviews the spreadsheet, not caveats about the
# data. A suggested alternative wording is something to weigh with the sheet
# and the source page side by side; in a NOMAD entry comment it is noise, and
# it was pushing some comments past 1900 characters.
SHEET_ONLY_NOTE = re.compile(
    r"\s*;?\s*source page calls this (?:event|contribution) '[^']*'", re.I
)


def sheet_only_note(note: str) -> str:
    """The curation note with the spreadsheet-only remarks taken out."""
    return SHEET_ONLY_NOTE.sub('', note).strip(' ;')


def build_event(name: str, rows: list[dict], problems: list[str]) -> tuple[FAIRmatEvent, list[str]]:
    starts = [as_datetime(r['Start date (curated)']) for r in rows]
    ends = [as_datetime(r['End date (curated)']) for r in rows]
    starts = [d for d in starts if d]
    ends = [d for d in ends if d]

    event = FAIRmatEvent(
        event_name=name,
        start_date=min(starts) if starts else None,
        end_date=max(ends) if ends else None,
    )

    def first(column: str) -> str:
        for row in rows:
            value = text(row.get(column))
            if value:
                return value
        return ''

    event.event_type = first('Event type') or None
    event.location_name = first('Location name') or None
    event.city = first('City') or None
    event.country = first('Country') or None
    event.mode = first('Mode') or None

    # Series, role, purpose and URL are all decided in the sheet now, where a
    # human can see and correct them, instead of being re-derived here where
    # nobody could. This script only carries across what the sheet says.
    event.event_series = first('Event series') or None
    event.event_url = first('Event URL') or None

    role = first('FAIRmat role')
    if role:
        if role not in FAIRMAT_ROLE:
            problems.append(f'{name}: FAIRmat role {role!r} is not in the schema')
            role = ''
        event.fairmat_role = role or None

    purpose = [p for p in text(first('Event purpose')).split('; ') if p]
    unknown = [p for p in purpose if p not in EVENT_PURPOSE]
    if unknown:
        problems.append(f'{name}: event purpose {unknown} is not in the schema')
    event.event_purpose = [p for p in purpose if p in EVENT_PURPOSE] or None

    contributions = []
    for row in rows:
        title = text(row.get('Contribution title'))
        if not title:
            continue
        fabio = FABIO_TO_SCHEMA.get(text(row.get('FaBiO (canonical)')).lower())
        raw_fabio = text(row.get('FaBiO (canonical)'))
        if raw_fabio and not fabio:
            problems.append(f'{name}: FaBiO value {raw_fabio!r} has no schema equivalent')
        contributions.append(Contribution(
            title=title,
            start_date=as_datetime(row['Start date (curated)']),
            end_date=as_datetime(row['End date (curated)']),
            contribution_type=text(row.get('Contribution type')) or None,
            contribution_role=text(row.get('Contribution role')) or None,
            fairmat_contributors=people(row.get('FAIRmat contributors')),
            other_contributors=people(row.get('Other contributors')),
            fabio_type=fabio,
            fabio_url=text(row.get('Link to relevant output 1')) or None,
        ))
    event.contributions = contributions
    # contributions_overview is deliberately NOT set here: FAIRmatEvent.normalize()
    # rebuilds it as an HTML table from the contributions on every processing
    # run, so anything written here is overwritten the moment NOMAD ingests it.

    # Only meaningful without contributions: normalize() rebuilds this from the
    # contributions otherwise.
    if not contributions:
        names = people(first('FAIRmat contributors'))
        if names:
            event.fairmat_contributors = names

    reviews = []
    for row in rows:
        if not text(row.get('Needs check')):
            continue
        note = sheet_only_note(text(row.get('Curation note')))
        if note:
            reviews.append(f"row {text(row.get('ID')) or '?'}: {note}")
    return event, reviews


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    default_input = sorted(LOCAL.glob('Events database - CURATED_*.xlsx'))
    parser.add_argument(
        'sheet', nargs='?',
        default=str(default_input[-1]) if default_input else '',
        help='curated sheet (default: the newest local/Events database - CURATED_*.xlsx)',
    )
    parser.add_argument(
        '-o', '--output',
        default=str(LOCAL / f'archives_{datetime.date.today():%Y%m%d}'),
    )
    args = parser.parse_args()

    source = Path(args.sheet)
    if not source.exists():
        print(f'ERROR curated sheet not found: {source}', file=sys.stderr)
        return 1
    out_dir = Path(args.output)
    out_dir.mkdir(parents=True, exist_ok=True)

    rows = read_sheet(source)
    grouped: dict[str, list[dict]] = defaultdict(list)
    for row in rows:
        grouped[text(row['Canonical event name'])].append(row)

    problems: list[str] = []
    flagged: list[tuple[str, list[str]]] = []
    created = errors = 0
    used: set[str] = set()

    for name, event_rows in grouped.items():
        try:
            event, reviews = build_event(name, event_rows, problems)
        except Exception as exc:  # noqa: BLE001 - report and carry on
            errors += 1
            print(f'ERROR {name!r}: {exc}', file=sys.stderr)
            continue

        metadata = EntryMetadata(entry_name=name)
        if reviews:
            metadata.comment = 'NEEDS REVIEW - ' + '; '.join(reviews)
            flagged.append((name, reviews))

        stem = slugify(name) or 'event'
        filename = f'{FILENAME_PREFIX}_{stem}'
        if filename in used:
            filename = f'{filename}_{len(used)}'
        used.add(filename)

        archive = EntryArchive(data=event, metadata=metadata)
        with open(out_dir / f'{filename}.archive.json', 'w', encoding='utf-8') as handle:
            json.dump(archive.m_to_dict(), handle, indent=2, ensure_ascii=False)
        created += 1

    report = out_dir / 'review_report.md'
    with open(report, 'w', encoding='utf-8') as handle:
        handle.write(f'# Review report\n\nSource: `{source.name}`\n\n')
        handle.write(
            f'- {created} entries written from {len(rows)} rows\n'
            f'- {errors} events failed\n'
            f'- {len(flagged)} entries carry a NEEDS REVIEW comment\n\n'
        )
        handle.write('## Entries flagged for review\n\n')
        for name, reasons in sorted(flagged):
            handle.write(f'### {name}\n\n')
            for reason in reasons:
                handle.write(f'- {reason}\n')
            handle.write('\n')
        handle.write('## Mapping problems\n\n')
        handle.write(''.join(f'- {p}\n' for p in problems) if problems else 'None.\n')

    for problem in problems[:20]:
        print(f'WARNING {problem}', file=sys.stderr)
    print(
        f'{created} archive files written to {out_dir}\n'
        f'  {len(rows)} rows -> {created} events, {errors} failed\n'
        f'  {len(flagged)} flagged for review, {len(problems)} mapping problems\n'
        f'  review report: {report}'
    )
    return 1 if errors else 0


if __name__ == '__main__':
    sys.exit(main())
