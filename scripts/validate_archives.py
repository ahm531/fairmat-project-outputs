#!/usr/bin/env python3
"""Check generated *.archive.json files before they go anywhere near an Oasis.

Writing a file NOMAD rejects costs a whole upload cycle to find out, so every
check that can be made locally is made here:

  1. the file parses as JSON and carries the right ``m_def``
  2. NOMAD itself can rebuild the section from it (``m_from_dict``), which is
     the same code path the Oasis uses -- this is the check that matters
  3. every controlled value is a member of its MEnum
  4. dates parse, and no event or contribution ends before it starts
  5. no two files collide on a filename or an entry name
  6. every event in the curated sheet produced exactly one file

Usage::

    uv run python packages/fairmat-project-outputs/scripts/validate_archives.py [dir]
"""

from __future__ import annotations

import datetime
import importlib
import json
import sys
import warnings
from collections import Counter, defaultdict
from pathlib import Path

warnings.filterwarnings('ignore', module='openpyxl')

import openpyxl  # noqa: E402
from build_curated_sheet import HEADERS  # noqa: E402
from nomad.datamodel import EntryArchive  # noqa: E402
from nomad.utils import get_logger  # noqa: E402

from fairmat_project_outputs.schema_packages.schema_package import (  # noqa: E402
    CONTRIBUTION_ROLE,
    CONTRIBUTION_TYPE,
    EVENT_MODE,
    EVENT_PURPOSE,
    EVENT_SERIES,
    EVENT_TYPE,
    FABIO_TERMS,
    FAIRMAT_ROLE,
    FAIRmatEvent,
)

ROOT = Path(__file__).resolve().parents[1]
LOCAL = ROOT / 'local'
EXPECTED_DEF = (
    'fairmat_project_outputs.schema_packages.schema_package.FAIRmatEvent'
)

EVENT_ENUMS = {
    'event_type': EVENT_TYPE,
    'event_series': EVENT_SERIES,
    'mode': EVENT_MODE,
    'fairmat_role': FAIRMAT_ROLE,
}
LIST_ENUMS = {'event_purpose': EVENT_PURPOSE}
CONTRIBUTION_ENUMS = {
    'contribution_type': CONTRIBUTION_TYPE,
    'contribution_role': CONTRIBUTION_ROLE,
    'fabio_type': FABIO_TERMS,
}


def resolve_section(reference: str):
    """Resolve an m_def string the way an Oasis resolves it: by import.

    A plugin m_def is a Python path. If the plugin is not installed, this is
    exactly where the upload would fail, so it is worth failing here instead.
    """
    module_path, _, name = reference.rpartition('.')
    module = importlib.import_module(module_path)
    return getattr(module, name).m_def


def members(enum) -> set[str]:
    return set(getattr(enum, 'values', None) or enum)


def parse_date(value: str):
    return datetime.datetime.fromisoformat(str(value).replace('Z', '+00:00'))


def check_enums(where: str, data: dict, enums: dict, fail) -> None:
    for field, enum in enums.items():
        value = data.get(field)
        if value is not None and value not in members(enum):
            fail(f'{where}: {field}={value!r} is not in the schema vocabulary')


def check_dates(where: str, data: dict, fail) -> tuple:
    start = end = None
    for field in ('start_date', 'end_date'):
        if data.get(field) is None:
            continue
        try:
            parsed = parse_date(data[field])
        except (ValueError, TypeError):
            fail(f'{where}: {field}={data[field]!r} is not a valid date')
            continue
        if field == 'start_date':
            start = parsed
        else:
            end = parsed
    if start and end and end < start:
        fail(f'{where}: ends {end.date()} before it starts {start.date()}')
    return start, end


def sheet_events(path: Path) -> set[str]:
    sheet = openpyxl.load_workbook(path, data_only=True)['Events']
    by_header = {header: key for key, header in HEADERS.items()}
    header = [by_header.get(c.value, c.value) for c in sheet[1]]
    column = header.index('Canonical event name')
    return {
        str(row[column].value).strip()
        for row in sheet.iter_rows(min_row=2)
        if row[column].value
    }


class Tally:
    """What one pass over the files found."""

    def __init__(self) -> None:
        self.failures: list[str] = []
        self.names: Counter = Counter()
        self.stats: Counter = Counter()
        self.filled: Counter = Counter()
        self.contribution_filled: Counter = Counter()
        self.years: Counter = Counter()

    def fail(self, problem: str) -> None:
        self.failures.append(problem)


def check_list_enums(label: str, data: dict, tally: Tally) -> None:
    for field, enum in LIST_ENUMS.items():
        values = data.get(field) or []
        if not isinstance(values, list):
            tally.fail(f'{label}: {field} should be a list, got {type(values).__name__}')
            continue
        for value in values:
            if value not in members(enum):
                tally.fail(f'{label}: {field} value {value!r} is not in the schema')
        if len(values) != len(set(values)):
            tally.fail(f'{label}: {field} repeats a value')


def check_contributions(label: str, data: dict, tally: Tally) -> None:
    for index, contribution in enumerate(data.get('contributions') or []):
        sub = f'{label}[contribution {index}]'
        if not str(contribution.get('title') or '').strip():
            tally.fail(f'{sub}: no title')
        check_enums(sub, contribution, CONTRIBUTION_ENUMS, tally.fail)
        check_dates(sub, contribution, tally.fail)
        for field in contribution:
            tally.contribution_filled[field] += 1
        tally.stats['contributions'] += 1


def check_processing(label: str, data: dict, archive, tally: Tally) -> None:
    """Run what the Oasis runs: normalize(), then check nothing was dropped."""
    try:
        archive.data.normalize(archive, get_logger(__name__))
        tally.stats['normalized'] += 1
    except Exception as exc:  # noqa: BLE001
        tally.fail(f'{label}: normalize() failed -- {type(exc).__name__}: {exc}')

    # Anything NOMAD did not understand is dropped silently on the way in, so
    # a value that does not survive the round trip was never really stored.
    try:
        round_tripped = archive.data.m_to_dict()
    except Exception as exc:  # noqa: BLE001
        tally.fail(f'{label}: m_to_dict() failed -- {type(exc).__name__}: {exc}')
        return
    for field, value in data.items():
        # normalize() rewrites the overview and rebuilds the contributor list
        # from the contributions; both are meant to change.
        if field in ('m_def', 'contributions', 'contributions_overview',
                     'fairmat_contributors'):
            continue
        if field not in round_tripped:
            tally.fail(f'{label}: {field}={value!r} was dropped by NOMAD')


def check_file(path: Path, tally: Tally) -> None:
    label = path.name
    try:
        raw = json.loads(path.read_text(encoding='utf-8'))
    except json.JSONDecodeError as exc:
        tally.fail(f'{label}: not valid JSON -- {exc}')
        return

    data = raw.get('data')
    if not isinstance(data, dict):
        tally.fail(f'{label}: no "data" section')
        return
    if data.get('m_def') != EXPECTED_DEF:
        tally.fail(f'{label}: m_def is {data.get("m_def")!r}')

    # The checks that matter: NOMAD's own deserialiser and then the schema's
    # own normalize(), which is what the Oasis runs on upload. A file that
    # parses but blows up in normalize() still fails processing.
    archive = None
    try:
        archive = EntryArchive.m_from_dict(raw)
    except Exception as exc:  # noqa: BLE001 - any failure is a failure
        tally.fail(f'{label}: NOMAD cannot read this file -- {type(exc).__name__}: {exc}')

    if archive is not None:
        check_processing(label, data, archive, tally)

    if not str(data.get('event_name') or '').strip():
        tally.fail(f'{label}: no event_name')
    tally.names[str(data.get('event_name') or '')] += 1

    check_enums(label, data, EVENT_ENUMS, tally.fail)
    check_list_enums(label, data, tally)
    start, _ = check_dates(label, data, tally.fail)
    if start:
        tally.years[start.year] += 1

    for field in data:
        if field != 'm_def':
            tally.filled[field] += 1
    check_contributions(label, data, tally)
    tally.stats['entries'] += 1


def report(where: Path, tally: Tally) -> None:
    entries = tally.stats['entries']
    print(f'checked {entries} entries / {tally.stats["contributions"]} '
          f'contributions in {where}')
    print(f'  schema: {EXPECTED_DEF.rsplit(".", 1)[1]}')
    print('\nfield coverage (event):')
    for field in FAIRmatEvent.m_def.all_quantities:
        print(f'  {field:26s} {tally.filled.get(field, 0):4d} / {entries}')
    print(f'  {"contributions":26s} {tally.filled.get("contributions", 0):4d} / {entries}')
    print('\nfield coverage (contribution):')
    for field, count in tally.contribution_filled.most_common():
        print(f'  {field:26s} {count:4d} / {tally.stats["contributions"]}')
    if tally.years:
        print('\nentries per year:',
              ', '.join(f'{y}: {n}' for y, n in sorted(tally.years.items())))
    print()


def check_against_sheet(tally: Tally) -> None:
    """Every event in the sheet became exactly one entry, and vice versa."""
    # Filenames are unique by construction; entry names must be too, or two
    # entries silently describe the same event.
    for name, count in tally.names.items():
        if count > 1:
            tally.fail(f'entry name {name!r} used by {count} files')

    sheets = sorted(LOCAL.glob('Events database - CURATED_*.xlsx'))
    if not sheets:
        return
    expected = sheet_events(sheets[-1])
    produced = set(tally.names)
    for missing in sorted(expected - produced):
        tally.fail(f'event in the sheet but not in the archives: {missing!r}')
    for extra in sorted(produced - expected):
        tally.fail(f'archive with no event in the sheet: {extra!r}')


def main() -> int:
    where = Path(sys.argv[1]) if len(sys.argv) > 1 else None
    if where is None:
        folders = sorted(LOCAL.glob('archives_*'))
        if not folders:
            raise SystemExit('no archives_* folder found')
        where = folders[-1]

    files = sorted(where.glob('*.archive.json'))
    if not files:
        raise SystemExit(f'no *.archive.json files in {where}')

    # The m_def strings are Python references into the plugin. They resolve
    # only if the plugin is installed in the Oasis, so prove it resolves here
    # rather than discovering it after an upload.
    try:
        resolved = resolve_section(EXPECTED_DEF)
        print(f'm_def resolves to: {resolved.qualified_name()}')
        print(
            f"  quantities: {len(resolved.all_quantities)}, "
            f"sub-sections: {len(resolved.all_sub_sections)}"
        )
    except Exception as exc:  # noqa: BLE001
        print(f'm_def DOES NOT RESOLVE: {type(exc).__name__}: {exc}')
        return 1

    tally = Tally()
    for path in files:
        check_file(path, tally)

    check_against_sheet(tally)
    report(where, tally)
    failures = tally.failures
    if failures:
        print(f'FAILED -- {len(failures)} problems')
        grouped: dict[str, int] = defaultdict(int)
        for problem in failures:
            grouped[problem.split(':', 1)[-1].strip()[:70]] += 1
        for problem, count in sorted(grouped.items(), key=lambda kv: -kv[1])[:25]:
            print(f'  x{count:<4d} {problem}')
        return 1
    print('PASSED -- every file is valid against the schema and self-consistent')
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
