#!/usr/bin/env python3
"""Decide whether an event speaker is a FAIRmat person or an outside guest.

The events spreadsheet records every speaker in one ``Speaker/Participant``
cell and says nothing about affiliation, but the schema separates them:
``Contribution.fairmat_contributors`` vs ``Contribution.other_contributors``.
This module supplies the missing bit by matching each name against a roster of
FAIRmat people.

The roster lives in ``local/fairmat_roster.json`` (names only, no e-mail
addresses) and is the union of three sources:

* the public team page, <https://fair-di.eu/fairmat/about-fairmat/team-fairmat>
  -- including its **Previous members** section, which matters because someone
  who has since left was still FAIRmat when they gave their talk;
* ``packages/fairmat-events-form/fairmat_team.json``;
* ``packages/fairmat-members/local/members.csv``.

Matching is on **surname + first initial**, because the sheet writes
``C. Draxl`` while every roster writes ``Claudia Draxl``. That key is safe
here: across all three rosters no two people share a surname and a first
initial, so a match is never a coin toss.

A name that matches nothing is treated as external. On the current sheet that
leaves 41 people -- and they are overwhelmingly the invited seminar speakers
(Taylor Sparks, Barend Mons, Isao Tanaka, ...), which is the expected answer.

Confidence is reported, never hidden:

``exact``         surname + initial agree with exactly one person.
``fuzzy``         surname is a near-miss (``M. Aeshlimann``); repaired.
``surname-only``  no initial in the source; one roster surname matched.
``ambiguous``     more than one roster spelling matched.
``none``          external.

Only ``exact`` is written unreviewed. Everything else is flagged so a human
decides -- ``J. Marques`` resolves to *Miguel A. L. Marques* on surname alone
but is really a typo for ``J. Marquez`` (*José*), which is exactly the kind of
call that must not be made silently.
"""

from __future__ import annotations

import difflib
import json
import re
import unicodedata
from collections import defaultdict
from pathlib import Path

ROSTER_PATH = Path(__file__).resolve().parents[1] / 'local' / 'fairmat_roster.json'

# 'ü' -> 'ue' before the accent strip, so 'Bürki' and 'Buerki' land on the
# same key. A plain NFKD would give 'burki' and miss it.
_UMLAUT = str.maketrans(
    {'ü': 'ue', 'ö': 'oe', 'ä': 'ae', 'ß': 'ss',
     'Ü': 'Ue', 'Ö': 'Oe', 'Ä': 'Ae'}
)
_TITLES = re.compile(r'\b(prof|dr|pd|jun|ph\s*\.?\s*d|phd|habil|mr|ms|mrs)\b\.?', re.I)
_PARTICLES = {'von', 'van', 'de', 'der', 'di', 'del', 'dos', 'da', 'v'}
_JOB_WORDS = re.compile(r'\b(director|lead|head|professor|engineering|canonical)\b', re.I)

MIN_SURNAME_LENGTH = 2
FUZZY_SURNAME_CUTOFF = 0.84

SKIP_TOKENS = {
    'various', 'n/a', 'na', 'tbd', 'tbc', '-', '',
    'organisation team', 'organization team',
}


def _fold(text: str) -> str:
    text = text.translate(_UMLAUT)
    return unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode()


def clean_person(raw: str) -> str:
    """Strip academic titles, parentheticals and stray punctuation."""
    text = _TITLES.sub(' ', raw)
    text = re.sub(r'\(.*?\)', ' ', text)
    text = text.replace(',', ' ')
    return re.sub(r'\s+', ' ', _fold(text)).strip(' .,;:')


def name_key(raw: str) -> tuple[str, str, str] | None:
    """``'H. v. Wenckstern'`` -> ``('h', 'von wenckstern', 'H. v. Wenckstern')``."""
    text = clean_person(raw)
    # 'C.Draxl' -> 'C. Draxl'
    text = re.sub(r'\b([A-Z])\.(?=[A-Z][a-z])', r'\1. ', text)
    parts = [p for p in text.split() if p]
    if not parts:
        return None
    last = len(parts) - 1
    surname = [parts[last]]
    while last - 1 >= 1 and parts[last - 1].lower().strip('.') in _PARTICLES:
        last -= 1
        surname.insert(0, parts[last].strip('.'))
    first = parts[0].lower().strip('.') if last > 0 else ''
    return first[:1], ' '.join(surname).lower(), text


def split_speakers(cell: str) -> list[str]:
    """Split one ``Speaker/Participant`` cell into individual people.

    Handles the separators actually present in the sheet: commas, semicolons,
    ' and ', a leading 'Organisation team:', a trailing job description, and
    rows where a period stands in for a comma ('A. Wojas. S. Nakhaie') or a
    comma stands in for a period ('N, Slepenko').
    """
    if not cell:
        return []
    text = re.sub(
        r'^\s*(Organisation team|Organization team|Organizers?)\s*:\s*', '', cell, flags=re.I
    )
    if _JOB_WORDS.search(text):
        text = re.sub(r':\s*.*$', '', text)
    text = re.sub(r'\s+and\s+', ',', text)
    # Protect a particle's period ('v. Wenckstern') before reading periods as
    # separators, then restore it.
    text = re.sub(
        r'\b(v|von|van|de|der|di|del)\.\s+', lambda m: m.group(1) + '\x00 ', text, flags=re.I
    )
    text = re.sub(r'(?<=[a-z])\.\s+(?=[A-Z])', ', ', text).replace('\x00', '.')

    parts = [p.strip(' .;:') for p in re.split(r'[,;]', text)]
    people: list[str] = []
    index = 0
    while index < len(parts):
        part = parts[index]
        if re.fullmatch(r'[A-Z]', part) and index + 1 < len(parts):
            people.append(f'{part}. {parts[index + 1]}')
            index += 2
            continue
        if part and part.lower() not in SKIP_TOKENS:
            people.append(part)
        index += 1
    return people


class Roster:
    """Surname+initial lookup over the FAIRmat people roster."""

    def __init__(self, path: Path = ROSTER_PATH):
        data = json.loads(path.read_text(encoding='utf-8'))
        self.names: list[str] = []
        for field in (
            'current_members', 'previous_members', 'team_json_only', 'members_csv_only'
        ):
            self.names.extend(data.get(field, []))
        self.current = set(data.get('current_members', []))

        self._exact: dict[tuple[str, str], set[str]] = defaultdict(set)
        self._by_surname: dict[str, set[str]] = defaultdict(set)
        for name in self.names:
            key = name_key(name)
            if not key:
                continue
            initial, surname, cleaned = key
            if initial:
                self._exact[(initial, surname)].add(cleaned)
            self._by_surname[surname].add(cleaned)
        self._surnames = list(self._by_surname)

    def match(self, token: str) -> tuple[str | None, str]:
        """Return ``(canonical name or None, confidence)``."""
        key = name_key(token)
        if not key or len(key[1]) < MIN_SURNAME_LENGTH:
            return None, 'none'
        initial, surname, _ = key

        if initial and (initial, surname) in self._exact:
            found = self._exact[(initial, surname)]
            return sorted(found)[0], 'exact' if len(found) == 1 else 'ambiguous'

        if surname in self._by_surname:
            found = self._by_surname[surname]
            return sorted(found)[0], 'surname-only' if len(found) == 1 else 'ambiguous'

        near = difflib.get_close_matches(surname, self._surnames, n=1, cutoff=FUZZY_SURNAME_CUTOFF)
        if near:
            found = self._by_surname[near[0]]
            # A fuzzy surname is only trusted when the initial still agrees.
            same_initial = [
                n for n in found if not initial or clean_person(n)[:1].lower() == initial
            ]
            if same_initial:
                return sorted(same_initial)[0], 'fuzzy'
        return None, 'none'

    def classify(self, cell: str) -> tuple[list[str], list[str], list[str]]:
        """Split a Speaker cell into (fairmat people, outside people, flags)."""
        internal: list[str] = []
        external: list[str] = []
        flags: list[str] = []
        for token in split_speakers(cell):
            name, confidence = self.match(token)
            if confidence == 'none':
                external.append(token)
                continue
            internal.append(name or token)
            if confidence != 'exact':
                flags.append(f'{token!r} matched {name!r} ({confidence})')
        return internal, external, flags


if __name__ == '__main__':
    roster = Roster()
    print(f'{len(roster.names)} people in the roster ({len(roster.current)} current)')
    for probe in ('C. Draxl', 'H. v. Wenckstern', 'Taylor Sparks', 'J. Marques'):
        print(f'  {probe!r:22} -> {roster.match(probe)}')
