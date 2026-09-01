"""
Generic parser for event CSV files whose names end with ``_events.csv``
(e.g. ``Cordi_events.csv``, ``workshop_events.csv``).

CSV layout (with header row, columns 0-indexed):
  Col 0  (From)                  – Start date
  Col 1  (To)                    – End date
  Col 2  (Name)                  – Event name / contribution name
  Col 3  (Speaker/Participant)   – Contributor(s); comma-separated → fairmat_contributors
  Col 4  (Event type main)       – Event type → schema EVENT_TYPE (default: Conference)
  Col 5  (Event type sub)        – Contribution type → schema CONTRIBUTION_TYPE
  Col 6  (FAIRmat role)          – FAIRmat primary role → schema FAIRMAT_ROLE
  Col 7  (Link / URL)            – FaBiO URL
  Col 8  (FaBiO type)            – FaBiO type
  Col 9  (Educational Purpose)   – Truthy → event_purpose "Educational event"
  Col 10 (Location)              – Location name
  Col 11 (City)                  – City
  Col 12 (Country)               – Country

**Row grouping (prefix matching)**:
  • The first row whose ``Name`` cell is NOT a suffix-extension of any
    previously-registered event name is treated as the **main event row**.
    It sets event-level metadata (type, role, dates, location, …).
  • Every subsequent row whose ``Name`` starts with a registered event's base
    name (plus extra text, e.g. `` talk: …``) is a **contribution row** for
    that event.  The contribution title and keyword-based type are extracted
    from the text that follows the base event name.
"""

from __future__ import annotations

import re
from datetime import datetime
from typing import TYPE_CHECKING

import yaml
from nomad.parsing.parser import MatchingParser

if TYPE_CHECKING:
    from nomad.datamodel.datamodel import EntryArchive
    from structlog.stdlib import BoundLogger

# ---------------------------------------------------------------------------
# Column indices  (0-based, after header row is consumed by pandas)
# ---------------------------------------------------------------------------

_COL_START = 0
_COL_END = 1
_COL_NAME = 2
_COL_CONTRIBUTORS = 3
_COL_EVENT_TYPE_MAIN = 4
_COL_EVENT_TYPE_SUB = 5
_COL_FAIRMAT_ROLE = 6
_COL_FABIO_URL = 7
_COL_FABIO_TYPE = 8
_COL_EDUCATIONAL = 9
_COL_LOCATION = 10
_COL_CITY = 11
_COL_COUNTRY = 12

# ---------------------------------------------------------------------------
# Date parsing
# ---------------------------------------------------------------------------

_DATE_FORMATS = (
    "%d-%b-%y", "%d-%b-%Y",
    "%Y-%m-%d", "%d/%m/%Y",
    "%m/%d/%Y", "%d.%m.%Y",
    "%B %d, %Y", "%b %d, %Y",
    "%d %B %Y", "%d %b %Y",
)


def _parse_date(raw) -> datetime | None:
    if raw is None:
        return None
    s = str(raw).strip()
    if not s or s.lower() in ("nan", "none", ""):
        return None
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


# ---------------------------------------------------------------------------
# Cell cleaner
# ---------------------------------------------------------------------------

def _clean(value) -> str:
    if value is None:
        return ""
    s = re.sub(r"\s+", " ", str(value).replace("\r", " ").strip())
    return "" if s.lower() in ("nan", "none", "") else s


def _icell(row, idx: int) -> str:
    """Return a cleaned string from a DataFrame row by positional index."""
    try:
        return _clean(row.iloc[idx])
    except (IndexError, KeyError):
        return ""


# ---------------------------------------------------------------------------
# Event type (main) normalisation  →  schema EVENT_TYPE
# ---------------------------------------------------------------------------

_MAIN_TYPE_MAP: dict[str, str] = {
    "s": "Seminar",
    "scholarly talks (s)": "Seminar",
    "scholarly talks": "Seminar",
    "ce": "Workshop",
    "community engagement (ce)": "Workshop",
    "community engagement": "Workshop",
    "cm": "Conference",
    "conferences and meetings (cm)": "Conference",
    "conferences and meetings": "Conference",
    "e": "Tutorial",
    "educational events (e)": "Tutorial",
    "educational events": "Tutorial",
    "pe": "Public event",
    "public engagement (pe)": "Public event",
    "public engagement": "Public event",
}

_EVENT_TYPE_KEYWORD_MAP: list[tuple[str, str]] = [
    ("colloquium", "Colloquium"),
    ("symposium", "Symposium"),
    ("hackathon", "Hackathon"),
    ("conference", "Conference"),
    ("workshop", "Workshop"),
    ("school", "School"),
    ("tutorial", "Tutorial"),
    ("course", "Course"),
    ("seminar", "Seminar"),
    ("meeting", "Meeting"),
    ("demonstration", "Demonstration"),
    ("booth", "Public event"),
]

_VALID_EVENT_TYPES = {
    "Conference", "Workshop", "School", "Meeting", "Seminar", "Tutorial",
    "Hackathon", "Symposium", "Colloquium", "Course", "Demonstration", "Public event",
}


def _normalise_event_type(raw: str, event_name: str = "") -> str:
    """Map raw event-type-main string (+ optional event name) → schema EVENT_TYPE."""
    if raw:
        lower = raw.strip().lower()
        if lower in _MAIN_TYPE_MAP:
            return _MAIN_TYPE_MAP[lower]
        for valid in _VALID_EVENT_TYPES:
            if valid.lower() == lower:
                return valid
        for key, val in _MAIN_TYPE_MAP.items():
            if lower.startswith(key):
                return val
    # Fallback: infer from event name keywords
    for keyword, etype in _EVENT_TYPE_KEYWORD_MAP:
        if keyword in event_name.lower():
            return etype
    return "Conference"


# ---------------------------------------------------------------------------
# Event type (sub) normalisation  →  schema CONTRIBUTION_TYPE
# ---------------------------------------------------------------------------

_PREFIX_RE = re.compile(r"^(?:S|Ce|CM|E)_", re.IGNORECASE)

_SUBTYPE_TO_CTYPE: dict[str, str] = {
    "invited talks": "Invited talk",
    "invited talk": "Invited talk",
    "plenary talks": "Plenary talk",
    "plenary talk": "Plenary talk",
    "contributed talks": "Contributed talk",
    "contributed talk": "Contributed talk",
    "information booths": "Information booth",
    "information booth": "Information booth",
    "poster": "Poster",
    "posters": "Poster",
    "tutorial": "Tutorial",
    "tutorials": "Tutorial",
    "workshop session": "Workshop session",
    "workshop sessions": "Workshop session",
    "panel discussion": "Panel discussion",
    "panel discussions": "Panel discussion",
    "conference session": "Conference session",
    "conference sessions": "Conference session",
    "symposium": "Symposium",
}


def _ctype_from_subtype(raw_sub: str) -> str | None:
    """Map Event type (sub) value → CONTRIBUTION_TYPE."""
    if not raw_sub:
        return None
    stripped = _PREFIX_RE.sub("", raw_sub).strip().lower()
    return _SUBTYPE_TO_CTYPE.get(stripped)


# ---------------------------------------------------------------------------
# Contribution keyword parsing from name suffix
# ---------------------------------------------------------------------------

# Ordered most-specific first; keys are lower-cased
_KEYWORD_TO_CTYPE: list[tuple[str, str]] = [
    ("plenary talk", "Plenary talk"),
    ("invited talk", "Invited talk"),
    ("contributed talk", "Contributed talk"),
    ("information booth", "Information booth"),
    ("market place", "Information booth"),
    ("marketplace", "Information booth"),
    ("panel discussion", "Panel discussion"),
    ("panel", "Panel discussion"),
    ("poster session", "Conference session"),
    ("poster", "Poster"),
    ("tutorial", "Tutorial"),
    ("workshop session", "Workshop session"),
    ("workshop", "Workshop session"),
    ("conference session", "Conference session"),
    ("session", "Conference session"),
    ("talk", "Contributed talk"),
    ("lecture", "Invited talk"),
    ("presentation", "Contributed talk"),
    ("demonstration", "Contributed talk"),
    ("keynote", "Plenary talk"),
]


def _parse_contribution_suffix(suffix: str) -> tuple[str, str | None]:
    """
    Given the text after the base event name (e.g. `` talk: Introduction to…``),
    return ``(contribution_title, keyword_ctype_or_None)``.

    The suffix is expected to start with an optional space, then a keyword,
    an optional colon, and then the actual title.
    """
    suffix = suffix.strip()
    if not suffix:
        return "", None

    # Try to find the first ": " — everything before it is the keyword phrase,
    # everything after is the title.
    colon_idx = suffix.find(": ")
    if colon_idx != -1:
        keyword_phrase = suffix[:colon_idx].strip().lower()
        title = suffix[colon_idx + 2:].strip()
    else:
        # No colon — the whole suffix is the "keyword phrase" with no explicit title
        keyword_phrase = suffix.lower()
        title = suffix

    # Map keyword phrase to contribution type
    ctype = None
    for kw, ct in _KEYWORD_TO_CTYPE:
        if kw in keyword_phrase:
            ctype = ct
            break

    return title, ctype


# ---------------------------------------------------------------------------
# FAIRmat role normalisation
# ---------------------------------------------------------------------------

_FAIRMAT_ROLE_MAP: dict[str, str] = {
    "organizer": "Organizer",
    "co-organizer": "Co-organizer",
    "coorganizer": "Co-organizer",
    "co organizer": "Co-organizer",
    "supporter": "Supporter",
    "support": "Supporter",
    "participant": "Participant",
    "speaker": "Participant",
    "exhibitor": "Participant",
    "host": "Organizer",
    "presenter": "Participant",
    "invited talk": "Participant",
}


def _normalise_fairmat_role(raw: str) -> str | None:
    if not raw:
        return None
    lower = raw.strip().lower()
    if lower in _FAIRMAT_ROLE_MAP:
        return _FAIRMAT_ROLE_MAP[lower]
    for key, val in _FAIRMAT_ROLE_MAP.items():
        if key in lower:
            return val
    return None


# ---------------------------------------------------------------------------
# FaBiO type normalisation
# ---------------------------------------------------------------------------

_VALID_FABIO_LOWER: dict[str, str] = {
    "fabio: abstract": "fabio: Abstract",
    "fabio: announcement": "fabio: Announcement",
    "fabio: book chapter": "fabio: Book chapter",
    "fabio: computer program": "fabio: Computer program",
    "fabio: conference proceedings": "fabio: Conference proceedings",
    "fabio: entity metadata": "fabio: Entity metadata",
    "fabio: instructional work": "fabio: Instructional work",
    "fabio: interview": "fabio: Interview",
    "fabio: journal article": "fabio: Journal article",
    "fabio: meeting report": "fabio: Meeting report",
    "fabio: periodical issue": "fabio: Periodical issue",
    "fabio: poster": "fabio: Poster",
    "fabio: preprint": "fabio: Preprint",
    "fabio: presentation": "fabio: Presentation",
    "fabio: scholarly work": "fabio: Scholarly work",
    "fabio: software dataset": "fabio: Software dataset",
    "fabio: timetable": "fabio: Timetable",
    "fabio: web page": "fabio: Web page",
}


def _normalise_fabio_type(raw: str) -> str | None:
    if not raw:
        return None
    # Strip surrounding brackets: [fabio: Web page] → fabio: Web page
    clean = re.sub(r"^\[|\]$", "", raw.strip())
    # Also handle missing space: "fabio:Web page" → "fabio: Web page"
    clean = re.sub(r"(?i)^fabio:\s*", "fabio: ", clean).strip()
    lower = clean.lower()
    if lower in _VALID_FABIO_LOWER:
        return _VALID_FABIO_LOWER[lower]
    # Partial suffix match
    for key, val in _VALID_FABIO_LOWER.items():
        suffix = key.replace("fabio: ", "")
        if suffix == lower or suffix in lower:
            return val
    return None


# ---------------------------------------------------------------------------
# Educational flag  →  event_purpose
# ---------------------------------------------------------------------------

_TRUTHY = {"yes", "true", "1", "y", "x", "educational", "educational event"}


def _is_educational(raw: str) -> bool:
    return raw.strip().lower() in _TRUTHY


# ---------------------------------------------------------------------------
# Contributor splitting
# ---------------------------------------------------------------------------

_SKIP_CONTRIBUTORS = {"various", "n/a", "na", "tbd", "tbc", ""}


def _split_contributors(raw: str) -> list[str] | None:
    if not raw:
        return None
    names = [n.strip() for n in re.split(r"[,;]", raw) if n.strip()]
    names = [n for n in names if n.lower() not in _SKIP_CONTRIBUTORS]
    return names or None


# ---------------------------------------------------------------------------
# CSV / DataFrame loader
# ---------------------------------------------------------------------------

def _load_dataframe(mainfile: str):
    import pandas as pd

    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            return pd.read_csv(
                mainfile,
                header=0,       # first row is the header
                dtype=str,
                encoding=enc,
                skip_blank_lines=True,
            )
        except UnicodeDecodeError:
            continue
    raise ValueError(f"Cannot decode {mainfile} with any supported encoding")


# ---------------------------------------------------------------------------
# Row grouping — prefix matching
# ---------------------------------------------------------------------------

def _find_parent_key(
    name_lower: str,
    registry: list[tuple[str, str]],
) -> str | None:
    """Return the event key whose base name is a prefix of *name_lower*, or None."""
    for base_lower, key in registry:
        if (
            name_lower.startswith(base_lower)
            and len(name_lower) > len(base_lower)
            and name_lower[len(base_lower)] in (" ", ":", ",", "(", "-", "_")
        ):
            return key
    return None


def _register_event(base_lower: str, key: str, registry: list[tuple[str, str]]) -> None:
    """Insert *key* into *registry* so that longer base names come first."""
    for i, (b, _k) in enumerate(registry):
        if len(base_lower) >= len(b):
            registry.insert(i, (base_lower, key))
            return
    registry.append((base_lower, key))


def _expand_dates(info: dict, start_dt: datetime | None, end_dt: datetime | None) -> None:
    if start_dt and (info["start"] is None or start_dt < info["start"]):
        info["start"] = start_dt
    if end_dt and (info["end"] is None or end_dt > info["end"]):
        info["end"] = end_dt


def _process_contribution_row(
    row, name: str, matched_key: str, event_map: dict
) -> None:
    info = event_map[matched_key]
    base_name = info["event_name"]
    suffix = name[len(base_name):]
    contrib_title, keyword_ctype = _parse_contribution_suffix(suffix)
    sub_raw = _icell(row, _COL_EVENT_TYPE_SUB)
    ctype = _ctype_from_subtype(sub_raw) or keyword_ctype
    start_dt = _parse_date(_icell(row, _COL_START))
    end_dt = _parse_date(_icell(row, _COL_END))
    _expand_dates(info, start_dt, end_dt)
    info["contributions"].append({
        "title": contrib_title,
        "contribution_type": ctype,
        "contributors": _split_contributors(_icell(row, _COL_CONTRIBUTORS)),
        "fabio_url": _icell(row, _COL_FABIO_URL) or None,
        "fabio_type": _normalise_fabio_type(_icell(row, _COL_FABIO_TYPE)),
        "start_dt": start_dt,
        "end_dt": end_dt,
    })


def _process_main_event_row(
    row, name: str, registry: list[tuple[str, str]], event_map: dict
) -> None:
    base_name = name.strip()
    base_lower = base_name.lower()
    start_dt = _parse_date(_icell(row, _COL_START))
    end_dt = _parse_date(_icell(row, _COL_END))

    existing_key = next((k for b, k in registry if b == base_lower), None)
    if existing_key:
        _expand_dates(event_map[existing_key], start_dt, end_dt)
        return

    key = re.sub(r"[^\w]", "_", base_lower)[:80]
    educational = _is_educational(_icell(row, _COL_EDUCATIONAL))
    event_map[key] = {
        "event_name": base_name,
        "start": start_dt,
        "end": end_dt,
        "event_type": _normalise_event_type(_icell(row, _COL_EVENT_TYPE_MAIN), base_name),
        "fairmat_role": _normalise_fairmat_role(_icell(row, _COL_FAIRMAT_ROLE)),
        "event_purpose": ["Educational event"] if educational else None,
        "location": _icell(row, _COL_LOCATION) or None,
        "city": _icell(row, _COL_CITY) or None,
        "country": _icell(row, _COL_COUNTRY) or None,
        "event_url": _icell(row, _COL_FABIO_URL) or None,
        "contributions": [],
    }
    _register_event(base_lower, key, registry)


def _group_rows(df, logger) -> dict:
    """
    Build an ordered dict ``{key: event_info}`` by processing rows in order.

    Rows are classified by prefix matching against the running registry of
    known event base names (longest first).  A row whose Name starts with a
    registered base name (+ extra text) is a contribution; otherwise it
    registers a new event.
    """
    registry: list[tuple[str, str]] = []   # (base_name_lower, key), longest-first
    event_map: dict[str, dict] = {}

    for _, row in df.dropna(how="all").iterrows():
        name = _icell(row, _COL_NAME)
        if not name:
            continue
        matched_key = _find_parent_key(name.strip().lower(), registry)
        if matched_key:
            _process_contribution_row(row, name.strip(), matched_key, event_map)
        else:
            _process_main_event_row(row, name, registry, event_map)

    return event_map


# ---------------------------------------------------------------------------
# Building schema objects
# ---------------------------------------------------------------------------

def _build_contributions(contrib_rows: list[dict]) -> list:
    from fairmat_project_outputs.schema_packages.schema_package import Contribution

    contributions = []
    for row in contrib_rows:
        contributions.append(
            Contribution(
                title=row["title"] or None,
                contribution_type=row["contribution_type"],
                fairmat_contributors=row["contributors"],
                start_date=row["start_dt"],
                end_date=row["end_dt"],
                fabio_type=row["fabio_type"],
                fabio_url=row["fabio_url"],
            )
        )
    return contributions


def _build_event_record(key: str, info: dict):
    from fairmat_project_outputs.schema_packages.schema_package import (
        EventRecord,
        FAIRmatEvent,
    )

    contributions = _build_contributions(info["contributions"])

    common = {
        "event_name": info["event_name"],
        "start_date": info["start"],
        "end_date": info["end"],
        "event_type": info["event_type"],
        "fairmat_role": info["fairmat_role"],
        "event_purpose": info["event_purpose"],
        "location_name": info["location"],
        "city": info["city"],
        "country": info["country"],
        "event_url": info.get("event_url"),
        "contributions": contributions,
    }
    event_name = info["event_name"]
    return EventRecord(**common), FAIRmatEvent(**common), event_name


# ---------------------------------------------------------------------------
# Child-archive writer
# ---------------------------------------------------------------------------

def _write_child_archive(
    event,
    event_name: str,
    key: str,
    archive: EntryArchive,
    logger: BoundLogger,
) -> None:
    from nomad.datamodel import EntryArchive as EA
    from nomad.datamodel import EntryMetadata

    safe_name = re.sub(r"[^\w\-]", "_", event_name or "event")[:60]
    safe_key = re.sub(r"[^\w]", "_", key)[:30]
    filename = f"generic_event_{safe_key}_{safe_name}.archive.yaml"

    child_archive = EA(
        data=event,
        m_context=archive.m_context,
        metadata=EntryMetadata(
            upload_id=archive.m_context.upload_id,
            entry_name=event_name or "Event",
        ),
    )
    try:
        with archive.m_context.raw_file(filename, "w") as outfile:
            yaml.dump(child_archive.m_to_dict(), outfile)
        archive.m_context.upload.process_updated_raw_file(filename, allow_modify=True)
        logger.info("created child entry", filename=filename, event_name=event_name)
    except Exception as exc:  # noqa: BLE001
        logger.error("Failed to create child entry", event_name=event_name, exc_info=exc)


# ---------------------------------------------------------------------------
# Parser
# ---------------------------------------------------------------------------


class GenericEventParser(MatchingParser):
    """
    Parses any ``*_events.csv`` file into ``FAIRmatEvent`` entries.

    The CSV is expected to have a header row. Rows are grouped using
    **prefix matching on the Name column**:

    * A row whose Name is not an extension of any previously-seen event name
      becomes a **main event** (sets event-level metadata).
    * A row whose Name starts with an existing event's name (followed by
      additional text such as `` talk: …``) becomes a **Contribution**
      sub-section of that event.

    The contribution title is extracted from the text after the base event
    name; the contribution type is inferred from the ``Event type (sub)``
    column and/or keywords in the name suffix (e.g., ``talk``, ``poster``,
    ``tutorial``).
    """

    def parse(
        self,
        mainfile: str,
        archive: EntryArchive,
        logger: BoundLogger,
        child_archives: dict[str, EntryArchive] = None,
    ) -> None:
        from fairmat_project_outputs.schema_packages.schema_package import (
            FAIRmatEventsFile,
        )

        logger.info("GenericEventParser.parse", mainfile=mainfile)

        try:
            df = _load_dataframe(mainfile)
        except Exception as exc:  # noqa: BLE001
            logger.error("Failed to load CSV", exc_info=exc)
            return

        logger.info("CSV loaded", rows=len(df))

        event_map = _group_rows(df, logger)
        event_records = []

        for key, info in event_map.items():
            try:
                record, event, event_name = _build_event_record(key, info)
                event_records.append(record)
                _write_child_archive(event, event_name, key, archive, logger)
            except Exception as exc:  # noqa: BLE001
                logger.error(
                    "Failed to build event record",
                    event_key=key,
                    event_name=info.get("event_name", "?"),
                    exc_info=exc,
                )

        archive.data = FAIRmatEventsFile(events=event_records)
        logger.info("GenericEventParser done", total=len(event_records))
