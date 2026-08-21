"""
Parser for DPG event CSV files named ``dpg_event.csv``.

CSV layout (no header row, 3 or 4 columns):
  - Column 0: start date
  - Column 1: end date
  - Column 2: raw text that contains both the event name and the contribution
              title.  The contribution title follows one of the recognised
              keywords (see ``_DPG_KEYWORD_RE``).
  - Column 3 (optional): speaker / presenter name

Multiple rows that share the same (start_date, end_date) are treated as
contributions to the same event.
"""

from __future__ import annotations

import csv
import re
from collections import defaultdict
from datetime import datetime
from typing import TYPE_CHECKING

import yaml
from nomad.parsing.parser import MatchingParser

if TYPE_CHECKING:
    from nomad.datamodel.datamodel import EntryArchive
    from structlog.stdlib import BoundLogger

_MIN_CSV_COLUMNS = 3
_SPEAKER_COL = 3
_EXTRA_START_COL = 4  # columns from here onward are scanned for FaBiO term / URL

# Contribution type → FaBiO term (used when no explicit FaBiO column value)
_CTYPE_TO_FABIO: dict[str, str] = {
    "Contributed talk": "fabio: Presentation",
    "Invited talk": "fabio: Presentation",
    "Plenary talk": "fabio: Presentation",
    "Poster": "fabio: Poster",
    "Tutorial": "fabio: Instructional work",
    "Information booth": "fabio: Announcement",
    "Panel discussion": "fabio: Meeting report",
    "Workshop session": "fabio: Meeting report",
    "Conference session": "fabio: Meeting report",
    "Symposium": "fabio: Conference proceedings",
}

# Valid FaBiO enum values (lower-cased for case-insensitive matching)
_VALID_FABIO: set[str] = {
    "fabio: abstract", "fabio: announcement", "fabio: book chapter",
    "fabio: computer program", "fabio: conference proceedings",
    "fabio: entity metadata", "fabio: instructional work", "fabio: interview",
    "fabio: journal article", "fabio: meeting report", "fabio: periodical issue",
    "fabio: poster", "fabio: preprint", "fabio: presentation",
    "fabio: scholarly work", "fabio: software dataset", "fabio: timetable",
    "fabio: web page",
}

# ---------------------------------------------------------------------------
# Keyword regex – ordered most-specific first
# ---------------------------------------------------------------------------

_DPG_KEYWORD_RE = re.compile(
    r"(?<![\w])(?:"  # not preceded by a word char (boundary without \b)
    r"plenary\s+talk\s+discussion"
    r"|focus(?:e)?\s+session\s*:"
    r"|information\s+booth"
    r"|hacky\s+hour"
    r"|talk\s+title"
    r"|exhibition"
    r"|symposium"
    r"|tutorial"
    r"|talk"
    r")\s*:?\s*",
    re.IGNORECASE,
)

# Keyword (lower-cased match group) → CONTRIBUTION_TYPE value
_KEYWORD_TO_CONTRIBUTION_TYPE: dict[str, str] = {
    "plenary talk discussion": "Panel discussion",
    "focuse session": "Conference session",
    "focus session": "Conference session",
    "information booth": "Information booth",
    "hacky hour": "Workshop session",
    "talk title": "Contributed talk",
    "exhibition": "Information booth",
    "symposium": "Symposium",
    "tutorial": "Tutorial",
    "talk": "Contributed talk",
}

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
# Event name normalisation
# ---------------------------------------------------------------------------

def _normalise_event_name(raw: str, year: int | None) -> str:
    """Replace SKM (standalone or embedded like DPG_SKM21) with the full DPG name."""
    if re.search(r'SKM', raw, re.IGNORECASE):
        yr = str(year) if year else ""
        return f"DPG Spring Meeting of the Condensed Matter Section {yr}".strip()
    return raw.strip()


# ---------------------------------------------------------------------------
# Text splitting helpers
# ---------------------------------------------------------------------------

def _keyword_ctype(keyword_text: str) -> str:
    """Return CONTRIBUTION_TYPE for the raw keyword match group."""
    normalised = keyword_text.strip().rstrip(":").strip().lower()
    for key, val in _KEYWORD_TO_CONTRIBUTION_TYPE.items():
        if normalised.startswith(key) or key.startswith(normalised):
            return val
    return ""


def _split_text(raw_text: str) -> tuple[str, str, str]:
    """
    Return ``(event_name_part, contribution_title, contribution_type)``
    extracted from *raw_text*.

    When no text follows the keyword (e.g. "... information booth" with nothing
    after it) the contribution type itself is used as the title.
    """
    m = _DPG_KEYWORD_RE.search(raw_text)
    if not m:
        return raw_text.strip(), "", ""
    event_part = raw_text[: m.start()].strip()
    contrib_title = raw_text[m.end():].strip()
    ctype = _keyword_ctype(m.group(0))
    # If the keyword was the last thing in the string, use the type as the title
    if not contrib_title:
        contrib_title = ctype
    return event_part, contrib_title, ctype


# ---------------------------------------------------------------------------
# CSV loader
# ---------------------------------------------------------------------------

def _load_csv(mainfile: str) -> list[list[str]]:
    """Load a headerless CSV. Tries multiple encodings."""
    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            with open(mainfile, newline="", encoding=enc) as fh:
                return [row for row in csv.reader(fh) if any(c.strip() for c in row)]
        except UnicodeDecodeError:
            continue
    raise ValueError(f"Cannot decode {mainfile} with any supported encoding")


def _extract_fabio_and_url(row: list[str]) -> tuple[str, str]:
    """
    Scan all extra columns (index >= _EXTRA_START_COL) and return
    ``(fabio_term, url)`` where each may be an empty string if not found.
    A cell is a FaBiO term if it matches _VALID_FABIO; it is a URL if it
    starts with 'http'.
    """
    fabio_term = ""
    url = ""
    for cell in row[_EXTRA_START_COL:]:
        val = str(cell).strip()
        if not val or val.lower() in ("nan", "none"):
            continue
        if not fabio_term and val.lower() in _VALID_FABIO:
            fabio_term = val
        elif not url and val.lower().startswith("http"):
            url = val
    return fabio_term, url



def _normalise_event_key(raw_event_part: str) -> str:
    """
    Return a canonical lower-cased key for an event-name fragment.

    All SKM variants ('DPG_SKM21', 'DPG SKM 2021', 'SKM 22', …) are
    collapsed by extracting the trailing year digits and producing a
    stable key ``skm:<year>``.  Non-SKM names are lowercased and stripped.
    """
    if re.search(r'SKM', raw_event_part, re.IGNORECASE):
        year_m = re.search(r'(\d{2,4})', raw_event_part)
        year_str = year_m.group(1) if year_m else ""
        if len(year_str) == 2:
            year_str = "20" + year_str
        return f"skm:{year_str}"
    return re.sub(r'\s+', ' ', raw_event_part.strip().lower())


def _group_rows(
    rows: list[list[str]], logger: BoundLogger | None
) -> dict[str, dict]:
    """
    Group CSV rows by **normalised event name** (the text before the first
    keyword in column 2).

    Using the event name rather than the date pair means that contributions
    belonging to the same conference but listed on different days of that
    conference (each row having its own per-day date) are still merged into
    one event entry.
    """
    event_map: dict[str, dict] = defaultdict(
        lambda: {"start": None, "end": None, "entries": [], "raw_event_name": ""}
    )
    for row_idx, row in enumerate(rows):
        if len(row) < _MIN_CSV_COLUMNS:
            if logger:
                logger.warning("Skipping short row", row_idx=row_idx)
            continue
        raw_text = str(row[2]).strip()
        if not raw_text:
            continue
        start_dt = _parse_date(row[0])
        end_dt = _parse_date(row[1])
        # Extract the event-name portion (text before any keyword).
        raw_event_part, _, _ = _split_text(raw_text)
        key: str = _normalise_event_key(raw_event_part)
        speaker = str(row[_SPEAKER_COL]).strip() if len(row) > _SPEAKER_COL else ""
        fabio_raw, output_url = _extract_fabio_and_url(row)
        info = event_map[key]
        if info["start"] is None:
            info["start"] = start_dt
            info["end"] = end_dt
            info["raw_event_name"] = raw_event_part
        else:
            # Expand date range to cover all contribution rows
            if start_dt and (info["start"] is None or start_dt < info["start"]):
                info["start"] = start_dt
            if end_dt and (info["end"] is None or end_dt > info["end"]):
                info["end"] = end_dt
        info["entries"].append((raw_text, speaker, fabio_raw, output_url))
    return dict(event_map)


# ---------------------------------------------------------------------------
# Contribution builder
# ---------------------------------------------------------------------------

def _build_contributions(entries: list[tuple[str, str, str, str]], start_dt, end_dt) -> list:
    from fairmat_project_outputs.schema_packages.schema_package import Contribution

    contributions = []
    for raw_text, speaker, fabio_raw, output_url in entries:
        _, contrib_title, ctype = _split_text(raw_text)
        if contrib_title:
            # Split comma-separated names; treat all as FAIRmat contributors for now
            # (to be manually curated later)
            other_names = [n.strip() for n in speaker.split(',') if n.strip()] if speaker else []
            # FaBiO: column-extracted term already validated; fall back to inferred term
            fabio_term = fabio_raw if fabio_raw else _CTYPE_TO_FABIO.get(ctype, "")
            contributions.append(
                Contribution(
                    title=contrib_title,
                    contribution_type=ctype or None,
                    fairmat_contributors=other_names or None,
                    start_date=start_dt,
                    end_date=end_dt,
                    fabio_type=fabio_term or None,
                    fabio_url=output_url or None,
                )
            )
    return contributions


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

    safe_name = re.sub(r"[^\w\-]", "_", event_name or "dpg_event")[:60]
    safe_key = re.sub(r"[^\w]", "_", key)
    filename = f"dpg_event_{safe_key}_{safe_name}.archive.yaml"

    child_archive = EA(
        data=event,
        m_context=archive.m_context,
        metadata=EntryMetadata(
            upload_id=archive.m_context.upload_id,
            entry_name=event_name or "DPG Event",
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
# Event builder
# ---------------------------------------------------------------------------

def _build_event_record(key: str, info: dict):
    from fairmat_project_outputs.schema_packages.schema_package import (
        EventRecord,
        FAIRmatEvent,
    )

    start_dt: datetime | None = info["start"]
    end_dt: datetime | None = info["end"]
    entries: list[str] = info["entries"]

    year = start_dt.year if start_dt else None
    # Use the stored raw_event_name from the first row of this group
    raw_event_part = info.get("raw_event_name") or _split_text(entries[0][0])[0]
    event_name = _normalise_event_name(raw_event_part, year)
    contributions = _build_contributions(entries, start_dt, end_dt)

    common = {
        "event_name": event_name,
        "start_date": start_dt,
        "end_date": end_dt,
        "event_type": "Conference",
        "country": "Germany",
        "mode": "Onsite",
        "contributions": contributions,
    }
    return EventRecord(**common), FAIRmatEvent(**common), event_name


# ---------------------------------------------------------------------------
# Parser
# ---------------------------------------------------------------------------

class DPGEventParser(MatchingParser):
    """
    Parses a ``dpg_event.csv`` file into ``FAIRmatEvent`` entries.

    One ``FAIRmatEvent`` is created per unique (start_date, end_date) pair.
    All rows sharing that pair become ``Contribution`` sub-sections.
    """

    def parse(
        self,
        mainfile: str,
        archive: EntryArchive,
        logger: BoundLogger,
        child_archives: dict[str, EntryArchive] = None,
    ) -> None:
        from fairmat_project_outputs.schema_packages.schema_package import FAIRmatEventsFile

        logger.info("DPGEventParser.parse", mainfile=mainfile)

        try:
            rows = _load_csv(mainfile)
        except Exception as exc:  # noqa: BLE001
            logger.error("Failed to load DPG CSV", exc_info=exc)
            return

        event_map = _group_rows(rows, logger)
        event_records = []

        for key, info in event_map.items():
            if not info["entries"]:
                continue
            record, event, event_name = _build_event_record(key, info)
            event_records.append(record)
            _write_child_archive(event, event_name, key, archive, logger)

        archive.data = FAIRmatEventsFile(events=event_records)
        logger.info("DPGEventParser done", total=len(event_records))
