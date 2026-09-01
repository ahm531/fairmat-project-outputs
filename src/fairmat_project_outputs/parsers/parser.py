from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from nomad.datamodel.datamodel import EntryArchive
    from structlog.stdlib import BoundLogger

import re
from datetime import datetime

from nomad.parsing.parser import MatchingParser

# ---------------------------------------------------------------------------
# Regex - title-extraction markers inside the Name cell
# ---------------------------------------------------------------------------

_TITLE_MARKER_RE = re.compile(
    r"\n?\s*(?:talk|presentation|poster|lecture)\s+title\s*:\s*",
    re.IGNORECASE,
)

# ---------------------------------------------------------------------------
# Subtype: strip old column prefixes (S_, Ce_, CM_, E_)
# ---------------------------------------------------------------------------

_PREFIX_RE = re.compile(r"^(?:S|Ce|CM|E)_", re.IGNORECASE)


def _strip_prefix(raw: str) -> str:
    return _PREFIX_RE.sub("", raw).strip()


# ---------------------------------------------------------------------------
# Subtype normalisation map (keys lower-cased after prefix stripping)
# ---------------------------------------------------------------------------

_SUBTYPE_CLEAN_MAP: dict[str, str] = {
    "fair-di - fairmat colloquium series": "FAIR-DI - FAIRmat colloquium series",
    "fairmat seminar series": "FAIRmat seminar series",
    "nfdi physical sciences joint colloquium": "NFDI Physical Sciences Joint Colloquium",
    "invited talks": "Invited talk",
    "invited talk": "Invited talk",
    "plenary talks": "Plenary talk",
    "plenary talk": "Plenary talk",
    "contributed talks": "Contributed talk",
    "contributed talk": "Contributed talk",
    "contributed presentations": "Contributed presentation",
    "contributed presentation": "Contributed presentation",
    "public workshops": "Public workshop",
    "public workshop": "Public workshop",
    "fairmat internal workshops": "FAIRmat internal workshop",
    "fairmat internal workshop": "FAIRmat internal workshop",
    "hackathon": "Hackathon",
    "tech. partners workshop series": "Tech partners workshop series",
    "tech partners workshop series": "Tech partners workshop series",
    "project meeting": "Project meeting",
    "conference session": "Conference session",
    "users meeting": "Users meeting",
    "information booths": "Information booth",
    "information booth": "Information booth",
    "fairmat tutorials series": "FAIRmat tutorials series",
    "external tutorials": "External tutorial",
    "external tutorial": "External tutorial",
    "schools": "School",
    "school": "School",
    "course instance": "Course instance",
    "lecture": "Lecture",
    "lectures": "Lecture",
    "demonstrations": "Demonstration",
    "demonstration": "Demonstration",
    "fair-di conference": "FAIR-DI conference",
}

# ---------------------------------------------------------------------------
# Subtype -> event_type  (path A)
# ---------------------------------------------------------------------------

_SUBTYPE_TO_TYPE: dict[str, str] = {
    "FAIR-DI - FAIRmat colloquium series": "Colloquium",
    "FAIRmat seminar series": "Seminar",
    "NFDI Physical Sciences Joint Colloquium": "Colloquium",
    "Invited talk": "Seminar",
    "Plenary talk": "Conference",
    "Contributed talk": "Conference",
    "Contributed presentation": "Conference",
    "Public workshop": "Workshop",
    "FAIRmat internal workshop": "Workshop",
    "Hackathon": "Hackathon",
    "Tech partners workshop series": "Workshop",
    "Project meeting": "Meeting",
    "Conference session": "Conference",
    "Users meeting": "Meeting",
    "Information booth": "Public event",
    "FAIRmat tutorials series": "Tutorial",
    "External tutorial": "Tutorial",
    "School": "School",
    "Course instance": "Course",
    "Lecture": "Seminar",
    "Demonstration": "Demonstration",
    "FAIR-DI conference": "Conference",
}

# ---------------------------------------------------------------------------
# Event type (main) fallback  (path B)
# ---------------------------------------------------------------------------

_MAIN_TYPE_MAP: dict[str, str] = {
    "s": "Seminar",
    "scholarly talks (s)": "Seminar",
    "scholarly talks": "Seminar",
    "ce": "Workshop",
    "community engagement (ce)": "Workshop",
    "community engagement": "Workshop",
    "cm": "Meeting",
    "conferences and meetings (cm)": "Meeting",
    "conferences and meetings": "Meeting",
    "e": "Tutorial",
    "educational events (e)": "Tutorial",
    "educational events": "Tutorial",
    "pe": "Public event",
    "public engagement (pe)": "Public event",
    "public engagement": "Public event",
}

# ---------------------------------------------------------------------------
# event_name keyword fallback  (path C, ordered most-specific first)
# ---------------------------------------------------------------------------

_NAME_KEYWORDS: list[tuple[str, str]] = [
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

# ---------------------------------------------------------------------------
# FAIRmat role normalisation
# ---------------------------------------------------------------------------

_ROLE_MAP: dict[str, str] = {
    "organizer": "Organizer",
    "co-organizer": "Co-organizer",
    "coorganizer": "Co-organizer",
    "co organizer": "Co-organizer",
    "supporter": "Supporter",
    "support": "Supporter",
    "participant": "Participant",
    # map extended roles to nearest existing value
    "speaker": "Participant",
    "exhibitor": "Participant",
    "host": "Organizer",
}

# ---------------------------------------------------------------------------
# FaBiO terms (case-insensitive)
# ---------------------------------------------------------------------------

_FABIO_TERMS_LIST = [
    "fabio: Abstract", "fabio: Announcement", "fabio: Book chapter",
    "fabio: Computer program", "fabio: Conference proceedings",
    "fabio: Entity metadata", "fabio: Instructional work", "fabio: Interview",
    "fabio: Journal article", "fabio: Meeting report", "fabio: Periodical issue",
    "fabio: Poster", "fabio: Preprint", "fabio: Presentation",
    "fabio: Scholarly work", "fabio: Software dataset", "fabio: Timetable",
    "fabio: Web page",
]
_FABIO_MAP: dict[str, str] = {v.lower(): v for v in _FABIO_TERMS_LIST}
_FABIO_RE = re.compile(r"\[?(fabio:\s*[^\]]+)\]?", re.IGNORECASE)

# ---------------------------------------------------------------------------
# Utilities
# ---------------------------------------------------------------------------


def _clean(value) -> str:
    if value is None:
        return ""
    s = re.sub(r"\s+", " ", str(value).replace("\r", " ").strip())
    return "" if s.lower() in ("nan", "none", "") else s


def _parse_date(value) -> datetime | None:
    s = _clean(value)
    if not s:
        return None
    for fmt in ("%d-%b-%y", "%d-%b-%Y", "%Y-%m-%d", "%d/%m/%Y",
                "%m/%d/%Y", "%d.%m.%Y", "%B %d, %Y"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


def _extract_fabio(raw: str) -> str | None:
    m = _FABIO_RE.search(raw)
    if not m:
        return None
    captured = m.group(1).strip()
    return _FABIO_MAP.get(captured.lower(), captured)


def _infer_event_type_from_name(name: str) -> str | None:
    lower = name.lower()
    for keyword, etype in _NAME_KEYWORDS:
        if keyword in lower:
            return etype
    return None


def _extract_names(raw_name: str) -> tuple[str, str]:
    m = _TITLE_MARKER_RE.search(raw_name)
    if m:
        return raw_name[: m.start()].strip(), raw_name[m.end():].strip()
    return raw_name.strip(), ""


def _load_dataframe(mainfile: str):
    import pandas as pd
    if mainfile.lower().endswith((".xlsx", ".xls")):
        return pd.read_excel(mainfile, header=0, dtype=str)
    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            return pd.read_csv(mainfile, header=0, dtype=str, encoding=enc)
        except UnicodeDecodeError:
            continue
    raise ValueError(f"Cannot decode {mainfile} with any supported encoding")


# ---------------------------------------------------------------------------
# Row parser  (12-step priority order from spec)
# ---------------------------------------------------------------------------


def _parse_rows(df, logger=None):
    df = df.dropna(how="all")
    if df.columns[0].startswith("Unnamed"):
        df = df.iloc[:, 1:]

    col = {c.strip(): c for c in df.columns}

    def get(row, *candidates):
        for name in candidates:
            if name in col:
                return _clean(row.get(col[name], ""))
        return ""

    last_event_name = ""
    last_event_type_raw = ""
    last_sub_raw = ""
    last_role_raw = ""

    for idx, row in df.iterrows():
        review_flags: list[str] = []

        # Step 1 - Row ID
        raw_id = get(row, "ID")
        if raw_id and raw_id.endswith(".0"):
            raw_id = raw_id[:-2]
        row_id = raw_id if raw_id else str(idx)
        if row_id.lower() == "id":
            continue
        is_sub_entry = "." in row_id

        # Step 2 - Name  (inherit for sub-entries)
        name_raw = get(row, "Name")
        if not name_raw:
            if is_sub_entry and last_event_name:
                name_raw = last_event_name
            else:
                continue

        # Step 4 - Extract event_name / contribution_name
        event_name, contribution_name = _extract_names(name_raw)
        if not event_name:
            review_flags.append("Could not extract event name")
            continue
        last_event_name = event_name

        # Step 3 - Dates
        start_date = _parse_date(get(row, "From"))
        end_date = _parse_date(get(row, "To"))

        # Step 5/6 - Clean Event type (sub) for use in event_type inference
        sub_raw = get(row, "Event type (sub)", "Event type (old)") or (
            last_sub_raw if is_sub_entry else ""
        )
        if sub_raw:
            last_sub_raw = sub_raw
        cleaned_sub = _strip_prefix(sub_raw)

        # Step 7 - event_type from subtype  (path A)
        event_type: str | None = None
        canonical_sub = _SUBTYPE_CLEAN_MAP.get(cleaned_sub.lower())
        if canonical_sub:
            event_type = _SUBTYPE_TO_TYPE.get(canonical_sub)

        # Step 8 - event_type from main  (path B)
        main_raw = get(row, "Event type (main)") or (
            last_event_type_raw if is_sub_entry else ""
        )
        if main_raw:
            last_event_type_raw = main_raw
        if event_type is None:
            event_type = _MAIN_TYPE_MAP.get(main_raw.strip().lower())

        # Step 9 - event_type from name keywords  (path C)
        if event_type is None:
            event_type = _infer_event_type_from_name(event_name)
        if event_type is None:
            review_flags.append("Missing event_type")

        # Step 10 - FAIRmat role
        role_raw = get(row, "FAIRmat role") or (last_role_raw if is_sub_entry else "")
        if role_raw:
            last_role_raw = role_raw
        fairmat_role = _ROLE_MAP.get(role_raw.strip().lower())
        if role_raw and fairmat_role is None:
            fairmat_role = role_raw.strip()
            review_flags.append(f"Unknown FAIRmat role: {role_raw!r}")

        # Step 11 - FaBiO
        fabio_raw = get(row, "FaBio Outpu1", "FaBio Output1", "FaBiO Output 1")
        fabio_type = _extract_fabio(fabio_raw) if fabio_raw else None
        output_url = get(row, "Link to relevant output 1") or None
        event_url = get(row, "Event URL", "URL", "Link to event") or None

        participants_raw = get(row, "Speaker/Participant", "Speaker/Participants")
        participants = [
            p.strip() for p in re.split(r"[;,]", participants_raw) if p.strip()
        ]

        # Step 12 - Log review flags
        if review_flags and logger:
            logger.warning(
                "Row needs review",
                row_id=row_id,
                event_name=event_name,
                flags=review_flags,
            )

        yield {
            "row_id": row_id,
            "event_name": event_name or None,
            "contribution_name": contribution_name or None,
            "start_date": start_date,
            "end_date": end_date,
            "event_type": event_type,
            "fairmat_role": fairmat_role,
            "participants": participants or None,
            "event_url": event_url,
            "output_url": output_url,
            "fabio_type": fabio_type,
            "review_flags": review_flags,
        }


# ---------------------------------------------------------------------------
# Parser
# ---------------------------------------------------------------------------


class FAIRmatEventsParser(MatchingParser):
    """
    Parses an Excel (.xlsx/.xls) or CSV file of FAIRmat event records.

    For every data row it:
      1. Adds an EventRecord sub-section to a FAIRmatEventsFile in archive.data.
      2. Writes a .archive.yaml child entry as an individual FAIRmatEvent.
    """

    def parse(
        self,
        mainfile: str,
        archive: "EntryArchive",
        logger: "BoundLogger",
        child_archives: dict[str, "EntryArchive"] = None,
    ) -> None:
        from nomad.datamodel import EntryArchive as EA
        from nomad.datamodel import EntryMetadata

        from fairmat_project_outputs.schema_packages.schema_package import (
            Contribution,
            EventRecord,
            FabioEventOutput,
            FAIRmatEvent,
            FAIRmatEventsFile,
        )

        logger.info("FAIRmatEventsParser.parse", mainfile=mainfile)

        try:
            df = _load_dataframe(mainfile)
        except Exception as exc:
            logger.error("Failed to load file", exc_info=exc)
            return

        row_dicts = list(_parse_rows(df, logger=logger))
        logger.info("rows parsed", count=len(row_dicts))

        event_records = []
        import yaml as _yaml

        for data in row_dicts:
            try:
                fabio_outputs = []
                if data["fabio_type"]:
                    fabio_outputs = [
                        FabioEventOutput(
                            fabio_type=data["fabio_type"],
                            url=data["output_url"],
                        )
                    ]

                contributions = []
                if data["contribution_name"] or data["participants"] or fabio_outputs:
                    contributions = [
                        Contribution(
                            title=data["contribution_name"],
                            contributors=data["participants"],
                            fabio_outputs=fabio_outputs,
                        )
                    ]

                record = EventRecord(
                    event_name=data["event_name"],
                    start_date=data["start_date"],
                    end_date=data["end_date"],
                    event_type=data["event_type"],
                    fairmat_role=data["fairmat_role"],
                    event_url=data["event_url"],
                    contributions=contributions,
                )
                event_records.append(record)

                event = FAIRmatEvent(
                    event_name=data["event_name"],
                    start_date=data["start_date"],
                    end_date=data["end_date"],
                    event_type=data["event_type"],
                    fairmat_role=data["fairmat_role"],
                    event_url=data["event_url"],
                    contributions=contributions,
                )

                safe_name = re.sub(r"[^\w\-]", "_", data["event_name"] or "event")[:60]
                safe_id = re.sub(r"[^\w]", "_", data["row_id"])
                filename = f"event_{safe_id}_{safe_name}.archive.yaml"
                entry_name = data["event_name"] or f"FAIRmat Event {data['row_id']}"

                child_archive = EA(
                    data=event,
                    m_context=archive.m_context,
                    metadata=EntryMetadata(
                        upload_id=archive.m_context.upload_id,
                        entry_name=entry_name,
                    ),
                )

                try:
                    with archive.m_context.raw_file(filename, "w") as outfile:
                        _yaml.dump(child_archive.m_to_dict(), outfile)
                    archive.m_context.upload.process_updated_raw_file(
                        filename, allow_modify=True
                    )
                    logger.info(
                        "created child entry",
                        row_id=data["row_id"],
                        filename=filename,
                    )
                except Exception as exc:
                    logger.error(
                        "Failed to create child entry",
                        row_id=data["row_id"],
                        exc_info=exc,
                    )

            except Exception as exc:
                logger.error(
                    "Skipping row due to error",
                    row_id=data.get("row_id", "?"),
                    event_name=data.get("event_name", "?"),
                    exc_info=exc,
                )

        archive.data = FAIRmatEventsFile(events=event_records)
        logger.info("FAIRmatEventsParser done", total=len(event_records))
