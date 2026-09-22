from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from nomad.datamodel.datamodel import EntryArchive
    from structlog.stdlib import BoundLogger

from nomad.config import config
from nomad.datamodel.data import ArchiveSection, Schema, UseCaseElnCategory
from nomad.datamodel.metainfo.annotations import ELNAnnotation, ELNComponentEnum
from nomad.metainfo import (
    Datetime,
    MEnum,
    Quantity,
    SchemaPackage,
    Section,
    SubSection,
)

configuration = config.get_plugin_entry_point(
    'fairmat_project_outputs.schema_packages:schema_package_entry_point'
)

m_package = SchemaPackage()

# ---------------------------------------------------------------------------
# Controlled vocabularies
# ---------------------------------------------------------------------------

# Every term below is spelled exactly as the ontology spells it: the prefix
# names the ontology, the rest is the class's own rdfs:label. FaBiO labels are
# lowercase throughout, checked against the current release of
# http://purl.org/spar/fabio (254 classes).
#
# Three terms this list used to carry are not FaBiO classes at all, so they
# could never resolve to a URI and were counted as nothing:
#   'Poster'           -> fabio: conference poster
#   'Software dataset' -> fabio: dataset
#   'Interview'        -> bibo: Interview, the only published class for the
#                         genre. FaBiO has no interview; its nearest term,
#                         'movie', names the medium and loses the meaning.
#                         BIBO is FaBiO's usual companion -- FaBiO itself uses
#                         bibo:doi and bibo:status -- so the mix is normal, and
#                         the prefix keeps it honest about which ontology owns
#                         the term. Note BIBO capitalises its labels.
FABIO_TERMS = MEnum(
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
    'bibo: Interview',
)

EVENT_TYPE = MEnum(
    'Conference',
    'Workshop',
    'School',
    'Meeting',
    'Seminar',
    'Tutorial',
    'Hackathon',
    'Symposium',
    'Colloquium',
    'Course',
    'Demonstration',
    'Public event',
)

EVENT_SERIES = MEnum(
    'FAIR-DI / FAIRmat colloquium series',
    'FAIRmat seminar series',
    'NFDI Physical Sciences Joint Colloquium',
    'NFDI Physical Sciences Workshop',
    'Tech. partners workshop series',
    'FAIRmat tutorials series',
    'FAIRmat project meeting',
    'FAIRmat users meeting',
)

EVENT_INTERACTION_APPROACH = MEnum(
    'Community engagement',
    'Community activation',
    'Interconnectivity',
    'Public engagement',
    'Internal',
)

EVENT_PURPOSE = MEnum(
    'Educational event',
    'Scholarly event',
    'Informational event'
)

EVENT_MODE = MEnum(
    'Onsite',
    'Online',
    'Hybrid',
)

FAIRMAT_ROLE = MEnum(
    'Organizer',
    'Co-organizer',
    'Supporter',
    'Participant',
)

CONTRIBUTION_TYPE = MEnum(
    'Invited talk',
    'Plenary talk',
    'Contributed talk',
    'Poster',
    'Tutorial',
    'Workshop session',
    'Information booth',
    'Panel discussion',
    'Symposium',
    'Conference session',
)

CONTRIBUTION_ROLE = MEnum(
    'Organization',
    'Co-organization',
    'Participation',
    'Moderation',
)

# ---------------------------------------------------------------------------
# Sub-sections
# ---------------------------------------------------------------------------


class Contribution(ArchiveSection):
    """A single contribution to an event (talk, poster, session, booth, etc.)."""

    m_def = Section(label_quantity='title')

    title = Quantity(
        type=str,
        label='Title',
        description=(
            'Title of the contribution, e.g. a talk title, session name, or booth name.'
        ),
        a_eln=ELNAnnotation(component=ELNComponentEnum.StringEditQuantity),
    )

    start_date = Quantity(
        type=Datetime,
        label='Start date',
        description='Start date of this contribution.',
        a_eln=ELNAnnotation(component=ELNComponentEnum.DateEditQuantity),
    )

    end_date = Quantity(
        type=Datetime,
        label='End date',
        description='End date of this contribution.',
        a_eln=ELNAnnotation(component=ELNComponentEnum.DateEditQuantity),
    )

    contribution_type = Quantity(
        type=CONTRIBUTION_TYPE,
        label='Contribution type',
        description='Format/type of the contribution (talk, poster, session, etc.).',
        a_eln=ELNAnnotation(component=ELNComponentEnum.EnumEditQuantity),
    )

    contribution_role = Quantity(
        type=CONTRIBUTION_ROLE,
        label='Contribution role',
        description='Role of FAIRmat in this contribution (organization, participation, etc.).',
        a_eln=ELNAnnotation(component=ELNComponentEnum.EnumEditQuantity),
    )

    fairmat_contributors = Quantity(
        type=str,
        shape=['*'],
        label='FAIRmat speaker / contributor',
        description='FAIRmat team members presenting or contributing. One name per entry.',
        a_eln=ELNAnnotation(component=ELNComponentEnum.StringEditQuantity),
    )

    other_contributors = Quantity(
        type=str,
        shape=['*'],
        label='Other speaker / contributor',
        description='Non-FAIRmat speakers or contributors. One name per entry.',
        a_eln=ELNAnnotation(component=ELNComponentEnum.StringEditQuantity),
    )

    fabio_type = Quantity(
        type=FABIO_TERMS,
        label='FaBiO type',
        description='FaBiO ontology class for this contribution output.',
        a_eln=ELNAnnotation(component=ELNComponentEnum.EnumEditQuantity),
    )

    fabio_url = Quantity(
        type=str,
        label='FaBiO URL',
        description='Link to the output resource.',
        a_eln=ELNAnnotation(component=ELNComponentEnum.URLEditQuantity),
    )


# ---------------------------------------------------------------------------
# Schema 1 – FAIRmat Event
# ---------------------------------------------------------------------------


class FAIRmatEvent(Schema):
    """Schema for events that FAIRmat organized, co-organized, supported, or participated in."""

    m_def = Section(
        label='FAIRmat Event',
        categories=[UseCaseElnCategory],
        a_eln={
            'hide': ['lab_id'],
        },
    )

    # -- Identity -------------------------------------------------------------

    event_name = Quantity(
        type=str,
        label='Event title',
        description='Official name or title of the event.',
        a_eln=ELNAnnotation(component=ELNComponentEnum.StringEditQuantity),
    )

    event_url = Quantity(
        type=str,
        label='Event URL',
        description='URL of the event page or invitation.',
        default='https://',
        a_eln=ELNAnnotation(component=ELNComponentEnum.URLEditQuantity),
    )

    start_date = Quantity(
        type=Datetime,
        label='Start date',
        description='Start date of the event.',
        a_eln=ELNAnnotation(component=ELNComponentEnum.DateEditQuantity),
    )

    end_date = Quantity(
        type=Datetime,
        label='End date',
        description='End date of the event.',
        a_eln=ELNAnnotation(component=ELNComponentEnum.DateEditQuantity),
    )

    # -- Classification -------------------------------------------------------

    event_type = Quantity(
        type=EVENT_TYPE,
        label='Event type',
        description='High-level format of the event.',
        a_eln=ELNAnnotation(component=ELNComponentEnum.EnumEditQuantity),
    )

    event_series = Quantity(
        type=EVENT_SERIES,
        label='Event series',
        description='Named series this event belongs to (optional).',
        a_eln=ELNAnnotation(component=ELNComponentEnum.AutocompleteEditQuantity),
    )

    event_interaction_approach = Quantity(
        type=EVENT_INTERACTION_APPROACH,
        label='Interaction approach',
        description='How FAIRmat interacted with the community through this event.',
        a_eln=ELNAnnotation(component=ELNComponentEnum.EnumEditQuantity),
    )

    event_purpose = Quantity(
        type=EVENT_PURPOSE,
        shape=['*'],
        label='Event purpose',
        description='Purpose(s) of the event. Multiple values can be selected.',
        a_eln=ELNAnnotation(component=ELNComponentEnum.AutocompleteEditQuantity),
    )

    # -- Location & mode ------------------------------------------------------

    location_name = Quantity(
        type=str,
        label='Location name',
        description='Name of the venue or institution.',
        a_eln=ELNAnnotation(component=ELNComponentEnum.StringEditQuantity),
    )

    city = Quantity(
        type=str,
        label='City',
        description='City where the event takes place.',
        a_eln=ELNAnnotation(component=ELNComponentEnum.StringEditQuantity),
    )

    country = Quantity(
        type=str,
        label='Country',
        description='Country where the event takes place.',
        a_eln=ELNAnnotation(component=ELNComponentEnum.StringEditQuantity),
    )

    mode = Quantity(
        type=EVENT_MODE,
        label='Mode',
        description='Attendance mode: onsite, online, or hybrid.',
        a_eln=ELNAnnotation(component=ELNComponentEnum.EnumEditQuantity),
    )

    # -- FAIRmat involvement --------------------------------------------------

    fairmat_role = Quantity(
        type=FAIRMAT_ROLE,
        label='FAIRmat primary role',
        description='Primary role of FAIRmat in this event.',
        a_eln=ELNAnnotation(component=ELNComponentEnum.EnumEditQuantity),
    )

    fairmat_contributors = Quantity(
        type=str,
        shape=['*'],
        label='FAIRmat contributors',
        description='FAIRmat team members involved in this event. Add one name per line.',
        a_eln=ELNAnnotation(component=ELNComponentEnum.StringEditQuantity),
    )

    # -- Contributions (repeatable) -------------------------------------------

    contributions_overview = Quantity(
        type=str,
        label='Contributions overview',
        description='Auto-populated bullet list of contributions at this event.',
        a_eln=ELNAnnotation(component=ELNComponentEnum.RichTextEditQuantity),
    )

    contributions = SubSection(
        section_def=Contribution,
        label='Contributions',
        description='Individual contributions at this event (talks, posters, sessions, etc.).',
        repeats=True,
    )

    def normalize(self, archive: 'EntryArchive', logger: 'BoundLogger') -> None:
        super().normalize(archive, logger)
        fairmat_names: list = []
        rows: list = []
        for contrib in (self.contributions or []):
            for name in (getattr(contrib, 'fairmat_contributors', None) or []):
                if name and name not in fairmat_names:
                    fairmat_names.append(name)
            title = getattr(contrib, 'title', None) or ''
            ctype = getattr(contrib, 'contribution_type', None) or ''
            all_speakers = (
                list(getattr(contrib, 'fairmat_contributors', None) or []) +
                list(getattr(contrib, 'other_contributors', None) or [])
            )
            speakers = ', '.join(all_speakers)
            if title:
                fabio_terms = (getattr(contrib, 'fabio_type', None) or '').strip()
                rows.append((ctype, title, speakers, fabio_terms))
        if fairmat_names:
            self.fairmat_contributors = fairmat_names
        if rows:
            header = (
                '<table style="width:100%;border-collapse:collapse">'
                '<thead><tr>'
                '<th style="border:1px solid #ccc;padding:4px 8px">#</th>'
                '<th style="border:1px solid #ccc;padding:4px 8px">Type</th>'
                '<th style="border:1px solid #ccc;padding:4px 8px">Title</th>'
                '<th style="border:1px solid #ccc;padding:4px 8px">Presenters</th>'
                '<th style="border:1px solid #ccc;padding:4px 8px">FaBiO class</th>'
                '</tr></thead><tbody>'
            )
            body = ''
            for idx, (ctype, title, speakers, fabio_terms) in enumerate(rows, start=1):
                body += (
                    f'<tr>'
                    f'<td style="border:1px solid #ccc;padding:4px 8px;text-align:center">{idx}</td>'
                    f'<td style="border:1px solid #ccc;padding:4px 8px"><strong>{ctype}</strong></td>'
                    f'<td style="border:1px solid #ccc;padding:4px 8px">{title}</td>'
                    f'<td style="border:1px solid #ccc;padding:4px 8px">{speakers}</td>'
                    f'<td style="border:1px solid #ccc;padding:4px 8px">{fabio_terms}</td>'
                    f'</tr>'
                )
            self.contributions_overview = header + body + '</tbody></table>'


class EventRecord(ArchiveSection):
    """A single event record used inside FAIRmatEventsFile (parsed from a spreadsheet)."""

    m_def = Section(label_quantity='event_name')

    event_name = Quantity(type=str, label='Event title')
    start_date = Quantity(type=Datetime, label='Start date')
    end_date = Quantity(type=Datetime, label='End date')
    event_type = Quantity(type=EVENT_TYPE, label='Event type')
    event_series = Quantity(type=EVENT_SERIES, label='Event series')
    event_interaction_approach = Quantity(
        type=EVENT_INTERACTION_APPROACH, label='Interaction approach'
    )
    event_purpose = Quantity(type=EVENT_PURPOSE, shape=['*'], label='Event purpose')
    location_name = Quantity(type=str, label='Location name')
    city = Quantity(type=str, label='City')
    country = Quantity(type=str, label='Country')
    mode = Quantity(type=EVENT_MODE, label='Mode')
    fairmat_role = Quantity(type=FAIRMAT_ROLE, label='FAIRmat primary role')
    fairmat_contributors = Quantity(type=str, shape=['*'], label='FAIRmat contributors')
    event_url = Quantity(type=str, label='Event URL', default='https://')

    contributions = SubSection(
        section_def=Contribution,
        label='Contributions',
        repeats=True,
    )


class FAIRmatEventsFile(Schema):
    """Container schema produced by the events spreadsheet parser."""

    m_def = Section(
        label='FAIRmat Events File',
        categories=[UseCaseElnCategory],
    )

    events = SubSection(
        section_def=EventRecord,
        label='Events',
        repeats=True,
    )

    def normalize(self, archive: 'EntryArchive', logger: 'BoundLogger') -> None:
        super().normalize(archive, logger)


# ---------------------------------------------------------------------------
# Schema 2 – FAIRmat Project Output
# ---------------------------------------------------------------------------


class FAIRmatOutput(Schema):
    """Schema for FAIRmat project outputs such as publications, software, and presentations."""

    m_def = Section(
        label='FAIRmat Output',
        categories=[UseCaseElnCategory],
        a_eln={
            'hide': ['lab_id'],
        },
    )

    title = Quantity(
        type=str,
        label='Title',
        description='Title of the output.',
        a_eln=ELNAnnotation(component=ELNComponentEnum.StringEditQuantity),
    )

    authors = Quantity(
        type=str,
        shape=['*'],
        label='Authors',
        description='List of authors. Add one name per line.',
        a_eln=ELNAnnotation(component=ELNComponentEnum.StringEditQuantity),
    )

    year = Quantity(
        type=int,
        label='Year',
        description='Year of publication or release.',
        a_eln=ELNAnnotation(component=ELNComponentEnum.NumberEditQuantity),
    )

    link_or_doi = Quantity(
        type=str,
        label='Link / DOI',
        description='URL or DOI pointing to the output.',
        a_eln=ELNAnnotation(component=ELNComponentEnum.URLEditQuantity),
    )

    fabio_class = Quantity(
        type=FABIO_TERMS,
        label='FaBiO class',
        description='FaBiO ontology class describing the type of output.',
        a_eln=ELNAnnotation(component=ELNComponentEnum.EnumEditQuantity),
    )

    is_open_access = Quantity(
        type=bool,
        label='Open access',
        description='Check if this output is openly accessible.',
        a_eln=ELNAnnotation(component=ELNComponentEnum.BoolEditQuantity),
    )

    def normalize(self, archive: 'EntryArchive', logger: 'BoundLogger') -> None:
        super().normalize(archive, logger)


m_package.__init_metainfo__()
