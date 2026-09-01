from nomad.config.models.plugins import ParserEntryPoint


class FAIRmatEventsParserEntryPoint(ParserEntryPoint):
    def load(self):
        from fairmat_project_outputs.parsers.parser import FAIRmatEventsParser

        return FAIRmatEventsParser(**self.model_dump())


class DPGEventParserEntryPoint(ParserEntryPoint):
    def load(self):
        from fairmat_project_outputs.parsers.dpg_event_parser import DPGEventParser

        return DPGEventParser(**self.model_dump())


parser_entry_point = FAIRmatEventsParserEntryPoint(
    name='FAIRmatEventsParser',
    description='Parses Excel/CSV files containing FAIRmat event records into FAIRmatEvent entries.',
    mainfile_name_re=r'.*[Ee]vents.*\.(xlsx|xls|csv)$',
    mainfile_mime_re='(application/vnd\.openxmlformats-officedocument\.spreadsheetml\.sheet|application/vnd\.ms-excel|text/.*)',
)

dpg_event_parser_entry_point = DPGEventParserEntryPoint(
    name='DPGEventParser',
    description='Parses dpg_event.csv files (DPG conference contributions) into FAIRmatEvent entries.',
    mainfile_name_re=r'.*dpg_event\.csv$',
    mainfile_mime_re='text/.*',
)

users_meetings_parser_entry_point = DPGEventParserEntryPoint(
    name='UsersMeetingsParser',
    description='Parses Users_Meetings.csv files into FAIRmatEvent entries.',
    mainfile_name_re=r'.*Users_Meetings\.csv$',
    mainfile_mime_re='text/.*',
)


class GenericEventParserEntryPoint(ParserEntryPoint):
    def load(self):
        from fairmat_project_outputs.parsers.generic_event_parser import (
            GenericEventParser,
        )

        return GenericEventParser(**self.model_dump())


generic_event_parser_entry_point = GenericEventParserEntryPoint(
    name='GenericEventParser',
    description='Parses any *_events.csv files into FAIRmatEvent entries.',
    mainfile_name_re=r'.*_events\.csv$',
    mainfile_mime_re='text/.*',
)
