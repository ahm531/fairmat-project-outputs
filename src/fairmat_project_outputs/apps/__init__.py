from nomad.config.models.plugins import AppEntryPoint
from nomad.config.models.ui import (
    App,
    Axis,
    Column,
    Dashboard,
    Layout,
    Menu,
    MenuItemHistogram,
    MenuItemTerms,
    Pagination,
    Rows,
    SearchQuantities,
    WidgetHistogram,
    WidgetTerms,
)

SCHEMA = 'fairmat_project_outputs.schema_packages.schema_package.FAIRmatEvent'

# ---------------------------------------------------------------------------
# Search quantity paths
# ---------------------------------------------------------------------------
Q_EVENT_NAME = f'data.event_name#{SCHEMA}'
Q_START_DATE = f'data.start_date#{SCHEMA}'
Q_END_DATE = f'data.end_date#{SCHEMA}'
Q_EVENT_TYPE = f'data.event_type#{SCHEMA}'
Q_FAIRMAT_ROLE = f'data.fairmat_role#{SCHEMA}'

# ---------------------------------------------------------------------------
# App definition
# ---------------------------------------------------------------------------
fairmat_events_app = App(
    label='FAIRmat Events',
    path='fairmat-events',
    category='FAIRmat',
    description='Browse and filter FAIRmat events: talks, workshops, schools, conferences and more.',
    filters_locked={
        'section_defs.definition_qualified_name': [SCHEMA],
    },
    pagination=Pagination(
        order_by='upload_create_time',
        order='desc',
        page_size=20,
    ),
    search_quantities=SearchQuantities(
        include=[
            Q_EVENT_NAME,
            Q_START_DATE,
            Q_END_DATE,
            Q_EVENT_TYPE,
            Q_FAIRMAT_ROLE,

        ]
    ),
    columns=[
        Column(search_quantity=Q_START_DATE, title='Start date', selected=True,
               format={'mode': 'date', 'decimals': 0}),
        Column(search_quantity=Q_END_DATE, title='End date', selected=True,
               format={'mode': 'date', 'decimals': 0}),
        Column(search_quantity=Q_EVENT_NAME, title='Event name', selected=True),
        Column(search_quantity=Q_EVENT_TYPE, title='Type', selected=True),
        Column(search_quantity=Q_FAIRMAT_ROLE, title='FAIRmat role', selected=False),
    ],
    rows=Rows(),
    menu=Menu(
        title='Filters',
        items=[
            Menu(
                title='Event name',
                items=[
                    MenuItemTerms(
                        search_quantity=Q_EVENT_NAME,
                        title='Search event name',
                        show_input=True,
                        options=20,
                    ),
                ],
            ),
            Menu(
                title='Event type',
                items=[
                    MenuItemTerms(
                        search_quantity=Q_EVENT_TYPE,
                        title='Type',
                        show_input=False,
                        options=10,
                    ),
                ],
            ),
            Menu(
                title='FAIRmat role',
                items=[
                    MenuItemTerms(
                        search_quantity=Q_FAIRMAT_ROLE,
                        title='Role',
                        show_input=False,
                        options=10,
                    ),
                ],
            ),
            Menu(
                title='Date range',
                items=[
                    MenuItemHistogram(
                        x=Axis(search_quantity=Q_START_DATE, title='Start date'),
                        y=Axis(search_quantity='count'),
                        title='Start date',
                        show_input=True,
                        n_bins=20,
                    ),
                ],
            ),
        ],
    ),
    dashboard=Dashboard(
        widgets=[
            # ------------------------------------------------------------------
            # 1. Events over time  (histogram – full width top row)
            # ------------------------------------------------------------------
            WidgetHistogram(
                title='Events over time',
                description='Distribution of FAIRmat events by start date.',
                type='histogram',
                x=Axis(search_quantity=Q_START_DATE, title='Start date'),
                y=Axis(search_quantity='count'),
                show_input=False,
                n_bins=30,
                autorange=True,
                layout={
                    'lg': Layout(h=6, w=12, x=0, y=0),
                    'md': Layout(h=6, w=12, x=0, y=0),
                    'sm': Layout(h=6, w=12, x=0, y=0),
                },
            ),
            # ------------------------------------------------------------------
            # 2. Events by type
            # ------------------------------------------------------------------
            WidgetTerms(
                title='Events by type',
                type='terms',
                search_quantity=Q_EVENT_TYPE,
                scale='linear',
                show_input=False,
                layout={
                    'lg': Layout(h=6, w=4, x=0, y=6),
                    'md': Layout(h=6, w=6, x=0, y=6),
                    'sm': Layout(h=6, w=12, x=0, y=6),
                },
            ),
            # ------------------------------------------------------------------
            # 3. Events by FAIRmat role
            # ------------------------------------------------------------------
            WidgetTerms(
                title='FAIRmat role',
                type='terms',
                search_quantity=Q_FAIRMAT_ROLE,
                scale='linear',
                show_input=False,
                layout={
                    'lg': Layout(h=6, w=4, x=4, y=6),
                    'md': Layout(h=6, w=6, x=6, y=6),
                    'sm': Layout(h=6, w=12, x=0, y=12),
                },
            ),
        ]
    ),
)

app_entry_point = AppEntryPoint(
    name='FAIRmatEventsApp',
    description='Search and visualise FAIRmat events: talks, workshops, schools, conferences and more.',
    app=fairmat_events_app,
)
