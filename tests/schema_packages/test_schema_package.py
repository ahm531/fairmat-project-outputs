import os.path

from nomad.client import normalize_all, parse

from fairmat_project_outputs.schema_packages.schema_package import FABIO_TERMS


def test_fabio_terms_are_spelt_as_the_ontologies_spell_them():
    """FaBiO labels are lowercase; BIBO capitalises its own.

    Checked against the 2026-09 release of http://purl.org/spar/fabio. The three
    terms this list used to carry -- 'Poster', 'Software dataset' and
    'Interview' -- are not FaBiO classes at all, so they could never resolve to
    a URI. Anything written into an archive has to match this list exactly.
    """
    terms = set(FABIO_TERMS)

    for term in terms:
        prefix, _, label = term.partition(': ')
        assert prefix in ('fabio', 'bibo'), f'{term!r} names no ontology'
        if prefix == 'fabio':
            assert label == label.lower(), f'{term!r} is not the FaBiO label'

    assert 'fabio: conference poster' in terms  # not 'Poster'
    assert 'fabio: dataset' in terms  # not 'Software dataset'
    assert 'bibo: Interview' in terms  # FaBiO has no interview class
    for gone in ('fabio: Poster', 'fabio: Software dataset', 'fabio: Interview'):
        assert gone not in terms


def test_schema_package():
    test_file = os.path.join('tests', 'data', 'test.archive.yaml')
    entry_archive = parse(test_file)[0]
    normalize_all(entry_archive)

    data = entry_archive.data
    assert data.event_name == 'Test Users Meeting'
    assert [c.title for c in data.contributions] == [
        'Introduction to FAIRmat',
        'FAIR data in practice',
    ]

    # normalize() collects the contributors of every contribution onto the event
    assert data.fairmat_contributors == ['Jane Doe', 'John Roe']

    # ... and renders the overview table
    assert 'Introduction to FAIRmat' in data.contributions_overview
