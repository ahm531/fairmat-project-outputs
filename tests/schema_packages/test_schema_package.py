import os.path

from nomad.client import normalize_all, parse


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
