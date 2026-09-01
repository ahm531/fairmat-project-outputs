"""Tests for the event parsers' text- and vocabulary-normalisation helpers."""

from fairmat_project_outputs.parsers.dpg_event_parser import (
    _normalise_event_key,
    _normalise_event_name,
    _split_text,
)
from fairmat_project_outputs.parsers.generic_event_parser import (
    _normalise_event_type,
    _normalise_fabio_type,
    _split_contributors,
)


def test_split_text_separates_event_and_contribution():
    event, title, ctype = _split_text('DPG-SKM 2025 talk: Inverse Design of Materials')

    assert event == 'DPG-SKM 2025'
    assert title == 'Inverse Design of Materials'
    assert ctype == 'Contributed talk'


def test_split_text_without_keyword_is_event_only():
    event, title, ctype = _split_text('DPG Fall Meeting')

    assert event == 'DPG Fall Meeting'
    assert not title
    assert not ctype


def test_skm_variants_share_one_event_key():
    assert (
        _normalise_event_key('DPG SKM Frühjahrstagung 2022 (DPG_SKM22)')
        == _normalise_event_key('DPG_SKM22')
        == 'skm:2022'
    )


def test_normalise_event_name_expands_skm():
    assert (
        _normalise_event_name('DPG_SKM21', 2021)
        == 'DPG Spring Meeting of the Condensed Matter Section 2021'
    )


def test_fabio_type_accepts_bracketed_and_unspaced_values():
    assert _normalise_fabio_type('[fabio: Web page]') == 'fabio: Web page'
    assert _normalise_fabio_type('[fabio:Web page]') == 'fabio: Web page'
    assert _normalise_fabio_type('not a fabio term') is None


def test_event_type_falls_back_to_name_keywords():
    assert _normalise_event_type('Conferences and Meetings (CM)', 'x') == 'Conference'
    assert _normalise_event_type('', 'Big Workshop') == 'Workshop'


def test_split_contributors_drops_placeholders():
    assert _split_contributors('A. One, B. Two') == ['A. One', 'B. Two']
    assert _split_contributors('Various') is None
