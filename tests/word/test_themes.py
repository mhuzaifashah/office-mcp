"""Document theme — `doc.theme` (colours/fonts/apply/list), `bgr_to_hex`, the
`apply_theme` / `set_theme_colors` / `set_theme_fonts` ops, and the CLI.

`doc.DocumentTheme` is a `_FakeOfficeTheme`: `ThemeColorScheme.Colors(i).RGB` is a
settable BGR int, `ThemeFontScheme.Major/MinorFont.Item(1).Name` is settable, and
`Load`/`ApplyDocumentTheme` are recording MagicMocks. Built-in *name* resolution
is exercised by monkeypatching `wordcom.model._themes._themes_dir` onto a tmp library.
"""

from __future__ import annotations

import os

import pytest

from office_mcp.word import wordcom
from office_mcp.word.wordcom.exceptions import OpError
from office_mcp.word.wordcom.model._format import bgr_to_hex, to_bgr
from office_mcp.word.wordcom.runtime._ops import run_batch


def _theme(fake_word):
    return fake_word.ActiveDocument.DocumentTheme


def _make_library(tmp_path):
    """A fake Office theme library; returns its root path."""
    (tmp_path / "Facet.thmx").write_text("thmx")
    (tmp_path / "Ion.thmx").write_text("thmx")
    colors = tmp_path / "Theme Colors"
    colors.mkdir()
    (colors / "Blue.xml").write_text("<a:clrScheme/>")
    (colors / "Orange.xml").write_text("<a:clrScheme/>")
    fonts = tmp_path / "Theme Fonts"
    fonts.mkdir()
    (fonts / "Garamond.xml").write_text("<a:fontScheme/>")
    return str(tmp_path)


# --- bgr_to_hex ----------------------------------------------------------------


def test_bgr_to_hex_round_trips_to_bgr():
    for spec in ("#FF0000", "#1A73E8", "#000000", "#FFFFFF"):
        assert bgr_to_hex(to_bgr(spec)) == spec.upper()
    assert bgr_to_hex(to_bgr("navy")) == "#000080"  # named colour


def test_bgr_to_hex_byte_order():
    assert bgr_to_hex(255) == "#FF0000"  # low byte is red
    assert bgr_to_hex(16711680) == "#0000FF"  # high byte is blue


def test_bgr_to_hex_bad_input_raises():
    with pytest.raises(TypeError):
        bgr_to_hex(True)
    with pytest.raises(ValueError):
        bgr_to_hex(0x1000000)


# --- reads ---------------------------------------------------------------------


def test_colors_read(fake_word):
    with wordcom.attach() as word:
        colors = word.documents.active.theme.colors
    assert set(colors) == {
        "text1",
        "background1",
        "text2",
        "background2",
        "accent1",
        "accent2",
        "accent3",
        "accent4",
        "accent5",
        "accent6",
        "hyperlink",
        "followed_hyperlink",
    }
    # accent1 is Colors(5).RGB == 5 * 0x010101 in the fake.
    assert colors["accent1"] == "#050505"


def test_fonts_read(fake_word):
    with wordcom.attach() as word:
        theme = word.documents.active.theme
        assert theme.major_font == "Aptos Display"
        assert theme.minor_font == "Aptos"


def test_to_dict(fake_word):
    with wordcom.attach() as word:
        data = word.documents.active.theme.to_dict()
    assert set(data) == {"colors", "major_font", "minor_font"}
    assert data["major_font"] == "Aptos Display"


# --- set_colors ----------------------------------------------------------------


def test_set_colors_overrides(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("theme"):
            result = doc.theme.set_colors(accent1="#FF0000", text1="navy")
    assert _theme(fake_word).ThemeColorScheme.Colors(5).RGB == to_bgr("#FF0000")
    assert _theme(fake_word).ThemeColorScheme.Colors(1).RGB == to_bgr("navy")
    assert result["accent1"] == "#FF0000"


def test_set_colors_scheme_path_loads(fake_word, tmp_path):
    xml = tmp_path / "brand.xml"
    xml.write_text("<a:clrScheme/>")
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("theme"):
            doc.theme.set_colors(scheme=str(xml))
    _theme(fake_word).ThemeColorScheme.Load.assert_called_once_with(str(xml))


def test_set_colors_builtin_scheme_name(fake_word, tmp_path, monkeypatch):
    lib = _make_library(tmp_path)
    monkeypatch.setattr(wordcom.model._themes, "_themes_dir", lambda app: lib)
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("theme"):
            doc.theme.set_colors(scheme="Blue")
    _theme(fake_word).ThemeColorScheme.Load.assert_called_once_with(
        os.path.join(lib, "Theme Colors", "Blue.xml")
    )


def test_set_colors_unknown_color_raises(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.theme.set_colors(banana="#FFFFFF")


def test_set_colors_unknown_scheme_raises(fake_word, tmp_path, monkeypatch):
    lib = _make_library(tmp_path)
    monkeypatch.setattr(wordcom.model._themes, "_themes_dir", lambda app: lib)
    with wordcom.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.theme.set_colors(scheme="Nonexistent")


# --- set_fonts -----------------------------------------------------------------


def test_set_fonts_explicit(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("theme"):
            result = doc.theme.set_fonts(major="Arial", minor="Calibri")
    assert _theme(fake_word).ThemeFontScheme.MajorFont.Item(1).Name == "Arial"
    assert _theme(fake_word).ThemeFontScheme.MinorFont.Item(1).Name == "Calibri"
    assert result == {"major_font": "Arial", "minor_font": "Calibri"}


def test_set_fonts_scheme_path_loads(fake_word, tmp_path):
    xml = tmp_path / "fonts.xml"
    xml.write_text("<a:fontScheme/>")
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("theme"):
            doc.theme.set_fonts(scheme=str(xml))
    _theme(fake_word).ThemeFontScheme.Load.assert_called_once_with(str(xml))


# --- apply ---------------------------------------------------------------------


def test_apply_theme_path(fake_word, tmp_path):
    thmx = tmp_path / "Brand.thmx"
    thmx.write_text("thmx")
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("theme"):
            name = doc.theme.apply(str(thmx))
    fake_word.ActiveDocument.ApplyDocumentTheme.assert_called_once_with(str(thmx))
    assert name == "Brand"


def test_apply_theme_builtin_name(fake_word, tmp_path, monkeypatch):
    lib = _make_library(tmp_path)
    monkeypatch.setattr(wordcom.model._themes, "_themes_dir", lambda app: lib)
    with wordcom.attach() as word:
        doc = word.documents.active
        with doc.edit("theme"):
            name = doc.theme.apply("Facet")
    fake_word.ActiveDocument.ApplyDocumentTheme.assert_called_once_with(
        os.path.join(lib, "Facet.thmx")
    )
    assert name == "Facet"


def test_apply_theme_unknown_raises(fake_word, tmp_path, monkeypatch):
    lib = _make_library(tmp_path)
    monkeypatch.setattr(wordcom.model._themes, "_themes_dir", lambda app: lib)
    with wordcom.attach() as word:
        doc = word.documents.active
        with pytest.raises(OpError):
            doc.theme.apply("DoesNotExist")


# --- list_available ------------------------------------------------------------


def test_list_available(fake_word, tmp_path, monkeypatch):
    lib = _make_library(tmp_path)
    monkeypatch.setattr(wordcom.model._themes, "_themes_dir", lambda app: lib)
    with wordcom.attach() as word:
        data = word.documents.active.theme.list_available()
    assert data["themes"] == ["Facet", "Ion"]
    assert data["color_schemes"] == ["Blue", "Orange"]
    assert data["font_schemes"] == ["Garamond"]


def test_list_available_missing_dir(fake_word, monkeypatch):
    monkeypatch.setattr(wordcom.model._themes, "_themes_dir", lambda app: r"Z:\nope")
    with wordcom.attach() as word:
        data = word.documents.active.theme.list_available()
    assert data == {"themes": [], "color_schemes": [], "font_schemes": []}


# --- exec ops ------------------------------------------------------------------


def test_exec_apply_theme(fake_word, tmp_path):
    thmx = tmp_path / "Corp.thmx"
    thmx.write_text("thmx")
    with wordcom.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(doc, [{"op": "apply_theme", "theme": str(thmx)}], label="t")
    assert exc is None and result["ok"] is True
    assert result["outputs"][0]["theme"] == "Corp"


def test_exec_set_theme_colors(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(
            doc,
            [{"op": "set_theme_colors", "colors": {"accent1": "#FF0000"}}],
            label="t",
        )
    assert exc is None and result["ok"] is True
    assert result["outputs"][0]["colors"]["accent1"] == "#FF0000"
    assert _theme(fake_word).ThemeColorScheme.Colors(5).RGB == to_bgr("#FF0000")


def test_exec_set_theme_fonts(fake_word):
    with wordcom.attach() as word:
        doc = word.documents.active
        result, exc = run_batch(doc, [{"op": "set_theme_fonts", "major": "Arial"}], label="t")
    assert exc is None and result["ok"] is True
    assert result["outputs"][0]["major_font"] == "Arial"


# --- CLI -----------------------------------------------------------------------
