"""Book presets, style-name resolution and Word power features."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from lxml import etree

from docx_mcp.document import DocxDocument

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"


@pytest.fixture
def doc(tmp_path: Path):
    d = DocxDocument.create(str(tmp_path / "b.docx"))
    d._first = d._trees["word/document.xml"].find(f".//{W}p").get(f"{W14}paraId")
    yield d
    d.close()


def _style(d, sid):
    for s in d._trees["word/styles.xml"].findall(f"{W}style"):
        if s.get(f"{W}styleId") == sid:
            return s
    return None


def test_style_names_and_labels_resolve_to_ids(doc):
    doc.define_style("Narration Number", "character", aliases=["رقم الحديث"], run={"bold": True})
    pid = doc.insert_paragraph(doc._first, "باب", style="Heading 1")["para_id"]
    p = doc._find_para(doc._trees["word/document.xml"], pid)
    assert p.find(f"{W}pPr/{W}pStyle").get(f"{W}val") == "Heading1"
    doc.update_paragraph(pid, style="heading 2")
    assert p.find(f"{W}pPr/{W}pStyle").get(f"{W}val") == "Heading2"
    doc.format_matches("باب", char_style="رقم الحديث")
    assert any(
        r.get(f"{W}val") == "NarrationNumber"
        for r in doc._trees["word/document.xml"].iter(f"{W}rStyle")
    )
    # Unknown values pass through untouched.
    assert doc._style_id_for("NoSuchStyle") == "NoSuchStyle"


def test_tahqiq_preset(doc):
    r = doc.apply_book_preset("tahqiq", title="رسالة", author="المؤلف", running_head="fleuron")
    assert r["applied"][-1] == "settings"
    normal = _style(doc, "Normal")
    assert normal.find(f"{W}rPr/{W}rFonts").get(f"{W}cs") == "Amiri"
    assert normal.find(f"{W}rPr/{W}szCs").get(f"{W}val") == "26"
    h1 = _style(doc, "Heading1")
    # headings are typographic (no ruled boxes); the design sits in the art
    assert h1.find(f"{W}pPr/{W}pBdr") is None
    assert h1.find(f"{W}pPr/{W}jc").get(f"{W}val") == "center"
    assert h1.find(f"{W}rPr/{W}color").get(f"{W}val") == "520912"
    assert h1.find(f"{W}pPr/{W}pageBreakBefore") is None or h1.find(
        f"{W}pPr/{W}pageBreakBefore"
    ).get(f"{W}val") in ("0", "false")
    fn = _style(doc, "FootnoteText")
    sp = fn.find(f"{W}pPr/{W}spacing")
    # vocalised text: "at least" spacing so harakat are never clipped
    assert sp.get(f"{W}lineRule") == "atLeast" and sp.get(f"{W}line") == "300"
    props = doc.get_custom_properties()
    assert props["docx-mcp.note-marks"] == "parentheses"
    assert props["docx-mcp.notes-tashkil"] == "strip"
    assert props["docx-mcp.numerals"] == "hindi"
    settings = doc._trees["word/settings.xml"]
    for flag in (
        "evenAndOddHeaders",
        "hideSpellingErrors",
        "updateFields",
        "stylePaneFormatFilter",
    ):
        assert settings.find(f"{W}{flag}") is not None, flag
    sect = doc._trees["word/document.xml"].find(f".//{W}sectPr")
    assert sect.find(f"{W}titlePg") is not None
    variants = {h.get(f"{W}type") for h in sect.findall(f"{W}headerReference")}
    assert variants == {"default", "even", "first"}
    doc.save()


def test_kitab_preset_new_pages_and_kashida(doc):
    doc.apply_book_preset("kitab", justify="kashida", honorifics=False, overrides={"body_pt": 15})
    h1 = _style(doc, "Heading1")
    assert h1.find(f"{W}pPr/{W}pageBreakBefore") is not None
    # medium/high kashida drop vocalised lines in Word: low kashida only
    assert _style(doc, "Normal").find(f"{W}pPr/{W}jc").get(f"{W}val") == "lowKashida"
    assert _style(doc, "Normal").find(f"{W}rPr/{W}sz").get(f"{W}val") == "30"


def test_preset_errors(doc):
    with pytest.raises(ValueError):
        doc.apply_book_preset("comic")
    with pytest.raises(ValueError):
        doc.apply_book_preset("tahqiq", running_head="sidebar")


def test_separator(doc):
    r = doc.insert_separator(doc._first)
    p = doc._find_para(doc._trees["word/document.xml"], r["para_id"])
    assert "".join(p.itertext()) == "❁ ❁ ❁"
    r2 = doc.insert_separator(doc._first, ornament="divider-dots")
    assert r2["kind"] == "divider-dots"


def test_restrict_formatting(doc):
    doc.define_style("Matn", "character", run={"bold": True})
    r = doc.restrict_formatting_to_styles(["Normal", "Heading 1", "Matn"])
    assert r["locked_styles"] > 0
    assert _style(doc, "Matn").find(f"{W}locked") is None
    assert _style(doc, "Heading2").find(f"{W}locked") is not None
    prot = doc._trees["word/settings.xml"].find(f"{W}documentProtection")
    assert prot.get(f"{W}formatting") == "1" and prot.get(f"{W}enforcement") == "1"
    doc.save()
    doc.restrict_formatting_to_styles(enabled=False)
    assert doc._trees["word/settings.xml"].find(f"{W}documentProtection") is None
    assert _style(doc, "Heading2").find(f"{W}locked") is None


def test_style_pane_and_flags(doc):
    r = doc.set_style_pane("in-use", sort="name")
    assert r["filter"] == "4808"
    with pytest.raises(ValueError):
        doc.set_style_pane("everything")
    doc.set_font_embedding(True)
    doc.hide_proofing_marks(True)
    s = doc._trees["word/settings.xml"]
    assert s.find(f"{W}embedTrueTypeFonts") is not None
    assert s.find(f"{W}saveSubsetFonts") is not None
    doc.set_font_embedding(False)
    assert s.find(f"{W}embedTrueTypeFonts") is None
    doc.save()


def test_attach_template(doc):
    r = doc.attach_template("C:\\Users\\Hp\\Templates\\Book.dotx")
    assert r["template"] == "file:///C:/Users/Hp/Templates/Book.dotx"
    doc.save()
    with zipfile.ZipFile(doc.source_path) as z:
        rels = z.read("word/_rels/settings.xml.rels").decode()
        settings = etree.fromstring(z.read("word/settings.xml"))
    assert "attachedTemplate" in rels and 'TargetMode="External"' in rels
    assert settings.find(f"{W}attachedTemplate") is not None
    assert settings.find(f"{W}linkStyles") is not None
