"""Book production: house colours, tashkil, headings, faharis, English books,
note marks and rules — and a guard that every design method is a server tool
with the same defaults."""

from __future__ import annotations

import inspect
import re
from pathlib import Path

import pytest

from docx_mcp.document import DocxDocument
from docx_mcp.document.arabic import (
    has_tashkil,
    loose_pattern,
    normalize_tashkil,
    strip_case_endings,
    strip_tashkil,
    tashkil_ratio,
)
from docx_mcp.document.fontmetrics import text_width_pt
from docx_mcp.document.islamic_art import note_rule

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
HADITH = "قَالَ رَسُولُ اللَّهِ ﷺ: إِنَّمَا الأَعْمَالُ بِالنِّيَّاتِ"


@pytest.fixture
def doc(tmp_path: Path):
    d = DocxDocument.create(str(tmp_path / "book.docx"))
    d._first = d._trees["word/document.xml"].find(f".//{W}p").get(f"{W14}paraId")
    yield d
    d.close()


def _para(d, pid):
    return d._find_para(d._trees["word/document.xml"], pid)


def _colors(el):
    return {c.get(f"{W}val") for c in el.iter(f"{W}color")}


# ── Arabic text helpers ──────────────────────────────────────────────────────


def test_strip_and_keep_shadda():
    assert strip_tashkil(HADITH).startswith("قال رسول الله")
    assert "ّ" in strip_tashkil("اللَّهِ", keep_shadda=True)
    assert not has_tashkil(strip_tashkil(HADITH))


def test_loose_pattern_matches_vocalised_text():
    assert re.search(loose_pattern("قال الله"), "قَالَ اللَّهُ")
    assert re.search(loose_pattern("الأعمال"), "الأَعْمَالُ")
    assert not re.search(loose_pattern("قال"), "قيل")


def test_ratio_normalize_and_light():
    assert tashkil_ratio("قال") == 0
    assert tashkil_ratio("قَالَ") > 0.5
    # shadda goes before the vowel whatever order it was typed in
    assert normalize_tashkil("لَّ") == "لَّ"
    assert strip_case_endings("رَسُولُ") == "رَسُول"


def test_font_width_measures_or_estimates():
    w = text_width_pt("باب الطهارة", "Amiri", 20)
    assert 40 < w < 200
    assert text_width_pt("abc", "No Such Font", 10) > 0


@pytest.mark.parametrize("kind", ["tapered", "khatam", "swash", "diamond"])
def test_note_rule_svgs(kind):
    svg = note_rule(kind, "1A1A1A")
    assert svg.startswith("<svg") and "1A1A1A" in svg.upper()


# ── House colours ────────────────────────────────────────────────────────────


def test_house_palettes(doc):
    pals = doc.list_house_palettes()
    assert {"maroon", "navy-gold", "green-gold", "ink"} <= set(pals)
    r = doc.set_house_colors("teal")
    assert r["accent"] == "0F4C5C"
    assert doc.get_house_colors()["accent"] == "0F4C5C"
    r = doc.set_house_colors(accent="#123456")
    assert r == {**r, "palette": "custom", "accent": "123456"}
    with pytest.raises(ValueError):
        doc.set_house_colors("neon")


def test_design_tools_follow_house_colours(doc):
    doc.set_house_colors("navy-gold")
    pid = doc.insert_separator(doc._first)["para_id"]
    assert "C00000" not in _colors(_para(doc, pid))
    r = doc.insert_ornament("divider", after_para_id=doc._first)
    assert r
    doc.save()
    media = [
        n for n in doc._trees if n.endswith(".svg")
    ]  # the art is stored as SVG parts in the package
    svgs = "".join(
        Path(doc.workdir, n).read_text(encoding="utf-8")
        for n in media
        if Path(doc.workdir, n).exists()
    )
    if svgs:
        assert "520912" not in svgs.upper()


# ── Presets, headings, note marks ────────────────────────────────────────────


@pytest.mark.parametrize("design", ["plain", "fleuron", "flanked", "cartouche", "bullet"])
def test_heading_designs(doc, design):
    doc.apply_book_preset("tahqiq", title="رسالة")
    level = 2 if design == "bullet" else 1
    r = doc.insert_heading("باب الطهارة", level=level, after_para_id=doc._first, design=design)
    p = _para(doc, r["para_id"])
    assert p.find(f"{W}pPr/{W}pStyle").get(f"{W}val") == f"Heading{level}"
    assert _para(doc, r["after_para_id"]) is not None  # every design says where to continue
    # no inline picture in the heading itself (it would be copied into the TOC)
    assert not list(
        p.iter("{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}inline")
    )
    doc.save()


def test_heading_bad_design(doc):
    with pytest.raises(ValueError):
        doc.insert_heading("x", design="neon")


@pytest.mark.parametrize("preset", ["tahqiq", "kitab", "urdu-nastaliq", "english", "bilingual"])
def test_presets_save(doc, preset):
    r = doc.apply_book_preset(preset, title="Book", author="A")
    assert r["applied"]
    doc.save()


def test_note_marks_parentheses_and_dash(doc, tmp_path):
    doc.apply_book_preset("tahqiq")
    pid = doc.insert_paragraph(doc._first, HADITH)["para_id"]
    doc.add_footnote(pid, "صَحِيحُ البُخَارِيِّ")
    doc.save()
    notes = doc.get_footnotes()
    # parentheses around the live mark, and the reference is unvocalised
    assert notes[0]["text"].startswith("()") and not has_tashkil(notes[0]["text"])

    eng = DocxDocument.create(str(tmp_path / "en.docx"))
    eng.apply_book_preset("english")
    p0 = eng._trees["word/document.xml"].find(f".//{W}p").get(f"{W14}paraId")
    pid = eng.insert_paragraph(p0, "Narrated Anas.")["para_id"]
    eng.add_footnote(pid, "Related by al-Bukhari.")
    eng.save()
    fn = eng._trees["word/footnotes.xml"]
    texts = [t.text for t in fn.iter(f"{W}t")]
    assert "-" in [x.strip() for x in texts if x]
    assert any(x and x.startswith("‎") for x in texts)  # LRM fences the number
    eng.close()


def test_note_mark_styles_switch_cleanly(doc):
    pid = doc.insert_paragraph(doc._first, "نص")["para_id"]
    doc.add_footnote(pid, "حاشية")
    for style in ("parentheses", "brackets", "dash", "plain", "parentheses"):
        doc.set_note_mark_style(style)
    body = doc._trees["word/document.xml"]
    texts = [t.text for t in body.iter(f"{W}t") if t.text in ("(", ")", "[", "]")]
    assert texts == ["(", ")"]  # earlier wrappings were removed, not stacked
    with pytest.raises(ValueError):
        doc.set_note_mark_style("stars")


@pytest.mark.parametrize(
    "style", ["rule", "double", "tapered", "khatam", "diamond", "swash", "text", "full", "none"]
)
def test_note_separator_styles(doc, style):
    doc.set_note_separator("footnote", style, text="⁂", width_pct=60)
    doc.save()


# ── Tashkil on the document ─────────────────────────────────────────────────


def test_tashkil_tools(doc):
    pid = doc.insert_paragraph(doc._first, HADITH)["para_id"]
    rep = doc.tashkil_report()
    assert rep["partial"] + rep["full"] >= 1
    r = doc.color_tashkil()
    assert r["color"] == doc.get_house_colors()["accent"] and r["runs_split"] == 0
    assert doc.get_custom_properties()["docx-mcp.diacritic-color"] == r["color"]
    doc.color_tashkil(per_run=True)
    # a second call recolours the already split mark runs
    r2 = doc.color_tashkil("0000FF", per_run=True)
    assert r2["runs_split"] > 0 and "0000FF" in _colors(_para(doc, pid))
    # every split piece stays an RTL run, or Word breaks the word apart
    runs = _para(doc, pid).findall(f"{W}r")
    assert all(r.find(f"{W}rPr/{W}rtl") is not None for r in runs if "".join(r.itertext()).strip())
    doc.color_tashkil("auto", per_run=True)
    assert "0000FF" not in _colors(_para(doc, pid))
    assert "docx-mcp.diacritic-color" not in doc.get_custom_properties()
    doc.tashkil("strip", para_ids=[pid])
    assert not has_tashkil("".join(_para(doc, pid).itertext()))
    with pytest.raises(ValueError):
        doc.tashkil("shout")
    doc.save()


def test_format_matches_ignores_tashkil(doc):
    doc.insert_paragraph(doc._first, HADITH)
    r = doc.format_matches("رسول الله", bold=True)
    assert r["formatted"] >= 1
    r = doc.format_matches("رسول الله", bold=True, ignore_tashkil=False)
    assert r["formatted"] == 0


# ── Faharis and English book parts ──────────────────────────────────────────


def test_build_faharis(doc):
    doc.apply_book_preset("tahqiq", title="رسالة")
    pid = doc.insert_paragraph(doc._first, HADITH, style="Normal")["para_id"]
    doc.format_matches("إنما الأعمال بالنيات", char_style="Prophetic Speech")
    doc.insert_paragraph(pid, "﴿وَمَا أُمِرُوا إِلَّا لِيَعْبُدُوا اللَّهَ﴾", style="Normal")
    r = doc.build_faharis(
        people=["رسول الله"],
        places=[["المدينة", "المدينة النبوية"]],
        sources=["صَحِيحُ البُخَارِيِّ", "سنن أبي داود"],
    )
    assert r
    doc.save()
    text = "".join(doc._trees["word/document.xml"].itertext())
    assert "الفهارس" in text and "صحيح البخاري" in text  # sources unvocalised
    instr = " ".join(t.text or "" for t in doc._trees["word/document.xml"].iter(f"{W}instrText"))
    assert "INDEX" in instr and "XE" in instr


def test_english_entries(doc):
    doc.apply_book_preset("bilingual", title="Bulugh al-Maram")
    h = doc.insert_bilingual_heading(
        "كتاب الطهارة", "Book 1: Purification", after_para_id=doc._first
    )
    e = doc.insert_hadith_entry(
        94,
        "وَعَنْ أَنَسٍ رضي الله عنه قَالَ",
        'Narrated Anas: "Whenever he entered the privy."',
        after_para_id=h["para_id"],
        source="Related by the seven Imams.",
    )
    g = doc.insert_glossary_entry("Ghusl", "The full ritual washing.")
    assert e and g
    text = "".join(doc._trees["word/document.xml"].itertext())
    assert "٩٤" in text and "‎94." in text
    # Arabic text style never uses medium/high kashida (Word drops lines)
    jcs = {j.get(f"{W}val") for j in doc._trees["word/styles.xml"].iter(f"{W}jc")}
    assert not jcs & {"mediumKashida", "highKashida"}
    doc.save()


# ── Every design method is a server tool with the same defaults ─────────────

_MODULES = {
    "house", "vectors", "notes", "bookparts", "bookdesign", "arabic",
    "typography", "indexing", "pdfexport", "pagedesign",
}  # fmt: skip
_SERVER_SHAPES_DIFFER = {"format_matches"}  # run props are spelled out in the tool


def test_server_tools_match_methods():
    import docx_mcp.server as server

    problems = []
    for name, fn in inspect.getmembers(DocxDocument, inspect.isfunction):
        if name.startswith("_") or fn.__module__.split(".")[-1] not in _MODULES:
            continue
        tool = getattr(server, name, None)
        if tool is None:
            problems.append(f"no tool: {name}")
            continue
        tool = getattr(tool, "fn", tool)
        sp = inspect.signature(tool).parameters
        for k, p in inspect.signature(fn).parameters.items():
            if k == "self":
                continue
            if k not in sp and name not in _SERVER_SHAPES_DIFFER:
                problems.append(f"{name}: missing {k}")
            elif k not in sp:
                continue
            elif (
                p.default is not inspect._empty
                and sp[k].default != p.default
                and not (k == "sides" and sp[k].default is None)
            ):
                problems.append(f"{name}.{k}: {sp[k].default!r} != {p.default!r}")
    assert not problems, problems
