"""Publishing conventions: running heads, note marks, verse, margin notes,
numbered entries, part-title pages."""

from __future__ import annotations

from pathlib import Path

import pytest

from docx_mcp.document import DocxDocument

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"


@pytest.fixture
def doc(tmp_path: Path):
    d = DocxDocument.create(str(tmp_path / "p.docx"))
    d.set_document_direction("rtl", complex_script_font="Amiri")
    first = d._trees["word/document.xml"].find(f".//{W}p").get(f"{W14}paraId")
    d._pids = [d.insert_paragraph(first, f"فقرة {i}")["para_id"] for i in range(3)]
    yield d
    d.close()


def _body(d):
    return d._trees["word/document.xml"]


@pytest.mark.parametrize("style", ["rule", "boxed", "vignette", "cartouche"])
def test_running_heads(doc, style):
    r = doc.design_running_head(style, title="الكتاب")
    assert len(r["headers"]) == 2
    sect = _body(doc).find(f".//{W}sectPr")
    assert {h.get(f"{W}type") for h in sect.findall(f"{W}headerReference")} >= {"default", "even"}
    head = doc._trees[r["headers"][0]]
    instr = " ".join(i.text for i in head.iter(f"{W}instrText"))
    assert "STYLEREF" in instr or style == "cartouche"
    if style == "boxed":
        assert head.find(f".//{W}bdr") is not None
    if style in ("vignette", "cartouche"):
        assert head.find(f".//{WP}inline") is not None or head.find(f".//{WP}anchor") is not None
    doc.save()
    with pytest.raises(ValueError):
        doc.design_running_head("neon")


def test_note_marks_parentheses_idempotent(doc):
    doc.add_footnote(doc._pids[0], "الأولى")
    doc.add_footnote(doc._pids[1], "الثانية")
    assert doc.set_note_mark_style("parentheses")["marks"] == 2
    doc.set_note_mark_style("parentheses")  # again: no double wrapping
    p = doc._find_para(_body(doc), doc._pids[0])
    texts = [t.text for t in p.iter(f"{W}t")]
    assert texts.count("(") == 1 and texts.count(")") == 1
    notes = doc._trees["word/footnotes.xml"]
    note_texts = "".join(t.text or "" for t in notes.iter(f"{W}t"))
    assert note_texts.count("(") == 2 and ") " in note_texts
    doc.set_note_mark_style("brackets")
    assert "[" in [t.text for t in p.iter(f"{W}t")]
    doc.save()


def test_verses_table(doc):
    r = doc.insert_verses(
        [
            "الحمد لله على التمام *** في البدء والختام والدوام",
            ["ثم الصلاة مع سلام شائع", "على النبي وصحبه والتابع"],
        ],
        after_para_id=doc._pids[0],
        numbered=True,
        start_at=48,
    )
    assert r["verses"] == 2
    tbl = _body(doc).find(f".//{W}tbl")
    assert tbl.find(f"{W}tblPr/{W}bidiVisual") is not None
    rows = tbl.findall(f"{W}tr")
    assert "".join(rows[0].itertext()).startswith("٤٨-")
    assert rows[0].findall(f"{W}tc")[1].find(f".//{W}jc").get(f"{W}val") == "distribute"
    with pytest.raises(ValueError):
        doc.insert_verses([])
    doc.save()


def test_margin_note(doc):
    doc.add_margin_note(doc._pids[1], "الخاتمة")
    an = _body(doc).find(f".//{WP}anchor")
    assert an.find(f"{WP}positionH").get("relativeFrom") == "outsideMargin"
    assert "الخاتمة" in "".join(an.itertext())
    with pytest.raises(ValueError):
        doc.add_margin_note(doc._pids[1], "x", side="top")
    doc.save()


def test_numbered_style(doc):
    r = doc.define_numbered_style("Tarjama Entry", template="%1 .", aliases=["ترجمة مرقمة"])
    num = doc._trees["word/numbering.xml"]
    lvl = [a for a in num.findall(f"{W}abstractNum")][-1].find(f"{W}lvl")
    assert lvl.find(f"{W}lvlText").get(f"{W}val") == "%1 ."
    pid = doc.insert_paragraph(doc._pids[0], "سليمان بن عمرو:", style="ترجمة مرقمة")["para_id"]
    p = doc._find_para(_body(doc), pid)
    assert p.find(f"{W}pPr/{W}pStyle").get(f"{W}val") == r["style_id"]
    with pytest.raises(ValueError):
        doc.define_numbered_style("X", number_format="klingon")
    doc.save()


def test_part_title_page(doc):
    r = doc.insert_part_title("الفهارس", after_para_id=doc._pids[1], items=["فهرس الموضوعات"])
    sects = list(_body(doc).iter(f"{W}sectPr"))
    assert len(sects) == 3
    centred = sects[1]
    assert centred.find(f"{W}vAlign").get(f"{W}val") == "center"
    assert centred.find(f"{W}headerReference") is not None  # frame lives here
    # the section after it has its own heads, so it does not inherit the frame
    assert sects[2].find(f"{W}headerReference") is not None
    assert r["para_id"]
    doc.save()
