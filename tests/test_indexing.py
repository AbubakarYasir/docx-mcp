"""Indexes (XE / INDEX) and style-based tables of contents."""

from __future__ import annotations

from pathlib import Path

import pytest

from docx_mcp.document import DocxDocument

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"


@pytest.fixture
def doc(tmp_path: Path):
    d = DocxDocument.create(str(tmp_path / "i.docx"))
    pid = d._trees["word/document.xml"].find(f".//{W}p").get(f"{W14}paraId")
    d.set_document_direction("rtl", complex_script_font="Amiri")
    d.define_style("Chapter", "paragraph", based_on="Normal", run={"bold": True})
    d.define_style("Bab", "paragraph", based_on="Normal")
    lines = [
        ("Chapter", "كتاب الإيمان"),
        ("Bab", "باب النية"),
        (None, "حدثنا الحميدي عن سفيان عن عمر بن الخطاب بالمدينة."),
        (None, "وقال البخاري: حدثنا مسدد عن أنس بن مالك، والبخاري ثقة."),
        ("Chapter", "كتاب العلم"),
        (None, "ورواه مسلم عن قتيبة عن عمر بن الخطاب بمكة."),
    ]
    pids = []
    for style, text in lines:
        kw = {"style": style} if style else {}
        pid = d.insert_paragraph(pid, text, **kw)["para_id"]
        pids.append(pid)
    d._pids = pids
    yield d
    d.close()


def _instrs(d):
    return [it.text for it in d._trees["word/document.xml"].iter(f"{W}instrText")]


def test_mark_terms_with_types_and_proclitics(doc):
    r = doc.mark_index_entries(["عمر بن الخطاب", "البخاري"], index_type="P")
    assert r["per_term"] == {"عمر بن الخطاب": 2, "البخاري": 1}  # first per paragraph
    r = doc.mark_index_entries(["البخاري"], index_type="P", first_per_paragraph=False)
    assert r["per_term"]["البخاري"] == 2  # «والبخاري» counts, filed under البخاري
    xe = [i for i in _instrs(doc) if " XE " in f" {i}"]
    assert all('\\f "P"' in i for i in xe)
    assert not any("والبخاري" in i for i in xe)
    listed = doc.list_index_entries("P")
    names = [e["entry"] for e in listed["entries"]]
    assert names == sorted(names, key=lambda n: n.replace("أ", "ا"))
    assert {e["entry"]: e["count"] for e in listed["entries"]}["البخاري"] == 3


def test_entry_subentry_see_and_escaping(doc):
    doc.mark_index_entries(
        [{"text": "أنس بن مالك", "entry": "أنس: ابن مالك", "subentry": "روايته", "bold_page": True}]
    )
    doc.mark_index_entry(doc._pids[2], "الحميدي", see="عبد الله بن الزبير")
    instr = " ".join(_instrs(doc))
    assert 'XE "أنس\\: ابن مالك:روايته" \\b' in instr
    assert '\\t "عبد الله بن الزبير"' in instr
    e = [x for x in doc.list_index_entries()["entries"] if x["entry"].startswith("أنس")][0]
    assert e["entry"] == "أنس: ابن مالك" and e["subentries"] == ["روايته"] and e["bold_page"]


def test_bad_index_type(doc):
    with pytest.raises(ValueError):
        doc.mark_index_entries(["مسلم"], index_type="people")


def test_insert_index_is_one_multi_paragraph_field(doc):
    doc.mark_index_entries(["عمر بن الخطاب", "مسلم"], index_type="P")
    doc.mark_index_entries(["المدينة"], index_type="L")
    r = doc.insert_index(title="فهرس الأعلام", index_type="P", columns=2, heading_color="520912")
    assert '\\c "2"' in r["field"] and '\\f "P"' in r["field"] and '\\z "1025"' in r["field"]
    assert r["entries"] == 2
    body = doc._trees["word/document.xml"].find(f"{W}body")
    kinds = []
    for p in body.iter(f"{W}p"):
        for fc in p.iter(f"{W}fldChar"):
            kinds.append(fc.get(f"{W}fldCharType"))
    # XE fields: begin/end pairs; INDEX: begin, separate, end in different paragraphs.
    assert kinds.count("begin") == kinds.count("end")
    idx_para = doc._find_para(doc._trees["word/document.xml"], r["para_id"])
    assert idx_para.find(f"{W}pPr/{W}pStyle").get(f"{W}val").lower().startswith("index")
    assert doc._trees["word/settings.xml"].find(f"{W}updateFields") is not None
    doc.save()


def test_index_option_validation(doc):
    with pytest.raises(ValueError):
        doc.insert_index(columns=6)
    with pytest.raises(ValueError):
        doc.insert_index(leader="stars")
    with pytest.raises(ValueError):
        doc.insert_index(language="xx")


def test_toc_from_styles_has_no_u_switch(doc):
    r = doc.insert_toc(title="فهرس الموضوعات", levels=None, styles={"Chapter": 1, "Bab": 2})
    assert r["field"] == 'TOC \\t "Chapter,1,Bab,2" \\h \\z'
    assert r["entries"] == 3
    r2 = doc.insert_toc(levels="1-2", new_page=False)
    assert "\\u" in r2["field"]
    doc.save()


def test_toc_errors(doc):
    with pytest.raises(ValueError):
        doc.insert_toc(levels=None, styles=None)
    with pytest.raises(ValueError):
        doc.insert_toc(styles={"Nope": 1})
    with pytest.raises(ValueError):
        doc.insert_toc(levels="one-three")


def test_toc_after_paragraph(doc):
    r = doc.insert_toc(
        after_para_id=doc._pids[0], levels=None, styles={"Chapter": 1}, title="المحتويات"
    )
    body = list(doc._trees["word/document.xml"].find(f"{W}body"))
    first = doc._find_para(doc._trees["word/document.xml"], doc._pids[0])
    nxt = body[body.index(first) + 1]
    assert "".join(nxt.itertext()) == "المحتويات"
    assert nxt.find(f"{W}pPr/{W}pageBreakBefore") is not None
    assert r["entries"] == 2


def test_remove_index_entries(doc):
    doc.mark_index_entries(["عمر بن الخطاب"], index_type="P")
    doc.mark_index_entries(["المدينة"], index_type="L")
    assert doc.remove_index_entries(index_type="L")["removed"] == 1
    assert doc.list_index_entries()["count"] == 1
    assert doc.remove_index_entries()["removed"] == 2
    assert not any("XE" in (i or "") for i in _instrs(doc))
    doc.save()
