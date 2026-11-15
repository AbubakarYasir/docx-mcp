"""Note design: continuation marks, separators, numbering, custom marks, conversion."""

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
    d = DocxDocument.create(str(tmp_path / "n.docx"))
    pid = d._trees["word/document.xml"].find(f".//{W}p").get(f"{W14}paraId")
    d.set_document_direction("rtl", complex_script_font="Amiri")
    pids = []
    for _ in range(3):
        pid = d.insert_paragraph(pid, "قال رسول الله ﷺ: «إنما الأعمال بالنيات».")["para_id"]
        pids.append(pid)
    d._pids = pids
    yield d
    d.close()


def _part(path: Path, name: str) -> etree._Element:
    with zipfile.ZipFile(path) as z:
        return etree.fromstring(z.read(name))


def _specials(root, kind="footnote"):
    return {n.get(f"{W}type"): n for n in root.findall(f"{W}{kind}") if n.get(f"{W}type")}


def test_continuation_notice_registered(doc):
    r = doc.set_note_continuation("footnote", color="520912", bold=True)
    assert r["notice"] == "=" and r["direction"] == "rtl"
    assert r["compatibility_mode"] == 15 and "note" in r
    doc.add_footnote(doc._pids[0], "تخريج")
    doc.save()
    path = doc.source_path
    fn = _part(path, "word/footnotes.xml")
    sp = _specials(fn)
    assert {"separator", "continuationSeparator", "continuationNotice"} <= set(sp)
    notice_id = sp["continuationNotice"].get(f"{W}id")
    assert "".join(sp["continuationNotice"].itertext()) == "="
    st = _part(path, "word/settings.xml")
    ids = [e.get(f"{W}id") for e in st.find(f"{W}footnotePr").findall(f"{W}footnote")]
    assert ids[:2] == ["-1", "0"] and notice_id in ids
    # The real note does not reuse the notice id.
    normal = [n.get(f"{W}id") for n in fn.findall(f"{W}footnote") if not n.get(f"{W}type")]
    assert notice_id not in normal
    # Survives a reopen + save (orphan repair must not drop the notice).
    d2 = DocxDocument(str(path))
    d2.open()
    d2.save()
    assert "continuationNotice" in _specials(_part(path, "word/footnotes.xml"))
    d2.close()


def test_legacy_layout_sets_mode_14(doc):
    r = doc.set_note_continuation("footnote", legacy_layout=True)
    assert r["compatibility_mode"] == 14 and "note" not in r
    doc.save()
    st = _part(doc.source_path, "word/settings.xml")
    modes = [
        c.get(f"{W}val")
        for c in st.iter(f"{W}compatSetting")
        if c.get(f"{W}name") == "compatibilityMode"
    ]
    assert modes == ["14"]


def test_no_notice_removes_it(doc):
    doc.set_note_continuation("footnote")
    doc.set_note_continuation("footnote", notice_text="")
    assert "continuationNotice" not in _specials(doc._trees["word/footnotes.xml"])


@pytest.mark.parametrize("style", ["default", "full", "none", "text"])
def test_separator_styles(doc, style):
    doc.set_note_separator("footnote", style, text="✽ ✽ ✽", color="520912", thickness_pt=1)
    sep = _specials(doc._trees["word/footnotes.xml"])["separator"]
    xml = etree.tostring(sep).decode()
    if style == "default":
        assert "separator/>" in xml
    if style == "full":
        assert 'w:color="520912"' in xml and "pBdr" in xml
    if style == "text":
        assert "✽ ✽ ✽" in "".join(sep.itertext())
    if style == "none":
        assert "<w:r" not in xml and "pBdr" not in xml


def test_separator_bad_style(doc):
    with pytest.raises(ValueError):
        doc.set_note_separator("footnote", "wavy")


def test_numbering_settings_and_sections(doc):
    doc.set_note_numbering("footnote", number_format="arabic-abjad", restart="page", start_at=1)
    st = doc._trees["word/settings.xml"].find(f"{W}footnotePr")
    assert st.find(f"{W}numFmt").get(f"{W}val") == "arabicAbjad"
    assert st.find(f"{W}numRestart").get(f"{W}val") == "eachPage"
    sect = doc._trees["word/document.xml"].find(f".//{W}sectPr/{W}footnotePr")
    assert sect.find(f"{W}numFmt").get(f"{W}val") == "arabicAbjad"
    with pytest.raises(ValueError):
        doc.set_note_numbering("endnote", restart="page")
    with pytest.raises(ValueError):
        doc.set_note_numbering("footnote", number_format="klingon")


def test_custom_mark_note(doc):
    r = doc.add_custom_mark_note(
        doc._pids[1],
        "رواه مسلم.",
        "[١]",
        where="start",
        mark_in_body={"bold": True, "color": "520912"},
    )
    body = doc._trees["word/document.xml"]
    refs = [e for e in body.iter(f"{W}footnoteReference") if e.get(f"{W}id") == str(r["id"])]
    assert refs and refs[0].get(f"{W}customMarkFollows") == "1"
    note = [
        n
        for n in doc._trees["word/footnotes.xml"].findall(f"{W}footnote")
        if n.get(f"{W}id") == str(r["id"])
    ][0]
    assert "".join(note.itertext()).startswith("[١]")
    doc.save()


def test_custom_mark_after_text(doc):
    doc.add_custom_mark_note(doc._pids[0], "ن", "*", where="after", after_text="الله")
    with pytest.raises(ValueError):
        doc.add_custom_mark_note(doc._pids[0], "ن", "*", where="after", after_text="غير موجود")


def test_convert_round_trip(doc):
    doc.add_footnote(doc._pids[0], "الأولى")
    doc.add_footnote(doc._pids[1], "الثانية")
    r = doc.convert_notes("footnote")
    assert r["moved"] == 2
    body = doc._trees["word/document.xml"]
    assert not list(body.iter(f"{W}footnoteReference"))
    assert len(list(body.iter(f"{W}endnoteReference"))) == 2
    doc.save()
    back = doc.convert_notes("endnote")
    assert back["moved"] == 2
    assert len(list(body.iter(f"{W}footnoteReference"))) == 2
    doc.save()
