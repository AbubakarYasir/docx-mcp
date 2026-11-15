"""Right-to-left (Arabic) support: automatic markup on save and explicit tools."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from lxml import etree

from docx_mcp import server
from docx_mcp.document import DocxDocument
from docx_mcp.document.bidi import first_strong, segment

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"

ARABIC = "قال رسول الله ﷺ: إنما الأعمال بالنيات"


def _new(tmp_path: Path) -> tuple[DocxDocument, str]:
    doc = DocxDocument.create(str(tmp_path / "t.docx"))
    first = doc._trees["word/document.xml"].find(f".//{W}p").get(f"{W14}paraId")
    return doc, first


def _saved_para(tmp_path: Path, para_id: str, part: str = "word/document.xml"):
    with zipfile.ZipFile(tmp_path / "t.docx") as zf:
        root = etree.fromstring(zf.read(part))
    return next(p for p in root.iter(f"{W}p") if p.get(f"{W14}paraId") == para_id)


def _runs(p):
    return [r for r in p.iter(f"{W}r") if r.find(f"{W}t") is not None]


class TestHelpers:
    def test_first_strong(self):
        assert first_strong("  (1) قال") == "rtl"
        assert first_strong("1. Hello قال") == "ltr"
        assert first_strong("123 ...") is None

    def test_segment_mixed(self):
        parts = segment("the word النية means", "l")
        assert [d for d, _ in parts] == ["l", "r", "l"]
        assert parts[1][1].strip() == "النية"

    def test_segment_digits_inside_arabic_stay_arabic(self):
        assert [d for d, _ in segment("رواه البخاري (1) ومسلم", "r")] == ["r"]


class TestAutomatic:
    def test_new_arabic_paragraph_is_rtl(self, tmp_path):
        doc, first = _new(tmp_path)
        pid = doc.insert_paragraph(first, ARABIC)["para_id"]
        doc.save()
        p = _saved_para(tmp_path, pid)
        assert p.find(f"{W}pPr/{W}bidi") is not None
        for r in _runs(p):
            assert r.find(f"{W}rPr/{W}rtl") is not None
            assert r.find(f"{W}rPr/{W}lang").get(f"{W}bidi") == "ar-SA"

    def test_english_paragraph_untouched(self, tmp_path):
        doc, first = _new(tmp_path)
        pid = doc.insert_paragraph(first, "Plain English text")["para_id"]
        doc.save()
        p = _saved_para(tmp_path, pid)
        assert p.find(f"{W}pPr/{W}bidi") is None
        assert all(r.find(f"{W}rPr/{W}rtl") is None for r in _runs(p))

    def test_mixed_run_split_and_marked(self, tmp_path):
        doc, first = _new(tmp_path)
        pid = doc.insert_paragraph(first, "The word النية means intention")["para_id"]
        doc.save()
        p = _saved_para(tmp_path, pid)
        assert p.find(f"{W}pPr/{W}bidi") is None  # starts with a Latin letter
        runs = _runs(p)
        assert len(runs) == 3
        assert [r.find(f"{W}rPr/{W}rtl") is not None for r in runs] == [False, True, False]
        assert "".join(r.find(f"{W}t").text for r in runs) == "The word النية means intention"

    def test_bold_arabic_gets_bcs(self, tmp_path):
        doc, first = _new(tmp_path)
        pid = doc.insert_paragraph(first, ARABIC)["para_id"]
        doc.set_formatting(pid, "إنما الأعمال", bold=True)
        doc.accept_all_changes()
        doc.save()
        p = _saved_para(tmp_path, pid)
        bold = [r for r in _runs(p) if r.find(f"{W}rPr/{W}b") is not None]
        assert bold and all(r.find(f"{W}rPr/{W}bCs") is not None for r in bold)

    def test_text_added_to_existing_rtl_paragraph(self, tmp_path):
        doc, first = _new(tmp_path)
        pid = doc.insert_paragraph(first, ARABIC)["para_id"]
        doc.save()
        doc2 = DocxDocument(str(tmp_path / "t.docx"))
        doc2.open()
        doc2.insert_text(pid, " وهو حديث عظيم", tracked=False)
        doc2.save()
        p = _saved_para(tmp_path, pid)
        assert all(r.find(f"{W}rPr/{W}rtl") is not None for r in _runs(p))

    def test_original_paragraphs_not_redirected(self, tmp_path):
        doc, first = _new(tmp_path)
        pid = doc.insert_paragraph(first, ARABIC)["para_id"]
        doc.save()
        doc.close()
        # Rewrite the file as if authored elsewhere: Arabic paragraph, no bidi.
        src = tmp_path / "t.docx"
        with zipfile.ZipFile(src) as zf:
            items = {n: zf.read(n) for n in zf.namelist()}
        root = etree.fromstring(items["word/document.xml"])
        para = next(p for p in root.iter(f"{W}p") if p.get(f"{W14}paraId") == pid)
        para.remove(para.find(f"{W}pPr"))
        items["word/document.xml"] = etree.tostring(root)
        with zipfile.ZipFile(src, "w") as zf:
            for n, data in items.items():
                zf.writestr(n, data)
        doc2 = DocxDocument(str(src))
        doc2.open()
        doc2.insert_paragraph(pid, "English")
        doc2.save()
        assert _saved_para(tmp_path, pid).find(f"{W}pPr/{W}bidi") is None

    def test_arabic_footnote_is_rtl(self, tmp_path):
        doc, first = _new(tmp_path)
        pid = doc.insert_paragraph(first, ARABIC)["para_id"]
        doc.add_footnote(pid, "أخرجه البخاري (1).")
        doc.save()
        with zipfile.ZipFile(tmp_path / "t.docx") as zf:
            notes = etree.fromstring(zf.read("word/footnotes.xml"))
        note_paras = [p for p in notes.iter(f"{W}p") if "البخاري" in "".join(p.itertext())]
        assert note_paras and note_paras[0].find(f"{W}pPr/{W}bidi") is not None

    def test_heading_styles_bold_arabic(self, tmp_path):
        doc, _ = _new(tmp_path)
        doc.save()
        with zipfile.ZipFile(tmp_path / "t.docx") as zf:
            styles = etree.fromstring(zf.read("word/styles.xml"))
        h1 = styles.find(f"{W}style[@{W}styleId='Heading1']/{W}rPr")
        assert h1.find(f"{W}bCs") is not None and h1.find(f"{W}szCs") is not None

    def test_can_be_disabled(self, tmp_path, monkeypatch):
        monkeypatch.setenv("DOCX_MCP_AUTO_BIDI", "0")
        doc, first = _new(tmp_path)
        pid = doc.insert_paragraph(first, ARABIC)["para_id"]
        doc.save()
        assert _saved_para(tmp_path, pid).find(f"{W}pPr/{W}bidi") is None


class TestTables:
    def test_new_arabic_table_is_rtl(self, tmp_path):
        from docx_mcp import server as srv

        md = "| الراوي | المصدر |\n|---|---|\n| عمر | البخاري |\n"
        srv.create_from_markdown(str(tmp_path / "m.docx"), markdown=md)
        with zipfile.ZipFile(tmp_path / "m.docx") as zf:
            body = etree.fromstring(zf.read("word/document.xml"))
        assert body.find(f".//{W}tblPr/{W}bidiVisual") is not None

    def test_new_english_table_stays_ltr(self, tmp_path):
        doc, first = _new(tmp_path)
        doc.add_table(first, 2, 2)
        doc.modify_cell(0, 0, 0, "Name", tracked=False)
        doc.save()
        assert doc._trees["word/document.xml"].find(f".//{W}bidiVisual") is None


class TestTools:
    def test_set_paragraph_direction(self, tmp_path):
        doc, first = _new(tmp_path)
        pid = doc.insert_paragraph(first, "Hello")["para_id"]
        doc.set_paragraph_direction(pid, "rtl")
        assert doc._find_para_any(pid)[1].find(f"{W}pPr/{W}bidi") is not None
        doc.set_paragraph_direction(pid, "ltr")
        assert doc._find_para_any(pid)[1].find(f"{W}pPr/{W}bidi").get(f"{W}val") == "0"
        assert doc.set_paragraph_direction(pid, "auto")["direction"] == "ltr"
        with pytest.raises(ValueError):
            doc.set_paragraph_direction(pid, "sideways")

    def test_set_document_direction(self, tmp_path):
        doc, first = _new(tmp_path)
        eng = doc.insert_paragraph(first, "English stays left")["para_id"]
        doc.add_table(eng, 2, 2)
        res = doc.set_document_direction(
            "rtl", complex_script_font="Amiri", complex_script_size_pt=16
        )
        assert res["tables"] == 1 and res["ltr_paragraphs_kept"] >= 1
        doc.save()
        with zipfile.ZipFile(tmp_path / "t.docx") as zf:
            styles = etree.fromstring(zf.read("word/styles.xml"))
            body = etree.fromstring(zf.read("word/document.xml"))
        drpr = styles.find(f"{W}docDefaults/{W}rPrDefault/{W}rPr")
        assert drpr.find(f"{W}rFonts").get(f"{W}cs") == "Amiri"
        assert drpr.find(f"{W}szCs").get(f"{W}val") == "32"
        assert drpr.find(f"{W}lang").get(f"{W}bidi") == "ar-SA"
        assert styles.find(f"{W}docDefaults/{W}pPrDefault/{W}pPr/{W}bidi") is not None
        assert body.find(f".//{W}sectPr/{W}bidi") is not None
        assert body.find(f".//{W}tblPr/{W}bidiVisual") is not None
        eng_p = next(p for p in body.iter(f"{W}p") if p.get(f"{W14}paraId") == eng)
        assert eng_p.find(f"{W}pPr/{W}bidi").get(f"{W}val") == "0"
        # and back
        doc.set_document_direction("ltr")
        assert doc._trees["word/document.xml"].find(f".//{W}sectPr/{W}bidi") is None

    def test_direction_report(self, tmp_path):
        doc, first = _new(tmp_path)
        doc.insert_paragraph(first, ARABIC)
        before = doc.get_direction_report()
        assert len(before["rtl_text_without_paragraph_direction"]) == 1
        doc.save()
        after = doc.get_direction_report()
        assert after["rtl_text_without_paragraph_direction"] == []
        assert after["rtl_runs_without_rtl_flag"] == 0

    def test_server_tools(self, tmp_path):
        path = tmp_path / "s.docx"
        server.create_document(str(path))
        first = server._doc._trees["word/document.xml"].find(f".//{W}p").get(f"{W14}paraId")
        server.set_paragraph_direction(first, "rtl")
        server.set_document_direction("rtl", complex_script_font="Amiri")
        assert '"rtl_paragraphs"' in server.get_direction_report()


class TestFormattingRevisions:
    def test_accept_removes_rprchange(self, tmp_path):
        doc, first = _new(tmp_path)
        pid = doc.insert_paragraph(first, "Hello world")["para_id"]
        doc.set_formatting(pid, "Hello", bold=True)
        doc.accept_all_changes()
        root = doc._trees["word/document.xml"]
        assert root.find(f".//{W}rPrChange") is None
        assert root.find(f".//{W}b") is not None

    def test_reject_restores_old_formatting(self, tmp_path):
        doc, first = _new(tmp_path)
        pid = doc.insert_paragraph(first, "Hello world")["para_id"]
        doc.set_formatting(pid, "Hello", bold=True)
        doc.reject_all_changes()
        root = doc._trees["word/document.xml"]
        assert root.find(f".//{W}rPrChange") is None
        assert root.find(f".//{W}b") is None
