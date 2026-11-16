"""PDF export: engines, print/digital profiles, page boxes, numerals."""

from __future__ import annotations

from pathlib import Path
from unittest.mock import MagicMock, patch

import pytest
from pypdf import PdfReader, PdfWriter

from docx_mcp.document import DocxDocument, pdfexport
from docx_mcp.document.pdfexport import MM, finish_pdf


def _make_doc(tmp_path: Path) -> DocxDocument:
    return DocxDocument.create(str(tmp_path / "test.docx"))


def _blank_pdf(path: Path, pages: int = 2, w: float = 425.2, h: float = 637.8) -> Path:
    wr = PdfWriter()
    for _ in range(pages):
        wr.add_blank_page(width=w, height=h)
    with open(path, "wb") as fh:
        wr.write(fh)
    return path


def _fake_lo(docx, pdf, *, bookmarks, pdfa):
    _blank_pdf(Path(pdf))


class TestFinishPdf:
    def test_print_boxes_bleed_and_marks(self, tmp_path):
        p = _blank_pdf(tmp_path / "a.pdf")
        info = finish_pdf(p, profile="print", bleed_mm=3, crop_marks=True, title="T")
        assert info["crop_marks"] == 2
        page = PdfReader(str(p)).pages[0]
        trim = [float(v) for v in page.trimbox]
        bleed = [float(v) for v in page.bleedbox]
        media = [float(v) for v in page.mediabox]
        assert trim == pytest.approx([0, 0, 425.2, 637.8])
        assert bleed[0] == pytest.approx(-3 * MM)
        assert media[0] < bleed[0]  # room for the crop marks outside the bleed
        reader = PdfReader(str(p))
        assert reader.trailer["/Root"]["/ViewerPreferences"]["/PrintScaling"] == "/None"
        assert reader.metadata.title == "T"

    def test_print_without_marks_keeps_media_at_bleed(self, tmp_path):
        p = _blank_pdf(tmp_path / "a.pdf", pages=1)
        finish_pdf(p, profile="print", bleed_mm=5, crop_marks=False)
        page = PdfReader(str(p)).pages[0]
        assert float(page.mediabox[0]) == pytest.approx(-5 * MM)
        assert float(page.trimbox[0]) == pytest.approx(0)

    def test_digital_rtl_spreads(self, tmp_path):
        p = _blank_pdf(tmp_path / "a.pdf")
        finish_pdf(p, profile="digital", rtl=True)
        root = PdfReader(str(p)).trailer["/Root"]
        assert root["/PageMode"] == "/UseOutlines"
        assert root["/PageLayout"] == "/TwoPageRight"
        assert root["/ViewerPreferences"]["/Direction"] == "/R2L"
        assert "/TrimBox" not in PdfReader(str(p)).pages[0]

    def test_digital_single_page_ltr(self, tmp_path):
        p = _blank_pdf(tmp_path / "a.pdf")
        finish_pdf(p, profile="digital", two_page=False)
        root = PdfReader(str(p)).trailer["/Root"]
        assert root["/PageLayout"] == "/SinglePage"
        assert "/Direction" not in root["/ViewerPreferences"]


class TestExportPdf:
    def test_libreoffice_fallback_print(self, tmp_path):
        doc = _make_doc(tmp_path)
        out = tmp_path / "out.pdf"
        with (
            patch.object(pdfexport, "_powershell", return_value=None),
            patch.object(pdfexport, "_libreoffice_export", side_effect=_fake_lo),
        ):
            r = doc.export_pdf(str(out), "print")
        assert r["engine"] == "libreoffice"
        assert r["bleed_mm"] == 3.0 and r["crop_marks"] == 2
        assert out.exists()

    def test_word_engine_gets_numerals(self, tmp_path):
        doc = _make_doc(tmp_path)
        doc.set_custom_property("docx-mcp.numerals", "hindi")
        out = tmp_path / "out.pdf"
        seen = {}

        def fake_word(docx, pdf, **kw):
            seen.update(kw)
            _blank_pdf(Path(pdf))
            return 2

        with (
            patch.object(pdfexport, "_powershell", return_value="powershell"),
            patch.object(pdfexport, "_word_export", side_effect=fake_word),
        ):
            r = doc.export_pdf(str(out), "digital")
        assert r["engine"] == "word" and r["numerals"] == "hindi"
        assert seen["numerals"] == "hindi" and seen["optimize_print"] is False

    def test_word_gets_diacritic_colour(self, tmp_path):
        doc = _make_doc(tmp_path)
        doc.color_tashkil("B08D57")
        seen = {}

        def fake_word(docx, pdf, **kw):
            seen.update(kw)
            _blank_pdf(Path(pdf))
            return 1

        with (
            patch.object(pdfexport, "_powershell", return_value="powershell"),
            patch.object(pdfexport, "_word_export", side_effect=fake_word),
        ):
            doc.export_pdf(str(tmp_path / "o.pdf"))
        assert seen["diacritic_color"] == "B08D57"
        assert pdfexport._bgr("B08D57") == 0x578DB0 and pdfexport._bgr(None) == -1

    def test_default_numerals_follow_direction(self, tmp_path):
        doc = _make_doc(tmp_path)
        assert doc._pdf_numerals() == "context"
        doc.set_document_direction("rtl")
        assert doc._pdf_numerals() == "hindi"

    def test_word_failure_falls_back_on_auto(self, tmp_path):
        doc = _make_doc(tmp_path)
        out = tmp_path / "out.pdf"
        with (
            patch.object(pdfexport, "_powershell", return_value="powershell"),
            patch.object(pdfexport, "_word_export", side_effect=RuntimeError("no word")),
            patch.object(pdfexport, "_libreoffice_export", side_effect=_fake_lo),
        ):
            r = doc.export_pdf(str(out))
        assert r["engine"] == "libreoffice" and "no word" in r["word_error"]

    def test_word_engine_required_off_windows(self, tmp_path):
        doc = _make_doc(tmp_path)
        with (
            patch.object(pdfexport, "_powershell", return_value=None),
            pytest.raises(RuntimeError, match="Word"),
        ):
            doc.export_pdf(str(tmp_path / "o.pdf"), engine="word")

    def test_bad_arguments(self, tmp_path):
        doc = _make_doc(tmp_path)
        with pytest.raises(ValueError):
            doc.export_pdf(str(tmp_path / "o.pdf"), "poster")
        with pytest.raises(ValueError):
            doc.export_pdf(str(tmp_path / "o.pdf"), numerals="roman")

    def test_raises_if_no_tools(self, tmp_path):
        doc = _make_doc(tmp_path)
        with (
            patch.object(pdfexport, "_powershell", return_value=None),
            patch("shutil.which", return_value=None),
            pytest.raises(RuntimeError, match="LibreOffice"),
        ):
            doc.export_pdf(str(tmp_path / "o.pdf"))

    def test_libreoffice_failure_reported(self, tmp_path):
        doc = _make_doc(tmp_path)
        with (
            patch.object(pdfexport, "_powershell", return_value=None),
            patch("shutil.which", return_value="/usr/bin/soffice"),
            patch("subprocess.run", return_value=MagicMock(returncode=1, stderr="boom")),
            pytest.raises(RuntimeError, match="LibreOffice export failed"),
        ):
            doc.export_pdf(str(tmp_path / "o.pdf"))

    def test_convert_to_pdf_wrapper(self, tmp_path):
        doc = _make_doc(tmp_path)
        with (
            patch.object(pdfexport, "_powershell", return_value=None),
            patch.object(pdfexport, "_libreoffice_export", side_effect=_fake_lo),
        ):
            r = doc.convert_to_pdf(str(tmp_path / "o.pdf"))
        assert set(r) == {"pdf_path", "engine"}

    def test_raises_if_no_document_open(self):
        doc = DocxDocument.__new__(DocxDocument)
        doc.workdir = None
        with pytest.raises(RuntimeError, match="No document"):
            doc.convert_to_pdf("x.pdf")


class TestPdfPreferences:
    def test_store_and_apply(self, tmp_path):
        doc = _make_doc(tmp_path)
        pref = doc.set_pdf_preferences(
            bleed_mm=5, crop_marks=False, two_page=False, numerals="arabic"
        )
        assert pref == {"bleed_mm": 5.0, "crop_marks": False, "pdfa": None,
                        "two_page": False, "numerals": "arabic"}  # fmt: skip
        out = tmp_path / "o.pdf"
        with (
            patch.object(pdfexport, "_powershell", return_value=None),
            patch.object(pdfexport, "_libreoffice_export", side_effect=_fake_lo),
        ):
            r = doc.export_pdf(str(out), "print")
        assert r["bleed_mm"] == 5.0 and r["crop_marks"] == 0
        assert float(PdfReader(str(out)).pages[0].mediabox[0]) == pytest.approx(-5 * MM)

    def test_digital_ignores_bleed(self, tmp_path):
        doc = _make_doc(tmp_path)
        doc.set_pdf_preferences(bleed_mm=5)
        with (
            patch.object(pdfexport, "_powershell", return_value=None),
            patch.object(pdfexport, "_libreoffice_export", side_effect=_fake_lo),
        ):
            r = doc.export_pdf(str(tmp_path / "o.pdf"), "digital")
        assert r["bleed_mm"] == 0 and r["crop_marks"] == 0

    def test_bad_values(self, tmp_path):
        doc = _make_doc(tmp_path)
        with pytest.raises(ValueError):
            doc.set_pdf_preferences(bleed_mm=50)
        with pytest.raises(ValueError):
            doc.set_pdf_preferences(numerals="roman")

    def test_preset_sets_numerals(self, tmp_path):
        doc = _make_doc(tmp_path)
        doc.apply_book_preset("english")
        assert doc._pdf_numerals() == "context"
        doc2 = DocxDocument.create(str(tmp_path / "ar.docx"))
        doc2.apply_book_preset("tahqiq")
        assert doc2._pdf_numerals() == "hindi"
