"""Regression tests: output that Microsoft Word must accept.

Each test pins one schema rule that docx-mcp used to break. They were found by
validating real output against the ECMA-376 XSDs.
"""

from __future__ import annotations

import struct
import zipfile
import zlib
from pathlib import Path

from lxml import etree

from docx_mcp.document import DocxDocument
from docx_mcp.document.schema_order import normalize_tree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
PIC = "{http://schemas.openxmlformats.org/drawingml/2006/picture}"
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"


def _new(tmp_path: Path) -> tuple[DocxDocument, str]:
    doc = DocxDocument.create(str(tmp_path / "t.docx"))
    first = doc._trees["word/document.xml"].find(f".//{W}p").get(f"{W14}paraId")
    para = doc.insert_paragraph(first, "Hello world")["para_id"]
    return doc, para


def _part(path: Path, name: str) -> etree._Element:
    with zipfile.ZipFile(path) as zf:
        return etree.fromstring(zf.read(name))


def _png(path: Path, w: int, h: int) -> Path:
    def chunk(kind: bytes, data: bytes) -> bytes:
        return (
            struct.pack(">I", len(data))
            + kind
            + data
            + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)
        )

    raw = b"".join(b"\x00" + b"\x00\x00\x00" * w for _ in range(h))
    path.write_bytes(
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(raw))
        + chunk(b"IEND", b"")
    )
    return path


def _tags(el: etree._Element) -> list[str]:
    return [etree.QName(c).localname for c in el]


class TestSchemaOrder:
    def test_ppr_children_sorted(self):
        ppr = etree.fromstring(
            f'<w:pPr xmlns:w="{W[1:-1]}"><w:jc w:val="left"/><w:keepNext/>'
            f'<w:spacing w:after="0"/><w:pStyle w:val="Normal"/></w:pPr>'
        )
        normalize_tree(ppr)
        assert _tags(ppr) == ["pStyle", "keepNext", "spacing", "jc"]

    def test_rpr_children_sorted(self):
        rpr = etree.fromstring(
            f'<w:rPr xmlns:w="{W[1:-1]}"><w:sz w:val="24"/><w:rtl/>'
            f'<w:rFonts w:ascii="A"/><w:b/></w:rPr>'
        )
        normalize_tree(rpr)
        assert _tags(rpr) == ["rFonts", "b", "sz", "rtl"]

    def test_unknown_elements_keep_position(self):
        rpr = etree.fromstring(
            f'<w:rPr xmlns:w="{W[1:-1]}" xmlns:x="urn:x"><w:b/><x:ext/><w:sz w:val="2"/>'
            f'<w:rFonts w:ascii="A"/></w:rPr>'
        )
        normalize_tree(rpr)
        assert _tags(rpr) == ["rFonts", "b", "ext", "sz"]

    def test_ppr_moved_to_front_of_paragraph(self):
        p = etree.fromstring(
            f'<w:p xmlns:w="{W[1:-1]}"><w:r><w:t>x</w:t></w:r>'
            f'<w:pPr><w:jc w:val="left"/></w:pPr></w:p>'
        )
        normalize_tree(p)
        assert _tags(p) == ["pPr", "r"]

    def test_numbering_abstract_before_num(self):
        root = etree.fromstring(
            f'<w:numbering xmlns:w="{W[1:-1]}"><w:num w:numId="1"/>'
            f'<w:abstractNum w:abstractNumId="0"/></w:numbering>'
        )
        normalize_tree(root)
        assert _tags(root) == ["abstractNum", "num"]

    def test_formatting_calls_save_in_order(self, tmp_path):
        doc, para = _new(tmp_path)
        doc.set_keep_with_next(para, True)
        doc.set_line_spacing(para, line_rule="auto", line_value=360)
        doc.set_paragraph_shading(para, "EEEEEE")
        doc.set_run_size(para, 0, 14)
        doc.set_run_font(para, 0, "Amiri")
        doc.save()
        body = _part(tmp_path / "t.docx", "word/document.xml")
        p = next(x for x in body.iter(f"{W}p") if x.get(f"{W14}paraId") == para)
        assert _tags(p.find(f"{W}pPr")) == ["keepNext", "shd", "spacing"]
        assert _tags(p.find(f"{W}r/{W}rPr")) == ["rFonts", "sz", "szCs"]


class TestBlankDocument:
    def test_has_a4_section(self, tmp_path):
        doc, _ = _new(tmp_path)
        doc.save()
        sect = _part(tmp_path / "t.docx", "word/document.xml").find(f"{W}body/{W}sectPr")
        assert sect is not None
        pg_sz = sect.find(f"{W}pgSz")
        assert (pg_sz.get(f"{W}w"), pg_sz.get(f"{W}h")) == ("11906", "16838")

    def test_styles_valid_order(self, tmp_path):
        doc, _ = _new(tmp_path)
        doc.save()
        styles = _part(tmp_path / "t.docx", "word/styles.xml")
        for ppr in styles.iter(f"{W}pPr"):
            tags = _tags(ppr)
            if "pBdr" in tags and "spacing" in tags:
                assert tags.index("pBdr") < tags.index("spacing")


class TestSections:
    def test_partial_margins_get_required_attributes(self, tmp_path):
        doc, _ = _new(tmp_path)
        doc.set_page_margins(top_mm=20)
        doc.save()
        pg_mar = _part(tmp_path / "t.docx", "word/document.xml").find(f".//{W}pgMar")
        for attr in ("top", "right", "bottom", "left", "header", "footer", "gutter"):
            assert pg_mar.get(f"{W}{attr}") is not None

    def test_landscape_swaps_dimensions(self, tmp_path):
        doc, _ = _new(tmp_path)
        doc.set_section_properties(orientation="landscape")
        pg_sz = doc._trees["word/document.xml"].find(f".//{W}pgSz")
        assert int(pg_sz.get(f"{W}w")) > int(pg_sz.get(f"{W}h"))

    def test_columns_on_new_document(self, tmp_path):
        doc, _ = _new(tmp_path)
        assert doc.set_section_columns(0, 2)["num_columns"] == 2


class TestImages:
    def test_inline_image_complete(self, tmp_path):
        doc, para = _new(tmp_path)
        img = _png(tmp_path / "a.png", 200, 100)
        res = doc.insert_image(para, str(img), alt_text="A red box")
        doc.save()
        body = _part(tmp_path / "t.docx", "word/document.xml")
        inline = body.find(f".//{WP}inline")
        assert _tags(inline) == [
            "extent",
            "effectExtent",
            "docPr",
            "cNvGraphicFramePr",
            "graphic",
        ]
        assert inline.find(f"{WP}docPr").get("descr") == "A red box"
        pic = body.find(f".//{PIC}pic")
        assert _tags(pic) == ["nvPicPr", "blipFill", "spPr"]
        # aspect ratio 2:1 kept
        assert res["width_emu"] == 2 * res["height_emu"]

    def test_width_only_keeps_aspect(self, tmp_path):
        doc, para = _new(tmp_path)
        img = _png(tmp_path / "a.png", 300, 100)
        res = doc.insert_image(para, str(img), width_emu=3 * 360000)
        assert res["height_emu"] == 360000

    def test_second_image_does_not_overwrite_first(self, tmp_path):
        doc, para = _new(tmp_path)
        a = doc.insert_image(para, str(_png(tmp_path / "a.png", 4, 4)))
        b = doc.insert_image(para, str(_png(tmp_path / "b.png", 8, 8)))
        assert a["filename"] != b["filename"]
        assert a["rId"] != b["rId"]

    def test_floating_image_has_nvpicpr(self, tmp_path):
        doc, para = _new(tmp_path)
        doc.insert_floating_image(para, str(_png(tmp_path / "a.png", 4, 4)), 2, 2)
        pic = doc._trees["word/document.xml"].find(f".//{PIC}pic")
        assert pic.find(f"{PIC}nvPicPr") is not None


class TestTrackedTables:
    def test_no_ins_in_tblpr(self, tmp_path):
        doc, para = _new(tmp_path)
        doc.add_table(para, 2, 2)
        tbl = doc._trees["word/document.xml"].find(f".//{W}tbl")
        assert tbl.find(f"{W}tblPr/{W}ins") is None
        assert all(tr.find(f"{W}trPr/{W}ins") is not None for tr in tbl.findall(f"{W}tr"))

    def test_reject_removes_table(self, tmp_path):
        doc, para = _new(tmp_path)
        doc.add_table(para, 2, 2)
        doc.reject_changes()
        assert doc._trees["word/document.xml"].find(f".//{W}tbl") is None

    def test_accept_keeps_table_without_markers(self, tmp_path):
        doc, para = _new(tmp_path)
        doc.add_table(para, 2, 2)
        doc.accept_all_changes()
        root = doc._trees["word/document.xml"]
        assert root.find(f".//{W}tbl") is not None
        assert root.find(f".//{W}ins") is None

    def test_reject_all_is_saved(self, tmp_path):
        doc, para = _new(tmp_path)
        doc.insert_text(para, " extra")
        doc.save()
        doc.reject_all_changes()
        doc.save()
        reopened = DocxDocument(str(tmp_path / "t.docx"))
        reopened.open()
        assert "extra" not in reopened.get_body_text()["body"]


class TestCharts:
    def test_bar_chart_has_axes(self, tmp_path):
        doc, para = _new(tmp_path)
        res = doc.insert_bar_chart(para, "T", [{"name": "s", "values": [1, 2]}], ["a", "b"])
        doc.save()
        chart = _part(tmp_path / "t.docx", res["part_path"])
        bar = chart.find(f".//{C}barChart")
        assert len(bar.findall(f"{C}axId")) == 2
        assert chart.find(f".//{C}catAx") is not None
        assert chart.find(f".//{C}valAx") is not None

    def test_pie_series_before_first_slice_angle(self, tmp_path):
        doc, para = _new(tmp_path)
        res = doc.insert_pie_chart(para, "T", [{"name": "s", "values": [1, 2]}], ["a", "b"])
        doc.update_chart_data(res["chart_id"], [{"name": "n", "values": [3, 4]}])
        doc.save()
        pie = _part(tmp_path / "t.docx", res["part_path"]).find(f".//{C}pieChart")
        tags = _tags(pie)
        assert tags.index("ser") < tags.index("firstSliceAng")


class TestPackageParts:
    def test_custom_properties_part_is_linked(self, tmp_path):
        doc, _ = _new(tmp_path)
        doc.set_custom_property("Project", "Risalah")
        doc.save()
        with zipfile.ZipFile(tmp_path / "t.docx") as zf:
            rels = zf.read("_rels/.rels").decode()
            ct = zf.read("[Content_Types].xml").decode()
        assert "docProps/custom.xml" in rels
        assert "/docProps/custom.xml" in ct

    def test_new_document_not_in_compatibility_mode(self, tmp_path):
        doc, _ = _new(tmp_path)
        doc.save()
        settings = _part(tmp_path / "t.docx", "word/settings.xml")
        modes = [
            c.get(f"{W}val")
            for c in settings.iter(f"{W}compatSetting")
            if c.get(f"{W}name") == "compatibilityMode"
        ]
        assert modes == ["15"]
