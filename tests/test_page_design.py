"""Header/footer designer, page numbering, book layout."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from lxml import etree

from docx_mcp import server
from docx_mcp.document import DocxDocument

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"


def _new(tmp_path: Path) -> DocxDocument:
    return DocxDocument.create(str(tmp_path / "t.docx"))


def _sect(doc):
    return doc._trees["word/document.xml"].find(f"{W}body/{W}sectPr")


def _saved(tmp_path, name):
    with zipfile.ZipFile(tmp_path / "t.docx") as zf:
        return etree.fromstring(zf.read(name))


class TestHeaderFooter:
    def test_three_zones_with_fields_and_rule(self, tmp_path):
        doc = _new(tmp_path)
        res = doc.design_header_footer(
            "header",
            start="Book title",
            center="{STYLEREF:Heading 1}",
            end="{PAGE}",
            rule="below",
            rule_color="520912",
            rule_style="double",
        )
        part = res["parts"][0]["part"]
        doc.save()
        root = _saved(tmp_path, part)
        instr = [t.text.strip() for t in root.iter(f"{W}instrText")]
        assert instr == ['STYLEREF "Heading 1"', "PAGE"]
        assert len(root.findall(f".//{W}tabs/{W}tab")) == 2
        bottom = root.find(f".//{W}pBdr/{W}bottom")
        assert bottom.get(f"{W}val") == "double" and bottom.get(f"{W}color") == "520912"
        refs = _saved(tmp_path, "word/document.xml").findall(f".//{W}sectPr/{W}headerReference")
        rels = _saved(tmp_path, "word/_rels/document.xml.rels")
        targets = {r.get("Id"): r.get("Target") for r in rels}
        assert targets[refs[0].get(f"{R}id")] == part.split("/")[-1]

    def test_rtl_direction_auto(self, tmp_path):
        doc = _new(tmp_path)
        res = doc.design_header_footer("footer", center="الصفحة {PAGE}")
        assert res["direction"] == "rtl"

    def test_first_page_variant_turns_on_title_page(self, tmp_path):
        doc = _new(tmp_path)
        doc.design_header_footer("header", variant="first", center="")
        assert _sect(doc).find(f"{W}titlePg") is not None

    def test_even_variant_and_gap_filling(self, tmp_path):
        doc = _new(tmp_path)
        doc.design_header_footer("footer", center="{PAGE}")
        res = doc.design_header_footer("header", variant="even", start="Title")
        assert doc._trees["word/settings.xml"].find(f"{W}evenAndOddHeaders") is not None
        assert res["even_pages_reuse_default"] == ["section 0 footer"]
        kinds = {
            (r.tag.split("}")[1], r.get(f"{W}type"))
            for r in _sect(doc)
            if r.tag.endswith("Reference")
        }
        assert ("footerReference", "even") in kinds

    def test_redesign_replaces_reference(self, tmp_path):
        doc = _new(tmp_path)
        doc.design_header_footer("header", center="A")
        doc.design_header_footer("header", center="B")
        defaults = [
            r for r in _sect(doc).findall(f"{W}headerReference") if r.get(f"{W}type") == "default"
        ]
        assert len(defaults) == 1

    def test_remove(self, tmp_path):
        doc = _new(tmp_path)
        doc.design_header_footer("footer", center="{PAGE}")
        assert doc.remove_header_footer("footer")["removed"] == 1

    def test_bad_arguments(self, tmp_path):
        doc = _new(tmp_path)
        with pytest.raises(ValueError):
            doc.design_header_footer("sidebar")
        with pytest.raises(ValueError):
            doc.design_header_footer("header", variant="odd")


class TestNumberingAndLayout:
    def test_page_numbering(self, tmp_path):
        doc = _new(tmp_path)
        doc.set_page_numbering("arabic-abjad", start_at=1)
        pn = _sect(doc).find(f"{W}pgNumType")
        assert pn.get(f"{W}fmt") == "arabicAbjad" and pn.get(f"{W}start") == "1"
        with pytest.raises(ValueError):
            doc.set_page_numbering("hindi")

    def test_book_layout(self, tmp_path):
        doc = _new(tmp_path)
        doc.set_book_layout(
            page_size="17x24",
            inside_cm=2,
            outside_cm=1.5,
            gutter_cm=0.5,
            mirror_margins=True,
            rtl_binding=True,
        )
        doc.save()
        sect = _saved(tmp_path, "word/document.xml").find(f".//{W}body/{W}sectPr")
        pg = sect.find(f"{W}pgSz")
        assert (pg.get(f"{W}w"), pg.get(f"{W}h")) == ("9638", "13606")
        assert sect.find(f"{W}rtlGutter") is not None
        mar = sect.find(f"{W}pgMar")
        assert mar.get(f"{W}left") == "1134" and mar.get(f"{W}gutter") == "283"
        settings = _saved(tmp_path, "word/settings.xml")
        assert settings.find(f"{W}mirrorMargins") is not None
        names = [etree.QName(c).localname for c in settings]
        assert names[0] == "mirrorMargins" or names.index("mirrorMargins") < names.index("compat")

    def test_update_fields_on_open(self, tmp_path):
        doc = _new(tmp_path)
        doc.set_update_fields_on_open(True)
        assert doc._trees["word/settings.xml"].find(f"{W}updateFields") is not None

    def test_server_tools(self, tmp_path):
        server.create_document(str(tmp_path / "s.docx"))
        assert '"parts"' in server.design_header_footer("footer", center="{PAGE}")
        assert "17x24" not in server.set_book_layout(page_size="A5")
