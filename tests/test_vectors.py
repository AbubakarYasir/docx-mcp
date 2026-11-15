"""Vector art: SVG with PNG fallback, ornaments, page frames, page borders."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from lxml import etree

from docx_mcp.document import DocxDocument
from docx_mcp.document.vectors import ORNAMENTS, ornament_svg, render_png

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
ASVG = "{http://schemas.microsoft.com/office/drawing/2016/SVG/main}"


@pytest.fixture
def doc(tmp_path: Path):
    d = DocxDocument.create(str(tmp_path / "v.docx"))
    d._first = d._trees["word/document.xml"].find(f".//{W}p").get(f"{W14}paraId")
    yield d
    d.close()


@pytest.mark.parametrize("kind", ORNAMENTS)
def test_ornaments_are_valid_svg_and_render(kind):
    svg = ornament_svg(kind, color="520912", aspect=8 if kind in ("divider", "band") else None)
    root = etree.fromstring(svg.encode())
    assert root.tag == "{http://www.w3.org/2000/svg}svg"
    png = render_png(svg, 120)
    assert png[:8] == b"\x89PNG\r\n\x1a\n"


def test_bad_ornament():
    with pytest.raises(ValueError):
        ornament_svg("dragon")


def test_insert_ornament_svg_and_fallback(doc):
    r = doc.insert_ornament("divider", after_para_id=doc._first, accent="E9D9B8")
    assert r["width_cm"] > 5
    doc.save()
    with zipfile.ZipFile(doc.source_path) as z:
        names = z.namelist()
        assert any(n.endswith(".svg") for n in names)
        assert any(n.startswith("word/media/vector") and n.endswith(".png") for n in names)
        body = etree.fromstring(z.read("word/document.xml"))
        ct = z.read("[Content_Types].xml").decode()
        rels = z.read("word/_rels/document.xml.rels").decode()
    blip = next(body.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}blip"))
    svg_blip = next(blip.iter(f"{ASVG}svgBlip"))
    rid = svg_blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
    assert f'Id="{rid}"' in rels and 'Extension="svg"' in ct


def test_insert_svg_markup_keeps_aspect(doc):
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 50">'
        '<rect width="200" height="50"/></svg>'
    )
    r = doc.insert_svg(svg, width_cm=8)
    assert r["height_cm"] == 2.0
    with pytest.raises(ValueError):
        doc.insert_svg("<svg")
    with pytest.raises(FileNotFoundError):
        doc.insert_svg("missing.svg")


def test_page_frame_lives_in_header_behind_text(doc):
    doc.design_header_footer("header", center="عنوان")
    r = doc.add_page_frame(margin_cm=1)
    assert r["frame_in"]
    doc.save()
    with zipfile.ZipFile(doc.source_path) as z:
        hdr = etree.fromstring(z.read(r["frame_in"][0]))
        rels_name = r["frame_in"][0].replace("word/", "word/_rels/") + ".rels"
        assert rels_name in z.namelist()
    anchor = next(
        hdr.iter("{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}anchor")
    )
    assert anchor.get("behindDoc") == "1"
    assert "عنوان" in "".join(hdr.itertext())  # existing header kept


def test_page_frame_creates_header_when_missing(doc):
    r = doc.add_page_frame()
    assert len(r["frame_in"]) == 1
    sect = doc._trees["word/document.xml"].find(f".//{W}sectPr")
    assert sect.find(f"{W}headerReference") is not None
    doc.save()


@pytest.mark.parametrize("style,sz", [("double", "12"), ("art-celtic", "12"), ("thick-thin", "12")])
def test_page_border(doc, style, sz):
    doc.set_page_border(style, width_pt=12 if style.startswith("art") else 1.5)
    pb = doc._trees["word/document.xml"].find(f".//{W}pgBorders")
    sides = [c.tag.split("}")[1] for c in pb]
    assert sides == ["top", "left", "bottom", "right"]
    assert pb.find(f"{W}top").get(f"{W}sz") == sz
    doc.set_page_border("none")
    assert doc._trees["word/document.xml"].find(f".//{W}pgBorders") is None


def test_page_border_errors(doc):
    with pytest.raises(ValueError):
        doc.set_page_border("rainbow")
    with pytest.raises(ValueError):
        doc.set_page_border("double", pages="odd")
