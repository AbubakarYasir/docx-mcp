"""Images mixin: list and insert embedded images."""

from __future__ import annotations

import contextlib
import shutil
from pathlib import Path

from lxml import etree

from .base import CT, RELS, W14, WP, A, R, W

# Additional namespace constants for floating images
_PIC = "{http://schemas.openxmlformats.org/drawingml/2006/picture}"
_PIC_NS = _PIC[1:-1]  # bare URI (no braces) for attribute values / nsmap keys
_WP_NS = WP[1:-1]  # bare URI derived from imported WP constant
_A_NS = A[1:-1]  # bare URI derived from imported A constant
_R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"  # bare URI for nsmap


# ── Image helpers ───────────────────────────────────────────────────────────

_EMU_PER_INCH = 914400
_EMU_PER_CM = 360000
# Word's default text width on A4/Letter with standard margins is about 16 cm.
_MAX_AUTO_WIDTH_EMU = 16 * _EMU_PER_CM

_CONTENT_TYPES = {
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "jpe": "image/jpeg",
    "gif": "image/gif",
    "bmp": "image/bmp",
    "tif": "image/tiff",
    "tiff": "image/tiff",
    "emf": "image/x-emf",
    "wmf": "image/x-wmf",
}


def _image_size_px(path: Path) -> tuple[int, int, float, float] | None:
    """Return (width_px, height_px, dpi_x, dpi_y) from the file header, or None.

    Supports PNG (with pHYs), JPEG (JFIF density), GIF and BMP without any
    imaging dependency. DPI defaults to 96 when the file does not say.
    """
    import struct

    try:
        data = path.read_bytes()
    except OSError:
        return None
    dpi_x = dpi_y = 96.0
    if data[:8] == b"\x89PNG\r\n\x1a\n" and len(data) >= 24:
        w, h = struct.unpack(">II", data[16:24])
        i = data.find(b"pHYs")
        if i > 0 and len(data) >= i + 13 and data[i + 12] == 1:
            px, py = struct.unpack(">II", data[i + 4 : i + 12])
            if px and py:
                dpi_x, dpi_y = px * 0.0254, py * 0.0254
        return (w, h, dpi_x, dpi_y) if w and h else None
    if data[:6] in (b"GIF87a", b"GIF89a") and len(data) >= 10:
        w, h = struct.unpack("<HH", data[6:10])
        return (w, h, dpi_x, dpi_y) if w and h else None
    if data[:2] == b"BM" and len(data) >= 26:
        w, h = struct.unpack("<ii", data[18:26])
        return (abs(w), abs(h), dpi_x, dpi_y) if w and h else None
    if data[:2] == b"\xff\xd8":
        i = 2
        while i + 9 < len(data):
            if data[i] != 0xFF:
                i += 1
                continue
            marker = data[i + 1]
            if marker in (0xD8, 0x01) or 0xD0 <= marker <= 0xD7:
                i += 2
                continue
            seg_len = struct.unpack(">H", data[i + 2 : i + 4])[0]
            if marker == 0xE0 and data[i + 4 : i + 9] == b"JFIF\x00":
                unit = data[i + 11]
                dx, dy = struct.unpack(">HH", data[i + 12 : i + 16])
                if dx and dy and unit in (1, 2):
                    f = 1.0 if unit == 1 else 2.54
                    dpi_x, dpi_y = dx * f, dy * f
            if marker in (
                0xC0,
                0xC1,
                0xC2,
                0xC3,
                0xC5,
                0xC6,
                0xC7,
                0xC9,
                0xCA,
                0xCB,
                0xCD,
                0xCE,
                0xCF,
            ):
                h, w = struct.unpack(">HH", data[i + 5 : i + 9])
                return (w, h, dpi_x, dpi_y) if w and h else None
            i += 2 + seg_len
    return None


def _auto_extent(path: Path, width_emu: int, height_emu: int) -> tuple[int, int]:
    """Fill in a missing width or height so the image keeps its aspect ratio.

    Both given: used as-is. One given: the other follows the image's aspect
    ratio. Neither given: native size at the file's DPI, scaled down to fit a
    16 cm text width. Unknown formats fall back to 5 cm square.
    """
    if width_emu > 0 and height_emu > 0:
        return width_emu, height_emu
    size = _image_size_px(path)
    if size is None:
        side = 5 * _EMU_PER_CM
        return (width_emu or height_emu or side), (height_emu or width_emu or side)
    w_px, h_px, dpi_x, dpi_y = size
    native_w = int(w_px / dpi_x * _EMU_PER_INCH)
    native_h = int(h_px / dpi_y * _EMU_PER_INCH)
    if width_emu > 0:
        return width_emu, max(1, int(width_emu * native_h / native_w))
    if height_emu > 0:
        return max(1, int(height_emu * native_w / native_h)), height_emu
    if native_w > _MAX_AUTO_WIDTH_EMU:
        native_h = int(native_h * _MAX_AUTO_WIDTH_EMU / native_w)
        native_w = _MAX_AUTO_WIDTH_EMU
    return max(1, native_w), max(1, native_h)


def _build_pic(rid: str, cx: int, cy: int, pic_id: int, name: str, descr: str) -> etree._Element:
    """Build a complete, schema-valid pic:pic element."""
    pic = etree.Element(f"{_PIC}pic", nsmap={"pic": _PIC_NS})
    nv = etree.SubElement(pic, f"{_PIC}nvPicPr")
    cnv = etree.SubElement(nv, f"{_PIC}cNvPr", id=str(pic_id), name=name)
    if descr:
        cnv.set("descr", descr)
    cnv_pic = etree.SubElement(nv, f"{_PIC}cNvPicPr")
    etree.SubElement(cnv_pic, f"{A}picLocks", noChangeAspect="1", noChangeArrowheads="1")
    blip_fill = etree.SubElement(pic, f"{_PIC}blipFill")
    blip = etree.SubElement(blip_fill, f"{A}blip")
    blip.set(f"{R}embed", rid)
    stretch = etree.SubElement(blip_fill, f"{A}stretch")
    etree.SubElement(stretch, f"{A}fillRect")
    sp_pr = etree.SubElement(pic, f"{_PIC}spPr", bwMode="auto")
    xfrm = etree.SubElement(sp_pr, f"{A}xfrm")
    etree.SubElement(xfrm, f"{A}off", x="0", y="0")
    etree.SubElement(xfrm, f"{A}ext", cx=str(cx), cy=str(cy))
    geom = etree.SubElement(sp_pr, f"{A}prstGeom", prst="rect")
    etree.SubElement(geom, f"{A}avLst")
    return pic


class ImagesMixin:
    """Image operations."""

    def get_images(self) -> list[dict]:
        """Get all embedded images with metadata."""
        doc = self._tree("word/document.xml")
        rels = self._tree("word/_rels/document.xml.rels")
        if doc is None:
            return []
        images = []
        for blip in doc.iter(f"{A}blip"):
            embed = blip.get(f"{R}embed")
            if not embed:
                continue
            info: dict = {"rId": embed, "filename": "", "content_type": ""}
            if rels is not None:
                rel = rels.find(f'{RELS}Relationship[@Id="{embed}"]')
                if rel is not None:
                    info["filename"] = rel.get("Target", "").split("/")[-1]
            # Get dimensions from wp:extent
            drawing = blip.getparent()
            while drawing is not None and drawing.tag != f"{W}drawing":
                drawing = drawing.getparent()
            if drawing is not None:
                extent = drawing.find(f".//{WP}extent")
                if extent is not None:
                    info["width_emu"] = int(extent.get("cx", "0"))
                    info["height_emu"] = int(extent.get("cy", "0"))
            # Content type from [Content_Types].xml
            ct_tree = self._tree("[Content_Types].xml")
            if ct_tree is not None:
                ext = info["filename"].rsplit(".", 1)[-1] if "." in info["filename"] else ""
                for default in ct_tree.findall(f"{CT}Default"):
                    if default.get("Extension") == ext:
                        info["content_type"] = default.get("ContentType", "")
            images.append(info)
        return images

    def _store_image(self, src: Path) -> tuple[str, str]:
        """Copy an image into word/media, add its relationship and content type.

        Returns (rId, filename). Never overwrites an existing media file.
        """
        if not src.is_file():
            raise FileNotFoundError(f"Image not found: {src}")
        ext = src.suffix.lstrip(".").lower() or "png"

        media_dir = self.workdir / "word" / "media"
        media_dir.mkdir(parents=True, exist_ok=True)
        taken = {p.name.lower() for p in media_dir.iterdir()}
        idx = 1
        while f"image{idx}.{ext}" in taken:
            idx += 1
        filename = f"image{idx}.{ext}"
        shutil.copy2(str(src), str(media_dir / filename))

        rels = self._tree("word/_rels/document.xml.rels")
        existing_rids = [r.get("Id") or "" for r in rels.findall(f"{RELS}Relationship")]
        rid_num = (
            max(
                (int(r[3:]) for r in existing_rids if r.startswith("rId") and r[3:].isdigit()),
                default=0,
            )
            + 1
        )
        rid = f"rId{rid_num}"
        rel = etree.SubElement(rels, f"{RELS}Relationship")
        rel.set("Id", rid)
        rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
        rel.set("Target", f"media/{filename}")
        self._mark("word/_rels/document.xml.rels")

        ct_tree = self._tree("[Content_Types].xml")
        content_type = _CONTENT_TYPES.get(ext, f"image/{ext}")
        has_ext = any(
            (d.get("Extension") or "").lower() == ext for d in ct_tree.findall(f"{CT}Default")
        )
        if not has_ext:
            default = etree.SubElement(ct_tree, f"{CT}Default")
            default.set("Extension", ext)
            default.set("ContentType", content_type)
            self._mark("[Content_Types].xml")
        return rid, filename

    def insert_image(
        self,
        para_id: str,
        image_path: str,
        *,
        width_emu: int = 0,
        height_emu: int = 0,
        alt_text: str = "",
    ) -> dict:
        """Insert an inline image in a new paragraph after a paragraph.

        Args:
            para_id: paraId of the paragraph to insert after.
            image_path: Path to the image file on disk.
            width_emu: Image width in EMUs (914400 = 1 inch, 360000 = 1 cm).
                0 = automatic.
            height_emu: Image height in EMUs. 0 = automatic. When only one of
                width/height is given the other keeps the image's aspect ratio;
                when neither is given the native size is used, scaled down to
                fit a 16 cm text width.
            alt_text: Description read by screen readers.
        """
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        src = Path(image_path)
        width_emu, height_emu = _auto_extent(src, int(width_emu or 0), int(height_emu or 0))
        rid, filename = self._store_image(src)
        drawing_id = self._next_drawing_id(doc)
        name = f"Picture {drawing_id}"

        new_para = etree.Element(f"{W}p")
        new_para.set(f"{W14}paraId", self._new_para_id())
        new_para.set(f"{W14}textId", "77777777")
        run = etree.SubElement(new_para, f"{W}r")
        drawing = etree.SubElement(run, f"{W}drawing")
        inline = etree.SubElement(
            drawing, f"{WP}inline", nsmap={"wp": _WP_NS}, distT="0", distB="0", distL="0", distR="0"
        )
        etree.SubElement(inline, f"{WP}extent", cx=str(width_emu), cy=str(height_emu))
        etree.SubElement(inline, f"{WP}effectExtent", l="0", t="0", r="0", b="0")
        doc_pr = etree.SubElement(inline, f"{WP}docPr", id=str(drawing_id), name=name)
        if alt_text:
            doc_pr.set("descr", alt_text)
        cnv_frame = etree.SubElement(inline, f"{WP}cNvGraphicFramePr")
        etree.SubElement(cnv_frame, f"{A}graphicFrameLocks", nsmap={"a": _A_NS}, noChangeAspect="1")
        graphic = etree.SubElement(inline, f"{A}graphic", nsmap={"a": _A_NS})
        gdata = etree.SubElement(graphic, f"{A}graphicData", uri=_PIC_NS)
        gdata.append(_build_pic(rid, width_emu, height_emu, 0, filename, alt_text))

        para.addnext(new_para)
        self._mark("word/document.xml")

        return {
            "filename": filename,
            "rId": rid,
            "para_id": new_para.get(f"{W14}paraId"),
            "width_emu": width_emu,
            "height_emu": height_emu,
            "width_cm": round(width_emu / _EMU_PER_CM, 2),
            "height_cm": round(height_emu / _EMU_PER_CM, 2),
        }

    # ── Image management ────────────────────────────────────────────────────

    def delete_image(self, rId: str) -> dict:
        """Remove the drawing containing image *rId* from the document body.

        Also removes the relationship entry from word/_rels/document.xml.rels.
        Raises ValueError if rId is not found.
        """
        doc = self._require("word/document.xml")
        rels = self._require("word/_rels/document.xml.rels")

        # Find the blip element
        blip = None
        for b in doc.iter(f"{A}blip"):
            if b.get(f"{R}embed") == rId:
                blip = b
                break
        if blip is None:
            raise ValueError(f"Image rId '{rId}' not found in document")

        # Walk up to the enclosing w:drawing
        drawing = blip.getparent()
        while drawing is not None and drawing.tag != f"{W}drawing":
            drawing = drawing.getparent()
        if drawing is None:
            raise ValueError(f"Could not find w:drawing for rId '{rId}'")

        # Walk up one more level: handle w:r wrapper or direct w:p child
        parent = drawing.getparent()
        if parent is None:
            raise ValueError(f"Drawing has no parent for rId '{rId}'")

        if parent.tag == f"{W}r":
            # normal case: drawing wrapped in a run — remove the whole run
            grandparent = parent.getparent()
            if grandparent is not None:
                grandparent.remove(parent)
            else:
                parent.remove(drawing)
        else:
            # drawing is a direct child of w:p or other element
            parent.remove(drawing)

        # Remove the relationship entry
        rel = rels.find(f'{RELS}Relationship[@Id="{rId}"]')
        if rel is not None:
            rels.remove(rel)

        self._mark("word/document.xml")
        self._mark("word/_rels/document.xml.rels")
        return {"deleted": rId}

    def update_image(self, rId: str, new_image_path: str) -> dict:
        """Replace the binary for an existing image in-place.

        Stores the new bytes in self._binaries keyed by zip path so they are
        written to the workdir on the next save().
        Raises ValueError if rId is not found.
        """
        rels = self._require("word/_rels/document.xml.rels")
        rel = rels.find(f'{RELS}Relationship[@Id="{rId}"]')
        if rel is None:
            raise ValueError(f"Image rId '{rId}' not found in relationships")

        target = rel.get("Target", "")
        zip_path = target.lstrip("/") if target.startswith("/") else f"word/{target}"

        new_bytes = Path(new_image_path).read_bytes()
        self._binaries[zip_path] = new_bytes
        return {"rId": rId, "new_path": new_image_path}

    def set_image_size(self, rId: str, width_cm: float, height_cm: float) -> dict:
        """Resize an embedded image by updating wp:extent (and a:ext if present).

        1 cm = 914400 / 2.54 ≈ 360000 EMU.
        Raises ValueError if rId is not found.
        """
        doc = self._require("word/document.xml")

        blip = None
        for b in doc.iter(f"{A}blip"):
            if b.get(f"{R}embed") == rId:
                blip = b
                break
        if blip is None:
            raise ValueError(f"Image rId '{rId}' not found in document")

        cx = round(width_cm * 360000)
        cy = round(height_cm * 360000)

        # Walk up to w:drawing
        drawing = blip.getparent()
        while drawing is not None and drawing.tag != f"{W}drawing":
            drawing = drawing.getparent()
        if drawing is None:
            raise ValueError(f"Could not find w:drawing for rId '{rId}'")

        # Update wp:extent
        extent = drawing.find(f".//{WP}extent")
        if extent is not None:
            extent.set("cx", str(cx))
            extent.set("cy", str(cy))

        # Update a:ext inside pic:spPr/a:xfrm if present
        for ext_el in drawing.iter(f"{A}ext"):
            ext_el.set("cx", str(cx))
            ext_el.set("cy", str(cy))

        self._mark("word/document.xml")
        return {
            "rId": rId,
            "width_cm": width_cm,
            "height_cm": height_cm,
            "width_emu": cx,
            "height_emu": cy,
        }

    def set_image_alt_text(self, rId: str, alt_text: str, title: str = "") -> dict:
        """Set accessibility alt text (descr) and title on a drawing's wp:docPr.

        Raises ValueError if rId is not found.
        """
        doc = self._require("word/document.xml")

        blip = None
        for b in doc.iter(f"{A}blip"):
            if b.get(f"{R}embed") == rId:
                blip = b
                break
        if blip is None:
            raise ValueError(f"Image rId '{rId}' not found in document")

        # Walk up to w:drawing
        drawing = blip.getparent()
        while drawing is not None and drawing.tag != f"{W}drawing":
            drawing = drawing.getparent()
        if drawing is None:
            raise ValueError(f"Could not find w:drawing for rId '{rId}'")

        # Find wp:docPr (under wp:inline or wp:anchor)
        doc_pr = drawing.find(f".//{WP}docPr")
        if doc_pr is None:
            # Create one under the inline/anchor container
            container = drawing.find(f"{WP}inline")
            if container is None:
                container = drawing.find(f"{WP}anchor")
            if container is None:
                # Fallback: attach directly to drawing
                container = drawing
            doc_pr = container.makeelement(f"{WP}docPr", {})
            doc_pr.set("id", "1")
            doc_pr.set("name", "Image")
            container.insert(0, doc_pr)

        doc_pr.set("descr", alt_text)
        if title:
            doc_pr.set("title", title)
        else:
            # Remove title attribute if present to keep XML clean
            if "title" in doc_pr.attrib:
                del doc_pr.attrib["title"]

        self._mark("word/document.xml")
        return {"rId": rId, "alt_text": alt_text}

    def set_image_border(self, rId: str, border_pt: float, color: str = "000000") -> dict:
        """Set or remove a border on an embedded image.

        Args:
            rId: The r:embed relationship ID of the image.
            border_pt: Border width in points (0 removes the border).
            color: RGB hex color string (default "000000").

        Returns:
            {"rId": str, "border_pt": float, "color": str}
        Raises:
            ValueError: if rId is not found in the document.
        """
        doc = self._require("word/document.xml")

        blip = None
        for b in doc.iter(f"{A}blip"):
            if b.get(f"{R}embed") == rId:
                blip = b
                break
        if blip is None:
            raise ValueError(f"Image rId '{rId}' not found in document")

        # Walk up to pic:pic
        pic_el = blip.getparent()
        while pic_el is not None and pic_el.tag != f"{_PIC}pic":
            pic_el = pic_el.getparent()
        if pic_el is None:
            raise ValueError(f"Image rId '{rId}' not found in document")

        # Get or create pic:spPr
        sp_pr = pic_el.find(f"{_PIC}spPr")
        if sp_pr is None:
            sp_pr = etree.SubElement(pic_el, f"{_PIC}spPr")

        # Remove existing a:ln
        existing_ln = sp_pr.find(f"{A}ln")
        if existing_ln is not None:
            sp_pr.remove(existing_ln)

        if border_pt == 0:
            self._mark("word/document.xml")
            return {"rId": rId, "border_pt": 0, "color": ""}

        # Create a:ln with width in EMU (1pt = 12700)
        ln = etree.SubElement(sp_pr, f"{A}ln")
        ln.set("w", str(round(border_pt * 12700)))

        solid_fill = etree.SubElement(ln, f"{A}solidFill")
        srgb_clr = etree.SubElement(solid_fill, f"{A}srgbClr")
        srgb_clr.set("val", color)

        self._mark("word/document.xml")
        return {"rId": rId, "border_pt": border_pt, "color": color}

    # ── Floating image helpers ───────────────────────────────────────────────

    def _next_drawing_id(self, doc_tree: etree._Element) -> int:
        """Return max existing wp:docPr id + 1 for unique id allocation."""
        max_id = 0
        for el in doc_tree.iter(f"{WP}docPr"):
            with contextlib.suppress(ValueError):
                max_id = max(max_id, int(el.get("id", "0")))
        return max_id + 1

    def insert_floating_image(
        self,
        para_id: str,
        image_path: str,
        width_cm: float,
        height_cm: float,
        h_pos: float = 0.0,
        v_pos: float = 0.0,
        wrap: str = "square",
    ) -> dict:
        """Insert a floating (anchored) image positioned at (h_pos, v_pos) cm from page origin.

        Returns {"filename": str, "rId": str, "width_emu": int, "height_emu": int, "wrap": str}.
        Raises ValueError if para_id not found.
        """
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")

        # EMU conversion: 1 cm = 914400/2.54 EMU
        width_emu = int(width_cm * 914400 / 2.54)
        height_emu = int(height_cm * 914400 / 2.54)
        h_pos_emu = int(h_pos * 914400 / 2.54)
        v_pos_emu = int(v_pos * 914400 / 2.54)

        rid, filename = self._store_image(Path(image_path))

        # Allocate unique drawing id
        drawing_id = self._next_drawing_id(doc)

        # Build wp:anchor element
        anchor = etree.Element(f"{WP}anchor")
        anchor.set("distT", "0")
        anchor.set("distB", "0")
        anchor.set("distL", "114300")
        anchor.set("distR", "114300")
        anchor.set("simplePos", "0")
        anchor.set("locked", "0")
        anchor.set("layoutInCell", "1")
        anchor.set("allowOverlap", "1")
        anchor.set("behindDoc", "0")
        anchor.set("relativeHeight", "251658240")

        simple_pos = etree.SubElement(anchor, f"{WP}simplePos")
        simple_pos.set("x", "0")
        simple_pos.set("y", "0")

        pos_h = etree.SubElement(anchor, f"{WP}positionH")
        pos_h.set("relativeFrom", "page")
        pos_offset_h = etree.SubElement(pos_h, f"{WP}posOffset")
        pos_offset_h.text = str(h_pos_emu)

        pos_v = etree.SubElement(anchor, f"{WP}positionV")
        pos_v.set("relativeFrom", "page")
        pos_offset_v = etree.SubElement(pos_v, f"{WP}posOffset")
        pos_offset_v.text = str(v_pos_emu)

        extent = etree.SubElement(anchor, f"{WP}extent")
        extent.set("cx", str(width_emu))
        extent.set("cy", str(height_emu))

        effect_extent = etree.SubElement(anchor, f"{WP}effectExtent")
        effect_extent.set("l", "0")
        effect_extent.set("t", "0")
        effect_extent.set("r", "0")
        effect_extent.set("b", "0")

        # Wrap element
        if wrap == "square":
            wrap_el = etree.SubElement(anchor, f"{WP}wrapSquare")
            wrap_el.set("wrapText", "bothSides")
        elif wrap == "topbottom":
            etree.SubElement(anchor, f"{WP}wrapTopAndBottom")
        else:  # "none"
            etree.SubElement(anchor, f"{WP}wrapNone")

        doc_pr = etree.SubElement(anchor, f"{WP}docPr")
        doc_pr.set("id", str(drawing_id))
        doc_pr.set("name", f"Image {drawing_id}")

        cnv_frame = etree.SubElement(anchor, f"{WP}cNvGraphicFramePr")
        frame_locks = etree.SubElement(cnv_frame, f"{A}graphicFrameLocks")
        frame_locks.set("noChangeAspect", "1")

        graphic = etree.SubElement(anchor, f"{A}graphic")
        gdata = etree.SubElement(graphic, f"{A}graphicData")
        gdata.set("uri", _PIC_NS)

        gdata.append(_build_pic(rid, width_emu, height_emu, 0, filename, ""))

        # Wrap anchor in w:drawing > w:r > w:p
        new_para = etree.Element(f"{W}p")
        new_para.set(f"{W14}paraId", self._new_para_id())
        new_para.set(f"{W14}textId", "77777777")
        run = etree.SubElement(new_para, f"{W}r")
        drawing_el = etree.SubElement(run, f"{W}drawing")
        drawing_el.append(anchor)

        para.addnext(new_para)
        self._mark("word/document.xml")

        return {
            "filename": filename,
            "rId": rid,
            "width_emu": width_emu,
            "height_emu": height_emu,
            "wrap": wrap,
        }
