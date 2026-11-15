"""Vector art in Word: native SVG, generated ornaments, page frames and borders.

Word 2016+ (Microsoft 365) draws SVG natively and keeps it sharp at any zoom
and in print. Each SVG is stored with a PNG fallback (rendered by resvg) for
older Word, LibreOffice and previews.

Ornaments are generated as clean geometry — no fonts, no clip art:

- ``star8``   — khatam (two interlaced squares), the eight-point star.
- ``rosette`` — n-point star polygon with a ring (default 12).
- ``shamsa``  — medallion: petals around concentric rings, for title pages.
- ``divider`` — rule tapering into a central star, with diamonds.
- ``divider-dots`` — light rule with a central diamond cluster.
- ``corner``  — corner piece (use four, rotated, or the frame).
- ``frame``   — full-page double frame with corner stars and side diamonds.
- ``band``    — repeated interlace of eight-point stars, for headers.

``add_page_frame`` anchors a frame behind the text in the header, so it
repeats on every page of the section without touching the body text.
"""

from __future__ import annotations

import contextlib
import math
import re
from pathlib import Path

from lxml import etree

from .ooxml import RELS, W14, W, add_relationship, color_hex, ensure_default

WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
PIC = "{http://schemas.openxmlformats.org/drawingml/2006/picture}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
ASVG_NS = "http://schemas.microsoft.com/office/drawing/2016/SVG/main"
SVG_EXT_URI = "{96DAC541-7B7A-43D3-8B79-37D633B846F1}"
EMU_PER_CM = 360000
EMU_PER_PT = 12700

ORNAMENTS = ("star8", "rosette", "shamsa", "divider", "divider-dots", "corner", "frame", "band")

PAGE_BORDER_STYLES = {
    "single": "single",
    "double": "double",
    "thick": "thick",
    "thick-thin": "thickThinMediumGap",
    "thin-thick": "thinThickMediumGap",
    "thin-thick-thin": "thinThickThinMediumGap",
    "triple": "triple",
    "dotted": "dotted",
    "dashed": "dashed",
    "wave": "wave",
    "double-wave": "doubleWave",
    "3d": "threeDEmboss",
    # Word's art borders (pictures repeated around the page)
    "art-celtic": "celticKnotwork",
    "art-weave": "weavingBraid",
    "art-weave-ribbon": "weavingRibbon",
    "art-weave-angles": "weavingAngles",
    "art-zigzag": "zigZag",
    "art-triangles": "triangles",
    "art-diamonds": "diamondsGray",
    "art-mosaic": "mosaic",
    "art-flowers": "flowersModern1",
    "art-vine": "vine",
    "art-double-d": "doubleD",
    "art-basic-wide-outline": "basicWideOutline",
    "art-basic-black-dots": "basicBlackDots",
    "art-twisted-lines": "twistedLines1",
}


# ── geometry ────────────────────────────────────────────────────────────


def _pt(x: float, y: float) -> str:
    return f"{x:.2f},{y:.2f}"


def _star_points(cx: float, cy: float, r_out: float, r_in: float, n: int, rot: float = -90):
    pts = []
    for i in range(2 * n):
        r = r_out if i % 2 == 0 else r_in
        a = math.radians(rot + i * 180 / n)
        pts.append((cx + r * math.cos(a), cy + r * math.sin(a)))
    return pts


def _polygon(pts) -> str:
    return " ".join(_pt(x, y) for x, y in pts)


def _khatam(cx: float, cy: float, r: float, color: str, fill: str, sw: float) -> str:
    """The eight-point star (khatam): one outline, inner octagon and ring.

    Outer points on radius r; the inner vertices of two interlaced squares
    sit at r·cos45°/cos22.5°.
    """
    r_in = r * math.cos(math.radians(45)) / math.cos(math.radians(22.5))
    star = _star_points(cx, cy, r, r_in, 8, rot=-90)
    inner = _star_points(cx, cy, r * 0.56, r * 0.56 * r_in / r, 8, rot=-90 + 22.5)
    return (
        f'<polygon points="{_polygon(star)}" fill="{fill}" stroke="{color}" '
        f'stroke-width="{sw}" stroke-linejoin="miter"/>'
        f'<polygon points="{_polygon(inner)}" fill="none" stroke="{color}" '
        f'stroke-width="{sw * 0.7}" stroke-linejoin="miter"/>'
        f'<circle cx="{cx:.2f}" cy="{cy:.2f}" r="{r * 0.2:.2f}" fill="none" '
        f'stroke="{color}" stroke-width="{sw * 0.8}"/>'
    )


def _diamond(cx: float, cy: float, w: float, h: float, color: str) -> str:
    pts = [(cx, cy - h), (cx + w, cy), (cx, cy + h), (cx - w, cy)]
    return f'<polygon points="{_polygon(pts)}" fill="{color}"/>'


def _svg(w: float, h: float, body: str) -> str:
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{w:.0f}" height="{h:.0f}" '
        f'viewBox="0 0 {w:.2f} {h:.2f}">{body}</svg>'
    )


def ornament_svg(
    kind: str,
    *,
    color: str = "520912",
    accent: str | None = None,
    aspect: float | None = None,
    points: int = 12,
    weight: float = 1.0,
) -> str:
    """SVG markup for a generated ornament.

    Args:
        kind: One of ORNAMENTS.
        color: Main colour (hex).
        accent: Second colour for fills (default: a light tint of color).
        aspect: Width / height (dividers, bands, frames). Frames default to
            the page shape.
        points: Points of a rosette (8–24).
        weight: Line weight multiplier.
    """
    if kind not in ORNAMENTS:
        raise ValueError(f"kind must be one of {ORNAMENTS}")
    c = "#" + (color_hex(color) or "520912")
    acc = "#" + color_hex(accent) if accent else _tint(c, 0.82)
    sw = 1.6 * weight

    if kind == "star8":
        return _svg(
            100,
            100,
            _khatam(50, 50, 47, c, acc, sw * 1.2) + f'<circle cx="50" cy="50" r="4" fill="{c}"/>',
        )

    if kind == "rosette":
        n = max(8, min(24, int(points)))
        star = _star_points(50, 50, 47, 30, n)
        inner = _star_points(50, 50, 26, 17, n, rot=-90 + 180 / n)
        return _svg(
            100,
            100,
            f'<polygon points="{_polygon(star)}" fill="{acc}" stroke="{c}" '
            f'stroke-width="{sw}" stroke-linejoin="round"/>'
            f'<circle cx="50" cy="50" r="28" fill="none" stroke="{c}" stroke-width="{sw}"/>'
            f'<polygon points="{_polygon(inner)}" fill="none" stroke="{c}" '
            f'stroke-width="{sw * 0.8}"/>'
            f'<circle cx="50" cy="50" r="7" fill="{c}"/>',
        )

    if kind == "shamsa":
        body = []
        n = 16
        for i in range(n):
            a = 360 / n * i
            body.append(
                f'<ellipse cx="50" cy="16" rx="5.2" ry="13" fill="{acc}" stroke="{c}" '
                f'stroke-width="{sw * 0.7}" transform="rotate({a:.2f} 50 50)"/>'
            )
        body.append(
            f'<circle cx="50" cy="50" r="30" fill="#ffffff" stroke="{c}" stroke-width="{sw}"/>'
        )
        body.append(
            f'<circle cx="50" cy="50" r="26.5" fill="none" stroke="{c}" stroke-width="{sw * 0.6}"/>'
        )
        body.append(_khatam(50, 50, 20, c, "none", sw * 0.8))
        body.append(f'<circle cx="50" cy="50" r="3.5" fill="{c}"/>')
        return _svg(100, 100, "".join(body))

    if kind == "divider":
        ar = aspect or 12.0
        h = 40.0
        w = h * ar
        cx, cy = w / 2, h / 2
        body = [
            # tapering rules: thin at the ends, thick near the centre
            f'<polygon points="{_polygon([(4, cy), (cx - 26, cy - 1.6 * weight), (cx - 26, cy + 1.6 * weight)])}" fill="{c}"/>',
            f'<polygon points="{_polygon([(w - 4, cy), (cx + 26, cy - 1.6 * weight), (cx + 26, cy + 1.6 * weight)])}" fill="{c}"/>',
            _khatam(cx, cy, 15, c, acc, sw),
            f'<circle cx="{cx:.2f}" cy="{cy:.2f}" r="2.6" fill="{c}"/>',
            _diamond(cx - 34, cy, 5, 4, c),
            _diamond(cx + 34, cy, 5, 4, c),
            _diamond(cx - w * 0.22, cy, 3.5, 3, c),
            _diamond(cx + w * 0.22, cy, 3.5, 3, c),
        ]
        return _svg(w, h, "".join(body))

    if kind == "divider-dots":
        ar = aspect or 10.0
        h = 24.0
        w = h * ar
        cx, cy = w / 2, h / 2
        body = [
            f'<line x1="{w * 0.12:.2f}" y1="{cy}" x2="{cx - 22:.2f}" y2="{cy}" stroke="{c}" stroke-width="{0.9 * weight}"/>',
            f'<line x1="{cx + 22:.2f}" y1="{cy}" x2="{w * 0.88:.2f}" y2="{cy}" stroke="{c}" stroke-width="{0.9 * weight}"/>',
            _diamond(cx, cy, 6, 6, c),
            _diamond(cx - 13, cy, 3.5, 3.5, c),
            _diamond(cx + 13, cy, 3.5, 3.5, c),
            f'<circle cx="{w * 0.12:.2f}" cy="{cy}" r="1.8" fill="{c}"/>',
            f'<circle cx="{w * 0.88:.2f}" cy="{cy}" r="1.8" fill="{c}"/>',
        ]
        return _svg(w, h, "".join(body))

    if kind == "corner":
        body = [
            f'<path d="M4,96 L4,24 Q4,4 24,4 L96,4" fill="none" stroke="{c}" stroke-width="{sw * 1.4}"/>',
            f'<path d="M12,96 L12,28 Q12,12 28,12 L96,12" fill="none" stroke="{c}" stroke-width="{sw * 0.7}"/>',
            _khatam(28, 28, 11, c, acc, sw * 0.8),
            _diamond(60, 12, 4, 3, c),
            _diamond(12, 60, 3, 4, c),
        ]
        return _svg(100, 100, "".join(body))

    if kind == "band":
        ar = aspect or 12.0
        h = 30.0
        w = h * ar
        n = max(3, int(w // 30))
        step = w / n
        body = [
            f'<line x1="0" y1="2" x2="{w:.2f}" y2="2" stroke="{c}" stroke-width="{sw}"/>',
            f'<line x1="0" y1="{h - 2}" x2="{w:.2f}" y2="{h - 2}" stroke="{c}" stroke-width="{sw}"/>',
        ]
        for i in range(n):
            cx = step * (i + 0.5)
            body.append(_khatam(cx, h / 2, 10.5, c, acc if i % 2 == 0 else "none", sw * 0.7))
            if i:
                body.append(_diamond(step * i, h / 2, 3, 3, c))
        return _svg(w, h, "".join(body))

    # frame
    ar = aspect or (210 / 297)
    h = 1000.0
    w = h * ar
    m1, m2 = 10.0, 22.0
    body = [
        f'<rect x="{m1}" y="{m1}" width="{w - 2 * m1:.2f}" height="{h - 2 * m1:.2f}" '
        f'fill="none" stroke="{c}" stroke-width="{sw * 2.2}"/>',
        f'<rect x="{m2}" y="{m2}" width="{w - 2 * m2:.2f}" height="{h - 2 * m2:.2f}" '
        f'fill="none" stroke="{c}" stroke-width="{sw * 0.9}"/>',
    ]
    for x, y in (
        (m1 + 6, m1 + 6),
        (w - m1 - 6, m1 + 6),
        (m1 + 6, h - m1 - 6),
        (w - m1 - 6, h - m1 - 6),
    ):
        body.append(
            f'<rect x="{x - 16:.2f}" y="{y - 16:.2f}" width="32" height="32" fill="#ffffff"/>'
        )
        body.append(_khatam(x, y, 15, c, acc, sw * 0.9))
    for x, y, dw, dh in (
        (w / 2, m1 + 6, 10, 6),
        (w / 2, h - m1 - 6, 10, 6),
        (m1 + 6, h / 2, 6, 10),
        (w - m1 - 6, h / 2, 6, 10),
    ):
        body.append(
            f'<rect x="{x - dw - 3:.2f}" y="{y - dh - 3:.2f}" width="{2 * dw + 6}" '
            f'height="{2 * dh + 6}" fill="#ffffff"/>'
        )
        body.append(_diamond(x, y, dw, dh, c))
    return _svg(w, h, "".join(body))


def _tint(hex_color: str, amount: float) -> str:
    h = hex_color.lstrip("#")
    r, g, b = (int(h[i : i + 2], 16) for i in (0, 2, 4))
    r, g, b = (round(v + (255 - v) * amount) for v in (r, g, b))
    return f"#{r:02x}{g:02x}{b:02x}"


def _svg_size(svg: str) -> tuple[float, float]:
    root = etree.fromstring(svg.encode("utf-8"))
    vb = root.get("viewBox")
    if vb:
        parts = [float(v) for v in re.split(r"[ ,]+", vb.strip())]
        if len(parts) == 4 and parts[2] > 0 and parts[3] > 0:
            return parts[2], parts[3]

    def num(v: str | None) -> float:
        m = re.match(r"[\d.]+", v or "")
        return float(m.group(0)) if m else 0.0

    w, h = num(root.get("width")), num(root.get("height"))
    return (w or 100.0), (h or 100.0)


def render_png(svg: str, width_px: int) -> bytes:
    """PNG fallback via resvg (bundled wheels on Windows, macOS, Linux)."""
    try:
        import resvg_py
    except ImportError:  # pragma: no cover - dependency is declared
        # 1×1 transparent PNG: Word still uses the SVG; old readers show nothing.
        return bytes.fromhex(
            "89504e470d0a1a0a0000000d4948445200000001000000010806000000"
            "1f15c4890000000d49444154789c6360000002000154a24f5d0000000049454e44ae426082"
        )
    return bytes(resvg_py.svg_to_bytes(svg_string=svg, width=max(16, int(width_px))))


# ── Word plumbing ───────────────────────────────────────────────────────


class VectorArtMixin:
    """SVG pictures, ornaments, page frames and page borders."""

    def _rels_for(self, part: str) -> tuple[str, etree._Element]:
        folder, name = part.rsplit("/", 1)
        path = f"{folder}/_rels/{name}.rels"
        rels = self._tree(path)
        if rels is None:
            rels = etree.Element(
                f"{RELS}Relationships",
                nsmap={None: RELS[1:-1]},
            )
            self._trees[path] = rels
        return path, rels

    def _store_svg(self, part: str, svg: str, width_px: int) -> tuple[str, str]:
        """Write the SVG and its PNG fallback to word/media; return (png rId, svg rId)."""
        taken = {p.rsplit("/", 1)[-1] for p in self._binaries}
        if self.workdir is not None:
            media = Path(self.workdir) / "word" / "media"
            if media.is_dir():
                taken |= {p.name for p in media.iterdir()}
        i = 1
        while f"vector{i}.svg" in taken or f"vector{i}.png" in taken:
            i += 1
        svg_name, png_name = f"vector{i}.svg", f"vector{i}.png"
        self._binaries[f"word/media/{svg_name}"] = svg.encode("utf-8")
        self._binaries[f"word/media/{png_name}"] = render_png(svg, width_px)
        ct = self._tree("[Content_Types].xml")
        ensure_default(ct, "svg", "image/svg+xml")
        ensure_default(ct, "png", "image/png")
        self._mark("[Content_Types].xml")
        rels_path, rels = self._rels_for(part)
        rid_png = add_relationship(rels, "image", f"media/{png_name}")
        rid_svg = add_relationship(rels, "image", f"media/{svg_name}")
        self._mark(rels_path)
        return rid_png, rid_svg

    def _next_docpr_id(self) -> int:
        top = 0
        for root in self._trees.values():
            if root is None:
                continue
            for dp in root.iter(f"{WP}docPr"):
                with contextlib.suppress(ValueError):
                    top = max(top, int(dp.get("id", "0")))
        return top + 1

    def _graphic(
        self, rid_png: str, rid_svg: str, cx: int, cy: int, pid: int, name: str, descr: str
    ) -> etree._Element:
        graphic = etree.Element(f"{A}graphic", nsmap={"a": A[1:-1]})
        gd = etree.SubElement(
            graphic,
            f"{A}graphicData",
            uri="http://schemas.openxmlformats.org/drawingml/2006/picture",
        )
        pic = etree.SubElement(gd, f"{PIC}pic", nsmap={"pic": PIC[1:-1]})
        nv = etree.SubElement(pic, f"{PIC}nvPicPr")
        cnv = etree.SubElement(nv, f"{PIC}cNvPr", id=str(pid), name=name)
        if descr:
            cnv.set("descr", descr)
        etree.SubElement(etree.SubElement(nv, f"{PIC}cNvPicPr"), f"{A}picLocks", noChangeAspect="1")
        bf = etree.SubElement(pic, f"{PIC}blipFill")
        blip = etree.SubElement(bf, f"{A}blip")
        blip.set(f"{R}embed", rid_png)
        ext_lst = etree.SubElement(blip, f"{A}extLst")
        ext = etree.SubElement(ext_lst, f"{A}ext", uri=SVG_EXT_URI)
        svg_blip = etree.SubElement(ext, f"{{{ASVG_NS}}}svgBlip", nsmap={"asvg": ASVG_NS})
        svg_blip.set(f"{R}embed", rid_svg)
        etree.SubElement(etree.SubElement(bf, f"{A}stretch"), f"{A}fillRect")
        sp = etree.SubElement(pic, f"{PIC}spPr")
        xfrm = etree.SubElement(sp, f"{A}xfrm")
        etree.SubElement(xfrm, f"{A}off", x="0", y="0")
        etree.SubElement(xfrm, f"{A}ext", cx=str(cx), cy=str(cy))
        geom = etree.SubElement(sp, f"{A}prstGeom", prst="rect")
        etree.SubElement(geom, f"{A}avLst")
        return graphic

    def _inline_drawing(
        self, part: str, svg: str, cx: int, cy: int, name: str, descr: str
    ) -> etree._Element:
        rid_png, rid_svg = self._store_svg(part, svg, int(cx / EMU_PER_CM * 118))  # ~300 dpi
        pid = self._next_docpr_id()
        drawing = etree.Element(f"{W}drawing")
        inline = etree.SubElement(
            drawing,
            f"{WP}inline",
            nsmap={"wp": WP[1:-1]},
            distT="0",
            distB="0",
            distL="0",
            distR="0",
        )
        etree.SubElement(inline, f"{WP}extent", cx=str(cx), cy=str(cy))
        etree.SubElement(inline, f"{WP}effectExtent", l="0", t="0", r="0", b="0")
        dp = etree.SubElement(inline, f"{WP}docPr", id=str(pid), name=name)
        if descr:
            dp.set("descr", descr)
        etree.SubElement(
            etree.SubElement(inline, f"{WP}cNvGraphicFramePr"),
            f"{A}graphicFrameLocks",
            nsmap={"a": A[1:-1]},
            noChangeAspect="1",
        )
        inline.append(self._graphic(rid_png, rid_svg, cx, cy, pid, name, descr))
        return drawing

    def _anchor_drawing(
        self,
        part: str,
        svg: str,
        cx: int,
        cy: int,
        x: int,
        y: int,
        name: str,
        behind: bool = True,
        rel_from: str = "page",
    ) -> etree._Element:
        rid_png, rid_svg = self._store_svg(part, svg, int(cx / EMU_PER_CM * 60))
        pid = self._next_docpr_id()
        drawing = etree.Element(f"{W}drawing")
        an = etree.SubElement(
            drawing,
            f"{WP}anchor",
            nsmap={"wp": WP[1:-1]},
            distT="0",
            distB="0",
            distL="0",
            distR="0",
            simplePos="0",
            relativeHeight=str(251658240 + pid),
            behindDoc="1" if behind else "0",
            locked="1",
            layoutInCell="1",
            allowOverlap="1",
        )
        etree.SubElement(an, f"{WP}simplePos", x="0", y="0")
        ph = etree.SubElement(an, f"{WP}positionH", relativeFrom=rel_from)
        etree.SubElement(ph, f"{WP}posOffset").text = str(x)
        pv = etree.SubElement(an, f"{WP}positionV", relativeFrom=rel_from)
        etree.SubElement(pv, f"{WP}posOffset").text = str(y)
        etree.SubElement(an, f"{WP}extent", cx=str(cx), cy=str(cy))
        etree.SubElement(an, f"{WP}effectExtent", l="0", t="0", r="0", b="0")
        etree.SubElement(an, f"{WP}wrapNone")
        etree.SubElement(an, f"{WP}docPr", id=str(pid), name=name)
        etree.SubElement(
            etree.SubElement(an, f"{WP}cNvGraphicFramePr"),
            f"{A}graphicFrameLocks",
            nsmap={"a": A[1:-1]},
            noChangeAspect="1",
        )
        an.append(self._graphic(rid_png, rid_svg, cx, cy, pid, name, ""))
        return drawing

    # ── public API ──────────────────────────────────────────────────────

    def insert_svg(
        self,
        svg: str,
        *,
        after_para_id: str | None = None,
        width_cm: float | None = None,
        height_cm: float | None = None,
        alignment: str = "center",
        alt_text: str = "",
        space_before_pt: float = 6,
        space_after_pt: float = 6,
    ) -> dict:
        """Insert SVG artwork (markup or a .svg file path) as a sharp vector picture.

        Width/height: give one to keep proportions; neither = 5 cm wide.
        Placed in a new paragraph after after_para_id (default: end).
        """
        if not svg.lstrip().startswith("<"):
            path = Path(svg)
            if not path.is_file():
                raise FileNotFoundError(f"SVG not found: {svg}")
            svg = path.read_text(encoding="utf-8")
        try:
            vw, vh = _svg_size(svg)
        except etree.XMLSyntaxError as e:
            raise ValueError(f"Not valid SVG: {e}") from e
        if width_cm and height_cm:
            w, h = width_cm, height_cm
        elif width_cm:
            w, h = width_cm, width_cm * vh / vw
        elif height_cm:
            w, h = height_cm * vw / vh, height_cm
        else:
            w = 5.0
            h = w * vh / vw
        return self._place_inline(
            svg,
            w,
            h,
            after_para_id,
            alignment,
            alt_text or "Vector art",
            space_before_pt,
            space_after_pt,
        )

    def _place_inline(self, svg, w_cm, h_cm, after_para_id, alignment, alt, sb, sa) -> dict:
        from .ooxml import make_paragraph

        doc = self._require("word/document.xml")
        body = doc.find(f"{W}body")
        p = make_paragraph(
            self._new_para_id(),
            {
                "alignment": alignment,
                "space_before_pt": sb,
                "space_after_pt": sa,
                "keep_with_next": False,
            },
        )
        r = etree.SubElement(p, f"{W}r")
        r.append(
            self._inline_drawing(
                "word/document.xml",
                svg,
                int(w_cm * EMU_PER_CM),
                int(h_cm * EMU_PER_CM),
                "Vector art",
                alt,
            )
        )
        if after_para_id:
            anchor = self._find_para(doc, after_para_id)
            if anchor is None:
                raise ValueError(f"Paragraph '{after_para_id}' not found")
            parent = anchor.getparent()
            parent.insert(list(parent).index(anchor) + 1, p)
        else:
            kids = list(body)
            pos = len(kids) - 1 if kids and kids[-1].tag == f"{W}sectPr" else len(kids)
            body.insert(pos, p)
        self._mark("word/document.xml")
        return {
            "para_id": p.get(f"{W14}paraId"),
            "width_cm": round(w_cm, 2),
            "height_cm": round(h_cm, 2),
        }

    def insert_ornament(
        self,
        kind: str = "divider",
        *,
        after_para_id: str | None = None,
        width_cm: float | None = None,
        color: str = "520912",
        accent: str | None = None,
        points: int = 12,
        weight: float = 1.0,
        alignment: str = "center",
    ) -> dict:
        """Insert a generated ornament (see ORNAMENTS) as vector art.

        Dividers and bands default to 60% / 100% of the text width; stars,
        rosettes, shamsa and corners to 2.5 cm.
        """
        if kind == "frame":
            raise ValueError("Use add_page_frame for a page frame")
        text_w = self._text_width_cm() if hasattr(self, "_text_width_cm") else 15.0
        if width_cm is None:
            width_cm = {"divider": text_w * 0.6, "divider-dots": text_w * 0.45, "band": text_w}.get(
                kind, 2.5
            )
        aspect = None
        if kind in ("divider", "divider-dots", "band"):
            heights = {"divider": 0.7, "divider-dots": 0.45, "band": 0.8}
            aspect = width_cm / heights[kind]
        svg = ornament_svg(
            kind, color=color, accent=accent, aspect=aspect, points=points, weight=weight
        )
        vw, vh = _svg_size(svg)
        r = self._place_inline(
            svg, width_cm, width_cm * vh / vw, after_para_id, alignment, f"Ornament: {kind}", 6, 6
        )
        return {"kind": kind, **r}

    def add_page_frame(
        self,
        *,
        color: str = "520912",
        accent: str | None = None,
        margin_cm: float = 0.9,
        weight: float = 1.0,
        section_index: int | None = None,
        svg: str | None = None,
    ) -> dict:
        """Draw a decorative frame on every page, behind the text.

        The frame lives in the header (anchored to the page, behind text), so
        it repeats on all pages of the section and never moves the body.
        Existing headers keep their content; first-page and even-page headers
        get the frame too. Pass your own svg to use custom artwork.
        """
        doc = self._require("word/document.xml")
        sects = (
            self._pick_sections(section_index)
            if hasattr(self, "_pick_sections")
            else list(doc.iter(f"{W}sectPr"))
        )
        done = []
        for sect in sects:
            pg = sect.find(f"{W}pgSz")
            pw = int(pg.get(f"{W}w", "11906")) if pg is not None else 11906
            ph = int(pg.get(f"{W}h", "16838")) if pg is not None else 16838
            if pg is not None and pg.get(f"{W}orient") == "landscape" and pw < ph:
                pw, ph = ph, pw
            pw_cm, ph_cm = pw / 567.0, ph / 567.0
            fw, fh = pw_cm - 2 * margin_cm, ph_cm - 2 * margin_cm
            art = svg or ornament_svg(
                "frame", color=color, accent=accent, aspect=fw / fh, weight=weight
            )
            refs = [r for r in sect.findall(f"{W}headerReference")]
            if not refs:
                path, root, rid = self._new_story_part("header")
                from .ooxml import make_paragraph

                root.append(make_paragraph(self._new_para_id(), {"space_after_pt": 0}))
                self._attach(sect, "header", "default", rid)
                refs = [r for r in sect.findall(f"{W}headerReference")]
            rels = self._tree("word/_rels/document.xml.rels")
            for ref in refs:
                rid = ref.get(f"{R}id")
                rel = rels.find(f'{RELS}Relationship[@Id="{rid}"]')
                if rel is None:
                    continue
                part = "word/" + rel.get("Target").lstrip("/").removeprefix("word/")
                root = self._require(part)
                p = root.find(f"{W}p")
                if p is None:
                    from .ooxml import make_paragraph

                    p = make_paragraph(self._new_para_id(), {"space_after_pt": 0})
                    root.append(p)
                run = etree.Element(f"{W}r")
                run.append(
                    self._anchor_drawing(
                        part,
                        art,
                        int(fw * EMU_PER_CM),
                        int(fh * EMU_PER_CM),
                        int(margin_cm * EMU_PER_CM),
                        int(margin_cm * EMU_PER_CM),
                        "Page frame",
                    )
                )
                ppr = p.find(f"{W}pPr")
                p.insert(0 if ppr is None else list(p).index(ppr) + 1, run)
                self._mark(part)
                done.append(part)
        self._mark("word/document.xml")
        return {"frame_in": done, "margin_cm": margin_cm}

    def set_page_border(
        self,
        style: str = "double",
        *,
        color: str = "520912",
        width_pt: float = 1.5,
        space_pt: float = 24,
        measure_from: str = "page",
        pages: str = "all",
        sides: tuple[str, ...] | list[str] = ("top", "left", "bottom", "right"),
        section_index: int | None = None,
    ) -> dict:
        """Word's page border: lines (single, double, thick-thin …) or art borders
        (art-celtic, art-weave, art-zigzag … — Word draws the repeated picture).

        Args:
            style: A key of PAGE_BORDER_STYLES, or "none" to remove.
            width_pt: Line width (art borders: picture size, 1–31 pt).
            space_pt: Distance from the page edge (measure_from="page") or
                from the text ("text"), 0–31 pt.
            pages: all, first, not-first.
        """
        if style != "none" and style not in PAGE_BORDER_STYLES:
            raise ValueError(f"style must be one of {sorted(PAGE_BORDER_STYLES)} or none")
        if measure_from not in ("page", "text"):
            raise ValueError("measure_from must be page or text")
        display = {"all": "allPages", "first": "firstPage", "not-first": "notFirstPage"}
        if pages not in display:
            raise ValueError("pages must be all, first or not-first")
        doc = self._require("word/document.xml")
        sects = (
            self._pick_sections(section_index)
            if hasattr(self, "_pick_sections")
            else list(doc.iter(f"{W}sectPr"))
        )
        val = PAGE_BORDER_STYLES.get(style)
        art = style.startswith("art-")
        for sect in sects:
            old = sect.find(f"{W}pgBorders")
            if old is not None:
                sect.remove(old)
            if style == "none":
                continue
            pb = etree.SubElement(sect, f"{W}pgBorders")
            pb.set(f"{W}offsetFrom", measure_from)
            if pages != "all":
                pb.set(f"{W}display", display[pages])
            for side in ("top", "left", "bottom", "right"):
                if side not in sides:
                    continue
                el = etree.SubElement(pb, f"{W}{side}")
                el.set(f"{W}val", val)
                el.set(
                    f"{W}sz",
                    str(max(1, min(31, round(width_pt))))
                    if art
                    else str(max(2, min(96, round(width_pt * 8)))),
                )
                el.set(f"{W}space", str(max(0, min(31, round(space_pt)))))
                el.set(f"{W}color", color_hex(color) or "auto")
        self._mark("word/document.xml")
        return {"style": style, "sections": len(sects)}

    def list_ornaments(self) -> dict:
        return {"ornaments": list(ORNAMENTS), "page_borders": sorted(PAGE_BORDER_STYLES)}
