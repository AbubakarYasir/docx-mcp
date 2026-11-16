"""Book conventions from published Arabic scholarship, as Word features.

Studied from printed books (Dar Rawai' al-Kitab, Adwa' al-Salaf, Dar
al-Tadmuriyya, the Saudi Ministry's manasik booklet …):

- ``design_running_head`` — running heads as publishers set them: page
  number boxed between double rules, between ornamental brackets, or on a
  cartouche band with a medallion folio; chapter (STYLEREF) on odd pages,
  book title on even, number always on the outer side.
- ``set_note_mark_style`` — note numbers in parentheses, «(١)», in the text
  and in the note (Word has no such number format; the marks stay live).
- ``insert_verses`` — nazm: two hemistichs per line, each justified to its
  half, optional verse numbers, in a borderless right-to-left table.
- ``add_margin_note`` — a side heading in the outer margin beside a
  paragraph (moves with it; flips sides on odd/even pages).
- ``define_numbered_style`` — auto-numbered entries («١٨٨٠ . سليمان بن عمرو:»):
  a paragraph style tied to a numbering definition, so inserting an entry
  renumbers everything after it.
- ``insert_part_title`` — a part / فهارس title page: own section,
  vertically centred, with a woven frame drawn on that page only.
"""

from __future__ import annotations

import copy
from typing import Any

from lxml import etree

from . import islamic_art as ia
from .bidi import document_rtl_default, has_rtl
from .ooxml import (
    W14,
    W,
    apply_para_props,
    color_hex,
    content_runs,
    make_paragraph,
    make_run,
)

WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
WPS = "{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
EMU_PER_CM = 360000

HEAD_STYLES = ("rule", "boxed", "vignette", "cartouche")


def _bracket_svg(color: str, side: str) -> str:
    """Half a fleuron (leaves and curl) as an ornamental bracket for a folio."""
    from shapely import affinity
    from shapely.geometry import box

    f = ia.fleuron(color)
    half = f.intersection(box(0.8, -20, 30, 20))
    half = affinity.translate(half, -0.8, 0)
    if side == "left":
        half = ia.mirror_x(half, 0)
    minx, miny, maxx, maxy = half.bounds
    half = affinity.translate(half, -minx + 0.5, -miny + 0.5)
    return ia.svg(maxx - minx + 1, maxy - miny + 1, ia.fill(half, ia.hexc(color)))


class BookPartsMixin:
    """Running heads, note marks, verse, margin notes, numbered entries,
    part-title pages."""

    # ── shared helpers ──────────────────────────────────────────────────

    def _is_rtl(self) -> bool:
        return bool(document_rtl_default(self._tree("word/styles.xml")))

    def _anchor_behind(
        self,
        part: str,
        p: etree._Element,
        art: str,
        w_cm: float,
        h_cm: float,
        line_pt: float,
        pad_pt: float,
        name: str = "Ornament",
    ) -> None:
        """Anchor art behind paragraph p, centred on its first line."""
        run = etree.Element(f"{W}r")
        drawing = self._anchor_drawing(
            part, art, int(w_cm * EMU_PER_CM), int(h_cm * EMU_PER_CM), 0, 0, name, behind=True
        )
        an = drawing.find(f"{WP}anchor")
        ph = an.find(f"{WP}positionH")
        ph.set("relativeFrom", "margin")
        for c in list(ph):
            ph.remove(c)
        etree.SubElement(ph, f"{WP}align").text = "center"
        pv = an.find(f"{WP}positionV")
        pv.set("relativeFrom", "paragraph")
        pv.find(f"{WP}posOffset").text = str(
            int((pad_pt + line_pt / 2) / 72 * 2.54 * EMU_PER_CM - h_cm * EMU_PER_CM / 2)
        )
        an.set("locked", "0")
        run.append(drawing)
        ppr = p.find(f"{W}pPr")
        p.insert(0 if ppr is None else list(p).index(ppr) + 1, run)

    def _place_after(self, p: etree._Element, after_para_id: str | None) -> None:
        doc = self._require("word/document.xml")
        if after_para_id:
            anchor = self._find_para(doc, after_para_id)
            if anchor is None:
                raise ValueError(f"Paragraph '{after_para_id}' not found")
            parent = anchor.getparent()
            parent.insert(list(parent).index(anchor) + 1, p)
        else:
            body = doc.find(f"{W}body")
            kids = list(body)
            body.insert(len(kids) - 1 if kids and kids[-1].tag == f"{W}sectPr" else len(kids), p)
        self._mark("word/document.xml")

    # ── running heads ───────────────────────────────────────────────────

    def design_running_head(
        self,
        style: str = "boxed",
        *,
        title: str | None = None,
        chapter_style: str = "Heading 1",
        color: str = "520912",
        font: str | None = None,
        size_pt: float = 12,
        bold: bool = True,
        section_index: int | None = None,
    ) -> dict:
        """Running heads with the page number on the outer side.

        Odd pages: the current chapter (STYLEREF chapter_style) and the
        number; even pages: the book title and the number.

        style:
          rule      — text and number over a thin rule.
          boxed     — number in a small box, double rule below (as in
                      «الرد القويم»).
          vignette  — number between ornamental brackets, rule below (as in
                      «جمهرة مستخرجات الجرح والتعديل»).
          cartouche — the head centred on a banner cartouche, and the folio
                      in a star medallion at the foot (as in «مناسك الحج»).
        """
        if style not in HEAD_STYLES:
            raise ValueError(f"style must be one of {HEAD_STYLES}")
        rtl = self._is_rtl()
        c = color_hex(color) or "520912"
        book = title or "{TITLE}"
        chap = "{STYLEREF:" + chapter_style + "}"
        base: dict[str, Any] = {"size_pt": size_pt, "bold": bold}
        if font:
            base["font"] = font
            base["cs_font"] = font
        if rtl:
            base["rtl"] = True
        out = []
        for sect in self._pick_sections(section_index):
            width_cm = self._text_width_twips(sect) * 2.54 / 1440
            for variant, text in (("default", chap), ("even", book)):
                path, root, rid = self._new_story_part("header")
                if style == "cartouche":
                    p = self._cartouche_head(path, text, base, c, width_cm, rtl)
                else:
                    p = self._ruled_head(path, style, text, variant, base, c, width_cm, rtl)
                root.append(p)
                self._attach(sect, "header", variant, rid)
                out.append(path)
            for variant in ("default", "even"):
                path, root, rid = self._new_story_part("footer")
                if style == "cartouche":
                    root.append(self._medallion_folio(path, base, c, rtl))
                else:
                    root.append(make_paragraph(self._new_para_id(), {"space_after_pt": 0}))
                self._attach(sect, "footer", variant, rid)
        self._mark("word/document.xml")
        return {"style": style, "headers": out}

    def _ruled_head(self, part, style, text, variant, base, color, width_cm, rtl):
        """Text at the inner side, number at the outer side, over a rule."""
        props: dict[str, Any] = {
            "direction": "rtl" if rtl else "ltr",
            "space_after_pt": 2,
            "space_before_pt": 0,
            "tabs": [{"align": "end", "pos_cm": round(width_cm, 2)}],
        }
        rule = (
            {"style": "double", "color": color, "size_pt": 0.75, "space_pt": 3}
            if style == "boxed"
            else {"style": "single", "color": color, "size_pt": 0.75, "space_pt": 3}
        )
        props["border_bottom"] = rule
        p = make_paragraph(self._new_para_id(), props)
        text_runs = content_runs(text, {**base, "color": "000000"})
        num_runs = self._folio_runs(part, style, base, color)
        # odd (default): text at start, number at end (outer side for RTL and
        # LTR odd pages); even: number at start, text at end.
        first, second = (text_runs, num_runs) if variant == "default" else (num_runs, text_runs)
        for r in first:
            p.append(r)
        p.append(make_run("\t", base))
        for r in second:
            p.append(r)
        return p

    def _folio_runs(self, part, style, base, color):
        runs = content_runs("{PAGE}", {**base, "color": color})
        if style == "boxed":
            pad = make_run(" ", {**base, "color": color})
            runs = [pad, *runs, copy.deepcopy(pad)]
            for r in runs:
                rpr = r.find(f"{W}rPr")
                if rpr is None:
                    rpr = etree.SubElement(r, f"{W}rPr")
                    r.insert(0, rpr)
                bdr = etree.SubElement(rpr, f"{W}bdr")
                bdr.set(f"{W}val", "single")
                bdr.set(f"{W}sz", "8")
                bdr.set(f"{W}space", "1")
                bdr.set(f"{W}color", color)
        elif style == "vignette":
            h_cm = 0.62
            out = []
            for side in ("right", "left"):
                svg = _bracket_svg("#" + color, side)
                vw, vh = svg_size(svg)
                r = etree.Element(f"{W}r")
                r.append(
                    self._inline_drawing(
                        part,
                        svg,
                        int(h_cm * vw / vh * EMU_PER_CM),
                        int(h_cm * EMU_PER_CM),
                        "Folio ornament",
                        "",
                    )
                )
                out.append(r)
            sp = make_run(" ", base)
            runs = [out[0], sp, *runs, copy.deepcopy(sp), out[1]]
        return runs

    def _cartouche_head(self, part, text, base, color, width_cm, rtl):
        w = min(width_cm * 0.62, 9.0)
        h = 0.95
        art = ia.cartouche(
            "banner", color="#ffffff", border="#" + color, aspect=w / h, knob=False, finials=True
        )
        vw, vh = svg_size(art)
        draw_w, draw_h = w, w * vh / vw
        line_pt = base["size_pt"] * 1.3
        pad_pt = max(0.0, draw_h / 2.54 * 72 - line_pt) / 2 + 1
        p = make_paragraph(
            self._new_para_id(),
            {
                "alignment": "center",
                "direction": "rtl" if rtl else "ltr",
                "space_before_pt": round(pad_pt, 1),
                "space_after_pt": round(pad_pt, 1),
                "line_exact_pt": round(line_pt, 1),
            },
        )
        for r in content_runs(text, {**base, "color": color}):
            p.append(r)
        self._anchor_behind(part, p, art, draw_w, draw_h, line_pt, pad_pt, "Head cartouche")
        return p

    def _medallion_folio(self, part, base, color, rtl):
        size = 1.05
        art = ia.star_medallion("#" + color, rings=2)
        line_pt = base["size_pt"] * 1.2
        pad_pt = max(0.0, size / 2.54 * 72 - line_pt) / 2 + 1
        p = make_paragraph(
            self._new_para_id(),
            {
                "alignment": "center",
                "direction": "rtl" if rtl else "ltr",
                "space_before_pt": round(pad_pt, 1),
                "space_after_pt": 0,
                "line_exact_pt": round(line_pt, 1),
            },
        )
        for r in content_runs("{PAGE}", {**base, "color": color, "size_pt": base["size_pt"] - 1}):
            p.append(r)
        self._anchor_behind(part, p, art, size, size, line_pt, pad_pt, "Folio medallion")
        return p

    # ── note marks ──────────────────────────────────────────────────────

    def set_note_mark_style(self, style: str = "parentheses", kind: str = "footnote") -> dict:
        """Note numbers as «(١)» in the text and in the note — the Arabic
        publishing convention. Word has no such number format, so the live
        reference is wrapped in parenthesis runs (re-run after adding notes;
        it skips marks already wrapped). style: parentheses, brackets, plain.
        """
        if style not in ("parentheses", "brackets", "plain"):
            raise ValueError("style must be parentheses, brackets or plain")
        if kind not in ("footnote", "endnote"):
            raise ValueError("kind must be footnote or endnote")
        opening, closing = {"parentheses": ("(", ")"), "brackets": ("[", "]"), "plain": ("", "")}[
            style
        ]
        rtl_doc = self._is_rtl()
        tag_ref = f"{W}{kind}Reference"
        tag_mark = f"{W}{kind}Ref"
        done = 0

        def text_of(r) -> str | None:
            if r is None or r.tag != f"{W}r":
                return None
            kids = [c for c in r if c.tag != f"{W}rPr"]
            if len(kids) == 1 and kids[0].tag == f"{W}t":
                return kids[0].text
            return None

        def wrap(run: etree._Element, superscript: bool) -> bool:
            parent = run.getparent()
            idx = list(parent).index(run)
            prev = parent[idx - 1] if idx > 0 else None
            nxt = parent[idx + 1] if idx + 1 < len(parent) else None
            # unwrap an earlier style
            if text_of(prev) in ("(", "[") and (text_of(nxt) or "").strip() in (")", "]"):
                parent.remove(prev)
                parent.remove(nxt)
                idx = list(parent).index(run)
            rpr = run.find(f"{W}rPr")
            if not superscript:
                if rpr is None:
                    rpr = etree.Element(f"{W}rPr")
                    run.insert(0, rpr)
                va = rpr.find(f"{W}vertAlign")
                if va is None:
                    va = etree.SubElement(rpr, f"{W}vertAlign")
                va.set(f"{W}val", "baseline")
            if not opening:
                return True
            after = nxt_text = None
            if not superscript:
                k = list(parent).index(run)
                nxt_run = parent[k + 1] if k + 1 < len(parent) else None
                nxt_text = text_of(nxt_run) or ""
                after = closing if nxt_text.startswith((" ", "\u00a0")) else closing + " "
                if rtl_doc and nxt_run is not None and not nxt_text.strip():
                    # a bare space between the mark and RTL text is weak; flag it
                    # RTL or Word moves it to the far side of the number
                    nrpr = nxt_run.find(f"{W}rPr")
                    if nrpr is None:
                        nrpr = etree.Element(f"{W}rPr")
                        nxt_run.insert(0, nrpr)
                    if nrpr.find(f"{W}rtl") is None:
                        etree.SubElement(nrpr, f"{W}rtl")
            for pos, ch in ((idx, opening), (idx + 2, after or closing)):
                r = etree.Element(f"{W}r")
                if rpr is not None:
                    r.append(copy.deepcopy(rpr))
                t = etree.SubElement(r, f"{W}t")
                t.text = ch
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                parent.insert(pos, r)
            return True

        doc = self._require("word/document.xml")
        for ref in list(doc.iter(tag_ref)):
            if ref.get(f"{W}customMarkFollows") in ("1", "true"):
                continue
            if wrap(ref.getparent(), superscript=True):
                done += 1
        notes = self._tree(f"word/{kind}s.xml")
        if notes is not None:
            for mark in list(notes.iter(tag_mark)):
                wrap(mark.getparent(), superscript=False)
            self._mark(f"word/{kind}s.xml")
        self._mark("word/document.xml")
        return {"style": style, "marks": done}

    # ── verse ───────────────────────────────────────────────────────────

    def insert_verses(
        self,
        verses: list[Any],
        *,
        after_para_id: str | None = None,
        numbered: bool = False,
        start_at: int = 1,
        color: str | None = None,
        number_color: str = "520912",
        font: str | None = None,
        size_pt: float | None = None,
        bold: bool = False,
        gap_cm: float = 1.0,
        separator: str | None = None,
    ) -> dict:
        """Poetry (nazm): each verse's two hemistichs side by side, each
        stretched across its half (Word "distribute"), in a borderless
        right-to-left table.

        verses: ["صدر *** عجز", …] (split on "***", "…" or "|"), or pairs.
        numbered: verse numbers («٤٨-») before each verse, in number_color.
        separator: a small mark between the hemistichs (e.g. "❁"), else a gap.
        """
        import re

        rows = []
        for v in verses:
            if isinstance(v, (list, tuple)):
                a, b = (v + ["", ""])[:2] if isinstance(v, list) else (tuple(v) + ("", ""))[:2]
            else:
                parts = re.split(r"\s*(?:\*\*\*|…|\|)\s*", str(v), maxsplit=1)
                a, b = (parts + [""])[:2]
            rows.append((a.strip(), b.strip()))
        if not rows:
            raise ValueError("Give at least one verse")
        rtl = self._is_rtl() or any(has_rtl(a + b) for a, b in rows)
        doc = self._require("word/document.xml")
        sect = doc.find(f"{W}body/{W}sectPr")
        width_tw = self._text_width_twips(sect) if sect is not None else 9000
        num_tw = 700 if numbered else 0
        mid_tw = int(gap_cm * 567)
        half = (width_tw - num_tw - mid_tw) // 2
        cols = ([num_tw] if numbered else []) + [half, mid_tw, half]

        tbl = etree.Element(f"{W}tbl")
        tpr = etree.SubElement(tbl, f"{W}tblPr")
        if rtl:
            etree.SubElement(tpr, f"{W}bidiVisual")
        tw = etree.SubElement(tpr, f"{W}tblW")
        tw.set(f"{W}w", str(sum(cols)))
        tw.set(f"{W}type", "dxa")
        jc = etree.SubElement(tpr, f"{W}jc")
        jc.set(f"{W}val", "center")
        borders = etree.SubElement(tpr, f"{W}tblBorders")
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            b = etree.SubElement(borders, f"{W}{side}")
            b.set(f"{W}val", "nil")
        lay = etree.SubElement(tpr, f"{W}tblLayout")
        lay.set(f"{W}type", "fixed")
        cm = etree.SubElement(tpr, f"{W}tblCellMar")
        for side in ("left", "right"):
            e = etree.SubElement(cm, f"{W}{side}")
            e.set(f"{W}w", "0")
            e.set(f"{W}type", "dxa")
        look = etree.SubElement(tpr, f"{W}tblLook")
        look.set(f"{W}val", "0000")
        grid = etree.SubElement(tbl, f"{W}tblGrid")
        for wv in cols:
            etree.SubElement(grid, f"{W}gridCol").set(f"{W}w", str(wv))

        run_props: dict[str, Any] = {}
        if color:
            run_props["color"] = color
        if font:
            run_props["font"] = font
            run_props["cs_font"] = font
        if size_pt:
            run_props["size_pt"] = size_pt
        if bold:
            run_props["bold"] = True
        if rtl:
            run_props["rtl"] = True
        digits = "٠١٢٣٤٥٦٧٨٩"

        def cell(width: int, text: str, align: str, props: dict) -> etree._Element:
            tc = etree.Element(f"{W}tc")
            tcpr = etree.SubElement(tc, f"{W}tcPr")
            w = etree.SubElement(tcpr, f"{W}tcW")
            w.set(f"{W}w", str(width))
            w.set(f"{W}type", "dxa")
            va = etree.SubElement(tcpr, f"{W}vAlign")
            va.set(f"{W}val", "bottom")
            p = make_paragraph(
                self._new_para_id(),
                {
                    "alignment": align,
                    "direction": "rtl" if rtl else "ltr",
                    "space_after_pt": 2,
                    "space_before_pt": 2,
                    "keep_together": True,
                },
            )
            if text:
                p.append(make_run(text, props or None))
            tc.append(p)
            return tc

        for i, (a, b) in enumerate(rows):
            tr = etree.SubElement(tbl, f"{W}tr")
            trpr = etree.SubElement(tr, f"{W}trPr")
            etree.SubElement(trpr, f"{W}cantSplit")
            if numbered:
                n = str(start_at + i)
                n = "".join(digits[int(ch)] for ch in n) if rtl else n
                tr.append(
                    cell(
                        num_tw, f"{n}-", "start", {**run_props, "color": number_color, "bold": True}
                    )
                )
            tr.append(cell(half, a, "distribute", run_props))
            tr.append(cell(mid_tw, separator or "", "center", {**run_props, "color": number_color}))
            tr.append(cell(half, b, "distribute", run_props))

        if after_para_id:
            anchor = self._find_para(doc, after_para_id)
            if anchor is None:
                raise ValueError(f"Paragraph '{after_para_id}' not found")
            parent = anchor.getparent()
            parent.insert(list(parent).index(anchor) + 1, tbl)
        else:
            body = doc.find(f"{W}body")
            kids = list(body)
            body.insert(len(kids) - 1 if kids and kids[-1].tag == f"{W}sectPr" else len(kids), tbl)
        # a paragraph after the table, so the next text has somewhere to go
        tail = make_paragraph(
            self._new_para_id(), {"space_after_pt": 0, "direction": "rtl" if rtl else "ltr"}
        )
        tbl.addnext(tail)
        self._mark("word/document.xml")
        return {"verses": len(rows), "para_id_after": tail.get(f"{W14}paraId")}

    # ── margin notes ────────────────────────────────────────────────────

    def add_margin_note(
        self,
        para_id: str,
        text: str,
        *,
        color: str = "520912",
        font: str | None = None,
        size_pt: float = 9,
        bold: bool = True,
        width_cm: float = 1.6,
        side: str = "outside",
    ) -> dict:
        """A side heading in the margin beside a paragraph (a text box
        anchored to it, so it travels with the text). side: outside (flips
        with odd/even pages), inside, left or right."""
        rel = {
            "outside": "outsideMargin",
            "inside": "insideMargin",
            "left": "leftMargin",
            "right": "rightMargin",
        }.get(side)
        if rel is None:
            raise ValueError("side must be outside, inside, left or right")
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")
        rtl = self._is_rtl() or has_rtl(text)
        lines = max(1, len(text) // 9 + 1)
        h_cm = max(0.6, lines * size_pt * 1.35 / 72 * 2.54)
        cx, cy = int(width_cm * EMU_PER_CM), int(h_cm * EMU_PER_CM)
        pid = self._next_docpr_id()
        run = etree.Element(f"{W}r")
        drawing = etree.SubElement(run, f"{W}drawing")
        an = etree.SubElement(
            drawing,
            f"{WP}anchor",
            nsmap={"wp": WP[1:-1]},
            distT="0",
            distB="0",
            distL="0",
            distR="0",
            simplePos="0",
            relativeHeight=str(251659264 + pid),
            behindDoc="0",
            locked="0",
            layoutInCell="1",
            allowOverlap="1",
        )
        etree.SubElement(an, f"{WP}simplePos", x="0", y="0")
        ph = etree.SubElement(an, f"{WP}positionH", relativeFrom=rel)
        etree.SubElement(ph, f"{WP}align").text = "center"
        pv = etree.SubElement(an, f"{WP}positionV", relativeFrom="paragraph")
        etree.SubElement(pv, f"{WP}posOffset").text = "0"
        etree.SubElement(an, f"{WP}extent", cx=str(cx), cy=str(cy))
        etree.SubElement(an, f"{WP}effectExtent", l="0", t="0", r="0", b="0")
        etree.SubElement(an, f"{WP}wrapNone")
        etree.SubElement(an, f"{WP}docPr", id=str(pid), name=f"Margin note {pid}")
        etree.SubElement(an, f"{WP}cNvGraphicFramePr")
        graphic = etree.SubElement(an, f"{A}graphic", nsmap={"a": A[1:-1]})
        gd = etree.SubElement(
            graphic,
            f"{A}graphicData",
            uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
        )
        wsp = etree.SubElement(gd, f"{WPS}wsp", nsmap={"wps": WPS[1:-1]})
        etree.SubElement(wsp, f"{WPS}cNvSpPr", txBx="1")
        sppr = etree.SubElement(wsp, f"{WPS}spPr")
        xfrm = etree.SubElement(sppr, f"{A}xfrm")
        etree.SubElement(xfrm, f"{A}off", x="0", y="0")
        etree.SubElement(xfrm, f"{A}ext", cx=str(cx), cy=str(cy))
        geom = etree.SubElement(sppr, f"{A}prstGeom", prst="rect")
        etree.SubElement(geom, f"{A}avLst")
        etree.SubElement(sppr, f"{A}noFill")
        ln = etree.SubElement(sppr, f"{A}ln")
        etree.SubElement(ln, f"{A}noFill")
        txbx = etree.SubElement(wsp, f"{WPS}txbx")
        content = etree.SubElement(txbx, f"{W}txbxContent")
        props: dict[str, Any] = {"color": color, "size_pt": size_pt, "bold": bold}
        if font:
            props["font"] = font
            props["cs_font"] = font
        if rtl:
            props["rtl"] = True
        np_ = make_paragraph(
            self._new_para_id(),
            {
                "alignment": "center",
                "direction": "rtl" if rtl else "ltr",
                "space_after_pt": 0,
                "space_before_pt": 0,
                "line_spacing": 1.0,
            },
        )
        np_.append(make_run(text, props))
        content.append(np_)
        body_pr = etree.SubElement(
            wsp,
            f"{WPS}bodyPr",
            rot="0",
            vert="horz",
            wrap="square",
            lIns="0",
            tIns="0",
            rIns="0",
            bIns="0",
            anchor="t",
        )
        etree.SubElement(body_pr, f"{A}spAutoFit")
        ppr = para.find(f"{W}pPr")
        para.insert(0 if ppr is None else list(para).index(ppr) + 1, run)
        self._mark("word/document.xml")
        return {"para_id": para_id, "text": text, "side": side}

    # ── numbered entries ────────────────────────────────────────────────

    def define_numbered_style(
        self,
        name: str,
        *,
        template: str = "%1 .",
        number_format: str = "decimal",
        start_at: int = 1,
        color: str = "520912",
        bold: bool = True,
        font: str | None = None,
        size_pt: float | None = None,
        aliases: list[str] | None = None,
        follow: str = "space",
    ) -> dict:
        """A paragraph style whose paragraphs number themselves — tarajim,
        narrations or masa'il like «١٨٨٠ . سليمان بن عمرو بن الأحوص:».

        Inserting or deleting an entry renumbers everything after it.
        template: the number's text, %1 being the number ("%1 .", "[%1]",
        "(%1)", "%1-"). number_format: decimal (shows ١٢٣ in Arabic text),
        arabic-abjad, arabic-alpha, roman. follow: space, tab or nothing.
        """
        fmts = {
            "decimal": "decimal",
            "arabic-abjad": "arabicAbjad",
            "arabic-alpha": "arabicAlpha",
            "roman": "upperRoman",
        }
        if number_format not in fmts:
            raise ValueError(f"number_format must be one of {sorted(fmts)}")
        if follow not in ("space", "tab", "nothing"):
            raise ValueError("follow must be space, tab or nothing")
        rtl = self._is_rtl()
        run: dict[str, Any] = {"bold": bold, "color": color}
        if font:
            run["font"] = font
            run["cs_font"] = font
        if size_pt:
            run["size_pt"] = size_pt
        res = self.define_style(
            name,
            "paragraph",
            based_on="Normal",
            next_style="Normal",
            aliases=aliases,
            gallery=True,
            run=run,
            paragraph={
                "keep_with_next": True,
                "space_before_pt": 6,
                "direction": "rtl" if rtl else "ltr",
            },
        )
        sid = res["style_id"]
        numbering = self._tree("word/numbering.xml")
        if numbering is None:
            from .creation import _ensure_numbering

            _ensure_numbering(self)
            numbering = self._require("word/numbering.xml")
        abs_ids = [int(a.get(f"{W}abstractNumId")) for a in numbering.findall(f"{W}abstractNum")]
        num_ids = [int(n.get(f"{W}numId")) for n in numbering.findall(f"{W}num")]
        aid = max(abs_ids, default=-1) + 1
        nid = max(num_ids, default=0) + 1
        absn = etree.Element(f"{W}abstractNum")
        absn.set(f"{W}abstractNumId", str(aid))
        etree.SubElement(absn, f"{W}multiLevelType").set(f"{W}val", "singleLevel")
        lvl = etree.SubElement(absn, f"{W}lvl")
        lvl.set(f"{W}ilvl", "0")
        etree.SubElement(lvl, f"{W}start").set(f"{W}val", str(start_at))
        etree.SubElement(lvl, f"{W}numFmt").set(f"{W}val", fmts[number_format])
        etree.SubElement(lvl, f"{W}pStyle").set(f"{W}val", sid)
        etree.SubElement(lvl, f"{W}suff").set(f"{W}val", follow)
        etree.SubElement(lvl, f"{W}lvlText").set(f"{W}val", template)
        etree.SubElement(lvl, f"{W}lvlJc").set(f"{W}val", "left")
        lppr = etree.SubElement(lvl, f"{W}pPr")
        ind = etree.SubElement(lppr, f"{W}ind")
        ind.set(f"{W}left", "0")
        ind.set(f"{W}firstLine", "0")
        lrpr = etree.SubElement(lvl, f"{W}rPr")
        from .ooxml import apply_run_props

        apply_run_props(lrpr, run)
        first_num = numbering.find(f"{W}num")
        if first_num is not None:
            first_num.addprevious(absn)
        else:
            numbering.append(absn)
        num = etree.SubElement(numbering, f"{W}num")
        num.set(f"{W}numId", str(nid))
        etree.SubElement(num, f"{W}abstractNumId").set(f"{W}val", str(aid))
        self._mark("word/numbering.xml")
        styles = self._require("word/styles.xml")
        st = next(s for s in styles.findall(f"{W}style") if s.get(f"{W}styleId") == sid)
        ppr = st.find(f"{W}pPr")
        if ppr is None:
            ppr = etree.SubElement(st, f"{W}pPr")
        old = ppr.find(f"{W}numPr")
        if old is not None:
            ppr.remove(old)
        numpr = etree.SubElement(ppr, f"{W}numPr")
        etree.SubElement(numpr, f"{W}numId").set(f"{W}val", str(nid))
        self._mark("word/styles.xml")
        return {"style_id": sid, "num_id": nid, "template": template}

    # ── part-title page ─────────────────────────────────────────────────

    def insert_part_title(
        self,
        title: str,
        *,
        after_para_id: str | None = None,
        subtitle: str | None = None,
        items: list[str] | None = None,
        color: str = "520912",
        font: str | None = None,
        size_pt: float = 30,
        frame: bool = True,
        ornament: str | None = "divider",
    ) -> dict:
        """A part-title page (a «الفهارس» or «القسم الثاني» page): its own
        section, vertically centred, with a woven frame drawn on that page
        only. items: lines listed under the title (e.g. the indexes that
        follow)."""
        doc = self._require("word/document.xml")
        body = doc.find(f"{W}body")
        body_sect = body.find(f"{W}sectPr")
        if body_sect is None:
            raise ValueError("Document has no section properties")
        rtl = self._is_rtl() or has_rtl(title)
        dirn = "rtl" if rtl else "ltr"

        def sect_copy(center: bool) -> etree._Element:
            s = copy.deepcopy(body_sect)
            for tag in ("headerReference", "footerReference", "titlePg"):
                for e in s.findall(f"{W}{tag}"):
                    if center:
                        s.remove(e)
            t = s.find(f"{W}type")
            if t is None:
                t = etree.SubElement(s, f"{W}type")
            t.set(f"{W}val", "nextPage")
            if center:
                va = s.find(f"{W}vAlign")
                if va is None:
                    va = etree.SubElement(s, f"{W}vAlign")
                va.set(f"{W}val", "center")
            return s

        # 1. end the previous section on the anchor paragraph
        if after_para_id:
            anchor = self._find_para(doc, after_para_id)
            if anchor is None:
                raise ValueError(f"Paragraph '{after_para_id}' not found")
        else:
            kids = [k for k in body if k.tag == f"{W}p"]
            anchor = kids[-1] if kids else None
            if anchor is None:
                anchor = make_paragraph(self._new_para_id(), {})
                body.insert(0, anchor)
        appr = anchor.find(f"{W}pPr")
        if appr is None:
            appr = etree.Element(f"{W}pPr")
            anchor.insert(0, appr)
        if appr.find(f"{W}sectPr") is None:
            appr.append(sect_copy(center=False))

        # 2. the title block
        block: list[etree._Element] = []
        tprops: dict[str, Any] = {"bold": True, "size_pt": size_pt, "color": color}
        if font:
            tprops["font"] = font
            tprops["cs_font"] = font
        if rtl:
            tprops["rtl"] = True
        tp = make_paragraph(
            self._new_para_id(),
            {
                "alignment": "center",
                "direction": dirn,
                "space_after_pt": 12,
                "keep_with_next": True,
            },
        )
        tp.append(make_run(title, tprops))
        block.append(tp)
        if subtitle:
            sp = make_paragraph(
                self._new_para_id(),
                {"alignment": "center", "direction": dirn, "space_after_pt": 12},
            )
            sp.append(
                make_run(
                    subtitle,
                    {**tprops, "size_pt": size_pt * 0.55, "bold": False, "color": "000000"},
                )
            )
            block.append(sp)
        for it in items or []:
            ip = make_paragraph(
                self._new_para_id(), {"alignment": "center", "direction": dirn, "space_after_pt": 8}
            )
            ip.append(
                make_run(
                    "◆  " + it if not rtl else it + "  ◆",
                    {**tprops, "size_pt": size_pt * 0.55, "color": "000000"},
                )
            )
            block.append(ip)
        pos = list(anchor.getparent()).index(anchor) + 1
        for k, el in enumerate(block):
            anchor.getparent().insert(pos + k, el)
        last = block[-1]
        if ornament:
            last_id = last.get(f"{W14}paraId")
            r = self.insert_ornament(ornament, after_para_id=last_id, color=color)
            last = self._find_para(doc, r["para_id"])
        frame_art = None
        if frame:
            pg = body_sect.find(f"{W}pgSz")
            pw = int(pg.get(f"{W}w", "11906")) / 567.0
            phh = int(pg.get(f"{W}h", "16838")) / 567.0
            m = 1.0
            frame_art = (
                ia.frame("#" + (color_hex(color) or "520912"), aspect=(pw - 2 * m) / (phh - 2 * m)),
                pw,
                phh,
                m,
            )
        # 3. close the centred section on the block's last paragraph
        lppr = last.find(f"{W}pPr")
        if lppr is None:
            lppr = etree.Element(f"{W}pPr")
            last.insert(0, lppr)
        title_sect = sect_copy(center=True)
        lppr.append(title_sect)
        # Empty heads on the title page; the frame is drawn from its header
        # (anchored in the body it would stop Word centring the page).
        for kind in ("header", "footer"):
            for variant in ("default", "even"):
                path, root, rid = self._new_story_part(kind)
                hp = make_paragraph(self._new_para_id(), {"space_after_pt": 0})
                if kind == "header" and frame_art is not None:
                    art, pw, phh, m = frame_art
                    run = etree.SubElement(hp, f"{W}r")
                    run.append(
                        self._anchor_drawing(
                            path,
                            art,
                            int((pw - 2 * m) * EMU_PER_CM),
                            int((phh - 2 * m) * EMU_PER_CM),
                            int(m * EMU_PER_CM),
                            int(m * EMU_PER_CM),
                            "Part frame",
                            behind=True,
                            rel_from="page",
                        )
                    )
                root.append(hp)
                self._attach(title_sect, kind, variant, rid)
        # The section after the title page must not inherit the frame: give it
        # explicit (empty) heads where it has none of its own.
        have = {
            (e.tag.split("}")[1][:6], e.get(f"{W}type"))
            for e in body_sect
            if e.tag in (f"{W}headerReference", f"{W}footerReference")
        }
        for kind in ("header", "footer"):
            for variant in ("default", "even"):
                if (kind, variant) not in have:
                    _, root, rid = self._new_story_part(kind)
                    root.append(make_paragraph(self._new_para_id(), {"space_after_pt": 0}))
                    self._attach(body_sect, kind, variant, rid)
        apply_para_props(lppr, {})
        self._mark("word/document.xml")
        return {"para_id": tp.get(f"{W14}paraId"), "after_para_id": last.get(f"{W14}paraId")}


def svg_size(svg: str) -> tuple[float, float]:
    from .vectors import _svg_size

    return _svg_size(svg)
