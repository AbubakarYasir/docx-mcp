"""Page design: headers and footers, page numbering, book layout.

* ``design_header_footer`` builds a header or footer from up to three zones
  (start / centre / end) with fields — page numbers, page X of Y, and
  running heads that follow the current chapter (Word's STYLEREF field, a
  trick that saves retyping every chapter title in the header). Variants for
  the first page and for even pages; a rule line above or below.
* ``set_page_numbering``  number format (decimal, roman, Arabic abjad /
  alphabetic, letters, "- 1 -") and restart value, per section.
* ``set_book_layout``  page size presets (A4, A5, B5, 17×24 Arabic book,
  14×21, Royal, Letter), mirror margins with inside/outside, gutter, and a
  right-hand binding gutter for right-to-left books.
"""

from __future__ import annotations

from typing import Any

from lxml import etree

from .ooxml import (
    CT_BASE,
    R_NS,
    W_NS,
    R,
    W,
    add_relationship,
    cm_to_twips,
    color_hex,
    content_runs,
    ensure_override,
    make_paragraph,
    unique_part_name,
)

PAGE_SIZES_MM = {
    "A4": (210, 297),
    "A5": (148, 210),
    "B5": (176, 250),
    "17x24": (170, 240),
    "14x21": (140, 210),
    "16x23": (160, 230),
    "royal": (156, 234),
    "letter": (215.9, 279.4),
}

NUMBER_FORMATS = {
    "decimal": "decimal",
    "roman": "lowerRoman",
    "upper-roman": "upperRoman",
    "letters": "lowerLetter",
    "upper-letters": "upperLetter",
    "arabic-abjad": "arabicAbjad",  # أ ب ج د (abjad order)
    "arabic-alpha": "arabicAlpha",  # أ ب ت ث (alphabetical order)
    "dash": "numberInDash",  # - 1 -
}

VARIANTS = ("default", "first", "even")


def _mm_twips(mm: float) -> int:
    return round(float(mm) * 1440 / 25.4)


class PageDesignMixin:
    """Headers/footers, page numbering and book page layout."""

    # ── helpers ──────────────────────────────────────────────────────────

    def _sections(self) -> list[etree._Element]:
        doc = self._require("word/document.xml")
        body = doc.find(f"{W}body")
        sects = [s for s, _ in self._collect_sectprs(body)]
        if not sects:
            sects = [etree.SubElement(body, f"{W}sectPr")]
            self._mark("word/document.xml")
        return sects

    def _pick_sections(self, section_index: int | None) -> list[etree._Element]:
        sects = self._sections()
        if section_index is None:
            return sects
        if not 0 <= section_index < len(sects):
            raise ValueError(f"section_index {section_index} out of range (0..{len(sects) - 1})")
        return [sects[section_index]]

    def _text_width_twips(self, sect: etree._Element) -> int:
        pg = sect.find(f"{W}pgSz")
        mar = sect.find(f"{W}pgMar")
        width = int(pg.get(f"{W}w", "11906")) if pg is not None else 11906
        if mar is None:
            return width - 2880
        left = int(mar.get(f"{W}left", "1440"))
        right = int(mar.get(f"{W}right", "1440"))
        gutter = int(mar.get(f"{W}gutter", "0"))
        return max(1440, width - left - right - gutter)

    def _new_story_part(self, kind: str) -> tuple[str, etree._Element, str]:
        """Create word/headerN.xml or word/footerN.xml; return (path, root, rId)."""
        tag = "hdr" if kind == "header" else "ftr"
        existing = {p.split("/")[-1] for p in self._trees} | (
            {p.name for p in (self.workdir / "word").glob(f"{kind}*.xml")}
            if self.workdir
            else set()
        )
        name = unique_part_name(existing, kind)
        path = f"word/{name}"
        root = etree.Element(f"{W}{tag}", nsmap={"w": W_NS, "r": R_NS})
        self._trees[path] = root
        ct = self._tree("[Content_Types].xml")
        ensure_override(ct, f"/{path}", f"{CT_BASE}{kind}+xml")
        self._mark("[Content_Types].xml")
        rels = self._tree("word/_rels/document.xml.rels")
        rid = add_relationship(rels, kind, name)
        self._mark("word/_rels/document.xml.rels")
        self._mark(path)
        return path, root, rid

    def _attach(self, sect: etree._Element, kind: str, variant: str, rid: str) -> None:
        ref_tag = f"{W}{kind}Reference"
        for ref in sect.findall(ref_tag):
            if ref.get(f"{W}type", "default") == variant:
                sect.remove(ref)
        ref = etree.Element(ref_tag)
        ref.set(f"{W}type", variant)
        ref.set(f"{R}id", rid)
        sect.insert(0, ref)
        if variant == "first" and sect.find(f"{W}titlePg") is None:
            etree.SubElement(sect, f"{W}titlePg")
        if variant == "even":
            settings = self._tree("word/settings.xml")
            if settings is not None and settings.find(f"{W}evenAndOddHeaders") is None:
                settings.insert(0, etree.Element(f"{W}evenAndOddHeaders"))
                self._mark("word/settings.xml")

    # ── headers and footers ──────────────────────────────────────────────

    def design_header_footer(
        self,
        kind: str = "header",
        *,
        start: str = "",
        center: str = "",
        end: str = "",
        variant: str = "default",
        section_index: int | None = None,
        direction: str = "auto",
        font: str | None = None,
        cs_font: str | None = None,
        size_pt: float | None = None,
        color: str | None = None,
        bold: bool | None = None,
        rule: str = "none",
        rule_color: str = "000000",
        rule_style: str = "single",
        rule_size_pt: float = 0.75,
        distance_cm: float | None = None,
        zone_formats: dict[str, dict[str, Any]] | None = None,
    ) -> dict:
        """Design a header or footer from three zones with live fields.

        Zones: ``start`` (right side in an RTL document, left in LTR),
        ``center``, ``end``. Text may contain {PAGE}, {NUMPAGES},
        {SECTIONPAGES}, {DATE}, {TITLE}, {STYLEREF:Heading 1} (running head:
        the current chapter's title, updated by Word), {FIELD:any Word field}.

        Args:
            kind: "header" or "footer".
            variant: "default", "first" (first page of the section) or
                "even" (even pages; odd pages use default). Turns on the
                matching Word setting automatically.
            section_index: One section, or None for every section.
            direction: "rtl", "ltr" or "auto" (rtl if any zone has Arabic).
            font / cs_font / size_pt / color / bold: formatting for all zones.
            rule: "below", "above", "both" or "none" — a line separating the
                header/footer from the page text.
            rule_color / rule_style (single, double, thick, dotted, dashed,
                thinThickSmallGap…) / rule_size_pt.
            distance_cm: Header (or footer) distance from the page edge.
            zone_formats: per-zone overrides, e.g. {"center": {"bold": True}}.
        """
        if kind not in ("header", "footer"):
            raise ValueError("kind must be 'header' or 'footer'")
        if variant not in VARIANTS:
            raise ValueError(f"variant must be one of {VARIANTS}")
        from .bidi import has_rtl

        zones = {"start": start, "center": center, "end": end}
        if direction == "auto":
            from .bidi import document_rtl_default, first_strong

            strong = [first_strong(z) for z in zones.values() if z]
            strong = [d for d in strong if d]
            if strong:
                direction = (
                    "rtl" if "rtl" in strong or any(has_rtl(z) for z in zones.values()) else "ltr"
                )
            else:
                direction = "rtl" if document_rtl_default(self._tree("word/styles.xml")) else "ltr"
        base: dict[str, Any] = {}
        if font:
            base["font"] = font
        if cs_font:
            base["cs_font"] = cs_font
        if size_pt is not None:
            base["size_pt"] = size_pt
        if color:
            base["color"] = color_hex(color)
        if bold is not None:
            base["bold"] = bold
        if direction == "rtl":
            base["rtl"] = True

        def runs_for(zone: str) -> list:
            props = {**base, **((zone_formats or {}).get(zone, {}))}
            return content_runs(zones[zone], props or None)

        results = []
        for sect in self._pick_sections(section_index):
            path, root, rid = self._new_story_part(kind)
            width = self._text_width_twips(sect)
            filled = [z for z in ("start", "center", "end") if zones[z]]
            para_props: dict[str, Any] = {
                "direction": direction,
                "space_after_pt": 0,
                "space_before_pt": 0,
            }
            rules = {
                "below": ("border_bottom",),
                "above": ("border_top",),
                "both": ("border_top", "border_bottom"),
            }.get(rule, ())
            for side in rules:
                para_props[side] = {
                    "style": rule_style,
                    "color": rule_color,
                    "size_pt": rule_size_pt,
                    "space_pt": 4,
                }
            if filled == ["center"] or not filled:
                para_props["alignment"] = "center"
                p = make_paragraph(self._new_para_id(), para_props, runs_for("center"))
            else:
                # Tab stops: centre and far end of the text area. Start text
                # goes first, then a tab to the centre, then a tab to the end.
                para_props["tabs"] = [
                    {"pos_cm": width / 2 * 2.54 / 1440, "align": "center"},
                    {"pos_cm": width * 2.54 / 1440, "align": "end"},
                ]
                p = make_paragraph(self._new_para_id(), para_props)
                for r in runs_for("start"):
                    p.append(r)
                tab = content_runs("\t", base or None)
                if zones["center"] or zones["end"]:
                    p.extend(tab)
                    for r in runs_for("center"):
                        p.append(r)
                if zones["end"]:
                    p.extend(content_runs("\t", base or None))
                    for r in runs_for("end"):
                        p.append(r)
            root.append(p)
            self._attach(sect, kind, variant, rid)
            if distance_cm is not None:
                mar = sect.find(f"{W}pgMar")
                if mar is not None:
                    mar.set(f"{W}{kind}", str(cm_to_twips(distance_cm)))
            results.append({"part": path, "rId": rid})
        filled_even = self._fill_even_gaps()
        self._mark("word/document.xml")
        return {
            "kind": kind,
            "variant": variant,
            "direction": direction,
            "parts": results,
            "even_pages_reuse_default": filled_even,
        }

    def _fill_even_gaps(self) -> list[str]:
        """With different odd/even pages on, a header or footer that has no
        even variant leaves even pages blank (a common Word surprise). Reuse
        the default variant on even pages until an even one is designed."""
        settings = self._tree("word/settings.xml")
        if settings is None or settings.find(f"{W}evenAndOddHeaders") is None:
            return []
        filled = []
        for i, sect in enumerate(self._sections()):
            for kind in ("header", "footer"):
                refs = {
                    r.get(f"{W}type", "default"): r for r in sect.findall(f"{W}{kind}Reference")
                }
                if "even" in refs or "default" not in refs:
                    continue
                ref = etree.Element(f"{W}{kind}Reference")
                ref.set(f"{W}type", "even")
                ref.set(f"{R}id", refs["default"].get(f"{R}id"))
                sect.insert(0, ref)
                filled.append(f"section {i} {kind}")
        return filled

    def remove_header_footer(
        self, kind: str = "header", variant: str = "default", section_index: int | None = None
    ) -> dict:
        """Detach a header/footer variant from sections (the page shows none)."""
        removed = 0
        for sect in self._pick_sections(section_index):
            for ref in sect.findall(f"{W}{kind}Reference"):
                if ref.get(f"{W}type", "default") == variant:
                    sect.remove(ref)
                    removed += 1
        self._mark("word/document.xml")
        return {"removed": removed}

    # ── numbering and layout ─────────────────────────────────────────────

    def set_page_numbering(
        self,
        number_format: str = "decimal",
        *,
        start_at: int | None = None,
        section_index: int | None = None,
    ) -> dict:
        """Page number format and restart for a section (or all).

        number_format: decimal, roman, upper-roman, letters, upper-letters,
        arabic-abjad (أ ب ج د), arabic-alpha (أ ب ت ث), dash (- 1 -).
        Front matter in roman and the body restarting at 1 is two calls with
        section_index. Arabic-Indic digits (١٢٣) come from Word's own digit
        setting ("Context" numerals), not from this format.
        """
        if number_format not in NUMBER_FORMATS:
            raise ValueError(f"number_format must be one of {sorted(NUMBER_FORMATS)}")
        for sect in self._pick_sections(section_index):
            pn = sect.find(f"{W}pgNumType")
            if pn is None:
                pn = etree.SubElement(sect, f"{W}pgNumType")
            pn.set(f"{W}fmt", NUMBER_FORMATS[number_format])
            if start_at is not None:
                pn.set(f"{W}start", str(int(start_at)))
            elif f"{W}start" in pn.attrib:
                del pn.attrib[f"{W}start"]
        self._mark("word/document.xml")
        return {"format": number_format, "start_at": start_at, "section_index": section_index}

    def set_book_layout(
        self,
        *,
        page_size: str | None = None,
        width_mm: float | None = None,
        height_mm: float | None = None,
        top_cm: float | None = None,
        bottom_cm: float | None = None,
        inside_cm: float | None = None,
        outside_cm: float | None = None,
        gutter_cm: float | None = None,
        mirror_margins: bool | None = None,
        rtl_binding: bool | None = None,
        header_cm: float | None = None,
        footer_cm: float | None = None,
        section_index: int | None = None,
    ) -> dict:
        """Lay pages out as a printed book.

        Args:
            page_size: A4, A5, B5, 17x24 (common Arabic book), 14x21, 16x23,
                royal, letter — or width_mm/height_mm.
            top_cm / bottom_cm: Top and bottom margins.
            inside_cm / outside_cm: Margins at the binding and the outer edge
                (with mirror_margins, these swap on facing pages).
            gutter_cm: Extra binding space added to the inside margin.
            mirror_margins: Facing pages (left and right pages mirror).
            rtl_binding: Put the gutter on the right — for right-to-left
                books bound on the right.
            header_cm / footer_cm: Header and footer distance from the edge.
        """
        if page_size:
            key = page_size if page_size in PAGE_SIZES_MM else page_size.upper()
            if key not in PAGE_SIZES_MM:
                raise ValueError(f"page_size must be one of {sorted(PAGE_SIZES_MM)}")
            width_mm, height_mm = PAGE_SIZES_MM[key]
        for sect in self._pick_sections(section_index):
            if width_mm and height_mm:
                pg = sect.find(f"{W}pgSz")
                if pg is None:
                    pg = etree.SubElement(sect, f"{W}pgSz")
                pg.set(f"{W}w", str(_mm_twips(width_mm)))
                pg.set(f"{W}h", str(_mm_twips(height_mm)))
                pg.attrib.pop(f"{W}orient", None)
            mar = sect.find(f"{W}pgMar")
            if mar is None:
                mar = etree.SubElement(sect, f"{W}pgMar")
            # With mirror margins Word reads left as inside and right as outside.
            for attr, val in (
                ("top", top_cm),
                ("bottom", bottom_cm),
                ("left", inside_cm),
                ("right", outside_cm),
                ("gutter", gutter_cm),
                ("header", header_cm),
                ("footer", footer_cm),
            ):
                if val is not None:
                    mar.set(f"{W}{attr}", str(cm_to_twips(val)))
            if rtl_binding is not None:
                el = sect.find(f"{W}rtlGutter")
                if rtl_binding and el is None:
                    etree.SubElement(sect, f"{W}rtlGutter")
                elif not rtl_binding and el is not None:
                    sect.remove(el)
        if mirror_margins is not None:
            settings = self._require("word/settings.xml")
            el = settings.find(f"{W}mirrorMargins")
            if mirror_margins and el is None:
                settings.insert(0, etree.Element(f"{W}mirrorMargins"))
            elif not mirror_margins and el is not None:
                settings.remove(el)
            self._mark("word/settings.xml")
        self._mark("word/document.xml")
        return {
            "page_mm": [width_mm, height_mm] if width_mm else None,
            "mirror_margins": mirror_margins,
            "rtl_binding": rtl_binding,
        }

    def set_update_fields_on_open(self, enabled: bool = True) -> dict:
        """Ask Word to refresh all fields (TOC, index, page refs) when the file opens."""
        settings = self._require("word/settings.xml")
        el = settings.find(f"{W}updateFields")
        if enabled:
            if el is None:
                el = etree.SubElement(settings, f"{W}updateFields")
            el.set(f"{W}val", "true")
        elif el is not None:
            settings.remove(el)
        self._mark("word/settings.xml")
        return {"update_fields_on_open": enabled}
