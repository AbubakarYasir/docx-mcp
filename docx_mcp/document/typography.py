"""Rich text with several fonts and colours, format-by-pattern, and themes."""

from __future__ import annotations

import copy
import re
from typing import Any

from lxml import etree

from .ooxml import (
    CT_BASE,
    W14,
    XML_SPACE,
    W,
    add_relationship,
    apply_run_props,
    content_runs,
    ensure_override,
    make_paragraph,
)

_A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
A = f"{{{_A_NS}}}"
THEME_PATH = "word/theme/theme1.xml"

# A complete Office theme with an Arabic script font slot. Colours are a
# restrained book palette; fonts default to Amiri (Arabic) and Cambria/Calibri.
_THEME_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="docx-mcp">
<a:themeElements>
<a:clrScheme name="docx-mcp">
<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="1F2A44"/></a:dk2><a:lt2><a:srgbClr val="F4EFE6"/></a:lt2>
<a:accent1><a:srgbClr val="520912"/></a:accent1><a:accent2><a:srgbClr val="9C7A3C"/></a:accent2>
<a:accent3><a:srgbClr val="2F5233"/></a:accent3><a:accent4><a:srgbClr val="1F4E79"/></a:accent4>
<a:accent5><a:srgbClr val="7F6000"/></a:accent5><a:accent6><a:srgbClr val="595959"/></a:accent6>
<a:hlink><a:srgbClr val="1F4E79"/></a:hlink><a:folHlink><a:srgbClr val="520912"/></a:folHlink>
</a:clrScheme>
<a:fontScheme name="docx-mcp">
<a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface="Sakkal Majalla"/>
<a:font script="Arab" typeface="Sakkal Majalla"/></a:majorFont>
<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface="Amiri"/>
<a:font script="Arab" typeface="Amiri"/></a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Office">
<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
</a:fmtScheme>
</a:themeElements>
<a:objectDefaults/><a:extraClrSchemeLst/>
</a:theme>"""

PALETTES = {
    "scholarly-maroon": {
        "dk2": "1F2A44",
        "lt2": "F4EFE6",
        "accent1": "520912",
        "accent2": "9C7A3C",
        "accent3": "2F5233",
        "accent4": "1F4E79",
        "accent5": "7F6000",
        "accent6": "595959",
        "hlink": "1F4E79",
        "folHlink": "520912",
    },
    "emerald-gold": {
        "dk2": "0B2E24",
        "lt2": "F6F1E3",
        "accent1": "0F5132",
        "accent2": "B08D3C",
        "accent3": "1B4965",
        "accent4": "7A2E2E",
        "accent5": "556B2F",
        "accent6": "5E5E5E",
        "hlink": "0F5132",
        "folHlink": "7A2E2E",
    },
    "midnight-ink": {
        "dk2": "14213D",
        "lt2": "F1F1EE",
        "accent1": "14213D",
        "accent2": "C9A227",
        "accent3": "6B2737",
        "accent4": "2E5E4E",
        "accent5": "3D5A80",
        "accent6": "6C757D",
        "hlink": "3D5A80",
        "folHlink": "6B2737",
    },
    "manuscript": {
        "dk2": "3B2F2F",
        "lt2": "FAF3E0",
        "accent1": "8B1E1E",
        "accent2": "A67C2E",
        "accent3": "3E5641",
        "accent4": "2C3E50",
        "accent5": "6D4C41",
        "accent6": "7D7461",
        "hlink": "2C3E50",
        "folHlink": "8B1E1E",
    },
}


def _theme_color_names() -> tuple[str, ...]:
    return (
        "dk1",
        "lt1",
        "dk2",
        "lt2",
        "accent1",
        "accent2",
        "accent3",
        "accent4",
        "accent5",
        "accent6",
        "hlink",
        "folHlink",
    )


class TypographyMixin:
    """Rich multi-font text, find-and-format, theme fonts and palettes."""

    # ── rich text ────────────────────────────────────────────────────────

    def insert_rich_paragraph(
        self,
        after_para_id: str,
        spans: list[dict[str, Any]],
        *,
        paragraph: dict[str, Any] | None = None,
    ) -> dict:
        """Insert a paragraph built from spans, each with its own formatting.

        Args:
            after_para_id: Insert after this paragraph (body, notes, headers, footers).
            spans: [{"text": "...", ...run formatting}] — keys per span: font,
                latin_font, cs_font, size_pt, cs_size_pt, bold, italic,
                underline, color, highlight, shading, caps, small_caps,
                superscript, subscript, spacing_pt, char_style, rtl, language.
                Text may contain {PAGE}, {NUMPAGES}, {STYLEREF:Heading 1},
                {FIELD:...} tokens, tabs and newlines.
            paragraph: Paragraph formatting (style, alignment, direction,
                spacing, indents, borders …) — see define_style.
        """
        part, anchor = self._find_para_any(after_para_id)
        runs = []
        for span in spans:
            span = self._resolve_style_props(dict(span))
            text = str(span.pop("text", ""))
            runs.extend(content_runs(text, span or None))
        p = make_paragraph(self._new_para_id(), self._resolve_style_props(paragraph), runs)
        anchor.addnext(p)
        self._mark(part)
        return {"para_id": p.get(f"{W14}paraId"), "runs": len(runs), "part": part}

    def append_rich_text(self, para_id: str, spans: list[dict[str, Any]]) -> dict:
        """Append formatted spans to the end of an existing paragraph."""
        part, p = self._find_para_any(para_id)
        n = 0
        for span in spans:
            span = self._resolve_style_props(dict(span))
            text = str(span.pop("text", ""))
            for r in content_runs(text, span or None):
                p.append(r)
                n += 1
        self._mark(part)
        return {"para_id": para_id, "runs_added": n}

    def format_matches(
        self,
        pattern: str,
        *,
        regex: bool = False,
        whole_word: bool = False,
        ignore_tashkil: bool = True,
        parts: str = "all",
        **run_props: Any,
    ) -> dict:
        """Apply character formatting to every occurrence of a text or pattern.

        Examples: colour every ﷺ maroon; bold every «…» quotation
        (pattern r"«[^»]+»", regex=True); put every Qur'anic ﴿…﴾ in a
        Quran font; make [١] [٢] numbers maroon.

        Args:
            pattern: Text, or a Python regex when regex=True.
            regex: Treat pattern as a regular expression.
            whole_word: Match whole words only.
            ignore_tashkil: Literal patterns match vocalised text too
                («البخاري» finds «البُخَارِيُّ»).
            parts: "all" or "body".
            **run_props: Formatting keys as in insert_rich_paragraph spans.
        """
        if not run_props:
            raise ValueError("Give at least one formatting property")
        if regex:
            rx_src = pattern
        elif ignore_tashkil:
            from .arabic import loose_pattern

            rx_src = loose_pattern(pattern)
        else:
            rx_src = re.escape(pattern)
        if whole_word:
            rx_src = rf"(?<!\w)(?:{rx_src})(?!\w)"
        rx = re.compile(rx_src)
        run_props = self._resolve_style_props(run_props)
        prefixes = (
            ("word/document.xml",)
            if parts == "body"
            else ("word/document", "word/footnotes", "word/endnotes", "word/header", "word/footer")
        )
        total = 0
        for part, root in list(self._trees.items()):
            if root is None or not part.startswith(prefixes):
                continue
            n = 0
            for p in list(root.iter(f"{W}p")):
                n += _format_in_paragraph(p, rx, run_props)
            if n:
                self._mark(part)
                total += n
        return {"pattern": pattern, "formatted": total}

    # ── themes ────────────────────────────────────────────────────────────

    def _ensure_theme(self) -> etree._Element:
        theme = self._tree(THEME_PATH)
        if theme is not None:
            return theme
        theme = etree.fromstring(_THEME_XML.encode("utf-8"))
        self._trees[THEME_PATH] = theme
        ct = self._tree("[Content_Types].xml")
        ensure_override(
            ct, "/word/theme/theme1.xml", "application/vnd.openxmlformats-officedocument.theme+xml"
        )
        self._mark("[Content_Types].xml")
        rels = self._tree("word/_rels/document.xml.rels")
        add_relationship(rels, "theme", "theme/theme1.xml")
        self._mark("word/_rels/document.xml.rels")
        self._mark(THEME_PATH)
        _ = CT_BASE
        return theme

    def set_theme_fonts(
        self,
        *,
        heading_latin: str | None = None,
        body_latin: str | None = None,
        heading_arabic: str | None = None,
        body_arabic: str | None = None,
    ) -> dict:
        """Set the theme's heading and body fonts, including the Arabic script slot.

        Styles that use theme fonts (Word's "+Headings" / "+Body") follow
        these, so one change restyles the whole document. Creates a theme if
        the document has none.
        """
        theme = self._ensure_theme()
        fs = theme.find(f"{A}themeElements/{A}fontScheme")
        changed = {}
        for which, latin, arabic in (
            ("majorFont", heading_latin, heading_arabic),
            ("minorFont", body_latin, body_arabic),
        ):
            font = fs.find(f"{A}{which}")
            if latin:
                font.find(f"{A}latin").set("typeface", latin)
                changed[f"{which}.latin"] = latin
            if arabic:
                cs = font.find(f"{A}cs")
                if cs is None:
                    cs = etree.SubElement(font, f"{A}cs")
                cs.set("typeface", arabic)
                arab = None
                for f in font.findall(f"{A}font"):
                    if f.get("script") == "Arab":
                        arab = f
                if arab is None:
                    arab = etree.SubElement(font, f"{A}font", script="Arab")
                arab.set("typeface", arabic)
                changed[f"{which}.arabic"] = arabic
        self._mark(THEME_PATH)
        return {"theme_fonts": changed}

    def apply_color_palette(
        self, palette: str | None = None, colors: dict[str, str] | None = None
    ) -> dict:
        """Set theme colours from a named palette and/or explicit slots.

        Palettes: scholarly-maroon, emerald-gold, midnight-ink, manuscript.
        Slots: dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink.
        """
        values: dict[str, str] = {}
        if palette:
            if palette not in PALETTES:
                raise ValueError(f"palette must be one of {sorted(PALETTES)}")
            values.update(PALETTES[palette])
        values.update({k: v.lstrip("#").upper() for k, v in (colors or {}).items()})
        if not values:
            raise ValueError("Give a palette or colors")
        theme = self._ensure_theme()
        scheme = theme.find(f"{A}themeElements/{A}clrScheme")
        for slot, val in values.items():
            if slot not in _theme_color_names():
                raise ValueError(f"Unknown colour slot {slot!r}")
            el = scheme.find(f"{A}{slot}")
            if el is None:
                el = etree.SubElement(scheme, f"{A}{slot}")
            for c in list(el):
                el.remove(c)
            etree.SubElement(el, f"{A}srgbClr", val=val)
        self._mark(THEME_PATH)
        return {"palette": palette, "colors": values}

    def list_palettes(self) -> dict:
        """Built-in colour palettes for apply_color_palette."""
        return PALETTES


def _format_in_paragraph(p: etree._Element, rx: re.Pattern, props: dict) -> int:
    count = 0
    for r in list(p.iter(f"{W}r")):
        kids = [c for c in r if c.tag != f"{W}rPr"]
        if len(kids) != 1 or kids[0].tag != f"{W}t" or not kids[0].text:
            continue
        anc = r.getparent()
        if anc is not None and anc.tag in (f"{W}del", f"{W}moveFrom"):
            continue
        text = kids[0].text
        spans = [m.span() for m in rx.finditer(text) if m.end() > m.start()]
        if not spans:
            continue
        parent = r.getparent()
        idx = list(parent).index(r)
        rpr = r.find(f"{W}rPr")
        pieces = []
        pos = 0
        for s, e in spans:
            if s > pos:
                pieces.append((text[pos:s], False))
            pieces.append((text[s:e], True))
            pos = e
        if pos < len(text):
            pieces.append((text[pos:], False))
        for i, (chunk, hit) in enumerate(pieces):
            nr = etree.Element(f"{W}r")
            nrpr = copy.deepcopy(rpr) if rpr is not None else None
            if hit:
                if nrpr is None:
                    nrpr = etree.Element(f"{W}rPr")
                apply_run_props(nrpr, props)
            if nrpr is not None:
                nr.append(nrpr)
            nt = etree.SubElement(nr, f"{W}t")
            nt.text = chunk
            nt.set(XML_SPACE, "preserve")
            parent.insert(idx + i, nr)
        parent.remove(r)
        count += len(spans)
    return count
