"""Style design: define complete styles, label and order them, audit usage.

Consistent Word documents come from styles, not from formatting each
paragraph by hand. These tools make styles first-class:

* ``define_style``  create or update a full style in one call (Latin and
  Arabic fonts, sizes, colours, spacing, borders, kashida justification, keep
  with next, outline level), with aliases (labels shown in Word's style list)
  and a place in the Styles gallery.
* ``organize_styles``  order the gallery, hide styles nobody uses.
* ``rename_style``  rename and update every reference to it.
* ``style_report``  which styles are used where, and which paragraphs carry
  direct formatting that should be a style.
* ``apply_style_where``  give a style to every paragraph matching a pattern,
  current style or text — the fastest way to clean up a manuscript.
* ``clear_direct_formatting``  strip hand formatting so styles take over.
"""

from __future__ import annotations

import re
from typing import Any

from lxml import etree

from .ooxml import (
    PARA_KEYS,
    RUN_KEYS,
    W14,
    W,
    apply_para_props,
    apply_run_props,
)

STYLE_TYPES = ("paragraph", "character", "table", "numbering")


def _style_id(name: str) -> str:
    sid = re.sub(r"[^0-9A-Za-z؀-ۿ]", "", name)
    return sid or "Style"


def _all_styles(styles: etree._Element) -> list[etree._Element]:
    return styles.findall(f"{W}style")


def _find(styles: etree._Element, name_or_id: str) -> etree._Element | None:
    lo = name_or_id.lower()
    for s in _all_styles(styles):
        if (s.get(f"{W}styleId") or "").lower() == lo:
            return s
        n = s.find(f"{W}name")
        if n is not None and (n.get(f"{W}val") or "").lower() == lo:
            return s
        al = s.find(f"{W}aliases")
        if al is not None and lo in [
            a.strip().lower() for a in (al.get(f"{W}val") or "").split(",")
        ]:
            return s
    return None


def _set_child(style: etree._Element, tag: str, val: str | None) -> None:
    el = style.find(f"{W}{tag}")
    if val is None:
        if el is not None:
            style.remove(el)
        return
    if el is None:
        el = etree.SubElement(style, f"{W}{tag}")
    el.set(f"{W}val", val)


def _flag(style: etree._Element, tag: str, on: bool | None) -> None:
    if on is None:
        return
    el = style.find(f"{W}{tag}")
    if on and el is None:
        etree.SubElement(style, f"{W}{tag}")
    elif not on and el is not None:
        style.remove(el)


def _used_style_ids(trees: dict) -> dict[str, int]:
    used: dict[str, int] = {}
    for part, root in trees.items():
        if root is None or not part.startswith("word/") or part == "word/styles.xml":
            continue
        for tag in ("pStyle", "rStyle", "tblStyle"):
            for el in root.iter(f"{W}{tag}"):
                v = el.get(f"{W}val")
                if v:
                    used[v] = used.get(v, 0) + 1
    return used


class StyleDesignMixin:
    """Create, update, organise and audit styles."""

    def define_style(
        self,
        name: str,
        style_type: str = "paragraph",
        *,
        based_on: str | None = None,
        next_style: str | None = None,
        linked_style: str | None = None,
        aliases: list[str] | None = None,
        gallery: bool | None = None,
        ui_priority: int | None = None,
        hidden: bool | None = None,
        run: dict[str, Any] | None = None,
        paragraph: dict[str, Any] | None = None,
        replace_formatting: bool = False,
    ) -> dict:
        """Create a style, or update it if it exists (matched by name, id or alias).

        Args:
            name: Display name, e.g. "Matn", "Isnad", "Heading 1", "Prophetic Speech".
            style_type: "paragraph", "character" or "table".
            based_on: Parent style name or id (inherits its formatting).
            next_style: Style of the paragraph that follows when pressing Enter.
            linked_style: Link a paragraph and a character style.
            aliases: Other names shown in Word, e.g. ["متن", "Body"] — a label.
            gallery: Show in the Home ▸ Styles gallery (Word "quick style").
            ui_priority: Order in style lists (lower comes first).
            hidden: Hide from style lists.
            run: Character formatting — keys: font, latin_font, cs_font,
                size_pt, cs_size_pt, bold, italic, underline, color, highlight,
                caps, small_caps, spacing_pt, rtl, language, cs_language.
            paragraph: Paragraph formatting — keys: alignment (start, end,
                center, justify, kashida, kashida-low/-high, distribute),
                direction, space_before_pt, space_after_pt, line_spacing,
                line_exact_pt, line_at_least_pt, indent_start_cm,
                indent_end_cm, first_line_cm, hanging_cm, keep_with_next,
                keep_together, page_break_before, widow_control,
                outline_level (0 = Heading 1, for TOC/navigation),
                border_top/bottom/start/end ({style,color,size_pt,space_pt}),
                shading, tabs, suppress_hyphenation.
            replace_formatting: Clear the style's existing run/paragraph
                formatting first instead of merging into it.
        """
        if style_type not in STYLE_TYPES:
            raise ValueError(f"style_type must be one of {STYLE_TYPES}")
        if paragraph and style_type == "character":
            raise ValueError("A character style cannot have paragraph formatting")
        if run and set(run) - RUN_KEYS:
            raise ValueError(f"Unknown run keys: {sorted(set(run) - RUN_KEYS)}")
        if paragraph and set(paragraph) - PARA_KEYS:
            raise ValueError(f"Unknown paragraph keys: {sorted(set(paragraph) - PARA_KEYS)}")
        styles = self._require("word/styles.xml")
        style = _find(styles, name)
        created = style is None
        if created:
            sid = _style_id(name)
            base = sid
            i = 2
            while _find(styles, sid) is not None:
                sid = f"{base}{i}"
                i += 1
            style = etree.SubElement(styles, f"{W}style")
            style.set(f"{W}type", style_type)
            style.set(f"{W}styleId", sid)
            _set_child(style, "name", name)
            if style_type == "character" and run is not None and "char_style" in run:
                raise ValueError("char_style is not valid inside a character style")
        elif style.get(f"{W}type") != style_type:
            raise ValueError(
                f"Style '{name}' exists as a {style.get(f'{W}type')} style, not {style_type}"
            )
        if based_on is not None:
            parent = _find(styles, based_on)
            if parent is None:
                raise ValueError(f"based_on style '{based_on}' not found")
            _set_child(style, "basedOn", parent.get(f"{W}styleId"))
        if next_style is not None:
            nxt = _find(styles, next_style)
            _set_child(style, "next", nxt.get(f"{W}styleId") if nxt is not None else next_style)
        if linked_style is not None:
            lk = _find(styles, linked_style)
            _set_child(style, "link", lk.get(f"{W}styleId") if lk is not None else linked_style)
        if aliases is not None:
            _set_child(style, "aliases", ",".join(a.strip() for a in aliases if a.strip()) or None)
        if ui_priority is not None:
            _set_child(style, "uiPriority", str(int(ui_priority)))
        _flag(style, "qFormat", gallery)
        if hidden is not None:
            _flag(style, "semiHidden", hidden)
            _flag(style, "unhideWhenUsed", hidden)
        if replace_formatting:
            for tag in ("rPr", "pPr"):
                el = style.find(f"{W}{tag}")
                if el is not None:
                    style.remove(el)
        if paragraph:
            ppr = style.find(f"{W}pPr")
            if ppr is None:
                ppr = etree.SubElement(style, f"{W}pPr")
            apply_para_props(ppr, paragraph)
        if run:
            rpr = style.find(f"{W}rPr")
            if rpr is None:
                rpr = etree.SubElement(style, f"{W}rPr")
            apply_run_props(rpr, run)
        self._mark("word/styles.xml")
        return {
            "style_id": style.get(f"{W}styleId"),
            "name": name,
            "type": style_type,
            "created": created,
        }

    def rename_style(self, name_or_id: str, new_name: str, *, new_id: bool = False) -> dict:
        """Rename a style. With new_id=True also change its id and every reference."""
        styles = self._require("word/styles.xml")
        style = _find(styles, name_or_id)
        if style is None:
            raise ValueError(f"Style '{name_or_id}' not found")
        if _find(styles, new_name) not in (None, style):
            raise ValueError(f"A style named '{new_name}' already exists")
        old_id = style.get(f"{W}styleId")
        _set_child(style, "name", new_name)
        refs = 0
        sid = old_id
        if new_id:
            sid = _style_id(new_name)
            style.set(f"{W}styleId", sid)
            for part, root in self._trees.items():
                if root is None or not part.startswith("word/"):
                    continue
                hit = False
                for tag in ("pStyle", "rStyle", "tblStyle", "basedOn", "next", "link"):
                    for el in root.iter(f"{W}{tag}"):
                        if el.get(f"{W}val") == old_id:
                            el.set(f"{W}val", sid)
                            refs += 1
                            hit = True
                if hit:
                    self._mark(part)
        self._mark("word/styles.xml")
        return {"old_id": old_id, "style_id": sid, "name": new_name, "references_updated": refs}

    def organize_styles(
        self,
        order: list[str] | None = None,
        *,
        hide_unused: bool = False,
        keep_visible: list[str] | None = None,
        gallery_only: list[str] | None = None,
    ) -> dict:
        """Tidy Word's style lists.

        Args:
            order: Style names in the order they should appear (sets uiPriority
                1, 2, 3 … and puts them in the gallery).
            hide_unused: Hide styles the document never uses (they come back
                automatically when used: Word's "unhide when used").
            keep_visible: Never hide these.
            gallery_only: Show only these styles in the Home ▸ Styles gallery.
        """
        styles = self._require("word/styles.xml")
        used = _used_style_ids(self._trees)
        ordered = []
        for i, nm in enumerate(order or [], start=1):
            s = _find(styles, nm)
            if s is None:
                raise ValueError(f"Style '{nm}' not found")
            _set_child(s, "uiPriority", str(i))
            _flag(s, "qFormat", True)
            _flag(s, "semiHidden", False)
            ordered.append(s.get(f"{W}styleId"))
        keep = {
            (_find(styles, k).get(f"{W}styleId") if _find(styles, k) is not None else k)
            for k in (keep_visible or [])
        } | set(ordered)
        # Styles that used styles inherit from count as used.
        by_id = {s.get(f"{W}styleId"): s for s in _all_styles(styles)}
        frontier = [sid for sid in list(used) + ordered if sid in by_id]
        seen = set(frontier)
        while frontier:
            b = by_id[frontier.pop()].find(f"{W}basedOn")
            bid = b.get(f"{W}val") if b is not None else None
            if bid and bid in by_id and bid not in seen:
                seen.add(bid)
                frontier.append(bid)
        used = {k: used.get(k, 1) for k in seen}
        hidden = []
        if hide_unused:
            for s in _all_styles(styles):
                sid = s.get(f"{W}styleId")
                if sid in used or sid in keep or s.get(f"{W}default") == "1":
                    continue
                _flag(s, "semiHidden", True)
                _flag(s, "unhideWhenUsed", True)
                _flag(s, "qFormat", False)
                hidden.append(sid)
        if gallery_only is not None:
            allow = {
                _find(styles, g).get(f"{W}styleId")
                for g in gallery_only
                if _find(styles, g) is not None
            }
            for s in _all_styles(styles):
                _flag(s, "qFormat", s.get(f"{W}styleId") in allow)
        self._mark("word/styles.xml")
        return {"ordered": ordered, "hidden": hidden}

    def style_report(self) -> dict:
        """How styles are used, and where direct formatting should become a style."""
        styles = self._require("word/styles.xml")
        doc = self._require("word/document.xml")
        used = _used_style_ids(self._trees)
        catalog = []
        for s in _all_styles(styles):
            sid = s.get(f"{W}styleId")
            n = s.find(f"{W}name")
            al = s.find(f"{W}aliases")
            catalog.append(
                {
                    "id": sid,
                    "name": n.get(f"{W}val") if n is not None else sid,
                    "type": s.get(f"{W}type"),
                    "aliases": al.get(f"{W}val") if al is not None else "",
                    "uses": used.get(sid, 0),
                    "gallery": s.find(f"{W}qFormat") is not None,
                    "hidden": s.find(f"{W}semiHidden") is not None,
                }
            )
        direct = []
        signature_counts: dict[str, int] = {}
        for p in doc.iter(f"{W}p"):
            fmt = set()
            ppr = p.find(f"{W}pPr")
            if ppr is not None:
                for c in ppr:
                    if c.tag not in (f"{W}pStyle", f"{W}rPr", f"{W}sectPr", f"{W}bidi"):
                        fmt.add(etree.QName(c).localname)
            for r in p.findall(f".//{W}r"):
                rpr = r.find(f"{W}rPr")
                if rpr is None:
                    continue
                for c in rpr:
                    if c.tag not in (f"{W}rStyle", f"{W}rtl", f"{W}lang"):
                        fmt.add("run:" + etree.QName(c).localname)
            if fmt:
                text = "".join(t.text or "" for t in p.iter(f"{W}t"))
                sig = ",".join(sorted(fmt))
                signature_counts[sig] = signature_counts.get(sig, 0) + 1
                if len(direct) < 40:
                    direct.append(
                        {"para_id": p.get(f"{W14}paraId"), "text": text[:50], "direct": sorted(fmt)}
                    )
        repeated = sorted(
            (
                {"formatting": k.split(","), "paragraphs": v}
                for k, v in signature_counts.items()
                if v > 2
            ),
            key=lambda x: -x["paragraphs"],
        )
        return {
            "styles": catalog,
            "unused": [c["name"] for c in catalog if c["uses"] == 0],
            "paragraphs_with_direct_formatting": sum(signature_counts.values()),
            "examples": direct,
            "repeated_direct_formatting_make_a_style": repeated[:15],
        }

    def apply_style_where(
        self,
        style: str,
        *,
        text_regex: str | None = None,
        current_style: str | None = None,
        para_ids: list[str] | None = None,
        clear_direct: bool = False,
    ) -> dict:
        """Apply a paragraph or character style to every match.

        For a paragraph style, matching paragraphs get it. For a character
        style, text matching ``text_regex`` is split out into runs with it
        (e.g. every «…» quotation, every [١] number).

        Args:
            style: Style name, id or alias.
            text_regex: Python regex over the paragraph text.
            current_style: Only paragraphs that currently have this style.
            para_ids: Only these paragraphs.
            clear_direct: Remove direct formatting from matched paragraphs.
        """
        styles = self._require("word/styles.xml")
        st = _find(styles, style)
        if st is None:
            raise ValueError(f"Style '{style}' not found")
        sid = st.get(f"{W}styleId")
        kind = st.get(f"{W}type")
        rx = re.compile(text_regex) if text_regex else None
        cur = None
        if current_style:
            cs = _find(styles, current_style)
            cur = cs.get(f"{W}styleId") if cs is not None else current_style
        wanted = {x.upper() for x in para_ids} if para_ids else None
        n = 0
        for part, root in list(self._trees.items()):
            if root is None or not part.startswith(
                ("word/document", "word/footnotes", "word/endnotes", "word/header", "word/footer")
            ):
                continue
            hit = False
            for p in list(root.iter(f"{W}p")):
                if wanted and (p.get(f"{W14}paraId") or "").upper() not in wanted:
                    continue
                ps = p.find(f"{W}pPr/{W}pStyle")
                if cur and (ps.get(f"{W}val") if ps is not None else "Normal") != cur:
                    continue
                text = "".join(t.text or "" for t in p.iter(f"{W}t"))
                if kind == "character":
                    if rx is None:
                        raise ValueError("A character style needs text_regex")
                    n += _char_style_matches(p, rx, sid)
                    hit = True
                    continue
                if rx and not rx.search(text):
                    continue
                ppr = p.find(f"{W}pPr")
                if ppr is None:
                    ppr = etree.Element(f"{W}pPr")
                    p.insert(0, ppr)
                if ps is None:
                    ps = etree.SubElement(ppr, f"{W}pStyle")
                ps.set(f"{W}val", sid)
                if clear_direct:
                    _clear_paragraph(p)
                n += 1
                hit = True
            if hit:
                self._mark(part)
        return {"style_id": sid, "type": kind, "applied": n}

    def clear_direct_formatting(
        self, para_ids: list[str] | None = None, *, keep: list[str] | None = None
    ) -> dict:
        """Remove hand-applied formatting so styles control the look.

        Args:
            para_ids: Paragraphs to clean. None = the whole body.
            keep: Run properties to keep, e.g. ["b", "i", "color", "rtl"].
                Direction (bidi/rtl), language, character styles and
                tracked-change records are always kept.
        """
        doc = self._require("word/document.xml")
        wanted = {x.upper() for x in para_ids} if para_ids else None
        n = 0
        for p in doc.iter(f"{W}p"):
            if wanted and (p.get(f"{W14}paraId") or "").upper() not in wanted:
                continue
            n += _clear_paragraph(p, set(keep or []))
        if n:
            self._mark("word/document.xml")
        return {"properties_removed": n}


_KEEP_PPR = {"pStyle", "rPr", "sectPr", "bidi", "numPr", "pPrChange"}
_KEEP_RPR = {"rStyle", "rtl", "lang", "rPrChange", "ins", "del", "cs"}


def _clear_paragraph(p: etree._Element, keep: set[str] | None = None) -> int:
    keep = keep or set()
    n = 0
    ppr = p.find(f"{W}pPr")
    if ppr is not None:
        for c in list(ppr):
            if etree.QName(c).localname not in _KEEP_PPR:
                ppr.remove(c)
                n += 1
    for r in p.iter(f"{W}r"):
        rpr = r.find(f"{W}rPr")
        if rpr is None:
            continue
        for c in list(rpr):
            ln = etree.QName(c).localname
            if ln not in _KEEP_RPR and ln not in keep:
                rpr.remove(c)
                n += 1
    return n


def _char_style_matches(p: etree._Element, rx: re.Pattern, sid: str) -> int:
    """Split simple text runs so regex matches carry the character style."""
    import copy

    from .ooxml import XML_SPACE

    count = 0
    for r in list(p.iter(f"{W}r")):
        kids = [c for c in r if c.tag != f"{W}rPr"]
        if len(kids) != 1 or kids[0].tag != f"{W}t" or not kids[0].text:
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
        for i, (chunk, styled) in enumerate(pieces):
            nr = etree.Element(f"{W}r")
            nrpr = copy.deepcopy(rpr) if rpr is not None else None
            if styled:
                if nrpr is None:
                    nrpr = etree.Element(f"{W}rPr")
                rs = nrpr.find(f"{W}rStyle")
                if rs is None:
                    rs = etree.SubElement(nrpr, f"{W}rStyle")
                rs.set(f"{W}val", sid)
            if nrpr is not None:
                nr.append(nrpr)
            nt = etree.SubElement(nr, f"{W}t")
            nt.text = chunk
            nt.set(XML_SPACE, "preserve")
            parent.insert(idx + i, nr)
        parent.remove(r)
        count += len(spans)
    return count
