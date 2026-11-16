"""Footnote and endnote design.

* ``set_note_continuation``  the classical "=" marks for a note that runs
  onto the next page: Word's continuation *notice* prints at the foot of the
  page where the note breaks, and its continuation *separator* prints above
  the continued part on the next page. Both are ordinary stories in
  footnotes.xml that almost nobody edits — so they can carry "=" (or any
  text) instead of the default rule.
* ``set_note_numbering``  number format (decimal, Arabic abjad /
  alphabetical, roman, letters, symbols *, †), restart every page / section,
  start value, and position.
* ``set_note_separator``  the rule above notes: Word's short default, a full
  text-width rule in any colour, none, or an ornament text.
* ``add_custom_mark_note``  a footnote whose mark is your own text — e.g.
  the narration number [١] carries the footnote, as in a musnad layout
  (Word's ``customMarkFollows``).
* ``convert_notes``  footnotes to endnotes or back.
"""

from __future__ import annotations

from typing import Any

from lxml import etree

from .ooxml import (
    W14,
    XML_SPACE,
    W,
    apply_para_props,
    apply_run_props,
    color_hex,
    make_run,
)

NOTE_FORMATS = {
    "decimal": "decimal",
    "arabic-abjad": "arabicAbjad",
    "arabic-alpha": "arabicAlpha",
    "roman": "lowerRoman",
    "upper-roman": "upperRoman",
    "letters": "lowerLetter",
    "upper-letters": "upperLetter",
    "symbols": "chicago",  # * † ‡ §
    "decimal-parentheses": "decimalEnclosedParen",
    "decimal-circle": "decimalEnclosedCircle",
}
RESTARTS = {"continuous": "continuous", "page": "eachPage", "section": "eachSect"}
POSITIONS = {
    "footnote": {"page-bottom": "pageBottom", "below-text": "beneathText"},
    "endnote": {"section-end": "sectEnd", "document-end": "docEnd"},
}


def _kind_parts(kind: str) -> tuple[str, str, str]:
    if kind not in ("footnote", "endnote"):
        raise ValueError("kind must be 'footnote' or 'endnote'")
    return f"word/{kind}s.xml", kind, f"{kind}Pr"


def _special(root: etree._Element, kind: str, note_type: str) -> etree._Element | None:
    for n in root.findall(f"{W}{kind}"):
        if n.get(f"{W}type") == note_type:
            return n
    return None


class NotesDesignMixin:
    """Continuation marks, numbering, separators, custom marks, conversion."""

    def _note_tree(self, kind: str) -> etree._Element:
        path, _, _ = _kind_parts(kind)
        return self._require(path)

    def _register_special(self, kind: str, note_id: int) -> None:
        """settings.xml lists the special notes (separators, notice) by id."""
        settings = self._require("word/settings.xml")
        pr = settings.find(f"{W}{kind}Pr")
        if pr is None:
            pr = etree.SubElement(settings, f"{W}{kind}Pr")
        ids = [n.get(f"{W}id") for n in pr.findall(f"{W}{kind}")]
        if str(note_id) not in ids:
            etree.SubElement(pr, f"{W}{kind}").set(f"{W}id", str(note_id))
        self._mark("word/settings.xml")

    def set_note_continuation(
        self,
        kind: str = "footnote",
        *,
        notice_text: str = "=",
        separator_text: str = "=",
        keep_rule: bool = True,
        font: str | None = None,
        cs_font: str | None = None,
        size_pt: float | None = None,
        color: str | None = None,
        bold: bool | None = None,
        notice_alignment: str = "end",
        separator_alignment: str = "start",
        direction: str = "auto",
        legacy_layout: bool = False,
    ) -> dict:
        """Mark notes that continue onto the next page.

        Args:
            kind: "footnote" or "endnote".
            notice_text: Printed at the foot of the page where a note breaks
                ("=" by Arabic convention, or e.g. "(continued)"). "" = none.
            separator_text: Printed above the continued part on the next page.
                "" = only the rule.
            keep_rule: Keep Word's continuation rule above the continued part.
            font / cs_font / size_pt / color / bold: formatting of the marks.
            notice_alignment / separator_alignment: start, end or center.
            direction: rtl, ltr or auto (from the document default).
            legacy_layout: Word 2013+ layout (compatibility mode 15) ignores the
                continuation separator and reuses the normal separator, so
                separator_text never prints. True switches the file to Word
                2010 layout (mode 14) so it does; Word then shows
                "Compatibility Mode" in the title bar. The notice at the foot
                of the page prints in every mode.
        """
        from .bidi import document_rtl_default

        root = self._note_tree(kind)
        if direction == "auto":
            direction = "rtl" if document_rtl_default(self._tree("word/styles.xml")) else "ltr"
        props: dict[str, Any] = {}
        for key, val in (
            ("font", font),
            ("cs_font", cs_font),
            ("size_pt", size_pt),
            ("bold", bold),
        ):
            if val is not None:
                props[key] = val
        if color:
            props["color"] = color_hex(color)
        if direction == "rtl":
            props["rtl"] = True

        def para(text: str, align: str) -> etree._Element:
            p = etree.Element(f"{W}p")
            p.set(f"{W14}paraId", self._new_para_id())
            p.set(f"{W14}textId", "77777777")
            ppr = etree.SubElement(p, f"{W}pPr")
            apply_para_props(
                ppr,
                {
                    "alignment": align,
                    "direction": direction,
                    "space_before_pt": 0,
                    "space_after_pt": 0,
                    "line_spacing": 1,
                },
            )
            if text:
                p.append(make_run(text, props or None))
            return p

        # Continuation separator (above the continued part).
        sep = _special(root, kind, "continuationSeparator")
        if sep is None:
            sep = etree.SubElement(root, f"{W}{kind}")
            sep.set(f"{W}type", "continuationSeparator")
            sep.set(f"{W}id", "0")
        for c in list(sep):
            sep.remove(c)
        if keep_rule:
            rule_p = para("", "start")
            r = etree.SubElement(rule_p, f"{W}r")
            etree.SubElement(r, f"{W}continuationSeparator")
            sep.append(rule_p)
        if separator_text or not keep_rule:
            sep.append(para(separator_text, separator_alignment))

        # Continuation notice (foot of the page where the note breaks).
        notice = _special(root, kind, "continuationNotice")
        if notice_text:
            if notice is None:
                used = {int(n.get(f"{W}id", "0")) for n in root.findall(f"{W}{kind}")}
                nid = max(used | {0}) + 1
                notice = etree.SubElement(root, f"{W}{kind}")
                notice.set(f"{W}type", "continuationNotice")
                notice.set(f"{W}id", str(nid))
            for c in list(notice):
                notice.remove(c)
            notice.append(para(notice_text, notice_alignment))
            self._register_special(kind, int(notice.get(f"{W}id")))
        elif notice is not None:
            root.remove(notice)
        # Make sure the separators are registered too (Word expects all three).
        for n in root.findall(f"{W}{kind}"):
            if n.get(f"{W}type") in ("separator", "continuationSeparator"):
                self._register_special(kind, int(n.get(f"{W}id")))
        self._mark(f"word/{kind}s.xml")
        if legacy_layout:
            self._set_compat_mode(14)
        mode = self._compat_mode()
        result = {
            "kind": kind,
            "notice": notice_text,
            "separator": separator_text,
            "rule": keep_rule,
            "direction": direction,
            "compatibility_mode": mode,
        }
        if separator_text and mode >= 15:
            result["note"] = (
                "Word 2013+ layout prints the notice but not separator_text; the rule "
                "above continued notes is the normal separator (style it with "
                "set_note_separator). Pass legacy_layout=True to print separator_text."
            )
        return result

    def _compat_mode(self) -> int:
        settings = self._tree("word/settings.xml")
        if settings is None:
            return 15
        for cs in settings.iter(f"{W}compatSetting"):
            if cs.get(f"{W}name") == "compatibilityMode":
                try:
                    return int(cs.get(f"{W}val", "15"))
                except ValueError:
                    return 15
        return 12

    def _set_compat_mode(self, mode: int) -> None:
        settings = self._require("word/settings.xml")
        compat = settings.find(f"{W}compat")
        if compat is None:
            compat = etree.SubElement(settings, f"{W}compat")
        for cs in compat.findall(f"{W}compatSetting"):
            if cs.get(f"{W}name") == "compatibilityMode":
                cs.set(f"{W}val", str(mode))
                break
        else:
            cs = etree.SubElement(compat, f"{W}compatSetting")
            cs.set(f"{W}name", "compatibilityMode")
            cs.set(f"{W}uri", "http://schemas.microsoft.com/office/word")
            cs.set(f"{W}val", str(mode))
        self._mark("word/settings.xml")

    def set_note_separator(
        self,
        kind: str = "footnote",
        style: str = "rule",
        *,
        width_pct: float = 65,
        text: str = "",
        color: str | None = None,
        thickness_pt: float = 0.6,
        direction: str = "auto",
    ) -> dict:
        """The line above the notes.

        style:
          rule     — a plain rule, width_pct of the text width, at the start
                     side (as printed books do: 60–70 %).
          double   — two hairlines.
          tapered  — a rule thinning from the text edge outward (vector).
          khatam   — hairline opening with a small eight-point star (vector).
          diamond  — hairline with a centred diamond and end beads (vector).
          swash    — a small calligraphic swash (vector).
          text     — an ornament in type, e.g. "⁂" or "✽ ✽ ✽".
          default  — Word's own short rule.  full — rule across the width.
          none     — no separator.
        In Word 2013+ the same line also sits above continued notes.
        """
        from .bidi import document_rtl_default

        styles = (
            "rule",
            "double",
            "tapered",
            "khatam",
            "diamond",
            "swash",
            "text",
            "default",
            "full",
            "none",
        )
        if style not in styles:
            raise ValueError(f"style must be one of {styles}")
        if not 10 <= width_pct <= 100:
            raise ValueError("width_pct must be 10–100")
        color = (
            color_hex(color)
            if color
            else (self._house("ink") if hasattr(self, "_house") else "000000")
        )
        root = self._note_tree(kind)
        sep = _special(root, kind, "separator")
        if sep is None:
            sep = etree.SubElement(root, f"{W}{kind}")
            sep.set(f"{W}type", "separator")
            sep.set(f"{W}id", "-1")
        for c in list(sep):
            sep.remove(c)
        if direction == "auto":
            direction = "rtl" if document_rtl_default(self._tree("word/styles.xml")) else "ltr"
        rtl = direction == "rtl"
        text_w = self._text_width_cm() if hasattr(self, "_text_width_cm") else 15.0
        width = 100 if style == "full" else width_pct
        cut_cm = round(text_w * (1 - width / 100), 2)
        p = etree.SubElement(sep, f"{W}p")
        p.set(f"{W14}paraId", self._new_para_id())
        p.set(f"{W14}textId", "77777777")
        ppr = etree.SubElement(p, f"{W}pPr")
        props: dict[str, Any] = {
            "space_before_pt": 2,
            "space_after_pt": 3,
            "line_spacing": 1,
            "direction": direction,
            "alignment": "start",
        }
        if style == "default":
            props.pop("alignment")
            r = etree.SubElement(p, f"{W}r")
            etree.SubElement(r, f"{W}separator")
        elif style in ("rule", "full", "double"):
            props["indent_end_cm"] = cut_cm
            props["border_bottom"] = {
                "style": "double" if style == "double" else "single",
                "color": color,
                "size_pt": thickness_pt * (0.75 if style == "double" else 1),
                "space_pt": 1,
            }
            props["line_exact_pt"] = 2
        elif style in ("tapered", "khatam", "diamond", "swash"):
            from . import islamic_art as ia

            w_cm = text_w * width / 100
            h_cm = 0.18 if style != "swash" else 0.32
            art = ia.note_rule(style, "#" + color, aspect=w_cm / h_cm, rtl=rtl)
            from .vectors import _svg_size

            vw, vh = _svg_size(art)
            r = etree.SubElement(p, f"{W}r")
            r.append(
                self._inline_drawing(
                    f"word/{kind}s.xml",
                    art,
                    int(w_cm * 360000),
                    int(w_cm * vh / vw * 360000),
                    "Note separator",
                    "",
                )
            )
        elif style == "text":
            props["alignment"] = "center"
            p.append(make_run(text or "⁂", {"color": color}))
        apply_para_props(ppr, props)
        self._register_special(kind, -1)
        self._mark(f"word/{kind}s.xml")
        return {"kind": kind, "style": style, "width_pct": width}

    def set_note_numbering(
        self,
        kind: str = "footnote",
        *,
        number_format: str | None = None,
        restart: str | None = None,
        start_at: int | None = None,
        position: str | None = None,
    ) -> dict:
        """Note numbering for the whole document (settings and every section).

        number_format: decimal, arabic-abjad (أ ب ج), arabic-alpha (أ ب ت),
            roman, upper-roman, letters, upper-letters, symbols (* † ‡),
            decimal-parentheses, decimal-circle.
        restart: continuous, page (restart every page — common in Arabic
            books), section.
        start_at: First number.
        position: footnotes: page-bottom | below-text;
            endnotes: section-end | document-end.
        """
        _, tag, pr_tag = _kind_parts(kind)
        if number_format and number_format not in NOTE_FORMATS:
            raise ValueError(f"number_format must be one of {sorted(NOTE_FORMATS)}")
        if restart and restart not in RESTARTS:
            raise ValueError(f"restart must be one of {sorted(RESTARTS)}")
        if position and position not in POSITIONS[kind]:
            raise ValueError(f"position must be one of {sorted(POSITIONS[kind])}")
        if restart == "page" and kind == "endnote":
            raise ValueError("Endnotes cannot restart every page")
        targets = [self._require("word/settings.xml")]
        doc = self._require("word/document.xml")
        targets += list(doc.iter(f"{W}sectPr"))
        for holder in targets:
            pr = holder.find(f"{W}{pr_tag}")
            if pr is None:
                pr = etree.SubElement(holder, f"{W}{pr_tag}")
            for key, val in (
                ("pos", POSITIONS[kind].get(position or "")),
                ("numFmt", NOTE_FORMATS.get(number_format or "")),
                ("numStart", None if start_at is None else str(int(start_at))),
                ("numRestart", RESTARTS.get(restart or "")),
            ):
                if val is None:
                    continue
                el = pr.find(f"{W}{key}")
                if el is None:
                    el = etree.SubElement(pr, f"{W}{key}")
                el.set(f"{W}val", val)
        self._mark("word/settings.xml")
        self._mark("word/document.xml")
        _ = tag
        return {
            "kind": kind,
            "number_format": number_format,
            "restart": restart,
            "start_at": start_at,
            "position": position,
        }

    def add_custom_mark_note(
        self,
        para_id: str,
        text: str,
        mark: str,
        *,
        kind: str = "footnote",
        where: str = "end",
        after_text: str | None = None,
        mark_in_body: dict[str, Any] | None = None,
        mark_in_note: dict[str, Any] | None = None,
        note_text_props: dict[str, Any] | None = None,
    ) -> dict:
        """Add a note whose mark is your own text instead of an automatic number.

        The mark (e.g. "[١]" — a narration number) appears in the body and
        opens the note, and clicking it in Word jumps to the note. Word's
        automatic numbering skips it.

        Args:
            para_id: Body paragraph.
            text: Note text.
            mark: The mark text, e.g. "[١]", "*", "(أ)".
            kind: "footnote" or "endnote".
            where: "start" or "end" of the paragraph, or "after" with after_text.
            after_text: Put the mark right after this text in the paragraph.
            mark_in_body: Run formatting of the mark in the body (e.g.
                {"bold": True, "color": "520912"}). Default: not superscript.
            mark_in_note: Run formatting of the mark at the start of the note
                (default bold).
            note_text_props: Run formatting of the note text.
        """
        path, tag, _ = _kind_parts(kind)
        root = self._require(path)
        doc = self._require("word/document.xml")
        mark_in_body = self._resolve_style_props(mark_in_body)
        mark_in_note = self._resolve_style_props(mark_in_note)
        note_text_props = self._resolve_style_props(note_text_props)
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")
        used = {int(n.get(f"{W}id", "0")) for n in root.findall(f"{W}{tag}")}
        nid = max(used | {0}) + 1
        note = etree.SubElement(root, f"{W}{tag}")
        note.set(f"{W}id", str(nid))
        np_ = etree.SubElement(note, f"{W}p")
        np_.set(f"{W14}paraId", self._new_para_id())
        np_.set(f"{W14}textId", "77777777")
        ppr = etree.SubElement(np_, f"{W}pPr")
        etree.SubElement(ppr, f"{W}pStyle").set(
            f"{W}val", "FootnoteText" if kind == "footnote" else "EndnoteText"
        )
        np_.append(make_run(mark, {"bold": True, **(mark_in_note or {})}))
        np_.append(make_run(" " + text, note_text_props or None))
        self._mark(path)

        ref = etree.Element(f"{W}r")
        rpr = etree.SubElement(ref, f"{W}rPr")
        if mark_in_body:
            apply_run_props(rpr, mark_in_body)
        ref_el = etree.SubElement(ref, f"{W}{tag}Reference")
        ref_el.set(f"{W}customMarkFollows", "1")
        ref_el.set(f"{W}id", str(nid))
        mark_run = make_run(mark, mark_in_body or None)
        new_runs = [ref, mark_run]

        if where == "start":
            ppr_b = para.find(f"{W}pPr")
            idx = 0 if ppr_b is None else list(para).index(ppr_b) + 1
            for i, r in enumerate(new_runs):
                para.insert(idx + i, r)
        elif where == "after":
            if not after_text:
                raise ValueError("where='after' needs after_text")
            if not _insert_after_text(para, after_text, new_runs):
                raise ValueError(f"Text {after_text!r} not found in a single run")
        else:
            para.extend(new_runs)
        self._mark("word/document.xml")
        return {"kind": kind, "id": nid, "mark": mark, "para_id": para_id}

    def convert_notes(self, from_kind: str = "footnote", ids: list[int] | None = None) -> dict:
        """Turn footnotes into endnotes (from_kind="footnote") or back.

        Text, formatting and references move; marks renumber automatically.
        """
        to_kind = "endnote" if from_kind == "footnote" else "footnote"
        src_path, src_tag, _ = _kind_parts(from_kind)
        dst_path, dst_tag, _ = _kind_parts(to_kind)
        src = self._require(src_path)
        dst = self._require(dst_path)
        doc = self._require("word/document.xml")
        used = {int(n.get(f"{W}id", "0")) for n in dst.findall(f"{W}{dst_tag}")}
        next_id = max(used | {0}) + 1
        style_map = {
            "FootnoteText": "EndnoteText",
            "FootnoteReference": "EndnoteReference",
        }
        if from_kind == "endnote":
            style_map = {v: k for k, v in style_map.items()}
        moved = {}
        for note in list(src.findall(f"{W}{src_tag}")):
            if note.get(f"{W}type") not in (None, "normal"):
                continue
            old = int(note.get(f"{W}id"))
            if ids and old not in ids:
                continue
            src.remove(note)
            note.tag = f"{W}{dst_tag}"
            note.set(f"{W}id", str(next_id))
            for el in note.iter(f"{W}{src_tag}Ref"):
                el.tag = f"{W}{dst_tag}Ref"
            for st in note.iter(f"{W}pStyle", f"{W}rStyle"):
                v = st.get(f"{W}val")
                if v in style_map:
                    st.set(f"{W}val", style_map[v])
            dst.append(note)
            moved[old] = next_id
            next_id += 1
        for ref in doc.iter(f"{W}{src_tag}Reference"):
            old = int(ref.get(f"{W}id", "0"))
            if old in moved:
                ref.tag = f"{W}{dst_tag}Reference"
                ref.set(f"{W}id", str(moved[old]))
                rpr = ref.getparent().find(f"{W}rPr")
                if rpr is not None:
                    rs = rpr.find(f"{W}rStyle")
                    if rs is not None and rs.get(f"{W}val") in style_map:
                        rs.set(f"{W}val", style_map[rs.get(f"{W}val")])
        self._ensure_note_styles(to_kind)
        for p in (src_path, dst_path, "word/document.xml"):
            self._mark(p)
        return {
            "from": from_kind,
            "to": to_kind,
            "moved": len(moved),
            "id_map": {str(k): v for k, v in moved.items()},
        }

    def _ensure_note_styles(self, kind: str) -> None:
        """Endnote/footnote text and reference styles exist after a conversion."""
        styles = self._require("word/styles.xml")
        cap = kind.capitalize()
        have = {s.get(f"{W}styleId") for s in styles.findall(f"{W}style")}
        if f"{cap}Text" not in have:
            st = etree.SubElement(
                styles, f"{W}style", {f"{W}type": "paragraph", f"{W}styleId": f"{cap}Text"}
            )
            etree.SubElement(st, f"{W}name").set(f"{W}val", f"{kind} text")
            etree.SubElement(st, f"{W}basedOn").set(f"{W}val", "Normal")
            rpr = etree.SubElement(st, f"{W}rPr")
            apply_run_props(rpr, {"size_pt": 10})
        if f"{cap}Reference" not in have:
            st = etree.SubElement(
                styles, f"{W}style", {f"{W}type": "character", f"{W}styleId": f"{cap}Reference"}
            )
            etree.SubElement(st, f"{W}name").set(f"{W}val", f"{kind} reference")
            rpr = etree.SubElement(st, f"{W}rPr")
            apply_run_props(rpr, {"superscript": True})
        self._mark("word/styles.xml")


def _insert_after_text(para: etree._Element, needle: str, new_runs: list) -> bool:
    import copy

    for r in para.iter(f"{W}r"):
        t = r.find(f"{W}t")
        if t is None or not t.text or needle not in t.text:
            continue
        before, after = t.text.split(needle, 1)
        t.text = before + needle
        t.set(XML_SPACE, "preserve")
        parent = r.getparent()
        idx = list(parent).index(r)
        for i, nr in enumerate(new_runs):
            parent.insert(idx + 1 + i, nr)
        if after:
            tail = etree.Element(f"{W}r")
            rpr = r.find(f"{W}rPr")
            if rpr is not None:
                tail.append(copy.deepcopy(rpr))
            tt = etree.SubElement(tail, f"{W}t")
            tt.text = after
            tt.set(XML_SPACE, "preserve")
            parent.insert(idx + 1 + len(new_runs), tail)
        return True
    return False
