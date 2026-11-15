"""Indexes and tables of contents that Word can rebuild.

- Index entries are real XE fields, so Word's References ▸ Update Index
  rebuilds the فهرس with correct page numbers after any edit.
- Several separate indexes in one book (أعلام، أماكن، أطراف الأحاديث …)
  through index types (XE \\f / INDEX \\f).
- INDEX and TOC are written as proper multi-paragraph fields: the cached
  entries sit inside the field result, so an update replaces them instead of
  leaving a stale copy behind.
- TOC from any styles (TOC \\t), not only Heading 1–9.
"""

from __future__ import annotations

import copy
import re
import unicodedata
from typing import Any

from lxml import etree

from .bidi import document_rtl_default, has_rtl
from .ooxml import W14, XML_SPACE, W, make_run

_TASHKEEL = re.compile("[\u0610-\u061a\u064b-\u065f\u0670\u06d6-\u06ed\u0640]")
_ALEF = str.maketrans({"أ": "ا", "إ": "ا", "آ": "ا", "ٱ": "ا", "ى": "ي", "ة": "ه"})
LANG_IDS = {"ar": 1025, "en": 1033, "ur": 1056, "fa": 1065}
LEADERS = {"dot": "dot", "hyphen": "hyphen", "underscore": "underscore", "none": None}


def _sort_key(text: str) -> str:
    """Arabic-aware sort: ignore harakat/tatweel and hamza forms on alef."""
    t = _TASHKEEL.sub("", unicodedata.normalize("NFC", text)).translate(_ALEF)
    return t.casefold()


def _heading_letter(text: str) -> str:
    key = _sort_key(text).lstrip("«(\"'[")
    return key[:1].upper() if key else ""


def _quote(text: str) -> str:
    """Escape text for a field argument: quotes and colons are special in XE."""
    return text.replace('"', '\\"').replace(":", "\\:")


def _xe_instr(
    entry: str,
    subentry: str | None,
    see: str | None,
    index_type: str | None,
    bold_page: bool,
    italic_page: bool,
) -> str:
    body = _quote(entry)
    if subentry:
        body += ":" + ":".join(_quote(s) for s in subentry.split(":"))
    instr = f'XE "{body}"'
    if see:
        instr += f' \\t "{see}"'
    if index_type:
        instr += f' \\f "{index_type}"'
    if bold_page:
        instr += " \\b"
    if italic_page:
        instr += " \\i"
    return instr


def _xe_runs(instr: str, rtl: bool) -> list[etree._Element]:
    """An XE field: begin, instruction, end (XE has no result)."""
    props = {"rtl": True} if rtl else None
    begin = make_run("", props)
    etree.SubElement(begin, f"{W}fldChar").set(f"{W}fldCharType", "begin")
    mid = make_run("", props)
    it = etree.SubElement(mid, f"{W}instrText")
    it.text = f" {instr} "
    it.set(XML_SPACE, "preserve")
    end = make_run("", props)
    etree.SubElement(end, f"{W}fldChar").set(f"{W}fldCharType", "end")
    return [begin, mid, end]


def _parse_xe(instr: str) -> dict[str, Any] | None:
    m = re.match(r'\s*XE\s+"((?:[^"\\]|\\.)*)"(.*)$', instr, re.S)
    if not m:
        return None
    raw, rest = m.group(1), m.group(2)
    parts = [p.replace("\\:", ":").replace('\\"', '"') for p in re.split(r"(?<!\\):", raw)]
    see = re.search(r'\\t\s+"([^"]*)"', rest)
    typ = re.search(r'\\f\s+"([^"]*)"', rest)
    return {
        "entry": parts[0],
        "subentries": parts[1:],
        "see": see.group(1) if see else None,
        "index_type": typ.group(1) if typ else None,
        "bold_page": "\\b" in rest,
        "italic_page": "\\i" in rest,
    }


def _check_type(index_type: str | None) -> None:
    if index_type is not None and not re.fullmatch(r"[A-Za-z]", index_type):
        raise ValueError("index_type must be a single letter, e.g. 'P' for people")


class IndexingMixin:
    """XE marking, INDEX fields and style-based tables of contents."""

    # ── marking ─────────────────────────────────────────────────────────

    def mark_index_entries(
        self,
        terms: list[Any],
        *,
        index_type: str | None = None,
        whole_word: bool = True,
        first_per_paragraph: bool = True,
        parts: str = "body",
    ) -> dict:
        """Mark every occurrence of terms as index entries (XE fields).

        Args:
            terms: Each a string, or a dict with keys text (what to find),
                regex (bool), entry (index wording; default the text found),
                subentry ("main:sub" levels), see (cross-reference "انظر …"),
                bold_page / italic_page (page number style for main
                discussions), index_type (overrides the default).
            index_type: One letter putting the marks in a separate index,
                e.g. "P" أعلام, "L" أماكن, "H" أطراف الأحاديث.
            whole_word: Match whole words only (Arabic prefixes such as و ف ب
                ل ك ال still match: «والبخاري» counts for «البخاري»).
            first_per_paragraph: One mark per paragraph per term (Word lists
                a page once anyway).
            parts: "body" or "all" (body, footnotes and endnotes).
        """
        _check_type(index_type)
        prefixes = (
            ("word/document.xml",)
            if parts == "body"
            else ("word/document.xml", "word/footnotes.xml", "word/endnotes.xml")
        )
        report: dict[str, int] = {}
        for term in terms:
            spec = {"text": term} if isinstance(term, str) else dict(term)
            text = spec.get("text") or spec.get("entry")
            if not text:
                raise ValueError("Each term needs text")
            typ = spec.get("index_type", index_type)
            _check_type(typ)
            src = text if spec.get("regex") else re.escape(text)
            if whole_word:
                # Allow attached Arabic proclitics before the word.
                src = rf"(?<![\w])(?:[وفبلك]{{0,2}})(?:{src})(?![\w])"
            rx = re.compile(src)
            count = 0
            for part in prefixes:
                root = self._tree(part)
                if root is None:
                    continue
                n = 0
                for p in list(root.iter(f"{W}p")):
                    n += self._mark_in_paragraph(p, rx, spec, typ, first_per_paragraph)
                if n:
                    self._mark(part)
                    count += n
            report[spec.get("entry") or text] = count
        return {"marked": sum(report.values()), "per_term": report}

    def _mark_in_paragraph(
        self,
        p: etree._Element,
        rx: re.Pattern,
        spec: dict,
        index_type: str | None,
        first_only: bool,
    ) -> int:
        done = 0
        for r in list(p.iter(f"{W}r")):
            if first_only and done:
                break
            kids = [c for c in r if c.tag != f"{W}rPr"]
            if len(kids) != 1 or kids[0].tag != f"{W}t" or not kids[0].text:
                continue
            anc = r.getparent()
            if anc is None or anc.tag in (f"{W}del", f"{W}moveFrom"):
                continue
            text = kids[0].text
            matches = [m for m in rx.finditer(text) if m.end() > m.start()]
            if first_only:
                matches = matches[:1]
            if not matches:
                continue
            idx = list(anc).index(r)
            rpr = r.find(f"{W}rPr")
            pieces: list[etree._Element] = []
            pos = 0
            for m in matches:
                found = m.group(0)
                entry = spec.get("entry") or _strip_proclitic(found, spec.get("text", ""))
                instr = _xe_instr(
                    entry,
                    spec.get("subentry"),
                    spec.get("see"),
                    index_type,
                    bool(spec.get("bold_page")),
                    bool(spec.get("italic_page")),
                )
                pieces.append(_text_run(text[pos : m.end()], rpr))
                pieces.extend(_xe_runs(instr, has_rtl(entry)))
                pos = m.end()
            if pos < len(text):
                pieces.append(_text_run(text[pos:], rpr))
            for i, piece in enumerate(pieces):
                anc.insert(idx + i, piece)
            anc.remove(r)
            done += len(matches)
        return done

    def mark_index_entry(
        self,
        para_id: str,
        entry: str,
        *,
        subentry: str | None = None,
        see: str | None = None,
        after_text: str | None = None,
        index_type: str | None = None,
        bold_page: bool = False,
        italic_page: bool = False,
    ) -> dict:
        """Mark one index entry in a paragraph (at its end, or after after_text)."""
        _check_type(index_type)
        doc = self._require("word/document.xml")
        para = self._find_para(doc, para_id)
        if para is None:
            raise ValueError(f"Paragraph '{para_id}' not found")
        instr = _xe_instr(entry, subentry, see, index_type, bold_page, italic_page)
        runs = _xe_runs(instr, has_rtl(entry))
        if after_text:
            from .notes import _insert_after_text

            if not _insert_after_text(para, after_text, runs):
                raise ValueError(f"Text {after_text!r} not found in a single run")
        else:
            para.extend(runs)
        self._mark("word/document.xml")
        return {"para_id": para_id, "field": instr}

    def list_index_entries(self, index_type: str | None = None) -> dict:
        """All XE marks, grouped by entry, with how often each is marked."""
        entries: dict[tuple, dict] = {}
        for part in ("word/document.xml", "word/footnotes.xml", "word/endnotes.xml"):
            root = self._tree(part)
            if root is None:
                continue
            for it in root.iter(f"{W}instrText"):
                info = _parse_xe(it.text or "")
                if info is None:
                    continue
                if index_type is not None and info["index_type"] != index_type:
                    continue
                key = (info["entry"], tuple(info["subentries"]), info["index_type"])
                e = entries.setdefault(key, {**info, "count": 0, "parts": set()})
                e["count"] += 1
                e["parts"].add(part.split("/")[-1].split(".")[0])
        out = sorted(entries.values(), key=lambda e: (_sort_key(e["entry"]), e["subentries"]))
        for e in out:
            e["parts"] = sorted(e["parts"])
        return {"count": len(out), "entries": out}

    def remove_index_entries(self, entry: str | None = None, index_type: str | None = None) -> dict:
        """Delete XE marks (all, one entry, or one index type)."""
        removed = 0
        for part in ("word/document.xml", "word/footnotes.xml", "word/endnotes.xml"):
            root = self._tree(part)
            if root is None:
                continue
            for it in list(root.iter(f"{W}instrText")):
                info = _parse_xe(it.text or "")
                if info is None:
                    continue
                if entry is not None and info["entry"] != entry:
                    continue
                if index_type is not None and info["index_type"] != index_type:
                    continue
                mid = it.getparent()
                parent = mid.getparent()
                i = list(parent).index(mid)
                kids = list(parent)
                for r in (
                    kids[i - 1] if i > 0 else None,
                    mid,
                    kids[i + 1] if i + 1 < len(kids) else None,
                ):
                    if r is not None and r.getparent() is parent:
                        parent.remove(r)
                removed += 1
                self._mark(part)
        return {"removed": removed}

    # ── INDEX field ─────────────────────────────────────────────────────

    def insert_index(
        self,
        *,
        after_para_id: str | None = None,
        title: str | None = None,
        index_type: str | None = None,
        columns: int = 2,
        letter_headings: bool = True,
        right_align_page_numbers: bool = True,
        leader: str = "dot",
        run_in: bool = False,
        language: str = "ar",
        font: str | None = None,
        cs_font: str | None = None,
        size_pt: float | None = None,
        heading_color: str | None = None,
        new_page: bool = True,
        update_on_open: bool = True,
    ) -> dict:
        """Insert an index (فهرس) built from the XE marks.

        Args:
            after_para_id: Put it after this paragraph (default: end of the
                document).
            title: Heading above the index, e.g. "فهرس الأعلام".
            index_type: Build only from marks of this type (one letter).
            columns: 1–4 columns.
            letter_headings: A heading per letter (أ، ب، ت …).
            right_align_page_numbers: Page numbers at the line end with a
                leader.
            leader: dot, hyphen, underscore or none.
            run_in: Sub-entries on the same line as the main entry.
            language: Sorting language: ar, en, ur, fa.
            font / cs_font / size_pt: Index text formatting (styles "index 1"…).
            heading_color: Colour of the letter headings.
            new_page: Start the index (its title) on a new page.
            update_on_open: Word fills in page numbers when the file opens.
        """
        _check_type(index_type)
        if not 1 <= columns <= 4:
            raise ValueError("columns must be 1–4")
        if leader not in LEADERS:
            raise ValueError(f"leader must be one of {sorted(LEADERS)}")
        if language not in LANG_IDS:
            raise ValueError(f"language must be one of {sorted(LANG_IDS)}")
        rtl = document_rtl_default(self._tree("word/styles.xml")) or language in ("ar", "ur", "fa")
        width_cm = self._text_width_cm() / columns - (0.6 if columns > 1 else 0)
        ids = self._ensure_index_styles(
            rtl, width_cm, leader, right_align_page_numbers, font, cs_font, size_pt, heading_color
        )

        instr = "INDEX"
        if right_align_page_numbers:
            instr += ' \\e "\t"'
        if letter_headings:
            instr += ' \\h "A"'
        if columns > 1:
            instr += f' \\c "{columns}"'
        if run_in:
            instr += " \\r"
        if index_type:
            instr += f' \\f "{index_type}"'
        instr += f' \\z "{LANG_IDS[language]}"'

        listed = self.list_index_entries(index_type)["entries"]
        if index_type is None:
            listed = [e for e in listed if e["index_type"] is None]
        lines: list[tuple[str, str, str]] = []  # (style id, text, page placeholder)
        last_letter = None
        for e in listed:
            if letter_headings:
                letter = _heading_letter(e["entry"])
                if letter != last_letter:
                    lines.append((ids["heading"], letter, ""))
                    last_letter = letter
            if e["see"]:
                lines.append(
                    (
                        ids[1],
                        f"{e['entry']}، انظر {e['see']}"
                        if rtl
                        else f"{e['entry']}. See {e['see']}",
                        "",
                    )
                )
                continue
            if e["subentries"]:
                if not any(t == e["entry"] and s == ids[1] for s, t, _ in lines):
                    lines.append((ids[1], e["entry"], ""))
                for depth, sub in enumerate(e["subentries"][:2], start=2):
                    lines.append((ids[depth], sub, "\t" if right_align_page_numbers else ", "))
            else:
                lines.append((ids[1], e["entry"], "\t" if right_align_page_numbers else ", "))
        if not lines:
            lines = [
                (
                    ids[1],
                    "حدِّث الحقول لبناء الفهرس" if rtl else "Update fields to build the index",
                    "",
                )
            ]
        placed = self._insert_field_block(
            instr, lines, after_para_id, title, rtl, title_style="IndexTitle", new_page=new_page
        )
        if update_on_open:
            self.set_update_fields_on_open(True)
        return {"field": instr, "entries": len(listed), "para_id": placed}

    # ── TOC from styles ─────────────────────────────────────────────────

    def insert_toc(
        self,
        *,
        after_para_id: str | None = None,
        title: str | None = None,
        levels: str | None = "1-3",
        styles: dict[str, int] | None = None,
        hyperlinks: bool = True,
        leader: str = "dot",
        page_numbers: bool = True,
        font: str | None = None,
        cs_font: str | None = None,
        size_pt: float | None = None,
        level1_bold: bool = True,
        level1_color: str | None = None,
        new_page: bool = True,
        update_on_open: bool = True,
    ) -> dict:
        """Table of contents (فهرس الموضوعات) from headings and/or your own styles.

        Args:
            after_para_id: Put it after this paragraph (default: end, as in
                Arabic books; pass the title page's last paragraph to put it
                at the front).
            title: Heading above it, e.g. "فهرس الموضوعات".
            levels: Heading/outline levels, e.g. "1-3"; None for styles only.
            styles: Your styles and their TOC level, e.g. {"Chapter": 1,
                "Bab": 2, "Narration Title": 3}.
            hyperlinks: Entries are clickable in Word and PDF.
            leader / page_numbers: Dotted leaders, show page numbers.
            font / cs_font / size_pt / level1_bold / level1_color: Look of the
                entries (styles "toc 1" … "toc 9").
            new_page: Start it (its title) on a new page.
            update_on_open: Word fills in page numbers when the file opens.
        """
        if leader not in LEADERS:
            raise ValueError(f"leader must be one of {sorted(LEADERS)}")
        if not levels and not styles:
            raise ValueError("Give levels, styles or both")
        styles_root = self._require("word/styles.xml")
        rtl = document_rtl_default(styles_root)
        style_level: dict[str, int] = {}
        names: list[str] = []
        for name, lvl in (styles or {}).items():
            if not 1 <= int(lvl) <= 9:
                raise ValueError("TOC levels are 1–9")
            st = _style_by_name_or_id(styles_root, name)
            if st is None:
                raise ValueError(f"Style {name!r} not found")
            style_level[st.get(f"{W}styleId")] = int(lvl)
            names.append(f"{st.find(f'{W}name').get(f'{W}val')},{int(lvl)}")
        lo, hi = 1, 0
        if levels:
            m = re.fullmatch(r"(\d)-(\d)", levels)
            if not m:
                raise ValueError('levels looks like "1-3"')
            lo, hi = int(m.group(1)), int(m.group(2))
            for st in styles_root.findall(f"{W}style"):
                nm = (
                    st.find(f"{W}name").get(f"{W}val") if st.find(f"{W}name") is not None else ""
                ).lower()
                hm = re.fullmatch(r"heading (\d)", nm)
                if hm and lo <= int(hm.group(1)) <= hi:
                    style_level.setdefault(st.get(f"{W}styleId"), int(hm.group(1)))

        instr = "TOC"
        if levels:
            instr += f' \\o "{levels}"'
        if names:
            instr += f' \\t "{",".join(names)}"'
        if hyperlinks:
            instr += " \\h"
        instr += " \\z"
        if not names:
            # \u together with \t makes Word find no entries (checked in Word).
            instr += " \\u"
        if not page_numbers:
            instr += " \\n"

        max_level = max(list(style_level.values()) + [hi, 1])
        ids = self._ensure_toc_styles(
            rtl, max_level, leader, page_numbers, font, cs_font, size_pt, level1_bold, level1_color
        )
        doc = self._require("word/document.xml")
        lines: list[tuple[str, str, str]] = []
        for p in doc.iter(f"{W}p"):
            ps = p.find(f"{W}pPr/{W}pStyle")
            lvl = style_level.get(ps.get(f"{W}val")) if ps is not None else None
            if lvl is None and levels:
                ol = p.find(f"{W}pPr/{W}outlineLvl")
                if ol is not None and lo <= int(ol.get(f"{W}val")) + 1 <= hi:
                    lvl = int(ol.get(f"{W}val")) + 1
            if lvl is None:
                continue
            text = "".join(t.text or "" for t in p.iter(f"{W}t")).strip()
            if text:
                lines.append((ids[lvl], text, "\t" if page_numbers else ""))
        if not lines:
            lines = [
                (
                    ids[1],
                    "حدِّث الحقول لبناء الفهرس"
                    if rtl
                    else "Update fields to build the table of contents",
                    "",
                )
            ]
        placed = self._insert_field_block(
            instr, lines, after_para_id, title, rtl, title_style="TOCHeading", new_page=new_page
        )
        if update_on_open:
            self.set_update_fields_on_open(True)
        return {"field": instr, "entries": len(lines), "para_id": placed}

    # ── shared building blocks ──────────────────────────────────────────

    def _text_width_cm(self) -> float:
        doc = self._require("word/document.xml")
        sect = doc.find(f"{W}body/{W}sectPr")
        try:
            w = int(sect.find(f"{W}pgSz").get(f"{W}w"))
            mar = sect.find(f"{W}pgMar")
            w -= int(mar.get(f"{W}left", 0)) + int(mar.get(f"{W}right", 0))
            w -= int(mar.get(f"{W}gutter", 0) or 0)
            return w / 567.0
        except (AttributeError, TypeError, ValueError):
            return 15.0

    def _ensure_index_styles(
        self, rtl, width_cm, leader, right_align, font, cs_font, size_pt, heading_color
    ) -> dict:
        run: dict[str, Any] = {}
        if font:
            run["font"] = font
        if cs_font:
            run["cs_font"] = cs_font
        if size_pt:
            run["size_pt"] = size_pt
        tabs = (
            [{"align": "end", "pos_cm": round(width_cm, 2), "leader": LEADERS[leader]}]
            if right_align
            else []
        )
        if tabs and tabs[0]["leader"] is None:
            del tabs[0]["leader"]
        direction = "rtl" if rtl else "ltr"
        ids: dict[Any, str] = {}
        for lvl in (1, 2, 3):
            para = {
                "direction": direction,
                "space_after_pt": 0,
                "indent_start_cm": round(0.5 * (lvl - 1), 2),
                "hanging_cm": 0.5,
            }
            if tabs:
                para["tabs"] = tabs
            r = self.define_style(
                f"index {lvl}",
                "paragraph",
                based_on="Normal",
                next_style="Normal",
                ui_priority=99,
                run=run or None,
                paragraph=para,
            )
            ids[lvl] = r["style_id"]
        head_run = dict(run)
        head_run["bold"] = True
        if size_pt:
            head_run["size_pt"] = size_pt + 2
        if heading_color:
            head_run["color"] = heading_color
        r = self.define_style(
            "index heading",
            "paragraph",
            based_on="Normal",
            next_style="index 1",
            ui_priority=99,
            run=head_run,
            paragraph={
                "direction": direction,
                "space_before_pt": 6,
                "space_after_pt": 2,
                "keep_with_next": True,
            },
        )
        ids["heading"] = r["style_id"]
        return ids

    def _ensure_toc_styles(
        self,
        rtl,
        max_level,
        leader,
        page_numbers,
        font,
        cs_font,
        size_pt,
        level1_bold,
        level1_color,
    ) -> dict:
        width = round(self._text_width_cm(), 2)
        ids: dict[int, str] = {}
        for lvl in range(1, max_level + 1):
            run: dict[str, Any] = {}
            if font:
                run["font"] = font
            if cs_font:
                run["cs_font"] = cs_font
            if size_pt:
                run["size_pt"] = size_pt
            if lvl == 1 and level1_bold:
                run["bold"] = True
            if lvl == 1 and level1_color:
                run["color"] = level1_color
            para: dict[str, Any] = {
                "direction": "rtl" if rtl else "ltr",
                "indent_start_cm": round(0.6 * (lvl - 1), 2),
                "space_after_pt": 2 if lvl > 1 else 4,
            }
            if page_numbers:
                tab = {"align": "end", "pos_cm": width}
                if LEADERS[leader]:
                    tab["leader"] = LEADERS[leader]
                para["tabs"] = [tab]
            r = self.define_style(
                f"toc {lvl}",
                "paragraph",
                based_on="Normal",
                next_style="Normal",
                ui_priority=39,
                run=run or None,
                paragraph=para,
            )
            ids[lvl] = r["style_id"]
        self.define_style(
            "TOC Heading",
            "paragraph",
            based_on="Normal",
            next_style="Normal",
            ui_priority=39,
            run={"bold": True, "size_pt": (size_pt or 12) + 4},
            paragraph={
                "direction": "rtl" if rtl else "ltr",
                "alignment": "center",
                "space_after_pt": 12,
                "keep_with_next": True,
            },
        )
        return ids

    def _insert_field_block(
        self,
        instr: str,
        lines: list[tuple[str, str, str]],
        after_para_id: str | None,
        title: str | None,
        rtl: bool,
        title_style: str,
        new_page: bool = False,
    ) -> str:
        """Title + one field whose cached result spans several paragraphs."""
        doc = self._require("word/document.xml")
        body = doc.find(f"{W}body")
        if after_para_id:
            anchor = self._find_para(doc, after_para_id)
            if anchor is None:
                raise ValueError(f"Paragraph '{after_para_id}' not found")
            container = anchor.getparent()
            pos = list(container).index(anchor) + 1
        else:
            container = body
            kids = list(body)
            pos = len(kids) - 1 if kids and kids[-1].tag == f"{W}sectPr" else len(kids)

        props = {"rtl": True} if rtl else None
        paras: list[etree._Element] = []
        if title:
            styles_root = self._require("word/styles.xml")
            st = _style_by_name_or_id(styles_root, title_style)
            if st is None:
                st = _style_by_name_or_id(styles_root, "TOC Heading")
            if st is None:
                st_id = self.define_style(
                    "TOC Heading",
                    "paragraph",
                    based_on="Normal",
                    run={"bold": True},
                    paragraph={"alignment": "center"},
                )["style_id"]
            else:
                st_id = st.get(f"{W}styleId")
            tp = self._new_block_para(st_id, rtl)
            tp.append(make_run(title, props))
            paras.append(tp)
        if new_page:
            first = paras[0] if paras else None
            if first is None:
                first = self._new_block_para(None, rtl)
                paras.append(first)
            ppr = first.find(f"{W}pPr")
            etree.SubElement(ppr, f"{W}pageBreakBefore")
        first_id = None
        for i, (style_id, text, tail) in enumerate(lines):
            p = self._new_block_para(style_id, rtl)
            if first_id is None:
                first_id = p.get(f"{W14}paraId")
            if i == 0:
                begin = make_run("", props)
                fc = etree.SubElement(begin, f"{W}fldChar")
                fc.set(f"{W}fldCharType", "begin")
                fc.set(f"{W}dirty", "true")
                mid = make_run("", props)
                it = etree.SubElement(mid, f"{W}instrText")
                it.text = f" {instr} "
                it.set(XML_SPACE, "preserve")
                sep = make_run("", props)
                etree.SubElement(sep, f"{W}fldChar").set(f"{W}fldCharType", "separate")
                p.extend([begin, mid, sep])
            p.append(make_run(text, props if has_rtl(text) else None))
            if tail:
                p.append(make_run(tail, None))
            if i == len(lines) - 1:
                end = make_run("", props)
                etree.SubElement(end, f"{W}fldChar").set(f"{W}fldCharType", "end")
                p.append(end)
            paras.append(p)
        for offset, p in enumerate(paras):
            container.insert(pos + offset, p)
        self._mark("word/document.xml")
        return first_id or ""

    def _new_block_para(self, style_id: str | None, rtl: bool) -> etree._Element:
        p = etree.Element(f"{W}p")
        p.set(f"{W14}paraId", self._new_para_id())
        p.set(f"{W14}textId", "77777777")
        ppr = etree.SubElement(p, f"{W}pPr")
        if style_id:
            etree.SubElement(ppr, f"{W}pStyle").set(f"{W}val", style_id)
        if rtl:
            etree.SubElement(ppr, f"{W}bidi")
        return p


def _text_run(text: str, rpr: etree._Element | None) -> etree._Element:
    r = etree.Element(f"{W}r")
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = etree.SubElement(r, f"{W}t")
    t.text = text
    t.set(XML_SPACE, "preserve")
    return r


def _strip_proclitic(found: str, wanted: str) -> str:
    """«والبخاري» found for «البخاري» → index under «البخاري»."""
    if wanted and found.endswith(wanted):
        return wanted
    return found


def _style_by_name_or_id(styles: etree._Element, key: str) -> etree._Element | None:
    lo = key.lower()
    for st in styles.findall(f"{W}style"):
        if (st.get(f"{W}styleId") or "").lower() == lo:
            return st
        nm = st.find(f"{W}name")
        if nm is not None and (nm.get(f"{W}val") or "").lower() == lo:
            return st
    return None
