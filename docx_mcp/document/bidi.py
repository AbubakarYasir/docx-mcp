"""Right-to-left (Arabic, Urdu, Persian, Hebrew) support.

Word does not infer direction from text. A paragraph is right-to-left only when
its ``w:pPr`` has ``w:bidi``; a run is treated as right-to-left (and uses its
complex-script font, size, bold and italic: ``rFonts/@cs``, ``szCs``, ``bCs``,
``iCs``) only when its ``w:rPr`` has ``w:rtl``. Without these, Arabic text is
left-aligned, punctuation lands on the wrong side, and bold/size settings are
silently ignored for the Arabic letters.

Two layers:

* :func:`normalize_bidi` runs on save. It fixes paragraphs created by
  docx-mcp in this session (tracked by paraId) and runs added to paragraphs
  that are already right-to-left. Paragraphs that came from the original file
  are never re-directed.
* :class:`BidiMixin` exposes explicit tools: paragraph direction, document
  direction (with an Arabic font, size and language), and table direction.

Set ``DOCX_MCP_AUTO_BIDI=0`` to turn the automatic layer off.
"""

from __future__ import annotations

import copy
import os
import unicodedata

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"

_STRONG_RTL = {"R", "AL"}
_STRONG_LTR = {"L"}
_DEFAULT_RTL_LANG = os.environ.get("DOCX_MCP_RTL_LANG", "ar-SA")
_XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def auto_bidi_enabled() -> bool:
    return os.environ.get("DOCX_MCP_AUTO_BIDI", "1").strip().lower() not in ("0", "false", "no")


def char_dir(ch: str) -> str:
    """'r' for strong right-to-left, 'l' for strong left-to-right, 'n' otherwise."""
    cls = unicodedata.bidirectional(ch)
    if cls in _STRONG_RTL:
        return "r"
    if cls in _STRONG_LTR:
        return "l"
    return "n"


def first_strong(text: str) -> str | None:
    """Direction of the first strong character ('rtl' / 'ltr'), as Word's 'auto'."""
    for ch in text:
        d = char_dir(ch)
        if d == "r":
            return "rtl"
        if d == "l":
            return "ltr"
    return None


def has_rtl(text: str) -> bool:
    return any(char_dir(c) == "r" for c in text)


def document_rtl_default(styles: etree._Element | None) -> bool:
    """True when styles.xml makes paragraphs right-to-left by default."""
    if styles is None:
        return False
    bidi = styles.find(f"{W}docDefaults/{W}pPrDefault/{W}pPr/{W}bidi")
    return bidi is not None and bidi.get(f"{W}val", "1") not in ("0", "false", "off")


def mirror_complex_script(root: etree._Element) -> int:
    """In style definitions, give b/i/sz their complex-script twins.

    Word applies ``w:b`` only to Latin letters and ``w:bCs`` only to
    Arabic/Hebrew letters, so a bold heading style without ``bCs`` leaves
    Arabic headings plain. Same for i/iCs and sz/szCs.
    """
    n = 0
    for rpr in root.iter(f"{W}rPr"):
        n += _mirror_toggle(rpr, "b", "bCs") + _mirror_toggle(rpr, "i", "iCs")
        sz = rpr.find(f"{W}sz")
        if sz is not None and rpr.find(f"{W}szCs") is None:
            etree.SubElement(rpr, f"{W}szCs").set(f"{W}val", sz.get(f"{W}val", "24"))
            n += 1
    return n


def _is_arabic_script(text: str) -> bool:
    return any(
        "؀" <= c <= "ۿ" or "ݐ" <= c <= "ݿ" or "ࢠ" <= c <= "ࣿ" or "ﭐ" <= c <= "﷿" or "ﹰ" <= c <= "﻿"
        for c in text
    )


def segment(text: str, base: str) -> list[tuple[str, str]]:
    """Split text into (direction, chunk) pieces, 'r' or 'l'.

    Neutral characters (spaces, digits, punctuation) join their neighbours
    when both sides have the same direction, otherwise the paragraph's base
    direction (``'r'`` or ``'l'``), which is how the Unicode bidi algorithm
    resolves them.
    """
    dirs = [char_dir(c) for c in text]
    n = len(dirs)
    resolved = dirs[:]
    i = 0
    while i < n:
        if dirs[i] != "n":
            i += 1
            continue
        j = i
        while j < n and dirs[j] == "n":
            j += 1
        before = dirs[i - 1] if i > 0 else None
        after = dirs[j] if j < n else None
        if before and after and before == after:
            fill = before
        elif before is None and after is not None:
            fill = after if after == base else base
        elif after is None and before is not None:
            fill = before if before == base else base
        else:
            fill = base
        for k in range(i, j):
            resolved[k] = fill
        i = j
    out: list[tuple[str, str]] = []
    for ch, d in zip(text, resolved, strict=True):
        if out and out[-1][0] == d:
            out[-1] = (d, out[-1][1] + ch)
        else:
            out.append((d, ch))
    return out


# ── XML helpers ─────────────────────────────────────────────────────────────


def _get_or_add(parent: etree._Element, tag: str) -> etree._Element:
    el = parent.find(tag)
    if el is None:
        el = etree.SubElement(parent, tag)
    return el


def _ppr(p: etree._Element) -> etree._Element:
    ppr = p.find(f"{W}pPr")
    if ppr is None:
        ppr = etree.Element(f"{W}pPr")
        p.insert(0, ppr)
    return ppr


def _rpr(r: etree._Element) -> etree._Element:
    rpr = r.find(f"{W}rPr")
    if rpr is None:
        rpr = etree.Element(f"{W}rPr")
        r.insert(0, rpr)
    return rpr


def _para_text(p: etree._Element) -> str:
    return "".join(t.text or "" for t in p.iter(f"{W}t"))


def _para_is_rtl(p: etree._Element) -> bool:
    bidi = p.find(f"{W}pPr/{W}bidi")
    return bidi is not None and bidi.get(f"{W}val", "1") not in ("0", "false", "off")


def _mirror_toggle(rpr: etree._Element, latin: str, cs: str) -> bool:
    """Copy b -> bCs, i -> iCs when the complex-script twin is missing."""
    src = rpr.find(f"{W}{latin}")
    if src is None or rpr.find(f"{W}{cs}") is not None:
        return False
    twin = etree.SubElement(rpr, f"{W}{cs}")
    val = src.get(f"{W}val")
    if val is not None:
        twin.set(f"{W}val", val)
    return True


def mark_run_rtl(r: etree._Element, text: str) -> int:
    """Give a right-to-left run the properties Word needs. Returns changes made."""
    changes = 0
    rpr = _rpr(r)
    if rpr.find(f"{W}rtl") is None:
        etree.SubElement(rpr, f"{W}rtl")
        changes += 1
    changes += _mirror_toggle(rpr, "b", "bCs") + _mirror_toggle(rpr, "i", "iCs")
    sz = rpr.find(f"{W}sz")
    if sz is not None and rpr.find(f"{W}szCs") is None:
        etree.SubElement(rpr, f"{W}szCs").set(f"{W}val", sz.get(f"{W}val", "24"))
        changes += 1
    if _is_arabic_script(text):
        lang = rpr.find(f"{W}lang")
        if lang is None:
            lang = etree.SubElement(rpr, f"{W}lang")
        if lang.get(f"{W}bidi") is None:
            lang.set(f"{W}bidi", _DEFAULT_RTL_LANG)
            changes += 1
    if len(rpr) == 0:
        r.remove(rpr)
    return changes


def _simple_text_run(r: etree._Element) -> etree._Element | None:
    """Return the run's single w:t if the run is just [rPr] + one w:t."""
    kids = [c for c in r if c.tag != f"{W}rPr"]
    if len(kids) == 1 and kids[0].tag == f"{W}t":
        return kids[0]
    return None


def _split_mixed_run(r: etree._Element, base: str) -> list[tuple[etree._Element, str, str]]:
    """Split a single-w:t run at direction changes. Returns [(run, dir, text)]."""
    t = _simple_text_run(r)
    text = t.text or "" if t is not None else ""
    if t is None or not text:
        return [(r, "n", "".join(x.text or "" for x in r.iter(f"{W}t")))]
    parts = segment(text, base)
    if len(parts) == 1:
        return [(r, parts[0][0], text)]
    parent = r.getparent()
    idx = list(parent).index(r)
    rpr = r.find(f"{W}rPr")
    out = []
    for i, (d, chunk) in enumerate(parts):
        nr = etree.Element(f"{W}r")
        for k, v in r.attrib.items():
            nr.set(k, v)
        if rpr is not None:
            nr.append(copy.deepcopy(rpr))
        nt = etree.SubElement(nr, f"{W}t")
        nt.text = chunk
        nt.set(_XML_SPACE, "preserve")
        parent.insert(idx + i, nr)
        out.append((nr, d, chunk))
    parent.remove(r)
    return out


def _runs_in_para(p: etree._Element) -> list[etree._Element]:
    """Runs that belong to the paragraph's own text (not inside deletions)."""
    runs = []
    for r in p.iter(f"{W}r"):
        anc = r.getparent()
        inside_del = False
        while anc is not None and anc is not p:
            if anc.tag in (f"{W}del", f"{W}moveFrom"):
                inside_del = True
                break
            anc = anc.getparent()
        if not inside_del:
            runs.append(r)
    return runs


def normalize_bidi(
    root: etree._Element, created_para_ids: set[str], rtl_default: bool = False
) -> int:
    """Apply right-to-left markup. Returns the number of changes.

    * Paragraphs created this session: direction from the first strong
      character (Word's own rule); mixed-script runs are split so each piece
      can be marked; right-to-left pieces get ``w:rtl`` and their
      complex-script properties.
    * Other paragraphs that are already right-to-left: right-to-left runs
      that lack ``w:rtl`` (text added by an edit) get it.
    Paragraphs with an explicit direction (``w:bidi`` present, either value)
    keep it. When the document default is right-to-left (``rtl_default``),
    new paragraphs that start with a left-to-right letter are marked
    left-to-right explicitly so they do not flip.
    """
    changes = 0
    for p in list(root.iter(f"{W}p")):
        pid = (p.get(f"{W14}paraId") or "").upper()
        is_new = pid in created_para_ids
        if not is_new and not _para_is_rtl(p):
            continue
        text = _para_text(p)
        if (
            is_new
            and rtl_default
            and p.find(f"{W}pPr/{W}bidi") is None
            and first_strong(text) == "ltr"
        ):
            etree.SubElement(_ppr(p), f"{W}bidi").set(f"{W}val", "0")
            changes += 1
            continue
        if not has_rtl(text):
            continue
        if is_new and p.find(f"{W}pPr/{W}bidi") is None and first_strong(text) == "rtl":
            etree.SubElement(_ppr(p), f"{W}bidi")
            changes += 1
        base = "r" if _para_is_rtl(p) else "l"
        for r in _runs_in_para(p):
            if r.getparent() is None:
                continue
            pieces = _split_mixed_run(r, base) if is_new else [(r, None, None)]
            for run, d, chunk in pieces:
                chunk = (
                    chunk if chunk is not None else "".join(x.text or "" for x in run.iter(f"{W}t"))
                )
                if d is None:
                    d = "r" if has_rtl(chunk) and first_strong(chunk) != "ltr" else "l"
                    if any(char_dir(c) == "l" for c in chunk):
                        continue  # mixed run in an original paragraph: leave as is
                if d == "r" and has_rtl(chunk):
                    changes += mark_run_rtl(run, chunk)
                if len(pieces) > 1:
                    changes += 1
    # New tables whose text starts right-to-left get right-to-left columns.
    for tbl in root.iter(f"{W}tbl"):
        first_row = tbl.find(f"{W}tr")
        if first_row is None:
            continue
        cell_ids = {(p.get(f"{W14}paraId") or "").upper() for p in first_row.iter(f"{W}p")} | {
            (first_row.get(f"{W14}paraId") or "").upper()
        }
        if not cell_ids & created_para_ids:
            continue
        tbl_pr = tbl.find(f"{W}tblPr")
        if tbl_pr is not None and tbl_pr.find(f"{W}bidiVisual") is not None:
            continue
        if first_strong(_para_text(tbl)) != "rtl":
            continue
        if tbl_pr is None:
            tbl_pr = etree.Element(f"{W}tblPr")
            tbl.insert(0, tbl_pr)
        etree.SubElement(tbl_pr, f"{W}bidiVisual")
        changes += 1
    return changes


# ── Tools ───────────────────────────────────────────────────────────────────


class BidiMixin:
    """Paragraph, table and document direction."""

    def set_paragraph_direction(self, para_id: str, direction: str = "rtl") -> dict:
        """Set a paragraph's direction: 'rtl', 'ltr', or 'auto' (from its text)."""
        direction = direction.lower()
        if direction not in ("rtl", "ltr", "auto"):
            raise ValueError("direction must be 'rtl', 'ltr' or 'auto'")
        part, para = self._find_para_any(para_id)
        text = _para_text(para)
        if direction == "auto":
            direction = first_strong(text) or "ltr"
        ppr = _ppr(para)
        bidi = ppr.find(f"{W}bidi")
        if direction == "rtl":
            if bidi is None:
                bidi = etree.SubElement(ppr, f"{W}bidi")
            bidi.attrib.pop(f"{W}val", None)
            for r in _runs_in_para(para):
                chunk = "".join(x.text or "" for x in r.iter(f"{W}t"))
                if has_rtl(chunk) and not any(char_dir(c) == "l" for c in chunk):
                    mark_run_rtl(r, chunk)
        else:
            if bidi is None:
                bidi = etree.SubElement(ppr, f"{W}bidi")
            bidi.set(f"{W}val", "0")
        self._mark(part)
        return {"para_id": para_id, "direction": direction}

    def set_table_direction(self, table_idx: int, direction: str = "rtl") -> dict:
        """Lay a table out right-to-left (first column on the right) or left-to-right."""
        direction = direction.lower()
        if direction not in ("rtl", "ltr"):
            raise ValueError("direction must be 'rtl' or 'ltr'")
        doc = self._require("word/document.xml")
        tables = list(doc.iter(f"{W}tbl"))
        if not 0 <= table_idx < len(tables):
            raise ValueError(f"table_idx {table_idx} out of range (0..{len(tables) - 1})")
        tbl = tables[table_idx]
        tbl_pr = tbl.find(f"{W}tblPr")
        if tbl_pr is None:
            tbl_pr = etree.Element(f"{W}tblPr")
            tbl.insert(0, tbl_pr)
        el = tbl_pr.find(f"{W}bidiVisual")
        if direction == "rtl" and el is None:
            etree.SubElement(tbl_pr, f"{W}bidiVisual")
        elif direction == "ltr" and el is not None:
            tbl_pr.remove(el)
        self._mark("word/document.xml")
        return {"table_index": table_idx, "direction": direction}

    def set_document_direction(
        self,
        direction: str = "rtl",
        *,
        complex_script_font: str = "",
        complex_script_size_pt: float | None = None,
        language: str = "",
        tables: bool = True,
    ) -> dict:
        """Make the whole document right-to-left (or back to left-to-right).

        Sets the document defaults (paragraph direction, complex-script font,
        size and language), every section's direction, and optionally every
        table's direction. Paragraphs with their own explicit direction keep it.

        Args:
            direction: 'rtl' or 'ltr'.
            complex_script_font: Font for Arabic/Hebrew letters, e.g. 'Amiri',
                'Traditional Arabic', 'Sakkal Majalla'. Empty = unchanged.
            complex_script_size_pt: Size for Arabic/Hebrew letters. None = unchanged.
            language: Proofing language for right-to-left text, e.g. 'ar-SA',
                'ur-PK', 'fa-IR'. Empty = 'ar-SA' when switching to rtl.
            tables: Also lay out all tables right-to-left (or left-to-right).
        """
        direction = direction.lower()
        if direction not in ("rtl", "ltr"):
            raise ValueError("direction must be 'rtl' or 'ltr'")
        rtl = direction == "rtl"
        styles = self._require("word/styles.xml")

        doc_defaults = styles.find(f"{W}docDefaults")
        if doc_defaults is None:
            doc_defaults = etree.Element(f"{W}docDefaults")
            styles.insert(0, doc_defaults)
        rpr_default = _get_or_add(doc_defaults, f"{W}rPrDefault")
        ppr_default = doc_defaults.find(f"{W}pPrDefault")
        if ppr_default is None:
            ppr_default = etree.SubElement(doc_defaults, f"{W}pPrDefault")
        drpr = _get_or_add(rpr_default, f"{W}rPr")
        dppr = _get_or_add(ppr_default, f"{W}pPr")

        bidi = dppr.find(f"{W}bidi")
        if rtl and bidi is None:
            etree.SubElement(dppr, f"{W}bidi")
        elif not rtl and bidi is not None:
            dppr.remove(bidi)

        if complex_script_font:
            fonts = _get_or_add(drpr, f"{W}rFonts")
            fonts.set(f"{W}cs", complex_script_font)
            for theme_attr in ("csTheme",):
                fonts.attrib.pop(f"{W}{theme_attr}", None)
        if complex_script_size_pt is not None:
            _get_or_add(drpr, f"{W}szCs").set(f"{W}val", str(round(complex_script_size_pt * 2)))
        lang_code = language or (_DEFAULT_RTL_LANG if rtl else "")
        if lang_code:
            _get_or_add(drpr, f"{W}lang").set(f"{W}bidi", lang_code)

        # The Normal style can override the defaults: keep it consistent.
        for style in styles.findall(f"{W}style"):
            if style.get(f"{W}type") == "paragraph" and style.get(f"{W}default") == "1":
                srpr = style.find(f"{W}rPr")
                if srpr is not None:
                    sfonts = srpr.find(f"{W}rFonts")
                    if complex_script_font and sfonts is not None and sfonts.get(f"{W}cs"):
                        sfonts.set(f"{W}cs", complex_script_font)
                    sszcs = srpr.find(f"{W}szCs")
                    if complex_script_size_pt is not None and sszcs is not None:
                        sszcs.set(f"{W}val", str(round(complex_script_size_pt * 2)))
        self._mark("word/styles.xml")

        doc = self._require("word/document.xml")
        body = doc.find(f"{W}body")
        sections = 0
        for sect in body.iter(f"{W}sectPr"):
            el = sect.find(f"{W}bidi")
            if rtl and el is None:
                etree.SubElement(sect, f"{W}bidi")
            elif not rtl and el is not None:
                sect.remove(el)
            sections += 1
        if sections == 0 and rtl:
            etree.SubElement(etree.SubElement(body, f"{W}sectPr"), f"{W}bidi")
            sections = 1

        table_count = 0
        if tables:
            for i, _ in enumerate(list(doc.iter(f"{W}tbl"))):
                self.set_table_direction(i, direction)
                table_count += 1

        # With a right-to-left default, paragraphs that start with a
        # left-to-right letter must say so explicitly or they flip.
        kept_ltr = 0
        if rtl:
            for p in doc.iter(f"{W}p"):
                if p.find(f"{W}pPr/{W}bidi") is None and first_strong(_para_text(p)) == "ltr":
                    etree.SubElement(_ppr(p), f"{W}bidi").set(f"{W}val", "0")
                    kept_ltr += 1

        # Right-to-left runs everywhere need w:rtl for the complex-script
        # settings to apply.
        runs_marked = 0
        if rtl:
            for p in doc.iter(f"{W}p"):
                for r in _runs_in_para(p):
                    chunk = "".join(x.text or "" for x in r.iter(f"{W}t"))
                    if has_rtl(chunk) and not any(char_dir(c) == "l" for c in chunk):
                        runs_marked += bool(mark_run_rtl(r, chunk))
        self._mark("word/document.xml")
        return {
            "direction": direction,
            "sections": sections,
            "tables": table_count,
            "runs_marked": runs_marked,
            "ltr_paragraphs_kept": kept_ltr,
            "complex_script_font": complex_script_font or None,
            "complex_script_size_pt": complex_script_size_pt,
            "language": lang_code or None,
        }

    def get_direction_report(self) -> dict:
        """Summarise direction: RTL paragraphs, RTL text lacking markup, sections, tables."""
        doc = self._require("word/document.xml")
        paras = list(doc.iter(f"{W}p"))
        rtl_paras = sum(_para_is_rtl(p) for p in paras)
        unmarked_paras = []
        unmarked_runs = 0
        for p in paras:
            text = _para_text(p)
            if first_strong(text) == "rtl" and p.find(f"{W}pPr/{W}bidi") is None:
                unmarked_paras.append({"para_id": p.get(f"{W14}paraId"), "text": text[:60]})
            for r in _runs_in_para(p):
                chunk = "".join(x.text or "" for x in r.iter(f"{W}t"))
                if has_rtl(chunk) and r.find(f"{W}rPr/{W}rtl") is None:
                    unmarked_runs += 1
        sects = list(doc.iter(f"{W}sectPr"))
        tables = list(doc.iter(f"{W}tbl"))
        return {
            "paragraphs": len(paras),
            "rtl_paragraphs": rtl_paras,
            "rtl_text_without_paragraph_direction": unmarked_paras[:50],
            "rtl_runs_without_rtl_flag": unmarked_runs,
            "rtl_sections": sum(s.find(f"{W}bidi") is not None for s in sects),
            "sections": len(sects),
            "rtl_tables": sum(t.find(f"{W}tblPr/{W}bidiVisual") is not None for t in tables),
            "tables": len(tables),
        }

    def _find_para_any(self, para_id: str) -> tuple[str, etree._Element]:
        """Find a paragraph by paraId in the body, notes, headers or footers."""
        for part, root in self._trees.items():
            if not part.startswith("word/") or root is None:
                continue
            for p in root.iter(f"{W}p"):
                if (p.get(f"{W14}paraId") or "").upper() == para_id.upper():
                    return part, p
        raise ValueError(f"Paragraph '{para_id}' not found")
