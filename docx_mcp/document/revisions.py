"""Revision extraction mixin: read tracked changes as structured JSON."""

from __future__ import annotations

import contextlib

from lxml import etree

from .base import CT, RELS, W14, W


class RevisionsMixin:
    """Extract pending tracked changes (w:ins / w:del) as structured data."""

    def get_tracked_changes(self) -> list[dict]:
        """Return all pending tracked changes in document order.

        Each entry:
          {
            "type":      "insertion" | "deletion",
            "change_id": int,
            "author":    str,
            "date":      str,
            "para_id":   str | None,
            "text":      str
          }
        """
        doc = self._require("word/document.xml")
        results: list[dict] = []

        for para in doc.iter(f"{W}p"):
            para_id = para.get(f"{W14}paraId")

            for child in para:
                tag = child.tag
                if tag == f"{W}ins":
                    text = "".join(t.text or "" for t in child.iter(f"{W}t"))
                    results.append(
                        {
                            "type": "insertion",
                            "change_id": _int_id(child.get(f"{W}id")),
                            "author": child.get(f"{W}author", ""),
                            "date": child.get(f"{W}date", ""),
                            "para_id": para_id,
                            "text": text,
                        }
                    )
                elif tag == f"{W}del":
                    text = "".join(dt.text or "" for dt in child.iter(f"{W}delText"))
                    results.append(
                        {
                            "type": "deletion",
                            "change_id": _int_id(child.get(f"{W}id")),
                            "author": child.get(f"{W}author", ""),
                            "date": child.get(f"{W}date", ""),
                            "para_id": para_id,
                            "text": text,
                        }
                    )

        return results

    def set_track_changes(self, enabled: bool, author: str = "") -> dict:
        """Enable or disable revision tracking in word/settings.xml.

        Args:
            enabled: True to enable tracking, False to disable.
            author: Author name for reference (returned in response only).

        Returns:
            {"track_changes": bool, "author": str}
        """
        # Get or create settings tree
        settings = self._tree("word/settings.xml")
        if settings is None:
            settings = etree.Element(
                f"{W}settings",
                nsmap={"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"},
            )
            self._trees["word/settings.xml"] = settings

            # Write to disk so workdir is consistent
            fp = self.workdir / "word" / "settings.xml"
            fp.parent.mkdir(parents=True, exist_ok=True)
            etree.ElementTree(settings).write(str(fp), xml_declaration=True, encoding="UTF-8")

            # Register content-type Override
            ct = self._tree("[Content_Types].xml")
            if ct is not None:
                existing = {e.get("PartName") for e in ct.findall(f"{CT}Override")}
                if "/word/settings.xml" not in existing:
                    ov = etree.SubElement(ct, f"{CT}Override")
                    ov.set("PartName", "/word/settings.xml")
                    ov.set(
                        "ContentType",
                        "application/vnd.openxmlformats-officedocument"
                        ".wordprocessingml.settings+xml",
                    )
                    self._mark("[Content_Types].xml")

            # Register relationship entry
            rels = self._tree("word/_rels/document.xml.rels")
            if rels is not None:
                existing_targets = {r.get("Target") for r in rels.findall(f"{RELS}Relationship")}
                if "settings.xml" not in existing_targets:
                    max_rid = 0
                    for r in rels.findall(f"{RELS}Relationship"):
                        rid = r.get("Id", "")
                        if rid.startswith("rId"):
                            with contextlib.suppress(ValueError):
                                max_rid = max(max_rid, int(rid[3:]))
                    rel = etree.SubElement(rels, f"{RELS}Relationship")
                    rel.set("Id", f"rId{max_rid + 1}")
                    rel.set(
                        "Type",
                        "http://schemas.openxmlformats.org/officeDocument"
                        "/2006/relationships/settings",
                    )
                    rel.set("Target", "settings.xml")
                    self._mark("word/_rels/document.xml.rels")

        tc_tag = f"{W}trackChanges"
        existing = settings.find(tc_tag)

        if enabled:
            if existing is None:
                etree.SubElement(settings, tc_tag)
        else:
            if existing is not None:
                settings.remove(existing)

        self._mark("word/settings.xml")
        return {"track_changes": enabled, "author": author}

    def accept_change(self, change_id: int) -> dict:
        kind = self._change_kind(change_id)
        count, parts = resolve_all(self._trees, accept=True, change_id=change_id)
        for p in parts:
            self._mark(p)
        return {"change_id": change_id, "action": "accepted", "type": kind}

    def _change_kind(self, change_id: int) -> str:
        id_str = str(change_id)
        for part in _story_parts(self._trees):
            root = self._trees[part]
            for el in root.iter(f"{W}ins", f"{W}del"):
                if el.get(f"{W}id") == id_str:
                    return "insertion" if el.tag == f"{W}ins" else "deletion"
        raise ValueError(f"No tracked change with id={change_id}")

    def reject_change(self, change_id: int) -> dict:
        kind = self._change_kind(change_id)
        count, parts = resolve_all(self._trees, accept=False, change_id=change_id)
        for p in parts:
            self._mark(p)
        return {"change_id": change_id, "action": "rejected", "type": kind}

    def accept_all_changes(self) -> dict:
        count, parts = resolve_all(self._trees, accept=True)
        for p in parts:
            self._mark(p)
        return {"accepted": count}

    def flatten_document(self) -> dict:
        """Accept all tracked changes and strip all revision markup.

        Accepts every w:ins / w:del element (via accept_all_changes) and then
        removes all w:rPrChange and w:pPrChange elements so no revision markup
        remains in the document.

        Returns:
            {"changes_accepted": int, "formatting_changes_removed": int}
        """
        accepted_result = self.accept_all_changes()
        changes_accepted: int = accepted_result["accepted"]

        doc = self._require("word/document.xml")
        fmt_count = 0
        for tag in (f"{W}rPrChange", f"{W}pPrChange"):
            for el in list(doc.iter(tag)):
                el.getparent().remove(el)
                fmt_count += 1

        if fmt_count:
            self._mark("word/document.xml")
        # accept_all_changes already marks dirty when count > 0

        return {
            "changes_accepted": changes_accepted,
            "formatting_changes_removed": fmt_count,
        }

    def reject_all_changes(self) -> dict:
        count, parts = resolve_all(self._trees, accept=False)
        for p in parts:
            self._mark(p)
        return {"rejected": count}


def _find_change(doc, change_id: int):
    id_str = str(change_id)
    for tag in (f"{W}ins", f"{W}del"):
        for el in doc.iter(tag):
            if el.get(f"{W}id") == id_str:
                return el
    return None


def _unwrap(el) -> None:
    parent = el.getparent()
    idx = list(parent).index(el)
    children = list(el)
    parent.remove(el)
    for i, child in enumerate(children):
        parent.insert(idx + i, child)


def _unwrap_del(el) -> None:
    for r in el.findall(f"{W}r"):
        for dt in r.findall(f"{W}delText"):
            t = etree.Element(f"{W}t")
            t.text = dt.text
            if dt.text and any(c.isspace() for c in dt.text):
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            r.remove(dt)
            r.append(t)
    _unwrap(el)


def _int_id(val: str | None) -> int:
    try:
        return int(val or 0)
    except (ValueError, TypeError):
        return 0


# ── Unified revision resolution ─────────────────────────────────────────────

_STORY_PREFIXES = (
    "word/document.xml",
    "word/footnotes.xml",
    "word/endnotes.xml",
    "word/header",
    "word/footer",
)


def _story_parts(trees: dict) -> list[str]:
    return [p for p in trees if p.startswith(_STORY_PREFIXES) and p.endswith(".xml")]


def _kind(el) -> str:
    """Classify a w:ins / w:del by where it sits.

    'row'  : row insertion/deletion (inside w:trPr)
    'mark' : paragraph-mark insertion/deletion (inside w:pPr/w:rPr)
    'inline': wraps runs (the common case)
    'other': a property-level marker we only strip
    """
    parent = el.getparent()
    if parent is None:
        return "other"
    if parent.tag == f"{W}trPr":
        return "row"
    if parent.tag == f"{W}rPr":
        gp = parent.getparent()
        if gp is not None and gp.tag == f"{W}pPr":
            return "mark"
        return "other"
    return "inline"


def _merge_with_next(para) -> None:
    """Join a paragraph with the next one (its paragraph mark goes away)."""
    nxt = para.getnext()
    if nxt is None or nxt.tag != f"{W}p":
        return
    content = [c for c in para if c.tag != f"{W}pPr"]
    anchor = nxt.find(f"{W}pPr")
    idx = 0 if anchor is None else list(nxt).index(anchor) + 1
    for i, c in enumerate(content):
        nxt.insert(idx + i, c)
    para.getparent().remove(para)


def _para_has_content(para) -> bool:
    return any(c.tag != f"{W}pPr" for c in para)


def _remove_row(tr) -> None:
    tbl = tr.getparent()
    tbl.remove(tr)
    if tbl.tag == f"{W}tbl" and tbl.find(f"{W}tr") is None and tbl.getparent() is not None:
        tbl.getparent().remove(tbl)


def resolve_revision(el, accept: bool) -> None:
    """Accept or reject one w:ins / w:del, whatever it is attached to."""
    if el.getparent() is None:
        return
    is_ins = el.tag == f"{W}ins"
    kind = _kind(el)
    if kind == "inline":
        if is_ins:
            _unwrap(el) if accept else el.getparent().remove(el)
        else:
            el.getparent().remove(el) if accept else _unwrap_del(el)
        return
    parent = el.getparent()
    parent.remove(el)
    if kind == "row":
        tr = parent.getparent()
        drop_row = (is_ins and not accept) or (not is_ins and accept)
        if drop_row and tr is not None and tr.getparent() is not None:
            _remove_row(tr)
        return
    if kind == "mark":
        ppr = parent.getparent()
        para = ppr.getparent() if ppr is not None else None
        if len(parent) == 0 and not parent.attrib:
            ppr.remove(parent)
        if para is None or para.getparent() is None:
            return
        if (is_ins and not accept) or (not is_ins and accept):
            if not _para_has_content(para) and para.getnext() is not None:
                para.getparent().remove(para)
            else:
                _merge_with_next(para)


def resolve_all(
    trees: dict, accept: bool, author: str | None = None, change_id: int | None = None
) -> tuple[int, list[str]]:
    """Resolve matching revisions in every story part.

    Inline changes are handled first, then paragraph marks, then table rows,
    so merges and row removals see already-resolved content.
    Returns (count, modified_parts).
    """
    count = 0
    modified: list[str] = []
    id_str = None if change_id is None else str(change_id)
    for part in _story_parts(trees):
        root = trees[part]
        if root is None:
            continue
        found = [
            el
            for el in root.iter(f"{W}ins", f"{W}del")
            if (author is None or el.get(f"{W}author") == author)
            and (id_str is None or el.get(f"{W}id") == id_str)
        ]
        if not found:
            continue
        order = {"inline": 0, "other": 0, "mark": 1, "row": 2}
        found.sort(key=lambda e: order[_kind(e)])
        # Inline revisions nested inside another matching revision go with it.
        for el in found:
            if el.getparent() is None:
                continue
            resolve_revision(el, accept)
            count += 1
        modified.append(part)
    return count, modified
