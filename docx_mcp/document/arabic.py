"""Arabic text utilities: tashkil-aware matching and tashkil operations.

Scholarly text is often fully vocalised. Every search in the toolkit
(honorifics, index marking, formatting by match, styles by match) goes
through ``loose_pattern`` so «صَلَّى اللَّهُ عَلَيْهِ وَسَلَّمَ» is found by
«صلى الله عليه وسلم» and «البُخَارِيُّ» by «البخاري».
"""

from __future__ import annotations

import re
import unicodedata

# harakat, tanwin, shadda, sukun, superscript alef, Quranic marks, tatweel
HARAKAT = "ًٌٍَُِّْٰٕٖٓٔٗ٘ٞ"
QURANIC = "".join(chr(c) for c in range(0x06D6, 0x06EE)) + "ؘؙؚؐؑؒؓؔؕؖؗ"
TATWEEL = "ـ"
MARKS = HARAKAT + QURANIC
_MARK_CLASS = "[" + re.escape(MARKS + TATWEEL) + "]*"
_ARABIC_LETTER = re.compile(r"[ء-يٱ-ۓۺ-ۿ]")
SHADDA = "ّ"
VOWELS = "ًٌٍَُِ"
SUKUN = "ْ"
_ALEF_VARIANTS = {
    "ا": "[اأإآٱ]",
    "أ": "[اأإآٱ]",
    "إ": "[اأإآٱ]",
    "آ": "[اأإآٱ]",
    "ٱ": "[اأإآٱ]",
    "ي": "[يى]",
    "ى": "[يى]",
    "ة": "[ةه]",
    "ه": "[هة]",
}


def strip_tashkil(text: str, keep_shadda: bool = False) -> str:
    marks = MARKS.replace(SHADDA, "") if keep_shadda else MARKS
    return "".join(ch for ch in text if ch not in marks and ch != TATWEEL)


def has_tashkil(text: str) -> bool:
    return any(ch in HARAKAT for ch in text)


def tashkil_ratio(text: str) -> float:
    """Share of Arabic letters followed by at least one haraka."""
    letters = vocal = 0
    for i, ch in enumerate(text):
        if _ARABIC_LETTER.match(ch):
            letters += 1
            j = i + 1
            while j < len(text) and text[j] == TATWEEL:
                j += 1
            if j < len(text) and text[j] in HARAKAT:
                vocal += 1
    return vocal / letters if letters else 0.0


def loose_pattern(literal: str, *, alef_loose: bool = False) -> str:
    """A regex source for literal that ignores tashkil and tatweel (and, with
    alef_loose, hamza forms of alef, alef maqsura and ta marbuta)."""
    base = strip_tashkil(literal)
    out = []
    for ch in base:
        if alef_loose and ch in _ALEF_VARIANTS:
            out.append(_ALEF_VARIANTS[ch])
        else:
            out.append(re.escape(ch))
        if _ARABIC_LETTER.match(ch):
            out.append(_MARK_CLASS)
    return "".join(out)


def normalize_tashkil(text: str) -> str:
    """Canonical mark order (shadda before its vowel), duplicate marks
    removed, NFC — fixes text pasted from different sources that renders
    harakat in the wrong place or twice."""
    text = unicodedata.normalize("NFC", text)
    out: list[str] = []
    i = 0
    while i < len(text):
        ch = text[i]
        if ch in MARKS:
            j = i
            cluster = []
            while j < len(text) and text[j] in MARKS:
                if text[j] not in cluster:
                    cluster.append(text[j])
                j += 1
            cluster.sort(key=lambda m: (m != SHADDA, m not in VOWELS, m))
            out.extend(cluster)
            i = j
        else:
            out.append(ch)
            i += 1
    return "".join(out)


def strip_case_endings(text: str) -> str:
    """Remove the final haraka/tanwin of each word (i'rab), keeping internal
    vocalisation — the "light tashkil" many books use."""

    def fix(m: re.Match) -> str:
        word = m.group(0)
        k = len(word)
        while k > 0 and word[k - 1] in MARKS and word[k - 1] != SHADDA:
            k -= 1
        return word[:k]

    return re.sub(r"[؀-ۿ]+", fix, text)


# ── document operations ─────────────────────────────────────────────────

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
_PARTS = {
    "body": ("word/document.xml",),
    "notes": ("word/footnotes.xml", "word/endnotes.xml"),
    "all": ("word/document.xml", "word/footnotes.xml", "word/endnotes.xml"),
}


class TashkilMixin:
    """Tashkil (harakat) operations on a document."""

    def _tashkil_texts(self, parts: str, para_ids: list[str] | None):
        if parts not in _PARTS:
            raise ValueError("parts must be body, notes or all")
        for part in _PARTS[parts]:
            root = self._tree(part)
            if root is None:
                continue
            for p in root.iter(f"{_W}p"):
                if para_ids and p.get(f"{_W14}paraId") not in para_ids:
                    continue
                for t in p.iter(f"{_W}t"):
                    yield part, p, t

    def tashkil(
        self, mode: str = "normalize", *, parts: str = "all", para_ids: list[str] | None = None
    ) -> dict:
        """Change the vocalisation of the text.

        mode:
          normalize    — canonical mark order, duplicates removed (fixes
                         harakat that render misplaced after pasting).
          strip        — remove all tashkil.
          keep-shadda  — remove all but shadda.
          light        — remove case endings only (i'rab), keep the rest.
        parts: body, notes or all. para_ids: limit to these paragraphs.
        """
        ops = {
            "normalize": normalize_tashkil,
            "strip": strip_tashkil,
            "keep-shadda": lambda t: strip_tashkil(t, keep_shadda=True),
            "light": strip_case_endings,
        }
        if mode not in ops:
            raise ValueError(f"mode must be one of {sorted(ops)}")
        changed = 0
        touched = set()
        for part, _p, t in self._tashkil_texts(parts, para_ids):
            if not t.text:
                continue
            new = ops[mode](t.text)
            if new != t.text:
                t.text = new
                changed += 1
                touched.add(part)
        for part in touched:
            self._mark(part)
        return {"mode": mode, "runs_changed": changed}

    def tashkil_report(self, parts: str = "all") -> dict:
        """How vocalised the text is: overall and per paragraph (fully,
        partly, unvocalised), so a book's consistency can be checked."""
        paras = {}
        for part, p, t in self._tashkil_texts(parts, None):
            key = (part, p.get(f"{_W14}paraId"))
            paras.setdefault(key, []).append(t.text or "")
        rows = []
        full = partial = none = 0
        all_text = []
        for (part, pid), texts in paras.items():
            text = "".join(texts)
            if not _ARABIC_LETTER.search(text):
                continue
            r = tashkil_ratio(text)
            all_text.append(text)
            band = "full" if r >= 0.7 else "partial" if r >= 0.1 else "none"
            full += band == "full"
            partial += band == "partial"
            none += band == "none"
            rows.append(
                {
                    "para_id": pid,
                    "part": part.split("/")[-1],
                    "ratio": round(r, 2),
                    "level": band,
                    "start": text[:40],
                }
            )
        return {
            "overall": round(tashkil_ratio("".join(all_text)), 2),
            "full": full,
            "partial": partial,
            "unvocalised": none,
            "paragraphs": rows,
        }

    def color_tashkil(
        self,
        color: str | None = None,
        *,
        parts: str = "body",
        para_ids: list[str] | None = None,
        per_run: bool = False,
    ) -> dict:
        """Draw the harakat in their own colour (a teaching / Qur'an
        convention), letters unchanged.

        Word colours diacritics with one colour for the whole document (its
        "Use this color for diacritics" option), not run by run. The colour is
        stored in the document (docx-mcp.diacritic-color) and export_pdf turns
        that option on for the export, so the PDF shows it; in Word itself the
        user can switch the option on to see it. per_run additionally colours
        the marks run by run (parts / para_ids) for LibreOffice and other
        renderers that honour it. color None uses the house accent; "auto" or
        "none" removes the colouring.
        """
        import copy

        from lxml import etree

        from .ooxml import color_hex

        if color is None:
            color = self._house("accent") if hasattr(self, "_house") else "auto"
        hx = None if str(color).lower() in ("auto", "none", "") else color_hex(color)
        if hx:
            self.set_custom_property("docx-mcp.diacritic-color", hx)
        elif "docx-mcp.diacritic-color" in self.get_custom_properties():
            self.delete_custom_property("docx-mcp.diacritic-color")
        if not per_run and hx:
            return {"color": hx, "runs_split": 0, "scope": "document (Word diacritic colour)"}

        def paint(rpr) -> None:
            old = rpr.find(f"{_W}color")
            if old is not None:
                rpr.remove(old)
            if hx:
                c = etree.Element(f"{_W}color")
                c.set(f"{_W}val", hx)
                # CT_RPr order: color sits before spacing/w/kern/position/sz…
                after = {"spacing", "w", "kern", "position", "sz", "szCs", "highlight",
                         "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl",
                         "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"}  # fmt: skip
                for i, ch in enumerate(rpr):
                    if etree.QName(ch).localname in after:
                        rpr.insert(i, c)
                        break
                else:
                    rpr.append(c)

        split = 0
        for part in _PARTS[parts]:
            root = self._tree(part)
            if root is None:
                continue
            for r in list(root.iter(f"{_W}r")):
                p = r.getparent()
                while p is not None and p.tag != f"{_W}p":
                    p = p.getparent()
                if para_ids and (p is None or p.get(f"{_W14}paraId") not in para_ids):
                    continue
                kids = [c for c in r if c.tag != f"{_W}rPr"]
                if len(kids) != 1 or kids[0].tag != f"{_W}t" or not kids[0].text:
                    continue
                text = kids[0].text
                if not any(ch in HARAKAT for ch in text):
                    continue
                segs: list[tuple[bool, str]] = []
                for ch in text:
                    is_mark = ch in HARAKAT
                    if segs and segs[-1][0] == is_mark:
                        segs[-1] = (is_mark, segs[-1][1] + ch)
                    else:
                        segs.append((is_mark, ch))
                if len(segs) == 1:
                    if segs[0][0]:  # a marks-only run from an earlier call
                        rpr = r.find(f"{_W}rPr")
                        if rpr is None:
                            rpr = etree.Element(f"{_W}rPr")
                            r.insert(0, rpr)
                        paint(rpr)
                        split += 1
                    continue
                base_rpr = r.find(f"{_W}rPr")
                parent = r.getparent()
                idx = list(parent).index(r)
                for k, (is_mark, chunk) in enumerate(segs):
                    nr = etree.Element(f"{_W}r")
                    rpr = (
                        copy.deepcopy(base_rpr)
                        if base_rpr is not None
                        else etree.Element(f"{_W}rPr")
                    )
                    if is_mark:
                        paint(rpr)
                    # letters and marks must both be RTL runs, or Word draws
                    # the word as separate letters
                    if rpr.find(f"{_W}rtl") is None:
                        etree.SubElement(rpr, f"{_W}rtl")
                    lang = rpr.find(f"{_W}lang")
                    if lang is None:
                        lang = etree.SubElement(rpr, f"{_W}lang")
                    if lang.get(f"{_W}bidi") is None:
                        lang.set(f"{_W}bidi", "ar-SA")
                    if len(rpr):
                        nr.append(rpr)
                    t = etree.SubElement(nr, f"{_W}t")
                    t.text = chunk
                    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    parent.insert(idx + k, nr)
                parent.remove(r)
                split += 1
            self._mark(part)
        return {"color": hx, "runs_split": split, "scope": "runs" if hx else "removed"}
