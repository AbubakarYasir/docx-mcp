"""Honorific phrases as their Unicode ligature glyphs (ﷺ ﵁ ﵀ ﷻ …).

Typing «صلى الله عليه وسلم» or the abbreviation (ص) gives plain text; the
Unicode standard has single ligature characters for these phrases (U+FDFA
and the Arabic Extended block U+FD40–FD4F added in Unicode 14). Few fonts
draw them: of the common Arabic fonts only Noto Naskh Arabic and Scheherazade
New cover the full set, Amiri / Traditional Arabic / Sakkal Majalla only ﷺ.
So each glyph is put in its own run in a font that has it, with its own
colour and size, while the surrounding text keeps its font.

Matching ignores tashkeel and tatweel, accepts the phrase in brackets or
parentheses, and (optionally) the common abbreviations.
"""

from __future__ import annotations

import copy
import re

from lxml import etree

from .ooxml import XML_SPACE, W, apply_run_props

# (glyph, Unicode name, full phrases, abbreviations)
HONORIFICS: list[tuple[str, str, list[str], list[str]]] = [
    ("ﷺ", "SALLALLAHOU ALAYHE WASALLAM", ["صلى الله عليه وسلم"], ["صلعم", "ص", "صلى"]),
    (
        "﵌",
        "SALLALLAHU ALAYHI WAAALIHEE WA-SALLAM",
        ["صلى الله عليه وآله وسلم", "صلى الله عليه وعلى آله وسلم"],
        [],
    ),
    ("﵆", "SALLALLAAHU ALAYHI WA-AALIH", ["صلى الله عليه وآله"], []),
    ("﵄", "RADI ALLAAHU ANHUMAA", ["رضي الله عنهما"], []),
    ("﵅", "RADI ALLAAHU ANHUNNA", ["رضي الله عنهن"], []),
    ("﵃", "RADI ALLAAHU ANHUM", ["رضي الله عنهم"], []),
    ("﵂", "RADI ALLAAHU ANHAA", ["رضي الله عنها"], []),
    ("﵁", "RADI ALLAAHU ANH", ["رضي الله عنه"], ["رض", "رضه"]),
    ("﵏", "RAHIMAHUM ALLAAH", ["رحمهم الله"], []),
    ("﵀", "RAHIMAHU ALLAAH", ["رحمه الله"], ["رح", "رحه"]),
    ("﵊", "ALAYHI AS-SALAATU WAS-SALAAM", ["عليه الصلاة والسلام"], []),
    ("﵉", "ALAYHIMAA AS-SALAAM", ["عليهما السلام"], []),
    ("﵈", "ALAYHIM AS-SALAAM", ["عليهم السلام"], []),
    ("﵍", "ALAYHAA AS-SALAAM", ["عليها السلام"], []),
    ("﵇", "ALAYHI AS-SALAAM", ["عليه السلام"], ["ع", "عم"]),
    ("ﷻ", "JALLAJALALOUHOU", ["جل جلاله"], []),
    ("﷾", "SUBHAANAHU WA TAAALAA", ["سبحانه وتعالى"], []),
    ("﷿", "AZZA WA JALL", ["عز وجل"], []),
    ("﵎", "TABAARAKA WA-TAAALAA", ["تبارك وتعالى"], []),
    ("﵋", "QUDDISA SIRRAH", ["قدس سره"], []),
]

GLYPHS = {g for g, *_ in HONORIFICS}
DEFAULT_GLYPH_FONT = "Scheherazade New"
FULL_COVERAGE_FONTS = ("Scheherazade New", "Noto Naskh Arabic")

_MARKS = "ً-ٰٟۖ-ۭـ"  # harakat, Quranic marks, tatweel
_SKIP = f"[{_MARKS}]*"
_ALEF = {"ا": "[اأإآٱ]", "ى": "[ىي]", "ي": "[يى]", "ه": "[هة]", "ة": "[ةه]"}


def _flex(phrase: str) -> str:
    """Regex for a phrase: optional tashkeel/tatweel after every letter,
    alef/ya/ta-marbuta variants, flexible spacing."""
    out = []
    for ch in phrase:
        if ch == " ":
            out.append(r"\s+")
            continue
        out.append(_ALEF.get(ch, re.escape(ch)) + _SKIP)
    return "".join(out)


def _pattern(phrases: list[str], abbreviations: list[str]) -> str:
    alts = [_flex(p) for p in sorted(phrases, key=len, reverse=True)]
    full = "|".join(alts)
    # The phrase itself, optionally in () or [] — brackets are consumed.
    parts = [rf"[\(\[]\s*(?:{full})\s*[\)\]]", rf"(?:{full})"]
    if abbreviations:
        abbr = "|".join(_flex(a) for a in sorted(abbreviations, key=len, reverse=True))
        # Abbreviations only inside brackets: «(ص)», «(رض)» — bare letters are words.
        parts.insert(0, rf"[\(\[]\s*(?:{abbr})\s*[\)\]]")
    return "|".join(parts)


_ARABIC_LETTER = r"[ء-يٮ-ۓۺ-ۿ]"


def build_matcher(
    include_abbreviations: bool = True, only: list[str] | None = None
) -> list[tuple[re.Pattern, str]]:
    """[(compiled regex, glyph)] ordered longest phrase first."""
    out = []
    for glyph, _name, phrases, abbr in HONORIFICS:
        if only and glyph not in only and not set(phrases) & set(only):
            continue
        pat = _pattern(phrases, abbr if include_abbreviations else [])
        # Not part of a longer word on either side.
        rx = re.compile(rf"(?<!{_ARABIC_LETTER})(?:{pat})(?!{_ARABIC_LETTER})")
        out.append((rx, glyph))
    return out


def find_matches(text: str, matcher) -> list[tuple[int, int, str]]:
    """Non-overlapping (start, end, glyph) spans, longest first."""
    spans: list[tuple[int, int, str]] = []
    taken = [False] * len(text)
    for rx, glyph in matcher:
        for m in rx.finditer(text):
            s, e = m.span()
            if any(taken[s:e]):
                continue
            spans.append((s, e, glyph))
            for k in range(s, e):
                taken[k] = True
    spans.sort()
    return spans


def _text_runs(p: etree._Element) -> list[etree._Element]:
    """Runs whose only content is rPr + one w:t (safe to split)."""
    runs = []
    for r in p.iter(f"{W}r"):
        kids = [c for c in r if c.tag != f"{W}rPr"]
        if len(kids) == 1 and kids[0].tag == f"{W}t" and kids[0].text:
            anc = r.getparent()
            if anc is not None and anc.tag in (f"{W}del", f"{W}moveFrom"):
                continue
            runs.append(r)
    return runs


def replace_in_run(r: etree._Element, matcher, glyph_props: dict) -> int:
    """Split one run around honorific matches; glyphs get their own run."""
    t = r.find(f"{W}t")
    text = t.text or ""
    spans = find_matches(text, matcher)
    if not spans:
        return 0
    parent = r.getparent()
    idx = list(parent).index(r)
    rpr = r.find(f"{W}rPr")
    pieces: list[tuple[str, bool]] = []
    pos = 0
    for s, e, glyph in spans:
        before = text[pos:s]
        # A space before the honorific is kept; brackets were consumed.
        if before:
            pieces.append((before, False))
        pieces.append((glyph, True))
        pos = e
    if pos < len(text):
        pieces.append((text[pos:], False))
    new_runs = []
    for chunk, is_glyph in pieces:
        nr = etree.Element(f"{W}r")
        if rpr is not None:
            nr.append(copy.deepcopy(rpr))
        if is_glyph:
            nrpr = nr.find(f"{W}rPr")
            if nrpr is None:
                nrpr = etree.SubElement(nr, f"{W}rPr")
            apply_run_props(nrpr, glyph_props)
            if nrpr.find(f"{W}rtl") is None:
                etree.SubElement(nrpr, f"{W}rtl")
        nt = etree.SubElement(nr, f"{W}t")
        nt.text = chunk
        nt.set(XML_SPACE, "preserve")
        new_runs.append(nr)
    for i, nr in enumerate(new_runs):
        parent.insert(idx + i, nr)
    parent.remove(r)
    return len(spans)


def restyle_glyph_runs(root: etree._Element, glyph_props: dict) -> int:
    """Give runs that already contain honorific glyphs the glyph formatting."""
    n = 0
    for p in root.iter(f"{W}p"):
        for r in _text_runs(p):
            t = r.find(f"{W}t")
            text = t.text or ""
            if not any(c in GLYPHS for c in text):
                continue
            if all(c in GLYPHS or c.isspace() for c in text):
                rpr = r.find(f"{W}rPr")
                if rpr is None:
                    rpr = etree.Element(f"{W}rPr")
                    r.insert(0, rpr)
                apply_run_props(rpr, glyph_props)
                n += 1
                continue
            # Split glyphs out of a mixed run.
            parent = r.getparent()
            idx = list(parent).index(r)
            rpr = r.find(f"{W}rPr")
            chunks = re.findall(rf"[{''.join(GLYPHS)}]+|[^{''.join(GLYPHS)}]+", text)
            for i, chunk in enumerate(chunks):
                nr = etree.Element(f"{W}r")
                if rpr is not None:
                    nr.append(copy.deepcopy(rpr))
                if chunk[0] in GLYPHS:
                    nrpr = nr.find(f"{W}rPr")
                    if nrpr is None:
                        nrpr = etree.SubElement(nr, f"{W}rPr")
                    apply_run_props(nrpr, glyph_props)
                    n += 1
                nt = etree.SubElement(nr, f"{W}t")
                nt.text = chunk
                nt.set(XML_SPACE, "preserve")
                parent.insert(idx + i, nr)
            parent.remove(r)
    return n


class HonorificsMixin:
    """Replace honorific phrases with ligature glyphs; style the glyphs."""

    def list_honorifics(self) -> list[dict]:
        """The supported honorific glyphs, the phrases they replace, and abbreviations."""
        return [
            {
                "glyph": g,
                "codepoint": f"U+{ord(g):04X}",
                "name": name,
                "phrases": phrases,
                "abbreviations_in_brackets": abbr,
            }
            for g, name, phrases, abbr in HONORIFICS
        ]

    def apply_honorifics(
        self,
        *,
        font: str = DEFAULT_GLYPH_FONT,
        color: str | None = None,
        size_pt: float | None = None,
        bold: bool | None = None,
        include_abbreviations: bool = True,
        only: list[str] | None = None,
        restyle_existing: bool = True,
        parts: str = "all",
    ) -> dict:
        """Replace honorific phrases with their glyphs and format every glyph.

        Args:
            font: Font for the glyphs. Scheherazade New and Noto Naskh Arabic
                draw the full set; most other Arabic fonts only draw ﷺ.
            color: Glyph colour (RRGGBB), e.g. "520912" for maroon. None = text colour.
            size_pt: Glyph size. None = same as the text.
            bold: Force bold on/off for glyphs. None = keep.
            include_abbreviations: Also replace bracketed abbreviations such as
                (ص), (صلعم), (رض), (رح), (ع).
            only: Limit to these glyphs or phrases.
            restyle_existing: Also format glyphs already in the text.
            parts: "all" (body, footnotes, endnotes, headers, footers, comments)
                or "body".
        """
        matcher = build_matcher(include_abbreviations, only)
        props: dict = {"font": font}
        if color:
            props["color"] = color
        if size_pt is not None:
            props["size_pt"] = size_pt
        if bold is not None:
            props["bold"] = bold
        prefixes = (
            ("word/document.xml",)
            if parts == "body"
            else (
                "word/document",
                "word/footnotes",
                "word/endnotes",
                "word/header",
                "word/footer",
                "word/comments",
            )
        )
        replaced: dict[str, int] = {}
        restyled = 0
        total = 0
        for part, root in list(self._trees.items()):
            if root is None or not part.startswith(prefixes) or not part.endswith(".xml"):
                continue
            n = 0
            for p in root.iter(f"{W}p"):
                for r in _text_runs(p):
                    n += replace_in_run(r, matcher, props)
            m = restyle_glyph_runs(root, props) if restyle_existing else 0
            if n or m:
                self._mark(part)
            if n:
                replaced[part] = n
            restyled += m
            total += n
        return {
            "replaced": total,
            "by_part": replaced,
            "glyph_runs_formatted": restyled,
            "font": font,
            "font_has_full_set": font in FULL_COVERAGE_FONTS,
        }
