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

# One entry per honorific. A glyph is only meaningful together with the font
# it was drawn for: Unicode assigns official meanings (U+FDFA, U+FD40-FD4F …),
# but some popular Urdu fonts draw their own artwork at other code points —
# and even reuse official ones for a different phrase (Jameel Noori Nastaleeq
# draws حفظه الله at U+FDFE, which Unicode defines as سبحانه وتعالى). So every
# glyph set carries its font, and a phrase missing from the chosen set falls
# back to another set *with that set's font*, never to a bare code point.
#
# fields: key, phrases, bracketed abbreviations, {glyph set: code point}
HONORIFICS: list[tuple[str, list[str], list[str], dict[str, str]]] = [
    (
        "salla",
        ["صلى الله عليه وسلم"],
        ["صلعم", "ص", "صلى"],
        {"unicode": "\ufdfa", "jameel": "\ufdfa"},
    ),
    (
        "salla-wa-alih",
        ["صلى الله عليه وآله وسلم", "صلى الله عليه وعلى آله وسلم"],
        [],
        {"unicode": "\ufd4c"},
    ),
    ("salla-alih", ["صلى الله عليه وآله"], [], {"unicode": "\ufd46"}),
    ("radi-anhuma", ["رضي الله عنهما"], [], {"unicode": "\ufd44", "jameel": "\ufde0"}),
    ("radi-anhunna", ["رضي الله عنهن"], [], {"unicode": "\ufd45"}),
    ("radi-anhum", ["رضي الله عنهم"], [], {"unicode": "\ufd43", "jameel": "\ufde2"}),
    ("radi-anha", ["رضي الله عنها"], [], {"unicode": "\ufd42", "jameel": "\ufdde"}),
    ("radi-anhu", ["رضي الله عنه"], ["رض", "رضه"], {"unicode": "\ufd41", "jameel": "\ufddc"}),
    ("rahimahum", ["رحمهم الله"], [], {"unicode": "\ufd4f"}),
    ("rahimahu", ["رحمه الله"], ["رح", "رحه"], {"unicode": "\ufd40"}),
    ("rahmatullah-alayhima", ["رحمة الله عليهما"], [], {"jameel": "\ufdee"}),
    ("rahmatullah-alayhim", ["رحمة الله عليهم"], [], {"jameel": "\ufded"}),
    ("rahmatullah-alayhinna", ["رحمة الله عليهن"], [], {"jameel": "\ue010"}),
    ("rahmatullah-alayha", ["رحمة الله عليها"], [], {"jameel": "\ue007"}),
    ("rahmatullah-alayh", ["رحمة الله عليه"], [], {"jameel": "\ufdec"}),
    ("hafizahum", ["حفظهم الله"], [], {"jameel": "\ue006"}),
    ("hafizaha", ["حفظها الله"], [], {"jameel": "\ue005"}),
    ("hafizahu", ["حفظه الله"], [], {"jameel": "\ufdfe"}),
    ("alayhi-salat", ["عليه الصلاة والسلام"], [], {"unicode": "\ufd4a", "jameel": "\ufdea"}),
    ("alayhima-salam", ["عليهما السلام"], [], {"unicode": "\ufd49", "jameel": "\ufde7"}),
    ("alayhim-salam", ["عليهم السلام"], [], {"unicode": "\ufd48", "jameel": "\ufde9"}),
    ("alayha-salam", ["عليها السلام"], [], {"unicode": "\ufd4d", "jameel": "\ufde5"}),
    ("alayhi-salam", ["عليه السلام"], ["ع", "عم"], {"unicode": "\ufd47", "jameel": "\ufde4"}),
    ("jalla-jalaluh", ["جل جلاله"], [], {"unicode": "\ufdfb", "jameel": "\ufdfb"}),
    ("subhanahu", ["سبحانه وتعالى"], [], {"unicode": "\ufdfe"}),
    ("azza-wa-jall", ["عز وجل"], [], {"unicode": "\ufdff", "jameel": "\ufdef"}),
    ("tabaraka", ["تبارك وتعالى"], [], {"unicode": "\ufd4e"}),
    ("quddisa-sirruh", ["قدس سره"], [], {"unicode": "\ufd4b"}),
]

GLYPH_SETS = {
    "unicode": {
        "font": "Scheherazade New",
        "fonts_with_full_set": ("Scheherazade New", "Noto Naskh Arabic"),
    },
    "jameel": {
        "font": "Jameel Noori Nastaleeq",
        "fonts_with_full_set": ("Jameel Noori Nastaleeq",),
    },
}
DEFAULT_GLYPH_FONT = GLYPH_SETS["unicode"]["font"]
FULL_COVERAGE_FONTS = GLYPH_SETS["unicode"]["fonts_with_full_set"]
GLYPHS = {g for *_, sets in HONORIFICS for g in sets.values()} | {
    "\ufdd0",
    "\ufdd1",
    "\ufddb",
    "\ufddd",
    "\ufddf",
    "\ufde1",
    "\ufde3",
    "\ufde6",
    "\ufde8",
    "\ufdeb",
    "\ufdff",
}

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
    """[(compiled regex, honorific key)] ordered longest phrase first."""
    out = []
    for key, phrases, abbr, sets in HONORIFICS:
        if (
            only
            and key not in only
            and not set(phrases) & set(only)
            and not (set(sets.values()) & set(only))
        ):
            continue
        pat = _pattern(phrases, abbr if include_abbreviations else [])
        rx = re.compile(rf"(?<!{_ARABIC_LETTER})(?:{pat})(?!{_ARABIC_LETTER})")
        out.append((max(len(p) for p in phrases), rx, key))
    # Longest phrase first everywhere: «رحمة الله عليه» before «رحمه الله».
    out.sort(key=lambda t: -t[0])
    return [(rx, key) for _, rx, key in out]


_BY_KEY = {key: sets for key, _p, _a, sets in HONORIFICS}


def resolve_glyph(key: str, glyph_set: str, fonts: dict[str, str]) -> tuple[str, str] | None:
    """(glyph, font) for an honorific in the chosen set, else another set."""
    sets = _BY_KEY[key]
    order = [glyph_set] + [s for s in GLYPH_SETS if s != glyph_set]
    for name in order:
        if name in sets:
            return sets[name], fonts.get(name) or GLYPH_SETS[name]["font"]
    return None


def find_matches(text: str, matcher) -> list[tuple[int, int, str]]:
    """Non-overlapping (start, end, honorific key) spans, longest first."""
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


def replace_in_run(r: etree._Element, matcher, resolve, base_props: dict, stats: dict) -> int:
    """Split one run around honorific matches; each glyph gets its own run
    in the font of its glyph set."""
    t = r.find(f"{W}t")
    text = t.text or ""
    spans = []
    for s, e, key in find_matches(text, matcher):
        res = resolve(key)
        if res is None:
            stats["no_glyph"][key] = stats["no_glyph"].get(key, 0) + 1
            continue
        spans.append((s, e, key, *res))
    if not spans:
        return 0
    parent = r.getparent()
    idx = list(parent).index(r)
    rpr = r.find(f"{W}rPr")
    pieces: list[tuple[str, dict | None]] = []
    pos = 0
    for s, e, key, glyph, font in spans:
        if text[pos:s]:
            pieces.append((text[pos:s], None))
        pieces.append((glyph, {**base_props, "font": font}))
        stats["by_key"][key] = stats["by_key"].get(key, 0) + 1
        stats["fonts"][font] = stats["fonts"].get(font, 0) + 1
        pos = e
    if pos < len(text):
        pieces.append((text[pos:], None))
    for i, (chunk, props) in enumerate(pieces):
        nr = etree.Element(f"{W}r")
        if rpr is not None:
            nr.append(copy.deepcopy(rpr))
        if props is not None:
            nrpr = nr.find(f"{W}rPr")
            if nrpr is None:
                nrpr = etree.SubElement(nr, f"{W}rPr")
            apply_run_props(nrpr, props)
            if nrpr.find(f"{W}rtl") is None:
                etree.SubElement(nrpr, f"{W}rtl")
        nt = etree.SubElement(nr, f"{W}t")
        nt.text = chunk
        nt.set(XML_SPACE, "preserve")
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
        """Supported honorifics: phrases, bracketed abbreviations, and the glyph
        in each glyph set (unicode = Scheherazade New / Noto Naskh Arabic,
        jameel = Jameel Noori Nastaleeq)."""
        return [
            {
                "key": key,
                "phrases": phrases,
                "abbreviations_in_brackets": abbr,
                "glyphs": {k: f"{v} U+{ord(v):04X}" for k, v in sets.items()},
            }
            for key, phrases, abbr, sets in HONORIFICS
        ]

    def apply_honorifics(
        self,
        *,
        glyph_set: str = "unicode",
        font: str | None = None,
        fallback_font: str | None = None,
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
            glyph_set: "unicode" (official code points; Scheherazade New or
                Noto Naskh Arabic draw all of them) or "jameel" (Jameel Noori
                Nastaleeq's own artwork, incl. حفظه الله / حفظها الله /
                حفظهم الله and رحمة الله عليه …). A phrase the chosen set
                lacks uses the other set, in that set's font.
            font: Font for the chosen set (e.g. "Noto Naskh Arabic" instead of
                Scheherazade New; "Jameel Noori Nastaleeq Kasheeda").
            fallback_font: Font for glyphs taken from the other set.
            color: Glyph colour (RRGGBB), e.g. "520912". None = text colour.
            size_pt: Glyph size. None = same as the text.
            bold: Force bold on/off for glyphs. None = keep.
            include_abbreviations: Also replace bracketed abbreviations such as
                (ص), (صلعم), (رض), (رح), (ع).
            only: Limit to these keys, phrases or glyphs (see list_honorifics).
            restyle_existing: Give glyphs already in the text the colour /
                size / bold (their font is left alone, since the same code
                point means different phrases in different fonts).
            parts: "all" (body, footnotes, endnotes, headers, footers, comments)
                or "body".
        """
        if glyph_set not in GLYPH_SETS:
            raise ValueError(f"glyph_set must be one of {sorted(GLYPH_SETS)}")
        fonts = {glyph_set: font or GLYPH_SETS[glyph_set]["font"]}
        for other in GLYPH_SETS:
            if other != glyph_set:
                fonts[other] = fallback_font or GLYPH_SETS[other]["font"]
        matcher = build_matcher(include_abbreviations, only)
        base: dict = {}
        if color:
            base["color"] = color
        if size_pt is not None:
            base["size_pt"] = size_pt
        if bold is not None:
            base["bold"] = bold
        stats: dict = {"by_key": {}, "fonts": {}, "no_glyph": {}}
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
        total = 0
        restyled = 0
        for part, root in list(self._trees.items()):
            if root is None or not part.startswith(prefixes) or not part.endswith(".xml"):
                continue
            n = 0
            for p in root.iter(f"{W}p"):
                for r in _text_runs(p):
                    n += replace_in_run(
                        r, matcher, lambda k: resolve_glyph(k, glyph_set, fonts), base, stats
                    )
            m = restyle_glyph_runs(root, base) if (restyle_existing and base) else 0
            if n or m:
                self._mark(part)
            total += n
            restyled += m
        return {
            "replaced": total,
            "glyph_set": glyph_set,
            "by_honorific": stats["by_key"],
            "fonts_used": stats["fonts"],
            "left_as_text_no_glyph": stats["no_glyph"],
            "existing_glyph_runs_restyled": restyled,
        }
