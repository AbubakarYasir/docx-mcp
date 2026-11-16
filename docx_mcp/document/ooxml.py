"""Shared WordprocessingML builders used by the design tools.

Run properties, paragraph properties, fields and new package parts are built
in one place so every tool produces the same, schema-valid markup. Element
order inside property containers is fixed on save by ``schema_order``, so the
builders here only need to produce the right elements.
"""

from __future__ import annotations

import re
from typing import Any

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
R = f"{{{R_NS}}}"
RELS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
CT = "{http://schemas.openxmlformats.org/package/2006/content-types}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

REL_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
CT_BASE = "application/vnd.openxmlformats-officedocument.wordprocessingml."

HEX_RE = re.compile(r"^[0-9A-Fa-f]{6}$")

# Paragraph alignment values. Arabic kashida justification is a Word feature
# few people know: it stretches letters instead of widening spaces.
ALIGNMENTS = {
    "left": "left",
    "start": "start",
    "right": "right",
    "end": "end",
    "center": "center",
    "centre": "center",
    "justify": "both",
    "both": "both",
    "distribute": "distribute",
    # Word drops whole lines of vocalised Arabic (Amiri) under medium/high
    # kashida justification — checked in Word 365. Low kashida is safe.
    "kashida": "lowKashida",
    "kashida-low": "lowKashida",
    "kashida-medium": "mediumKashida",
    "kashida-high": "highKashida",
    "thai": "thaiDistribute",
}

UNDERLINES = {
    "single",
    "double",
    "thick",
    "dotted",
    "dash",
    "dotDash",
    "dotDotDash",
    "wave",
    "words",
    "none",
}


def color_hex(value: str | None) -> str | None:
    """Accept 'RRGGBB', '#RRGGBB' or 'auto'. Returns the value Word expects."""
    if value is None or value == "":
        return None
    v = value.strip().lstrip("#")
    if v.lower() == "auto":
        return "auto"
    if not HEX_RE.match(v):
        raise ValueError(f"Colour must be RRGGBB hex (got {value!r})")
    return v.upper()


def half_points(size_pt: float) -> str:
    return str(round(float(size_pt) * 2))


def twips(value_pt: float) -> str:
    return str(round(float(value_pt) * 20))


def cm_to_twips(cm: float) -> int:
    return round(float(cm) * 1440 / 2.54)


def _get(parent: etree._Element, tag: str) -> etree._Element:
    el = parent.find(f"{W}{tag}")
    if el is None:
        el = etree.SubElement(parent, f"{W}{tag}")
    return el


def _toggle(rpr: etree._Element, tag: str, on: bool | None) -> None:
    if on is None:
        return
    el = _get(rpr, tag)
    if on:
        el.attrib.pop(f"{W}val", None)
    else:
        el.set(f"{W}val", "0")


# ── Run properties ──────────────────────────────────────────────────────────

RUN_KEYS = {
    "font",
    "latin_font",
    "cs_font",
    "east_asia_font",
    "size_pt",
    "cs_size_pt",
    "bold",
    "italic",
    "underline",
    "strike",
    "color",
    "highlight",
    "shading",
    "caps",
    "small_caps",
    "superscript",
    "subscript",
    "spacing_pt",
    "position_pt",
    "char_style",
    "rtl",
    "language",
    "cs_language",
    "hidden",
}


def apply_run_props(rpr: etree._Element, props: dict[str, Any]) -> None:
    """Write run properties into an existing w:rPr.

    ``font`` sets the font for every script (Latin, Arabic/Hebrew and East
    Asian); ``latin_font`` / ``cs_font`` set one script only. Size, bold and
    italic are applied to both Latin and complex-script letters so Arabic text
    follows the same formatting.
    """
    unknown = set(props) - RUN_KEYS
    if unknown:
        raise ValueError(f"Unknown run properties: {sorted(unknown)}")
    if props.get("char_style"):
        _get(rpr, "rStyle").set(f"{W}val", str(props["char_style"]))
    font = props.get("font")
    latin = props.get("latin_font") or font
    cs = props.get("cs_font") or font
    ea = props.get("east_asia_font") or font
    if latin or cs or ea:
        fonts = _get(rpr, "rFonts")
        for theme_attr in ("asciiTheme", "hAnsiTheme", "cstheme", "eastAsiaTheme"):
            if (theme_attr.startswith(("ascii", "hAnsi")) and latin) or (
                theme_attr == "cstheme" and cs
            ):
                fonts.attrib.pop(f"{W}{theme_attr}", None)
        if latin:
            fonts.set(f"{W}ascii", latin)
            fonts.set(f"{W}hAnsi", latin)
        if cs:
            fonts.set(f"{W}cs", cs)
        if ea:
            fonts.set(f"{W}eastAsia", ea)
    if props.get("bold") is not None:
        _toggle(rpr, "b", props["bold"])
        _toggle(rpr, "bCs", props["bold"])
    if props.get("italic") is not None:
        _toggle(rpr, "i", props["italic"])
        _toggle(rpr, "iCs", props["italic"])
    _toggle(rpr, "caps", props.get("caps"))
    _toggle(rpr, "smallCaps", props.get("small_caps"))
    _toggle(rpr, "strike", props.get("strike"))
    _toggle(rpr, "vanish", props.get("hidden"))
    color = color_hex(props.get("color"))
    if color:
        _get(rpr, "color").set(f"{W}val", color)
    if props.get("spacing_pt") is not None:
        _get(rpr, "spacing").set(f"{W}val", twips(props["spacing_pt"]))
    if props.get("position_pt") is not None:
        _get(rpr, "position").set(f"{W}val", half_points(props["position_pt"]))
    size = props.get("size_pt")
    cs_size = props.get("cs_size_pt", size)
    if size is not None:
        _get(rpr, "sz").set(f"{W}val", half_points(size))
    if cs_size is not None:
        _get(rpr, "szCs").set(f"{W}val", half_points(cs_size))
    if props.get("highlight"):
        _get(rpr, "highlight").set(f"{W}val", str(props["highlight"]))
    ul = props.get("underline")
    if ul:
        ul = "single" if ul is True else str(ul)
        if ul not in UNDERLINES:
            raise ValueError(f"underline must be one of {sorted(UNDERLINES)}")
        _get(rpr, "u").set(f"{W}val", ul)
    shading = color_hex(props.get("shading"))
    if shading:
        shd = _get(rpr, "shd")
        shd.set(f"{W}val", "clear")
        shd.set(f"{W}color", "auto")
        shd.set(f"{W}fill", shading)
    if props.get("superscript"):
        _get(rpr, "vertAlign").set(f"{W}val", "superscript")
    elif props.get("subscript"):
        _get(rpr, "vertAlign").set(f"{W}val", "subscript")
    if props.get("rtl") is not None:
        _toggle(rpr, "rtl", props["rtl"])
    if props.get("language") or props.get("cs_language"):
        lang = _get(rpr, "lang")
        if props.get("language"):
            lang.set(f"{W}val", str(props["language"]))
        if props.get("cs_language"):
            lang.set(f"{W}bidi", str(props["cs_language"]))


def make_run(text: str = "", props: dict[str, Any] | None = None) -> etree._Element:
    """A w:r with optional properties and text (tabs and line breaks honoured)."""
    r = etree.Element(f"{W}r")
    if props:
        apply_run_props(etree.SubElement(r, f"{W}rPr"), props)
    parts = re.split(r"(\t|\n)", text)
    for part in parts:
        if part == "":
            continue
        if part == "\t":
            etree.SubElement(r, f"{W}tab")
        elif part == "\n":
            etree.SubElement(r, f"{W}br")
        else:
            t = etree.SubElement(r, f"{W}t")
            t.text = part
            t.set(XML_SPACE, "preserve")
    return r


# ── Paragraph properties ────────────────────────────────────────────────────

PARA_KEYS = {
    "style",
    "alignment",
    "direction",
    "space_before_pt",
    "space_after_pt",
    "line_spacing",
    "line_exact_pt",
    "line_at_least_pt",
    "indent_start_cm",
    "indent_end_cm",
    "first_line_cm",
    "hanging_cm",
    "keep_with_next",
    "keep_together",
    "page_break_before",
    "widow_control",
    "outline_level",
    "contextual_spacing",
    "border_top",
    "border_bottom",
    "border_start",
    "border_end",
    "shading",
    "tabs",
    "suppress_hyphenation",
}


def _border(parent: etree._Element, side: str, spec: Any) -> None:
    """spec: True | 'RRGGBB' | {'style','color','size_pt','space_pt'}."""
    if not spec:
        return
    if spec is True:
        spec = {}
    elif isinstance(spec, str):
        spec = {"color": spec}
    el = _get(parent, side)
    el.set(f"{W}val", spec.get("style", "single"))
    el.set(f"{W}sz", str(round(float(spec.get("size_pt", 0.75)) * 8)))
    el.set(f"{W}space", str(round(float(spec.get("space_pt", 1)))))
    el.set(f"{W}color", color_hex(spec.get("color", "auto")) or "auto")


def apply_para_props(ppr: etree._Element, props: dict[str, Any]) -> None:
    """Write paragraph properties into an existing w:pPr."""
    unknown = set(props) - PARA_KEYS
    if unknown:
        raise ValueError(f"Unknown paragraph properties: {sorted(unknown)}")
    if props.get("style"):
        _get(ppr, "pStyle").set(f"{W}val", str(props["style"]))
    for key, tag in (
        ("keep_with_next", "keepNext"),
        ("keep_together", "keepLines"),
        ("page_break_before", "pageBreakBefore"),
        ("widow_control", "widowControl"),
        ("contextual_spacing", "contextualSpacing"),
        ("suppress_hyphenation", "suppressAutoHyphens"),
    ):
        _toggle(ppr, tag, props.get(key))
    if props.get("direction"):
        d = str(props["direction"]).lower()
        if d not in ("rtl", "ltr"):
            raise ValueError("direction must be 'rtl' or 'ltr'")
        _toggle(ppr, "bidi", d == "rtl")
    if props.get("alignment"):
        a = str(props["alignment"]).lower()
        if a not in ALIGNMENTS:
            raise ValueError(f"alignment must be one of {sorted(ALIGNMENTS)}")
        _get(ppr, "jc").set(f"{W}val", ALIGNMENTS[a])
    spacing_keys = (
        "space_before_pt",
        "space_after_pt",
        "line_spacing",
        "line_exact_pt",
        "line_at_least_pt",
    )
    if any(props.get(k) is not None for k in spacing_keys):
        sp = _get(ppr, "spacing")
        if props.get("space_before_pt") is not None:
            sp.set(f"{W}before", twips(props["space_before_pt"]))
        if props.get("space_after_pt") is not None:
            sp.set(f"{W}after", twips(props["space_after_pt"]))
        if props.get("line_exact_pt") is not None:
            sp.set(f"{W}line", twips(props["line_exact_pt"]))
            sp.set(f"{W}lineRule", "exact")
        elif props.get("line_at_least_pt") is not None:
            sp.set(f"{W}line", twips(props["line_at_least_pt"]))
            sp.set(f"{W}lineRule", "atLeast")
        elif props.get("line_spacing") is not None:
            sp.set(f"{W}line", str(round(float(props["line_spacing"]) * 240)))
            sp.set(f"{W}lineRule", "auto")
    indent_keys = ("indent_start_cm", "indent_end_cm", "first_line_cm", "hanging_cm")
    if any(props.get(k) is not None for k in indent_keys):
        ind = _get(ppr, "ind")
        if props.get("indent_start_cm") is not None:
            ind.set(f"{W}start", str(cm_to_twips(props["indent_start_cm"])))
        if props.get("indent_end_cm") is not None:
            ind.set(f"{W}end", str(cm_to_twips(props["indent_end_cm"])))
        if props.get("first_line_cm") is not None:
            ind.set(f"{W}firstLine", str(cm_to_twips(props["first_line_cm"])))
            ind.attrib.pop(f"{W}hanging", None)
        if props.get("hanging_cm") is not None:
            ind.set(f"{W}hanging", str(cm_to_twips(props["hanging_cm"])))
            ind.attrib.pop(f"{W}firstLine", None)
    if props.get("outline_level") is not None:
        _get(ppr, "outlineLvl").set(f"{W}val", str(int(props["outline_level"])))
    sides = {
        "border_top": "top",
        "border_bottom": "bottom",
        "border_start": "left",
        "border_end": "right",
    }
    if any(props.get(k) for k in sides):
        pbdr = _get(ppr, "pBdr")
        for key, side in sides.items():
            _border(pbdr, side, props.get(key))
    shading = color_hex(props.get("shading"))
    if shading:
        shd = _get(ppr, "shd")
        shd.set(f"{W}val", "clear")
        shd.set(f"{W}color", "auto")
        shd.set(f"{W}fill", shading)
    if props.get("tabs"):
        tabs = _get(ppr, "tabs")
        for c in list(tabs):
            tabs.remove(c)
        for tab in props["tabs"]:
            tab_el = etree.SubElement(tabs, f"{W}tab")
            tab_el.set(f"{W}val", tab.get("align", "left"))
            tab_el.set(f"{W}pos", str(cm_to_twips(tab["pos_cm"])))
            if tab.get("leader"):
                tab_el.set(f"{W}leader", tab["leader"])


def make_paragraph(
    para_id: str, props: dict[str, Any] | None = None, runs: list | None = None
) -> etree._Element:
    p = etree.Element(f"{W}p")
    p.set(f"{W14}paraId", para_id)
    p.set(f"{W14}textId", "77777777")
    if props:
        apply_para_props(etree.SubElement(p, f"{W}pPr"), props)
    for r in runs or []:
        p.append(r)
    return p


# ── Fields ──────────────────────────────────────────────────────────────────


def field_runs(
    instr: str, result: str = "", props: dict[str, Any] | None = None, dirty: bool = True
) -> list[etree._Element]:
    """A complex field as runs: begin, instruction, separate, result, end.

    ``dirty`` asks Word to recalculate the field when the document opens,
    so page numbers, STYLEREF running heads and indexes are always current.
    """
    runs = []
    begin = make_run("", props)
    fc = etree.SubElement(begin, f"{W}fldChar")
    fc.set(f"{W}fldCharType", "begin")
    if dirty:
        fc.set(f"{W}dirty", "true")
    runs.append(begin)
    instr_run = make_run("", props)
    it = etree.SubElement(instr_run, f"{W}instrText")
    it.text = f" {instr.strip()} "
    it.set(XML_SPACE, "preserve")
    runs.append(instr_run)
    sep = make_run("", props)
    etree.SubElement(sep, f"{W}fldChar").set(f"{W}fldCharType", "separate")
    runs.append(sep)
    runs.append(make_run(result or "1", props))
    end = make_run("", props)
    etree.SubElement(end, f"{W}fldChar").set(f"{W}fldCharType", "end")
    runs.append(end)
    return runs


# Tokens accepted by header/footer and rich-text content.
FIELD_TOKENS = {
    "PAGE": ("PAGE", "1"),
    "NUMPAGES": ("NUMPAGES", "1"),
    "SECTIONPAGES": ("SECTIONPAGES", "1"),
    "SECTION": ("SECTION", "1"),
    "DATE": ('DATE \\@ "d MMMM yyyy"', ""),
    "TITLE": ("TITLE", ""),
    "AUTHOR": ("AUTHOR", ""),
    "FILENAME": ("FILENAME", ""),
}
TOKEN_RE = re.compile(
    r"\{(PAGE|NUMPAGES|SECTIONPAGES|SECTION|DATE|TITLE|AUTHOR|FILENAME|"
    r"STYLEREF(?::[^}]+)?|FIELD:[^}]+)\}"
)


def content_runs(text: str, props: dict[str, Any] | None = None) -> list[etree._Element]:
    """Text with {PAGE}, {NUMPAGES}, {STYLEREF:Heading 1}, {FIELD:...} tokens -> runs."""
    out: list[etree._Element] = []
    pos = 0
    for m in TOKEN_RE.finditer(text):
        if m.start() > pos:
            out.append(make_run(text[pos : m.start()], props))
        tok = m.group(1)
        if tok.startswith("STYLEREF"):
            style = tok.split(":", 1)[1] if ":" in tok else "Heading 1"
            out.extend(field_runs(f'STYLEREF "{style}"', " ", props))
        elif tok.startswith("FIELD:"):
            out.extend(field_runs(tok.split(":", 1)[1], " ", props))
        else:
            instr, res = FIELD_TOKENS[tok]
            out.extend(field_runs(instr, res, props))
        pos = m.end()
    if pos < len(text):
        out.append(make_run(text[pos:], props))
    return out


# ── Package parts ───────────────────────────────────────────────────────────


def next_rid(rels: etree._Element) -> str:
    nums = []
    for rel in rels.findall(f"{RELS}Relationship"):
        rid = rel.get("Id", "")
        if rid.startswith("rId") and rid[3:].isdigit():
            nums.append(int(rid[3:]))
    return f"rId{max(nums, default=0) + 1}"


def add_relationship(
    rels: etree._Element, rel_type: str, target: str, external: bool = False
) -> str:
    rid = next_rid(rels)
    rel = etree.SubElement(rels, f"{RELS}Relationship")
    rel.set("Id", rid)
    rel.set("Type", rel_type if rel_type.startswith("http") else REL_BASE + rel_type)
    rel.set("Target", target)
    if external:
        rel.set("TargetMode", "External")
    return rid


def ensure_override(ct_root: etree._Element, part_name: str, content_type: str) -> bool:
    for ov in ct_root.findall(f"{CT}Override"):
        if ov.get("PartName") == part_name:
            return False
    ov = etree.SubElement(ct_root, f"{CT}Override")
    ov.set("PartName", part_name)
    ov.set("ContentType", content_type)
    return True


def ensure_default(ct_root: etree._Element, ext: str, content_type: str) -> bool:
    for d in ct_root.findall(f"{CT}Default"):
        if (d.get("Extension") or "").lower() == ext.lower():
            return False
    d = etree.SubElement(ct_root, f"{CT}Default")
    d.set("Extension", ext)
    d.set("ContentType", content_type)
    return True


def unique_part_name(existing: set[str], prefix: str, ext: str = "xml") -> str:
    i = 1
    while f"{prefix}{i}.{ext}" in existing:
        i += 1
    return f"{prefix}{i}.{ext}"
