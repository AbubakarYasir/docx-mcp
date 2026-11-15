"""Put WordprocessingML property children back into schema order.

ECMA-376 defines most property containers (``w:pPr``, ``w:rPr``, ``w:sectPr``,
``w:tblPr``, ``w:tcPr``, ``w:style``) as ``xsd:sequence``: the children must
appear in a fixed order. Many editing operations append a new child with
``SubElement``, which puts it at the end. LibreOffice tolerates that, but
Microsoft Word validates the order and reports "unreadable content" (or
silently drops formatting) when it is wrong.

Rather than teaching every one of the ~100 call sites the right insertion
point, :func:`normalize_tree` is run on every XML part before it is written.
It is a stable sort: children of the same kind keep their relative order, and
children this module does not know (extension elements such as ``w14:*``)
keep their position relative to the known element before them.
"""

from __future__ import annotations

from lxml import etree

_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _q(names: str) -> list[str]:
    return [f"{{{_W}}}{n}" for n in names.split()]


# CT_PPr = CT_PPrBase + rPr, sectPr, pPrChange
_PPR = _q(
    "pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr "
    "suppressLineNumbers pBdr shd tabs suppressAutoHyphens kinsoku wordWrap "
    "overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd "
    "snapToGrid spacing ind contextualSpacing mirrorIndents suppressOverlap jc "
    "textDirection textAlignment textboxTightWrap outlineLvl divId cnfStyle "
    "rPr sectPr pPrChange"
)

# CT_ParaRPr (paragraph-mark run properties) and CT_RPr share the base order.
_RPR = _q(
    "ins del moveFrom moveTo "
    "rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow "
    "emboss imprint noProof snapToGrid vanish webHidden color spacing w kern "
    "position sz szCs highlight u effect bdr shd fitText vertAlign rtl cs em "
    "lang eastAsianLayout specVanish oMath rPrChange"
)

_SECTPR = _q(
    "headerReference footerReference footnotePr endnotePr type pgSz pgMar "
    "paperSrc pgBorders lnNumType pgNumType cols formProt vAlign noEndnote "
    "titlePg textDirection bidi rtlGutter docGrid printerSettings sectPrChange"
)
# Header and footer references form one unordered group at the front.
_SECTPR_SAME_RANK = {f"{{{_W}}}footerReference": f"{{{_W}}}headerReference"}

_TBLPR = _q(
    "tblStyle tblpPr tblOverlap bidiVisual tblStyleRowBandSize "
    "tblStyleColBandSize tblW jc tblCellSpacing tblInd tblBorders shd "
    "tblLayout tblCellMar tblLook tblCaption tblDescription tblPrChange"
)

_TCPR = _q(
    "cnfStyle tcW gridSpan hMerge vMerge tcBorders shd noWrap tcMar "
    "textDirection tcFitText vAlign hideMark headers cellIns cellDel "
    "cellMerge tcPrChange"
)

_STYLE = _q(
    "name aliases basedOn next link autoRedefine hidden uiPriority semiHidden "
    "unhideWhenUsed qFormat locked personal personalCompose personalReply rsid "
    "pPr rPr tblPr trPr tcPr tblStylePr"
)

# Borders inside pBdr / tblBorders / tcBorders.
_PBDR = _q("top left bottom right between bar")
_TBLBORDERS = _q("top left start bottom right end insideH insideV")
_TCBORDERS = _q("top left start bottom right end insideH insideV tl2br tr2bl")

# numbering.xml root: all abstractNum definitions before any num instance.
_NUMBERING = _q("numPicBullet abstractNum num numIdMacAtCleanup")

# CT_TrPr: the base properties may come in any order, then ins, del, trPrChange.
_TRPR_BASE = (
    "cnfStyle divId gridBefore gridAfter wBefore wAfter cantSplit trHeight "
    "tblHeader tblCellSpacing jc hidden"
)
_TRPR = _q("_base ins del trPrChange")
_TRPR_SAME_RANK = {t: f"{{{_W}}}_base" for t in _q(_TRPR_BASE)}

# CT_SdtPr: fixed lead sequence, then one type element (all of equal rank).
_SDT_TYPES = (
    "equation comboBox date docPartObj docPartList dropDownList picture "
    "richText text citation group bibliography"
)
_SDTPR = _q(
    "rPr alias tag id lock placeholder temporary showingPlcHdr dataBinding label tabIndex _type"
)
_SDTPR_SAME_RANK = {t: f"{{{_W}}}_type" for t in _q(_SDT_TYPES)}

_ORDERS: dict[str, tuple[list[str], dict[str, str]]] = {
    f"{{{_W}}}numbering": (_NUMBERING, {}),
    f"{{{_W}}}trPr": (_TRPR, _TRPR_SAME_RANK),
    f"{{{_W}}}sdtPr": (_SDTPR, _SDTPR_SAME_RANK),
    f"{{{_W}}}pPr": (_PPR, {}),
    f"{{{_W}}}rPr": (_RPR, {}),
    f"{{{_W}}}sectPr": (_SECTPR, _SECTPR_SAME_RANK),
    f"{{{_W}}}tblPr": (_TBLPR, {}),
    f"{{{_W}}}tcPr": (_TCPR, {}),
    f"{{{_W}}}style": (_STYLE, {}),
    f"{{{_W}}}pBdr": (_PBDR, {}),
    f"{{{_W}}}tblBorders": (_TBLBORDERS, {}),
    f"{{{_W}}}tcBorders": (_TCBORDERS, {}),
}

_RANKS: dict[str, dict[str, int]] = {}
for _tag, (_seq, _same) in _ORDERS.items():
    ranks = {name: i for i, name in enumerate(_seq)}
    for alias, target in _same.items():
        ranks[alias] = ranks[target]
    _RANKS[_tag] = ranks

# Elements whose properties child must come first.
_FIRST_CHILD = {
    f"{{{_W}}}p": f"{{{_W}}}pPr",
    f"{{{_W}}}r": f"{{{_W}}}rPr",
    f"{{{_W}}}tbl": f"{{{_W}}}tblPr",
    f"{{{_W}}}tc": f"{{{_W}}}tcPr",
}
# w:tr allows tblPrEx before trPr; both must precede the cells.
_TR = f"{{{_W}}}tr"
_TR_LEAD = (f"{{{_W}}}tblPrEx", f"{{{_W}}}trPr")


def _sort_children(parent: etree._Element, ranks: dict[str, int]) -> bool:
    children = list(parent)
    if len(children) < 2:
        return False
    keyed = []
    last_rank = -1
    for idx, child in enumerate(children):
        tag = child.tag if isinstance(child.tag, str) else None
        rank = ranks.get(tag) if tag else None
        if rank is None:
            # Unknown element or comment: keep it after the known element before it.
            rank = last_rank
        else:
            last_rank = rank
        keyed.append((rank, idx, child))
    ordered = [c for _, _, c in sorted(keyed, key=lambda t: (t[0], t[1]))]
    if ordered == children:
        return False
    for child in children:
        parent.remove(child)
    parent.extend(ordered)
    return True


def _move_first(parent: etree._Element, tags: tuple[str, ...]) -> bool:
    lead = [c for c in parent if c.tag in tags]
    if not lead:
        return False
    lead.sort(key=lambda c: tags.index(c.tag))
    if list(parent)[: len(lead)] == lead:
        return False
    for c in lead:
        parent.remove(c)
    for i, c in enumerate(lead):
        parent.insert(i, c)
    return True


def normalize_tree(root: etree._Element) -> int:
    """Reorder property children in ``root`` in place. Returns elements fixed."""
    fixed = 0
    for el in root.iter():
        tag = el.tag
        if not isinstance(tag, str):
            continue
        ranks = _RANKS.get(tag)
        if ranks is not None and _sort_children(el, ranks):
            fixed += 1
        first = _FIRST_CHILD.get(tag)
        lead = (first,) if first is not None else (_TR_LEAD if tag == _TR else None)
        if lead is not None and _move_first(el, lead):
            fixed += 1
    return fixed
