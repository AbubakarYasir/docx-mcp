"""Construction checks for the generated Islamic ornament."""

from __future__ import annotations

from collections import Counter

import pytest
from lxml import etree

from docx_mcp.document import islamic_art as ia


def _degrees(lines):
    deg = Counter()
    for g in lines.geoms:
        cs = list(g.coords)
        deg[ia._key(cs[0])] += 1
        deg[ia._key(cs[-1])] += 1
    return Counter(deg.values())


@pytest.mark.parametrize("n", [8, 10, 12, 16])
def test_rosette_network_crosses_at_every_tip(n):
    lines = ia.rosette_lines(n, 40, 50, 50)
    assert _degrees(lines)[4] == n  # one woven crossing per star tip


def test_weave_alternates_on_interlocking_rings():
    lines = ia.link_chain_lines(0, 52, 0, 15, sides=0, pitch=1.45)
    import inspect

    src = inspect.getsource(ia.weave).replace(
        "    return unary_union(pieces), len(occ)", "    return strands, occ, base"
    )
    ns: dict = {}
    exec(src, ia.__dict__, ns)
    strands, occ, base = ns["weave"](lines, 2.8, 0.75)
    states = {k: [(base[s] ^ (p & 1)) for s, p, _ in v] for k, v in occ.items()}
    assert len(states) == 2
    for pair in states.values():
        assert sorted(pair) == [0, 1]  # one over, one under
    top, bottom = sorted(states, key=lambda k: k[1])
    assert states[top] != states[bottom]  # the same ring alternates


def test_tilings_are_clean():
    t = ia.tiling_488(10, (-40, -40, 40, 40))
    assert set(len(p.exterior.coords) - 1 for p in t) == {8, 4}
    t = ia.tiling_31212(6, (-40, -40, 40, 40))
    assert set(len(p.exterior.coords) - 1 for p in t) == {12, 3}


@pytest.mark.parametrize(
    "fn",
    [
        lambda: ia.star_medallion(number="١٢"),
        lambda: ia.rosette(fold=12),
        lambda: ia.rosette(fold=8),
        lambda: ia.shamsa(),
        lambda: ia.ring(),
        lambda: ia.swash_divider(),
        lambda: ia.dots_divider(),
        lambda: ia.corner(),
        lambda: ia.band(kind="rings"),
        lambda: ia.band(kind="chain"),
        lambda: ia.band(kind="khatam"),
        lambda: ia.band(kind="8"),
        lambda: ia.cartouche("lobed"),
        lambda: ia.cartouche("banner", finials=True, knob=False),
        lambda: ia.cartouche("arch", knob=False),
        lambda: ia.frame(aspect=0.7),
    ],
)
def test_motifs_are_valid_svg(fn):
    root = etree.fromstring(fn().encode())
    assert root.tag.endswith("svg")
    assert len(root) > 0


def test_cartouche_rejects_unknown():
    with pytest.raises(ValueError):
        ia.cartouche_shape("cloud", 10, 4)
