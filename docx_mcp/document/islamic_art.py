"""Production-grade Islamic ornament, generated as clean vector geometry.

Construction follows the traditional methods rather than clip-art shortcuts:

- Girih (geometric strapwork) is built with Hankin's polygons-in-contact
  method on a regular tiling (4.8.8 for eight-fold stars, 3.12.12 for
  twelve-fold rosettes, 6.6.6 for six-fold), clipped to the motif's shape.
- Strapwork is woven: each strand alternates over / under at successive
  crossings (solved as a parity 2-colouring), and the under strand is cut
  back so the over strand reads on top.
- Outlines use true parallel offsets (shapely buffers), so double and triple
  rules keep an even gap round every point of a star.
- Fleurons, swashes and cartouche ends are cubic Béziers, mirrored for
  symmetry and merged into single filled shapes.

Every function returns SVG markup (a viewBox in abstract units); the caller
scales it. Shapes are filled paths (fill-rule evenodd), which Word, PDF and
print RIPs render identically.
"""

from __future__ import annotations

import math
from collections.abc import Iterable, Sequence

from shapely import affinity
from shapely.geometry import (
    GeometryCollection,
    LineString,
    MultiLineString,
    MultiPolygon,
    Point,
    Polygon,
    box,
)
from shapely.ops import unary_union

Pt = tuple[float, float]

# ── SVG output ──────────────────────────────────────────────────────────


def _ring_d(coords: Iterable[Sequence[float]]) -> str:
    pts = list(coords)
    if len(pts) > 1 and pts[0] == pts[-1]:
        pts = pts[:-1]
    if len(pts) < 3:
        return ""
    head = f"M{pts[0][0]:.2f},{pts[0][1]:.2f}"
    body = "".join(f"L{x:.2f},{y:.2f}" for x, y in pts[1:])
    return head + body + "Z"


def path_d(geom) -> str:
    """SVG path data for a (Multi)Polygon, holes included."""
    if geom is None or geom.is_empty:
        return ""
    if isinstance(geom, Polygon):
        parts = [_ring_d(geom.exterior.coords)]
        parts += [_ring_d(i.coords) for i in geom.interiors]
        return "".join(parts)
    if isinstance(geom, (MultiPolygon, GeometryCollection)):
        return "".join(path_d(g) for g in geom.geoms if isinstance(g, (Polygon, MultiPolygon)))
    return ""


def fill(geom, color: str, opacity: float | None = None) -> str:
    d = path_d(geom)
    if not d:
        return ""
    op = f' fill-opacity="{opacity}"' if opacity is not None else ""
    return f'<path d="{d}" fill="{color}" fill-rule="evenodd"{op}/>'


def svg(w: float, h: float, body: str) -> str:
    return (
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{w:.0f}" height="{h:.0f}" '
        f'viewBox="0 0 {w:.2f} {h:.2f}">{body}</svg>'
    )


def hexc(c: str | None, default: str = "#000000") -> str:
    if not c:
        return default
    c = c.strip()
    return c if c.startswith("#") else "#" + c


def tint(color: str, amount: float) -> str:
    h = hexc(color).lstrip("#")
    r, g, b = (int(h[i : i + 2], 16) for i in (0, 2, 4))
    r, g, b = (round(v + (255 - v) * amount) for v in (r, g, b))
    return f"#{r:02x}{g:02x}{b:02x}"


# ── basic geometry ──────────────────────────────────────────────────────


def regular(n: int, r: float, cx: float = 0, cy: float = 0, rot: float = 0) -> Polygon:
    return Polygon(
        [
            (
                cx + r * math.cos(math.radians(rot + 360 * i / n)),
                cy + r * math.sin(math.radians(rot + 360 * i / n)),
            )
            for i in range(n)
        ]
    )


def star(
    n: int, r_out: float, r_in: float, cx: float = 0, cy: float = 0, rot: float = -90
) -> Polygon:
    pts = []
    for i in range(2 * n):
        r = r_out if i % 2 == 0 else r_in
        a = math.radians(rot + i * 180 / n)
        pts.append((cx + r * math.cos(a), cy + r * math.sin(a)))
    return Polygon(pts)


def khatam(r: float, cx: float = 0, cy: float = 0, rot: float = 0) -> Polygon:
    """Eight-point star of two interlaced squares (outer radius r)."""
    r_in = r * math.cos(math.radians(45)) / math.cos(math.radians(22.5))
    return star(8, r, r_in, cx, cy, rot=-90 + rot)


def circle(r: float, cx: float = 0, cy: float = 0, res: int = 64) -> Polygon:
    return Point(cx, cy).buffer(r, quad_segs=res)


def outline(shape, width: float):
    """A rule of `width` drawn on the inside of shape's edge."""
    return shape.difference(shape.buffer(-width, join_style="mitre", mitre_limit=8))


def cubic(p0: Pt, p1: Pt, p2: Pt, p3: Pt, n: int = 24) -> list[Pt]:
    out = []
    for i in range(n + 1):
        t = i / n
        u = 1 - t
        out.append(
            (
                u**3 * p0[0] + 3 * u * u * t * p1[0] + 3 * u * t * t * p2[0] + t**3 * p3[0],
                u**3 * p0[1] + 3 * u * u * t * p1[1] + 3 * u * t * t * p2[1] + t**3 * p3[1],
            )
        )
    return out


def bezier_shape(segments: list[tuple[Pt, Pt, Pt, Pt]], n: int = 20) -> Polygon:
    pts: list[Pt] = []
    for seg in segments:
        pts.extend(cubic(*seg, n=n)[:-1])
    poly = Polygon(pts)
    return poly if poly.is_valid else poly.buffer(0)


def mirror_x(geom, x0: float = 0):
    return affinity.scale(geom, xfact=-1, yfact=1, origin=(x0, 0))


def mirror_y(geom, y0: float = 0):
    return affinity.scale(geom, xfact=1, yfact=-1, origin=(0, y0))


def symmetric_x(geom, x0: float = 0):
    return unary_union([geom, mirror_x(geom, x0)])


def arc_stroke(
    cx: float,
    cy: float,
    r: float,
    a0: float,
    a1: float,
    width: float,
    taper: bool = True,
    ball: float = 0.0,
):
    """A calligraphic curl: an arc stroke that thins toward its tip, with an
    optional ball terminal."""
    n = 40
    pieces = []
    for i in range(n):
        t0, t1 = i / n, (i + 1) / n
        a = math.radians(a0 + (a1 - a0) * t0)
        b = math.radians(a0 + (a1 - a0) * t1)
        w = width * ((1 - 0.75 * t0) if taper else 1)
        seg = LineString(
            [
                (cx + r * math.cos(a), cy + r * math.sin(a)),
                (cx + r * math.cos(b), cy + r * math.sin(b)),
            ]
        )
        pieces.append(seg.buffer(w / 2, cap_style="round"))
    shape = unary_union(pieces)
    if ball:
        a = math.radians(a1)
        shape = shape.union(circle(ball, cx + r * math.cos(a), cy + r * math.sin(a), 16))
    return shape


# ── tilings and Hankin girih ────────────────────────────────────────────


def _holes(region, tiles: list[Polygon], sides: int) -> list[Polygon]:
    """The gaps between tiles that are clean regular polygons of `sides`
    (boolean noding leaves in-line vertices, which would spawn stray rays)."""
    rest = region.difference(unary_union(tiles))
    geoms = rest.geoms if hasattr(rest, "geoms") else [rest]
    out = []
    for g in geoms:
        if not isinstance(g, Polygon) or g.area < 1e-6:
            continue
        g = g.simplify(1e-3)
        if len(g.exterior.coords) - 1 == sides:
            out.append(g)
    return out


def tiling_488(s: float, bounds: tuple[float, float, float, float]) -> list[Polygon]:
    """Octagons + squares (4.8.8), octagon side s, covering bounds."""
    a = s * (1 + math.sqrt(2))
    rc = s / (2 * math.sin(math.radians(22.5)))
    x0, y0, x1, y1 = bounds
    tiles = []
    i0, i1 = math.floor(x0 / a) - 1, math.ceil(x1 / a) + 1
    j0, j1 = math.floor(y0 / a) - 1, math.ceil(y1 / a) + 1
    for i in range(i0, i1 + 1):
        for j in range(j0, j1 + 1):
            tiles.append(regular(8, rc, i * a, j * a, rot=22.5))
    region = box(i0 * a, j0 * a, i1 * a, j1 * a)
    squares = [h for h in _holes(region, tiles, 4) if abs(h.area - s * s) < s * s * 0.05]
    return tiles + squares


def tiling_31212(s: float, bounds: tuple[float, float, float, float]) -> list[Polygon]:
    """Dodecagons + triangles (3.12.12), dodecagon side s."""
    ap = s / (2 * math.tan(math.radians(15)))
    rc = s / (2 * math.sin(math.radians(15)))
    dx = 2 * ap
    dy = dx * math.sqrt(3) / 2
    x0, y0, x1, y1 = bounds
    tiles = []
    j0, j1 = math.floor(y0 / dy) - 1, math.ceil(y1 / dy) + 1
    i0, i1 = math.floor(x0 / dx) - 2, math.ceil(x1 / dx) + 2
    for j in range(j0, j1 + 1):
        for i in range(i0, i1 + 1):
            cx = i * dx + (dx / 2 if j % 2 else 0)
            tiles.append(regular(12, rc, cx, j * dy, rot=15))
    region = box(i0 * dx, j0 * dy, i1 * dx, j1 * dy)
    tri_area = math.sqrt(3) / 4 * s * s
    tris = [h for h in _holes(region, tiles, 3) if abs(h.area - tri_area) < tri_area * 0.05]
    return tiles + tris


def tiling_hex(s: float, bounds: tuple[float, float, float, float]) -> list[Polygon]:
    ap = s * math.sqrt(3) / 2
    dx = 2 * ap
    dy = 1.5 * s
    x0, y0, x1, y1 = bounds
    tiles = []
    for j in range(math.floor(y0 / dy) - 1, math.ceil(y1 / dy) + 2):
        for i in range(math.floor(x0 / dx) - 1, math.ceil(x1 / dx) + 2):
            tiles.append(regular(6, s, i * dx + (ap if j % 2 else 0), j * dy, rot=30))
    return tiles


def hankin(tiles: Iterable[Polygon], theta: float) -> list[LineString]:
    """Polygons-in-contact: from each edge midpoint, two rays at `theta`
    degrees to the edge, meeting the rays of the neighbouring edges."""
    segs: list[LineString] = []
    th = math.radians(theta)
    for poly in tiles:
        pts = list(poly.exterior.coords)[:-1]
        # counter-clockwise, so the interior lies to the left of each edge
        area = sum(
            pts[i][0] * pts[(i + 1) % len(pts)][1] - pts[(i + 1) % len(pts)][0] * pts[i][1]
            for i in range(len(pts))
        )
        if area < 0:
            pts.reverse()
        n = len(pts)
        mids, fwd, back = [], [], []
        for i in range(n):
            ax, ay = pts[i]
            bx, by = pts[(i + 1) % n]
            ex, ey = bx - ax, by - ay
            ln = math.hypot(ex, ey)
            ex, ey = ex / ln, ey / ln
            mids.append(((ax + bx) / 2, (ay + by) / 2))
            # rotate edge direction by +theta (forward, inward) and the
            # reversed direction by -theta (backward, inward)
            fwd.append(
                (ex * math.cos(th) - ey * math.sin(th), ex * math.sin(th) + ey * math.cos(th))
            )
            bx2, by2 = -ex, -ey
            back.append(
                (
                    bx2 * math.cos(-th) - by2 * math.sin(-th),
                    bx2 * math.sin(-th) + by2 * math.cos(-th),
                )
            )
        for i in range(n):
            j = (i + 1) % n
            p, r = mids[i], fwd[i]
            q, s = mids[j], back[j]
            den = r[0] * s[1] - r[1] * s[0]
            if abs(den) < 1e-9:
                continue
            t = ((q[0] - p[0]) * s[1] - (q[1] - p[1]) * s[0]) / den
            x = (p[0] + r[0] * t, p[1] + r[1] * t)
            segs.append(LineString([p, x]))
            segs.append(LineString([x, q]))
    return segs


def girih_lines(kind: str, cell: float, bounds, theta: float | None = None):
    """The line network of a girih pattern (MultiLineString)."""
    if kind == "8":
        tiles = tiling_488(cell, bounds)
        th = theta or 67.5
    elif kind == "12":
        tiles = tiling_31212(cell, bounds)
        th = theta or 67.5
    elif kind == "6":
        tiles = tiling_hex(cell, bounds)
        th = theta or 60
    else:
        raise ValueError("girih kind must be 6, 8 or 12")
    return unary_union(hankin(tiles, th))


def khatam_chain_lines(
    x0: float, x1: float, y: float, r: float, vertical: bool = False, overlap: float = 1.72
) -> MultiLineString:
    """An interlocking row of khatams (two interlaced squares each) from x0
    to x1 on the line y: neighbours overlap so every junction is a woven
    crossing and every strand a closed loop — the classic border chain."""
    length = x1 - x0
    pitch_target = overlap * r
    count = max(1, round((length - 2 * r) / pitch_target) + 1)
    pitch = (length - 2 * r) / max(1, count - 1) if count > 1 else 0
    segs = []
    for k in range(count):
        cx = x0 + r + pitch * k
        for rot in (0, 45):
            cs = list(regular(4, r, cx, y, rot=rot).exterior.coords)
            segs.extend(LineString([a, b]) for a, b in zip(cs[:-1], cs[1:], strict=True))
    g = unary_union(segs)
    if vertical:
        g = affinity.rotate(g, 90, origin=(x0, y))
        g = affinity.translate(g, y - x0, x0 - y)
    return g


def _ring_centres(x0: float, x1: float, r: float, pitch: float) -> list[float]:
    length = x1 - x0
    count = max(1, round((length - 2 * r) / (pitch * r)) + 1)
    step = (length - 2 * r) / max(1, count - 1) if count > 1 else 0
    return [x0 + r + step * k for k in range(count)]


def link_chain_lines(
    x0: float, x1: float, y: float, r: float, sides: int = 8, pitch: float = 1.45
) -> MultiLineString:
    """Guilloche chain: interlocking rings (sides=0 → circles, 8 → octagons,
    6 → hexagons) from x0 to x1, each crossing only its neighbours."""
    length = x1 - x0
    target = pitch * r
    count = max(1, round((length - 2 * r) / target) + 1)
    step = (length - 2 * r) / max(1, count - 1) if count > 1 else 0
    segs = []
    for k in range(count):
        cx = x0 + r + step * k
        n = sides or 48
        rot = 180 / n if sides else 0
        cs = list(regular(n, r, cx, y, rot=rot).exterior.coords)
        segs.extend(LineString([a, b]) for a, b in zip(cs[:-1], cs[1:], strict=True))
    return unary_union(segs)


# ── weaving ─────────────────────────────────────────────────────────────


def _key(p: Sequence[float]) -> tuple[int, int]:
    return (round(p[0] * 100), round(p[1] * 100))


def weave(lines, width: float, gap: float):
    """Turn a line network into over/under strapwork.

    Returns (straps, crossings) where straps is the union of strap shapes
    with the under strand cut back at every crossing.
    """
    merged = unary_union(lines)
    segs = [merged] if isinstance(merged, LineString) else list(getattr(merged, "geoms", []))
    edges: list[tuple[Pt, Pt]] = []
    for g in segs:
        cs = list(g.coords)
        edges.extend(zip(cs[:-1], cs[1:], strict=True))
    adj: dict[tuple[int, int], list[int]] = {}
    for idx, (a, b) in enumerate(edges):
        adj.setdefault(_key(a), []).append(idx)
        adj.setdefault(_key(b), []).append(idx)

    def other(idx: int, k: tuple[int, int]) -> tuple[int, int]:
        a, b = edges[idx]
        return _key(b) if _key(a) == k else _key(a)

    def direction(idx: int, k: tuple[int, int]) -> tuple[float, float]:
        a, b = edges[idx]
        p, q = (a, b) if _key(a) == k else (b, a)
        dx, dy = q[0] - p[0], q[1] - p[1]
        ln = math.hypot(dx, dy) or 1
        return dx / ln, dy / ln

    def straight_on(idx: int, k: tuple[int, int]) -> int | None:
        """At a crossing node, the edge that continues idx straight on."""
        cands = [e for e in adj[k] if e != idx]
        if len(adj[k]) != 4:
            return cands[0] if len(cands) == 1 else None
        dx, dy = direction(idx, k)
        best, best_dot = None, 0.0
        for e in cands:
            ex, ey = direction(e, k)
            dot = -(dx * ex + dy * ey)
            if dot > best_dot:
                best, best_dot = e, dot
        return best if best_dot > 0.99 else None

    used = [False] * len(edges)
    strands: list[list[tuple[int, int]]] = []  # node keys
    coords: dict[tuple[int, int], Pt] = {}
    for a, b in edges:
        coords[_key(a)] = a
        coords[_key(b)] = b
    for start in range(len(edges)):
        if used[start]:
            continue
        used[start] = True
        a, b = edges[start]
        chain = [_key(a), _key(b)]
        for forward in (True, False):
            idx = start
            node = chain[-1] if forward else chain[0]
            while True:
                nxt = straight_on(idx, node)
                if nxt is None or used[nxt]:
                    break
                used[nxt] = True
                node2 = other(nxt, node)
                if forward:
                    chain.append(node2)
                else:
                    chain.insert(0, node2)
                idx, node = nxt, node2
        strands.append(chain)

    crossing_nodes = {k for k, v in adj.items() if len(v) == 4}
    closed = [len(ch) > 3 and ch[0] == ch[-1] for ch in strands]
    # strand -> crossings in order: (key, position along strand, chain index)
    occ: dict[tuple[int, int], list[tuple[int, int, int]]] = {}
    for si, ch in enumerate(strands):
        idxs = [i for i in range(1, len(ch) - 1) if ch[i] in crossing_nodes]
        if closed[si] and ch[0] in crossing_nodes:
            idxs = [0] + idxs
        for pos, i in enumerate(idxs):
            occ.setdefault(ch[i], []).append((si, pos, i))
    base: dict[int, int] = {}
    links: dict[int, list[tuple[int, int]]] = {}
    for lst in occ.values():
        if len(lst) == 2:
            (s1, p1, _), (s2, p2, _) = lst
            par = 1 ^ (p1 & 1) ^ (p2 & 1)
            links.setdefault(s1, []).append((s2, par))
            links.setdefault(s2, []).append((s1, par))
    for s in range(len(strands)):
        if s in base:
            continue
        base[s] = 0
        stack = [s]
        while stack:
            u = stack.pop()
            for v, par in links.get(u, []):
                if v not in base:
                    base[v] = base[u] ^ par
                    stack.append(v)

    def unit_at(si: int, i: int) -> tuple[float, float]:
        ch = strands[si]
        a = coords[ch[i - 1]] if i > 0 else coords[ch[-2]]
        b = coords[ch[i + 1]] if i + 1 < len(ch) else coords[ch[1]]
        dx, dy = b[0] - a[0], b[1] - a[1]
        ln = math.hypot(dx, dy) or 1.0
        return dx / ln, dy / ln

    strap_of = {}
    for si, ch in enumerate(strands):
        pts = [coords[k] for k in ch]
        if closed[si]:
            from shapely.geometry import LinearRing

            geom = LinearRing(pts[:-1])
        else:
            geom = LineString(pts)
        strap_of[si] = geom.buffer(width / 2, cap_style="flat", join_style="mitre", mitre_limit=6)
    cut = {si: [] for si in strap_of}
    for k, lst in occ.items():
        if len(lst) != 2:
            continue
        (s1, p1, i1), (s2, p2, i2) = lst
        over1 = (base[s1] ^ (p1 & 1)) == 0
        (so, io), (su, iu) = ((s1, i1), (s2, i2)) if over1 else ((s2, i2), (s1, i1))
        do = unit_at(so, io)
        du = unit_at(su, iu)
        sin = abs(do[0] * du[1] - do[1] * du[0]) or 1.0
        reach = (width / 2 + gap) / sin + gap
        x, y = coords[k]
        bar = LineString(
            [(x - do[0] * reach, y - do[1] * reach), (x + do[0] * reach, y + do[1] * reach)]
        )
        cut[su].append(bar.buffer(width / 2 + gap, cap_style="flat"))
    pieces = []
    for si, shape in strap_of.items():
        if cut[si]:
            shape = shape.difference(unary_union(cut[si]))
        pieces.append(shape)
    return unary_union(pieces), len(occ)


# ── motifs ──────────────────────────────────────────────────────────────


def girih_fill(
    region,
    kind: str = "8",
    cell: float = 10,
    line: float = 1.6,
    color: str = "#520912",
    ground: str | None = None,
    woven: bool = False,
    strap: float | None = None,
    rotate: float = 0,
    offset: Pt = (0, 0),
) -> str:
    """Fill a region with girih: either cut lines through a coloured ground
    (ground given) or woven double-line straps."""
    minx, miny, maxx, maxy = region.bounds
    pad = cell * 4
    lines = girih_lines(kind, cell, (minx - pad, miny - pad, maxx + pad, maxy + pad))
    if rotate or offset != (0, 0):
        lines = affinity.translate(affinity.rotate(lines, rotate, origin=(0, 0)), *offset)
    lines = lines.intersection(region.buffer(cell))
    if not woven:
        cuts = lines.buffer(line / 2, cap_style="flat", join_style="mitre")
        return fill(region.difference(cuts), hexc(ground or color))
    w = strap or line * 3
    straps, _ = weave(lines, w, line * 0.9)
    straps = straps.intersection(region)
    inner = straps.buffer(-line, join_style="mitre")
    return fill(straps, hexc(color)) + fill(inner, hexc(ground or "#ffffff"))


def star_medallion(
    color: str = "#520912",
    accent: str | None = None,
    rings: int = 3,
    number: str | None = None,
    font: str = "Amiri",
    number_color: str | None = None,
) -> str:
    """Eight-point star frame with parallel outlines (as in numbered star
    frames), optional number in the centre."""
    c = hexc(color)
    acc = hexc(accent) if accent else "#ffffff"
    R = 48.0
    outer = khatam(R, 50, 50)
    band = 3.2
    gap = 2.0
    body = [fill(outer, c)]
    cur = outer.buffer(-band, join_style="mitre", mitre_limit=10)
    body.append(fill(cur, acc))
    for _ in range(max(0, rings - 1)):
        cur = cur.buffer(-gap, join_style="mitre", mitre_limit=10)
        body.append(fill(outline(cur, 1.4), c))
        cur = cur.buffer(-1.4, join_style="mitre", mitre_limit=10)
    disc = circle(R * 0.47, 50, 50)
    body.append(fill(outline(disc, 2.2), c))
    body.append(fill(outline(disc.buffer(-3.6), 0.9), c))
    if number:
        body.append(
            # Word's SVG renderer ignores dominant-baseline: place the baseline.
            f'<text x="50" y="59.5" text-anchor="middle" '
            f'font-family="{font}" font-weight="bold" font-size="27" '
            f'fill="{hexc(number_color or color)}">{number}</text>'
        )
    return svg(100, 100, "".join(body))


def rosette_lines(
    n: int = 12,
    R: float = 1.0,
    cx: float = 0,
    cy: float = 0,
    r_tip: float = 0.50,
    inner: float = 0.62,
    r_edge: float = 0.80,
    r_cap: float = 0.96,
) -> MultiLineString:
    """Classical rosette line network.

    A central n-point star (tips at r_tip, inner points at inner·r_tip). The
    two edges meeting at each tip cross there and run on outward, forming the
    sides of the petals that sit between the star's points; each petal is
    capped on its axis at r_cap. Tips are the woven crossings.
    """

    def pol(r, a):
        a = math.radians(a - 90)
        return (cx + R * r * math.cos(a), cy + R * r * math.sin(a))

    step = 360 / n
    tips = [pol(r_tip, i * step) for i in range(n)]
    ins = [pol(r_tip * inner, i * step + step / 2) for i in range(n)]

    def extend(p, q, r):
        """Point where the ray p→q, continued beyond q, meets radius r·R."""
        dx, dy = q[0] - p[0], q[1] - p[1]
        fx, fy = q[0] - cx, q[1] - cy
        a = dx * dx + dy * dy
        b = 2 * (fx * dx + fy * dy)
        c = fx * fx + fy * fy - (r * R) ** 2
        t = (-b + math.sqrt(max(0.0, b * b - 4 * a * c))) / (2 * a)
        return (q[0] + dx * t, q[1] + dy * t)

    segs = []
    for i in range(n):
        t = tips[i]
        v_prev, v_next = ins[(i - 1) % n], ins[i]
        segs.append(LineString([v_prev, t]))
        segs.append(LineString([v_next, t]))
        # crossing at the tip: v_prev→t runs on into the petal after tip i,
        # v_next→t into the petal before it
        e_after = extend(v_prev, t, r_edge)
        e_before = extend(v_next, t, r_edge)
        segs.append(LineString([t, e_after]))
        segs.append(LineString([t, e_before]))
        segs.append(LineString([e_after, pol(r_cap, i * step + step / 2)]))
        segs.append(LineString([e_before, pol(r_cap, i * step - step / 2)]))
    return unary_union(segs)


def rosette(
    color: str = "#520912",
    accent: str | None = None,
    fold: int = 12,
    woven: bool = True,
    petal_fill: str | None = None,
) -> str:
    """Woven strapwork rosette in a double ring: a central star of `fold`
    points (8–16) with its ring of petals."""
    c = hexc(color)
    acc = hexc(accent) if accent else tint(c, 0.86)
    R = 48.0
    disc = circle(R, 50, 50, 128)
    field_r = R - 5.5
    inner = {8: 0.56, 10: 0.60, 12: 0.64, 16: 0.70}.get(fold, 0.64)
    lines = rosette_lines(fold, field_r, 50, 50, r_tip=0.52, inner=inner, r_edge=0.80, r_cap=0.985)
    strap, line = 3.4, 0.9
    straps, _ = weave(lines, strap, line)
    field = circle(field_r, 50, 50, 128)
    straps = straps.intersection(field.buffer(strap / 2))
    body = [fill(disc, c), fill(disc.buffer(-1.8), acc)]
    if petal_fill:
        body.append(fill(field.difference(straps), hexc(petal_fill)))
    body.append(fill(straps, c))
    body.append(fill(straps.buffer(-line, join_style="mitre"), "#ffffff" if woven else acc))
    body.append(fill(outline(circle(field_r + strap / 2 + 0.6, 50, 50, 128), 0.9), c))
    body.append(fill(star(fold, 4.2, 2.4, 50, 50), c))
    return svg(100, 100, "".join(body))


def ring(
    color: str = "#0b3a5a",
    sectors: int = 8,
    inner: float = 0.5,
    gap: float = 2.6,
    cell: float = 6.5,
) -> str:
    """Segmented girih ring (sectors of a coloured annulus with white girih
    cut through) — the Markaz Athari mark's construction."""
    c = hexc(color)
    R = 49.0
    ann = circle(R, 50, 50, 128).difference(circle(R * inner, 50, 50, 96))
    cuts = []
    for i in range(sectors):
        a = math.radians(-90 + 360 * i / sectors + 180 / sectors)
        cuts.append(LineString([(50, 50), (50 + 60 * math.cos(a), 50 + 60 * math.sin(a))]))
    ann = ann.difference(unary_union(cuts).buffer(gap / 2))
    lines = girih_lines("8", cell, (-60, -60, 60, 60))
    lines = affinity.translate(lines, 50, 50).intersection(ann.buffer(2))
    out = ann.difference(lines.buffer(0.7, cap_style="flat", join_style="mitre"))
    return svg(100, 100, fill(out, c))


def shamsa(color: str = "#520912", accent: str | None = None, petals: int = 24) -> str:
    """Title-page sunburst: a woven rosette inside a ring of pointed petals."""
    c = hexc(color)
    acc = hexc(accent) if accent else tint(c, 0.82)
    body = []
    petal_shapes = []
    for i in range(petals):
        a = 360 * i / petals
        p = bezier_shape(
            [
                ((50, 3), (54.5, 10), (54.5, 14), (50, 20)),
                ((50, 20), (45.5, 14), (45.5, 10), (50, 3)),
            ]
        )
        petal_shapes.append(affinity.rotate(p, a, origin=(50, 50)))
    petals_g = unary_union(petal_shapes)
    body.append(fill(petals_g, c))
    body.append(fill(petals_g.buffer(-1.0, join_style="mitre"), acc))
    disc = circle(32.5, 50, 50, 96)
    body.append(fill(disc, c))
    inner = rosette(color, accent, fold=12)
    # embed the rosette scaled into the disc
    rs = 31.0 / 48.0
    body.append(
        f'<g transform="translate({50 - 50 * rs:.3f},{50 - 50 * rs:.3f}) scale({rs:.4f})">'
        + inner[inner.index(">") + 1 : inner.rindex("</svg>")]
        + "</g>"
    )
    return svg(100, 100, "".join(body))


def fleuron(color: str = "#520912") -> Polygon:
    """Central lily-and-curls fleuron (the MSN divider centrepiece), in a
    40×24 box centred at (0,0)."""
    lily = bezier_shape(
        [
            ((0, -12), (2.4, -7), (4.2, -3), (0, 0)),
            ((0, 0), (4.2, 3), (2.4, 7), (0, 12)),
            ((0, 12), (-2.4, 7), (-4.2, 3), (0, 0)),
            ((0, 0), (-4.2, -3), (-2.4, -7), (0, -12)),
        ]
    )
    lily = lily.difference(
        bezier_shape(
            [
                ((0, -7), (1, -4.5), (1.3, -2.5), (0, -1.2)),
                ((0, -1.2), (-1.3, -2.5), (-1, -4.5), (0, -7)),
            ]
        )
    ).difference(
        bezier_shape(
            [
                ((0, 7), (1, 4.5), (1.3, 2.5), (0, 1.2)),
                ((0, 1.2), (-1.3, 2.5), (-1, 4.5), (0, 7)),
            ]
        )
    )
    # side leaves sweeping out and up, plus C-curls
    leaf = bezier_shape(
        [
            ((1.5, -1.0), (5, -5.5), (9, -6.5), (12.5, -4.2)),
            ((12.5, -4.2), (9, -4.4), (5.5, -2.6), (1.5, 0.6)),
            ((1.5, 0.6), (1.4, 0.2), (1.4, -0.5), (1.5, -1.0)),
        ]
    )
    curl = arc_stroke(14.5, 1.0, 3.0, 200, -110, 1.9, ball=0.0)
    right = unary_union([leaf, mirror_y(leaf, 0).buffer(-0.15), curl])
    return unary_union([lily, right, mirror_x(right, 0)])


def swash_divider(color: str = "#520912", aspect: float = 12.0, weight: float = 1.0) -> str:
    """Calligraphic divider: lens-tapered rules into a central fleuron."""
    h = 28.0
    w = h * aspect
    cx, cy = w / 2, h / 2
    t0 = 1.7 * weight
    pts_top, pts_bot = [], []
    L = cx - 19
    for i in range(41):
        f = i / 40
        x = 2 + f * L
        th = t0 * math.sin(math.pi / 2 * f) ** 1.6
        pts_top.append((x, cy - th))
        pts_bot.append((x, cy + th))
    left = Polygon(pts_top + pts_bot[::-1])
    rules = unary_union([left, mirror_x(left, cx)])
    orn = affinity.translate(affinity.scale(fleuron(color), 1.05, 1.05, origin=(0, 0)), cx, cy)
    beads = unary_union([circle(1.1, cx - L * 0.55, cy, 12), circle(1.1, cx + L * 0.55, cy, 12)])
    return svg(w, h, fill(unary_union([rules, orn, beads]), hexc(color)))


def cartouche_shape(kind: str, w: float, h: float) -> Polygon:
    """Cartouche outline centred at (0,0): lobed, banner or arch."""
    if kind == "banner":
        r = h * 0.28
        return box(-w / 2, -h / 2, w / 2, h / 2).buffer(-r).buffer(r, quad_segs=24)
    if kind == "arch":
        # straight top and bottom, ogee-pointed ends
        x0 = w / 2 - h * 0.55
        top = [(-x0, -h / 2), (x0, -h / 2)]
        right = cubic(
            (x0, -h / 2), (x0 + h * 0.42, -h / 2), (w / 2 - h * 0.12, -h * 0.06), (w / 2, 0), 18
        )
        right += cubic(
            (w / 2, 0), (w / 2 - h * 0.12, h * 0.06), (x0 + h * 0.42, h / 2), (x0, h / 2), 18
        )[1:]
        left = [(-x, y) for x, y in reversed(right)]
        return Polygon(top + right[1:] + [(-x0, h / 2)] + left[1:]).buffer(0)
    if kind == "lobed":
        base = box(-w / 2, -h / 2, w / 2, h / 2)
        lobe_r = h * 0.075
        n = max(4, int(w / (lobe_r * 3.2)))
        bumps = []
        for i in range(n + 1):
            x = -w / 2 + h * 0.25 + (w - h * 0.5) * i / n
            bumps.append(circle(lobe_r, x, -h / 2, 16))
            bumps.append(circle(lobe_r, x, h / 2, 16))
        m = max(3, int(h / (lobe_r * 3.2)))
        for i in range(m + 1):
            y = -h / 2 + h * 0.25 + (h * 0.5) * i / m
            bumps.append(circle(lobe_r, -w / 2, y, 16))
            bumps.append(circle(lobe_r, w / 2, y, 16))
        shape = base.union(unary_union(bumps))
        notch = h * 0.16
        for sx in (-1, 1):
            for sy in (-1, 1):
                shape = shape.difference(circle(notch, sx * w / 2, sy * h / 2, 24))
        return shape.buffer(0.3).buffer(-0.3)
    raise ValueError("cartouche kind must be lobed, banner or arch")


def cartouche(
    kind: str = "lobed",
    color: str = "#264237",
    border: str = "#b79a5b",
    aspect: float = 3.2,
    double: bool = True,
    knob: bool = True,
    finials: bool = False,
    ground: str | None = None,
) -> str:
    """Title/label cartouche: filled body with a gold border (and inner rule),
    optional knob on top and fleuron finials at the ends."""
    h = 40.0
    w = h * aspect
    pad = 8.0 + (9.0 if finials else 0)
    shape = cartouche_shape(kind, w, h)
    body = []
    edge = 1.8
    body.append(fill(shape, hexc(border)))
    inner = shape.buffer(-edge, join_style="round")
    body.append(fill(inner, hexc(color) if ground is None else hexc(ground)))
    if double:
        body.append(fill(outline(inner.buffer(-2.2, join_style="round"), 0.7), hexc(border)))
    extras = []
    if knob:
        k = circle(h * 0.11, 0, -h / 2 - h * 0.02, 32)
        extras.append(fill(k, hexc(border)))
        extras.append(fill(k.buffer(-edge * 0.8), hexc(color)))
    if finials:
        f = affinity.scale(fleuron(border), 0.55, 0.55, origin=(0, 0))
        f = f.intersection(box(0, -20, 30, 20))  # right half: leaves and curl
        f = affinity.translate(f, w / 2 - 1.5, 0)
        extras.append(fill(unary_union([f, mirror_x(f, 0)]), hexc(border)))
    content = "".join(body) + "".join(extras)
    W, H = w + 2 * pad, h + 2 * 8 + (h * 0.15 if knob else 0)
    return svg(
        W,
        H,
        f'<g transform="translate({W / 2:.2f},{H / 2 + (h * 0.07 if knob else 0):.2f})">'
        f"{content}</g>",
    )


def band(
    color: str = "#520912", accent: str | None = None, aspect: float = 12.0, kind: str = "rings"
) -> str:
    """Woven border band between double rules: "rings" (interlocking guilloche
    with a star in each link), "chain" (interlocking octagons), "hex-chain",
    "khatam" (interlocking khatams) or a girih field ("8" or "6")."""
    c = hexc(color)
    acc = hexc(accent) if accent else "#ffffff"
    h = 30.0
    w = h * aspect
    inner = box(0, 4.2, w, h - 4.2)
    if kind == "khatam":
        lines = khatam_chain_lines(0, w, h / 2, (h - 8.4) / 2 - 1.2)
        strap, line = 2.6, 0.7
    elif kind in ("chain", "rings", "hex-chain"):
        sides = {"chain": 8, "rings": 0, "hex-chain": 6}[kind]
        r = (h - 8.4) / 2 - 1.6
        lines = link_chain_lines(
            -4,
            w + 4,
            h / 2,
            r,
            sides=sides,
            pitch={"chain": 1.62, "rings": 1.5, "hex-chain": 1.6}[kind],
        )
        strap, line = 2.6, 0.7
    else:
        apothem = (h - 8.4) / 2
        if kind == "8":
            cell = 2 * apothem * math.tan(math.radians(22.5))
        else:
            cell = apothem / math.cos(math.radians(30))
        lines = girih_lines(kind, cell, (-cell * 4, -cell * 4, w + cell * 4, h + cell * 4))
        lines = affinity.translate(lines, 0, h / 2).intersection(inner.buffer(3))
        strap, line = 2.6, 0.7
    straps, _ = weave(lines, strap, line)
    straps = straps.intersection(inner)
    stars = []
    if kind in ("rings", "chain", "hex-chain"):
        r = (h - 8.4) / 2 - 1.6
        n = len(_ring_centres(-4, w + 4, r, {"chain": 1.62, "rings": 1.5, "hex-chain": 1.6}[kind]))
        for cx in _ring_centres(
            -4, w + 4, r, {"chain": 1.62, "rings": 1.5, "hex-chain": 1.6}[kind]
        ):
            if 0 < cx < w:
                stars.append(khatam(r * 0.36, cx, h / 2))
        _ = n
    body = [
        fill(unary_union(stars), c) if stars else "",
        fill(box(0, 0, w, 1.6), c),
        fill(box(0, h - 1.6, w, h), c),
        fill(box(0, 2.6, w, 3.2), c),
        fill(box(0, h - 3.2, w, h - 2.6), c),
        fill(straps, c),
        fill(straps.buffer(-line, join_style="mitre"), acc),
    ]
    return svg(w, h, "".join(body))


def corner(color: str = "#520912", accent: str | None = None) -> str:
    """Corner piece: double rule turning the corner round a woven rosette."""
    c = hexc(color)
    acc = hexc(accent) if accent else tint(c, 0.86)
    L = LineString([(5, 100), (5, 5), (100, 5)])
    rules = unary_union(
        [
            L.buffer(1.4, cap_style="flat", join_style="mitre"),
            LineString([(11, 100), (11, 11), (100, 11)]).buffer(
                0.55, cap_style="flat", join_style="mitre"
            ),
        ]
    )
    disc = circle(15, 24, 24, 64)
    rules = rules.difference(disc.buffer(1.6))
    ros = rosette(color, accent, fold=8)
    sc = 15 / 48
    leaf = affinity.translate(
        affinity.rotate(affinity.scale(fleuron(color), 0.7, 0.7, origin=(0, 0)), 0), 62, 11
    )
    leaf2 = affinity.translate(
        affinity.rotate(affinity.scale(fleuron(color), 0.7, 0.7, origin=(0, 0)), 90), 11, 62
    )
    body = fill(unary_union([rules, leaf, leaf2]), c)
    body += (
        f'<g transform="translate({24 - 50 * sc:.3f},{24 - 50 * sc:.3f}) scale({sc:.4f})">'
        + ros[ros.index(">") + 1 : ros.rindex("</svg>")]
        + "</g>"
    )
    _ = acc
    return svg(100, 100, body)


def frame(
    color: str = "#520912",
    accent: str | None = None,
    aspect: float = 210 / 297,
    band_width: float = 22.0,
    kind: str = "rings",
) -> str:
    """Page/cover frame: a woven border (khatam chain or girih field) between
    double rules all round, with woven rosettes in the corners."""
    c = hexc(color)
    acc = hexc(accent) if accent else "#ffffff"
    H = 1000.0
    W = H * aspect
    m = 6.0
    outer = box(m, m, W - m, H - m)
    inner_edge = outer.buffer(-band_width, join_style="mitre")
    strip = outer.buffer(-3.5, join_style="mitre").difference(
        inner_edge.buffer(3.5, join_style="mitre")
    )
    mid = m + band_width / 2
    r_corner = band_width * 0.95
    stars = []
    if kind in ("rings", "chain", "khatam"):
        r = (band_width - 7.0) / 2 - 0.8

        def run(x0, x1, y):
            if kind == "khatam":
                return khatam_chain_lines(x0, x1, y, r)
            sides, pitch = (0, 1.5) if kind == "rings" else (8, 1.62)
            for cx in _ring_centres(x0, x1, r, pitch):
                stars.append((cx, y))
            return link_chain_lines(x0, x1, y, r, sides=sides, pitch=pitch)

        a, b = mid + r_corner + 2, W - mid - r_corner - 2
        top = run(a, b, mid)
        bottom = run(a, b, H - mid)
        n_h = len(stars)
        a2, b2 = mid + r_corner + 2, H - mid - r_corner - 2
        vert = run(a2, b2, 0)
        stars[n_h:] = [(y, x) for x, y in stars[n_h:]]
        vert = affinity.affine_transform(vert, [0, 1, 1, 0, 0, 0])
        left = affinity.translate(vert, mid, 0)
        right = affinity.translate(vert, W - mid, 0)
        star_pts = (
            stars[:n_h]
            + [(mid, y) for _, y in stars[n_h:]]
            + [(W - mid, y) for _, y in stars[n_h:]]
        )
        stars = star_pts
        lines = unary_union([top, bottom, left, right])
    else:
        apothem = (band_width - 7.0) / 2
        cell = 2 * apothem * math.tan(math.radians(22.5))
        lines = girih_lines("8", cell, (-cell * 3, -cell * 3, W + cell * 3, H + cell * 3))
        lines = affinity.translate(lines, mid, mid).intersection(strip.buffer(2))
    straps, _ = weave(lines, band_width * 0.11, band_width * 0.03)
    straps = straps.intersection(strip)
    rules = unary_union(
        [
            outline(outer, 2.2),
            outline(outer.buffer(-3.0, join_style="mitre"), 0.8),
            inner_edge.buffer(3.0, join_style="mitre").difference(
                inner_edge.buffer(2.2, join_style="mitre")
            ),
            outline(inner_edge, 2.2),
        ]
    )
    body = [
        fill(rules, c),
        fill(straps, c),
        fill(straps.buffer(-band_width * 0.035, join_style="mitre"), acc),
    ]
    if stars and kind != "khatam":
        rr = (band_width - 7.0) / 2 - 0.8
        body.append(fill(unary_union([khatam(rr * 0.36, x, y) for x, y in stars]), c))
    ros = rosette(color, accent, fold=8)
    sc = r_corner / 48
    inner_svg = ros[ros.index(">") + 1 : ros.rindex("</svg>")]
    for x, y in ((mid, mid), (W - mid, mid), (mid, H - mid), (W - mid, H - mid)):
        body.append(fill(circle(r_corner + 2.5, x, y, 48), "#ffffff"))
        body.append(
            f'<g transform="translate({x - 50 * sc:.3f},{y - 50 * sc:.3f}) '
            f'scale({sc:.4f})">{inner_svg}</g>'
        )
    return svg(W, H, "".join(body))


def note_rule(
    kind: str = "tapered", color: str = "#1A1A1A", aspect: float = 40.0, rtl: bool = True
) -> str:
    """Footnote separator art (a short rule, start-aligned).

    tapered — a rule that thins from the text edge outward;
    khatam  — hairline rule opening with a small eight-point star;
    swash   — the calligraphic swash divider in miniature;
    diamond — hairline rule with a centred diamond and end beads.
    rtl puts the heavy end / star at the right (the start of an RTL line).
    """
    c = hexc(color)
    h = 10.0
    w = h * aspect
    y = h / 2
    shapes = []
    if kind == "tapered":
        pts_t, pts_b = [], []
        for i in range(41):
            f = i / 40
            th = 0.85 * (1 - f) ** 1.3 + 0.12
            x = f * w
            pts_t.append((x, y - th))
            pts_b.append((x, y + th))
        shapes.append(Polygon(pts_t + pts_b[::-1]))
    elif kind == "khatam":
        k = khatam(4.2, 5.0, y)
        shapes.append(k.difference(k.buffer(-0.9, join_style="mitre")))
        shapes.append(circle(1.0, 5.0, y, 12))
        shapes.append(box(10.5, y - 0.3, w - 2.5, y + 0.3))
        shapes.append(regular(4, 1.4, w - 1.5, y))
    elif kind == "diamond":
        shapes.append(box(2.5, y - 0.28, w - 2.5, y + 0.28))
        shapes.append(regular(4, 2.6, w / 2, y))
        shapes.append(circle(1.0, 1.2, y, 12))
        shapes.append(circle(1.0, w - 1.2, y, 12))
    elif kind == "swash":
        return swash_divider(c, aspect=aspect / 2.6, weight=0.7)
    else:
        raise ValueError("note rule kind must be tapered, khatam, swash or diamond")
    g = unary_union(shapes)
    if rtl:
        g = mirror_x(g, w / 2)
    return svg(w, h, fill(g, c))


def dots_divider(color: str = "#520912", aspect: float = 10.0) -> str:
    """Light divider: hairline rules, a small khatam, diamonds and beads."""
    c = hexc(color)
    h = 20.0
    w = h * aspect
    cx, cy = w / 2, h / 2
    rules = unary_union(
        [
            box(w * 0.08, cy - 0.35, cx - 16, cy + 0.35),
            box(cx + 16, cy - 0.35, w * 0.92, cy + 0.35),
        ]
    )
    k = khatam(7.5, cx, cy)
    shapes = [rules, k.difference(k.buffer(-1.3, join_style="mitre")), circle(1.6, cx, cy, 16)]
    for dx in (-11.5, 11.5):
        shapes.append(regular(4, 2.4, cx + dx, cy))
    for x in (w * 0.08, w * 0.92):
        shapes.append(circle(1.3, x, cy, 12))
    return svg(w, h, fill(unary_union(shapes), c))


__all__ = [
    "band",
    "cartouche",
    "corner",
    "dots_divider",
    "frame",
    "girih_fill",
    "ring",
    "rosette",
    "shamsa",
    "star_medallion",
    "swash_divider",
    "weave",
    "MultiLineString",
]
