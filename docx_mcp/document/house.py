"""House colours: one place that decides a book's colours.

Every design tool (ornaments, running heads, cartouches, separators,
numbered entries, frames …) takes its colour from the document's house
palette unless a colour is passed explicitly. Change the palette once and
everything made afterwards follows it.

The palette is kept in the theme (accent1 = accent, accent2 = secondary,
dk2 = ink) so Word's own colour pickers show it, and in custom properties so
it survives theme edits.

Palettes are taken from studied books and the author's own work; restrained
on purpose — colour marks structure (numbers, ornaments, the Prophet's
words), the reading text stays ink.
"""

from __future__ import annotations

from typing import Any

from .ooxml import color_hex

HOUSE_PALETTES: dict[str, dict[str, str]] = {
    # the author's tahqiq house style
    "maroon": {
        "accent": "520912",
        "secondary": "A8874E",
        "ink": "1A1A1A",
        "label": "Deep maroon and old gold (author's house style)",
    },
    # جمهرة مستخرجات الجرح والتعديل — muted brick maroon on ivory
    "brick": {
        "accent": "7A2E2A",
        "secondary": "B08D57",
        "ink": "1C1A17",
        "label": "Brick maroon (Dar Rawai' al-Kitab)",
    },
    # الرد القويم — near-black navy with antique gold
    "navy-gold": {
        "accent": "1F2E4A",
        "secondary": "B8955A",
        "ink": "151515",
        "label": "Navy and antique gold (cover style)",
    },
    # مناسك الحج — calm Islamic green with sand gold
    "green-gold": {
        "accent": "2F6E54",
        "secondary": "C2A36A",
        "ink": "1A1A1A",
        "label": "Islamic green and sand (manasik booklet)",
    },
    # Markaz Athari mark — deep teal-blue
    "teal": {
        "accent": "0F4C5C",
        "secondary": "C29B57",
        "ink": "1A1A1A",
        "label": "Deep teal (Markaz Athari)",
    },
    # maroon silk cover mock-up
    "burgundy-silk": {
        "accent": "6B1F33",
        "secondary": "C6A15B",
        "ink": "1A1A1A",
        "label": "Burgundy silk and gold",
    },
    # plain scholarly print: ink only, grey for ornaments
    "ink": {
        "accent": "1A1A1A",
        "secondary": "6B6B6B",
        "ink": "000000",
        "label": "Ink only (classic print)",
    },
}

_PROP = "docx-mcp.house."
DEFAULT_PALETTE = "maroon"


class HouseStyleMixin:
    """The document's house colours and the fallback for every design tool."""

    def set_house_colors(
        self,
        palette: str | None = None,
        *,
        accent: str | None = None,
        secondary: str | None = None,
        ink: str | None = None,
    ) -> dict:
        """Set the book's colours: a palette (see list_house_palettes) and/or
        explicit colours. accent — numbers, ornaments, the Prophet's words;
        secondary — gold edges, fills; ink — the reading text."""
        values = dict(HOUSE_PALETTES[DEFAULT_PALETTE])
        if palette:
            if palette not in HOUSE_PALETTES:
                raise ValueError(f"palette must be one of {sorted(HOUSE_PALETTES)}")
            values = dict(HOUSE_PALETTES[palette])
        for key, val in (("accent", accent), ("secondary", secondary), ("ink", ink)):
            if val:
                hx = color_hex(val)
                if not hx:
                    raise ValueError(f"{key} must be a hex colour such as 520912")
                values[key] = hx
        for key in ("accent", "secondary", "ink"):
            self.set_custom_property(_PROP + key, values[key])
        self.apply_color_palette(
            colors={
                "accent1": values["accent"],
                "accent2": values["secondary"],
                "dk2": values["ink"],
            }
        )
        return {
            "palette": palette or "custom",
            **{k: values[k] for k in ("accent", "secondary", "ink")},
        }

    def get_house_colors(self) -> dict:
        return {k: self._house(k) for k in ("accent", "secondary", "ink")}

    def list_house_palettes(self) -> dict:
        return HOUSE_PALETTES

    def _house(self, slot: str = "accent") -> str:
        """The house colour for slot (accent, secondary, ink)."""
        try:
            props: dict[str, Any] = self.get_custom_properties()
        except Exception:  # noqa: BLE001 - a missing part means no house colours yet
            props = {}
        val = props.get(_PROP + slot)
        if isinstance(val, dict):
            val = val.get("value")
        if val:
            return str(val)
        return HOUSE_PALETTES[DEFAULT_PALETTE][slot]

    def _c(self, color: str | None, slot: str = "accent") -> str:
        """An explicit colour, or the house colour for slot (hex, no #)."""
        if color:
            return color_hex(color) or self._house(slot)
        return self._house(slot)
