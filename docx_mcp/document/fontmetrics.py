"""Measure text with the real font when it is installed (for ornaments that
must fit round a heading). Falls back to an average advance width."""

from __future__ import annotations

import functools
import glob
import os
import sys


@functools.lru_cache(maxsize=1)
def _font_index() -> dict[str, str]:
    dirs = []
    if sys.platform.startswith("win"):
        dirs += [
            os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts"),
            os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts"),
        ]
    else:
        dirs += [
            "/usr/share/fonts",
            "/usr/local/share/fonts",
            os.path.expanduser("~/.fonts"),
            os.path.expanduser("~/.local/share/fonts"),
            "/Library/Fonts",
            os.path.expanduser("~/Library/Fonts"),
        ]
    index: dict[str, str] = {}
    try:
        from PIL import ImageFont
    except ImportError:  # pragma: no cover
        return index
    for d in dirs:
        for path in glob.glob(os.path.join(d, "**", "*.[ot]tf"), recursive=True):
            try:
                fam, style = ImageFont.truetype(path, 10).getname()
            except Exception:  # noqa: BLE001 - unreadable font files are skipped
                continue
            key = fam.lower()
            if key not in index or (style or "").lower() in ("regular", "normal", "book"):
                index[key] = path
    return index


def text_width_pt(text: str, font: str | None, size_pt: float, bold: bool = False) -> float:
    """Advance width of text in points (shaped with raqm when available)."""
    try:
        from PIL import ImageFont, features

        path = _font_index().get((font or "").lower())
        if path:
            engine = ImageFont.Layout.RAQM if features.check("raqm") else ImageFont.Layout.BASIC
            f = ImageFont.truetype(path, 100, layout_engine=engine)
            w = (
                f.getlength(
                    text, direction="rtl" if any("\u0600" <= c <= "\u06ff" for c in text) else None
                )
                if engine == ImageFont.Layout.RAQM
                else f.getlength(text)
            )
            return w / 100 * size_pt * (1.04 if bold else 1.0)
    except Exception:  # noqa: BLE001 - measuring is best effort
        pass
    return len(text) * size_pt * 0.48
