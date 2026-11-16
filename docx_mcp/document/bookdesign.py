"""Whole-book presets and Word power features most people never touch.

``apply_book_preset("tahqiq")`` sets up a scholarly Arabic book in one call:
page, margins, fonts, a labelled style set, running head and folio, footnote
design ("=" continuation, maroon rule), honorific glyphs, fields that refresh
on open. Everything it creates is a named style, so the author keeps editing
with styles in Word.

Power features:

- ``restrict_formatting_to_styles`` — Word's "Limit formatting to a selection
  of styles": hand formatting is switched off, so the book stays consistent.
- ``set_style_pane`` — show only styles in use (or recommended) in the Styles
  pane, with their Arabic labels.
- ``set_font_embedding`` — Word embeds the fonts (Amiri, Jameel Noori…) next
  time it saves, so the file looks right on a computer without them.
- ``hide_proofing_marks`` — no red/blue squiggles across Arabic text.
- ``attach_template`` — link a .dotx; with auto-update, editing the template's
  styles restyles every book that uses it.
- ``insert_separator`` — the section separator (❁ ❁ ❁ or an ornament).
"""

from __future__ import annotations

from typing import Any

from lxml import etree

from .ooxml import W14, W, add_relationship, apply_para_props, make_paragraph, make_run

R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"

PRESETS: dict[str, dict[str, Any]] = {
    "tahqiq": {
        "label": "Tahqiq / risalah (author's house style)",
        "direction": "rtl",
        "page_size": "17x24",
        "margins_cm": (1.9, 1.9, 2.0, 1.7),  # top, bottom, inside, outside
        "body_font": "Amiri",
        "heading_font": "Sakkal Majalla",
        "body_pt": 13,
        "line": 1.15,
        "chapter_pt": 20,
        "section_pt": 15,
        "note_pt": 10,
        "note_line_pt": 15,
        "palette": "maroon",
        "chapters_new_page": False,
        "running_head": "fleuron",
        "folio": "head",
        "first_line_cm": 0.0,
    },
    "kitab": {
        "label": "Full book: chapters on new pages, larger type",
        "direction": "rtl",
        "page_size": "17x24",
        "margins_cm": (2.0, 2.0, 2.2, 1.7),
        "body_font": "Amiri",
        "heading_font": "Sakkal Majalla",
        "body_pt": 14,
        "line": 1.2,
        "chapter_pt": 24,
        "section_pt": 16,
        "note_pt": 10.5,
        "note_line_pt": 16,
        "palette": "brick",
        "chapters_new_page": True,
        "running_head": "fleuron",
        "folio": "head",
        "first_line_cm": 0.0,
    },
    "urdu-nastaliq": {
        "label": "Urdu book in Jameel Noori Nastaleeq",
        "direction": "rtl",
        "page_size": "17x24",
        "margins_cm": (2.0, 2.0, 2.0, 1.8),
        "body_font": "Jameel Noori Nastaleeq",
        "heading_font": "Jameel Noori Nastaleeq",
        "body_pt": 15,
        "line": 1.6,
        "chapter_pt": 24,
        "section_pt": 18,
        "note_pt": 11,
        "note_line_pt": 21,
        "palette": "maroon",
        "chapters_new_page": True,
        "running_head": "fleuron",
        "folio": "head",
        "first_line_cm": 0.0,
        "language": "ur-PK",
        "glyph_set": "jameel",
    },
    "english": {
        "label": "English Islamic book (translation conventions)",
        "direction": "ltr",
        "page_size": "16x23",
        "margins_cm": (2.0, 2.2, 2.2, 1.8),
        "body_font": "Palatino Linotype",
        "heading_font": "Palatino Linotype",
        "arabic_font": "Amiri",
        "body_pt": 11.5,
        "line": 1.15,
        "chapter_pt": 17,
        "section_pt": 13,
        "note_pt": 9,
        "note_line_pt": 11.5,
        "palette": "ink",
        "chapters_new_page": True,
        "running_head": "none",
        "folio": "foot",
        "first_line_cm": 0.6,
        "language": "en-GB",
    },
    "bilingual": {
        "label": "Arabic text with English translation (e.g. Bulugh al-Maram)",
        "direction": "ltr",
        "page_size": "royal",
        "margins_cm": (2.0, 2.2, 2.2, 1.8),
        "body_font": "Palatino Linotype",
        "heading_font": "Palatino Linotype",
        "arabic_font": "Amiri",
        "arabic_pt": 15,
        "body_pt": 11.5,
        "line": 1.15,
        "chapter_pt": 16,
        "section_pt": 13,
        "note_pt": 9,
        "note_line_pt": 11.5,
        "palette": "ink",
        "chapters_new_page": False,
        "running_head": "none",
        "folio": "foot",
        "first_line_cm": 0.6,
        "language": "en-GB",
    },
}

HEADING_DESIGNS = ("plain", "fleuron", "flanked", "cartouche", "bullet")

# STYLE PANE filter bits ([MS-OI29500] CT_StylePaneFilter@val)
_PANE_BITS = {
    "all": 0x0001,
    "custom": 0x0002,
    "in-use": 0x0008,
    "headings": 0x0010,
    "clear": 0x0800,
    "recommended": 0x2000,
    "alternate-names": 0x4000,
}
_SORTS = {
    "name": "name",
    "priority": "priority",
    "font": "font",
    "based-on": "basedOn",
    "type": "type",
    "default": "default",
}


class BookDesignMixin:
    """One-call book presets and Word power features."""

    # ── preset ──────────────────────────────────────────────────────────

    def apply_book_preset(
        self,
        preset: str = "tahqiq",
        *,
        title: str | None = None,
        author: str | None = None,
        palette: str | None = None,
        running_head: str | None = None,
        chapters_new_page: bool | None = None,
        justify: str = "justify",
        honorifics: bool = True,
        tashkil: bool = True,
        notes_tashkil: str = "strip",
        note_separator: str = "rule",
        note_marks: str | None = None,
        note_continuation: bool = True,
        page_numbers: str = "decimal",
        title_page: bool = True,
        overrides: dict[str, Any] | None = None,
    ) -> dict:
        """Set up a whole book in one call.

        Args:
            preset: tahqiq (Amiri 13 / Sakkal Majalla, 17×24, the author's
                house style), kitab (chapters on new pages), urdu-nastaliq
                (Jameel Noori Nastaleeq), english (an English Islamic book),
                bilingual (Arabic text with English translation, as in
                Bulugh al-Maram).
            title / author: Document properties; the title runs in the head.
            palette: House colours (list_house_palettes): maroon, brick,
                navy-gold, green-gold, teal, burgundy-silk, ink. Default per
                preset. Colour marks structure only; reading text stays ink.
            running_head: fleuron (default for Arabic), rule, boxed, vignette,
                cartouche, or none (English: folio at the foot only).
            chapters_new_page: Override the preset.
            justify: justify, or kashida (low kashida: stretches with tatweel
                as in print; Word drops lines of vocalised text under medium
                or high kashida, so those are not used).
            honorifics: Replace honorific phrases with their glyphs (matches
                vocalised text too).
            tashkil: The text carries harakat — line spacing is "at least",
                never "exactly", so fatha/kasra/shadda are never clipped.
            notes_tashkil: footnotes and references carry no tashkil in
                printed books — "strip" (default) removes it on every save,
                "keep-shadda" keeps only shadda, "keep" leaves notes alone.
            note_separator: rule (65 %), double, tapered, khatam, diamond,
                swash, default, none.
            note_marks: parentheses «(١)» (Arabic default), dash «57-» in the
                note with a plain superscript in the text (English default),
                brackets, plain.
            note_continuation: "=" where a note runs on.
            page_numbers: decimal, arabic-abjad, arabic-alpha …
            title_page: no head or folio on the first page.
            overrides: any preset value, e.g. {"body_pt": 14, "line": 1.3}.

        Styles (with Arabic labels): Normal (متن), Title (العنوان), Heading 1
        (باب), Heading 2 (فصل), Heading 3, Narration Number (رقم الحديث),
        Prophetic Speech (كلام النبي ﷺ), Matn (نص الحديث), Isnad (الإسناد),
        Quran (آية), Separator (فاصل), footnote text (الحاشية); English
        presets add Arabic Text, Translation, Transliteration, Glossary Term.
        Headings carry no rules: decorate one with insert_heading(design=…).
        """
        if preset not in PRESETS:
            raise ValueError(f"preset must be one of {sorted(PRESETS)}")
        if justify not in ("justify", "kashida"):
            raise ValueError("justify must be justify or kashida")
        cfg = {**PRESETS[preset], **(overrides or {})}
        if chapters_new_page is not None:
            cfg["chapters_new_page"] = chapters_new_page
        head_style = running_head or cfg["running_head"]
        from .bookparts import HEAD_STYLES

        if head_style not in (*HEAD_STYLES, "none"):
            raise ValueError(f"running_head must be one of {(*HEAD_STYLES, 'none')}")
        rtl = cfg["direction"] == "rtl"
        dirn = cfg["direction"]
        body, head = cfg["body_font"], cfg["heading_font"]
        arabic = cfg.get("arabic_font", body)
        lang = cfg.get("language", "ar-SA")
        done: list[str] = []

        house = self.set_house_colors(palette or cfg["palette"])
        accent, ink = house["accent"], house["ink"]
        done.append(f"palette:{palette or cfg['palette']}")

        top, bottom, inside, outside = cfg["margins_cm"]
        self.set_book_layout(
            page_size=cfg["page_size"],
            top_cm=top,
            bottom_cm=bottom,
            inside_cm=inside,
            outside_cm=outside,
            mirror_margins=True,
            rtl_binding=rtl,
            header_cm=1.0,
            footer_cm=0.9,
        )
        done.append("page")
        if rtl:
            self.set_document_direction(
                "rtl",
                complex_script_font=body,
                complex_script_size_pt=cfg["body_pt"],
                language=lang,
            )
            self.set_theme_fonts(heading_arabic=head, body_arabic=body)
        else:
            self.set_theme_fonts(
                heading_latin=head, body_latin=body, heading_arabic=arabic, body_arabic=arabic
            )
        done.append("direction+theme")

        line = cfg["line"]
        if tashkil and rtl:
            line = max(line, 1.15)
        body_para: dict[str, Any] = {
            "direction": dirn,
            "alignment": "kashida" if (justify == "kashida" and rtl) else "justify",
            "space_after_pt": 4,
            "widow_control": True,
        }
        if tashkil:
            body_para["line_at_least_pt"] = round(cfg["body_pt"] * 1.2 * line, 1)
        else:
            body_para["line_spacing"] = line
        if cfg.get("first_line_cm"):
            body_para["first_line_cm"] = cfg["first_line_cm"]
        run_body: dict[str, Any] = {
            "font": body,
            "cs_font": arabic,
            "size_pt": cfg["body_pt"],
            "color": ink,
        }
        run_body["cs_language" if rtl else "language"] = lang
        self.define_style(
            "Normal", run=run_body, aliases=["متن"] if rtl else None, paragraph=body_para
        )

        hrun = {"font": head, "cs_font": head if rtl else arabic, "bold": True, "color": ink}
        self.define_style(
            "Title",
            based_on="Normal",
            next_style="Normal",
            aliases=["العنوان"] if rtl else None,
            gallery=True,
            ui_priority=1,
            run={**hrun, "size_pt": cfg["chapter_pt"] * 1.8, "color": accent},
            paragraph={
                "alignment": "center",
                "space_before_pt": 0,
                "space_after_pt": 14,
                "keep_with_next": True,
                "first_line_cm": 0,
                "line_spacing": 1.0,
            },
        )
        # title page lines under the title (author, editor, publisher)
        self.define_style(
            "Subtitle",
            based_on="Normal",
            next_style="Normal",
            aliases=["عنوان فرعي"] if rtl else None,
            gallery=True,
            ui_priority=2,
            run={**hrun, "bold": False, "size_pt": cfg["section_pt"], "color": ink},
            paragraph={
                "alignment": "center",
                "space_before_pt": 0,
                "space_after_pt": 8,
                "keep_with_next": True,
                "first_line_cm": 0,
                "line_spacing": 1.0,
            },
        )
        h1_run = {**hrun, "size_pt": cfg["chapter_pt"]}
        if rtl:
            h1_run["color"] = accent
        self.define_style(
            "Heading 1",
            based_on="Normal",
            next_style="Normal",
            aliases=["باب"] if rtl else None,
            gallery=True,
            ui_priority=2,
            run=h1_run,
            paragraph={
                "alignment": "center",
                "space_before_pt": 18,
                "space_after_pt": 10,
                "keep_with_next": True,
                "keep_together": True,
                "outline_level": 0,
                "first_line_cm": 0,
                "page_break_before": bool(cfg["chapters_new_page"]),
                "line_spacing": 1.0,
            },
        )
        h2_run = {**hrun, "size_pt": cfg["section_pt"]}
        if not rtl:
            h2_run["italic"] = True
        self.define_style(
            "Heading 2",
            based_on="Normal",
            next_style="Normal",
            aliases=["فصل"] if rtl else None,
            gallery=True,
            ui_priority=3,
            run=h2_run,
            paragraph={
                "alignment": "start" if rtl else "center",
                "space_before_pt": 10,
                "space_after_pt": 6,
                "keep_with_next": True,
                "outline_level": 1,
                "first_line_cm": 0,
                "line_spacing": 1.0,
            },
        )
        self.define_style(
            "Heading 3",
            based_on="Normal",
            next_style="Normal",
            aliases=["مسألة"] if rtl else None,
            gallery=True,
            ui_priority=4,
            run={**hrun, "size_pt": cfg["body_pt"] + 0.5},
            paragraph={
                "alignment": "start",
                "space_before_pt": 6,
                "space_after_pt": 3,
                "keep_with_next": True,
                "outline_level": 2,
                "first_line_cm": 0,
            },
        )
        for name, label, run in (
            ("Narration Number", "رقم الحديث", {"bold": True, "color": accent}),
            ("Prophetic Speech", "كلام النبي ﷺ", {"bold": True, "color": accent if rtl else ink}),
            ("Matn", "نص الحديث", {"bold": True, "color": ink}),
            ("Isnad", "الإسناد", {"color": ink}),
            ("Quran", "آية", {"bold": True, "color": ink}),
        ):
            self.define_style(name, "character", aliases=[label], gallery=True, run=run)
        self.define_style(
            "Separator",
            based_on="Normal",
            next_style="Normal",
            aliases=["فاصل"],
            gallery=True,
            run={"color": accent, "size_pt": cfg["body_pt"] - 2},
            paragraph={
                "alignment": "center",
                "space_before_pt": 6,
                "space_after_pt": 8,
                "first_line_cm": 0,
                "line_spacing": 1.0,
            },
        )
        note_para: dict[str, Any] = {
            "direction": dirn,
            "alignment": "justify",
            "space_after_pt": 1,
            "space_before_pt": 0,
            "first_line_cm": 0,
        }
        if tashkil:
            note_para["line_at_least_pt"] = cfg["note_line_pt"]
        else:
            note_para["line_exact_pt"] = cfg["note_line_pt"]
        self.define_style(
            "footnote text",
            based_on="Normal",
            aliases=["الحاشية"],
            run={"font": body, "cs_font": arabic, "size_pt": cfg["note_pt"], "color": ink},
            paragraph=note_para,
        )
        if not rtl:
            self._english_styles(cfg, arabic, ink, accent)
        done.append("styles")

        if head_style != "none":
            # book names in running heads go unvocalised, like references
            head_title = title
            if title and notes_tashkil in ("strip", "keep-shadda"):
                from .arabic import strip_tashkil

                head_title = strip_tashkil(title, keep_shadda=notes_tashkil == "keep-shadda")
            self.design_running_head(
                head_style,
                title=head_title,
                color=accent,
                font=head,
                size_pt=11 if rtl else 9.5,
                bold=rtl,
            )
        if cfg["folio"] == "foot" or head_style == "none":
            self.design_header_footer(
                "footer",
                center="{PAGE}",
                font=body,
                cs_font=body,
                size_pt=cfg["body_pt"] - 1,
                color=ink,
            )
            if head_style == "none":
                self.design_header_footer("header", center="")
        if title_page:
            self.design_header_footer("header", center="", variant="first")
            self.design_header_footer("footer", center="", variant="first")
        if page_numbers != "decimal":
            self.set_page_numbering(number_format=page_numbers)
        done.append("header/footer")

        self.set_note_separator("footnote", note_separator)
        if note_continuation:
            self.set_note_continuation(
                "footnote", color=accent, bold=True, notice_text="=" if rtl else "(cont.)"
            )
        note_marks = note_marks or ("parentheses" if rtl else "dash")
        self._note_marks_style = note_marks
        done.append("footnotes")

        if honorifics:
            self.apply_honorifics(glyph_set=cfg.get("glyph_set", "unicode"))
            done.append("honorifics")
        if title or author:
            props = {}
            if title:
                props["title"] = title
            if author:
                props["creator"] = author
            self.set_properties(**props)
        self.set_custom_property("docx-mcp.note-marks", note_marks)
        # how Word draws note/page numbers in exported PDFs (see export_pdf)
        self.set_custom_property("docx-mcp.numerals", "hindi" if rtl else "context")
        if notes_tashkil not in ("strip", "keep-shadda", "keep"):
            raise ValueError("notes_tashkil must be strip, keep-shadda or keep")
        self.set_custom_property("docx-mcp.notes-tashkil", notes_tashkil if rtl else "keep")
        self.hide_proofing_marks(True)
        self.set_style_pane(show="in-use", sort="priority", alternate_names=True)
        self.set_update_fields_on_open(True)
        done.append("settings")
        return {"preset": preset, "applied": done, "config": cfg, "house": house}

    def _english_styles(self, cfg: dict, arabic: str, ink: str, accent: str) -> None:
        """Styles of an English Islamic book (studied from a Bulugh al-Maram
        translation): Arabic text blocks, numbered translations, glossary."""
        apt = cfg.get("arabic_pt", cfg["body_pt"] + 3)
        self.define_style(
            "Arabic Text",
            based_on="Normal",
            next_style="Translation",
            gallery=True,
            run={
                "font": arabic,
                "cs_font": arabic,
                "size_pt": apt,
                "rtl": True,
                "cs_language": "ar-SA",
                "color": ink,
            },
            paragraph={
                "direction": "rtl",
                "alignment": "justify",
                "first_line_cm": 0,
                "space_before_pt": 8,
                "space_after_pt": 4,
                "keep_with_next": True,
                "line_at_least_pt": round(apt * 1.45, 1),
            },
        )
        self.define_style(
            "Translation",
            based_on="Normal",
            next_style="Arabic Text",
            gallery=True,
            run={"color": ink},
            paragraph={
                "direction": "ltr",
                "alignment": "justify",
                "first_line_cm": cfg.get("first_line_cm", 0.6),
                "space_after_pt": 6,
            },
        )
        self.define_style(
            "Arabic Heading",
            based_on="Arabic Text",
            next_style="Heading 2",
            gallery=True,
            run={"size_pt": apt + 3, "bold": False},
            paragraph={
                "alignment": "center",
                "keep_with_next": True,
                "space_before_pt": 14,
                "space_after_pt": 2,
            },
        )
        self.define_style("Transliteration", "character", gallery=True, run={"italic": True})
        self.define_style(
            "Glossary Term", "character", gallery=True, run={"bold": True, "italic": True}
        )
        self.define_style(
            "Glossary Entry",
            based_on="Normal",
            gallery=True,
            run={"italic": True},
            paragraph={"first_line_cm": 0, "hanging_cm": 0.6, "space_after_pt": 3},
        )
        self.define_style(
            "Signature",
            based_on="Normal",
            gallery=True,
            run={"bold": True, "italic": True},
            paragraph={"alignment": "end", "first_line_cm": 0, "space_before_pt": 12},
        )

    # ── headings ────────────────────────────────────────────────────────

    def insert_heading(
        self,
        text: str,
        *,
        level: int = 1,
        after_para_id: str | None = None,
        design: str = "plain",
        color: str | None = None,
        ornament_width_pct: float = 30,
    ) -> dict:
        """A heading in the book's Heading N style (so it reaches the TOC and
        the running head), optionally decorated as the studied books do:

          plain     — the style alone.
          fleuron   — a small calligraphic swash centred under it.
          flanked   — set between two ornamental scroll brackets.
          cartouche — on a lobed cartouche (real text over vector art).
          bullet    — opened by a small eight-point star (start-aligned).
        """
        if design not in HEADING_DESIGNS:
            raise ValueError(f"design must be one of {HEADING_DESIGNS}")
        if not 1 <= level <= 3:
            raise ValueError("level must be 1–3")
        color = self._c(color)
        style = f"Heading {level}"
        if design == "cartouche":
            sid = self._style_id_for(style, "paragraph")
            res = self.insert_cartouche_title(
                text, after_para_id=after_para_id, kind="lobed", style=sid, size_pt=16
            )
            # same shape as the other designs: continue writing after the heading
            return {"design": design, **res, "after_para_id": res["para_id"]}
        res = self.insert_paragraph(after_para_id, text, style=style) if after_para_id else None
        if res is None:
            doc = self._require("word/document.xml")
            body = doc.find(f"{W}body")
            p = make_paragraph(
                self._new_para_id(), {"style": self._style_id_for(style, "paragraph")}
            )
            p.append(make_run(text, None))
            kids = list(body)
            body.insert(len(kids) - 1 if kids and kids[-1].tag == f"{W}sectPr" else len(kids), p)
            self._mark("word/document.xml")
            pid = p.get(f"{W14}paraId")
        else:
            pid = res["para_id"]
        doc = self._require("word/document.xml")
        p = self._find_para(doc, pid)
        last = pid
        if design == "fleuron":
            text_w = self._text_width_cm()
            r = self.insert_ornament(
                "divider",
                after_para_id=pid,
                color=color,
                width_cm=text_w * ornament_width_pct / 100,
                height_cm=0.5,
            )
            orn = self._find_para(doc, r["para_id"])
            apply_para_props(
                orn.find(f"{W}pPr"), {"space_before_pt": 0, "space_after_pt": 8, "first_line_cm": 0}
            )
            ppr = p.find(f"{W}pPr")
            apply_para_props(ppr, {"space_after_pt": 2, "keep_with_next": True})
            last = r["para_id"]
        elif design in ("flanked", "bullet"):
            # Background art, not inline pictures: Word copies inline pictures
            # into the table of contents and the running head (STYLEREF).
            from . import islamic_art as ia
            from .bookparts import _bracket_svg, svg_size
            from .fontmetrics import text_width_pt

            st = next(
                x
                for x in self._require("word/styles.xml").findall(f"{W}style")
                if x.get(f"{W}styleId") == self._style_id_for(style, "paragraph")
            )
            sz = st.find(f"{W}rPr/{W}szCs")
            if sz is None:
                sz = st.find(f"{W}rPr/{W}sz")
            size = int(sz.get(f"{W}val")) / 2 if sz is not None else 14.0
            rf = st.find(f"{W}rPr/{W}rFonts")
            font = None
            if rf is not None:
                font = rf.get(f"{W}cs") if self._is_rtl() else rf.get(f"{W}ascii")
            line_pt = size * 1.3
            ppr = p.find(f"{W}pPr")
            sp = ppr.find(f"{W}spacing")
            before = int(sp.get(f"{W}before", "0")) / 20 if sp is not None else 0
            before = before or (18 if level == 1 else 10)
            rtl = self._is_rtl()
            if design == "flanked":
                text_w = text_width_pt(text, font, size, bold=True)
                h_pt = size * 0.95
                left = _bracket_svg("#" + color, "left")
                right = _bracket_svg("#" + color, "right")
                lw, lh = svg_size(left)
                bw_pt = h_pt * lw / lh
                gap_pt = size * 0.6
                total_pt = text_w + 2 * (gap_pt + bw_pt)
                scale = 100 / h_pt
                W_, H_ = total_pt * scale, 100.0
                inner_l = left[left.index(">") + 1 : left.rindex("</svg>")]
                inner_r = right[right.index(">") + 1 : right.rindex("</svg>")]
                k = 100 / lh
                art = ia.svg(
                    W_,
                    H_,
                    f'<g transform="scale({k:.4f})">{inner_l}</g>'
                    f'<g transform="translate({W_ - bw_pt * scale:.2f},0) scale({k:.4f})">'
                    f"{inner_r}</g>",
                )
                w_cm, h_cm = total_pt / 72 * 2.54, h_pt / 72 * 2.54
                self._anchor_behind(
                    "word/document.xml", p, art, w_cm, h_cm, line_pt, before, "Heading brackets"
                )
            else:
                star_pt = size * 0.75
                art = ia.svg(
                    100,
                    100,
                    ia.fill(ia.khatam(46, 50, 50), "#" + color)
                    + ia.fill(ia.circle(13, 50, 50, 24), "#ffffff"),
                )
                s_cm = star_pt / 72 * 2.54
                apply_para_props(ppr, {"indent_start_cm": round(s_cm + 0.25, 2)})
                self._anchor_behind(
                    "word/document.xml", p, art, s_cm, s_cm, line_pt, before, "Heading star"
                )
                an = list(p.iter(f"{{{_WP}}}anchor"))[-1]
                ph = an.find(f"{{{_WP}}}positionH")
                ph.find(f"{{{_WP}}}align").text = "right" if rtl else "left"
                # Arabic letters sit high in the line: lift the star to them
                off = an.find(f"{{{_WP}}}positionV/{{{_WP}}}posOffset")
                off.text = str(int(off.text) - int(size * 0.18 / 72 * 2.54 * 360000))
            self._mark("word/document.xml")
        return {"design": design, "para_id": pid, "after_para_id": last}

    # ── separators ──────────────────────────────────────────────────────

    def insert_separator(
        self,
        after_para_id: str | None = None,
        *,
        text: str = "❁ ❁ ❁",
        ornament: str | None = None,
        color: str | None = None,
    ) -> dict:
        """Section separator after a paragraph: text (❁ ❁ ❁, ✽ ✽ ✽, ⁂) in the
        Separator style, or ornament="divider" / "divider-dots" / "star8"."""
        color = self._c(color, "accent")
        if ornament:
            return self.insert_ornament(ornament, after_para_id=after_para_id, color=color)
        styles = self._require("word/styles.xml")
        sid = None
        for st in styles.findall(f"{W}style"):
            nm = st.find(f"{W}name")
            if nm is not None and nm.get(f"{W}val", "").lower() == "separator":
                sid = st.get(f"{W}styleId")
        if sid is None:
            sid = self.define_style(
                "Separator",
                based_on="Normal",
                aliases=["فاصل"],
                run={"color": color},
                paragraph={"alignment": "center", "space_before_pt": 6, "space_after_pt": 8},
            )["style_id"]
        doc = self._require("word/document.xml")
        p = make_paragraph(self._new_para_id(), {"style": sid})
        # ornament characters (❁ ✽ ⁂ …) in a symbol font: Arabic text fonts
        # lack them and Word would draw substitutes
        import re as _re

        for chunk in _re.findall(
            r"[\u2700-\u27bf\u2600-\u26ff\u2042\u2055\u25c6\u25c7]+|[^\u2700-\u27bf\u2600-\u26ff\u2042\u2055\u25c6\u25c7]+",
            text,
        ):
            if _re.match(r"[\u2700-\u27bf\u2600-\u26ff\u2042\u2055\u25c6\u25c7]", chunk):
                p.append(make_run(chunk, {"font": "Segoe UI Symbol", "cs_font": "Segoe UI Symbol"}))
            else:
                p.append(
                    make_run(
                        chunk,
                        {"rtl": True} if any("\u0600" <= c <= "\u06ff" for c in chunk) else None,
                    )
                )
        if after_para_id:
            anchor = self._find_para(doc, after_para_id)
            if anchor is None:
                raise ValueError(f"Paragraph '{after_para_id}' not found")
            parent = anchor.getparent()
            parent.insert(list(parent).index(anchor) + 1, p)
        else:
            body = doc.find(f"{W}body")
            kids = list(body)
            body.insert(len(kids) - 1 if kids and kids[-1].tag == f"{W}sectPr" else len(kids), p)
        self._mark("word/document.xml")
        return {"para_id": p.get(f"{W14}paraId"), "style": sid}

    # ── power features ──────────────────────────────────────────────────

    def _settings_flag(self, name: str, on: bool, **attrs: str) -> None:
        settings = self._require("word/settings.xml")
        el = settings.find(f"{W}{name}")
        if on:
            if el is None:
                el = etree.SubElement(settings, f"{W}{name}")
            for k, v in attrs.items():
                el.set(f"{W}{k}", v)
        elif el is not None:
            settings.remove(el)
        self._mark("word/settings.xml")

    def restrict_formatting_to_styles(
        self,
        allowed: list[str] | None = None,
        *,
        enabled: bool = True,
        block_theme_switching: bool = True,
    ) -> dict:
        """Word's "Limit formatting to a selection of styles".

        Bold/colour/font buttons are disabled; only the allowed styles can be
        applied, so the book stays consistent. No password: Review ▸ Restrict
        Editing ▸ Stop Protection turns it off.

        Args:
            allowed: Style names or ids to keep usable (default: every style
                defined in the document). Others are locked.
            enabled: False removes the restriction and unlocks styles.
            block_theme_switching: Also block changing the theme / style set.
        """
        styles = self._require("word/styles.xml")
        keep: set[str] | None = None
        if allowed is not None:
            keep = {a.lower() for a in allowed}
        locked = 0
        for st in styles.findall(f"{W}style"):
            name = (
                st.find(f"{W}name").get(f"{W}val") if st.find(f"{W}name") is not None else ""
            ).lower()
            sid = (st.get(f"{W}styleId") or "").lower()
            old = st.find(f"{W}locked")
            lock = enabled and keep is not None and name not in keep and sid not in keep
            if lock and old is None:
                etree.SubElement(st, f"{W}locked")
                locked += 1
            elif not lock and old is not None:
                st.remove(old)
        latent = styles.find(f"{W}latentStyles")
        if latent is not None:
            if enabled and keep is not None:
                latent.set(f"{W}defLockedState", "1")
                for ex in latent.findall(f"{W}lsdException"):
                    nm = (ex.get(f"{W}name") or "").lower()
                    ex.set(f"{W}locked", "0" if nm in keep else "1")
            else:
                latent.set(f"{W}defLockedState", "0")
                for ex in latent.findall(f"{W}lsdException"):
                    if ex.get(f"{W}locked") is not None:
                        del ex.attrib[f"{W}locked"]
        self._mark("word/styles.xml")
        settings = self._require("word/settings.xml")
        prot = settings.find(f"{W}documentProtection")
        if enabled:
            if prot is None:
                prot = etree.SubElement(settings, f"{W}documentProtection")
            prot.set(f"{W}formatting", "1")
            prot.set(f"{W}enforcement", "1")
        elif prot is not None and prot.get(f"{W}edit") in (None, "none"):
            settings.remove(prot)
        elif prot is not None:
            prot.attrib.pop(f"{W}formatting", None)
        self._settings_flag("styleLockTheme", enabled and block_theme_switching)
        self._settings_flag("styleLockQFSet", enabled and block_theme_switching)
        return {
            "enabled": enabled,
            "locked_styles": locked,
            "allowed": sorted(keep) if keep is not None else "all document styles",
        }

    def set_style_pane(
        self,
        show: str = "in-use",
        *,
        sort: str = "priority",
        alternate_names: bool = True,
    ) -> dict:
        """What Word's Styles pane lists and in what order.

        show: in-use, recommended, all, custom (your own styles), headings.
        sort: name, priority, font, based-on, type, default.
        alternate_names: Show the style labels (e.g. «باب», «رقم الحديث»).
        """
        if show not in ("in-use", "recommended", "all", "custom", "headings"):
            raise ValueError("show must be in-use, recommended, all, custom or headings")
        if sort not in _SORTS:
            raise ValueError(f"sort must be one of {sorted(_SORTS)}")
        bits = _PANE_BITS[show] | _PANE_BITS["clear"]
        if alternate_names:
            bits |= _PANE_BITS["alternate-names"]
        self._settings_flag("stylePaneFormatFilter", True, val=f"{bits:04X}")
        self._settings_flag("stylePaneSortMethod", True, val=_SORTS[sort])
        return {"show": show, "sort": sort, "filter": f"{bits:04X}"}

    def set_font_embedding(
        self, embed: bool = True, *, subset: bool = True, include_system_fonts: bool = False
    ) -> dict:
        """Ask Word to embed the fonts in the file when it next saves it.

        Use it before sending a book typeset in Amiri / Jameel Noori Nastaleeq
        to someone who does not have the fonts. subset: only the characters
        used (smaller file; the reader can view and print but not add text in
        that font).
        """
        self._settings_flag("embedTrueTypeFonts", embed)
        self._settings_flag("saveSubsetFonts", embed and subset)
        self._settings_flag("embedSystemFonts", embed and include_system_fonts)
        return {
            "embed": embed,
            "subset": embed and subset,
            "note": "Word embeds the fonts the next time the file is saved in Word.",
        }

    def hide_proofing_marks(self, hide: bool = True) -> dict:
        """Hide spelling and grammar squiggles (they cover Arabic/Urdu text)."""
        self._settings_flag("hideSpellingErrors", hide)
        self._settings_flag("hideGrammaticalErrors", hide)
        return {"hidden": hide}

    def attach_template(self, template_path: str, *, auto_update_styles: bool = True) -> dict:
        """Link the document to a Word template (.dotx/.dotm) on this computer.

        With auto_update_styles, Word copies the template's styles into the
        document every time it opens — one place to change the house style of
        every book.
        """
        settings = self._require("word/settings.xml")
        rels_path, rels = self._rels_for("word/settings.xml")
        rel_type = (
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate"
        )
        for rel in list(rels):
            if rel.get("Type") == rel_type:
                rels.remove(rel)
        uri = template_path.replace("\\", "/")
        if not uri.startswith("file:"):
            uri = "file:///" + uri.lstrip("/")
        rid = add_relationship(rels, rel_type, uri, external=True)
        self._mark(rels_path)
        el = settings.find(f"{W}attachedTemplate")
        if el is None:
            el = etree.SubElement(settings, f"{W}attachedTemplate")
        el.set(f"{R}id", rid)
        self._settings_flag("linkStyles", auto_update_styles)
        self._mark("word/settings.xml")
        return {"template": uri, "auto_update_styles": auto_update_styles}


class EnglishBookMixin:
    """English Islamic book conventions (studied from a Bulugh al-Maram
    translation): bilingual hadith entries, chapter titles in both
    languages, glossary entries."""

    def _has_style(self, style_id: str) -> bool:
        styles = self._tree("word/styles.xml")
        return styles is not None and any(
            s.get(f"{W}styleId") == style_id for s in styles.findall(f"{W}style")
        )

    def insert_hadith_entry(
        self,
        number: int | str,
        arabic: str,
        english: str,
        *,
        after_para_id: str | None = None,
        bold_quotes: bool = True,
        source: str | None = None,
        arabic_number_dash: str = " ـــ ",
    ) -> dict:
        """A numbered hadith: the vocalised Arabic («٩٤ ـــ …», justified
        right-to-left) then the English («94. Narrated …»).

        bold_quotes: the Prophet's words — text inside “…” or "…" in the
        English — in bold, as translations set them. source: appended to the
        English ("Related by Muslim.").
        """
        import re as _re

        digits = "٠١٢٣٤٥٦٧٨٩"
        n = str(number)
        n_ar = "".join(digits[int(c)] if c.isdigit() else c for c in n)
        ar_id = self._style_id_for("Arabic Text", "paragraph")
        tr_id = self._style_id_for("Translation", "paragraph")
        if not (self._has_style(ar_id) and self._has_style(tr_id)):
            raise ValueError(
                "Apply the english or bilingual book preset first "
                "(it defines the Arabic Text and Translation styles)"
            )
        doc = self._require("word/document.xml")
        a = make_paragraph(self._new_para_id(), {"style": ar_id})
        a.append(make_run(n_ar + arabic_number_dash, {"rtl": True, "bold": True}))
        a.append(make_run(arabic, {"rtl": True}))
        e = make_paragraph(self._new_para_id(), {"style": tr_id})
        # LRM fences the number: Word's "context" numerals would otherwise
        # draw it in Arabic-Indic after the Arabic paragraph
        e.append(make_run(f"\u200e{n}.\u2002", {"bold": False}))
        text = english + (f" {source}" if source else "")
        if bold_quotes:
            for chunk in _re.split(r"(“[^”]+”|\"[^\"]+\")", text):
                if chunk:
                    e.append(make_run(chunk, {"bold": True} if chunk[:1] in '“"' else None))
        else:
            e.append(make_run(text, None))
        if after_para_id:
            anchor = self._find_para(doc, after_para_id)
            if anchor is None:
                raise ValueError(f"Paragraph '{after_para_id}' not found")
            anchor.addnext(a)
        else:
            body = doc.find(f"{W}body")
            kids = list(body)
            body.insert(len(kids) - 1 if kids and kids[-1].tag == f"{W}sectPr" else len(kids), a)
        a.addnext(e)
        self._mark("word/document.xml")
        return {"arabic_para_id": a.get(f"{W14}paraId"), "para_id": e.get(f"{W14}paraId")}

    def insert_bilingual_heading(
        self,
        arabic: str,
        english: str,
        *,
        level: int = 1,
        after_para_id: str | None = None,
    ) -> dict:
        """A chapter title in both languages: the vocalised Arabic title over
        the English heading («بَابُ الْمَوَاقِيتِ» / "Chapter II: Mawâqît"). The
        English line carries the Heading style, so it reaches the contents."""
        ah = self._style_id_for("Arabic Heading", "paragraph")
        if not self._has_style(ah):
            raise ValueError("Apply the english or bilingual book preset first")
        doc = self._require("word/document.xml")
        a = make_paragraph(self._new_para_id(), {"style": ah})
        a.append(make_run(arabic, {"rtl": True}))
        h = make_paragraph(
            self._new_para_id(),
            {"style": self._style_id_for(f"Heading {level}", "paragraph"), "direction": "ltr"},
        )
        h.append(make_run(english, None))
        if after_para_id:
            anchor = self._find_para(doc, after_para_id)
            if anchor is None:
                raise ValueError(f"Paragraph '{after_para_id}' not found")
            anchor.addnext(a)
        else:
            body = doc.find(f"{W}body")
            kids = list(body)
            body.insert(len(kids) - 1 if kids and kids[-1].tag == f"{W}sectPr" else len(kids), a)
        a.addnext(h)
        self._mark("word/document.xml")
        return {"arabic_para_id": a.get(f"{W14}paraId"), "para_id": h.get(f"{W14}paraId")}

    def insert_glossary_entry(
        self, term: str, definition: str, *, after_para_id: str | None = None
    ) -> dict:
        """A glossary line: the term in bold italic, a colon, the definition
        in italic ("Ghusl: The full ritual washing …")."""
        gid = self._style_id_for("Glossary Entry", "paragraph")
        if not self._has_style(gid):
            raise ValueError("Apply the english or bilingual book preset first")
        doc = self._require("word/document.xml")
        p = make_paragraph(self._new_para_id(), {"style": gid})
        p.append(
            make_run(term + ":", {"char_style": self._style_id_for("Glossary Term", "character")})
        )
        p.append(make_run(" " + definition, None))
        if after_para_id:
            anchor = self._find_para(doc, after_para_id)
            if anchor is None:
                raise ValueError(f"Paragraph '{after_para_id}' not found")
            anchor.addnext(p)
        else:
            body = doc.find(f"{W}body")
            kids = list(body)
            body.insert(len(kids) - 1 if kids and kids[-1].tag == f"{W}sectPr" else len(kids), p)
        self._mark("word/document.xml")
        return {"para_id": p.get(f"{W14}paraId")}
