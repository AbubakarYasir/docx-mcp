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

from .ooxml import W14, W, add_relationship, make_paragraph, make_run

R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

PRESETS = {
    "tahqiq": {
        "label": "Tahqiq / risalah (house style)",
        "page_size": "17x24",
        "margins_cm": (1.8, 1.8, 1.8, 1.8),  # top, bottom, inside, outside
        "body_font": "Amiri",
        "heading_font": "Sakkal Majalla",
        "body_pt": 13,
        "line": 1.05,
        "chapter_pt": 20,
        "section_pt": 16,
        "note_pt": 9,
        "note_line_pt": 15,
        "accent": "520912",
        "ink": "000000",
        "separator": "❁ ❁ ❁",
        "chapters_new_page": False,
        "honorific_color": None,
    },
    "kitab": {
        "label": "Full book: chapters on new pages, larger type",
        "page_size": "17x24",
        "margins_cm": (2.0, 2.0, 2.2, 1.6),
        "body_font": "Amiri",
        "heading_font": "Sakkal Majalla",
        "body_pt": 14,
        "line": 1.15,
        "chapter_pt": 24,
        "section_pt": 17,
        "note_pt": 10,
        "note_line_pt": 16,
        "accent": "520912",
        "ink": "000000",
        "separator": "❁ ❁ ❁",
        "chapters_new_page": True,
        "honorific_color": None,
    },
    "urdu-nastaliq": {
        "label": "Urdu book in Jameel Noori Nastaleeq",
        "page_size": "17x24",
        "margins_cm": (2.0, 2.0, 2.0, 1.8),
        "body_font": "Jameel Noori Nastaleeq",
        "heading_font": "Jameel Noori Nastaleeq",
        "body_pt": 15,
        "line": 1.5,
        "chapter_pt": 24,
        "section_pt": 18,
        "note_pt": 11,
        "note_line_pt": 20,
        "accent": "520912",
        "ink": "000000",
        "separator": "❁ ❁ ❁",
        "chapters_new_page": True,
        "honorific_color": None,
        "language": "ur-PK",
        "glyph_set": "jameel",
    },
}

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
        running_head: str = "title",
        chapters_new_page: bool | None = None,
        justify: str = "justify",
        honorifics: bool = True,
        note_continuation: bool = True,
        page_numbers: str = "decimal",
        title_page: bool = True,
        overrides: dict[str, Any] | None = None,
    ) -> dict:
        """Set up a whole Arabic/Urdu book in one call.

        Args:
            preset: "tahqiq" (Amiri 13 / Sakkal Majalla, 17×24, maroon
                #520912, 9 pt notes on exact 15 pt lines, chapters run on),
                "kitab" (chapters on new pages, larger type) or
                "urdu-nastaliq" (Jameel Noori Nastaleeq, Urdu).
            title / author: Written to the document properties; the title is
                the running head.
            running_head: "title" (book title on every page), "chapter" (the
                current chapter via STYLEREF), "title-chapter" (title on even
                pages, chapter on odd), or "none".
            chapters_new_page: Override the preset.
            justify: "justify" or "kashida" (stretch with tatweel, as in print).
            honorifics: Replace phrases with honorific glyphs.
            note_continuation: "=" at the foot of a page where a note runs on.
            page_numbers: decimal, arabic-abjad, arabic-alpha …
            title_page: The first page is a title page: no running head or
                folio on it.
            overrides: Any preset value, e.g. {"body_pt": 14, "accent": "1B4D3E"}.

        Styles created (English name, with Arabic/Urdu label): Normal (متن),
        Title (العنوان), Heading 1 (باب / chapter), Heading 2 (ترجمة / section),
        Heading 3, Narration Number (رقم الحديث), Prophetic Speech (كلام
        النبي ﷺ), Matn (نص الحديث), Isnad (الإسناد), Quran (آية), Separator
        (فاصل), footnote text (الحاشية).
        """
        if preset not in PRESETS:
            raise ValueError(f"preset must be one of {sorted(PRESETS)}")
        if running_head not in ("title", "chapter", "title-chapter", "none"):
            raise ValueError("running_head must be title, chapter, title-chapter or none")
        if justify not in ("justify", "kashida"):
            raise ValueError("justify must be justify or kashida")
        cfg = {**PRESETS[preset], **(overrides or {})}
        if chapters_new_page is not None:
            cfg["chapters_new_page"] = chapters_new_page
        body, head = cfg["body_font"], cfg["heading_font"]
        accent, ink = cfg["accent"], cfg["ink"]
        lang = cfg.get("language", "ar-SA")
        done: list[str] = []

        top, bottom, inside, outside = cfg["margins_cm"]
        self.set_book_layout(
            page_size=cfg["page_size"],
            top_cm=top,
            bottom_cm=bottom,
            inside_cm=inside,
            outside_cm=outside,
            mirror_margins=True,
            rtl_binding=True,
            header_cm=1.0,
            footer_cm=0.9,
        )
        done.append("page")
        self.set_document_direction(
            "rtl", complex_script_font=body, complex_script_size_pt=cfg["body_pt"], language=lang
        )
        self.set_theme_fonts(heading_arabic=head, body_arabic=body)
        done.append("direction+theme")

        run_body = {
            "font": body,
            "cs_font": body,
            "size_pt": cfg["body_pt"],
            "color": ink,
            "cs_language": lang,
        }
        self.define_style(
            "Normal",
            run=run_body,
            aliases=["متن"],
            paragraph={
                "direction": "rtl",
                "alignment": "kashida" if justify == "kashida" else "justify",
                "line_spacing": cfg["line"],
                "space_after_pt": 4,
                "widow_control": True,
            },
        )
        hrun = {"font": head, "cs_font": head, "bold": True, "color": ink}
        self.define_style(
            "Title",
            based_on="Normal",
            next_style="Normal",
            aliases=["العنوان"],
            gallery=True,
            ui_priority=1,
            run={**hrun, "size_pt": cfg["chapter_pt"] * 2},
            paragraph={
                "alignment": "center",
                "space_before_pt": 0,
                "space_after_pt": 18,
                "keep_with_next": True,
                "border_top": {"style": "double", "color": accent, "size_pt": 0.75, "space_pt": 16},
                "border_bottom": {
                    "style": "double",
                    "color": accent,
                    "size_pt": 0.75,
                    "space_pt": 16,
                },
            },
        )
        self.define_style(
            "Heading 1",
            based_on="Normal",
            next_style="Normal",
            aliases=["باب"],
            gallery=True,
            ui_priority=2,
            run={**hrun, "size_pt": cfg["chapter_pt"]},
            paragraph={
                "alignment": "center",
                "space_before_pt": 12,
                "space_after_pt": 14,
                "keep_with_next": True,
                "keep_together": True,
                "outline_level": 0,
                "page_break_before": bool(cfg["chapters_new_page"]),
                "line_spacing": 1.0,
                "border_top": {"style": "double", "color": accent, "size_pt": 0.75, "space_pt": 10},
                "border_bottom": {
                    "style": "double",
                    "color": accent,
                    "size_pt": 0.75,
                    "space_pt": 10,
                },
            },
        )
        self.define_style(
            "Heading 2",
            based_on="Normal",
            next_style="Normal",
            aliases=["ترجمة"],
            gallery=True,
            ui_priority=3,
            run={**hrun, "size_pt": cfg["section_pt"]},
            paragraph={
                "alignment": "start",
                "space_before_pt": 3,
                "space_after_pt": 6,
                "keep_with_next": True,
                "outline_level": 1,
                "line_spacing": 1.0,
            },
        )
        self.define_style(
            "Heading 3",
            based_on="Normal",
            next_style="Normal",
            aliases=["فصل"],
            gallery=True,
            ui_priority=4,
            run={**hrun, "size_pt": cfg["body_pt"] + 1},
            paragraph={
                "alignment": "start",
                "space_before_pt": 3,
                "space_after_pt": 3,
                "keep_with_next": True,
                "outline_level": 2,
            },
        )
        self.define_style(
            "Narration Number",
            "character",
            aliases=["رقم الحديث"],
            gallery=True,
            ui_priority=5,
            run={"bold": True, "color": accent},
        )
        self.define_style(
            "Prophetic Speech",
            "character",
            aliases=["كلام النبي ﷺ"],
            gallery=True,
            ui_priority=6,
            run={"bold": True, "color": accent},
        )
        self.define_style(
            "Matn",
            "character",
            aliases=["نص الحديث"],
            gallery=True,
            ui_priority=7,
            run={"bold": True, "color": ink},
        )
        self.define_style(
            "Isnad",
            "character",
            aliases=["الإسناد"],
            gallery=True,
            ui_priority=8,
            run={"color": ink},
        )
        self.define_style(
            "Quran",
            "character",
            aliases=["آية"],
            gallery=True,
            ui_priority=9,
            run={"bold": True, "color": ink},
        )
        self.define_style(
            "Separator",
            based_on="Normal",
            next_style="Normal",
            aliases=["فاصل"],
            gallery=True,
            ui_priority=10,
            run={"color": accent, "size_pt": 11},
            paragraph={
                "alignment": "center",
                "space_before_pt": 6,
                "space_after_pt": 8,
                "keep_with_next": False,
                "line_spacing": 1.0,
            },
        )
        self.define_style(
            "footnote text",
            based_on="Normal",
            aliases=["الحاشية"],
            run={"font": body, "cs_font": body, "size_pt": cfg["note_pt"], "color": ink},
            paragraph={
                "direction": "rtl",
                "alignment": "justify",
                "line_exact_pt": cfg["note_line_pt"],
                "space_after_pt": 1,
                "space_before_pt": 0,
            },
        )
        done.append("styles")

        hf = {"font": head, "cs_font": head, "color": ink}
        rule = {"rule": "below", "rule_color": accent, "rule_size_pt": 0.5}
        book_title = title or "{TITLE}"
        if running_head == "title":
            self.design_header_footer("header", center=book_title, size_pt=12, **hf, **rule)
        elif running_head == "chapter":
            self.design_header_footer(
                "header", center="{STYLEREF:Heading 1}", size_pt=12, **hf, **rule
            )
        elif running_head == "title-chapter":
            self.design_header_footer(
                "header", center="{STYLEREF:Heading 1}", size_pt=12, **hf, **rule
            )
            self.design_header_footer(
                "header", center=book_title, variant="even", size_pt=12, **hf, **rule
            )
        self.design_header_footer("footer", center="— {PAGE} —", size_pt=13, **hf)
        if title_page:
            self.design_header_footer("header", center="", variant="first")
            self.design_header_footer("footer", center="", variant="first")
        if page_numbers != "decimal":
            self.set_page_numbering(number_format=page_numbers)
        done.append("header/footer")

        self.set_note_separator("footnote", "full", color=accent, thickness_pt=0.75)
        if note_continuation:
            self.set_note_continuation("footnote", color=accent, bold=True)
        done.append("footnotes")

        if honorifics:
            self.apply_honorifics(
                glyph_set=cfg.get("glyph_set", "unicode"), color=cfg.get("honorific_color")
            )
            done.append("honorifics")
        if title or author:
            props = {}
            if title:
                props["title"] = title
            if author:
                props["creator"] = author
            self.set_properties(**props)
        self.hide_proofing_marks(True)
        self.set_style_pane(show="in-use", sort="priority", alternate_names=True)
        self.set_update_fields_on_open(True)
        done.append("settings")
        return {"preset": preset, "applied": done, "config": cfg}

    # ── separators ──────────────────────────────────────────────────────

    def insert_separator(
        self,
        after_para_id: str | None = None,
        *,
        text: str = "❁ ❁ ❁",
        ornament: str | None = None,
        color: str = "520912",
    ) -> dict:
        """Section separator after a paragraph: text (❁ ❁ ❁, ✽ ✽ ✽, ⁂) in the
        Separator style, or ornament="divider" / "divider-dots" / "star8"."""
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
        p.append(make_run(text, None))
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
