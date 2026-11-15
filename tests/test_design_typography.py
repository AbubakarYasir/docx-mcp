"""Styles, rich text, honorific glyphs and themes."""

from __future__ import annotations

import zipfile
from pathlib import Path

import pytest
from lxml import etree

from docx_mcp import server
from docx_mcp.document import DocxDocument
from docx_mcp.document.honorifics import build_matcher, find_matches

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def _new(tmp_path: Path) -> tuple[DocxDocument, str]:
    doc = DocxDocument.create(str(tmp_path / "t.docx"))
    first = doc._trees["word/document.xml"].find(f".//{W}p").get(f"{W14}paraId")
    return doc, first


def _text(doc) -> str:
    return doc.get_body_text()["body"]


def _style(doc, sid):
    return doc._trees["word/styles.xml"].find(f"{W}style[@{W}styleId='{sid}']")


class TestHonorifics:
    @pytest.mark.parametrize(
        ("text", "glyph"),
        [
            ("صلى الله عليه وسلم", "ﷺ"),
            ("صَلَّى اللهُ عَلَيْهِ وَسَلَّمَ", "ﷺ"),
            ("(ص)", "ﷺ"),
            ("(صلعم)", "ﷺ"),
            ("رضي الله عنه", "﵁"),
            ("رضي الله عنها", "﵂"),
            ("رضي الله عنهم", "﵃"),
            ("رضي الله عنهما", "﵄"),
            ("(رض)", "﵁"),
            ("رحمه الله", "﵀"),
            ("رحمهم الله", "﵏"),
            ("عليه السلام", "﵇"),
            ("عليه الصلاة والسلام", "﵊"),
            ("عز وجل", "﷿"),
            ("جل جلاله", "ﷻ"),
            ("سبحانه وتعالى", "﷾"),
            ("تبارك وتعالى", "﵎"),
            ("[رضي الله عنه]", "﵁"),
        ],
    )
    def test_phrase_to_glyph(self, text, glyph):
        spans = find_matches(f"قال {text} ثم", build_matcher())
        assert [g for _, _, g in spans] == [glyph]

    def test_longest_phrase_wins(self):
        spans = find_matches("ابن عمر رضي الله عنهما", build_matcher())
        assert spans[0][2] == "﵄"

    def test_bare_abbreviation_is_not_replaced(self):
        # «ص» alone is a letter (e.g. a page abbreviation), not an honorific.
        assert find_matches("انظر ص 12", build_matcher()) == []

    def test_not_inside_longer_word(self):
        assert find_matches("عليه السلامة", build_matcher()) == []

    def test_abbreviations_can_be_disabled(self):
        assert find_matches("عمر (رض)", build_matcher(include_abbreviations=False)) == []

    def test_apply_in_document_with_font_and_colour(self, tmp_path):
        doc, first = _new(tmp_path)
        doc.insert_paragraph(first, "قال رسول الله صلى الله عليه وسلم وعمر (رض).")
        res = doc.apply_honorifics(color="520912")
        assert res["replaced"] == 2 and res["font_has_full_set"]
        assert _text(doc).startswith("قال رسول الله ﷺ وعمر ﵁")
        glyph_runs = [
            r for r in doc._trees["word/document.xml"].iter(f"{W}r")
            if (r.findtext(f"{W}t") or "") in ("ﷺ", "﵁")
        ]
        assert len(glyph_runs) == 2
        for r in glyph_runs:
            assert r.find(f"{W}rPr/{W}rFonts").get(f"{W}cs") == "Scheherazade New"
            assert r.find(f"{W}rPr/{W}color").get(f"{W}val") == "520912"

    def test_footnotes_too(self, tmp_path):
        doc, first = _new(tmp_path)
        pid = doc.insert_paragraph(first, "نص")["para_id"]
        doc.add_footnote(pid, "رواه أنس رضي الله عنه.")
        doc.apply_honorifics()
        notes = "".join(doc._trees["word/footnotes.xml"].itertext())
        assert "﵁" in notes

    def test_existing_glyphs_restyled(self, tmp_path):
        doc, first = _new(tmp_path)
        doc.insert_paragraph(first, "قال النبي ﷺ ذلك")
        res = doc.apply_honorifics(color="520912")
        assert res["glyph_runs_formatted"] >= 1


class TestStyles:
    def test_define_creates_then_updates(self, tmp_path):
        doc, _ = _new(tmp_path)
        a = doc.define_style("Matn", run={"cs_font": "Amiri", "size_pt": 14},
                             paragraph={"alignment": "kashida", "direction": "rtl"},
                             aliases=["متن"], gallery=True, ui_priority=1)
        assert a["created"]
        b = doc.define_style("متن", run={"color": "333333"})  # found by alias
        assert not b["created"] and b["style_id"] == "Matn"
        st = _style(doc, "Matn")
        assert st.find(f"{W}aliases").get(f"{W}val") == "متن"
        assert st.find(f"{W}pPr/{W}jc").get(f"{W}val") == "mediumKashida"
        assert st.find(f"{W}pPr/{W}bidi") is not None
        rpr = st.find(f"{W}rPr")
        assert rpr.find(f"{W}rFonts").get(f"{W}cs") == "Amiri"
        assert rpr.find(f"{W}szCs").get(f"{W}val") == "28"
        assert rpr.find(f"{W}color").get(f"{W}val") == "333333"
        assert st.find(f"{W}qFormat") is not None

    def test_type_mismatch_and_bad_keys(self, tmp_path):
        doc, _ = _new(tmp_path)
        doc.define_style("X", "character", run={"bold": True})
        with pytest.raises(ValueError):
            doc.define_style("X", "paragraph")
        with pytest.raises(ValueError):
            doc.define_style("Y", run={"colour": "red"})
        with pytest.raises(ValueError):
            doc.define_style("Z", "character", paragraph={"alignment": "center"})

    def test_rename_updates_references(self, tmp_path):
        doc, first = _new(tmp_path)
        doc.define_style("Old Name")
        pid = doc.insert_paragraph(first, "x", style="OldName")["para_id"]
        res = doc.rename_style("Old Name", "New Name", new_id=True)
        assert res["references_updated"] == 1
        p = doc._find_para_any(pid)[1]
        assert p.find(f"{W}pPr/{W}pStyle").get(f"{W}val") == "NewName"

    def test_organize_keeps_base_styles_visible(self, tmp_path):
        doc, first = _new(tmp_path)
        doc.define_style("Chapter", based_on="Heading1")
        doc.insert_paragraph(first, "x", style="Chapter")
        res = doc.organize_styles(["Chapter"], hide_unused=True)
        assert "Heading1" not in res["hidden"]
        assert _style(doc, "Chapter").find(f"{W}uiPriority").get(f"{W}val") == "1"

    def test_apply_style_where_paragraph_and_character(self, tmp_path):
        doc, first = _new(tmp_path)
        doc.define_style("Quote", "character", run={"bold": True})
        doc.define_style("Narration")
        doc.insert_paragraph(first, "[١] قال: «إنما الأعمال بالنيات» تم")
        doc.insert_paragraph(first, "سطر آخر")
        assert doc.apply_style_where("Narration", text_regex=r"^\[")["applied"] == 1
        assert doc.apply_style_where("Quote", text_regex=r"«[^»]+»")["applied"] == 1
        styled = [
            r.findtext(f"{W}t") for r in doc._trees["word/document.xml"].iter(f"{W}r")
            if r.find(f"{W}rPr/{W}rStyle") is not None
        ]
        assert styled == ["«إنما الأعمال بالنيات»"]

    def test_style_report_and_clear(self, tmp_path):
        doc, first = _new(tmp_path)
        for i in range(3):
            pid = doc.insert_paragraph(first, f"p{i}")["para_id"]
            doc.set_run_size(pid, 0, 18)
        rep = doc.style_report()
        assert rep["paragraphs_with_direct_formatting"] == 3
        assert rep["repeated_direct_formatting_make_a_style"]
        assert doc.clear_direct_formatting()["properties_removed"] >= 6
        assert doc.style_report()["paragraphs_with_direct_formatting"] == 0


class TestRichTextAndThemes:
    def test_rich_paragraph_spans_and_fields(self, tmp_path):
        doc, first = _new(tmp_path)
        res = doc.insert_rich_paragraph(
            first,
            [{"text": "A", "font": "Cambria", "italic": True},
             {"text": "ب", "cs_font": "Amiri", "color": "#1F4E79"},
             {"text": " page {PAGE} of {NUMPAGES}"}],
            paragraph={"alignment": "center"},
        )
        p = doc._find_para_any(res["para_id"])[1]
        instr = [t.text.strip() for t in p.iter(f"{W}instrText")]
        assert instr == ["PAGE", "NUMPAGES"]
        colours = [c.get(f"{W}val") for c in p.iter(f"{W}color")]
        assert "1F4E79" in colours

    def test_bad_colour_rejected(self, tmp_path):
        doc, first = _new(tmp_path)
        with pytest.raises(ValueError):
            doc.insert_rich_paragraph(first, [{"text": "x", "color": "maroon"}])

    def test_format_matches(self, tmp_path):
        doc, first = _new(tmp_path)
        doc.insert_paragraph(first, "قال تعالى ﴿الحمد لله﴾ و ﴿رب العالمين﴾")
        res = doc.format_matches("﴿[^﴾]+﴾", regex=True, cs_font="Amiri Quran", color="2F5233")
        assert res["formatted"] == 2

    def test_theme_created_and_saved(self, tmp_path):
        doc, _ = _new(tmp_path)
        doc.set_theme_fonts(heading_arabic="Sakkal Majalla", body_arabic="Amiri")
        doc.apply_color_palette("emerald-gold", {"accent6": "#123456"})
        doc.set_theme_color("accent2", "AA8833")
        doc.save()
        with zipfile.ZipFile(tmp_path / "t.docx") as zf:
            theme = etree.fromstring(zf.read("word/theme/theme1.xml"))
            rels = zf.read("word/_rels/document.xml.rels").decode()
        assert "theme/theme1.xml" in rels
        minor = theme.find(f".//{A}minorFont")
        assert [f.get("typeface") for f in minor.findall(f"{A}font")] == ["Amiri"]
        scheme = theme.find(f".//{A}clrScheme")
        assert scheme.find(f"{A}accent6/{A}srgbClr").get("val") == "123456"
        assert scheme.find(f"{A}accent2/{A}srgbClr").get("val") == "AA8833"

    def test_server_tools(self, tmp_path):
        server.create_document(str(tmp_path / "s.docx"))
        server.define_style("Matn", run={"cs_font": "Amiri"})
        assert '"replaced"' in server.apply_honorifics()
        assert "scholarly-maroon" in server.apply_color_palette("scholarly-maroon")
