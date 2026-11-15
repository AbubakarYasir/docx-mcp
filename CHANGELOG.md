# Changelog (fork)

Changes in this fork on top of upstream `SecurityRonin/docx-mcp` v0.7.4.

## Unreleased

### Added: vector art

- `insert_svg`: SVG artwork as a native vector picture (sharp at any zoom and
  in print in Word 2016+), stored with a PNG fallback rendered by resvg for
  older Word, LibreOffice and previews. New dependency: `resvg-py` (prebuilt
  wheels for Windows, macOS and Linux).
- `insert_ornament`: generated Islamic-geometric ornaments — khatam
  (eight-point star), rosette, shamsa medallion, divider, dotted divider,
  corner, band — in any colour with a tinted fill.
- `add_page_frame`: a double frame with khatam corners on every page, drawn
  behind the text from the header so it repeats without moving the body;
  existing header content is kept.
- `set_page_border`: Word's line and art page borders (celtic, weave,
  zigzag …), all pages / first / not first.
- Checked in Microsoft Word: ornaments and frame export to PDF as vectors.

### Added: indexes and tables of contents

- `mark_index_entries`: marks every occurrence of names, places or phrases as
  real index entries (XE fields), so Word's Update Index always has correct
  pages. Arabic-aware: «والبخاري» and «بمكة» are filed under البخاري and مكة.
  Sub-entries, "انظر" cross-references, bold/italic page numbers for the main
  discussion. `mark_index_entry`, `list_index_entries` (Arabic sorting that
  ignores harakat and hamza forms), `remove_index_entries`.
- Several indexes in one book through index types: one letter per index
  (e.g. P أعلام, L أماكن, H أطراف الأحاديث).
- `insert_index`: title on a new page, 1–4 columns, letter headings in any
  colour, dotted leaders to right-aligned page numbers, Arabic sorting.
- `insert_toc`: table of contents from heading levels and/or your own styles
  ("Chapter" → 1, "Bab" → 2), clickable, dotted leaders, coloured first level.
- INDEX and TOC are written as one field whose cached entries sit inside it,
  so updating replaces them. Checked in Microsoft Word: TOC with `\t` and
  `\u` together finds nothing, so `\u` is used only for heading-level TOCs.

### Added: footnote and endnote design

- `set_note_continuation`: a mark ("=" by Arabic convention, or "(continued)")
  at the foot of every page where a note runs on, in any font, size and
  colour, aligned to the reading direction. Checked in Microsoft Word: the
  notice prints in every layout mode. Text above the continued part prints
  only in Word 2010 layout (Word 2013+ ignores the continuation separator and
  reuses the normal separator), so it is opt-in with `legacy_layout=True` and
  the tool says what will actually print.
- `set_note_separator`: Word's short rule, a full-width rule in any colour and
  thickness, an ornament ("⁂", "✽ ✽ ✽"), or none.
- `set_note_numbering`: Arabic abjad / alphabetical, roman, letters, symbols,
  circled and parenthesised numbers; restart every page or section; start
  number; footnotes below text, endnotes at section or document end. Written
  to settings and every section so Word cannot disagree.
- `add_custom_mark_note`: a note opened by your own mark (a narration number
  "[١]", "*", "(أ)") instead of an automatic number, at the start, end, or
  after given text, with separate formatting in the body and in the note.
- `convert_notes`: footnotes to endnotes and back, with styles and references.

### Fixed

- Saving no longer deletes separator and continuation notes as "orphans"
  (they have no body reference by design), and footnote listings skip them.

### Added: page design

- `design_header_footer`: headers/footers from start/centre/end zones with
  live fields (`{PAGE}`, `{NUMPAGES}`, `{SECTIONPAGES}`, `{DATE}`, `{TITLE}`,
  `{STYLEREF:…}` running heads, any `{FIELD:…}`), first-page and even-page
  variants (Word settings switched on automatically), rule lines above/below,
  per-zone formatting, distance from the edge. With odd/even pages on, a
  header or footer without an even variant is reused on even pages instead
  of leaving them blank. `remove_header_footer`.
- `set_page_numbering`: decimal, roman, letters, Arabic abjad / alphabetical,
  "- 1 -", with restart.
- `set_book_layout`: page presets (A4, A5, B5, 17×24, 14×21, 16×23, royal,
  letter), inside/outside margins, gutter, mirror margins, right-hand binding
  for right-to-left books, header/footer distance.
- `set_update_fields_on_open`: Word refreshes TOC, index and running heads.

### Added: glyph sets for honorifics

- `apply_honorifics(glyph_set="unicode" | "jameel")`. Jameel Noori Nastaleeq
  draws its own honorific artwork at non-standard code points and reuses
  U+FDFE for حفظه الله (Unicode: سبحانه وتعالى), so a glyph is always written
  with the font of its set; a phrase missing from the chosen set uses the
  other set in that set's font. New phrases: حفظه الله / حفظها الله /
  حفظهم الله, رحمة الله عليه / عليها / عليهم / عليهما / عليهن. Existing glyphs
  are recoloured but never re-fonted. Longest phrase always wins.

### Fixed

- New documents opened in Word's Compatibility Mode (Word 2007 layout engine)
  because settings.xml was empty. The template now sets compatibility mode 15
  and Word's standard defaults.
- settings.xml, footnotePr/endnotePr and pgBorders children are kept in schema
  order.
- `set_custom_property` created docProps/custom.xml without linking it in the
  package relationships and content types (Word saw an unknown stray file).

### Added: styles, typography, honorific glyphs and themes

- **Style designer** (`styledesign.py`): `define_style` creates or updates a
  complete paragraph/character/table style in one call — Latin and Arabic
  fonts and sizes, bold/italic for both scripts, colour, spacing (multiple,
  exact, at-least), start/end indents, kashida justification, keep-with-next,
  outline level, borders, shading, tabs, labels (aliases) and gallery
  placement. `rename_style` (optionally updates every reference),
  `organize_styles` (gallery order; hide unused styles but keep the styles
  used ones are based on), `style_report` (usage, and repeated direct
  formatting that should become a style), `apply_style_where` (paragraph
  styles by pattern/current style, character styles on regex matches),
  `clear_direct_formatting`.
- **Rich text** (`typography.py`): `insert_rich_paragraph` / `append_rich_text`
  build paragraphs from spans with their own fonts, sizes and colours, with
  `{PAGE}`, `{NUMPAGES}`, `{STYLEREF:…}` and `{FIELD:…}` tokens.
  `format_matches` formats every occurrence of a text or regex across the
  body, notes, headers and footers.
- **Honorifics** (`honorifics.py`): `apply_honorifics` replaces honorific
  phrases (ignoring tashkeel and spelling variants) and bracketed
  abbreviations — (ص) (صلعم) (رض) (رح) (ع) — with their full glyphs, each
  in a font that draws it (Scheherazade New / Noto Naskh Arabic; Amiri,
  Traditional Arabic and Sakkal Majalla only draw ﷺ) and an optional colour
  and size. `list_honorifics` lists them.
- **Themes**: `set_theme_fonts` (heading/body, Latin and the Arabic script
  slot) and `apply_color_palette` (four book palettes or explicit slots).
  Documents without a theme get one, so `set_theme_color` no longer fails.

### Fixed

- Saving removed the relationship to any part added in the same session
  before it was written to disk (it looked "broken").

### Added: your name on tracked changes

- `DOCX_MCP_AUTHOR` sets the name Word shows for tracked changes and comments
  when a tool is called without an explicit author (default stays "Claude").

### Added: Arabic and right-to-left support

Word does not work out direction from the text: an Arabic paragraph is only
right-to-left when it is marked so, and bold, italic, size and font only reach
Arabic letters through their complex-script twins (`bCs`, `iCs`, `szCs`,
`rFonts/@cs`). Upstream set none of these, so Arabic came out left-aligned,
with punctuation on the wrong side and formatting silently ignored.

- **Automatic, on save** (`docx_mcp/document/bidi.py`), for every paragraph
  docx-mcp creates, in the body, footnotes, endnotes, headers, footers and
  comments:
  - direction from the first strong letter, as Word itself does;
  - runs mixing Arabic and English are split so each piece is marked;
  - right-to-left runs get `w:rtl`, `bCs`/`iCs`/`szCs` twins, and the Arabic
    proofing language (`ar-SA`, change with `DOCX_MCP_RTL_LANG`);
  - new tables whose text starts in Arabic get right-to-left columns;
  - text added to an existing right-to-left paragraph is marked too.
  Paragraphs that came with the original file keep their direction.
  Turn the automatic layer off with `DOCX_MCP_AUTO_BIDI=0`.
- **New tools:** `set_document_direction` (whole document, with Arabic font,
  size and language), `set_paragraph_direction`, `set_table_direction`,
  `get_direction_report` (finds Arabic text Word will render wrongly).
- Built-in heading and other styles carry `bCs`/`szCs`, so Arabic headings are
  bold and sized like Latin ones.

### Fixed: formatting revisions

- Accepting changes left `w:rPrChange` / `w:pPrChange` records behind (Word and
  LibreOffice still showed the text as changed). Accept now drops them; reject
  restores the previous formatting.

### Performance

- New paragraph ids are drawn from a cached set instead of rescanning every
  XML part each time; long documents no longer slow down quadratically.

### Fixed: files Word may reject or "repair"

Found by validating output from every tool against the ECMA-376 schemas.

- **Element order.** Word requires the children of `w:pPr`, `w:rPr`,
  `w:sectPr`, `w:tblPr`, `w:tcPr`, `w:trPr`, `w:sdtPr`, `w:style`, border
  groups and `w:numbering` in a fixed order. Many tools appended new settings
  at the end. A new step on every save (`schema_order.py`) restores the order
  for all parts, so every tool is covered at once.
- **Blank documents** now include a section (A4, 2.54 cm margins). Before, new
  documents had no page setup, which also made `set_section_columns` crash.
- **Page margins.** Setting one margin wrote an incomplete `w:pgMar`; the
  missing required attributes are now filled in. Changing only the orientation
  now swaps width and height.
- **Images.** Inserted pictures were missing required parts (`wp:docPr`,
  `pic:nvPicPr`). Images now keep their aspect ratio (native size, fitted to a
  16 cm text width, or derived from one given side), accept alt text, and a new
  image can no longer overwrite an existing media file.
- **Tracked new tables** put `w:ins` inside `w:tblPr`, which is invalid. Rows
  and cell paragraph marks are now marked instead.
- **Charts.** Bar and line charts had no axes; pie charts had children out of
  order. Updating chart data keeps the order valid.

### Fixed: accepting and rejecting tracked changes

- Rejecting a tracked table, row or inserted paragraph left it in the document.
  One routine now handles text, paragraph-mark and row revisions in the body,
  footnotes, endnotes, headers and footers, for all accept/reject tools.
- `reject_all_changes` never marked the document as changed, so the rejection
  was lost on save.

### Fixed: compatibility

- Works with `mcp` 2.x (where `FastMCP` became `MCPServer`) as well as 1.x, so
  the `--with "mcp<2"` workaround is no longer needed.
