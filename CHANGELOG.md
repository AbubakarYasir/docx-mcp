# Changelog (fork)

Changes in this fork on top of upstream `SecurityRonin/docx-mcp` v0.7.4.

## Unreleased

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
