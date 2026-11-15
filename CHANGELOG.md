# Changelog (fork)

Changes in this fork on top of upstream `SecurityRonin/docx-mcp` v0.7.4.

## Unreleased

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
