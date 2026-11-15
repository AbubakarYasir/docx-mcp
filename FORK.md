# About this fork

This is Abu Bakar Gondal's fork of
[SecurityRonin/docx-mcp](https://github.com/SecurityRonin/docx-mcp).
It runs the upstream server with fixes and additions for producing Word files
that open cleanly in Microsoft Word, and for Arabic / right-to-left work.

Every change is listed in [CHANGELOG.md](CHANGELOG.md) with the reason for it.

## Arabic and right-to-left documents

Write Arabic as you would English: new paragraphs, footnotes, table cells,
headings and Markdown imports are marked right-to-left automatically when the
document is saved, including lines that mix Arabic and English.

For a fully Arabic document, call once:

```
set_document_direction("rtl", complex_script_font="Amiri", complex_script_size_pt=16)
```

English paragraphs inside it stay left-to-right. `get_direction_report` lists
any Arabic text that is still missing direction markup.

## Designing documents

- **A whole book in one call:** `apply_book_preset("tahqiq", title="…",
  running_head="title-chapter")` sets page, fonts, a labelled style set,
  running heads, folio, footnote design and honorifics. Then write with
  styles: «باب» for chapters, «ترجمة» for sections, «رقم الحديث» for [١],
  «كلام النبي ﷺ» for the Prophet's words, `insert_separator` for ❁ ❁ ❁.
- **Keep it consistent:** `restrict_formatting_to_styles()` turns off hand
  formatting in Word; `set_style_pane("in-use")` keeps the Styles pane short.

- **Styles first.** `define_style` builds complete styles (Arabic and Latin
  fonts, kashida justification, borders, spacing, labels in Arabic or
  English); `apply_style_where` applies them by pattern; `style_report` shows
  where hand formatting should become a style; `organize_styles` tidies the
  gallery.
- **Several fonts in one line:** `insert_rich_paragraph` with spans.
- **Honorifics as real glyphs:** `apply_honorifics(color="520912")` turns
  «صلى الله عليه وسلم», «رضي الله عنهما», (رض) … into ﷺ ﵄ ﵁ in Scheherazade
  New, coloured; `glyph_set="jameel"` uses Jameel Noori Nastaleeq's artwork
  (with حفظه الله, رحمة الله عليه …).
- **Book pages:** `set_book_layout(page_size="17x24", mirror_margins=True,
  rtl_binding=True)`, `design_header_footer` with `{STYLEREF:Chapter}`
  running heads and `{PAGE}`, `set_page_numbering`.
- **Footnotes like a printed Arabic book:** `set_note_continuation(color=
  "520912")` puts "=" at the foot of a page where a note carries over,
  `set_note_separator("full", color="520912")`, `set_note_numbering(
  restart="page")`, and `add_custom_mark_note(..., mark="[١]")` for a
  narration number that carries its own takhrij note.
- **Indexes (فهارس):** `mark_index_entries(["عمر بن الخطاب", "البخاري"],
  index_type="P")` then `insert_index(title="فهرس الأعلام", index_type="P")`;
  `insert_toc(title="فهرس الموضوعات", styles={"Chapter": 1, "Bab": 2},
  levels=None)`. Word fills the page numbers when the file opens.
- **Ornaments and frames:** `insert_ornament("shamsa")` on a title page,
  `insert_ornament("divider")` between sections, `add_page_frame()` on every
  page, `insert_svg(...)` for your own vector artwork.
- **Themes:** `set_theme_fonts`, `apply_color_palette("scholarly-maroon")`.

## How every change is checked

1. The full upstream test suite still passes.
2. Each fix gets its own regression test (see `tests/test_schema_validity.py`).
3. Output files are validated against the official ECMA-376 / ISO 29500
   WordprocessingML schemas, the same rules Word uses when it decides a file
   is "unreadable".

## Running this fork

In the Claude desktop app settings file:

```json
"docx-mcp": {
  "command": "uvx",
  "args": ["--from", "docx-mcp-server[equations] @ git+https://github.com/AbubakarYasir/docx-mcp@main", "docx-mcp-server"],
  "env": { "DOCX_MCP_AUTHOR": "Abu Bakar Gondal" }
}
```

`uvx` installs the latest commit of this fork.

Settings (environment variables):

| Variable | Default | Effect |
|---|---|---|
| `DOCX_MCP_AUTHOR` | `Claude` | Name on tracked changes and comments |
| `DOCX_MCP_AUTO_BIDI` | `1` | `0` turns off automatic right-to-left markup |
| `DOCX_MCP_RTL_LANG` | `ar-SA` | Proofing language for right-to-left text |
 No version pin on the `mcp`
library is needed any more: the server works with both mcp 1.x and 2.x.

## Keeping up with upstream

Upstream changes are merged in deliberately, never automatically, and the
checks above are re-run after each merge.
