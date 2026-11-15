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
  "args": ["--from", "git+https://github.com/AbubakarYasir/docx-mcp", "docx-mcp-server"]
}
```

`uvx` installs the latest commit of this fork. No version pin on the `mcp`
library is needed any more: the server works with both mcp 1.x and 2.x.

## Keeping up with upstream

Upstream changes are merged in deliberately, never automatically, and the
checks above are re-run after each merge.
