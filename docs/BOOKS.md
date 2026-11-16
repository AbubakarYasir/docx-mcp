# Book guide — docx-mcp Book Edition

How to produce Arabic, Urdu, English and bilingual books with docx-mcp: the
workflow, the design rules the tools follow, recipes for common jobs, and a
reference for every book tool. (The [README](../README.md) has the overview
and the gallery.)

- [The workflow](#the-workflow)
- [Design rules](#design-rules-the-tools-follow)
- [Recipes](#recipes)
- [Word behaviours worth knowing](#word-behaviours-worth-knowing)
- [Tool reference](#tool-reference)

## The workflow

```text
1. create_document("book.docx")              or open_document(...) on a manuscript
2. apply_book_preset(preset, title=, author=, palette=)
3. write with styles: headings, paragraphs, notes, separators, verses
4. format_matches(...) to style Qur'an, prophetic speech, names — by pattern
5. apply_honorifics()
6. build_faharis(...)
7. audit_document(); save_document()
8. export_pdf("book-print.pdf", "print"); export_pdf("book.pdf", "digital")
```

Everything stays live Word: change a heading and the contents follow, add a
footnote and the marks renumber, open the file and edit it by hand.

### Presets

| Preset | For | Page | Body / headings | Palette | Notes |
|---|---|---|---|---|---|
| `tahqiq` | Arabic critical editions, risalahs | 17×24 cm | Amiri 13 / Sakkal Majalla | maroon | «(١)» marks, unvocalised notes, fleuron head, chapters run on |
| `kitab` | Arabic books | 17×24 cm | Amiri 14 / Sakkal Majalla | brick | chapters open a new page |
| `urdu-nastaliq` | Urdu books | 17×24 cm | Jameel Noori Nastaleeq 15 | maroon | chapters open a new page |
| `english` | English Islamic books | 16×23 cm | Palatino Linotype 11.5, Amiri for Arabic | ink | «57-» notes, folio at the foot, no running head |
| `bilingual` | Arabic text with English translation | Royal (15.6×23.4 cm) | Palatino Linotype 11.5 + Amiri | ink | the Bulugh al-Maram layout |

Any value can be overridden: `overrides={"body_pt": 14, "margins_cm": (2.0, 2.0, 2.2, 1.8)}`.

## Design rules the tools follow

These come from studying printed tahqiq editions and English hadith books.
They are defaults, and every one can be changed per call.

1. **Colour marks structure, never reading text.** The *accent* colour goes
   on chapter titles, numbers, ornament and rules; the *secondary* (usually
   gold) on ornament detail. Body text, footnotes and headings' text are
   ink. Leave `color` unset and the house palette applies.
2. **Headings are typography.** No boxes. Decoration is optional, and it is
   anchored behind the text so the table of contents stays clean.
3. **The author's harakat are sacred; the editor's apparatus is plain.**
   Vocalised body text gets "at least" line spacing. Footnotes, references,
   book names and running-head titles are unvocalised.
4. **Notes like a printed book:** a fine rule over about 65% of the line,
   «(١)» marks in Arabic or «57-» in English, and «=» where a note runs on.
5. **Folios on the outer side**, with the chapter on odd pages and the book
   title on even ones.
6. **Safe justification:** justify, or low kashida. Medium and high kashida
   are never used.

### House palettes

| Palette | Accent | Secondary | Character |
|---|---|---|---|
| `maroon` | `#520912` | `#A8874E` | classic tahqiq |
| `brick` | `#7A2E2A` | `#B08D57` | warm, Maghribi |
| `navy-gold` | `#1F2E4A` | `#B8955A` | formal, English editions |
| `green-gold` | `#2F6E54` | `#C2A36A` | Hijazi, mushaf-like |
| `teal` | `#0F4C5C` | `#C29B57` | modern scholarly |
| `burgundy-silk` | `#6B1F33` | `#C6A15B` | gift editions |
| `ink` | `#1A1A1A` | `#6B6B6B` | austere, one colour |

A client's own colours: `set_house_colors(accent="7A1F2B", secondary="C9A55C", ink="222222")`.

## Recipes

### Tahqiq: musnad numbers that carry their takhrij

```text
p = insert_paragraph(after, "عَنْ أَبِي هُرَيْرَةَ … «هُوَ الطَّهُورُ مَاؤُهُ».")
add_custom_mark_note(p, "أخرجه أبو داود والترمذي …", mark="[١]", where="start",
                     mark_in_body={"bold": True})
format_matches("«[^»]+»", regex=True, char_style="Prophetic Speech")
```

The «[١]» opens the narration, is the footnote mark, and opens the note in
bold. Word's automatic numbering skips it.

### Poetry

```text
insert_verses(["العِلْمُ صَيْدٌ وَالكِتَابَةُ قَيْدُهُ *** قَيِّدْ صُيُودَكَ بِالحِبَالِ الوَاثِقَهْ"],
              after_para_id=p, numbered=False, width_pct=80)
```

### Self-numbering tarajim

```text
define_numbered_style("Tarjama Entry", template="%1 .", aliases=["ترجمة مرقمة"])
insert_paragraph(after, "سُلَيْمَانُ بْنُ عَمْرٍو:", style="Tarjama Entry")   → «١٨٨٠ . سليمان …»
```

### Faharis

```text
build_faharis(
    people=["أبو هريرة", ["ابن عمر", "عبد الله بن عمر"]],   # text to find, or [text, entry]
    places=["المدينة", "مكة"],
    sources=["صحيح البخاري، دار طوق النجاة، ١٤٢٢هـ."],
)
```

The Qur'an index comes from the `Quran` style (or ﴿…﴾), the atraf from
`Prophetic Speech` / `Matn`, and the topics index is the table of contents.
Matching ignores harakat. Word fills in the page numbers on export.

### English hadith book

```text
apply_book_preset("bilingual", title="Bulugh al-Maram", palette="navy-gold")
insert_bilingual_heading("كتاب الطهارة", "Book 1: Purification")
insert_hadith_entry(1, arabic, "Narrated Abu Hurairah (RAA): … “Its water is pure …”",
                    source="Related by al-Arba'ah.")
insert_glossary_entry("Junub", "In a state of major ritual impurity.")
```

### One manuscript, several clients

```text
apply_book_preset("tahqiq", palette="maroon")                      # client A
design_running_head("vignette"); set_note_separator("footnote", "khatam")
set_house_colors("green-gold")                                       # client B
```

Restyle, export, repeat. The text is never touched.

### Print and digital PDFs

```text
set_pdf_preferences(bleed_mm=5, crop_marks=True)   # the printer's spec, saved in the file
export_pdf("book-print.pdf", "print")              # TrimBox, BleedBox, crop marks, no scaling
export_pdf("book.pdf", "digital")                  # bookmarks open, spreads, RTL page order
export_pdf("book-archive.pdf", "digital", pdfa=True)
```

Page size, margins, gutter and mirror margins come from `set_book_layout`
(e.g. `set_book_layout(page_size="17x24", mirror_margins=True, gutter_cm=0.5)`).

## Word behaviours worth knowing

All of these were found by rendering in Microsoft Word 365. The server
already handles them.

- **Kashida:** Word drops lines of vocalised Arabic under medium and high
  kashida. Only low kashida is used.
- **Numerals:** automatic numbers follow the computer's *Numeral* option. The
  PDF export sets it per book and restores it afterwards. In Word itself, see
  File → Options → Advanced → Show document content → Numeral.
- **Diacritic colour:** one colour for the whole document (File → Options →
  Advanced → "Use this color for diacritics"). `color_tashkil` stores it and
  the export applies it.
- **Top of page:** Word drops the space above a paragraph that starts a page.
  Heading art is anchored to the text line, so it stays on its title.
- **Mark-only runs:** a run that holds only harakat is kept right-to-left, or
  Word breaks the word apart.
- **Pictures in headings** get copied into the table of contents, so heading
  art is anchored behind the text.

## Tool reference

Generated from the server's own tool definitions. Every tool also accepts
`document_handle` for parallel sessions.

### Whole book

#### `apply_book_preset`

```text
apply_book_preset(preset='tahqiq', title=None, author=None, palette=None, running_head=None, chapters_new_page=None, justify='justify', honorifics=True, tashkil=True, notes_tashkil='strip', note_separator='rule', note_marks=None, note_continuation=True, page_numbers='decimal', title_page=True, overrides=None)
```

Set up a whole book in one call: page, margins, fonts, a labelled style
set (Title, Heading 1 باب, Heading 2 ترجمة, Narration Number, Prophetic
Speech ﷺ, Matn, Isnad, Quran, Separator, footnote text; English presets
add Arabic Text, Translation, Source, Glossary), running head, folio,
designed note rule, note marks, honorific glyphs, house colours.
preset: tahqiq (Amiri / Sakkal Majalla, 17×24), kitab (chapters on new
pages), urdu-nastaliq, english, bilingual (Arabic with English, as in
Bulugh al-Maram).
palette: maroon, brick, navy-gold, green-gold, teal, burgundy-silk, ink
(list_house_palettes). Colour marks structure; reading text stays ink.
running_head: fleuron, rule, boxed, vignette, cartouche, none.
tashkil: the body carries harakat, so line spacing is "at least" and
vowel marks are never clipped. notes_tashkil: strip (default — notes,
references and book titles are unvocalised), keep-shadda, keep.
note_separator: any set_note_separator style. note_marks: parentheses
«(١)» (Arabic default), dash «57-» (English default), brackets, plain.
justify: justify or kashida (low kashida; safe for vocalised text).
overrides: any preset value, e.g. {"body_pt": 14}.

#### `set_book_layout`

```text
set_book_layout(page_size=None, width_mm=None, height_mm=None, top_cm=None, bottom_cm=None, inside_cm=None, outside_cm=None, gutter_cm=None, mirror_margins=None, rtl_binding=None, header_cm=None, footer_cm=None, section_index=None)
```

Book page layout: page_size A4 | A5 | B5 | 17x24 | 14x21 | 16x23 | royal |
letter, margins (inside/outside for facing pages), gutter, mirror margins,
and rtl_binding (gutter on the right for right-to-left books).

#### `set_house_colors`

```text
set_house_colors(palette=None, accent=None, secondary=None, ink=None)
```

Set the document's house colours — every design tool (headings, running
heads, frames, ornaments, note rules, separators, verses) uses them when no
colour is given. Pick a palette and/or give hex colours to match a
client's brand or a reference book. Also sets theme accent1/accent2/dk2.

#### `get_house_colors`

```text
get_house_colors()
```

The document's current house colours (accent, secondary, ink).

#### `list_house_palettes`

```text
list_house_palettes()
```

House colour palettes for books (accent for structure, secondary for
ornament, ink for text and rules), drawn from printed Arabic and English
Islamic books: maroon, brick, navy-gold, green-gold, teal, burgundy-silk,
ink.

#### `set_theme_fonts`

```text
set_theme_fonts(heading_latin=None, body_latin=None, heading_arabic=None, body_arabic=None)
```

Set theme heading/body fonts (Latin and Arabic). Styles using theme fonts follow.

#### `apply_color_palette`

```text
apply_color_palette(palette=None, colors=None)
```

Set theme colours: palette scholarly-maroon | emerald-gold | midnight-ink |
manuscript, and/or explicit slots {accent1: "520912", ...}.

### Headings, titles and parts

#### `insert_heading`

```text
insert_heading(text, level=1, after_para_id=None, design='plain', color=None, ornament_width_pct=30)
```

A chapter/section heading in the book's heading style (stays in the
table of contents, no art copied into it). design: plain; fleuron (a small
fleuron rule under the heading); flanked (calligraphic brackets drawn
either side of the measured title); cartouche; bullet (a small khatam at
the start, for sub-headings). Colour defaults to the house accent.

#### `insert_bilingual_heading`

```text
insert_bilingual_heading(arabic, english, level=1, after_para_id=None)
```

A heading in Arabic above its English form («كتاب الطهارة» /
«Book 1: Purification»); only the English line enters the table of
contents.

#### `insert_cartouche_title`

```text
insert_cartouche_title(text, after_para_id=None, kind='lobed', color=None, border=None, text_color='FFFFFF', font=None, size_pt=18, bold=True, width_cm=None, height_cm=None, style=None)
```

A title on a cartouche: real, editable Word text centred on a vector
cartouche anchored behind it. kind: lobed (cusped label with knob, green
in the house colours by default), banner (rounded with fleuron ends), arch (pointed
ogee ends, cover style). style: paragraph style for the title (e.g.
"Heading 1" so it appears in the TOC).

#### `insert_part_title`

```text
insert_part_title(title, after_para_id=None, subtitle=None, items=None, color=None, font=None, size_pt=30, frame=True, ornament='divider')
```

A part-title page («الفهارس», «القسم الثاني»): own page, vertically
centred, woven frame on that page only, no running head. items: lines
listed under the title.

#### `define_numbered_style`

```text
define_numbered_style(name, template='%1 .', number_format='decimal', start_at=1, color=None, bold=True, font=None, size_pt=None, aliases=None, follow='space')
```

A paragraph style that numbers itself (tarajim, narrations, masa'il:
«١٨٨٠ . سليمان بن عمرو:»). Inserting an entry renumbers the rest.
template: "%1 .", "[%1]", "(%1)", "%1-".

### Writing with styles

#### `define_style`

```text
define_style(name, style_type='paragraph', based_on=None, next_style=None, linked_style=None, aliases=None, gallery=None, ui_priority=None, hidden=None, run=None, paragraph=None, replace_formatting=False)
```

Create a complete style, or update it if it exists (by name, id or alias).

Use styles instead of formatting paragraphs one by one: change the style
once and every paragraph using it follows.

run keys: font (all scripts), latin_font, cs_font (Arabic/Hebrew), size_pt,
  cs_size_pt, bold, italic, underline, strike, color, highlight, shading,
  caps, small_caps, spacing_pt, position_pt, rtl, language, cs_language, hidden.
paragraph keys: alignment (start|end|center|justify|kashida|kashida-low|
  kashida-high|distribute), direction (rtl|ltr), space_before_pt,
  space_after_pt, line_spacing (multiple), line_exact_pt, line_at_least_pt,
  indent_start_cm, indent_end_cm, first_line_cm, hanging_cm, keep_with_next,
  keep_together, page_break_before, widow_control, outline_level (0=H1),
  border_top/border_bottom/border_start/border_end ({style: single|double|
  thick|dotted|dashed, color, size_pt, space_pt}), shading, tabs
  ([{pos_cm, align, leader}]), suppress_hyphenation.
aliases: labels shown in Word's style list, e.g. ["متن", "Body"].
gallery: show in Home ▸ Styles. ui_priority: order (lower first).

#### `apply_style_where`

```text
apply_style_where(style, text_regex=None, current_style=None, para_ids=None, clear_direct=False)
```

Apply a style to every match. Paragraph style: paragraphs matching text_regex
/ current_style / para_ids. Character style: text matching text_regex
(e.g. r"«[^»]+»" for quotations, r"\\[[٠-٩0-9]+\\]" for narration numbers).

#### `format_matches`

```text
format_matches(pattern, regex=False, whole_word=False, parts='all', font=None, latin_font=None, cs_font=None, size_pt=None, bold=None, italic=None, underline=None, color=None, highlight=None, shading=None, char_style=None, ignore_tashkil=True)
```

Format every occurrence of a text or regex across body, notes, headers, footers.

E.g. every ﴿…﴾ verse in a Quran font (pattern "﴿[^﴾]+﴾", regex=True,
cs_font="Amiri Quran"), every «…» bold, every name in a colour.
ignore_tashkil: a plain search term also finds the vocalised text
(«قال» matches «قَالَ»), so the same words are formatted however they
are written.

#### `insert_rich_paragraph`

```text
insert_rich_paragraph(after_para_id, spans, paragraph=None)
```

Insert a paragraph made of spans with different fonts, sizes and colours.

spans: [{"text": "...", "font"|"latin_font"|"cs_font": ..., "size_pt": ...,
"bold": ..., "color": "520912", "char_style": ..., ...}]. Text may hold
{PAGE}, {NUMPAGES}, {STYLEREF:Heading 1}, {FIELD:...}, tabs and newlines.
paragraph: style/alignment/direction/spacing/indent/border keys (see define_style).

#### `insert_separator`

```text
insert_separator(after_para_id=None, text='❁ ❁ ❁', ornament=None, color=None)
```

Section separator after a paragraph: text (❁ ❁ ❁, ✽ ✽ ✽, ⁂) in the
Separator style, or a vector ornament ("divider", "divider-dots", "star8").

#### `insert_verses`

```text
insert_verses(verses, after_para_id=None, numbered=False, start_at=1, color=None, number_color=None, font=None, size_pt=None, bold=False, gap_cm=1.0, separator=None, width_pct=80)
```

Poetry (nazm): two hemistichs per line, each spread across its half.
verses: ["صدر *** عجز", …] or [["صدر", "عجز"], …]. numbered adds «٤٨-». width_pct: how much of the line
the couplet spans (80 keeps the halves close, as printed).

#### `add_margin_note`

```text
add_margin_note(para_id, text, color=None, font=None, size_pt=9, bold=True, width_cm=1.6, side='outside')
```

A side heading in the margin beside a paragraph (travels with it).
side: outside (flips on odd/even pages), inside, left, right.

#### `apply_honorifics`

```text
apply_honorifics(glyph_set='unicode', font=None, fallback_font=None, color=None, size_pt=None, bold=None, include_abbreviations=True, only=None, restyle_existing=True, parts='all')
```

Replace honorific phrases with full glyphs, each in a font that draws it.

Phrases (tashkeel and spelling variants match): صلى الله عليه وسلم,
رضي الله عنه/عنها/عنهم/عنهما/عنهن, رحمه الله, رحمهم الله, رحمة الله عليه/
عليها/عليهم/عليهما/عليهن, حفظه الله/حفظها الله/حفظهم الله, عليه السلام …,
عز وجل, جل جلاله, سبحانه وتعالى, تبارك وتعالى, قدس سره; and bracketed
abbreviations (ص) (صلعم) (رض) (رح) (ع).

glyph_set: "unicode" (official code points; Scheherazade New / Noto Naskh
Arabic draw all) or "jameel" (Jameel Noori Nastaleeq's own artwork with
full diacritics). The same code point means different phrases in different
fonts (U+FDFE is سبحانه وتعالى in Unicode but حفظه الله in Jameel), so each
glyph is always written with the font of its set; phrases missing from the
chosen set use the other set in its font. Glyphs already in the text get
colour/size/bold only — their font is never changed.

#### `list_honorifics`

```text
list_honorifics()
```

Honorific ligature glyphs (ﷺ ﵁ ﵀ ﷻ …), the phrases and abbreviations they replace.

### Tashkil

#### `tashkil`

```text
tashkil(mode='normalize', parts='all', para_ids=None)
```

Work on the harakat (tashkil) already in the text. mode: normalize
(canonical order of shadda and vowels, stray marks removed), strip (remove
all harakat — e.g. for references and book titles in the notes),
keep-shadda (remove all but shadda), light (drop case endings, keep the
word-internal vowels). parts: body, notes, headers, all; or para_ids.

#### `tashkil_report`

```text
tashkil_report(parts='all')
```

How vocalised each part is: harakat count, share of Arabic letters
carrying a mark, and paragraphs that are fully, partly or not vocalised —
to check a manuscript before typesetting.

#### `color_tashkil`

```text
color_tashkil(color=None, parts='body', para_ids=None, per_run=False)
```

Draw the harakat in their own colour (as in teaching Qur'ans and
vocalised matn), letters unchanged. Word colours diacritics with one
colour for the whole document: the colour is saved in the document and
export_pdf shows it in the PDF (in Word, File > Options > Advanced >
"Use this color for diacritics" shows it on screen). per_run also colours
marks run by run for LibreOffice and other renderers. Colour defaults to
the house accent; color="auto" removes it.

### Footnotes

#### `add_footnote`

```text
add_footnote(para_id, text, url='')
```

Add a footnote to a paragraph. url, if provided, is rendered as a hotlink.

#### `add_custom_mark_note`

```text
add_custom_mark_note(para_id, text, mark, kind='footnote', where='end', after_text=None, mark_in_body=None, mark_in_note=None, note_text_props=None)
```

Footnote/endnote with your own mark instead of a number, e.g. a
narration number "[١]" that itself carries the takhrij note. where: start,
end, or after (with after_text). mark_in_body / mark_in_note /
note_text_props take run formatting such as {"bold": true, "color":
"520912", "cs_font": "Amiri"}.

#### `set_note_mark_style`

```text
set_note_mark_style(style='parentheses', kind='footnote')
```

Note numbers as «(١)» in the text and the note (parentheses), «[١]»
(brackets) or plain. Numbers stay live; re-run after adding notes.

#### `set_note_separator`

```text
set_note_separator(kind='footnote', style='rule', width_pct=65, text='', color=None, thickness_pt=0.6, direction='auto')
```

The line above footnotes/endnotes. style: rule (a fine rule over
width_pct of the text width, default 65), double, tapered (a calligraphic
line thick at the middle), khatam / diamond (a rule with a small star or
lozenge at its centre), swash, text (an ornament such as "⁂" or "✽ ✽ ✽"),
full (full width), default (Word's short rule), none. Colour defaults to
the house ink. In Word 2013+ this also styles the line above continued
notes.

#### `set_note_continuation`

```text
set_note_continuation(kind='footnote', notice_text='=', separator_text='=', keep_rule=True, font=None, cs_font=None, size_pt=None, color=None, bold=None, notice_alignment='end', separator_alignment='start', direction='auto', legacy_layout=False)
```

Mark notes that run onto the next page: notice_text ("=" by Arabic
convention, or "(continued)") prints at the foot of the page where a note
breaks. separator_text prints above the continued part only in Word 2010
layout (legacy_layout=True; Word then shows "Compatibility Mode"); Word
2013+ layout reuses the normal separator there.

#### `set_note_numbering`

```text
set_note_numbering(kind='footnote', number_format=None, restart=None, start_at=None, position=None)
```

Note numbering for the whole document: number_format decimal |
arabic-abjad | arabic-alpha | roman | upper-roman | letters | upper-letters
| symbols | decimal-parentheses | decimal-circle; restart continuous | page
| section; start_at; position (footnotes: page-bottom | below-text;
endnotes: section-end | document-end).

#### `convert_notes`

```text
convert_notes(from_kind='footnote', ids=None)
```

Turn footnotes into endnotes (from_kind="footnote") or endnotes into
footnotes; ids limits it to some notes.

### Pages: running heads, folios, frames

#### `design_running_head`

```text
design_running_head(style='fleuron', title=None, chapter_style='Heading 1', color=None, font=None, size_pt=12, bold=True, section_index=None)
```

Running heads as Arabic publishers set them, page number on the outer
side; chapter (STYLEREF) on odd pages, book title on even. style:
fleuron (a fine ink rule with a small fleuron at its centre, default);
rule; boxed (number in a box, double rule); vignette (number between ornamental
brackets); cartouche (head on a banner cartouche, folio in a star
medallion at the foot). Colours default to the house palette.

#### `design_header_footer`

```text
design_header_footer(kind='header', start='', center='', end='', variant='default', section_index=None, direction='auto', font=None, cs_font=None, size_pt=None, color=None, bold=None, rule='none', rule_color='000000', rule_style='single', rule_size_pt=0.75, distance_cm=None, zone_formats=None)
```

Design a header or footer: three zones (start, center, end) with live fields.

Tokens: {PAGE}, {NUMPAGES}, {SECTIONPAGES}, {DATE}, {TITLE},
{STYLEREF:Heading 1} (running head = current chapter title, kept up to date
by Word), {FIELD:any field code}. variant: default | first | even (turns on
the matching Word setting; even pages reuse the default variant until an
even one is designed). rule: below | above | both | none, with rule_color,
rule_style (single, double, thick, dotted, dashed, thinThickSmallGap…).
zone_formats: per-zone overrides, e.g. {"center": {"bold": true}}.

#### `set_page_numbering`

```text
set_page_numbering(number_format='decimal', start_at=None, section_index=None)
```

Page number format and restart: decimal, roman, upper-roman, letters,
upper-letters, arabic-abjad (أ ب ج د), arabic-alpha (أ ب ت ث), dash (- 1 -).

#### `add_page_frame`

```text
add_page_frame(color=None, accent=None, margin_cm=1.3, weight=1.0, section_index=None, svg=None, style='rings')
```

Decorative frame on every page (double rule, khatam corners, side
diamonds), drawn behind the text from the header so it repeats without
moving the body. style: rings (interlocking guilloche with stars), chain,
khatam, 8. margin_cm keeps the art inside the printable safe zone.
Colours default to the house palette. Pass svg for your own artwork.

#### `set_page_border`

```text
set_page_border(style='double', color=None, width_pt=1.5, space_pt=24, measure_from='page', pages='all', sides=None, section_index=None)
```

Word page border: single, double, thick, thick-thin, thin-thick,
thin-thick-thin, triple, dotted, dashed, wave, double-wave, 3d, or art
borders art-celtic, art-weave, art-weave-ribbon, art-weave-angles,
art-zigzag, art-triangles, art-diamonds, art-mosaic, art-flowers, art-vine,
art-double-d, art-basic-wide-outline, art-basic-black-dots,
art-twisted-lines; "none" removes. pages: all | first | not-first.

### Indexes and contents (faharis)

#### `build_faharis`

```text
build_faharis(after_para_id=None, include=None, people=None, places=None, sources=None, part_title='الفهارس', tarf_words=6, heading_design='fleuron', columns_people=2)
```

The indexes (الفهارس) at the end of a tahqiq book, each a live Word index
with page numbers: quran (verses, from the Quran style / ﴿﴾), hadith
(atraf: first words of each narration, from Prophetic Speech/Matn), people
(الأعلام), places (الأماكن والبلدان), sources (المصادر والمراجع —
unvocalised, sorted, numbered) and topics (الموضوعات — the table of
contents). include picks and orders them. people / places: names to find
in the text (tashkil-insensitive), or [name, entry] pairs. Starts with a
part-title page unless part_title is None.

#### `mark_index_entries`

```text
mark_index_entries(terms, index_type=None, whole_word=True, first_per_paragraph=True, parts='body')
```

Mark every occurrence of terms as index entries (real XE fields, so
Word's Update Index rebuilds page numbers). terms: strings, or dicts with
text, regex, entry (index wording), subentry ("main:sub"), see ("انظر …"),
bold_page, italic_page, index_type. index_type: one letter for a separate
index, e.g. "P" أعلام, "L" أماكن, "H" أطراف. Arabic proclitics (و ف ب ل ك)
before a word still match. parts: body or all (with notes).

#### `insert_index`

```text
insert_index(after_para_id=None, title=None, index_type=None, columns=2, letter_headings=True, right_align_page_numbers=True, leader='dot', run_in=False, language='ar', font=None, cs_font=None, size_pt=None, heading_color=None, new_page=True, update_on_open=True)
```

Insert an index (فهرس) from the marks: title, one index type, 1–4
columns, letter headings, dotted leaders to right-aligned page numbers,
Arabic sorting (language ar/en/ur/fa), heading colour. Default: end of the
document, on a new page; Word fills page numbers on open.

#### `insert_toc`

```text
insert_toc(after_para_id=None, title=None, levels='1-3', styles=None, hyperlinks=True, leader='dot', page_numbers=True, font=None, cs_font=None, size_pt=None, level1_bold=True, level1_color=None, new_page=True, update_on_open=True)
```

Table of contents (فهرس الموضوعات) from heading levels and/or your own
styles, e.g. styles={"Chapter": 1, "Bab": 2}, levels=None. Clickable
entries, dotted leaders, coloured first level. Default: end of the
document (pass after_para_id to place it at the front).

### English and bilingual books

#### `insert_hadith_entry`

```text
insert_hadith_entry(number, arabic, english, after_para_id=None, bold_quotes=True, source=None, arabic_number_dash=' ـــ ')
```

One narration as set in Bulugh al-Maram (English): the Arabic text
right-to-left with its number «٩٤ ـــ», then the English translation
«94. Narrated …» with the Prophet's words in bold (bold_quotes), then an
optional source line («Related by …»). Use after apply_book_preset
("bilingual" or "english").

#### `insert_glossary_entry`

```text
insert_glossary_entry(term, definition, after_para_id=None)
```

A glossary entry for English Islamic books: bold term, italic
definition, hanging indent.

### Ornament and art

#### `insert_ornament`

```text
insert_ornament(kind='divider', after_para_id=None, width_cm=None, color=None, accent=None, points=12, weight=1.0, alignment='center', style=None, number=None, border=None, height_cm=None)
```

Insert Islamic ornament as native vector art, built with traditional
construction (woven girih, parallel offsets, Bézier fleurons):
star8 (eight-point star frame; number="١٢" puts a number inside),
rosette (woven strapwork, points=8/10/12/16), shamsa (title sunburst),
ring (segmented girih ring), divider (calligraphic swash with fleuron),
divider-dots, corner, band (style=rings | chain | hex-chain | khatam | 8 | 6),
cartouche-lobed / cartouche-banner / cartouche-arch (color = body,
border = edge). color = lines, accent = fill (default: house palette).

#### `list_ornaments`

```text
list_ornaments()
```

The vector ornaments insert_ornament can draw.

#### `insert_svg`

```text
insert_svg(svg, after_para_id=None, width_cm=None, height_cm=None, alignment='center', alt_text='', space_before_pt=6, space_after_pt=6)
```

Insert SVG artwork (markup or a .svg path) as a native vector picture:
sharp at any zoom and in print in Word 2016+, with a PNG fallback for older
readers. Give width or height to keep proportions.

### Direction (RTL)

#### `set_document_direction`

```text
set_document_direction(direction='rtl', complex_script_font='', complex_script_size_pt=None, language='', tables=True)
```

Make a whole document right-to-left (Arabic, Urdu, Persian, Hebrew) or left-to-right.

Sets paragraph direction defaults, every section's direction, the font,
size and proofing language used for Arabic/Hebrew letters, and (optionally)
every table's column order. Paragraphs with their own explicit direction
keep it.

#### `set_paragraph_direction`

```text
set_paragraph_direction(para_id, direction='rtl')
```

Set one paragraph's direction: "rtl", "ltr", or "auto" (from its first letter).

Works for paragraphs in the body, footnotes, endnotes, headers and footers.
New paragraphs containing Arabic/Hebrew text are already made right-to-left
automatically on save; use this to override.

#### `get_direction_report`

```text
get_direction_report()
```

Report right-to-left status: RTL paragraphs, sections and tables, and any
Arabic/Hebrew text that is missing direction markup (which Word renders wrongly).

### PDF

#### `export_pdf`

```text
export_pdf(output_path, profile='digital', engine='auto', numerals=None, bleed_mm=None, crop_marks=None, pdfa=None, two_page=None, update_fields=True)
```

Export a print-ready or digital-ready PDF.

profile: digital — screen-optimised, heading bookmarks open, links kept,
book spreads, right-to-left page order for Arabic/Urdu books. print —
print-optimised, TrimBox at the page edge, bleed (default 3 mm) and crop
marks, no viewer scaling. engine: auto (Microsoft Word when available —
fields, indexes and contents refreshed first — else LibreOffice), word,
libreoffice. numerals: how Word draws note and page numbers (hindi «١٢٣»
for Arabic books, context for English; default from the book preset).
pdfa: archival PDF/A. Unset options follow set_pdf_preferences.

#### `set_pdf_preferences`

```text
set_pdf_preferences(bleed_mm=None, crop_marks=None, pdfa=None, two_page=None, numerals=None)
```

Save a client's PDF wishes in the document so every export follows them:
print bleed in mm, crop marks, PDF/A, spreads in the digital view, numeral
shapes (hindi, arabic, context, system). Page size and margins are set with
set_book_layout / set_page_margins.

#### `get_pdf_preferences`

```text
get_pdf_preferences()
```

The PDF preferences saved in the document.

#### `convert_to_pdf`

```text
convert_to_pdf(output_path)
```

Convert the open document to PDF using LibreOffice headless.

Requires LibreOffice to be installed ('libreoffice' or 'soffice' on PATH).

### Word power settings

#### `restrict_formatting_to_styles`

```text
restrict_formatting_to_styles(allowed=None, enabled=True, block_theme_switching=True)
```

Word's "Limit formatting to a selection of styles": hand formatting is
disabled so everyone uses the book's styles. allowed: names, ids or labels
to keep usable (default all document styles). No password; enabled=False
lifts it.

#### `set_style_pane`

```text
set_style_pane(show='in-use', sort='priority', alternate_names=True)
```

Word's Styles pane: show in-use | recommended | all | custom | headings,
sort name | priority | font | based-on | type, and show style labels such as
«باب».

#### `set_font_embedding`

```text
set_font_embedding(embed=True, subset=True, include_system_fonts=False)
```

Have Word embed the fonts (Amiri, Jameel Noori Nastaleeq …) on its next
save, so the book looks right on computers without them.

#### `hide_proofing_marks`

```text
hide_proofing_marks(hide=True)
```

Hide spelling/grammar squiggles across Arabic and Urdu text.

#### `attach_template`

```text
attach_template(template_path, auto_update_styles=True)
```

Link a Word template (.dotx) on this computer; with auto_update_styles
Word copies the template's styles in on every open — one place to change
the house style of every book.

#### `set_update_fields_on_open`

```text
set_update_fields_on_open(enabled=True)
```

Make Word refresh every field (TOC, index, page references, running heads) on open.
