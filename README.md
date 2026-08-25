# docformat.py — consistent formatting for your raw docx files (v3)

## What it does
Run it on any raw `.docx` you write, and it applies **one consistent visual
style** without touching your wording:

- First line → styled Title
- `Part A:` / `Section 1:` / `Part H —` style lines → **Heading 1**
- `1) ...` / `2. ...` style lines → **Heading 2**
- `a) ...` / `1.1 ...` style lines → **Heading 3**
- A short line you already made fully bold, that doesn't match any of the
  above, still becomes a heading (ALL CAPS → Heading 1, else Heading 2) —
  a safety net for whatever phrasing didn't match a pattern
- Bullet lists → consistent font/spacing (bullets stay bullets, nesting kept)
- Body paragraphs → one consistent font/color/size/justification, but any
  bold/italic you already typed on specific words is kept, not erased
- **`Note: ...` / `Warning:` or `Caution:` / `Danger:` or `Critical:`** at the
  start of a line → a colored callout box (blue / amber / red), the label
  stripped and reused as the box heading — if you already write asides this
  way (a lot of people do without thinking about it), this is free
- A **Table of Contents** is inserted after the title from whatever got
  detected as Heading 1/2/3 — see the note on this below
- Manual blank-line spacing → removed (spacing comes from styles now)
- Every image → resized to one consistent width and centered — including
  floating/"anchored" screenshots, converted to inline first
- ASCII/box-drawing diagrams → collected into a bordered, monospaced box
- Any table → shaded header row, banded rows, clean borders
- A running header (doc title) and footer (`Page X of Y`) on every page

## What it deliberately does NOT do
It never rewrites your words, invents section titles, or writes captions.
Structure it can detect from what's already there (numbering, bold, a "Note:"
prefix) gets formatted; content it would have to *understand* to add is out
of scope — that stays a small, separate ask of me afterward if you want it.

## About the Table of Contents
Word can't know page numbers until it lays the document out, so like every
docx-generation tool, this inserts a **live TOC field**, not fixed text.
Open the file in Word and it fills in automatically (the file is flagged to
auto-update fields on open — no manual step in Word). In LibreOffice, or if
Word doesn't auto-update, right-click the placeholder text under "Table of
Contents" → **Update Field**. One click, then it's done for that copy.

## Requirements
`pip install python-docx`

## Usage
```bash
python3 docformat.py input.docx output.docx
python3 docformat.py input.docx output.docx myconfig.json   # your own styling
```
No API key, no network calls — runs entirely on your machine.

## Customizing the look
Copy `config.example.json`, change only what you care about. Everything
tunable:

| Key | What it controls |
|---|---|
| `fonts.body` / `fonts.mono` | Body font / code & diagram font |
| `colors.*` | Every color — body, each heading level, table shading, diagram box, and each callout type (`note_bg`/`note_border`/`note_text`, same for `warning_*` and `danger_*`) |
| `sizes.*` | Point size for title and each heading level, and body text |
| `bold.*` / `italic.*` | Bold/italic per heading level |
| `spacing.*` / `page.*` | Paragraph spacing, margins, image width |
| `header.enabled` / `footer.enabled` / `footer.text_format` | Running header/footer, with `{PAGE}`/`{NUMPAGES}` placeholders |
| `toc.enabled` | Turn the auto-generated Table of Contents off |
| `heading_patterns.heading1/2/3` | Regex list per level — add your own if you head sections a way the defaults don't catch |
| `bold_line_fallback` / `bold_line_max_words` | Turn the "bold short line → heading" safety net off, or change how many words counts as "short" |
| `callout_patterns.note/warning/danger` | The line-start patterns that trigger each callout box — add your own trigger words here (e.g. add `tip` to the note pattern) |

## The actual goal here
This isn't meant to produce a finished document zero-touch — it's meant to
take the mechanical 80–90% off your plate (structure, spacing, images,
tables, now boxes and a TOC) so whatever manual tuning you still do is
small, not a full reformat pass. If a chunk of your raw writing doesn't
match any pattern, it just stays as plain, clean body text — never guessed
into the wrong shape.

## Known limitations
- Heading detection is pattern-based. Two different conceptual levels that
  reuse the exact same numbering in your raw text (e.g. a "1." section
  immediately followed by its own "1." sub-step) can't be told apart from
  text alone.
- The bold-line heading fallback and the callout-box detection both require
  the paragraph to already read that way in your raw draft (fully bold, or
  starting with `Note:`/`Warning:`/`Danger:`) — nothing is invented from
  paragraphs that don't already signal it.
