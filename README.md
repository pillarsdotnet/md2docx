# md2docx

Convert Markdown to `.docx` with every formatting decision driven by a
configuration file.

`md2docx` is a single Python script with no third-party dependencies. It wraps
pandoc, but instead of accepting pandoc's default look it compiles your config
into two throwaway artifacts and hands them to pandoc:

1. a **`reference.docx`** whose `styles.xml` carries the run and paragraph
   properties you asked for, and
2. a **Lua filter** that routes inline constructs pandoc would otherwise emit
   as direct formatting (`*emphasis*`, `**strong**`, `~~strikeout~~`, …)
   through named character styles, so the config can actually reach them.

Out of the box, code — both `` `backticks` `` and fenced blocks — is set in a
monospace face on a light grey `#dddddd` background.

## Requirements

- Python 3.11 or newer (3.8+ works if you `pip install tomli`)
- [pandoc](https://pandoc.org/installing.html) on `PATH`
- optional: PyYAML, only if you want to write your config in YAML

## Install

```
git clone https://github.com/pillarsdotnet/md2docx
cd md2docx
./md2docx.py --install
```

`--install` copies the script to `md2docx` in the first writable directory on
your `$PATH`, `chmod`s it to 0755, and warns if an existing `md2docx` earlier
on `$PATH` would shadow it. Name a directory to choose one yourself:

```
./md2docx.py --install ~/.local/bin
```

If nothing on `$PATH` is writable it says so rather than guessing — the usual
fix is `mkdir -p ~/.local/bin` and putting that on your `PATH`. Re-running
`--install` upgrades an existing copy in place; it refuses to overwrite a file
that isn't md2docx unless you add `--force`.

Nothing requires installation, though. The repo ships a `md2docx` symlink to
`md2docx.py`, so `./md2docx` works straight from a clone.

## Usage

```
md2docx README.md                    # writes README.docx
md2docx notes.md -o report.docx
md2docx -c house-style.toml notes.md -o report.docx
md2docx chapter*.md -o book.docx     # several inputs, one document
cat notes.md | md2docx - -o notes.docx
```

Useful flags:

| Flag | Effect |
|------|--------|
| `--install [DIR]` | copy this script to `DIR/md2docx`, or to the first writable `$PATH` entry |
| `-c, --config FILE` | use this config instead of the discovered one |
| `--init [FILE]` | write the default config (default `./md2docx.toml`; `--init user` writes `~/.config/md2docx/default.toml`) |
| `--list-styles` | list every configurable style key |
| `--print-reference FILE` | write the generated `reference.docx` and stop |
| `--print-filter` | print the generated Lua filter and stop |
| `--reference-doc FILE` | start from your own `.docx` instead of pandoc's |
| `--toc` / `--no-toc`, `--number-sections` | the usual |
| `-n, --dry-run`, `-v, --verbose` | show the pandoc command |
| `--keep-temp` | keep the generated reference doc and filter for inspection |

Anything after `--` goes straight to pandoc:

```
md2docx notes.md -o notes.docx -- --shift-heading-level-by=1
```

## Configuration

Run `md2docx --init` to drop a fully commented `md2docx.toml` next to your
document. Everything in it is optional — whatever you leave out keeps its
built-in default, so a config can be as small as:

```toml
[styles.inline_code]
background = "#ffe0b2"
```

### Where the config comes from

1. `--config FILE`, if given
2. the first `md2docx.toml` / `.yaml` / `.yml` / `.json` found walking up from
   the input file's directory (the search stops at a repository root)
3. your personal stylesheet, `~/.config/md2docx/default.toml` — also `.yaml`,
   `.yml` or `.json`, and honouring `$XDG_CONFIG_HOME` if you set it
4. the built-in defaults

So the usual arrangement is one personal stylesheet that applies everywhere:

```
md2docx --init user        # writes ~/.config/md2docx/default.toml
$EDITOR ~/.config/md2docx/default.toml
md2docx anything.md        # picks it up with no -c needed
```

Drop an `md2docx.toml` beside a document to override that stylesheet for one
project, or pass `-c` to override both.

Your file is deep-merged *over* the defaults, never replacing them wholesale.

### Sections

`[document]` — page size, orientation, margins, default font and language.

`[pandoc]` — reader format, TOC, section numbering, syntax highlighting, plus
`extra_args` and additional `filters`.

`[styles.KEY]` — one block per Markdown construct. `md2docx --list-styles`
prints all of them; the most-used are:

| Key | Markdown | Word style |
|-----|----------|------------|
| `emphasis` | `*text*` | Emphasis |
| `strong` | `**text**` | Strong |
| `strikeout` | `~~text~~` | Strikeout |
| `inline_code` | `` `text` `` | Verbatim Char |
| `code_block` | ` ``` fenced ``` ` | Source Code |
| `blockquote` | `> text` | Block Text |
| `link` | `[text](url)` | Hyperlink |
| `heading1` … `heading9` | `#` … `#########` | heading 1 … 9 |
| `superscript` / `subscript` | `x^2^` / `H~2~O` | Superscript / Subscript |
| `body`, `normal`, `compact` | paragraphs, list items, cells | Body Text, Normal, Compact |
| `title`, `subtitle`, `author`, `date` | YAML metadata | Title, Subtitle, … |
| `caption`, `table_caption`, `image_caption` | captions | Caption, … |
| `footnote_text`, `footnote_reference` | `[^1]` | Footnote Text, Footnote Reference |

`[table]` — borders, cell padding, and a `[table.header]` sub-block for the
header row.

### Style settings

Character settings (valid in every block):

`font`, `size`, `bold`, `italic`, `underline`, `underline_color`, `strike`,
`double_strike`, `all_caps`, `small_caps`, `color`, `background`, `highlight`,
`vert_align`, `letter_spacing`, `position`, `kerning`, `language`, `border`,
and the per-script variants `font_ascii`, `font_hansi`, `font_cs`,
`font_east_asia`.

Paragraph settings (valid in paragraph blocks):

`alignment`, `space_before`, `space_after`, `line_spacing`, `line_spacing_rule`,
`indent_left`, `indent_right`, `indent_first_line`, `indent_hanging`,
`keep_with_next`, `keep_lines`, `page_break_before`, `widow_control`,
`contextual_spacing`, `outline_level`, `paragraph_background`, `border`,
`suppress_line_numbers`.

Metadata: `based_on`, `next`, `priority`, `quick_format`, `hidden`, and
`docx_style` — which retargets the block at a differently-named Word style.
`docx_style` is only accepted on the seven filter-driven inline styles below,
because those are the only ones whose style name `md2docx` controls; for
everything else pandoc picks the name and retargeting would silently do
nothing.

Units and colors:

- lengths are points when bare (`size = 10`), or a string with a unit
  (`"1in"`, `"2.54cm"`, `"12pt"`, `"16px"`)
- `line_spacing` is a multiple (`1.15`), unless you set
  `line_spacing_rule = "exact"` or `"atLeast"`, when it becomes a length
- colors are `"#dddddd"`, `"dddddd"`, `"#ddd"`, a common name (`"lightgrey"`),
  or `"auto"`
- borders are tables: `border = { left = { style = "single", size = 0.75,
  color = "#bbbbbb", space = 4 } }`. Sides are `top`, `left`, `bottom`,
  `right`, plus `between` (a rule between consecutive paragraphs of the style)
  and `bar`. `all` is shorthand for the four edges; `between` and `bar` must be
  named explicitly. A character-style `border` takes the properties directly
  rather than a table of sides.

### `via_filter` and `keep_native`

Pandoc writes emphasis, strong, strikeout, underline, small caps, superscript
and subscript as *direct* formatting, which no reference document can override.
Those seven blocks therefore take two extra settings:

```toml
[styles.emphasis]
via_filter  = true    # wrap in the "Emphasis" character style
keep_native = true    # ...and keep pandoc's own <w:i/> as well
italic      = true
color       = "#8b0000"
```

- `via_filter = false` leaves pandoc's native output completely alone; the rest
  of the block is then ignored for that construct.
- `keep_native = true` (the default) layers your style *on top of* pandoc's
  direct formatting. This is what keeps `***bold italic***` bold **and**
  italic, because a Word run can only carry one character style and the
  innermost one wins.
- `keep_native = false` gives the style total control — useful when you want
  `*emphasis*` to be, say, coloured and **not** italic, at the cost of nested
  `***bold italic***` rendering with the inner style only.

## How it works

```
md2docx.toml ─┬─> styles.xml surgery ──> reference.docx ─┐
              │                                          ├─> pandoc ──> out.docx
              └─> generated Lua filter ──────────────────┘
```

The base reference document is pandoc's own (`pandoc --print-default-data-file
reference.docx`) unless you pass `--reference-doc`. `md2docx` rewrites
`word/styles.xml` — creating styles that don't exist yet, clearing theme
references that would otherwise outrank your literal values, and inserting each
property at its schema-mandated position — and rewrites the section properties
in `word/document.xml` for page size and margins.

Because the result is an ordinary reference document, everything stays
inspectable:

```
md2docx --print-reference ref.docx    # open it in Word and look around
md2docx --print-filter                # read the generated Lua
md2docx notes.md -o notes.docx --keep-temp -v
```

## Example

`examples/sample.md` exercises every construct the default config styles:

```
md2docx examples/sample.md -o sample.docx
```

## License

MIT
