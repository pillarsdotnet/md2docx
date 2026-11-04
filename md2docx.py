#!/usr/bin/env python3
# Copyright (c) 2026 Robert August Vincent II <pillarsdotnet@gmail.com>
# Co-author: Claude Code.
"""md2docx - convert Markdown to .docx with formatting driven by a config file.

The tool wraps pandoc.  Everything about the look of the output document is
read from a configuration file and compiled into two throwaway artifacts:

  * a ``reference.docx`` whose ``styles.xml`` carries the configured run and
    paragraph properties, and
  * a Lua filter that routes inline constructs pandoc would otherwise render
    with direct formatting (emphasis, strong, strikeout, ...) through named
    character styles so the configuration can reach them.

Only the Python standard library is required; PyYAML is used when present so
that YAML configuration files work too.
"""

from __future__ import annotations

import argparse
import contextlib
import copy
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from collections.abc import Iterable, Iterator

__version__ = "1.0.0"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def w(tag: str) -> str:
    """Return the namespace-qualified OOXML tag name."""
    return f"{{{W}}}{tag}"


# --------------------------------------------------------------------------
# Child-element ordering.  OOXML validates against a sequence, so properties
# have to be inserted in schema order or Word refuses to open the file.
# --------------------------------------------------------------------------

STYLE_ORDER = [
    "name", "aliases", "basedOn", "next", "link", "autoRedefine", "hidden",
    "uiPriority", "semiHidden", "unhideWhenUsed", "qFormat", "locked",
    "personal", "personalCompose", "personalReply", "rsid", "pPr", "rPr",
    "tblPr", "trPr", "tcPr", "tblStylePr",
]

RPR_ORDER = [
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike",
    "dstrike", "outline", "shadow", "emboss", "imprint", "noProof",
    "snapToGrid", "vanish", "webHidden", "color", "spacing", "w", "kern",
    "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd",
    "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
    "specVanish", "oMath",
]

PPR_ORDER = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
    "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs",
    "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
    "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
    "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
    "suppressOverlap", "jc", "textDirection", "textAlignment",
    "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr",
    "pPrChange",
]

TBLPR_ORDER = [
    "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
    "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd",
    "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption",
    "tblDescription",
]

BORDER_SIDES_P = ["top", "left", "bottom", "right", "between", "bar"]
BORDER_SIDES_T = ["top", "left", "bottom", "right", "insideH", "insideV"]

# --------------------------------------------------------------------------
# Style registry: config key -> (docx style name, style type)
# --------------------------------------------------------------------------

STYLE_MAP = {
    "normal":             ("Normal", "paragraph"),
    "body":               ("Body Text", "paragraph"),
    "first_paragraph":    ("First Paragraph", "paragraph"),
    "compact":            ("Compact", "paragraph"),
    "title":              ("Title", "paragraph"),
    "subtitle":           ("Subtitle", "paragraph"),
    "author":             ("Author", "paragraph"),
    "date":               ("Date", "paragraph"),
    "abstract":           ("Abstract", "paragraph"),
    "abstract_title":     ("Abstract Title", "paragraph"),
    "bibliography":       ("Bibliography", "paragraph"),
    "heading1":           ("heading 1", "paragraph"),
    "heading2":           ("heading 2", "paragraph"),
    "heading3":           ("heading 3", "paragraph"),
    "heading4":           ("heading 4", "paragraph"),
    "heading5":           ("heading 5", "paragraph"),
    "heading6":           ("heading 6", "paragraph"),
    "heading7":           ("heading 7", "paragraph"),
    "heading8":           ("heading 8", "paragraph"),
    "heading9":           ("heading 9", "paragraph"),
    "blockquote":         ("Block Text", "paragraph"),
    "code_block":         ("Source Code", "paragraph"),
    "footnote_text":      ("Footnote Text", "paragraph"),
    "definition_term":    ("Definition Term", "paragraph"),
    "definition":         ("Definition", "paragraph"),
    "caption":            ("Caption", "paragraph"),
    "table_caption":      ("Table Caption", "paragraph"),
    "image_caption":      ("Image Caption", "paragraph"),
    "figure":             ("Figure", "paragraph"),
    "captioned_figure":   ("Captioned Figure", "paragraph"),
    "toc_heading":        ("TOC Heading", "paragraph"),
    "inline_code":        ("Verbatim Char", "character"),
    "link":               ("Hyperlink", "character"),
    "footnote_reference": ("Footnote Reference", "character"),
    "section_number":     ("Section Number", "character"),
    "emphasis":           ("Emphasis", "character"),
    "strong":             ("Strong", "character"),
    "strikeout":          ("Strikeout", "character"),
    "underline":          ("Underline", "character"),
    "small_caps":         ("Small Caps", "character"),
    "superscript":        ("Superscript", "character"),
    "subscript":          ("Subscript", "character"),
    "table":              ("Table", "table"),
}

# Config keys whose formatting pandoc emits as direct formatting rather than a
# named style.  These are routed through the generated Lua filter.
FILTERED_INLINES = {
    "emphasis":    "Emph",
    "strong":      "Strong",
    "strikeout":   "Strikeout",
    "underline":   "Underline",
    "small_caps":  "SmallCaps",
    "superscript": "Superscript",
    "subscript":   "Subscript",
}

PAGE_SIZES = {          # width, height in twips (portrait)
    "letter":  (12240, 15840),
    "legal":   (12240, 20160),
    "tabloid": (15840, 24480),
    "a3":      (16838, 23811),
    "a4":      (11906, 16838),
    "a5":      (8391, 11906),
    "b5":      (9978, 14173),
    "executive": (10440, 15120),
}

NAMED_COLORS = {
    "black": "000000", "white": "FFFFFF", "red": "FF0000", "green": "008000",
    "blue": "0000FF", "yellow": "FFFF00", "gray": "808080", "grey": "808080",
    "lightgray": "D3D3D3", "lightgrey": "D3D3D3", "darkgray": "A9A9A9",
    "darkgrey": "A9A9A9", "navy": "000080", "maroon": "800000",
    "olive": "808000", "purple": "800080", "teal": "008080",
    "silver": "C0C0C0", "orange": "FFA500", "crimson": "DC143C",
    "gold": "FFD700", "pink": "FFC0CB", "brown": "A52A2A",
    "indigo": "4B0082", "violet": "EE82EE", "cyan": "00FFFF",
    "aqua": "00FFFF", "magenta": "FF00FF", "fuchsia": "FF00FF",
    "lime": "00FF00", "beige": "F5F5DC", "coral": "FF7F50",
    "salmon": "FA8072", "khaki": "F0E68C", "tan": "D2B48C",
    "ivory": "FFFFF0", "whitesmoke": "F5F5F5", "gainsboro": "DCDCDC",
    "steelblue": "4682B4", "royalblue": "4169E1", "seagreen": "2E8B57",
    "forestgreen": "228B22", "darkred": "8B0000", "darkblue": "00008B",
    "darkgreen": "006400", "slategray": "708090", "slategrey": "708090",
    "dimgray": "696969", "dimgrey": "696969",
}

ALIGNMENTS = {
    "left": "left", "start": "start", "right": "right", "end": "end",
    "center": "center", "centre": "center", "justify": "both",
    "both": "both", "distribute": "distribute",
}


# --------------------------------------------------------------------------
# Default configuration (also what ``--init`` writes out)
# --------------------------------------------------------------------------

DEFAULT_CONFIG_TOML = r"""# md2docx configuration
#
# Every section below is optional; anything you leave out keeps its built-in
# default.  Write `md2docx.py --init` to regenerate this file.
#
# Units
#   font sizes, spacing, indents, borders : points (bare number = points;
#                                           strings may use pt/in/cm/mm/px)
#   page size and margins                 : any of the above, e.g. "1in"
#   line_spacing                          : a multiple, e.g. 1.15
# Colors
#   "#dddddd", "dddddd", "#ddd", a CSS-ish name ("lightgrey"), or "auto".

[document]
page_size    = "letter"     # letter | legal | tabloid | a3 | a4 | a5 | b5
orientation  = "portrait"   # portrait | landscape
margin_top    = "1in"
margin_bottom = "1in"
margin_left   = "1in"
margin_right  = "1in"
default_font     = "Calibri"
default_size     = 11
default_language = "en-US"

[pandoc]
from                = "markdown"   # any pandoc reader, e.g. "gfm", "commonmark_x"
toc                 = false
toc_depth           = 3
number_sections     = false
syntax_highlighting = true         # false makes styles.code_block.color authoritative
# pygments | tango | kate | monochrome | breezedark | haddock | zenburn
highlight_style     = "tango"
extra_args          = []           # e.g. ["--shift-heading-level-by=1"]
filters             = []           # additional --lua-filter paths, applied after ours

# ---------------------------------------------------------------- inline ---
# Constructs marked `via_filter` are wrapped in a named character style by a
# generated Lua filter, because pandoc otherwise writes them as direct
# formatting that no reference document can reach.
#
#   via_filter  = false  -> leave pandoc's native output alone; this section
#                           is then ignored for that construct
#   keep_native = true   -> keep pandoc's direct formatting *and* add the
#                           style on top (safest for nested **bold *italic***)
#                 false  -> the style below is the only thing applied

# *emphasis*
[styles.emphasis]
via_filter  = true
keep_native = true
italic      = true

# **strong**
[styles.strong]
via_filter  = true
keep_native = true
bold        = true

# ~~strikeout~~
[styles.strikeout]
via_filter  = true
keep_native = true
strike      = true

# [underline]{.underline}
[styles.underline]
via_filter  = true
keep_native = true
underline   = "single"

# [small caps]{.smallcaps}
[styles.small_caps]
via_filter  = true
keep_native = true
small_caps  = true

# X^2^
[styles.superscript]
via_filter  = true
keep_native = true
vert_align  = "superscript"

# H~2~O
[styles.subscript]
via_filter  = true
keep_native = true
vert_align  = "subscript"

# `backticks` -- inline code
[styles.inline_code]
font       = "Consolas"
size       = 10
color      = "#1a1a1a"
background = "#dddddd"

# [links](https://example.com)
[styles.link]
color     = "#0563c1"
underline = "single"

[styles.footnote_reference]
vert_align = "superscript"

# ----------------------------------------------------------------- blocks ---

# Base style every other style inherits from.
[styles.normal]
font         = "Calibri"
size         = 11
color        = "#000000"
line_spacing = 1.15
space_after  = 8

[styles.body]
space_after = 8

# Paragraphs inside tight lists, table cells, etc.
[styles.compact]
space_before = 0
space_after  = 0

# ``` fenced code blocks ```
[styles.code_block]
font                = "Consolas"
size                = 10
color               = "#1a1a1a"
background          = "#dddddd"
line_spacing        = 1.0
space_before        = 0
space_after         = 0
contextual_spacing  = true
indent_left         = 6
indent_right        = 6

# A hairline box drawn around the whole run of code lines.  Comment the two
# lines below out for a plain grey band with no rule.
[styles.code_block.border]
left  = { style = "single", size = 0.75, color = "#bbbbbb", space = 4 }
right = { style = "single", size = 0.75, color = "#bbbbbb", space = 4 }

# > blockquotes
[styles.blockquote]
italic      = false
color       = "#404040"
indent_left = 36
space_after = 8
[styles.blockquote.border]
left = { style = "single", size = 2.25, color = "#cccccc", space = 8 }

[styles.title]
font         = "Calibri Light"
size         = 28
color        = "#1f3864"
alignment    = "center"
space_after  = 12

[styles.subtitle]
size      = 15
color     = "#595959"
italic    = true
alignment = "center"

[styles.heading1]
font           = "Calibri Light"
size           = 20
bold           = true
color          = "#1f3864"
space_before   = 18
space_after    = 6
keep_with_next = true

[styles.heading2]
font           = "Calibri Light"
size           = 16
bold           = true
color          = "#2e5496"
space_before   = 14
space_after    = 4
keep_with_next = true

[styles.heading3]
font           = "Calibri Light"
size           = 13
bold           = true
color          = "#2e5496"
space_before   = 12
space_after    = 4
keep_with_next = true

[styles.heading4]
size           = 11
bold           = true
italic         = true
color          = "#2e5496"
space_before   = 10
space_after    = 3
keep_with_next = true

[styles.heading5]
size           = 11
bold           = true
color          = "#404040"
space_before   = 10
space_after    = 3
keep_with_next = true

[styles.heading6]
size           = 11
italic         = true
color          = "#404040"
space_before   = 10
space_after    = 3
keep_with_next = true

[styles.caption]
size      = 9
italic    = true
color     = "#595959"
alignment = "center"

[styles.image_caption]
size      = 9
italic    = true
color     = "#595959"
alignment = "center"

[styles.table_caption]
size      = 9
italic    = true
color     = "#595959"

[styles.footnote_text]
size         = 9
line_spacing = 1.0
space_after  = 0

[styles.definition_term]
bold = true

[styles.definition]
indent_left = 18

# ----------------------------------------------------------------- tables ---
[table]
borders        = true
border_style   = "single"   # single | double | dashed | dotted | thick | none
border_size    = 0.5        # points
border_color   = "#999999"
inside_borders = true
cell_padding   = 5          # points, left/right
cell_padding_v = 2          # points, top/bottom

[table.header]
bold       = true
background = "#eeeeee"
"""


# --------------------------------------------------------------------------
# Small helpers
# --------------------------------------------------------------------------

class ConfigError(Exception):
    """A problem with the configuration, the input, or the environment."""


_UNITS_PER_PT = {
    "pt": 1.0, "px": 0.75, "in": 72.0, "cm": 72.0 / 2.54, "mm": 7.2 / 2.54,
    "pc": 12.0, "twip": 1.0 / 20.0, "twips": 1.0 / 20.0,
}


def to_points(value: object, what: str) -> float:
    """Convert a config length into points."""
    if isinstance(value, bool):
        msg = f"{what}: expected a length, got a boolean"
        raise ConfigError(msg)
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        m = re.fullmatch(r"\s*(-?\d*\.?\d+)\s*([a-z]*)\s*", value, re.IGNORECASE)
        if m:
            num, unit = float(m.group(1)), m.group(2).lower()
            if not unit:
                return num
            if unit in _UNITS_PER_PT:
                return num * _UNITS_PER_PT[unit]
    msg = f"{what}: cannot parse length {value!r}"
    raise ConfigError(msg)


def twips(value: object, what: str) -> int:
    """Convert a config length into twentieths of a point."""
    return round(to_points(value, what) * 20)


def half_points(value: object, what: str) -> int:
    """Convert a config length into half-points."""
    return round(to_points(value, what) * 2)


def eighth_points(value: object, what: str) -> int:
    """Convert a config length into eighth-points, clamped to the schema range."""
    return max(2, min(96, round(to_points(value, what) * 8)))


def to_color(value: object, what: str) -> str:
    """Normalize a config color into six-digit upper-case hex or ``auto``."""
    if not isinstance(value, str):
        msg = f"{what}: expected a color string, got {value!r}"
        raise ConfigError(msg)
    v = value.strip()
    if v.lower() in ("auto", "none"):
        return "auto"
    if v.lower() in NAMED_COLORS:
        return NAMED_COLORS[v.lower()]
    v = v.lstrip("#")
    if re.fullmatch(r"[0-9A-Fa-f]{3}", v):
        v = "".join(c * 2 for c in v)
    if not re.fullmatch(r"[0-9A-Fa-f]{6}", v):
        msg = (f'{what}: {value!r} is not a valid color; use a hex value such as '
               f'"#dddddd", "auto", or one of: {", ".join(sorted(NAMED_COLORS))}')
        raise ConfigError(msg)
    return v.upper()


def onoff(value: object, what: str) -> str:
    """Normalize a config boolean into an OOXML toggle value."""
    if isinstance(value, bool):
        return "1" if value else "0"
    if isinstance(value, str) and value.lower() in ("true", "yes", "on", "1"):
        return "1"
    if isinstance(value, str) and value.lower() in ("false", "no", "off", "0"):
        return "0"
    msg = f"{what}: expected true or false, got {value!r}"
    raise ConfigError(msg)


def deep_merge(base: dict, override: dict) -> dict:
    """Merge ``override`` into ``base`` recursively and return the result."""
    out = copy.deepcopy(base)
    for key, val in override.items():
        if isinstance(val, dict) and isinstance(out.get(key), dict):
            out[key] = deep_merge(out[key], val)
        else:
            out[key] = copy.deepcopy(val)
    return out


# --------------------------------------------------------------------------
# XML element helpers
# --------------------------------------------------------------------------

def ensure_child(parent: ET.Element, tag: str, order: list) -> ET.Element:
    """Find ``tag`` under ``parent``, creating it at the schema-correct spot."""
    found = parent.find(w(tag))
    if found is not None:
        return found
    el = ET.Element(w(tag))
    try:
        target = order.index(tag)
    except ValueError:
        parent.append(el)
        return el
    pos = len(parent)
    for i, child in enumerate(parent):
        name = child.tag.split("}")[-1]
        rank = order.index(name) if name in order else len(order)
        if rank > target:
            pos = i
            break
    parent.insert(pos, el)
    return el


def set_val(parent: ET.Element, tag: str, order: list, val: str) -> ET.Element:
    """Set the ``w:val`` attribute of ``tag`` under ``parent``, creating it."""
    el = ensure_child(parent, tag, order)
    el.set(w("val"), val)
    return el


# `on` mirrors the OOXML toggle this sets and is always a computed flag, so a
# positional boolean is the natural signature here.
def set_toggle(parent: ET.Element, tag: str, on: bool, order: list) -> None:  # noqa: FBT001
    """Switch an OOXML toggle property on or off."""
    el = ensure_child(parent, tag, order)
    if on:
        el.attrib.pop(w("val"), None)
    else:
        el.set(w("val"), "0")


def set_shading(parent: ET.Element, order: list, fill: str) -> None:
    """Fill the element's shading, clearing any theme attributes."""
    shd = ensure_child(parent, "shd", order)
    shd.set(w("val"), "clear")
    shd.set(w("color"), "auto")
    shd.set(w("fill"), fill)
    for name in ("themeFill", "themeFillTint", "themeFillShade",
                 "themeColor", "themeTint", "themeShade"):
        shd.attrib.pop(w(name), None)


def set_border(el: ET.Element, spec: object, what: str) -> None:
    """Populate a CT_Border element from a config mapping or shorthand."""
    if spec is False or spec is None:
        el.set(w("val"), "none")
        el.set(w("sz"), "0")
        el.set(w("space"), "0")
        el.set(w("color"), "auto")
        return
    if not isinstance(spec, dict):
        msg = f"{what}: expected a table of border properties"
        raise ConfigError(msg)
    el.set(w("val"), str(spec.get("style", "single")))
    el.set(w("sz"), str(eighth_points(spec.get("size", 0.5), what + ".size")))
    el.set(w("space"), str(round(to_points(spec.get("space", 0), what + ".space"))))
    el.set(w("color"), to_color(spec.get("color", "auto"), what + ".color"))


# --------------------------------------------------------------------------
# Applying a config style block to a docx style element
# --------------------------------------------------------------------------

RUN_KEYS = {
    "font", "font_ascii", "font_hansi", "font_cs", "font_east_asia", "size",
    "bold", "italic", "underline", "underline_color", "strike",
    "double_strike", "all_caps", "small_caps", "color", "background",
    "highlight", "vert_align", "letter_spacing", "position", "kerning",
    "outline", "shadow", "hidden_text", "language", "border",
}

PARA_KEYS = {
    "alignment", "space_before", "space_after", "space_before_auto",
    "space_after_auto", "line_spacing", "line_spacing_rule", "indent_left",
    "indent_right", "indent_first_line", "indent_hanging", "keep_with_next",
    "keep_lines", "page_break_before", "widow_control", "contextual_spacing",
    "outline_level", "paragraph_background", "border", "suppress_line_numbers",
}

META_KEYS = {"docx_style", "based_on", "next", "priority", "quick_format",
             "hidden", "via_filter", "keep_native"}


def drop_attrs(el: ET.Element, names: Iterable[str]) -> None:
    """Remove the named attributes from ``el``."""
    for name in names:
        el.attrib.pop(w(name), None)


# A theme reference takes precedence over the literal value sitting next to it,
# so any theme attribute has to be cleared or the configured value is ignored.
THEME_FONT_ATTRS = {"ascii": "asciiTheme", "hAnsi": "hAnsiTheme",
                    "cs": "cstheme", "eastAsia": "eastAsiaTheme"}
THEME_COLOR_ATTRS = ("themeColor", "themeTint", "themeShade")
THEME_FILL_ATTRS = ("themeFill", "themeFillTint", "themeFillShade")


def apply_run_props(  # noqa: C901, PLR0912 — one branch per run property reads best as a unit
        rpr: ET.Element, cfg: dict, what: str, *, run_border: bool = True) -> None:
    """Apply the run (character) settings in ``cfg`` to an ``rPr`` element."""
    fonts = {}
    if "font" in cfg:
        for slot in ("ascii", "hAnsi", "cs", "eastAsia"):
            fonts[slot] = str(cfg["font"])
    for key, slot in (("font_ascii", "ascii"), ("font_hansi", "hAnsi"),
                      ("font_cs", "cs"), ("font_east_asia", "eastAsia")):
        if key in cfg:
            fonts[slot] = str(cfg[key])
    if fonts:
        rf = ensure_child(rpr, "rFonts", RPR_ORDER)
        for slot, name in fonts.items():
            rf.set(w(slot), name)
            drop_attrs(rf, [THEME_FONT_ATTRS[slot]])

    if "size" in cfg:
        hp = str(half_points(cfg["size"], what + ".size"))
        set_val(rpr, "sz", RPR_ORDER, hp)
        set_val(rpr, "szCs", RPR_ORDER, hp)

    for key, tags in (("bold", ("b", "bCs")), ("italic", ("i", "iCs")),
                      ("all_caps", ("caps",)), ("small_caps", ("smallCaps",)),
                      ("strike", ("strike",)), ("double_strike", ("dstrike",)),
                      ("outline", ("outline",)), ("shadow", ("shadow",)),
                      ("hidden_text", ("vanish",))):
        if key in cfg:
            on = onoff(cfg[key], what + "." + key) == "1"
            for tag in tags:
                set_toggle(rpr, tag, on, RPR_ORDER)

    if "color" in cfg:
        el = set_val(rpr, "color", RPR_ORDER, to_color(cfg["color"], what + ".color"))
        drop_attrs(el, THEME_COLOR_ATTRS)

    if "background" in cfg:
        set_shading(rpr, RPR_ORDER, to_color(cfg["background"], what + ".background"))

    if "highlight" in cfg:
        set_val(rpr, "highlight", RPR_ORDER, str(cfg["highlight"]))

    if "underline" in cfg:
        val = cfg["underline"]
        if isinstance(val, bool):
            val = "single" if val else "none"
        u = set_val(rpr, "u", RPR_ORDER, str(val))
        if "underline_color" in cfg:
            u.set(w("color"), to_color(cfg["underline_color"], what + ".underline_color"))

    if "vert_align" in cfg:
        set_val(rpr, "vertAlign", RPR_ORDER, str(cfg["vert_align"]))

    if "letter_spacing" in cfg:
        set_val(rpr, "spacing", RPR_ORDER,
                str(twips(cfg["letter_spacing"], what + ".letter_spacing")))

    if "position" in cfg:
        set_val(rpr, "position", RPR_ORDER,
                str(half_points(cfg["position"], what + ".position")))

    if "kerning" in cfg:
        set_val(rpr, "kern", RPR_ORDER,
                str(half_points(cfg["kerning"], what + ".kerning")))

    if "language" in cfg:
        ensure_child(rpr, "lang", RPR_ORDER).set(w("val"), str(cfg["language"]))

    # For paragraph styles `border` means paragraph borders, handled elsewhere.
    if run_border and "border" in cfg:
        set_border(ensure_child(rpr, "bdr", RPR_ORDER), cfg["border"], what + ".border")


def apply_para_props(  # noqa: C901, PLR0912, PLR0915 — one branch per paragraph property
        ppr: ET.Element, cfg: dict, what: str) -> None:
    """Apply the paragraph settings in ``cfg`` to a ``pPr`` element."""
    for key, tag in (("keep_with_next", "keepNext"), ("keep_lines", "keepLines"),
                     ("page_break_before", "pageBreakBefore"),
                     ("widow_control", "widowControl"),
                     ("contextual_spacing", "contextualSpacing"),
                     ("suppress_line_numbers", "suppressLineNumbers")):
        if key in cfg:
            set_toggle(ppr, tag, onoff(cfg[key], what + "." + key) == "1", PPR_ORDER)

    if "paragraph_background" in cfg:
        set_shading(ppr, PPR_ORDER,
                    to_color(cfg["paragraph_background"], what + ".paragraph_background"))
    elif "background" in cfg:
        # A run-level shade alone leaves gaps between words in a code block, so
        # paragraph styles get the same fill behind the whole line.
        set_shading(ppr, PPR_ORDER, to_color(cfg["background"], what + ".background"))

    spacing_keys = {"space_before", "space_after", "line_spacing",
                    "space_before_auto", "space_after_auto"}
    if spacing_keys & cfg.keys():
        sp = ensure_child(ppr, "spacing", PPR_ORDER)
        if "space_before" in cfg:
            sp.set(w("before"), str(twips(cfg["space_before"], what + ".space_before")))
            sp.set(w("beforeAutospacing"), "0")
        if "space_after" in cfg:
            sp.set(w("after"), str(twips(cfg["space_after"], what + ".space_after")))
            sp.set(w("afterAutospacing"), "0")
        if "space_before_auto" in cfg:
            sp.set(w("beforeAutospacing"),
                   onoff(cfg["space_before_auto"], what + ".space_before_auto"))
        if "space_after_auto" in cfg:
            sp.set(w("afterAutospacing"),
                   onoff(cfg["space_after_auto"], what + ".space_after_auto"))
        if "line_spacing" in cfg:
            rule = str(cfg.get("line_spacing_rule", "auto"))
            if rule == "auto":
                sp.set(w("line"), str(round(
                    to_points(cfg["line_spacing"], what + ".line_spacing") * 240)))
            else:
                sp.set(w("line"), str(twips(cfg["line_spacing"], what + ".line_spacing")))
            sp.set(w("lineRule"), rule)

    indent_keys = {"indent_left", "indent_right", "indent_first_line", "indent_hanging"}
    if indent_keys & cfg.keys():
        ind = ensure_child(ppr, "ind", PPR_ORDER)
        for key, attr in (("indent_left", "left"), ("indent_right", "right"),
                          ("indent_first_line", "firstLine"),
                          ("indent_hanging", "hanging")):
            if key in cfg:
                ind.set(w(attr), str(twips(cfg[key], what + "." + key)))

    if "alignment" in cfg:
        val = str(cfg["alignment"]).lower()
        if val not in ALIGNMENTS:
            msg = f"{what}.alignment: unknown value {cfg['alignment']!r}"
            raise ConfigError(msg)
        set_val(ppr, "jc", PPR_ORDER, ALIGNMENTS[val])

    if "outline_level" in cfg:
        set_val(ppr, "outlineLvl", PPR_ORDER, str(int(cfg["outline_level"])))

    if "border" in cfg:
        spec = cfg["border"]
        if not isinstance(spec, dict):
            msg = f"{what}.border: expected a table of sides"
            raise ConfigError(msg)
        unknown = set(spec) - set(BORDER_SIDES_P) - {"all"}
        if unknown:
            msg = (f"{what}.border: unknown side(s): {', '.join(sorted(unknown))} "
                   f"(expected any of {', '.join(['all', *BORDER_SIDES_P])})")
            raise ConfigError(msg)
        pbdr = ensure_child(ppr, "pBdr", PPR_ORDER)
        # `all` means the four edges; `between` and `bar` must be asked for by
        # name, since a rule between consecutive paragraphs is rarely wanted.
        sides = dict.fromkeys(["top", "left", "bottom", "right"],
                              spec["all"]) if "all" in spec else {}
        for side in BORDER_SIDES_P:
            if side in spec:
                sides[side] = spec[side]
        for side in BORDER_SIDES_P:          # schema order
            if side in sides:
                set_border(ensure_child(pbdr, side, BORDER_SIDES_P),
                           sides[side], f"{what}.border.{side}")


def apply_style(style_el: ET.Element, cfg: dict, stype: str, what: str) -> None:
    """Apply one config style block to a docx style element."""
    unknown = set(cfg) - RUN_KEYS - PARA_KEYS - META_KEYS
    if unknown:
        msg = f"{what}: unknown setting(s): {', '.join(sorted(unknown))}"
        raise ConfigError(msg)

    if "based_on" in cfg:
        set_val(style_el, "basedOn", STYLE_ORDER, style_id_for(cfg["based_on"]))
    if "next" in cfg:
        set_val(style_el, "next", STYLE_ORDER, style_id_for(cfg["next"]))
    if "priority" in cfg:
        set_val(style_el, "uiPriority", STYLE_ORDER, str(int(cfg["priority"])))
    if "quick_format" in cfg:
        set_toggle(style_el, "qFormat",
                   onoff(cfg["quick_format"], what + ".quick_format") == "1", STYLE_ORDER)
    if "hidden" in cfg:
        set_toggle(style_el, "semiHidden",
                   onoff(cfg["hidden"], what + ".hidden") == "1", STYLE_ORDER)

    if stype == "paragraph" and (PARA_KEYS & cfg.keys() or "background" in cfg):
        apply_para_props(ensure_child(style_el, "pPr", STYLE_ORDER), cfg, what)
    if RUN_KEYS & cfg.keys():
        apply_run_props(ensure_child(style_el, "rPr", STYLE_ORDER), cfg, what,
                        run_border=(stype == "character"))


def style_id_for(name: str) -> str:
    """Derive the OOXML style id from a style name."""
    return re.sub(r"[^0-9A-Za-z]", "", str(name))


# --------------------------------------------------------------------------
# Reference document construction
# --------------------------------------------------------------------------

NS_RE = re.compile(r'xmlns:([A-Za-z0-9_.\-]+)\s*=\s*"([^"]+)"')


def register_namespaces(xml_text: str) -> None:
    """Register the namespace prefixes of ``xml_text`` so they round-trip."""
    for prefix, uri in NS_RE.findall(xml_text[:4000]):
        with contextlib.suppress(ValueError):
            ET.register_namespace(prefix, uri)


def serialize(tree_root: ET.Element, original: str, root_name: str) -> bytes:
    """Serialize, restoring the original root tag so no xmlns is dropped."""
    body = ET.tostring(tree_root, encoding="unicode")
    original_root = re.search(rf"<{re.escape(root_name)}\b[^>]*>", original)
    if original_root:
        body = re.sub(rf"^<{re.escape(root_name)}\b[^>]*>",
                      lambda _m: original_root.group(0), body, count=1)
    decl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    return (decl + body).encode("utf-8")


def find_style_el(root: ET.Element, name: str, stype: str) -> ET.Element | None:
    """Find a style by name, falling back to its derived style id."""
    want = name.lower()
    sid = style_id_for(name).lower()
    fallback = None
    for st in root.findall(w("style")):
        if st.get(w("type")) != stype:
            continue
        nm = st.find(w("name"))
        if nm is not None and (nm.get(w("val")) or "").lower() == want:
            return st
        if (st.get(w("styleId")) or "").lower() == sid:
            fallback = st
    return fallback


def create_style_el(root: ET.Element, name: str, stype: str) -> ET.Element:
    """Create a new style element with sensible defaults."""
    st = ET.SubElement(root, w("style"))
    st.set(w("type"), stype)
    st.set(w("styleId"), style_id_for(name))
    ET.SubElement(st, w("name")).set(w("val"), name)
    base = {"character": "DefaultParagraphFont", "paragraph": "Normal"}.get(stype)
    if base and find_style_el(root, base, stype) is not None:
        ET.SubElement(st, w("basedOn")).set(w("val"), style_id_for(base))
    ET.SubElement(st, w("qFormat"))
    return st


def apply_table_config(root: ET.Element, cfg: dict) -> None:  # noqa: C901 — one branch per table setting
    """Apply the ``[table]`` section to the Table style."""
    style = find_style_el(root, "Table", "table")
    if style is None:
        style = create_style_el(root, "Table", "table")
    tblpr = ensure_child(style, "tblPr", STYLE_ORDER)

    borders_on = cfg.get("borders", True)
    spec = {
        "style": cfg.get("border_style", "single"),
        "size": cfg.get("border_size", 0.5),
        "color": cfg.get("border_color", "#999999"),
        "space": 0,
    }
    tbl_borders = ensure_child(tblpr, "tblBorders", TBLPR_ORDER)
    for side in BORDER_SIDES_T:
        inside = side.startswith("inside")
        on = borders_on and (cfg.get("inside_borders", True) if inside else True)
        set_border(ensure_child(tbl_borders, side, BORDER_SIDES_T),
                   spec if on else False, "table.border")

    if "cell_padding" in cfg or "cell_padding_v" in cfg:
        mar = ensure_child(tblpr, "tblCellMar", TBLPR_ORDER)
        pad_h = twips(cfg.get("cell_padding", 5), "table.cell_padding")
        pad_v = twips(cfg.get("cell_padding_v", 0), "table.cell_padding_v")
        for side, val in (("top", pad_v), ("left", pad_h),
                          ("bottom", pad_v), ("right", pad_h)):
            el = ensure_child(mar, side, ["top", "left", "bottom", "right"])
            el.set(w("w"), str(val))
            el.set(w("type"), "dxa")

    header = cfg.get("header") or {}
    if header:
        cond = None
        for existing in style.findall(w("tblStylePr")):
            if existing.get(w("type")) == "firstRow":
                cond = existing
                break
        if cond is None:
            cond = ensure_child(style, "tblStylePr", STYLE_ORDER)
            cond.set(w("type"), "firstRow")
        # `background` becomes cell shading below; run shading would be redundant.
        run_cfg = {k: v for k, v in header.items()
                   if k in RUN_KEYS and k != "background"}
        if run_cfg:
            apply_run_props(ensure_child(cond, "rPr", ["pPr", "rPr", "tblPr",
                                                       "trPr", "tcPr"]),
                            run_cfg, "table.header")
        if "background" in header:
            tcpr = ensure_child(cond, "tcPr", ["pPr", "rPr", "tblPr", "trPr", "tcPr"])
            set_shading(tcpr, ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge",
                               "tcBorders", "shd"],
                        to_color(header["background"], "table.header.background"))


def build_styles_xml(source: str, config: dict) -> bytes:  # noqa: C901 — cohesive rewrite pass
    """Rewrite ``styles.xml`` according to the configuration."""
    register_namespaces(source)
    root = ET.fromstring(source)  # noqa: S314 — parses the user's own reference document

    doc = config.get("document") or {}
    if {"default_font", "default_size", "default_language"} & doc.keys():
        docdefaults = ensure_child(root, "docDefaults", ["docDefaults", "latentStyles"])
        rpr_default = ensure_child(docdefaults, "rPrDefault",
                                   ["rPrDefault", "pPrDefault"])
        rpr = ensure_child(rpr_default, "rPr", ["rPr"])
        defaults = {}
        if "default_font" in doc:
            defaults["font"] = doc["default_font"]
        if "default_size" in doc:
            defaults["size"] = doc["default_size"]
        if "default_language" in doc:
            defaults["language"] = doc["default_language"]
        apply_run_props(rpr, defaults, "document")

    for key, block in (config.get("styles") or {}).items():
        if key not in STYLE_MAP:
            msg = f"styles.{key}: unknown style key (see --list-styles)"
            raise ConfigError(msg)
        if not isinstance(block, dict):
            msg = f"styles.{key}: expected a table of settings"
            raise ConfigError(msg)
        name, stype = STYLE_MAP[key]
        if "docx_style" in block:
            if key not in FILTERED_INLINES:
                msg = (f"styles.{key}.docx_style: only the filter-driven inline "
                       f"styles ({', '.join(sorted(FILTERED_INLINES))}) can be "
                       f"retargeted; pandoc picks the style name {name!r} for "
                       f"{key} itself.")
                raise ConfigError(msg)
            name = block["docx_style"]
        if stype == "table":
            continue
        style_el = find_style_el(root, name, stype)
        if style_el is None:
            style_el = create_style_el(root, name, stype)
        apply_style(style_el, block, stype, "styles." + key)

    if config.get("table"):
        apply_table_config(root, config["table"])

    return serialize(root, source, "w:styles")


def build_document_xml(source: str, doc_cfg: dict) -> bytes | None:  # noqa: C901 — one branch per page setting
    """Rewrite the section properties of ``document.xml``, where configured."""
    keys = {"page_size", "page_width", "page_height", "orientation",
            "margin_top", "margin_bottom", "margin_left", "margin_right",
            "margin_header", "margin_footer", "margin_gutter"}
    if not keys & doc_cfg.keys():
        return None

    register_namespaces(source)
    root = ET.fromstring(source)  # noqa: S314 — parses the user's own reference document
    body = root.find(w("body"))
    if body is None:
        return None
    sect = body.find(w("sectPr"))
    if sect is None:
        sect = ET.SubElement(body, w("sectPr"))

    sect_order = ["headerReference", "footerReference", "footnotePr", "endnotePr",
                  "type", "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType",
                  "pgNumType", "cols", "formProt", "vAlign", "noEndnote",
                  "titlePg", "textDirection", "bidi", "rtlGutter", "docGrid",
                  "printerSettings"]

    size_name = str(doc_cfg.get("page_size", "letter")).lower()
    if size_name not in PAGE_SIZES:
        msg = (f"document.page_size: unknown size {doc_cfg.get('page_size')!r} "
               f"(known: {', '.join(sorted(PAGE_SIZES))})")
        raise ConfigError(msg)
    width, height = PAGE_SIZES[size_name]
    if "page_width" in doc_cfg:
        width = twips(doc_cfg["page_width"], "document.page_width")
    if "page_height" in doc_cfg:
        height = twips(doc_cfg["page_height"], "document.page_height")
    landscape = str(doc_cfg.get("orientation", "portrait")).lower() == "landscape"
    if landscape:
        width, height = height, width

    pgsz = ensure_child(sect, "pgSz", sect_order)
    pgsz.set(w("w"), str(width))
    pgsz.set(w("h"), str(height))
    pgsz.set(w("orient"), "landscape" if landscape else "portrait")

    pgmar = ensure_child(sect, "pgMar", sect_order)
    for key, attr, default in (("margin_top", "top", None),
                               ("margin_right", "right", None),
                               ("margin_bottom", "bottom", None),
                               ("margin_left", "left", None),
                               ("margin_header", "header", None),
                               ("margin_footer", "footer", None),
                               ("margin_gutter", "gutter", None)):
        if key in doc_cfg:
            pgmar.set(w(attr), str(twips(doc_cfg[key], "document." + key)))
        elif pgmar.get(w(attr)) is None and default is not None:
            pgmar.set(w(attr), str(default))

    return serialize(root, source, "w:document")


def build_reference_docx(base: bytes, config: dict, dest: Path) -> None:
    """Write a copy of ``base`` with the configured styles into ``dest``."""
    with zipfile.ZipFile(io.BytesIO(base)) as zin:
        entries = [(info, zin.read(info.filename)) for info in zin.infolist()]

    styles_src = next((data.decode("utf-8") for info, data in entries
                       if info.filename == "word/styles.xml"), None)
    if styles_src is None:
        msg = "reference document contains no word/styles.xml"
        raise ConfigError(msg)
    new_styles = build_styles_xml(styles_src, config)

    doc_src = next((data.decode("utf-8") for info, data in entries
                    if info.filename == "word/document.xml"), None)
    new_doc = build_document_xml(doc_src, config.get("document") or {}) if doc_src else None

    with zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as zout:
        for info, data in entries:
            if info.filename == "word/styles.xml":
                payload = new_styles
            elif info.filename == "word/document.xml" and new_doc is not None:
                payload = new_doc
            else:
                payload = data
            out = zipfile.ZipInfo(info.filename, date_time=info.date_time)
            out.compress_type = zipfile.ZIP_DEFLATED
            out.external_attr = info.external_attr
            zout.writestr(out, payload)


# --------------------------------------------------------------------------
# Lua filter construction
# --------------------------------------------------------------------------

LUA_HEADER = """-- Generated by md2docx; edits will be overwritten.
-- Routes inline constructs that pandoc renders with direct formatting through
-- named character styles so the configuration can control them.
local styles = {
%s}
local keep_native = {
%s}

local function restyle(el, tag)
  local name = styles[tag]
  if name == nil then return nil end
  local span = pandoc.Span(el.content, { ['custom-style'] = name })
  if keep_native[tag] then
    el.content = { span }
    return el
  end
  return span
end

%s
"""


def build_lua_filter(config: dict) -> str:
    """Generate the Lua filter for the inline constructs marked ``via_filter``."""
    styles_cfg = config.get("styles") or {}
    style_lines, keep_lines, fn_lines = [], [], []
    for key, element in FILTERED_INLINES.items():
        block = styles_cfg.get(key) or {}
        if not block.get("via_filter", False):
            continue
        name = block.get("docx_style", STYLE_MAP[key][0])
        style_lines.append(f"  {element} = {lua_string(name)},\n")
        keep = "true" if block.get("keep_native", True) else "false"
        keep_lines.append(f"  {element} = {keep},\n")
        fn_lines.append(f"function {element}(el) return restyle(el, '{element}') end")
    if not fn_lines:
        return ""
    # The Lua template is full of literal braces, so %-formatting stays readable
    # where an f-string or str.format would need every brace doubled.
    return LUA_HEADER % (
        "".join(style_lines), "".join(keep_lines), "\n".join(fn_lines))


def lua_string(value: str) -> str:
    """Quote ``value`` as a single-quoted Lua string literal."""
    return "'" + str(value).replace("\\", "\\\\").replace("'", "\\'") + "'"


# --------------------------------------------------------------------------
# Configuration loading
# --------------------------------------------------------------------------

def load_toml(text: str, origin: str) -> dict:
    """Parse TOML text, via the standard library or the ``tomli`` fallback."""
    try:
        # Deferred import: JSON and YAML configs must keep working without it.
        import tomllib  # noqa: PLC0415
    except ModuleNotFoundError:                     # Python < 3.11
        try:
            import tomli as tomllib  # noqa: PLC0415
        except ModuleNotFoundError:
            msg = (f"{origin}: reading TOML needs Python 3.11+ or the 'tomli' "
                   "package; alternatively use a .json or .yaml config")
            raise ConfigError(msg) from None
    try:
        return tomllib.loads(text)
    except Exception as exc:
        msg = f"{origin}: {exc}"
        raise ConfigError(msg) from exc


def load_config_text(path: Path) -> dict:
    """Load a configuration mapping from a TOML, YAML or JSON file."""
    text = path.read_text(encoding="utf-8")
    suffix = path.suffix.lower()
    if suffix in (".yaml", ".yml"):
        try:
            # Deferred import: PyYAML is optional and only YAML configs need it.
            import yaml  # noqa: PLC0415
        except ModuleNotFoundError:
            msg = f"{path}: YAML configs need PyYAML (pip install pyyaml)"
            raise ConfigError(msg) from None
        try:
            data = yaml.safe_load(text) or {}
        except Exception as exc:
            msg = f"{path}: {exc}"
            raise ConfigError(msg) from exc
    elif suffix == ".json":
        try:
            data = json.loads(text)
        except Exception as exc:
            msg = f"{path}: {exc}"
            raise ConfigError(msg) from exc
    else:
        data = load_toml(text, str(path))
    if not isinstance(data, dict):
        msg = f"{path}: top level must be a table/object"
        raise ConfigError(msg)
    return data


CONFIG_NAMES = ["md2docx.toml", "md2docx.yaml", "md2docx.yml", "md2docx.json",
                ".md2docx.toml"]

USER_CONFIG_NAMES = ["default.toml", "default.yaml", "default.yml",
                     "default.json"]


def user_config_dir() -> Path:
    """Return the per-user stylesheet directory: ~/.config/md2docx by default."""
    base = os.environ.get("XDG_CONFIG_HOME")
    return (Path(base) if base else Path.home() / ".config") / "md2docx"


def discover_config(start: Path) -> Path | None:
    """Find the nearest project configuration, else the per-user default."""
    for directory in [start, *list(start.parents)]:
        for name in CONFIG_NAMES:
            candidate = directory / name
            if candidate.is_file():
                return candidate
        if (directory / ".git").exists():
            break
    for name in USER_CONFIG_NAMES:
        candidate = user_config_dir() / name
        if candidate.is_file():
            return candidate
    return None


def resolve_config(explicit: str | None, search_from: Path, *,
                   quiet: bool) -> tuple[dict, Path | None]:
    """Merge the built-in defaults with the discovered or named config file."""
    defaults = load_toml(DEFAULT_CONFIG_TOML, "<built-in defaults>")
    path: Path | None
    if explicit is not None:
        path = Path(explicit)
        if not path.is_file():
            msg = f"config file not found: {path}"
            raise ConfigError(msg)
    else:
        path = discover_config(search_from)
    if path is None:
        return defaults, None
    merged = deep_merge(defaults, load_config_text(path))
    if not quiet:
        print(f"md2docx: using config {path}", file=sys.stderr)  # noqa: T201 — progress
    return merged, path


# --------------------------------------------------------------------------
# Pandoc invocation
# --------------------------------------------------------------------------

def pandoc_path(explicit: str | None = None) -> str:
    """Return the pandoc executable to run, or fail with installation advice."""
    found = explicit or shutil.which("pandoc")
    if not found:
        msg = ("pandoc was not found on PATH.  Install it (apt install pandoc, "
               "brew install pandoc, or https://pandoc.org/installing.html) or "
               "pass --pandoc /path/to/pandoc.")
        raise ConfigError(msg)
    return found


def default_reference_docx(pandoc: str) -> bytes:
    """Read pandoc's built-in reference.docx."""
    # S603: fixed arguments; pandoc is the converter the user asked us to run.
    proc = subprocess.run([pandoc, "--print-default-data-file", "reference.docx"],  # noqa: S603
                          capture_output=True, check=False)
    if proc.returncode != 0:
        msg = ("could not read pandoc's default reference.docx: "
               f"{proc.stderr.decode('utf-8', 'replace').strip()}")
        raise ConfigError(msg)
    return proc.stdout


def build_pandoc_args(  # noqa: PLR0913, PLR0917 — one parameter per ingredient of the command line
        args: argparse.Namespace, config: dict, reference: Path,
        lua: Path | None, inputs: list, output: Path) -> list:
    """Assemble the full pandoc command line."""
    pcfg = config.get("pandoc") or {}
    cmd = [args.pandoc_bin]

    reader = args.from_format or pcfg.get("from", "markdown")
    cmd += ["--from", reader, "--to", "docx", "--standalone"]
    cmd += ["--reference-doc", str(reference)]
    if lua is not None:
        cmd += ["--lua-filter", str(lua)]
    for extra in pcfg.get("filters") or []:
        cmd += ["--lua-filter", str(extra)]

    if args.toc or (args.toc is None and pcfg.get("toc", False)):
        cmd += ["--toc", "--toc-depth", str(pcfg.get("toc_depth", 3))]
    if args.number_sections or (args.number_sections is None
                                and pcfg.get("number_sections", False)):
        cmd.append("--number-sections")

    if pcfg.get("syntax_highlighting", True):
        cmd += ["--highlight-style", str(pcfg.get("highlight_style", "tango"))]
    else:
        cmd.append("--no-highlight")

    for key, value in (pcfg.get("metadata") or {}).items():
        cmd += ["--metadata", f"{key}={value}"]

    cmd += [str(a) for a in (pcfg.get("extra_args") or [])]
    cmd += list(args.pandoc_args)

    resource_dirs = {str(Path(p).resolve().parent) for p in inputs if p != "-"}
    if resource_dirs:
        cmd += ["--resource-path", os.pathsep.join(sorted(resource_dirs))]

    cmd += ["--output", str(output)]
    cmd += list(inputs)
    return cmd


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------

def make_parser() -> argparse.ArgumentParser:
    """Build the command-line argument parser."""
    p = argparse.ArgumentParser(
        prog="md2docx",
        description="Convert Markdown to .docx with formatting driven by a config file.",
        epilog="Anything after -- is passed straight through to pandoc.")
    p.add_argument("inputs", nargs="*", metavar="FILE",
                   help="Markdown input file(s); '-' reads standard input")
    p.add_argument("-o", "--output", metavar="FILE",
                   help="output .docx (default: first input with a .docx suffix)")
    p.add_argument("-c", "--config", metavar="FILE",
                   help="configuration file (.toml, .yaml or .json)")
    p.add_argument("--init", nargs="?", const="md2docx.toml", metavar="FILE",
                   help="write the default configuration to FILE and exit "
                        "(default ./md2docx.toml; 'user' means "
                        "~/.config/md2docx/default.toml)")
    p.add_argument("--list-styles", action="store_true",
                   help="list the configurable style keys and exit")
    p.add_argument("--install", nargs="?", const=True, metavar="DIR",
                   help="copy this script to DIR/md2docx and exit; without DIR, "
                        "the first writable directory on $PATH is used")
    p.add_argument("--force", action="store_true",
                   help="with --install, replace an existing unrelated file")
    p.add_argument("--print-reference", metavar="FILE",
                   help="write the generated reference.docx to FILE and exit")
    p.add_argument("--print-filter", action="store_true",
                   help="print the generated Lua filter and exit")
    p.add_argument("--reference-doc", metavar="FILE",
                   help="start from this .docx instead of pandoc's default")
    p.add_argument("--from", dest="from_format", metavar="FORMAT",
                   help="override the pandoc reader (default: config [pandoc] from)")
    p.add_argument("--toc", dest="toc", action="store_true", default=None,
                   help="add a table of contents")
    p.add_argument("--no-toc", dest="toc", action="store_false",
                   help="suppress the table of contents")
    p.add_argument("--number-sections", dest="number_sections",
                   action="store_true", default=None,
                   help="number section headings")
    p.add_argument("--pandoc", dest="pandoc_bin", metavar="PATH",
                   help="path to the pandoc executable")
    p.add_argument("--keep-temp", action="store_true",
                   help="keep (and report) the generated reference doc and filter")
    p.add_argument("-n", "--dry-run", action="store_true",
                   help="show the pandoc command without running it")
    p.add_argument("-v", "--verbose", action="store_true",
                   help="print the pandoc command before running it")
    p.add_argument("-q", "--quiet", action="store_true",
                   help="suppress progress messages")
    p.add_argument("--version", action="version",
                   version=f"md2docx {__version__}")
    return p


def split_pandoc_args(argv: list) -> tuple[list, list]:
    """Split ``argv`` at ``--`` into md2docx arguments and pandoc passthrough."""
    if "--" in argv:
        idx = argv.index("--")
        return argv[:idx], argv[idx + 1:]
    return argv, []


def list_styles() -> None:
    """Print the configurable style keys."""
    print("Configurable style keys ([styles.KEY] in the config file):\n")  # noqa: T201 — output
    width = max(len(k) for k in STYLE_MAP)
    for key in sorted(STYLE_MAP, key=lambda k: (STYLE_MAP[k][1], k)):
        name, stype = STYLE_MAP[key]
        mark = "  (via Lua filter)" if key in FILTERED_INLINES else ""
        print(f"  {key:<{width}}  {stype:<10} -> {name}{mark}")  # noqa: T201 — output
    print("\nTable formatting lives in its own [table] section.")  # noqa: T201 — output


def path_dirs() -> Iterator[Path]:
    """Yield the $PATH components, in order, that we could install into."""
    for raw in os.environ.get("PATH", "").split(os.pathsep):
        if not raw:
            continue                      # an empty entry means "."
        yield Path(raw).expanduser()


def writable_path_dirs() -> Iterator[Path]:
    """Yield the $PATH directories we can actually write into."""
    for directory in path_dirs():
        if directory.is_dir() and os.access(directory, os.W_OK | os.X_OK):
            yield directory


def looks_like_md2docx(path: Path) -> bool:
    """Report whether ``path`` looks like an installed copy of this script."""
    try:
        with path.open("rb") as handle:
            return b"DEFAULT_CONFIG_TOML" in handle.read(200_000)
    except OSError:
        return False


def install(target_dir: str | None, *, force: bool, quiet: bool) -> int:
    """Copy this script to ``target_dir`` or the first writable $PATH entry."""
    source = Path(__file__).resolve()

    dest_dir: Path | None
    if target_dir is not None:
        dest_dir = Path(target_dir).expanduser()
        if not dest_dir.is_dir():
            msg = f"install directory does not exist: {dest_dir}"
            raise ConfigError(msg)
        if not os.access(dest_dir, os.W_OK | os.X_OK):
            msg = f"install directory is not writable: {dest_dir}"
            raise ConfigError(msg)
    else:
        dest_dir = next(iter(writable_path_dirs()), None)
        if dest_dir is None:
            msg = (
                "no writable directory on PATH.  Either name one:\n"
                "    md2docx --install ~/bin\n"
                "or create the usual one and put it on your PATH:\n"
                "    mkdir -p ~/.local/bin\n"
                '    export PATH="$HOME/.local/bin:$PATH"')
            raise ConfigError(msg)

    dest = dest_dir / "md2docx"

    if dest.is_symlink() or dest.exists():
        if dest.resolve() == source:
            print(f"md2docx: {dest} already points at this script")  # noqa: T201 — output
            return 0
        if not force and not looks_like_md2docx(dest):
            msg = (f"{dest} already exists and does not look like md2docx; "
                   "pass --force to replace it")
            raise ConfigError(msg)
        dest.unlink()                     # replace, never write through a symlink

    shutil.copyfile(source, dest)
    dest.chmod(0o755)
    print(f"md2docx: installed {dest}")  # noqa: T201 — output

    found = shutil.which("md2docx")
    if found and Path(found).resolve() != dest.resolve():
        print(  # noqa: T201 — warning for the person installing
            f"md2docx: warning: {found} comes earlier on PATH and will shadow it",
            file=sys.stderr)
    elif not quiet:
        print("md2docx: run 'md2docx --init user' to create "  # noqa: T201 — output
              f"{user_config_dir() / 'default.toml'}")
    return 0


# run() is the whole command-line surface: a flat sequence of early-exit modes
# followed by the conversion, which reads better than it would split apart.
def run(argv: list) -> int:  # noqa: C901, PLR0911, PLR0912, PLR0915
    """Run the conversion, or one of the utility modes, and return a status."""
    argv, passthrough = split_pandoc_args(argv)
    parser = make_parser()
    args = parser.parse_args(argv)
    args.pandoc_args = passthrough

    if args.list_styles:
        list_styles()
        return 0

    if args.install:
        target = None if args.install is True else args.install
        return install(target, force=args.force, quiet=args.quiet)

    if args.init:
        dest = Path(args.init).expanduser()
        if args.init == "user":
            dest = user_config_dir() / "default.toml"
        if dest.exists():
            print(f"md2docx: {dest} already exists; not overwriting",  # noqa: T201 — error report
                  file=sys.stderr)
            return 1
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_text(DEFAULT_CONFIG_TOML, encoding="utf-8")
        print(f"md2docx: wrote {dest}")  # noqa: T201 — output
        return 0

    search_from = Path(args.inputs[0]).resolve().parent if (
        args.inputs and args.inputs[0] != "-") else Path.cwd()
    config, _config_path = resolve_config(args.config, search_from, quiet=args.quiet)

    if args.print_filter:
        sys.stdout.write(build_lua_filter(config) or
                         "-- no inline constructs are routed through the filter\n")
        return 0

    args.pandoc_bin = pandoc_path(args.pandoc_bin)

    if args.reference_doc:
        base = Path(args.reference_doc).read_bytes()
    else:
        base = default_reference_docx(args.pandoc_bin)

    if args.print_reference:
        build_reference_docx(base, config, Path(args.print_reference))
        print(f"md2docx: wrote {args.print_reference}")  # noqa: T201 — output
        return 0

    if not args.inputs:
        parser.error("no input files (use - to read standard input)")

    for name in args.inputs:
        if name != "-" and not Path(name).is_file():
            msg = f"input file not found: {name}"
            raise ConfigError(msg)

    if args.output:
        output = Path(args.output)
    elif args.inputs[0] != "-":
        output = Path(args.inputs[0]).with_suffix(".docx")
    else:
        parser.error("--output is required when reading from standard input")

    tmpdir = Path(tempfile.mkdtemp(prefix="md2docx-"))
    try:
        reference = tmpdir / "reference.docx"
        build_reference_docx(base, config, reference)

        lua_text = build_lua_filter(config)
        lua = None
        if lua_text:
            lua = tmpdir / "md2docx.lua"
            lua.write_text(lua_text, encoding="utf-8")

        cmd = build_pandoc_args(args, config, reference, lua, args.inputs, output)

        if args.verbose or args.dry_run:
            print(" ".join(shell_quote(c) for c in cmd), file=sys.stderr)  # noqa: T201 — output
        if args.dry_run:
            args.keep_temp = True
            return 0

        # S603: the command is pandoc plus the options assembled above.
        proc = subprocess.run(cmd, check=False)  # noqa: S603
        if proc.returncode != 0:
            print(f"md2docx: pandoc exited with status {proc.returncode}",  # noqa: T201 — error
                  file=sys.stderr)
            return proc.returncode

        if not args.quiet:
            print(f"md2docx: wrote {output}")  # noqa: T201 — output
        return 0
    finally:
        if args.keep_temp:
            print(f"md2docx: kept generated files in {tmpdir}",  # noqa: T201 — progress
                  file=sys.stderr)
        else:
            shutil.rmtree(tmpdir, ignore_errors=True)


def shell_quote(value: str) -> str:
    """Quote ``value`` for display as part of a shell command."""
    if re.fullmatch(r"[\w@%+=:,./-]+", value):
        return value
    escaped = value.replace("'", r"'\''")
    return f"'{escaped}'"


def main() -> int:
    """Run the command-line interface."""
    try:
        return run(sys.argv[1:])
    except ConfigError as exc:
        print(f"md2docx: {exc}", file=sys.stderr)  # noqa: T201 — error report
        return 2
    except KeyboardInterrupt:
        return 130


if __name__ == "__main__":
    sys.exit(main())
