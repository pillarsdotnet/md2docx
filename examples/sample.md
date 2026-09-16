---
title: md2docx Sample Document
subtitle: Every construct the default configuration styles
author: Robert August Vincent II
date: 2026-09-16
---

# Heading 1

Body text sits in the `Body Text` style. This paragraph exercises *emphasis*,
**strong emphasis**, ***both at once***, ~~strikeout~~, [a hyperlink][pandoc],
`inline code in backticks`, [small caps]{.smallcaps}, H~2~O and E = mc^2^.

Inline code such as `git rebase --interactive` should render in a monospace
face on a light grey `#dddddd` background, matching the fenced blocks below.

## Heading 2

> A block quote picks up the `Block Text` style: indented, greyed, and hung off
> a thin left rule. It can contain *emphasis* and `code` too.

### Heading 3

A fenced code block:

```python
def greet(name: str) -> str:
    """Return a greeting."""
    return f"Hello, {name}!"


if __name__ == "__main__":
    print(greet("world"))
```

An indented code block with no language:

```
$ ./md2docx.py README.md -o README.docx
md2docx: wrote README.docx
```

#### Heading 4

1. Ordered list item with `code`.
2. Second item.
   - Nested bullet.
   - Another nested bullet.

##### Heading 5

Term
:   A definition list entry.

Another term
:   Its definition.

###### Heading 6

| Option | Type | Default |
|--------|------|---------|
| `page_size` | string | `letter` |
| `default_size` | number | `11` |
| `syntax_highlighting` | bool | `true` |

: Table caption text.

A footnote reference lives here.[^note]

[^note]: And the footnote text itself, in the `Footnote Text` style.

[pandoc]: https://pandoc.org/
