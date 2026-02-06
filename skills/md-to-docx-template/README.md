# md-to-docx

Convert markdown to **beautifully styled** Word documents. Not plain, unstyled exports — professional documents with your fonts, colors, and table styling applied automatically.

## Why

Write in markdown, publish as branded DOCX. One command, no manual formatting. Edit your markdown, rebuild, get consistent output every time.

The goal: **one-shot export from markdown to templated DOCX without intervention.**

## Prerequisites

- `pandoc` (install: `brew install pandoc`)
- Python packages: `python-docx`, `lxml`, `click`, `pyyaml` (in automations venv)

## Quick Start

```bash
# Convert markdown with default styling
./md-to-docx.py input.md -o output.docx

# Use a custom template
./md-to-docx.py input.md -o output.docx --template my-brand.docx

# Assemble multiple files from a manifest
./md-to-docx.py manifest.yaml --open
```

## Template Customization

Two approaches for custom branding:

### Option 1: Generate a Fresh Template

Create a template with your fonts and colors:

```bash
./create-reference-template.py my-template.docx \
  --font-body "DM Sans" \
  --font-heading "DM Sans" \
  --font-mono "Consolas" \
  --accent-color "0D494D" \
  --heading-color "0D494D"
```

Options:
- `--font-body` — Body text font (default: Calibri)
- `--font-heading` — Heading font (default: Calibri)
- `--font-mono` — Code/monospace font (default: Consolas)
- `--accent-color` — Hex color for table headers/accents (default: 4472C4)
- `--heading-color` — Hex color for headings (default: black)

### Option 2: Extract Styles from an Existing Document

Have a branded Word document? Extract its styling:

```bash
# Extract fonts, colors, and styling from existing document
./extract-styles.py existing-branded-doc.docx my-template.docx

# Then use it for conversion
./md-to-docx.py input.md -o output.docx --template my-template.docx
```

The extractor detects:
- Body and heading fonts (from styles, docDefaults, or theme)
- Heading colors
- Accent/highlight colors
- Monospace fonts

### Option 3: Edit the Vanilla Template in Word

1. Run a conversion to generate `templates/vanilla.docx`
2. Open in Word and modify styles (Format → Style → Modify)
3. Save and use as your template

**Important:** When editing in Word, modify the **styles** (Heading 1, Normal, etc.) rather than direct formatting. Use Design → Fonts to set theme fonts.

## Full Pipeline for Extracted Styles

When extracting from an existing document, run the full pipeline for best results:

```bash
# 1. Extract styles from source document
./extract-styles.py source.docx template.docx

# 2. Convert markdown using template
./md-to-docx.py input.md -o output.docx --template template.docx --no-fix-indent

# 3. Apply template styles (ensures colors transfer)
./apply-template-styles.py template.docx output.docx

# 4. Fix table header text color
./fix-table-headers.py output.docx
```

## Manifest Mode

For multi-file documents, create a `manifest.yaml`:

```yaml
template: ../templates/brand.docx
output: ../My Document.docx
page_break_between_sections: true
toc: false
sections:
  - 00-front-matter.md
  - 01-introduction.md
  - 02-main-content.md
```

| Key | Required | Description |
|-----|----------|-------------|
| `output` | Yes | Path for the built DOCX |
| `sections` | Yes | Ordered list of markdown files |
| `template` | No | Reference DOCX (default: vanilla template) |
| `page_break_between_sections` | No | Insert breaks between files (default: true) |
| `toc` | No | Generate table of contents (default: false) |

## Supported Elements

| Markdown | Result |
|----------|--------|
| `# Heading` | Styled headings (H1-H6) with custom fonts/colors |
| Paragraphs | Body text with proper font and spacing |
| `**bold**`, `*italic*` | Strong/emphasis formatting |
| `` `code` `` | Inline code with monospace font |
| Code blocks | Shaded code blocks with monospace font |
| `> blockquote` | Indented block with left border |
| Lists | Bullet and numbered lists with compact indentation |
| Tables | Styled tables with colored headers |
| `![](image.png)` | Embedded images (local or remote) |
| `[link](url)` | Hyperlinks |

## Limitations

This tool produces good results for most documents, but don't expect perfection:

- **Font availability**: Custom fonts must be installed on machines opening the documents
- **Complex tables**: Nested tables or complex layouts may not render perfectly
- **SVG images**: Require `rsvg-convert` to be installed
- **Theme conflicts**: Some Word features rely on theme definitions that may conflict with explicit styling
- **Pandoc quirks**: Pandoc generates its own formatting that sometimes overrides template styles — the post-processing scripts fix the most common issues

For best results:
- Use common fonts (or embed fonts in the document)
- Keep table structures simple
- Test output on target machines

## Command Reference

```bash
# Simple conversion
./md-to-docx.py input.md -o output.docx [options]

# Manifest mode
./md-to-docx.py manifest.yaml [options]

Options:
  -o, --output PATH     Output DOCX (required for .md input)
  --template PATH       Custom reference template
  --toc                 Generate table of contents
  --open                Open output after building
  --dry-run             Preview without building
  --no-fix-indent       Skip list indentation fix (use with extract pipeline)
```

## Files

| File | Purpose |
|------|---------|
| `md-to-docx.py` | Main conversion script |
| `create-reference-template.py` | Generate templates with custom fonts/colors |
| `extract-styles.py` | Extract styling from existing Word documents |
| `apply-template-styles.py` | Copy styles from template to output |
| `fix-table-headers.py` | Apply white text to table headers |
| `fix-list-indent.py` | Compact list indentation |
| `fix-fonts.py` | Remove direct font formatting |
| `templates/vanilla.docx` | Auto-generated default template |
