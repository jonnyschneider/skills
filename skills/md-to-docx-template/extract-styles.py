#!/usr/bin/env python3
"""
Extract styles from an existing DOCX and create a fresh reference template.

Reads font, color, and sizing from a source document's XML and generates
a clean template using those values.

Usage:
    extract-styles.py <source.docx> <output-template.docx>
"""

import zipfile
import tempfile
from pathlib import Path

import click
from lxml import etree


NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
}


def get_text(elem):
    """Get text content or attribute value."""
    if elem is not None:
        return elem.get(f'{{{NAMESPACES["w"]}}}val') or elem.text
    return None


def extract_style_from_xml(style_elem) -> dict:
    """Extract style properties from XML element."""
    info = {}

    # Get style name and type
    name_elem = style_elem.find('w:name', NAMESPACES)
    if name_elem is not None:
        info['name'] = name_elem.get(f'{{{NAMESPACES["w"]}}}val')

    # Get run properties (font, color, etc.)
    rPr = style_elem.find('.//w:rPr', NAMESPACES)
    if rPr is not None:
        # Font name - check multiple locations
        rFonts = rPr.find('w:rFonts', NAMESPACES)
        if rFonts is not None:
            # Try ascii, then hAnsi, then others
            for attr in ['ascii', 'hAnsi', 'cs', 'eastAsia']:
                font_name = rFonts.get(f'{{{NAMESPACES["w"]}}}{attr}')
                if font_name and not font_name.startswith('theme'):
                    info['font'] = font_name
                    break

        # Color
        color = rPr.find('w:color', NAMESPACES)
        if color is not None:
            color_val = color.get(f'{{{NAMESPACES["w"]}}}val')
            if color_val and color_val != 'auto':
                info['color'] = color_val

        # Size
        sz = rPr.find('w:sz', NAMESPACES)
        if sz is not None:
            size_val = sz.get(f'{{{NAMESPACES["w"]}}}val')
            if size_val:
                info['size'] = int(size_val) // 2  # Half-points to points

    # Get paragraph properties
    pPr = style_elem.find('.//w:pPr', NAMESPACES)
    if pPr is not None:
        # Check for font in pPr > rPr as well
        pPr_rPr = pPr.find('w:rPr', NAMESPACES)
        if pPr_rPr is not None and 'font' not in info:
            rFonts = pPr_rPr.find('w:rFonts', NAMESPACES)
            if rFonts is not None:
                for attr in ['ascii', 'hAnsi', 'cs']:
                    font_name = rFonts.get(f'{{{NAMESPACES["w"]}}}{attr}')
                    if font_name and not font_name.startswith('theme'):
                        info['font'] = font_name
                        break

    return info


def extract_fonts_from_theme(theme_path: Path) -> dict:
    """Extract font names from theme XML.

    Theme fonts are defined in a:fontScheme with majorFont and minorFont.
    - majorFont = headings (headingTheme, majorHAnsi, etc.)
    - minorFont = body text (bodyTheme, minorHAnsi, etc.)
    """
    fonts = {'major': None, 'minor': None}

    if not theme_path.exists():
        return fonts

    tree = etree.parse(str(theme_path))
    root = tree.getroot()
    a_ns = 'http://schemas.openxmlformats.org/drawingml/2006/main'

    # Find fontScheme
    font_scheme = root.find(f'.//{{{a_ns}}}fontScheme')
    if font_scheme is None:
        return fonts

    # Major font (headings)
    major = font_scheme.find(f'{{{a_ns}}}majorFont')
    if major is not None:
        latin = major.find(f'{{{a_ns}}}latin')
        if latin is not None:
            fonts['major'] = latin.get('typeface')

    # Minor font (body)
    minor = font_scheme.find(f'{{{a_ns}}}minorFont')
    if minor is not None:
        latin = minor.find(f'{{{a_ns}}}latin')
        if latin is not None:
            fonts['minor'] = latin.get('typeface')

    return fonts


def extract_accent_from_theme(theme_path: Path) -> str:
    """Extract accent color from theme XML."""
    if not theme_path.exists():
        return None

    tree = etree.parse(str(theme_path))
    root = tree.getroot()

    # Look for accent1 color in theme
    # Theme colors are in a:clrScheme/a:accent1
    for accent in root.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}accent1'):
        # Check for srgbClr (direct RGB)
        srgb = accent.find('{http://schemas.openxmlformats.org/drawingml/2006/main}srgbClr')
        if srgb is not None:
            return srgb.get('val')
        # Check for sysClr
        sys_clr = accent.find('{http://schemas.openxmlformats.org/drawingml/2006/main}sysClr')
        if sys_clr is not None:
            return sys_clr.get('lastClr')

    return None


def extract_table_header_color(styles_root) -> str:
    """Extract table header background color from table styles."""
    for style in styles_root.findall('.//w:style[@w:type="table"]', NAMESPACES):
        # Look for firstRow conditional formatting
        first_row = style.find('.//w:tblStylePr[@w:type="firstRow"]', NAMESPACES)
        if first_row is not None:
            shd = first_row.find('.//w:shd', NAMESPACES)
            if shd is not None:
                fill = shd.get(f'{{{NAMESPACES["w"]}}}fill')
                if fill and fill != 'auto':
                    return fill
    return None


def extract_styles(source_path: Path) -> dict:
    """Extract key styles from a source document."""

    extracted = {
        'fonts': {
            'body': None,
            'heading': None,
            'mono': None,
        },
        'colors': {
            'accent': None,
            'heading': None,
            'body': None,
        },
        'styles': {},
    }

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)

        with zipfile.ZipFile(source_path) as zf:
            zf.extractall(tmp)

        styles_path = tmp / 'word' / 'styles.xml'
        theme_path = tmp / 'word' / 'theme' / 'theme1.xml'

        if not styles_path.exists():
            return extracted

        tree = etree.parse(str(styles_path))
        root = tree.getroot()

        # Check docDefaults for default fonts (Google Docs often puts fonts here)
        doc_defaults_font = None
        doc_defaults = root.find('.//w:docDefaults', NAMESPACES)
        if doc_defaults is not None:
            rPr_default = doc_defaults.find('.//w:rPrDefault/w:rPr/w:rFonts', NAMESPACES)
            if rPr_default is not None:
                for attr in ['ascii', 'hAnsi']:
                    font_name = rPr_default.get(f'{{{NAMESPACES["w"]}}}{attr}')
                    if font_name and not font_name.startswith('theme'):
                        # Clean up font name (Google Docs sometimes appends size like "DM Sans 14pt")
                        if font_name.endswith('pt'):
                            # Remove the trailing size (e.g., "DM Sans 14pt" -> "DM Sans")
                            font_name = ' '.join(font_name.split()[:-1])
                        doc_defaults_font = font_name
                        extracted['fonts']['body'] = font_name
                        break

        # Map style IDs to our categories
        # Check multiple heading levels - some may have explicit fonts, others inherit
        style_mapping = {
            'Normal': 'body',
            'Title': 'heading',  # Title often has the heading font
            'Heading1': 'heading',
            'Heading2': 'heading',
            'Heading3': 'heading',
            'Heading4': 'heading',
            'Heading 1': 'heading',
            'Heading 2': 'heading',
            'Heading 3': 'heading',
            'Heading 4': 'heading',
            'SourceCode': 'mono',
            'VerbatimChar': 'mono',
        }

        # Extract all styles
        for style_elem in root.findall('.//w:style', NAMESPACES):
            style_id = style_elem.get(f'{{{NAMESPACES["w"]}}}styleId')
            if not style_id:
                continue

            info = extract_style_from_xml(style_elem)
            if info:
                extracted['styles'][style_id] = info

            # Map to our categories
            if style_id in style_mapping:
                category = style_mapping[style_id]

                if info.get('font'):
                    if category == 'body' and not extracted['fonts']['body']:
                        extracted['fonts']['body'] = info['font']
                    elif category == 'heading':
                        # For heading, prefer Title font, then any heading with a font
                        if style_id == 'Title' or not extracted['fonts']['heading']:
                            extracted['fonts']['heading'] = info['font']
                    elif category == 'mono' and not extracted['fonts']['mono']:
                        extracted['fonts']['mono'] = info['font']

                if info.get('color'):
                    if category == 'heading':
                        # For heading color, prefer Heading1/Heading2
                        if style_id in ('Heading1', 'Heading 1', 'Heading2', 'Heading 2') or not extracted['colors']['heading']:
                            extracted['colors']['heading'] = info['color']
                    elif category == 'body' and not extracted['colors']['body']:
                        extracted['colors']['body'] = info['color']

        # Prefer heading color as accent (most likely to be intentional brand color)
        if extracted['colors']['heading']:
            if extracted['colors']['heading'] not in ('000000', 'auto'):
                extracted['colors']['accent'] = extracted['colors']['heading']

        # Or from table header background
        if not extracted['colors']['accent']:
            table_accent = extract_table_header_color(root)
            if table_accent:
                extracted['colors']['accent'] = table_accent

        # Or from theme accent
        if not extracted['colors']['accent']:
            accent = extract_accent_from_theme(theme_path)
            if accent:
                extracted['colors']['accent'] = accent

        # Use docDefaults font for headings if no explicit heading font found
        # (headings inherit from docDefaults unless they override)
        if not extracted['fonts']['heading'] and doc_defaults_font:
            extracted['fonts']['heading'] = doc_defaults_font

        # Try to get fonts from theme if not found in explicit styles or docDefaults
        theme_fonts = extract_fonts_from_theme(theme_path)
        if not extracted['fonts']['body'] and theme_fonts.get('minor'):
            extracted['fonts']['body'] = theme_fonts['minor']
        if not extracted['fonts']['heading'] and theme_fonts.get('major'):
            extracted['fonts']['heading'] = theme_fonts['major']

    # Defaults
    if not extracted['fonts']['body']:
        extracted['fonts']['body'] = 'Calibri'
    if not extracted['fonts']['heading']:
        extracted['fonts']['heading'] = extracted['fonts']['body']
    if not extracted['fonts']['mono']:
        extracted['fonts']['mono'] = 'Consolas'
    if not extracted['colors']['accent']:
        extracted['colors']['accent'] = '4472C4'

    return extracted


def create_template_from_extracted(extracted: dict, output_path: Path):
    """Create a fresh template using extracted style values."""
    import importlib.util
    spec = importlib.util.spec_from_file_location(
        "create_reference_template",
        Path(__file__).parent / "create-reference-template.py"
    )
    crt = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(crt)

    result = crt.create_reference_template(
        str(output_path),
        font_body=extracted['fonts']['body'],
        font_heading=extracted['fonts']['heading'],
        font_mono=extracted['fonts']['mono'],
        accent_color=extracted['colors']['accent'],
        heading_color=extracted['colors'].get('heading'),
    )

    return result


@click.command()
@click.argument('source', type=click.Path(exists=True))
@click.argument('output', type=click.Path())
@click.option('--show-styles', is_flag=True, help='Show all extracted styles')
def main(source, output, show_styles):
    """Extract styles from SOURCE and create a fresh template at OUTPUT."""
    source_path = Path(source)
    output_path = Path(output)

    click.echo(f"Extracting styles from: {source_path.name}")

    extracted = extract_styles(source_path)

    click.echo(f"\nDetected:")
    click.echo(f"  Body font:    {extracted['fonts']['body']}")
    click.echo(f"  Heading font: {extracted['fonts']['heading']}")
    click.echo(f"  Mono font:    {extracted['fonts']['mono']}")
    click.echo(f"  Accent color: #{extracted['colors']['accent']}")
    if extracted['colors']['heading']:
        click.echo(f"  Heading color: #{extracted['colors']['heading']}")

    if show_styles:
        click.echo(f"\nAll styles found:")
        for name, info in sorted(extracted['styles'].items()):
            click.echo(f"  {name}: {info}")

    click.echo(f"\nCreating template...")
    result = create_template_from_extracted(extracted, output_path)

    click.echo(f"\nTemplate created: {output_path}")
    click.echo(f"Styles configured: {len(result['styles'])}")


if __name__ == '__main__':
    main()
