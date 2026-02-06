#!/usr/bin/env python3
"""
Remove direct font formatting from document so style fonts apply.

Pandoc often applies direct formatting that overrides style definitions.
This script removes direct font (rFonts) elements from paragraphs so the
style fonts take effect.

Usage:
    fix-fonts.py <document.docx>
"""

import zipfile
import tempfile
from pathlib import Path

import click
from lxml import etree


NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
}


def fix_fonts(doc_path: Path):
    """Remove direct font formatting so styles apply."""

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)

        # Extract
        with zipfile.ZipFile(doc_path) as zf:
            zf.extractall(tmp)

        # Parse document.xml
        doc_xml = tmp / 'word' / 'document.xml'

        parser = etree.XMLParser(remove_blank_text=False)
        tree = etree.parse(str(doc_xml), parser)
        root = tree.getroot()

        fonts_removed = 0

        # Find all rFonts elements and remove them (let styles control fonts)
        for rFonts in root.findall('.//w:rFonts', NAMESPACES):
            parent = rFonts.getparent()
            if parent is not None:
                parent.remove(rFonts)
                fonts_removed += 1

        # Write back
        tree.write(str(doc_xml), xml_declaration=True, encoding='UTF-8', standalone=True)

        # Repack
        doc_path.unlink()
        with zipfile.ZipFile(doc_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for f in tmp.rglob('*'):
                if f.is_file():
                    zf.write(f, f.relative_to(tmp))

        return fonts_removed


@click.command()
@click.argument('document', type=click.Path(exists=True))
def main(document):
    """Remove direct font formatting so style fonts apply."""
    doc_path = Path(document)

    click.echo(f"Fixing fonts in: {doc_path.name}")

    count = fix_fonts(doc_path)

    click.echo(f"Removed {count} direct font references")


if __name__ == '__main__':
    main()
