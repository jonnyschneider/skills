#!/usr/bin/env python3
"""
Fix table header text color in DOCX files.

Applies white text color to the first row of all tables.

Usage:
    fix-table-headers.py <document.docx> [--color FFFFFF]
"""

import zipfile
import tempfile
from pathlib import Path

import click
from lxml import etree


# OOXML namespaces
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}


def fix_table_header_colors(doc_path: Path, text_color: str = 'FFFFFF'):
    """Apply text color to first row of all tables."""

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)

        # Extract
        with zipfile.ZipFile(doc_path) as zf:
            zf.extractall(tmp)

        # Parse document.xml
        doc_xml = tmp / 'word' / 'document.xml'

        # Use lxml for better namespace handling
        parser = etree.XMLParser(remove_blank_text=False)
        tree = etree.parse(str(doc_xml), parser)
        root = tree.getroot()

        tables_fixed = 0

        # Find all tables
        for tbl in root.findall('.//w:tbl', NAMESPACES):
            # Get first row
            first_row = tbl.find('w:tr', NAMESPACES)
            if first_row is None:
                continue

            # Process each cell in first row
            for tc in first_row.findall('w:tc', NAMESPACES):
                # Process each paragraph in the cell
                for p in tc.findall('w:p', NAMESPACES):
                    # Process each run in the paragraph
                    for r in p.findall('w:r', NAMESPACES):
                        # Get or create run properties
                        rPr = r.find('w:rPr', NAMESPACES)
                        if rPr is None:
                            rPr = etree.Element(f'{{{NAMESPACES["w"]}}}rPr')
                            r.insert(0, rPr)

                        # Remove existing color if any
                        existing_color = rPr.find('w:color', NAMESPACES)
                        if existing_color is not None:
                            rPr.remove(existing_color)

                        # Add white color
                        color = etree.Element(f'{{{NAMESPACES["w"]}}}color')
                        color.set(f'{{{NAMESPACES["w"]}}}val', text_color)
                        rPr.append(color)

                        # Also make bold if not already
                        if rPr.find('w:b', NAMESPACES) is None:
                            bold = etree.Element(f'{{{NAMESPACES["w"]}}}b')
                            rPr.append(bold)

            tables_fixed += 1

        # Write back
        tree.write(str(doc_xml), xml_declaration=True, encoding='UTF-8', standalone=True)

        # Repack
        doc_path.unlink()
        with zipfile.ZipFile(doc_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for f in tmp.rglob('*'):
                if f.is_file():
                    zf.write(f, f.relative_to(tmp))

        return tables_fixed


@click.command()
@click.argument('document', type=click.Path(exists=True))
@click.option('--color', default='FFFFFF', help='Text color hex (default: FFFFFF white)')
def main(document, color):
    """Apply light text color to table header rows."""
    doc_path = Path(document)

    click.echo(f"Fixing table headers in: {doc_path.name}")

    count = fix_table_header_colors(doc_path, color)

    click.echo(f"Fixed {count} table(s)")


if __name__ == '__main__':
    main()
