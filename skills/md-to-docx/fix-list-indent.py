#!/usr/bin/env python3
"""
Fix list indentation in pandoc-generated DOCX files.

Pandoc generates its own numbering definitions with 720 twips (0.5") indent.
This script modifies them to use compact indentation suitable for tables.

Usage:
    fix-list-indent.py <input.docx> [output.docx]

If output is not specified, modifies the input file in place.
"""

import sys
import zipfile
import tempfile
import re
from pathlib import Path
from xml.etree import ElementTree as ET


def fix_numbering_indent(doc_path: Path, output_path: Path = None,
                         base_indent: int = 360, hanging: int = 216):
    """
    Fix list indentation in a DOCX file.

    Args:
        doc_path: Input DOCX file
        output_path: Output DOCX file (or None to modify in place)
        base_indent: Base indent in twips (216 = 0.15")
        hanging: Hanging indent in twips
    """
    if output_path is None:
        output_path = doc_path

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)

        # Extract
        with zipfile.ZipFile(doc_path) as zf:
            zf.extractall(tmp)

        numbering_path = tmp / 'word' / 'numbering.xml'
        if not numbering_path.exists():
            print("No numbering.xml found - no lists to fix")
            return

        # Parse and fix
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        ET.register_namespace('w', ns['w'])

        tree = ET.parse(numbering_path)
        root = tree.getroot()

        changes = 0
        w_ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

        # Find all abstractNum elements and fix their level indents
        for abstract_num in root.findall(f'.//{w_ns}abstractNum', ns):
            for lvl in abstract_num.findall(f'{w_ns}lvl', ns):
                ilvl = lvl.get(f'{w_ns}ilvl', '0')
                level = int(ilvl)

                pPr = lvl.find(f'{w_ns}pPr', ns)
                if pPr is not None:
                    ind = pPr.find(f'{w_ns}ind', ns)
                    if ind is not None:
                        left = ind.get(f'{w_ns}left')
                        if left and int(left) >= 500:  # Only fix large indents
                            new_left = base_indent + (level * base_indent)
                            ind.set(f'{w_ns}left', str(new_left))
                            ind.set(f'{w_ns}hanging', str(hanging))
                            changes += 1

        if changes > 0:
            tree.write(numbering_path, xml_declaration=True, encoding='UTF-8')
            print(f"Fixed {changes} indent definitions")
        else:
            print("No large indents found to fix")

        # Repack
        if output_path != doc_path:
            output_path.parent.mkdir(parents=True, exist_ok=True)
        if output_path.exists() and output_path != doc_path:
            output_path.unlink()

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for f in tmp.rglob('*'):
                if f.is_file():
                    zf.write(f, f.relative_to(tmp))


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    input_path = Path(sys.argv[1])
    output_path = Path(sys.argv[2]) if len(sys.argv) > 2 else None

    if not input_path.exists():
        print(f"Error: {input_path} not found")
        sys.exit(1)

    fix_numbering_indent(input_path, output_path)
    print(f"Output: {output_path or input_path}")


if __name__ == '__main__':
    main()
