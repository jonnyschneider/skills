#!/usr/bin/env python3
"""
Apply template styles to a pandoc-generated document.

Pandoc doesn't fully copy style properties (especially colors) from reference docs.
This script copies the styles.xml from the template to the output document,
preserving all formatting including colors.

Usage:
    apply-template-styles.py <template.docx> <output.docx>
"""

import sys
import zipfile
import tempfile
from pathlib import Path

import click


def copy_styles_from_template(template_path: Path, output_path: Path):
    """Copy styles.xml from template to output document (wholesale copy)."""

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)

        # Extract output document
        output_dir = tmp / 'output'
        output_dir.mkdir()
        with zipfile.ZipFile(output_path) as zf:
            zf.extractall(output_dir)

        # Extract template
        template_dir = tmp / 'template'
        template_dir.mkdir()
        with zipfile.ZipFile(template_path) as zf:
            zf.extractall(template_dir)

        # Wholesale copy styles.xml from template to output
        template_styles = template_dir / 'word' / 'styles.xml'
        output_styles = output_dir / 'word' / 'styles.xml'

        if template_styles.exists():
            # Just copy the file directly - no XML parsing
            output_styles.write_bytes(template_styles.read_bytes())
            click.echo(f"  Styles copied from template")

        # Also copy theme if present (themes contain color schemes)
        template_theme = template_dir / 'word' / 'theme' / 'theme1.xml'
        output_theme_dir = output_dir / 'word' / 'theme'
        output_theme = output_theme_dir / 'theme1.xml'

        if template_theme.exists() and output_theme.exists():
            # Only copy if output already has a theme (to avoid adding relationships)
            output_theme.write_bytes(template_theme.read_bytes())
            click.echo(f"  Theme copied")

        # Repack
        output_path.unlink()
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for f in output_dir.rglob('*'):
                if f.is_file():
                    zf.write(f, f.relative_to(output_dir))


@click.command()
@click.argument('template', type=click.Path(exists=True))
@click.argument('output', type=click.Path(exists=True))
def main(template, output):
    """Copy styles from TEMPLATE into OUTPUT document."""
    template_path = Path(template)
    output_path = Path(output)

    click.echo(f"Applying styles from: {template_path.name}")
    click.echo(f"To document: {output_path.name}")

    copy_styles_from_template(template_path, output_path)

    click.echo(f"\nDone: {output_path}")


if __name__ == '__main__':
    main()
