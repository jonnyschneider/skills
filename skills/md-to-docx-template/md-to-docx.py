#!/usr/bin/env python3
"""
Convert markdown to styled DOCX using pandoc.

Two modes:
  1. Simple: Convert a single markdown file
     md-to-docx.py input.md -o output.docx

  2. Manifest: Assemble multiple files from a manifest.yaml
     md-to-docx.py manifest.yaml

Features:
  - Auto-generates a vanilla template if none specified
  - Post-processes list indentation for compact tables
  - Supports custom templates with --template

Usage:
    md-to-docx.py <input.md> -o <output.docx> [--template template.docx]
    md-to-docx.py <manifest.yaml> [--open]
"""

import os
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import click
import yaml

# Default template location
DEFAULT_TEMPLATE_DIR = Path(__file__).parent / 'templates'
DEFAULT_TEMPLATE = DEFAULT_TEMPLATE_DIR / 'vanilla.docx'


def resolve_path(path, base_dir, allow_escape=False):
    """Resolve a path relative to a base directory.

    By default, rejects paths that resolve outside base_dir (path traversal guard).
    Set allow_escape=True for output paths that legitimately live elsewhere.
    """
    resolved = os.path.normpath(os.path.join(base_dir, path)) if not os.path.isabs(path) else os.path.normpath(path)
    if not allow_escape and not resolved.startswith(os.path.normpath(base_dir) + os.sep) and resolved != os.path.normpath(base_dir):
        raise click.ClickException(f"Path escapes base directory: {path}")
    return resolved


def ensure_default_template():
    """Generate the default vanilla template if it doesn't exist."""
    if DEFAULT_TEMPLATE.exists():
        return DEFAULT_TEMPLATE

    DEFAULT_TEMPLATE_DIR.mkdir(parents=True, exist_ok=True)

    # Import from same directory
    import importlib.util
    spec = importlib.util.spec_from_file_location(
        "create_reference_template",
        Path(__file__).parent / "create-reference-template.py"
    )
    crt = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(crt)

    click.echo("Generating default template...")
    crt.create_reference_template(
        str(DEFAULT_TEMPLATE),
        font_body='Calibri',
        font_heading='Calibri',
        font_mono='Consolas',
        accent_color='666666',  # Neutral gray
    )
    click.echo(f"Created: {DEFAULT_TEMPLATE}")
    return DEFAULT_TEMPLATE


def fix_list_indentation(doc_path: Path, base_indent: int = 360, hanging: int = 216):
    """Fix list indentation in the generated document for compact tables."""
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)

        with zipfile.ZipFile(doc_path) as zf:
            zf.extractall(tmp)

        numbering_path = tmp / 'word' / 'numbering.xml'
        if not numbering_path.exists():
            return  # No lists

        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        ET.register_namespace('w', ns['w'])

        tree = ET.parse(numbering_path)
        root = tree.getroot()
        w_ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

        changes = 0
        for abstract_num in root.findall(f'.//{w_ns}abstractNum', ns):
            for lvl in abstract_num.findall(f'{w_ns}lvl', ns):
                ilvl = lvl.get(f'{w_ns}ilvl', '0')
                level = int(ilvl)

                pPr = lvl.find(f'{w_ns}pPr', ns)
                if pPr is not None:
                    ind = pPr.find(f'{w_ns}ind', ns)
                    if ind is not None:
                        left = ind.get(f'{w_ns}left')
                        if left and int(left) >= 500:
                            new_left = base_indent + (level * base_indent)
                            ind.set(f'{w_ns}left', str(new_left))
                            ind.set(f'{w_ns}hanging', str(hanging))
                            changes += 1

        if changes > 0:
            tree.write(numbering_path, xml_declaration=True, encoding='UTF-8')

        # Repack
        doc_path.unlink()
        with zipfile.ZipFile(doc_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for f in tmp.rglob('*'):
                if f.is_file():
                    zf.write(f, f.relative_to(tmp))


def read_manifest(manifest_path):
    """Read and validate the manifest file."""
    with open(manifest_path, 'r') as f:
        manifest = yaml.safe_load(f)

    # Template is optional (uses default if not specified)
    required = ['output', 'sections']
    for key in required:
        if key not in manifest:
            raise click.ClickException(f"Missing required key '{key}' in manifest")

    if not manifest['sections']:
        raise click.ClickException("No sections listed in manifest")

    return manifest


def assemble_markdown(sections, base_dir, page_breaks=True):
    """Concatenate markdown files, optionally inserting page breaks."""
    parts = []

    for i, filename in enumerate(sections):
        filepath = resolve_path(filename, base_dir)

        if not os.path.exists(filepath):
            raise click.ClickException(f"Section file not found: {filepath}")

        with open(filepath, 'r') as f:
            content = f.read().rstrip('\n')

        parts.append(content)

        if page_breaks and i < len(sections) - 1:
            parts.append('\n\n\\newpage\n')

    return '\n\n'.join(parts) + '\n'


def run_pandoc(input_path, output_path, template_path, toc=False, highlight=False, lua_filter=None):
    """Invoke pandoc to convert markdown to DOCX."""
    cmd = [
        'pandoc',
        input_path,
        f'--reference-doc={template_path}',
        '-o', output_path,
        '--from=markdown',
        '--request-header=User-Agent: pandoc/3.0',  # For remote images
    ]

    if toc:
        cmd.append('--toc')

    if not highlight:
        cmd.append('--syntax-highlighting=none')

    if lua_filter:
        cmd.extend(['--lua-filter', lua_filter])

    result = subprocess.run(cmd, capture_output=True, text=True)

    if result.returncode != 0:
        raise click.ClickException(f"pandoc failed:\n{result.stderr}")

    # Print any warnings
    if result.stderr:
        for line in result.stderr.strip().split('\n'):
            if line.strip():
                click.echo(f"  Warning: {line}", err=True)

    return result


@click.command()
@click.argument('input_file', type=click.Path(exists=True))
@click.option('-o', '--output', type=click.Path(), help='Output DOCX path (required for .md input)')
@click.option('--template', type=click.Path(exists=True), help='Custom reference template')
@click.option('--toc', is_flag=True, help='Generate table of contents')
@click.option('--open', 'open_file', is_flag=True, help='Open the output file after building')
@click.option('--dry-run', is_flag=True, help='Show what would be done without building')
@click.option('--no-fix-indent', is_flag=True, help='Skip list indentation fix')
def main(input_file, output, template, toc, open_file, dry_run, no_fix_indent):
    """Convert markdown to styled DOCX.

    INPUT_FILE can be a .md file or a manifest.yaml file.

    \b
    Examples:
      md-to-docx input.md -o output.docx
      md-to-docx input.md -o output.docx --template custom.docx
      md-to-docx manifest.yaml --open
    """
    input_path = Path(input_file).resolve()

    # Determine mode based on file extension
    if input_path.suffix in ('.yaml', '.yml'):
        # Manifest mode
        _run_manifest_mode(input_path, template, open_file, dry_run, no_fix_indent)
    elif input_path.suffix == '.md':
        # Simple mode
        if not output:
            raise click.ClickException("Output path required (-o output.docx) for markdown files")
        _run_simple_mode(input_path, Path(output), template, toc, open_file, dry_run, no_fix_indent)
    else:
        raise click.ClickException(f"Unsupported input file type: {input_path.suffix}")


def _run_simple_mode(input_path, output_path, template, toc, open_file, dry_run, no_fix_indent):
    """Convert a single markdown file to DOCX."""
    # Determine template
    if template:
        template_path = Path(template)
    else:
        template_path = ensure_default_template()

    if dry_run:
        click.echo(f"Input:     {input_path}")
        click.echo(f"Output:    {output_path}")
        click.echo(f"Template:  {template_path}")
        click.echo(f"TOC:       {toc}")
        return

    click.echo(f"Converting {input_path.name}...")
    run_pandoc(str(input_path), str(output_path), str(template_path), toc)

    # Post-process list indentation
    if not no_fix_indent:
        fix_list_indentation(output_path)

    size = output_path.stat().st_size
    click.echo(f"Built: {output_path} ({size:,} bytes)")

    if open_file:
        subprocess.run(['open', str(output_path)])


def _run_manifest_mode(manifest_path, template_override, open_file, dry_run, no_fix_indent):
    """Assemble multiple markdown files from a manifest."""
    manifest_dir = manifest_path.parent
    config = read_manifest(str(manifest_path))

    # Template: CLI override > manifest > default
    if template_override:
        template_path = Path(template_override)
    elif config.get('template'):
        template_path = Path(resolve_path(config['template'], str(manifest_dir)))
    else:
        template_path = ensure_default_template()

    output_path = Path(resolve_path(config['output'], str(manifest_dir), allow_escape=True))
    page_breaks = config.get('page_break_between_sections', True)
    toc = config.get('toc', False)

    sections = config['sections']

    if dry_run:
        click.echo(f"Manifest:  {manifest_path}")
        click.echo(f"Template:  {template_path}")
        click.echo(f"Output:    {output_path}")
        click.echo(f"TOC:       {toc}")
        click.echo(f"Breaks:    {page_breaks}")
        click.echo(f"Sections:  {len(sections)}")
        for s in sections:
            path = resolve_path(s, str(manifest_dir))
            exists = "✓" if os.path.exists(path) else "✗"
            click.echo(f"  {exists} {s}")
        return

    # Validate template exists
    if not template_path.exists():
        raise click.ClickException(f"Template not found: {template_path}")

    # Assemble
    click.echo(f"Assembling {len(sections)} sections...")
    combined = assemble_markdown(sections, str(manifest_dir), page_breaks)

    # Write to temp file and run pandoc
    with tempfile.NamedTemporaryFile(mode='w', suffix='.md', delete=False) as tmp:
        tmp.write(combined)
        tmp_path = tmp.name

    try:
        click.echo(f"Building DOCX...")
        run_pandoc(tmp_path, str(output_path), str(template_path), toc)
    finally:
        os.unlink(tmp_path)

    # Post-process list indentation
    if not no_fix_indent:
        fix_list_indentation(output_path)

    size = output_path.stat().st_size
    click.echo(f"Built: {output_path} ({size:,} bytes)")

    if open_file:
        subprocess.run(['open', str(output_path)])


if __name__ == '__main__':
    main()
