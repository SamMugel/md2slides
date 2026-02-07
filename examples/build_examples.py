#!/usr/bin/env python3
"""
Build script to generate all example PPTX files from markdown sources.

This script converts all .md files in the examples/ directory to .pptx files.
It uses only local resources (images in examples/ or templates/) and does not
require network access.

Usage:
    python examples/build_examples.py

Output:
    PPTX files are generated in the examples/ directory alongside their
    source markdown files.
"""

import os
import sys
from pathlib import Path

# Add parent directory to path for imports
script_dir = Path(__file__).parent
project_root = script_dir.parent
sys.path.insert(0, str(project_root))

from md2slides.converter import convert_markdown_to_pptx


def build_all_examples():
    """Convert all markdown files in examples/ to PPTX."""
    examples_dir = script_dir
    md_files = sorted(examples_dir.glob("*.md"))

    if not md_files:
        print("No markdown files found in examples/")
        return 1

    print(f"Found {len(md_files)} markdown files to convert\n")

    success_count = 0
    error_count = 0

    for md_file in md_files:
        pptx_file = md_file.with_suffix(".pptx")
        print(f"Converting: {md_file.name} -> {pptx_file.name}")

        try:
            # Read markdown content from file
            markdown_content = md_file.read_text(encoding="utf-8")

            # Change to examples directory so relative image paths resolve
            original_cwd = os.getcwd()
            os.chdir(examples_dir)

            try:
                convert_markdown_to_pptx(markdown_content, str(pptx_file))
            finally:
                os.chdir(original_cwd)

            size_kb = pptx_file.stat().st_size / 1024
            print(f"  ✓ Success ({size_kb:.1f} KB)")
            success_count += 1
        except Exception as e:
            print(f"  ✗ Error: {e}")
            error_count += 1

    print(f"\n{'='*50}")
    print(f"Results: {success_count} succeeded, {error_count} failed")

    if error_count > 0:
        return 1
    return 0


def main():
    """Main entry point."""
    print("md2slides Example Builder")
    print("=" * 50)
    print()

    return build_all_examples()


if __name__ == "__main__":
    sys.exit(main())
