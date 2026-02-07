"""CLI entry point for md2slides."""

import argparse
import sys
from pathlib import Path

from md2slides import __version__
from md2slides.converter import PresentationConverter

THEMES = ["default", "dark", "light", "corporate"]


def main():
    """Convert Markdown files to PowerPoint presentations."""
    parser = argparse.ArgumentParser(
        prog="md2slides",
        description="Convert Markdown files to PowerPoint presentations.",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"md2slides {__version__}",
    )
    parser.add_argument(
        "input_file",
        nargs="?",
        help="Path to the markdown file to convert",
    )
    parser.add_argument(
        "-o",
        "--output",
        help="Output PPTX file path",
    )
    parser.add_argument(
        "--theme",
        choices=THEMES,
        default="default",
        help="Theme to use for the presentation (default: default)",
    )
    parser.add_argument(
        "--list-themes",
        action="store_true",
        help="List available themes",
    )

    args = parser.parse_args()

    if args.list_themes:
        print("Available themes:")
        for theme in THEMES:
            print(f"  - {theme}")
        return

    if not args.input_file:
        print("Error: Missing argument 'INPUT_FILE'.", file=sys.stderr)
        print("Run 'md2slides --help' for usage information.", file=sys.stderr)
        sys.exit(1)

    # Validate input file exists
    input_path = Path(args.input_file)
    if not input_path.exists():
        print(f"Error: File not found: {args.input_file}", file=sys.stderr)
        sys.exit(1)

    if not input_path.is_file():
        print(f"Error: Not a file: {args.input_file}", file=sys.stderr)
        sys.exit(1)

    # Determine output path
    if args.output:
        output_path = Path(args.output)
    else:
        output_path = input_path.with_suffix(".pptx")

    print(f"Converting {input_path} to PowerPoint...")

    try:
        converter = PresentationConverter(theme=args.theme)
        converter.convert_file(str(input_path), str(output_path))
        print(f"Created: {output_path}")
    except Exception as e:
        print(f"Error: Conversion failed: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
