"""CLI entry point for md2slides."""

import argparse
import sys

from md2slides import __version__

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

    print(f"Converting {args.input_file} to PowerPoint...")
    print(f"Theme: {args.theme}")
    if args.output:
        print(f"Output: {args.output}")
    else:
        print("Output: (default based on input file)")

    # TODO: Implement actual conversion
    print("Conversion not yet implemented.")


if __name__ == "__main__":
    main()
