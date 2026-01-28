"""Command-line interface for md2slides."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import List, Optional

from md2slides.converter import convert_file
from md2slides.parser import ValidationError


def main(args: Optional[List[str]] = None) -> int:
    """Main entry point for the CLI.

    Args:
        args: Command line arguments. If None, uses sys.argv.

    Returns:
        Exit code (0 for success, 1 for error).
    """
    parser = argparse.ArgumentParser(
        prog="md2slides",
        description="Convert Markdown files to PowerPoint (PPTX) presentations",
    )
    parser.add_argument(
        "input",
        type=str,
        help="Path to the input Markdown file",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=str,
        default=None,
        help="Path for the output PPTX file (default: same as input with .pptx extension)",
    )
    parser.add_argument(
        "-v",
        "--version",
        action="version",
        version="%(prog)s 0.1.0",
    )

    parsed_args = parser.parse_args(args)

    try:
        output_path = convert_file(parsed_args.input, parsed_args.output)
        print(f"Created: {output_path}")
        return 0
    except FileNotFoundError as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1
    except ValidationError as e:
        print(f"Validation error: {e}", file=sys.stderr)
        return 1
    except Exception as e:
        print(f"Unexpected error: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
