"""Utility functions for md2slides."""

from pathlib import Path
from typing import Optional


def get_template_path(name: str) -> Path:
    """Get the path to a template file in the templates/ directory.

    Args:
        name: The template filename (e.g., 'template.pptx', 'multiverse_logo.png')

    Returns:
        Path to the template file

    Raises:
        FileNotFoundError: If the template file does not exist
    """
    # Templates directory is at the project root, parallel to md2slides/
    package_dir = Path(__file__).parent
    project_root = package_dir.parent
    templates_dir = project_root / "templates"

    template_path = templates_dir / name

    if not template_path.exists():
        raise FileNotFoundError(f"Template not found: {name}")

    return template_path


def get_templates_dir() -> Path:
    """Get the path to the templates/ directory.

    Returns:
        Path to the templates directory
    """
    package_dir = Path(__file__).parent
    project_root = package_dir.parent
    return project_root / "templates"


def list_templates() -> list[str]:
    """List all available template files.

    Returns:
        List of template filenames
    """
    templates_dir = get_templates_dir()
    if not templates_dir.exists():
        return []
    return [f.name for f in templates_dir.iterdir() if f.is_file()]
