"""Base theme class for md2slides."""

from dataclasses import dataclass
from typing import Tuple


@dataclass
class ThemeColors:
    """Color definitions for a theme."""
    # RGB tuples (0-255)
    background: Tuple[int, int, int] = (255, 255, 255)
    title_text: Tuple[int, int, int] = (0, 0, 0)
    body_text: Tuple[int, int, int] = (51, 51, 51)
    accent: Tuple[int, int, int] = (0, 112, 192)
    code_background: Tuple[int, int, int] = (245, 245, 245)


@dataclass
class ThemeFonts:
    """Font definitions for a theme."""
    title_font: str = "Calibri"
    body_font: str = "Calibri"
    code_font: str = "Courier New"
    title_size: int = 44  # points
    subtitle_size: int = 32  # points
    body_size: int = 18  # points
    code_size: int = 12  # points


class BaseTheme:
    """Base class for presentation themes."""

    name: str = "base"
    colors: ThemeColors = ThemeColors()
    fonts: ThemeFonts = ThemeFonts()

    def apply_to_presentation(self, prs):
        """Apply theme settings to a presentation."""
        pass

    def apply_to_title_slide(self, slide):
        """Apply theme settings to a title slide."""
        pass

    def apply_to_content_slide(self, slide):
        """Apply theme settings to a content slide."""
        pass
