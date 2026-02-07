"""Theme definitions for md2slides."""

from md2slides.themes.corporate import CorporateTheme
from md2slides.themes.dark import DarkTheme
from md2slides.themes.default import DefaultTheme
from md2slides.themes.light import LightTheme

THEMES = {
    "default": DefaultTheme,
    "dark": DarkTheme,
    "light": LightTheme,
    "corporate": CorporateTheme,
}


def get_theme(name: str):
    """Get a theme class by name."""
    if name not in THEMES:
        raise ValueError(f"Unknown theme: {name}. Available: {list(THEMES.keys())}")
    return THEMES[name]()


def list_themes():
    """List available theme names."""
    return list(THEMES.keys())
