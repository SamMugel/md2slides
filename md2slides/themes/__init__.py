"""Theme definitions for md2slides."""

from md2slides.themes.default import DefaultTheme

THEMES = {
    "default": DefaultTheme,
}


def get_theme(name: str):
    """Get a theme class by name."""
    if name not in THEMES:
        raise ValueError(f"Unknown theme: {name}. Available: {list(THEMES.keys())}")
    return THEMES[name]()
