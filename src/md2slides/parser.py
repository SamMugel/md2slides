"""Markdown parser for extracting slide structure."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import List, Optional


class ValidationError(Exception):
    """Raised when markdown content fails validation."""

    pass


@dataclass
class TextRun:
    """A run of text with formatting."""

    text: str
    bold: bool = False
    italic: bool = False
    url: Optional[str] = None


@dataclass
class ListItem:
    """A list item with optional nesting."""

    content: List[TextRun]
    level: int = 0
    ordered: bool = False
    number: Optional[int] = None


@dataclass
class Slide:
    """Represents a single slide."""

    title: str
    content: List[ListItem | TextRun] = field(default_factory=list)
    is_title_slide: bool = False
    subtitle: Optional[str] = None


class MarkdownParser:
    """Parse markdown content into slide structures."""

    def __init__(self, content: str) -> None:
        """Initialize parser with markdown content.

        Args:
            content: The markdown string to parse.

        Raises:
            ValidationError: If content is empty or not a string.
        """
        self._validate_content(content)
        self.content = content
        self.slides: List[Slide] = []

    def _validate_content(self, content: str) -> None:
        """Validate the input content.

        Args:
            content: The content to validate.

        Raises:
            ValidationError: If content is invalid.
        """
        if not isinstance(content, str):
            raise ValidationError(
                f"Content must be a string, got {type(content).__name__}"
            )
        if not content.strip():
            raise ValidationError("Content cannot be empty")

    def parse(self) -> List[Slide]:
        """Parse the markdown content into slides.

        Returns:
            List of Slide objects representing the presentation.

        Raises:
            ValidationError: If no H1 heading is found.
        """
        lines = self.content.split("\n")
        self.slides = []

        current_slide: Optional[Slide] = None
        collecting_subtitle = False
        subtitle_lines: List[str] = []

        i = 0
        while i < len(lines):
            line = lines[i]

            # Check for H1 (title slide)
            if line.startswith("# ") and not line.startswith("## "):
                title = line[2:].strip()
                current_slide = Slide(
                    title=title, is_title_slide=True, content=[], subtitle=None
                )
                self.slides.append(current_slide)
                collecting_subtitle = True
                subtitle_lines = []
                i += 1
                continue

            # Check for H2 (content slide)
            if line.startswith("## "):
                if collecting_subtitle and subtitle_lines:
                    # Finalize subtitle for previous title slide
                    if self.slides and self.slides[0].is_title_slide:
                        self.slides[0].subtitle = "\n".join(subtitle_lines).strip()
                collecting_subtitle = False
                subtitle_lines = []

                title = line[3:].strip()
                current_slide = Slide(title=title, is_title_slide=False, content=[])
                self.slides.append(current_slide)
                i += 1
                continue

            # Handle content
            if current_slide is not None:
                if collecting_subtitle and current_slide.is_title_slide:
                    # Collect subtitle content
                    if line.strip():
                        subtitle_lines.append(line.strip())
                elif not current_slide.is_title_slide:
                    # Parse content for content slides
                    self._parse_content_line(line, current_slide)

            i += 1

        # Finalize any remaining subtitle
        if collecting_subtitle and subtitle_lines and self.slides:
            if self.slides[0].is_title_slide:
                self.slides[0].subtitle = "\n".join(subtitle_lines).strip()

        if not self.slides:
            raise ValidationError(
                "Document must contain at least one heading (# or ##)"
            )

        return self.slides

    def _parse_content_line(self, line: str, slide: Slide) -> None:
        """Parse a content line and add to slide.

        Args:
            line: The line to parse.
            slide: The slide to add content to.
        """
        stripped = line.rstrip()

        # Check for bullet list
        bullet_match = re.match(r"^(\s*)[-*+]\s+(.+)$", stripped)
        if bullet_match:
            indent = len(bullet_match.group(1))
            level = indent // 2  # 2 spaces per indent level
            content_text = bullet_match.group(2)
            text_runs = self._parse_inline_formatting(content_text)
            slide.content.append(ListItem(content=text_runs, level=level, ordered=False))
            return

        # Check for numbered list
        numbered_match = re.match(r"^(\s*)(\d+)[.)]\s+(.+)$", stripped)
        if numbered_match:
            indent = len(numbered_match.group(1))
            level = indent // 2
            number = int(numbered_match.group(2))
            content_text = numbered_match.group(3)
            text_runs = self._parse_inline_formatting(content_text)
            slide.content.append(
                ListItem(content=text_runs, level=level, ordered=True, number=number)
            )
            return

        # Plain text (non-empty)
        if stripped:
            text_runs = self._parse_inline_formatting(stripped)
            for run in text_runs:
                slide.content.append(run)

    def _parse_inline_formatting(self, text: str) -> List[TextRun]:
        """Parse inline formatting (bold, italic, URLs) in text.

        Args:
            text: The text to parse.

        Returns:
            List of TextRun objects with formatting applied.
        """
        runs: List[TextRun] = []

        # First, handle markdown links [caption](url)
        # Then handle bold+italic, bold, or italic
        # Order matters: check longer patterns first
        # Link pattern: [caption](url) - caption can be empty
        link_pattern = r"\[([^\]]*)\]\(([^)]+)\)"
        format_pattern = r"(\*\*\*|___)(.+?)(\*\*\*|___)|(\*\*|__)(.+?)(\*\*|__)|(\*|_)(.+?)(\*|_)"

        # Combined pattern: links first, then formatting
        combined_pattern = f"({link_pattern})|{format_pattern}"

        pos = 0
        for match in re.finditer(combined_pattern, text):
            # Add any text before this match as plain text
            if match.start() > pos:
                plain_text = text[pos : match.start()]
                if plain_text:
                    runs.append(TextRun(text=plain_text))

            # Check if this is a link match
            if match.group(1):  # Link match [caption](url)
                caption = match.group(2)
                url = match.group(3)
                # Use URL as display text if caption is empty
                display_text = caption if caption else url
                runs.append(TextRun(text=display_text, url=url))
            # Determine formatting type (groups shifted by 3 due to link pattern)
            elif match.group(4):  # Bold + Italic (*** or ___)
                runs.append(TextRun(text=match.group(5), bold=True, italic=True))
            elif match.group(7):  # Bold (** or __)
                runs.append(TextRun(text=match.group(8), bold=True, italic=False))
            elif match.group(10):  # Italic (* or _)
                runs.append(TextRun(text=match.group(11), bold=False, italic=True))

            pos = match.end()

        # Add any remaining text
        if pos < len(text):
            remaining = text[pos:]
            if remaining:
                runs.append(TextRun(text=remaining))

        # If no formatting found, return single plain run
        if not runs:
            runs.append(TextRun(text=text))

        return runs
