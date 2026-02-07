"""Markdown parsing logic for md2slides."""

from dataclasses import dataclass, field
from typing import List, Optional, Union

import mistune
from mistune.plugins.table import table as table_plugin


@dataclass
class InlineSpan:
    """An inline text span, possibly with a hyperlink."""
    text: str
    href: Optional[str] = None  # URL if this is a hyperlink
    bold: bool = False
    italic: bool = False
    code: bool = False


@dataclass
class TextElement:
    """A text element (paragraph or inline text)."""
    text: str
    bold: bool = False
    italic: bool = False
    code: bool = False
    spans: List["InlineSpan"] = field(default_factory=list)  # Rich inline content


@dataclass
class ListItem:
    """A single list item, which may contain nested lists."""
    text: str
    children: List["ListItem"] = field(default_factory=list)
    spans: List["InlineSpan"] = field(default_factory=list)  # Rich inline content


@dataclass
class BulletList:
    """A bullet list element."""
    items: List[ListItem] = field(default_factory=list)
    ordered: bool = False


@dataclass
class CodeBlock:
    """A fenced code block."""
    code: str
    language: Optional[str] = None


@dataclass
class Table:
    """A markdown table."""
    headers: List[str] = field(default_factory=list)
    rows: List[List[str]] = field(default_factory=list)
    alignments: List[Optional[str]] = field(default_factory=list)


@dataclass
class Image:
    """An image reference."""
    src: str
    alt: str = ""
    title: str = ""


# Union type for all slide content elements
SlideElement = Union[TextElement, BulletList, CodeBlock, Table, Image]


@dataclass
class Slide:
    """A single slide in the presentation."""
    title: str
    elements: List[SlideElement] = field(default_factory=list)


@dataclass
class Section:
    """A section of the presentation (defined by H1)."""
    title: str
    subtitle: str = ""  # Text immediately after H1, before any H2
    slides: List[Slide] = field(default_factory=list)


@dataclass
class Presentation:
    """The complete parsed presentation structure."""
    title: str = ""
    sections: List[Section] = field(default_factory=list)


class MarkdownParser:
    """Parse markdown into a structured presentation format."""

    def __init__(self):
        self._md = mistune.create_markdown(renderer=None, plugins=[table_plugin])

    def parse(self, markdown_text: str) -> Presentation:
        """Parse markdown text into a Presentation structure."""
        tokens = self._md(markdown_text)
        return self._build_presentation(tokens)

    def _build_presentation(self, tokens: list) -> Presentation:
        """Build a Presentation from mistune tokens."""
        presentation = Presentation()
        current_section: Optional[Section] = None
        current_slide: Optional[Slide] = None
        awaiting_subtitle = False  # True when we just saw an H1 and expect a subtitle

        for token in tokens:
            token_type = token.get("type")

            if token_type == "heading":
                level = token["attrs"]["level"]
                text = self._extract_text(token.get("children", []))

                if level == 1:
                    # H1 creates a new section
                    if not presentation.title:
                        presentation.title = text
                    current_section = Section(title=text)
                    presentation.sections.append(current_section)
                    current_slide = None
                    awaiting_subtitle = True  # Next paragraph could be a subtitle

                elif level == 2:
                    # H2 creates a new slide (and consumes any pending subtitle expectation)
                    awaiting_subtitle = False
                    current_slide = Slide(title=text)
                    if current_section is None:
                        # Create a default section if none exists
                        current_section = Section(title=presentation.title or "Untitled")
                        presentation.sections.append(current_section)
                    current_section.slides.append(current_slide)

            elif token_type == "paragraph":
                # Check if this paragraph should be treated as a section subtitle
                if awaiting_subtitle and current_section is not None and current_slide is None:
                    # This is the first paragraph after H1, before any H2 - use as subtitle
                    text = self._extract_text_excluding_images(token.get("children", []))
                    if text:
                        current_section.subtitle = text
                    awaiting_subtitle = False
                    continue  # Don't process this paragraph as slide content

                awaiting_subtitle = False  # Any other content clears the subtitle expectation

                if current_slide is not None:
                    # Check for images in paragraph
                    images = self._extract_images(token.get("children", []))
                    for img in images:
                        current_slide.elements.append(img)

                    # Extract inline spans with hyperlink support
                    spans = self._extract_inline_spans(token.get("children", []))
                    # Also get plain text for backwards compatibility
                    text = self._extract_text_excluding_images(token.get("children", []))
                    if text:
                        current_slide.elements.append(TextElement(text=text, spans=spans))

            elif token_type == "list":
                if current_slide is not None:
                    bullet_list = self._parse_list(token)
                    current_slide.elements.append(bullet_list)

            elif token_type == "block_code":
                if current_slide is not None:
                    code = token.get("raw", "")
                    info = token.get("attrs", {}).get("info")
                    current_slide.elements.append(CodeBlock(code=code, language=info))

            elif token_type == "table":
                if current_slide is not None:
                    table = self._parse_table(token)
                    current_slide.elements.append(table)

        return presentation

    def _extract_text(self, children: list) -> str:
        """Extract plain text from a list of inline tokens."""
        parts = []
        for child in children:
            child_type = child.get("type")
            if child_type == "text":
                parts.append(child.get("raw", ""))
            elif child_type == "codespan":
                parts.append(child.get("raw", ""))
            elif child_type == "emphasis":
                parts.append(self._extract_text(child.get("children", [])))
            elif child_type == "strong":
                parts.append(self._extract_text(child.get("children", [])))
            elif child_type == "link":
                parts.append(self._extract_text(child.get("children", [])))
            elif child_type == "image":
                parts.append(child.get("attrs", {}).get("alt", ""))
            elif child_type == "softbreak":
                parts.append(" ")
            elif child_type == "linebreak":
                parts.append("\n")
        return "".join(parts)

    def _extract_images(self, children: list) -> List[Image]:
        """Extract images from inline tokens.

        Mistune can parse images in two ways:
        1. As 'image' tokens (when on their own line)
        2. As text("!") followed by link (when inline with text)
        """
        images = []
        i = 0
        while i < len(children):
            child = children[i]
            child_type = child.get("type")

            # Handle direct image tokens
            if child_type == "image":
                attrs = child.get("attrs", {})
                alt = self._extract_text(child.get("children", []))
                src = attrs.get("url", "")
                title = attrs.get("title", "")
                images.append(Image(src=src, alt=alt, title=title))
                i += 1
                continue

            # Look for pattern: text with "!" followed by link (inline images)
            if (
                child_type == "text"
                and child.get("raw", "").endswith("!")
                and i + 1 < len(children)
                and children[i + 1].get("type") == "link"
            ):
                link = children[i + 1]
                attrs = link.get("attrs", {})
                alt = self._extract_text(link.get("children", []))
                src = attrs.get("url", "")
                title = attrs.get("title", "")
                images.append(Image(src=src, alt=alt, title=title))
                i += 2  # Skip both the "!" text and the link
            else:
                i += 1
        return images

    def _extract_text_excluding_images(self, children: list) -> str:
        """Extract text from inline tokens, excluding image patterns."""
        parts = []
        i = 0
        while i < len(children):
            child = children[i]
            child_type = child.get("type")

            # Skip image tokens
            if child_type == "image":
                i += 1
                continue

            # Skip image patterns (text ending with "!" followed by link)
            if (
                child_type == "text"
                and child.get("raw", "").endswith("!")
                and i + 1 < len(children)
                and children[i + 1].get("type") == "link"
            ):
                # Add any text before the "!"
                raw = child.get("raw", "")
                if len(raw) > 1:
                    parts.append(raw[:-1])
                i += 2  # Skip both
                continue

            if child_type == "text":
                parts.append(child.get("raw", ""))
            elif child_type == "codespan":
                parts.append(child.get("raw", ""))
            elif child_type == "emphasis":
                parts.append(self._extract_text(child.get("children", [])))
            elif child_type == "strong":
                parts.append(self._extract_text(child.get("children", [])))
            elif child_type == "link":
                parts.append(self._extract_text(child.get("children", [])))
            elif child_type == "softbreak":
                parts.append(" ")
            elif child_type == "linebreak":
                parts.append("\n")
            i += 1

        return "".join(parts).strip()

    def _extract_inline_spans(self, children: list) -> List[InlineSpan]:
        """Extract inline spans from tokens, preserving hyperlinks.

        Returns a list of InlineSpan objects that can represent both plain text
        and clickable hyperlinks.
        """
        spans = []
        i = 0
        while i < len(children):
            child = children[i]
            child_type = child.get("type")

            # Skip image tokens
            if child_type == "image":
                i += 1
                continue

            # Skip image patterns (text ending with "!" followed by link)
            if (
                child_type == "text"
                and child.get("raw", "").endswith("!")
                and i + 1 < len(children)
                and children[i + 1].get("type") == "link"
            ):
                # Add any text before the "!" as a plain span
                raw = child.get("raw", "")
                if len(raw) > 1:
                    spans.append(InlineSpan(text=raw[:-1]))
                i += 2  # Skip both
                continue

            if child_type == "text":
                text = child.get("raw", "")
                if text:
                    spans.append(InlineSpan(text=text))

            elif child_type == "codespan":
                text = child.get("raw", "")
                if text:
                    spans.append(InlineSpan(text=text, code=True))

            elif child_type == "emphasis":
                text = self._extract_text(child.get("children", []))
                if text:
                    spans.append(InlineSpan(text=text, italic=True))

            elif child_type == "strong":
                text = self._extract_text(child.get("children", []))
                if text:
                    spans.append(InlineSpan(text=text, bold=True))

            elif child_type == "link":
                # Extract link text and URL
                link_text = self._extract_text(child.get("children", []))
                href = child.get("attrs", {}).get("url", "")
                if link_text:
                    spans.append(InlineSpan(text=link_text, href=href))

            elif child_type == "softbreak":
                spans.append(InlineSpan(text=" "))

            elif child_type == "linebreak":
                spans.append(InlineSpan(text="\n"))

            i += 1

        return spans

    def _parse_list(self, token: dict) -> BulletList:
        """Parse a list token into a BulletList."""
        ordered = token.get("attrs", {}).get("ordered", False)
        items = []

        for item_token in token.get("children", []):
            if item_token.get("type") == "list_item":
                list_item = self._parse_list_item(item_token)
                items.append(list_item)

        return BulletList(items=items, ordered=ordered)

    def _parse_list_item(self, token: dict) -> ListItem:
        """Parse a list_item token into a ListItem."""
        text_parts = []
        all_spans = []
        children = []

        for child in token.get("children", []):
            child_type = child.get("type")
            if child_type in ("block_text", "paragraph"):
                text_parts.append(self._extract_text(child.get("children", [])))
                all_spans.extend(self._extract_inline_spans(child.get("children", [])))
            elif child_type == "list":
                # Nested list
                nested_list = self._parse_list(child)
                children.extend(nested_list.items)

        return ListItem(text=" ".join(text_parts), children=children, spans=all_spans)

    def _parse_table(self, token: dict) -> Table:
        """Parse a table token into a Table."""
        headers = []
        rows = []
        alignments = []

        for child in token.get("children", []):
            child_type = child.get("type")

            if child_type == "table_head":
                # table_head contains table_cell directly (headers)
                for cell in child.get("children", []):
                    if cell.get("type") == "table_cell":
                        text = self._extract_text(cell.get("children", []))
                        headers.append(text)
                        align = cell.get("attrs", {}).get("align")
                        alignments.append(align)

            elif child_type == "table_body":
                # table_body contains table_row elements
                for row in child.get("children", []):
                    if row.get("type") == "table_row":
                        row_data = []
                        for cell in row.get("children", []):
                            if cell.get("type") == "table_cell":
                                text = self._extract_text(cell.get("children", []))
                                row_data.append(text)
                        rows.append(row_data)

        return Table(headers=headers, rows=rows, alignments=alignments)


def parse_markdown(markdown_text: str) -> Presentation:
    """Convenience function to parse markdown text."""
    parser = MarkdownParser()
    return parser.parse(markdown_text)
