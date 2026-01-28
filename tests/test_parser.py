"""Tests for the markdown parser."""

import pytest

from md2slides.parser import (
    ListItem,
    MarkdownParser,
    Slide,
    TextRun,
    ValidationError,
)


class TestValidation:
    """Test input validation."""

    def test_empty_content_raises_error(self):
        """Empty content should raise ValidationError."""
        with pytest.raises(ValidationError, match="cannot be empty"):
            MarkdownParser("")

    def test_whitespace_only_content_raises_error(self):
        """Whitespace-only content should raise ValidationError."""
        with pytest.raises(ValidationError, match="cannot be empty"):
            MarkdownParser("   \n\t  ")

    def test_non_string_content_raises_error(self):
        """Non-string content should raise ValidationError."""
        with pytest.raises(ValidationError, match="must be a string"):
            MarkdownParser(123)  # type: ignore

    def test_none_content_raises_error(self):
        """None content should raise ValidationError."""
        with pytest.raises(ValidationError, match="must be a string"):
            MarkdownParser(None)  # type: ignore

    def test_no_headings_raises_error(self):
        """Content without headings should raise ValidationError."""
        parser = MarkdownParser("Just some text without any headings")
        with pytest.raises(ValidationError, match="at least one heading"):
            parser.parse()


class TestTitleSlide:
    """Test title slide parsing."""

    def test_h1_creates_title_slide(self):
        """H1 should create a title slide."""
        parser = MarkdownParser("# My Presentation")
        slides = parser.parse()

        assert len(slides) == 1
        assert slides[0].is_title_slide is True
        assert slides[0].title == "My Presentation"

    def test_h1_with_subtitle(self):
        """Content after H1 should become subtitle."""
        content = """# My Presentation

This is the subtitle text
"""
        parser = MarkdownParser(content)
        slides = parser.parse()

        assert slides[0].subtitle == "This is the subtitle text"

    def test_h1_without_subtitle(self):
        """H1 without following content should have no subtitle."""
        parser = MarkdownParser("# Title Only")
        slides = parser.parse()

        assert slides[0].subtitle is None

    def test_h1_with_multiline_subtitle(self):
        """Multiple lines after H1 should be joined as subtitle."""
        content = """# Quarterly Report

Q3 2024
Financial Summary
"""
        parser = MarkdownParser(content)
        slides = parser.parse()

        assert slides[0].subtitle == "Q3 2024\nFinancial Summary"


class TestContentSlides:
    """Test content slide parsing."""

    def test_h2_creates_content_slide(self):
        """H2 should create a content slide."""
        content = """# Title

## First Section

Some content here
"""
        parser = MarkdownParser(content)
        slides = parser.parse()

        assert len(slides) == 2
        assert slides[1].is_title_slide is False
        assert slides[1].title == "First Section"

    def test_multiple_h2_creates_multiple_slides(self):
        """Multiple H2s should create multiple content slides."""
        content = """# Title

## Section One

Content one

## Section Two

Content two

## Section Three

Content three
"""
        parser = MarkdownParser(content)
        slides = parser.parse()

        assert len(slides) == 4  # 1 title + 3 content
        assert slides[1].title == "Section One"
        assert slides[2].title == "Section Two"
        assert slides[3].title == "Section Three"

    def test_h2_only_document(self):
        """Document with only H2 headings should work."""
        content = """## First Slide

Content

## Second Slide

More content
"""
        parser = MarkdownParser(content)
        slides = parser.parse()

        assert len(slides) == 2
        assert slides[0].title == "First Slide"
        assert slides[1].title == "Second Slide"


class TestBulletLists:
    """Test bullet list parsing."""

    def test_simple_bullet_list(self):
        """Simple bullet list should be parsed correctly."""
        content = """## Slide

- Item one
- Item two
- Item three
"""
        parser = MarkdownParser(content)
        slides = parser.parse()

        assert len(slides[0].content) == 3
        assert isinstance(slides[0].content[0], ListItem)
        assert slides[0].content[0].ordered is False
        assert slides[0].content[0].content[0].text == "Item one"

    def test_nested_bullet_list(self):
        """Nested bullet list should preserve indentation levels."""
        content = """## Slide

- Level 0
  - Level 1
    - Level 2
  - Back to Level 1
"""
        parser = MarkdownParser(content)
        slides = parser.parse()

        assert slides[0].content[0].level == 0
        assert slides[0].content[1].level == 1
        assert slides[0].content[2].level == 2
        assert slides[0].content[3].level == 1

    def test_bullet_styles(self):
        """Different bullet styles (-, *, +) should work."""
        content = """## Slide

- Dash bullet
* Asterisk bullet
+ Plus bullet
"""
        parser = MarkdownParser(content)
        slides = parser.parse()

        assert len(slides[0].content) == 3
        for item in slides[0].content:
            assert isinstance(item, ListItem)
            assert item.ordered is False


class TestNumberedLists:
    """Test numbered list parsing."""

    def test_simple_numbered_list(self):
        """Numbered list should be parsed with numbers."""
        content = """## Slide

1. First item
2. Second item
3. Third item
"""
        parser = MarkdownParser(content)
        slides = parser.parse()

        assert len(slides[0].content) == 3
        assert slides[0].content[0].ordered is True
        assert slides[0].content[0].number == 1
        assert slides[0].content[1].number == 2
        assert slides[0].content[2].number == 3

    def test_numbered_list_with_parentheses(self):
        """Numbered list with parentheses should work."""
        content = """## Slide

1) First
2) Second
"""
        parser = MarkdownParser(content)
        slides = parser.parse()

        assert slides[0].content[0].ordered is True
        assert slides[0].content[0].number == 1


class TestInlineFormatting:
    """Test inline text formatting."""

    def test_bold_text(self):
        """Bold text should be parsed correctly."""
        content = """## Slide

- This has **bold** text
"""
        parser = MarkdownParser(content)
        slides = parser.parse()

        item = slides[0].content[0]
        assert isinstance(item, ListItem)
        # Should have: "This has ", "bold", " text"
        runs = item.content
        assert runs[0].text == "This has "
        assert runs[0].bold is False
        assert runs[1].text == "bold"
        assert runs[1].bold is True
        assert runs[2].text == " text"
        assert runs[2].bold is False

    def test_italic_text(self):
        """Italic text should be parsed correctly."""
        content = """## Slide

- This is *italic* text
"""
        parser = MarkdownParser(content)
        slides = parser.parse()

        item = slides[0].content[0]
        runs = item.content
        assert runs[1].text == "italic"
        assert runs[1].italic is True

    def test_bold_italic_text(self):
        """Bold+italic text should be parsed correctly."""
        content = """## Slide

- This is ***bold italic*** text
"""
        parser = MarkdownParser(content)
        slides = parser.parse()

        item = slides[0].content[0]
        runs = item.content
        assert runs[1].text == "bold italic"
        assert runs[1].bold is True
        assert runs[1].italic is True

    def test_underscore_formatting(self):
        """Underscore-based formatting should work."""
        content = """## Slide

- __bold__ and _italic_
"""
        parser = MarkdownParser(content)
        slides = parser.parse()

        item = slides[0].content[0]
        runs = item.content
        assert runs[0].text == "bold"
        assert runs[0].bold is True
        assert runs[2].text == "italic"
        assert runs[2].italic is True


class TestPlainText:
    """Test plain text content."""

    def test_plain_text_paragraph(self):
        """Plain text should be parsed as TextRun."""
        content = """## Slide

Just some plain text here
"""
        parser = MarkdownParser(content)
        slides = parser.parse()

        assert len(slides[0].content) == 1
        assert isinstance(slides[0].content[0], TextRun)
        assert slides[0].content[0].text == "Just some plain text here"


class TestExampleDocument:
    """Test the example from INSTRUCTIONS.md."""

    def test_quarterly_report_example(self):
        """Parse the example document from specifications."""
        content = """# Quarterly Report

Q3 2024 Financial Summary

## Revenue Overview

- Total revenue: $2.5M
- Growth: 15% YoY
- Key drivers:
  - New product launches
  - Expanded market reach

## Next Steps

- Increase marketing spend
- Launch in **two new regions**
- Hire additional sales staff
"""
        parser = MarkdownParser(content)
        slides = parser.parse()

        # Check structure
        assert len(slides) == 3

        # Title slide
        assert slides[0].is_title_slide is True
        assert slides[0].title == "Quarterly Report"
        assert slides[0].subtitle == "Q3 2024 Financial Summary"

        # Revenue Overview slide
        assert slides[1].title == "Revenue Overview"
        assert len(slides[1].content) == 5  # 3 top-level + 2 nested

        # Next Steps slide
        assert slides[2].title == "Next Steps"
        # Check bold formatting preserved
        second_item = slides[2].content[1]
        assert isinstance(second_item, ListItem)
        has_bold = any(run.bold for run in second_item.content)
        assert has_bold is True
