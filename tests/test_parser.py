"""Tests for the markdown parser."""

import pytest

from md2slides.parser import (
    Image,
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


class TestUrlParsing:
    """Test URL/hyperlink parsing (issue #6)."""

    def test_url_with_caption(self):
        """URL with caption should be parsed correctly."""
        content = """## Slide

- Visit [Google](https://google.com)
"""
        parser = MarkdownParser(content)
        slides = parser.parse()

        item = slides[0].content[0]
        assert isinstance(item, ListItem)
        runs = item.content
        # Should have: "Visit ", "Google"
        assert runs[0].text == "Visit "
        assert runs[0].url is None
        assert runs[1].text == "Google"
        assert runs[1].url == "https://google.com"

    def test_url_without_caption(self):
        """URL without caption should use URL as display text."""
        content = """## Slide

- Check [](https://example.com)
"""
        parser = MarkdownParser(content)
        slides = parser.parse()

        item = slides[0].content[0]
        runs = item.content
        # Should have: "Check ", "https://example.com"
        assert runs[1].text == "https://example.com"
        assert runs[1].url == "https://example.com"

    def test_multiple_urls_in_line(self):
        """Multiple URLs in one line should all be parsed."""
        content = """## Slide

- Visit [Google](https://google.com) or [Bing](https://bing.com)
"""
        parser = MarkdownParser(content)
        slides = parser.parse()

        item = slides[0].content[0]
        runs = item.content
        # Should have: "Visit ", "Google", " or ", "Bing"
        assert runs[1].text == "Google"
        assert runs[1].url == "https://google.com"
        assert runs[3].text == "Bing"
        assert runs[3].url == "https://bing.com"

    def test_url_in_plain_text(self):
        """URLs should work in plain text paragraphs."""
        content = """## Slide

Learn more at [our website](https://multiverse.com)
"""
        parser = MarkdownParser(content)
        slides = parser.parse()

        # Content should be TextRuns
        runs = slides[0].content
        # Find the URL run
        url_run = None
        for run in runs:
            if isinstance(run, TextRun) and run.url:
                url_run = run
                break

        assert url_run is not None
        assert url_run.text == "our website"
        assert url_run.url == "https://multiverse.com"

    def test_url_with_formatting(self):
        """URLs mixed with formatting should work correctly."""
        content = """## Slide

- **Bold** and [Link](https://example.com) text
"""
        parser = MarkdownParser(content)
        slides = parser.parse()

        item = slides[0].content[0]
        runs = item.content
        # Should have: "Bold", " and ", "Link", " text"
        has_bold = any(r.bold for r in runs)
        has_link = any(r.url == "https://example.com" for r in runs)
        assert has_bold is True
        assert has_link is True


class TestImageParsing:
    """Test image parsing (issue #5)."""

    def test_image_with_caption(self):
        """Image with caption should be parsed correctly."""
        content = """## Slide

- Some text
![Figure 1](image.png)
"""
        parser = MarkdownParser(content)
        slides = parser.parse()

        assert slides[0].image is not None
        assert slides[0].image.path == "image.png"
        assert slides[0].image.caption == "Figure 1"

    def test_image_without_caption(self):
        """Image without caption should have None caption."""
        content = """## Slide

- Some text
![](image.png)
"""
        parser = MarkdownParser(content)
        slides = parser.parse()

        assert slides[0].image is not None
        assert slides[0].image.path == "image.png"
        assert slides[0].image.caption is None

    def test_image_with_path(self):
        """Image with path should preserve full path."""
        content = """## Slide

![Chart](resources/charts/figure1.png)
"""
        parser = MarkdownParser(content)
        slides = parser.parse()

        assert slides[0].image is not None
        assert slides[0].image.path == "resources/charts/figure1.png"
        assert slides[0].image.caption == "Chart"

    def test_slide_without_image(self):
        """Slide without image should have None image."""
        content = """## Slide

- Just bullet points
- No image here
"""
        parser = MarkdownParser(content)
        slides = parser.parse()

        assert slides[0].image is None

    def test_image_with_text_content(self):
        """Image and text content should both be parsed."""
        content = """## Slide

- First item
- Second item
![My Image](photo.jpg)
"""
        parser = MarkdownParser(content)
        slides = parser.parse()

        # Check text content
        assert len(slides[0].content) == 2
        assert slides[0].content[0].content[0].text == "First item"

        # Check image
        assert slides[0].image is not None
        assert slides[0].image.path == "photo.jpg"
        assert slides[0].image.caption == "My Image"

    def test_multiple_images_uses_last(self):
        """If multiple images are present, the last one is used."""
        content = """## Slide

![First](first.png)
![Second](second.png)
"""
        parser = MarkdownParser(content)
        slides = parser.parse()

        # Last image wins
        assert slides[0].image is not None
        assert slides[0].image.path == "second.png"
        assert slides[0].image.caption == "Second"
