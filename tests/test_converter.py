"""Tests for the PPTX converter."""

import os
import tempfile
from pathlib import Path

import pytest
from pptx import Presentation

from md2slides.converter import MarkdownToPptxConverter, convert_file
from md2slides.parser import ValidationError


class TestConverterValidation:
    """Test converter input validation."""

    def test_empty_output_path_raises_error(self):
        """Empty output path should raise ValidationError."""
        converter = MarkdownToPptxConverter()
        with pytest.raises(ValidationError, match="cannot be empty"):
            converter.convert("# Title", "")

    def test_whitespace_output_path_raises_error(self):
        """Whitespace-only output path should raise ValidationError."""
        converter = MarkdownToPptxConverter()
        with pytest.raises(ValidationError, match="cannot be empty"):
            converter.convert("# Title", "   ")

    def test_non_string_output_path_raises_error(self):
        """Non-string output path should raise ValidationError."""
        converter = MarkdownToPptxConverter()
        with pytest.raises(ValidationError, match="must be a string"):
            converter.convert("# Title", 123)  # type: ignore

    def test_wrong_extension_raises_error(self):
        """Non-.pptx extension should raise ValidationError."""
        converter = MarkdownToPptxConverter()
        with pytest.raises(ValidationError, match=".pptx extension"):
            converter.convert("# Title", "output.pdf")


class TestConverterOutput:
    """Test converter PPTX output."""

    def test_creates_pptx_file(self):
        """Converter should create a valid PPTX file."""
        converter = MarkdownToPptxConverter()
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "test.pptx")
            result = converter.convert("# Test Title", output_path)

            assert os.path.exists(result)
            # Verify it's a valid PPTX
            prs = Presentation(result)
            assert len(prs.slides) == 1

    def test_returns_absolute_path(self):
        """Converter should return absolute path."""
        converter = MarkdownToPptxConverter()
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "test.pptx")
            result = converter.convert("# Test", output_path)

            assert os.path.isabs(result)

    def test_creates_output_directory(self):
        """Converter should create output directory if needed."""
        converter = MarkdownToPptxConverter()
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "subdir", "nested", "test.pptx")
            result = converter.convert("# Test", output_path)

            assert os.path.exists(result)


class TestSlideContent:
    """Test slide content generation."""

    def test_title_slide_content(self):
        """Title slide should have correct title."""
        converter = MarkdownToPptxConverter()
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "test.pptx")
            converter.convert("# My Presentation", output_path)

            prs = Presentation(output_path)
            slide = prs.slides[0]

            # Find text in shapes
            texts = []
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    texts.append(shape.text)

            assert "My Presentation" in texts

    def test_title_slide_with_subtitle(self):
        """Title slide should include subtitle."""
        content = """# Main Title

This is the subtitle
"""
        converter = MarkdownToPptxConverter()
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "test.pptx")
            converter.convert(content, output_path)

            prs = Presentation(output_path)
            slide = prs.slides[0]

            texts = []
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    texts.append(shape.text)

            assert "Main Title" in texts
            assert "This is the subtitle" in texts

    def test_content_slide_title(self):
        """Content slide should have correct title."""
        content = """# Title

## Section Header

Some content
"""
        converter = MarkdownToPptxConverter()
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "test.pptx")
            converter.convert(content, output_path)

            prs = Presentation(output_path)
            assert len(prs.slides) == 2

            # Check second slide has the section header
            slide = prs.slides[1]
            texts = []
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    texts.append(shape.text)

            assert any("Section Header" in t for t in texts)

    def test_bullet_list_content(self):
        """Bullet list should appear in content slide."""
        content = """## Slide

- First item
- Second item
- Third item
"""
        converter = MarkdownToPptxConverter()
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "test.pptx")
            converter.convert(content, output_path)

            prs = Presentation(output_path)
            slide = prs.slides[0]

            # Find content shape
            all_text = ""
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    all_text += shape.text

            assert "First item" in all_text
            assert "Second item" in all_text
            assert "Third item" in all_text

    def test_multiple_slides(self):
        """Multiple H2s should create multiple slides."""
        content = """# Title

## Slide 1

Content 1

## Slide 2

Content 2

## Slide 3

Content 3
"""
        converter = MarkdownToPptxConverter()
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "test.pptx")
            converter.convert(content, output_path)

            prs = Presentation(output_path)
            assert len(prs.slides) == 4  # 1 title + 3 content


class TestConvertFile:
    """Test the convert_file function."""

    def test_convert_file_basic(self):
        """convert_file should work with basic markdown file."""
        with tempfile.TemporaryDirectory() as tmpdir:
            # Create input file
            input_path = os.path.join(tmpdir, "input.md")
            with open(input_path, "w") as f:
                f.write("# Test Presentation\n\n## Slide One\n\nContent")

            output_path = os.path.join(tmpdir, "output.pptx")
            result = convert_file(input_path, output_path)

            assert os.path.exists(result)
            prs = Presentation(result)
            assert len(prs.slides) == 2

    def test_convert_file_default_output(self):
        """convert_file should use default output path."""
        with tempfile.TemporaryDirectory() as tmpdir:
            input_path = os.path.join(tmpdir, "presentation.md")
            with open(input_path, "w") as f:
                f.write("# Test\n\n## Content\n\nText")

            result = convert_file(input_path)

            expected_output = os.path.join(tmpdir, "presentation.pptx")
            assert result == os.path.abspath(expected_output)
            assert os.path.exists(result)

    def test_convert_file_nonexistent_raises_error(self):
        """convert_file should raise FileNotFoundError for missing file."""
        with pytest.raises(FileNotFoundError, match="not found"):
            convert_file("/nonexistent/path/file.md")

    def test_convert_file_empty_path_raises_error(self):
        """convert_file should raise ValidationError for empty path."""
        with pytest.raises(ValidationError, match="cannot be empty"):
            convert_file("")

    def test_convert_file_non_string_raises_error(self):
        """convert_file should raise ValidationError for non-string path."""
        with pytest.raises(ValidationError, match="must be a string"):
            convert_file(123)  # type: ignore

    def test_convert_file_directory_raises_error(self):
        """convert_file should raise ValidationError for directory path."""
        with tempfile.TemporaryDirectory() as tmpdir:
            with pytest.raises(ValidationError, match="not a file"):
                convert_file(tmpdir)


class TestFormattingPreservation:
    """Test that text formatting is preserved in output."""

    def test_bold_text_preserved(self):
        """Bold formatting should be preserved in PPTX."""
        content = """## Slide

- This has **bold** text
"""
        converter = MarkdownToPptxConverter()
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "test.pptx")
            converter.convert(content, output_path)

            prs = Presentation(output_path)
            slide = prs.slides[0]

            # Find content shape and check for bold runs
            has_bold = False
            for shape in slide.shapes:
                if hasattr(shape, "text_frame"):
                    for para in shape.text_frame.paragraphs:
                        for run in para.runs:
                            if run.font.bold and "bold" in run.text:
                                has_bold = True

            assert has_bold is True

    def test_italic_text_preserved(self):
        """Italic formatting should be preserved in PPTX."""
        content = """## Slide

- This has *italic* text
"""
        converter = MarkdownToPptxConverter()
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "test.pptx")
            converter.convert(content, output_path)

            prs = Presentation(output_path)
            slide = prs.slides[0]

            # Find content shape and check for italic runs
            has_italic = False
            for shape in slide.shapes:
                if hasattr(shape, "text_frame"):
                    for para in shape.text_frame.paragraphs:
                        for run in para.runs:
                            if run.font.italic and "italic" in run.text:
                                has_italic = True

            assert has_italic is True
