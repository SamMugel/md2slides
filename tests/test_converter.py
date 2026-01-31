"""Tests for the PPTX converter."""

import os
import tempfile
from pathlib import Path

import pytest
from pptx import Presentation
from pptx.oxml.ns import qn

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


class TestListFormatting:
    """Test that list formatting renders properly in PPTX."""

    def test_bullet_points_have_bullet_char(self):
        """Bullet points should have proper bullet character in XML."""
        content = """## Slide

- First bullet
- Second bullet
"""
        converter = MarkdownToPptxConverter()
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "test.pptx")
            converter.convert(content, output_path)

            prs = Presentation(output_path)
            slide = prs.slides[0]

            # Find paragraphs with bullet characters
            bullet_count = 0
            for shape in slide.shapes:
                if hasattr(shape, "text_frame"):
                    for para in shape.text_frame.paragraphs:
                        pPr = para._p.get_or_add_pPr()
                        buChar = pPr.find(qn('a:buChar'))
                        if buChar is not None:
                            bullet_count += 1

            assert bullet_count == 2

    def test_nested_bullets_have_different_chars(self):
        """Nested bullet points should have different bullet characters."""
        content = """## Slide

- Parent item
  - Child item
"""
        converter = MarkdownToPptxConverter()
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "test.pptx")
            converter.convert(content, output_path)

            prs = Presentation(output_path)
            slide = prs.slides[0]

            bullet_chars = []
            for shape in slide.shapes:
                if hasattr(shape, "text_frame"):
                    for para in shape.text_frame.paragraphs:
                        pPr = para._p.get_or_add_pPr()
                        buChar = pPr.find(qn('a:buChar'))
                        if buChar is not None:
                            bullet_chars.append(buChar.get('char'))

            assert len(bullet_chars) == 2
            # Parent and child should have different bullet chars
            assert bullet_chars[0] != bullet_chars[1]

    def test_numbered_list_has_auto_numbering(self):
        """Numbered lists should use buAutoNum element."""
        content = """## Slide

1. First step
2. Second step
3. Third step
"""
        converter = MarkdownToPptxConverter()
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "test.pptx")
            converter.convert(content, output_path)

            prs = Presentation(output_path)
            slide = prs.slides[0]

            auto_num_count = 0
            for shape in slide.shapes:
                if hasattr(shape, "text_frame"):
                    for para in shape.text_frame.paragraphs:
                        pPr = para._p.get_or_add_pPr()
                        buAutoNum = pPr.find(qn('a:buAutoNum'))
                        if buAutoNum is not None:
                            auto_num_count += 1
                            # Should be arabicPeriod type (1. 2. 3.)
                            assert buAutoNum.get('type') == 'arabicPeriod'

            assert auto_num_count == 3

    def test_nested_numbered_list_uses_alpha(self):
        """Nested numbered lists should use alphabetic numbering."""
        content = """## Slide

1. Parent step
   1. Child step
"""
        converter = MarkdownToPptxConverter()
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "test.pptx")
            converter.convert(content, output_path)

            prs = Presentation(output_path)
            slide = prs.slides[0]

            num_types = []
            for shape in slide.shapes:
                if hasattr(shape, "text_frame"):
                    for para in shape.text_frame.paragraphs:
                        pPr = para._p.get_or_add_pPr()
                        buAutoNum = pPr.find(qn('a:buAutoNum'))
                        if buAutoNum is not None:
                            num_types.append(buAutoNum.get('type'))

            assert len(num_types) == 2
            assert num_types[0] == 'arabicPeriod'  # Parent: 1. 2. 3.
            assert num_types[1] == 'alphaLcPeriod'  # Child: a. b. c.

    def test_mixed_list_formatting(self):
        """Mixed lists should have correct bullet/number formatting."""
        content = """## Slide

- Bullet item
  1. Numbered sub-item
"""
        converter = MarkdownToPptxConverter()
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "test.pptx")
            converter.convert(content, output_path)

            prs = Presentation(output_path)
            slide = prs.slides[0]

            has_bullet = False
            has_number = False
            for shape in slide.shapes:
                if hasattr(shape, "text_frame"):
                    for para in shape.text_frame.paragraphs:
                        pPr = para._p.get_or_add_pPr()
                        buChar = pPr.find(qn('a:buChar'))
                        buAutoNum = pPr.find(qn('a:buAutoNum'))
                        if buChar is not None:
                            has_bullet = True
                        if buAutoNum is not None:
                            has_number = True

            assert has_bullet is True
            assert has_number is True


class TestBrandStyling:
    """Test Multiverse Computing brand styling."""

    def test_slide_has_brand_background_color(self):
        """Slides should have Catskill White background (#F8FAFC)."""
        converter = MarkdownToPptxConverter()
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "test.pptx")
            converter.convert("# Test Title", output_path)

            prs = Presentation(output_path)
            slide = prs.slides[0]

            # Check background color
            bg_fill = slide.background.fill
            assert bg_fill.fore_color.rgb == (0xF8, 0xFA, 0xFC)

    def test_title_uses_brand_text_color(self):
        """Title text should use Woodsmoke color (#111417)."""
        converter = MarkdownToPptxConverter()
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "test.pptx")
            converter.convert("# Test Title", output_path)

            prs = Presentation(output_path)
            slide = prs.slides[0]

            # Find title text and check color
            for shape in slide.shapes:
                if hasattr(shape, "text_frame"):
                    for para in shape.text_frame.paragraphs:
                        for run in para.runs:
                            if "Test Title" in run.text:
                                assert run.font.color.rgb == (0x11, 0x14, 0x17)

    def test_title_uses_montserrat_font(self):
        """Title text should use Montserrat font."""
        converter = MarkdownToPptxConverter()
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "test.pptx")
            converter.convert("# Test Title", output_path)

            prs = Presentation(output_path)
            slide = prs.slides[0]

            for shape in slide.shapes:
                if hasattr(shape, "text_frame"):
                    for para in shape.text_frame.paragraphs:
                        for run in para.runs:
                            if "Test Title" in run.text:
                                assert run.font.name == "Montserrat"

    def test_body_uses_open_sans_font(self):
        """Body text should use Open Sans font."""
        content = """## Slide

- Body text here
"""
        converter = MarkdownToPptxConverter()
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "test.pptx")
            converter.convert(content, output_path)

            prs = Presentation(output_path)
            slide = prs.slides[0]

            for shape in slide.shapes:
                if hasattr(shape, "text_frame"):
                    for para in shape.text_frame.paragraphs:
                        for run in para.runs:
                            if "Body text here" in run.text:
                                assert run.font.name == "Open Sans"

    def test_h1_uses_correct_size(self):
        """H1 title should use 28pt font size (issue #3)."""
        converter = MarkdownToPptxConverter()
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "test.pptx")
            converter.convert("# Test Title", output_path)

            prs = Presentation(output_path)
            slide = prs.slides[0]

            for shape in slide.shapes:
                if hasattr(shape, "text_frame"):
                    for para in shape.text_frame.paragraphs:
                        for run in para.runs:
                            if "Test Title" in run.text:
                                assert run.font.size.pt == 28

    def test_h2_uses_correct_size(self):
        """H2 title should use 28pt font size (issue #3)."""
        content = """## Section Header

Content here
"""
        converter = MarkdownToPptxConverter()
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "test.pptx")
            converter.convert(content, output_path)

            prs = Presentation(output_path)
            slide = prs.slides[0]

            for shape in slide.shapes:
                if hasattr(shape, "text_frame"):
                    for para in shape.text_frame.paragraphs:
                        for run in para.runs:
                            if "Section Header" in run.text:
                                assert run.font.size.pt == 28

    def test_body_uses_correct_size(self):
        """Body text should use 18pt font size (issue #4)."""
        content = """## Slide

- Body text here
"""
        converter = MarkdownToPptxConverter()
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "test.pptx")
            converter.convert(content, output_path)

            prs = Presentation(output_path)
            slide = prs.slides[0]

            for shape in slide.shapes:
                if hasattr(shape, "text_frame"):
                    for para in shape.text_frame.paragraphs:
                        for run in para.runs:
                            if "Body text here" in run.text:
                                assert run.font.size.pt == 18


class TestStyleEnhancements:
    """Test style enhancements from issue #3."""

    def test_child_items_use_dark_grey(self):
        """Child list items should use dark grey color (#404040)."""
        content = """## Slide

- Parent item
  - Child item
"""
        converter = MarkdownToPptxConverter()
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "test.pptx")
            converter.convert(content, output_path)

            prs = Presentation(output_path)
            slide = prs.slides[0]

            found_child = False
            for shape in slide.shapes:
                if hasattr(shape, "text_frame"):
                    for para in shape.text_frame.paragraphs:
                        for run in para.runs:
                            if "Child item" in run.text:
                                assert run.font.color.rgb == (0x40, 0x40, 0x40)
                                found_child = True
            assert found_child is True

    def test_parent_items_use_woodsmoke(self):
        """Parent list items should use Woodsmoke color (#111417)."""
        content = """## Slide

- Parent item
  - Child item
"""
        converter = MarkdownToPptxConverter()
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "test.pptx")
            converter.convert(content, output_path)

            prs = Presentation(output_path)
            slide = prs.slides[0]

            found_parent = False
            for shape in slide.shapes:
                if hasattr(shape, "text_frame"):
                    for para in shape.text_frame.paragraphs:
                        for run in para.runs:
                            if "Parent item" in run.text:
                                assert run.font.color.rgb == (0x11, 0x14, 0x17)
                                found_parent = True
            assert found_parent is True

    def test_list_items_have_spacing(self):
        """List items should have 4 spaces after bullet/number."""
        content = """## Slide

- Test item
"""
        converter = MarkdownToPptxConverter()
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "test.pptx")
            converter.convert(content, output_path)

            prs = Presentation(output_path)
            slide = prs.slides[0]

            # Find the content shape and verify spacing run exists
            found_spacing = False
            for shape in slide.shapes:
                if hasattr(shape, "text_frame"):
                    for para in shape.text_frame.paragraphs:
                        runs = list(para.runs)
                        # First run should be 4 spaces
                        for i, run in enumerate(runs):
                            if run.text == "    ":
                                found_spacing = True
            assert found_spacing is True


class TestLogoPlacement:
    """Test Multiverse Computing logo placement."""

    def test_logo_added_to_title_slide(self):
        """Title slide should have logo when logo file exists."""
        from pptx.enum.shapes import MSO_SHAPE_TYPE

        converter = MarkdownToPptxConverter()
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "test.pptx")
            converter.convert("# Test Title", output_path)

            prs = Presentation(output_path)
            slide = prs.slides[0]

            # Check if logo path was found (depends on resources folder existing)
            if converter._logo_path:
                has_picture = any(
                    shape.shape_type == MSO_SHAPE_TYPE.PICTURE
                    for shape in slide.shapes
                )
                assert has_picture is True

    def test_logo_added_to_content_slide(self):
        """Content slide should have logo when logo file exists."""
        from pptx.enum.shapes import MSO_SHAPE_TYPE

        content = """## Slide

Content here
"""
        converter = MarkdownToPptxConverter()
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "test.pptx")
            converter.convert(content, output_path)

            prs = Presentation(output_path)
            slide = prs.slides[0]

            if converter._logo_path:
                has_picture = any(
                    shape.shape_type == MSO_SHAPE_TYPE.PICTURE
                    for shape in slide.shapes
                )
                assert has_picture is True


class TestCosmeticChanges:
    """Test cosmetic changes from issue #4."""

    def test_subtitle_uses_dark_grey(self):
        """Subtitle should use dark grey color (#404040)."""
        content = """# Title

Subtitle text
"""
        converter = MarkdownToPptxConverter()
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "test.pptx")
            converter.convert(content, output_path)

            prs = Presentation(output_path)
            slide = prs.slides[0]

            found_subtitle = False
            for shape in slide.shapes:
                if hasattr(shape, "text_frame"):
                    for para in shape.text_frame.paragraphs:
                        for run in para.runs:
                            if "Subtitle text" in run.text:
                                assert run.font.color.rgb == (0x40, 0x40, 0x40)
                                found_subtitle = True
            assert found_subtitle is True

    def test_text_has_line_spacing(self):
        """Text should have half-space line spacing (space_after = 9pt)."""
        content = """## Slide

- Item one
- Item two
"""
        converter = MarkdownToPptxConverter()
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "test.pptx")
            converter.convert(content, output_path)

            prs = Presentation(output_path)
            slide = prs.slides[0]

            found_spacing = False
            for shape in slide.shapes:
                if hasattr(shape, "text_frame"):
                    for para in shape.text_frame.paragraphs:
                        # Check if space_after is set
                        if para.space_after is not None and para.space_after.pt == 9:
                            found_spacing = True
            assert found_spacing is True

    def test_text_size_is_18pt(self):
        """Body text should be 18pt (issue #4)."""
        content = """## Slide

- Body text here
"""
        converter = MarkdownToPptxConverter()
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "test.pptx")
            converter.convert(content, output_path)

            prs = Presentation(output_path)
            slide = prs.slides[0]

            found_text = False
            for shape in slide.shapes:
                if hasattr(shape, "text_frame"):
                    for para in shape.text_frame.paragraphs:
                        for run in para.runs:
                            if "Body text here" in run.text:
                                assert run.font.size.pt == 18
                                found_text = True
            assert found_text is True


class TestTextScaling:
    """Test auto-scaling text to fit slide (issue #7)."""

    def test_content_frame_has_auto_size(self):
        """Content frame should have TEXT_TO_FIT_SHAPE auto size mode."""
        from pptx.enum.text import MSO_AUTO_SIZE

        content = """## Slide

- Item one
- Item two
"""
        converter = MarkdownToPptxConverter()
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "test.pptx")
            converter.convert(content, output_path)

            prs = Presentation(output_path)
            slide = prs.slides[0]

            # Find content shape (not the title)
            found_auto_size = False
            for shape in slide.shapes:
                if hasattr(shape, "text_frame"):
                    tf = shape.text_frame
                    # Content frame has list items
                    all_text = "".join(r.text for p in tf.paragraphs for r in p.runs)
                    if "Item one" in all_text:
                        assert tf.auto_size == MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
                        found_auto_size = True
            assert found_auto_size is True

    def test_long_content_fits_slide(self):
        """Long content should be contained within slide boundaries."""
        # Create content with many items that would normally overflow
        items = "\n".join([f"- Item {i} with some longer text" for i in range(20)])
        content = f"""## Slide with Many Items

{items}
"""
        converter = MarkdownToPptxConverter()
        with tempfile.TemporaryDirectory() as tmpdir:
            output_path = os.path.join(tmpdir, "test.pptx")
            # Should not raise any errors
            result = converter.convert(content, output_path)
            assert os.path.exists(result)

            prs = Presentation(output_path)
            slide = prs.slides[0]

            # Verify content shape exists
            found_content = False
            for shape in slide.shapes:
                if hasattr(shape, "text_frame"):
                    all_text = "".join(r.text for p in shape.text_frame.paragraphs for r in p.runs)
                    if "Item 0" in all_text:
                        found_content = True
                        # Verify auto_size is set for shrinking
                        from pptx.enum.text import MSO_AUTO_SIZE
                        assert shape.text_frame.auto_size == MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
            assert found_content is True
