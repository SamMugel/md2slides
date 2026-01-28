"""PowerPoint converter for creating PPTX from parsed markdown."""

from __future__ import annotations

import os
from pathlib import Path
from typing import List, Union

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

from md2slides.parser import ListItem, MarkdownParser, Slide, TextRun, ValidationError


class MarkdownToPptxConverter:
    """Convert markdown content to PowerPoint presentations."""

    def __init__(self) -> None:
        """Initialize the converter."""
        self._prs: Presentation | None = None

    def convert(self, markdown_content: str, output_path: str) -> str:
        """Convert markdown content to a PPTX file.

        Args:
            markdown_content: The markdown string to convert.
            output_path: Path where the PPTX file will be saved.

        Returns:
            The absolute path to the created PPTX file.

        Raises:
            ValidationError: If markdown content or output path is invalid.
        """
        self._validate_output_path(output_path)

        parser = MarkdownParser(markdown_content)
        slides = parser.parse()

        self._prs = Presentation()
        self._prs.slide_width = Inches(13.333)
        self._prs.slide_height = Inches(7.5)

        for slide_data in slides:
            if slide_data.is_title_slide:
                self._create_title_slide(slide_data)
            else:
                self._create_content_slide(slide_data)

        # Ensure output directory exists
        output_dir = os.path.dirname(output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)

        self._prs.save(output_path)
        return os.path.abspath(output_path)

    def _validate_output_path(self, output_path: str) -> None:
        """Validate the output path.

        Args:
            output_path: The path to validate.

        Raises:
            ValidationError: If the path is invalid.
        """
        if not isinstance(output_path, str):
            raise ValidationError(
                f"Output path must be a string, got {type(output_path).__name__}"
            )
        if not output_path.strip():
            raise ValidationError("Output path cannot be empty")
        if not output_path.lower().endswith(".pptx"):
            raise ValidationError("Output path must have .pptx extension")

    def _create_title_slide(self, slide_data: Slide) -> None:
        """Create a title slide.

        Args:
            slide_data: The slide data to render.
        """
        blank_layout = self._prs.slide_layouts[6]  # Blank layout
        slide = self._prs.slides.add_slide(blank_layout)

        # Title text box
        title_left = Inches(0.5)
        title_top = Inches(2.5)
        title_width = Inches(12.333)
        title_height = Inches(1.5)

        title_shape = slide.shapes.add_textbox(
            title_left, title_top, title_width, title_height
        )
        title_frame = title_shape.text_frame
        title_frame.word_wrap = True

        title_para = title_frame.paragraphs[0]
        title_para.alignment = PP_ALIGN.CENTER
        title_run = title_para.add_run()
        title_run.text = slide_data.title
        title_run.font.size = Pt(44)
        title_run.font.bold = True
        title_run.font.color.rgb = RGBColor(0, 0, 0)

        # Subtitle text box (if present)
        if slide_data.subtitle:
            subtitle_top = Inches(4.2)
            subtitle_height = Inches(1.0)

            subtitle_shape = slide.shapes.add_textbox(
                title_left, subtitle_top, title_width, subtitle_height
            )
            subtitle_frame = subtitle_shape.text_frame
            subtitle_frame.word_wrap = True

            subtitle_para = subtitle_frame.paragraphs[0]
            subtitle_para.alignment = PP_ALIGN.CENTER
            subtitle_run = subtitle_para.add_run()
            subtitle_run.text = slide_data.subtitle
            subtitle_run.font.size = Pt(24)
            subtitle_run.font.color.rgb = RGBColor(80, 80, 80)

    def _create_content_slide(self, slide_data: Slide) -> None:
        """Create a content slide.

        Args:
            slide_data: The slide data to render.
        """
        blank_layout = self._prs.slide_layouts[6]  # Blank layout
        slide = self._prs.slides.add_slide(blank_layout)

        # Title text box
        title_left = Inches(0.5)
        title_top = Inches(0.4)
        title_width = Inches(12.333)
        title_height = Inches(0.8)

        title_shape = slide.shapes.add_textbox(
            title_left, title_top, title_width, title_height
        )
        title_frame = title_shape.text_frame
        title_frame.word_wrap = True

        title_para = title_frame.paragraphs[0]
        title_para.alignment = PP_ALIGN.LEFT
        title_run = title_para.add_run()
        title_run.text = slide_data.title
        title_run.font.size = Pt(32)
        title_run.font.bold = True
        title_run.font.color.rgb = RGBColor(0, 0, 0)

        # Content text box
        content_left = Inches(0.5)
        content_top = Inches(1.4)
        content_width = Inches(12.333)
        content_height = Inches(5.6)

        content_shape = slide.shapes.add_textbox(
            content_left, content_top, content_width, content_height
        )
        content_frame = content_shape.text_frame
        content_frame.word_wrap = True

        self._render_content(content_frame, slide_data.content)

    def _render_content(
        self, text_frame, content: List[Union[ListItem, TextRun]]
    ) -> None:
        """Render content to a text frame.

        Args:
            text_frame: The PowerPoint text frame to render to.
            content: The list of content items to render.
        """
        first_item = True

        for item in content:
            if isinstance(item, ListItem):
                if first_item:
                    para = text_frame.paragraphs[0]
                    first_item = False
                else:
                    para = text_frame.add_paragraph()

                # Set indentation based on level
                para.level = item.level

                # Add bullet or number
                if item.ordered and item.number is not None:
                    # For ordered lists, we prepend the number
                    # (python-pptx doesn't have great numbered list support)
                    prefix_run = para.add_run()
                    prefix_run.text = f"{item.number}. "
                    prefix_run.font.size = Pt(18)
                else:
                    # Bullet point - use bullet character
                    para.bullet = True

                # Add content with formatting
                for text_run in item.content:
                    run = para.add_run()
                    run.text = text_run.text
                    run.font.size = Pt(18)
                    run.font.bold = text_run.bold
                    run.font.italic = text_run.italic

            elif isinstance(item, TextRun):
                if first_item:
                    para = text_frame.paragraphs[0]
                    first_item = False
                else:
                    para = text_frame.add_paragraph()

                run = para.add_run()
                run.text = item.text
                run.font.size = Pt(18)
                run.font.bold = item.bold
                run.font.italic = item.italic


def convert_file(input_path: str, output_path: str | None = None) -> str:
    """Convert a markdown file to PPTX.

    Args:
        input_path: Path to the markdown file.
        output_path: Optional path for the output PPTX. If not provided,
            uses the same name as input with .pptx extension.

    Returns:
        The absolute path to the created PPTX file.

    Raises:
        ValidationError: If the input file doesn't exist or is invalid.
        FileNotFoundError: If the input file doesn't exist.
    """
    if not isinstance(input_path, str):
        raise ValidationError(
            f"Input path must be a string, got {type(input_path).__name__}"
        )

    input_path = input_path.strip()
    if not input_path:
        raise ValidationError("Input path cannot be empty")

    path = Path(input_path)
    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    if not path.is_file():
        raise ValidationError(f"Input path is not a file: {input_path}")

    # Read markdown content
    content = path.read_text(encoding="utf-8")

    # Determine output path
    if output_path is None:
        output_path = str(path.with_suffix(".pptx"))

    converter = MarkdownToPptxConverter()
    return converter.convert(content, output_path)
