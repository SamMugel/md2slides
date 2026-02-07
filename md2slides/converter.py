"""Core conversion logic for md2slides."""

from pathlib import Path
from typing import Optional

from pptx import Presentation as PptxPresentation
from pptx.util import Inches, Pt

from md2slides.parser import (
    BulletList,
    CodeBlock,
    Image,
    Presentation,
    Section,
    Slide,
    Table,
    TextElement,
    parse_markdown,
)
from md2slides.themes import get_theme


class PresentationConverter:
    """Convert a parsed Presentation to PowerPoint format."""

    # Slide layout indices for the default template
    LAYOUT_TITLE = 0        # Title Slide
    LAYOUT_CONTENT = 1      # Title and Content
    LAYOUT_SECTION = 2      # Section Header
    LAYOUT_BLANK = 6        # Blank

    def __init__(self, theme: str = "default"):
        self.theme_name = theme
        self._theme = get_theme(theme)
        self._prs: Optional[PptxPresentation] = None
        self._input_dir: Optional[Path] = None  # For resolving relative image paths

    def convert(self, presentation: Presentation) -> PptxPresentation:
        """Convert a parsed Presentation to a PowerPoint presentation."""
        self._prs = PptxPresentation()

        # Create title slide from presentation title
        if presentation.title:
            self._add_title_slide(presentation.title)

        # Process each section
        for section in presentation.sections:
            self._add_section(section)

        return self._prs

    def convert_markdown(self, markdown_text: str) -> PptxPresentation:
        """Parse markdown and convert to PowerPoint."""
        presentation = parse_markdown(markdown_text)
        return self.convert(presentation)

    def convert_file(self, input_path: str, output_path: str) -> None:
        """Convert a markdown file to a PowerPoint file."""
        input_path = Path(input_path)
        output_path = Path(output_path)

        # Store input directory for resolving relative image paths
        self._input_dir = input_path.parent

        with open(input_path, "r", encoding="utf-8") as f:
            markdown_text = f.read()

        pptx = self.convert_markdown(markdown_text)
        pptx.save(str(output_path))

    def _add_title_slide(self, title: str, subtitle: str = "") -> None:
        """Add a title slide to the presentation."""
        slide_layout = self._prs.slide_layouts[self.LAYOUT_TITLE]
        slide = self._prs.slides.add_slide(slide_layout)

        title_shape = slide.shapes.title
        if title_shape:
            title_shape.text = title

        # Add subtitle if there's a placeholder for it
        if subtitle and len(slide.placeholders) > 1:
            subtitle_shape = slide.placeholders[1]
            subtitle_shape.text = subtitle

        # Apply theme styling
        self._theme.apply_to_title_slide(slide)

    def _add_section(self, section: Section) -> None:
        """Add a section's slides to the presentation."""
        for slide in section.slides:
            self._add_content_slide(slide)

    def _add_content_slide(self, slide: Slide) -> None:
        """Add a content slide to the presentation."""
        slide_layout = self._prs.slide_layouts[self.LAYOUT_CONTENT]
        pptx_slide = self._prs.slides.add_slide(slide_layout)

        # Set the title
        title_shape = pptx_slide.shapes.title
        if title_shape:
            title_shape.text = slide.title

        # Find the content placeholder
        content_placeholder = None
        for shape in pptx_slide.placeholders:
            if shape.placeholder_format.idx == 1:  # Content placeholder
                content_placeholder = shape
                break

        if content_placeholder is None:
            return

        # Process slide elements
        text_frame = content_placeholder.text_frame
        text_frame.clear()  # Clear default text

        first_element = True
        for element in slide.elements:
            if isinstance(element, TextElement):
                self._add_text_element(text_frame, element, first_element)
                first_element = False

            elif isinstance(element, BulletList):
                self._add_bullet_list(text_frame, element, first_element)
                first_element = False

            elif isinstance(element, CodeBlock):
                self._add_code_block(text_frame, element, first_element)
                first_element = False

            elif isinstance(element, Table):
                # Tables need special handling - add as a shape
                self._add_table(pptx_slide, element)

            elif isinstance(element, Image):
                # Images are added as separate shapes
                self._add_image(pptx_slide, element)

        # Apply theme styling
        self._theme.apply_to_content_slide(pptx_slide)

    def _add_text_element(
        self, text_frame, element: TextElement, is_first: bool
    ) -> None:
        """Add a text element to the text frame."""
        if is_first:
            p = text_frame.paragraphs[0]
        else:
            p = text_frame.add_paragraph()

        p.text = element.text
        p.level = 0

    def _add_bullet_list(
        self, text_frame, bullet_list: BulletList, is_first: bool
    ) -> None:
        """Add a bullet list to the text frame."""
        for i, item in enumerate(bullet_list.items):
            if is_first and i == 0:
                p = text_frame.paragraphs[0]
            else:
                p = text_frame.add_paragraph()

            p.text = item.text
            p.level = 0

            # Add nested items
            for child in item.children:
                p = text_frame.add_paragraph()
                p.text = child.text
                p.level = 1

                # Handle deeper nesting (grandchildren)
                for grandchild in child.children:
                    p = text_frame.add_paragraph()
                    p.text = grandchild.text
                    p.level = 2

    def _add_code_block(
        self, text_frame, code_block: CodeBlock, is_first: bool
    ) -> None:
        """Add a code block to the text frame."""
        lines = code_block.code.strip().split("\n")

        for i, line in enumerate(lines):
            if is_first and i == 0:
                p = text_frame.paragraphs[0]
            else:
                p = text_frame.add_paragraph()

            p.text = line
            p.level = 0

            # Use monospace font for code
            for run in p.runs:
                run.font.name = "Courier New"
                run.font.size = Pt(10)

    def _add_table(self, pptx_slide, table: Table) -> None:
        """Add a table to the slide as a separate shape."""
        if not table.headers and not table.rows:
            return

        rows = len(table.rows) + (1 if table.headers else 0)
        cols = len(table.headers) if table.headers else len(table.rows[0])

        # Position the table below the content area
        left = Inches(1)
        top = Inches(4)
        width = Inches(8)
        height = Inches(0.5 * rows)

        tbl = pptx_slide.shapes.add_table(rows, cols, left, top, width, height).table

        # Add headers
        if table.headers:
            for col_idx, header in enumerate(table.headers):
                cell = tbl.cell(0, col_idx)
                cell.text = header

        # Add data rows
        start_row = 1 if table.headers else 0
        for row_idx, row_data in enumerate(table.rows):
            for col_idx, cell_text in enumerate(row_data):
                cell = tbl.cell(start_row + row_idx, col_idx)
                cell.text = cell_text

    def _add_image(self, pptx_slide, image: Image) -> None:
        """Add an image to the slide."""
        image_path = self._resolve_image_path(image.src)

        if image_path is None:
            # Image not found, skip it
            return

        # Position and size the image
        left = Inches(2)
        top = Inches(2)
        # Let python-pptx auto-calculate width to maintain aspect ratio
        # Max width is 6 inches to leave room on the slide
        max_width = Inches(6)

        try:
            pptx_slide.shapes.add_picture(str(image_path), left, top, width=max_width)
        except Exception:
            # If image fails to load, skip it silently
            pass

    def _resolve_image_path(self, src: str) -> Optional[Path]:
        """Resolve an image path (local or URL)."""
        # Handle URLs - not supported yet, return None
        if src.startswith(("http://", "https://")):
            # TODO: Download remote images
            return None

        # Handle local paths
        path = Path(src)

        # If absolute path, use as-is
        if path.is_absolute():
            if path.exists():
                return path
            return None

        # Relative path - resolve relative to input file directory
        if self._input_dir:
            resolved = self._input_dir / path
            if resolved.exists():
                return resolved

        # Try relative to current directory
        if path.exists():
            return path

        return None


def convert_markdown_to_pptx(
    markdown_text: str,
    output_path: Optional[str] = None,
    theme: str = "default"
) -> PptxPresentation:
    """Convenience function to convert markdown to PowerPoint."""
    converter = PresentationConverter(theme=theme)
    pptx = converter.convert_markdown(markdown_text)

    if output_path:
        pptx.save(output_path)

    return pptx
