"""Core conversion logic for md2slides."""

from pathlib import Path
from typing import Optional

from pptx import Presentation as PptxPresentation
from pptx.util import Inches, Pt

from md2slides.parser import (
    BulletList,
    CodeBlock,
    Image,
    InlineSpan,
    ListItem,
    Presentation,
    Section,
    Slide,
    Table,
    TextElement,
    parse_markdown,
)
from md2slides.themes import get_theme
from md2slides.utils import get_template_path


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

        # Create title slide from presentation title (with subtitle from first section)
        if presentation.title:
            # Get subtitle from the first section (if it's the title section)
            subtitle = ""
            if presentation.sections and presentation.sections[0].title == presentation.title:
                subtitle = presentation.sections[0].subtitle
            self._add_title_slide(presentation.title, subtitle)

        # Process each section
        for section in presentation.sections:
            self._add_section(section)

        # Post-processing: add logo to all slides
        self._add_logo_to_all_slides()

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
        # Check if this slide qualifies for 2-column layout (text + exactly 1 image)
        images = [e for e in slide.elements if isinstance(e, Image)]
        text_elements = [e for e in slide.elements if isinstance(e, (TextElement, BulletList, CodeBlock))]
        tables = [e for e in slide.elements if isinstance(e, Table)]

        if len(images) == 1 and len(text_elements) > 0 and len(tables) == 0:
            # Use 2-column layout: text on left, image on right
            self._add_two_column_slide(slide, text_elements, images[0])
            return

        if len(images) == 2 and len(text_elements) == 0 and len(tables) == 0:
            # Use vertically stacked layout for two images
            self._add_two_image_stacked_slide(slide, images)
            return

        # Default layout for other slides
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

    def _add_two_column_slide(
        self,
        slide: Slide,
        text_elements: list,
        image: Image
    ) -> None:
        """Add a slide with 2-column layout: text on left, image on right."""
        from pptx.util import Emu
        from pptx.enum.text import MSO_ANCHOR

        # Use blank layout for custom positioning
        slide_layout = self._prs.slide_layouts[self.LAYOUT_BLANK]
        pptx_slide = self._prs.slides.add_slide(slide_layout)

        # Slide dimensions (standard 10" x 7.5")
        slide_width = Inches(10)
        slide_height = Inches(7.5)

        # Content area margins
        margin_left = Inches(0.5)
        margin_right = Inches(0.5)
        margin_top = Inches(1.5)  # Space for title
        margin_bottom = Inches(0.75)  # Space for logo

        # Title area
        title_left = margin_left
        title_top = Inches(0.3)
        title_width = slide_width - margin_left - margin_right
        title_height = Inches(1.0)

        # Add title text box
        title_box = pptx_slide.shapes.add_textbox(
            title_left, title_top, title_width, title_height
        )
        title_tf = title_box.text_frame
        title_tf.word_wrap = True
        p = title_tf.paragraphs[0]
        p.text = slide.title
        p.font.size = Pt(32)
        p.font.bold = True

        # Content area dimensions
        content_top = margin_top
        content_bottom = slide_height - margin_bottom
        content_height = content_bottom - content_top
        content_width = slide_width - margin_left - margin_right

        # 2-column split: 55% left (text), 45% right (image), with gutter
        gutter = Inches(0.3)
        left_width = content_width * 0.55 - gutter / 2
        right_width = content_width * 0.45 - gutter / 2
        left_left = margin_left
        right_left = margin_left + left_width + gutter

        # Add text content on the left
        text_box = pptx_slide.shapes.add_textbox(
            left_left, content_top, left_width, content_height
        )
        text_tf = text_box.text_frame
        text_tf.word_wrap = True
        text_tf.auto_size = None  # Disable auto-resize

        first_element = True
        for element in text_elements:
            if isinstance(element, TextElement):
                self._add_text_element(text_tf, element, first_element)
                first_element = False
            elif isinstance(element, BulletList):
                self._add_bullet_list(text_tf, element, first_element)
                first_element = False
            elif isinstance(element, CodeBlock):
                self._add_code_block(text_tf, element, first_element)
                first_element = False

        # Add image on the right, scaled to fit
        image_path = self._resolve_image_path(image.src)
        if image_path is not None:
            self._add_scaled_image(
                pptx_slide,
                image_path,
                left=right_left,
                top=content_top,
                max_width=right_width,
                max_height=content_height
            )

        # Apply theme styling
        self._theme.apply_to_content_slide(pptx_slide)

    def _add_two_image_stacked_slide(
        self,
        slide: Slide,
        images: list
    ) -> None:
        """Add a slide with two images stacked vertically (top and bottom)."""
        # Use blank layout for custom positioning
        slide_layout = self._prs.slide_layouts[self.LAYOUT_BLANK]
        pptx_slide = self._prs.slides.add_slide(slide_layout)

        # Slide dimensions (standard 10" x 7.5")
        slide_width = Inches(10)
        slide_height = Inches(7.5)

        # Content area margins
        margin_left = Inches(0.5)
        margin_right = Inches(0.5)
        margin_top = Inches(1.5)  # Space for title
        margin_bottom = Inches(0.75)  # Space for logo

        # Title area
        title_left = margin_left
        title_top = Inches(0.3)
        title_width = slide_width - margin_left - margin_right
        title_height = Inches(1.0)

        # Add title text box
        title_box = pptx_slide.shapes.add_textbox(
            title_left, title_top, title_width, title_height
        )
        title_tf = title_box.text_frame
        title_tf.word_wrap = True
        p = title_tf.paragraphs[0]
        p.text = slide.title
        p.font.size = Pt(32)
        p.font.bold = True

        # Image region (full width, excluding title and logo areas)
        content_left = margin_left
        content_top = margin_top
        content_width = slide_width - margin_left - margin_right
        content_bottom = slide_height - margin_bottom
        content_height = content_bottom - content_top

        # Split into two vertical slots with gutter
        gutter = Inches(0.2)
        slot_height = (content_height - gutter) / 2

        # Top slot
        top_slot_top = content_top
        top_slot_height = slot_height

        # Bottom slot
        bottom_slot_top = content_top + slot_height + gutter
        bottom_slot_height = slot_height

        # Add first image (top)
        image1_path = self._resolve_image_path(images[0].src)
        if image1_path is not None:
            self._add_scaled_image(
                pptx_slide,
                image1_path,
                left=content_left,
                top=top_slot_top,
                max_width=content_width,
                max_height=top_slot_height
            )

        # Add second image (bottom)
        image2_path = self._resolve_image_path(images[1].src)
        if image2_path is not None:
            self._add_scaled_image(
                pptx_slide,
                image2_path,
                left=content_left,
                top=bottom_slot_top,
                max_width=content_width,
                max_height=bottom_slot_height
            )

        # Apply theme styling
        self._theme.apply_to_content_slide(pptx_slide)

    def _add_scaled_image(
        self,
        pptx_slide,
        image_path,
        left,
        top,
        max_width,
        max_height
    ) -> None:
        """Add an image scaled to fit within the given bounds, preserving aspect ratio."""
        from PIL import Image as PILImage

        try:
            # Get original image dimensions
            with PILImage.open(image_path) as img:
                orig_width, orig_height = img.size

            # Calculate scale factors
            width_scale = max_width / Inches(orig_width / 96)  # Assume 96 DPI
            height_scale = max_height / Inches(orig_height / 96)

            # Use the smaller scale to fit within bounds
            scale = min(width_scale, height_scale, 1.0)  # Don't upscale beyond original

            # Calculate final dimensions
            final_width = Inches(orig_width / 96) * scale
            final_height = Inches(orig_height / 96) * scale

            # Center the image within the right region
            centered_left = left + (max_width - final_width) / 2
            centered_top = top + (max_height - final_height) / 2

            pptx_slide.shapes.add_picture(
                str(image_path),
                int(centered_left),
                int(centered_top),
                width=int(final_width)
            )
        except Exception:
            # Fallback: add image with max width constraint
            try:
                pptx_slide.shapes.add_picture(
                    str(image_path),
                    int(left),
                    int(top),
                    width=int(max_width)
                )
            except Exception:
                pass

    def _add_text_element(
        self, text_frame, element: TextElement, is_first: bool
    ) -> None:
        """Add a text element to the text frame."""
        if is_first:
            p = text_frame.paragraphs[0]
        else:
            p = text_frame.add_paragraph()

        p.level = 0

        # If we have spans with rich content (hyperlinks), render those
        if element.spans:
            self._add_spans_to_paragraph(p, element.spans)
        else:
            # Fallback to plain text
            p.text = element.text

    def _add_spans_to_paragraph(self, paragraph, spans: list) -> None:
        """Add inline spans to a paragraph, supporting hyperlinks."""
        from pptx.dml.color import RGBColor

        # Clear any existing text
        paragraph.clear()

        for span in spans:
            run = paragraph.add_run()
            run.text = span.text

            # Apply hyperlink if present
            if span.href:
                run.hyperlink.address = span.href
                # Optionally apply link styling (blue, underline)
                run.font.color.rgb = RGBColor(0x05, 0x63, 0xC1)  # A standard link blue
                run.font.underline = True

            # Apply other formatting
            if span.bold:
                run.font.bold = True
            if span.italic:
                run.font.italic = True
            if span.code:
                run.font.name = self._theme.fonts.code_font
                run.font.size = Pt(self._theme.fonts.code_size)

    def _add_bullet_list(
        self, text_frame, bullet_list: BulletList, is_first: bool
    ) -> None:
        """Add a bullet list to the text frame."""
        for i, item in enumerate(bullet_list.items):
            if is_first and i == 0:
                p = text_frame.paragraphs[0]
            else:
                p = text_frame.add_paragraph()

            p.level = 0
            self._add_list_item_content(p, item)

            # Add nested items
            for child in item.children:
                p = text_frame.add_paragraph()
                p.level = 1
                self._add_list_item_content(p, child)

                # Handle deeper nesting (grandchildren)
                for grandchild in child.children:
                    p = text_frame.add_paragraph()
                    p.level = 2
                    self._add_list_item_content(p, grandchild)

    def _add_list_item_content(self, paragraph, item: ListItem) -> None:
        """Add a list item's content to a paragraph, supporting hyperlinks."""
        if item.spans:
            self._add_spans_to_paragraph(paragraph, item.spans)
        else:
            paragraph.text = item.text

    def _add_code_block(
        self, text_frame, code_block: CodeBlock, is_first: bool
    ) -> None:
        """Add a code block to the text frame with monospace styling."""
        from pptx.dml.color import RGBColor

        lines = code_block.code.strip().split("\n")
        code_font = self._theme.fonts.code_font
        code_size = Pt(self._theme.fonts.code_size)
        code_color = RGBColor(*self._theme.colors.body_text)

        for i, line in enumerate(lines):
            if is_first and i == 0:
                p = text_frame.paragraphs[0]
            else:
                p = text_frame.add_paragraph()

            # Use add_run to ensure we can style the text
            run = p.add_run()
            run.text = line
            run.font.name = code_font
            run.font.size = code_size
            run.font.color.rgb = code_color
            p.level = 0

    def _add_table(self, pptx_slide, table: Table) -> None:
        """Add a table to the slide as a separate shape."""
        from pptx.dml.color import RGBColor
        from pptx.enum.text import PP_ALIGN

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

        # Map alignment strings to PowerPoint alignment
        alignment_map = {
            "left": PP_ALIGN.LEFT,
            "center": PP_ALIGN.CENTER,
            "right": PP_ALIGN.RIGHT,
            None: PP_ALIGN.LEFT,
        }

        # Get theme colors for styling
        header_color = RGBColor(*self._theme.colors.title_text)
        body_color = RGBColor(*self._theme.colors.body_text)
        body_font = self._theme.fonts.body_font
        body_size = Pt(self._theme.fonts.body_size - 4)  # Slightly smaller for tables

        # Add headers
        if table.headers:
            for col_idx, header in enumerate(table.headers):
                cell = tbl.cell(0, col_idx)
                cell.text = header

                # Style header cell
                para = cell.text_frame.paragraphs[0]
                para.font.name = body_font
                para.font.size = body_size
                para.font.bold = True
                para.font.color.rgb = header_color

                # Apply alignment
                if col_idx < len(table.alignments):
                    para.alignment = alignment_map.get(table.alignments[col_idx], PP_ALIGN.LEFT)

        # Add data rows
        start_row = 1 if table.headers else 0
        for row_idx, row_data in enumerate(table.rows):
            for col_idx, cell_text in enumerate(row_data):
                cell = tbl.cell(start_row + row_idx, col_idx)
                cell.text = cell_text

                # Style data cell
                para = cell.text_frame.paragraphs[0]
                para.font.name = body_font
                para.font.size = body_size
                para.font.color.rgb = body_color

                # Apply alignment
                if col_idx < len(table.alignments):
                    para.alignment = alignment_map.get(table.alignments[col_idx], PP_ALIGN.LEFT)

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

    def _add_logo_to_all_slides(self) -> None:
        """Add the company logo to the bottom-right corner of every slide."""
        try:
            logo_path = get_template_path("multiverse_logo.png")
        except FileNotFoundError:
            # Logo not available, skip silently
            return

        # Logo dimensions and positioning
        # Standard slide is 10 inches wide by 7.5 inches tall
        logo_max_height = Inches(0.5)  # Fixed max height for logo
        margin_right = Inches(0.25)    # Margin from right edge
        margin_bottom = Inches(0.25)   # Margin from bottom edge

        # Slide dimensions (standard 10x7.5)
        slide_width = Inches(10)
        slide_height = Inches(7.5)

        for slide in self._prs.slides:
            try:
                # Add the logo with height constraint (width auto-calculated to maintain aspect ratio)
                picture = slide.shapes.add_picture(
                    str(logo_path),
                    left=Inches(0),  # Temporary position
                    top=Inches(0),   # Temporary position
                    height=logo_max_height
                )

                # Calculate actual position for bottom-right corner
                # After adding, we know the actual width from the aspect ratio
                picture.left = int(slide_width - picture.width - margin_right)
                picture.top = int(slide_height - picture.height - margin_bottom)
            except Exception:
                # If logo fails to load on any slide, continue with others
                pass


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
