# md2slides

Convert Markdown files to PowerPoint (PPTX) presentations.

## Installation

```bash
pip install -e .
```

For development with test dependencies:

```bash
pip install -e ".[dev]"
```

## Usage

### Command Line

```bash
# Convert with automatic output name (input.md -> input.pptx)
md2slides presentation.md

# Specify output path
md2slides presentation.md -o slides.pptx
```

### Python API

```python
from md2slides import MarkdownToPptxConverter
from md2slides.converter import convert_file

# Convert a file
convert_file("presentation.md", "output.pptx")

# Convert from string
converter = MarkdownToPptxConverter()
converter.convert("# Title\n\n## Slide 1\n\nContent", "output.pptx")
```

## Markdown Format

### Document Structure

| Element | Markdown | Result |
|---------|----------|--------|
| Title slide | `# Heading` | Creates title slide with heading as title |
| Subtitle | Text after `#` | Becomes subtitle on title slide |
| Content slide | `## Heading` | Creates new content slide |
| Slide content | Text/lists after `##` | Becomes slide body |

### Supported Formatting

- **Bold**: `**text**` or `__text__`
- *Italic*: `*text*` or `_text_`
- ***Bold italic***: `***text***` or `___text___`
- Bullet lists with `-`, `*`, or `+`
- Numbered lists with `1.` or `1)`
- Nested lists with indentation

### Example

```markdown
# Quarterly Report

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
```

This produces:

1. **Title Slide**: "Quarterly Report" with subtitle "Q3 2024 Financial Summary"
2. **Content Slide**: "Revenue Overview" with bullet list (nested items preserved)
3. **Content Slide**: "Next Steps" with bullet list (bold formatting preserved)

## API Reference

### `MarkdownParser`

Parses markdown content into slide structures.

```python
from md2slides.parser import MarkdownParser

parser = MarkdownParser(markdown_content)
slides = parser.parse()
```

**Raises:**
- `ValidationError`: If content is empty, not a string, or has no headings

### `MarkdownToPptxConverter`

Converts parsed markdown to PPTX.

```python
from md2slides.converter import MarkdownToPptxConverter

converter = MarkdownToPptxConverter()
output_path = converter.convert(markdown_content, "output.pptx")
```

**Parameters:**
- `markdown_content`: Markdown string to convert
- `output_path`: Path for the output PPTX file (must end in `.pptx`)

**Returns:** Absolute path to created file

**Raises:**
- `ValidationError`: If output path is invalid or content is invalid

### `convert_file`

Convenience function to convert a markdown file.

```python
from md2slides.converter import convert_file

# With explicit output path
convert_file("input.md", "output.pptx")

# With auto-generated output path
convert_file("input.md")  # Creates input.pptx
```

**Parameters:**
- `input_path`: Path to the markdown file
- `output_path`: Optional output path (defaults to input with `.pptx` extension)

**Returns:** Absolute path to created file

**Raises:**
- `FileNotFoundError`: If input file doesn't exist
- `ValidationError`: If paths are invalid

## Data Classes

### `Slide`

Represents a single slide.

```python
@dataclass
class Slide:
    title: str
    content: List[ListItem | TextRun]
    is_title_slide: bool = False
    subtitle: Optional[str] = None
```

### `ListItem`

Represents a list item with optional nesting.

```python
@dataclass
class ListItem:
    content: List[TextRun]
    level: int = 0
    ordered: bool = False
    number: Optional[int] = None
```

### `TextRun`

A run of text with formatting.

```python
@dataclass
class TextRun:
    text: str
    bold: bool = False
    italic: bool = False
```

## Running Tests

```bash
# Install dev dependencies
pip install -e ".[dev]"

# Run tests with coverage
pytest

# View HTML coverage report
open htmlcov/index.html
```

## License

MIT
