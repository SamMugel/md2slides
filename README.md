# md2slides

Convert Markdown files to PowerPoint presentations.

## Overview

md2slides is a Python CLI tool that converts Markdown files to PPTX presentations. It parses your markdown content and generates professional-looking slides with support for:

- Headings (H1 for sections/title, H2 for slide titles)
- Bullet points with nested lists
- Code blocks with monospace font
- Tables with alignment
- Images (local files)
- Multiple themes

## Installation

### Requirements

- Python 3.9+
- python-pptx
- mistune

### Install from source

```bash
git clone https://github.com/your-repo/md2slides.git
cd md2slides
pip install -e .
```

## Usage

### Basic conversion

```bash
md2slides input.md -o output.pptx
```

### With a theme

```bash
md2slides input.md -o output.pptx --theme dark
```

### List available themes

```bash
md2slides --list-themes
```

Output:
```
Available themes:
  - default
  - dark
  - light
  - corporate
```

### Help

```bash
md2slides --help
```

## Supported Markdown Syntax

### Headings

- `# H1` - Creates the presentation title (first H1) and sections
- `## H2` - Creates a new slide

### Text and Lists

```markdown
## My Slide

This is a paragraph.

- Bullet point
- Another point
  - Nested item
  - Another nested item
```

### Code Blocks

````markdown
## Code Example

```python
def hello():
    print("Hello, World!")
```
````

### Tables

```markdown
## Data Table

| Name | Value |
|------|-------|
| Alpha | 100 |
| Beta | 200 |
```

Alignment is supported:

```markdown
| Left | Center | Right |
|:-----|:------:|------:|
| L | C | R |
```

### Images

```markdown
## Slide with Image

![Alt text](path/to/image.png)
```

## Themes

### Default
Clean, professional look with blue titles and gray body text.

### Dark
Dark background with light text, ideal for presentations in dim rooms.

### Light
Minimal white background with subtle gray accents.

### Corporate
Navy blue titles with Georgia font, professional and formal.

## Examples

See the `examples/` directory for sample markdown files:

- `simple.md` - Basic presentation with headings and lists
- `full-featured.md` - All features: code, tables, text
- `theme-showcase.md` - Demonstrates all themes
- `code-test.md` - Code block examples
- `table-test.md` - Table examples

Convert an example:

```bash
md2slides examples/simple.md -o simple.pptx
```

## Project Structure

```
md2slides/
├── __init__.py      # Package init with version
├── cli.py           # CLI entry point
├── parser.py        # Markdown parsing logic
├── converter.py     # PPTX generation
└── themes/          # Theme definitions
    ├── __init__.py
    ├── base.py
    ├── default.py
    ├── dark.py
    ├── light.py
    └── corporate.py
```

## License

MIT License
