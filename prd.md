# md2slides - Product Requirements Document

## Overview
md2slides is a Python CLI tool that converts Markdown files to PowerPoint (PPTX) presentations. It solves the problem of users needing to quickly create professional presentations from structured text content without manually designing slides.

## Target Audience
General users who need to convert markdown files to PPTX presentations. These users want a simple, reliable tool that produces good-looking slides without requiring PowerPoint expertise or design skills.

## Core Features

1. **Basic Markdown to PPTX Conversion**
   - Parse markdown files and convert to PowerPoint slides
   - H1 headings define sections, H2 headings create new slides
   - Body content (paragraphs, lists) becomes slide content

2. **Built-in Theme Support**
   - Pre-built themes: default, dark, light, corporate
   - Consistent color schemes, fonts, and styling
   - Theme selection via CLI flag

3. **Image Handling**
   - Embed images referenced in markdown into slides
   - Support local file paths and URLs
   - Automatic image sizing and positioning

4. **Code Syntax Highlighting**
   - Render fenced code blocks with syntax highlighting
   - Preserve code formatting in slides
   - Readable monospace fonts

5. **Table Support**
   - Convert markdown tables to PowerPoint tables
   - Proper cell alignment and borders
   - Styled to match selected theme

## Tech Stack
- **Language**: Python 3.9+
- **Core Library**: python-pptx (PowerPoint generation)
- **Markdown Parsing**: markdown-it-py or mistune
- **Syntax Highlighting**: Pygments
- **CLI Framework**: Click or argparse
- **Package Manager**: pip/pyproject.toml

## Architecture
Simple CLI application with the following structure:
```
md2slides/
├── __init__.py
├── cli.py           # CLI entry point and argument parsing
├── parser.py        # Markdown parsing logic
├── converter.py     # Core conversion logic
├── themes/          # Built-in theme definitions
│   ├── __init__.py
│   ├── default.py
│   ├── dark.py
│   ├── light.py
│   └── corporate.py
├── elements/        # Slide element handlers
│   ├── __init__.py
│   ├── text.py
│   ├── images.py
│   ├── code.py
│   └── tables.py
└── utils.py         # Helper functions
```

## Data Model
- **Presentation**: Collection of slides with theme settings
- **Slide**: Title, content elements, layout type
- **Element**: Text block, image, code block, or table
- **Theme**: Colors, fonts, spacing definitions

## CLI Interface
```bash
# Basic usage
md2slides input.md -o output.pptx

# With theme
md2slides input.md -o output.pptx --theme dark

# List available themes
md2slides --list-themes

# Help
md2slides --help
```

## Security Considerations
- Validate file paths to prevent path traversal
- Sanitize image URLs before fetching
- No execution of code blocks - display only

## Constraints & Assumptions
- Python 3.9+ required
- python-pptx library is stable and well-maintained
- Users have basic command-line familiarity
- Output format is .pptx only (no PDF, Google Slides, etc.)

## Success Criteria
- All features working: themes, images, code blocks, tables
- Complete with README, examples, and usage documentation
- Clean, maintainable codebase
- Passes basic test cases

---

## Task List

```json
[
  {
    "category": "setup",
    "description": "Initialize Python project structure",
    "steps": [
      "Create pyproject.toml with project metadata and dependencies",
      "Create md2slides/ package directory with __init__.py",
      "Create basic cli.py with Click framework and --help command",
      "Verify: running 'python -m md2slides --help' shows usage"
    ],
    "passes": true
  },
  {
    "category": "setup",
    "description": "Install core dependencies",
    "steps": [
      "Add python-pptx, click, markdown-it-py, pygments to dependencies",
      "Create requirements.txt for development",
      "Install dependencies with pip install -e .",
      "Verify: imports work in Python REPL"
    ],
    "passes": true
  },
  {
    "category": "feature",
    "description": "Implement markdown parser",
    "steps": [
      "Create parser.py with MarkdownParser class",
      "Parse H1 as section markers, H2 as slide boundaries",
      "Extract paragraphs, bullet lists, and nested lists",
      "Return structured representation of slides",
      "Verify: parse test markdown file and print structure"
    ],
    "passes": true
  },
  {
    "category": "feature",
    "description": "Implement basic PPTX converter",
    "steps": [
      "Create converter.py with PresentationConverter class",
      "Create title slide from first H1",
      "Create content slides from H2 headings",
      "Add body text and bullet points to slides",
      "Verify: generate basic PPTX from test markdown"
    ],
    "passes": true
  },
  {
    "category": "feature",
    "description": "Wire up CLI to parser and converter",
    "steps": [
      "Add input file and -o/--output arguments to CLI",
      "Read markdown file, parse, convert, save PPTX",
      "Add error handling for missing files",
      "Verify: md2slides test.md -o test.pptx creates valid file"
    ],
    "passes": true
  },
  {
    "category": "feature",
    "description": "Implement default theme",
    "steps": [
      "Create themes/ directory with __init__.py",
      "Create default.py with color scheme, fonts, sizes",
      "Apply theme colors to slide backgrounds and text",
      "Apply theme fonts to headings and body",
      "Verify: generated PPTX uses consistent styling"
    ],
    "passes": true
  },
  {
    "category": "feature",
    "description": "Add dark, light, and corporate themes",
    "steps": [
      "Create dark.py theme (dark background, light text)",
      "Create light.py theme (white background, minimal colors)",
      "Create corporate.py theme (professional blue palette)",
      "Add --theme flag to CLI with choices",
      "Add --list-themes flag to show available themes",
      "Verify: each theme produces visually distinct output"
    ],
    "passes": false
  },
  {
    "category": "feature",
    "description": "Implement image handling",
    "steps": [
      "Detect image references in markdown (![alt](path))",
      "Load local images and embed in slides",
      "Download remote images from URLs",
      "Position and size images appropriately",
      "Verify: markdown with images produces slides with embedded images"
    ],
    "passes": false
  },
  {
    "category": "feature",
    "description": "Implement code block handling",
    "steps": [
      "Detect fenced code blocks with language hints",
      "Use Pygments to syntax highlight code",
      "Render highlighted code as formatted text box",
      "Use monospace font and appropriate sizing",
      "Verify: code blocks appear with syntax highlighting in slides"
    ],
    "passes": false
  },
  {
    "category": "feature",
    "description": "Implement table support",
    "steps": [
      "Detect markdown tables in parsed content",
      "Create PowerPoint table with correct rows/columns",
      "Apply theme styling to table borders and cells",
      "Handle cell alignment (left, center, right)",
      "Verify: markdown tables render as styled PPTX tables"
    ],
    "passes": false
  },
  {
    "category": "testing",
    "description": "Create test markdown files",
    "steps": [
      "Create examples/ directory",
      "Create simple.md with basic headings and text",
      "Create full-featured.md with images, code, tables",
      "Create theme-showcase.md to demonstrate all themes",
      "Verify: all example files convert without errors"
    ],
    "passes": false
  },
  {
    "category": "documentation",
    "description": "Write README and documentation",
    "steps": [
      "Create README.md with project description",
      "Document installation instructions",
      "Document CLI usage with examples",
      "Document supported markdown syntax",
      "Add screenshots of example outputs",
      "Verify: README is clear and complete"
    ],
    "passes": false
  }
]
```

---

## Agent Instructions

1. Read `activity.md` first to understand current state
2. Find next task with `"passes": false`
3. Complete all steps for that task
4. Verify the task works as described
5. Update task to `"passes": true`
6. Log completion in `activity.md`
7. Repeat until all tasks pass

**Important:** Only modify the `passes` field. Do not remove or rewrite tasks.

---

## Completion Criteria
All tasks marked with `"passes": true`
