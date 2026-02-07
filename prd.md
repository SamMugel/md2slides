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
    "passes": true
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
    "passes": true
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
    "passes": true
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
    "passes": true
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
    "passes": true
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
    "passes": true
  },
{
  "category": "refactor",
  "description": "Move PPTX templates into a dedicated templates/ folder and update codepaths",
  "steps": [
    "Create top-level templates/ directory (or move existing template assets into it).",
    "Move any existing .pptx template files into templates/ (preserve filenames).",
    "Update any hardcoded paths or config to reference templates/ as the canonical template root.",
    "Add a small helper in utils.py (e.g., get_template_path(name)) used by converter/theme loading code.",
    "Verify: running the CLI with template usage still finds templates from templates/ without manual path edits."
  ],
  "passes": true
},
{
  "category": "feature",
  "description": "Add company logo to bottom-right corner of every slide from templates/multiverse_logo.png",
  "steps": [
    "Add a presentation-wide post-processing step in converter.py that iterates all slides after creation.",
    "Load templates/multiverse_logo.png once (cache image bytes/path) and reuse for all slides.",
    "Place logo at bottom-right with consistent margin (e.g., 0.2–0.3in from right/bottom) and a fixed max height (maintain aspect ratio).",
    "Ensure logo is behind no content overlaps: if footer area is used, reserve padding or adjust content boxes upward slightly.",
    "Verify: generate PPTX from an example markdown and confirm logo appears on every slide in the same position."
  ],
  "passes": true
},
{
  "category": "feature",
  "description": "When a slide has both text and one image: use a 2-column layout with autoscaled image on the right",
  "steps": [
    "Define a layout rule in elements/images.py (or a slide layout resolver) that detects slides containing both text blocks and exactly one image.",
    "Compute slide content bounds (account for title area, margins, and the bottom-right logo safe area).",
    "Split content bounds into left and right regions (e.g., 55% left text, 45% right image, with gutter).",
    "Render all text content into the left region (wrapping enabled; keep lists as lists).",
    "Autoscale the image to fit the right region preserving aspect ratio: scale by min(width_scale, height_scale) so it fully fits; center within the right region.",
    "Verify: create example markdown with text + image; confirm text is entirely left and image fills right region as much as possible without cropping."
  ],
  "passes": true
},
{
  "category": "feature",
  "description": "If two images are present on a slide, stack them vertically (top and bottom) with consistent spacing",
  "steps": [
    "Extend slide content detection to identify when a slide has exactly two images (and optional text rules if applicable; prefer image-only stacking when text absent).",
    "Compute an image region (full width inside margins; exclude title + logo safe area).",
    "Split the image region into two vertical slots with a gutter (e.g., 0.2in).",
    "Autoscale each image to fit its slot (preserve aspect ratio, center in slot).",
    "Define deterministic ordering (markdown order).",
    "Verify: example markdown with two images produces a slide with one image above the other, properly scaled and not overlapping."
  ],
  "passes": true
},
{
  "category": "feature",
  "description": "Render markdown hyperlinks as clickable links in PPTX",
  "steps": [
    "Ensure the markdown parser emits link tokens (text + href) for inline links [label](url).",
    "In elements/text.py, when rendering runs, create separate runs for linked spans and apply hyperlink relationship using python-pptx APIs.",
    "Preserve link label styling consistent with theme (optionally underline/blue if your theme defines it; otherwise keep minimal but clickable).",
    "Handle edge cases: multiple links in a paragraph, links inside list items, and trailing punctuation.",
    "Verify: example markdown with multiple hyperlinks results in clickable links in PowerPoint."
  ],
  "passes": true
},
{
  "category": "feature",
  "description": "Treat text immediately after an H1 as a subtitle for the section/title slide",
  "steps": [
    "Update parser.py section handling: when an H1 is followed by a paragraph (before any H2), attach that paragraph as section_subtitle.",
    "Update converter.py title/section slide creation to add a subtitle text box under the H1 title using theme subtitle styles.",
    "Ensure subtitle does not get rendered as a normal body paragraph elsewhere (consume it as subtitle).",
    "Verify: example markdown with H1 + paragraph + H2 renders a title slide with subtitle and subsequent slides unaffected."
  ],
  "passes": true
},
{
  "category": "feature",
  "description": "If text appears after an H2 and is not in a list, render it inline (not as a bullet list) in PPTX",
  "steps": [
    "In parser.py, distinguish paragraph blocks vs list blocks under each H2 slide.",
    "In converter/elements/text.py, ensure paragraphs are rendered as normal text frames (no bullet) and do not get coerced into list structures.",
    "Confirm that any paragraph following H2 remains a paragraph even if later lists appear on the same slide.",
    "Verify in the resulting PPTX XML/PowerPoint UI that the paragraph is not a bulleted list (no bullet indentation, no list formatting).",
    "Verify: example markdown H2 + paragraph + list produces one inline paragraph followed by bullets."
  ],
  "passes": true
},
{
  "category": "feature",
  "description": "Render H3 headings as bold, red, and inline with subsequent content",
  "steps": [
    "Update parser.py to keep H3 tokens as inline heading elements within a slide rather than creating new slides.",
    "In elements/text.py, render H3 as a run with bold enabled, font color red, and placed inline (same text box) with the following paragraph/list item if applicable.",
    "Define spacing rule: add a separator like ' — ' or ': ' between H3 and following text when in the same paragraph, unless user already includes punctuation.",
    "Ensure theme doesn’t override the red/bold requirement for H3 (explicitly set run properties).",
    "Verify: example markdown with H3 followed by text shows bold red inline heading within the body area."
  ],
  "passes": false
},
{
  "category": "testing",
  "description": "Add examples demonstrating new layout and markdown behavior; output generated PPTX examples into examples/ folder",
  "steps": [
    "Create new markdown files in examples/ covering: logo placement, text+image two-column, two-image stacking, hyperlinks, H1 subtitle, H2 inline paragraph, H3 inline styling.",
    "Add (or update) a simple script/Make target to build all examples into examples/*.pptx (or a dedicated examples/output/ if you prefer, but keep under examples/).",
    "Ensure the build script does not require external network access (use local images in templates/ or examples/assets/).",
    "Verify: running the script generates PPTX outputs into examples/ as required."
  ],
  "passes": false
},
{
  "category": "chore",
  "description": "Ensure PPTX files are not version controlled (ignore outputs and templates as appropriate)",
  "steps": [
    "Update .gitignore to include: *.pptx (and optionally examples/*.pptx if you want narrower scope).",
    "If you must keep template .pptx files, ignore only generated outputs and keep templates tracked; otherwise ignore templates/*.pptx too (choose consistent policy).",
    "Remove any already-tracked PPTX files from git index (git rm --cached ...) while keeping them locally.",
    "Verify: git status does not show generated PPTX files as untracked/modified after generation."
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
