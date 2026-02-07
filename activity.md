# md2slides - Activity Log

## Current Status
**Last Updated:** 2026-02-07
**Tasks Completed:** 13
**Current Task:** Move PPTX templates into dedicated templates/ folder

---

## Session Log

<!--
The Ralph Wiggum loop will append dated entries here.
Each entry should include:
- Date and time
- Task worked on
- Changes made
- Commands run
- Screenshot filename (if applicable)
- Any issues and resolutions
-->

### 2026-02-07 - Tasks 1 & 2: Project Setup and Dependencies

**Tasks completed:**
1. Initialize Python project structure
2. Install core dependencies

**Changes made:**
- Created `md2slides/` package directory
- Created `md2slides/__init__.py` with version info
- Created `md2slides/__main__.py` for `python -m md2slides` support
- Created `md2slides/cli.py` with argparse-based CLI (adapted from Click due to network issues)
- Updated `pyproject.toml` to use `mistune` (already installed) instead of `markdown-it-py`

**Commands run:**
```bash
python -m py_compile md2slides/*.py  # Syntax check passed
python -m md2slides --help           # Shows usage
python -m md2slides --version        # Shows 0.1.0
python -m md2slides --list-themes    # Lists 4 themes
```

**Issues and resolutions:**
- Network issues prevented pip/conda from installing packages
- Resolved by using `argparse` (stdlib) instead of `click`, and `mistune` (already installed) instead of `markdown-it-py`
- All dependencies now available: python-pptx, mistune are installed

### 2026-02-07 - Task 3: Implement Markdown Parser

**Task completed:**
3. Implement markdown parser

**Changes made:**
- Created `md2slides/parser.py` with MarkdownParser class
- Implemented parsing of H1 (sections), H2 (slides), paragraphs, bullet lists, nested lists
- Added support for code blocks with language detection
- Added table parsing support using mistune's table plugin
- Created data classes: Presentation, Section, Slide, TextElement, BulletList, ListItem, CodeBlock, Table, Image

**Commands run:**
```bash
python -m py_compile md2slides/parser.py  # Syntax check passed
python -c "from md2slides.parser import parse_markdown; ..."  # Parse test passed
```

**Verification:**
- Parser correctly extracts presentation title from first H1
- Parser creates slides from H2 headings
- Bullet lists with nesting are properly parsed
- Code blocks retain language info (e.g., python, javascript)
- Tables are parsed with headers, rows, and alignments

### 2026-02-07 - Task 4: Implement Basic PPTX Converter

**Task completed:**
4. Implement basic PPTX converter

**Changes made:**
- Created `md2slides/converter.py` with PresentationConverter class
- Implemented title slide creation from first H1
- Implemented content slides from H2 headings
- Added text paragraphs and bullet lists (including nested) to slides
- Added basic code block and table support

**Commands run:**
```bash
python -m py_compile md2slides/converter.py  # Syntax check passed
python -c "from md2slides.converter import convert_markdown_to_pptx; ..."  # Test passed
```

**Verification:**
- Generated test-output.pptx with 3 slides (title + 2 content)
- Generated example-output.pptx from resources/example.md with 5 slides
- All PPTX files open correctly

### 2026-02-07 - Task 5: Wire Up CLI to Parser and Converter

**Task completed:**
5. Wire up CLI to parser and converter

**Changes made:**
- Updated `md2slides/cli.py` to integrate with converter
- Added file validation (exists, is_file)
- Added automatic output path generation (input.md -> input.pptx)
- Added error handling with proper exit codes

**Commands run:**
```bash
python -m md2slides --help                            # Shows updated usage
python -m md2slides resources/example.md -o cli-test.pptx  # Creates PPTX
python -m md2slides nonexistent.md                    # Returns error (exit 1)
```

**Verification:**
- CLI converts markdown files to PPTX successfully
- Missing file error is handled gracefully
- Output file is created with correct content

### 2026-02-07 - Task 6: Implement Default Theme

**Task completed:**
6. Implement default theme

**Changes made:**
- Created `md2slides/themes/` directory
- Created `md2slides/themes/__init__.py` with get_theme() function
- Created `md2slides/themes/base.py` with ThemeColors, ThemeFonts, BaseTheme classes
- Created `md2slides/themes/default.py` with DefaultTheme (blue/gray color scheme)
- Updated converter.py to apply theme to title and content slides

**Commands run:**
```bash
python -m py_compile md2slides/themes/*.py  # Syntax check passed
python -m md2slides resources/example.md -o themed-test.pptx  # Creates themed PPTX
```

**Verification:**
- Theme applies title colors (dark blue) and body colors (dark gray)
- Fonts are applied consistently (Calibri for title/body)

### 2026-02-07 - Task 7: Add Dark, Light, and Corporate Themes

**Task completed:**
7. Add dark, light, and corporate themes

**Changes made:**
- Created `md2slides/themes/dark.py` (dark background, light text, bright blue accent)
- Created `md2slides/themes/light.py` (white background, minimal gray palette, Arial fonts)
- Created `md2slides/themes/corporate.py` (navy blue titles, Georgia font, bold headings)
- Updated themes/__init__.py to register all four themes

**Commands run:**
```bash
python -m md2slides --list-themes                           # Shows all 4 themes
python -m md2slides resources/example.md --theme dark -o dark-test.pptx
python -m md2slides resources/example.md --theme light -o light-test.pptx
python -m md2slides resources/example.md --theme corporate -o corporate-test.pptx
```

**Verification:**
- All four themes (default, dark, light, corporate) produce distinct PPTX files
- --list-themes shows all available themes
- --theme flag properly selects the theme

### 2026-02-07 - Task 8: Implement Image Handling

**Task completed:**
8. Implement image handling

**Changes made:**
- Updated parser.py to detect image tokens (both standalone and inline patterns)
- Added _extract_images() and _extract_text_excluding_images() methods
- Updated converter.py to embed local images in slides
- Added _add_image() and _resolve_image_path() methods
- Images are automatically sized to max 6 inches width

**Commands run:**
```bash
python -m md2slides resources/example.md -o image-test.pptx  # Creates PPTX with image
```

**Verification:**
- Images are detected in markdown (![alt](src) syntax)
- Local images are embedded in slides (file size increased from 32KB to 38KB)
- Image is positioned and sized appropriately (6.00in x 3.36in)

### 2026-02-07 - Task 9: Implement Code Block Handling

**Task completed:**
9. Implement code block handling

**Changes made:**
- Enhanced _add_code_block() in converter.py to use add_run() for proper font styling
- Updated all theme files to skip code runs when applying body font
- Code blocks now use Courier New monospace font at 12pt
- Created examples/code-test.md for testing

**Commands run:**
```bash
python -m md2slides examples/code-test.md -o code-test.pptx  # Creates PPTX with code
```

**Verification:**
- Code blocks use Courier New font while body text uses Calibri
- Code blocks preserve indentation and formatting
- Python and JavaScript code blocks render correctly

**Note:** Pygments syntax highlighting not implemented (package unavailable due to network issues)

### 2026-02-07 - Task 10: Implement Table Support

**Task completed:**
10. Implement table support

**Changes made:**
- Enhanced _add_table() in converter.py with theme styling
- Added cell alignment support (left, center, right)
- Headers are bold, data cells use theme body font
- Created examples/table-test.md for testing

**Commands run:**
```bash
python -m md2slides examples/table-test.md -o table-test.pptx  # Creates PPTX with tables
```

**Verification:**
- Tables render with correct rows and columns
- Headers are bold, data cells use normal weight
- Alignment is applied to cells

### 2026-02-07 - Task 11: Create Test Markdown Files

**Task completed:**
11. Create test markdown files

**Changes made:**
- Created examples/simple.md with basic headings and bullet lists
- Created examples/full-featured.md with code, tables, and text
- Created examples/theme-showcase.md to demonstrate all themes

**Commands run:**
```bash
python -m md2slides examples/simple.md -o simple.pptx        # OK
python -m md2slides examples/full-featured.md -o full.pptx   # OK
python -m md2slides examples/theme-showcase.md -o themes.pptx # OK
```

**Verification:**
- All 5 example files convert without errors
- Each example demonstrates different md2slides features

### 2026-02-07 - Task 12: Write README and Documentation

**Task completed:**
12. Write README and documentation

**Changes made:**
- Replaced README.md with md2slides project documentation
- Added installation instructions
- Documented CLI usage with examples
- Documented all supported markdown syntax
- Listed all themes with descriptions
- Added project structure reference

**Verification:**
- README is comprehensive and well-organized
- All features are documented
- Examples directory is referenced

### 2026-02-07 - Task 13: Move PPTX Templates into Dedicated templates/ Folder

**Task completed:**
13. Move PPTX templates into a dedicated templates/ folder and update codepaths

**Changes made:**
- templates/ directory already existed with multiverse_logo.png
- Moved `archived-instructions/template.pptx` to `templates/template.pptx`
- Created `md2slides/utils.py` with helper functions:
  - `get_template_path(name)` - Get path to a template file
  - `get_templates_dir()` - Get path to templates directory
  - `list_templates()` - List all available template files

**Commands run:**
```bash
mv archived-instructions/template.pptx templates/template.pptx  # Moved template
python -m py_compile md2slides/utils.py  # Syntax check passed
python -c "from md2slides.utils import get_template_path, list_templates; ..."  # Helper test passed
python -m md2slides --help  # CLI still works
python -m md2slides examples/simple.md -o test-output.pptx  # Conversion works
```

**Verification:**
- templates/ directory contains: template.pptx, multiverse_logo.png
- get_template_path() correctly resolves template paths
- list_templates() returns ['template.pptx', 'multiverse_logo.png']
- CLI continues to function normally

### 2026-02-07 - Task 14: Add Company Logo to Every Slide

**Task completed:**
14. Add company logo to bottom-right corner of every slide from templates/multiverse_logo.png

**Changes made:**
- Updated `md2slides/converter.py`:
  - Added import for `get_template_path` from utils
  - Added `_add_logo_to_all_slides()` method that iterates all slides after creation
  - Logo loaded from `templates/multiverse_logo.png` using the template helper
  - Logo placed at bottom-right with 0.25" margin from edges
  - Logo scaled to max height of 0.5" while maintaining aspect ratio
  - Called as post-processing step in `convert()` after all slides are created

**Commands run:**
```bash
python -m py_compile md2slides/converter.py  # Syntax check passed
python -m md2slides --help                   # CLI works
python -m md2slides examples/simple.md -o test-output.pptx  # Creates PPTX with logo
python -m md2slides examples/full-featured.md -o final-test.pptx  # Works with content
```

**Verification:**
- Logo appears on all slides (verified with Python inspection)
- Logo position: 6.64" from left, 6.75" from top (bottom-right corner)
- Logo size: 3.11" x 0.50" (aspect ratio maintained)
- Logo does not overlap with content area
- File size increased from ~32KB to ~62KB (logo embedded)

### 2026-02-07 - Task 15: Two-Column Layout for Text+Image Slides

**Task completed:**
15. When a slide has both text and one image: use a 2-column layout with autoscaled image on the right

**Changes made:**
- Updated `md2slides/converter.py`:
  - Added layout detection in `_add_content_slide()` to identify slides with text elements + exactly 1 image
  - Added `_add_two_column_slide()` method using LAYOUT_BLANK for custom positioning
  - Content split: 55% left (text), 45% right (image) with 0.3" gutter
  - Title positioned at top, content below with margins for logo safe area
  - Added `_add_scaled_image()` method using PIL to calculate optimal image scaling
  - Image is scaled to fit right region preserving aspect ratio, centered within region
- Created `examples/two-column-test.md` for testing various slide configurations

**Commands run:**
```bash
python -m py_compile md2slides/converter.py  # Syntax check passed
python -m md2slides examples/two-column-test.md -o two-column-test.pptx  # Creates PPTX
python -c "from pptx import Presentation; ..."  # Verified layout structure
```

**Verification:**
- Slides with text + 1 image use "Blank" layout with 2-column positioning
- Text content on left: 0.50" from left, 4.80" width (55%)
- Image on right: 5.99" from left, 3.12" width, centered in right region
- Slides with no images or 2+ images use standard "Title and Content" layout
- Logo appears on all slides in correct position
- Image aspect ratio preserved during scaling

### 2026-02-07 - Task 16: Two-Image Vertical Stacking Layout

**Task completed:**
16. If two images are present on a slide, stack them vertically (top and bottom) with consistent spacing

**Changes made:**
- Updated `md2slides/converter.py`:
  - Extended `_add_content_slide()` to detect slides with exactly 2 images (no text elements)
  - Added `_add_two_image_stacked_slide()` method that:
    - Uses LAYOUT_BLANK for custom positioning
    - Computes image region (full width, excludes title + logo safe area)
    - Splits region into two vertical slots with 0.2" gutter
    - Scales each image to fit its slot preserving aspect ratio
    - Centers images horizontally within full-width region
  - Deterministic ordering: images appear in markdown order (top to bottom)
- Created `examples/two-images-test.md` for testing the feature

**Commands run:**
```bash
python -m py_compile md2slides/converter.py  # Syntax check passed
python -m md2slides examples/two-images-test.md -o two-images-output.pptx  # Creates PPTX
python -c "from pptx import Presentation; ..."  # Verified layout structure
```

**Verification:**
- Slide with 2 images uses "Blank" layout with stacked positioning
- First image: 3.12"w x 1.75"h at (3.44", 1.89") - top position
- Second image: 3.12"w x 1.75"h at (3.44", 4.61") - bottom position
- Images are horizontally centered with consistent sizing
- Gutter spacing of 0.2" between images
- Logo appears on all slides in bottom-right corner
- Slides with 1 image + text still use 2-column layout correctly
- Text-only slides still use standard layout correctly

### 2026-02-07 - Task 17: Render Markdown Hyperlinks as Clickable Links in PPTX

**Task completed:**
17. Render markdown hyperlinks as clickable links in PPTX

**Changes made:**
- Updated `md2slides/parser.py`:
  - Added `InlineSpan` dataclass to represent inline text with optional hyperlink (href), bold, italic, code attributes
  - Added `spans` field to `TextElement` dataclass for rich inline content
  - Added `spans` field to `ListItem` dataclass for hyperlinks in list items
  - Added `_extract_inline_spans()` method to parse inline tokens into spans, preserving hyperlinks
  - Updated paragraph parsing to extract spans alongside plain text
  - Updated list item parsing to extract spans for hyperlink support in lists
- Updated `md2slides/converter.py`:
  - Added imports for `InlineSpan` and `ListItem`
  - Added `_add_spans_to_paragraph()` method that renders spans with hyperlinks using `run.hyperlink.address`
  - Hyperlinks are styled with blue color (0x0563C1) and underline
  - Updated `_add_text_element()` to use spans when available
  - Added `_add_list_item_content()` method to render list items with hyperlink support
  - Updated `_add_bullet_list()` to use the new method for all list item levels
- Created `examples/hyperlinks-test.md` for testing various hyperlink scenarios

**Commands run:**
```bash
python -m py_compile md2slides/*.py  # Syntax check passed
python -m md2slides examples/hyperlinks-test.md -o hyperlinks-test.pptx  # Creates PPTX
python -c "from pptx import Presentation; ..."  # Verified hyperlinks in PPTX
python -m md2slides examples/simple.md -o test-output.pptx  # Existing examples still work
```

**Verification:**
- Hyperlinks in paragraphs: "Google" -> https://google.com, "GitHub" -> https://github.com
- Hyperlinks in lists: "Anthropic" -> https://anthropic.com, "Stack Overflow" -> https://stackoverflow.com
- Multiple hyperlinks per paragraph work correctly
- Links are styled with blue color and underline
- Edge cases handled: multiple links in paragraph, links in list items, trailing punctuation
- Existing examples convert without errors

### 2026-02-07 - Task 18: Treat Text After H1 as Subtitle

**Task completed:**
18. Treat text immediately after an H1 as a subtitle for the section/title slide

**Changes made:**
- Updated `md2slides/parser.py`:
  - Added `subtitle` field to `Section` dataclass (default empty string)
  - Added `awaiting_subtitle` flag in `_build_presentation()` to track when a paragraph should be captured as subtitle
  - When a paragraph appears immediately after H1 (before any H2), it is consumed as `section.subtitle`
  - Subtitle paragraph is not rendered as regular slide content elsewhere
- Updated `md2slides/converter.py`:
  - Modified `convert()` to extract subtitle from first section when it matches the presentation title
  - Passed subtitle to `_add_title_slide()` which renders it using the existing subtitle placeholder
- Created `examples/subtitle-test.md` for testing the feature

**Commands run:**
```bash
python -m py_compile md2slides/*.py  # Syntax check passed
python -m md2slides --help            # CLI works
python -m md2slides examples/subtitle-test.md -o subtitle-test.pptx  # Creates PPTX with subtitle
python -m md2slides examples/simple.md -o test-output.pptx  # No subtitle when none expected
python -m md2slides examples/hyperlinks-test.md -o hyperlinks-test.pptx  # Existing features work
```

**Verification:**
- Title slide shows: "My Presentation Title" with subtitle "This is the subtitle text that appears right after the H1 title."
- Subtitle does not appear as body text on subsequent slides
- Existing examples (simple.md, hyperlinks-test.md) continue to work correctly
- When no paragraph follows H1, no subtitle appears (correct behavior)

### 2026-02-07 - Task 19: Render H2 Paragraphs as Inline Text (Not Bullets)

**Task completed:**
19. If text appears after an H2 and is not in a list, render it inline (not as a bullet list) in PPTX

**Changes made:**
- Updated `md2slides/parser.py`:
  - Added `is_paragraph: bool = True` field to `TextElement` dataclass
  - Paragraphs are marked with `is_paragraph=True` (default) to distinguish them from list content
- Updated `md2slides/converter.py`:
  - Added import for `PP_PARAGRAPH_ALIGNMENT` from `pptx.enum.text`
  - Added `_disable_bullet()` method that uses XML manipulation to set `<a:buNone>` element
  - Updated `_add_text_element()` to call `_disable_bullet()` when `element.is_paragraph` is True
  - This explicitly removes bullet formatting for paragraphs while preserving bullets for list items
- Created `examples/paragraph-test.md` for testing various paragraph/list combinations

**Commands run:**
```bash
python -m py_compile md2slides/*.py  # Syntax check passed
python -m md2slides --help           # CLI works
python -m md2slides examples/paragraph-test.md -o paragraph-test.pptx  # Creates PPTX
python -m md2slides examples/simple.md -o test-output.pptx  # Existing examples work
python -m md2slides examples/hyperlinks-test.md -o hyperlinks-test.pptx  # Existing examples work
```

**Verification:**
- Paragraphs in PPTX have `<a:buNone>` element set (verified via XML inspection)
- List items retain default bullet formatting
- Slide 2 (paragraphs only): All paragraphs display without bullets
- Slide 3 (paragraph + list): Intro/outro paragraphs have no bullets, list items have bullets
- Slide 4 (list only): All items have bullets as expected
- Slide 5 (mixed): Paragraphs before/after list have no bullets, list items have bullets
- PPTX XML confirmed: paragraphs show "NO BULLET", list items show "(default/bullet)"

### 2026-02-07 - Task 20: Render H3 Headings as Bold, Red, and Inline

**Task completed:**
20. Render H3 headings as bold, red, and inline with subsequent content

**Changes made:**
- Updated `md2slides/parser.py`:
  - Added `InlineHeading` dataclass with `text`, `following_text`, and `following_spans` fields
  - Added `pending_h3` tracking in `_build_presentation()` to capture H3 headings
  - H3 tokens are now stored and merged with the following paragraph content
  - When H3 is followed by a list, it's rendered as a standalone inline heading
- Updated `md2slides/converter.py`:
  - Added import for `InlineHeading` type
  - Added `_add_inline_heading()` method that renders H3 as bold red with separator
  - Added `_add_spans_to_run()` helper for appending spans without clearing paragraph
  - H3 text is styled with `font.bold = True` and `font.color.rgb = RGBColor(0xFF, 0x00, 0x00)`
  - Separator " — " is added between H3 and following text, unless H3 ends with punctuation (.:;!?—-)
  - Updated layout detection to include `InlineHeading` in text elements
- Updated all theme files (`default.py`, `dark.py`, `light.py`, `corporate.py`):
  - Added check to skip runs with H3 red color (`FF0000`) to prevent theme from overriding the styling
- Created `examples/h3-test.md` for testing various H3 scenarios

**Commands run:**
```bash
python -m py_compile md2slides/parser.py md2slides/converter.py  # Syntax check passed
python -m py_compile md2slides/themes/*.py  # Themes syntax OK
python -m md2slides --help  # CLI works
python -m md2slides examples/h3-test.md -o h3-test.pptx  # Creates PPTX
python -m md2slides examples/simple.md -o test-output.pptx  # Existing examples work
python -m md2slides examples/hyperlinks-test.md -o hyperlinks-test.pptx  # Existing examples work
```

**Verification:**
- H3 headings render as bold with red color (`FF0000`)
- H3 followed by paragraph: "Key Point — This text..." (inline with em-dash separator)
- H3 followed by list: "Features:" appears as standalone bold red heading, list follows
- H3 ending with punctuation: "Important Note:" (no extra separator added)
- Theme does not override the red/bold H3 styling
- Existing examples (simple.md, hyperlinks-test.md) continue to work correctly

### 2026-02-07 - Task 21: Add Examples and Build Script for PPTX Generation

**Task completed:**
21. Add examples demonstrating new layout and markdown behavior; output generated PPTX examples into examples/ folder

**Changes made:**
- Created `examples/all-features.md` - comprehensive example demonstrating all new features:
  - Logo placement on every slide
  - Two-column layouts for text and images
  - Vertical stacking for multiple images
  - Clickable hyperlinks
  - H1 subtitle support
  - Paragraph rendering without bullets
  - H3 inline headings with bold red styling
  - Code blocks and tables
- Created `examples/build_examples.py` - build script that:
  - Finds all .md files in examples/ directory
  - Converts each to .pptx using `convert_markdown_to_pptx()`
  - Changes to examples/ directory for relative image path resolution
  - Reports success/failure with file sizes
  - No external network access required (uses local images only)

**Commands run:**
```bash
python -m py_compile md2slides/*.py  # Syntax check passed
python -m md2slides --help            # CLI works
python examples/build_examples.py     # Builds all 13 examples successfully
```

**Verification:**
- Build script generates 13 PPTX files in examples/
- All conversions succeed with varying file sizes (58-74 KB)
- all-features.pptx has 10 slides with proper content
- Layouts verified: Title Slide, Blank (for custom layouts), Title and Content
- Logo appears on every slide
- No network access required - uses examples/test_image.jpeg for all images

### 2026-02-07 - Task 22: Ensure PPTX Files Are Not Version Controlled

**Task completed:**
22. Ensure PPTX files are not version controlled (ignore outputs and templates as appropriate)

**Changes made:**
- Updated `.gitignore` to add:
  ```
  # Generated PPTX files (keep templates/template.pptx tracked)
  *.pptx
  !templates/template.pptx
  ```
- Policy: Ignore all generated PPTX outputs, but keep `templates/template.pptx` tracked (as it's a base template file)

**Commands run:**
```bash
git ls-files "*.pptx"                    # Confirmed templates/template.pptx is tracked
git check-ignore -v examples/*.pptx      # Verified generated files are ignored
git check-ignore templates/template.pptx  # Verified template is NOT ignored
python examples/build_examples.py         # Regenerated all PPTX files
git status                                # Confirmed no PPTX files show as untracked
```

**Verification:**
- All generated PPTX files (examples/*.pptx, *.pptx in root) are ignored by git
- templates/template.pptx remains tracked and is not ignored
- After running build_examples.py, git status shows no PPTX files as untracked/modified
- No PPTX files need to be removed from git index (none were tracked except template)
