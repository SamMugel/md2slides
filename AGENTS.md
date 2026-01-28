# AGENTS.md

This file provides guidance to WARP (warp.dev) when working with code in this repository.

## Common commands

### Install (editable)
```bash
pip install -e .
```

### Install with dev/test deps
```bash
pip install -e ".[dev]"
```

### Run the CLI
The console script `md2slides` is defined in `pyproject.toml` under `[project.scripts]`.

```bash
# Convert with automatic output name (presentation.md -> presentation.pptx)
md2slides presentation.md

# Specify output path
md2slides presentation.md -o slides.pptx

# Show version
md2slides --version
```

### Run tests
Pytest is configured in `pyproject.toml` under `[tool.pytest.ini_options]`.

```bash
pytest
```

Run a single file:
```bash
pytest tests/test_parser.py
```

Run a single test:
```bash
pytest tests/test_parser.py::TestTitleSlide::test_h1_creates_title_slide
```

Run tests matching a substring:
```bash
pytest -k bold
```

Coverage outputs:
- Terminal report is enabled by default via `addopts`.
- HTML report is written to `htmlcov/`.

```bash
open htmlcov/index.html
```

### Lint / format
No linter/formatter is configured in this repo (no ruff/black/isort/mypy config in `pyproject.toml`).

## High-level architecture

### End-to-end flow
1. **Input**: Markdown content from a `.md` file (CLI) or a string (Python API).
2. **Parse**: `md2slides.parser.MarkdownParser` converts Markdown text into a list of `Slide` objects.
3. **Render**: `md2slides.converter.MarkdownToPptxConverter` renders `Slide` objects into a `python-pptx` `Presentation`.
4. **Output**: A `.pptx` file is written to disk.

### Key modules
- `src/md2slides/cli.py`
  - Minimal argparse wrapper around `md2slides.converter.convert_file`.
  - Maps exceptions to exit codes (0 success, 1 error).

- `src/md2slides/parser.py`
  - Owns the “Markdown → slide model” transformation.
  - Core data model:
    - `Slide` (title slide vs content slide; optional subtitle)
    - `ListItem` (ordered/bulleted, nesting level)
    - `TextRun` (text + bold/italic flags)
  - Parsing strategy is line-based:
    - `# ` starts a title slide; subsequent non-empty lines become the subtitle until the first `## `.
    - `## ` starts a new content slide; subsequent lines become slide body.
    - List detection is regex-based (bullets `-/*/+` and numbered `1.` or `1)`).
    - Inline formatting is handled by `_parse_inline_formatting()` and produces multiple `TextRun`s per line.

- `src/md2slides/converter.py`
  - Owns the “slide model → PPTX” transformation using `python-pptx`.
  - Uses blank slide layout (`slide_layouts[6]`) and manually positions text boxes.
  - `_render_content()` is where `ListItem` vs `TextRun` are mapped to paragraphs/runs.
  - Note: ordered lists are rendered by prefixing text with `"{n}. "` (due to limited numbered-list support).

### Project specs
- `INSTRUCTIONS.md` is the authoritative spec for how headings and content map to slides.
- Tests in `tests/` are organized by layer:
  - `tests/test_parser.py` validates parsing + formatting extraction.
  - `tests/test_converter.py` validates PPTX creation and basic content/formatting presence.
  - `tests/test_cli.py` validates the CLI wrapper behavior.

### Agent/tooling notes
- `.gitignore` intentionally excludes common local artifacts (venv, caches, coverage output, `.claude/`, etc.).
- There is no repo-level `CLAUDE.md`/Cursor/Copilot instruction file; `.claude/settings.local.json` exists but is intended to be local-only (and `.claude/` is ignored).