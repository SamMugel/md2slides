# md2slides - Execution plan

Context: prd.md

---

## Task List

```json
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
  "passes": false
}
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
  "passes": false
}
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
  "passes": false
}
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
  "passes": false
}
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
  "passes": false
}
{
  "category": "feature",
  "description": "Treat text immediately after an H1 as a subtitle for the section/title slide",
  "steps": [
    "Update parser.py section handling: when an H1 is followed by a paragraph (before any H2), attach that paragraph as section_subtitle.",
    "Update converter.py title/section slide creation to add a subtitle text box under the H1 title using theme subtitle styles.",
    "Ensure subtitle does not get rendered as a normal body paragraph elsewhere (consume it as subtitle).",
    "Verify: example markdown with H1 + paragraph + H2 renders a title slide with subtitle and subsequent slides unaffected."
  ],
  "passes": false
}
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
  "passes": false
}
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
}
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
}
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
```

---

## Agent Instructions

0. Read `prd.md` first to understand the vision
1. Read `activity.md` to understand current state
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
