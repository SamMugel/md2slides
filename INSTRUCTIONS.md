# md2slides - Markdown to PPTX Conversion Specifications

This document describes how md2slides transforms Markdown files into PowerPoint (PPTX) presentations.

## Document Structure

### Heading 1 (`#`)
- Creates the **title slide**
- The heading text becomes the presentation title
- Only one H1 should be used per document (at the beginning)

### Content Under Heading 1
- Any text immediately following the H1 heading becomes the **subtitle** on the title slide
- If no content exists under H1, the title slide displays only the title

### Heading 2 (`##`)
- Each H2 heading starts a **new content slide**
- The heading text becomes the slide title

### Content Under Heading 2
- All content following an H2 (until the next H2 or end of document) becomes the slide body

## Formatting Rules

### Text Formatting
All standard Markdown formatting is preserved in the PPTX output:
- **Bold** (`**text**` or `__text__`)
- *Italic* (`*text*` or `_text_`)
- Combined formatting (e.g., ***bold italic***)

### Lists
- Bullet point lists render as native PPTX bullet lists
- Nested lists maintain proper indentation levels
- Numbered lists are supported

### Content Fitting
- All slide content must fit within slide boundaries
- Text sizing automatically scales to accommodate all content on a single slide

## Output Format

The converter produces a single PPTX file containing:

1. **Title Slide**
   - Title (from H1)
   - Subtitle (from content under H1, if present)

2. **Content Slides**
   - One slide per H2 heading
   - Slide title from H2 text
   - Slide body from content under that H2

## Example

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
- **Slide 1 (Title)**: "Quarterly Report" with subtitle "Q3 2024 Financial Summary"
- **Slide 2**: "Revenue Overview" with bullet list content
- **Slide 3**: "Next Steps" with bullet list content (bold formatting preserved)
