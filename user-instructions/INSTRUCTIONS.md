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

#### Bullet Points (Unordered Lists)
Bullet points in Markdown use `-`, `*`, or `+` characters at the start of a line:

```markdown
- First item
- Second item
- Third item
```

These render as native PPTX bullet points with:
- Each line starting with `-`, `*`, or `+` becomes a separate bullet point
- Standard bullet styling applied
- Proper spacing between items

#### Nested Bullet Points
Nested lists are created by indenting items with 2-4 spaces before the `-`, `*`, or `+` character:

```markdown
- Parent item
  - Child item 1
  - Child item 2
- Another parent item
  - Another child item
```

Nested bullet points render with:
- Proper indentation levels (parent bullets at level 1, child bullets at level 2, etc.)
- Indented bullets display with smaller bullet symbols or different styling
- All nesting levels must maintain consistent indentation (2-4 spaces per level)

#### Numbered Lists (Ordered Lists)
Numbered lists use `1.`, `2.`, `3.`, etc. at the start of a line:

```markdown
1. First step
2. Second step
3. Third step
```

Numbered lists render as native PPTX ordered lists with:
- Sequential numbering in the presentation
- Numbers are automatically applied regardless of input numbering
- Proper spacing between items

#### Nested Numbered Lists
Nested numbered lists follow the same indentation pattern as bullet points:

```markdown
1. Parent step
   1. Child step A
   2. Child step B
2. Another parent step
   1. Another child step
```

#### Mixed Lists
You can mix bullet and numbered lists in the same slide. Each maintains its own format:

```markdown
- Top-level bullet point
  1. Numbered sub-item
  2. Another numbered sub-item
- Another top-level bullet
```

### Content Fitting
- All slide content must fit within slide boundaries
- Text sizing automatically scales to accommodate all content on a single slide
- List items automatically adjust font size if there are many items

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

## Implementation Steps

1. Increase marketing spend
2. Launch in **two new regions**
3. Hire additional sales staff
   - Technical roles
   - Sales roles
```

This produces:
- **Slide 1 (Title)**: "Quarterly Report" with subtitle "Q3 2024 Financial Summary"
- **Slide 2**: "Revenue Overview" with unordered bullet list including nested items
- **Slide 3**: "Implementation Steps" with ordered list including nested items (bold formatting preserved)
