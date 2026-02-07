# md2slides Feature Showcase

A comprehensive demonstration of all md2slides features including layouts, hyperlinks, subtitles, and inline headings.

## Introduction

This presentation showcases the complete feature set of md2slides, converting markdown to professional PowerPoint presentations.

Key features demonstrated:

- Logo placement on every slide
- Two-column layouts for text and images
- Vertical stacking for multiple images
- Clickable hyperlinks
- H1 subtitle support
- Paragraph rendering without bullets
- H3 inline headings

## Text and Image Layout

This slide demonstrates the two-column layout feature.

When a slide contains both text content and exactly one image, md2slides automatically arranges them in a side-by-side layout with text on the left and the image on the right.

- Text fills the left column (55% width)
- Image auto-scales to fit the right column
- Aspect ratio is preserved

![Demo Image](test_image.jpeg)

## Two Image Stacking

This slide demonstrates vertical image stacking with two images.

![First Image](test_image.jpeg)
![Second Image](test_image.jpeg)

## Hyperlinks in Action

Visit [md2slides on GitHub](https://github.com) for the source code. You can also check out [Python-PPTX documentation](https://python-pptx.readthedocs.io) for the underlying library.

Links also work in lists:

- Learn more at [Python.org](https://python.org)
- Find packages at [PyPI](https://pypi.org)
- Get help on [Stack Overflow](https://stackoverflow.com)

## Inline Paragraphs

This paragraph appears without bullet formatting because it's plain text after an H2 heading.

This second paragraph also has no bullets. md2slides distinguishes between paragraphs and list items.

Now here's a bullet list for comparison:

- This item has a bullet
- So does this one
- And this final item

## Inline H3 Headings

### Key Concept

This text follows an H3 heading and is rendered inline with it. The heading appears bold and red, with an em-dash separator.

### Important Note:

When the H3 ends with punctuation, no separator is added.

### Multiple Sections

You can have multiple H3 sections on the same slide, each rendered with its following content.

## Code Blocks

Here's an example of syntax-highlighted code:

```python
def convert_markdown(input_file, output_file, theme="default"):
    """Convert a markdown file to PowerPoint presentation."""
    parser = MarkdownParser()
    presentation = parser.parse(input_file)

    converter = PresentationConverter(theme=theme)
    converter.convert(presentation)
    converter.save(output_file)
```

## Data Tables

Tables are rendered with theme-consistent styling:

| Feature | Status | Notes |
|---------|:------:|-------|
| Logo placement | Complete | Bottom-right corner |
| Two-column layout | Complete | Text + image |
| Image stacking | Complete | Two images vertical |
| Hyperlinks | Complete | Clickable links |
| Subtitles | Complete | After H1 |
| Inline paragraphs | Complete | No bullets |
| H3 styling | Complete | Bold red inline |

## Conclusion

This presentation demonstrated all major md2slides features:

- Automatic logo placement on every slide
- Smart layout detection for images
- Rich text formatting with hyperlinks
- Semantic heading support (H1, H2, H3)
- Tables and code blocks
- Theme consistency throughout

Visit [the project repository](https://github.com) to learn more!
