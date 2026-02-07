# Theme Showcase

## About Themes

md2slides comes with four built-in themes:

- **Default** - Clean blue and gray professional look
- **Dark** - Dark background with light text
- **Light** - Minimal, white with gray accents
- **Corporate** - Navy blue with Georgia headings

## Sample Content

This slide has mixed content to show theme styling:

- Regular bullet point
- Another point here
  - Nested item
  - Another nested item

## Code Sample

```javascript
function displayTheme(name) {
    console.log(`Current theme: ${name}`);
    return true;
}
```

## Data Table

| Theme | Background | Text Color |
|-------|------------|------------|
| Default | White | Dark Blue |
| Dark | Dark Gray | White |
| Light | White | Gray |
| Corporate | White | Navy |

## Try Each Theme

Generate this presentation with each theme:

```bash
md2slides theme-showcase.md --theme default -o default.pptx
md2slides theme-showcase.md --theme dark -o dark.pptx
md2slides theme-showcase.md --theme light -o light.pptx
md2slides theme-showcase.md --theme corporate -o corporate.pptx
```
