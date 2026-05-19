# PPTXLib Documentation

## Overview

PPTXLib is a pure-Ring library for creating Microsoft PowerPoint (.pptx) files using the Ring programming language. It generates fully compatible Office Open XML (ECMA-376) files that can be opened in Microsoft PowerPoint, LibreOffice Impress, Google Slides, and other presentation applications.

## Features

- **Multiple Slides** - Create presentations with unlimited slides
- **Text Boxes** - Formatted text with fonts, colors, alignment
- **Rich Text** - Multiple formats within a single text box
- **Shapes** - Rectangles, circles, ovals, rounded rectangles, lines
- **Images** - PNG, JPG, BMP, GIF support
- **Tables** - With headers, styling, and zebra striping
- **Lists** - Bullet and numbered lists
- **Backgrounds** - Solid colors per slide
- **Layouts** - 16:9, 16:10, 4:3 aspect ratios
- **Quick Layouts** - Title slides, two-column layouts
- **Speaker Notes** - Add notes to slides
- **No Dependencies** - Pure Ring implementation

## Installation

	ringpm install pptxlib from ringpackages

---

## Quick Start

### Simple Presentation

```ring
load "pptxlib.ring"

ppt = new PPTWriter()
ppt.addTitleSlide("My Presentation", "Subtitle here")
ppt.addSlide()
ppt.addTitle("First Slide")
ppt.addTextBox("Hello, World!", 0.5, 1.5, 9, 2, NULL)
ppt.save("output.pptx")
```

### Quick Presentation Function

```ring
load "pptxlib.ring"

slides = [
    ["Welcome", "Introduction text"],
    ["Features", "Feature description"],
    ["Conclusion", "Thank you!"]
]
quickPPT("output.pptx", "My Presentation", slides)
```

---

## API Reference

### PPTWriter Class

#### Constructor

```ring
ppt = new PPTWriter()
```

Creates a new PowerPoint presentation.

---

### Document Properties

| Method | Description |
|--------|-------------|
| `setAuthor(author)` | Set document author |
| `setTitle(title)` | Set document title |
| `setSubject(subject)` | Set document subject |
| `setCompany(company)` | Set company name |

---

### Layout Settings

| Method | Description |
|--------|-------------|
| `setLayout(layout)` | Set slide layout (see constants) |
| `setCustomSize(width, height)` | Set custom size in inches |

**Layout Constants:**
- `PPT_LAYOUT_16x9` - 10" × 5.625" (default, widescreen)
- `PPT_LAYOUT_16x10` - 10" × 6.25"
- `PPT_LAYOUT_4x3` - 10" × 7.5" (traditional)

```ring
ppt.setLayout(PPT_LAYOUT_4x3)
ppt.setCustomSize(13.333, 7.5)  # Custom widescreen
```

---

### Slide Management

| Method | Description |
|--------|-------------|
| `addSlide()` | Add a blank slide |
| `addTitleSlide(title, subtitle)` | Add a title slide |
| `addContentSlide(title, content)` | Add slide with title and content |
| `addTwoColumnSlide(title, left, right)` | Add two-column slide |
| `selectSlide(index)` | Select slide by index (1-based) |
| `getSlideCount()` | Get number of slides |

---

### Slide Background

| Method | Description |
|--------|-------------|
| `setBackground(color)` | Set solid background color |
| `setBackgroundImage(path)` | Set background image |

```ring
ppt.addSlide()
ppt.setBackground("1E2761")  # Dark blue
ppt.setBackground("navy")    # Named color
```

---

### Text Elements

#### addTitle

```ring
ppt.addTitle(text)
```

Adds a title at the top of the slide (36pt, bold).

#### addSubtitle

```ring
ppt.addSubtitle(text)
```

Adds a subtitle below the title area.

#### addTextBox

```ring
ppt.addTextBox(text, x, y, width, height, options)
```

**Parameters:**
- `text` - Text content
- `x`, `y` - Position in inches from top-left
- `width`, `height` - Size in inches
- `options` - Formatting options (or NULL)

**Options:**

| Option | Type | Description |
|--------|------|-------------|
| `:fontSize` | Number | Font size in points |
| `:fontName` | String | Font name (e.g., "Arial") |
| `:bold` | Boolean | Bold text |
| `:italic` | Boolean | Italic text |
| `:underline` | Boolean | Underlined text |
| `:color` | String | Text color (name or hex) |
| `:align` | String | "left", "center", "right" |
| `:valign` | String | "top", "middle", "bottom" |
| `:bgColor` | String | Background color |
| `:borderColor` | String | Border color |
| `:borderWidth` | Number | Border width in points |

```ring
ppt.addTextBox("Hello", 0.5, 1, 9, 1, [
    :fontSize = 24,
    :bold = true,
    :color = "navy",
    :align = "center"
])
```

---

### Rich Text

```ring
ppt.addRichText(runs, x, y, width, height, options)
```

Multiple formats in one text box:

```ring
ppt.addRichText([
    ["Normal ", []],
    ["Bold", [:bold = true]],
    [" and ", []],
    ["Red", [:color = "red"]]
], 0.5, 1.5, 9, 1, NULL)
```

---

### Lists

#### Bullet List

```ring
ppt.addBulletList(items, x, y, width, height, options)
```

#### Numbered List

```ring
ppt.addNumberedList(items, x, y, width, height, options)
```

**Options:**
- `:fontSize` - Font size
- `:color` - Text color
- `:bulletColor` - Bullet color (bullet list only)

```ring
ppt.addBulletList([
    "First point",
    "Second point",
    "Third point"
], 0.5, 1.5, 4, 3, [:fontSize = 18])
```

---

### Shapes

#### addShape

```ring
ppt.addShape(shapeType, x, y, width, height, options)
```

**Shape Types:**
- `PPT_SHAPE_RECT` - Rectangle
- `PPT_SHAPE_ROUND_RECT` - Rounded rectangle
- `PPT_SHAPE_ELLIPSE` - Ellipse/Circle
- `PPT_SHAPE_LINE` - Line
- `PPT_SHAPE_TRIANGLE` - Triangle
- `PPT_SHAPE_DIAMOND` - Diamond
- `PPT_SHAPE_ARROW` - Arrow

**Options:**
- `:fillColor` - Fill color
- `:lineColor` - Border color
- `:lineWidth` - Border width
- `:text` - Text inside shape
- `:fontSize` - Text size
- `:fontColor` - Text color
- `:noFill` - No fill (outline only)

#### Convenience Methods

```ring
ppt.addRectangle(x, y, w, h, options)
ppt.addRoundedRectangle(x, y, w, h, options)
ppt.addCircle(x, y, size, options)
ppt.addOval(x, y, w, h, options)
ppt.addLine(x1, y1, x2, y2, options)
```

```ring
ppt.addRectangle(1, 1, 3, 2, [
    :fillColor = "4472C4",
    :text = "Click Here",
    :fontColor = "FFFFFF"
])

ppt.addCircle(5, 1, 1.5, [:fillColor = "green"])

ppt.addLine(0, 3, 10, 3, [:color = "red", :width = 2])
```

---

### Images

```ring
ppt.addImage(imagePath, x, y, width, height)
ppt.addImageCentered(imagePath, y, width, height)
```

**Supported Formats:** PNG, JPG, JPEG, BMP, GIF

```ring
ppt.addImage("logo.png", 0.5, 1, 3, 2)
ppt.addImageCentered("chart.jpg", 2, 6, 3)
```

---

### Tables

```ring
ppt.addTable(data, x, y, width, height, options)
ppt.addSimpleTable(data, x, y, width, height)
ppt.addStyledTable(data, x, y, w, h, headerColor, evenRowColor)
```

**Options:**
- `:headerRow` - First row is header
- `:headerBgColor` - Header background
- `:headerFontColor` - Header text color
- `:borderColor` - Border color
- `:borderWidth` - Border width
- `:evenRowBgColor` - Even row color (zebra)
- `:fontSize` - Font size

```ring
data = [
    ["Name", "Score", "Grade"],
    ["Alice", "95", "A"],
    ["Bob", "87", "B"]
]

ppt.addSimpleTable(data, 0.5, 1.5, 9, 2)

ppt.addStyledTable(data, 0.5, 1.5, 9, 2, "4472C4", "E3F2FD")
```

---

### Speaker Notes

```ring
ppt.addNotes(text)
```

Adds speaker notes to the current slide.

---

### Saving

```ring
result = ppt.save(filename)
```

Returns `true` if successful.

---

## Quick Functions

### quickPPT

```ring
quickPPT(filename, title, slides)
```

Creates a simple presentation:

```ring
slides = [
    ["Slide Title 1", "Content 1"],
    ["Slide Title 2", "Content 2"]
]
quickPPT("output.pptx", "Presentation Title", slides)
```

---

## Constants

### Layout Constants
- `PPT_LAYOUT_16x9` - Widescreen (default)
- `PPT_LAYOUT_16x10` - Widescreen variant
- `PPT_LAYOUT_4x3` - Traditional

### Alignment Constants
- `PPT_ALIGN_LEFT`, `PPT_ALIGN_CENTER`, `PPT_ALIGN_RIGHT`
- `PPT_VALIGN_TOP`, `PPT_VALIGN_MIDDLE`, `PPT_VALIGN_BOTTOM`

### Shape Constants
- `PPT_SHAPE_RECT`, `PPT_SHAPE_ROUND_RECT`
- `PPT_SHAPE_ELLIPSE`, `PPT_SHAPE_LINE`
- `PPT_SHAPE_TRIANGLE`, `PPT_SHAPE_DIAMOND`, `PPT_SHAPE_ARROW`

---

## Color Reference

### Named Colors
black, white, red, green, blue, yellow, orange, purple, gray/grey, navy, teal, maroon, silver, lime, aqua, fuchsia, olive, darkblue, darkgreen, coral, charcoal

### Hex Colors
```ring
ppt.setBackground("1E2761")     # Without #
ppt.addTextBox("Hi", 1, 1, 2, 1, [:color = "#FF5733"])  # With #
```

---

## Complete Examples

### Business Presentation

```ring
load "pptxlib.ring"

ppt = new PPTWriter()
ppt.setTitle("Q4 Review")
ppt.setAuthor("Sales Team")

# Title slide with dark background
ppt.addSlide()
ppt.setBackground("1E2761")
ppt.addTextBox("Q4 2024", 0.5, 2, 9, 0.8, [
    :fontSize = 48, :bold = true, :color = "FFFFFF", :align = "center"
])
ppt.addTextBox("Business Review", 0.5, 3, 9, 0.5, [
    :fontSize = 24, :color = "CADCFC", :align = "center"
])

# Content slide with KPI cards
ppt.addSlide()
ppt.addTitle("Key Metrics")

ppt.addRectangle(0.5, 1.5, 2.8, 1.5, [:fillColor = "E3F2FD"])
ppt.addTextBox("Revenue", 0.7, 1.6, 2.4, 0.3, [:fontSize = 12, :color = "1565C0"])
ppt.addTextBox("$12.5M", 0.7, 2, 2.4, 0.5, [:fontSize = 28, :bold = true, :color = "1565C0"])

ppt.addRectangle(3.6, 1.5, 2.8, 1.5, [:fillColor = "E8F5E9"])
ppt.addTextBox("Growth", 3.8, 1.6, 2.4, 0.3, [:fontSize = 12, :color = "2E7D32"])
ppt.addTextBox("+18%", 3.8, 2, 2.4, 0.5, [:fontSize = 28, :bold = true, :color = "2E7D32"])

# Table slide
ppt.addSlide()
ppt.addTitle("Sales by Region")

data = [
    ["Region", "Q3", "Q4", "Change"],
    ["North", "$3.2M", "$3.8M", "+18%"],
    ["South", "$2.8M", "$3.1M", "+11%"],
    ["East", "$2.5M", "$2.9M", "+16%"]
]
ppt.addStyledTable(data, 0.5, 1.5, 9, 2.5, "4472C4", "D9E2F3")

# Closing slide
ppt.addSlide()
ppt.setBackground("1E2761")
ppt.addTextBox("Thank You!", 0.5, 2.5, 9, 0.8, [
    :fontSize = 44, :bold = true, :color = "FFFFFF", :align = "center"
])

ppt.save("business_review.pptx")
```

### Product Showcase

```ring
load "pptxlib.ring"

ppt = new PPTWriter()

ppt.addTitleSlide("Product Launch", "New Features Overview")

ppt.addSlide()
ppt.addTitle("Key Features")
ppt.addBulletList([
    "Lightning fast performance",
    "Intuitive user interface",
    "Enterprise-grade security",
    "24/7 customer support"
], 0.5, 1.3, 9, 3, [:fontSize = 20])

ppt.addSlide()
ppt.addTitle("Comparison")

# Two-column layout
ppt.addRectangle(0.3, 1.2, 4.5, 3.5, [:fillColor = "FFEBEE"])
ppt.addTextBox("Before", 0.5, 1.4, 4, 0.5, [:bold = true, :color = "C62828"])
ppt.addBulletList(["Slow", "Complex", "Expensive"], 0.5, 2, 4, 2, [:fontSize = 16])

ppt.addRectangle(5.2, 1.2, 4.5, 3.5, [:fillColor = "E8F5E9"])
ppt.addTextBox("After", 5.4, 1.4, 4, 0.5, [:bold = true, :color = "2E7D32"])
ppt.addBulletList(["Fast", "Simple", "Affordable"], 5.4, 2, 4, 2, [:fontSize = 16])

ppt.save("product_launch.pptx")
```

---

## Performance Tips

1. **Pre-calculate Length in Loops**
```ring
itemsLen = len(items)
for i = 1 to itemsLen
    # process
next
```

2. **Reuse Slide Patterns** - Create similar slides with loops

3. **Batch Image Loading** - Load images once, use multiple times

---

## Troubleshooting

| Issue | Solution |
|-------|----------|
| File won't open | Check .pptx extension |
| Text cut off | Increase text box height |
| Images not showing | Verify file exists, check format |
| Shapes overlapping | Adjust x, y coordinates |

---

## Technical Notes

- **Format:** Office Open XML (ECMA-376)
- **Coordinates:** Inches (converted to EMUs internally)
- **Default Layout:** 16:9 widescreen (10" × 5.625")
