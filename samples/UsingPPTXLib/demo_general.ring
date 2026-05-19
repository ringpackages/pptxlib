/*
    PPTXLib Demo - PowerPoint Creation in Ring
    ==============================================
    This demo showcases various features of PPTXLib
*/

load "pptxlib.ring"

? "=============================================="
? "   PPTXLib Demo - PowerPoint in Ring"
? "=============================================="
? ""

# Demo 1: Simple Presentation
? "Demo 1: Creating a simple presentation..."
ppt = new PPTWriter()
ppt.setTitle("Simple Presentation")
ppt.setAuthor("Ring Programmer")

ppt.addTitleSlide("Welcome to PPTXLib", "Creating PowerPoint in Ring")

ppt.addSlide()
ppt.addTitle("Introduction")
ppt.addTextBox("This is a simple presentation created with PPTXLib.", 0.5, 1.5, 9, 1, NULL)
ppt.addTextBox("The library supports text, shapes, images, tables, and more!", 0.5, 2.5, 9, 1, NULL)

if ppt.save("demo1_simple.pptx")
    ? "  Created: demo1_simple.pptx"
else
    ? "  FAILED: demo1_simple.pptx"
ok

# Demo 2: Text Formatting
? "Demo 2: Creating presentation with formatted text..."
ppt = new PPTWriter()
ppt.setTitle("Text Formatting Demo")

ppt.addSlide()
ppt.addTitle("Text Formatting Options")

ppt.addTextBox("Normal Text", 0.5, 1.2, 4, 0.5, NULL)
ppt.addTextBox("Bold Text", 0.5, 1.8, 4, 0.5, [:bold = true])
ppt.addTextBox("Italic Text", 0.5, 2.4, 4, 0.5, [:italic = true])
ppt.addTextBox("Underlined Text", 0.5, 3.0, 4, 0.5, [:underline = true])
ppt.addTextBox("Colored Text", 0.5, 3.6, 4, 0.5, [:color = "red"])
ppt.addTextBox("Large Text", 0.5, 4.2, 4, 0.5, [:fontSize = 28])

ppt.addTextBox("Right Aligned", 5, 1.8, 4.5, 0.5, [:align = "right"])
ppt.addTextBox("Center Aligned", 5, 2.4, 4.5, 0.5, [:align = "center"])
ppt.addTextBox("With Background", 5, 3.0, 4.5, 0.5, [:bgColor = "FFFF00"])
ppt.addTextBox("With Border", 5, 3.6, 4.5, 0.5, [:borderColor = "000000", :borderWidth = 2])

if ppt.save("demo2_text_formatting.pptx")
    ? "  Created: demo2_text_formatting.pptx"
else
    ? "  FAILED: demo2_text_formatting.pptx"
ok

# Demo 3: Shapes
? "Demo 3: Creating presentation with shapes..."
ppt = new PPTWriter()
ppt.setTitle("Shapes Demo")

ppt.addSlide()
ppt.addTitle("Basic Shapes")

ppt.addRectangle(0.5, 1.5, 2, 1.2, [:fillColor = "4472C4", :text = "Rectangle", :fontColor = "FFFFFF"])
ppt.addRoundedRectangle(3, 1.5, 2, 1.2, [:fillColor = "ED7D31", :text = "Rounded", :fontColor = "FFFFFF"])
ppt.addCircle(5.5, 1.5, 1.2, [:fillColor = "70AD47", :text = "Circle", :fontColor = "FFFFFF"])
ppt.addOval(7.5, 1.5, 2, 1.2, [:fillColor = "FFC000", :text = "Oval", :fontColor = "000000"])

ppt.addRectangle(0.5, 3.2, 2, 1.2, [:noFill = true, :lineColor = "4472C4", :lineWidth = 3, :text = "Border Only", :fontColor = "4472C4"])

ppt.addLine(3, 3.5, 5, 3.5, [:color = "FF0000", :width = 3])
ppt.addLine(3, 4.0, 5, 4.0, [:color = "00FF00", :width = 2])
ppt.addLine(3, 4.5, 5, 4.5, [:color = "0000FF", :width = 1])

ppt.addTextBox("Horizontal Lines", 5.2, 3.8, 2, 0.5, [:fontSize = 12])

if ppt.save("demo3_shapes.pptx")
    ? "  Created: demo3_shapes.pptx"
else
    ? "  FAILED: demo3_shapes.pptx"
ok

# Demo 4: Lists
? "Demo 4: Creating presentation with lists..."
ppt = new PPTWriter()
ppt.setTitle("Lists Demo")

ppt.addSlide()
ppt.addTitle("Bullet and Numbered Lists")

ppt.addTextBox("Bullet List:", 0.5, 1.2, 4, 0.4, [:bold = true])
ppt.addBulletList([
    "First item",
    "Second item",
    "Third item",
    "Fourth item"
], 0.5, 1.6, 4, 2.5, [:fontSize = 16])

ppt.addTextBox("Numbered List:", 5, 1.2, 4, 0.4, [:bold = true])
ppt.addNumberedList([
    "Step one",
    "Step two",
    "Step three",
    "Step four"
], 5, 1.6, 4, 2.5, [:fontSize = 16])

if ppt.save("demo4_lists.pptx")
    ? "  Created: demo4_lists.pptx"
else
    ? "  FAILED: demo4_lists.pptx"
ok

# Demo 5: Tables
? "Demo 5: Creating presentation with tables..."
ppt = new PPTWriter()
ppt.setTitle("Tables Demo")

ppt.addSlide()
ppt.addTitle("Simple Table")

data = [
    ["Product", "Q1", "Q2", "Q3", "Q4"],
    ["Laptops", "150", "180", "200", "220"],
    ["Phones", "300", "350", "400", "450"],
    ["Tablets", "100", "120", "140", "160"]
]

ppt.addSimpleTable(data, 0.5, 1.5, 9, 2.5)

ppt.addSlide()
ppt.addTitle("Styled Tables")

ppt.addStyledTable(data, 0.5, 1.3, 9, 2, "1565C0", "E3F2FD")

if ppt.save("demo5_tables.pptx")
    ? "  Created: demo5_tables.pptx"
else
    ? "  FAILED: demo5_tables.pptx"
ok

# Demo 6: Backgrounds
? "Demo 6: Creating presentation with backgrounds..."
ppt = new PPTWriter()
ppt.setTitle("Backgrounds Demo")

ppt.addSlide()
ppt.addTitle("Default White Background")
ppt.addTextBox("This slide has the default white background.", 0.5, 2, 9, 1, NULL)

ppt.addSlide()
ppt.setBackground("1E2761")
ppt.addTextBox("Dark Blue Background", 0.5, 0.3, 9, 0.8, [:fontSize = 36, :bold = true, :color = "FFFFFF"])
ppt.addTextBox("This slide has a dark blue background.", 0.5, 1.5, 9, 1, [:color = "FFFFFF"])

ppt.addSlide()
ppt.setBackground("2C5F2D")
ppt.addTextBox("Forest Green", 0.5, 0.3, 9, 0.8, [:fontSize = 36, :bold = true, :color = "FFFFFF"])
ppt.addTextBox("Nature-inspired color scheme.", 0.5, 1.5, 9, 1, [:color = "97BC62"])

if ppt.save("demo6_backgrounds.pptx")
    ? "  Created: demo6_backgrounds.pptx"
else
    ? "  FAILED: demo6_backgrounds.pptx"
ok

# Demo 7: Two-Column Layout
? "Demo 7: Creating two-column layout..."
ppt = new PPTWriter()
ppt.setTitle("Two Column Layout")

ppt.addSlide()
ppt.addTitle("Two Column Comparison")

ppt.addRectangle(0.3, 1.2, 4.5, 4, [:fillColor = "E3F2FD", :lineColor = "1565C0"])
ppt.addTextBox("Left Panel", 0.5, 1.4, 4, 0.5, [:bold = true, :color = "1565C0"])
ppt.addBulletList(["Feature 1", "Feature 2", "Feature 3"], 0.5, 2, 4, 2, [:fontSize = 14])

ppt.addRectangle(5.2, 1.2, 4.5, 4, [:fillColor = "FFF3E0", :lineColor = "E65100"])
ppt.addTextBox("Right Panel", 5.4, 1.4, 4, 0.5, [:bold = true, :color = "E65100"])
ppt.addBulletList(["Benefit 1", "Benefit 2", "Benefit 3"], 5.4, 2, 4, 2, [:fontSize = 14])

if ppt.save("demo7_two_column.pptx")
    ? "  Created: demo7_two_column.pptx"
else
    ? "  FAILED: demo7_two_column.pptx"
ok

# Demo 8: Rich Text
? "Demo 8: Creating presentation with rich text..."
ppt = new PPTWriter()
ppt.setTitle("Rich Text Demo")

ppt.addSlide()
ppt.addTitle("Rich Text Formatting")

ppt.addRichText([
    ["This sentence has ", []],
    ["bold", [:bold = true]],
    [", ", []],
    ["italic", [:italic = true]],
    [", and ", []],
    ["colored", [:color = "FF0000"]],
    [" text mixed together.", []]
], 0.5, 1.5, 9, 1, NULL)

ppt.addRichText([
    ["Large ", [:fontSize = 28]],
    ["and ", [:fontSize = 18]],
    ["small ", [:fontSize = 12]],
    ["text in one line.", [:fontSize = 18]]
], 0.5, 2.5, 9, 1, NULL)

if ppt.save("demo8_rich_text.pptx")
    ? "  Created: demo8_rich_text.pptx"
else
    ? "  FAILED: demo8_rich_text.pptx"
ok

# Demo 9: Images
? "Demo 9: Creating presentation with images..."
ppt = new PPTWriter()
ppt.setTitle("Images Demo")

ppt.addSlide()
ppt.addTitle("Images in Presentations")

if fexists("images/test1.png")
    ppt.addImage("images/test1.png", 0.5, 1.5, 3, 2.5)
    ppt.addTextBox("test1.png", 0.5, 4.1, 3, 0.4, [:align = "center", :fontSize = 12])
else
    ppt.addTextBox("(test1.png not found)", 0.5, 2.5, 3, 0.5, [:italic = true])
ok

if fexists("images/test2.jpg")
    ppt.addImage("images/test2.jpg", 3.7, 1.5, 3, 2.5)
    ppt.addTextBox("test2.jpg", 3.7, 4.1, 3, 0.4, [:align = "center", :fontSize = 12])
else
    ppt.addTextBox("(test2.jpg not found)", 3.7, 2.5, 3, 0.5, [:italic = true])
ok

if fexists("images/test3.bmp")
    ppt.addImage("images/test3.bmp", 6.9, 1.5, 3, 2.5)
    ppt.addTextBox("test3.bmp", 6.9, 4.1, 3, 0.4, [:align = "center", :fontSize = 12])
else
    ppt.addTextBox("(test3.bmp not found)", 6.9, 2.5, 3, 0.5, [:italic = true])
ok

if ppt.save("demo9_images.pptx")
    ? "  Created: demo9_images.pptx"
else
    ? "  FAILED: demo9_images.pptx"
ok

# Demo 10: Complete Business Presentation
? "Demo 10: Creating complete business presentation..."
ppt = new PPTWriter()
ppt.setTitle("Q4 Business Review")
ppt.setAuthor("Business Team")
ppt.setCompany("ABC Corporation")

# Title slide
ppt.addSlide()
ppt.setBackground("1E2761")
ppt.addTextBox("Q4 2025", 0.5, 1.5, 9, 0.8, [:fontSize = 48, :bold = true, :color = "FFFFFF", :align = "center"])
ppt.addTextBox("Business Review", 0.5, 2.4, 9, 0.6, [:fontSize = 32, :color = "CADCFC", :align = "center"])
ppt.addTextBox("ABC Corporation", 0.5, 4.5, 9, 0.4, [:fontSize = 18, :color = "FFFFFF", :align = "center"])

# Agenda slide
ppt.addSlide()
ppt.addTitle("Agenda")
ppt.addNumberedList([
    "Financial Highlights",
    "Sales Performance", 
    "Regional Breakdown",
    "Key Achievements",
    "2025 Outlook"
], 0.5, 1.3, 9, 3.5, [:fontSize = 20])

# Financial Highlights
ppt.addSlide()
ppt.addTitle("Financial Highlights")

ppt.addRectangle(0.5, 1.3, 2.8, 1.8, [:fillColor = "E3F2FD", :lineColor = "1565C0"])
ppt.addTextBox("Revenue", 0.7, 1.5, 2.4, 0.4, [:fontSize = 14, :color = "1565C0"])
ppt.addTextBox("$12.5M", 0.7, 1.9, 2.4, 0.6, [:fontSize = 28, :bold = true, :color = "1565C0"])
ppt.addTextBox("+18% YoY", 0.7, 2.6, 2.4, 0.4, [:fontSize = 12, :color = "2E7D32"])

ppt.addRectangle(3.6, 1.3, 2.8, 1.8, [:fillColor = "E8F5E9", :lineColor = "2E7D32"])
ppt.addTextBox("Profit", 3.8, 1.5, 2.4, 0.4, [:fontSize = 14, :color = "2E7D32"])
ppt.addTextBox("$3.2M", 3.8, 1.9, 2.4, 0.6, [:fontSize = 28, :bold = true, :color = "2E7D32"])
ppt.addTextBox("+22% YoY", 3.8, 2.6, 2.4, 0.4, [:fontSize = 12, :color = "2E7D32"])

ppt.addRectangle(6.7, 1.3, 2.8, 1.8, [:fillColor = "FFF3E0", :lineColor = "E65100"])
ppt.addTextBox("Customers", 6.9, 1.5, 2.4, 0.4, [:fontSize = 14, :color = "E65100"])
ppt.addTextBox("2,450", 6.9, 1.9, 2.4, 0.6, [:fontSize = 28, :bold = true, :color = "E65100"])
ppt.addTextBox("+340 new", 6.9, 2.6, 2.4, 0.4, [:fontSize = 12, :color = "2E7D32"])

# Sales Table
ppt.addSlide()
ppt.addTitle("Sales Performance")

salesData = [
    ["Quarter", "Target", "Actual", "Variance"],
    ["Q1", "$2.5M", "$2.7M", "+8%"],
    ["Q2", "$2.8M", "$3.0M", "+7%"],
    ["Q3", "$3.0M", "$3.2M", "+7%"],
    ["Q4", "$3.2M", "$3.6M", "+12%"]
]
ppt.addStyledTable(salesData, 0.5, 1.3, 9, 2.8, "4472C4", "D9E2F3")

# Key Achievements
ppt.addSlide()
ppt.addTitle("Key Achievements")
ppt.addBulletList([
    "Launched 3 new products successfully",
    "Expanded into 5 new markets",
    "Achieved 98% customer satisfaction",
    "Reduced operational costs by 12%",
    "Hired 50+ talented team members"
], 0.5, 1.3, 9, 3.5, [:fontSize = 20])

# Thank You slide
ppt.addSlide()
ppt.setBackground("1E2761")
ppt.addTextBox("Thank You!", 0.5, 2, 9, 1, [:fontSize = 48, :bold = true, :color = "FFFFFF", :align = "center"])
ppt.addTextBox("Questions?", 0.5, 3.2, 9, 0.6, [:fontSize = 24, :color = "CADCFC", :align = "center"])

if ppt.save("demo10_business.pptx")
    ? "  Created: demo10_business.pptx"
else
    ? "  FAILED: demo10_business.pptx"
ok

# Demo 11: Quick Function
? "Demo 11: Using quick function..."
slides = [
    ["Welcome", "This presentation was created with quickPPT function"],
    ["Features", "Simple and fast presentation creation"],
    ["Conclusion", "Thank you for using PPTXLib!"]
]
quickPPT("demo11_quick.pptx", "Quick Presentation", slides)
? "  Created: demo11_quick.pptx"

# Demo 12: Different Layout (4:3)
? "Demo 12: Creating with 4:3 layout..."
ppt = new PPTWriter()
ppt.setLayout(PPT_LAYOUT_4x3)
ppt.setTitle("4:3 Layout")

ppt.addTitleSlide("4:3 Aspect Ratio", "Traditional presentation format")

ppt.addSlide()
ppt.addTitle("Standard 4:3 Layout")
ppt.addTextBox("This presentation uses the classic 4:3 aspect ratio.", 0.5, 1.5, 9, 2, [:fontSize = 18])

if ppt.save("demo12_layout_4x3.pptx")
    ? "  Created: demo12_layout_4x3.pptx"
else
    ? "  FAILED: demo12_layout_4x3.pptx"
ok

? ""
? "=============================================="
? "   All demos completed!"
? "=============================================="
? ""
? "Created files:"
? "  1. demo1_simple.pptx"
? "  2. demo2_text_formatting.pptx"
? "  3. demo3_shapes.pptx"
? "  4. demo4_lists.pptx"
? "  5. demo5_tables.pptx"
? "  6. demo6_backgrounds.pptx"
? "  7. demo7_two_column.pptx"
? "  8. demo8_rich_text.pptx"
? "  9. demo9_images.pptx"
? "  10. demo10_business.pptx"
? "  11. demo11_quick.pptx"
? "  12. demo12_layout_4x3.pptx"
