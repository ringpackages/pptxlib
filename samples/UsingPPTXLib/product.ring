load "pptxlib.ring"

cFileName = substr(filename(),".ring",".pptx")
? "Generate File: " + cFileName

ppt = new PPTWriter()

ppt.addTitleSlide("Product Launch", "New Features Overview")

ppt.addSlide()
ppt.addTitle("Key Features")
ppt.addBulletList([
    "Lightning-fast performance",
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

ppt.save(cFileName)