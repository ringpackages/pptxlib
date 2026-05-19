load "pptxlib.ring"

cFileName = substr(filename(),".ring",".pptx")
? "Generate File: " + cFileName

ppt = new PPTWriter()
ppt.setTitle("Q4 Review")
ppt.setAuthor("Sales Team")

# Title slide with dark background
ppt.addSlide()
ppt.setBackground("1E2761")
ppt.addTextBox("Q4 2025", 0.5, 2, 9, 0.8, [
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

ppt.save(cFileName)