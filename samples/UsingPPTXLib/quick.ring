load "pptxlib.ring"

cFileName = substr(filename(),".ring",".pptx")
? "Generate File: " + cFileName

slides = [
    ["Welcome", "Introduction text"],
    ["Features", "Feature description"],
    ["Conclusion", "Thank you!"]
]

quickPPT(cFileName, "My Presentation", slides)