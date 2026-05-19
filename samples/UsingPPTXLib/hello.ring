load "pptxlib.ring"

cFileName = substr(filename(),".ring",".pptx")
? "Generate File: " + cFileName

new PPTWriter() {
	addTitleSlide("My Presentation", "Subtitle here")
	addSlide()
	addTitle("First Slide")
	addTextBox("Hello, World!", 0.5, 1.5, 9, 2, NULL)
	save(cFileName)
}