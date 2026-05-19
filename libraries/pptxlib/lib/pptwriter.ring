/*
    PPTXLib - PowerPoint Library for Ring Programming Language
*/

# ============================================================================
# PPTWriter Class
# ============================================================================

class PPTWriter

    # Presentation settings
    cLayout
    nSlideWidth
    nSlideHeight
    
    # Slides
    aSlides
    nCurrentSlide
    
    # Images
    aImages
    nImageId
    
    # Document properties
    cAuthor
    cTitle
    cSubject
    cCompany
    
    # Relationships counter
    nRelId
    
    func init
        # Default to 16:9 layout (10" x 5.625")
        cLayout = PPT_LAYOUT_16x9
        nSlideWidth = pptEmuFromInches(10)
        nSlideHeight = pptEmuFromInches(5.625)
        
        aSlides = []
        nCurrentSlide = 0
        aImages = []
        nImageId = 0
        nRelId = 10
        
        cAuthor = "PPTXLib"
        cTitle = "Presentation"
        cSubject = ""
        cCompany = ""
        
        return self
    
    # ========================================================================
    # Document Properties
    # ========================================================================
    
    func setAuthor author
        cAuthor = author
        return self
    
    func setTitle title
        cTitle = title
        return self
    
    func setSubject subject
        cSubject = subject
        return self
    
    func setCompany company
        cCompany = company
        return self
    
    # ========================================================================
    # Layout Settings
    # ========================================================================
    
    func setLayout layout
        cLayout = layout
        switch layout
            on PPT_LAYOUT_16x9
                nSlideWidth = pptEmuFromInches(10)
                nSlideHeight = pptEmuFromInches(5.625)
            on PPT_LAYOUT_16x10
                nSlideWidth = pptEmuFromInches(10)
                nSlideHeight = pptEmuFromInches(6.25)
            on PPT_LAYOUT_4x3
                nSlideWidth = pptEmuFromInches(10)
                nSlideHeight = pptEmuFromInches(7.5)
        off
        return self
    
    func setCustomSize widthInches, heightInches
        nSlideWidth = pptEmuFromInches(widthInches)
        nSlideHeight = pptEmuFromInches(heightInches)
        return self
    
    # ========================================================================
    # Slide Management
    # ========================================================================
    
    func addSlide
        slide = [
            :elements = [],
            :background = NULL,
            :notes = "",
            :transition = NULL
        ]
        aSlides + slide
        nCurrentSlide = len(aSlides)
        return self
    
    func addTitleSlide title, subtitle
        addSlide()
        
        # Title
        addTextBox(title, 0.5, 2, 9, 1.2, [
            :fontSize = 44,
            :bold = true,
            :align = "center"
        ])
        
        # Subtitle
        if subtitle != NULL and len(subtitle) > 0
            addTextBox(subtitle, 0.5, 3.3, 9, 0.8, [
                :fontSize = 24,
                :align = "center",
                :color = "666666"
            ])
        ok
        
        return self
    
    func addContentSlide title, content
        addSlide()
        addTitle(title)
        addTextBox(content, 0.5, 1.5, 9, 4, [:fontSize = 18])
        return self
    
    func addTwoColumnSlide title, leftContent, rightContent
        addSlide()
        addTitle(title)
        addTextBox(leftContent, 0.5, 1.5, 4.25, 4, [:fontSize = 16])
        addTextBox(rightContent, 5.25, 1.5, 4.25, 4, [:fontSize = 16])
        return self
    
    func selectSlide index
        if index >= 1 and index <= len(aSlides)
            nCurrentSlide = index
        ok
        return self
    
    func getSlideCount
        return len(aSlides)
    
    # ========================================================================
    # Slide Background
    # ========================================================================
    
    func setBackground color
        if nCurrentSlide = 0 return self ok
        aSlides[nCurrentSlide][:background] = [:type = "solid", :color = pptColorToHex(color)]
        return self
    
    func setBackgroundImage imagePath
        if nCurrentSlide = 0 return self ok
        if !fexists(imagePath)
            ? "Warning: Background image not found: " + imagePath
            return self
        ok
        
        imageData = read(imagePath)
        ext = pptGetImageExtension(imagePath)
        
        nImageId++
        img = [
            :id = nImageId,
            :filename = "image" + nImageId + "." + ext,
            :data = imageData,
            :contentType = pptGetImageContentType(ext),
            :isBackground = true
        ]
        aImages + img
        
        aSlides[nCurrentSlide][:background] = [:type = "image", :imageId = nImageId]
        return self
    
    # ========================================================================
    # Text Elements
    # ========================================================================
    
    func addTitle text
        return addTextBox(text, 0.5, 0.3, 9, 0.8, [
            :fontSize = 36,
            :bold = true,
            :align = "left"
        ])
    
    func addSubtitle text
        return addTextBox(text, 0.5, 1.0, 9, 0.5, [
            :fontSize = 20,
            :color = "666666",
            :align = "left"
        ])
    
    func addTextBox text, x, y, w, h, options
        if nCurrentSlide = 0 return self ok
        if options = NULL options = [] ok
        
        element = [
            :type = "textbox",
            :text = text,
            :x = pptEmuFromInches(x),
            :y = pptEmuFromInches(y),
            :w = pptEmuFromInches(w),
            :h = pptEmuFromInches(h),
            :fontSize = 18,
            :fontName = "Calibri",
            :bold = false,
            :italic = false,
            :underline = false,
            :color = "000000",
            :align = "left",
            :valign = "top",
            :bgColor = NULL,
            :borderColor = NULL,
            :borderWidth = 0
        ]
        
        # Apply options
        if options[:fontSize] != NULL element[:fontSize] = options[:fontSize] ok
        if options[:fontName] != NULL element[:fontName] = options[:fontName] ok
        if options[:bold] = true element[:bold] = true ok
        if options[:italic] = true element[:italic] = true ok
        if options[:underline] = true element[:underline] = true ok
        if options[:color] != NULL element[:color] = pptColorToHex(options[:color]) ok
        if options[:align] != NULL element[:align] = options[:align] ok
        if options[:valign] != NULL element[:valign] = options[:valign] ok
        if options[:bgColor] != NULL element[:bgColor] = pptColorToHex(options[:bgColor]) ok
        if options[:borderColor] != NULL element[:borderColor] = pptColorToHex(options[:borderColor]) ok
        if options[:borderWidth] != NULL element[:borderWidth] = options[:borderWidth] ok
        
        aSlides[nCurrentSlide][:elements] + element
        return self
    
    func addRichText runs, x, y, w, h, options
        /*
            Add text with multiple formatting styles
            runs = [ ["text1", [:bold = true]], ["text2", [:italic = true]], ... ]
        */
        if nCurrentSlide = 0 return self ok
        if options = NULL options = [] ok
        
        element = [
            :type = "richtext",
            :runs = runs,
            :x = pptEmuFromInches(x),
            :y = pptEmuFromInches(y),
            :w = pptEmuFromInches(w),
            :h = pptEmuFromInches(h),
            :align = "left",
            :valign = "top"
        ]
        
        if options[:align] != NULL element[:align] = options[:align] ok
        if options[:valign] != NULL element[:valign] = options[:valign] ok
        
        aSlides[nCurrentSlide][:elements] + element
        return self
    
    # ========================================================================
    # Lists
    # ========================================================================
    
    func addBulletList items, x, y, w, h, options
        if nCurrentSlide = 0 return self ok
        if options = NULL options = [] ok
        
        element = [
            :type = "bulletlist",
            :items = items,
            :x = pptEmuFromInches(x),
            :y = pptEmuFromInches(y),
            :w = pptEmuFromInches(w),
            :h = pptEmuFromInches(h),
            :fontSize = 18,
            :color = "000000",
            :bulletColor = NULL
        ]
        
        if options[:fontSize] != NULL element[:fontSize] = options[:fontSize] ok
        if options[:color] != NULL element[:color] = pptColorToHex(options[:color]) ok
        if options[:bulletColor] != NULL element[:bulletColor] = pptColorToHex(options[:bulletColor]) ok
        
        aSlides[nCurrentSlide][:elements] + element
        return self
    
    func addNumberedList items, x, y, w, h, options
        if nCurrentSlide = 0 return self ok
        if options = NULL options = [] ok
        
        element = [
            :type = "numberedlist",
            :items = items,
            :x = pptEmuFromInches(x),
            :y = pptEmuFromInches(y),
            :w = pptEmuFromInches(w),
            :h = pptEmuFromInches(h),
            :fontSize = 18,
            :color = "000000"
        ]
        
        if options[:fontSize] != NULL element[:fontSize] = options[:fontSize] ok
        if options[:color] != NULL element[:color] = pptColorToHex(options[:color]) ok
        
        aSlides[nCurrentSlide][:elements] + element
        return self
    
    # ========================================================================
    # Shapes
    # ========================================================================
    
    func addShape shapeType, x, y, w, h, options
        if nCurrentSlide = 0 return self ok
        if options = NULL options = [] ok
        
        element = [
            :type = "shape",
            :shapeType = shapeType,
            :x = pptEmuFromInches(x),
            :y = pptEmuFromInches(y),
            :w = pptEmuFromInches(w),
            :h = pptEmuFromInches(h),
            :fillColor = "4472C4",
            :lineColor = NULL,
            :lineWidth = 1,
            :text = "",
            :fontSize = 14,
            :fontColor = "FFFFFF",
            :align = "center"
        ]
        
        if options[:fillColor] != NULL element[:fillColor] = pptColorToHex(options[:fillColor]) ok
        if options[:lineColor] != NULL element[:lineColor] = pptColorToHex(options[:lineColor]) ok
        if options[:lineWidth] != NULL element[:lineWidth] = options[:lineWidth] ok
        if options[:text] != NULL element[:text] = options[:text] ok
        if options[:fontSize] != NULL element[:fontSize] = options[:fontSize] ok
        if options[:fontColor] != NULL element[:fontColor] = pptColorToHex(options[:fontColor]) ok
        if options[:align] != NULL element[:align] = options[:align] ok
        if options[:noFill] = true element[:fillColor] = NULL ok
        
        aSlides[nCurrentSlide][:elements] + element
        return self
    
    func addRectangle x, y, w, h, options
        return addShape(PPT_SHAPE_RECT, x, y, w, h, options)
    
    func addRoundedRectangle x, y, w, h, options
        return addShape(PPT_SHAPE_ROUND_RECT, x, y, w, h, options)
    
    func addCircle x, y, size, options
        return addShape(PPT_SHAPE_ELLIPSE, x, y, size, size, options)
    
    func addOval x, y, w, h, options
        return addShape(PPT_SHAPE_ELLIPSE, x, y, w, h, options)
    
    func addLine x1, y1, x2, y2, options
        if nCurrentSlide = 0 return self ok
        if options = NULL options = [] ok
        
        element = [
            :type = "line",
            :x1 = pptEmuFromInches(x1),
            :y1 = pptEmuFromInches(y1),
            :x2 = pptEmuFromInches(x2),
            :y2 = pptEmuFromInches(y2),
            :color = "000000",
            :width = 1
        ]
        
        if options[:color] != NULL element[:color] = pptColorToHex(options[:color]) ok
        if options[:width] != NULL element[:width] = options[:width] ok
        
        aSlides[nCurrentSlide][:elements] + element
        return self
    
    # ========================================================================
    # Images
    # ========================================================================
    
    func addImage imagePath, x, y, w, h
        if nCurrentSlide = 0 return self ok
        
        if !fexists(imagePath)
            ? "Warning: Image file not found: " + imagePath
            return self
        ok
        
        imageData = read(imagePath)
        if len(imageData) = 0
            ? "Warning: Could not read image file: " + imagePath
            return self
        ok
        
        ext = pptGetImageExtension(imagePath)
        
        nImageId++
        img = [
            :id = nImageId,
            :filename = "image" + nImageId + "." + ext,
            :data = imageData,
            :contentType = pptGetImageContentType(ext),
            :slide = nCurrentSlide,
            :isBackground = false
        ]
        aImages + img
        
        element = [
            :type = "image",
            :imageId = nImageId,
            :x = pptEmuFromInches(x),
            :y = pptEmuFromInches(y),
            :w = pptEmuFromInches(w),
            :h = pptEmuFromInches(h)
        ]
        
        aSlides[nCurrentSlide][:elements] + element
        return self
    
    func addImageCentered imagePath, y, w, h
        # Center horizontally
        slideWidthInches = nSlideWidth / 914400
        x = (slideWidthInches - w) / 2
        return addImage(imagePath, x, y, w, h)
    
    # ========================================================================
    # Tables
    # ========================================================================
    
    func addTable data, x, y, w, h, options
        if nCurrentSlide = 0 return self ok
        if options = NULL options = [] ok
        
        element = [
            :type = "table",
            :data = data,
            :x = pptEmuFromInches(x),
            :y = pptEmuFromInches(y),
            :w = pptEmuFromInches(w),
            :h = pptEmuFromInches(h),
            :headerRow = false,
            :headerBgColor = "4472C4",
            :headerFontColor = "FFFFFF",
            :borderColor = "000000",
            :borderWidth = 1,
            :evenRowBgColor = NULL,
            :fontSize = 14
        ]
        
        if options[:headerRow] = true element[:headerRow] = true ok
        if options[:headerBgColor] != NULL element[:headerBgColor] = pptColorToHex(options[:headerBgColor]) ok
        if options[:headerFontColor] != NULL element[:headerFontColor] = pptColorToHex(options[:headerFontColor]) ok
        if options[:borderColor] != NULL element[:borderColor] = pptColorToHex(options[:borderColor]) ok
        if options[:borderWidth] != NULL element[:borderWidth] = options[:borderWidth] ok
        if options[:evenRowBgColor] != NULL element[:evenRowBgColor] = pptColorToHex(options[:evenRowBgColor]) ok
        if options[:fontSize] != NULL element[:fontSize] = options[:fontSize] ok
        
        aSlides[nCurrentSlide][:elements] + element
        return self
    
    func addSimpleTable data, x, y, w, h
        return addTable(data, x, y, w, h, [:headerRow = true])
    
    func addStyledTable data, x, y, w, h, headerColor, evenRowColor
        return addTable(data, x, y, w, h, [
            :headerRow = true,
            :headerBgColor = headerColor,
            :evenRowBgColor = evenRowColor
        ])
    
    # ========================================================================
    # Charts
    # ========================================================================
    
    func addChart chartType, chartData, x, y, w, h, options
        if nCurrentSlide = 0 return self ok
        if options = NULL options = [] ok
        
        element = [
            :type = "chart",
            :chartType = chartType,
            :chartData = chartData,
            :x = pptEmuFromInches(x),
            :y = pptEmuFromInches(y),
            :w = pptEmuFromInches(w),
            :h = pptEmuFromInches(h),
            :title = "",
            :showLegend = true,
            :colors = ["4472C4", "ED7D31", "A5A5A5", "FFC000", "5B9BD5"]
        ]
        
        if options[:title] != NULL element[:title] = options[:title] ok
        if options[:showLegend] = false element[:showLegend] = false ok
        if options[:colors] != NULL element[:colors] = options[:colors] ok
        
        aSlides[nCurrentSlide][:elements] + element
        return self
    
    func addBarChart labels, values, x, y, w, h, options
        chartData = [:labels = labels, :values = values]
        return addChart(PPT_CHART_BAR, chartData, x, y, w, h, options)
    
    func addLineChart labels, values, x, y, w, h, options
        chartData = [:labels = labels, :values = values]
        return addChart(PPT_CHART_LINE, chartData, x, y, w, h, options)
    
    func addPieChart labels, values, x, y, w, h, options
        chartData = [:labels = labels, :values = values]
        return addChart(PPT_CHART_PIE, chartData, x, y, w, h, options)
    
    # ========================================================================
    # Speaker Notes
    # ========================================================================
    
    func addNotes text
        if nCurrentSlide = 0 return self ok
        aSlides[nCurrentSlide][:notes] = text
        return self
    
    # ========================================================================
    # Slide Transitions
    # ========================================================================
    
    func setTransition transitionType, duration
        if nCurrentSlide = 0 return self ok
        if duration = NULL duration = 1000 ok
        aSlides[nCurrentSlide][:transition] = [:type = transitionType, :duration = duration]
        return self
    
    # ========================================================================
    # Save Document
    # ========================================================================
    
    func save filename
        sep = pptGetSep()
        
        # Create temp directory structure
        tempDir = filename + "_temp" + sep
        pptMakeDir(tempDir)
        pptMakeDir(tempDir + "_rels")
        pptMakeDir(tempDir + "docProps")
        pptMakeDir(tempDir + "ppt")
        pptMakeDir(tempDir + "ppt" + sep + "_rels")
        pptMakeDir(tempDir + "ppt" + sep + "slides")
        pptMakeDir(tempDir + "ppt" + sep + "slides" + sep + "_rels")
        pptMakeDir(tempDir + "ppt" + sep + "slideLayouts")
        pptMakeDir(tempDir + "ppt" + sep + "slideLayouts" + sep + "_rels")
        pptMakeDir(tempDir + "ppt" + sep + "slideMasters")
        pptMakeDir(tempDir + "ppt" + sep + "slideMasters" + sep + "_rels")
        pptMakeDir(tempDir + "ppt" + sep + "theme")
        
        # Create media folder if we have images
        if len(aImages) > 0
            pptMakeDir(tempDir + "ppt" + sep + "media")
        ok
        
        # Write all XML files
        writeContentTypes(tempDir)
        writeRels(tempDir)
        writeCore(tempDir)
        writeApp(tempDir)
        writePresentation(tempDir)
        writePresentationRels(tempDir)
        writeTheme(tempDir)
        writeSlideMaster(tempDir)
        writeSlideMasterRels(tempDir)
        writeSlideLayout(tempDir)
        writeSlideLayoutRels(tempDir)
        
        # Write slides
        slidesLen = len(aSlides)
        for i = 1 to slidesLen
            writeSlide(tempDir, i)
            writeSlideRels(tempDir, i)
        next
        
        # Create ZIP file
        result = createZip(tempDir, filename)
        
        # Clean up temp directory
        if isWindows()
            system('rmdir /s /q "' + tempDir + '" 2>nul')
        else
            system("rm -rf '" + tempDir + "'")
        ok
        
        return result
    
    # ========================================================================
    # XML Generation Methods (Private)
    # ========================================================================
    
    func writeContentTypes tempDir
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        c += '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        c += '<Default Extension="xml" ContentType="application/xml"/>'
        
        # Image types
        imagesLen = len(aImages)
        hasPng = false
        hasJpg = false
        hasGif = false
        hasBmp = false
        
        for i = 1 to imagesLen
            ct = aImages[i][:contentType]
            if ct = "image/png" and !hasPng
                c += '<Default Extension="png" ContentType="image/png"/>'
                hasPng = true
            ok
            if ct = "image/jpeg" and !hasJpg
                c += '<Default Extension="jpg" ContentType="image/jpeg"/>'
                c += '<Default Extension="jpeg" ContentType="image/jpeg"/>'
                hasJpg = true
            ok
            if ct = "image/gif" and !hasGif
                c += '<Default Extension="gif" ContentType="image/gif"/>'
                hasGif = true
            ok
            if ct = "image/bmp" and !hasBmp
                c += '<Default Extension="bmp" ContentType="image/bmp"/>'
                hasBmp = true
            ok
        next
        
        c += '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
        c += '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
        c += '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
        c += '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
        
        slidesLen = len(aSlides)
        for i = 1 to slidesLen
            c += '<Override PartName="/ppt/slides/slide' + i + '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        next
        
        c += '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        c += '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        c += '</Types>'
        
        write(tempDir + "[Content_Types].xml", c)
    
    func writeRels tempDir
        sep = pptGetSep()
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        c += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'
        c += '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        c += '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
        c += '</Relationships>'
        
        write(tempDir + "_rels" + sep + ".rels", c)
    
    func writeCore tempDir
        sep = pptGetSep()
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        c += 'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
        c += 'xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        c += '<dc:title>' + pptXmlEsc(cTitle) + '</dc:title>'
        c += '<dc:subject>' + pptXmlEsc(cSubject) + '</dc:subject>'
        c += '<dc:creator>' + pptXmlEsc(cAuthor) + '</dc:creator>'
        c += '<cp:lastModifiedBy>' + pptXmlEsc(cAuthor) + '</cp:lastModifiedBy>'
        c += '<dcterms:created xsi:type="dcterms:W3CDTF">2025-01-01T00:00:00Z</dcterms:created>'
        c += '<dcterms:modified xsi:type="dcterms:W3CDTF">2025-01-01T00:00:00Z</dcterms:modified>'
        c += '</cp:coreProperties>'
        
        write(tempDir + "docProps" + sep + "core.xml", c)
    
    func writeApp tempDir
        sep = pptGetSep()
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
        c += '<Application>PPTXLib</Application>'
        c += '<AppVersion>1.0</AppVersion>'
        c += '<Slides>' + len(aSlides) + '</Slides>'
        if len(cCompany) > 0
            c += '<Company>' + pptXmlEsc(cCompany) + '</Company>'
        ok
        c += '</Properties>'
        
        write(tempDir + "docProps" + sep + "app.xml", c)
    
    func writePresentation tempDir
        sep = pptGetSep()
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        c += 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        c += 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1">'
        c += '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
        c += '<p:sldIdLst>'
        
        slidesLen = len(aSlides)
        for i = 1 to slidesLen
            c += '<p:sldId id="' + (255 + i) + '" r:id="rId' + (i + 3) + '"/>'
        next
        
        c += '</p:sldIdLst>'
        c += '<p:sldSz cx="' + nSlideWidth + '" cy="' + nSlideHeight + '"/>'
        c += '<p:notesSz cx="' + nSlideHeight + '" cy="' + nSlideWidth + '"/>'
        c += '</p:presentation>'
        
        write(tempDir + "ppt" + sep + "presentation.xml", c)
    
    func writePresentationRels tempDir
        sep = pptGetSep()
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        c += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>'
        c += '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
        
        slidesLen = len(aSlides)
        for i = 1 to slidesLen
            c += '<Relationship Id="rId' + (i + 3) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide' + i + '.xml"/>'
        next
        
        c += '</Relationships>'
        
        write(tempDir + "ppt" + sep + "_rels" + sep + "presentation.xml.rels", c)
    
    func writeTheme tempDir
        sep = pptGetSep()
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">'
        c += '<a:themeElements>'
        c += '<a:clrScheme name="Office">'
        c += '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
        c += '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
        c += '<a:dk2><a:srgbClr val="44546A"/></a:dk2>'
        c += '<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>'
        c += '<a:accent1><a:srgbClr val="4472C4"/></a:accent1>'
        c += '<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>'
        c += '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>'
        c += '<a:accent4><a:srgbClr val="FFC000"/></a:accent4>'
        c += '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>'
        c += '<a:accent6><a:srgbClr val="70AD47"/></a:accent6>'
        c += '<a:hlink><a:srgbClr val="0563C1"/></a:hlink>'
        c += '<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>'
        c += '</a:clrScheme>'
        c += '<a:fontScheme name="Office">'
        c += '<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
        c += '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>'
        c += '</a:fontScheme>'
        c += '<a:fmtScheme name="Office">'
        c += '<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>'
        c += '<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>'
        c += '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
        c += '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>'
        c += '</a:fmtScheme>'
        c += '</a:themeElements>'
        c += '</a:theme>'
        
        write(tempDir + "ppt" + sep + "theme" + sep + "theme1.xml", c)
    
    func writeSlideMaster tempDir
        sep = pptGetSep()
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        c += 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        c += 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        c += '<p:cSld><p:spTree>'
        c += '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        c += '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        c += '</p:spTree></p:cSld>'
        c += '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
        c += '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>'
        c += '</p:sldMaster>'
        
        write(tempDir + "ppt" + sep + "slideMasters" + sep + "slideMaster1.xml", c)
    
    func writeSlideMasterRels tempDir
        sep = pptGetSep()
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        c += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
        c += '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>'
        c += '</Relationships>'
        
        write(tempDir + "ppt" + sep + "slideMasters" + sep + "_rels" + sep + "slideMaster1.xml.rels", c)
    
    func writeSlideLayout tempDir
        sep = pptGetSep()
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        c += 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        c += 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">'
        c += '<p:cSld name="Blank"><p:spTree>'
        c += '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        c += '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        c += '</p:spTree></p:cSld>'
        c += '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
        c += '</p:sldLayout>'
        
        write(tempDir + "ppt" + sep + "slideLayouts" + sep + "slideLayout1.xml", c)
    
    func writeSlideLayoutRels tempDir
        sep = pptGetSep()
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        c += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>'
        c += '</Relationships>'
        
        write(tempDir + "ppt" + sep + "slideLayouts" + sep + "_rels" + sep + "slideLayout1.xml.rels", c)
    
    func writeSlide tempDir, slideIndex
        sep = pptGetSep()
        slide = aSlides[slideIndex]
        elements = slide[:elements]
        background = slide[:background]
        
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        c += 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        c += 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        c += '<p:cSld>'
        
        # Background
        if background != NULL
            if background[:type] = "solid"
                c += '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="' + background[:color] + '"/></a:solidFill></p:bgPr></p:bg>'
            ok
        ok
        
        c += '<p:spTree>'
        c += '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        c += '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        
        # Render elements
        elemId = 2
        elementsLen = len(elements)
        for i = 1 to elementsLen
            elem = elements[i]
            elemType = elem[:type]
            
            switch elemType
                on "textbox"
                    c += renderTextBox(elem, elemId)
                    elemId++
                on "richtext"
                    c += renderRichText(elem, elemId)
                    elemId++
                on "bulletlist"
                    c += renderBulletList(elem, elemId)
                    elemId++
                on "numberedlist"
                    c += renderNumberedList(elem, elemId)
                    elemId++
                on "shape"
                    c += renderShape(elem, elemId)
                    elemId++
                on "line"
                    c += renderLine(elem, elemId)
                    elemId++
                on "image"
                    c += renderImage(elem, elemId, slideIndex)
                    elemId++
                on "table"
                    c += renderTable(elem, elemId)
                    elemId++
            off
        next
        
        c += '</p:spTree></p:cSld>'
        c += '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
        c += '</p:sld>'
        
        write(tempDir + "ppt" + sep + "slides" + sep + "slide" + slideIndex + ".xml", c)
    
    func writeSlideRels tempDir, slideIndex
        sep = pptGetSep()
        slide = aSlides[slideIndex]
        elements = slide[:elements]
        
        c = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        c += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        c += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
        
        # Add image relationships
        relId = 2
        elementsLen = len(elements)
        for i = 1 to elementsLen
            elem = elements[i]
            if elem[:type] = "image"
                imgId = elem[:imageId]
                # Find image filename
                imagesLen = len(aImages)
                for j = 1 to imagesLen
                    if aImages[j][:id] = imgId
                        c += '<Relationship Id="rId' + relId + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/' + aImages[j][:filename] + '"/>'
                        relId++
                        exit
                    ok
                next
            ok
        next
        
        c += '</Relationships>'
        
        write(tempDir + "ppt" + sep + "slides" + sep + "_rels" + sep + "slide" + slideIndex + ".xml.rels", c)
    
    # ========================================================================
    # Element Rendering Methods
    # ========================================================================
    
    func renderTextBox elem, elemId
        c = '<p:sp>'
        c += '<p:nvSpPr><p:cNvPr id="' + elemId + '" name="TextBox ' + elemId + '"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        c += '<p:spPr>'
        c += '<a:xfrm><a:off x="' + elem[:x] + '" y="' + elem[:y] + '"/><a:ext cx="' + elem[:w] + '" cy="' + elem[:h] + '"/></a:xfrm>'
        c += '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        
        if elem[:bgColor] != NULL
            c += '<a:solidFill><a:srgbClr val="' + elem[:bgColor] + '"/></a:solidFill>'
        else
            c += '<a:noFill/>'
        ok
        
        if elem[:borderColor] != NULL and elem[:borderWidth] > 0
            c += '<a:ln w="' + (elem[:borderWidth] * 12700) + '"><a:solidFill><a:srgbClr val="' + elem[:borderColor] + '"/></a:solidFill></a:ln>'
        else
            c += '<a:ln><a:noFill/></a:ln>'
        ok
        
        c += '</p:spPr>'
        c += '<p:txBody>'
        c += '<a:bodyPr wrap="square" rtlCol="0"'
        
        # Vertical alignment
        if elem[:valign] = "middle" or elem[:valign] = "center"
            c += ' anchor="ctr"'
        elseif elem[:valign] = "bottom"
            c += ' anchor="b"'
        else
            c += ' anchor="t"'
        ok
        c += '/>'
        
        c += '<a:lstStyle/>'
        c += '<a:p>'
        
        # Horizontal alignment
        c += '<a:pPr'
        if elem[:align] = "center"
            c += ' algn="ctr"'
        elseif elem[:align] = "right"
            c += ' algn="r"'
        else
            c += ' algn="l"'
        ok
        c += '/>'
        
        c += '<a:r>'
        c += '<a:rPr lang="en-US" sz="' + (elem[:fontSize] * 100) + '"'
        if elem[:bold] c += ' b="1"' ok
        if elem[:italic] c += ' i="1"' ok
        if elem[:underline] c += ' u="sng"' ok
        c += '>'
        c += '<a:solidFill><a:srgbClr val="' + elem[:color] + '"/></a:solidFill>'
        c += '<a:latin typeface="' + elem[:fontName] + '"/>'
        c += '</a:rPr>'
        c += '<a:t>' + pptXmlEsc(elem[:text]) + '</a:t>'
        c += '</a:r>'
        c += '</a:p></p:txBody></p:sp>'
        
        return c
    
    func renderRichText elem, elemId
        c = '<p:sp>'
        c += '<p:nvSpPr><p:cNvPr id="' + elemId + '" name="TextBox ' + elemId + '"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        c += '<p:spPr>'
        c += '<a:xfrm><a:off x="' + elem[:x] + '" y="' + elem[:y] + '"/><a:ext cx="' + elem[:w] + '" cy="' + elem[:h] + '"/></a:xfrm>'
        c += '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln>'
        c += '</p:spPr>'
        c += '<p:txBody><a:bodyPr wrap="square" rtlCol="0"/><a:lstStyle/>'
        c += '<a:p>'
        
        # Alignment
        c += '<a:pPr'
        if elem[:align] = "center"
            c += ' algn="ctr"'
        elseif elem[:align] = "right"
            c += ' algn="r"'
        ok
        c += '/>'
        
        runs = elem[:runs]
        runsLen = len(runs)
        for i = 1 to runsLen
            run = runs[i]
            text = run[1]
            opts = []
            if len(run) > 1 opts = run[2] ok
            
            fontSize = 18
            fontName = "Calibri"
            color = "000000"
            bold = false
            italic = false
            underline = false
            
            if opts[:fontSize] != NULL fontSize = opts[:fontSize] ok
            if opts[:fontName] != NULL fontName = opts[:fontName] ok
            if opts[:color] != NULL color = pptColorToHex(opts[:color]) ok
            if opts[:bold] = true bold = true ok
            if opts[:italic] = true italic = true ok
            if opts[:underline] = true underline = true ok
            
            c += '<a:r>'
            c += '<a:rPr lang="en-US" sz="' + (fontSize * 100) + '"'
            if bold c += ' b="1"' ok
            if italic c += ' i="1"' ok
            if underline c += ' u="sng"' ok
            c += '>'
            c += '<a:solidFill><a:srgbClr val="' + color + '"/></a:solidFill>'
            c += '<a:latin typeface="' + fontName + '"/>'
            c += '</a:rPr>'
            c += '<a:t>' + pptXmlEsc(text) + '</a:t>'
            c += '</a:r>'
        next
        
        c += '</a:p></p:txBody></p:sp>'
        return c
    
    func renderBulletList elem, elemId
        c = '<p:sp>'
        c += '<p:nvSpPr><p:cNvPr id="' + elemId + '" name="TextBox ' + elemId + '"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        c += '<p:spPr>'
        c += '<a:xfrm><a:off x="' + elem[:x] + '" y="' + elem[:y] + '"/><a:ext cx="' + elem[:w] + '" cy="' + elem[:h] + '"/></a:xfrm>'
        c += '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln>'
        c += '</p:spPr>'
        c += '<p:txBody><a:bodyPr wrap="square" rtlCol="0"/><a:lstStyle/>'
        
        items = elem[:items]
        itemsLen = len(items)
        for i = 1 to itemsLen
            c += '<a:p>'
            c += '<a:pPr>'
            c += '<a:buFont typeface="Arial"/>'
            if elem[:bulletColor] != NULL
                c += '<a:buClr><a:srgbClr val="' + elem[:bulletColor] + '"/></a:buClr>'
            ok
            c += '<a:buChar char="•"/>'
            c += '</a:pPr>'
            c += '<a:r>'
            c += '<a:rPr lang="en-US" sz="' + (elem[:fontSize] * 100) + '">'
            c += '<a:solidFill><a:srgbClr val="' + elem[:color] + '"/></a:solidFill>'
            c += '<a:latin typeface="Calibri"/>'
            c += '</a:rPr>'
            c += '<a:t>' + pptXmlEsc(items[i]) + '</a:t>'
            c += '</a:r>'
            c += '</a:p>'
        next
        
        c += '</p:txBody></p:sp>'
        return c
    
    func renderNumberedList elem, elemId
        c = '<p:sp>'
        c += '<p:nvSpPr><p:cNvPr id="' + elemId + '" name="TextBox ' + elemId + '"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        c += '<p:spPr>'
        c += '<a:xfrm><a:off x="' + elem[:x] + '" y="' + elem[:y] + '"/><a:ext cx="' + elem[:w] + '" cy="' + elem[:h] + '"/></a:xfrm>'
        c += '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln>'
        c += '</p:spPr>'
        c += '<p:txBody><a:bodyPr wrap="square" rtlCol="0"/><a:lstStyle/>'
        
        items = elem[:items]
        itemsLen = len(items)
        for i = 1 to itemsLen
            c += '<a:p>'
            c += '<a:pPr>'
            c += '<a:buFont typeface="Arial"/>'
            c += '<a:buAutoNum type="arabicPeriod"/>'
            c += '</a:pPr>'
            c += '<a:r>'
            c += '<a:rPr lang="en-US" sz="' + (elem[:fontSize] * 100) + '">'
            c += '<a:solidFill><a:srgbClr val="' + elem[:color] + '"/></a:solidFill>'
            c += '<a:latin typeface="Calibri"/>'
            c += '</a:rPr>'
            c += '<a:t>' + pptXmlEsc(items[i]) + '</a:t>'
            c += '</a:r>'
            c += '</a:p>'
        next
        
        c += '</p:txBody></p:sp>'
        return c
    
    func renderShape elem, elemId
        c = '<p:sp>'
        c += '<p:nvSpPr><p:cNvPr id="' + elemId + '" name="Shape ' + elemId + '"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        c += '<p:spPr>'
        c += '<a:xfrm><a:off x="' + elem[:x] + '" y="' + elem[:y] + '"/><a:ext cx="' + elem[:w] + '" cy="' + elem[:h] + '"/></a:xfrm>'
        c += '<a:prstGeom prst="' + elem[:shapeType] + '"><a:avLst/></a:prstGeom>'
        
        if elem[:fillColor] != NULL
            c += '<a:solidFill><a:srgbClr val="' + elem[:fillColor] + '"/></a:solidFill>'
        else
            c += '<a:noFill/>'
        ok
        
        if elem[:lineColor] != NULL
            c += '<a:ln w="' + (elem[:lineWidth] * 12700) + '"><a:solidFill><a:srgbClr val="' + elem[:lineColor] + '"/></a:solidFill></a:ln>'
        else
            c += '<a:ln><a:noFill/></a:ln>'
        ok
        
        c += '</p:spPr>'
        
        # Add text if provided
        if len(elem[:text]) > 0
            c += '<p:txBody><a:bodyPr wrap="square" rtlCol="0" anchor="ctr"/><a:lstStyle/>'
            c += '<a:p><a:pPr algn="ctr"/>'
            c += '<a:r><a:rPr lang="en-US" sz="' + (elem[:fontSize] * 100) + '">'
            c += '<a:solidFill><a:srgbClr val="' + elem[:fontColor] + '"/></a:solidFill>'
            c += '<a:latin typeface="Calibri"/>'
            c += '</a:rPr>'
            c += '<a:t>' + pptXmlEsc(elem[:text]) + '</a:t>'
            c += '</a:r></a:p></p:txBody>'
        ok
        
        c += '</p:sp>'
        return c
    
    func renderLine elem, elemId
        c = '<p:cxnSp>'
        c += '<p:nvCxnSpPr><p:cNvPr id="' + elemId + '" name="Line ' + elemId + '"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>'
        c += '<p:spPr>'
        
        # Calculate position and size
        x = elem[:x1]
        y = elem[:y1]
        cx = elem[:x2] - elem[:x1]
        cy = elem[:y2] - elem[:y1]
        
        if cx < 0
            x = elem[:x2]
            cx = -cx
        ok
        if cy < 0
            y = elem[:y2]
            cy = -cy
        ok
        
        c += '<a:xfrm><a:off x="' + x + '" y="' + y + '"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm>'
        c += '<a:prstGeom prst="line"><a:avLst/></a:prstGeom>'
        c += '<a:ln w="' + (elem[:width] * 12700) + '"><a:solidFill><a:srgbClr val="' + elem[:color] + '"/></a:solidFill></a:ln>'
        c += '</p:spPr>'
        c += '</p:cxnSp>'
        return c
    
    func renderImage elem, elemId, slideIndex
        # Find the relationship ID for this image
        relId = 2
        slide = aSlides[slideIndex]
        elements = slide[:elements]
        elementsLen = len(elements)
        for i = 1 to elementsLen
            e = elements[i]
            if e[:type] = "image"
                if e[:imageId] = elem[:imageId]
                    exit
                ok
                relId++
            ok
        next
        
        c = '<p:pic>'
        c += '<p:nvPicPr><p:cNvPr id="' + elemId + '" name="Picture ' + elemId + '"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
        c += '<p:blipFill>'
        c += '<a:blip r:embed="rId' + relId + '"/>'
        c += '<a:stretch><a:fillRect/></a:stretch>'
        c += '</p:blipFill>'
        c += '<p:spPr>'
        c += '<a:xfrm><a:off x="' + elem[:x] + '" y="' + elem[:y] + '"/><a:ext cx="' + elem[:w] + '" cy="' + elem[:h] + '"/></a:xfrm>'
        c += '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        c += '</p:spPr>'
        c += '</p:pic>'
        return c
    
    func renderTable elem, elemId
        data = elem[:data]
        if len(data) = 0 return "" ok
        
        rows = len(data)
        cols = len(data[1])
        
        cellW = elem[:w] / cols
        cellH = elem[:h] / rows
        
        c = '<p:graphicFrame>'
        c += '<p:nvGraphicFramePr><p:cNvPr id="' + elemId + '" name="Table ' + elemId + '"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>'
        c += '<p:xfrm><a:off x="' + elem[:x] + '" y="' + elem[:y] + '"/><a:ext cx="' + elem[:w] + '" cy="' + elem[:h] + '"/></p:xfrm>'
        c += '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
        c += '<a:tbl>'
        c += '<a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr>'
        c += '<a:tblGrid>'
        for col = 1 to cols
            c += '<a:gridCol w="' + cellW + '"/>'
        next
        c += '</a:tblGrid>'
        
        for row = 1 to rows
            c += '<a:tr h="' + cellH + '">'
            rowData = data[row]
            rowDataLen = len(rowData)
            for col = 1 to rowDataLen
                cellValue = rowData[col]
                if !isString(cellValue) cellValue = "" + cellValue ok
                
                c += '<a:tc>'
                c += '<a:txBody><a:bodyPr/><a:lstStyle/>'
                c += '<a:p><a:r>'
                c += '<a:rPr lang="en-US" sz="' + (elem[:fontSize] * 100) + '"'
                
                if elem[:headerRow] and row = 1
                    c += ' b="1">'
                    c += '<a:solidFill><a:srgbClr val="' + elem[:headerFontColor] + '"/></a:solidFill>'
                else
                    c += '>'
                    c += '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>'
                ok
                
                c += '<a:latin typeface="Calibri"/>'
                c += '</a:rPr>'
                c += '<a:t>' + pptXmlEsc(cellValue) + '</a:t>'
                c += '</a:r></a:p></a:txBody>'
                c += '<a:tcPr>'
                
                # Cell background
                if elem[:headerRow] and row = 1
                    c += '<a:solidFill><a:srgbClr val="' + elem[:headerBgColor] + '"/></a:solidFill>'
                elseif elem[:evenRowBgColor] != NULL and row % 2 = 0
                    c += '<a:solidFill><a:srgbClr val="' + elem[:evenRowBgColor] + '"/></a:solidFill>'
                ok
                
                c += '</a:tcPr>'
                c += '</a:tc>'
            next
            c += '</a:tr>'
        next
        
        c += '</a:tbl></a:graphicData></a:graphic></p:graphicFrame>'
        return c
    
    func createZip tempDir, filename
        sep = pptGetSep()
        
        filesList = []
        
        filesList + ["[Content_Types].xml", read(tempDir + "[Content_Types].xml")]
        filesList + ["_rels/.rels", read(tempDir + "_rels" + sep + ".rels")]
        filesList + ["docProps/core.xml", read(tempDir + "docProps" + sep + "core.xml")]
        filesList + ["docProps/app.xml", read(tempDir + "docProps" + sep + "app.xml")]
        filesList + ["ppt/presentation.xml", read(tempDir + "ppt" + sep + "presentation.xml")]
        filesList + ["ppt/_rels/presentation.xml.rels", read(tempDir + "ppt" + sep + "_rels" + sep + "presentation.xml.rels")]
        filesList + ["ppt/theme/theme1.xml", read(tempDir + "ppt" + sep + "theme" + sep + "theme1.xml")]
        filesList + ["ppt/slideMasters/slideMaster1.xml", read(tempDir + "ppt" + sep + "slideMasters" + sep + "slideMaster1.xml")]
        filesList + ["ppt/slideMasters/_rels/slideMaster1.xml.rels", read(tempDir + "ppt" + sep + "slideMasters" + sep + "_rels" + sep + "slideMaster1.xml.rels")]
        filesList + ["ppt/slideLayouts/slideLayout1.xml", read(tempDir + "ppt" + sep + "slideLayouts" + sep + "slideLayout1.xml")]
        filesList + ["ppt/slideLayouts/_rels/slideLayout1.xml.rels", read(tempDir + "ppt" + sep + "slideLayouts" + sep + "_rels" + sep + "slideLayout1.xml.rels")]
        
        # Add slides
        slidesLen = len(aSlides)
        for i = 1 to slidesLen
            filesList + ["ppt/slides/slide" + i + ".xml", read(tempDir + "ppt" + sep + "slides" + sep + "slide" + i + ".xml")]
            filesList + ["ppt/slides/_rels/slide" + i + ".xml.rels", read(tempDir + "ppt" + sep + "slides" + sep + "_rels" + sep + "slide" + i + ".xml.rels")]
        next
        
        # Add images
        imagesLen = len(aImages)
        for i = 1 to imagesLen
            filesList + ["ppt/media/" + aImages[i][:filename], aImages[i][:data]]
        next
        
        return pptZipCreateFile(filename, filesList)
