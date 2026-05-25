/*
    PPTXLib - PowerPoint Library for Ring Programming Language
*/

# ============================================================================
# Helper Functions
# ============================================================================

func pptGetSep
    if isWindows()
        return "\"
    else
        return "/"
    ok

func pptXmlEsc str
    str = "" + str
    str = substr(str, "&", "&amp;")
    str = substr(str, "<", "&lt;")
    str = substr(str, ">", "&gt;")
    str = substr(str, '"', "&quot;")
    str = substr(str, "'", "&apos;")
    return str

func pptMakeDir path
    if isWindows()
        path = substr(path, "/", "\")
        system('mkdir "' + path + '" 2>nul')
    else
        system("mkdir -p '" + path + "'")
    ok

func pptColorToHex color
    if color = NULL return "000000" ok
    color = "" + color
    
    color = lower(color)
    
    if aCommonColors[color] != NULL
        return aCommonColors[color]
    ok
    
    if left(color, 1) = "#"
        color = substr(color, 2)
    ok
    
    return upper(color)

func pptEmuFromInches inches
    # Convert inches to EMUs (English Metric Units)
    # 1 inch = 914400 EMUs
    return floor(inches * 914400)

func pptEmuFromCm cm
    # 1 cm = 360000 EMUs
    return floor(cm * 360000)

func pptGetImageExtension filepath
    dotPos = 0
    fpLen = len(filepath)
    for i = 1 to fpLen
        if substr(filepath, i, 1) = "."
            dotPos = i
        ok
    next
    if dotPos > 0 and dotPos < fpLen
        return lower(substr(filepath, dotPos + 1, fpLen - dotPos))
    ok
    return "png"

func pptGetImageContentType ext
    ext = lower(ext)
    switch ext
        on "png"
            return "image/png"
        on "jpg"
            return "image/jpeg"
        on "jpeg"
            return "image/jpeg"
        on "gif"
            return "image/gif"
        on "bmp"
            return "image/bmp"
        other
            return "image/png"
    off

# Quick functions
func quickPPT filename, title, slides
    /*
        Quick function to create a simple presentation
        slides = [ ["Title", "Content"], ["Title2", "Content2"], ... ]
    */
    ppt = new PPTWriter()
    ppt.setTitle(title)
    
    slidesLen = len(slides)
    for i = 1 to slidesLen
        slideData = slides[i]
        ppt.addSlide()
        if len(slideData) >= 1
            ppt.addTitle(slideData[1])
        ok
        if len(slideData) >= 2
            ppt.addTextBox(slideData[2], 0.5, 1.5, 9, 4, NULL)
        ok
    next
    
    return ppt.save(filename)

# ============================================================================
# ZIP Functions for PowerPoint-Compatible ZIP Creation
# ============================================================================

func pptZipInitCRC
    if len(aPPTZipCRC32Table) > 0
        return
    ok
    
    for i = 0 to 255
        crc = i
        for j = 1 to 8
            if (crc & 1) = 1
                crc = (crc >> 1) ^ 0xEDB88320
            else
                crc = crc >> 1
            ok
        next
        aPPTZipCRC32Table + crc
    next

func pptZipCRC32 data
    pptZipInitCRC()
    
    crc = 0xFFFFFFFF
    dataLen = len(data)
    for i = 1 to dataLen
        b = ascii(data[i])
        idx = ((crc ^ b) & 0xFF) + 1
        crc = (crc >> 8) ^ aPPTZipCRC32Table[idx]
    next
    
    return crc ^ 0xFFFFFFFF

func pptZipWord value
    return char(value & 0xFF) + char((value >> 8) & 0xFF)

func pptZipDWord value
    return char(value & 0xFF) + char((value >> 8) & 0xFF) + char((value >> 16) & 0xFF) + char((value >> 24) & 0xFF)

func pptZipCreateFile filename, filesList
    output = ""
    centralDir = ""
    offset = 0
    
    filesCount = len(filesList)
    for i = 1 to filesCount
        entry = filesList[i]
        zipPath = entry[1]
        content = entry[2]
        
        crc32 = pptZipCRC32(content)
        uncompSize = len(content)
        compMethod = 0
        compSize = uncompSize
        fileData = content
        
        dosTime = 0x0000
        dosDate = 0x0021
        
        localHeader = ""
        localHeader += char(0x50) + char(0x4B) + char(0x03) + char(0x04)
        localHeader += pptZipWord(20)
        localHeader += pptZipWord(0)
        localHeader += pptZipWord(compMethod)
        localHeader += pptZipWord(dosTime)
        localHeader += pptZipWord(dosDate)
        localHeader += pptZipDWord(crc32)
        localHeader += pptZipDWord(compSize)
        localHeader += pptZipDWord(uncompSize)
        localHeader += pptZipWord(len(zipPath))
        localHeader += pptZipWord(0)
        localHeader += zipPath
        
        output += localHeader
        output += fileData
        
        centralEntry = ""
        centralEntry += char(0x50) + char(0x4B) + char(0x01) + char(0x02)
        centralEntry += pptZipWord(20)
        centralEntry += pptZipWord(20)
        centralEntry += pptZipWord(0)
        centralEntry += pptZipWord(compMethod)
        centralEntry += pptZipWord(dosTime)
        centralEntry += pptZipWord(dosDate)
        centralEntry += pptZipDWord(crc32)
        centralEntry += pptZipDWord(compSize)
        centralEntry += pptZipDWord(uncompSize)
        centralEntry += pptZipWord(len(zipPath))
        centralEntry += pptZipWord(0)
        centralEntry += pptZipWord(0)
        centralEntry += pptZipWord(0)
        centralEntry += pptZipWord(0)
        centralEntry += pptZipDWord(0)
        centralEntry += pptZipDWord(offset)
        centralEntry += zipPath
        
        centralDir += centralEntry
        offset += len(localHeader) + len(fileData)
    next
    
    eocd = ""
    eocd += char(0x50) + char(0x4B) + char(0x05) + char(0x06)
    eocd += pptZipWord(0)
    eocd += pptZipWord(0)
    eocd += pptZipWord(filesCount)
    eocd += pptZipWord(filesCount)
    eocd += pptZipDWord(len(centralDir))
    eocd += pptZipDWord(offset)
    eocd += pptZipWord(0)
    
    output += centralDir
    output += eocd
    
    write(filename, output)
    return fexists(filename)

