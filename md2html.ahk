#Requires AutoHotkey v2.0
#SingleInstance Force

; Markdown to HTML Converter
; This script converts markdown files to HTML and downloads images for local use

; Add a dummy hotkey to keep script running (never triggered)
#HotIf false
F24::
	{
	}
#HotIf

; Global variables
global folderName := ""  ; Will be set based on first headline
global baseOutputDir := ""  ; Will be set by user selection
global processingGui := ""  ; Simple processing window
global iniFile := A_ScriptDir "\md2html.ini"  ; Path to INI file
global currentHotkey := "!m"  ; Default hotkey (Alt+M)
global LogFile := A_ScriptDir "\md2html_log.txt"
global MAX_LOG_SIZE := 1048576  ; 1MB in bytes

;------------------------------------------------------------------------------
; Variables
;------------------------------------------------------------------------------
VarScriptName := "MD2HTML"
VarVersionNo := "v1.0"
VarBlurb := " Press " currentHotkey " or this icon to convert Markdown to HTML."

;------------------------------------------------------------------------------
;Icon Tip
;------------------------------------------------------------------------------
A_IconTip := VarScriptName " " VarVersionNo " " VarBlurb

;------------------------------------------------------------------------------
; Add ICON to your tray
;------------------------------------------------------------------------------
Try {
    TraySetIcon(A_ScriptDir "\" VarScriptName ".ico")
}
Catch {
    TrayTip "Remember to add " VarScriptName ".ico to same folder as " VarScriptName ".ahk", VarScriptName
}

; Create the hotkey
Hotkey currentHotkey, Main

; Logging function
LogMessage(message) {
    try {
        ; Check file size
        if FileExist(LogFile) {
            fileSize := FileGetSize(LogFile)
            if (fileSize > MAX_LOG_SIZE) {
                ; Read the file
                content := FileRead(LogFile)
                ; Split into lines
                lines := StrSplit(content, "`n", "`r")
                ; Keep only the latter half of the lines
                newContent := ""
                startIndex := lines.Length // 2
                Loop lines.Length - startIndex {
                    newContent .= lines[startIndex + A_Index] "`n"
                }
                ; Write back truncated content
                FileDelete(LogFile)
                FileAppend(newContent, LogFile)
            }
        }

        ; Append new log message
        timestamp := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss")
        logEntry := timestamp " - " message "`n"
        FileAppend(logEntry, LogFile)
    } catch Error as err {
        MsgBox "Error writing to log file: " err.Message
    }
}

; Load settings from INI file on startup
LoadSettings()

ShowInstructions(*) {
    LogMessage("Showing instructions dialog")
    result := MsgBox("Welcome to " VarScriptName ", a Markdown to HTML converter!`n`n"
        . "HOW TO USE" "`n"
        . "1. Click the tray icon or press " currentHotkey " to open the converter`n"
        . "2. Choose to open a Markdown file or paste Markdown text`n"
        . "3. The script will convert to HTML and download any images`n"
        . "   (You can change the hotkey from the tray menu)`n`n"
        . "FEATURES" "`n"
        . "- Converts Markdown headings, lists, and images to HTML`n"
        . "- Automatically downloads referenced images`n"
        . "- Creates a separate folder for each conversion`n"
        . "- Customizable hotkey via tray menu`n`n",
        VarScriptName " Setup & Usage Guide", "OK")
}

ChangeHotkey(*) {
    ; Create GUI
    myGui := Gui("+Resize", "Change Hotkey")
    myGui.SetFont("s10")

    ; Add instructions
    myGui.Add("Text",, "Press your desired hotkey combination...")

    ; Add hotkey control
    hkControl := myGui.Add("Hotkey", "vChosenHotkey w200", currentHotkey)

    ; Add buttons
    myGui.Add("Button", "Default w80", "OK").OnEvent("Click", ProcessNewHotkey)
    myGui.Add("Button", "x+10 w80", "Cancel").OnEvent("Click", (*) => myGui.Destroy())
    myGui.Add("Button", "x+10 w120", "Reset to Default").OnEvent("Click", ResetToDefault)

    ResetToDefault(*) {
        defaultHotkey := "!m"
        hkControl.Value := defaultHotkey
    }

    ; Show GUI
    myGui.Show()

    ProcessNewHotkey(*) {
        ; Get the new hotkey
        newHotkey := hkControl.Value

        ; If no hotkey was chosen, show error and return
        if (newHotkey = "") {
            MsgBox("Please choose a valid hotkey.", "Error", "Icon!")
            return
        }

        ; Disable old hotkey
        try {
            Hotkey currentHotkey, "Off"
        }

        ; Enable new hotkey
        try {
            Hotkey newHotkey, Main
            currentHotkey := newHotkey
            ; Update the tooltip text
            VarBlurb := " Press " currentHotkey " or this icon to convert Markdown to HTML."
            A_IconTip := VarScriptName " " VarVersionNo " " VarBlurb
            LogMessage("Hotkey changed to: " currentHotkey)
            MsgBox("Hotkey successfully changed to " newHotkey, "Success")
            myGui.Destroy()
        } catch as err {
            LogMessage("Error changing hotkey: " err.Message)
            MsgBox("Failed to set new hotkey: " err.Message, "Error", "Icon!")
            ; Re-enable old hotkey
            Hotkey currentHotkey, Main
        }
    }
}

; Create the tray menu
A_TrayMenu.Delete()
A_TrayMenu.Add "Open " VarScriptName, (*) => Main()
A_TrayMenu.Add "Select Output Folder", (*) => SelectDestinationFolder()
A_TrayMenu.Add "Change Hotkey", (*) => ChangeHotkey()
A_TrayMenu.Add
A_TrayMenu.Add "About " VarScriptName, (*) => ShowInstructions()
A_TrayMenu.Add
A_TrayMenu.Add "Exit", (*) => ExitApp()
A_TrayMenu.Default := "Open " VarScriptName
A_TrayMenu.ClickCount := 1

; Show startup notification
ToolTip(VarScriptName " " VarVersionNo " loaded`nPress " currentHotkey " to open converter")
SetTimer () => ToolTip(), -3000  ; Hide tooltip after 3 seconds

LoadSettings() {
    global baseOutputDir, iniFile

    ; Try to read last used folder from INI file
    if (FileExist(iniFile)) {
        baseOutputDir := IniRead(iniFile, "Settings", "LastFolder", "")

        ; Also read hotkey if available
        savedHotkey := IniRead(iniFile, "Settings", "Hotkey", "")
        if (savedHotkey != "") {
            try {
                Hotkey currentHotkey, "Off"  ; Disable current hotkey
                Hotkey savedHotkey, Main     ; Enable saved hotkey
                currentHotkey := savedHotkey
                LogMessage("Loaded saved hotkey: " currentHotkey)
            } catch as err {
                LogMessage("Error loading saved hotkey: " err.Message)
                ; Keep using default hotkey if there's an error
            }
        }
    }

    LogMessage("Settings loaded. Output directory: " baseOutputDir)
}

SaveSettings() {
    global baseOutputDir, iniFile, currentHotkey

    ; Save current folder to INI file
    IniWrite(baseOutputDir, iniFile, "Settings", "LastFolder")

    ; Save current hotkey
    IniWrite(currentHotkey, iniFile, "Settings", "Hotkey")

    LogMessage("Settings saved. Output directory: " baseOutputDir)
}

SelectDestinationFolder() {
    global baseOutputDir  ; Explicitly mark as global

    LogMessage("Selecting destination folder")

    ; Prompt user to select a destination folder
    selectedDir := DirSelect("*" A_ScriptDir, 3, "Select destination folder for converted files")

    if (selectedDir != "") {
        baseOutputDir := selectedDir
        SaveSettings()  ; Save the selected folder
        LogMessage("New destination folder selected: " baseOutputDir)
    } else {
        ; If user cancels, use last saved folder or default to script directory
        if (baseOutputDir = "") {
            baseOutputDir := A_ScriptDir
            LogMessage("Using script directory as default output folder")
        }
    }
}

; Main setup - NOT run automatically
Main(*) {
    LogMessage("Starting MD2HTML conversion process")

    ; If baseOutputDir is not set (first run or empty INI)
    if (baseOutputDir = "") {
        SelectDestinationFolder()
    }

    ; Create the GUI
    CreateGui()
}

CreateGui() {
    myGui := Gui("", VarScriptName " " VarVersionNo " - Markdown to HTML Converter")
    myGui.SetFont("s10")
    myGui.Add("Text", "w400", "Choose how to input your Markdown content:")

    myGui.Add("Button", "w180 h30", "Open Markdown File").OnEvent("Click", OpenFile)
    myGui.Add("Button", "x+20 w180 h30", "Paste Markdown Text").OnEvent("Click", PasteText)

    ; Add button to change output folder
    myGui.Add("Button", "w380 h30 xm", "Change Output Folder").OnEvent("Click", (*) => SelectDestinationFolder())

    ; Show current output folder
    myGui.Add("Text", "w400 xm", "Current output folder: " baseOutputDir)

    myGui.Show("w400 h180")  ; Increased height to show all elements

    LogMessage("Main GUI displayed")
}

OpenFile(*) {
    LogMessage("Open file button clicked")

    ; Get the markdown file
    mdFile := FileSelect(1, A_ScriptDir, "Select Markdown File", "Markdown Files (*.md)")
    if (mdFile = "") {
        LogMessage("File selection cancelled")
        return  ; User cancelled
    }

    LogMessage("Selected file: " mdFile)

    ; Show processing message
    ShowProcessingMessage("Processing... Please wait")

    ; Read markdown content
    Try {
        mdContent := FileRead(mdFile)
        LogMessage("File read successfully")
    } Catch as e {
        CloseProcessingMessage()
        LogMessage("Error reading file: " e.Message)
        MsgBox("Error reading markdown file: " e.Message, "Error", "Icon!")
        return
    }

    ; Process the markdown
    ProcessMarkdown(mdContent)
}

PasteText(*) {
    LogMessage("Paste text button clicked")

    ; Create new GUI for text input
    inputGui := Gui("", VarScriptName " - Paste Markdown Text")
    inputGui.SetFont("s10")
    inputGui.Add("Text", "w600", "Paste your Markdown content below:")

    ; Add multiline edit control
    editControl := inputGui.Add("Edit", "w600 h400 vMarkdownText")
    editControl.Value := A_Clipboard  ; Pre-fill with clipboard content if available

    ; Add process button
    inputGui.Add("Button", "w100 h30 Default", "Process").OnEvent("Click", ProcessButton)

    ; Show the GUI
    inputGui.Show("w620 h500")

    LogMessage("Paste text GUI displayed")

    ProcessButton(*) {
        mdContent := editControl.Value
        if (mdContent = "") {
            LogMessage("No content to process")
            MsgBox("No content to process.", "Error", "Icon!")
            return
        }

        LogMessage("Processing pasted content")
        inputGui.Destroy()  ; Close the input GUI

        ; Show processing message
        ShowProcessingMessage("Processing... Please wait")

        ProcessMarkdown(mdContent)
    }
}

ProcessMarkdown(mdContent) {
    global baseOutputDir  ; Explicitly reference the global variable

    LogMessage("Processing markdown content")

    ; Extract the first headline to use as folder/file name
    firstHeadline := ExtractFirstHeadline(mdContent)
    if (firstHeadline = "") {
        firstHeadline := "converted_" FormatTime(A_Now, "yyyyMMdd_HHmmss")
        LogMessage("No headline found, using timestamp: " firstHeadline)
    } else {
        LogMessage("Using headline as folder name: " firstHeadline)
    }

    ; Clean the headline for use as a filename
    folderName := CleanFilename(firstHeadline)

    ; Create a unique folder for this conversion
    conversionDir := baseOutputDir "\" folderName
    imgDir := folderName  ; Use the same folder for images

    ; Create directory if it doesn't exist
    Try {
        if !DirExist(conversionDir) {
            DirCreate(conversionDir)
            LogMessage("Created directory: " conversionDir)
        } else {
            ; If folder exists, make it unique by adding a timestamp
            folderName := folderName "_" FormatTime(A_Now, "HHmmss")
            conversionDir := baseOutputDir "\" folderName
            imgDir := folderName
            DirCreate(conversionDir)
            LogMessage("Directory already exists, created with timestamp: " conversionDir)
        }
    } Catch as e {
        CloseProcessingMessage()
        LogMessage("Error creating directory: " e.Message)
        MsgBox("Error creating directory: " e.Message "`n`nWill use script directory instead.", "Error", "Icon!")

        ; Fall back to script directory
        baseOutputDir := A_ScriptDir
        conversionDir := baseOutputDir "\" folderName
        imgDir := folderName

        ; Try again with script directory
        Try {
            if !DirExist(conversionDir) {
                DirCreate(conversionDir)
                LogMessage("Created directory in script location: " conversionDir)
            }
        } Catch as e2 {
            LogMessage("Critical error creating directory: " e2.Message)
            MsgBox("Could not create directory in script location. Please run the script with administrator privileges.", "Critical Error", "Icon!")
            return
        }
    }

    ; Save the current output directory for future use
    SaveSettings()

    ; Process the markdown content
    LogMessage("Converting markdown to HTML")
    htmlContent := ConvertMarkdownToHtml(mdContent, imgDir)

    ; Save HTML file
    htmlFile := conversionDir "\00_" folderName ".html"
    Try {
        FileDelete(htmlFile)  ; Delete if exists
    } Catch {
        ; File doesn't exist, continue
    }

    Try {
        FileAppend(htmlContent, htmlFile, "UTF-8")
        LogMessage("HTML file saved: " htmlFile)
    } Catch as e {
        CloseProcessingMessage()
        LogMessage("Error saving HTML file: " e.Message)
        MsgBox("Error saving HTML file: " e.Message, "Error", "Icon!")
        return
    }

    ; Close the processing message BEFORE showing success message
    CloseProcessingMessage()

    ; Add a short delay to ensure processing window is closed
    Sleep(50)

    LogMessage("Conversion complete")
    MsgBox("Conversion complete!`n`nHTML and Image files have been saved to:`n" conversionDir, "Success", "Icon!")

    ; Open the HTML file
    Run(htmlFile)

    ; Open the directory
    Run("explorer.exe " conversionDir)
}

ExtractFirstHeadline(mdContent) {
    ; Look for the first headline (# Headline)
    contentArray := StrSplit(mdContent, "`n", "`r")
    for i, line in contentArray {
        if (RegExMatch(line, "^#+\s+(.+)$", &match)) {
            return match[1]
        }
    }
    return ""
}

CleanFilename(filename) {
    ; Remove invalid characters from filename
    cleaned := ""
    invalidChars := "\/:" . "*" . "?" . Chr(34) . "<" . ">" . "|"

    ; Go through each character and replace invalid ones
    Loop Parse filename {
        if InStr(invalidChars, A_LoopField)
            cleaned .= "_"
        else
            cleaned .= A_LoopField
    }

    ; Limit length
    if (StrLen(cleaned) > 50)
        cleaned := SubStr(cleaned, 1, 50)

    return cleaned
}

ConvertMarkdownToHtml(mdContent, imgDir) {
    ; Basic HTML template
    htmlStart := "<!DOCTYPE html>"
    htmlStart .= "`n<html>"
    htmlStart .= "`n<head>"
    htmlStart .= "`n    <meta charset='UTF-8'>"
    htmlStart .= "`n    <style>"
    htmlStart .= "`n        /* Minimal styling to avoid conflicts with Freshdesk */"
    htmlStart .= "`n        img { max-width: 800px; display: block; }"
    htmlStart .= "`n        .step { margin-bottom: 10px; }"
    htmlStart .= "`n    </style>"
    htmlStart .= "`n</head>"
    htmlStart .= "`n<body>"

    htmlEnd := "</body>"
    htmlEnd .= "`n</html>"
    ; Process the content line by line
    contentArray := StrSplit(mdContent, "`n", "`r")
    htmlBody := ""

    ; Keep track of list state
    inList := false
    imageCount := 0

    for i, line in contentArray {
        ; Skip lines with Scribe reference
        if (InStr(line, "Scribe"))
            continue

        ; Process heading (lines starting with #)
        if (RegExMatch(line, "^#+\s+(.+)$", &match)) {
            headingLevel := StrLen(RegExReplace(line, "^(#+).*$", "$1"))
            headingText := match[1]
            htmlBody .= "<h" headingLevel ">" headingText "</h" headingLevel ">`n"
            continue
        }

        ; Process numbered list items (like "1. ")
        if (RegExMatch(line, "^(\d+)\.(.+)$", &match)) {
            number := match[1]
            text := Trim(match[2])

            if (!inList) {
                htmlBody .= "<div class='step'>`n"
                inList := true
            } else {
                htmlBody .= "</div>`n<div class='step'>`n"
            }

            htmlBody .= "<span class='step-number'>" number ".</span> " text "`n"
            continue
        }

        ; Process images
        if (RegExMatch(line, "!\[\]\((.+?)\)", &match)) {
            imageUrl := match[1]
            imageCount++
            LogMessage("Processing image " imageCount ": " imageUrl)

            ; Format image count with leading zero for sorting
            paddedCount := Format("{:02}", imageCount)

            ; Download the image
            localImage := DownloadImage(imageUrl, paddedCount, imgDir)
            if (localImage) {
                htmlBody .= "<img src='" localImage "' alt='Step Image'>`n"
            } else {
                ; If download fails, use original URL
                htmlBody .= "<img src='" imageUrl "' alt='Step Image'>`n"
            }
            continue
        }

        ; Handle empty lines
        if (Trim(line) = "") {
            if (inList) {
                htmlBody .= "</div>`n"
                inList := false
            }
            htmlBody .= "<br>`n"
            continue
        }

        ; Default: treat as paragraph
        htmlBody .= "<p>" line "</p>`n"
    }

    ; Close any open list
    if (inList)
        htmlBody .= "</div>`n"

    return htmlStart htmlBody htmlEnd
}

DownloadImage(url, paddedCount, imgDir) {
    global baseOutputDir  ; Explicitly reference the global variable
    global folderName  ; Explicitly reference the global variable

    ; Clean URL
    url := Trim(url)
    LogMessage("Downloading image: " url)

    ; Extract filename from URL
    SplitPath(url, &origFileName)

    ; Create new filename - using the naming pattern requested
    newFileName := paddedCount "_" folderName ".jpg"
    fullImagePath := baseOutputDir "\" imgDir "\" newFileName
    LogMessage("Saving image to: " fullImagePath)

    ; For image error handling - keep track if we should continue
    imageDownloaded := false

    ; Download the image
    Try {
        whr := ComObject("WinHttp.WinHttpRequest.5.1")
        whr.Open("GET", url, true)
        whr.Send()
        whr.WaitForResponse()

        ; Check response status
        if (whr.Status = 200) {
            Try {
                ; Read binary data
                ADO := ComObject("ADODB.Stream")
                ADO.Type := 1  ; Binary
                ADO.Open()
                ADO.Write(whr.ResponseBody)
                ADO.SaveToFile(fullImagePath, 2)  ; 2 = overwrite if exists
                ADO.Close()

                imageDownloaded := true
                LogMessage("Image downloaded successfully")
            } Catch as e {
                ; Create a placeholder text file instead
                LogMessage("Error saving image: " e.Message)
                FileAppend("Image download failed: " url, fullImagePath ".txt", "UTF-8")
            }
        } else {
            LogMessage("HTTP error downloading image. Status: " whr.Status)
        }
    } Catch as e {
        ; Create a placeholder text file instead
        LogMessage("Error downloading image: " e.Message)
        FileAppend("Image download failed: " url, fullImagePath ".txt", "UTF-8")
    }

    ; If download succeeded, return image filename
    if (imageDownloaded) {
        return newFileName
    } else {
        ; If download failed, use original URL
        return url
    }
}

; Simple processing message functions
ShowProcessingMessage(text := "Processing... Please wait") {
    global processingGui  ; Ensure we use the global variable
    LogMessage("Showing processing message: " text)

    ; Destroy any existing processing GUI first
    CloseProcessingMessage()

    ; Create new processing GUI
    processingGui := Gui("+AlwaysOnTop -SysMenu -Caption", "Processing")
    processingGui.SetFont("s10")
    processingGui.Add("Text", "w300 Center", text)
    processingGui.Show("w320 h60")
}

CloseProcessingMessage() {
    global processingGui  ; Ensure we use the global variable

    if (IsSet(processingGui) && processingGui) {
        try {
            processingGui.Destroy()
            processingGui := ""  ; Clear the reference
            LogMessage("Closed processing message")
        } catch {
            ; Ignore errors if window is already gone
        }
    }
}
