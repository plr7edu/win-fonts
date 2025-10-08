; =====================================================================
; Silent Font Installer Script - AutoHotkey v2 (FONTREGISTER LIST METHOD)
; Installs fonts and captures exact registered names from FontRegister
; Uses FontRegister's list command to get actual registered names
; Requires Administrator Rights
; Creates font_metadata.json for uninstaller
; =====================================================================

#Requires AutoHotkey v2.0+
#SingleInstance Force

; Check if running as administrator
if not A_IsAdmin {
    try {
        Run('*RunAs "' . A_ScriptFullPath . '"')
    }
    ExitApp
}

scriptDir := A_ScriptDir
logFile := scriptDir . "\font_installation.log"
metadataFile := scriptDir . "\font_metadata.json"

; Initialize log
WriteLog("=" . StrRepeat("=", 68))
WriteLog("FONT INSTALLATION LOG - " . FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss"))
WriteLog("=" . StrRepeat("=", 68))
WriteLog("")

; Create temporary directory
tempDir := A_Temp . "\FontRegisterTemp_" . A_TickCount
DirCreate(tempDir)
WriteLog("Created temporary directory: " . tempDir)

; Extract embedded FontRegister files
try {
    FileInstall("FontRegister.exe", tempDir . "\FontRegister.exe", 1)
    FileInstall("FontRegister.exe.config", tempDir . "\FontRegister.exe.config", 1)
    FileInstall("FontRegister.pdb", tempDir . "\FontRegister.pdb", 1)
    FileInstall("FontRegister.xml", tempDir . "\FontRegister.xml", 1)
    WriteLog("Extracted embedded FontRegister files successfully")
} catch as err {
    WriteLog("ERROR: Failed to extract embedded files - " . err.Message)
    CleanupAndExit(tempDir, 1)
}

fontRegisterExe := tempDir . "\FontRegister.exe"

if !FileExist(fontRegisterExe) {
    WriteLog("ERROR: FontRegister.exe not found after extraction")
    CleanupAndExit(tempDir, 1)
}

; Collect font files
fontExtensions := ["*.ttf", "*.otf", "*.ttc", "*.fon", "*.fnt"]
fontFiles := []

WriteLog("Scanning for font files in: " . scriptDir)
WriteLog("")

for index, pattern in fontExtensions {
    Loop Files, scriptDir . "\" . pattern {
        fontFiles.Push(A_LoopFileFullPath)
        WriteLog("Found: " . A_LoopFileName)
    }
}

WriteLog("")
WriteLog("Total fonts found: " . fontFiles.Length)

if fontFiles.Length = 0 {
    WriteLog("WARNING: No font files found")
    CleanupAndExit(tempDir, 0)
}

WriteLog("")
WriteLog("=" . StrRepeat("=", 68))
WriteLog("GETTING PRE-INSTALLATION FONT LIST")
WriteLog("=" . StrRepeat("=", 68))
WriteLog("")

; Get list of fonts BEFORE installation
preInstallList := tempDir . "\fonts_before.txt"
WriteLog("Capturing font list before installation...")

try {
    RunWait('"' . fontRegisterExe . '" list --machine > "' . preInstallList . '"', tempDir, "Hide")
    WriteLog("✓ Pre-installation font list captured")
} catch as err {
    WriteLog("WARNING: Could not capture pre-installation list - " . err.Message)
}

WriteLog("")
WriteLog("=" . StrRepeat("=", 68))
WriteLog("INSTALLING FONTS")
WriteLog("=" . StrRepeat("=", 68))
WriteLog("")

; Build command line for FontRegister
fontPathsArg := ""
for index, fontPath in fontFiles {
    fontPathsArg .= '"' . fontPath . '" '
}

command := '"' . fontRegisterExe . '" install --machine ' . fontPathsArg

; Execute installation
WriteLog("Executing: FontRegister.exe install --machine [" . fontFiles.Length . " fonts]")
startTime := A_TickCount

try {
    RunWait(command, tempDir, "Hide")
    elapsedTime := (A_TickCount - startTime) / 1000
    WriteLog("Installation completed in " . Round(elapsedTime, 2) . " seconds")
} catch as err {
    WriteLog("ERROR: Failed to execute FontRegister.exe - " . err.Message)
    CleanupAndExit(tempDir, 1)
}

; Wait for registry to update
Sleep(1000)

WriteLog("")
WriteLog("=" . StrRepeat("=", 68))
WriteLog("GETTING POST-INSTALLATION FONT LIST")
WriteLog("=" . StrRepeat("=", 68))
WriteLog("")

; Get list of fonts AFTER installation
postInstallList := tempDir . "\fonts_after.txt"
WriteLog("Capturing font list after installation...")

try {
    RunWait('"' . fontRegisterExe . '" list --machine > "' . postInstallList . '"', tempDir, "Hide")
    WriteLog("✓ Post-installation font list captured")
} catch as err {
    WriteLog("ERROR: Could not capture post-installation list - " . err.Message)
    CleanupAndExit(tempDir, 1)
}

WriteLog("")
WriteLog("=" . StrRepeat("=", 68))
WriteLog("ANALYZING NEWLY INSTALLED FONTS")
WriteLog("=" . StrRepeat("=", 68))
WriteLog("")

; Read both lists
beforeFonts := Map()
afterFonts := Map()

if FileExist(preInstallList) {
    try {
        content := FileRead(preInstallList, "UTF-8")
        Loop Parse, content, "`n", "`r" {
            line := Trim(A_LoopField)
            if (line != "" && !InStr(line, "Installed fonts")) {
                beforeFonts[line] := true
            }
        }
        WriteLog("Fonts before installation: " . beforeFonts.Count)
    }
}

if FileExist(postInstallList) {
    try {
        content := FileRead(postInstallList, "UTF-8")
        Loop Parse, content, "`n", "`r" {
            line := Trim(A_LoopField)
            if (line != "" && !InStr(line, "Installed fonts")) {
                afterFonts[line] := true
            }
        }
        WriteLog("Fonts after installation: " . afterFonts.Count)
    }
}

; Find newly installed fonts
newFonts := []
for fontName in afterFonts {
    if !beforeFonts.Has(fontName) {
        newFonts.Push(fontName)
        WriteLog("NEW: " . fontName)
    }
}

WriteLog("")
WriteLog("Total new fonts registered: " . newFonts.Length)

WriteLog("")
WriteLog("=" . StrRepeat("=", 68))
WriteLog("CREATING METADATA")
WriteLog("=" . StrRepeat("=", 68))
WriteLog("")

; Create metadata
metadata := Map()
metadata["installation_date"] := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss")
metadata["installation_scope"] := "machine"
metadata["fonts"] := []

; Define installation directory
systemFontsDir := "C:\Windows\Fonts"

; Map font files to their registered names
for index, fontPath in fontFiles {
    SplitPath(fontPath, &fileName)

    fontEntry := Map()
    fontEntry["filename"] := fileName
    fontEntry["source_path"] := fontPath
    fontEntry["installed_path"] := systemFontsDir . "\" . fileName
    fontEntry["registered_names"] := []

    ; Try to match this file with registered font names
    ; For now, store all new fonts (we'll improve matching if needed)
    if (index <= newFonts.Length) {
        fontEntry["registered_names"].Push(newFonts[index])
    }

    metadata["fonts"].Push(fontEntry)

    WriteLog("Metadata entry:")
    WriteLog("  File: " . fileName)
    WriteLog("  Path: " . fontEntry["installed_path"])
    WriteLog("  Registered as: " . (fontEntry["registered_names"].Length > 0 ? fontEntry["registered_names"][1] : "Unknown"))
}

WriteLog("")
WriteLog("=" . StrRepeat("=", 68))
WriteLog("ALTERNATIVE METHOD: CHECKING WINDOWS REGISTRY")
WriteLog("=" . StrRepeat("=", 68))
WriteLog("")

; Read from Windows Registry for more accurate mapping
WriteLog("Reading font registry entries...")
try {
    Loop Reg, "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts" {
        fontRegName := A_LoopRegName
        fontFileName := RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts", fontRegName)

        ; Check if this registry entry matches any of our font files
        for index, fontEntry in metadata["fonts"] {
            if (InStr(fontFileName, fontEntry["filename"])) {
                ; Extract font name (remove file extension info from registry name)
                fontName := RegExReplace(fontRegName, "\s*\(.*?\)\s*$", "")

                ; Check if already in list
                alreadyExists := false
                for , existingName in fontEntry["registered_names"] {
                    if (existingName = fontName) {
                        alreadyExists := true
                        break
                    }
                }

                if !alreadyExists {
                    fontEntry["registered_names"].Push(fontName)
                    WriteLog("✓ Found: " . fontEntry["filename"] . " → " . fontName)
                }
            }
        }
    }
    WriteLog("")
    WriteLog("Registry scan completed")
} catch as err {
    WriteLog("WARNING: Could not read registry - " . err.Message)
}

WriteLog("")
WriteLog("=" . StrRepeat("=", 68))
WriteLog("SAVING METADATA")
WriteLog("=" . StrRepeat("=", 68))
WriteLog("")

WriteLog("Saving metadata to: " . metadataFile)

try {
    jsonContent := MapToJSON(metadata, 0)
    if FileExist(metadataFile)
        FileDelete(metadataFile)
    FileAppend(jsonContent, metadataFile, "UTF-8")
    WriteLog("✓ Metadata file created successfully")
    WriteLog("")
    WriteLog("Metadata summary:")
    for index, fontEntry in metadata["fonts"] {
        WriteLog("  " . fontEntry["filename"])
        WriteLog("    Location: " . fontEntry["installed_path"])
        WriteLog("    Registered names (" . fontEntry["registered_names"].Length . "):")
        for , regName in fontEntry["registered_names"] {
            WriteLog("      - " . regName)
        }
    }
} catch as err {
    WriteLog("ERROR: Failed to save metadata file - " . err.Message)
    WriteLog("Uninstaller may not work properly without this file!")
}

WriteLog("")
WriteLog("=" . StrRepeat("=", 68))
WriteLog("INSTALLATION SUMMARY")
WriteLog("=" . StrRepeat("=", 68))
WriteLog("")
WriteLog("✓ SUCCESS: Installed " . fontFiles.Length . " font file(s) system-wide")
WriteLog("✓ Installation directory: " . systemFontsDir)
WriteLog("✓ Total registered font names: " . newFonts.Length)
WriteLog("")
WriteLog("All fonts are now available for all users and applications")
WriteLog("")
WriteLog("To uninstall these fonts, run B_Font_Uninstaller.ahk")
WriteLog("(The uninstaller will use: font_metadata.json)")
WriteLog("")
WriteLog("=" . StrRepeat("=", 68))

CleanupAndExit(tempDir, 0)

; =====================================================================
; HELPER FUNCTIONS
; =====================================================================

WriteLog(message) {
    global logFile
    try {
        FileAppend(message . "`n", logFile, "UTF-8")
    }
}

CleanupAndExit(tempDir, exitCode) {
    global logFile

    if DirExist(tempDir) {
        try {
            DirDelete(tempDir, 1)
            WriteLog("Cleaned up temporary directory")
        } catch {
            WriteLog("WARNING: Could not delete temporary directory: " . tempDir)
        }
    }

    if exitCode = 0 {
        WriteLog("Script finished successfully")
    } else {
        WriteLog("Script finished with errors (Exit Code: " . exitCode . ")")
    }

    ExitApp(exitCode)
}

StrRepeat(str, count) {
    result := ""
    Loop count {
        result .= str
    }
    return result
}

MapToJSON(obj, indent := 0) {
    spaces := StrRepeat("  ", indent)
    nextSpaces := StrRepeat("  ", indent + 1)

    if Type(obj) = "Map" {
        json := "{`n"
        first := true
        for key, value in obj {
            if !first
                json .= ",`n"
            json .= nextSpaces . '"' . JSONEscape(key) . '": ' . MapToJSON(value, indent + 1)
            first := false
        }
        json .= "`n" . spaces . "}"
        return json
    }
    else if Type(obj) = "Array" {
        json := "[`n"
        first := true
        for index, value in obj {
            if !first
                json .= ",`n"
            json .= nextSpaces . MapToJSON(value, indent + 1)
            first := false
        }
        json .= "`n" . spaces . "]"
        return json
    }
    else if Type(obj) = "String" {
        return '"' . JSONEscape(obj) . '"'
    }
    else if Type(obj) = "Integer" or Type(obj) = "Float" {
        return String(obj)
    }
    else {
        return '""'
    }
}

JSONEscape(str) {
    str := StrReplace(str, "\", "\\")
    str := StrReplace(str, '"', '\"')
    str := StrReplace(str, "`n", "\n")
    str := StrReplace(str, "`r", "\r")
    str := StrReplace(str, "`t", "\t")
    return str
}