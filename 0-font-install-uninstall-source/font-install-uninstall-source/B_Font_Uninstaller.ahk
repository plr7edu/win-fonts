; =====================================================================
; Silent Font Uninstaller Script - AutoHotkey v2 (REGISTRY-BASED)
; Uninstalls fonts using actual registered names from metadata
; Uses embedded FontRegister.exe for system-wide uninstallation
; Requires Administrator Rights
; Creates font_uninstallation.log in current directory
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
logFile := scriptDir . "\font_uninstallation.log"
metadataFile := scriptDir . "\font_metadata.json"

; Initialize log file
WriteLog("=" . StrRepeat("=", 68))
WriteLog("FONT UNINSTALLATION LOG - " . FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss"))
WriteLog("=" . StrRepeat("=", 68))
WriteLog("")

; Check if metadata file exists
if !FileExist(metadataFile) {
    WriteLog("ERROR: Metadata file not found: " . metadataFile)
    WriteLog("")
    WriteLog("The font_metadata.json file is required for uninstallation.")
    WriteLog("This file should have been created by A_Font_Installer.ahk")
    WriteLog("")
    WriteLog("Script aborted.")
    MsgBox("ERROR: font_metadata.json not found!`n`nThis file is required for uninstallation.`nPlease run the installer first or uninstall fonts manually.", "Font Uninstaller", "Icon!")
    ExitApp(1)
}

; Read metadata file
WriteLog("Reading metadata file: " . metadataFile)
try {
    metadataContent := FileRead(metadataFile, "UTF-8")
    metadata := JSONToMap(metadataContent)
    WriteLog("✓ Metadata loaded successfully")
    WriteLog("  Installation date: " . metadata["installation_date"])
    WriteLog("  Installation scope: " . metadata["installation_scope"])
    WriteLog("  Fonts to uninstall: " . metadata["fonts"].Length)
} catch as err {
    WriteLog("ERROR: Failed to read or parse metadata file - " . err.Message)
    WriteLog("Script aborted.")
    MsgBox("ERROR: Could not read metadata file!`n`n" . err.Message, "Font Uninstaller", "IconX")
    ExitApp(1)
}

WriteLog("")

; Create temporary directory for embedded files
tempDir := A_Temp . "\FontRegisterTemp_" . A_TickCount
DirCreate(tempDir)
WriteLog("Created temporary directory: " . tempDir)

; Extract embedded FontRegister files to temp directory
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

WriteLog("")
WriteLog("=" . StrRepeat("=", 68))
WriteLog("PRE-UNINSTALL VERIFICATION")
WriteLog("=" . StrRepeat("=", 68))
WriteLog("")

WriteLog("Verifying font files and registered names:")
WriteLog("")
fontsFound := 0

for index, fontEntry in metadata["fonts"] {
    fileName := fontEntry["filename"]
    installedPath := fontEntry["installed_path"]
    registeredNames := fontEntry["registered_names"]

    WriteLog("Font: " . fileName)

    if FileExist(installedPath) {
        fontsFound++
        WriteLog("  ✓ File exists: " . installedPath)
    } else {
        WriteLog("  ✗ File not found: " . installedPath)
    }

    WriteLog("  Registered names (" . registeredNames.Length . "):")
    for , regName in registeredNames {
        WriteLog("    - " . regName)
    }
    WriteLog("")
}

WriteLog("Font files found: " . fontsFound . " / " . metadata["fonts"].Length)

WriteLog("")
WriteLog("=" . StrRepeat("=", 68))
WriteLog("STARTING UNINSTALLATION PROCESS")
WriteLog("=" . StrRepeat("=", 68))
WriteLog("")

; Create PowerShell script for uninstallation
psScript := tempDir . "\uninstall.ps1"
psOutput := tempDir . "\uninstall_output.txt"

psContent := "Set-Location '" . tempDir . "'`n"
psContent .= "$ErrorActionPreference = 'Continue'`n"
psContent .= "`n"
psContent .= "# Create output file`n"
psContent .= "$outputFile = '" . psOutput . "'`n"
psContent .= "New-Item -ItemType File -Path $outputFile -Force | Out-Null`n"
psContent .= "`n"

scope := metadata["installation_scope"]
scopeFlag := (scope = "machine") ? "--machine" : "--user"

; Uninstall using registered names from metadata
for index, fontEntry in metadata["fonts"] {
    fileName := fontEntry["filename"]
    registeredNames := fontEntry["registered_names"]
    installedPath := fontEntry["installed_path"]

    psContent .= "Write-Output '=== UNINSTALLING: " . fileName . " ==='`n"
    psContent .= "Add-Content -Path '" . psOutput . "' -Value '=== " . fileName . " ==='`n"

    if (registeredNames.Length = 0) {
        psContent .= "Write-Output '  WARNING: No registered names found for this font'`n"
        psContent .= "Add-Content -Path '" . psOutput . "' -Value '  WARNING: No registered names in metadata'`n"
    } else {
        ; Uninstall each registered name
        for , regName in registeredNames {
            psContent .= "Write-Output '  Uninstalling: " . regName . "'`n"
            psContent .= "Add-Content -Path '" . psOutput . "' -Value '  Uninstalling: " . regName . "'`n"
            psContent .= "$result = & '" . fontRegisterExe . "' uninstall " . scopeFlag . " '" . regName . "' 2>&1 | Out-String`n"
            psContent .= "Add-Content -Path '" . psOutput . "' -Value $result`n"
        }
    }

    psContent .= "Write-Output ''`n"
}

; Clear font cache
psContent .= "`nWrite-Output '=== CLEARING FONT CACHE ==='`n"
psContent .= "Add-Content -Path '" . psOutput . "' -Value '=== CLEARING FONT CACHE ==='`n"
psContent .= "$result = & '" . fontRegisterExe . "' --clear-cache 2>&1 | Out-String`n"
psContent .= "Add-Content -Path '" . psOutput . "' -Value $result`n"
psContent .= "Add-Content -Path '" . psOutput . "' -Value '=== UNINSTALLATION COMPLETE ==='`n"

; Write PowerShell script
try {
    if FileExist(psScript)
        FileDelete(psScript)
    FileAppend(psContent, psScript, "UTF-8-RAW")
    WriteLog("Created PowerShell script for uninstallation")
} catch as err {
    WriteLog("ERROR: Failed to create PowerShell script - " . err.Message)
    CleanupAndExit(tempDir, 1)
}

; Execute PowerShell script
WriteLog("Executing FontRegister to uninstall fonts...")
WriteLog("Using installation scope: " . scope)
WriteLog("")
startTime := A_TickCount

try {
    RunWait('powershell.exe -ExecutionPolicy Bypass -NoProfile -File "' . psScript . '"', tempDir, "Hide")
    elapsedTime := (A_TickCount - startTime) / 1000
    WriteLog("Uninstallation process completed in " . Round(elapsedTime, 2) . " seconds")
} catch as err {
    WriteLog("ERROR: Failed to execute PowerShell - " . err.Message)
    CleanupAndExit(tempDir, 1)
}

; Wait for file system
Sleep(500)

; Read and log the output
WriteLog("")
WriteLog("=" . StrRepeat("=", 68))
WriteLog("FONTREGISTER OUTPUT:")
WriteLog("=" . StrRepeat("=", 68))
WriteLog("")

if FileExist(psOutput) {
    try {
        output := FileRead(psOutput, "UTF-8")
        if (Trim(output) = "") {
            WriteLog("(No output captured - FontRegister may be running silently)")
        } else {
            WriteLog(output)
        }
    } catch as err {
        WriteLog("ERROR: Could not read output file - " . err.Message)
    }
} else {
    WriteLog("WARNING: Output file was not created")
}

WriteLog("")
WriteLog("=" . StrRepeat("=", 68))
WriteLog("POST-UNINSTALL VERIFICATION")
WriteLog("=" . StrRepeat("=", 68))
WriteLog("")

; Check if fonts still exist after uninstallation
WriteLog("Verifying font removal:")
WriteLog("")
stillExist := []
successfullyRemoved := 0

for index, fontEntry in metadata["fonts"] {
    fileName := fontEntry["filename"]
    installedPath := fontEntry["installed_path"]
    registeredNames := fontEntry["registered_names"]

    if FileExist(installedPath) {
        stillExist.Push(fontEntry)
        WriteLog("⚠ STILL EXISTS: " . fileName)
        WriteLog("  Location: " . installedPath)
    } else {
        successfullyRemoved++
        WriteLog("✓ REMOVED: " . fileName)
        WriteLog("  Was at: " . installedPath)
    }
}

WriteLog("")
WriteLog("=" . StrRepeat("=", 68))
WriteLog("UNINSTALLATION SUMMARY")
WriteLog("=" . StrRepeat("=", 68))
WriteLog("")
WriteLog("Total fonts processed: " . metadata["fonts"].Length)
WriteLog("Successfully removed: " . successfullyRemoved)
WriteLog("Still present: " . stillExist.Length)
WriteLog("")

if stillExist.Length > 0 {
    WriteLog("⚠ WARNING: " . stillExist.Length . " font file(s) still present in system")
    WriteLog("")
    WriteLog("Files that still exist:")
    for index, fontEntry in stillExist {
        WriteLog("  - " . fontEntry["filename"])
        WriteLog("    Path: " . fontEntry["installed_path"])
        WriteLog("    Had " . fontEntry["registered_names"].Length . " registered name(s)")
    }
    WriteLog("")
    WriteLog("Possible reasons:")
    WriteLog("1. Fonts are locked by running applications")
    WriteLog("2. Font names in metadata don't match actual registered names")
    WriteLog("3. Insufficient permissions")
    WriteLog("")
    WriteLog("Solutions to try:")
    WriteLog("1. Close all applications and restart Windows")
    WriteLog("2. Manually uninstall from: Settings > Personalization > Fonts")
    WriteLog("3. Check installation log for correct registered font names")
    WriteLog("")
} else {
    WriteLog("✓ SUCCESS: All font files have been successfully removed!")
    WriteLog("")
    WriteLog("The metadata file (font_metadata.json) can now be safely deleted.")
}

WriteLog("=" . StrRepeat("=", 68))

; If successful, offer to delete metadata file
if stillExist.Length = 0 {
    try {
        FileDelete(metadataFile)
        WriteLog("Deleted metadata file: font_metadata.json")
    } catch {
        WriteLog("Note: Could not delete metadata file (you can remove it manually)")
    }
}

; Cleanup and exit
CleanupAndExit(tempDir, (stillExist.Length > 0) ? 1 : 0)

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

    WriteLog("")
    if exitCode = 0 {
        WriteLog("Script finished successfully - All fonts removed")
    } else {
        WriteLog("Script finished with warnings - Some fonts may still be present")
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

JSONToMap(jsonStr) {
    jsonStr := Trim(jsonStr)

    if (SubStr(jsonStr, 1, 1) = "{")
        jsonStr := SubStr(jsonStr, 2, StrLen(jsonStr) - 2)

    result := Map()
    pos := 1

    while (pos <= StrLen(jsonStr)) {
        keyStart := InStr(jsonStr, '"', , pos)
        if !keyStart
            break
        keyEnd := InStr(jsonStr, '"', , keyStart + 1)
        key := SubStr(jsonStr, keyStart + 1, keyEnd - keyStart - 1)

        colonPos := InStr(jsonStr, ":", , keyEnd)
        valueStart := colonPos + 1
        while (SubStr(jsonStr, valueStart, 1) = " " or SubStr(jsonStr, valueStart, 1) = "`n" or SubStr(jsonStr, valueStart, 1) = "`r" or SubStr(jsonStr, valueStart, 1) = "`t")
            valueStart++

        valueChar := SubStr(jsonStr, valueStart, 1)

        if (valueChar = '"') {
            valueEnd := InStr(jsonStr, '"', , valueStart + 1)
            value := SubStr(jsonStr, valueStart + 1, valueEnd - valueStart - 1)
            value := StrReplace(value, '\n', "`n")
            value := StrReplace(value, '\r', "`r")
            value := StrReplace(value, '\t', "`t")
            value := StrReplace(value, '\"', '"')
            value := StrReplace(value, '\\', '\')
            result[key] := value
            pos := valueEnd + 1
        }
        else if (valueChar = "[") {
            bracketCount := 1
            valueEnd := valueStart + 1
            while (bracketCount > 0 and valueEnd <= StrLen(jsonStr)) {
                char := SubStr(jsonStr, valueEnd, 1)
                if (char = "[")
                    bracketCount++
                else if (char = "]")
                    bracketCount--
                valueEnd++
            }
            arrayStr := SubStr(jsonStr, valueStart, valueEnd - valueStart)
            result[key] := JSONToArray(arrayStr)
            pos := valueEnd
        }
        else {
            valueEnd := InStr(jsonStr, ",", , valueStart)
            if !valueEnd
                valueEnd := StrLen(jsonStr) + 1
            value := Trim(SubStr(jsonStr, valueStart, valueEnd - valueStart))
            result[key] := value
            pos := valueEnd + 1
        }
    }

    return result
}

JSONToArray(arrayStr) {
    arrayStr := Trim(arrayStr)

    if (SubStr(arrayStr, 1, 1) = "[")
        arrayStr := SubStr(arrayStr, 2, StrLen(arrayStr) - 2)

    result := []

    if (Trim(arrayStr) = "")
        return result

    pos := 1
    while (pos <= StrLen(arrayStr)) {
        while (SubStr(arrayStr, pos, 1) = " " or SubStr(arrayStr, pos, 1) = "`n" or SubStr(arrayStr, pos, 1) = "`r" or SubStr(arrayStr, pos, 1) = "`t")
            pos++

        if (pos > StrLen(arrayStr))
            break

        char := SubStr(arrayStr, pos, 1)

        if (char = "{") {
            braceCount := 1
            elementEnd := pos + 1
            while (braceCount > 0 and elementEnd <= StrLen(arrayStr)) {
                c := SubStr(arrayStr, elementEnd, 1)
                if (c = "{")
                    braceCount++
                else if (c = "}")
                    braceCount--
                elementEnd++
            }
            objectStr := SubStr(arrayStr, pos, elementEnd - pos)
            result.Push(JSONToMap(objectStr))
            pos := elementEnd
        }
        else if (char = '"') {
            stringEnd := InStr(arrayStr, '"', , pos + 1)
            value := SubStr(arrayStr, pos + 1, stringEnd - pos - 1)
            value := StrReplace(value, '\n', "`n")
            value := StrReplace(value, '\r', "`r")
            value := StrReplace(value, '\t', "`t")
            value := StrReplace(value, '\"', '"')
            value := StrReplace(value, '\\', '\')
            result.Push(value)
            pos := stringEnd + 1
        }
        else {
            pos++
        }

        commaPos := InStr(arrayStr, ",", , pos)
        if commaPos
            pos := commaPos + 1
        else
            break
    }

    return result
}