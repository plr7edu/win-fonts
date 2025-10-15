Option Explicit

Dim fso, folder, file
Dim currentPath, fontExtensions, ext
Dim baseName, fontFamily, fontWeight, fontStyle, newName
Dim renamedCount, skippedCount, errorCount
Dim logMessage

' Initialize
Set fso = CreateObject("Scripting.FileSystemObject")
currentPath = fso.GetParentFolderName(WScript.ScriptFullName)
Set folder = fso.GetFolder(currentPath)

' Font file extensions to process
fontExtensions = Array(".ttf", ".otf", ".ttc", ".woff", ".woff2")

renamedCount = 0
skippedCount = 0
errorCount = 0
logMessage = ""

' Process each file in the current directory
For Each file In folder.Files
    ext = LCase(fso.GetExtensionName(file.Name))
    
    ' Check if it's a font file
    If IsInArray("." & ext, fontExtensions) Then
        On Error Resume Next
        
        ' Get base filename without extension
        baseName = fso.GetBaseName(file.Name)
        
        ' Extract weight and style from filename (only if they exist)
        fontWeight = ExtractWeightIfExists(baseName)
        fontStyle = ExtractStyle(baseName)
        
        ' Remove weight and style from base name to get clean font family
        fontFamily = RemoveWeightAndStyle(baseName)
        
        ' Remove all spaces from font family
        fontFamily = Replace(fontFamily, " ", "")
        
        ' Clean up any trailing hyphens or underscores
        Do While Right(fontFamily, 1) = "-" Or Right(fontFamily, 1) = "_"
            fontFamily = Left(fontFamily, Len(fontFamily) - 1)
        Loop
        
        ' Build new filename with pattern: FontFamily-Weight-Style.ext
        newName = fontFamily
        If fontWeight <> "" Then
            newName = newName & "-" & fontWeight
        End If
        If fontStyle <> "" Then
            newName = newName & "-" & fontStyle
        End If
        newName = newName & "." & ext
        
        ' Rename if different from current name
        If LCase(file.Name) <> LCase(newName) Then
            Dim newPath
            newPath = fso.BuildPath(folder.Path, newName)
            
            ' Check if target file already exists
            If Not fso.FileExists(newPath) Then
                file.Move newPath
                If Err.Number = 0 Then
                    renamedCount = renamedCount + 1
                    logMessage = logMessage & file.Name & " -> " & newName & vbCrLf
                Else
                    errorCount = errorCount + 1
                    logMessage = logMessage & "ERROR: " & file.Name & " (" & Err.Description & ")" & vbCrLf
                    Err.Clear
                End If
            Else
                skippedCount = skippedCount + 1
                logMessage = logMessage & "SKIPPED: " & file.Name & " (target already exists)" & vbCrLf
            End If
        Else
            skippedCount = skippedCount + 1
        End If
        
        On Error GoTo 0
    End If
Next

' Show completion message
Dim message
message = "Font File Renaming Complete!" & vbCrLf & vbCrLf
message = message & "Renamed: " & renamedCount & " file(s)" & vbCrLf
message = message & "Skipped: " & skippedCount & " file(s)" & vbCrLf
message = message & "Errors: " & errorCount & " file(s)" & vbCrLf & vbCrLf

If renamedCount > 0 Or errorCount > 0 Or skippedCount > 0 Then
    message = message & "Details:" & vbCrLf & logMessage
End If

MsgBox message, vbInformation + vbOKOnly, "Font Renamer - Complete"

' Cleanup
Set file = Nothing
Set folder = Nothing
Set fso = Nothing

' Helper Functions
Function IsInArray(value, arr)
    Dim i
    IsInArray = False
    For i = LBound(arr) To UBound(arr)
        If LCase(arr(i)) = LCase(value) Then
            IsInArray = True
            Exit Function
        End If
    Next
End Function

Function ExtractWeightIfExists(filename)
    Dim weights, weight
    ' Order matters - check longer names first to avoid partial matches
    weights = Array("ExtraLight", "ExtraBold", "SemiBold", "DemiBold", "Thin", "Light", "Regular", "Medium", "Bold", "Black", "Heavy", "Book")
    
    ExtractWeightIfExists = ""
    For Each weight In weights
        If InStr(1, filename, weight, vbTextCompare) > 0 Then
            ExtractWeightIfExists = weight
            Exit Function
        End If
    Next
    
    ' Return empty string if no weight found (don't default to Regular)
    ExtractWeightIfExists = ""
End Function

Function ExtractWeight(filename)
    Dim weights, weight
    ' Order matters - check longer names first to avoid partial matches
    weights = Array("ExtraLight", "ExtraBold", "SemiBold", "DemiBold", "Thin", "Light", "Regular", "Medium", "Bold", "Black", "Heavy", "Book")
    
    ExtractWeight = ""
    For Each weight In weights
        If InStr(1, filename, weight, vbTextCompare) > 0 Then
            ExtractWeight = weight
            Exit Function
        End If
    Next
    
    ' Default to Regular if not found
    ExtractWeight = "Regular"
End Function

Function ExtractStyle(filename)
    ExtractStyle = ""
    If InStr(1, filename, "Italic", vbTextCompare) > 0 Then
        ExtractStyle = "Italic"
    ElseIf InStr(1, filename, "Oblique", vbTextCompare) > 0 Then
        ExtractStyle = "Oblique"
    End If
End Function

Function RemoveWeightAndStyle(filename)
    Dim result, weights, weight, styles, style
    result = filename
    
    ' List of weights to remove
    weights = Array("ExtraLight", "ExtraBold", "SemiBold", "DemiBold", "Thin", "Light", "Regular", "Medium", "Bold", "Black", "Heavy", "Book")
    
    ' List of styles to remove
    styles = Array("Italic", "Oblique")
    
    ' Remove weights (case-insensitive)
    For Each weight In weights
        result = ReplaceWord(result, weight, "")
    Next
    
    ' Remove styles (case-insensitive)
    For Each style In styles
        result = ReplaceWord(result, style, "")
    Next
    
    ' Clean up any trailing/leading spaces, hyphens or underscores
    result = Trim(result)
    Do While Right(result, 1) = "-" Or Right(result, 1) = "_" Or Right(result, 1) = " "
        result = Left(result, Len(result) - 1)
    Loop
    Do While Left(result, 1) = "-" Or Left(result, 1) = "_" Or Left(result, 1) = " "
        result = Mid(result, 2)
    Loop
    
    RemoveWeightAndStyle = Trim(result)
End Function

Function ReplaceWord(text, word, replacement)
    Dim result, pos
    result = text
    
    ' Case-insensitive replacement - find the word
    pos = InStr(1, result, word, vbTextCompare)
    
    If pos > 0 Then
        result = Left(result, pos - 1) & replacement & Mid(result, pos + Len(word))
    End If
    
    ReplaceWord = result
End Function