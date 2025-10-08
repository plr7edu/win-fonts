'=== Cleanup.vbs ===
Option Explicit

Dim shell, fso, scriptDir, buildDir, distDir, specFile

Set shell = CreateObject("Shell.Application")
Set fso = CreateObject("Scripting.FileSystemObject")

' Get current directory where the script is located
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)

' Define paths
buildDir = scriptDir & "\build"
distDir = scriptDir & "\dist"
specFile = scriptDir & "\Font_Name_Editor.spec"

' Move to Recycle Bin function
' FOF_ALLOWUNDO = &H40 (allows undo from recycle bin)
' FOF_NOCONFIRMATION = &H10 (no confirmation dialog)
Const FOF_ALLOWUNDO = &H40
Const FOF_NOCONFIRMATION = &H10
Const FOF_SILENT = &H4

' Delete build folder
If fso.FolderExists(buildDir) Then
    shell.NameSpace(0).ParseName(buildDir).InvokeVerb "delete"
End If

' Delete dist folder
If fso.FolderExists(distDir) Then
    shell.NameSpace(0).ParseName(distDir).InvokeVerb "delete"
End If

' Delete spec file
If fso.FileExists(specFile) Then
    shell.NameSpace(0).ParseName(specFile).InvokeVerb "delete"
End If

MsgBox "Cleanup complete! Build files moved to Recycle Bin.", vbInformation, "Cleanup"

' Cleanup
Set fso = Nothing
Set shell = Nothing