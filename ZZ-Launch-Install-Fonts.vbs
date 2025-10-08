' Admin Launcher for PowerShell Script
' Launches ZZ-Install-Fonts.ps1 with administrator privileges in Windows Terminal

Set objShell = CreateObject("Shell.Application")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Get the current directory where this VBS script is located
strCurrentPath = objFSO.GetParentFolderName(WScript.ScriptFullName)

' PowerShell script name
strScriptName = "ZZ-Install-Fonts.ps1"
strScriptPath = objFSO.BuildPath(strCurrentPath, strScriptName)

' Check if the PowerShell script exists
If Not objFSO.FileExists(strScriptPath) Then
    MsgBox "Error: " & strScriptName & " not found in current directory!" & vbCrLf & vbCrLf & "Path: " & strCurrentPath, vbCritical, "Font Installer"
    WScript.Quit 1
End If

' Launch with Windows Terminal as administrator
objShell.ShellExecute "wt.exe", "powershell.exe -NoProfile -ExecutionPolicy Bypass -NoExit -File """ & strScriptPath & """", strCurrentPath, "runas", 1

' Clean up
Set objShell = Nothing
Set objFSO = Nothing