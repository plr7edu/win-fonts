Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Get the directory where this script is located
strScriptPath = objFSO.GetParentFolderName(WScript.ScriptFullName)

' Path to requirements.txt (assumes it's in the same folder as the script)
strRequirementsFile = objFSO.BuildPath(strScriptPath, "requirements.txt")

' Check if requirements.txt exists
If Not objFSO.FileExists(strRequirementsFile) Then
    MsgBox "Error: requirements.txt not found in: " & strScriptPath, vbCritical, "File Not Found"
    WScript.Quit
End If

' Build the command to run in Windows Terminal
' The command installs from requirements.txt and then pauses
strCommand = "wt.exe cmd /k ""pip install -r """ & strRequirementsFile & """ && pause"""

' Execute the command
objShell.Run strCommand, 1, False

' Clean up
Set objShell = Nothing
Set objFSO = Nothing