Set fso = CreateObject("Scripting.FileSystemObject")
Set sh  = CreateObject("WScript.Shell")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
sh.CurrentDirectory = scriptDir

venvPython = scriptDir & "\.venv\Scripts\pythonw.exe"
If fso.FileExists(venvPython) Then
    pyExe = """" & venvPython & """"
Else
    pyExe = "pythonw"
End If

sh.Run pyExe & " app.py", 0, False
