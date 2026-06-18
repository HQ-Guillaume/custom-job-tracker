Option Explicit

Dim shell, fso, scriptDirectory, command

Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
scriptDirectory = fso.GetParentFolderName(WScript.ScriptFullName)

command = "powershell.exe -NoLogo -NoProfile -WindowStyle Hidden -STA -ExecutionPolicy Bypass -File " & _
    """" & scriptDirectory & "\app\cli\Launch-AnalyticsJobCrawlerGui.ps1" & """"

shell.CurrentDirectory = scriptDirectory
shell.Run command, 0, False
