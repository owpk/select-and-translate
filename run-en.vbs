Dim lang
Dim WshShell
lang = "en"
Set WshShell = CreateObject("WScript.Shell")
WshShell.Run chr(34) & ".\ntf.bat " & Chr(34) & lang, 0
Set WshShell = Nothing