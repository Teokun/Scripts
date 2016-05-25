'
'
'
'
'
'
'	Script Inventaire déploiement v1.0

Option Explicit


If WScript.Arguments.Count = 0 Then
	Dim objshell 
	Set objshell = CreateObject("Shell.Application")
	objshell.ShellExecute "wscript.exe", Chr(34) & WScript.ScriptFullName & Chr(34) & " Run", , "Runas", 1
Else
	Inventaire
End If 

Function Inventaire
	Dim shl
	Dim rootFolder
	rootFolder = "\\fmi-data\DATA\_DEP\Postes"
	Set shl = CreateObject("WScript.Shell")
	shl.run "net use \\fmi-data /user:dep\dep dep", 1 , True
	shl.run Chr(34) & rootFolder & "\Inventaire.bat" & chr(34)
End Function