'#--------------------------------------------------------------------------------- 
'This script is to change Windows or office product key and try to active it.
' Use ONLY with administrator privileges

Option Explicit

	ChangeWinKey 

'This function is to change Windows key and active it
Function ChangeWinKey
	Dim objshell
	Set objshell = CreateObject("Wscript.shell")
	
	If Wscript.Arguments.Count < 1 Then
		objshell.popup "Longueur de la clé incorrect" & vbCrlf & vbCrlf & "Echec de l'activation", 5
		Set objshell = Nothing 
		Exit Function
	End If

	Dim VBpath, SysRoot, Result, Flag, value, str
	Dim objArgs
	Set objArgs = Wscript.Arguments
	
	If objArgs(0).Length <> 29 Then 
		objshell.popup "Longueur de la clé incorrect" & vbCrlf & vbCrlf & "Echec de l'activation", 5
		Set objshell = Nothing 
		set objArgs = Nothing
		Exit Function
	End If
	
	'Get system root path 
	SysRoot = objshell.ExpandEnvironmentStrings("%SystemRoot%")
	'Get the vbscript path 
	VBpath =  SysRoot & "\System32\slmgr.vbs"
	'Get Windows key 
	value = objArgs(0)

    	'Import Windows product key 
		Set Result = objshell.Exec("Cscript.exe " & VBPath & " -ipk " & value)
		Flag = False 
		Do While Not result.StdOut.AtEndOfStream 
			str = result.StdOut.ReadLine()
			If InStr(UCase(str),UCase("Error")) <> 0 Then 
				Flag = True 
			End If 
			WScript.Echo str
		Loop 
		'If importing failed, try again.
		
		If Flag = True  Then 
			objshell.popup "Longueur de la clé incorrect" & vbCrlf & vbCrlf & "Echec de l'activation", 5
			Set objshell = Nothing 
			set objArgs = Nothing
			Exit Function	
		Else
		'Try to active Windows 
			WScript.Echo "Activation de la clé de produit Wndows."
			Set Result = objshell.Exec("Cscript.exe " & VBPath & " -ato")
			Flag = False 
			Do While Not result.StdOut.AtEndOfStream 
				str = result.StdOut.ReadLine()
				WScript.Echo str
			Loop 
		End If 

	Set objshell = Nothing 
	Set result = Nothing 
End Function 

