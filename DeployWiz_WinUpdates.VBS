' // ***************************************************************************
' // 
' // Copyright (c) Microsoft Corporation.  All rights reserved.
' // 
' // Microsoft Deployment Toolkit Solution Accelerator
' //
' // File:      DeployWiz_WinUpdates.vbs
' // 
' // Version:   6.2.5019.0
' // 
' // Purpose:   Windows Updates wizard pane validation
' // 
' // ***************************************************************************

Option Explicit
'''''''''''''''''''''''''''''''''''''
' Initialize Windows Update pane
'
Function InitializeWindowsUpdate
	If Len(WSUSServer.Value) > 1 Then WURadio3.checked = True
End Function

'''''''''''''''''''''''''''''''''''''
' Validate Windows Update pane
'
Function ValidateWindowsUpdate
	If WURadio3.checked Then
		WSUSServer.disabled = False
		
		If Not CheckURL(WSUSServer.Value) Then
   		ValidateWindowsUpdate = False
			InvalidWSUSServer.style.display = "inline"
		Else
			InvalidWSUSServer.style.display = "none"
			
			ValidateWindowsUpdate = True
		End If
	Else
		WSUSServer.disabled = True
		ValidateWindowsUpdate = True
		InvalidWSUSServer.style.display = "none"
	End If
End Function

'''''''''''''''''''''''''''''''''''''
' Validate_Final Windows Update pane
'
Function ValidateWindowsUpdate_Final
	If not WURadio3.checked Then
		 WSUSServer.Value = ""
	End if
	ValidateWindowsUpdate_Final = True
End Function


'''''''''''''''''''''''''''''''''''''
' CheckURL function
'
Function CheckURL(URL)
	'Dim oHttp
	'Set oHttp = CreateObject("MSXML2.XMLHTTP")
	' Set timeouts to 5s for name resolution, connect, send, and receive
	'oHttp.setTimeouts 5000, 5000, 5000, 5000
	'oHttp.open "GET", URL, false
	'On Error Resume Next
	'	oHttp.send
	'On Error Goto 0
	
	'If oHttp.status = 200 Then
    CheckURL = True
	'Else
	'	CheckURL = False
	'End If
	'Set oHttp = Nothing
End Function