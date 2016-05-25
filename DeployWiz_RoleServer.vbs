'
'
'
'
'
'  VALIDATION IP IMM
Option Explicit

Function InitializeRoleServer

'	If oProperties("SRVNOMCLIENT") = "" then
'		SRVNOMCLIENT.Value = "CLIENT"
'	End if

	If oProperties("IMMIP") = "" then
		IMMIP.Value = "*.*.*.*"
	End if
	
	If oProperties("IMMMASK") = "" then
		IMMMASK.Value = "255.255.255.0"
	End if

	If oProperties("IMMGW") = "" then
		IMMGW.Value = "*.*.*.*"
	End if

End Function

Function ValidateChoice
	document.getElementById("IMMPWD").innerHTML = ""
	SRVNOMCLIENT.disabled = JDRadio1.checked
	IMMIP.disabled = JDRadio1.checked
	IMMMASK.disabled = JDRadio1.checked
	IMMGW.disabled = JDRadio1.checked
	ButtonNext.disabled = not JDRadio1.checked
End Function

'Function ValidateSRVNOMCLIENT
	' Check Warnings
'	ParseAllWarningLabels
'	If Len(SRVNOMCLIENT.Value) > 15 then
'		InvalidCharNC.style.display = "none"
'		TooLongNC.style.display = "inline"
'		ValidateSRVNOMCLIENT = false
'		ButtonNext.disabled = true
'	ElseIf IsValidNOM ( SRVNOMCLIENT.Value ) then
'		ValidateSRVNOMCLIENT = TRUE
'		InvalidCharNC.style.display = "none"
'		TooLongNC.style.display = "none"
'		
'		document.getElementById("IMMPWD").innerHTML = "Le mot de passe IMM sera " & "" & SRVNOMCLIENT.Value & "@fmi" & ""
'		
'	Else
'		InvalidCharNC.style.display = "inline"
'		TooLongNC.style.display = "none"
'		ValidateSRVNOMCLIENT = false
'		ButtonNext.disabled = true
'	End if
'End function

Function ValidateRoleServer

End Function

Function ValidateIMMIP
	' Check Warnings
	ParseAllWarningLabels
	If Len(IMMIP.Value) > 15 then
		InvalidCharIP.style.display = "none"
		TooLongIP.style.display = "inline"
		ValidateIMMIP = false
		ButtonNext.disabled = true
	ElseIf IsValidIP ( IMMIP.Value ) then
		ValidateIMMIP = TRUE
		InvalidCharIP.style.display = "none"
		TooLongIP.style.display = "none"
	Else
		InvalidCharIP.style.display = "inline"
		TooLongIP.style.display = "none"
		ValidateIMMIP = false
		ButtonNext.disabled = true
	End if
End function

Function ValidateIMMGW
	' Check Warnings
	ParseAllWarningLabels
	If Len(IMMGW.Value) > 15 then
		InvalidCharGW.style.display = "none"
		TooLongGW.style.display = "inline"
		ValidateIMMGW = false
		ButtonNext.disabled = true
	ElseIf IsValidIP ( IMMGW.Value ) then
		ValidateIMMGW = TRUE
		InvalidCharGW.style.display = "none"
		TooLongGW.style.display = "none"
	Else
		InvalidCharGW.style.display = "inline"
		TooLongGW.style.display = "none"
		ValidateIMMGW = false
		ButtonNext.disabled = true
	End if
End function

Function ValidateIMMMASK
	' Check Warnings
	ParseAllWarningLabels
	If Len(IMMMASK.value) > 15 then
		InvalidCharMK.style.display = "none"
		TooLongMK.style.display = "inline"
		ValidateIMMMASK = false
		ButtonNext.disabled = true
	ElseIf IsValidIP ( IMMMASK.Value ) then
		ValidateIMMMASK = TRUE
		InvalidCharMK.style.display = "none"
		TooLongMK.style.display = "none"
	Else
		InvalidCharMK.style.display = "inline"
		TooLongMK.style.display = "none"
		ValidateIMMMASK = false
		ButtonNext.disabled = true
	End if
End function

Function IsValidNOM ( NOM )
    Dim regEx
    'RegExMatch=False
    Set regEx = New RegExp     
    regEx.Pattern = "[^A-Z0-9\-\_]"
		IsValidNOM = not regEx.Test ( NOM )	and Len( NOM ) <= 15
End Function 

Function IsValidIP ( IP )
    Dim regEx
    Dim strBM
    Dim strIpMatch
    'RegExMatch=False
    Set regEx = New RegExp     
		regEx.Pattern = "((25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)(\.|$)){4}"
		IsValidIP = regEx.Test ( IP ) and	Len( IP ) <= 15
End Function 