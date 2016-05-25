' // ***************************************************************************
' // 
' // Copyright (c) Microsoft Corporation.  All rights reserved.
' // 
' // Microsoft Deployment Toolkit Solution Accelerator
' //
' // File:      WelcomeWiz_Choice.vbs
' // 
' // Version:   6.2.5019.0
' // 
' // Purpose:   Scripts for welcome wizard choice pane
' // 
' // ***************************************************************************


Option Explicit

Dim sSelectedBtn

Function StartMenu
	Dim iRetVal, sArch
	
	sArch = "x86"
	If oEnvironment.Item("Architecture") = "X64" then sArch = "x64"
	
	
	ButtonNext.disabled=True
	
		' Local Variables
		
		Dim sExecutable, sBGInfo, sWindowHide, sWinVnc
		Dim sCmdString, iCmdRetVal
		Dim colProcesses
		Dim objWMIService
		
	' If in WinPE
	If oEnvironment.Item("OSVersion") = "WinPE" Then
	
		Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
		
		' Figure out the executable to use
		
		If oEnvironment.Item("Architecture") = "X64" then
			sExecutable = "BGInfo64.exe"
			sWinVnc = "\Tools\Programs\UltraVNC64\winvnc.exe"
		Else
			sExecutable = "BGInfo.exe"
			sWinVnc = "\Tools\Programs\UltraVNC\winvnc.exe"
		End if
		
		' Activate winvnc
		
		Set colProcesses = objWMIService.ExecQuery _ 
		("SELECT * FROM Win32_Process WHERE Name = " & _
		"'winvnc.exe'")
		
		If oFso.FileExists(oEnvironment.Item("DeployRoot") & "\Tools\Programs\UltraVNC\winvnc.exe") and colProcesses.Count = 0 Then
			sCmdString = "wpeutil DisableFirewall"
			iCmdRetVal = oShell.Run(sCmdString)
			sCmdString = "cmd /c start " & oEnvironment.Item("DeployRoot") & sWinVnc
			iCmdRetVal = oShell.Run(sCmdString)
		End If
		
		' Find the BGInfo executable
		
		iRetVal = oUtility.FindFile(sExecutable, sBGInfo)
		If iRetVal <> Success then
			oLogging.CreateEntry "Unable to find " & sExecutable & ", exiting.", LogTypeInfo
			Exit Function
		End if
		
		' Run BGInfo
		
		sCmdString = """" & sBGInfo & """ " & oEnvironment.Item("DeployRoot") & "\Tools\X86\bgconfig.bgi /nolicprompt /silent /timer:0"
		oShell.CurrentDirectory = oEnvironment.Item("DeployRoot") & "\Tools\X86"
		
		
		On Error Resume Next
			iCmdRetVal = oShell.Run(sCmdString)
			TestAndLog iCmdRetVal, "Failed to set the background"
		On Error Goto 0
		' We never want to fail
		
	End If
	
		'' Disable Buttons for items not present
		If not oFso.FileExists("\\fmi-data\_DEP\Transfert\Sources\CMB\poste2015\poste-cmb.bat") or oEnvironment.Item("OSVersion") <> "WinPE" Then
			buttonItem4.Style.display = "none"
		End if

		'B1 - MDT
		'B2 - LOTC
		'B3 - HBCD
		'B4 - Deploy CMB
		'B7 - Restore HP EliteBook
			
		If not oFSO.FileExists(oEnvironment.Item("DeployRoot") & "\Tools\recovery\" & sArch & "\RecEnv.exe") or oEnvironment.Item("OSVersion") <> "WinPE" Then
			buttonItem5.Style.display = "none"
		End if
	
		If not oFso.FileExists(oEnvironment.Item("DeployRoot") & "Tools\recovery\tools\" & sArch & "\MSDartTools.exe") Then
			buttonItem6.Style.display = "none"
		End if

		If sArch <> "x86" Then
			buttonItem3.Style.display = "none"
		End If

		If UCase(Property("Make")) <> "HEWLETT-PACKARD" or Instr(1, UCase(Property("Model")), "HP ELITEBOOK 2530P", 1) <> 1 or oEnvironment.Item("OSVersion") <> "WinPE" Then 
			buttonItem7.Style.display = "none"
		End IF
	
	
End Function

Function GetValueFromID( oItem ) 

	Select Case oItem.ID
		Case buttonitem1.ID
			GetValueFromID = "DEPLOYWIZARD"
		Case buttonitem2.ID
			GetValueFromID = "RECOVERY"
		Case buttonitem3.ID
			GetValueFromID = "HIRENSBOOT"
		Case buttonitem4.ID
			GetValueFromID = "POSTECMB2015"
		Case buttonitem5.ID
			GetValueFromID = "RECENV"
		Case buttonitem6.ID
			GetValueFromID = "DART"			
		Case buttonitem7.ID
			GetValueFromID = "TEHP2530P"			
			
	End select

End Function 


Function RunSelCmd

	Dim sValue
	sValue = GetValueFromID(window.event.srcElement)

	Select Case (window.event.type)
	
		Case "mouseout", "deactivate"
		
			If window.event.srcElement.ID <> sSelectedBtn then
				window.event.srcElement.style.backgroundimage = "url(btnout.png)"
			End if
		
		Case "mouseover"
		
			If window.event.srcElement.ID <> sSelectedBtn then
				window.event.srcElement.style.backgroundimage = "url(btnover.png)"
			End if

		Case "activate"
		
			ActivateItem window.event.srcElement
		
		Case "click", "dblclick"
			ActivateItem window.event.srcElement
		'	ButtonNextClick
		RunSelectedCommand
			
	End Select

End function


Function ActivateItem ( oItemNew ) 

	if sSelectedBtn <> "" then
		document.GetElementByID(sSelectedBtn).style.backgroundimage = "url(btnout.png)"
	End if
	oItemNew.style.backgroundimage = "url(btnsel.png)"

	sSelectedBtn = oItemNew.ID
	oItemNew.Focus

End function


Sub KeyHandlerCustom

	if window.event.srcElement.tagName = "INPUT" then
		KeyHandler
		exit sub
	End if

	select case window.event.KeyCode

		Case 40  ' Down

			If window.event.srcElement.ID = "buttonItem1" and buttonItem2.style.display <> "none" then
				ActivateItem buttonItem2
			Elseif (window.event.srcElement.ID = "buttonItem1" or window.event.srcElement.ID = "buttonItem2") and buttonItem3.style.display <> "none" then
				ActivateItem buttonItem3
			Else
				ActivateItem buttonItem4
			End if

		Case 38  ' Up		

			If window.event.srcElement.ID = "buttonItem4" and buttonItem3.style.display <> "none" then
				ActivateItem buttonItem3
			ElseIf (window.event.srcElement.ID = "buttonItem4" or window.event.srcElement.ID = "buttonItem3") and buttonItem2.style.display <> "none" then
				ActivateItem buttonItem2
			Else
				ActivateItem buttonItem1
			End if

		End select
	
End sub



Function RunSelectedCommand 

	Dim sCmd, sArch
	
	sArch = "x86"
	If oEnvironment.Item("Architecture") = "X64" then sArch = "x64"

	Select case GetValueFromID(document.GetElementByID(sSelectedBtn))

		Case "RECOVERY"

		'	RunSelectedCommand = True
		'	ButtonNext.disabled = False
			oEnvironment.Item("SkipLOTC")="NO"
			'Alert(oEnvironment.Item("SkipLOTC"))
			ButtonNext.disabled=False
			ButtonNextClick
			'Exit function
		Case "HIRENSBOOT"
			
		'	sCmd = "x:\Tools\Programs\FreeCommanderXE\FreeCommander.exe"
		'	If oFso.FileExists(sCmd) Then
		'		oShell.Run sCmd & " x:\Tools\Programs", 1, true
		'	Else
		'		sCmd = "z:\Tools\Programs\FreeCommanderXE\FreeCommander.exe"
		'		If oFso.FileExists(sCmd) Then
		'			oShell.Run sCmd & " z:\Tools\Programs", 1, true
		'		Else
		'			Alert("Programme introuvable...")
		'		End If
		'	End If
			document.body.style.cursor = "Wait"
			sCmd = oEnvironment.Item("DeployRoot") & "\Tools\" & sArch & "\Explorer++.exe"
			sCmd = sCmd & " " & chr(34) & oEnvironment.Item("DeployDrive") & "\Tools\Programs" & chr(34)
			oShell.Run sCmd, 1, true
			document.body.style.cursor = "default"
			RunSelectedCommand = false
		
		Case "POSTECMB2015"
		
		If MsgBoxConfirm ("Formatage complet du poste") Then
			sCmd = "\\fmi-data\_DEP\Transfert\Sources\CMB\poste2015\poste-cmb.bat"
			If oFso.FileExists(sCmd) Then
				oShell.Run sCmd , 1, true
				window.close
				Exit function
			End If
		End If
		
		Case "RECENV"

			document.body.style.cursor = "Wait"
			sCmd = oEnvironment.Item("DeployRoot") & "\Tools\recovery\" & sArch & "\RecEnv.exe"
			oShell.Run sCmd, 1, true
			document.body.style.cursor = "default"
			RunSelectedCommand = false

		Case "DART"

			document.body.style.cursor = "Wait"
			sCmd = oEnvironment.Item("DeployRoot") & "Tools\recovery\tools\" & sArch & "\MSDartTools.exe"
			oShell.Run sCmd, 1, true
			document.body.style.cursor = "default"
			RunSelectedCommand = false
		
		Case "TEHP2530P"
			sCmd = oEnvironment.Item("DeployRoot") & "\Applications\FMI-TEHP2530p\poste-tech-2530p.bat"
			If MsgBoxConfirm ("Formatage complet du poste") Then
				If oFso.FileExists(sCmd) Then
					oShell.Run sCmd , 1, true
					window.close
					Exit function
				Else
					Alert("Script introuvable, contacter votre administrateur")
				End If
			End If
			
			'Alert sCmd
			'document.body.style.cursor = "Wait"
			'window.close
				
		Case else ' "DEPLOYWIZARD"
			
		'	RunSelectedCommand = True
			oEnvironment.Item("SkipLOTC")="YES"
			'Alert(oEnvironment.Item("SkipLOTC"))
			ButtonNext.disabled=False
			ButtonNextClick
			'Exit function

	End select

End function


Function SafeRegRead( KeyValue )
   on error resume next
      SafeRegRead = oShell.RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\WinPE\KeyboardLayouts\" & GetLocale & "\" & KeyValue)
   on error goto 0
End function


Function CustomInitializationCloseout
	buttonItem1.focus
End function 
