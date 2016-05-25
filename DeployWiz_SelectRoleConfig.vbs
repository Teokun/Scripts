' // ***************************************************************************
' // 
' // Copyright (c) Microsoft Corporation.  All rights reserved.
' // 
' // Microsoft Deployment Toolkit Solution Accelerator
' //
' // File:      DeployWiz_Applications.vbs
' // 
' // Version:   6.2.5019.0
' // 
' // Purpose:   Deployment Wizard ROLE Configurations
' // 
' // ***************************************************************************

Option Explicit

Dim oRoleSelect, oSiteSelect, oRoleOption, oSiteOption

Function InitializeRoleConfig

 Dim oSelRole, oSite
 Dim sCmd
 Dim iRetVal, i, j
 Dim strHolder
 
  
  SelectRoleConfig_ProcessingLock( False )
 
 If oEnvironment.listitem("SelectableRole") is nothing then
  Exit function
 ElseIf oEnvironment.listitem("SelectableRole").count < 1 Then
  Exit function
 End if
	
		
 ProcessRoleProperties.style.display = "none"
 ProcessReloadRole.style.display = "none"
 ProcessSite.style.display = "none"
 
	Set oSiteSelect=document.getElementById("hClientSiteSelect")
	'For Each oSiteOption in oSiteSelect.Options 
	'	oSiteOption.RemoveNode
	'Next
	
	Set oRoleSelect=document.getElementById("hRoleSelect")
	For Each oRoleOption in oRoleSelect.Options 
		oRoleOption.RemoveNode
	Next
	i=0

	For each oSelRole in Property("SelectableRole")
			If oSelRole <> "" Then
				i = i + 1
				ReDim Preserve arrListRole(i+1)
		  	arrListRole(i) = oSelRole
		  End If
	Next
	
	For i = ( UBound( arrListRole ) - 1 ) to 0 Step -1
		For j= 0 to i
			If UCase( arrListRole( j ) ) > UCase( arrListRole( j + 1 ) ) Then
				strHolder 				 = arrListRole( j + 1 )
				arrListRole( j + 1 ) = arrListRole( j )
				arrListRole( j )     = strHolder
			End If
		Next
	Next
	
 SetRoleOption  "-- Applications ONLY --", ""
 
 For i = 0 To UBound(arrListRole)
  If arrListRole(i) <> "" Then SetRoleOption arrListRole(i),arrListRole(i)
 Next

 If Not IsEmpty(Property("Role001")) Then
		hRoleSelect.Value = Property("Role001")
 End If
 	
	
	If CLIENTINVSITE.Value <> "" Then hClientSiteSelect.Value=CLIENTINVSITE.Value
	
	'Affichage information Profil en cache
	If oEnvironment.Item("WizardSelectionProfile") <> "" Then 
		document.getElementById("cCurrentProfile").innerHTML = "Profil en cache : " & oEnvironment.Item("WizardSelectionProfile")
	Else
		document.getElementById("cCurrentProfile").innerHTML = "Aucun profil en cache"
	End If
	
	
	If UCase(oEnvironment.Item("IsServer")) = "TRUE" Then 
		TypeMaterielSRV.checked = True
		Alert("Attention, cocher E-Backup selon le projet")
	Else
		TypeMaterielPC.checked = True
	End If

	' Affichage Organisation Pour Control Visuel
	'If Not IsEmpty(Property("TagSite")) Then
	'	document.getElementById("cCurrentProfile").innerHTML = document.getElementById("cCurrentProfile").innerHTML & "<br />" & _
	'	" TAG : " & "TAG_" & oEnvironment.Item("OrgName") & oEnvironment.Item("TagSite")
	'Else
	If Not IsEmpty(Property("OrgName")) Then
			document.getElementById("cCurrentProfile").innerHTML = document.getElementById("cCurrentProfile").innerHTML & "<br />" & _
			" TAG : " & "TAG_" & oEnvironment.Item("OrgName")
	Else
			document.getElementById("cCurrentProfile").innerHTML = document.getElementById("cCurrentProfile").innerHTML & "<br />"
	End If	

End Function

'''
Sub ClearSite
	Dim sVariable, i, sPadded
	
	Set oSiteSelect=document.getElementById("hClientSiteSelect")
	For Each oSiteOption in oSiteSelect.Options 
		oSiteOption.RemoveNode
	Next	
	
 	If ( oEnvironment.listitem("ClientSite") is Nothing ) or ( oEnvironment.listitem("ClientSite").count < 1 ) then
 		 SetSiteOption "-- Aucun --", ""
  	Exit Sub
 	End if
	

	
	sVariable = "CLIENTSITE"
		'Set ListI = CreateObject("Scripting.Dictionary")
		On Error Resume Next
		For i = 1 to 150
			sPadded = sVariable & Right("000" & CStr(i), 3)
			'Alert(sPadded & ":" & oEnvironment.Item(sPadded) )
			If oEnvironment.Item(sPadded)<> "" then
					oEnvironment.Item(sPadded) = ""
			ElseIf oEnvironment.Item(sVariable & CStr(i)) then
					oEnvironment.Item(sVariable & CStr(i)) = ""
			End if
		Next
		On Error Goto 0
		Err.Clear
	
End Sub


Function ProcessRoleConfig
	Dim sCmd, oItem, oSite, oClientSelect, iRetVal, ListI
	
		'If Property("Role001") = hRoleSelect.Value Then Exit Function
		oEnvironment.Item("Role001") = hRoleSelect.Value
		
		ButtonNext.disabled = true
		ProcessSite.style.display = "inline"
	
	sCmd = "wscript.exe """ & oUtility.ScriptDir & "\ZTIGather.wsf""" & " /inifile:""" & oUtility.ScriptDir &  "\..\Control\ClientSiteSettings.ini"""
	iRetVal = oShell.Run(sCmd, , true)
	
 If not( ( oEnvironment.listitem("ClientSite") is nothing ) or ( oEnvironment.listitem("ClientSite").count < 1 ) ) Then
  	
  	For each oSite in Property("ClientSite")
	  	SetSiteOption oSite,oSite
		Next
		
	End If
	
		ProcessSite.style.display = "none"
		ButtonNext.disabled = False

End Function

'''
Function SelectRoleConfig_ProcessingLock( sValue )
		hClientSiteSelect.disabled = sValue
		hRoleSelect.disabled = sValue
		buttonReloadRole.disabled = sValue
		buttonReloadSite.disabled = sValue

	  TypeMaterielPC.disabled = sValue
	  TypeMaterielSRV.disabled = sValue
	  TypeMaterielEBKP.disabled = sValue
End Function

Function ValidateRoleSelectList
 Dim sCmd, oItem
	
	SelectRoleConfig_ProcessingLock( True )
	
	oEnvironment.Item("TagSite") = ""
	oEnvironment.Item("OrgName") = ""
  oEnvironment.Item("AdmAcc") = ""
  oEnvironment.Item("AdmPwd") = "" 
  oEnvironment.Item("Role001") = hRoleSelect.Value
  
  CLIENTINVSITE.Value =  hClientSiteSelect.Value
  oEnvironment.Item("ClientSite001") = hClientSiteSelect.Value
	
  If IsEmpty(Property("Role001")) Then
 		ValidateRoleSelectList=True
	 	Exit function
  End If	
 
  If TypeMaterielSRV.checked or TypeMaterielEBKP.checked Then
 		oEnvironment.Item("SkipRoles")="NO"
 		oEnvironment.Item("WizardSelectionProfile")="--- SERVEUR ---"
 	End If
 
 CreateInventaireDir
 'TAG pour OCS Inventory client déployé
 CLIENTINV.Value=chr(34) & Property("Role001") & chr(34)
 
 'Alert(TagTmp.Value)
  
	' Process full rules (needed to pick up the role settings, apps, etc.)
   ProcessRoleProperties.style.display = "inline"
	 sCmd = "wscript.exe """ & oUtility.ScriptDir & "\ZTIGather.wsf""" & " /inifile:""" & oUtility.ScriptDir & "\..\Control\DepSettings.ini"""
	 oItem = oShell.Run(sCmd, , true)
	 ProcessRoleProperties.style.display = "none"

		If Len(oEnvironment.Item("TagSite")) > 0 Then 
			oEnvironment.Item("TagSite") = "_" & oEnvironment.Item("TagSite")
		Else
			oEnvironment.Item("TagSite") = ""
		End If
		Alert("Tag FusionInventory : " & "TAG_" & oEnvironment.Item("OrgName") & oEnvironment.Item("TagSite") )
		
		oEnvironment.Item("OrgName") = oEnvironment.Item("OrgName") & oEnvironment.Item("TagSite")
	
	If (not IsEmpty(Property("AdmAcc"))) and (not IsEmpty(Property("AdmPwd"))) Then
		'alert(Property("AdmAcc"))
		'alert(Property("AdmPwd"))
		AdminACC.Value = Property("AdmAcc")
		'AdminPWD.Value = Property("AdmPwd")		
		If TypeMaterielSRV.checked or TypeMaterielEBKP.checked Then
			' AdminACC.Value = ""
			AdminPWD.Value = 	UCASE(Property("AdmPwd"))
			AdminPWD.Value =  Replace(AdminPWD.Value, "@FMI","@fmi")
		Else
			AdminPWD.Value = Property("AdmPwd") & ".FR" 	'Modification effective au 21/04/2015
		End If
		
	End If
	

	
	'Ajout pour gestion des paramètres lors de la selection tache serveur
	'If TypeMaterielSRV.checked or TypeMaterielEBKP.checked Then
	'	AdminPWD.Value = 	UCASE(AdminPWD.Value)
	'	AdminPWD.Value =  Replace(UCASE(Property("AdmPwd")), "@FMI","@fmi")	
	'End If
	
	SelectRoleConfig_ProcessingLock( False )
	
	ValidateRoleSelectList = True

End Function

''''
'
'


Sub SetRoleOption(OptText,OptValue) 
	Dim oNewOption
	Set oNewOption = Document.CreateElement("OPTION")
 	oNewOption.Text = OptText 
	oNewOption.Value = OptValue 
	oRoleSelect.options.Add(oNewOption) 
End Sub

Sub SetSiteOption(OptText,OptValue) 
	Dim oNewOption
	Set oNewOption = Document.CreateElement("OPTION")
 	oNewOption.Text = OptText 
	oNewOption.Value = OptValue 
	oSiteSelect.options.Add(oNewOption) 
End Sub

Function CreateInventaireDir
	Dim sServeurInventaire
	Dim strDossierInventaire
	
	sServeurInventaire = "\\fmi-data\DATA\Inventaires\Data"
	
	strDossierInventaire = sServeurInventaire & "\" & Replace(Property("Role001")," ","_") & "-" & Replace(Replace(hClientSiteSelect.Value," ","_"),"/","_")
	
	If Not oFso.FolderExists(sServeurInventaire) Then
		Exit Function
	Else
		If Not oFso.FolderExists(strDossierInventaire) Then CreateChemin strDossierInventaire
	End If

End Function

Sub CreateChemin(ByVal Path)
  If Not oFSO.FolderExists(Path) Then
    CreateChemin oFSO.GetParentFolderName(Path)
    oFSO.CreateFolder Path
  End If
End Sub

Function ReloadRole
Dim sCmd, oItem, iRetVal

		ButtonNext.disabled = True		
	  ProcessReloadRole.style.display = "inline"
	 sCmd = "wscript.exe """ & oUtility.ScriptDir & "\ZTIGather.wsf""" & " /inifile:""" & oUtility.ScriptDir & "\..\Control\CustomRoleSettings.ini"""
	 oItem = oShell.Run(sCmd, , true)
	 	 
	 ProcessReloadRole.style.display = "none"
	 InitializeRoleConfig
	 ButtonNext.disabled = False
	  
End Function

Sub QuickCleanup2
		ClearSite
		oEnvironment.Item("WizardSelectionProfile")=""
		oEnvironment.Item("OSDComputerName")=""
		oEnvironment.Item("OrgName") = ""
		oEnvironment.Item("TagSite") = ""
		oEnvironment.Item("AdminPassword") = ""
		hRoleSelect.Value=""
		set g_AllOperatingSystems = Nothing
		window.location.reload()
End sub