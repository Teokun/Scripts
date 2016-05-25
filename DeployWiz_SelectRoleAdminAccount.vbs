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
' // Purpose:   Deployment Wizard ROLE Administrateur Account 
' // 
' // ***************************************************************************

Function InitializeRoleList

 Dim oSelRole
 Dim sFilteredXML
 Dim sCmd
 Dim iRetVal
 
 If oEnvironment.listitem("SelectableRole") is nothing then
  Exit function
 ElseIf oEnvironment.listitem("SelectableRole").count < 1 Then
  Exit function
 End if

 For each oSelRole in Property("SelectableRole")
  sFilteredXML = sFilteredXML & "<Selectablerole>"
  sFilteredXML = sFilteredXML & "<role>"
  sFilteredXML = sFilteredXML & oSelRole
  sFilteredXML = sFilteredXML & "</role>"
  sFilteredXML = sFilteredXML & "<comments>"
  sFilteredXML = sFilteredXML & "</comments>"
  sFilteredXML = sFilteredXML & "</Selectablerole>"
 Next

 If not IsEmpty(sFilteredXML) then
  broles.XMLDocument.LoadXML "<SelectableRoles>" & sFilteredXML & "</SelectableRoles>"
 End if
 
 ProcessRoleProperties.style.display = "none"
End Function

Function ReadyInitializeRoleList
 Dim oInput, oSelectableRoleList
 Dim bFound, oRoleItem
 
 ButtonNext.Disabled = False

 If not bRoleList.readystate = "complete" then
  Exit function  
 End if
  
 Set oSelectableRoleList = document.getElementsByName("Role")
  
 If oSelectableRoleList is nothing then
  Exit function
 ElseIf oSelectableRoleList.Length < 1 then
  Exit function
 End if
  
 For each oInput in oSelectableRoleList
  If UCase(document.all.item(oInput.SourceIndex - 1).TagName) = "INPUT" then
   If oInput.Value = "" then
    document.all.item(oInput.SourceIndex - 1).Disabled = TRUE
    document.all.item(oInput.SourceIndex - 1).Style.Display = "none"
   Else
    document.all.item(oInput.SourceIndex - 1).Style.Display = "inline"
    If not IsEmpty(Property("Role"))then
     For each oRoleItem in Property("Role")

      If UCase(oRoleItem) = UCase(oInput.Value) then     
       document.all.item(oInput.SourceIndex - 1).checked = TRUE
       'document.all.item(oInput.SourceIndex - 1).disabled = TRUE
       Exit for
      End if
     Next
    End if
   
   End if
  End if

 Next
End function

Sub RoleItemChange
 document.all.item(window.event.srcElement.SourceIndex + 1).Disabled = not window.event.SrcElement.checked
End Sub

Sub UpdateTitle
	'document.getElementById("RoleTitle").innerHTML = "En cours : "
End Sub

Function ValidateRoleList
 Dim oSelectedRoleList, oRole, p
 Dim iRetVal, sCmd
	
  oEnvironment.Item("AdmAcc") = ""
  oEnvironment.Item("AdmPwd") = "" 
  
 Set oSelectedRoleList = document.getElementsByName("SelectedItem")
 Set oEnvironment.ListItem("Role") = oSelectedRoleList
 
 	' Flush the value to variables.dat, before we continue.
	SaveAllDataElements
	SaveProperties
	' Exit if no role selected or role not change
 'If IsEmpty(Property("RoleTemp001")) or Property("RoleTemp001")=Property("Role001") Then
 '	ValidateRoleList=True
 ' Exit function
 'Else
 'Set oEnvironment.ListItem("Role") = oSelectedRoleList
 'End If		
	
 If IsEmpty(Property("Role001")) Then
 	ValidateRoleList=True
  Exit function
 End If	
 
 'TAG pour OCS Inventory client déployé
 CLIENTINV.Value=chr(34) & Property("Role001") & chr(34)
 
	' Process full rules (needed to pick up the role settings, apps, etc.)
   ProcessRoleProperties.style.display = "inline"
	 sCmd = "wscript.exe """ & oUtility.ScriptDir & "\ZTIGather.wsf"""
	 oItem = oShell.Run(sCmd, , true)
	 ProcessRoleProperties.style.display = "none"
	
	If (not IsEmpty(Property("AdmAcc"))) and (not IsEmpty(Property("AdmPwd"))) Then
	'alert(Property("AdmAcc"))
	'alert(Property("AdmPwd"))
		AdminACC.Value = Property("AdmAcc")
		AdminPWD.Value = Property("AdmPwd")
	End If
	
	ValidateRoleList = True
	
	ButtonNextClick
	

End Function
