' // ***************************************************************************
' // 
' // Copyright (c) Microsoft Corporation.  All rights reserved.
' // 
' // Microsoft Deployment Toolkit Solution Accelerator
' //
' // File:      DeployWiz_GhostDetect.vbs
' // 
' // Version:   6.2.5019.0
' // 
' // Purpose:   Ghost file detection - ask for Ghost getback
' // 
' // ***************************************************************************

Option Explicit

Function InitializeGhostDetect

	Dim strSN 'SerialNumber
	Dim strGF 'GhostFolder
	Dim strGFF	'GhostFolderFound
	Dim subf1, subf2

	strSN = UCase(oEnvironment.Item("SerialNumber"))
 	
 	'strGF = "\\fmi-data\DATA\_DEP\ISOs - IMAGES\GHOST"
 		
 	strGF = "\\dx\sove\GHOST"
 	
 	GDRadio2.checked = True
 	
 	If oFso.FolderExists(strGF) Then
 		SrvChemin.Value = strGF
 		For Each subf1 in oFso.GetFolder( strGF ).SubFolders
			For Each subf2 in subf1.SubFolders
				If LCase(cstr(subf2.Name)) = LCase(strSN) Then
					strGFF = cstr(subf2.Path)
					'wscript.echo "Une sauvegarde se trouve dans le chemin suivant : " & vbcrlf & vbcrlf & strGFF
					GhostChemin.Value = Right(strGFF,Len(strGFF)-Len(strGF)-1)
					GDRadio1.checked = True
					GhostSize.Value = Round(oFso.GetFolder(strGFF).Size / 1024 / 1024 / 1024 ,2) & " Go"
				End If
			Next
		Next
	Else
		SrvChemin.Value = "Non disponible"
	End If

End Function

Function ValidateGhostDetect_Final
	If GDRadio1.checked Then
		SrvChemin.Value = SrvChemin.Value & "\" & GhostChemin.Value
		GhostChemin.Value = "C:\" & GhostChemin.Value & "\"
		
		Applications900.Value="{e93a875e-93c1-4e0d-9abc-8e8a773b7114}"
		ValidateGhostDetect_Final = True
	Else 
		If GDRadio2.checked Then 
			ValidateGhostDetect_Final = True
		Else
			ValidateGhostDetect_Final = False
		End If
	End If
End Function