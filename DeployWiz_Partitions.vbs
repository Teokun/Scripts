'
'
'
' VARIABLES
'
Option Explicit

Dim oDisk
Dim SizeDisk
Dim colDisks
Dim oWMI

Set oDisk = New ZTIDisk

Set oWMI = GetObject("winmgmts:\\.\root\cimv2")
Set colDisks = oWMI.ExecQuery ("Select * from Win32_DiskDrive")

Sub UpdateDisk
	SizeDisk = Round(oDisk.oWMI.Size / 1024 / 1024 /1024,0)
	document.getElementById("Disk0Title").innerHTML = "Volumetrie du disque " & oDisk.Disk & " : " & oDisk.oWMI.Caption & " : " & SizeDisk & " Go"
End Sub

Function ValidateDisk
	If Property("OSDDiskIndex") = "" Then
		If HD0.checked Then
			oDisk.Disk = 0
			oEnvironment.Item("OSDDiskIndex") = 0
		ElseIf HD1.checked Then
			oDisk.Disk = 1
			oEnvironment.Item("OSDDiskIndex") = 1
		End If
	Else
		If Property("OSDDiskIndex") = 0 And HD1.checked Then 
			oDisk.Disk = 1
			oEnvironment.Item("OSDDiskIndex") = 1
		ElseIf Property("OSDDiskIndex") = 1 And HD0.checked Then 
			oDisk.Disk = 0
			oEnvironment.Item("OSDDiskIndex") = 0
		End If
	End If
	
	UpdateDisk
	ValidateSizeDisk
End Function


Function InitializePartitions
	'Alert("NB:" & colDisks.Count)
	If colDisks.Count > 1 Then
		HD1.disabled = False
	Else
		HD1.disabled = True
	End If
	
	If oProperties("doNotFormatAndPartition") = "YES" then
		KeepPartitions.checked = True
	Else
		KeepPartitions.checked = False
	End if
	
	If UCase(Property("SkipUserData")) = "YES" then
		OptionMig.checked = False
	Else
		OptionMig.checked = True
	End if
		
	' Si pas de partition définis, initialisations des paramètres par défaut
	'alert(Property("OSDPartitions"))
	If Property("OSDPartitions")="" Then
		' Initialisation des parametres de base : type GB / MB ou %
		TypeC.SelectedIndex = 0 ' GB
		TypeD.SelectedIndex = 2	' %
		SizeC.disabled = True
		SizeD.disabled = True
		TypeC.disabled = True
		TypeD.disabled = True
		SizeC.Value=50
		SizeD.Value=100	
	'ElseIf Property("OSDPartitions")=2 Then
	Else 'If Property("OSDPartitions")=2 Then
			JDradio2.checked = True
			SizeC.disabled = False
			SizeD.disabled = False
			TypeC.disabled = False
			TypeD.disabled = False
			TypeC.Value = Property("OSDPartitions0SIZEUNITS")
			TypeD.Value = Property("OSDPartitions1SIZEUNITS")
			SizeC.Value = Property("OSDPartitions0SIZE")
			SizeD.Value = Property("OSDPartitions1SIZE")
			
	End If
	
	If Property("Role001")="--- SERVEUR ---" Then JDradio2.checked = True
	
	oDisk.Disk = 0
	UpdateDisk
	
	InvalidConfDisk.style.display = "none"
	
' InitializeGhost
'
'
	Dim strSN 'SerialNumber
	Dim strGF 'GhostFolder
	Dim strGFF	'GhostFolderFound
	Dim subf1, subf2
	Dim bGF
	
	Dim strWimFF
	
	document.getElementById("ServLabel").style.display = "none"
	strSN = oEnvironment.Item("SerialNumber")
 	
 	'SrvWimChemin.Value = "\\fmi-data\DATA\_DEP\ISOs - IMAGES\WIM"
 	'SrvWimChemin.Value = "\\sove\WIM"
 	SrvWimChemin.Value = oEnvironment.Item("DeployRoot") & "\Captures"
 	'SrvWimChemin.Value = "\\dx\WIM"
 	
 	Dim bServer
 	bServer = False
 	If oFso.FolderExists(SrvWimChemin.Value) Then
 		bServer = True
 	Else
		oShell.Run "net use " & chr(34) & SrvWimChemin.Value & chr(34) & " /user:sauve@dep s@uve", 1, true
		If oFso.FolderExists(SrvWimChemin.Value) Then bServer = True
 	End If
 	
 	If bServer = True Then
 		document.getElementById("BackupLocationLabel").innerHTML= SrvWimChemin.Value & "\" & oEnvironment.Item("Model") & "\" & oEnvironment.Item("SerialNumber") & "<br/>(Recopie auto sur le poste sur C:)"
 	Else
 		EspaceBackupWim.style.display="none"
 	End If
 	
 	' Function testGhost sur DX et si non disponible sur FMI-DATA, le serveur DX etant le plus a jour
 	
 	strGF = "\\sove\ghost"
 	If oFso.FolderExists(strGF) Then
 		bServer = True
 	Else
		oShell.Run "net use " & chr(34) & strGF & chr(34) & " /user:sauve@dep s@uve", 1, true
		If oFso.FolderExists(strGF) Then
			bServer = True
		Else
		 	strGF = "\\fmi-data\DATA\_DEP\ISOs - IMAGES\GHOST"
 			If oFso.FolderExists(strGF) Then
 				bServer = True
 			Else
				oShell.Run "net use " & chr(34) & strGF & chr(34) & " /user:dep@dep dep", 1, true
				If oFso.FolderExists(strGF) Then bServer = True
			End If
		End If
 	End If
 	'strGF = "\\dx\sove\Ghost"
 	GDRadio2.checked = True
 	
 	bGF=False
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
					document.getElementById("GhostTitre").innerHTML = "Ghost de " & GhostSize.Value
					bGF = True
				End If
			Next
		Next
	End If
	
	
	
	If not bGF Then
		GDRadio1.disabled = True
		document.getElementById("GhostTitre").innerHTML = "Aucun Ghost"
		SrvChemin.Value = "Non disponible"
		GhostChemin.Value = "Non disponible"
		GhostModule.style.display = "none"
		Espace1.style.display = "inline"
		Espace2.style.display = "inline"
		Espace3.style.display = "inline"
	Else
		Espace1.style.display = "none"
		Espace2.style.display = "none"
		Espace3.style.display = "none"
	End If
	
End Function


Function ValidatePartitions

	If OptBackupWim.checked=True Then
		WimFileToCopy.Value = oEnvironment.Item("Model") & "\" & oEnvironment.Item("SerialNumber")
		oEnvironment.Item("ComputerBackupLocation") = "" & SrvWimChemin.Value & "\" & WimFileToCopy.Value & ""
		oEnvironment.Item("BackupFile") = UCase(oEnvironment.Item("SerialNumber")) & ".wim"
		oEnvironment.Item("BackupDrive") = "ALL"
		WimSrvPath.Value = oEnvironment.Item("ComputerBackupLocation")
		WimFilePath.Value = oEnvironment.Item("BackupFile")	
		
	Else
		oEnvironment.Item("ComputerBackupLocation") = "NONE"
	End If

	If OptionMig.checked then
		oProperties("SkipUserData") = "NO"
	Else
		oProperties("SkipUserData") = "YES"
	End If

	If KeepPartitions.Checked then
		oProperties("doNotFormatAndPartition") = "YES"
		
		
		JDRadio1.disabled = True
		JDRadio2.disabled = True
		SizeC.disabled = True
		SizeD.disabled = True
		TypeC.disabled = True
		TypeD.disabled = True
		
		ValidatePartitions = True
	Else

	RemovePropertyIfPresent "doNotFormatAndPartition"
	JDRadio1.disabled = False
	JDRadio2.disabled = False
	ValidatePartitions = False
	
	If JDRadio1.checked Then
		SizeC.disabled = True
		SizeD.disabled = True
		TypeC.disabled = True
		TypeD.disabled = True
		oEnvironment.Item("OSDPartitions")=""
		ValidatePartitions = True
	ElseIf ValidateSizeDisk Then
		SizeC.disabled = False
		SizeD.disabled = False
		TypeC.disabled = False
		TypeD.disabled = False
		oEnvironment.Item("OSDPartitions")=2
		oEnvironment.Item("OSDPartitions0SIZEUNITS")=CStr(document.getElementById("TypeC").Value)
		oEnvironment.Item("OSDPartitions1SIZEUNITS")=CStr(document.getElementById("TypeD").Value)
		oEnvironment.Item("OSDPartitions0SIZE")=SizeC.Value
		oEnvironment.Item("OSDPartitions1SIZE")=SizeD.Value
		oEnvironment.Item("OSDPartitions1VOLUMENAME")="Data"
		ValidatePartitions = True
	Else
		ValidatePartitions = False
	End If
 	End if
 
End Function

Function ValidateSizeC
	' Check Warnings
	ParseAllWarningLabels
	
	If IsNumeric(SizeC.Value) Then
		If TypeC.SelectedIndex = 2 Then
			If not ( SizeC.Value > 0 and SizeC.Value <= 100 ) Then
			InvalidCharC.style.display = "none"
			OnlyC.style.display = "inline"
			ValidateSizeC = False
			ButtonNext.disabled = True
			End If
		Else 
			InvalidCharC.style.display = "none"
			OnlyC.style.display = "none"
			ValidateSizeC = False
			ButtonNext.disabled = True
			ValidateSizeC = TRUE
		End If
	ElseIf IsValidNumber ( SizeC.Value ) Then
		ValidateSizeC = TRUE
		InvalidCharC.style.display = "none"
		OnlyC.style.display = "none"
		ButtonNext.disabled = False
	Else
		InvalidCharC.style.display = "inline"
		OnlyC.style.display = "none"
		ValidateSizeC = False
		ButtonNext.disabled = True
	End If
End Function

Function ValidateSizeD
	' Check Warnings
	ParseAllWarningLabels
	
	If IsNumeric(SizeD.Value) Then
	If TypeD.SelectedIndex = 2 Then
			If not ( SizeD.Value > 0 and SizeD.Value <= 100 ) Then
			InvalidCharD.style.display = "none"
			OnlyD.style.display = "inline"
			ValidateSizeD = false
			ButtonNext.disabled = true
			End If
		Else 
			InvalidCharD.style.display = "none"
			OnlyD.style.display = "none"
			ValidateSizeD = False
			ButtonNext.disabled = True
			ValidateSizeD = TRUE
		End If
	ElseIf IsValidNumber ( SizeD.Value ) Then
		ValidateSizeD = TRUE
		InvalidCharD.style.display = "none"
		OnlyD.style.display = "none"
		ButtonNext.disabled = False
	Else
		InvalidCharD.style.display = "inline"
		OnlyD.style.display = "none"
		ValidateSizeD = False
		ButtonNext.disabled = True
	End If
	
End Function

' Validation globale pour la configuration et vérification des abérrations
Function ValidateSizeDisk
		
	ValidateSizeC
	ValidateSizeD
	
		If (not TestSize(SizeDisk)) and JDRadio2.checked Then 
			InvalidConfDisk.style.display = "inline"
			ButtonNext.disabled = True
			ValidateSizeDisk = False
		Else
			InvalidConfDisk.style.display = "none"
			ButtonNext.disabled = False
			ValidateSizeDisk = True
		End If
		
End Function
	
Function TestSize(nSize)
	Dim cRatio
	Dim dRatio
	
	TestSize = True
	
		If TypeC.SelectedIndex < 2 And TypeD.SelectedIndex < 2 And IsNumeric(SizeC.Value) And IsNumeric(SizeD.Value) Then
			cRatio = 1000^TypeC.SelectedIndex
			dRatio = 1000^TypeD.SelectedIndex
			If ((SizeC.Value/cRatio)+(SizeD.Value/dRatio)) > nSize Then TestSize = False
		ElseIf TypeC.SelectedIndex < 2 And TypeD.SelectedIndex = 2 And IsNumeric(SizeC.Value) Then
			cRatio = 1000^TypeC.SelectedIndex
			If (SizeC.Value/cRatio) > nSize  Then TestSize = False
		ElseIf TypeC.SelectedIndex = 2 And TypeD.SelectedIndex < 2 And IsNumeric(SizeD.Value) Then
			dRatio = 1000^TypeD.SelectedIndex
			If (SizeD.Value/dRatio) > nSize Then TestSize = False
		End If
End Function


Function IsValidNumber ( ByVal Num )
    Dim regEx
    'RegExMatch=False
    Set regEx = New RegExp     
    regEx.Pattern = "[^0-9\-\_]"
		IsValidNumber = not regEx.Test ( Num )
End Function 

Function ValidateGhostDetect_Final
	If GDRadio1.checked Then
		SrvChemin.Value = SrvChemin.Value & "\" & GhostChemin.Value
		GhostChemin.Value = "C:\" & GhostChemin.Value & "\"
		
		'Applications900.Value="{e93a875e-93c1-4e0d-9abc-8e8a773b7114}"
		ValidateGhostDetect_Final = True
	Else 
		If GDRadio2.checked Then 
			ValidateGhostDetect_Final = True
		Else
			ValidateGhostDetect_Final = False
		End If
	End If
End Function