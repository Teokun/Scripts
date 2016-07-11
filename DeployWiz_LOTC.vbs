'
'
' // Module Ecran de démarrage
Option Explicit

' // Déclaration des variables
Dim objDisk
Dim colParts
Dim colDisks
Dim objWMI
Dim TailleDisk
Dim sBackupFile
Dim strWimSRV
Dim sWimBackupFile
Dim strWimFound
Dim oDiskSelect
Dim sOSBuild
Dim sArchitecture
Dim sDestinationDrive
Dim sImageIndex
Dim sImagePath
Dim strSN
Dim bWimFound

Set objDisk = New ZTIDisk
Set objWMI = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colParts = objWMI.ExecQuery ("Select * from Win32_LogicalDisk Where DriveType='3' ")
Set colDisks = objWMI.ExecQuery ("Select * from Win32_DiskDrive")

'
'			Initialization
'
Function InitializeLOTC
	
	' Désactivation du bouton suivant
	ButtonNext.disabled = True
	Dim sCmd, oDiskOption
	
	'' Initialisation pour réseau LANDEP
	'strWimSRV = "\\fmi-data\SOVCLI\WIM\_Clients"
	'strWimSRV = "\\sove\WIM\_Clients"
	strWimSRV = "Z:\Captures"
	strSN = UCase(oEnvironment.Item("SerialNumber"))
	sWimBackupFile = strSN & ".wim"
	
	DataPath.Value = "Z:\Captures"
	'DataPath.Value = strWimSRV
	
	' Montage du lecteur réseau
	
	If not oFso.FolderExists(strWimSRV) Then
		sCmd = "net use w: " & chr(34) & strWimSRV & chr(34) & " /user:dep\dep dep"
		oShell.Run sCmd, 1, true
	End If
	
	If oFso.FolderExists(strWimSRV) Then
		LanDEPopt.disabled = False
		LanDEPopt.checked = True
	Else
		LanDEPopt.disabled = False
		LanDEPopt.checked = False
	End If
	
	If oFso.FolderExists("w:") Then
		strWimSRV = "w:"
		WimSrvPath.Value = "w:"
	End If
	
	Set oDiskSelect=document.getElementById("hDiskSelect")
		

	
	If colParts.Count < 1 Then
		NoDisk.style.Display = "inline"
		NoPart.style.Display = "inline"
		Exit function
	End if

 	If colDisks.Count > 0 Then
		For Each oDiskOption in oDiskSelect.Options 
			oDiskOption.RemoveNode
		Next
		' Populate liste des disques
		For each oDiskOption in colDisks
			SetDiskOption oDiskOption.index + 1 & "/" & colDisks.Count & " " &  oDiskOption.caption & " (" &  Round(oDiskOption.Size / 1024 / 1024 /1024,0) & " Go)",oDiskOption.index
		Next
	End If
 
 	WimIndex.Value = 1
	
	' Si Destination non défini, on le fixe avec le numéro de série
	If sBackupFile <> "" then
		' Already set, leave it alone
	ElseIf oEnvironment.Item("BackupFile") = "" then
		sBackupFile = strSN & ".wim"
	Else
		sBackupFile = oEnvironment.Item("BackupFile")
	End if
	
	ItemChangeLOTC
	ValidateDiskLOTC
	MajTableauParts
	MajCaptureDescLOTC
 
End Function	

Function ValidateLanDEP
	
	If LanDEPopt.checked = True Then
		CaptureModuleDep.style.display="inline"
		CaptureModule1.style.display="none"
	Else
		CaptureModuleDep.style.display="none"
		CaptureModule1.style.display="inline"
		Exit Function
	End If
	
	Dim subf1, subf2
	Dim str,str2
	
	bWimFound = False
	'alert(strSN)
 	If oFso.FolderExists(strWimSRV) Then
 		WimSrvPath.Value = strWimSRV
 		For Each subf1 in oFso.GetFolder( strWimSRV ).SubFolders
			For Each subf2 in subf1.SubFolders
				If LCase(cstr(subf2.Name)) = LCase(strSN) Then
					bWimFound = True
					strWimFound = cstr(subf2.Path)
					'alert(strWimFound)
					'WimSrvPath.Value = Right(strWimFound,Len(strWimFound)-Len(strWimSRV)-1)
					str2 = Right(strWimFound,Len(strWimFound)-Len(strWimSRV)-1)
					document.getElementById("labelWimFileExist").innerHTML="Dossier "& str2 &" Existant Taille : " & Round(oFso.GetFolder(strWimFound).Size / 1024 / 1024 / 1024 ,2) & " Go"
					NomClientFMI.Value = Left(str2,Instr(1,str2,"-")-1)
					str = Right(str2, Len(str2) - Len (NomClientFMI.Value)-1)
					NumFMI.Value = Left(str,Instr(1,str,"\")-1)
					
				End If
			Next
		Next
	Else
		WimSrvPath.Value = ""
	End If	

	If bWimFound Then
		NomClientFMI.disabled = True
		NumFMI.disabled = True
	Else
		NomClientFMI.disabled = False
		NumFMI.disabled = False
	End If

End Function

Function MajTableauParts 
  Dim oSelPart
  Dim sFilteredXML
  Dim usedSize
  Dim totSize
  Dim freeSize
  For each oSelPart in objWMI.ExecQuery ("Select * from Win32_LogicalDisk Where DriveType='3' ")
  If oSelPart.Size > 0 and oSelPart.DeviceID <> "X:" Then
  	'usedSize = Round((oSelPart.Size - oSelPart.FreeSpace) / 1024 / 1024 / 1024 ,2)
  	totSize = Round(oSelPart.Size / 1024 / 1024 / 1024 ,2)
  	freeSize =  Round( oSelPart.FreeSpace / 1024 / 1024 / 1024 ,2)
  	
  	sFilteredXML = sFilteredXML & "<SelectablePart>"
  	sFilteredXML = sFilteredXML & "<part>"
  	sFilteredXML = sFilteredXML & oSelPart.DeviceID
  	sFilteredXML = sFilteredXML & "</part>"
  	sFilteredXML = sFilteredXML & "<comments>"
		sFilteredXML = sFilteredXML & oSelPart.VolumeName  & " (" & oSelPart.FileSystem & ")"
  	sFilteredXML = sFilteredXML & "</comments>"
  	sFilteredXML = sFilteredXML & "<PartSize>"
  	sFilteredXML = sFilteredXML & Round(freesize * 100 / totSize, 1) & "% libres -- " & Round(oSelPart.FreeSpace / 1024 / 1024 / 1024 ,2) & " Go libres sur " & Round(oSelPart.Size / 1024 / 1024 / 1024 ,2) & " Go"
  	sFilteredXML = sFilteredXML & "</PartSize>"
  	sFilteredXML = sFilteredXML & "</SelectablePart>"
  End If
	Next
 If not IsEmpty(sFilteredXML) then
  	bparts.XMLDocument.LoadXML "<SelectablePart>" & sFilteredXML & "</SelectablePart>"
 End if

End Function

Function ItemChangeLOTC

 If OptionCapture.Checked = True Then 
 	If CaptureDiskMode.checked Then
 		CaptureDiskModule1.style.display = "inline"
 		CaptureDiskModuleDep.style.display = "inline"
 		CaptureBrowser.style.display = "none"
 		Partlabel.style.display = "none"
 	Else
 		CaptureDiskModule1.style.display = "none"
 		CaptureDiskModuleDep.style.display = "none"
 		CaptureBrowser.style.display = "inline"
 		Partlabel.style.display = "inline"
 	End If
 		CaptureModule.style.display = "inline"
 		RestoreModule.style.display = "none"
 	'	HDinfo.style.display = "none"
 		RestoreBtn.style.display = "none"
 		CaptureBtn.style.display = "inline"
 		document.getElementById("PartLabel").innerHTML="Selectionner partition(s) source(s)"
 		DiskToWimBtn.style.display="inline"
 		DiskToVHDBtn.style.display="inline"
 		DiskToVHDDesc.style.display="inline"
 		VPC.style.display="inline"
 		'document.getElementById("RebootLabel").innerHTML="Reboot Automatique en fin de capture"
 	Else
 		Partlabel.style.display = "inline"
 		CaptureBrowser.style.display = "inline"
 		CaptureModule.style.display = "none"
 		RestoreModule.style.display = "inline" 		
 	'	HDinfo.style.display = "inline"
 		RestoreBtn.style.display = "inline"
 		CaptureBtn.style.display = "none"
 		document.getElementById("PartLabel").innerHTML="Selectionner 1 Partition Cible"
 		DiskToWimBtn.style.display="none"
 		DiskToVHDBtn.style.display="none"
 		DiskToVHDDesc.style.display="none"
 		VPC.style.display="none"
 		'document.getElementById("RebootLabel").innerHTML="Reboot Automatique en fin de restauration"
 		
 	End If

End Function

'' Fonction de validation des choix du disque sélectionné
Function ValidateDiskLOTC

		objDisk.Disk = hDiskSelect.Value
		oEnvironment.Item("OSDDiskIndex") = hDiskSelect.Value
	
	UpdateDiskLOTC

End Function

''  Fonction de mises à jours des informations relatif au disque sélectionné
Function UpdateDiskLOTC
	
	TailleDisk = Round(objDisk.oWMI.Size / 1024 / 1024 /1024,0)
	
	document.getElementById("Disk0Title").innerHTML = objDisk.Disk & " : " & objDisk.oWMI.Caption 
	
	If oEnvironment.Item("OSVersion") <> "WinPE" Then
		If ( TailleDisk < 127 ) and UCase(oEnvironment.Item("OSArchitecture")) <> "AMD64" Then
			VPC.style.background="green"
			document.getElementById("VPC").innerHTML="Le VHD sera compatible VirtualPC (HD<127Go / OS x86)"
		Else
			VPC.style.background="red"
			document.getElementById("VPC").innerHTML="Le VHD ne sera pas compatible VirtualPC (> 127Go ou non x86)"
		End If
	Else
			If ( TailleDisk < 127 ) Then
			VPC.style.background="yellow"
			document.getElementById("VPC").innerHTML="Le VHD sera compatible VirtualPC (HD<127Go / OS x86)"
		Else
			VPC.style.background="red"
			document.getElementById("VPC").innerHTML="Le VHD ne sera pas compatible VirtualPC (> 127Go ou non x86)"
		End If
	End If

End Function


'' Fonction restrictive des modules autorisée ou pas en WinPE ou mode OS online
Function MajCaptureDescLOTC

	If DataPath.Value <> ""  and oFSO.FolderExists(DataPath.Value) Then
		If oFso.FileExists( DataPath.Value & "\" & sWimBackupFile ) Then
			document.getElementById("CaptureDesc").InnerHTML = "Fusion sauvegarde avec " & DataPath.Value & "\" & sWimBackupFile 
			CaptureDesc.style.background = "lightblue"
		Else
			document.getElementById("CaptureDesc").InnerHTML = "Sauvegarde sous " & DataPath.Value & "\" & sWimBackupFile 
			CaptureDesc.style.background = "lightgreen"
		End If
 		CaptureBtn.disabled = False
 		DiskToWimBtn.disabled = False or oEnvironment.Item("OSVersion") <> "WinPE"
 		DiskToVHDBtn.disabled = False or oEnvironment.Item("OSVersion") = "WinPE"
 	Else
 		document.getElementById("CaptureDesc").InnerHTML = ""
 		CaptureBtn.disabled = True
 		DiskToWimBtn.disabled = True
 		DiskToVHDBtn.disabled = True
 	End If
End Function

Function MajCaptureDescDep

	If NumFMI.Value <> "" And NomClientFMI.Value <> "" And WimSrvPath.Value <> "" Then
		If bWimFound Then
			document.getElementById("CaptureDescDep").InnerHTML = "Fusion sauvegarde avec " & sWimBackupFile
			CaptureDescDep.style.background = "lightblue"
		Else
			document.getElementById("CaptureDescDep").InnerHTML = "Sauvegarde sous " & WimSrvPath.Value & "\" &NomClientFMI.Value & "-" & NumFMI.Value & "\"& strSN & "\" & sWimBackupFile 
			CaptureDescDep.style.background = "lightgreen"
		End If
 		CaptureBtn.disabled = False
 		DiskToWimBtn2.disabled = False or oEnvironment.Item("OSVersion") <> "WinPE"
 		DiskToVHDBtn.disabled = False or oEnvironment.Item("OSVersion") = "WinPE"
 	Else
 		document.getElementById("CaptureDescDep").InnerHTML = ""
 		CaptureBtn.disabled = True
 		DiskToWimBtn2.disabled = True
 		DiskToVHDBtn.disabled = True
 	End If
End Function

''''''''' Fonction formatage du disque sélectionné
Function FormatDiskLOTC
	Dim sCmd, iRc, ret
	
	' Definition des paramètre par défaut pour le formatage du HDD
	ValidateDiskLOTC
	oEnvironment.Item("OSDPartitions") = 1
	
	sImagePath = WimSourcePath.Value
	
	If Not oFSO.FileExists( sImagePath ) or sImagePath = "" Then
		Alert "Commencer par choisir votre image source !"
		Exit Function 
	End If
	
	oEnvironment.Item("OSDDiskIndex") = CInt(hDiskSelect.Value)
	' Set value for formatting depending on imageBuild
	oEnvironment.Item("ImageBuild") = SelectWimInfoLOTC ( sImagePath , "Version" )
	oEnvironment.Item("DEPLOYMENTTYPE") = "NEWCOMPUTER"
	oEnvironment.Item("OSDPartitions0SIZEUNITS")="%"
	oEnvironment.Item("OSDPartitions0SIZE")="100"
	oEnvironment.Item("OSDPartitions")= 1
	oEnvironment.Item("OSDPartitions0VOLUMENAME")="System"
	
	If Not MsgBoxConfirm("Vous allez formater le disque " & Property("OSDDiskIndex") & ", toutes les donnees seront perdus !" ) Then 
		Exit Function
	End If
	
		'Démarrage du formatage !!
	sCmd = "cscript.exe """ & oUtility.ScriptDir & "\ZTIDiskpart.wsf /debug:true"""
	iRc = RunCmdShell ( sCmd )
	
	'Alert(iRc)
	
	' Rafraichissement du module
	MajTableauParts
	'window.location.reload()
End Function

Function PartitionWizardLOTC
	Dim sCmd
	
	If oEnvironment.Item("Architecture") = "X64" then 
		Alert("Non pris en charge sur x64" & vbcrlf & "Veuillez utiliser diskpart" &vbcrlf& "(F8) pour ouvrir une console CMD")
		Exit Function
	Else
		sCmd = "x:\Tools\Programs\PartitionWizard\PartitionWizard.exe"
	End If
	
	document.body.style.cursor = "Wait"
	
	If oFso.FileExists(sCmd) Then
		oShell.Run sCmd, 1, true
		MajTableauParts
	Else
		sCmd = "z:\Tools\Programs\PartitionWizard\PartitionWizard.exe"
		If oFso.FileExists(sCmd) Then
			oShell.Run sCmd, 1, true
			MajTableauParts
		Else
			Alert("Programme introuvable...")
		End If
	End If
	
	document.body.style.cursor = "default"
	
End Function

Function ApplyImage1
	Dim sCmd, sParameters, sFoundFile
	Dim	objDest, objNode, iRC, objBootDrive
	Dim cpt

	' Si aucun fichier image sélectionner, exit function
	sImagePath = WimSourcePath.Value
	sImageIndex = WimIndex.Value

	If Not oFSO.FileExists( sImagePath ) or sImagePath = "" Then Exit Function 
	
	If not ValidateWimIndexLOTC Then
		Exit Function
	End If
	
	Set objDest = document.getElementsByName("SelectedItem")
	cpt = 0
	For Each objNode in objDest
		If document.all.item(objNode.SourceIndex).checked Then 
			sDestinationDrive = document.all.item(objNode.SourceIndex + 1).Value
			cpt = cpt +1
		End If
 	Next
	' on s'assure de n'avoir qu'une seule partition sélectionner
	If cpt <> 1 Then
		Alert("Veuillez choisir 1 partition !")
		Exit Function
	End If
	
	If Not MsgBoxConfirm("Vous allez restaurer " & sImagePath & ":" & sImageIndex & " sur " & sDestinationDrive  ) Then 
		Exit Function
	End If

	oEnvironment.Item("sWimImagePath") = chr(34) & sImagePath & chr(34)
	oEnvironment.Item("sWimImageIndex") = sImageIndex
	oEnvironment.Item("sWimImageBuild") = SelectWimInfoLOTC ( sImagePath , "Version" )
	oEnvironment.Item("sWimDestinationDrive") = sDestinationDrive
	oEnvironment.Item("sWimImageProcessor") = Left (SelectWimInfoLOTC ( sImagePath , "Architecture"), 3 )
	
	sCmd = "cscript.exe """ & oUtility.ScriptDir & "\ZTIApply.wsf"""
	iRc = RunCmdShell ( sCmd )
			
	If iRc <> 0 then
		'If iRc = 2 Then
		Msgbox "Erreur : " & iRc & " - not specified", vbExclamation
		Exit Function
	Else
	
	End If
	
	
	Alert("Redemarrer le poste pour continuer !")
	
End Function

Function GetSourcePathWim
	GetSourcePathWim = WimSourcePath.Value
End function

Function FindAndRunLOTC ( sCmd , sParameters )
	Dim iRC
	iRC = oUtility.FindFile ( sCmd, sFoundFile )
	iRC = """" & sFoundFile & """ " & sParameters
	iRC = RunCmdShell ( sCmd )
End Function

Function TestBCDLOTC
	Dim oBootDrive
	Set oBootDrive = GetBootDriveEx ( true, "0.0.0.0", false )
	If oBootDrive is Nothing Then
		Alert("Pas de Partition Bootable" & vbcrlf & "Utilisez PartitionWizard pour rendre une partition bootable (primary/active)")
		Exit Function
	End If
	
	Dim bootOK
		If ucase(oEnvironment.Item("IsUEFI")) = "TRUE" then
			If Not oFSO.FileExists(Left(oBootDrive.Drive,2) & "\efi\microsoft\boot\bcd") Then bootOK = False
		Else
			If Not oFSO.FileExists(Left(oBootDrive.Drive,2) & "\boot\bcd") Then bootOK = False
		End If
		
	If Not bootOK Then
		Alert("Le magasin BCD n'existe pas ! Le pc ne peux booter sans le BCD")
		TestBCDLOTC = False
	Else
		Msgbox "Le BCD existe, la machine devait booter normalement", vbInformation
		TestBCDLOTC = True
	End If
	
End Function


Function CaptureImageLOTC
	Dim sCmd, iRc, iErr
	Dim oDest, oNode	
	Dim sDestinationDrive
	Dim strBackupTarget
	Dim sCurrentDrive
	Dim sDriveErr
	
	If LanDEPopt.checked = True and ( NumFMI.Value = "" or NomClientFMI.Value = "" ) Then Exit Function
	
	If LanDEPopt.checked = False and DataPath.Value = "" Then Exit Function
	
	If LanDEPopt.checked = True Then
		strBackupTarget = WimSrvPath.Value & "\" &NomClientFMI.Value & "-" & NumFMI.Value & "\"& strSN
	Else
		strBackupTarget = DataPath.Value 
	End If
	
	sDestinationDrive = ""
	Set oDest = document.getElementsByName("SelectedItem")
	
	For Each oNode in oDest
 		If document.all.item(oNode.SourceIndex ).checked Then sDestinationDrive = sDestinationDrive & " " & document.all.item(oNode.SourceIndex + 1).Value
 	Next
 	
	If sDestinationDrive="" Then
		Alert "Selectionnez au moins 1 partition"
		Exit Function
	End If
	
	If Not MsgBoxConfirm("Capture de " & sDestinationDrive ) Then 
		Exit Function
	End If
	
	oEnvironment.Item("ComputerBackupLocation") = strBackupTarget	
	oEnvironment.Item("BackupFile") = sBackupFile	
	sDriveErr = ""
	sDestinationDrive = ""
	
	For Each oNode in oDest 	' démarrage capture pour chaque partition sélectionner

 		If document.all.item(oNode.SourceIndex ).checked Then 
			sCurrentDrive = document.all.item(oNode.SourceIndex + 1).Value
			
			oEnvironment.Item("BackupPartition") = sCurrentDrive
			sCmd = "cscript.exe """ & oUtility.ScriptDir & "\ZTIBackup.wsf"""
			iRc = RunCmdShell ( sCmd )			
			
			If iRc <> 0 then
				'If iRc = 2 Then msgbox "Erreur : " & iRc & " - Echec de la capture", vbExclamation
				sDriveErr = sDriveErr & " " &  sCurrentDrive
			Else
				sDestinationDrive = sDestinationDrive & " " &  sCurrentDrive
			End If

		End If
 	Next
	
	If sDriveErr <> "" Then msgbox "Erreur de capture de : " & sDriveErr & vbcrlf & "Capture impossible, Verifiez le disque ou l'OS..."& vbcrlf & "Des mises a jour sont peu-etre en attente d'application...", vbExclamation
	
	If sDestinationDrive <> "" Then msgbox "Capture OK de " & sDestinationDrive & " vers " & strBackupTarget & "\" & sBackupFile, vbInformation
	
	'Alert("Image Backup de " & sDestinationDrive & " vers " & strBackupTarget  )
	
	oEnvironment.Item("ComputerBackupLocation") = ""
	oEnvironment.Item("BackupFile") = ""
	
End Function


Function DiskToWimLOTC
	Dim sCmd, iRc, objDrive1, iErr
	Dim colDiskPartitions, objPartition, objDrive
	Dim cPartitions,cLogicalDisks, oLDisk
	Dim aPartition, sListPartition
	
	If DataPath.Value = "" Then Exit Function
	
	objDisk.Disk = CInt(hDiskSelect.Value)
	
	' Destination de la sauvegarde
	oEnvironment.Item("ComputerBackupLocation") = DataPath.Value	
	oEnvironment.Item("BackupPartition") = ""
	oEnvironment.Item("BackupFile") = sBackupFile 
	sListPartition = ""
	
	If Not MsgBoxConfirm("Sauvegarde du disque " & hDiskSelect.Value + 1 ) Then 
		Exit Function
	End If
	
	' Récupération de la liste des partition du disque sélectionné
	iErr = ""
	
	For Each objDrive In colDisks
		If objDrive.Index = objDisk.Disk Then
		
			'WScript.echo "Disk " & objDrive.Index & " (" & objDrive.InterfaceType & "): " & objDrive.Caption
			Set cPartitions = objWMI.ExecQuery("ASSOCIATORS OF {Win32_DiskDrive.DeviceID=""" _
				& Replace(objDrive.DeviceID, "\", "\\") & """} WHERE AssocClass = " & "Win32_DiskDriveToDiskPartition")
		
			For Each objPartition In cPartitions
				Set cLogicalDisks = objWMI.ExecQuery _
					("ASSOCIATORS OF {Win32_DiskPartition.DeviceID=""" & objPartition.DeviceID _
					& """} WHERE AssocClass = Win32_LogicalDiskToPartition")
		
				For Each oLDisk In cLogicalDisks
				  'If oLDisk.DeviceID <> "C:" Then
						sListPartition = sListPartition & oLDisk.DeviceID & " "
						oEnvironment.Item("BackupDrive") = oLDisk.DeviceID 
						' Lancement de la sauvegarde pour chaque partition du disque
					
						sCmd = "cscript.exe """ & oUtility.ScriptDir & "\ZTIBackup.wsf"""
						iRc = RunCmdShell ( sCmd )
						If iRc <> 0 then iErr = iErr & objPartition.DeviceID & " "
					'End If
					
				Next
			Next
		End If
	Next
	
	If iErr <> "" then
		MsgBox "Erreur de sauvegarde : " & iErr & vbcrlf & "Capture impossible, Verifiez le disque ou l'OS..."& vbcrlf & "Des mises a jour sont peu-etre en attente d'application...", vbExclamation
	Else
		If oFso.FileExists(DataPath.Value & "\" & sBackupFile ) Then
				msgbox "Capture OK vers " & DataPath.Value & "\" & sBackupFile , vbInformation
		Else
				Alert(" Pas de Sauvegarde ! ")
		End If
	End If
	
	oEnvironment.Item("ComputerBackupLocation") = ""
	oEnvironment.Item("BackupFile") = ""
	oEnvironment.Item("BackupDrive") = ""
	
End Function

Function DiskToWimDep
	Dim sCmd, iRc
	Dim sb
	Dim strBackupTarget
	
	If WimSrvPath.Value = "" or NumFMI.Value = "" or NomClientFMI.Value = "" Then Exit Function
	
	
	' If NumFMI.Value <> "" And NomClientFMI.Value <> ""
	If WimSrvPath.Value <> "" Then
		strBackupTarget = WimSrvPath.Value & "\" &NomClientFMI.Value & "-" & NumFMI.Value & "\"& strSN
	Else
		
		If NumFMI.Value = "" Or NomClientFMI.Value = "" Then Exit Function
		strBackupTarget = strWimSRV & "\" &NomClientFMI.Value & "-" & NumFMI.Value & "\"& strSN
		
	End If
	
	oEnvironment.Item("BackupDrive") = "ALL"
	oEnvironment.Item("ComputerBackupLocation") = strBackupTarget
	Alert("Disk Backup vers " & strBackupTarget )
	'sb = sBackupFile
	oEnvironment.Item("BackupFile") = sWimBackupFile
	sCmd = "cscript.exe """ & oUtility.ScriptDir & "\ZTIBackup.wsf"""
	iRc = RunCmdShell ( sCmd )
	
	If iRc <> 0 then
		'If iRc = 2 Then 
		MsgBox "Erreur de sauvegarde : " &iRc& vbcrlf & "Capture impossible, Verifiez le disque ou l'OS..."& vbcrlf & "Des mises a jour sont peu-etre en attente d'application...", vbExclamation
		Exit Function
	Else
		msgbox "Capture OK vers " & strBackupTarget , vbInformation		
	End If
	
	oEnvironment.Item("ComputerBackupLocation") = ""
	oEnvironment.Item("BackupFile") = ""
	oEnvironment.Item("BackupDrive") = ""
	
	ValidateLanDEP
	
End Function

Function fDisk2VHDLOTC
	Dim sCmd, iRc, strBackupTarget
	
	If LanDEPopt.checked = True Then
		If WimSrvPath.Value = "" or NumFMI.Value = "" And NomClientFMI.Value = "" Then Exit Function
		strBackupTarget = WimSrvPath.Value & "\" &NomClientFMI.Value & "-" & NumFMI.Value & "\"& strSN
	Else
		If Not oFso.FolderExists(DataPath.Value) Then Exit Function
		strBackupTarget = DataPath.Value
	End If

	'oEnvironment.Item("ComputerBackupLocation") = strBackupTarget
	
	If Not MsgBoxConfirm("En continuant votre machine physique changera de nom ! " & Property("OSDDiskIndex") & ",Voulez-vous continuer ?" ) Then 
		Exit Function
	End If
	
	' If DataPath.Value = "" Then Exit Function
	
	oEnvironment.Item("ComputerBackupLocation") = strBackupTarget
	oEnvironment.Item("TaskSequenceID")="P2V"
	
	oEnvironment.Item("SkipApplications")="YES"
	oEnvironment.Item("SkipComputerName")="YES"
	oEnvironment.Item("SkipTaskSequence")="YES"
	oEnvironment.Item("SkipRoleConfig")="YES"

	SaveAllDataElements
	SaveProperties
	
	ButtonNext.disabled=False
	ButtonNextClick
	
End Function


'''''''' Fonction affichage du contenu du WIM sélectionner

Function GetWimIndexLOTC
	Dim sCmd, iRc, objReadFile
	
	If Not oFSO.FileExists(WimSourcePath.Value) Then Exit Function
	sCmd = "dism.exe /get-wiminfo /Wimfile:" & chr(34) &  WimSourcePath.Value & chr(34)
	
	Set objReadFile = oFso.OpenTextFile("X:\WINDOWS\TEMP\wiminfo.txt", 2, 1)
	objReadFile.Write ExecShellAnsi ( sCmd )
	objReadFile.close
	sCmd = "notepad.exe X:\WINDOWS\TEMP\wiminfo.txt"
	RunCmdShell ( sCmd )
End Function


Function GetWimIndexInfoLOTC
	Dim sCmd, iRc
	On Error Resume Next
	If Not oFSO.FileExists(WimSourcePath.Value) Then Exit Function	
	sCmd = "dism.exe /get-Imageinfo /Imagefile:" & chr(34) &  WimSourcePath.Value & chr(34) & " /index:" & WimIndex.Value
	msgBox ExecShellAnsi(sCmd)
	On Error Goto 0		
End Function

Function TestInfo
	'SelectWimInfoLOTC WimSourcePath.Value, inTest.Value
End Function

Function SelectWimInfoLOTC ( sWim, sInfo )
	Dim sCmd, iRc, strData, intStart, strText, intStop
	On Error Resume Next
	If Not oFSO.FileExists(sWim) Then Exit Function	
	sCmd = "dism.exe /get-imageinfo /Imagefile:" & chr(34) & sWim & chr(34) & " /index:" & WimIndex.Value
	iRc = LCase(ExecShellAnsi(sCmd))
	
	intStart = InStr(iRc,"index")
	iRc = Mid( iRc, intStart, 1000 )
	
	intStart = InStr( iRc, LCase(sInfo) ) + Len(sInfo) + 3 
	iRc = Mid( iRc, intStart, 11 )
	
	SelectWimInfoLOTC = iRc
	
	On Error Goto 0		
End Function

' Conversion des caractères
Function ExecShellAnsi ( sCmd )
	Dim objExec
	Set objExec = oShell.Exec(sCmd)
	ExecShellAnsi=ToAnsi(objExec.StdOut.ReadAll)
End Function

Function ValidateWimIndexLOTC
	Dim sCmd, iRc
	
	If WimIndex.Value <> "" and oFSO.FileExists(WimSourcePath.Value) Then
	
		sCmd = "dism.exe /get-imageinfo /Imagefile:" & chr(34) &  WimSourcePath.Value & chr(34) & " /index:" & WimIndex.Value
		
		iRc = ExecShellAnsi ( sCmd )
		'Alert(iRc)
		If InStr(LCase(iRc),"error" ) > 0 Then 
			ValidateWimIndexLOTC = False
			Alert("Index non existant dans l'image")
		Else
			ValidateWimIndexLOTC = True
		End If
		
	End If
	
End Function

'''''''' Fonction Message de confirmation
Function MsgBoxConfirm ( sMessage )
	Dim ret
	MsgBoxConfirm = True
	ret = msgbox( sMessage & _
				vbCrLf & vbCrLf & _
				"Etes-vous sur de vouloir continuer ?" , vbYesNo + vbExclamation)
	
	If ret <> vbYes Then 
		MsgBoxConfirm = False
	End If
End Function

''''''' Fonction pour autoriser le formatage
Sub InitializeFormat
	If optionFormat.checked = True Then
		FormatBtn.disabled = False
	Else
		FormatBtn.disabled = True
	End If
End Sub

''''''' Fonction Lancement CMD
Function RunCmdShell ( sCmd )
	RunCmdShell = oShell.Run(sCmd, , true)
End Function

Sub SetDiskOption(OptText,OptValue) 
	Dim oNewOption
	Set oNewOption = Document.CreateElement("OPTION")
 	oNewOption.Text = OptText 
	oNewOption.Value = OptValue 
	oDiskSelect.options.Add(oNewOption) 
End Sub


Function ToAnsi (strOem)

Dim nCount
Dim strOemToCharCode: strOemToCharCode = _
    "\000\001\002\003\004\005\006\007\008\009\010" & _
    "\011\012\013\014\164\016\017\018\019\182\167" & _
    "\022\023\024\025\026\027\028\029\030\031\032" & _
    "\033\034\035\036\037\038\039\040\041\042\043" & _
    "\044\045\046\047\048\049\050\051\052\053\054" & _
    "\055\056\057\058\059\060\061\062\063\064\065" & _
    "\066\067\068\069\070\071\072\073\074\075\076" & _
    "\077\078\079\080\081\082\083\084\085\086\087" & _
    "\088\089\090\091\092\093\094\095\096\097\098" & _
    "\099\100\101\102\103\104\105\106\107\108\109" & _
    "\110\111\112\113\114\115\116\117\118\119\120" & _
    "\121\122\123\124\125\126\127\199\252\233\226" & _
    "\228\224\229\231\234\235\232\239\238\236\196" & _
    "\197\201\230\198\244\246\242\251\249\255\214" & _
    "\220\248\163\216\215\131\225\237\243\250\241" & _
    "\209\170\186\191\174\172\189\188\161\171\187" & _
    "\166\166\166\166\166\193\194\192\169\166\166" & _
    "\043\043\162\165\043\043\045\045\043\045\043" & _
    "\227\195\043\043\045\045\166\045\043\164\240" & _
    "\208\202\203\200\105\205\206\207\043\043\166" & _
    "\095\166\204\175\211\223\212\210\245\213\181" & _
    "\254\222\218\219\217\253\221\175\180\173\177" & _
    "\061\190\182\167\247\184\176\168\183\185\179" & _
    "\178\166\160"

    For nCount = 1 To Len (strOem)
     ToAnsi = ToAnsi & _
        Chr (Int ( _
         Mid ( _
            strOemToCharCode, Asc ( _
             Mid ( _
                strOem, nCount, 1)) * 4 + 2, 3)))
    Next

End Function

