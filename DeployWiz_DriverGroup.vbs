' //
' // Purpose:   Set DriverGroup wizard pane validation
' // 
' // ***************************************************************************

Option Explicit
'''''''''''''''''''''''''''''''''''''
'  Set DriverGroup Update pane
'
Dim Manufact
Dim oSelect

Function InitDriverGroup
	'UCase(oEnvironment.Item("Architecture")) & "\" & 
	Dim oItem
	Dim oDG
	Dim bFound
	Dim bMFound, bDFound 'Manufacturer Found & DirectoryFound
	Dim sDGName, sDGComment
	Dim oOption, sGeneric
	
	Dim oMember, sDriverCount
	
	Dim i
	Dim s
	Dim j, strHolder 
	Dim sParent
	Dim sModel
	Dim sgModel ' modèle global si contructeur LENOVO (xxxx-xxx)
	
	Dim oOSXMLselect
	Dim oOperatingSystems
	Dim sOSBuild
	Dim oOS
	
	Dim sOSPath
	
	
	' Groupe pour la sélection du dossier de pilotes correspondant à l'OS choisi pour l'installation
	Set oOSXMLselect =  New ConfigFile
	oOSXMLselect.sFileType = "OperatingSystems"
	
	Set oOperatingSystems = oOSXMLselect.FindItems
	Set oOS = oOperatingSystems.Item(oEnvironment.Item("OSGUID"))
	
	If Not oOS.SelectSingleNode("Build") Is Nothing Then 
		sOSBuild = oOS.SelectSingleNode("Build").Text
		document.getElementById("TLD1").innerHTML = Left(sOSBuild, 3)
	End If
	
	Select Case Left(sOSBuild, 3)
		Case "6.1":
			sOSPath = "\WINDOWS 7"
		Case "6.2","6.3", "10.":
			sOSPath = "\WINDOWS 8"
		Case Else :
			sOSPath = ""
	End Select
	
	'
	
	' Variable pour Constructeur\Modèle
	If Instr(1, UCase(oEnvironment.Item("Make")), "FUJITSU", 1) = 1 Then
		Manufact = "FUJITSU"
	Else
		Manufact = UCase(oEnvironment.Item("Make"))
	End If
	
	If UCase(oEnvironment.Item("IsServer")) = "TRUE" or Property("Role001")="--- SERVEUR ---" or Property("Role001")="--- E-BACKUP ---"  Then

		Select Case Manufact
			Case "FUJITSU" :
				sModel = "ServerView"
			Case "IBM" :
				sModel = "ServerGuide"
			Case "HP" :
				
			Case Else :
				sModel = "CHOISIR DANS LISTE PREDEFINIS"
		End Select
		
	Else
	
		' Fonctions de regroupement pour les pilotes LENOVO avec model XXXX
		If Instr(1, UCase(oEnvironment.Item("Make")), "LENOVO", 1) = 1 Then
				sgModel =  Left(UCase(oEnvironment.Item("Model")),4)
			Else
				sgModel = ""
		End If
		
		
		If sgModel <> "" Then 
			sModel = sgModel
		Else
			sModel = UCase(oEnvironment.Item("Model")) 
		End If
		
	End If
	'Initialisation chaine Constructeur\Modèle
	
	If DriverPath.Value = "" Then DriverPath.Value = Manufact & "\" & sModel

	bFound = False
	bMFound = False
	bDFound = False
	
	'Génération de la liste des drivers disponibles
	Set oSelect=document.getElementById("hDriverSelect")
	' Clean the List
	For Each oOption in oSelect.Options 
		oOption.RemoveNode
	Next
	
	' Recherche dans les fichiers de configurations si trouve Constructeur\Modèle
	
	' Récupération de tous les éléments pour la génération de la liste des drivers présent
	Set oDG = oUtility.CreateXMLDOMObjectEx(oEnvironment.Item("DeployRoot") & "\Control\DriverGroups.xml").selectNodes( "//*/*[@guid]" )
	
	
	i=0
	
	For each oItem in oDG
		' Récupération du noeud "Name"
		sDGName = oUtility.SelectSingleNodeString(oItem,"Name")
		' Récupération du noeud "Comments"
		
		On Error Resume Next
		sDGComment  ="---"
		sDGComment = oUtility.SelectSingleNodeString(oItem,"Comments")
		On Error Goto 0
		
		' Récupération du nombre de drivers pour le modèle		
		oMember = oItem.SelectNodes( "./Member" ).Length
		
		If Left(sDGName,1) <> "__" and sDGName <> "hidden" Then 
			ReDim Preserve arrListDG(i+1)
			
			If oMember <> 0 Then 
				sDriverCount = "> - " & oMember
			Else
				sDriverCount = ">"
			End If
			
			' Tableau contenant la liste des modèles
			arrListDG(i) = sDGName & " <" & sDGComment & sDriverCount
			i = i + 1
			
		End If
		
		'Si modèle exact trouvé AVEC Dossier "Version OS" alors on le sélectionne
		If ( Not bDFound ) And Instr(1,UCase(sDGName), DriverPath.Value & sOSPath ) = 1  Then
			bDFound = True
			DGModel.disabled=False
			DriverPath.Value =  UCase(sDGName)
			document.getElementById("ModelDriverCount").innerHTML = " (" & oMember & ")"
						If oMember = 0 Then msgbox " /!\ Attention, aucun pilote pour le model, Please Select Compatible Model /!\ "
						
		' Sinon on Sélectionne le model si disponible
		ElseIf ( Not bFound ) And Instr(1,UCase(sDGName), DriverPath.Value ) = 1 Then 
			bFound = True
			DGModel.disabled=False
			DriverPath.Value =  UCase(sDGName)
			document.getElementById("ModelDriverCount").innerHTML = " (" & oMember & ")"
						If oMember = 0 Then msgbox " /!\ Attention, aucun pilote pour le model, Please Select Compatible Model /!\ "
						
		' Sinon le Constructeur en dernier choix
		'ElseIf ( Not bMfound ) And Instr(1, sDGName, Manufact, 1 ) = 1 Then 
		
		'	bMFound = True
		
		End If
		
	Next
	
	'Sort(arrListDG)
	For i = ( UBound( arrListDG ) - 1 ) to 0 Step -1
		For j= 0 to i
			If UCase( arrListDG( j ) ) > UCase( arrListDG( j + 1 ) ) Then
				strHolder 				 = arrListDG( j + 1 )
				arrListDG( j + 1 ) = arrListDG( j )
				arrListDG( j )     = strHolder
			End If
		Next
	Next 
	
	' Populate Select Option List
	sParent = " "
	s = ""
	Dim pereOK, h, sTiret
	pereOK = True
	
	For i = 1 To UBound(arrListDG)
		If arrListDG(i) = "Generic <>" Then sGeneric = arrListDG(i)
		
		If arrListDG(i) <> "" Then 
		s = ""
			If i < UBound(arrListDG) Then
				' Test si le noeud actuel est un pere
				If Instr(1, arrListDG(i+1), Mid(arrListDG(i),1,Instr(1,arrListDG(i)," <",1)-1) , 1) Then

						sParent = Left(arrListDG(i+1),Instr(1,arrListDG(i+1),"\",1)-1)
						s = "+ " & arrListDG(i)
						pereOK = True

				' Test si le noeud est un fils
				ElseIf Instr(1, arrListDG(i), sParent , 1) = 1 Then
					sTiret = ""
					For h = 1 to Len(sParent)
						sTiret = sTiret + " "
					Next
					s = " | " & Replace(arrListDG(i),sParent,sTiret)
				Else
					s = arrListDG(i)
					pereOK = False
				End If
			End If
		SetOption s , arrListDG(i)
		End If
	Next
	
	DGexist.style.padding = "5px"
	'Alert(bMFound)
	'Alert(Manufact)
	
	'hDriverSelect.Value = sGeneric
	hDriverSelect.Value = DriverPath.Value
	'hDriverSelect.Disabled = True
	
	If bFound Then
			' Sélection modèle si trouvé
			DGModel.checked = True
			DGexist.style.background = "#00A651"
			Alert("Model Match ! Pouvoir de finition = 90% ")
	ElseIf bMFound Then
			' Sinon Sélection Constructeur si trouvé
				'DGMake.checked = True
				DGList.Checked = True
				DGMexist.style.background = "#FFFF01"
			Alert ("Constructeur only, pouvoir de finiton 50% " & vbcrlf & _
							"Votre seule chance repose sur le parametrage Liste Drivers !")
		
	Else
				' Sinon Sélection toute la base MDT
				'DGAll.checked = True
				'DGnoexist.style.background = "#FF4500"
				DGnone.checked = True
				'DGList.Checked = True
				DGNoDrivers.style.background = "#FF4500"
	End If 
	
	If Not bFound Then	
		Alert("Choix Manuel ! Veuillez choisir un equivalent dans Liste Drivers ! " & vbcrlf & _ 
				"Si serveur Fujitsu, choisir votre modele dans FUJITSU\ServerView" & vbcrlf & _
				"Si serveur IBM , choisir votre modele IBM\ServerGuide" _
				)
	End If
	
End Function


Function ValidateDriverGroup		
	If DGList.checked Then
		hDriverSelect.Disabled = False
	Else
		hDriverSelect.Disabled = True
	End If
	ValidateDriverGroup = True
End Function

Function ValidateDriverGroup_Final
	If DGMake.checked Then
		' Si constructeur
		oEnvironment.Item("DriverGroup001")=Manufact
		oEnvironment.Item("DriverSelectionProfile")="Nothing"
'	ElseIf DGAll.checked Then 
		' Si toute la base MDT
'		oEnvironment.Item("DriverGroup001")="default"
'		oEnvironment.Item("DriverSelectionProfile")="All Drivers"
	ElseIf DGNone.checked Then
		' Si aucun
		oEnvironment.Item("DriverGroup001")=""
		oEnvironment.Item("DriverSelectionProfile")="Nothing"
	ElseIf DGList.checked Then
		' Si Liste prédéfinis
		oEnvironment.Item("DriverGroup001")=Mid(hDriverSelect.Value,1,Instr(1,hDriverSelect.Value," <",1)-1)
		DriverPath.Value = Property("DriverGroup001")
		oEnvironment.Item("DriverSelectionProfile")="Nothing"
	ElseIf DGModel.checked Then
		' Si Par défaut, constructeur/modèle
		oEnvironment.Item("DriverGroup001")=DriverPath.Value
		oEnvironment.Item("DriverSelectionProfile")="Nothing"
	End If
	
	' Integration des pilotes constructeurs pour n'importe quel cas
	oEnvironment.Item("DriverGroup002")="Constructeurs"
	' Flush the value to variables.dat, before we continue.
	SaveAllDataElements
	SaveProperties
	
	ValidateDriverGroup_Final = true
End Function

' Fonction d'ajout des éléments SELECT OPTION
Sub SetOption(OptText,OptValue) 
	Dim oNewOption
	Set oNewOption = Document.CreateElement("OPTION")
 	oNewOption.Text = OptText 
	oNewOption.Value = OptValue 
	oSelect.options.Add(oNewOption) 
End Sub

