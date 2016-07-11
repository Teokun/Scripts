'
' DeployWiz_FusionInventory by JBCO 2016/06/24
'
'
'
'
Option Explicit
Dim dClient
Dim dClientSite
Dim oUniqCli, oUniqCliSite

Function InitializeFusionInventory

	Dim TagFile : TagFile = oEnvironment.Item("DeployRoot") & "\Control\TagFusion.csv"
	Dim i, inputFile
	Dim oCli, oCliSite
	Dim row : row = 0
	Dim fields
	
	Set dClient = CreateObject("Scripting.Dictionary")
	Set dClientSite = CreateObject("Scripting.Dictionary")	
	
	' Check TagFile and Build selection menu
	If oFso.FileExists( TagFile ) Then
		Const ForReading = 1    ' Declare constant for reading for more clarity
	
		Set inputFile = oFso.OpenTextFile(TagFile, ForReading, True) ' Set inputFile as file to be read from
		inputFile.ReadLine 'skip header 
	
		Do Until inputFile.AtEndOfStream  
			fields = Split(inputFile.Readline,";") 'store line in temp array  
			' fields(0) = Client 
			' fields(1) = ClientSite 
			' fields(2) = TagSite 
			' fields(3) = TagFusion
		
			On Error Resume Next
			dClient.add fields(0), fields(3)
			On Error Resume Next
			dClientSite.add fields(3) & " " & fields(1) , fields(2)
		
		Loop
		inputFile.close
	Else
		ValidateFusionInventory = True
		Exit Function
	End If

'wscript.echo Join(dClient.Keys)
'wscript.echo Join(dClient.Items)
'wscript.echo Join(dClientSite.Keys)
'wscript.echo Join(dClientSite.Items)

	'Populate client LisT
	Set oUniqCli=document.getElementById("hUniqCli")
	For Each oCli in oUniqCli.Options 
		oCli.RemoveNode
	Next
	
	Set oUniqCliSite=document.getElementById("hUniqCliSite")
	For Each oCliSite in oUniqCliSite.Options 
		oCliSite.RemoveNode
	Next
	
	For Each i In dClient
		SetCliOption i , dClient(i)
		'Alert( i & dClient(i) )
	Next

End Function

Function GetSiteFusionTag
	Dim oCliSite
	Dim i
	Dim selectCli
	
	For Each oCliSite in oUniqCliSite.Options 
		oCliSite.RemoveNode
	Next
	
	selectCli = hUniqCli.Value
	
	For Each i In dClientSite
		If InStr(1, i, selectCli, 1) = 1 Then 
			SetCliSiteOption Right( i , Len(i) - Len(selectCli) -1 )  , dClientSite(i)
		End If
	Next
	
End Function

Function ValidateFusionInventory

	FusionTag.Value = hUniqCli.Value
	FusionTagSite.Value = hUniqCliSite.Value
	
	If FusionTagSite.Value <> "" Then 
		oEnvironment.Item("OrgName") = FusionTag.Value & "_" & FusionTagSite.Value
	Else
		oEnvironment.Item("OrgName") = FusionTag.Value
	End If
	Alert("TAG_FUSION : TAG_" & oEnvironment.Item("OrgName") )
	
	oEnvironment.Item("MandatoryApplications001")="{0c58c4d9-4aea-41b5-b67a-00ab5179f9e3}"
	ValidateFusionInventory = True

End Function

Sub SetCliOption(OptText,OptValue) 
	Dim oNewOption
	Set oNewOption = Document.CreateElement("OPTION")
 	oNewOption.Text = OptText 
	oNewOption.Value = OptValue 
	oUniqCli.options.Add(oNewOption) 
End Sub

Sub SetCliSiteOption(OptText,OptValue) 
	Dim oNewOption
	Set oNewOption = Document.CreateElement("OPTION")
 	oNewOption.Text = OptText 
	oNewOption.Value = OptValue 
	oUniqCliSite.options.Add(oNewOption) 
End Sub