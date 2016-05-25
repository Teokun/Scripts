	
	
Function InitializeTestView
	
	Dim sProp, objTmp, objT
	Dim admOK, pwdOK
	
	admOK=True
	pwdOK=False
	
	 If oEnvironment.ListItem("Administrators") is Nothing then
   admOK = False
 ElseIf oEnvironment.ListItem("Administrators").count < 1 Then
   admOK = False
 End if
 
 If oEnvironment.ListItem("AdminPassword") is Nothing then
   pwdOK = False
 ElseIf oEnvironment.ListItem("AdminPassword").count < 1 Then
   pwdOK = False
 End if
  
 
If admOK Then
	For Each sProp in oEnvironment.ListItem("Administrators")
  objTmp = objTmp & vbCrLf & sProp
	Next
End If
	
If pwdOK Then
	objT = oEnvironment.ListItem("AdminPassword")
	objTmp = objTmp & vbCrLf & objT
	Next
End If
	optionalWindow1.InnerText = objTmp
	
End Function
	
Function ValidateTestView

InitializeTestView

End Function