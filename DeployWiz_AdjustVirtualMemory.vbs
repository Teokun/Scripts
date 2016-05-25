'
'
'
' Panneau de paramètre pour définir VMInitialSize et VMMaximumSize
'
'
'

Function InitializeAdjustVirtualMemory
	
	' the Adjust Application has the GUID : {17a5b435-eec3-44bf-af99-29e0f81b9e7f}

	' Size in MB
	VMInitialSize.Value = 512
	VMMamixumSize.Value = 10240

	If IsAdjustAppPresent=False Then ButtonNextClick

End Function

Function IsAdjustAppPresent
	Dim oAppItem
	IsAdjustAppPresent=False
	
	If Not IsEmpty(Property("Applications")) Then
		For Each oAppItem in Property("Applications")
			If UCase(oAppItem) = UCase("{17a5b435-eec3-44bf-af99-29e0f81b9e7f}") Then
				IsAdjustAppPresent=True
				Exit for
			End If
		Next
	End If
End Function

Function ValidateVMSize

	ValidateVMSize = True
	
	ErrSizeMin.style.display = "none"
	ErrSizeMax.style.display = "none"
	
	If (VMInitialSize.Value <> "" ) and ( VMMaximumSize.Value <> "" ) Then
		If IsNumeric(VMMamixumSize.Value) and IsNumeric(VMInitialSize.Value) Then
			If VMInitialSize.Value < 512 Then 
				ErrSizeMin.style.display = "inline"
				ValidateVMSize = False
			End If
		
			If ( VMMaximumSize.Value / VMInitialSize.Value ) < 1 Then
				ErrSizeMax.style.display = "inline"
				ValidateVMSize = False
			End If
		End If
	End If
	
End Function