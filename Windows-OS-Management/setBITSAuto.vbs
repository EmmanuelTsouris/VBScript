
'
' setBITSAuto
' Connects to WMI on a computer and sets the BITS service to auto
'

On Error Resume Next

' Computer Name
computerName = "computer.somedomain.com"

' Create a WMI Object
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & computerName & "\root\cimv2")

' Check for no Errors
If Err = 0 Then

	' Query WMI for the BITS service
	Set colServiceList = objWMIService.ExecQuery("Select * from Win32_Service where Name = 'BITS'")

	If Err = 0 Then	

		WScript.Echo computerName & vbCrLf
		
		' Loop through the BITS services (should only be one, but who knows what might happen in the future)
		For Each objService in colServiceList
		
			' Change the service to automatic and grab any error codes that might be returned		
			errReturnCode = objService.Change( , , , , "Automatic")
			
			' Echo out the change
			WScript.Echo vbTab & "Changed " & objService.Name & " to automatic"
			
			' Start the service
			objService.StartService()
		Next
	End If

	Set colServiceList = nothing

	
End If

Set objWMIService = nothing



