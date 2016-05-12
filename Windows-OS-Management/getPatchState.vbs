
'
' getPatchState
'

On Error Resume Next
Declare Count

Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_Patchstate_Extended",,48)

Count = 0

WScript.Echo "Found " & colItems.count & " item(s)."

For Each objItem in colItems

	If objItem.Status <> "Installed" Then
		WScript.Echo "AuthorizationName: " & objItem.AuthorizationName
		WScript.Echo "Title: " & objItem.Title
		WScript.Echo "QNumbers: " & objItem.QNumbers
		WScript.Echo "RebootType: " & objItem.RebootType
		WScript.Echo "ScanAgent: " & objItem.ScanAgent
		WScript.Echo "ScanDateTime: " & objItem.ScanDateTime
		WScript.Echo "Status: " & objItem.Status
		Count = Count + 1
	End If
Next 

If Count = 0 Then
	WScript.Echo "You are patched."
ElseIf Count = 1 Then
	WScript.Echo "You are missing one patch."
Else
	WScript.Echo "You are missing " & Count & " patches."
End If


