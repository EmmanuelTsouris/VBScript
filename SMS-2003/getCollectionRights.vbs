'
' Get Collection Rights
' v0.3 2/27/2008
' Emmanuel Tsouris
'
' server: SMS Primary Site Server
' site: SMS Site Code
' collection: Target Collection to Read
'
' example: getCollectionRights.vbs /server:serverName /site:ABC /col:ABC01234
'
On Error Resume Next 

'Get the named arguments
Set colNamedArguments = WScript.Arguments.Named

'SMS Primary Site Server Name
smsServerName = colNamedArguments.Item("server")

'SMS Primary Site Code
smsSiteCode = colNamedArguments.Item("site")

'Parent Collection ID
collectionID = colNamedArguments.Item("col")

'Connect to the server
Set objLocator = CreateObject("WbemScripting.SWbemLocator")
CatchError()

Set objSWbemServices = objLocator.ConnectServer(smsServerName , "root\sms\site_" & smsSiteCode )
CatchError()

WScript.Echo Now() & vbTab & smsServerName  & ", " & smsSiteCode & ", " & collectionID

Set targetCollectionRights = objSWbemServices.ExecQuery( "Select * From SMS_UserInstancePermissions " & _
	" WHERE ObjectKey=1 AND InstanceKey='" & collectionID & "'" )

For Each colRight in targetCollectionRights
	WScript.Echo collectionID & _
		vbTab & colRight.Username & _
		vbTab & colRight.InstancePermissions
	Next


WScript.Quit(0)

Sub CatchError
	If Err.Number <> 0 Then
		WScript.Echo Now() & vbTab & "Error: " & Err.Number
		WScript.Echo Now() & vbTab &  "Error (Hex): " & Hex(Err.Number)
		WScript.Echo Now() & vbTab &  "Source: " &  Err.Source
		WScript.Echo Now() & vbTab &  "Description: " &  Err.Description
		Err.Clear
	End If
End sub

' Reference & Notes
' InstancePermissions are a bitwise string value.
'
' Support Rights
' "4129" Read, Read Resource, Use Remote Tools
'
' All Rights but not delete
' "6883" Advertise, Delete Resource, Modify, Modify Resource, Read, Read Resource, Use Remote Tools, View Collected Files
'
' All Rights
' "6887" Advertise, Delete, Delete Resource, Modify, Modify Resource, Read, Read Resource, Use Remote Tools, View Collected Files
'
' RMD
' "7" Read, Modify, Delete
'
' Read
' "1" Read
'
' RMDR
' "4103" Read, Modify, Delete, Read Resource
'
'
'
'
'