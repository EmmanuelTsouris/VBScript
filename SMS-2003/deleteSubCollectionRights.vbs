'
' Set Subcollection Rights
' v1.0 7/16/2008
' Emmanuel Tsouris
'
'Example:  (be sure to use quotes around the group name)
' cscript setSubCollectionRights.vbs /server:serverName /site:ABC /col:ABC01234 /user:"domain\UserNameOrGroup"
'
'Does not delete permissions on the collectionID specified, only it's sub-collections.
'
'Spelling must be exact!
' SMS will allow us to put any group name in there and doesn't check for validity.
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

'User or GroupName
collectionUserName = colNamedArguments.Item("user")

'Connect to the server
Set objLocator = CreateObject("WbemScripting.SWbemLocator")
CatchError()

Set objSWbemServices = objLocator.ConnectServer(smsServerName , "root\sms\site_" & smsSiteCode )
CatchError()

'Begin copying rights to the subcollections
CopyRightsToSubCollection collectionID, collectionUserName

'Done
WScript.Quit(0)


'Subroutines to enumerate collections and set rights.

'Copy Rights to SubCollections
Sub CopyRightsToSubCollection(targetParentCollectionID, targetUserName)
	Set subCollections = objSWbemServices.ExecQuery("select * from SMS_CollectToSubCollect where ParentCollectionID = '" & targetParentCollectionID & "' order by Name")

	'Loop each sub-collection
	For each subCollection in subCollections
		
		'Set rights on the sub-collection
		setCollectionRight subCollection.SubCollectionID,targetUserName
		
		'loop through each of it's sub-collections
		CopyRightsToSubCollection subCollection.SubCollectionID,targetUserName
	Next
End Sub

'Set the Collection Rights
Sub setCollectionRight(targetCollectionID, targetGroupName)

	Set objNewRight = objSWbemServices.Get("SMS_UserInstancePermissions").SpawnInstance_()
	objNewRight.UserName = targetGroupName 
		
	' for complete list of .ObjectKey & .InstancePermissions Values
	' reference the SMS 2003 SDK documentation.
		
	'1 = collections
	objNewRight.ObjectKey = 1
	objNewRight.InstanceKey = targetCollectionID
		
	' bit field
	objNewRight.InstancePermissions = instancePerms
	objNewRight.Put_
	
	WScript.Echo Now() & vbTab & "set" & vbTab & objNewRight.InstanceKey & vbTab & objNewRight.Username & vbTab & objNewRight.InstancePermissions

End Sub

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
' All Rights but not delete
' "6883" Advertise, Delete Resource, Modify, Modify Resource, Read, Read Resource, Use Remote Tools, View Collected Files
'
' All Rights but not delete resource
' "6375" Advertise, Delete, Modify, Modify Resource, Read, Read Resource, Use Remote Tools, View Collected Files
'
' All Rights but not delete and delete resource
' "6371" Advertise, Modify, Modify Resource, Read, Read Resource, Use Remote Tools, View Collected Files
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
