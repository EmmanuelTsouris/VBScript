'
' Copy Collection Rights to Sub-Collection
' v0.3 2/28/2008
' Emmanuel Tsouris

'SMS Primary Site Server Name
smsServerName = "serverName"

'SMS Primary Site Code
smsSiteCode = "ABC"

'Parent Collection ID
collectionID = "ABC0001D"  'ID of the collection

Set objLocator = CreateObject("WbemScripting.SWbemLocator")
Set objSWbemServices = objLocator.ConnectServer(smsServerName , "root\sms\site_" & smsSiteCode )

Set parentCollectionRights = objSWbemServices.ExecQuery( "Select * From SMS_UserInstancePermissions WHERE ObjectKey=1 AND InstanceKey='" & targetParentCollectionID & "'" )


CopyRightsToSubCollection collectionID, parentCollectionRights

WScript.Quit(0)

'Subroutines to enumerate collections and set rights.

Sub CopyRightsToSubCollection(targetParentCollectionID, targetCollectionRights)
	Set subCollections = objSWbemServices.ExecQuery("select * from SMS_CollectToSubCollect where ParentCollectionID = '" & targetParentCollectionID & "' order by Name")

	For each subCollection in subCollections
		For Each colRight in targetCollectionRights
			setCollectionRight subCollection.SubCollectionID,colRight.Username,colRight.InstancePermissions
		Next
		
		CopyRightsToSubCollection subCollection.SubCollectionID,targetCollectionRights
	Next
End Sub

Sub setCollectionRight(targetCollectionID, targetGroupName, targetPermissions)

	Set objNewRight = objSWbemServices.Get("SMS_UserInstancePermissions").SpawnInstance_()
	objNewRight.UserName = targetGroupName 
		
	' for complete list of .ObjectKey & .InstancePermissions Values
	' reference the SMS 2003 SDK documentation.
		
	'1 = collections
	objNewRight.ObjectKey = 1
	objNewRight.InstanceKey = targetCollectionID
		
	' bit field
	objNewRight.InstancePermissions = targetPermissions
	objNewRight.Put_
	
	WScript.Echo Now() & vbTab & "set" & vbTab & objNewRight.InstanceKey & vbTab & objNewRight.Username & vbTab & objNewRight.InstancePermissions

End Sub

' Reference & Notes
' InstancePermissions are a bitwise string value.
'
' All Rights but not delete
' "6883" Advertise, Delete Resource, Modify, Modify Resource, Read, Read Resource, Use Remote Tools, View Collected Files
'
' All Rights
' "6887" Advertise, Delete, Delete Resource, Modify, Modify Resource, Read, Read Resource, Use Remote Tools, View Collected Files
