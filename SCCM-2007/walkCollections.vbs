'
' CollectionWalker
' v0.1
' Emmanuel Tsouris
' emmanuel.tsouris@lmco.com
'
'Example:
' cscript CollectionWalker.vbs
'
'

'On Error Resume Next 

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
'WScript.Echo "<?xml version=""1.0"" encoding=""ISO-8859-1""?>"

'Begin copying rights to the subcollections
Call WalkCollection(objSWbemServices, collectionID, 1, "")

'Done
WScript.Quit(0)


'Subroutines to enumerate collections and set rights.

'Copy Rights to SubCollections
Sub WalkCollection(connection, targetParentCollectionID, level, topPath)
	Set subCollections = connection.ExecQuery("select subCollectionID from SMS_CollectToSubCollect where ParentCollectionID = '" & targetParentCollectionID & "'")

	'Loop each sub-collection
	For each subCollection in subCollections

        Dim collectionName
        Dim path
        path = topPath

		path = path & "\" & subCollection.SubCollectionID

		collectionName =  getCollectionName(connection, subCollection.SubCollectionID)
		WScript.Echo path & ":" & collectionName

		Call WalkCollection(connection, subCollection.SubCollectionID, level+1, path)

	Next
End Sub


Function getCollectionName(connection, targetCollectionId)

	Set collections = connection.ExecQuery("select * from SMS_Collection where CollectionID = '" & targetCollectionId & "'")

	For each collection in collections

		getCollectionName =  collection.name
	Next
End Function

Sub CatchError
	If Err.Number <> 0 Then
		WScript.Echo Now() & vbTab & "Error: " & Err.Number
		WScript.Echo Now() & vbTab &  "Error (Hex): " & Hex(Err.Number)
		WScript.Echo Now() & vbTab &  "Source: " &  Err.Source
		WScript.Echo Now() & vbTab &  "Description: " &  Err.Description
		Err.Clear
	End If
End sub

