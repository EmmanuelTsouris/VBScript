'
' CreateCollections
' v0.1
' Emmanuel Tsouris
'
'Example:  (be sure to use quotes around the group name)
'
'On Error Resume Next 

'Get the named arguments
Set colNamedArguments = WScript.Arguments.Named

'Primary Site Server Name
smsServerName = colNamedArguments.Item("server")

'Primary Site Code
smsSiteCode =  colNamedArguments.Item("site")

'Parent Collection ID
collectionID = colNamedArguments.Item("col")

'Connect to the server
Set swbemLocator = CreateObject("WbemScripting.SWbemLocator")
CatchError()

Set swbemconnection = swbemLocator.ConnectServer(smsServerName , "root\sms\site_" & smsSiteCode )
CatchError()

WScript.Echo Now() & vbTab & "set" & vbTab & "Preparing to Create the Collection"

'This is where the magic happens
Call CreateDynamicCollection(swbemconnection, collectionID, collectionID, "My Test Collection 2", "New dynamic collection comment.", true, "SELECT * from SMS_R_System", "New Rule Name")



'Set the Collection Rights
Sub setCollectionRight(connection, targetCollectionID, targetGroupName, targetInstancePerms)

	Set objNewRight = connection.Get("SMS_UserInstancePermissions").SpawnInstance_()
	CatchError()
        
        objNewRight.UserName = targetGroupName 
	objNewRight.ObjectKey = 1 '1 = collections
	objNewRight.InstanceKey = targetCollectionID
		
	' bit field
	objNewRight.InstancePermissions = targetInstancePerms
	objNewRight.Put_
	CatchError()
	WScript.Echo Now() & vbTab & "set" & vbTab & objNewRight.InstanceKey & vbTab & objNewRight.Username & vbTab & objNewRight.InstancePermissions

End Sub

Sub CreateDynamicCollection(connection, existingParentCollectionID, limitToCollectionID, newCollectionName, newCollectionComment, ownedByThisSite, queryForRule, ruleName)

    ' Create the collection.
    Set newCollection = connection.Get("SMS_Collection").SpawnInstance_
    CatchError()
    newCollection.Comment = newCollectionComment
    newCollection.Name = newCollectionName
    newCollection.OwnedByThisSite = ownedByThisSite
    
    ' Save the new collection and save the collection path for later.
    Set collectionPath = newCollection.Put_
    CatchError()
    
    WScript.Echo Now() & vbTab & "set" & vbTab & "Created the Collection"
    
    newCollectionID = CStr(collectionPath.Keys("CollectionID"))
    
   ' Define to what collection the new collection is subordinate.
   ' IMPORTANT: If you do not specify the relationship, the new collection will not be visible in the console. 
    Set newSubCollectToSubCollect = connection.Get("SMS_CollectToSubCollect").SpawnInstance_
    CatchError()
    newSubCollectToSubCollect.parentCollectionID = existingParentCollectionID
    newSubCollectToSubCollect.subCollectionID = newCollectionID
    
    ' Save the subcollection information.
    newSubCollectToSubCollect.Put_
    CatchError()
    
    WScript.Echo Now() & vbTab & "set" & vbTab & "Set as subcollection"

    ' Create a new collection rule object for validation.
    Set queryRule = connection.Get("SMS_CollectionRuleQuery")
    CatchError()
    
    ' Validate the query (good practice before adding it to the collection). 
    validQuery = queryRule.ValidateQuery(queryForRule)
    CatchError()
    
    ' Continue with processing, if the query is valid.
    If validQuery Then
        
        ' Create the query rule.
        Set newQueryRule = QueryRule.SpawnInstance_
        newQueryRule.QueryExpression = queryForRule
        newQueryRule.RuleName = ruleName
        newQueryRule.LimitToCollectionID = limitToCollectionID
        
        ' Add the new query rule to a variable.
        Set newCollectionRule = newQueryRule
        
        ' Get the collection.
        Set newCollection = connection.Get(collectionPath.RelPath)
        
        ' Add the rules to the collection.
        newCollection.AddMembershipRule newCollectionRule

        ' Call RequestRefresh to initiate the collection evaluator.
        newCollection.RequestRefresh False
        WScript.Echo Now() & vbTab & "set" & vbTab & "Create query rule"
    End If
    
    Call setCollectionRight(connection, newCollectionID, "domain\username", 1)
    Call setCollectionRight(connection, newCollectionID, "domain\username", 1)
    
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
