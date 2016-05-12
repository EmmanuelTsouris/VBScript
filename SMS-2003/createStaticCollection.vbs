' Set up a connection to the local provider.
Set swbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set swbemconnection= swbemLocator.ConnectServer(".", "root\sms")
Set providerLoc = swbemconnection.InstancesOf("SMS_ProviderLocation")

For Each Location In providerLoc
    If location.ProviderForLocalSite = True Then
        Set swbemconnection = swbemLocator.ConnectServer(Location.Machine, "root\sms\site_" + Location.SiteCode)
        Exit For
    End If
Next

Call CreateStaticCollection(swbemconnection, "SMS00001", "New Static Collection Name", "New static collection comment.", true, "SMS_R_System", 2)


Sub CreateStaticCollection(connection, existingParentCollectionID, newCollectionName, newCollectionComment, ownedByThisSite, resourceClassName, resourceID)

    ' Create the collection.
    Set newCollection = connection.Get("SMS_Collection").SpawnInstance_
    newCollection.Comment = newCollectionComment
    newCollection.Name = newCollectionName
    newCollection.OwnedByThisSite = ownedByThisSite
    
    ' Save the new collection and save the collection path for later.
    Set collectionPath = newCollection.Put_    
    
   ' Define to what collection the new collection is subordinate.
   ' IMPORTANT: If you do not specify the relationship, the new collection will not be visible in the console. 
    Set newSubCollectToSubCollect = connection.Get("SMS_CollectToSubCollect").SpawnInstance_
    newSubCollectToSubCollect.parentCollectionID = existingParentCollectionID
    newSubCollectToSubCollect.subCollectionID = CStr(collectionPath.Keys("CollectionID"))
    
    ' Save the subcollection information.
    newSubCollectToSubCollect.Put_
        
    ' Create the direct rule.
    Set newDirectRule = connection.Get("SMS_CollectionRuleDirect").SpawnInstance_
    newDirectRule.ResourceClassName = resourceClassName
    newDirectRule.ResourceID = resourceID
    
    ' Add the new query rule to a variable.
    Set newCollectionRule = newDirectRule
    
    ' Get the collection.
    Set newCollection = connection.Get(collectionPath.RelPath)
    
    ' Add the rules to the collection.
    newCollection.AddMembershipRule newCollectionRule

    ' Call RequestRefresh to initiate the collection evaluator. 
    newCollection.RequestRefresh False
    
End Sub
