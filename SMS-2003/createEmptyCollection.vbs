' Set up a connection to the local provider.
Set swbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set swbemconnection= swbemLocator.ConnectServer("serverName" , "root\sms\site_ABC")
Set providerLoc = swbemconnection.InstancesOf("SMS_ProviderLocation")

Call CreateCollection(swbemconnection, "ABC00123", "Some Collection Name", "Comment Goes Here")


Sub CreateCollection(connection, existingParentCollectionID, newCollectionName, newCollectionComment)

    ' Create the collection.
    Set newCollection = connection.Get("SMS_Collection").SpawnInstance_
    newCollection.Comment = newCollectionComment
    newCollection.Name = newCollectionName
    newCollection.OwnedByThisSite = true
    
    ' Save the new collection and save the collection path for later.
    Set collectionPath = newCollection.Put_    
    
   ' Define to what collection the new collection is subordinate.
   ' IMPORTANT: If you do not specify the relationship, the new collection will not be visible in the console. 
    Set newSubCollectToSubCollect = connection.Get("SMS_CollectToSubCollect").SpawnInstance_
    newSubCollectToSubCollect.parentCollectionID = existingParentCollectionID
    newSubCollectToSubCollect.subCollectionID = CStr(collectionPath.Keys("CollectionID"))
    
    ' Save the subcollection information.
    newSubCollectToSubCollect.Put_

    
End Sub
