'
' Get Collection
' v0.1 4/11/2008
' Emmanuel Tsouris
'
' server: SMS Primary Site Server
' site: SMS Site Code
'
' This script offers two modes, search by name (wildcards ok) or search by collection ID.
' Name: Collection Name to Search For
' ID: Collection ID to Search For
'
' example: getCollection.vbs /name:"%some collection%"
' example: getCollection.vbs /id:ABC00123
'
'On Error Resume Next 

'Get the named arguments
Set colNamedArguments = WScript.Arguments.Named

'SMS Primary Site Server Name
smsServerName = colNamedArguments.Item("server")

'SMS Primary Site Code
smsSiteCode = colNamedArguments.Item("site")

'Collection Name
collectionName = colNamedArguments.Item("name")

'Collection Id
collectionID = colNamedArguments.Item("id")

If (colNamedArguments.Item("id") ="" and colNamedArguments.Item("name") ="") then
    WScript.Echo("Searches for and displays collection information based on Collection Name or Collection ID.")
    WScript.Echo(vbTab & "example: cscript getCollection.vbs /name:""%searchtext%""")
    WScript.Echo(vbTab & "example: cscript getCollection.vbs /id:ABC00123")
    WScript.Quit(0)
End If

Set oLocator = CreateObject("WbemScripting.SWbemLocator")
CatchError()

Set oSMSConnection = oLocator.ConnectServer(smsServerName , "root\sms\site_" & smsSiteCode )
CatchError()

If colNamedArguments.Item("id") <> "" Then
    Set oCollections = oSMSConnection.ExecQuery ("Select * from SMS_Collection where CollectionID = '" & collectionID & "' Order By Name")
    CatchError()

	    'Walk through the collections
    For Each oCollection in oCollections
	    WScript.Echo oCollection.CollectionID & _
		    vbTab & oCollection.Name
		WScript.Echo "Child of: "
    	subCollection oCollection.CollectionID              		
    Next
ElseIf colNamedArguments.Item("name") <> "" Then
    Set oCollections = oSMSConnection.ExecQuery ("Select * from SMS_Collection where Name like '" & collectionName & "' Order By Name")
    CatchError()

	    'Walk through the collections
    For Each oCollection in oCollections
	    WScript.Echo oCollection.CollectionID & _
		    vbTab & oCollection.Name

    Next
End If

WScript.Quit(0)

Sub subCollection(collectionID)
	Set subCollections = oSMSConnection.ExecQuery("select * from SMS_CollectToSubCollect where SubCollectionID = '" & collectionID & "'")
    CatchError()
    For Each oSubCollection in subCollections
	    WScript.Echo oSubCollection.parentCollectionID
		subCollection oSubCollection.parentCollectionID
    Next
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
'
'Class SMS_Collection : SMS_BaseClass
'{
'  string CollectionID;
'  SMS_CollectionRule CollectionRules[];
'  string Comment;
'  uint32 CurrentStatus;
'  datetime LastChangeTime;
'  datetime LastRefreshTime;
'  string MemberClassName;
'  string Name;
'  Boolean OwnedByThisSite;
'  SMS_ScheduleToken RefreshSchedule[];
'  uint32 RefreshType;
'  boolean ReplicateToSubSites;
'};
'