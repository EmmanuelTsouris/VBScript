'
' getSMSCollectionStructure
' v0.1 9/8/2011
' Emmanuel Tsouris
'
' Walks a Collection Hierarchy in SMS 2003 and formats the results for display
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



Sub WalkCollection(connection, targetParentCollectionID, level, path)
	Set subCollections = connection.ExecQuery("select subCollectionID from SMS_CollectToSubCollect where ParentCollectionID = '" & targetParentCollectionID & "'")

	'Loop each sub-collection
	For each subCollection in subCollections

		path = path & vbTab & subCollection.SubCollectionID

		collectionName =  getCollectionName(connection, subCollection.SubCollectionID)
		WScript.Echo path & " " & collectionName & ", "

		Call WalkCollection(connection, subCollection.SubCollectionID, level+1, path)

	Next
End Sub


Function getCollectionName(connection, targetCollectionId)
	Set collections = connection.ExecQuery("select * from SMS_Collection where CollectionID = '" & targetCollectionId & "'")

	For each collection in collections
		getCollectionName =  collection.name
	Next
	
End Function

sub showCollectionRefresh(connection, targetCollectionId)
	Set collections = connection.ExecQuery("select * from SMS_Collection where CollectionID = '" & targetCollectionId & "'")

	For each collection in collections
		getCollectionName =  collection.name

		For Each propValue In instCollection.RefreshSchedule

		WScript.Echo "$$$$$-" & VarType(propValue)
		WScript.Echo vbTab & "RefreshSchedule: " & propValue

		If VarType(propValue) = 0 Then Exit For

		set schedule=instCollection.RefreshSchedule(0)
		schedtype=Schedule.Path_.Class
		WScript.echo vbTab & schedtype
		WScript.echo vbTab & "Start Time: " & schedule.StartTime

		Select Case schedtype

		Case "SMS_ST_RecurMonthlyByDate"
			WScript.Echo instCollection.CollectionID & vbTab & "MONTHDAY" & vbTab & schedule.MonthDay & vbTab & instCollection.Name

		Case "SMS_ST_RecurInterval"
			WScript.Echo vbTab & "DaySpan: " & schedule.DaySpan

			WScript.Echo vbTab & "HourSpan: " & schedule.HourSpan
			WScript.Echo vbTab & "MinuteSpan: " & schedule.MinuteSpan

	Select Case True

	Case schedule.DaySpan <> 0
		Wscript.Echo instCollection.CollectionID & vbTab & "DAY" & vbTab & schedule.DaySpan & vbTab & instCollection.Name
	Case schedule.HourSpan <> 0
		Wscript.Echo instCollection.CollectionID & vbTab & "HOUR" & vbTab & schedule.HourSpan & vbTab & instCollection.Name
	Case schedule.MinuteSpan <> 0
		WScript.Echo instCollection.CollectionID & vbTab & "MINUTE" & vbTab & schedule.MinuteSpan & vbTab & instCollection.Name & vbTab & schedule.StartTime

	'***MAKE THE CHANGE TO THE SCHEDULE***
	'schedule.MinuteSpan = 0
	'schedule.DaySpan = 1
	'instCollection.Put_
	'*************************************
	Case Else
		WScript.Echo instCollection.CollectionID & vbTab & "NO SMS_ST_RecurInterval found" � � � � � � � � � � � � � � � � � � �End Select ��
	Case Else
		WScript.Echo instCollection.CollectionID & vbTab & "NO schedtype found" ��
	End Select

	If schedule.DaySpan <> 0 Then Wscript.Echo objItem.CollectionID & vbTab & objItem.Name & vbTab & "Day frequency: " & schedule.DaySpan �
	If schedule.HourSpan <> 0 Then Wscript.Echo objItem.CollectionID & vbTab & objItem.Name & vbTab & "Hour frequency: " & schedule.HourSpan
	If schedule.MinuteSpan <> 0 Then WScript.Echo objItem.CollectionID & vbTab & objItem.Name & vbTab & "Minute frequency: " & schedule.MinuteSpan
	If schedtype="SMS_ST_RecurMonthlyByDate" Then
	If schedule.MonthDay <> 0 Then WScript.Echo objItem.CollectionID & vbTab & objItem.Name & vbTab & "Day of the month: " & schedule.MonthDay
End If
End If
Next
On Error goto 0
Else � � � � � �'Collection doesn't update on a schedule
WScript.Echo instCollection.CollectionID & vbTab & "NOSCHED" & vbTab & vbTab & instCollection.Name
'strNoUpdateSched = strNoUpdateSched & objItem.CollectionID & vbTab & objItem.Name & vbcrlf
End If ��

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

