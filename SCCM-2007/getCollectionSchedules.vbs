'
' getCollectionSchedules
' v0.1
' Emmanuel Tsouris
'
'Example:
' cscript getCollectionSchedules.vbs /c:COL00001
'   where COL00001 is the collection ID
'

On Error Resume Next 

'Get the named arguments
Set colNamedArguments = WScript.Arguments.Named

'SMS Primary Site Server Name
smsServerName = colNamedArguments.Item("server")

'SMS Primary Site Code
smsSiteCode = colNamedArguments.Item("site")

'Parent Collection ID
collectionID = colNamedArguments.Item("c")

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
Sub WalkCollection(connection, targetParentCollectionID, level, path)
	Set subCollections = connection.ExecQuery("select subCollectionID from SMS_CollectToSubCollect where ParentCollectionID = '" & targetParentCollectionID & "'")
    CatchError()

	'Loop each sub-collection
	For each subCollection in subCollections
    CatchError()

        Set collections = connection.ExecQuery("select * from SMS_Collection where CollectionID = '" & subCollection.SubCollectionID & "'")
        CatchError()

	        For each collection in collections
            CatchError()
                
                WScript.Echo tabs

                path = path & "\" & subCollection.SubCollectionID
                WScript.Echo path

                Call SetSchedule(connection, collection.collectionID) 
                CatchError()

            Next
        
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

Function WMIDateStringToDate(dtmInstallDate)
    WMIDateStringToDate = CDate(Mid(dtmInstallDate, 5, 2) & "/" & _
    Mid(dtmInstallDate, 7, 2) & "/" & Left(dtmInstallDate, 4) _
    & " " & Mid (dtmInstallDate, 9, 2) & ":" & _
    Mid(dtmInstallDate, 11, 2) & ":" & Mid(dtmInstallDate, _
    13, 2))
End Function

Sub SetSchedule(connection, collectionID) 

    Set Collection = connection.Get("SMS_Collection.CollectionID='" & collectionid & "'")
    CatchError()

    set schedule=Collection.RefreshSchedule(0)
    CatchError()

    If Collection.RefreshType = 1 Then

        WScript.Echo Collection.Name & " Type:" & Collection.RefreshType & " LastRefresh:" & WMIDateStringToDate(Collection.LastRefreshTime)
    
        schedtype=Schedule.Path_.Class

        WScript.echo schedtype 
        WScript.echo "Start Time: " & WMIDateStringToDate(schedule.StartTime)
    
        if schedtype="SMS_ST_RecurInterval" then 
            if schedule.DaySpan<>0 then wscript.echo "Previous Day frequency: " & schedule.DaySpan 
            if schedule.HourSpan<>0 then wscript.echo "Previous Hour frequency: " & schedule.HourSpan 
            if schedule.MinuteSpan<>0 then wscript.echo "Previous Minute frequency: " & schedule.MinuteSpan 
        end if 

        if schedtype="SMS_ST_RecurMonthlyByDate" then 
            if schedule.MonthDay<>0 then wscript.echo "Day of the month: " & schedule.MonthDay 
        end if

    End If

'    exampletype="recur" 
'    
'    if exampletype="monthday" then 
'        set newschedtype=gService.Get("SMS_ST_RecurMonthlyByDate").SpawnInstance_ 
'        newschedtype.MonthDay=9 
'        Collection.RefreshSchedule(0)=newschedtype 
'    end if
'
   ' if exampletype="recur" then 
   '     set newschedtype=gService.Get("SMS_ST_RecurInterval").SpawnInstance_ 
   '     newschedtype.DaySpan=1 
   '     newschedtype.HourSpan=0 
   '     newschedtype.MinuteSpan=0 
   '     Collection.RefreshSchedule(0)=newschedtype 
   'end if

  'Collection.Put_ 


End Sub 


