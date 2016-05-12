'
' Get Package Status
' v0.1 6/18/2009
' Emmanuel Tsouris
'

'Get the named arguments
Set colNamedArguments = WScript.Arguments.Named

'Advertisement ID
smsPackageID = colNamedArguments.Item("package")


'SMS Primary Site Server Name
smsServerName = colNamedArguments.Item("server")

'SMS Primary Site Code
smsSiteCode = colNamedArguments.Item("site")

'Connect to the server
Set objLocator = CreateObject("WbemScripting.SWbemLocator")

Set oSWbemServices = objLocator.ConnectServer(smsServerName , "root\sms\site_" & smsSiteCode )

Dim oQueryResults
Dim oSingleResult

' Execute WMI Query
Set oQueryResults = oSWbemServices.ExecQuery("Select * From SMS_AdvertisementStatusRootSummarizer Where PackageID=""" & smsPackageID & """")

WScript.Echo "AdvertisementID, AdvertisementName, AdvertisementsFailed, AdvertisementsReceived, CollectionID, CollectionName, ExpirationTime, PackageID, PackageName, PackageVersion, PresentTime, ProgramName, ProgramsFailed, ProgramsStarted, ProgramsSucceeded, SourceSite"

' Loop through results and output collection information
For Each oSingleResult In oQueryResults

	WScript.Echo oSingleResult.AdvertisementID & ", " & _
		oSingleResult.AdvertisementName & ", " & _
		oSingleResult.AdvertisementsFailed & ", " & _
		oSingleResult.AdvertisementsReceived & ", " & _
		oSingleResult.CollectionID & ", " & _
		oSingleResult.CollectionName & ", " & _
		WMIDateStringToDate(oSingleResult.ExpirationTime) & ", " & _
		oSingleResult.PackageID & ", " & _
		oSingleResult.PackageName & ", " & _
		oSingleResult.PackageVersion & ", " & _
		WMIDateStringToDate(oSingleResult.PresentTime) & ", " & _
		oSingleResult.ProgramName & ", " & _
		oSingleResult.ProgramsFailed & ", " & _
		oSingleResult.ProgramsStarted & ", " & _
		oSingleResult.ProgramsSucceeded & ", " & _
		oSingleResult.SourceSite

Next

Set oSWbemServices=nothing
Set objLocator=nothing

Function WMIDateStringToDate(dtmInstallDate)
    WMIDateStringToDate = CDate(Mid(dtmInstallDate, 5, 2) & "/" & _
        Mid(dtmInstallDate, 7, 2) & "/" & Left(dtmInstallDate, 4) _
            & " " & Mid (dtmInstallDate, 9, 2) & ":" & _
                Mid(dtmInstallDate, 11, 2) & ":" & Mid(dtmInstallDate, _
                    13, 2))
End Function