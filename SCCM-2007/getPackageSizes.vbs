
'Get the named arguments
Set colNamedArguments = WScript.Arguments.Named

'SMS Primary Site Server Name
smsServerName = colNamedArguments.Item("server")

'SMS Primary Site Code
smsSiteCode =colNamedArguments.Item("site")

'Connect to the server
Set objLocator = CreateObject("WbemScripting.SWbemLocator")

Set oSWbemServices = objLocator.ConnectServer(smsServerName , "root\sms\site_" & smsSiteCode )

Dim oQueryResults
Dim oSingleResult

' Execute WMI Query
Set oQueryResults = oSWbemServices.ExecQuery("Select PackageID, Name, PkgSourcePath from SMS_Package")

	Wscript.echo "packageID, packageName, packagePath, folderSize, fileCount, folderCount"

' Loop through results and output collection information
For Each oSingleResult In oQueryResults

	
	getPackageSize oSingleResult.PackageID, oSingleResult.Name, oSingleResult.PkgSourcePath
	'WScript.Quit(0) 'Exit after showing the first one
Next

Set oSWbemServices=nothing
Set objLocator=nothing

sub getPackageSize(pkgId, pkgName, pkgPath)

	set oFS = WScript.CreateObject("Scripting.FileSystemObject")
	set oF = oFS.GetFolder("\\primarysiteserver\smspkgX$\" & pkgID)

	Wscript.echo "Package " & pkgId & "," & pkgName & "," & pkgPath & "," & oF.Size & "," & oF.Files.Count & "," & oF.Subfolders.count


	'dim F
	'for each F in oF.Subfolders
	'ShowFolderDetails(F)
	'next

end sub

Function WMIDateStringToDate(dtmInstallDate)
    WMIDateStringToDate = CDate(Mid(dtmInstallDate, 5, 2) & "/" & _
        Mid(dtmInstallDate, 7, 2) & "/" & Left(dtmInstallDate, 4) _
            & " " & Mid (dtmInstallDate, 9, 2) & ":" & _
                Mid(dtmInstallDate, 11, 2) & ":" & Mid(dtmInstallDate, _
                    13, 2))
End Function