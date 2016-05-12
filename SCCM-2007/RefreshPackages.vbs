'*********************************************************
' RefreshPackages
' Author: Emmanuel Tsouris
' Version: 0.2
' Date: 5/6/2011
' Purpose:  Refresh all packages on a specific DP. If a specific package is specified
'           then only that package is refreshed.
'
' Inputs:   /SERVER:ServerName sets the servername to the value passed.
'               Use for servers other than TEST or PROD
'           /TEST sets the servername to the DEFAULT_SERVER_NAME_TEST value
'           /PROD sets the servername to the DEFAULT_SERVER_NAME_PROD value
'
'           /DP:ServerName or ALL [default: ALL] ServerName can be a
'               partial name, such as SCCDP and will hit all matching DPs.
'
'           /Package:Eight Character Package ID [default: ALL]
'
' Returns:  Standard Output of Activity.
'
' Reference:none.
'*********************************************************

'Defaults
Const DEFAULT_SERVER_NAME_PROD = "serverName"
Const DEFAULT_SERVER_NAME_TEST = "serverNameTest"
Const DEFAULT_PACKAGEID = "ALL"
Const DEFAULT_DP = "ALL"

Set commandlineArguments = WScript.Arguments.Named

'Process Command Line Arguments and set Local Variables using Arguments or Defaults
'

'Set the server to connect to.
'NOTE: The script will quit if no server is passed and if the /Prod or /Test options are left out
'If both prod and test are passed, the script will use the test server.
If commandlineArguments.Exists("Test") Then
    sccmServerName = DEFAULT_SERVER_NAME_TEST
    consoleLog "SCCM Server [TEST]:" & sccmServerName
ElseIf commandlineArguments.Exists("Prod") Then
    sccmServerName = DEFAULT_SERVER_NAME_PROD
    consoleLog "SCCM Server [PROD]:" & sccmServerName
ElseIf commandlineArguments.Exists("Server") Then
    sccmServerName = commandlineArguments.Item("Server")
    consoleLog "SCCM Server [TEST]:" & sccmServerName
Else
    consoleLog "NOTE: You must provide either the /Prod or /Test options or /SERVER:servername."
    WScript.Quit(0)
End If

If commandlineArguments.Exists("DP") Then
    distributionPoint = commandlineArguments.Item("DP")
    consoleLog "Distribution Point:" & distributionPoint
Else
    distributionPoint = DEFAULT_DP
    consoleLog "Distribution Point [default]:" & distributionPoint
End If

If commandlineArguments.Exists("Package") Then
    packageID = commandlineArguments.Item("Package")
    consoleLog "Package:" & packageID
Else
    packageID = DEFAULT_PACKAGEID
    consoleLog "Package [default]:" & packageID
End If

'Create the SWbemLocator object 
Set swbemLocator = CreateObject("WbemScripting.SWbemLocator")
'Create a SWbemServices object that is bound to the root\sms namespace
Set swbemconnection= swbemLocator.ConnectServer(sccmServerName, "root\sms")
'Get instances of the SMS_ProviderLocation class
Set providerLocs = swbemconnection.InstancesOf("SMS_ProviderLocation")

'Bind to the site code
For each loc in providerLocs
    If loc.ProviderForLocalSite = True Then
        sccmSiteCode = loc.Sitecode
        consoleLog("Connecting to Provider on " & loc.Machine & " at root\sms\site_" & sccmSiteCode)
        Set swbemconnection= swbemLocator.ConnectServer(sccmServerName, "root\sms\site_" & sccmSiteCode)
    end if
Next

If packageID = "ALL" and distributionPoint = "ALL" Then
    'Execute a WMI Query that returns all the packages for a specific DP
    Set distributionPoints = swbemconnection.ExecQuery("Select * From SMS_DistributionPoint " & _
        "where Status <> 3")
ElseIf packageID = "ALL" and distributionPoint <> "ALL" Then
    'Execute a WMI Query that returns all the packages for a specific DP
    Set distributionPoints = swbemconnection.ExecQuery("Select * From SMS_DistributionPoint " & _
        "where ServerNALPath like '%" & distributionPoint & "%' and Status <> 3")
Else
    'Execute a WMI Query that returns the matching package for a specific DP
    Set distributionPoints = swbemconnection.ExecQuery("Select * From SMS_DistributionPoint " & _
        "where ServerNALPath like '%" & distributionPoint & "%' and PackageID = '" & packageID & _
        "' and Status <> 3")
End IF

For each DP in distributionPoints
    consoleLog("Refresh " & DP.PackageID & " on " & DP.ServerNALPath)
    DP.RefreshNow = True 
    DP.Put_
Next

'
'Some Common Functions
'

'Function to echo a line of text formatted with the date and time for logging
Function consoleLog( stringLine)
    WScript.Echo Date() & " " & Time() & vbTab & stringLine 
    consoleLog = true   
end Function