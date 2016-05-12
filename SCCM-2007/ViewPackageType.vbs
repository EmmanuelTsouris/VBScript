'*********************************************************
' ViewPackageType
' Author: Emmanuel Tsouris
' Version: 0.1
' Date: 5/6/2011
' Purpose:  Display Packages by the various PackageTypes
'
' Inputs:   /SERVER:ServerName sets the servername to the value passed.
'               Use for servers other than TEST or PROD
'           /TEST sets the servername to the DEFAULT_SERVER_NAME_TEST value
'           /PROD sets the servername to the DEFAULT_SERVER_NAME_PROD value
'
' Returns:  Standard Output of Activity.
'
' Reference:

'
'*********************************************************

'Defaults
Const DEFAULT_SERVER_NAME_PROD = "serverName"
Const DEFAULT_SERVER_NAME_TEST = "serverNameTest"

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

'Execute a WMI Query that returns all the packages for a specific DP
Set objects = swbemconnection.ExecQuery("Select * From SMS_DistributionPoint where PackageID = 'ABC12345'")

'PackageID
'SMS_OperatingSystemInstallPackage

For each obj in objects
    consoleLog(obj.PackageID & " -> " & obj.ServerNALPath)
Next

'
'Some Common Functions
'

'Function to echo a line of text formatted with the date and time for logging
Function consoleLog( stringLine)
    WScript.Echo Date() & " " & Time() & vbTab & stringLine 
    consoleLog = true   
end Function