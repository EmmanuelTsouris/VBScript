
' SetSiteCode
' Emmanuel Tsouris
' v0.3
' Sets the SCCM site code
'
' Example: /SiteCode:<the three digit site code>
'

On Error Resume Next

' Get the named arguments
Set colNamedArguments = WScript.Arguments.Named

Dim sccmClient

siteCode = colNamedArguments.Item("SiteCode")

Set sccmClient= CreateObject ("Microsoft.SMS.Client")

If Err.Number <> 0 Then 
    WScript.Echo Now() & " Error creating the Client Object, quitting.."
    WScript.quit
End If

Set serverNames = CreateObject("Scripting.Dictionary")

If Err.Number <> 0 Then 
    WScript.Echo Now() & " Error creating the dictionary Object, quitting."
    WScript.quit
End If

' Authorize only the server names below to run this script.
' Array of Server Names, along with a boolean value of true or false.
' if the servername is not in the array, or it is set to false, the script will quit.
serverNames.add "SERVER1", true
serverNames.add "SERVER2", true
serverNames.add "SERVER3", true

Set wshShell = WScript.CreateObject( "WScript.Shell" )

If Err.Number <> 0 Then 
    WScript.Echo Now() & " Error getting the computername, quitting."
    WScript.quit
End If

computerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )

WScript.Echo Now() & " Computer Name: " & computerName

If serverNames.item(computerName) = false Then
	Wscript.Echo Now() & " This script is not authorized to run on this system, quitting."
	WScript.Quit(0)
End If

If siteCode = "?" or len(SiteCode) <> 3 Then
	WScript.Echo Now() & " Query Mode: Site code is set to " & sccmClient.GetAssignedSite
	WScript.Quit(0)
End If

If (sccmClient.GetAssignedSite <> siteCode) and (serverNames.item(computerName)) then
	sccmClient.SetAssignedSite siteCode
	WScript.Echo Now() & " Set site code to " & siteCode
Else
	WScript.Echo Now() & " Site code already set to " & siteCode
End If

WScript.Echo Now() & " done." 

Set sccmClient=Nothing

