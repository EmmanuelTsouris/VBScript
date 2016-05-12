
'
' Shows the members of a local group on a given SMS Server
'

'List groups
serverName = "sccmServerName"
 showGroupMembers serverName,"Power users"
 showGroupMembers serverName,"Remote Desktop Users"

'List groups
serverName = "hdsccpn7"
 showGroupMembers serverName,"Power users"
 showGroupMembers serverName,"Remote Desktop Users"
 showGroupMembers serverName,"SMS Admins"
 showGroupMembers serverName,"Distributed COM Users"
 showGroupMembers serverName,"SMS_SiteSystemToSiteServerConnection_ENT"
 showGroupMembers serverName,"SMS_SiteToSiteConnection_ENT"
 'this one will be on the reporting point
 'showGroupMembers serverName,"SMS Reporting Users"


'
' Reusable Function to Display the Group Members
'

Function showGroupMembers(server, group)

    Set objGroup = GetObject("WinNT://" & server & "/" & group & ",group")
    wscript.echo 
    For Each Member In objGroup.Members
            wscript.echo server & vbTab & group & vbTab & Member.Name
    Next
    wscript.echo vbCrLf
    Set objGroup = Nothing

    showGroupMembers = true
end Function
