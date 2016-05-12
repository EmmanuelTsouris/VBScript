
' runSMSClientActions
'

' Set required variables.
actionNameToRun = "Software Metering Usage Report Cycle"

' Create a CPAppletMgr instance.
Dim CPAppletMgr
Set CPAppletMgr = CreateObject("CPApplet.CPAppletMgr")

' Get the available client actions.
Dim clientActions
Set clientActions = CPAppletMgr.GetClientActions()

' Loop through the available client actions. Run matching client action when found.
Dim clientAction
For Each clientAction In clientActions

WScript.Echo clientAction.Name

    If clientAction.Name = actionNameToRun Then
        'clientAction.PerformAction  
    End If
Next

' Output success message.
wscript.echo "Ran: " & actionNameToRun

