On Error Resume Next

sActionName="Hardware Inventory Collection Cycle"

' Create a CPAppletMgr instance

Dim oCPAppletMgr
Set oCPAppletMgr = CreateObject("CPApplet.CPAppletMgr")

 ' Get the available ClientActions

Dim oClientActions
Set oClientActions = oCPAppletMgr.GetClientActions()

 ' Loop through the available client actions

Dim oClientAction
For Each oClientAction In oClientActions

    ' Is this the action we want to start?

     If oClientAction.Name = sActionName Then

    ' Start the action

        oClientAction.PerformAction  
    End If
Next
