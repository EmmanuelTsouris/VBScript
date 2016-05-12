On Error Resume Next
 
 
Dim oCPAppletMgr        'Control Applet manager object.
Dim oClientComponent    'Individual client components.
Dim oClientComponents   'A collection of client components.
 
 
'Get the Control Panel applet manager object.
Set  oCPAppletMgr=CreateObject("CPApplet.CPAppletMgr")
If oCPAppletMgr Is Nothing Then
     Wscript.echo "Could not create control panel application manager"
     wscript.quit
End If
 
 
'Get a collection of components.
 
Set oClientComponents=oCPAppletMgr.GetClientComponents
If oClientComponents Is Nothing Then
     wscript.echo "Could not get the client components"
     Set oCPAppletMgr=Nothing
     wscript.quit
End If
 
 
 
wscript.echo "There are "  &oClientComponents.Count & " components"
wscript.echo
'Display each client action.
For Each oClientComponent In oClientComponents
 
     wscript.echo oClientComponent.DisplayName
     Select Case oClientComponent.State
     Case 0
         wscript.echo "installed"
     Case 1 
         wscript.echo "enabled"
     Case 2
         wscript.echo "disabled"
     End Select
     wscript.echo
 
Next
 
Set oClientComponents=Nothing
Set oCPAppletMgr=Nothing