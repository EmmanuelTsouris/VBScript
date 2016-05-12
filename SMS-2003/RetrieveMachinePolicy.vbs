On Error Resume Next

Dim objCPAppletMgr
Dim objClientActions
Dim objClientAction
Dim strActionName

strActionName="Request & Evaluate Machine Policy"

'Get the Control Panel applet manager object
set  objCPAppletMgr = CreateObject("CPApplet.CPAppletMgr")

'Get a collection of client actions

set objClientActions=objCPAppletMgr.GetClientActions

'Loop through the available client actions

For Each objClientAction In objClientActions
   
  If objClientAction.Name = strActionName Then
    objClientAction.PerformAction  
     WScript.Echo "Action " + objClientAction.Name + " initiated" 
  End If
Next
