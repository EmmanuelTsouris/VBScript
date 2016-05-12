
'
' getSCCMClientGUID
' Returns the SCCM Client GUID
' 3/27/12
'

On Error Resume Next
Declare Count

Set objWMIService = GetObject("winmgmts:\\.\root\ccm")
Set colItems = objWMIService.ExecQuery("Select ClientID From CCM_Client",,48)

For Each objItem in colItems
	Wscript.Echo "Client ID" & objItem.ClientID
Next 

