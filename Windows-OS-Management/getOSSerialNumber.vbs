
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")

' Get OS Serial Number

Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem",,48)

'Loop through the collection of Win32_PatchState_Extended items
For Each objItem in colItems
'Echo out the properties for each item
Wscript.Echo "SerialNumber: " & objItem.SerialNumber
Next