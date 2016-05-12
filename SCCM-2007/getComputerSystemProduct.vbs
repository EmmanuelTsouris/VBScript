
'
' Loops through a collection of Computer System Products
'

Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystemProduct",,48)

'Loop through the collection
For Each objItem in colItems
	'Echo out the properties for each item
	Wscript.Echo "Caption: " & objItem.Caption
	Wscript.Echo "Description: " & objItem.Description
	Wscript.Echo "IdentifyingNumber: " & objItem.IdentifyingNumber
	Wscript.Echo "Name: " & objItem.Name
	Wscript.Echo "SKUNumber: " & objItem.SKUNumber
	Wscript.Echo "UUID: " & objItem.UUID
	Wscript.Echo "Vendor: " & objItem.Vendor
	Wscript.Echo "Version: " & objItem.Version & vbCrLf
Next