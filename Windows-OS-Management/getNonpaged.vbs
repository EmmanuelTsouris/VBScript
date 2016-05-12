Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2") 
Set colItems = objWMIService.ExecQuery( _
    "SELECT * FROM Win32_PerfFormattedData_PerfOS_Memory",,48) 

For Each objItem in colItems 

	pagedMemory = objItem.PoolPagedBytes / 1024
	nonPaged = objItem.PoolNonpagedBytes / 1024

	totalMemory = 	pagedMemory + nonPaged 
	
	WScript.Echo Now() & vbTab & totalMemory
	WScript.Echo Now() & vbTab & pagedMemory
	WScript.Echo Now() & vbTab & nonPaged
Next