'
' Connects to a Windows Computer and lists the print queue information
'

' List of computers to connect to
arrComputers = Array("localhost")

For each sServer in arrComputers
	WScript.Echo("Getting Print Queue Info for " & sServer)
	set objPrinters = GetOBJect("WinNT://" & sServer &",computer")
	
	'If there is an error with the object then it will print Failed
	If Err.Number <> 0 Then
		WScript.Echo("Error " & Err.Number & ": Connection to " & sServer & " Failed")
		WScript.Echo(Err.description)
		WScript.Echo(Err.helpfile)
	Else

		 ' Filter Unwanted Properties
		objPrinters.Filter = Array("PrintQueue")
		
		For Each p In objPrinters
			i = i + 1

			WSCript.Echo(p.name)

			'Set pq = GetObject(p.ADsPath)
			'WSCript.Echo(pq.name)		
			'Set pq = nothing
			
		Next

		WScript.Echo("Count = " & i)

	end if

Next

WScript.Quit(0)

WScript.Echo(" Try WMI ")

server = arrComputers(0)

set wmi = getobject("winmgmts://" & server & "/root/cimv2")
wql = "select * from win32 share where type=1"
set results = wmi.execquery(wql)
WScript.Echo("Results = " & results.count)
for each printer in results
  'WSCript.Echo(printer.name)
  'printer.caption
  'printer.description
	WSCript.Echo("name=" & printer.name)
next

