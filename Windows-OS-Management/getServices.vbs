'
' Export Windows Services to a CSV
' 5/24/2006
'

Const ForAppending = 8

Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objLogFile = objFSO.OpenTextFile("%temp%\ServiceList.csv", ForAppending, True)
 
objLogFile.Write _
	("System Name,Service Name,Service Type,Service State,Exit " _
	& "Code,Process ID,Can Be Paused,Can Be Stopped,Caption," _
	& "Description,Can Interact with Desktop,Display Name,Error " _
	& "Control,Executable Path Name,Service Started," _
	& "Start Mode,Account Name ")
 
objLogFile.Writeline

strComputer = "."

Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colListOfServices = objWMIService.ExecQuery("SELECT * FROM Win32_Service")
 
For Each objService in colListOfServices
	objLogFile.Write(objService.SystemName) & ","
	objLogFile.Write(objService.Name) & ","
	objLogFile.Write(objService.ServiceType) & ","
	objLogFile.Write(objService.State) & ","
	objLogFile.Write(objService.ExitCode) & ","
	objLogFile.Write(objService.ProcessID) & ","
	objLogFile.Write(objService.AcceptPause) & ","
	objLogFile.Write(objService.AcceptStop) & ","
	objLogFile.Write(objService.Caption) & ","
	objLogFile.Write(objService.Description) & ","
	objLogFile.Write(objService.DesktopInteract) & ","
	objLogFile.Write(objService.DisplayName) & ","
	objLogFile.Write(objService.ErrorControl) & ","
	objLogFile.Write(objService.PathName) & ","
	objLogFile.Write(objService.Started) & ","
	objLogFile.Write(objService.StartMode) & ","
	objLogFile.Write(objService.StartName) & ","
	objLogFile.Writeline
Next

objLogFile.Close

