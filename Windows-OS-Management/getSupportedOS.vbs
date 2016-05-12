 
'
' Checks the OS Version to see if it matches the defined versions
' 2/1/2005
'

    IsSupportedOS = False
    Dim objWMI, operatingSystems, operatingSystem, strOSVer
    On Error Resume Next
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set operatingSystems = objWMI.ExecQuery("Select * from Win32_OperatingSystem")

    For Each operatingSystem In operatingSystems
      strOSVer = Left(operatingSystem.Version, 3)

	WScript.Echo "OS Version: " & operatingSystem.Version
	WScript.Echo "SP Version: " & operatingSystem.ServicePackMajorVersion

      Select Case StrOsVer
        Case "5.0"
          If operatingSystem.ServicePackMajorVersion > 2 Then IsSupportedOS = True
        Case "5.1"
          If operatingSystem.ServicePackMajorVersion > 0 Then IsSupportedOS = True
	Case "5.2"
		IsSupportedOS=True
      End Select

    Next

    Set colOperatingSystems = Nothing
    Set objWMI = Nothing

WScript.Echo "IsSupportedOS: " & IsSupportedOS