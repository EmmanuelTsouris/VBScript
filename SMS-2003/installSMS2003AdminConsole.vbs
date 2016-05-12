'Install SMS 2003 Console

Set WshShell = WScript.CreateObject("WScript.Shell")

'Installs SMS 2003 SP1 Administrators Console
WshShell.run "\\PATH\TO\SMS\SMS2003SP1\SMSSETUP\BIN\I386\Setup.EXE /script console.ini /NoUserInput", 1, true

'Upgrade the SMS Administrator Console to SMS 2003 SP2
WshShell.run "\\PATH\TO\SMS\SMS2003SP2\SMSSETUP\BIN\I386\Setup.EXE /Upgrade /NoUserInput", 1, true

'Installs the OSD Feature Pack
WshShell.run "msiexec /i \\PATH\TO\sms\SMS2003OSDFP\OSDeployment.msi /qbn!- REBOOT=R /L*v C:\tmp\smsvbs.log", 1, true

'Windows XP SP2 DCOM Configuration
WshSHell.run "regedit /s \\PATH\TO\sms\SMS2003OSDFP\sms.reg", 1, true

'firewall.cpl - uncheck Don't allow exceptions check box
WshSHell.run "regedit /s \\PATH\TO\sms\SMS2003OSDFP\firewall.reg", 1, true

'firewall.cpl - add Programs to the XP SP2 exceptions tab (unsecapp.exe,statview.exe)
WshSHell.run "regedit /s \\PATH\TO\sms\SMS2003OSDFP\exceptions.reg", 1, true

'firewall.cpl - add ports to XP SP2 the exception tab (Port=135, Port Name=SMS Admin Console 135, Select=TCP)
WshSHell.run "regedit /s \\PATH\TO\sms\SMS2003OSDFP\port.reg", 1, true