'*********************************************************
' setAppSyncMimeTypes
' Author: Emmanuel Tsouris
' Version: 0.3
' Purpose:  Adds MSI, DAT, and EXE mimetypes to IIS if they don't exist
'
' Inputs:   NONE
'
' Returns:  Standard Output of Activity.
'
' Reference:http://msdn.microsoft.com/en-us/library/ms752346.aspx
' 	    http://blogs.vmware.com/thinapp/2008/10/how-to-configur.html
'*********************************************************

Const ADS_PROPERTY_UPDATE = 2

dim mimeMapEntry, allMimeMaps

dim foundMSI, foundDAT, foundEXE

foundMSI = false
foundDAT = false
foundEXE = false

WScript.Echo(Now() & vbTab & "Begin Execution")

' Get the mimemap object.
Set mimeMapEntry = GetObject("IIS://localhost/MimeMap")
allMimeMaps = mimeMapEntry.GetEx("MimeMap")

' Display the mappings in the table.
For Each mimeMap In allMimeMaps
    If mimeMap.Extension = ".msi" Then
        WScript.Echo(Now() & vbTab & mimeMap.MimeType & " (" & mimeMap.Extension + ") exists")
        foundMSI = true
    End If

    If mimeMap.Extension = ".dat" Then
        WScript.Echo(Now() & vbTab & mimeMap.MimeType & " (" & mimeMap.Extension + ") exists")
        foundDAT = true
    End If
    
    If mimeMap.Extension = ".exe" Then
        WScript.Echo(Now() & vbTab & mimeMap.MimeType & " (" & mimeMap.Extension + ") exists")
        foundEXE = true
    End If
Next

If foundMSI = false Then
    AddMimeType ".msi", "application/octet-stream"
    WScript.Echo(Now() & vbTab & "Added MSI MimeType")
End If

If foundDAT = false Then
    AddMimeType ".dat", "application/octet-stream"
    WScript.Echo(Now() & vbTab & "Added DAT MimeType")
End If

If foundEXE = false Then
    AddMimeType ".exe", "application/octet-stream"
    WScript.Echo(Now() & vbTab & "Added EXE MimeType")
End If

' Create a Shell object
Set WshShell = CreateObject("WScript.Shell")

' Stop and Start the IIS Service
WScript.Echo(Now() & vbTab & "Stopping IIS")
Set oExec = WshShell.Exec("net stop w3svc")
Do While oExec.Status = 0
    WScript.Sleep 100
Loop

WScript.Echo(Now() & vbTab & "Starting IIS")
Set oExec = WshShell.Exec("net start w3svc")
Do While oExec.Status = 0
    WScript.Sleep 100
Loop

WScript.Echo(Now() & vbTab & "Execution Completed")


' AddMimeType Sub
Sub AddMimeType (Ext, MType)

    ' Get the mappings from the MimeMap property. 
    MimeMapArray = mimeMapEntry.GetEx("MimeMap") 

    ' Add a new mapping. 
    i = UBound(MimeMapArray) + 1 
    Redim Preserve MimeMapArray(i) 
    Set MimeMapArray(i) = CreateObject("MimeMap") 
    MimeMapArray(i).Extension = Ext 
    MimeMapArray(i).MimeType = MType 
    mimeMapEntry.PutEx ADS_PROPERTY_UPDATE, "MimeMap", MimeMapArray
    mimeMapEntry.SetInfo
    
End Sub