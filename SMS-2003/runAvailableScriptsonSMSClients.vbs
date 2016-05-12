On Error Resume Next

Dim oUIResource 
Dim sProgramID
Dim sPackageID
Dim sPackageName 
Dim oArgs
Dim objSWbemlocator
Dim objSWbemServices
Dim szNameSpacePath
Dim oProgramObjectSet
Dim oProgramObject
Dim sProgramObjectPath


szNameSpacePath = "root/microsoft/sms/client/swdist"

sPackageName = "PackageName"
sPackageID = "99999999"
sProgramID = "ProgramID" 'use the program name


Set objSWbemlocator = CreateObject("WbemScripting.SWbemLocator.1")
Set objSWbemServices = objSWbemlocator.ConnectServer(".",szNameSpacePath)

If Err.Number <> 0 Then

    ' Trying to create UIResource - assuming Advanced Client

    Set oUIResource = CreateObject ("UIResource.UIResourceMgr")
    If oUIResource Is Nothing Then 
        Wscript.Quit(2)
    End If

    ' Run the program

    oUIResource.ExecuteProgram sProgramID, sPackageID, TRUE

    Wscript.Echo "Successfully executed program"
    Wscript.Quit (0)

End If

' Get the program index

Set oProgramObjectSet = _
    oServices.ExecQuery _
    ("Select * from CLI_AvailableProgram where ProgramName=""" + _
        sProgramID + """ and sPackageName=""" + sPackageName + """")

' Run Each program in the objectset ... you might get more programs 
'because key is package name and not package ID

For Each oProgramObject In oProgramObjectSet 
    sProgramObjectPath = oProgramObject.Path_.RelPath

    ' Actually calling the WbemMethod to run the program now

    oServices.ExecMethod sProgramObjectPath,"RunNow"


