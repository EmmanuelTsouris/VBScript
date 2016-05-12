'*********************************************************
' SetPackageFlagDND - Do Not Download
' Author: Emmanuel Tsouris
' Version: 0.1
' Date: 4/29/2011
' Purpose:  Sets the DO_NOT_DOWNLOAD 0x01000000 (24) flag on all packages.
'           This flag tells SCCM that packages on BDPs will be pre-staged,
'           so don't download them to the BDP.
' Inputs:   none.
' Returns:  none.
'
' Reference:
'
'   0x01000000 (24) DO_NOT_DOWNLOAD
'   Do not download the package to branch distribution points, as it will be pre-staged.
'   
'   0x02000000 (25) PERSIST_IN_CACHE
'   Persist the package in the cache.
'   
'   0x04000000 (26) USE_BINARY_DELTA_REP
'   Marks the package to be replicated by distribution manager using binary delta replication.
'   
'   0x10000000 (28) NO_PACKAGE
'   The package does not require distribution points.
'   
'   0x20000000 (29) USE_SPECIAL_MIF
'   This value determines if Configuration Manager uses MIFName, MIFPublisher, and MIFVersion for MIF file status matching.
'   Otherwise, Configuration Manager uses Name, Manufacturer, and Version for status matching. For more information,
'   see the Remarks section later in this topic.
'   
'   0x40000000 (30) DISTRIBUTE_ON_DEMAND
'   The package is allowed to be distributed on demand to branch distribution points.
'   
'*********************************************************

'Set the Constants
DO_NOT_DOWNLOAD = &H01000000
PERSIST_IN_CACHE = &H02000000
USE_BINARY_DELTA_REP = &H04000000
NO_PACKAGE = &H10000000
USE_SPECIAL_MIF = &H20000000
DISTRIBUTE_ON_DEMAND = &H40000000

'Set the SCCM Server Name
sccmServerName = "serverName"

'Create the SWbemLocator object 
Set swbemLocator = CreateObject("WbemScripting.SWbemLocator")
'Create a SWbemServices object that is bound to the root\sms namespace
Set swbemconnection= swbemLocator.ConnectServer(sccmServerName, "root\sms")
'Get instances of the SMS_ProviderLocation class
Set providerLocs = swbemconnection.InstancesOf("SMS_ProviderLocation")

'Bind to the site code
For each loc in providerLocs
    If loc.ProviderForLocalSite = True Then
        sccmSiteCode = loc.Sitecode
        consoleLog("Connecting to Provider on " & loc.Machine & " at root\sms\site_" & sccmSiteCode)
        Set swbemconnection= swbemLocator.ConnectServer(sccmServerName, "root\sms\site_" & sccmSiteCode)
    end if
Next

'Execute a WMI Query that returns all the packages
Set packages = swbemconnection.ExecQuery("Select * From SMS_package")

consoleLog("PkgID" & vbTab & "Flag" & vbTab & "FlagText" & vbTab & "FlagNew")

'Future note:
'local var for flag to reduce cost of loop operation

'Loop through the packages
For Each package in packages

    If (package.PkgFlags AND DO_NOT_DOWNLOAD) = DO_NOT_DOWNLOAD Then
        consoleLog(package.PackageID & vbTab & "Detected DO_NOT_DOWNLOAD")
    End If
    
    If (package.PkgFlags AND PERSIST_IN_CACHE) = PERSIST_IN_CACHE Then
        consoleLog(package.PackageID & vbTab & "Detected PERSIST_IN_CACHE")
    End If
    
    If (package.PkgFlags AND USE_BINARY_DELTA_REP) = USE_BINARY_DELTA_REP Then
        consoleLog(package.PackageID & vbTab & "Detected USE_BINARY_DELTA_REP")
    End If
    
    If (package.PkgFlags AND NO_PACKAGE) = NO_PACKAGE Then
        consoleLog(package.PackageID & vbTab & "Detected NO_PACKAGE")
    End If
    
    If (package.PkgFlags AND USE_SPECIAL_MIF) = USE_SPECIAL_MIF Then
        consoleLog(package.PackageID & vbTab & "Detected USE_SPECIAL_MIF")
    End If
    
    If (package.PkgFlags AND DISTRIBUTE_ON_DEMAND) = DISTRIBUTE_ON_DEMAND Then
        consoleLog(package.PackageID & vbTab & "Detected DISTRIBUTE_ON_DEMAND")
    End If

    consoleLog(package.PackageID & vbTab & Hex(package.PkgFlags) & vbTab & Hex(package.PkgFlags or DO_NOT_DOWNLOAD))
    

    If (package.PkgFlags AND DO_NOT_DOWNLOAD) = DO_NOT_DOWNLOAD Then
    package.PkgFlags = package.PkgFlags xor DO_NOT_DOWNLOAD
    package.Put_
    End If

Next

'
'Some Common Functions
'

'Function to echo a line of text formatted with the date and time for logging
Function consoleLog( stringLine)
    WScript.Echo Date() & " " & Time() & vbTab & stringLine 
    consoleLog = true   
end Function

'Function to return a string with the package flag name instead of the hex value
Function getPackageFlagName(pkgFlags)
    Dim packageFlagName
    Select Case pkgFlags 
        Case DO_NOT_DOWNLOAD 
           packageFlagName = "DO_NOT_DOWNLOAD" 
        Case PERSIST_IN_CACHE
           packageFlagName = "PERSIST_IN_CACHE" 
        Case USE_BINARY_DELTA_REP
           packageFlagName = "USE_BINARY_DELTA_REP"
        Case NO_PACKAGE
            packageFlagName = "NO_PACKAGE"
        Case USE_SPECIAL_MIF
            packageFlagName = "USE_SPECIAL_MIF"
        Case DISTRIBUTE_ON_DEMAND
            packageFlagName = "DISTRIBUTE_ON_DEMAND"
        Case Else
            packageFlagName = Hex(pkgFlags)
    End Select 
    getPackageFlagName = packageFlagName
End Function




