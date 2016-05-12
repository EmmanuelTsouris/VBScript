Sub QueryPackageStatus (connection) 
 
    Dim query 
    Dim sink 
    Dim minutes 
 
 
    Set sink = WScript.CreateObject("wbemscripting.swbemsink","sink_") 
     
    ' You have to specify a polling interval because Configuration Manager 
    ' does not provide an intrinsic event provider for these classes. 
    Query = "SELECT * FROM __InstanceCreationEvent Within 120 " & _ 
            "WHERE TargetInstance.__Class = 'SMS_PackageStatusRootSummarizer' " 
    connection.ExecNotificationQueryAsync sink, query 
 
    query = "SELECT * FROM __InstanceModificationEvent Within 120 " & _ 
            "WHERE TargetInstance.__Class = 'SMS_PackageStatusRootSummarizer' " 
    connection.ExecNotificationQueryAsync sink, query 
 
    minutes = 0 
    
    ' Loop for 5 minutes. 
    While minutes < 300 
        wscript.sleep 1000 
        minutes = minutes + 1 
    Wend  
         
    sink.Cancel 
    Set sink = nothing 
      
 End Sub    
 
' The sink subroutine to handle the OnObjectReady  
' event. This is called as each object returns. 
Sub sink_OnObjectReady(statusEvent, octx) 
   Wscript.Echo "Name: " + statusEvent.TargetInstance.Name 
   Wscript.Echo "Targeted: " + CStr(statusEvent.TargetInstance.Targeted) 
   Wscript.Echo "Installed: " + CStr(statusEvent.TargetInstance.Installed) 
   Wscript.Echo 
End Sub 
 
Sub sink_OnCompleted(Hresult, oErr, oCtx) 
    Wscript.Echo "Finished" 
End Sub