On Error Resume Next

Set oUIResManager = CreateObject("UIResource.UIResourceMgr")
Set oCache=oUIResManager.GetCacheInfo()

If oCache Is Nothing Then
    Set oUIResManager=Nothing
    Wscript.Echo "Could not get cache info - quitting"
    Wscript.Quit
End If

Wscript.Echo "Total size:      " & FormatNumber(oCache.TotalSize,0) + "MB"
