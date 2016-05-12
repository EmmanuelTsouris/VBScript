On Error Resume Next

' Declare Variables
Dim checkCacheValue
Dim setCacheValue

' Location for the Cache
' default is "C:\WINDOWS\system32\CCM\Cache"
CacheLocation = "C:\WINDOWS\system32\CCM\Cache"

Dim oUIResourceMgr
Dim oCache

Set oUIResourceMgr = CreateObject("UIResource.UIResourceMgr")
Set oCacheInfo = oUIResourceMgr.GetCacheInfo

' Echo out the old location
WScript.Echo Now() & " SMS Client cache location was: " & oCacheInfo.Location

' Set the new location
oCacheInfo.Location = CacheLocation


' Echo out the new location
WScript.Echo Now() & " SMS Client cache location is now: " & oCacheInfo.Location

'Return the error so SMS can report it
WScript.Quit(Err)
