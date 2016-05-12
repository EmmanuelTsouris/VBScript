'
' Configure SMS Client Cache
' v1.2 6/21/2006
' Emmanuel Tsouris
'
' CacheMin: Size to check for
' CacheSize: Size to set if Cache is less than CacheMin
'
' example: configureclientcachesize.vbs /CacheMin:1024 /CacheSize:1024
' example: configureclientcachesize.vbs /CacheSize:1024
'
'

On Error Resume Next 

' Get the named arguments
Set colNamedArguments = WScript.Arguments.Named

checkCacheValue = colNamedArguments.Item("CacheMin")
setCacheValue = colNamedArguments.Item("CacheSize")

Dim oUIResourceMgr 
Dim oCache 

Set oUIResourceMgr = CreateObject("UIResource.UIResourceMgr") 
Set oCacheInfo = oUIResourceMgr.GetCacheInfo 

' Set a default 
If setCacheValue = 0 or setCacheValue = "" then
	setCacheValue = 1024
end if

If (checkCacheValue = 0) or (checkCacheValue  =  "") then

	' Set the cache to this new value regardless of the current size
	oCacheInfo.TotalSize = setCacheValue

Else

	' If the current cache size is less than or equal to checkCacheValue, then
	' set it to setCacheValue 
	
	if oCacheInfo.TotalSize <= checkCacheValue then 
		oCacheInfo.TotalSize = setCacheValue
	end if

End If

' Return the error code
WScript.Quit(Err)

