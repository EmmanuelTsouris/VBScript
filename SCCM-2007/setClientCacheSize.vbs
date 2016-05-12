'
' Set SMS Client Cache Configuration
' v1.0 5/2/2006
' Emmanuel Tsouris
'

	On Error Resume Next 

' Declare Variables
	Dim checkCacheValue 
	Dim setCacheValue

' Set the Cache Size to check for (less than or equal to)
	checkCacheValue = 1024 

' If less than or equal to the checkCacheValue
' set the Cache Size to this
	setCacheValue= 1024 

	Dim oUIResourceMgr 
	Dim oCache 

	Set oUIResourceMgr = CreateObject("UIResource.UIResourceMgr") 
	Set oCacheInfo = oUIResourceMgr.GetCacheInfo 

' Set it if it's less than or equal to checkCacheValue 
	if oCacheInfo.TotalSize <= checkCacheValue then 
		oCacheInfo.TotalSize = setCacheValue
	end if

WScript.Quit(Err)
