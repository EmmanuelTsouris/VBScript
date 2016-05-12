
'
' Get SMS Query Results
' Emmanuel Tsouris
' 3/28/08

Set objLocator = CreateObject("WbemScripting.SWbemLocator")
Set objSWbemServices = objLocator.ConnectServer("serverName" , "root\sms\site_" & "ABC" )

strQueryID="ABC123" 


Set objQuery=  _ 
    objSWbemServices.Get _ 
    ("SMS_Query.QueryID='" + strQueryID +" '" )

wscript.echo objQuery.Name 
wscript.echo "----------------------------------" 
Set colQueryResults=objSWbemServices.ExecQuery(objQuery.Expression) 

'Use this to get the query and property (field) names
WScript.Echo "Executing: " & objQuery.Expression

For Each objResult In colQueryResults 
	WScript.Echo objResult.Name
Next



' Since each query may have different properties, you will have to customize the loop above for the property names. 
'
' Example for strQueryID="SMS001" 
'
'    WScript.Echo objResult.Name & _
'	vbTab & objResult.OperatingSystemNameandVersion & _
'	vbTab & objResult.ResourceID
'
'	'Some properties are collections, like SMSAssignedSites
'	For each SMSAssignedSite in objResult.SMSAssignedSites
'		WScript.Echo vbTab & "Assigned Site = " & SMSAssignedSite 
'	Next