'
' Get Computer Collections
' v0.1 1/21/2009
' Emmanuel Tsouris
' emmanuel.tsouris@lmco.com
'

'On Error Resume Next 

'Get the named arguments
Set colNamedArguments = WScript.Arguments.Named

'SMS Primary Site Server Name
smsServerName = colNamedArguments.Item("server")

'SMS Primary Site Code
smsSiteCode = colNamedArguments.Item("site")

'Computer Name
ComputerName = colNamedArguments.Item("ComputerName")

'Connect to the server
Set objLocator = CreateObject("WbemScripting.SWbemLocator")

Set oSWbemServices = objLocator.ConnectServer(smsServerName , "root\sms\site_" & smsSiteCode )

Dim oCollectionSet
Dim oCollection

' Execute WMI Query
Set oCollectionSet = oSWbemServices.ExecQuery("select SMS_Collection.CollectionID, SMS_Collection.NAme from SMS_R_System inner join SMS_FullCollectionMembership on SMS_R_System.ResourceID = SMS_FullCollectionMembership.ResourceID inner join SMS_Collection on SMS_Collection.CollectionID = SMS_FullCollectionMembership.CollectionID Where SMS_R_System.Name = '" & ComputerName & "'")

' Loop through results and output collection information
For Each oCollection In oCollectionSet
    WScript.Echo oCollection.CollectionID & ", " & oCollection.Name
Next




