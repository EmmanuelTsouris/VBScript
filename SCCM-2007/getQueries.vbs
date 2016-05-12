
Set objLocator = CreateObject("WbemScripting.SWbemLocator")
Set objSWbemServices = objLocator.ConnectServer("serverName" , "root\sms\site_" & "ABC" )

Set colQueries=objSWbemServices.InstancesOf("SMS_Query")

wscript.echo "QueryID QueryName" 
wscript.echo "-----------------" 
For Each objQuery In colQueries 
    wscript.echo objQuery.QueryID + " "  + objQuery.Name 
Next

