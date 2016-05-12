' Request Something Demonstration Script
' SMS App Request Workflow Proof of Concept
' Emmanuel Tsouris
' October 3, 2006
'
' Proof of concept request app
' executes an HTTP GET to hit a web service and pass a value to it
' 

Dim RequestedApplication

RequestedApplication = "sample app"

Set WshNetwork = WScript.CreateObject("WScript.Network")

WScript.Echo WshNetwork.UserDomain & "\" & WshNetwork.UserName & " from " & WshNetwork.ComputerName & " requests " & RequestedApplication

WScript.Echo getResults()

function getResults()
Dim xmlDOC
Dim bOK
Dim HTTP
Set HTTP = CreateObject("MSXML2.XMLHTTP")
Set xmlDOC =CreateObject("MSXML.DOMDocument")
xmlDOC.Async=False
HTTP.Open "GET","http://webserviceURLGoesHere.com/RequestApp" , False
HTTP.Send()
bOK = xmlDOC.load(HTTP.responseXML)
if Not bOK then
	WScript.Echo "Error loading XML from HTTP"
	WScript.Quit(1)
end if
' Note: Instead of making an XSL transform Stylesheet, we will use the selectNodes Method 
' with the XPath "//" search directive to get the ifornmation we want to display out of the
Dim objNodeList1
DIm objNodeList2
Set objNodeList1=xmlDOC.documentElement.selectNodes("//Node1")
Set objNodeList2=xmlDOC.documentElement.selectNodes("//Node2")

WScript.Echo "Len: "& objNodeList1.Length

For I = 0 to objNodeList1.Length -1
	WScript.Echo ": " & objNodeList1(i).text
	WScript.Echo " " & objNodeList2(i).text
next

WScript.Echo "Http Status: " & HTTP.statusText

end Function
