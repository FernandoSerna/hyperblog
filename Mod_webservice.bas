Attribute VB_Name = "Module1"
Function invokewebservice(strSoap, strSOAPAction, strurl, ByRef xmlResponse) As Boolean
Dim xmlhttp As MSXML2.XMLHTTP30
Dim blnsucces As Boolean

Set xmlhttp = New MSXML2.XMLHTTP30
xmlhttp.Open "POST", strurl, False
xmlhttp.setRequestHeader "Man", "POST " & strurl & " HTTP/1.1"
xmlhttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
xmlhttp.setRequestHeader "SOAPAction", strSOPAAction
Call xmlhttp.send(strSoap)

If xmlhttp.status = 200 Then
   blnsuccess = True
Else
   blnsuccess = False
End If

'MsgBox xmlhttp.responseXML
Set xmlResponse = xmlhttp.responseXML
invokewebservice = blnsuccess
Set xmlhttp = Nothing


End Function
