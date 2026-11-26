<%
On Error Resume Next
Set o = server.CreateObject("foo.bar")
'On Error Goto 0
If Err.Number <> 0 Then
  response.write Err.description
Else
  Response.Write "ok"
End If


%>