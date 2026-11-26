<%
'Session.LCID=3082
'session.lcid=1034
'session.lcid=1052
'response.write session.lcid & "<br>"
'response.write Request.ServerVariables("HTTP_ACCEPT_LANGUAGE") & "<br>"
response.write clng(date -clng(cdate("1-1-1900")) )  & "<br>"
'SetLCID
'response.write session.lcid & "<br>"
'45385
'xfecha=45200 '03/10/2023
'xfecha=45385  '3/4/2024
xfecha=45734
'response.write cdate(xfecha + clng(cdate("1-1-1900"))) & "---<br>"

response.write GTFechadorsh(xfecha)


Function GTFechadorsh(GTWfecha)

	If CStr(GTWfecha & "") = "0" Or CStr(GTWfecha & "") = "" Then

		GTFechadorSH  = "-----"

	Else


		GTwfec = CDate(GTWfecha + CLng(CDate("01-01-1900")))

		GTWFDIAB = CStr(Day(GTwfec))
		GTWFMESB = CStr(Month(GTwfec))

		If Len(GTWFDIAB) = 1 Then GTWFDIAB = "0" & GTWFDIAB
		If Len(GTWFMESB) = 1 Then GTWFMESB = "0" & GTWFMESB

		GTwFechash = GTWFDIAB & "/" & GTWFMESB & "/" & Year(GTwfec)
		GTFechadorSH  = UCase(GTwFechash)

	End If

End Function

%>