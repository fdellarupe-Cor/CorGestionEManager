<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

response.expires=0

sHTML = "<html>" & chr(13)

sHTML = sHTML & "<body bgcolor=#ffffff "

sHTML = sHTML & " onload=" & chr(34) & "document.init.submit();" & chr(34)

sHTML = sHTML & " ><head><title>COR EManager - Edicion de Esquemas Normativos</title></head>"

sHTML = sHTML & "<p >"  & Session("ASP")  & "</p>"
sHTML = sHTML & "<input type=button onclick=" & chr(34) & "document.init.submit();" & chr(34) & " value='sigue'>"


sHTML = sHTML & "<form name=init action=" & Session("ASPSTR")
sHTML = sHTML & " target=info method=post>"
sHTML = sHTML & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
sHTML = sHTML & "</form>"

sHTML = sHTML & "</body></html>"

response.write sHTML

%>