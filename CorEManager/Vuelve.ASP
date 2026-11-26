<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

response.expires=0

sHTML = "<html><script language=JavaScript>" & chr(13)
sHTML = sHTML & "function inicio()" & chr(13)
sHTML = sHTML & "{ window.name = 'COREManager'; document.init.submit(); " & chr(13)
if ofv.explorador <> "MSIE" then
  sHTML = sHTML & "for (i = 0; i < 100; i++) {};"
end if
sHTML = sHTML & "myWin=window.close(); window.top.close();" & chr(13)
sHTML = sHTML & "}" & chr(13)
sHTML = sHTML & "</script>" & chr(13)
sHTML = sHTML & "<body style='display:none;' bgcolor=#ffffff "

sHTML = sHTML & " onload=" & chr(34) & "inicio();" & chr(34)

sHTML = sHTML & " ><head><title>COR EManager - Edicion de Esquemas Normativos</title></head>"

sHTML = sHTML & "<p >"  & Session("ASP")  & "</p>"
sHTML = sHTML & "<input type=button onclick=" & chr(34) & "inicio();" & chr(34) & " value='sigue'>"

'response.write Session("ASP")
sHTML = sHTML & "<form name=init action=" & Session("ASP")
sHTML = sHTML & " target=fesquemas  method=post>"
sHTML = sHTML & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
sHTML = sHTML & "</form>"

sHTML = sHTML & "</body></html>"

response.write sHTML

%>
