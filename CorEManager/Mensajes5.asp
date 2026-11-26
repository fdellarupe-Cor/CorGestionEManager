<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

response.expires=0

sMsg = ofv.convertircar(request.querystring("msg"),"_"," ")
sVolver = request.querystring("volver")
sNombre = request.querystring("Nombre")
sTipo = request.querystring("Tipo")
nBotones = CInt(request.querystring("Botones"))
sHTML = ofv.FormHeader("Atencion")

if sTipo <> "" and sNombre <> "" then
  if right(sVolver,3) = "asp" then
    sVolver = sVolver & "?Tipo=" & sTipo & "&Nombre=" & sNombre
  else
    sVolver = sVolver & "&Tipo=" & sTipo & "&Nombre=" & sNombre
  end if
end if

sHTML = sHTML & "<table width=100% ><tr><td align=center>"
sHTML = sHTML & "<font color=#336699 face=verdana size=1>"
sHTML = sHTML & "<b>" & sMsg & "</b></font></td></tr></table><hr size=1 color=#000099>"

sHTML = sHTML & "<table width=100% ><tr>"
if nBotones = 1 then
  sHTML = sHTML & "<form name=Mensajes id=Mensajes action=" & sVolver & " method=post>"
  sHTML = sHTML & "<td width=50% align=center><input type=submit value=Volver "
  sHTML = sHTML & "onclick=window.close()></td></form>"
else
  sHTML = sHTML & "<form action=" & sVolver & " method=post>"
  sHTML = sHTML & "<td width=40% align=center><input type=submit name=Accion "
  sHTML = sHTML & "Value=Eliminar></td></form>"
  sHTML = sHTML & "<form action=" & sVolver & " method=post>"
  sHTML = sHTML & "<td width=40% align=center><input type=submit name=Accion "
  sHTML = sHTML & "Value=Cancelar></td></form>"
end if
sHTML = sHTML & "</tr></table></BODY></html>"

response.write sHTML

%>