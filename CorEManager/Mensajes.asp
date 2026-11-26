<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else

response.expires=0

sMsg = ofv.convertircar(request.querystring("msg"),"_"," ")
sVolver = request.querystring("volver")
if sVolver = "" then sVolver = session("ASP")
sNombre = request.querystring("Nombre")
sTipo = request.querystring("Tipo")
nBotones = CInt(request.querystring("Botones"))

''sHTML = ofv.FormHeader("Atencion")


sVolver = replace(sVolver,"_","&")

if sTipo <> "" and sNombre <> "" then
  if right(sVolver,3) = "asp" then
    sVolver = sVolver & "?Tipo=" & sTipo & "&Nombre=" & sNombre
  else
    sVolver = sVolver & "&Tipo=" & sTipo & "&Nombre=" & sNombre
  end if
end if

 tabla = ofv.mensajes(sMsg,sVolver,nBotones)

 tablan = instr(tabla,"&Eliminar=")

 if tablan = 0 then
    tabla = replace(tabla,"asp?","asp?Eliminar=NO&")
 end if

sHTML = sHTML & Tabla & "</body></html>"

response.write sHTML
end if
%>