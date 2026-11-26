<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  error=request.form("texterror")
  sHTML = ofv.FormHeader(session("FormN"))



  'Encabezados
  Formulario = ofv.Encabezados("error")

 Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
  Celda = Celda & ofv.GenCelda("","vt","","80%","","","",error,"si")
 Fila = ofv.genrow("","","","","","","",Celda,"si")
   sHTML = sHTML & ofv.GenTabla("","","","80%","","0","0","0","0","","",Formulario & Fila,"si")
  sHTML = sHTML & "</body></html>"
end if

Response.Write SHTML

%>