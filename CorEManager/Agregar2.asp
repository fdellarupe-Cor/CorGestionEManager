<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  sNombre = request.querystring("Nombre")
  sTipo = request.querystring("Tipo")

  sHTML = ofv.FormHeader("Agregar " & sNombre)

    f1a = "asp/AgregarComplemento.asp?Tipo="
    f1a = f1a & sTipo & "&Nombre=" & sNombre
  sHTML = sHTML & ofv.GenForm("Agregar2b","",f1a,"","post","","","no")

  sHTML = sHTML & ofv.GenTabla("panel","","","80%","",0,0,0,0,"","","","no")

  'Encabezado Descripcion
   t1 = ofv.GenCelda("","tt","","2%","","","","Descripcion","si")
  'Texto Descripcion
   t1 = t1 & ofv.GenCelda("","","","","","","",ofv.GenerarInput("Descri","","Agregar","","60","dt",""),"si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",T1,"si")

   t2 = ofv.GenCelda("","","","","","","","&nbsp;","si")
   t2 = t2 & ofv.GenCelda("","","","","","","","&nbsp;","si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",T2,"si")

  'Encabezado Identificacion
   t3 = ofv.GenCelda("","tt","","2%","","","","Identificacion","si")
  'Texto Identificacion
   t3 = t3 & ofv.GenCelda("","","","","","","",ofv.GenerarInput("Codiden","","agregar","","40","dt",""),"si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",T3,"si")

  sHTML = sHTML & ofv.GenFTabla()

  'Botones Aceptar y Cancelar
  sAccion = "Complementos.asp?Nombre=" & sNombre & "&Tipo=" & sTipo
  sHTML = sHTML & ofv.AgregarBoton("Agregar2b",sAccion,"")

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>