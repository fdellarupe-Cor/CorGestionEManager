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
  sTipo   = request.querystring("Tipo")

  sHTML   = ofv.FormHeader("Agregar " & sNombre)

  f1a   = "asp/AgregarComplementocess.asp?Tipo="
  f1a   = f1a & sTipo & "&Nombre=" & sNombre
  sHTML = sHTML & ofv.GenForm("Agregar2b","",f1a,"","post","","","no")

  sHTML = sHTML & ofv.GenTabla("panel","aaa","","80%","",0,0,0,0,"","","","no")

  'Encabezado Descripcion
   t1   = ofv.GenCelda("","tt","","20%","","","","Descripcion","si")
  'Texto Descripcion
   t1   = t1 & ofv.GenCelda("","vti","","80%","","","",ofv.GenerarInput("Descri","","Agregar","","60","dt",""),"si")
  sHTML = sHTML & ofv.GenRow("","","","","25","","",T1,"si")


  'Encabezado Identificacion
   t3   = ofv.GenCelda("","tt","","","","","","Identificacion","si")
  'Texto Identificacion
   t3   = t3 & ofv.GenCelda("","vti","","","","","",ofv.GenerarInput("Codiden","","agregar","","40","dt",""),"si")
  sHTML = sHTML & ofv.GenRow("","","","","25","","",T3,"si")

  'codigo de empresa
   t3   = ofv.GenCelda("","","","","","","","&nbsp;","si")
   t3   = t3 & ofv.GenCelda("","","","","","","",ofv.GenerarInput("COMPUN",SESSION("UN"),"HIDDEN","","40","dt",""),"si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",T3,"si")

  sHTML = sHTML & ofv.GenFTabla()

  'Botones Aceptar y Cancelar
  sAccion = "CompTree.asp"
  sHTML   = sHTML & ofv.AgregarBoton("Agregar2b",sAccion,"_self")
  sHTML   = sHTML & "</body></html>"

end if

Response.Write sHTML

%>