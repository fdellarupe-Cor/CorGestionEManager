<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
  direct = request.querystring("direct")
  nomdir = request.querystring("nomdir")
  rutdir = request.querystring("rutdir")



  sHTML = ofv.FormHeader("Agregar Subdirectorio a " & nomdir)

  f1a = "Cordscr2.asp?direct=" & direct & "&nomdir=" & nomdir & "&rutdir=" & rutdir & "&OPC=SI"
  sHTML = sHTML & ofv.GenForm("Agregarb","",f1a,"","post","","","no")

  sHTML = sHTML & ofv.GenTabla("panel","aaa","","80%","",0,0,0,0,"","","","no")

  'Encabezados Sector y Frecuencia
  h1 = ofv.GenCelda("","tt","","","","","","Nombre Subdirectorio","si")

  sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")


  t1 = ofv.GenCelda("","","","","","","",ofv.GenerarInput("newdir","","Agregar","","140","dt",""),"si")


  sHTML = sHTML & ofv.GenRow("","","","","","","",t1,"si")
  sHTML = sHTML & ofv.GenFTabla()
  'Botones Aceptar y Cancelar
  sAccion = "Cordscr2.asp?direct=" & direct & "&nomdir=" & nomdir & "&rutdir=" & rutdir & "&OPC=NO"
  sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"_self")

  sHTML = sHTML & "</body></html>"

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
end if

response.write sHTML

%>