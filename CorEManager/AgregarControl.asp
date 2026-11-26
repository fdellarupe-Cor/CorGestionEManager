<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  Tabla = request.querystring("Tabla")
  Volver = request.querystring("Volver")

  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select * From " & tabla & " Where codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

  sHTML = ofv.FormHeader("Agregar Control")

  f1a = "asp/AgregarEN.asp?Tabla=" & Tabla
  f1a = f1a & "&volver=" & volver & "&Codobjeto=" & nCodigo
  f1a = f1a & "&Tipobjeto=" & sTipo
  sHTML = sHTML & ofv.GenForm("AgregarControlb","",f1a,"","post","","","no")

  sHTML = sHTML & ofv.GenTabla("panel","","","80%","",0,0,0,0,"","","","no")

  'Encabezados Sector y Frecuencia
  h1 = ofv.GenCelda("","tt","","2%","","","","Sector","si")
  h1 = h1 & ofv.GenCelda("","tt","","2%","","","","Frecuencia","si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")


  'Combo Sector
  sqleco = "Select Eco.* From Objsector, Eco WHERE "
  sqleco = sqleco & "Objsector.codsector = Eco.eco and codobjeto = " & nCodigo

  t1 = ofv.GenCelda("","","","","","","",ofv.GenerarCombo2(rs(3).name,"",sqleco,cn,0,1,""),"si")

  'Combo Frecuencia
  sqleco = "Select * From Frecuencia"
  t2 = ofv.GenCelda("","","","","","","",ofv.GenerarCombo2(rs(4).name,"",sqleco,cn,0,0,""),"si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",t1 & t2,"si")
  sHTML = sHTML & ofv.GenFTabla()
  'Botones Aceptar y Cancelar
  sAccion = "edControles.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  sHTML = sHTML & ofv.AgregarBoton("AgregarControlb",sAccion,"_self")

  sHTML = sHTML & "</body></html>"

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
end if

response.write sHTML

%>