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

  sHTML = ofv.FormHeader("Agregar Procedimiento")

  f1a = "asp/AgregarEN.asp?Tabla=" & Tabla
  f1a = f1a & "&volver=" & volver & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  sHTML = sHTML & ofv.GenForm("AgreArchb","",f1a,"","post","","","no")

  sHTML = sHTML & ofv.GenTabla("panel","","","80%","",0,0,0,0,"","","","no")

  'Encabezados
  h1 = ofv.GenCelda("","tt","","2%","","","","Evento","si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")

  'Combo Evento
  Strsql="Select * From Objeven Where codobjeto = " & nCodigo & " Order By secuen"
  t1 = ofv.GenCelda("","","","","","","",ofv.GenerarCombo2(rs(3).name,"",Strsql,cn,0,3,"Evento"),"si")
  sukey = GTkey(nCodigo)
  sInput1 = ofv.GenerarInput("codproced",sukey,"hidden","","10","","")

  sHTML = sHTML & ofv.GenRow("","","","","","","",t1 & sInput1,"si")
  sHTML = sHTML & ofv.GenFTabla()

  'Botones Aceptar y Cancelar
  sAccion = "edProcedimientos.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  sHTML = sHTML & ofv.AgregarBoton("AgreArchb",sAccion,"_self")

  sHTML = sHTML & "</body></html>"

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
end if

response.write sHTML

%>