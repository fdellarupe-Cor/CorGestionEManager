<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  nProced = request.querystring("Codproced")
  nCodigo = request.querystring("Codobjeto")
  sEven = request.querystring("Codeven")
  sTipo = request.querystring("Tipobjeto")
  Tabla = request.querystring("Tabla")
  Volver = request.querystring("Volver")

  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select * From " & tabla & " Where codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

  sHTML = ofv.FormHeader("Agregar Accion")

    f1a = "asp/AgregarEN.asp?Tabla=" & Tabla
    f1a = f1a & "&volver=" & volver & "&Codobjeto=" & nCodigo
    f1a = f1a & "&Tipobjeto=" & sTipo & "&Codproced=" & nProced
    f1a = f1a & "&Codeven=" & sEven
  sHTML = sHTML & ofv.GenForm("AgregarAccionb","",f1a,"","post","","","no")

  sHTML = sHTML & ofv.GenTabla("panel","","","80%","",0,0,0,0,"","","","no")

  'Encabezado Accion
  h1 = ofv.GenCelda("","tt","","2%","","","","Accion del Procedimiento","si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")

  'Combo Acciones
  Strsql="Select * From Tipoaccalt where estado=3 Order By descri"
  t1 = ofv.GenCelda("","","","","","","",ofv.GenerarCombo2(rs(4).name,"",Strsql,cn,0,1,"Accion"),"si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",t1,"si")

  sHTML = sHTML & ofv.GenFTabla()

  'Botones Aceptar y Cancelar
  sAccion = "edProcedimientos.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  sHTML = sHTML & ofv.AgregarBoton("AgregarAccionb",sAccion,"")

  sHTML = sHTML & "</body></html>"

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
end if

set ofv = nothing

response.write sHTML

%>