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
  sID = request.querystring("ID")
  sIDt = request.querystring("IDt")
  Tabla = request.querystring("Tabla")
  Volver = request.querystring("Volver")

  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select * From " & tabla & " Where tabblk = " & sID & " and tabcontipo = '" & sIDt
  strsql = strsql & "' Order By tabconsec desc"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
  if not rs.eof then
    nSecuencia = rs("tabconsec") + 10

  else
    nSecuencia = 10

  end if

  if sIDt = "1" then
     sxtitu = "Agregar Fila de Titulos"
  else
     sxtitu = "Agregar Fila de Contenidos"
  end if

  sHTML = ofv.FormHeader(sxtitu)

    f1a = "asp/AgregarEN.asp?Tabla=" & Tabla
    f1a = f1a & "&volver=" & volver & "&Codobjeto=" & nCodigo
    f1a = f1a & "&Tipobjeto=" & sTipo 
  sHTML = sHTML & ofv.GenForm("AgreArchb","",f1a,"","post","","","no")

  sHTML = sHTML & ofv.GenTabla("panel","","","80%","",0,0,0,0,"","","","no")

  'Encabezados Descripcion y Secuencia
  h1 = ofv.GenCelda("","tt","","2%","","","","Orden","si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")

  'Texto Secuencia
  t1 = ofv.GenCelda("","","","","","","",ofv.GenerarInput(rs(4).name,nSecuencia,"Agregar","","10","dtn",""),"si")
  'Texto Descripcion
  t2 = ofv.GenCelda("","","","","","","",ofv.GenerarInput(rs(5).name,sIDt,"hidden","","","",""),"si")

  t3 = ofv.GenCelda("","","","","","","",ofv.GenerarInput(rs(3).name,sID,"hidden","","","",""),"si")

  sHTML = sHTML & ofv.GenRow("","","","","","","",t1 & t2 & t3,"si")
  sHTML = sHTML & ofv.GenFTabla()
  'Botones Aceptar y Cancelar
   saccion = mid(volver,4)
   saccion = ofv.convertircar(saccion,"_","&") & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  sHTML = sHTML & ofv.AgregarBoton("AgreArchb",sAccion,"_self")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>