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

  strsql = "Select * From " & tabla & " Where codobjeto = " & nCodigo & " Order By consecuen desc"

''response.write strsql & " strsql " & chr(13)

  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
  if not rs.eof then
    nSecuencia = clng(rs("consecuen")) + 10
  else
    nSecuencia = 10
  end if

  sHTML = ofv.FormHeader("Agregar Indice")

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

'-------------------------------------------------

  sForm = "Agregarb"

  sFuncion = ""
  sAccion  = "asp/AgregarEN.asp?Tabla=" & Tabla
  sAccion  = sAccion & "&volver=" & volver & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo

    sInput = "AGREGAR DATOS"
    Celda = ofv.gencelda("","TTT COLSPAN=2 ","","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

     Celda  = ofv.gencelda("","tt","","20%","","","","SECUENCIA","si")
     sInput = ofv.GenerarInput(rs(4).name,nSecuencia,"Agregar","","5","dtn","")
     Celda  = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")
    Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")

     Celda  = ofv.gencelda("","tt","","20%","","","","CAPITULO","si")
     sInput = ofv.GenerarInput(rs(3).name,"","Agregar","","190","dt","")
     Celda  = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")
    Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")

    sInput = ofv.GenTabla("","aaa","","80%","","0","","0","0","","",Fila,"si")

    sHTML   = sHTML & ofv.GenForm("Agregarb","",sAccion,"","post",sFuncion,sInput,"no")

  'Botones Aceptar y Cancelar
  sAccion = "edestr.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  sHTML   = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"_self")

  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"

end if

response.write sHTML

%>