<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  if session("Nomodificar") = "S" then bModificar = False

  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  sTabla = request.querystring("Tabla")
  nID = request.querystring("ID")
  volver = request.querystring("volver")

  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

'-------------------------------------------------

  sHTML = sHTML & "<center>"

  sVuelve = ofv.convertircar(volver,"_","&")
  sVuelve = ofv.convertircar(sVuelve,"*","&")
  sVuelta = ofv.convertircar(sVuelve,"&","*")
  sVuelve = mid(sVuelve,4,len(sVuelve) - 3)

  esttxt = " STYLE={font-family:tahoma;color:#225588;font-size:8pt;background-color:#f7eedd;} "

  set cn = ofv.conectar(ofv.strconn2)


     sdato = "Referencia"

  scontdesc = "SD"

  Strsql = "SELECT * from cscontrolsel "
  Strsql = Strsql & "Where objid = " & nID
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
  if not rs.eof then
     scontdesc = rs(1)
  end if
  call ofv.cerrarconsulta(rs)


  strsql = "SELECT " & sdato & " From " & sTabla & " Where objid = " & nID

  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
  if not rs.eof then
    sAccion = ""
    if bModificar then
      'Encabezado Actualizar
      sAccion = "asp/ActualizarEN.asp?Tabla=" & sTabla & "&volver=" & volver
      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&ID="
      sAccion = sAccion & nID & "&Referencia=SI"
    end if
    sHTML = sHTML & ofv.GenForm("Actualizar","",sAccion,"","post","","","no")

    sHTML = sHTML & "<br>" & ofv.GenTabla("","","","60%","","0","0","0","0","","","","no")

    scontdesc = "<B>Control: " & scontdesc & "</b>"

    'Encabezado control
 '   Celda = ofv.GenCelda("","ut","center","100%","","","",scontdesc,"si")
 '   sHTML = sHTML & ofv.GenRow("","","","","","","",Celda,"si")

    'Encabezado Referencia
    Celda = ofv.GenCelda("","ttt","","100%","","","","Texto de Referencia" & scontdesc,"si")
    sHTML = sHTML & ofv.GenRow("","","","","25","","",Celda,"si")

    'Texto de Referencia
    celda = "<font class=at><textarea  name=" & sdato & " rows=20 cols=100 "
      if bModificar then
        celda = celda & " onkeypress=" & chr(34)
        celda = celda & " if (this.value.length > 1950) {alert('El texto es demasiado"
        celda = celda & " extenso'); return false;}" & chr(34)
      else
        celda = celda & " readonly"
      end if

    celda = celda & " >" & rs(0) & "</textarea></font>"

    Celda = ofv.GenCelda("","ut","center","","","","",celda,"si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenFTabla()

    sHTML = sHTML & ofv.GenTabla("","","","60%","","0","0","0","0","","","","no")
    Celda = ""
    if bModificar then
      'Encabezado Actualizar
      sInput = ofv.GenerarInput("boton","Actualizar","submit","","20","bt","")
      Celda = ofv.GenCelda("","","right","50%","","","",sInput,"si")
    end if
  end if
  Celda = Celda & ofv.GenFForm()

  'Encabezado volver
  sInput = ofv.GenerarInput("boton2","Volver","submit","","20","bt","")
  Celdan = ofv.GenCelda("","","left","50%","","","",sInput,"si")
  Form = ofv.GenForm("Retornar","",sVuelve & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo,"","post","",Celdan,"si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",Celda & Form,"si")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & ofv.GenFTabla() & "</center>"


'--------------------------------------------------------------------------------
  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------


  sHTML = sHTML & "</body></html>"


end if

response.write sHTML

%>