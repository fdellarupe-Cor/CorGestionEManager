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

  strsql = "Select * From " & tabla & " Where codobjeto = " & nCodigo & " Order By secuen desc"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
  if not rs.eof then
    nSecuencia = cint(rs("secuen")) + 10
  else
    nSecuencia = 10
  end if

  sHTML = ofv.FormHeader("Agregar Evento")

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

'-------------------------------------------------



  sForm = "Agregarb"
  sFormX = "AregarX"
    sInput = "AGREGAR EVENTO" '& session("formn")
    Celda = ofv.gencelda("","TTT COLSPAN=2 ","","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


     Celda  = ofv.gencelda("","tt","","20%","","","","DESCRIPCION","si")

     sInput = ofv.GenerarInput(rs(3).name,"NUEVO EVENTO","text",sFormX,"200","dt maxlength=70","")
     sInput = REPLACE(sInput,"{ document." & sFormX  & ".submit(); return true;}","{ if (this.value.length > 70) {alert('El texto es demasiado extenso'); return false;} else { document." & sFormX & ".submit(); return true;}}")
     Celda  = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")

     Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")

     Celda  = ofv.gencelda("","tt","","20%","","","","Secuencia","si")
     sInput = ofv.GenerarInput(rs(4).name,nSecuencia,"Agregar","","10","dtn","")
     Celda  = Celda & ofv.gencelda("","vti","","80%","","","",sInput,"si")
    Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")


  sFuncion = "onSubmit=" & chr(34) & " if (this." & rs(3).name & ".value.length==0) "
  sFuncion = sFuncion & "{alert('El campo Descripcion no debe ser nulo');"
  sFuncion = sFuncion & "return false;}"
  sFuncion = sFuncion & " else {if (this." & rs(4).name & ".value.length==0)"
  sFuncion = sFuncion & " {this." & rs(4).name & ".value = 0;} return true;}" & chr(34)

  sAccion = "asp/AgregarEN.asp?Tabla=" & Tabla
  sAccion = sAccion & "&volver=" & volver & "&Codobjeto=" & nCodigo
  sAccion = sAccion & "&Tipobjeto=" & sTipo


  Fila    = ofv.GenTabla("","aaa","","80%","","0","","0","0","","",fila,"si")

 sHTML = sHTML & ofv.GenForm("Agregarb","",sAccion,"","post",sFuncion,Fila,"si")


  'Aceptar y Cancelar
  sAccion = "edEventos.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"")

  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"

end if

response.write sHTML

%>