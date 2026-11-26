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
  nAccion = request.querystring("Accion")
  sTipo = request.querystring("Tipobjeto")
  Tabla = request.querystring("Tabla")
  Volver = request.querystring("Volver")

  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select * From " & tabla & " Where codproced = '" & nProced & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

  sHTML = ofv.FormHeader("Agregar Accion")

  shtml = shtml & CorPnlFnc("N","")
  sHTML = sHTML & "<DIV ID='DCEX1' style='display:block;' >"
  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:90%;overflow:auto;" & styl & "' ><span id=111>"

  sForm = "Agregarb"

    sInput = "AGREGAR pasos al circuito alternativo"
    Celda = ofv.gencelda("","TTT COLSPAN=2 ","","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


     Celda  = ofv.gencelda("","tt","","20%","","","","Paso","si")
''''     Strsql = "Select * From Tipoaccalt WHERE modo<>'CN'  Order By descri"

     Strsql = "Select a.* From Tipoaccalt a, modo b WHERE a.estado=3 and a.modo = b.tipomodo and  b.control = 0  Order By a.descri"
     sInput = ofv.GenerarCombo2(rs(2).name,"",Strsql,cn,0,1,"no")
     sInput = sInput & ofv.GenerarInput("Modi","1","hidden","","10","","")
     sInput = sInput & ofv.GenerarInput("TipModi","A","hidden","","10","","")



  sFuncion = ""

  sAccion = "EdProcedAlt.asp?Codproced=" & nProced & "&Codobjeto=" & nCodigo
  sAccion = sAccion & "&Accion=" & nAccion & "&Tipobjeto=" & sTipo & "&Tabla=" & Tabla
  sAccion = sAccion & "&Volver=" & Volver

  sAccionv1 = sAccion

  sAccionV = ofv.convertircar(sAccion,"&","_")
  sAccionV = "../" & sAccionV

  sAccion = "asp/AgregarEN.asp?Tabla=" & Tabla
  sAccion = sAccion & "&volver=" & sAccionV & "&Codobjeto=" & nCodigo
  sAccion = sAccion & "&Tipobjeto=" & sTipo & "&Codproced=" & nProced
  sAccion = sAccion & "&Accion=" & nAccion

  Formulario = Formulario & ofv.GenForm("Agregarb","",sAccion,"","post",sFuncion,sInput,"no")
   Celda  = Celda & ofv.gencelda("","vti","","80%","","","",Formulario,"si")
    Fila    = Fila & ofv.genrow("","","","","","","",Celda,"si")

  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Fila,"si")


  'Aceptar y Cancelar
  sAccion = sAccionv1
  sHTML = sHTML & ofv.AgregarBoton("Agregarb",sAccion,"")



  sHTML = sHTML & "</span></div>"
  sHTML = sHTML & "</div>"
  Celda = ofv.GenCelda("","ut","","","","","","&nbsp;","si")
  Filann = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%   ","","0","","1","0","","",Filann,"si")
  'Estructura del Indice
  sHTML = sHTML & CorPnlIndx("N","")


  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)



  sHTML = sHTML & "</body></html>"

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
end if

response.write sHTML

%>