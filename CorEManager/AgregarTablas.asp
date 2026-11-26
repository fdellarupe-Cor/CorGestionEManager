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

  strsql = "Select * From " & tabla & " Where codobjeto = " & nCodigo & " Order By tabsecuen desc"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
  if not rs.eof then
    nSecuencia = cint(rs("tabsecuen")) + 10
  else
    nSecuencia = 10
  end if


  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

'-------------------------------------------------

    f1a = "asp/AgregarEN.asp?Tabla=" & Tabla
    f1a = f1a & "&volver=" & volver & "&Codobjeto=" & nCodigo
    f1a = f1a & "&Tipobjeto=" & sTipo 
  sHTML = sHTML & ofv.GenForm("AgreArchb","",f1a,"","post","","","no")

  sHTML = sHTML & ofv.GenTabla("panel","","","80%","",0,0,0,0,"","","","no")

  'Encabezados Descripcion y Secuencia
  h1 = ofv.GenCelda("","tt","","2%","","","","Orden","si")
  h1 = h1 & ofv.GenCelda("","tt","","2%","","","","Titulo","si")
  h1 = h1 & ofv.GenCelda("","tt","","2%","","","","Modelo","si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")

  'Texto Secuencia
  t1 = ofv.GenCelda("","","","","","","",ofv.GenerarInput(rs(5).name,nSecuencia,"Agregar","","10","dtn",""),"si")
  'Texto Titulo
  t2 = ofv.GenCelda("","","","","","","",ofv.GenerarInput(rs(4).name,"","Agregar","","100","dt",""),"si")
  'combo Perfil
  Strsql = "Select * From Tabperfil Order By tabdescri"
  sInput = ofv.GenerarCombox3(Strsql,cn,0,1,rs(3).name)

   vFECHAHOY = clng(DATE) - clng(cdate("01-01-1900"))
   xxxx      = GETFECH()
   vHORAHOY  = XLHORDAT

   sInputMTM  = ""
   sInputMTM  = sInputMTM & ofv.GenerarInput("Modi","1","hidden","","10","","")
   sInputMTM  = sInputMTM & ofv.GenerarInput("TipModi","A","hidden","","10","","")
   sInputMTM  = sInputMTM & ofv.GenerarInput("MODAUTOR",SESSION("USUARIO"),"hidden","","10","","")
   sInputMTM  = sInputMTM & ofv.GenerarInput("MODFECHA",vFECHAHOY,"hidden","","10","","")
   sInputMTM  = sInputMTM & ofv.GenerarInput("MODHORA",vHORAHOY,"hidden","","10","","")

  t3 =  ofv.GenCelda("","","","","","","",sInput & sInputMTM,"si")

  sHTML = sHTML & ofv.GenRow("","","","","","","",t1 & t2 & t3,"si")
  sHTML = sHTML & ofv.GenFTabla()
  'Botones Aceptar y Cancelar
  sAccion = "edTablas.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  sHTML = sHTML & ofv.AgregarBoton("AgreArchb",sAccion,"_self")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)


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