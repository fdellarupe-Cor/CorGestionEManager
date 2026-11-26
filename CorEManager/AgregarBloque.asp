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
  sTipo   = request.querystring("Tipobjeto")
  sID     = request.querystring("ID")
  Tabla   = request.querystring("Tabla")
  Volver  = request.querystring("Volver")

  set cn  = ofv.conectar(ofv.strconn2)

  strsql  = "Select * From " & tabla & " Where codobjeto = " & nCodigo
  strsql  = strsql & " and tabdef = " & sID
  strsql  = strsql & " Order By tabblksec desc"
  set rs  = ofv.crearconsultaEx(StrSql,cn,1,volver)
  if not rs.eof then
    nSecuencia = clng(rs("tabblksec")) + 10
  else
    nSecuencia = 10
  end if

  sHTML = ofv.FormHeader("Agregar Subtitulos")

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
  h1 = ofv.GenCelda("","ttr","","2%","","","","Orden","si")
  h1 = h1 & ofv.GenCelda("","ttr","","2%","","","","Subtitulo","si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")

  'Texto Secuencia
  t1 = ofv.GenCelda("","","","","","","",ofv.GenerarInput(rs(4).name,nSecuencia,"Agregar","","5","dtn",""),"si")
  'Texto Descripcion
  t2 = ofv.GenCelda("","","","","","","",ofv.GenerarInput(rs(5).name," ","Agregar","","90","dt",""),"si")


   vFECHAHOY = clng(DATE) - clng(cdate("01-01-1900"))
   xxxx      = GETFECH()
   vHORAHOY  = XLHORDAT

   sInputMTM  = ""
   sInputMTM  = sInputMTM & ofv.GenerarInput("Modi","1","hidden","","10","","")
   sInputMTM  = sInputMTM & ofv.GenerarInput("TipModi","A","hidden","","10","","")
   sInputMTM  = sInputMTM & ofv.GenerarInput("MODAUTOR",SESSION("USUARIO"),"hidden","","10","","")
   sInputMTM  = sInputMTM & ofv.GenerarInput("MODFECHA",vFECHAHOY,"hidden","","10","","")
   sInputMTM  = sInputMTM & ofv.GenerarInput("MODHORA",vHORAHOY,"hidden","","10","","")


  t3 = ofv.GenCelda("","","","","","","",ofv.GenerarInput(rs(3).name,sID,"hidden","","","","") & sInputMTM,"si")

  sHTML = sHTML & ofv.GenRow("","","","","","","",t1 & t2 & t3,"si")
  sHTML = sHTML & ofv.GenFTabla()
  'Botones Aceptar y Cancelar
   saccion = mid(volver,4)
   saccion = ofv.convertircar(saccion,"_","&") & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  sHTML = sHTML & ofv.AgregarBoton("AgreArchb",sAccion,"_self")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

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