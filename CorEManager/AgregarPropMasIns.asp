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

  sHTML = ofv.FormHeader("Agregar origenes de documento adicionales")
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
  h1 = ofv.GenCelda("","tt","","2%","","","","Secuencia","si")
  h1 = h1 & ofv.GenCelda("","tt","","2%","","","","origen","si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")

  'Secuencia

  t1 = ofv.GenCelda("","","","","","","",ofv.GenerarInput(rs(6).name,nSecuencia,"Agregar","","10","dtn",""),"si")

  Sql1 = "Select * From Origen Order By descri"


  'Origen del documento
  Celda = ofv.gencelda("","","","","","","","","si")

 
  sInput = ofv.GenerarCombox2(sql1,cn,0,1,rs(3).name) 
  t2 =  ofv.gencelda("","","","","","","",sInput,"si")


  sHTML = sHTML & ofv.GenRow("","","","","","","",t1 & t2,"si")
  sHTML = sHTML & ofv.GenFTabla()
  'Botones Aceptar y Cancelar
  sAccion = "edPropiedadesMasIns.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  sHTML = sHTML & ofv.AgregarBoton("AgreArchb",sAccion,"_self")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  Agregar = ""
  Buscar  = ""
  Celda   = ""
 '--------------------------------------------------------------------------------
   sHTML = sHTML & "</div>"



     Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
     Fila = ofv.GenRow("","","","","","","",Celda,"si")

     sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")


     'Estructura del Indice

     sHTML = sHTML & CorPnlIndx("N","")

     'sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)


'--------------------------------------------------------------------------------
  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>