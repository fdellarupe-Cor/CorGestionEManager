<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
  nCodigo = request.querystring("Codcomp")
  sTipo   = request.querystring("Tipocomp")
  Tabla   = request.querystring("Tabla")
  Volver  = request.querystring("Volver")

  set cn  = ofv.conectar(ofv.strconn2)

  strsql  = "Select * From " & tabla & " Where codcomp = " & nCodigo
  set rs  = ofv.crearconsultaEx(StrSql,cn,1,volver)

  sHTML   = ofv.FormHeader("Agregar Estructura")

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

'-------------------------------------------------
      f1a = "asp/AgregarC.asp?Tabla=" & Tabla
      f1a = f1a & "&volver=" & volver & "&Codcomp=" & nCodigo
      f1a = f1a & "&Tipocomp=" & sTipo
  sHTML = sHTML & ofv.GenForm("AgreArchb","",f1a,"","post","","","no")

  sHTML = sHTML & ofv.GenTabla("panel","aaa","","80%","",0,0,0,0,"","","","no")

  'Encabezado estructura
  h1    = ofv.GenCelda("","tt","","20%","","","","Estructura","si")
  'Texto estructura
  t1    = ofv.GenCelda("","uti","","80%","","","",ofv.GenerarInput(rs(3).name,"","agregar","","100","dt",""),"si")
  sHTML = sHTML & ofv.GenRow("","","","","25","","",h1 & t1,"si")

  sHTML = sHTML & ofv.GenFTabla()

  'Botones Aceptar y Cancelar
''  sAccion = "edEstructura2.asp?Codcomp=" & nCodigo & "&Tipocomp=" & sTipo
  sAccion = replace(volver,"../","")
  sAccion = replace(sAccion,"_","&") & "&Codcomp=" & nCodigo & "&Tipocomp=" & sTipo
  sHTML = sHTML & ofv.AgregarBoton("AgreArchb",sAccion,"_self")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
'--------------------------------------------------------------------------------
  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------
  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>