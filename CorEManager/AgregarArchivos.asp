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

  strsql = "Select * From " & tabla & " Where codobjeto = " & nCodigo & "  AND TIPO = " & SESSION("SKIND") & " Order By secuen desc"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
  if not rs.eof then
    nSecuencia = clng(rs("secuen")) + 10
  else
    nSecuencia = 10
  end if

  sHTML = ofv.FormHeader("Agregar Documento")

  '----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

'-------------------------------------------------

    f1a = "asp/AgregarEN.asp?Tabla=" & Tabla
    f1a = f1a & "&volver=" & volver & "&Codobjeto=" & nCodigo
    f1a = f1a & "&Tipobjeto=" & sTipo
  sHTML = sHTML & ofv.GenForm("AgreArchb","",f1a,"","post","","","no")

  sHTML = sHTML & ofv.GenTabla("panel","aaa","","40%","",0,0,0,0,"","","","no")

sFormX = "AgregarX"
  'Encabezados Descripcion y Secuencia
  h1 = ofv.GenCelda("","tt","","1%","","","","Secuencia","si")
  h1 = h1 & ofv.GenCelda("","tt","","1%","","","","Descripcion","si")
  sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")

  'Texto Secuencia
  t1 = ofv.GenCelda("","","","","","","",ofv.GenerarInput(rs(3).name,nSecuencia,"Agregar","","47","dtn",""),"si")
  'Texto Descripcion
  sInput = ofv.GenerarInput(rs(4).name,"","text",sFormX,"200","dt maxlength=70","") & ofv.GenerarInput(rs(7).name,SESSION("SKIND"),"HIDDEN",sFormX,"1","dt","")
  sInput = REPLACE(sInput,"{ document." & sFormX  & ".submit(); return true;}","{ if (this.value.length > 70) {alert('El texto es demasiado extenso'); return false;} else { document." & sFormX & ".submit(); return true;}}")
  t2 = ofv.GenCelda("","","","","","","",sInput,"si")

  sHTML = sHTML & ofv.GenRow("","","","","","","",t1 & t2,"")
  sHTML = sHTML & ofv.GenFTabla()
  'Botones Aceptar y Cancelar
  sAccion = "edArchivos.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  sHTML = sHTML & ofv.AgregarBoton("AgreArchb",sAccion,"_self")

  '--------------------------------------------------------------------------------
  sHTML = sHTML & "</div>"

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