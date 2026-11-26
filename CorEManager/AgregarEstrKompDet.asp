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
  nID     = request.querystring("ID")
  sTipo   = request.querystring("Tipocomp")
  Tabla   = request.querystring("Tabla")
  Volver  = request.querystring("Volver")

  set cn  = ofv.conectar(ofv.strconn2)
 
  nsec    = 10

  strsql  = "SELECT max(datoformul) FROM " & tabla & " Where codestr=" & nID
  set rs  = ofv.crearconsultaEx(StrSql,cn,1,volver)
  if  not rs.eof then 
      if  isnull(rs(0)) then
          nsec = 10
      else
          nsec = clng(rs(0)) + 10
      end if
  end if
  call ofv.cerrarconsulta(rs)

  strsql = "Select * From " & tabla & " Where codestr=" & nID
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

  sHTML = ofv.FormHeader("Agregar Estructura - Detalles")

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

'-------------------------------------------------

    f1a = "asp/AgregarC.asp?Tabla=" & Tabla
    f1a = f1a & "&volver=" & volver & "&Codcomp=" & nCodigo
    f1a = f1a & "&Tipocomp=" & sTipo & "&ID=" & nID 
  sHTML = sHTML & ofv.GenForm("AgreArchb","",f1a,"","post","","","no")

  sHTML = sHTML & ofv.GenTabla("panel","","","80%","",0,0,0,0,"","","","no")

  'Encabezado sec
  h1    = ofv.GenCelda("","tt","","20%","","","","Secuencia","si")
 'Texto sec
  t1    = ofv.GenCelda("","uti","","80% ","","","",nsec & "." & ofv.GenerarInput(rs(6).name,nsec,"hidden","","45","dt",""),"si")
  sHTML = sHTML & ofv.GenRow("","","","","25","","",h1 & t1,"si")

  'Encabezado descripcion
  h1    = ofv.GenCelda("","tt","","","","","","Descripcion","si")
 'Texto Descripcion
  t1    = ofv.GenCelda("","uti","","","","","",ofv.GenerarInput(rs(4).name,"","agregar","","90","dt",""),"si")
  sHTML = sHTML & ofv.GenRow("","","","","25","","",h1 & t1,"si")

  'Encabezado Dato/Columna
  h1    = ofv.GenCelda("","tt","","","","","","Dato/Columna","si")
  'Texto Tipo de Dato/Columna
  t1    = ofv.GenCelda("","uti","","","","","",ofv.GenerarInput(rs(5).name,"","agregar","","45","dt",""),"si")
  sHTML = sHTML & ofv.GenRow("","","","","25","","",h1 & t1,"si")

  sHTML = sHTML & ofv.GenFTabla()

  'Botones Aceptar y Cancelar
  sAccion = replace(volver,"../","")
  sAccion = replace(sAccion,"_","&") & "&Codcomp=" & nCodigo & "&Tipocomp=" & sTipo
  sHTML   = sHTML & ofv.AgregarBoton("AgreArchb",sAccion,"_self")

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