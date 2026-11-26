<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  ID = request.querystring("ID")

  set cn = ofv.conectar(ofv.strconn2)

  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

'-------------------------------------------------

  Strsql = "Select * From Complementos Where codcomp = " & ID
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     sNombre = rs(3)
     sNombre = ofv.convertircar(sNombre," ","&nbsp;")
     sTipo   = rs(2)
     siden   = rs(1)
  end if
  call ofv.cerrarconsulta(rs)


  strsql = "Select descri,archivo From Tipocomp Where tipo = '" & sTipo & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     sComplemento = rs(0)
     sImagen = rs(1)
  end if
  call ofv.cerrarconsulta(rs)

  Fila    = ""

  'Encabezado

  Celda  = ofv.GenCelda("","tt","","20%","","","","Identificacion","si")

  Celda   = Celda & ofv.GenCelda("","uti","","100%","","","",siden,"si")
  Fila  = Fila & ofv.GenRow("","","","","25","","",Celda,"si")


  Celda  = ofv.GenCelda("","tt","","20%","","","","Descripcion","si")

  sImagen = ofv.GenImage("","","","","Iconos/" & sImagen,"0","")
  sLink   = sImagen

  Celda   = Celda & ofv.GenCelda("","uti","","100%","","","",sNombre,"si")
  Fila  = Fila & ofv.GenRow("","","","","25","","",Celda,"si")


  Strsql = "SELECT * FROM COMPPRTY Where codcomp = " & ID
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  nID    = 0
  Valor3 = ""
  Valor4 = ""
  if not rs.eof Then
     nID    = rs(0)
     Valor3 = rs(3)
     Valor4 = rs(4)
  end if


  'Tipo de Generacion
  Celda  = ofv.GenCelda("","tt","","30%","","","","Tipo&nbsp;de&nbsp;Generacion","si")
  Strsql = "Select tipogen From Tipogen where tipogen = '" & Valor3 & "' "
  set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs2.eof then
     sInput = rs2(0)
  else
     sInput = "-----"
  end if
  Celda = Celda & ofv.GenCelda("","uti","","70%","","","",sInput,"si")
  Fila  = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

  'Sistema Generador
  Celda = ofv.GenCelda("","tt","","","","","","Sistema&nbsp;Generador","si")
  Strsql = "Select sisgen From Sisgen where Sisgen = '" & Valor4 & "' "
  set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs2.eof then
     sInput = rs2(0)
  else
     sInput = "-----"
  end if
  Celda = Celda & ofv.GenCelda("","uti","","","","","",sInput,"si")
  Fila  = Fila & ofv.GenRow("","","","","25","","",Celda,"si")


''  sHTML = sHTML & "<br>" & Tabla
  sHTML = sHTML & ofv.GenTabla("","aaa","","100%","","0","","3","0","","",Fila,"si")

'--------------------------------------------------------------------------------
  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  ''sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------
 sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconsulta(rs2)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"

end if

Response.Write sHTML

%>