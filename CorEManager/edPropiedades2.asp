<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  ofv.ObtenerAtributos Session("Form"),""
  bModificar = ofv.AModificar

  nCodigo = request.querystring("Codcomp")
  sTipo = request.querystring("Tipocomp")
  vuelta = ofv.vueltaasp2("../")

  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select descri,archivo From Tipocomp Where tipo = '" & sTipo & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    sComplemento = rs(0)
    sImagen = rs(1)
  end if
  call ofv.cerrarconsulta(rs)

  Strsql = "Select * From Complementos Where codcomp=" & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then sNombre = rs(1) & "  -  " & rs(3)
  sNombre = ofv.convertircar(sNombre," ","&nbsp;")
  call ofv.cerrarconsulta(rs)

  Strsql = "SELECT * FROM COMPPRTY"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  Campo3 = rs(3).name
  Campo4 = rs(4).name
  call ofv.cerrarconsulta(rs)

  sHTML = ofv.FormHeader("Edicion de " & sComplemento & " - Propiedades")

  'Encabezado
  sImagen = ofv.GenImage("","","","","Iconos/" & sImagen,"0","")
  sHref = "EditarImagen2.asp?ID=" & nCodigo
  sLink = ofv.GenLink("","",sHref,"_blank","","",sImagen,"si")
  Celda = ofv.GenCelda("","","","","","","",sLink,"si")
  Celda = Celda & ofv.GenCelda("","dt","","100%","","","",sNombre,"si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  Tabla = ofv.GenTabla("","","","80%","","0","","0","0","","",Fila,"si")

  Strsql = "SELECT * FROM COMPPRTY Where codcomp=" & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  nID = 0
  sASP = "asp/AgregarC.asp?"
  Valor3 = ""
  Valor4 = ""
  if not rs.eof Then
    nID = rs(0)
    sASP = "asp/ActualizarC.asp?"
    Valor3 = rs(3)
    Valor4 = rs(4)
  end if

  'Tipo de Generacion
  Celda = ofv.GenCelda("","tt","","","","","","Tipo de Generacion","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  if bModificar then
    Strsql = "Select * From Tipogen Order By tipogen"
    sInput = ofv.GenerarCombo("edPropiedades2",Campo3,Valor3,Strsql,cn,0,0,"sb")
  else
    sInput = ofv.GenerarInput("",Valor3,"readonly","","60","dt","")
  end if
  Celda = ofv.GenCelda("","","","","","","",sInput,"si")
  Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

  Celda = ofv.GenCelda("","abc colspan=2","","","","","","","si")
  Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

  'Sistema Generador
  Celda = ofv.GenCelda("","tt","","","","","","Sistema Generador","si")
  Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")
  if bModificar then
    Strsql = "Select * From Sisgen where estado=3 Order By sisgen"
    sInput = ofv.GenerarCombo("edPropiedades2",Campo4,Valor4,Strsql,cn,0,0,"sb")
  else
    sInput = ofv.GenerarInput("",Valor4,"readonly","","60","dt","")
  end if
  Celda = ofv.GenCelda("","","","","","","",sInput,"si")
  Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

  sAccion = ""
  if bModificar then
    sAccion = sASP & "ID=" & nID & "&Tabla=Compprty&volver=" & vuelta & "&Codcomp=" & nCodigo
    sAccion = sAccion & "&Tipocomp=" & sTipo
  end if
  Formulario = ofv.GenForm("edPropiedades2","",sAccion,"","post","",Fila,"si")

  sHTML = sHTML & Tabla & "<br><br><br>"
  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>