<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  if session("Nomodificar") = "S" then bModificar = False

  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  vuelta = ofv.vueltaasp2("../")

  set cn = ofv.conectar(ofv.strconn2)

    sHTML = ofv.FormBody

    'Descripcion
    Celda = ofv.GenCelda("","tt colspan=2 ","","","","","","Titulos","si")
    Fila = ofv.GenRow("","","","","","","",Celda,"si")

    Celda = ofv.GenCelda("","abc colspan=2","","","","","","&nbsp;","si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    Formulario = formulario & Fila

  Strsql = "Select objid,titulo From ObjMain where CodObjeto=" & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    'Titulo bloque 1
    Celda = ofv.GenCelda("","tt","","","","","","bloque 1: ","si")
    if bModificar then
      sInput = ofv.GenerarInput(rs(1).name,rs(1),"text","edGeneral1","80","dt","")
    else
      sInput = ofv.GenerarInput("",rs(1),"readonly","","90","dt","")
    end if
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")
    Fila = ofv.GenRow("","","","","","","",Celda,"si")

    Celda = ofv.GenCelda("","abc colspan=2","","","","","","&nbsp;","si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=ObjMain&volver=" & vuelta
      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
    end if
    Formulario = formulario & ofv.GenForm("edGeneral1","",sAccion,"","post","",Fila,"si")


  end if 
  call ofv.cerrarconsulta(rs)

  Strsql = "Select objid,titulo From ObjAlcance where CodObjeto=" & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    'Titulo bloque 2
    Celda = ofv.GenCelda("","tt","","","","","","bloque 2: ","si")
    if bModificar then
      sInput = ofv.GenerarInput(rs(1).name,rs(1),"text","edGeneral2","80","dt","")
    else
      sInput = ofv.GenerarInput("",rs(1),"readonly","","90","dt","")
    end if
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")
    Fila = ofv.GenRow("","","","","","","",Celda,"si")

    Celda = ofv.GenCelda("","abc colspan=2","","","","","","&nbsp;","si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=ObjAlcance&volver=" & vuelta
      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
    end if
    Formulario = formulario & ofv.GenForm("edGeneral2","",sAccion,"","post","",Fila,"si")


  end if 
  call ofv.cerrarconsulta(rs)

  Strsql = "Select objid,titulo From ObjConsid where CodObjeto=" & nCodigo 
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

    'Titulo bloque 3
    Celda = ofv.GenCelda("","tt","","","","","","bloque 3: ","si")
    if bModificar then
      sInput = ofv.GenerarInput(rs(1).name,rs(1),"text","edGeneral3","80","dt","")
    else
      sInput = ofv.GenerarInput("",rs(1),"readonly","","90","dt","")
    end if
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")
    Fila = ofv.GenRow("","","","","","","",Celda,"si")

    Celda = ofv.GenCelda("","abc colspan=2","","","","","","&nbsp;","si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=ObjConsid&volver=" & vuelta
      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
    end if
    Formulario = formulario & ofv.GenForm("edGeneral3","",sAccion,"","post","",Fila,"si")

  end if 
  call ofv.cerrarconsulta(rs)



' Boton Volver
    sAccion = "EdConsideraciones.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Menu=x&Sector=x"
    sInput = ofv.GenerarInput("","Volver","submit","","20","bt","")
 '   Formulario = ofv.GenForm("FormX","",sAccion,"Transfer","post","",sInput,"si")
    Celda = ofv.GenCelda("","abc colspan=2","","","","","",sInput,"si")
    Fila = ofv.GenRow("","","","","","","",Celda,"si")

    Formulario = formulario & ofv.GenForm("edGeneralV","",sAccion,"","post","",Fila,"si")

    sHTML = sHTML & "<br><br><center>"
    sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")
    sHTML = sHTML & "<center>"



  call ofv.cerrarconn(cn)
  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>