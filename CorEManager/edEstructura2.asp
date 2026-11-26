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

  nCodigo          = request.querystring("Codcomp")
  sTipo            = request.querystring("Tipocomp")
  CantRegAMover    = cdbl(request.querystring("CantRegAMover"))
  CantRegAMostrar  = 1
  CantRegAMover2   = cdbl(request.querystring("CantRegAMover2"))
  CantRegAMostrar2 = 2
  vuelta           = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover)

  set cn = ofv.conectar(ofv.strconn2)

  sHTML  = ofv.FormHeader(Session("FormN"))


  'Encabezado Estructura
  Celda = ofv.GenCelda("","ttt","","2%","","","","Estructura","si")

  Strsql = "Select * From Compestr Where codcomp = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  nID = ""
  if not rs.eof then rs.move(CantRegAMover)
  if not rs.eof then
    nID = rs(0)

    'Eliminar
    if bModificar then
      sHref = "asp/EliminarC.asp?ID=" & rs(0) & "&Tabla=Compestr&volver=" & vuelta
      sHref = sHref & "&Codcomp=" & nCodigo & "&Tipocomp=" & sTipo & "&Eliminar=NSNC"
      Celda = Celda & ofv.BotonEliminar(rs(3),sHref)
    else
      Celda = Celda & ofv.GenCelda("","","","1%","","","","&nbsp;","si")
    end if

    'Estructura
    if bModificar then
      sInput = ofv.GenerarInput(rs(3).name,rs(3),"text","edEstructura2","120","dt","")
    else
      sInput = ofv.GenerarInput("",rs(3),"readonly","","120","dt","")
    end if
    Celda = Celda & ofv.GenCelda("","ut","","","","","",sInput,"si")

    rs.movenext
  else
      Celda = Celda & ofv.GenCelda("","ttt","","90%","","","","&nbsp;sin definir","si")
  end if

    Fila  = ofv.GenRow("","","","","25","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/ActualizarC.asp?ID=" & nID & "&Tabla=Compestr&volver=" & vuelta
      sAccion = sAccion & "&Codcomp=" & nCodigo & "&Tipocomp=" & sTipo
    end if
    Formulario = ofv.GenForm("edEstructura2","",sAccion,"","post","",Fila,"si")

  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")

  'Agregar
  Agregar = ""
  if bModificar then
    sInput = ofv.GenerarInput("","Agregar","submit","","10","bt","")
    Celda = ofv.GenCelda("","","","","","","",sInput,"si")
    sAccion = "AgregarEstructura2.asp?Tabla=Compestr&volver=" & vuelta & "&Codcomp="
    sAccion = sAccion & nCodigo & "&Tipocomp=" & sTipo & "&ID=" & nID
    Agregar = ofv.GenForm("Agregar","",sAccion,"","post","",Celda,"si")
  end if
  sHTML = sHTML & ofv.botoneraESSX("8",rs,CantRegAMover,CantRegAMostrar,"","","",nCodigo,sTipo,"no",Agregar,"no","")

  call ofv.cerrarconsulta(rs)

  sHTML = sHTML & "<hr color=#dddddd size=2>"

  'Encabezados
  Formulario = ofv.Encabezados("&nbsp;","Descripcion","Tipo&nbsp;de&nbsp;Dato/Columna")

  if nID <> "" then
    Strsql = "Select * From Compdet Where codestr=" & nID & " order by datoformul, objid"
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs2.eof then rs2.move(CantRegAMover2)
    do until rs2.eof
      sForm = "edEstructuraDetalle" & rs2(0)

      'Eliminar
      if bModificar then
        sHref = "asp/EliminarC.asp?ID=" & rs2(0) & "&Tabla=Compdet&volver=" & vuelta
        sHref = sHref & "&Codcomp=" & nCodigo & "&Tipocomp=" & sTipo & "&Eliminar=NSNC"
        Celda = ofv.BotonEliminar(rs2(4),sHref)
      else
        Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
      end if

      'Secuencia
      if bModificar then
        sInput = ofv.GenerarInput(rs2(6).name,rs2(6),"text",sForm,"2","dtn","SI")
      else
        sInput = ofv.GenerarInput("",rs2(6),"readonly","","2","dtn","")
      end if
      Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

      'Descripcion
      if bModificar then
        sInput = ofv.GenerarInput(rs2(4).name,rs2(4),"text",sForm,"80","dt","")
      else
        sInput = ofv.GenerarInput("",rs2(4),"readonly","","80","dt","")
      end if
      Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

      'Tipo de Dato/Columna
      if bModificar then
        sInput = ofv.GenerarInput(rs2(5).name,rs2(5),"text",sForm,"35","dt","")
      else
        sInput = ofv.GenerarInput("",rs2(5),"readonly","","35","dt","")
      end if
      Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

      Celda = Celda & ofv.GenCelda("","","","","","","","","si")
      Fila = ofv.GenRow("","","","","","","",Celda,"si")

      sAccion = ""
      if bModificar then
        sAccion = "asp/ActualizarC.asp?ID=" & rs2(0) & "&Tabla=Compdet&volver=" & vuelta
        sAccion = sAccion & "&Codcomp=" & nCodigo & "&Tipocomp=" & sTipo
      end if
      Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",Fila,"si")

      rs2.movenext
    loop
    sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")

    'Agregar
    Agregar = ""
    if bModificar then
      sInput = ofv.GenerarInput("","Agregar","submit","","10","bt","")
      Celda = ofv.GenCelda("","","","","","","",sInput,"si")
      sAccion = "AgregarEstructuraDetalle.asp?Tabla=Compdet&volver=" & vuelta & "&Codcomp="
      sAccion = sAccion & nCodigo & "&Tipocomp=" & sTipo & "&ID=" & nID
      Agregar = ofv.GenForm("Agregarb","",sAccion,"","post","",Celda,"si")
    end if
    sHTML = sHTML & ofv.botoneraESSX("9",rs2,"","",CantRegAMover2,CantRegAMostrar2,"",nCodigo,sTipo,"no",Agregar,"no","")

    call ofv.cerrarconsulta(rs2)
  end if

  sHTML = sHTML & "</body></html>"

  call ofv.cerrarconn(cn)
end if

response.write sHTML

%>