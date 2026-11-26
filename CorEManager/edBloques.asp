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
  if request.querystring("ID") <> "" then
     Session("TabID") = request.querystring("ID")
     sID = request.querystring("ID")
  else
     sID = Session("TabID")
  end if

  CantRegAMover = cdbl(request.querystring("CantRegAMover"))
  CantRegAMover2 = cdbl(request.querystring("CantRegAMover2"))
  CantRegAMostrar = 1
  If BrowserName() = "IE" OR BrowserName() = "CHROME" then
    CantRegAMostrar2 = 10000
  else
    CantRegAMostrar2 = 10000
  end if

  vuelta = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover) & "_ID=" & sID
  vuelta2 = ofv.vueltaaspEx("../",CantRegAMostrar2,CantRegAMover2) & "_ID=" & sID
  vueltaC = ofv.vueltaaspExC("../",CantRegAMover,CantRegAMover2) & "_ID=" & sID

  sHTML = ofv.FormBody2

  'Encabezados
  Formulario = ofv.Encabezados("Orden","Titulo")

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "SELECT * From Objtabblk Where Tabdef = " & sID
'response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegAMover)
  if not rs.eof then
    nID = rs(0)

    'Eliminar
    sHref = ""
    if bModificar then
      sHref = "asp/EliminarEN.asp?ID=" & nID & "&Tabla=Objtabblk&volver=" & vuelta
      sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
    end if
    if sHref <> "" then
      Celda = ofv.BotonEliminar(rs(5),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","","si")
    end if

    'Orden
    sInput = ofv.GenerarInput(rs(4).name,rs(4),"text","edBloques","10","dt","")
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

    'Titulo
    sInput = ofv.GenerarInput(rs(5).name,rs(5),"text","edBloques","120","dt","")
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

    Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
'    Fila = ofv.GenRow("","","","","","","",Celda,"si")

    if bModificar then
      sAccion = "asp/ActualizarEN.asp?ID=" & nID & "&Tabla=Objtabblk&volver=" & vuelta
      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      Formulario = Formulario & ofv.GenForm("edBloques","",sAccion,"","post","",celda,"si")
    else
      Formulario = Formulario & ofv.GenForm("","","","","post","",celda,"si")
    end if

    'test
    sInput = ofv.GenerarInput("","Test","submit","","20","bt","")
    Celda = ofv.GenCelda("","","","","","","",sInput,"si")
    sAccion = "CoTablas3.asp?Tabla=" & sID
    Formulario = Formulario & ofv.GenForm("testform","",sAccion,"_blank","post","",Celda,"si")

    'Volver
    sInput = ofv.GenerarInput("","Volver","submit","","20","bt","")
    Celda = ofv.GenCelda("","","","","","","",sInput,"si")
    sAccion = "edTablas.asp?Tipobjeto=" & sTipo & "&Codobjeto=" & nCodigo & "&Menu=x&Sector=x"
    Formulario = Formulario & ofv.GenForm("testform","",sAccion,"","post","",Celda,"si")

    Fila = ofv.GenRow("","","","","","","",Formulario,"si")

    rs.movenext
  else
    nID = "no"
  end if
  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",fila,"si")

  'Botonera
  if bModificar then
    sAccion = "AgregarBloque.asp?Tabla=Objtabblk&volver=" & vuelta & "&Codobjeto="
    sAccion = sAccion & nCodigo & "&Tipobjeto=" & sTipo & "&ID=" & sID
    Agregar = ofv.BotonAgregar(sAccion)
    sRetornar = "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
    sHTML = sHTML & ofv.botoneraESSX("3",rs,CantRegAMover,CantRegAMostrar,"no","no",posi,"","","no",Agregar,"no",sRetornar)
  else
    sHTML = sHTML & ofv.botoneraESSX("a",rs,CantRegAMover,CantRegAMostrar,"","","",nCodigo,sTipo,"no","no","no","")
  end if
  sHTML = sHTML & "<hr color=#000088 size=1>"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  redim aColumnas(0)
  if nID <> "no" then
    strsql = "Select tabperfil From Objtabdef Where objid = " & sID
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs2.eof then
    strsql = "Select * From Tabperfil Where Tabperfil = '" & rs2(0) & "'"
    call ofv.cerrarconsulta(rs2)
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs2.eof then
      nCols = rs2(7)
      redim preserve aColumnas(nCols - 1)
      for ix = 0 to UBound(aColumnas)
        aColumnas(ix) = rs2(ix + 16)
      next
    end if
    end if
    call ofv.cerrarconsulta(rs2)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Encabezados titulos
    Formulario = ofv.Encabezados("Orden","Titulos")

    strsql = "SELECT * From Objtabcon Where tabblk = " & nID & " and tabcontipo = '1' "
    strsql = strsql & " order by tabconsec "
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    posicion = 1
    do until rs2.eof
      nAccion = rs2(0)
      sForm = "edtitulo" & rs2(0)

      'Eliminar
      sHref = ""
      if bModificar then
        sHref = "asp/EliminarEN.asp?ID=" & nAccion & "&Tabla=Objtabcon&volver=" & vuelta
        sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
      end if
      if sHref <> "" then
        Celda = ofv.BotonEliminar(rs2(4),sHref)
      else
        Celda = ofv.GenCelda("","","","","","","","","si")
      end if

      'Orden y Tipo
      sInput = ofv.GenerarInput(rs2(4).name,rs2(4),"text",sForm,"10","dt","")
      sInput2 = ofv.GenerarInput(rs2(5).name,rs2(5),"hidden",sForm,"10","dt","")
      Celda = Celda & ofv.GenCelda("","","","2%","","","",sInput & sInput2,"si")

      'Columnas
      for ix = 0 to UBound(aColumnas)
        sSize = "10"
        if aColumnas(ix) <> 0 then sSize = cstr(aColumnas(ix))
        sInput = ofv.GenerarInput(rs2(ix + 6).name,rs2(ix + 6),"textsb",sForm,sSize,"dt","")
        Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")
      next
      Celda = Celda & ofv.GenCelda("","","","60%","","","","&nbsp;","si")
      Fila = ofv.GenRow("","","","","","","",Celda,"si")

      if bModificar then
        sAccion = "asp/ActualizarEN.asp?ID=" & nAccion & "&Tabla=Objtabcon&volver=" & vuelta
        sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
        Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",Fila,"si")
      else
        Formulario = Formulario & ofv.GenForm("","","","","post","",Fila,"si")
      end if

      rs2.movenext
      posicion = posicion + 1
    loop
    sHTML = sHTML & ofv.GenTabla("","","","50%","","0","","0","0","","",Formulario,"si")

    if bModificar then
      'Boton Agregar
      sRetornar = "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      sAccion = "AgregarCont.asp?Tabla=Objtabcon&volver=" & vuelta & "&Codobjeto="
      sAccion = sAccion & nCodigo & "&Tipobjeto=" & sTipo & "&ID=" & nID & "&IDt=1"
      sAccion = sAccion & "&CantRegAMover=" & CantRegAMover & "&CantRegAMover2=" & CantRegAMover2
      Agregar = ofv.BotonAgregar(sAccion)

      sHTML = sHTML & ofv.botoneraESSX("5",rs2,CantRegAMover,"",CantRegAMover2,CantRegAMostrar2,posicion,"","","no",Agregar,"no",sRetornar)
    else
      sHTML = sHTML & ofv.botoneraESSX("b",rs2,"","",CantRegAMover2,CantRegAMostrar2,"",nCodigo,sTipo,"no","no","no","")
    end if

    call ofv.cerrarconsulta(rs2)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Encabezados Contenido
    Formulario = ofv.Encabezados("Orden","Contenidos")

    strsql = "SELECT * From Objtabcon Where tabblk = " & nID & " and tabcontipo = '2' "
    strsql = strsql & " order by tabconsec "
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

    posicion = 1
    do until rs2.eof
      nAccion = rs2(0)
      sForm = "edContenido" & rs2(0)

      'Eliminar
      sHref = ""
      if bModificar then
        sHref = "asp/EliminarEN.asp?ID=" & nAccion & "&Tabla=Objtabcon&volver=" & vuelta
        sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
      end if
      if sHref <> "" then
        Celda = ofv.BotonEliminar(rs2(4),sHref)
      else
        Celda = ofv.GenCelda("","","","","","","","","si")
      end if

      'Orden y Tipo
      sInput = ofv.GenerarInput(rs2(4).name,rs2(4),"text",sForm,"10","dt","")
      sInput2 = ofv.GenerarInput(rs2(5).name,rs2(5),"hidden",sForm,"10","dt","")
      Celda = Celda & ofv.GenCelda("","","","2%","","","",sInput & sInput2,"si")


      'Columnas
      for ix = 0 to UBound(aColumnas)
        sSize = "10"
        if aColumnas(ix) <> 0 then sSize = cstr(aColumnas(ix))
        sInput = ofv.GenerarInput(rs2(ix + 6).name,rs2(ix + 6),"textsb",sForm,sSize,"dt","")
        Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")
      next
      Celda = Celda & ofv.GenCelda("","","","60%","","","","&nbsp;","si")
      Fila = ofv.GenRow("","","","","","","",Celda,"si")

      if bModificar then
        sAccion = "asp/ActualizarEN.asp?ID=" & nAccion & "&Tabla=Objtabcon&volver=" & vuelta
        sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
        Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",Fila,"si")
      else
        Formulario = Formulario & ofv.GenForm("","","","","post","",Fila,"si")
      end if

      rs2.movenext
      posicion = posicion + 1
    loop
    sHTML = sHTML & ofv.GenTabla("","","","50%","","0","","0","0","","",Formulario,"si")

    if bModificar then
      'Boton Agregar
      sRetornar = "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      sAccion = "AgregarCont.asp?Tabla=Objtabcon&volver=" & vuelta & "&Codobjeto="
      sAccion = sAccion & nCodigo & "&Tipobjeto=" & sTipo & "&ID=" & nID & "&IDt=2"
      sAccion = sAccion & "&CantRegAMover=" & CantRegAMover & "&CantRegAMover2=" & CantRegAMover2
      Agregar = ofv.BotonAgregar(sAccion)

      sHTML = sHTML & ofv.botoneraESSX("5",rs2,CantRegAMover,"",CantRegAMover2,CantRegAMostrar2,posicion,"","","no",Agregar,"no",sRetornar)
    else
      sHTML = sHTML & ofv.botoneraESSX("b",rs2,"","",CantRegAMover2,CantRegAMostrar2,"",nCodigo,sTipo,"no","no","no","")
    end if

    call ofv.cerrarconsulta(rs2)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  end if

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>