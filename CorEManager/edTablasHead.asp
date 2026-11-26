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
  sOpcion = request.querystring("OPC")
  if sOpcion <> "" then sOpcion = Ucase(sOpcion) 
  sconte = request.querystring("OConte")
  if sconte = "" then sconte = "NO"
  sconteK = request.querystring("OConteK")
  sconteD = request.querystring("OConteD")

  CantRegAMover = cdbl(request.querystring("CantRegAMover"))
  CantRegAMover2 = cdbl(request.querystring("CantRegAMover2"))
  CantRegAMostrar = 1
  CantRegAMostrar2 = 1


  vuelta = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover) & "_Oconte=" & sconte  & "_OconteK=" & sconteK
  vuelta2 = ofv.vueltaaspEx("../",CantRegAMostrar2,CantRegAMover2) & "_Oconte=" & sconte  & "_OconteK=" & sconteK
  vueltaC = ofv.vueltaaspExC("../",CantRegAMover,CantRegAMover2) & "_Oconte=" & sconte  & "_OconteK=" & sconteK

  sHTML = ofv.FormBody

  'Encabezados
    if sconte = "SI" then
       sdtab = "Edicion de Tabla: " & sconteD & " asociada al contenido"
       Celda = ofv.GenCelda("","tt","center colspan=5","","","","",sdtab,"si")
       Fila = ofv.GenRow("","","","","","","",Celda,"si")
    end if 

    Fila = fila & ofv.Encabezados("Orden","Titulo","Modelo","")

  set cn = ofv.conectar(ofv.strconn2)

  if sconte = "SI" then
     Strsql = "Select * From ObjTabdef Where codobjeto=" & nCodigo & " and tabdettablas = '"
     Strsql = Strsql & sconteK & "' "
  else
     Strsql = "Select * From ObjTabdef Where codobjeto=" & nCodigo & " and tabdettablas is null "
     Strsql = Strsql & " Order By tabsecuen"
  end if

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof and sconte = "SI" then
     strsql = "Insert into Objtabdef (codobjeto,tipobjeto,tabtitulo,tabsecuen,tabdettablas) "
     Strsql = Strsql & " values (" & nCodigo & ",'" & sTipo & "','" 
     Strsql = Strsql & ofv.convertircar(sconteD,"_"," ") & "',10,'"
     Strsql = Strsql & sconteK & "')"
     call ofv.crearconsultaEx(StrSql,cn,1,parametros)

     Strsql = "Select * From ObjTabdef Where codobjeto=" & nCodigo & " and tabdettablas = '"
     Strsql = Strsql & sconteK & "' "
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  end if

  if not rs.eof then rs.move(CantRegAMover)
  if not rs.eof then
    sForm = "edTablas" & rs(0)
    sid = rs(0)

    'Eliminar
    sHref = ""
    if bModificar and sconte <> "SI" then
      sHref = "asp/EliminarEN.asp?ID=" & rs(0) & "&Tabla=Objtabdef&volver=" & vuelta
      sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
    end if
    if sHref <> "" then
      Celda = ofv.BotonEliminar("Tabla",sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","","si")
    end if

    'Secuencia
    if bModificar and sconte <> "SI" then
      sInput = ofv.GenerarInput(rs(5).name,rs(5),"text",sForm,"5","dtn","SI")
    else
      sInput = ofv.GenerarInput("",rs(5),"readonly","","5","dtn","SI")
    end if
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

    'Descripcion
    if bModificar then
      sInput = ofv.GenerarInput(rs(4).name,rs(4),"text",sForm,"90","dt","")
    else
      sInput = ofv.GenerarInput("",rs(4),"readonly","","90","dt","")
    end if
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

    'Perfil
    Strsql = "Select * From Tabperfil Order By tabdescri"
    sInput = ofv.GenerarCombo(sForm,rs(3).name,rs(3).value,Strsql,cn,0,1,"")
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

    if bModificar then
      sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objtabdef&volver=" & vuelta
      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      Formulario = ofv.GenForm(sForm,"",sAccion,"","post","",Celda,"si")
    else
      Formulario = ofv.GenForm("","","","","post","",Celda,"si")
    end if

    if cstr(rs(3)) <> "" then
      'test
       sInput = ofv.GenerarInput("","Test","submit","","20","bt","")
       Celda = ofv.GenCelda("","","","","","","",sInput,"si")
       sAccion = "CoTablas3.asp?Tabla=" & sID
       Formulario = Formulario & ofv.GenForm("testform","",sAccion,"_blank","post","",Celda,"si")
    else
       sid = "NO"
    end if



    Fila = fila & ofv.GenRow("","","","","","","",Formulario,"si")

    rs.movenext
  else
     sid = "NO"
  end if

'  sHTML = sHTML & "<br><br><br><br>"
  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",fila,"si")

  'Botonera
    Agregar = ""

    sRetornar = "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
    sAccion = "AgregarTablas.asp?Tabla=ObjTabdef&volver=" & vuelta & "&Codobjeto=" & nCodigo
    sAccion = sAccion & "&Tipobjeto=" & sTipo
  if bModificar and sconte <> "SI"  then
    Agregar = ofv.BotonAgregar(sAccion)
  end if
    sHTML = sHTML & ofv.botoneraESSX("3",rs,CantRegAMover,CantRegAMostrar,"no","no",posi,"","","no",Agregar,"no",sRetornar)



  call ofv.cerrarconsulta(rs)

  if sid <> "NO" then
'''''''''''''''''''''''''''''''''''''''''''''''
  
  'Encabezados
  Formulario1 = ofv.Encabezados("Orden","Subtitulo")

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "SELECT * From Objtabblk Where Tabdef = " & sID & " order by tabblksec "

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then 
     Strsqlx = "insert into Objtabblk (codobjeto,tipobjeto,Tabdef,tabblksec) values ("
     Strsqlx = Strsqlx & nCodigo & ",'" & sTipo & "'," & sID & ",10)"

     call ofv.crearconsultaEx(StrSqlx,cn,1,vueltac)
  end if
  call ofv.cerrarconsulta(rs)


  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegAMover2)
  if not rs.eof then
    nID = rs(0)

    'Eliminar
    sHref = ""
    if bModificar then
      sHref = "asp/EliminarEN.asp?ID=" & nID & "&Tabla=Objtabblk&volver=" & vueltac
      sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
    end if
    if sHref <> "" then
      Celda = ofv.BotonEliminar("Subtitulo",sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","","si")
    end if

    'Orden
    sInput = ofv.GenerarInput(rs(4).name,rs(4),"text","edBloques","10","dt","")
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

    'Titulo

    sInput = ofv.GenerarInput(rs(5).name,rs(5),"textsb","edBloques","120","dt","")
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

    Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
'    Fila = ofv.GenRow("","","","","","","",Celda,"si")

    if bModificar then
      sAccion = "asp/ActualizarEN.asp?ID=" & nID & "&Tabla=Objtabblk&volver=" & vueltac
      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      Formulario = ofv.GenForm("edBloques","",sAccion,"","post","",celda,"si")
    else
      Formulario = ofv.GenForm("","","","","post","",celda,"si")
    end if



    Fila = ofv.GenRow("","","","","","","",Formulario1 & Formulario,"si")

    rs.movenext
  else 
    Fila = ofv.GenRow("","","","","","","",Formulario1,"si")
    nID = "no"
  end if
  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",fila,"si")

  'Botonera
  if bModificar then
    sAccion = "AgregarBloque.asp?Tabla=Objtabblk&volver=" & vueltac & "&Codobjeto="
    sAccion = sAccion & nCodigo & "&Tipobjeto=" & sTipo & "&ID=" & sID
    Agregar = ofv.BotonAgregar(sAccion)
    sRetornar = "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      sHTML = sHTML & ofv.botoneraESSX("5",rs,CantRegAMover,"",CantRegAMover2,CantRegAMostrar2,posicion,"","","no",Agregar,"no",sRetornar)
    else
      sRetornar = "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      sHTML = sHTML & ofv.botoneraESSX("5",rs,CantRegAMover,"",CantRegAMover2,CantRegAMostrar2,posicion,"","","no","no","no",sRetornar)
'      sHTML = sHTML & ofv.botoneraESSX("b",rs,"","",CantRegAMover2,CantRegAMostrar2,"",nCodigo,sTipo,"no","no","no","")

  end if

  call ofv.cerrarconsulta(rs)
  
'  sHTML = sHTML & "<hr color=#000088 size=1>"

  wss = "x"

  do until wss = "x"

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
      if cint(nCols) > 10 then nCols = 10 
      redim preserve aColumnas(cint(nCols) - 1)
      for ix = 0 to UBound(aColumnas)
        aColumnas(ix) = rs2(ix + 16)
      next
    end if
    end if
    call ofv.cerrarconsulta(rs2)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Encabezados titulos
    Formulario = ofv.Encabezados("Orden","Titulos&nbsp;de&nbsp;Columnas")
    sHTML = sHTML & ofv.GenTabla("","","","50%","","0","","0","0","","",Formulario,"si")
    Formulario = ""

    if sOpcion = "T" then

       strsqlx = "Select * From Objtabcon Where tabblk = " & nID
       strsqlx = strsqlx &" and tabcontipo = '1' "
       strsqlx = strsqlx & " Order By tabconsec desc"
       set rsx = ofv.crearconsultaEx(StrSqlx,cn,1,volver)
       if not rsx.eof then
          nSecuencia = cint(rsx("tabconsec")) + 10
       else
          nSecuencia = 10
       end if
       call ofv.cerrarconsulta(rsx)
     Strsqlx = "insert into Objtabcon (codobjeto,tipobjeto,Tabblk,tabconsec,tabcontipo) values ("
     Strsqlx = Strsqlx & nCodigo & ",'" & sTipo & "'," & nID & "," & nSecuencia & ",1)"
     call ofv.crearconsultaEx(StrSqlx,cn,1,vueltac)
    end if 

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
        sHref = "asp/EliminarEN.asp?ID=" & nAccion & "&Tabla=Objtabcon&volver=" & vueltac
        sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
      end if
      if sHref <> "" then
        Celda = ofv.BotonEliminar("Titulo",sHref)
      else
        Celda = ofv.GenCelda("","","","","","","","","si")
      end if

      'Orden y Tipo
      sInput = ofv.GenerarInput(rs2(4).name,rs2(4),"agregar",sForm,"10","dt","")
      sInput2 = ofv.GenerarInput(rs2(5).name,rs2(5),"hidden",sForm,"10","dt","")
      Celda = Celda & ofv.GenCelda("","cc","","2%","","","",sInput & sInput2,"si")

      'Columnas
      for ix = 0 to UBound(aColumnas)
        sSize = "10"
        if cint(aColumnas(ix)) <> 0 then sSize = cstr(aColumnas(ix))
        sInput = ofv.GenerarInput(rs2(ix + 6).name,rs2(ix + 6),"agregar",sForm,sSize,"dt","")
        Celda = Celda & ofv.GenCelda("","cc","","","","","",sInput,"si")
      next
      if bModificar then
         sInput = ofv.GenerarInput("boton1","Ok","submit",sForm,"5","bt","")
      else
         sInput = "&nbsp;"
      end if
      Celda = Celda & ofv.GenCelda("","cc","left","60%","","","",sInput,"si")
      Fila = ofv.GenRow("","","","","","","",Celda,"si")

      if bModificar then
        sAccion = "asp/ActualizarEN.asp?ID=" & nAccion & "&Tabla=Objtabcon&volver=" & vueltac
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
       saccion = mid(vueltac,4)
       saccion = ofv.convertircar(saccion,"_","&")
       saccion = saccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&OPC=T"

'      sAccion = "AgregarCont.asp?Tabla=Objtabcon&volver=" & vueltac & "&Codobjeto="
'      sAccion = sAccion & nCodigo & "&Tipobjeto=" & sTipo & "&ID=" & nID & "&IDt=1"
'      sAccion = sAccion & "&CantRegAMover=" & CantRegAMover & "&CantRegAMover2=" & CantRegAMover2
      Agregar = ofv.BotonAgregar(sAccion)
      sHTML = sHTML & ofv.GenTabla("","","","50%","","0","","0","0","","",Agregar,"si")
    end if

    call ofv.cerrarconsulta(rs2)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Encabezados Contenido
    Formulario = ofv.Encabezados("Orden","Contenidos&nbsp;de&nbsp;Columnas")
    sHTML = sHTML & ofv.GenTabla("","","","50%","","0","","0","0","","",Formulario,"si")
    Formulario = ""
    esttxt = " STYLE={font-family:tahoma;color:#225588;font-size:7pt;background-color:#f7eedd;} "
    if sOpcion = "C" then

       strsqlx = "Select * From Objtabcon Where tabblk = " & nID
       strsqlx = strsqlx &" and tabcontipo = '2' "
       strsqlx = strsqlx & " Order By tabconsec desc"
       set rsx = ofv.crearconsultaEx(StrSqlx,cn,1,volver)
       if not rsx.eof then
          nSecuencia = cint(rsx("tabconsec")) + 10
       else
          nSecuencia = 10
       end if
       call ofv.cerrarconsulta(rsx)
     Strsqlx = "insert into Objtabcon (codobjeto,tipobjeto,Tabblk,tabconsec,tabcontipo) values ("
     Strsqlx = Strsqlx & nCodigo & ",'" & sTipo & "'," & nID & "," & nSecuencia & ",2)"
     call ofv.crearconsultaEx(StrSqlx,cn,1,vueltac)
    end if 

    strsql = "SELECT * From Objtabcon Where tabblk = " & nID & " and tabcontipo = '2' "
    strsql = strsql & " order by tabconsec "   
    set rs5 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
 '         rs5.move 1
    posicion = 1
    do until rs5.eof 


      nAccion = rs5(0)
      sForm = "edContenido" & rs5(0)



      'Eliminar
      sHref = ""
      if bModificar then
        sHref = "asp/EliminarEN.asp?ID=" & nAccion & "&Tabla=Objtabcon&volver=" & vueltac
        sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
      end if
      if sHref <> "" then
        Celda = ofv.BotonEliminar("Contenido",sHref)
      else
        Celda = ofv.GenCelda("","","","","","","","","si")
      end if

      'Orden y Tipo
      sInput = ofv.GenerarInput(rs5(4).name,rs5(4),"agregar",sForm,"10","dt","")
      sInput2 = ofv.GenerarInput(rs5(5).name,rs5(5),"hidden",sForm,"10","dt","")
      Celda = Celda & ofv.GenCelda("","aaa valign=top ","","2%","","","",sInput & sInput2,"si")


      'Columnas
      for ix = 0 to UBound(aColumnas)
'      for ix = 0 to 4
        sSize = "10"
        if cint(aColumnas(ix)) <> 0 then sSize = cstr(aColumnas(ix))
        sInput = ofv.GenerarInput(rs5(ix + 6).name,rs5(ix + 6),"agregar",sForm,sSize,"dt","")

    t3b = "<font class=at><textarea "
    t3b = t3b & "name=" & rs5(ix + 6).name & " cols=" & sSize & " rows=3 " & esttxt
    t3b = t3b & " onkeypress=" & chr(34)
    t3b = t3b & " if (this.value.length > 600) {alert('El texto es demasiado"
    t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
    xdatos =  rs5(ix + 6)
    t3b = t3b & xdatos & "</textarea></font>"

        Celda = Celda & ofv.GenCelda("","","","","","","",t3b,"si")
      next
      if bModificar then
         sInput = ofv.GenerarInput("boton1","Ok","submit",sForm,"5","bt","")
      else
         sInput = "&nbsp;"
      end if
      Celda = Celda & ofv.GenCelda("","aaa valign=top ","left","60%","","","",sInput,"si")
      Fila = ofv.GenRow("","","","","","","",Celda,"si")

      if bModificar then
        sAccion = "asp/ActualizarEN.asp?ID=" & nAccion & "&Tabla=Objtabcon&volver=" & vueltac
        sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
        Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",Fila,"si")
      else
        Formulario = Formulario & ofv.GenForm("","","","","post","",Fila,"si")
      end if

      rs5.movenext
      posicion = posicion + 1

    loop
    sHTML = sHTML & ofv.GenTabla("","","","50%","","0","","0","0","","",Formulario,"si")

    if bModificar then
      'Boton Agregar
       saccion = mid(vueltac,4)
       saccion = ofv.convertircar(saccion,"_","&")
       saccion = saccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&OPC=C"

'      sAccion = "AgregarCont.asp?Tabla=Objtabcon&volver=" & vueltac & "&Codobjeto="
'      sAccion = sAccion & nCodigo & "&Tipobjeto=" & sTipo & "&ID=" & nID & "&IDt=2"
'      sAccion = sAccion & "&CantRegAMover=" & CantRegAMover & "&CantRegAMover2=" & CantRegAMover2
      Agregar = ofv.BotonAgregar(sAccion)
      sHTML = sHTML & ofv.GenTabla("","","","50%","","0","","0","0","","",Agregar,"si")
    end if

    call ofv.cerrarconsulta(rs5)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

   end if
  loop

  end if


  call ofv.cerrarconn(cn)

  sHTML = sHTML & " </body></html>"
end if

response.write sHTML

%>