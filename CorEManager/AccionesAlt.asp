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

  nAccion = request.querystring("Accion")
  nProced = request.querystring("Codproced")
  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  vuelta = request.querystring("volver")
  if session("volverproc") = "" then session("volverproc") = vuelta
  volver2 = ofv.vueltaasp2("../")

  esttxt = " STYLE={font-family:tahoma;color:#000000;font-size:10pt;background-color:#f7eedd;} "

  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select Version From Objprty Where Objprty.codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  bVersion = "NO"
  if not rs.eof then
    if rs(0) = "V" then bVersion = "SI"
  end if
  call ofv.cerrarconsulta(rs)

  strsql = "SELECT * FROM OBJACCIONCONTROL"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
  if not rs.eof then Campo = rs(2).name
  call ofv.cerrarconsulta(rs)



  strsql = "SELECT * from accionmaster "
  strsql = strsql & "Where codproced = '" & nProced & "'"
  strsql = strsql & " Order By secuen"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
  x = 1
  redim AccionesAlt(0)
  redim AccionesAlt2(0)
  do until rs.eof
    redim preserve AccionesAlt(x)
    AccionesAlt(x) = rs(3)
    redim preserve AccionesAlt2(x)
    AccionesAlt2(x) = rs(2) & " - " & rs(1)
    x = x + 1
    rs.movenext
  loop
  call ofv.cerrarconsulta(rs)



  strsql = "SELECT * from accionmaster2 "
  strsql = strsql & "Where idaccion = '" & nAccion & "'"
  if bVersion = "NO" then
   strsql = strsql & " and (modi = 0 or (modi <> 0 and "
   strsql = strsql & "tipmodi <> 'E')) "
  end if
  strsql = strsql & " order by secuencia, objid "

  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

  strsql = "SELECT * FROM TIPOACCALT WHERE TIPOACCALT.modo='OP' Or TIPOACCALT.modo='RF' "
  strsql = strsql & "ORDER BY descri"

 '' sHTML = ofv.FormBody

  sHTML = ofv.FormHeader(session("FormN"))

  sHTML = sHTML & ofv.GenTabla("panel","","","80%","",0,0,0,0,"","","","no")

  Valor = ""
  do until rs.eof
    Valor = rs(0)
    if not isnull(rs(8)) then Continuador = rs(8)
    ID = rs(2)

    if bModificar then
      sAccion = "asp/ActualizarEN.asp?Tabla=Objaccioncontrol"
      sAccion = sAccion & "&volver=" & volver2 & "&Codobjeto="
      sAccion = sAccion & nCodigo & "&Tipobjeto=" & sTipo & "&Accion=" & nAccion
      sAccion = sAccion & "&Codproced=" & nProced & "&ID=" & rs(2)
      sHTML = sHTML & ofv.GenForm("AccionesAlt" & rs(2),"",sAccion,"","post","","","no")
    else
      sHTML = sHTML & ofv.GenForm("","","","","post","","","no")
    end if

    'Encabezado Accion
    h1 = ofv.GenCelda("","tt","","2%","","","","&nbsp;","si")
    h1 = h1 & ofv.GenCelda("","tt colspan=4","","98%","","","","Accion","si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")

    if rs(9) And rs(10) = "E" then
      sImage = ofv.GenImage("","","","","Imagenes/Botdelx.jpg","0","")
      t1 = ofv.GenCelda("","","","","","","",sImage,"si")
    else
      if bModificar then
        if rs(9) And rs(10) = "A" then
          sHref = "asp/EliminarEN.asp?ID=" & rs(2) & "&Tabla=Objaccioncontrol&volver=" & volver2
          sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
          sHref = sHref & "&Accion=" & nAccion & "&Codproced=" & nProced
        else
          sHref = "asp/ActualizarEN.asp?ID=" & rs(2) & "&Tabla=Objaccioncontrol&volver="
          sHref = sHref & volver2 & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
          sHref = sHref & "&Accion=" & nAccion & "&Codproced="
          sHref = sHref & nProced & "&Eliminar=SI"
        end if
      end if
      if sHref <> "" then
        t1 = ofv.BotonEliminar("Accion",sHref)
      else
        t1 = ofv.GenCelda("","","","","","","","","si")
      end if
    end if

      'Secuencia
      if bModificar then
        sInput = ofv.GenerarInput(rs(6).name,rs(6),"text","AccionesAlt" & rs(2) ,"2","dtn","SI")
      else
        sInput = ofv.GenerarInput("",rs(6),"readonly","","2","dtn","SI")
      end if
      t2a = ofv.GenCelda("","","","","","","",sInput,"si")



    if bModificar then
      'Combo Accion Alt
      t2 = ofv.GenerarCombo("AccionesAlt" & rs(2),Campo,Valor,Strsql,cn,0,1,"")
    else
      'Texto Accion Alt
      if rs(2) <> "" then
        Strsql = "Select descri From Tipoaccalt Where tipoaccionalt = " & Valor
        set rsAcciones = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsAcciones.eof then sAccion = rsAcciones(0)
        call ofv.cerrarconsulta(rsAcciones)
      end if
      t2 = ofv.GenerarInput("",sAccion,"readonly","","120","dt","")
    end if
    t2 = ofv.GenCelda("","","","","","","",t2,"si")

      'Version
      if bVersion = "SI" then
      sSelect = ""
      if rs(9) then sSelect = "selected"
      if sSelect <> "" then
        if bModificar then
          sInput = "<font class=at><select name=" & rs(9).name
          sInput = sInput & " onchange=AccionesAlt" & rs(2) & ".submit(); >"
          sInput = sInput & "<option value=False>No&nbsp;Modificada</option>"
          sInput = sInput & "<option value=True " & sSelect & ">Modificada"
          sInput = sInput & "</option></select></font>"
        else
          sInput = ofv.GenerarInput(rs(9).name,rs(9),"readonly","","2","dt","")
        end if
        t2b = ofv.GenCelda("","","","","","","",sInput,"si")
        if rs(10) = "A" then
          sInput = "Agregada"
        elseif rs(10) = "E" then
          sInput = "Eliminada"
        elseif rs(10) = "M" then
          sInput = "Modificada"
        end if
        t2b = t2b & ofv.GenCelda("","dt","","","","","",sInput,"si")
      else
        sInput = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
        t2b = ofv.GenCelda("","","","","","","",sInput,"si")
        t2b = t2b & ofv.GenCelda("","","","","","","",sInput,"si")
      end if
      else
        t2b = ofv.GenCelda("","","","","","","","&nbsp;","si")
        t2b = t2b & ofv.GenCelda("","","","","","","","&nbsp;","si")
      end if

    sHTML = sHTML & ofv.GenRow("","","","","","","",t1 & t2a & t2 & t2b,"si")

    'Encabezado Referencia
    h2 = ofv.GenCelda("","tt","","","","","","&nbsp;","si")
    h2 = h2 & ofv.GenCelda("","tt colspan=4","","","","","","Referencia","si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h2,"si")

    'Texto de Referencia
    If BrowserName() = "IE" OR BrowserName() = "CHROME" then
      nCol = 80
    else
      nCol = 58
    end if
    t3 = ofv.GenCelda("","","","","","","","&nbsp;","si")
    t3b = "<font class=at><textarea "
    t3b = t3b & "name=referencia cols=" & nCol & " rows=4 " & esttxt
    if bModificar then
      t3b = t3b & " onblur=AccionesAlt" & rs(2) & ".submit() "
      t3b = t3b & " onkeypress=" & chr(34)
      t3b = t3b & " if (this.value.length > 2450) {alert('El texto es demasiado"
      t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
    else
      t3b = t3b & " readonly >"
    end if
    xdatos =  rs(7)
    t3b = t3b & xdatos & "</textarea></font>"
    t3 = t3 & ofv.GenCelda("","abc colspan=4","","80%","","","",t3b,"si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",t3,"si")
    sHTML = sHTML & ofv.GenFForm()
    rs.movenext
  loop
  sHTML = sHTML & ofv.GenFTabla()

  if Valor <> "" then
    if bModificar then
      f2a = "asp/ActualizarEN.asp?Tabla=Objaccioncontrol"
      f2a = f2a & "&volver=" & volver2 & "&Codobjeto="
      f2a = f2a & nCodigo & "&Tipobjeto=" & sTipo & "&Accion=" & nAccion
      f2a = f2a & "&Codproced=" & nProced & "&ID=" & ID  & "&conti=SI"
      sHTML = sHTML & ofv.GenForm("XContinuador","",f2a,"","post","","","no")
    else
      sHTML = sHTML & ofv.GenForm("","","","","post","","","no")
    end if

    sHTML = sHTML & ofv.GenTabla("panel2","","","80%","",0,0,0,0,"","","","no")

    'Encabezado Continuador
    t4 = ofv.GenCelda("","","","40%","","","","&nbsp;","si")
    t4 = t4 & ofv.GenCelda("","tt","","2%","","","","Continuador","si")
    if isnull(Continuador) then Continuador = 0
    if bModificar then
      'Combo Continuador
      t4b = "<font class=at><select name=Continuador onchange=XContinuador.submit()>"
      t4b = t4b & "<option value=nsnc>&nbsp;</option>"
      For n = 1 To Ubound(AccionesAlt)
        If CStr(AccionesAlt(n)) = CStr(Continuador) Then
          t4b = t4b & "<option value=" & ofv.convertircar(AccionesAlt(n)," ","_")
          t4b = t4b & " selected>" & AccionesAlt2(n) & "</option>"
        Else
          t4b = t4b & "<option value=" & ofv.convertircar(AccionesAlt(n)," ","_")
          t4b = t4b & ">" & AccionesAlt2(n) & "</option>"
        End If
      Next
      t4b = t4b & "</select></font>"
    else
      'Texto Continuador
      sContinuador = ""
      For n = 1 To Ubound(AccionesAlt)
        If CStr(AccionesAlt(n)) = CStr(Continuador) Then
           sContinuador = AccionesAlt2(n)
        End If
      Next
      t4b = ofv.GenerarInput("",sContinuador,"readonly","","40","dt","")
    end if
    t4 = t4 & ofv.GenCelda("","","","","","","",t4b,"si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",t4,"si")
    sHTML = sHTML & ofv.GenRow("","","","","","","","&nbsp;","si")
    sHTML = sHTML & ofv.GenFForm()
    sHTML = sHTML & ofv.GenFTabla()
  end if

  vuelta = session("volverproc")
  sVuelve = ofv.convertircar(vuelta,"_","&")
  sVuelve = ofv.convertircar(sVuelve,"*","&")
  sVuelta = ofv.convertircar(sVuelve,"&","*")

  on error resume next
  sVuelve = mid(sVuelve,4,len(sVuelve) - 3)

  Agregar = ""

  if bModificar then
     AgregarA = "AgregarAccionAlt.asp?Tabla=Objaccioncontrol"
     AgregarA = AgregarA & "&Codproced=" & nProced & "&Codobjeto="
     AgregarA = AgregarA & nCodigo & "&Tipobjeto=" & sTipo & "&Accion=" & nAccion

     Agregar = ofv.GenerarInput("bac1","Agregar","submit","","10","bt","")
     Agregar = ofv.GenCelda("","","center","50%","","","",Agregar,"si")
     Agregar = ofv.GenForm("Agregar","",AgregarA,"","post","",Agregar,"si")

     sAccion = AgregarA
     Agregar = ofv.BotonAgregar(sAccion)

  end if

  Retornar = ofv.GenerarInput("bvo1","Volver","submit","","10","bt","")
  Retornar = ofv.GenCelda("","","center","50%","","","",Retornar,"si")
  Retornar = ofv.GenForm("Retornar","",sVuelve,"","post","",Retornar,"si")

  'Boton Volver
  sAccion = sVuelve
  Retornar = ofv.BotonVolver(sAccion)


    'Botonera
    sHTML = sHTML & ofv.botoneraESSX("3",rs,CantRegAMover,CantRegAMostrar,"no","no",posi,"","","no",Agregar,Retornar,"")


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>