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
''response.write request.querystring & "<br>"
  nAccion = request.querystring("Accion")
  nProced = request.querystring("Codproced")
  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  vuelta = request.querystring("volver")

''response.write nAccion & " nAccion "
''response.write nProced & " nProced "
''response.write nCodigo & " nCodigo "
''response.write sTipo   & " sTipo "

  if nAccion & "" <>""  then
   session("SnAccion")= nAccion
  else
   nAccion= session("SnAccion")
  end if

  if nProced & "" <>""  then
     session("SnProced")= nProced
    else
     nProced= session("SnProced")
  end if

  if session("volverproc") = "" then session("volverproc") = vuelta
  volver2 = ofv.vueltaasp2("../")
''response.write volver2
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
''response.write strsql & " strsql "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
  if not rs.eof then Campo = rs(2).name
  call ofv.cerrarconsulta(rs)



  strsql = "SELECT * from accionmaster "
  strsql = strsql & "Where codproced = '" & nProced & "'"
  strsql = strsql & " Order By secuen"
''response.write strsql & " strsql "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
  x = 1
  redim AccionesAlt(0)
  redim AccionesAlt2(0)
  do until rs.eof
    redim preserve AccionesAlt(x)
    AccionesAlt(x) = rs(3)
    redim preserve AccionesAlt2(x)
    AccionesAlt2(x) = rs(2) & " - " & rs(1)

'''response.write AccionesAlt(x) & " AccionesAlt " & chr(13)
'''response.write AccionesAlt2(x) & " AccionesAlt2 " & chr(13)


    x = x + 1
    rs.movenext
  loop
  call ofv.cerrarconsulta(rs)

  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

'-------------------------------------------------

  Strsql = "SELECT * from csprocmaster "
  Strsql = Strsql & "Where codobjeto = " & nCodigo & " and codproced = '" & nProced & "' "
''response.write strsql & " strsql "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if  not rs.eof then

      'Secuencia evento
      sInput = rs(8) & "."

      'Secuencia procedimiento
        sInput = sInput & rs(7) & "&nbsp;"

    'Evento

      sEvento = ""
      if  rs(4) <> "" then
          Strsql = "Select descri From Objeven Where objid = " & rs(4)
        ''response.write strsql & " strsql "
          set rsEvento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if  not rsEvento.eof then sEvento = rsEvento(0)
          call ofv.cerrarconsulta(rsEvento)
      end if
      sInput = sInput & sEvento & "-"

    'Sector
      sSector = ""
      if  rs(5) <> "" then
          Strsql = "Select descri From Eco, Objsector where "
          Strsql = Strsql & "Eco.eco = Objsector.codsector and Objsector.objid = " & rs(5)
        ''response.write strsql & " strsql "
          set rsSector = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if  not rsSector.eof then sSector = rsSector(0)
          call ofv.cerrarconsulta(rsSector)
      end if
      sInput = sInput & sSector & "-"

    'Perfil
      sPerfil = ""
      if  rs(6) <> "" then
          Strsql = "Select descri From Tipoperfil Where tipoperfil = '" & rs(6) & "'"
        ''response.write strsql & " strsql "
          set rsPerfil = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if  not rsPerfil.eof then sPerfil = rsPerfil(0)
          call ofv.cerrarconsulta(rsPerfil)
      end if
      sInput = sInput & sPerfil & "-"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      set cn2 = ofv.conectar(ofv.strconn2)

  ''''Para obtener la descripcion de la Accion
    ''Primero obtiene el tipo de accion de la tabla objaccion
      Strsql = "SELECT * From OBJACCION "
      Strsql = Strsql & " WHERE codproced = '" & nProced & "' "
      Strsql = Strsql & " AND modo = 'CN' "
    ''Strsql = Strsql & " AND (objid = '" & nCodigo & "' "
    ''Strsql = Strsql & " OR codobjeto = '" & nCodigo & "' ) "
    ''response.write strsql & " strsql Objaccion "
      set rsz = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
      if not rsz.eof then
       nTAccion  = rsz(4)
      end if
      call ofv.cerrarconsulta(rsz)

    ''Despues obtiene la descri de la accion de la tabla Tipoaccalt
      Strsql = "Select descri From Tipoaccalt "
      Strsql = Strsql & " Where tipoaccionalt = '" & nTAccion & "' "
    ''response.write strsql & " strsql Tipoaccalt "
      set rsy = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
      if not rsy.eof then
        nTdescri  = rsy(0)
      end if
      call ofv.cerrarconsulta(rsy)

      sInput = sInput & nTdescri & "-"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      Celda = Celda & ofv.GenCelda("","TTT","","","","","",sInput,"si")



      fILA = ofv.GenRow("","","","","","","",Celda,"si")

  END IF

  sHTML = sHTML & ofv.GenTabla("panel1","","","80%","",0,0,0,0,"","",FILA,"si")









  sHTML = sHTML & ofv.GenTabla("panel","","","80%","",0,0,0,0,"","","","no")

  Valor = ""



  strsql = "SELECT * from accionmaster2 "
  strsql = strsql & "Where idaccion = '" & nAccion & "'"
  if bVersion = "NO" then
   strsql = strsql & " and (modi = 0 or (modi <> 0 and "
   strsql = strsql & "tipmodi <> 'E')) "
  end if
  strsql = strsql & " order by secuencia, objid "

  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)


  if not rs.eof then

 ' strsql = "SELECT * FROM TIPOACCALT WHERE TIPOACCALT.modo<>'CN' "
 ' strsql = strsql & "ORDER BY descri"
   Strsql = "Select a.* From Tipoaccalt a, modo b WHERE a.estado=3 and a.modo = b.tipomodo and  b.control = 0 Order By a.descri"
 ''response.write strsql & " strsql Tipoaccalt II "
    'Encabezado Accion
  '''  h1 = ofv.GenCelda("","tt","","2%","","","","&nbsp;","si")
    h1 = ofv.GenCelda("","ttr colspan=5","","98%","","","","Pasos del circuito alternativo.","si")
    sHTML = sHTML & ofv.GenRow("","","","","22","","",h1,"si")




  do until rs.eof
    Valor = rs(0)
    if not isnull(rs(8)) then Continuador = rs(8)
    ID = rs(2)

    if bModificar then
      sAccion = "asp/ActualizarEN.asp?Tabla=Objaccioncontrol"
      sAccion = sAccion & "&volver=" & volver2 & "&Codobjeto="
      sAccion = sAccion & nCodigo & "&Tipobjeto=" & sTipo & "&Accion=" & nAccion
      sAccion = sAccion & "&Codproced=" & nProced & "&ID=" & rs(2)
    ''  sHTML = sHTML & ofv.GenForm("AccionesAlt" & rs(2),"",sAccion,"","post","","","no")
  ''  else
  ''    sHTML = sHTML & ofv.GenForm("","","","","post","","","no")
    end if


    if rs(9) And rs(10) = "E" then
      sImage = ofv.GenImage("","","","","Imagenes/Botdelx.jpg","0","")
      t1 = ofv.GenCelda("","","","","","","",sImage,"si")
    else
      if bModificar then
        if rs(9) And rs(10) = "A" then
          sHref = "asp/EliminarEN.asp?ID=" & rs(2) & "&Tabla=Objaccioncontrol&volver=" & volver2
          sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
          sHref = sHref & "&Accion=" & nAccion & "&Codproced=" & nProced & "&Eliminar=NSNC"
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
        sInput = ofv.GenerarInput(rs(6).name,rs(6),"text","AccionesAlt" & rs(2) & "1x" ,"2","dtn","SI")
        sInput = ofv.GenForm("AccionesAlt" & rs(2) & "1x","",sAccion,"","post","",sInput,"si")
      else
        sInput = ofv.GenerarInput("",rs(6),"readonly","","2","dtn","SI")
      end if
      t2a = ofv.GenCelda("","","","","","","",sInput,"si")



    if bModificar then
      'Combo Accion Alt

       scomKK = ""
       scomDD = ""
      'response.write strsql & "<br>" & chr(13)
       set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       DO UNTIL rsniv.eof
         scomKK = scomKK & rsniv(0) & chr(9)
         scomDD = scomDD & rsniv(1) & chr(9)
         RSNIV.MOVENEXT
       LOOP
       call ofv.cerrarconsulta(rsniv)

            t2  = ofv.gencomboFX("AccionesAlt" & rs(2) & "xa",Campo,Valor,scomKK,scomDD,"") '& "arreglar"
            t2 = ofv.GenForm("AccionesAlt" & rs(2) & "xa","",sAccion,"","post","",t2,"si")
           't2 = ofv.GenerarCombo("AccionesAlt" & rs(2),Campo,Valor,Strsql,cn,0,1,"") & "arreglar"
    else
      'Texto Accion Alt
      if rs(2) <> "" then
        Strsql = "Select descri From Tipoaccalt Where tipoaccionalt = " & Valor
     ''response.write Strsql & " Strsql " & chr(13)
        set rsAcciones = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsAcciones.eof then sAccion = rsAcciones(0)
        call ofv.cerrarconsulta(rsAcciones)
      end if
      t2 = ofv.GenerarInput("",sAccion,"readonly","" & "vacio","120","dt","")
     '' t2 = ofv.GenForm("" & "vacio","",sAccion,"","post","",t2,"si")
    end if

    t2 = ofv.GenCelda("","","","","","","",t2,"si")

      'Version
      if bVersion = "SI" then
      sSelect = ""
      if rs(9) then sSelect = "selected"
      if sSelect <> "" then
        if bModificar then
          sInput = "<font class=at><select name=" & rs(9).name
          sInput = sInput & " onchange=AccionesAlt" & rs(2) & "x2a" & ".submit(); >"
          sInput = sInput & "<option value=False>No&nbsp;Modificada</option>"
          sInput = sInput & "<option value=True " & sSelect & ">Modificada"
          sInput = sInput & "</option></select></font>"
          'sInput = ofv.GenForm("AccionesAlt" & rs(2) & "x2a" ,"",sAccion,"","post","",sInput,"si")

        else
          sInput = ofv.GenerarInput(rs(9).name,rs(9),"readonly","" & "x0","2","dt","")
          'sInput = ofv.GenForm("" & "x0","",sAccion,"","post","",sInput,"si")
        end if
        'sInput = ofv.GenForm("AccionesAlt" & rs(2) & "x2a","",sAccion,"","post","",sInput,"si")
        t2b = ofv.GenCelda("","","","","","","",sInput,"si")
        if rs(10) = "A" then
          sInput = "Agregada"
        elseif rs(10) = "E" then
          sInput = "Eliminada"
        elseif rs(10) = "M" then
          sInput = "Modificada"
        end if
        'sInput = ofv.GenForm("AccionesAlt" & rs(2) & "x2a","",sAccion,"","post","",sInput,"si")
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

   ' Fila   = fila & ofv.GenRow("","","","","","","",t1 & t2a & t2 & t2b,"si")
    'sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","-2","-1","","",Fila,"si")
    'Fila3   =  ofv.GenRow("","","","","","","",t1,"si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",t1 & t2a & t2 & t2b,"si")
    'sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","-2","-1","","",Fila3,"si")

    'Encabezado Referencia
    h2 = ofv.GenCelda("","ut colspan=2 ","","","","","","&nbsp;","si")
    h2 = h2 & ofv.GenCelda("","ttt colspan=3","","","","","","Referencia","si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",h2,"si")


    'Texto de Referencia
     nCol = 70

    t3  = ofv.GenCelda("","ut colspan=2 ","","","","","","&nbsp;","si")
    t3b = "<font class=at><textarea class=dt "
    t3b = t3b & "name=referencia cols=" & nCol & " rows=4 "
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
    t3b = ofv.GenForm("AccionesAlt" & rs(2),"",sAccion,"","post","",t3b,"si")
    t3 = t3 & ofv.GenCelda("","uti colspan=3","","80%","","","",t3b,"si")
    sHTML = sHTML & ofv.GenRow("","","","","","","",t3,"si")
    'sHTML = sHTML & ofv.GenFForm()
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

       SCOMSEP = ""
       scomK   = ""
       scomD   = ""

       For i=1 To Ubound(AccionesAlt)
           scomK = scomK & SCOMSEP & ofv.convertircar(AccionesAlt(i)," ","_")
           scomD = scomD & SCOMSEP & AccionesAlt2(i)
           SCOMSEP = chr(9)
       next

       t4b = ofv.gencomboFX("XContinuador","Continuador",CStr(Continuador),scomK,scomD,"sb") '& "test"

      'Combo Continuador
   ''   t4b = "<font class=at><select name=Continuador onchange=XContinuador.submit()>"
   ''   t4b = t4b & "<option value=nsnc>&nbsp;</option>"
   ''   For n = 1 To Ubound(AccionesAlt)
   ''     If CStr(AccionesAlt(n)) = CStr(Continuador) Then
   ''       t4b = t4b & "<option value=" & ofv.convertircar(AccionesAlt(n)," ","_")
   ''       t4b = t4b & " selected>" & AccionesAlt2(n) & "</option>"
   ''     Else
   ''       t4b = t4b & "<option value=" & ofv.convertircar(AccionesAlt(n)," ","_")
   ''       t4b = t4b & ">" & AccionesAlt2(n) & "</option>"
   ''     End If
   ''   Next
   ''   t4b = t4b & "</select></font>"


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
    sHTML = sHTML & ofv.GenFForm()
    sHTML = sHTML & ofv.GenFTabla()
   end if

  else

     t4b   = "NO HAY CIRCUITO ALTERNATIVO INFORMADO."
     t4    = ofv.GenCelda("","vti","","","","","",t4b,"si")
     sHTML = sHTML & ofv.GenRow("","","","","60","","",t4,"si")
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

     sAccion = AgregarA
     Agregar = ofv.BotonAgregar(sAccion)

  end if


  'Boton Volver
  sAccion  = sVuelve
  Retornar = ofv.BotonVolver(sAccion)


    'Botonera



'--------------------------------------------------------------------------------
  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

sHTML = sHTML & ofv.botoneraESSX("3",rs,CantRegAMover,CantRegAMostrar,"no","no",posi,"","","no",Agregar,Retornar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>