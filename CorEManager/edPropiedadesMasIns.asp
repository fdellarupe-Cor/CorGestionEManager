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
  bModificaresp = ofv.AModificar
  if session("Nomodificar") = "S" then

     if LEFT(Session("Form2"),4) <> "DESA" then    bModificar = False

  end if

  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  vuelta = ofv.vueltaasp2("../")

  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

 shtml = shtml & CorPnlFnc("N","")


  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

'-------------------------------------------------

  'Encabezados
  Formulario = ofv.Encabezados("#","origen","instrumento","identificacion","lnk 1","lnk 2")

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From Objprtyadv Where codobjeto=" & nCodigo & " Order By secuen"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  do until rs.eof
    sForm = "edPropiedades" & rs(0)

     if bModificar then
   sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objprtyadv&volver=" & vuelta
   sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
   end if

    'Eliminar
    if bModificar then
      sHref = "asp/EliminarEN.asp?ID=" & rs(0) & "&Tabla=Objprtyadv&volver=" & vuelta
      sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
    end if
    if sHref <> "" then
      Celda = ofv.BotonEliminar(rs(2),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","","si")
    end if

    'secuencia
    if bModificar then
      sInput = ofv.GenerarInput(rs(6).name,rs(6),"text",sForm & "w1","5","dtn","SI")
      sInput = ofv.GenForm(sForm & "w1","",sAccion,"","post","",sInput,"si")
    else
      sInput = ofv.GenerarInput("",rs(6),"readonly","","5","dtn","SI")
    end if
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")


    'origen
    if bModificar then
      Strsql = "Select * From Origen Order By descri"

      scomKK = ""
    scomDD = ""
  ''  response.write strsql & "<br>" & chr(13)
    set rsniv = ofv.crearconsultaEx(Strsql,cn,1,parametros)
    DO UNTIL rsniv.eof
       scomKK = scomKK & rsniv(0) & chr(9)
       scomDD = scomDD & rsniv(1) & chr(9)
       RSNIV.MOVENEXT
    LOOP
    call ofv.cerrarconsulta(rsniv)

      sInput = ofv.gencomboFX(sForm & "w2",rs(3).name,rs(3).value,scomKK,scomDD,"")
      sInput = ofv.GenForm(sForm & "w2","",sAccion,"","post","",sInput,"si")
    ''  sInput = ofv.GenerarCombo(sForm,rs(3).name,rs(3).value,Strsql,cn,0,1,"")
    else
      if cstr(rs(3)) <> "" then
        Strsql = "Select descri From Origen Where objid = " & rs(3)
        set rsOrigen = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsOrigen.eof then sOrigen = rsOrigen(0)
        call ofv.cerrarconsulta(rsOrigen)
      end if
      sInput = ofv.GenerarInput("",sOrigen,"readonly","","60","dt","")
    end if
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

    'instrumento
    if bModificar then
      Strsql = "Select * From Instrumen Order By descri"


      scomKK = ""
    scomDD = ""
  ''  response.write strsql & "<br>" & chr(13)
    set rsniv = ofv.crearconsultaEx(Strsql,cn,1,parametros)
    DO UNTIL rsniv.eof
       scomKK = scomKK & rsniv(0) & chr(9)
       scomDD = scomDD & rsniv(1) & chr(9)
       RSNIV.MOVENEXT
    LOOP
    call ofv.cerrarconsulta(rsniv)

      sInput = ofv.gencomboFX(sForm & "w3",rs(4).name,rs(4).value,scomKK,scomDD,"")
      sInput = ofv.GenForm(sForm & "w3","",sAccion,"","post","",sInput,"si")
      'sInput = ofv.GenerarCombo(sForm,rs(4).name,rs(4).value,Strsql,cn,0,1,"")
    else
      if cstr(rs(4)) <> "" then
        Strsql = "Select descri From Instrumen Where objid = " & rs(4)
        set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsInstrumento.eof then sInstrumento = rsInstrumento(0)
        call ofv.cerrarconsulta(rsInstrumento)
      end if
      sInput = ofv.GenerarInput("",sInstrumento,"readonly","","60","dt","")
    end if
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

    'identificacion
    if bModificar then
      sInput = ofv.GenerarInput(rs(5).name,rs(5),"text",sForm & "w5","20","dt","")
      sInput = ofv.GenForm(sForm & "w5","",sAccion,"","post","",sInput,"si")
    else
      sInput = ofv.GenerarInput("",rs(5),"readonly","","20","dt","")
    end if
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

    'link
    if bModificar then
      sInput = ofv.GenerarInput(rs(7).name,rs(7),"textsb",sForm & "w6","20","dt","")
      sInput = ofv.GenForm(sForm & "w6","",sAccion,"","post","",sInput,"si")
    else
      sInput = ofv.GenerarInput("",rs(7),"readonly","","20","dt","")
    end if
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

    'link 2
    if bModificar then
      sInput = ofv.GenerarInput(rs(8).name,rs(8),"textsb",sForm & "w7","20","dt","")
      sInput = ofv.GenForm(sForm & "w7","",sAccion,"","post","",sInput,"si")
    else
      sInput = ofv.GenerarInput("",rs(8),"readonly","","20","dt","")
    end if
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

   '' Fila = ofv.GenRow("","","","","","","",Celda,"si")

' '   if bModificar then
' '     sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objprtyadv&volver=" & vuelta
' '     sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
' '     Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",Fila,"si")
' '   else
' '     Formulario = Formulario & ofv.GenForm("","","","","post","",Fila,"si")
' '   end if
    Formulario = Formulario & ofv.GenRow("","","","","","","",Celda,"si")
    rs.movenext
  loop

  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")

    Agregar = ""

  'Botonera
  if bModificar then
    sAccion = "AgregarPropMasIns.asp?Tabla=Objprtyadv&volver=" & vuelta & "&Codobjeto=" & nCodigo
    sAccion = sAccion & "&Tipobjeto=" & sTipo
    Agregar = ofv.BotonAgregar(sAccion)

  end if



  'Boton Volver
  sAccion = "edPropiedades.asp?codobjeto=" & Ncodigo & "&tipobjeto=" & stipo & "  "
  sRetornar = ofv.BotonVolver(sAccion)



'--------------------------------------------------------------------------------
  sHTML = sHTML & "</div>"



    Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
    Fila = ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")


    'Estructura del Indice

    sHTML = sHTML & CorPnlIndx("N","")

    'sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)


'--------------------------------------------------------------------------------

 sHTML = sHTML & ofv.botoneraESSX("0",rs,CantRegAMover,CantRegAMostrar,"no","no",posi,"","","no",Agregar,sRetornar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & " </body></html>"
end if

response.write sHTML

%>

