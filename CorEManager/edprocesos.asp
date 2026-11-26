<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  XXXX = GENLOGEXT("EDIDET")

  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  if session("Nomodificar") = "S" then bModificar = False

  nCodigo = request.querystring("Codobjeto")
  sTipo   = request.querystring("Tipobjeto")
  sopcion = request.querystring("opcion")
  sACTCOD = request.querystring("ACTCOD")
  sACTID  = request.querystring("ACTID")

  sHTML  = ofv.FormHeader(session("FormN"))
  vuelta = ofv.vueltaasp2("../")




  set cn = ofv.conectar(ofv.strconn2)

  IF SOPCION = "SI" THEN

     Strsql = "UPDATE Objeven SET CODPROCESO = '" & SACTCOD & "' Where OBJID= 0" & sACTID
     CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

  END IF


    Strsql = "Select * From csgenobjetos "
    strsql = strsql & " where codobjeto = 0" & nCodigo
 
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if Not rs.EOF then

       Celda = ""
       sInput = "Eventos de:&nbsp;&nbsp;" & rs(1)
       Celda = Celda & ofv.GenCelda("","ttt","","","","","",sInput,"si")

       Fila = ofv.genrow("","","","","25","","",Celda,"si")

       sHTML = sHTML & ofv.GenTabla("","","","80%","","0","0","0","0","","",Fila,"si") & "<br>"

     end if



  'Encabezados

    Celda = ""
    Celda = Celda & ofv.GenCelda("","ttr","","","","","","#","si")
    Celda = Celda & ofv.GenCelda("","ttr","","","","","","Descripcion","si")
    Celda = Celda & ofv.GenCelda("","ttr","","","","","","Proceso","si")
    Celda = Celda & ofv.GenCelda("","ttr","","","","","","Actualizar","si")
    Formulario = ofv.GenRow("","","","","","","",Celda,"si")





  Strsql = "Select * From Objeven Where codobjeto=" & nCodigo & " Order By secuen"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then

    Celda = ""
    Celda = Celda & ofv.GenCelda("","ut COLSPAN=4 ","center","","","","","EL DOCUMENTO NO POSEE EVENTOS ASOCIADOS","si")
    Formulario = Formulario & ofv.GenRow("","","","","60","","",Celda,"si")


  else



  do until rs.eof
    sForm = "edEventos" & rs(0)

    Celda = ""

    'Secuencia
    sInput = rs(4)
    Celda = Celda & ofv.GenCelda("","ut","","2%","","","",sInput,"si")

    'Descripcion
    sInput = rs(3)
    Celda = Celda & ofv.GenCelda("","ut","","40%","","","",sInput,"si")

    sInstrumento = "----------"

    'Proceso
    if  cstr(rs(11) & "") <> "" then
            Strsql = "Select DSCPROC From CSPROCESOS Where CODPROCESO = '" & rs(11) & "' "
            set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
            if  not rsInstrumento.eof then sInstrumento = rsInstrumento(0)
            call ofv.cerrarconsulta(rsInstrumento)
    end if
    sInput = sInstrumento
    Celda = Celda & ofv.GenCelda("","ut","","50%","","","",sInput,"si")


    soldproc = rs(11) & ""

    if soldproc = "" then soldproc = "NO"

    SHREF = "edgesproc.asp?opcion=ACT&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&OLDCOD=" & soldproc & "&CODEVEN=" & rs(0)

    sImagenP = "&#0232;"
    sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
    sinput = "<a class='ft1n' " & sFuncAnc & " href='" & SHREF & "' "
    sinput = sinput  & " title='Actualizar proceso para " & rs(3) & "' >" & sImagenP & "</a>" 
    Celda  = Celda & ofv.gencelda("","ft","center","","1%","","",sInput,"si")


    Fila = ofv.GenRow("","","","","","","",Celda,"si")

    if bModificar then
      sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objeven&volver=" & vuelta
      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",Fila,"si")
    else
      Formulario = Formulario & ofv.GenForm("","","","","post","",Fila,"si")
    end if

    rs.movenext
  loop

  end if

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & ofv.GenTabla("","","","90%","","0","","0","0","","",Formulario,"si")

  'Botonera
   Agregar = ""

    sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no","","","","no",Agregar,"no","")



  sHTML = sHTML & " </body></html>"
end if

response.write sHTML

%>