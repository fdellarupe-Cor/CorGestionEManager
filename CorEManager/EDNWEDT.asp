<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%
CONT = 0
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
  sTipo = request.querystring("Tipobjeto")

  sHTML = ofv.FormHeader(session("FormN"))

  histobs = request.querystring("histobs")

  if histobs = "S" then

  niv1x=request.querystring("NIV1")

  session("codSol")=niv1x

  end if

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;margin: 5 15px" & styl & "' ><span id=111>"

'-------------------------------------------------


  vuelta = ofv.vueltaasp2("../")

  'Encabezados
  fila = ofv.Encabezados("Orden","Titulo del Indice","DETALLE")

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From OBJNWEDT Where codobjeto=" & nCodigo & " Order By consecuen"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then
     xxxx = NWEDTVER()
  end if
  do until rs.eof
  CONT  = CONT + 1
    sForm = "edTablas" & rs(0)

     if bModificar then
      sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=OBJNWEDT&volver=" & vuelta
      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo

    end if

    'Eliminar
    sHref = ""
    if bModificar then
      sHref = "asp/EliminarEN.asp?ID=" & rs(0) & "&Tabla=OBJNWEDT&volver=" & vuelta
      sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
    end if
    if sHref <> "" then
      Celda = ofv.BotonEliminar(rs(3),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","","si")
    end if

    'Secuencia
    if bModificar then
      sInput = ofv.GenerarInput(rs(4).name,rs(4),"text",sForm & "x1","3","dtn","SI")
      sInput = ofv.GenForm(sForm & "x1","",sAccion,"","post","",sInput,"si")
    else
      sInput = rs(4)
    end if
    Celda = celda & ofv.GenCelda("","VTI","","","","","",sInput,"si")

    'Descripcion
    if bModificar then
      sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sForm & "x2","80","dt","")
      sInput =ofv.GenForm(sForm & "x2","",sAccion,"","post","",sInput,"si")
    else
      sInput = rs(3)
    end if
    Celda =  celda & ofv.GenCelda("","VTI","","","","","",sInput,"si")


    'DETALLE

    sImagenP = "<b><i class='las la-arrow-right' ></i></b>"

    sinput = "<a class=nt href='EDNWEDTDET.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&concod=" & rs(5)
    sinput = sinput  & "' title='Propiedades de " & rs(1) & "' >" & sImagenP & "</a>"
    Celda = Celda & ofv.gencelda("","vt ","center","","15%","","",sInput,"si")



    Fila = fila & ofv.GenRow("","","","","","","",celda,"si")

    rs.movenext
  loop

  'sHTML = sHTML & ""
  'sHTML = sHTML & ofv.GenTabla("","aaa","","80%","","0","","0","0","","",fila,"si")
  if CONT = 0 then
    	    sHTML = sHTML & "<div class='Sininfo' align='center'><i class='las la-broadcast-tower' style='margin-right: 0.3rem;'></i>No tienes consideraciones generales agregadas.</div>"
    	else
    	    tabla = ofv.GenTabla("","aaa","","80%","","0","","0","0","","",fila,"si")
    	    sHTML = sHTML & tabla
	end if

    Agregar = ""

  'Botonera
  if bModificar then
    sAccion = "AgregarNWEDT.asp?Tabla=OBJNWEDT&volver=" & vuelta & "&Codobjeto=" & nCodigo
    sAccion = sAccion & "&Tipobjeto=" & sTipo
    Agregar = ofv.BotonAgregar(sAccion)

  end if



'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1"," onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

 sHTML = sHTML & ofv.botoneraESSX("0",rs,CantRegAMover,CantRegAMostrar,"no","no",posi,"","","no",Agregar,"no","")


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & " </body></html>"
end if

response.write sHTML

function NWEDTVER()

  XNVER = "N"
  XNSEC = 0

  Xfer = server.mappath(request.servervariables("SCRIPT_NAME"))
  XferTarget0 = mid(Xfer,1,instrrev(Xfer,"\",-1)) & "CarpetasCMD\"
  XferTarget1 = mid(Xfer,1,instrrev(Xfer,"\",-1)) & "carpetaObjetivos\"
  XferTarget2 = mid(Xfer,1,instrrev(Xfer,"\",-1)) & "CarpetaAlcance\"
  XferTarget3 = mid(Xfer,1,instrrev(Xfer,"\",-1)) & "CarpetaConsideracion\"

  abrir1 = XferTarget1 & RtfObjetivo
  abrir2 = XferTarget2 & RtfAlcance
  abrir3 = XferTarget3 & RtfConsid

  set fff = server.createobject("scripting.filesystemobject")




'''''''''''''' Objetivo
''''
  Strsql = "Select ARCHIVO,titulo From ObjMain where CodObjeto=" & nCodigo & " and TipObjeto='"
  Strsql = Strsql & sTipo & "'"
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs1.eof then

    XNVER = "S"
    XNSEC = XNSEC + 10
    RtfObjetivo=rs1(0)
    TITObjetivo=rs1(1)

    XNNWCOD = GETNWEN(nCodigo,sTipo) & XNSEC

    XNABRIR = XferTarget1 & RtfObjetivo

    xarchivo = ofv.convertircar(XNABRIR,chr(34)," ")
    xxx = ""
    if fff.FileExists(trim(Xarchivo)) then
       Set fff1 = fff.openTextFile(trim(Xarchivo), 1, -1)
       do while fff1.AtEndOfstream <> true
          xxx = xxx & fff1.readline
       loop
       fff1.close
       set XNPSR = server.createobject("CorRtfParser.RtfParser")
       XNPSRFL=""
       if xxx <> "" then
        XNPSRFL   =  XNPSR.RtfToHtml(xxx)
       end if
       Set MyFile = fff.CreateTextFile(XferTarget0 & XNNWCOD & ".CSE")
       MyFile.Write XNPSRFL
       MyFile.Close

       strsql = "insert into OBJNWEDT (codobjeto,tipobjeto,contitulo,consecuen,concodigo) values(" & nCodigo & ",'"
       Strsql = Strsql & sTipo & "','" &  TITObjetivo & "'," &  XNSEC & ",'" &  XNNWCOD & "')"
       call ofv.crearconsultaEx(StrSql,cn,1,parametros)

    end if
  end if
  call ofv.cerrarconsulta(rs1)

'''''''''''''' Alcance
''''
  Strsql = "Select ARCHIVO,titulo From ObjAlcance where CodObjeto=" & nCodigo & " and TipObjeto='"
  Strsql = Strsql & sTipo & "'"
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs1.eof then

    XNVER = "S"
    XNSEC = XNSEC + 10
    RtfObjetivo=rs1(0)
    TITObjetivo=rs1(1)

    XNNWCOD = GETNWEN(nCodigo,sTipo) & XNSEC

    XNABRIR = XferTarget2 & RtfObjetivo

    xarchivo = ofv.convertircar(XNABRIR,chr(34)," ")
    xxx = ""
    if fff.FileExists(trim(Xarchivo)) then
       Set fff1 = fff.openTextFile(trim(Xarchivo), 1, -1)
       do while fff1.AtEndOfstream <> true
          xxx = xxx & fff1.readline
       loop
       fff1.close
       set XNPSR = server.createobject("CorRtfParser.RtfParser")
       XNPSRFL=""
       if xxx <> "" then
         XNPSRFL   =  XNPSR.RtfToHtml(xxx)
       end if
       Set MyFile = fff.CreateTextFile(XferTarget0 & XNNWCOD & ".CSE")
       MyFile.Write XNPSRFL
       MyFile.Close

       strsql = "insert into OBJNWEDT (codobjeto,tipobjeto,contitulo,consecuen,concodigo) values(" & nCodigo & ",'"
       Strsql = Strsql & sTipo & "','" &  TITObjetivo & "'," &  XNSEC & ",'" &  XNNWCOD & "')"
       call ofv.crearconsultaEx(StrSql,cn,1,parametros)

    end if
  end if
  call ofv.cerrarconsulta(rs1)

'''''''''''''' Consideraciones generales
''''
  Strsql = "Select ARCHIVO,titulo From ObjConsid where CodObjeto=" & nCodigo & " and TipObjeto='"
  Strsql = Strsql & sTipo & "'"
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs1.eof then

    XNVER = "S"
    XNSEC = XNSEC + 10
    RtfObjetivo=rs1(0)
    TITObjetivo=rs1(1)

    XNNWCOD = GETNWEN(nCodigo,sTipo) & XNSEC

    XNABRIR = XferTarget3 & RtfObjetivo

    xarchivo = ofv.convertircar(XNABRIR,chr(34)," ")
    xxx = ""
    if fff.FileExists(trim(Xarchivo)) then
       Set fff1 = fff.openTextFile(trim(Xarchivo), 1, -1)
       do while fff1.AtEndOfstream <> true
          xxx = xxx & fff1.readline
       loop
       fff1.close
       set XNPSR = server.createobject("CorRtfParser.RtfParser")
       XNPSRFL=""
       if xxx <> "" then
        XNPSRFL   =  XNPSR.RtfToHtml(xxx)
       end if
       Set MyFile = fff.CreateTextFile(XferTarget0 & XNNWCOD & ".CSE")
       MyFile.Write XNPSRFL
       MyFile.Close

       strsql = "insert into OBJNWEDT (codobjeto,tipobjeto,contitulo,consecuen,concodigo) values(" & nCodigo & ",'"
       Strsql = Strsql & sTipo & "','" &  TITObjetivo & "'," &  XNSEC & ",'" &  XNNWCOD & "')"
       call ofv.crearconsultaEx(StrSql,cn,1,parametros)

    end if
  end if
  call ofv.cerrarconsulta(rs1)

  if  XNVER = "S" then

      Strsql = "Select * From OBJNWEDT Where codobjeto=" & nCodigo & " Order By consecuen"
      set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  end if

end function

%>