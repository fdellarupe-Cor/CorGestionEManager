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

     session("comp31") = "TODOS"
     session("comp32") = "HTM"

  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  bModificaresp = ofv.AModificar
  if session("Nomodificar") = "S" then

     if LEFT(Session("Form2"),4) <> "DESA" then    bModificar = False

  end if

  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  sKIND = request.querystring("SKIND")

  IF SKIND <> "" THEN
     SESSION("SKIND") = SKIND
  ELSE
     SKIND = SESSION("SKIND")
  END IF 

  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

'-------------------------------------------------


  vuelta = ofv.vueltaasp2("../")

  'Encabezados
  fila = ofv.Encabezados("T","Descripcion","ACTIVO","Orden")

  set cn = ofv.conectar(ofv.strconn2)


  Strsql = "SELECT * FROM objacd  "
  Strsql = Strsql & " Where codobjeto=" & nCodigo
  Strsql = Strsql & " And tipobjeto = '" & stipo & "' "

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if  rs.eof then

      Strsql = "insert into objacd  (codobjeto,tipobjeto,tipoacd,codobjacd,tipobjacd,activo,descripcion,secuen) values ("
      Strsql = Strsql & nCodigo & ",'" & stipo & "','P','000',"
      Strsql = Strsql & "'FLDPRI','N','DEFAULT','10')"
      call ofv.crearconsultaEx(StrSql,cn,1,parametros)

      Strsql = "insert into objacd  (codobjeto,tipobjeto,tipoacd,codobjacd,tipobjacd,activo,descripcion,secuen) values ("
      Strsql = Strsql & nCodigo & ",'" & stipo & "','S','000',"
      Strsql = Strsql & "'FLDPRE','N','DEFAULT','10')"
      call ofv.crearconsultaEx(StrSql,cn,1,parametros)


      Strsql = "insert into objacd  (codobjeto,tipobjeto,tipoacd,codobjacd,tipobjacd,activo,descripcion,secuen) values ("
      Strsql = Strsql & nCodigo & ",'" & stipo & "','R','000',"
      Strsql = Strsql & "'FLDREL','N','DEFAULT','10')"
      call ofv.crearconsultaEx(StrSql,cn,1,parametros)



      Strsql = "SELECT * FROM objacd  "
      Strsql = Strsql & " Where codobjeto=" & nCodigo
      Strsql = Strsql & " And tipobjeto = '" & stipo & "' "
      set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  end if


  do until rs.eof
    sForm = "edarchivos" & rs(0)

    if bModificar  then
         sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objacd&volver=" & vuelta
         sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
    end if  


    'Eliminar
    if bModificar  then
      sHref = "asp/EliminarEN.asp?ID=" & rs(0) & "&Tabla=Objacd&volver=" & vuelta
      sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
    end if
    if sHref <> "" then
      Celda = ofv.BotonEliminar("acceso directo",sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","","si")
    end if

    'ICON

    if  RS(3) = "P" OR RS(3) = "S" OR RS(3) = "R" then
        'sxiconv = "<font  face='corflaticon' size=4 color='#F7AA22' >&#0204;</font>"
		sxiconv = "<i class='las la-folder' style='color:"& Session("CMFPNTIT") &"' ></i>"
    elseif  RS(3) = "C" then
        sxiconv = "<font  face='corflaticon' size=4 color='#336699' >&#0204;</font>"
    else
        sxiconv = "<font  face='corflaticon' size=4 color='#993333' >&#0158;</font>"
    end if

    Celda = Celda & ofv.GenCelda("","UTS","center","1%","","","",sxiconv,"si")

    if  RS(3) = "P" then
        sxiconv = "CARPETA DEL DOCUMENTO."
    elseif  RS(3) = "S" then
        sxiconv = "CARPETA SUPERIOR."
    elseif  RS(3) = "R" then
        sxiconv = "CARPETA DEPENDIENTE."
    else
        sxiconv = "<font  face='corflaticon' size=4 color='#993333' >&#0158;</font>"
    end if
    Celda = Celda & ofv.GenCelda("","UT","","40%","","","",sxiconv,"si")

   'ACTIVO
    if  RS(3) = "P" OR RS(3) = "S" OR RS(3) = "R" then
        scomK = "N" & chr(9) & "S"
        scomD = "NO" & chr(9) & "SI"

    if bModificar  then
           sInput = ofv.gencomboFX(sForm & "X5",rs(6).name,rs(6),scomK,scomD," ")
           sInput = ofv.GenForm(sForm & "X5","",sAccion,"","post","",sInput,"si")
        else
           sInput = ofv.GenerarInput("",rs(6),"readonly","","2","dt","")
        end if
        sxiconv = sInput
    else
        sxiconv = "SI"
    end if
    Celda = Celda & ofv.GenCelda("","UT","","10%","","","",sxiconv,"si")

    'Secuencia
     if bModificar  then
      sInput = ofv.GenerarInput(rs(8).name,rs(8),"text",sForm & "X3","5","dtn","SI")
      sInput = ofv.GenForm(sForm & "X3","",sAccion,"","post","",sInput,"si")
    else
      sInput = ofv.GenerarInput("",rs(8),"readonly","","5","dtn","SI")
    end if
    Celda = Celda & ofv.GenCelda("","Ut","","5%","","","",sInput,"si")


''   if bModificar  then
''     sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objacd&volver=" & vuelta
''     sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
''     Formulario = ofv.GenForm(sForm,"",sAccion,"","post","",Celda,"si")
''   else
''     Formulario = ofv.GenForm("","","","","post","",Celda,"si")
''   end if

    'Fila = fila & ofv.GenRow("","","","","","","",Formulario,"si")

  Formulario = Formulario & ofv.GenRow("","","","","","","",Celda,"si")

    rs.movenext
  loop


  sHTML = sHTML & ofv.GenTabla("","aaa","","90%","","0","","0","0","","",fila & Formulario,"si")

  'Botonera
    if bModificar  then
    sAccion = "Agregaracd.asp?Tabla=Objacd&volver=" & vuelta & "&Codobjeto=" & nCodigo
    sAccion = sAccion & "&Tipobjeto=" & sTipo
  '  Agregar = ofv.BotonAgregar(sAccion)
    Agregar = ""
    xxxx = ofv.botoneraESSX("0","","","","","","","","","no",Agregar,"no","")
  end if

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

'--------------------------------------------------------------------------------
  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------
sHTML = sHTML & xxxx

  sHTML = sHTML & " </body></html>"
end if

response.write sHTML

%>