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
  sTipo   = request.querystring("Tipobjeto")
  sSubTit = session("SubTit")
 
  sOpcion = request.querystring("OPC")

  if sOpcion <> "" then sOpcion = Ucase(sOpcion) 

  sFILAI   = request.querystring("FILA")
  sCELDAI  = request.querystring("CELDA")
  sTIPOI   = request.querystring("TIPO")


  vuelta   = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover)
  vuelta   = vuelta & "_FILA=" & sFILAI  & "_CELDA=" & sCELDAI  & "_TIPO=" & sTIPOI
 

  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

'-------------------------------------------------

  set cn   = ofv.conectar(ofv.strconn2)

  strsql   = "SELECT * From Objtabcmb Where tabobj = '" & sFILAI  & "' AND TABCOL = " & sCELDAI
  set rs   = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if  rs.eof then

     strsql = "INSERT INTO Objtabcmb (CODOBJETO,TIPOBJETO,TABBLK,TABOBJ,TABTIPO,TABCOL,CMBCOL,CMBROW,CMBSTYL)"
     strsql = strsql & " VALUES ('" & nCodigo & "','" & sTipo & "','" & sSubTit & "',"
     strsql = strsql & "'" & sFILAI & "','" & sTIPOI & "','" & sCELDAI & "','N','N','N') "
     CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

    call ofv.cerrarconsulta(rs)

     strsql = "SELECT * From Objtabcmb Where tabobj = '" & sFILAI  & "' AND TABCOL = " & sCELDAI
   ''  response.write strsql & "no hay<br>"  
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


  END IF




  if not rs.eof then
    do until rs.eof

     sForm = "edTablas" & rs(0)
     sid   = rs(0)

     if bModificar then
     sAccion = "asp/ActualizarEN.asp?ID=" & RS(0) & "&Tabla=OBJTABCMB&volver=" & vuelta
     sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
     end if

     'Fila = ""

      Celda1 = ofv.GenCelda("","tt","","30%","","","","COMBINA CELDAS","si")

    scomK = "2" & chr(9) & "3" & chr(9) & "4" & chr(9) & "5" & chr(9) & "99" & chr(9) & "N"
    scomD = "2" & chr(9) & "3" & chr(9) & "4" & chr(9) & "5" & chr(9) & "TODAS" & chr(9) & "NO"


    if bModificar then
        sInput = ofv.gencomboFX(sForm & "X3",rs(7).name,rs(7),scomK,scomD," ")
        sInput = ofv.GenForm(sForm & "X3","",sAccion,"","post","",sInput,"si")
    else
        sInput = ofv.GenerarInput("",rs(7),"readonly","","10","dt","")
    end if
    Celda =  Celda & ofv.gencelda("","vti","","70%","","","",sInput,"si")

      'Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")


      Celda1 =Celda1 & ofv.GenCelda("","tt","","","","","","COMBINA FILAS","si")

    scomK = "2" & chr(9) & "3" & chr(9) & "4" & chr(9) & "5" & chr(9) & "99" & chr(9) & "N"
    scomD = "2" & chr(9) & "3" & chr(9) & "4" & chr(9) & "5" & chr(9) & "TODAS" & chr(9) & "NO"


    if bModificar then
        sInput = ofv.gencomboFX(sForm& "X2",rs(8).name,rs(8),scomK,scomD," ")
        sInput = ofv.GenForm(sForm & "X2","",sAccion,"","post","",sInput,"si")
    else
        sInput = ofv.GenerarInput("",rs(8),"readonly","","10","dt","")
    end if
    Celda =  Celda & ofv.gencelda("","vti","","","","","",sInput,"si")

     '' Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

      Celda1 = Celda1 & ofv.GenCelda("","tt","","","","","","FORMATO","si")

    scomK = "S" & chr(9) & "N"
    scomD = "ESPECIAL" & chr(9) & "DEFAULT"


    if bModificar then
        sInput = ofv.gencomboFX(sForm & "X1",rs(9).name,rs(9),scomK,scomD," ")
        sInput = ofv.GenForm(sForm & "X1","",sAccion,"","post","",sInput,"si")
    else
        sInput = ofv.GenerarInput("",rs(9),"readonly","","10","dt","")
    end if
    Celda =  Celda & ofv.gencelda("","vti","","","","","",sInput,"si")

      'Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")




    '  if bModificar then
    '    sAccion = "asp/ActualizarEN.asp?ID=" & RS(0) & "&Tabla=OBJTABCMB&volver=" & vuelta
    '    sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
    '    Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",Fila,"si")
    '  else
    '    Formulario = Formulario & ofv.GenForm("","","","","post","",Fila,"si")
    '  end if

    Formulario = Formulario & ofv.GenRow("","","","","","","",Celda,"si")

    rs.movenext
    loop



    END IF
    sHTML = sHTML & ofv.GenTabla("","aaa","","80%","","0","","0","0","","",Celda1 & Formulario,"si")

    call ofv.cerrarconsulta(rs)

      Agregar = ""


  IF  sTIPOI = "1" THEN
      sAccion  = "EdTablasBLKTIT.asp?Tabla=" & sTabla & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      sAccion  = sAccion & "&SubTit=" & sSubTit
  ELSE
      sAccion  = "EdTablasBLKCON.asp?Tabla=" & sTabla & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      sAccion  = sAccion & "&SubTit=" & sSubTit
  END IF 

  sVolver  = ofv.BotonVolver(sAccion)


  

'--------------------------------------------------------------------------------
  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1"," onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,sVolver,"")


  call ofv.cerrarconn(cn)

  sHTML = sHTML & " </body></html>"
end if

response.write sHTML

%>