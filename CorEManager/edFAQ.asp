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

  sPath = Request.ServerVariables("PATH_INFO")
  Xruta = Request.ServerVariables("SERVER_NAME") & ":" & Request.ServerVariables("SERVER_PORT") & left(sPath,instrRev(sPath,"/",-1))

 ' response.write  Xruta

  'Encabezados
  fila = ofv.Encabezados("DIRECCION URL","&nbsp;")

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "SELECT * FROM OBJFAQ  "
  Strsql = Strsql & " Where codobjeto=" & nCodigo
  Strsql = Strsql & " And tipobjeto = '" & stipo & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then

     strsql = "Insert into OBJFAQ  (codobjeto, tipobjeto, REDURL, SECUEN "
     strsql = strsql & ") Values (" & nCodigo & ",'" & sTipo & "'"
     strsql = strsql & ",'" & Xruta & "NULO.ASP','10'"
     strsql = strsql & ")"



     call ofv.crearconsultaEx(StrSql,cn,1,volver)

     Strsql = "SELECT * FROM OBJFAQ  "
     Strsql = Strsql & " Where codobjeto=" & nCodigo
     Strsql = Strsql & " And tipobjeto = '" & stipo & "'"
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  end if
IF NOT  rs.eof THEN
    sForm = "EDFAQ"

    'Eliminar

      Celda = ofv.GenCelda("","","","","","","","","si")


    'URL REDIRECT
    if bModificar then
      sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sForm,"120","dt","")
    else
      sInput = ofv.GenerarInput("",rs(3),"readonly","","120","dt","")
    end if
    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")


    sinput = "<a  CLASS=TT TARGET='_blank' href='" & TipoSite("")  & rs(3) & "' "
    sinput = sinput  & " ><b>TEST</b></a>"
    Celda = Celda & ofv.gencelda("","tt style='border-style:outset;border-width:2;' ","center","","15%","","",sInput,"si")




    if bModificar then
      sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=ObjFAQ&volver=" & vuelta
      sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
      Formulario = ofv.GenForm(sForm,"",sAccion,"","post","",Celda,"si")
    else
      Formulario = ofv.GenForm("","","","","post","",Celda,"si")
    end if



    Fila = fila & ofv.GenRow("","","","","","","",Formulario,"si")

  END IF

  sHTML = sHTML & "<br><br><br><br>"
  sHTML = sHTML & ofv.GenTabla("","aaa","","80%","","0","","0","0","","",fila,"si")

  'Botonera


'--------------------------------------------------------------------------------
  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

  sHTML = sHTML & ofv.botoneraESSX("0",rs,CantRegAMover,CantRegAMostrar,"no","no",posi,"","","no","no","no","")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & " </body></html>"
end if

response.write sHTML

%>