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
  sKIND = request.querystring("SKIND")

  IF SKIND <> "" THEN
     SESSION("SKIND") = SKIND
  ELSE
     SKIND = SESSION("SKIND")
  END IF


' http://server-41/coresV3/coress20/corexpress.asp?opcion=IX&opt=CC&id=103-MNN-1&durl=codigo*103_tipo*MN_perfil*OBJARC


  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

'-------------------------------------------------

  vuelta = ofv.vueltaasp2("../")

  'Encabezados
  fila = ""

  set cn = ofv.conectar(ofv.strconn2)


  Strsql = "SELECT * FROM TIPOBJETO  "
  Strsql = Strsql & " Where  tipobjeto = '" & stipo & "' "

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then




    sForm = "edarchivos" & rs(0)


    Celda = ofv.GenCelda("","TT","","20%","","","","URL script:","si")

    wkhosturl = request.servervariables("HTTP_REFERER")

    wkhosturlX = ucase(wkhosturl & "NO")

    wkhosturln = instr(wkhosturlX,"/COREMANAGER/")

    if wkhosturln > 0 then

       wkhosturl = left(wkhosturl,wkhosturln)


    else

       wkhosturln = instr(UCASE(request.servervariables("SCRIPT_NAME") & "NO"),"/COREMANAGER/")

       wkhosturl = left(request.servervariables("SCRIPT_NAME") & "NO",wkhosturln)

       'wkhosturl = TipoSite("") & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & wkhosturl
        if SESSION("URLPORT") ="N" then
         wkhosturl = TipoSite("") & request.servervariables("SERVER_NAME")  & wkhosturl

	    else
	      wkhosturl = TipoSite("") & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & wkhosturl
        end if


    end if

    WKURLCONS = wkhosturl



    Celda = Celda & ofv.GenCelda("","UTG","","80%","","","",WKURLCONS,"si")



    Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")




    Celda = ofv.GenCelda("","TT","","20%","","","","URL PARA CONSULTA:","si")

    WKURLCONS = WKURLCONS & "coress20/corexpress.asp?opt=CC-" & nCodigo & "-" & stipo

    Celda = Celda & ofv.GenCelda("","UTG","","80%","","","",WKURLCONS,"si")



    Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

   if session("APLMDCLI")="BLP" then
      WKURLCONS = wkhosturl
	    Celda = ofv.GenCelda("","TT","","20%","","","","URL ALTERNATIVA:","si")

	    WKURLCONS = WKURLCONS & "coress20/corexpress.asp?opt=CC-" & nCodigo & "-" & stipo & "&xBus=S"

	    Celda = Celda & ofv.GenCelda("","UTG","","80%","","","",WKURLCONS,"si")



	    Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

  END IF

 end if


  sHTML = sHTML & ofv.GenTabla("","aaa","","90%","","0","","0","0","","",fila,"si")



  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

    sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no","","no","")

'--------------------------------------------------------------------------------
  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------


  sHTML = sHTML & " </body></html>"
end if

response.write sHTML

%>