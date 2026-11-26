<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  sCod = request.querystring("Cod")
  sCIRCUITO    = request.querystring("CIRCUITO")
  Volver = request.querystring("Volver")
  Volver = Volver & "?Cod=" & sCod & "_CIRCUITO=" & SCIRCUITO
  sxopc1 = ""


     session("RGPROC") = RGPROC

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

     sHTML = ofv.FormHeader("Agregar Usuario")

	'   sHTML = ofv.FormHeader(session("FormN"))

	 '----------------------------------------------

	   shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=1111>"

     SHTML = SHTML & GenBuscaUsr("N","N","S","S",CHR(9),CHR(9),"GRP-","USR-")

     IF SESSION("FBUSR") <> "NO" THEN

        set cn = ofv.conectar(ofv.strconn2)

        IF session("modalidad") = "PRGCOR" THEN

           strsql = "Select OBJID, * From ORRSKCIRCUITOSDKDET "

           WKSPTABLA = "ORRSKCIRCUITOSDKDET"


        ELSEIF session("modalidad") = "PRGESTADO" THEN

           strsql = "Select OBJID, * From ORESTADOSPER "

           WKSPTABLA = "ORESTADOSPER"


        ELSE




           strsql = "Select * From objsdkdet "

           WKSPTABLA = "OBJSDKdet"

       END IF


       set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)


        f1a = "asp/AgregarEN.asp?Tabla=" & WKSPTABLA
        f1a = f1a & "&volver=" & volver & "&Codobjeto=" & nCodigo
        f1a = f1a & "&Tipobjeto=" & sTipo

        IF session("modalidad") = "PRGESTADO" THEN
           f1a = "ORESTADOSSDKDET.ASP?OPCION=SI"
        end if

        sHTML = sHTML & ofv.GenForm("AgreArchb","",f1a,"","post","","","no")

        IF session("modalidad") = "PRGCOR" THEN
           sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & sCIRCUITO & "' NAME='codrsk' >"
        ELSEIF session("modalidad") = "PRGESTADO" THEN
           sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & sCIRCUITO & "' NAME='ESTADO' >"
        end if
        sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & SESSION("FBUSR") & "' NAME='" & RS(4).NAME & "' >"
        sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & sCod & "' NAME='" & RS(3).NAME & "' >"

        ' 02/08/2021 aca se agrega la nueva funcionalidad OBJMEMPALC -> ALE
		 '10317-0004/20

	    SELECT CASE session("APLMDCLI")
	     CASE "GP"
		  sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & session("SETIDEMPRE") & "' NAME='codemp' >"
		CASE ELSE
		  sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & SESSION("UN") & "' NAME='codemp' >"
	    END SELECT

       '-------------------------------------------------------------------------------------------


      ELSE

        sAccion = "edobjsdkdet.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Cod=" & sCod  & "&CIRCUITO=" & SCIRCUITO & "&volver=edobjsdk.asp"

        IF session("modalidad") = "PRGESTADO" THEN
           sAccion = "ORESTADOSSDKDET.ASP"
        end if



        sHTML = sHTML & ofv.GenForm("AgreArchb","",sAccion,"","post","","","no")


      END IF


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



XGUFILTRO =  request.form("XGUFILTRO")

if XGUFILTRO <> "TEFND" and XGUFILTRO <> "TE" then
  'Botones Aceptar y Cancelar
  sAccion = "edobjsdkdet.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Cod=" & sCod & "&volver=edobjsdk.asp" & "&CIRCUITO=" & SCIRCUITO

  IF session("modalidad") = "PRGESTADO" THEN
     sAccion = "ORESTADOSSDKDET.ASP"
  end if

  sHTML = sHTML & ofv.AgregarBoton("AgreArchb",sAccion,"_self")

end if
'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

 sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,"")
  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>