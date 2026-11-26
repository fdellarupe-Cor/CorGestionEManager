<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
  ncaller = request.querystring("caller")
  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  sCod = request.querystring("Cod")
  INDICE=request.querystring("Indice")
  Volver = request.querystring("Volver")
  Volver = Volver & "?Cod=" & sCod
  sxopc1 = ""



'''''''''''''''''''''''''''''''''''''''''''''''''''

     sHTML = ofv.FormHeader("Agregar Usuario")
 IF Indice <>"NO" THEN
'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"
END IF
'-------------------------------------------------
'RESPONSE.WRITE session("APLMDCLI")
    if  session("APLMDCLI") ="GP" THEN
     SHTML = SHTML & GenBuscaUsr("N","N","N","N",CHR(9),CHR(9),CHR(9),CHR(9))
    ELSE
     SHTML = SHTML & GenBuscaUsr("N","N","N","S",CHR(9),CHR(9),CHR(9),CHR(9))
    END IF


     IF SESSION("FBUSR") <> "NO" THEN

        set cn = ofv.conectar(ofv.strconn2)

        strsql = "Select * From objwfldet "
        set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

        f1a = "asp/AgregarEN.asp?Tabla=OBJwflDET"
        f1a = f1a & "&volver=" & volver & "&Codobjeto=" & nCodigo
        f1a = f1a & "&Tipobjeto=" & sTipo

        sHTML = sHTML & ofv.GenForm("AgreArchb","",f1a,"","post","","","no")

        if session("CKTTYPE")="" then session("CKTTYPE")="AUTO"


        sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & SESSION("FBUSR") & "' NAME='" & RS(4).NAME & "' >"
        sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='0' NAME='" & RS(3).NAME & "' >"
        sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & session("CKTTYPE") & "' NAME='CKTTYPE' >"
        'session("CKTTYPE")=""
        call ofv.cerrarconsulta(rs)
        call ofv.cerrarconn(cn)


      ELSE

         if ncaller = "AUT" then
            sAccion = "AutDefaultdet.asp"
         else
            sAccion = "edobjwfldet.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Cod=" & sCod & "&volver=edobjwfl.asp"
         end if

         sHTML = sHTML & ofv.GenForm("AgreArchb","",sAccion,"","post","","","no")


      END IF

''''''''''''''''''''''''''''''''''''''''''''''''''''


XGUFILTRO = request.form("XGUFILTRO")
'response.write request.form
if XGUFILTRO <> "TEFND" and XGUFILTRO <> "TE" then
  'Botones Aceptar y Cancelar
  if ncaller = "AUT" then
  sAccion = "AutDefaultdet.asp"
  else
  sAccion = "edobjwfldet.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Cod=" & sCod & "&volver=edobjwfl.asp"
  end if
  sHTML = sHTML & ofv.AgregarBoton("AgreArchb",sAccion,"_self")
end if

 sAccion = "AutDefaultDet.asp"
  'volver = "<table width=100% border ><tr><td>" & ofv.BotonVolver(sAccion) & "</td></tr></table>"
  volver = ofv.BotonVolver(sAccion)

  Agregar = ""
  Buscar  = ""
  Celda   = ""


  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,volver)
  IF Indice <>"NO" THEN

'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------
END IF

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>