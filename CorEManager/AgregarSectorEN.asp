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

     sHTML = ofv.FormHeader("Agregar sectores")
 IF Indice <>"NO" THEN
'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"
END IF
'-------------------------------------------------
     SHTML = SHTML & GenBuscaUsr("N","S","N","N",CHR(9),CHR(9),CHR(9),CHR(9))

     IF SESSION("FBUSR") <> "NO" THEN

        set cn = ofv.conectar(ofv.strconn2)

        strsql = "Select * From objsector "
        set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

        f1a = "asp/AgregarEN.asp?Tabla=objsector"
        f1a = f1a & "&volver=" & volver & "&Codobjeto=" & nCodigo
        f1a = f1a & "&Tipobjeto=" & sTipo

        sHTML = sHTML & ofv.GenForm("AgreArchb","",f1a,"","post","","","no")

        if session("CKTTYPE")="" then session("CKTTYPE")="AUTO"


        sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & SESSION("FBUSR") & "' NAME='" & RS(4).NAME & "' >"
        sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='O' NAME='" & RS(3).NAME & "' >"
      '  sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & nCodigo & "' NAME='codobjeto' >"
        'sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & sTipo & "' NAME='tipobjeto' >"
        'session("CKTTYPE")=""
        call ofv.cerrarconsulta(rs)
        call ofv.cerrarconn(cn)


      ELSE


            sAccion = "edsectores.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo


         sHTML = sHTML & ofv.GenForm("AgreArchb","",sAccion,"","post","","","no")


      END IF

''''''''''''''''''''''''''''''''''''''''''''''''''''




  'Botones Aceptar y Cancelar
  sAccion = "edsectores.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
  sHTML = sHTML & ofv.AgregarBoton("AgreArchb",sAccion,"_self")


  Agregar = ""
  Buscar  = ""
  Celda   = ""


  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)
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