<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  sCod = request.querystring("Cod")
  Volver = request.querystring("Volver")
  OPWK = request.querystring("OPWK")
   sorden = REQUEST.QUERYSTRING("orden")
   tipobjeto= REQUEST.QUERYSTRING("tipobjeto")
  ID=REQUEST.QUERYSTRING("ID")

  Volver = Volver & "?Cod=" & sCod & "_CIRCUITO=" & SCIRCUITO & "_OPWK=" & OPWK
  sxopc1 = ""
'response.write NOW() & "<br>"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

     sHTML = ofv.FormHeader("Agregar Usuario")
if OPWK = "REASIGNAR" then

       shtml = shtml & CorPnlFnc("N","")
'response.write "ID1.1:" & ID & "<br>"
	   sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

end if
'response.write "ID1.2: " & ID & "<br>"
idx2 = ID
     SHTML = SHTML & GenBuscaUsr("N","N","N","S",CHR(9),CHR(9),"GRP-","USR-")

'response.write "ID1.3: " & idx2 & "<br>"


     sAccion = "WKFREASIGNADOR.ASP?OPCION=SETUSR" & "&OPWK=" & OPWK & "&id=" & idx2 & "&TIPOBJETO=" & TIPOBJETO

     sHTML = sHTML & ofv.GenForm("AgreArchb","",sAccion,"","post","","","no")

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  'Botones Aceptar y Cancelar
  'response.write "ID2: " & idx2 & "<br>"
     sAccion = "WKFREASIGNADOR.ASP?OPCION=NO" & "&OPWK=" & OPWK & "&id=" & idx2 & "&TIPOBJETO=" & TIPOBJETO

	 sHTML = sHTML & ofv.AgregarBoton("AgreArchb",sAccion,"_self")

'--------------------------------------------------------------------------------
     sHTML = sHTML & "</span></div>"



     Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
     Fila = ofv.GenRow("","","","","","","",Celda,"si")
     sHTML = sHTML &  ofv.GenTabla("TCEX1"," onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

     'Estructura del Indice
if OPWK = "REASIGNAR" then
     sHTML = sHTML & CorPnlIndx("N","")


'--------------------------------------------------------------------------------"
 sHTML = sHTML & "<script language='javascript' > "
       IF REQUEST.FORM("XGUMASTER")="" THEN
         sHTML = sHTML & " clactfrm('U','G','NO','NO','NO');document.XGUFRM.submit(); "
       END IF
     sHTML = sHTML & "</script> "
end if
  sHTML = sHTML & "</body></html>"

end if



response.write sHTML

%>