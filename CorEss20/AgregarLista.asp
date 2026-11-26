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
op = request.querystring("OP")
'response.write "STRING: " & request.querystring

'''''''''''''''''''''''''''''''''''''''''''''''''''

     sHTML = ofv.FormHeader("Agregar Usuario")
 'IF Indice <>"NO" THEN
'----------------------------------------------

 ' shtml = shtml & CorPnlFnc("N","")

  'sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"
'END IF
'-------------------------------------------------
     SHTML = SHTML & GenBuscaUsr("N","N","N","N",CHR(9),CHR(9),CHR(9),CHR(9))

     IF SESSION("FBUSR") <> "NO" THEN



        f1a = "EnvMail1.asp?borrar=no"

        sHTML = sHTML & ofv.GenForm("AgreArchb","",f1a,"","post","","","no")




        sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & SESSION("FBUSR") & "' NAME='" & OP & "' >"



      ELSE

            sAccion = "EnvMail1.asp?borrar=no"


         sHTML = sHTML & ofv.GenForm("AgreArchb","",sAccion,"","post","","","no")


      END IF

''''''''''''''''''''''''''''''''''''''''''''''''''''




  'Botones Aceptar y Cancelar

  sAccion = "EnvMail1.asp?borrar=no"

  sHTML = sHTML & ofv.AgregarBoton("AgreArchb",sAccion,"_self")


  Agregar = ""
  Buscar  = ""
  Celda   = ""


  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)
 ' IF Indice <>"NO" THEN

'--------------------------------------------------------------------------------
  'sHTML = sHTML & "</span></div>"

  'Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  'Fila = ofv.GenRow("","","","","","","",Celda,"si")
  'sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();'","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  'sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------
'END IF
  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>