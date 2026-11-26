<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFMANAGER.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


'''  sHTML = ofv.MenuHeader("","")

  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

'-------------------------------------------------

  CantRegMostrarx = 12
  CantRegMoverx = 0
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn5)
  set cn4 = ofv.conectar(ofv.strconn6)
  set cn3 = ofv.conectar(ofv.strconn2)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  WKFAUTOB = "WKF00016631121722600921928"
  WKFAPROB = "WKF00016631121722600921928"
  WKFAUTO  = "WKF00016631121722600921928"
  WKFAPRO  = "WKF00016631121722600921928"

  DESLEVL = 9
  SUPLEVL = 9


     Strsql = "Select * From TIPOBJETO where TIPOBJETO = '" & session("TCliente") & "' "
     set rs = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if not rs.eof then
        WKFAUTO = RS(19)
        WKFAPRO = RS(20)
        DESLEVL = 9
        SUPLEVL = RS(18)
     END IF


if (WKFAUTO & "") = "" then WKFAUTO = WKFAUTOB
if (WKFAPRO & "") = "" then WKFAPRO = WKFAPROB



 sopcion = request.querystring("opcion")
 sACCION = request.querystring("ACCION")
 act = request.querystring("act")

if nFecini < clng(cdate("01-01-1900")) then
   nFecini = clng(cdate("01-01-1900"))
else
   nFecini = 0
end if

vNuevoValor = clng(DATE()) - nFecini

''RESPONSE.WRITE SOPCION & " " & SACCION

IF SOPCION = "" THEN SOPCION = "CONF"


 SELECT CASE SOPCION

        CASE "CONF"
           Celda = ofv.gencelda("","Vt colspan=4 ","center","30%","","","","&nbsp;","si")

           Fila = ofv.GenRow("","","","","","","",Celda,"si")




           sInput = ofv.GenerarInput("boton1","ANULA CANCELACION","submit","edPropiedades2","5","bt style='background-color:#993333;' ","")
           sAccion = "EDCANCELWkf.asp?OPCION=NO"
           sInput = ofv.GenForm("edPropiedades2","",sAccion,"","post","",sInput,"si")
           Celda = ofv.gencelda("","Vt colspan=2 style='BORDER-STYLE:RIDGE;BORDER-WIDTH:2;BORDER-COLOR:#EE0000;' ","center","30%","","","",sInput,"si")

           sInput = ofv.GenerarInput("boton1","CONFIRMA","submit","edPropiedades3","5","bt style='background-color:#005500;' ","")
           sAccion = "EDCANCELWkf.asp?OPCION=OK"
           sInput = ofv.GenForm("edPropiedades3","",sAccion,"","post","",sInput,"si")
           Celda = celda & ofv.gencelda("","Vt colspan=2 style='BORDER-STYLE:RIDGE;BORDER-WIDTH:2;BORDER-COLOR:#00EE00;' ","center","30%","","","",sInput,"si")

           Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

           sHTML = sHTML & "<br><br><br><br><br>"
           sHTML = sHTML & ofv.GenTabla("","aaa style='BORDER-STYLE:RIDGE;BORDER-WIDTH:2;BORDER-COLOR:#F7AA22;' ","","80% border=0 ","","0","0","1","0","","",Fila,"si")


        CASE "NO"
             IF (WKAPLIC = "OR" OR  WKAPLIC = "PL") AND session("TCliente") = "RSK" THEN
                RESPONSE.REDIRECT "edobjwflor.asp?PASO=AUTO&Tipobjeto=" & session("TCliente") & "&Codobjeto=" & session("NCliente") & "&GLBSOLN=ESPWFL"
             ELSE
                RESPONSE.REDIRECT "edObjWfl.asp"
             END IF 

        CASE "OK"
             SESSION("WKF") = WKFAPRO
             TITULO = "CANCELACION DEL CIRCUITO DE AUTORIZACION "



         SESSION("SEGCKT")    = "NO"

         StrSql = "Select * From SEGCIRCUITO WHERE OBJTIPO = '" & session("TCliente") & "' and  OBJCOD = '" & session("NCliente") & "' ORDER BY ID DESC "
         set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      ''  RESPONSE.WRITE "SQL " & Strsql & "<BR>"
         if not rs.eof then
            SESSION("SEGCKT")    = RS(6)
         END IF

         Strsql = "SELECT * FROM CKTVIEW WHERE SEGCKT = '" & session("SEGCKT") & "' "
         Strsql = Strsql & " and PSCFIN = 'NO'  "
     ''   RESPONSE.WRITE "SQL " & Strsql & "<BR>"
         set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         if NOT rsx.eof then
            do while not rsx.eof
               session("WKFPSC")    = RSX(9)
          '' ''  SESSION("SEGCKT")    = RSX(5)
               SESSION("WKF")       = RSX(0)
               SESSION("ORGUSR")    = RSX(15)
               session("SEGCKTPSC") = RSX(8)
               sobserv              = "CIRCUITO CANCELADO POR SOLICITUD DEL USUARIO:  " 
               sobserv              = sobserv & SESSION("USUARIO") & "&nbsp;" & SESSION("SEGUSUARIO")

	       strsql = "UPDATE SEGPASO SET FIN = 'CANCEL - " & NOW() & "', observaciones = '" & sobserv & "' WHERE SEGCKTPSC = '" & SESSION("SEGCKTPSC") & "' "
	       call ofv.crearconsultaEx(StrSql,cn,1,parametros)

             rsx.movenext
             loop

             strsql = "UPDATE SEGCIRCUITO SET FIN = 'CANCELADO-" & NOW() & "'  WHERE SEGCKT = '" & SESSION("SEGCKT") & "' "
             call ofv.crearconsultaEx(StrSql,cn,1,parametros)


             IF (WKAPLIC = "OR" OR  WKAPLIC = "PL") AND session("TCliente") = "RSK" THEN
                 strsql = "UPDATE ORRSKCIRCUITO SET ESTADO = 1  WHERE OBJID = '" & session("NCliente") & "' "
                 call ofv.crearconsultaEx(StrSql,cn3,1,parametros)
             ELSE
                XXXX = GETWKFMANAGER(session("NCliente"),session("TCliente"),1,"OBJPRTY","WKF")
             END IF 



           Celda = ofv.gencelda("","Vt colspan=4 ","center","30%","","","","&nbsp;","si")

           Fila = ofv.GenRow("","","","","","","",Celda,"si")




           sInput = TITULO 
           Celda = ofv.gencelda("","Vt colspan=2 style='BORDER-STYLE:RIDGE;BORDER-WIDTH:2;BORDER-COLOR:#EE0000;' ","center","30%","","","",sInput,"si")

           sInput = ofv.GenerarInput("boton1","FINALIZADA","submit","edPropiedades3","5","bt style='background-color:#005500;' ","")
           sAccion = "EDCANCELWkf.asp?OPCION=NO"
           sInput = ofv.GenForm("edPropiedades3","",sAccion,"","post","",sInput,"si")
           Celda = celda & ofv.gencelda("","Vt colspan=2 style='BORDER-STYLE:RIDGE;BORDER-WIDTH:2;BORDER-COLOR:#00EE00;' ","center","30%","","","",sInput,"si")

           Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

           sHTML = sHTML & "<br><br><br><br><br>"
           sHTML = sHTML & ofv.GenTabla("","aaa style='BORDER-STYLE:RIDGE;BORDER-WIDTH:2;BORDER-COLOR:#F7AA22;' ","","80% border=0 ","","0","0","1","0","","",Fila,"si")

         else


           Celda = ofv.gencelda("","Vt colspan=4 ","center","30%","","","","&nbsp;","si")

           Fila = ofv.GenRow("","","","","","","",Celda,"si")




           sInput = TITULO 
           Celda = ofv.gencelda("","Vt colspan=2 style='BORDER-STYLE:RIDGE;BORDER-WIDTH:2;BORDER-COLOR:#EE0000;' ","center","30%","","","",sInput,"si")

           sInput = ofv.GenerarInput("boton1","ERROR NO SE ENCUENTRA CIRCUITO VIGENTE","submit","edPropiedades3","5","bt style='background-color:#005500;' ","")
           sAccion = "EDCANCELWkf.asp?OPCION=NO"
           sInput = ofv.GenForm("edPropiedades3","",sAccion,"","post","",sInput,"si")
           Celda = celda & ofv.gencelda("","Vt colspan=2 style='BORDER-STYLE:RIDGE;BORDER-WIDTH:2;BORDER-COLOR:#00EE00;' ","center","30%","","","",sInput,"si")

           Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

           sHTML = sHTML & "<br><br><br><br><br>"
           sHTML = sHTML & ofv.GenTabla("","aaa style='BORDER-STYLE:RIDGE;BORDER-WIDTH:2;BORDER-COLOR:#F7AA22;' ","","80% border=0 ","","0","0","1","0","","",Fila,"si")

         END IF
         call ofv.cerrarconsulta(rsx)





 END SELECT


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------



  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML






%>