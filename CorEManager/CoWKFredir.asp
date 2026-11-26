<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
  response.write sHTML
else
  response.expires=0


''' verifica si es del circuito metodo 2.8.2 o anterior

     sTipo   = request.querystring("Tipobjeto")
     nCodigo = request.querystring("Codobjeto")

      sopcion = request.querystring("Sopcion")

  if sopcion <> "" then
     session("sopc") = sopcion
  else
     sopcion = session("sopc")
  end if

 IF SOPCION <> "NOVA" THEN
  set cn2 = ofv.conectar(ofv.strconn2)




  Strsql = "SELECT * FROM OBJWFL  "
  Strsql = Strsql & " Where codobjeto=" & nCodigo
  Strsql = Strsql & " And tipobjeto = '" & stipo & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
  if not rs.eof then

    ' response.write "v " & rs(6)  & "  "  & nCodigo & "  " & sopcion & "<br>"
     if rs(6)  then
        sMenu   = request.querystring("Menu")
        sSector = request.querystring("Sector")
        response.redirect "edObjWfl.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo  & "&Menu=" & sMenu & "&Sector=" & sSector & "&GLBSOLN=ESPWFL"
     end if

  end if
  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn2)

END IF

'''''' fin verifica

  set cn = ofv.conectar(ofv.strconn5)



 ' RESPONSE.WRITE REQUEST.QUERYSTRING & "<BR>"

     XUSR = " USUARIO: " & session("usuario") & "&nbsp;-&nbsp;" & session("SEGusuario")

  if sopcion = "VER" then

    Strsql = "SELECT * FROM CKTVIEW WHERE SEGCKTPSC = '" & session("SEGCKTPSC") & "' "
    set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT rsx.eof then
       session("WKFPSC")    = RSX(9)
       SESSION("SEGCKT")    = RSX(5)
       SESSION("WKF")       = RSX(0)
       SESSION("ORGUSR")    = RSX(15)

       response.redirect "SEGCKTESTADO.asp?CKT=" & RSx(5) & "&DSD=COR"

    else
        response.write ofv.FormHeader(session("FormN")) 'OFV.MENUHEADER("","")
	   	  			              response.write CorPnlFnc("N","")

	   			  response.write  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"
	          response.write MOSTPAN(XUSR,"ERROR BUSQUEDA DEL CIRCUITO VIGENTE!!","NO") '& "</body></html>"
	          response.write  "</div>"

	   	   					     response.write ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
	   	   						response.write ofv.GenRow("","","","","","","",Celda,"si")
	   	   						response.write ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")
			  response.write  CorPnlIndx("N","") & "</body></html>"


    END IF
    call ofv.cerrarconsulta(rsx)

  end if

  if sopcion = "VER2" then



     sTipo   = request.querystring("Tipobjeto")
     nCodigo = request.querystring("Codobjeto")
     session("TCliente") = sTipo
     session("NCliente") = nCodigo



     snova = 2



         SESSION("SEGCKT")    = "NO"

         StrSql = "Select * From SEGCIRCUITO WHERE OBJTIPO = '" & session("TCliente") & "' and  OBJCOD = '" & session("NCliente") & "' ORDER BY ID DESC "
         set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
'RESPONSE.WRITE "SQL " & Strsql & "<BR>"
         if not rs.eof then
            SESSION("SEGCKT")    = RS(6)
            snova = 1
         END IF


     if  snova = 1 then

           response.write ofv.FormHeader(session("FormN")) 'OFV.MENUHEADER("","")
		                 response.write CorPnlFnc("N","")

		   			  response.write  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"


		               response.write MOSTPAN(XUSR,"NO SE ENCUENTRA HABILITADO PARA MODIFICAR Y/O HABILITAR CIRCUITOS PARA ESTE DOCUMENTO","SI") '& "</body></html>"
		               response.write  "</div>"

		   								     response.write ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
		   									response.write ofv.GenRow("","","","","","","",Celda,"si")
		   									response.write ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")
			  response.write  CorPnlIndx("N","") & "</body></html>"

           ' response.write OFV.MENUHEADER("","")
           ' response.write MOSTPAN(XUSR,"NO SE ENCUENTRA HABILITADO PARA MODIFICAR Y/O HABILITAR CIRCUITOS PARA ESTE DOCUMENTO","SI") & "</body></html>"



     else
 response.write ofv.FormHeader(session("FormN")) 'OFV.MENUHEADER("","")
						              response.write CorPnlFnc("N","")

			  response.write  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"
            response.write MOSTPAN(XUSR,"NO EXISTEN CIRCUITOS HABILITADOS EN ESTE DOCUMENTO","NO")' & "</body></html>"response.write  "</div>"

					     response.write ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
						response.write ofv.GenRow("","","","","","","",Celda,"si")
						response.write ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")
			  response.write  CorPnlIndx("N","") & "</body></html>"
            'response.write OFV.MENUHEADER("","")
            'response.write MOSTPAN(XUSR,"NO EXISTEN CIRCUITOS HABILITADOS EN ESTE DOCUMENTO","NO") & "</body></html>"

     end if



  end if

  if sopcion = "NOVA" then


      response.write ofv.FormHeader(session("FormN"))
						              response.write CorPnlFnc("N","")

			  response.write  "<div id='mt2'  style='Xdisplay:none;height:350px;overflow:auto;" & styl & "' >"
            response.write MOSTPAN(XUSR,"ESTE CIRCUITO NO CORRESPONDE AL ESQUEMA - VERIFICAR EN TIPOS DE ESQUEMA LA CONDICION EN EL ITEM 'OBSERVACION'","NO") & "</DIV>"

					     response.write ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
						response.write ofv.GenRow("","","","","","","",Celda,"si")
						response.write ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")
			  response.write  CorPnlIndx("N","") & "</body></html>"

  end if


  if sopcion = "APR" then



     sTipo   = request.querystring("Tipobjeto")
     nCodigo = request.querystring("Codobjeto")
     session("TCliente") = sTipo
     session("NCliente") = nCodigo

  ' response.write   nCodigo & "  " & stipo & "  ddd<br>"

     snova = 2



         SESSION("SEGCKT")    = "NO"

         StrSql = "Select * From SEGCIRCUITO WHERE OBJTIPO = '" & session("TCliente") & "' and  OBJCOD = '" & session("NCliente") & "' ORDER BY ID DESC "
         set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
'RESPONSE.WRITE "SQL " & Strsql & "<BR>"
         if not rs.eof then
            SESSION("SEGCKT")    = RS(6)
         END IF

         StrSql = "Select * From SEGCKTRESP WHERE SEGCKT = '" & SESSION("SEGCKT") & "' "
         Strsql = Strsql & " and TIPO = 'USR' AND CODIGO = '" & session("usuario") & "' "
         set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
'RESPONSE.WRITE "SQL " & Strsql & "<BR>"
         if not rs.eof then
            snova = 1
         END IF

     if  snova = 1 then

         Strsql = "SELECT * FROM CKTVIEW WHERE SEGCKT = '" & session("SEGCKT") & "' "
         Strsql = Strsql & " and PSCFIN = 'NO' AND PSCRESPONSABLE = '" & session("usuario") & "' "

         set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         if NOT rsx.eof then
            session("WKFPSC")    = RSX(9)
            SESSION("SEGCKT")    = RSX(5)
            SESSION("WKF")       = RSX(0)
            SESSION("ORGUSR")    = RSX(15)
            session("SEGCKTPSC") = RSX(8)

            response.redirect "CoWKFselect.asp?SKIND=1&tipobjeto="  & stipo & "&codobjeto=" & nCodigo

         else

          '  response.write OFV.MENUHEADER("","")
          '  response.write MOSTPAN(XUSR,"EXISTEN AUTORIZANTES CON ORDEN PREVIO SIN CONFIRMAR","SI") & "</body></html>"
           response.write ofv.FormHeader(session("FormN")) 'OFV.MENUHEADER("","")
		  			              response.write CorPnlFnc("N","")

		  			  response.write  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"
		              response.write MOSTPAN(XUSR,"EXISTEN AUTORIZANTES CON ORDEN PREVIO SIN CONFIRMAR","SI")
		               response.write  "</div>"

		  					     response.write ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
		  						response.write ofv.GenRow("","","","","","","",Celda,"si")
		  						response.write ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")
			  response.write  CorPnlIndx("N","") & "</body></html>"

         END IF
         call ofv.cerrarconsulta(rsx)

     else

         '   response.write OFV.MENUHEADER("","")
         '   response.write MOSTPAN(XUSR,"NO SE ENCUENTRA HABILITADO PARA CONFIRMAR EN ESTE DOCUMENTO","NO") & "</body></html>"
         response.write ofv.FormHeader(session("FormN")) 'OFV.MENUHEADER("","")
		 			              response.write CorPnlFnc("N","")

		 			  response.write  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"
		             response.write MOSTPAN(XUSR,"NO SE ENCUENTRA HABILITADO PARA CONFIRMAR EN ESTE DOCUMENTO","NO")
		              response.write  "</div>"

		 		     response.write ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
		 			response.write ofv.GenRow("","","","","","","",Celda,"si")
		 			response.write ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")
			  response.write  CorPnlIndx("N","") & "</body></html>"

     end if



  end if

 END IF

FUNCTION MOSTPAN(dusr,TEXTO,OPC)

    sImagenRU = "<B><FONT FACE=corflaticon COLOR=#f7aa22 SIZE=3>&#0128;</FONT></B>"

           Celda = ofv.gencelda("","Vt BGCOLOR=#FFFFFF ","center","100%","","","","CIRCUITO DE AUTORIZACION","si")
           Fila = ofv.GenRow("","","","","","","",Celda,"si")

           SINPUT = "&nbsp;"
           Celda = ofv.gencelda("","Vt style='border-TOP-style:RIDGE;border-TOP-width:2;border-TOP-color:#f7aa22;' ","center","","","","",sInput,"si")
           Fila = FILA & ofv.GenRow("","","","","","","",Celda,"si")


           SINPUT = sImagenRU & "&nbsp;" & dusr
           Celda = ofv.gencelda("","Vt  ","center","","","","",sInput,"si")
           Fila = FILA & ofv.GenRow("","","","","","","",Celda,"si")

           SINPUT = "&nbsp;"
           Celda = ofv.gencelda("","Vt  ","center","","","","",sInput,"si")
           Fila = FILA & ofv.GenRow("","","","","","","",Celda,"si")


           SINPUT = "<FONT FACE=VERDANA COLOR=#993333 SIZE=2>" & TEXTO & "</FONT>"
           Celda = ofv.gencelda("","Vt  ","center","","","","",sInput,"si")
           Fila = FILA & ofv.GenRow("","","","","","","",Celda,"si")


           SINPUT = "&nbsp;"
           Celda = ofv.gencelda("","Vt  ","center","","","","",sInput,"si")
           Fila = FILA & ofv.GenRow("","","","","","","",Celda,"si")

           if opc = "SI" then

           sInput = ofv.GenerarInput("boton1","Consultar circuito vigente","submit","edPropiedades3","5","bt","")
           sAccion = "SEGCKTESTADO.asp?CKT=" & SESSION("SEGCKT") & "&DSD=COR"
           sInput = ofv.GenForm("edPropiedades3","",sAccion,"","post","",sInput,"si")
           Celda = ofv.gencelda("","ut colspan=2","center","30%","","","",sInput,"si")

           Fila = FILA & ofv.GenRow("","","","","","","",Celda,"si")

           end if

           Fila = "<br><br><br><br><center>" & ofv.GenTabla("","aaa style='wfilter:shadow(color=#cccccc,direction=125);border-style:RIDGE;border-width:2;border-color:#f7aa22;' ","","80%  ","","0","0","1","0","","",fila,"si") & "</center>"


           MOSTPAN =  Fila


END FUNCTION
%>