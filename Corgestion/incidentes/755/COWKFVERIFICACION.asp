<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->

<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  sHTML = ok
else
	response.expires = 0
	response.buffer = True
	APLCHECK()


	sCodigo   = request.querystring("Codobjeto")

	sHTML = ofv.FormHeader(session("FormN"))

	IF session("APLMDCLI")="BC" THEN

		Set cn = ofv.conectar(ofv.strconn2)


		sInput = "NO ES POSIBLE CONTINUAR CON EL CIRCUITO DE WORKFLOW:"
		Celda = ""
		Celda = Celda & ofv.GenCelda("","Vt BGCOLOR='' ","left","","","","",sInput,"si")
		Fila  = ofv.GenRow("","","","","","","",Celda,"si")
		sHTML = sHTML & "<br><br>" & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;' valign='middle' align='center' ","","100% ","30","0","","-1","1","","",Fila,"si")



		Strsql = "(select f.objid, f.requiereplan, f.requiereindicador,  isnull(f.incplan, '') AS incplan, isnull(o.idindicador, 0) as idindicador, F.ACTIVIDAD, F.CODRIESGO " & CHR(13)
		Strsql = Strsql & " from  oractivriesgo f  inner join ORRSKCIRCUITO b1  on (f.tipobjeto=b1.tipobjeto and f.periodo=b1.periodo and  f.codobjeto = b1.codobjeto )  " & CHR(13)
		Strsql = Strsql & " left join ORRELAINDIRIESGO o on (o.codriesgo = f.codriesgo ) where b1.objid = " & sCodigo & " ) as t1  " & CHR(13)



		'response.write strsql

		Set RS = ofv.crearconsultaEx(Strsql, cn, 1, parametros)

		REQUIEREPLAN        = FALSE
		REQUIEREINDICADORES = FALSE

		IF NOT RS.EOF THEN
			DO WHILE NOT RS.EOF
				IF RS("requiereplan") = "SI" AND (RS("incplan") = "NO" OR RS("incplan") = "" ) THEN
					'RESPONSE.WRITE "requiere plan"
					REQUIEREPLAN = TRUE
					'sHTML = sHTML & "<BR>" & sCodigo & "</BR>"

                                ELSEIF  RS("requiereplan") = "SI" AND RS("incplan") = "SI"  THEN

                                       strsql= "select * from ORAGENDA  where  actividad='" & rs("actividad")  & "' and activriesgo='" & RS("CODRIESGO") & "' "
			                'response.write chr(13) & strsql & chr(13)
	                               set rsAGENDA = ofv.crearconsultaEx(StrSql,cn,1,parametros)

	                               if not rsAGENDA.eof  then
                                          sCODAGENDA = rsAGENDA(10)
	                                  if (rsAGENDA(5) & "") < "2"  then REQUIEREPLAN = TRUE

                                       ELSE
					  REQUIEREPLAN = TRUE
                                       END IF


				END IF
				IF CINT(RS("idindicador")) = 0 AND RS("requiereindicador") = "SI" THEN
					'RESPONSE.WRITE "indicadores"
					REQUIEREINDICADORES = TRUE
					'sHTML = sHTML & "<BR>" & sCodigo & "</BR>"
				END IF
				RS.MOVENEXT
			LOOP
		ELSE
			'response.redirect "CoWKFselect.asp?SKIND=1"
		END IF


		IF REQUIEREPLAN THEN

			sInput = "EXISTEN PLANES PENDIENTES DE AGENDAR"
			Celda = ""
			Celda = Celda & ofv.GenCelda("","Vt BGCOLOR='#F7EEDD' ","center","","","","",sInput,"si")
			Fila  = ofv.GenRow("","","","","","","",Celda,"si")
			sHTML = sHTML & "<br><br>" & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;' valign='middle' align='center' ","","100% ","30","0","","-1","1","","",Fila,"si")

		END IF

	'	IF REQUIEREINDICADORES THEN
	'
	'		sInput = "FALTA ASOCIAR INDICADORES"
	'		Celda = ""
	'		Celda = Celda & ofv.GenCelda("","Vt BGCOLOR='#F7EEDD' ","center","","","","",sInput,"si")
	'		Fila  = ofv.GenRow("","","","","","","",Celda,"si")
	'		sHTML = sHTML & "<br><br>" & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;' valign='middle' align='center' ","","100% ","30","0","","-1","1","","",Fila,"si")
	'
	'	END IF

	'	IF NOT REQUIEREPLAN AND NOT REQUIEREINDICADORES THEN
	'		response.redirect "CoWKFselect.asp?SKIND=1"
	'	END IF

		IF NOT REQUIEREPLAN THEN
				response.redirect "CoWKFselect.asp?SKIND=1"
		END IF


		sHTML = sHTML & "</body></html>"

	ELSEif session("APLMDCLI")="BP" THEN
     		Set cn = ofv.conectar(ofv.strconn2)



	 		sInput = "NO ES POSIBLE CONTINUAR CON EL CIRCUITO DE WORKFLOW:"
	 		Celda = ""
	 		Celda = Celda & ofv.GenCelda("","Vt BGCOLOR='' ","left","","","","",sInput,"si")
	 		Fila  = ofv.GenRow("","","","","","","",Celda,"si")
	 		sHTML = sHTML & "<br><br>" & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;' valign='middle' align='center' ","","100% ","30","0","","-1","1","","",Fila,"si")



	 		Strsql = "(select f.objid, f.requiereplan, f.requiereindicador,  isnull(f.incplan, '') AS incplan, isnull(o.idindicador, 0) as idindicador " & CHR(13)
	 		Strsql = Strsql & "from  oractivriesgo f  inner join ORRSKCIRCUITO b1  on (f.tipobjeto=b1.tipobjeto and f.periodo=b1.periodo and  f.codobjeto = b1.codobjeto )  " & CHR(13)
	 		Strsql = Strsql & "left join ORRELAINDIRIESGO o on (o.codriesgo = f.codriesgo ) where b1.objid = " & sCodigo & " ) as t1  " & CHR(13)



	 		'response.write strsql

	 		Set RS = ofv.crearconsultaEx(Strsql, cn, 1, parametros)

	 		REQUIEREPLAN = FALSE
	 		REQUIEREINDICADORES = FALSE
	 		REQUIEREACTIVIDAD=FALSE

	 		IF NOT RS.EOF THEN
	 			DO WHILE NOT RS.EOF
	 				IF RS("requiereplan") = "SI" AND (RS("incplan") = "NO" OR RS("incplan") = "" ) THEN
	 					'RESPONSE.WRITE "requiere plan"
	 					REQUIEREPLAN = TRUE
	 					'sHTML = sHTML & "<BR>" & sCodigo & "</BR>"
	 				END IF
	 				IF CINT(RS("idindicador")) = 0 AND RS("requiereindicador") = "SI" THEN
	 					'RESPONSE.WRITE "indicadores"
	 					REQUIEREINDICADORES = TRUE
	 					'sHTML = sHTML & "<BR>" & sCodigo & "</BR>"
	 				END IF
	 				RS.MOVENEXT
	 			LOOP
	 		ELSE
	 		'	response.redirect "CoWKFselect.asp?SKIND=1"
	 		END IF

           CALL OFV.CERRARCONSULTA(RS)

	 		IF REQUIEREPLAN THEN

	 			sInput = "EXISTEN PLANES PENDIENTES DE AGENDAR"
	 			Celda = ""
	 			Celda = Celda & ofv.GenCelda("","Vt BGCOLOR='#F7EEDD' ","center","","","","",sInput,"si")
	 			Fila  = ofv.GenRow("","","","","","","",Celda,"si")
	 			sHTML = sHTML & "<br><br>" & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;' valign='middle' align='center' ","","100% ","30","0","","-1","1","","",Fila,"si")

	 		END IF

	 		IF REQUIEREINDICADORES THEN

	 			sInput = "FALTA ASOCIAR INDICADORES"
	 			Celda = ""
	 			Celda = Celda & ofv.GenCelda("","Vt BGCOLOR='#F7EEDD' ","center","","","","",sInput,"si")
	 			Fila  = ofv.GenRow("","","","","","","",Celda,"si")
	 			sHTML = sHTML & "<br><br>" & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;' valign='middle' align='center' ","","100% ","30","0","","-1","1","","",Fila,"si")

	 		END IF
            'verifica actividades
            STRSQL = "select B.IDEN4 from orrskcircuito A "
			STRSQL =STRSQL & " INNER JOIN OBJLNK B ON ( A.CODOBJETO=B.CODOBJETO)"
            STRSQL=STRSQL & " WHERE A.OBJID=" & sCodigo
            Set RSactividades = ofv.crearconsultaEx(Strsql, cn, 1, parametros)

            IF NOT RSactividades.EOF THEN
            CODPROCED=RSactividades(0)

            CALL OFV.CERRARCONSULTA(RSactividades)
            Call ofv.cerrarconn(cn)

            Set cnfd = ofv.conectar(ofv.strconn2)

            STRSQL= "select b.actividad,C.* from OBJACCION a  "
			STRSQL= STRSQL & " left join (select distinct actividad  from oractivriesgo where periodo = " & session("orperiodo") & ") b  on a.codaccion = b.actividad  "
			STRSQL=STRSQL & " LEFT JOIN oractivriesgo C ON ( C.ACTIVIDAD=B.ACTIVIDAD)"
			STRSQL =STRSQL & " WHERE a.codproced = '" & CODPROCED & "'  and (a.modi = 0 or (a.modi <> 0 and a.tipmodi <> 'E')) "
			STRSQL = STRSQL & "Order By a.secuen "
			'RESPONSE.WRITE STRSQL
            Set RSFD = ofv.crearconsultaEx(Strsql, cnfd, 1, parametros)
            	IF NOT RSFD.EOF THEN
				 			DO WHILE NOT RSFD.EOF
				 				IF CSTR(RSFD(0)& "")="" OR clng("0" & RSFD("itemriesgo")) = 0 THEN
				 				 REQUIEREACTIVIDAD=TRUE
				 				 EXIT DO
				 				END IF

				 				RSFD.MOVENEXT
				 			LOOP

	 		   END IF
	 		ELSE
	 		REQUIEREACTIVIDAD=TRUE
            END IF

            CALL OFV.CERRARCONSULTA(RSFD)
            Call ofv.cerrarconn(cnfd)
            IF REQUIEREACTIVIDAD THEN

				 			sInput = "EXISTEN ACTIVIDADES PENDIENTES DE EVALUAR"
				 			Celda = ""
				 			Celda = Celda & ofv.GenCelda("","Vt BGCOLOR='#F7EEDD' ","center","","","","",sInput,"si")
				 			Fila  = ofv.GenRow("","","","","","","",Celda,"si")
				 			sHTML = sHTML & "<br><br>" & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;' valign='middle' align='center' ","","100% ","30","0","","-1","1","","",Fila,"si")

	 		END IF



	 		IF NOT REQUIEREPLAN AND NOT REQUIEREINDICADORES AND not REQUIEREACTIVIDAD THEN
	 			response.redirect "CoWKFselect.asp?SKIND=1"
	 		END IF



		sHTML = sHTML & "</body></html>"
	''''end if
	else


		response.redirect "CoWKFselect.asp?SKIND=1"
	END IF


End If

response.write sHTML
%>