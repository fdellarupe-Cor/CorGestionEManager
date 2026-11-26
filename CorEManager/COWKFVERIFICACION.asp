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
		

		
		Strsql = "(select f.objid, f.requiereplan, f.requiereindicador,  isnull(f.incplan, '') AS incplan, isnull(o.idindicador, 0) as idindicador " & CHR(13)
		Strsql = Strsql & "from  oractivriesgo f  inner join ORRSKCIRCUITO b1  on (f.tipobjeto=b1.tipobjeto and f.periodo=b1.periodo and  f.codobjeto = b1.codobjeto )  " & CHR(13)
		Strsql = Strsql & "left join ORRELAINDIRIESGO o on (o.codriesgo = f.codriesgo ) where b1.objid = " & sCodigo & " ) as t1  " & CHR(13)


		
		'response.write strsql
		
		Set RS = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
		
		REQUIEREPLAN = FALSE
		REQUIEREINDICADORES = FALSE
		
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
			response.redirect "CoWKFselect.asp?SKIND=1"
		END IF

		
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
		
		IF NOT REQUIEREPLAN AND NOT REQUIEREINDICADORES THEN
			response.redirect "CoWKFselect.asp?SKIND=1"
		END IF
		
		sHTML = sHTML & "</body></html>"
		
	ELSE
		response.redirect "CoWKFselect.asp?SKIND=1"
	END IF

	
End If

response.write sHTML
%>