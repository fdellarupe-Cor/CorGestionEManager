<%

Function GENFMTWFLDET(FNCODOBJ, FNTIPOBJ)

	'response.write request.querystring
	FNSHTML = ""
	Celda = ""
	fila = ""
	Dim ATABLA(100)
	TIT = ""

	sImageAp = "<font face='corflaticon' size=4 color='1F7E28'>a</font>"
	sImageRc = "<font face='corflaticon' size=4 color='#BE1313'>r</font>"
	sImagePn = "<font face='corflaticon' size=4 color='D18742'>s</font>"

	StrSql = "SELECT *,b.ordenest FROM OBJwfldet a "
	StrSql = StrSql & " inner join  OBJwfl b on (b.codobjeto=a.codobjeto) "
	StrSql = StrSql & " Where a.codobjeto=" & nCodigo
	' StrSql = StrSql & " And a.tipobjeto = '" & sTipo & "' And a.CKTTYPE = '" & session("CKTTYPE") & "' order by a.orden,a.objid"
	StrSql = StrSql & " And a.tipobjeto = '" & sTipo & "' order by a.orden,a.objid"

	Set rx = ofv.crearconsultaEx(StrSql, cn, 1, parametros)
'RESPONSE.WRITE StrSql
	if not rx.eof then

	ordenwfl = ""
	If CBool(rx("ordenest")) Then
		ordenwfl = "ESTRICTO"
		'ordenwfl = "INDISTINTO"
	Else
		ordenwfl = "INDISTINTO"
		'ordenwfl = "ESTRICTO"
	End If

	If SACC = "PEND" Then TIT = "MODIFICACION"

	If session("base") = "K" Then
		Celda = ofv.gencelda("", "ttt colspan=4", "center", "90%", "", "", "", "<big>" & TIT & " CIRCUITO ENVIO DE FORMULARIO PARA " & WKOPCIONdsc & "</big>", "si")
	Else
		Celda = ofv.gencelda("", "ttt colspan=4", "center", "90%", "", "", "", "<big>" & TIT & " CIRCUITO DE  " & WKOPCIONdsc & "</big>", "si")
	End If
	fila = ofv.GenRow("", "", "", "", "20", "", "", Celda, "si")

	Celda = ofv.gencelda("", "ttt colspan=4", "center", "90%", "", "", "", "Lista de Autorizantes", "si")
	fila = fila & ofv.GenRow("", "", "", "", "20", "", "", Celda, "si")

	Celda = ofv.gencelda("", "ttt colspan=4", "", "90%", "", "", "", "Orden:&nbsp;" & ordenwfl, "si")
	fila = fila & ofv.GenRow("", "", "", "", "20", "", "", Celda, "si")

	Celda = ofv.gencelda("", "ttr", "", "2%", "", "", "", "&nbsp;", "si")
	Celda = Celda & ofv.gencelda("", "ttr", "", "80%", "", "", "", "Descripcion", "si")
	Celda = Celda & ofv.gencelda("", "ttr", "", "10%", "", "", "", "Estado", "si")
	'Celda = Celda & ofv.gencelda("", "ttr", "", "%28", "", "", "", "Aprobo", "si")
	fila = fila & ofv.GenRow("", "", "", "", "20", "", "", Celda, "si")

	Celda=""
	xactiva = 0
	xmanda = 0

	StrSql = "SELECT * FROM Defwfl"
	Set rs = ofv.crearconsultaEx(StrSql, cn, 1, parametros)

	If Not rs.EOF Then

		xactiva = rs(1)


	End If

	Call ofv.cerrarconsulta(rs)

	If xactiva Then xmanda = 1

	''session("CKTTYPE")

	Formulario = ""

	vinoSC = 0

	sImagenU = "<B><FONT FACE=corflaticon COLOR=#f7aa22 SIZE=2>€</FONT></B>"
	'If SACC = "PEND" Then
	'sql = " SELECT  count(*) FROM SEGCIRCUITO"
	'sql = sql & " inner join segpaso on ( segpaso.segckt=SEGCIRCUITO.segckt)"
	'sql = sql & " where SEGCIRCUITO.segckt='" & segckt & "' and  segpaso.fin='NO'"
	'response.write sql
	'Set rsckt = ofv.crearconsultaEx(sql, cn3, 1, parametros)
	'totalp = 0
	'If Not rsckt.EOF Then totalp = rsckt(0)
	' Call ofv.cerrarconsulta(rsckt)
	'end if
	' response.write totalp
	'  Set rsckt = ofv.crearconsultaEx(sql, cn3, 1, parametros)

	sql = " select a.*,b.segsalida from segcktresp a "
	sql = sql & " left join segpaso b on (a.segckt = b.segckt and a.id = b.segidresp) "
	sql = sql & " where a.segckt='" & SEGCKT & "' and a.modo='L' and b.segsalida is null "

	Set rf = ofv.crearconsultaEx(sql, cn3, 1, parametros)

	indice = 0
	Do While Not rf.eof
		ATABLA(indice) = rf(5)
		indice = indice + 1
		rf.movenext
	Loop


	sql = "select a.id,a.segckt,a.segcktpsc,a.modo,a.tipo,a.codigo,a.orden,a.nivel,a.subcodigo,b.segsalida,b.segidresp,b.fin,b.observaciones from segcktresp a "
	sql = sql & " left join segpaso b on (a.segckt = b.segckt and a.id = b.segidresp ) "
	sql = sql & " where a.segckt='" & SEGCKT & "' and a.modo='L'  "
	sql = sql & " order by orden "

	'response.write "sql: " & sql
	Set rs = ofv.crearconsultaEx(sql, cn3, 1, parametros)

	If Not rs.eof Then

		Do While Not rs.EOF


			sForm = "wfl" & rs(0)

			shref = ""

			bModificarR = bModificar

			'xorden = rs(3)

			'If Not IsNumeric(xorden) Then xorden = 0

			celda=""

			If xmanda And xorden > 8999 Then bModificarR = 0

			If SACC = "PEND" Then
				sql = " SELECT  SEGCIRCUITO.segckt,segpaso.responsable,segpaso.fin,segpaso.SEGCKTPSC FROM SEGCIRCUITO"
				sql = sql & " inner join segpaso on ( segpaso.segckt=SEGCIRCUITO.segckt)"
				sql = sql & " where SEGCIRCUITO.segckt='" & segckt & "' and segpaso.responsable='" & rs(5) & "'"
				Set rsckt = ofv.crearconsultaEx(sql, cn3, 1, parametros)

				'response.write sql & "<br>"
				Spaso = "N"
				If Not rsckt.EOF Then Spaso = "S"

				shref = ""

				'If bModificarR And Spaso = "N" And CBool(rx("ordenest")) Then

				'    sHref="edobjwfldetpend.asp?USR=" & rx("usr") & "&Codobjeto=" & nCodigo & "&TIPOBJETO=" &sTipo& "&SEGCKT=" & segckt &"&TACC=DELUSR&SEGCKTPSC=&ckt=" & Sckt & "&cod=" & sCod & "&ACC=" & SACC &"&VOLVER=" & sVolver

				'    Celda = ofv.BotonEliminar("Usuario", shref)
				'    Celda = ""



				'ElseIf bModificarR And Spaso = "S" And Not CBool(rs("ordenest")) Then


				'shref="edobjwfldetpend.asp?Codobjeto=" & nCodigo & "&TIPOBJETO=" &sTipo& "&SEGCKT=" & segckt &"&TACC=DEL&SEGCKTPSC=" & rsckt("SEGCKTPSC") & "&ckt=" & Sckt & "&cod=" & sCod & "&ACC=" & SACC &"&VOLVER=" & sVolver
				'  END IF
				'   response.write "aaa" & totalp
				'If totalp > 1 Then
				'     Celda = ofv.BotonEliminar("Usuario", shref)
				' Else
				'     Celda = ofv.gencelda("", "", "", "", "", "", "", "&nbsp;", "si")
				'End If



				'End If



				' Else

				'  response.write "w2"
				'response.write ordenwfl



				If ordenwfl = "ESTRICTO" Then


					sql = "select count(*) fin from segcktresp where segckt='" & SEGCKT & "' and orden > " & rs(6) & " and modo='L'"
					'response.write sql & "<br>"
					Set ro = ofv.crearconsultaEx(sql, cn3, 1, parametros)

					puede = True
					If ro(0) < 1 Then puede = False


					If bModificarR And (rs(10) & "" = "") Then
						href=""
						'shref = "edObJWfldetPEND.asp?ID=" & rs(0) & "&Tabla=segcktresp&volver="
						'shref = shref & vuelta & "?SEGCKT=" & segckt &"&TACC=DLT&codobjeto=" & ncodigo & "&tipobjeto=" & stipo & "&ACC=" & SACC

						'SINPUT = "<INPUT TYPE='BUTTON' VALUE='Elim' CLASS='BT' onclick=""disp_confirm('Desea eliminar autorizante?','" & sForm & "')"";>"

						SACCION = "edObJWfldetPEND.asp?ID=" & rs(0) & "&Tabla=segcktresp"
						SACCION = SACCION & "&SEGCKT=" & segckt &  "&TACC=DLT&codobjeto=" & ncodigo & "&tipobjeto=" & stipo & "&ACC=" & SACC

						'Formulario = ofv.GenForm(sForm,"",sAccion,"","POST","","","si")
						'Celda = ofv.gencelda("", "", "", "2%", "", "", "", sINPUT, "si")

					ElseIf rs(10) & "" <> "" And rs(11) = "NO" And puede Then

						shref = "edObJWfldetPEND.asp?ID=" & rs(0) & "&Tabla=segcktresp"
						shref = shref & "&SEGCKT=" & segckt &"&TACC=DEL&codobjeto=" & ncodigo & "&tipobjeto=" & stipo & "&ACC=" & SACC & "&orden=" & rs(6)

					End If

					If shref <> "" Then
						Celda = ofv.BotonEliminar("Usuario", shref)
					ElseIf shref = "" And saccion <> "" Then
						Celda = BotonEliminar("boton",SACCION,rs(0))
					Else
						Celda = ofv.gencelda("", "", "", "2%", "","", "", "", "si")
					End If

				ElseIf ordenwfl = "INDISTINTO" Then

					sql = "select count(*) fin from segpaso where segckt='" & SEGCKT & "' and fin='NO'"
					Set ro = ofv.crearconsultaEx(sql, cn3, 1, parametros)

					puede = True
					If ro(0) < 2 Then puede = False

					If bModificarR And rs(11) = "NO" And puede Then

						shref = "edObJWfldetPEND.asp?ID=" & rs(0) & "&Tabla=segcktresp"
						shref = shref & "&SEGCKT=" & segckt &"&TACC=DEL&codobjeto=" & ncodigo & "&tipobjeto=" & stipo & "&ACC=" & SACC

					End If

					If shref <> "" Then
						Celda = ofv.BotonEliminar("Usuario", shref)
					Else
						Celda = ofv.gencelda("", "", "", "2%", "", "", "", "&nbsp;", "si")
					End If


				End If
			End If


			sTabla = "Usuario"
			sBusca = rs(5)
			sGrpUsr = "S/D"


			xxxx = getInfoUser(sBusca)
			sGrpUsr = session("SGU_nombre")

			sGrpUsr = sImagenU & "&nbsp;" & sBusca & " - " & sGrpUsr


			sinput = ofv.gencelda("", "ttt ", "", "80%", "", "", "", sGrpUsr, "si")

			Celda = Celda & sinput

			sInput = ""


			besta = False

			For i = 0 To indice - 1

				If ATABLA(i) = rs(5) Then
					besta = True
					Exit For
				End If

			Next


			If rs(9) = "RCH" And Not besta And rs(11) <> "NO" Then

				'href = "edobjwfldetpend.asp?SEGCKT=" & segckt &"&TACC=RINS&usr=" & rsckt(1) & "&codobjeto=" & ncodigo & "&tipobjeto=" & stipo
				href = "edobjwfldetpend.asp?SEGCKT=" & segckt &"&TACC=RINS&ID=" & rs(0) & "&codobjeto=" & ncodigo & "&tipobjeto=" & stipo & "&ACC=" & SACC & "&sorden=" & rx("ordenest") & "&respasg=" & rs(5)

				sInput = "<a class=nt href='" & href & "' >"
				sInput = sInput & "Rechazado " & sImageRc
				sInput = sInput & "</a>"

			ElseIf rs(11) <> "NO" And rs(9) = "OK" Then

				sInput = "Aprobado " & sImageAp

			ElseIf rs(11) <> "NO" And rs(9) = "RCH" Then

				sInput = "Rechazado" & sImageRc

			ElseIf rs(11) = "NO" Then

				sInput = "Pendiente " & sImagePn

			ElseIf InStr(rs("observaciones") & "","REASIGNACION") > 0 Then

				sInput = "Reasignado"

			Else

				sInput = "----"

			End If

			If InStr(rs(11) & "", "FINALIZADO") > 0 Then sInput = "Finalizado por emisor"

			Celda = Celda & ofv.gencelda("", "ttt ", "center", "10%", "", "", "",sInput, "si")
			fila = fila & ofv.GenRow("", "", "", "", "", "", "", Celda, "si")
			Celda=""


			rs.movenext
		Loop
	End If

	'sHTML = sHTML & "<br><center>"
	sHTML = sHTML & ofv.GenTabla("", "aaa style='border-style:solid;border-width:1;border-color:#999999;' ", "", "95%", "", "0", "0", "-2", "1", "", "", Formulario & fila, "si")
	'sHTML = sHTML & "</center>"

	end if

	GENFMTWFLDET = "<div id=zglos class='norma-card'>" & FNSHTML  & "</div>" ' FNSHTML
End Function

Public Function BotonEliminar(sNombre, sHref,CLAVE)

	wktitdelete="Confirma la eliminacion "
	'' sImageELS = "<font face='corflaticon' size=3 color='#bb0000'>&#0253;</font>"
	'' sImageels = ofv.GenImage("","","","","Imagenes/Botdel.jpg","0","")
	sImageELS  = "<input  type=BUTTON  size=10%  value='&#052;' class=bte     >"
	sformobj   = "FRMOBJ" & CLAVE
	sAccion    = sHref
	formobj    = ofv.genform(sformobj,"AAA STYLE='display:none;' ",sAccion,"","post","","","si")
	sfunobj    = " onclick="" disp_confirm('" & wktitdelete & "','" & sformobj & "'); "" "

	BotonEliminar=ofv.GenCelda("","ut style='cursor:pointer;' title='" & wktitdelete & "' " & sfunobj & " ","","2","","","",sImageELs ,"si")
	xxform=xxform & formobj


End Function


%>