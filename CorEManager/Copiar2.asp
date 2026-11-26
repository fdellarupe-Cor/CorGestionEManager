<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorCopyGen.inc" -->
<%


ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0

	'response.write request.querystring & "<br>" & request.form

	XXXX = NEWEMPRE("NO")
	MOVERESQUEMA= Request.querystring("MOVERESQUEMA")
	nCodigo    = Request.querystring("Codigo")
	sTipo      = Request.querystring("Tipo")
	sPerfil    = Request.querystring("Perfil")
	bCopiar    = Request.querystring("Copiar")
	sOpcion    = Request.querystring("Opcion")
	sEmpresa   = Request.querystring("empresa")
	sBASEF     = Request.querystring("BASEF")
	SACC       = Request.querystring("ACC")
	IDPEDIDO   = Request.querystring("IDPEDIDO")
	sUsuario   = Session("Usuario")
	sTipobjeto = sOpcion
	volver = ""
	ESQSUP=Request.querystring("ESQSUP")
	' response.write WKAPLIC
	Select Case  WKAPLIC
		Case "EM"
		WHERETIPO= " WHERE APLIC IN ('EM','GLB')"
		If session("base")="K" Then
			WHERETIPO2= " AND APLIC IN ('EM','GLB') AND TIPOCPT = 'K'"
		Else
			WHERETIPO2= " AND APLIC IN ('EM','GLB') AND TIPOCPT <> 'K'"
		End If
		Case "GLB"
		WHERETIPO= " WHERE APLIC IN ('EM','GLB')"
		'WHERETIPO2= " AND APLIC IN ('EM','GLB')"
		If session("base")="K" Then
			WHERETIPO2= " AND APLIC IN ('EM','GLB') AND TIPOCPT = 'K'"
		Else
			WHERETIPO2= " AND APLIC IN ('EM','GLB') AND TIPOCPT <> 'K'"
		End If
		Case Else
		WHERETIPO=""
		WHERETIPO2=""
	End Select
	APLCHECK()


	'response.write WHERETIPO

	If sOpcion="TODOS" Then ESQSUP=sOpcion
	'''''''''''''
	'' RESPONSE.WRITE REQUEST.QUERYSTRING

	If moveresquema="SI" Then
		sREADONLY= " readonly "
	Else
		sREADONLY= ""
	End If

	If bCopiar = "SI" Then


		Server.scripttimeout = 600
		Response.write OFV.FormBody

		Response.write "<font face='verdana' size=3 color=#336699>"

		Response.write "<center><p style='border-style=solid;border-width:1;border-color:#eeeeee;' ><b>  Registro de la Copia </b></p><br>"

		styl = "SCROLLBAR-FACE-COLOR:#eaeaea; SCROLLBAR-HIGHLIGHT-COLOR: #fdfdfd; SCROLLBAR-SHADOW-COLOR: #eeeeee; "
		styl = styl & "SCROLLBAR-3DLIGHT-COLOR: #ffffff; SCROLLBAR-ARROW-COLOR:#6699cc; SCROLLBAR-TRACK-COLOR: #fafafa;"
		styl = styl & "SCROLLBAR-DARKSHADOW-COLOR: #e2e2e2;"


		Response.write "<div id='mt2'  style='border-style=solid;border-width:2;border-color:#eeeeee;height:350px;width:400px;overflow:auto;" & styl & "' >"

		Response.write "<font face='verdana' size=2 color=#336699>"

		ntime0     = Minute(Now) & ":" & Second(Now)
		Response.write "tiempo proc inicio: " & ntime0
		xCodigo    = nCodigo
		xtipobjeto = sTipobjeto

		sCodiden   = UCase(Request.Form("codiden"))
		sDescri    = UCase(Request.Form("descri"))
		sEsquema   = Request.Form("Esquema")

		'response.write request.form
		i          = InStr(sEsquema, "_")
		nEsquema   = Left(sEsquema, i - 1)
		sEsquema   = Mid(sEsquema, i + 1)



		CCGIOCOD = nCodigo
		CCGIOTIP = sTipo

		CCGINCOD = 0
		CCGINTIP = sTipobjeto

		CCGIEMPR = sEmpresa
		CCGINPRF = sPerfil
		CCGIDESC = sDescri
		CCGICDID = sCodiden
		CCGICESQ = nEsquema
		CCGITESQ = sEsquema


		CCGESQSUP= ESQSUP




		Set cn = OFV.conectar(OFV.strconn2)

		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'''' VERIFICA EL PERFIL DEL DOCUMENTO ''''''
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

		bSector        = "NO"
		bEvento        = "NO"
		bProcedimiento = "NO"
		bControl       = "NO"
		bEsquema       = "NO"
		bTablas        = "NO"
		bDocum         = "NO"
		bEditor        = "NO"
		bDescri        = "NO"
		bMain          = "NO"
		' IF session("APLMDCLI")="BC" THEN
		BPORTADA="SI"
		'  ELSE
		'  BPORTADA="NO"
		' END IF

		strsql = "Select * From Objperfildoc Where codobjeto =" & xCodigo
		'  response.write "sql " & strsql
		Set rsPerfil = OFV.crearconsultaEx(strsql, cn, 1, volver)
		If rsPerfil.EOF Then
			perfilsp = "N"
			'  response.write "no "

			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'''' VERIFICA EL PERFIL GENERAL DEL OBJETO ''''''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


			strsql = "Select Objperfil.* From Objperfil, Tipobjeto where "
			strsql = strsql & "Objperfil.tipoperfil = Tipobjeto.perfil "
			strsql = strsql & "and Tipobjeto.tipobjeto = '" & xtipobjeto & "'"
			Set rsPerfil = OFV.crearconsultaEx(strsql, cn, 1, volver)

		Else
			'  response.write "si "
			perfilsp = "s"

		End If

		If Not rsPerfil.EOF Then

			If rsPerfil("objsector") Then          bSector = "SI"
			'Sectores

			If rsPerfil("objeven") Then            bEvento = "SI"
			'Eventos

			If rsPerfil("objproced") Then          bProcedimiento = "SI"
			'Procedimientos

			If rsPerfil("objcontrol") Then         bControl = "SI"
			'Controles

			If rsPerfil("objesquema") Then         bEsquema = "SI"
			'Esquemas

			If rsPerfil("objtabla") Then           bTablas = "SI"
			'Tablas

			If rsPerfil("objhtm") Then             bDocum = "SI"
			'Documentos

			If rsPerfil("objobj1") Then            bEditor = "SI"
			'Editor estructurado

			If rsPerfil("objaccion") Then          bDescri = "SI"
			'Descripcion

			If rsPerfil("objmain") Then            bMain = "SI"
			'Consideraciones Generales

		End If


		Call OFV.cerrarconsulta(rsPerfil)
		Call OFV.cerrarconn(cn)

		'response.write MOVERESQUEMA
		Select Case MOVERESQUEMA

			Case "SI"

			xxxx = CopyGenFn("","MOVE")

			Case   Else ' COPIAR


			'''''''''''''''''' invoca generacion del objeto (BLKOBJ)

			xxxx = CopyGenFn("BLKOBJ","NO")




			'''''''''''''''''' invoca generacion del relaciones (BLKREL)

			xxxx = CopyGenFn("BLKREL","NO")


			'''''''''''''''''' invoca generacion propietario del docuemnto (BLKAUT)

			xxxx = CopyGenFn("BLKAUT","NO")

			'''''''''''''''''' invoca permisos propietario del docuemnto (BLKAUT)

			If session("APLMDCLI") = "BC" Or session("APLMDCLI") = "GP" Then

				xxxx = CopyGenFn("BLKSDK","NO")
			End If

			'''''''''''''''''' invoca generacion propiedades del docuemnto (BLKPRT)

			xxxx = CopyGenFn("BLKPRT","NO")

			'''''''''''''''''' invoca generacion perfil del docuemnto (BLKPRF)

			xxxx = CopyGenFn("BLKPRF","NO")


			'''''''''''''''''' invoca generacion agrupamiento y segmento (BLKAGS)

			xxxx = CopyGenFn("BLKAGS","NO")

			'''''''''''''''''' invoca generacion origen de normativa (BLKORG)

			xxxx = CopyGenFn("BLKORG","NO")


			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			''''' seccion de bloques segun perfil '''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'''' DESCRIPCION ''''''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

			If bDescri = "SI" Then

				'''''''''''''''''' invoca generacion bloque descripcion (BLKDSC)

				xxxx = CopyGenFn("BLKDSC","NO")
			End If


			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'''' SECTORES ''''''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

			'Sectores
			If bSector = "SI" Then
				'''''''''''''''''' invoca generacion bloque sectores (BLKSEC)

				xxxx = CopyGenFn("BLKSEC","NO")
			End If



			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'''' EVENTOS ''''''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

			If bEvento = "SI" Then
				'''''''''''''''''' invoca generacion bloque eventos (BLKEVE)

				xxxx = CopyGenFn("BLKEVE","NO")
			End If


			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'''' PROCEDIMIENTOS ''''''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

			If bProcedimiento = "SI" Then

				'''''''''''''''''' invoca generacion bloque procedimientos (BLKPCD)

				xxxx = CopyGenFn("BLKPCD","NO")


			End If




			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'''' COPIA DESTINOS DEL DOCUMENTO ("BLKDEST")''''''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

			'''''''''''''''''' invoca generacion bloque procedimientos (BLKPCD)

			If SESSION("BASE") = "K" And (SESSION("APLMDCLI") = "BGAL") Then

				xxxx = CopyGenFn("BLKDEST","NO")

			End If






			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'''' DOCUMENTOS ASOCIADOS ''''''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

			If bDocum = "SI" Then
				'''''''''''''''''' invoca generacion bloque docuemntos asociados (BLKDOC)

				xxxx = CopyGenFn("BLKDOC","NO")
			End If

			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'''' TABLAS ''''''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			If bTablas = "SI" Then
				'''''''''''''''''' invoca generacion bloque tablas (BLKTAB)

				xxxx = CopyGenFn("BLKTAB","NO")

			End If

			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'''' EDITOR ESTRUCTURADO ''''''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			If bEditor = "SI" Then
				'''''''''''''''''' invoca generacion bloque contenido (BLKEST)

				xxxx = CopyGenFn("BLKEST","NO")

			End If

			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'''' CONTROLES ''''''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			If bControl = "SI" Then
				'''''''''''''''''' invoca generacion bloque descripcion (BLKCNT)

				xxxx = CopyGenFn("BLKCNT","NO")
			End If

			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'''' ESQUEMAS CONTABLES ''''''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			If bEsquema = "SI" Then
				'''''''''''''''''' invoca generacion bloque descripcion (BLKESQ)

				xxxx = CopyGenFn("BLKESQ","NO")
			End If


			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'''' DOCUMENTOS REDIRECCIONADOS ''''''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

			'''''''''''''''''' invoca generacion bloque descripcion (BLKFAQ)

			xxxx = CopyGenFn("BLKFAQ","NO")



			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'''' CONSIDERACIONES GENERALES ''''''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			If bMain = "SI" Then

				'''''''''''''''''' invoca generacion bloque descripcion (BLKCON)

				xxxx = CopyGenFn("BLKCON","NO")


			End If


			'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			'''' PORTADA ''''''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

			If BPORTADA = "SI" Then
				'''''''''''''''''' invoca generacion bloque portada

				xxxx = CopyGenFn("BLKPORTADA","NO")
			End If
		End Select

		Response.write "tiempo  fin: " & Minute(Now) & ":" & Second(Now) & "<br>"



		Response.write "</div>"

		Response.write "<center>"

		sInput = OFV.GenerarInput("", "Aceptar", "submit", "", "10", "bt", "")
		Response.write OFV.GenForm("Copiar", "", "vuelve.asp", "", "post", "", sInput, "si")
		Response.write "</center>"

		sHTML = "</body></html>"


	Else
		sHTML = OFV.FormBody & corstyler2() & "<link rel=""stylesheet"" type=""text/css"" href=""../fonts/Corfont/Corflaticon.css"">"
		'Objetos

		Set cn = OFV.conectar(OFV.strconn2)
		Select Case ESQSUP



			Case "Y" ' SUBCARPETA INCLUIDA
			Set cn = OFV.conectar(OFV.strconn2)
			strsql = "select * from estados"
			Set rsTipo = OFV.crearconsultaEx(strsql, cn, 1, volver)
			inp = 0
			Do Until rsTipo.EOF
				ReDim Preserve aestr(inp)
				ReDim Preserve nestr(inp)
				aestr(inp) = UCase(rsTipo(0))
				nestr(inp) = UCase(rsTipo(1))
				inp = inp + 1
				rsTipo.movenext
			Loop

			If session("base")="K" Then
				strsql = "select tipobjeto,descri from subtema where icono='K'"
			Else
				strsql = "select tipobjeto,descri from subtema where icono <> 'K'"
			End If
			Set rsTipo = OFV.crearconsultaEx(strsql, cn, 1, volver)
			inp = 0
			Do Until rsTipo.EOF
				ReDim Preserve asector(inp)
				ReDim Preserve nsector(inp)
				asector(inp) = UCase(rsTipo(0))
				nsector(inp) = UCase(rsTipo(1))
				inp = inp + 1
				rsTipo.movenext
			Loop

			If session("base")="K" Then
				strsql = "select tipobjeto,descri from tipobjeto " & WHERETIPO & " and  tipocpt='K' "
			Else
				strsql = "select tipobjeto,descri from tipobjeto  " & WHERETIPO  & " and  tipocpt <> 'K' "
			End If
			Set rsTipo = OFV.crearconsultaEx(strsql, cn, 1, volver)
			inp = 0
			Do Until rsTipo.EOF
				ReDim Preserve aevento(inp)
				ReDim Preserve nevento(inp)
				aevento(inp) = UCase(rsTipo(0))
				nevento(inp) = UCase(rsTipo(1))
				inp = inp + 1
				rsTipo.movenext
			Loop
			Call OFV.cerrarconsulta(rsTipo)





			'''sEmpresa

			If  sEmpresa = "" Then sEmpresa = 1
			bempresa = sEmpresa
			sLink    = "&nbsp;"

			Fila = ""

			If  Session("EMPMUN") = "S" Then

				For IAEMP=1 To TAEMP

					If ATREMP(IAEMP,2) = 1 Then
						If ATREMP(IAEMP,0) = sEmpresa Then
							bempresa = ATREMP(IAEMP,4)
							sLink    = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(0,5),"0","")
							Exit For
						End If
					End If

				Next

				'empresa
				Celda = OFV.GenCelda("", "ttt ", "", "", "", "", "", sLink, "si")
				sInput = "&nbsp;&nbsp;" & bempresa
				Celda = Celda & OFV.GenCelda("", "ttt", "", "", "", "", "", sInput, "si")
				Fila = Fila & OFV.GenRow("", "", "", "", "30", "", "", Celda, "si")

				Celda = OFV.GenCelda("", "ut colspan=2 ", "", "", "", "", "", "&nbsp;&nbsp;", "si")
				Fila = Fila & OFV.GenRow("", "", "", "", "30", "", "", Celda, "si")

			End If
			If session("base")="K" Then



				Strsql = "Select * From Subtema WHERE tipobjeto='"& stipo &"' and (aplic = '" & wkaplic & "' or aplic = 'GLB') Order By descri"
			Else
				Strsql = "Select * From Subtema WHERE tipobjeto='"& stipo &"' and (aplic = '" & wkaplic & "' or aplic = 'GLB') Order By descri"
			End If
			Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

			If Not rs.eof Then sNombredOC = rs("descri") &":"
			Call ofv.cerrarconsulta(rs)

			strsql = "SELECT DESCRI FROM PRODSERV where codprodserv = " & nCodigo

			Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not rs.eof Then sNombredOC = sNombredOC & rs(0)
			Call ofv.cerrarconsulta(rs)




			'Identificacion
			Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "NOMBRE", "si")
			sInput = sNombredOC
			Celda = Celda & OFV.GenCelda("", "UTI", "", "", "", "", "", sInput, "si")
			Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")




			'Esquema superior
			Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "Mover a Tema: ", "si")
			sInput = "<font class=at><select name=Esquema>"
			If aTipos <> ")" Then
				sInput = sInput & "<option value='nsnc' STYLE='border-style=solid;border-width:1;border-color:#eeeeee;"
				sInput = sInput & "background-color:#6699cc;color:#ffffff;' >"
				sInput = sInput & "+++ Temas ++++++++++++++++++++++++++++++++++++"
				sInput = sInput & "</OPTION>"

				If session("base")="K" Then
					strsql="select * from tema where pstipo='K' order by  descri"
				Else
					strsql="select * from tema where pstipo <> 'K' order by  descri"
				End If
				Set rsTipo2 = OFV.crearconsultaEx(strsql, cn, 1, volver)
				Do Until rsTipo2.EOF

					sInput = sInput & "<option value='" & rsTipo2(0) & "_" & rsTipo2(0) &"' "
					sInput = sInput & " >" & rsTipo2("descri") & "</option>"


					rsTipo2.movenext
				Loop
			End If


			sInput = sInput & "</select></font>"
			Celda  = Celda & OFV.GenCelda("", "UTI", "", "", "", "", "", sInput, "si")
			Fila   = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

			Tabla  = OFV.GenTabla("", "", "", "80%", "", "0", "", OFV.wspacing, 1 + OFV.wspacing, "", "", Fila, "si")


			Tabla      = Tabla & "<br><br>" & OFV.GenTabla("", "", "", "100%", "", "0", "0", "0", "0", "", "", "", "no")
			Fila       = Tabla & OFV.GenRow("", "", "", "", "", "", "", "", "no")

			sInput     = OFV.GenerarInput("", "Aceptar", "submit", "", "10", "bt", "this.style.display='none';")
			Celda      = Fila & OFV.GenCelda("", "", "center", "50%", "", "", "", sInput, "si")
			sAccion    = "Copiar2.asp?Codigo=" & nCodigo & "&Tipo=" & sTipo & "&Perfil=" & sPerfil & "&empresa=" & sEmpresa
			sAccion    = sAccion & "&Opcion=" & sOpcion & "&Copiar=SI" & "&moveresquema=" & moveresquema & "&ESQSUP=" & ESQSUP
			Formulario = OFV.GenForm("Copiar", "", sAccion, "", "post", "", Celda, "si")

			sFunc      = "window.close()"
			sInput     = OFV.GenerarInput("", "Cancelar", "button", "", "10", "bt", sFunc)
			Celda      = OFV.GenCelda("", "", "center", "50%", "", "", "", sInput, "si")
			Formulario = Formulario & OFV.GenForm("", "", "", "", "post", "", Celda, "si")
			Fila       = Formulario & OFV.GenFRow()

			Tabla      = Fila & OFV.GenFTabla()

			sHTML      = sHTML & "<br><br><br><br><center>" & Tabla & "</center>"
			sHTML      = sHTML & "</body></html>"


			Case "TODOS" ' TODOS MENOS LA SUBCARPETA

			strsql = "select * from estados"
			Set rsTipo = OFV.crearconsultaEx(strsql, cn, 1, volver)
			inp = 0
			Do Until rsTipo.EOF
				ReDim Preserve aestr(inp)
				ReDim Preserve nestr(inp)
				aestr(inp) = UCase(rsTipo(0))
				nestr(inp) = UCase(rsTipo(1))
				inp = inp + 1
				rsTipo.movenext
			Loop

			If session("base")="K" Then
				strsql = "select tipobjeto,descri from subtema where icono = 'K' "
			Else
				strsql = "select tipobjeto,descri from subtema where icono <>'K' "
			End If
			Set rsTipo = OFV.crearconsultaEx(strsql, cn, 1, volver)
			inp = 0
			Do Until rsTipo.EOF
				ReDim Preserve asector(inp)
				ReDim Preserve nsector(inp)
				asector(inp) = UCase(rsTipo(0))
				nsector(inp) = UCase(rsTipo(1))
				inp = inp + 1
				rsTipo.movenext
			Loop

			If session("base")="K" Then
				strsql = "select tipobjeto,descri from tipobjeto" & WHERETIPO & " and tipocpt='K' "
			Else
				strsql = "select tipobjeto,descri from tipobjeto" & WHERETIPO & " and tipocpt <> 'K' "
			End If
			Set rsTipo = OFV.crearconsultaEx(strsql, cn, 1, volver)
			inp = 0
			Do Until rsTipo.EOF
				ReDim Preserve aevento(inp)
				ReDim Preserve nevento(inp)
				aevento(inp) = UCase(rsTipo(0))
				nevento(inp) = UCase(rsTipo(1))
				inp = inp + 1
				rsTipo.movenext
			Loop
			Call OFV.cerrarconsulta(rsTipo)

			strsql = "Select Objetos.*,objprty.estado From Objetos, Objprty where "
			strsql = strsql & "Objetos.codobjeto = Objprty.codobjeto and objetos.codobjeto = " & nCodigo

			'  response.write strsql
			Set rsObj = OFV.crearconsultaEx(strsql, cn, 1, volver)
			If  rsObj.EOF Then
				strsql = "Select a.descri,b.descri as descri2 From prodserv  a "
				strsql=strsql & "inner join subtema b on (a.marcaps=b.tipobjeto) "
				strsql=strsql & "where codprodserv=0" & nCodigo
				Set rsObjx = OFV.crearconsultaEx(strsql, cn, 1, volver)

				XTipo=rsObjx("descri2")
				xdescri=rsObjx("descri")
				Call ofv.cerrarconsulta(rsObjx)
			Else
				XTipo  = rsObj(2)
			End If

			strsql = "Select codobjmaster,tipobjmaster From relaobj Where codojasoc =" & nCodigo
			strsql = strsql & " and tipobjasoc = '" & XTipo	 & "' "
			' response.write strsql

			Set rsNew = OFV.crearconsultaEx(strsql, cn, 1, volver)



			bSector = "S/D"
			bControl = "N"
			Nsect = 0
			If Not rsNew.EOF Then
				Nsect = rsNew(0)
				For inp = 0 To UBound(asector)
					If asector(inp) = UCase(rsNew(1)) Then
						bSector  = nsector(inp)
						bControl = "S"
						Exit For
					End If
				Next

				If bControl = "N" Then
					For inp = 0 To UBound(aevento)
						If aevento(inp) = UCase(rsNew(1)) Then
							bSector  = nevento(inp)
							bControl = "O"
							Exit For
						End If
					Next
				End If

				If bControl = "O" Then
					strsql = "Select Objetos.*,objprty.estado From Objetos, Objprty where "
					strsql = strsql & "Objetos.codobjeto = Objprty.codobjeto and objetos.codobjeto = " & rsNew(0)

					Set rsNew = OFV.crearconsultaEx(strsql, cn, 1, volver)
					If Not rsNew.EOF Then
						bSector = bSector & ": " & rsNew(1) & " - " & rsNew(3)
						For inp = 0 To UBound(aestr)
							If aestr(inp) = UCase(rsNew(9)) Then
								bSector = bSector & " (" & nestr(inp) & ")"
								Exit For
							End If
						Next

					End If
				Else
					strsql = "Select Tema.descri, Prodserv.* From Tema, Prodserv where "
					strsql = strsql & "Tema.codtema = Prodserv.codtema and prodserv.codprodserv = " & rsNew(0)
					Set rsNew = OFV.crearconsultaEx(strsql, cn, 1, volver)
					If Not rsNew.EOF Then
						bSector = bSector & ": " & rsNew(0) & "/" & rsNew(2)
					End If

				End If

				Call OFV.cerrarconsulta(rsNew)
			End If

			bEvento = "S/D"




			'''sEmpresa

			If  sEmpresa = "" Then sEmpresa = 1
			bempresa = sEmpresa
			sLink    = "&nbsp;"

			Fila = ""

			If  Session("EMPMUN") = "S" Then

				For IAEMP=1 To TAEMP

					If ATREMP(IAEMP,2) = 1 Then
						If ATREMP(IAEMP,0) = sEmpresa Then
							bempresa = ATREMP(IAEMP,4)
							sLink    = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(0,5),"0","")
							Exit For
						End If
					End If

				Next

				'empresa
				Celda = OFV.GenCelda("", "ttt ", "", "", "", "", "", sLink, "si")
				sInput = "&nbsp;&nbsp;" & bempresa
				Celda = Celda & OFV.GenCelda("", "ttt", "", "", "", "", "", sInput, "si")
				Fila = Fila & OFV.GenRow("", "", "", "", "30", "", "", Celda, "si")

				Celda = OFV.GenCelda("", "ut colspan=2 ", "", "", "", "", "", "&nbsp;&nbsp;", "si")
				Fila = Fila & OFV.GenRow("", "", "", "", "30", "", "", Celda, "si")

			End If

			'origen
			'  Celda = "" 'OFV.GenCelda("", "tt", "", "", "", "", "", " Copiar:" , "si")
			'  sInput = "" 'bEvento & ": " & rsObj(1) & " - " & rsObj(3)
			'  For inp = 0 To UBound(aestr)
			'      If aestr(inp) = UCase(rsObj(9)) Then
			'         sInput = sInput & " (" & nestr(inp) & ")"
			'        Exit For
			'     End If
			' Next
			'Celda = Celda & OFV.GenCelda("", "UTI", "", "", "", "", "", sInput, "si")
			'Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")




			'Superior
			' Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "Perteneciente a:", "si")
			'sInput = bSector
			' Celda = Celda & OFV.GenCelda("", "UTI", "", "", "", "", "", sInput, "si")
			' Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

			'Como
			Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "Como:", "si")
			sInput = sOpcion
			For inp = 0 To UBound(aevento)
				If aevento(inp) = UCase(sOpcion) Then
					sInput = nevento(inp)
					Exit For
				End If
			Next
			'   Celda = Celda & OFV.GenCelda("", "UTI", "", "", "", "", "", sInput, "si")
			'   Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

			'Descripcion
			Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "Descripcion", "si")
			If rsObj.eof Then
				sInput=xdescri
			Else
				sInput = OFV.GenerarInput(rsObj(3).Name & sREADONLY, rsObj(3), "agregar", "Copiar", "80", "dt", "")
			End If

			Celda = Celda & OFV.GenCelda("", "UTI", "", "", "", "", "", sInput, "si")
			Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

			'Identificacion
			Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "Identificacion", "si")
			If rsObj.eof Then

				sInput=xtipo
			Else
				sInput = OFV.GenerarInput(rsObj(1).Name & sREADONLY, rsObj(1), "agregar", "copiar", "20", "dt", "")
			End If
			Celda = Celda & OFV.GenCelda("", "UTI", "", "", "", "", "", sInput, "si")
			Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
			aTipos=""
			aTipos0=""
			If aTipos <> ")" Then
				strsql = "Select Objetos.*,objprty.estado From Objetos, Objprty,TIPOBJETO where "
				If session("base")="K" Then
					strsql = strsql & " (TIPOBJETO.TIPOBJETO=OBJETOS.TIPOBJETO AND TIPOCPT='K') AND "
				Else
					strsql = strsql & " (TIPOBJETO.TIPOBJETO=OBJETOS.TIPOBJETO AND TIPOCPT <> 'K') AND "
				End If
				strsql = strsql & "Objetos.codobjeto = Objprty.codobjeto and (Objprty.estado = 1 "
				strsql = strsql & "Or Objprty.estado = 3) " '&"And Objetos.tipobjeto In " & aTipos  & "Select nivel From Tipobjeto Where tipobjeto = '" & sOpcion & "'"
				' response.write strsql & "<br>"
				Set rsTipo2 = OFV.crearconsultaEx(strsql, cn, 1, volver)
			End If
			If aTipos0 <> ")" Then
				strsql = "Select Tema.descri, Prodserv.* From Tema, Prodserv where "
				strsql = strsql & "Tema.codtema = Prodserv.codtema " '&" and Prodserv.marcaps In " & aTipos0
				If session("base")="K" Then
					strsql = strsql & " AND Tema.PSTIPO='K'"
				Else
					strsql = strsql & " AND Tema.PSTIPO <> 'K'"
				End If
				'  response.write strsql & "<br>"
				Set rsTipo0 = OFV.crearconsultaEx(strsql, cn, 1, volver)
			End If

			'Esquema superior
			Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "Esquema Superior" , "si")
			sInput = "<font class=at><select name=Esquema>"
			If aTipos <> ")" Then
				sInput = sInput & "<option value='nsnc' STYLE='border-style=solid;border-width:1;border-color:#eeeeee;"
				sInput = sInput & "background-color:#6699cc;color:#ffffff;' >"
				sInput = sInput & "+++ ESQUEMAS ++++++++++++++++++++++++++++++++++++"
				sInput = sInput & "</OPTION>"

				Do Until rsTipo2.EOF

					sInput = sInput & "<option value='" & rsTipo2(0) & "_" & rsTipo2(2) & "' "
					If  bControl = "O" Then
						If  Cstr(Nsect) = Cstr(rsTipo2(0)) Then
							sInput = sInput & " selected "
						End If
					End If
					sInput = sInput & " >"

					For inp = 0 To UBound(aevento)
						If  aevento(inp) = UCase(rsTipo2(2)) Then
							sInput       = sInput & nevento(inp) & ": "
							Exit For
						End If
					Next

					sInput = sInput & rsTipo2(1) & " - "

					wwwtipo2 = rsTipo2(3)

					If len(wwwtipo2 & "") > 40 Then
						sInput = sInput & left(wwwtipo2,40) & "..(+)"
					Else
						sInput = sInput & wwwtipo2
					End If


					''''      sInput = sInput & rsTipo2(3)

					For inp = 0 To UBound(aestr)
						If  aestr(inp) = UCase(rsTipo2(9)) Then
							sInput     = sInput & " (" & nestr(inp) & ")"
							Exit For
						End If
					Next

					rsTipo2.movenext
				Loop
			End If

			If aTipos0 <> ")" Then

				sInput = sInput & "<option value='nsnc' STYLE='border-style=solid;border-width:1;border-color:#eeeeee;"
				sInput = sInput & "background-color:#6699cc;color:#ffffff;' >"
				sInput = sInput & "+++ SUBTEMAS ++++++++++++++++++++++++++++++++++++"
				sInput = sInput & "</OPTION>"


				Do Until rsTipo0.EOF
					sInput = sInput & "<option value='" & rsTipo0(1) & "_" & rsTipo0(5) & "' "
					If  bControl = "S" Then
						If  CStr(Nsect) = CStr(rsTipo0(1)) Then
							sInput  = sInput & " selected "
						End If
					End If
					sInput = sInput & " >"

					For inp = 0 To UBound(asector)
						If  asector(inp) = UCase(rsTipo0(5)) Then
							sInput  = sInput & nsector(inp) & ": "
							Exit For
						End If
					Next

					sInput = sInput & rsTipo0(0)
					sInput = sInput & "/" & rsTipo0(2) & "/"

					rsTipo0.movenext
				Loop
			End If

			sInput = sInput & "</select></font>"
			Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "Esquema Superior" , "si")
			Celda  = Celda & OFV.GenCelda("", "UTI", "", "", "", "", "", sInput, "si")
			Fila   = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

			Tabla  = OFV.GenTabla("", "", "", "80%", "", "0", "", OFV.wspacing, 1 + OFV.wspacing, "", "", Fila, "si")

			If  aTipos  <> ")"  Then Call OFV.cerrarconsulta(rsTipo2)
			If  aTipos0 <> ")"  Then Call OFV.cerrarconsulta(rsTipo0)

			Tabla      = Tabla & "<br><br>" & OFV.GenTabla("", "", "", "100%", "", "0", "0", "0", "0", "", "", "", "no")
			Fila       = Tabla & OFV.GenRow("", "", "", "", "", "", "", "", "no")

			sInput     = OFV.GenerarInput("", "Aceptar", "submit", "", "10", "bt", "this.style.display='none';")
			Celda      = Fila & OFV.GenCelda("", "", "center", "50%", "", "", "", sInput, "si")
			sAccion    = "Copiar2.asp?Codigo=" & nCodigo & "&Tipo=" & sTipo & "&Perfil=" & sPerfil & "&empresa=" & sEmpresa
			sAccion    = sAccion & "&Opcion=" & sOpcion & "&Copiar=SI" & "&moveresquema=" & moveresquema & "&ESQSUP=" & ESQSUP
			Formulario = OFV.GenForm("Copiar", "", sAccion, "", "post", "", Celda, "si")

			sFunc      = "window.close()"
			sInput     = OFV.GenerarInput("", "Cancelar", "button", "", "10", "bt", sFunc)
			Celda      = OFV.GenCelda("", "", "center", "50%", "", "", "", sInput, "si")
			Formulario = Formulario & OFV.GenForm("", "", "", "", "post", "", Celda, "si")
			Fila       = Formulario & OFV.GenFRow()

			Tabla      = Fila & OFV.GenFTabla()

			sHTML      = sHTML & "<br><br><br><br><center>" & Tabla & "</center>"
			sHTML      = sHTML & "</body></html>"

			Call OFV.cerrarconsulta(rsObj)


			Case Else ' SOLO ESQUEMA
			'RESPONSE.WRITE "ACA " &  SBASEF

			If SBASEF ="PEDIDO" Then

				strsql = " select b.* ,a.accion,a.objun from pedidos a"
				strsql=strsql & " inner join pedidosdet b on (a.objid=b.idpedido)"
				strsql= strsql & "   where b.idPEDIDO=" & IDPEDIDO

				Set rsTipo0 = OFV.crearconsultaEx(strsql, cn, 1, volver)

				If Not rsTipo0.eof Then
					pedun=rsTipo0("objun")
				Else
					pedun=1
				End If
				Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "descripcion:", "si")


				SINPUT= "<input  type=text  size=100%  value='" &rsTipo0(15)&"' name=Descri class=dt >"
				Celda  = Celda & OFV.GenCelda("", "UTI", "", "", "", "", "", sInput, "si")
				Fila   =   OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

				Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "Como:", "si")

				Strsql = " Select * From tipobjeto where tipocpt='K' Order By descri "
				Set rsTipo0 = OFV.crearconsultaEx(strsql, cn, 1, volver)

				sInput = "<font class=at><select name=tipo>"

				Do Until rsTipo0.EOF
					sInput = sInput & "<option value='" & rsTipo0(0) & "' "

					sInput = sInput & " >"

					sInput = sInput & rsTipo0(1)
					'sInput = sInput & "/" & rsTipo0(2) & "/"

					rsTipo0.movenext
				Loop


				sInput = sInput & "</select></font>"
				Celda  = Celda & OFV.GenCelda("", "UTI", "", "", "", "", "", sInput, "si")



				Fila   =  fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

				strsql = "Select Tema.descri, Prodserv.* From Tema, Prodserv where "
				strsql = strsql & " Tema.codtema = Prodserv.codtema and Prodserv.pstipo='K' and marcaps <> 'PEDIDO' order by tema.descri,prodserv.descri"
				Set rsTipo0 = OFV.crearconsultaEx(strsql, cn, 1, volver)

				Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "Esquema Superior" , "si")
				sInput = "<font class=at><select name=elemento>"

				Do Until rsTipo0.EOF
					sInput = sInput & "<option value='" & rsTipo0("codprodserv") & "-" & rsTipo0("marcaps") & "Y-N' "

					sInput = sInput & " >"

					sInput = sInput & rsTipo0(0)
					sInput = sInput & "/" & rsTipo0(2) & "/"

					rsTipo0.movenext
				Loop


				sInput = sInput & "</select></font>"
				Celda  = Celda & OFV.GenCelda("", "UTI", "", "", "", "", "", sInput, "si")
				Fila   =  fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")



				Tabla  = OFV.GenTabla("", "", "", "80%", "", "0", "", OFV.wspacing, 1 + OFV.wspacing, "", "", Fila, "si")

				Tabla      = Tabla & "<br><br>" & OFV.GenTabla("", "", "", "100%", "", "0", "0", "0", "0", "", "", "", "no")
				Fila       = Tabla & OFV.GenRow("", "", "", "", "", "", "", "", "no")

				sInput     = OFV.GenerarInput("", "Aceptar", "submit", "", "10", "bt", "this.style.display='none';")
				If  Session("EMPMUN")  = "S"Then
					XDATO=  "<input  type=HIDDEN  value='"& pedun &"' name=OBJUN >"
				Else
					XDATO=  "<input  type=HIDDEN  value='1' name=OBJUN >"
				End If
				XDATO= XDATO & "<input  type=HIDDEN  value='EM' name=aplic >"
				XDATO= XDATO & "<input  type=HIDDEN  value='1' name=prioridad >"
				XDATO= XDATO & "<input  type=HIDDEN  size=17%  value='' name=Codiden class=dt >"


				Celda      = Fila & OFV.GenCelda("", "", "center", "50%", "", "", "", sInput & XDATO, "si")
				SACCION="asp/AgregarEsquema.asp?acc=ALTAPED&idpedido=" & idpedido
				Formulario = OFV.GenForm("Copiar", "", sAccion, "", "post", "", Celda, "si")

				sFunc      = "window.top.close()"
				sInput     = OFV.GenerarInput("", "Cancelar", "button", "", "10", "bt", sFunc)
				Celda      = OFV.GenCelda("", "", "center", "50%", "", "", "", sInput, "si")
				Formulario = Formulario & OFV.GenForm("", "", "", "", "post", "", Celda, "si")
				Fila       = Formulario & OFV.GenFRow()

				Tabla      = Fila & OFV.GenFTabla()

				sHTML      = sHTML & "<br><br><br><br><center>" & Tabla & "</center>"
				sHTML      = sHTML & "</body></html>"
				' RESPONSE.WRITE IDPEDIDO






			Else



				strsql = "select * from estados"
				Set rsTipo = OFV.crearconsultaEx(strsql, cn, 1, volver)
				inp = 0
				Do Until rsTipo.EOF
					ReDim Preserve aestr(inp)
					ReDim Preserve nestr(inp)
					aestr(inp) = UCase(rsTipo(0))
					nestr(inp) = UCase(rsTipo(1))
					inp = inp + 1
					rsTipo.movenext
				Loop

				If session("base")="K" Then
					strsql = "select tipobjeto,descri from subtema where icono='K'"
				Else
					strsql = "select tipobjeto,descri from subtema where icono <>'K'"
				End If
				'response.write strsql
				Set rsTipo = OFV.crearconsultaEx(strsql, cn, 1, volver)
				inp = 0
				Do Until rsTipo.EOF
					ReDim Preserve asector(inp)
					ReDim Preserve nsector(inp)
					asector(inp) = UCase(rsTipo(0))
					nsector(inp) = UCase(rsTipo(1))
					inp = inp + 1
					rsTipo.movenext
				Loop

				If session("base")="K" Then
					strsql = "select tipobjeto,descri from tipobjeto" & WHERETIPO & " and TIPOCPT='K'"
				Else
					strsql = "select tipobjeto,descri from tipobjeto" & WHERETIPO & " and TIPOCPT <>'K'"
				End If
				' response.write strsql
				Set rsTipo = OFV.crearconsultaEx(strsql, cn, 1, volver)
				inp = 0
				Do Until rsTipo.EOF
					ReDim Preserve aevento(inp)
					ReDim Preserve nevento(inp)
					aevento(inp) = UCase(rsTipo(0))
					nevento(inp) = UCase(rsTipo(1))
					inp = inp + 1
					rsTipo.movenext
				Loop
				Call OFV.cerrarconsulta(rsTipo)

				strsql = "Select Objetos.*,objprty.estado From Objetos, Objprty where "
				strsql = strsql & "Objetos.codobjeto = Objprty.codobjeto and objetos.codobjeto = " & nCodigo


				Set rsObj = OFV.crearconsultaEx(strsql, cn, 1, volver)
				If  rsObj.EOF Then
					XTipo=sTipo
				Else
					XTipo  = rsObj(2)
				End If
				strsql = "Select codobjmaster,tipobjmaster From relaobj Where codojasoc =" & nCodigo
				strsql = strsql & " and tipobjasoc = '" & XTipo	 & "' "
				' response.write strsql

				Set rsNew = OFV.crearconsultaEx(strsql, cn, 1, volver)



				bSector = "S/D"
				bControl = "N"
				Nsect = 0
				If Not rsNew.EOF Then
					Nsect = rsNew(0)
					For inp = 0 To UBound(asector)
						If asector(inp) = UCase(rsNew(1)) Then
							bSector  = nsector(inp)
							bControl = "S"
							Exit For
						End If
					Next

					If bControl = "N" Then
						For inp = 0 To UBound(aevento)
							If aevento(inp) = UCase(rsNew(1)) Then
								bSector  = nevento(inp)
								bControl = "O"
								Exit For
							End If
						Next
					End If

					If bControl = "O" Then
						strsql = "Select Objetos.*,objprty.estado From Objetos, Objprty where "
						strsql = strsql & "Objetos.codobjeto = Objprty.codobjeto and objetos.codobjeto = " & rsNew(0)

						Set rsNew = OFV.crearconsultaEx(strsql, cn, 1, volver)
						If Not rsNew.EOF Then
							bSector = bSector & ": " & rsNew(1) & " - " & rsNew(3)
							For inp = 0 To UBound(aestr)
								If aestr(inp) = UCase(rsNew(9)) Then
									bSector = bSector & " (" & nestr(inp) & ")"
									Exit For
								End If
							Next

						End If
					Else
						strsql = "Select Tema.descri, Prodserv.* From Tema, Prodserv where "
						strsql = strsql & "Tema.codtema = Prodserv.codtema and prodserv.codprodserv = " & rsNew(0)
						Set rsNew = OFV.crearconsultaEx(strsql, cn, 1, volver)
						If Not rsNew.EOF Then
							bSector = bSector & ": " & rsNew(0) & "/" & rsNew(2)
						End If

					End If

					Call OFV.cerrarconsulta(rsNew)
				End If

				bEvento = "S/D"

				For inp = 0 To UBound(aevento)
					If aevento(inp) = UCase(rsObj(2)) Then
						bEvento = nevento(inp)
						Exit For
					End If
				Next


				'''sEmpresa

				If  sEmpresa = "" Then sEmpresa = 1
				bempresa = sEmpresa
				sLink    = "&nbsp;"

				Fila = ""

				If  Session("EMPMUN") = "S" Then

					For IAEMP=1 To TAEMP

						If ATREMP(IAEMP,2) = 1 Then
							If ATREMP(IAEMP,0) = sEmpresa Then
								bempresa = ATREMP(IAEMP,4)
								sLink    = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(0,5),"0","")
								Exit For
							End If
						End If

					Next

					'empresa
					Celda = OFV.GenCelda("", "ttt ", "", "", "", "", "", sLink, "si")
					sInput = "&nbsp;&nbsp;" & bempresa
					Celda = Celda & OFV.GenCelda("", "ttt", "", "", "", "", "", sInput, "si")
					Fila = Fila & OFV.GenRow("", "", "", "", "30", "", "", Celda, "si")

					Celda = OFV.GenCelda("", "ut colspan=2 ", "", "", "", "", "", "&nbsp;&nbsp;", "si")
					Fila = Fila & OFV.GenRow("", "", "", "", "30", "", "", Celda, "si")

				End If

				'origen
				Celda = OFV.GenCelda("", "tt", "", "", "", "", "", " Copiar:" , "si")
				sInput = bEvento & ": " & rsObj(1) & " - " & rsObj(3)
				For inp = 0 To UBound(aestr)
					If aestr(inp) = UCase(rsObj(9)) Then
						sInput = sInput & " (" & nestr(inp) & ")"
						Exit For
					End If
				Next
				Celda = Celda & OFV.GenCelda("", "UTI", "", "", "", "", "", sInput, "si")
				Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")




				'Superior
				Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "Perteneciente a:", "si")
				sInput = bSector
				Celda = Celda & OFV.GenCelda("", "UTI", "", "", "", "", "", sInput, "si")
				Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

				'Como
				Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "Como:", "si")
				sInput = sOpcion
				For inp = 0 To UBound(aevento)
					If aevento(inp) = UCase(sOpcion) Then
						sInput = nevento(inp)
						Exit For
					End If
				Next
				Celda = Celda & OFV.GenCelda("", "UTI", "", "", "", "", "", sInput, "si")
				Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

				'Descripcion
				Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "Descripcion", "si")
				sInput = OFV.GenerarInput(rsObj(3).Name & sREADONLY, rsObj(3), "agregar", "Copiar", "80", "dt", "")
				Celda = Celda & OFV.GenCelda("", "UTI", "", "", "", "", "", sInput, "si")
				Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

				'Identificacion
				Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "Identificacion", "si")
				sInput = OFV.GenerarInput(rsObj(1).Name & sREADONLY, rsObj(1), "agregar", "copiar", "20", "dt", "")
				Celda = Celda & OFV.GenCelda("", "UTI", "", "", "", "", "", sInput, "si")
				Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

				strsql = "Select nivel From Tipobjeto Where tipobjeto = '" & sOpcion & "'"
				'response.write strsql
				Set rsTipo = OFV.crearconsultaEx(strsql, cn, 1, volver)
				sTipo = rsTipo(0)
				Call OFV.cerrarconsulta(rsTipo)
				If session("base")="K" Then
					ADDW=" ICONO='K' AND "
				Else
					ADDW=" ICONO <> 'K' AND "
				End If
				Select Case sTipo
					Case "C"
					strsql = "Select tipobjeto From Subtema Where " & ADDW & " levc = 1"
					Case "D"
					strsql = "Select tipobjeto From Subtema Where " & ADDW & " levd = 1"
					Case "E"
					strsql = "Select tipobjeto From Subtema Where " & ADDW & " leve = 1"
					Case "F"
					strsql = "Select tipobjeto From Subtema Where " & ADDW & " levf = 1"
					Case "G"
					strsql = "Select tipobjeto From Subtema Where " & ADDW & " levg = 1"
					Case "H"
					strsql = "Select tipobjeto From Subtema Where " & ADDW & " levh = 1"
					Case "I"
					strsql = "Select tipobjeto From Subtema Where " & ADDW & " levi = 1"
					Case "J"
					strsql = "Select tipobjeto From Subtema Where " & ADDW & " levj = 1"
					Case "K"
					strsql = "Select tipobjeto From Subtema Where " & ADDW & " levk = 1"
					Case "L"
					strsql = "Select tipobjeto From Subtema Where " & ADDW & " levl = 1"
				End Select
				' response.write strsql
				Set rsTipo0 = OFV.crearconsultaEx(strsql, cn, 1, volver)
				aTipos0 = "('"
				Do Until rsTipo0.EOF
					aTipos0 = aTipos0 & rsTipo0(0) & "','"
					rsTipo0.movenext
				Loop
				aTipos0 = Left(aTipos0, Len(aTipos0) - 2) & ")"

				Select Case sTipo
					Case "C"
					strsql = "Select tipobjeto From Tipobjeto Where levc = 1 " & wheretipo2
					Case "D"
					strsql = "Select tipobjeto From Tipobjeto Where levd = 1" & wheretipo2
					Case "E"
					strsql = "Select tipobjeto From Tipobjeto Where leve = 1" & wheretipo2
					Case "F"
					strsql = "Select tipobjeto From Tipobjeto Where levf = 1" & wheretipo2
					Case "G"
					strsql = "Select tipobjeto From Tipobjeto Where levg = 1" & wheretipo2
					Case "H"
					strsql = "Select tipobjeto From Tipobjeto Where levh = 1" & wheretipo2
					Case "I"
					strsql = "Select tipobjeto From Tipobjeto Where levi = 1" & wheretipo2
					Case "J"
					strsql = "Select tipobjeto From Tipobjeto Where levj = 1" & wheretipo2
					Case "K"
					strsql = "Select tipobjeto From Tipobjeto Where levk = 1" & wheretipo2
					Case "L"
					strsql = "Select tipobjeto From Tipobjeto Where levl = 1" & wheretipo2
				End Select
				Set rsTipo = OFV.crearconsultaEx(strsql, cn, 1, volver)
				aTipos = "('"
				Do Until rsTipo.EOF
					aTipos = aTipos & rsTipo(0) & "','"
					rsTipo.movenext
				Loop
				aTipos = Left(aTipos, Len(aTipos) - 2) & ")"
				If aTipos <> ")" Then
					strsql = "Select Objetos.*,objprty.estado From Objetos, Objprty where 1=2 and "
					strsql = strsql & "Objetos.codobjeto = Objprty.codobjeto and (Objprty.estado = 1 "
					strsql = strsql & "Or Objprty.estado = 3) And Objetos.tipobjeto In " & aTipos
					strsql = strsql &  " order by objetos.descri"
					' response.write strsql & "<br>"
					Set rsTipo2 = OFV.crearconsultaEx(strsql, cn, 1, volver)
				End If
				If aTipos0 <> ")" Then
					strsql = "Select Tema.descri, Prodserv.* From Tema, Prodserv where "
					strsql = strsql & "Tema.codtema = Prodserv.codtema and Prodserv.marcaps In " & aTipos0
					strsql = strsql &  " order by Tema.descri"
					' response.write strsql & "<br>"
					Set rsTipo0 = OFV.crearconsultaEx(strsql, cn, 1, volver)
				End If
				'response.write strsql
				'Esquema superior
				Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "Esquema Superior", "si")
				sInput = "<font class=at><select name=Esquema>"
				If aTipos <> ")" Then
					sInput = sInput & "<option value='nsnc' STYLE='border-style=solid;border-width:1;border-color:#eeeeee;"
					sInput = sInput & "background-color:#6699cc;color:#ffffff;' >"
					sInput = sInput & "+++ ESQUEMAS ++++++++++++++++++++++++++++++++++++"
					sInput = sInput & "</OPTION>"

					Do Until rsTipo2.EOF

						sInput = sInput & "<option value='" & rsTipo2(0) & "_" & rsTipo2(2) & "' "
						If  bControl = "O" Then
							If  Cstr(Nsect) = Cstr(rsTipo2(0)) Then
								sInput = sInput & " selected "
							End If
						End If
						sInput = sInput & " >"

						For inp = 0 To UBound(aevento)
							If  aevento(inp) = UCase(rsTipo2(2)) Then
								sInput       = sInput & nevento(inp) & ": "
								Exit For
							End If
						Next

						sInput = sInput & rsTipo2(1) & " - "

						wwwtipo2 = rsTipo2(3)

						If len(wwwtipo2 & "") > 40 Then
							sInput = sInput & left(wwwtipo2,40) & "..(+)"
						Else
							sInput = sInput & wwwtipo2
						End If


						''''      sInput = sInput & rsTipo2(3)

						For inp = 0 To UBound(aestr)
							If  aestr(inp) = UCase(rsTipo2(9)) Then
								sInput     = sInput & " (" & nestr(inp) & ")"
								Exit For
							End If
						Next

						rsTipo2.movenext
					Loop
				End If

				If aTipos0 <> ")" Then

					sInput = sInput & "<option value='nsnc' STYLE='border-style=solid;border-width:1;border-color:#eeeeee;"
					sInput = sInput & "background-color:#6699cc;color:#ffffff;' >"
					sInput = sInput & "+++ SUBTEMAS ++++++++++++++++++++++++++++++++++++"
					sInput = sInput & "</OPTION>"


					Do Until rsTipo0.EOF
						sInput = sInput & "<option value='" & rsTipo0(1) & "_" & rsTipo0(5) & "' "
						If  bControl = "S" Then
							If  CStr(Nsect) = CStr(rsTipo0(1)) Then
								sInput  = sInput & " selected "
							End If
						End If
						sInput = sInput & " >"

						For inp = 0 To UBound(asector)
							If  asector(inp) = UCase(rsTipo0(5)) Then
								sInput  = sInput & nsector(inp) & ": "
								Exit For
							End If
						Next

						sInput = sInput & rsTipo0(0)
						sInput = sInput & "/" & rsTipo0(2) & "/"

						rsTipo0.movenext
					Loop
				End If

				sInput = sInput & "</select></font>"
				' aca 19_9
				WSPFRMN="TS1"

				sImagenE = "<span style='display:inline-block; width: 10;'></span><B><FONT FACE=CorFlaticon COLOR=#F7AA22 SIZE=3>&#0232;</FONT></B>"

				' divdst="<p id=dst name=dst>www</p>"

				'sqlpath="select mc1.descri,mc2.descri,mc3.descri,mc4.descri,mc5.descri,mc6.descri,mc7.descri,a.mc7,a.mt7 from  csrelpath a "
				'sqlpath= sqlpath & " left join objetos mc1 on(mc1.codobjeto=a.mc1 and mc1.tipobjeto=a.mt1 )"
				'sqlpath= sqlpath & "left join objetos mc2 on(mc2.codobjeto=a.mc2 and mc2.tipobjeto=a.mt2 ) "
				'sqlpath= sqlpath & "left join objetos mc3 on(mc3.codobjeto=a.mc3 and mc3.tipobjeto=a.mt3 ) "
				'sqlpath= sqlpath & "left join objetos mc4 on(mc4.codobjeto=a.mc4 and mc4.tipobjeto=a.mt4 ) "
				'sqlpath= sqlpath & "left join objetos mc5 on(mc5.codobjeto=a.mc5 and mc5.tipobjeto=a.mt5) "
				'sqlpath= sqlpath & "left join objetos mc6 on(mc6.codobjeto=a.mc6 and mc6.tipobjeto=a.mt6) "
				'sqlpath= sqlpath & "left join objetos mc7 on(mc7.codobjeto=a.mc7 and mc7.tipobjeto=a.mt7 )"
				'sqlpath= sqlpath & " where a.cobj=0" & nCodigo


				sqlpath="select* FROM CSRELPATHDESCRIEX2  A where a.cobj=" & nCodigo

				'  response.write sqlpath
				Set rspth = ofv.crearconsultaEx(sqlpath,cn,1,parametros)
				stpath=""
				If Not rspth.eof Then
					stpath= rspth("temadescri") & "/" &   rspth("stemadescri")
					For i=1 To 6
						If trim(rspth("expr" & i) & "") <> "" Then
							stpath= stpath & "/" & rspth("expr" & i)
						End If
					Next
					sinput="<input  type=hidden id=esquema name=esquema value=" &rspth("mc7") & "_" & rspth("mt7") &">"

				End If
				Call ofv.cerrarconsulta(rspth)
				shrf="<a  title='ASIGNAR' style='cursor: pointer; cursor: hand;' onclick=document.getElementById('TS1').submit();><span id=dst name=dst>" & stpath & "</span> "& sImagenE & " </A>"

				Celda  = Celda & OFV.GenCelda("", "UTI", "", "", "", "", "", shrf &  sInput, "si")
				Fila   = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

				Tabla  = OFV.GenTabla("", "", "", "80%", "", "0", "", OFV.wspacing, 1 + OFV.wspacing, "", "", Fila, "si")

				If  aTipos  <> ")"  Then Call OFV.cerrarconsulta(rsTipo2)
				If  aTipos0 <> ")"  Then Call OFV.cerrarconsulta(rsTipo0)

				Tabla      = Tabla & "<br><br>" & OFV.GenTabla("", "", "", "100%", "", "0", "0", "0", "0", "", "", "", "no")
				Fila       = Tabla & OFV.GenRow("", "", "", "", "", "", "", "", "no")

				sInput     = OFV.GenerarInput("", "Aceptar", "submit", "", "10", "bt", "this.style.display='none';")
				Celda      = Fila & OFV.GenCelda("", "", "center", "50%", "", "", "", sInput, "si")
				sAccion    = "Copiar2.asp?Codigo=" & nCodigo & "&Tipo=" & sOpcion & "&Perfil=" & sPerfil & "&empresa=" & sEmpresa
				sAccion    = sAccion & "&Opcion=" & sOpcion & "&Copiar=SI" & "&moveresquema=" & moveresquema & "&ESQSUP=" & ESQSUP & "&ACC=" & SACC & "&IDPEDIDO=" & IDPEDIDO
				Formulario = OFV.GenForm("Copiar", "", sAccion, "", "post", "", Celda, "si")

				sFunc      = "window.close()"
				sInput     = OFV.GenerarInput("", "Cancelar", "button", "", "10", "bt", sFunc)
				Celda      = OFV.GenCelda("", "", "center", "50%", "", "", "", sInput, "si")
				Formulario = Formulario & OFV.GenForm("", "", "", "", "post", "", Celda, "si")
				Fila       = Formulario & OFV.GenFRow()

				Tabla      = Fila & OFV.GenFTabla()

				sHTML      = sHTML & "<br><br><br><br><center>" & Tabla & "</center>"

				WSPFRMA = "conscopiador.asp?id=" & ncodigo & "&tid=" & sTipobjeto

				WSPFORME = "<FORM ID='" & WSPFRMN & "' NAME='" & WSPFRMN & "' ACTION='" & WSPFRMA & "' TARGET='xxxx' style='display:none;visibility:hidden;' METHOD='POST' ></FORM>"
				jsc =   chr(13) & "<SCRIPT language=javascript >" & chr(13)
				jsc = jsc & chr(13)
				jsc = jsc & "function asignar(e) {" & chr(13)
				jsc = jsc & " document.getElementById('dst').innerText=document.getElementById(e).innerText;"
				jsc = jsc & " document.getElementById('esquema').value=e;"
				jsc = jsc & " document.getElementById('DTTEMAINT').style.visibility='hidden';"
				jsc = jsc & " } " & chr(13)
				jsc = jsc & " </SCRIPT> " & chr(13)

				sHTML      = sHTML & WSPFORME & "<HR><div id=DTTEMAINT style='overflow:scroll;height:400px;visibility:hidden;'></div><iframe style='display:none;visibility:hidden;' id=xxxx name=xxxx></iframe>" & jsc &"</body></html>"

				Call OFV.cerrarconsulta(rsObj)
			End If


		End Select ' esqsup

		Call OFV.cerrarconn(cn)

	End If








End If

response.write sHTML

%>