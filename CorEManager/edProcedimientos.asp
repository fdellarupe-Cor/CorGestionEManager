<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%
CONT = 0
ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0

	XXXX = GENLOGEXT("EDIDET")


	Set cn = ofv.conectar(ofv.strconn2)

	nCodigo = request.querystring("Codobjeto")
	sTipo = request.querystring("Tipobjeto")

	''sTest = request.querystring("Test")

	histobs = request.querystring("histobs")

	If histobs = "S" Then

		niv1x=request.querystring("NIV1")

		session("codSol")=niv1x

	End If

	sImageobs = "&nbsp;<font face='Corflaticon' color='#990000' size=3 >&#095;</font>"

	sopcion = request.querystring("sopcion")

	If sopcion <> "" Then

		sevento = request.querystring("sevento")
		sukey = GTkey(nCodigo)

		strsql = "insert into objproced  (codobjeto,tipobjeto,codproced,codeven,frecuencia,tipproced) values ("
		strsql = strsql & nCodigo & ",'" & stipo & "','" & sukey & "'," & sevento & ",'DIARIA','" & wkaplic & "')"
		Call ofv.crearconsultaEx(StrSql,cn,1,volver)

		'   response.write strsql

		response.redirect "edProcedDET.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&XEVEN=" & sukey


	Else

		vuelta = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover)


		sHTML = ofv.FormHeader(session("FormN"))

		'----------------------------------------------

		'shtml = shtml & CorPnlFnc("N","")

		sHTML = sHTML &  "<div id='mt2'  style='xdisplay:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

		'-------------------------------------------------

		%>
		<!-- #INCLUDE FILE="IncFile/CorVinculadorPL.inc" -->
		<%


		If WKSIGUE = "S" Then

			ofv.ObtenerAtributos Session("Form2"),""
			bModificar = ofv.AModificar
			If session("Nomodificar") = "S" Then bModificar = False



			strsql = "Select Version,estado From Objprty Where Objprty.codobjeto = " & nCodigo
			Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			bVersion = "NO"
			If Not rs.eof Then
				If rs(0) = "V" Then bVersion = "SI"
				nestado = rs(1)
			End If
			Call ofv.cerrarconsulta(rs)

			siconmas = "ICorAcciones.jpg"
			Strsql = "Select * From Tipobjeto Where tipobjeto ='" & stipo & "'"
			Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not rs.eof Then
				sEsquema = rs(1)
				siconmas = rs("archivo")
			End If

			Call ofv.cerrarconsulta(rs)

			Strsql = "Select * From Objetos Where codobjeto=" & nCodigo
			Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not rs.eof Then


				bstitulox = sEsquema & ": " & rs(1) & " - " & rs(3)
				h1 = "<IMG SRC=Iconos\" & siconmas & " border=0>"
			End If
			Call ofv.cerrarconsulta(rs)


			Formulario = ""

			Celda      = ofv.GenCelda("","ttt COLSPAN=6 ","","","","","","<big>LISTA DE EVENTOS</BIG>","si")
			Formulario = Formulario & ofv.GenRow("","","","","","","",Celda,"si")

			Strsql = "Select * From Objeven Where codobjeto=" & nCodigo & " Order By secuen"
			Set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			Do Until rs1.eof
			CONT = CONT + 1
				' if BrowserName()="IE" or BrowserName()="CHROME" THEN
				sicK = "<i class='las la-file-alt'  style=color:" & session("cmfmenu") & " ></i>"
				' else
				' sicK = "<font face='corflaticon' size=4 color='" & session("cmfmenu") & "' >&#128196;</font>&nbsp;"
				'end if

				If bModificar Then
					sAccion = "edProcedimientos.asp?Codobjeto="
					sAccion = sAccion & nCodigo & "&Tipobjeto=" & sTipo & "&sopcion=A&sevento=" & rs1(0)

					sImagenP = "<font face='CorFlaticon' size=3 color=#005500 >&#097;</font>"
					sinput  = "<a class='nt'  href='"
					sinput  = sinput  & sAccion
					sinput  = sinput  & "' title='Agregar circuito para el evento " & rs1(3) & "' >" & sImagenP & "</a>"
					Agregar = sinput

					sinput  = "<a class='nt'  href='"
					sinput  = sinput  & sAccion
					sinput  = sinput  & "' title='Agregar circuito para el evento " & rs1(3) & "' >" & sicK & "</a>"
					sicK = sinput


				Else
					Agregar = "&nbsp;"
				End If

				Agregar = "&nbsp;"


				Celda = ofv.GenCelda("","","","1%","","","","&nbsp;","si")
				Celda = Celda & ofv.GenCelda("","","","1%","","","",Agregar,"si")
				Celda = Celda & ofv.GenCelda("","ut style='filter:shadow(color=#dddddd,direction=125);' ","","1%","","","",sicK,"si")
				Celda = Celda & ofv.GenCelda("","UT COLSPAN=3 ","","","","",""," <U>" & rs1(4) & " - " & rs1(3) & "</U>","si")

				Formulario = Formulario & ofv.GenRow("","","","","","","",Celda,"si")

				strsql = "SELECT csprocmaster.*, b.codinx1 From csprocmaster "
				' strsql = strsql & " left join (select distinct codinx1 from OBJOBSCMT where actpend is null and codsol = " & session("codSol") & ") b on (csprocmaster.codproced = b.codinx1) "
				' strsql = strsql & " left join (select distinct codinx1 from OBJOBSCMT where   codsol = " & session("codSol") & ") b on (csprocmaster.codproced = b.codinx1) "
				strsql = strsql & " left join (select distinct codinx1 from OBJOBSCMT where TIPO='OBS' AND  codsol = " & session("codSol") & " and codobjeto=" & nCodigo & ") b on (csprocmaster.codproced = b.codinx1) "
				Strsql = Strsql & "Where codobjeto = " & nCodigo & " and codeven = " & rs1(0)
				Strsql = Strsql & " order by secuen, secuencia, objid"

				'  response.write strsql & "<br>"

				Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
				Do Until rs.eof

					'   if BrowserName()="IE" or BrowserName()="CHROME" THEN
					'       sicK = "<font face='corflaticon' size=3 color=#993333 >2</font>"
					'       else
					sicK = "<i class='las la-sitemap' style=color:" & session("cmfmenu") & "></i>"
					'    end if

					Celda = ofv.GenCelda("","","","1%","","","","&nbsp;","si")
					Celda = Celda & ofv.GenCelda("","","","1%","","","","&nbsp;","si")


					nID   = rs(0)
					nID2  = rs(9)
					sEven = rs(2)

					If Not IsNull(seven) Then sEven = ofv.convertircar(sEven," ","_")

					sSector = " "
					Strsql = "Select eco.descri,clienteplan.cliente From Eco, Objsector, clienteplan where eco.estado=3 and "
					Strsql = Strsql & "Eco.eco = Objsector.codsector and eco.ecoUN=clienteplan.id and Objsector.objid = " & rs(5)
					Set rsSector = ofv.crearconsultaEx(StrSql,cn,1,parametros)
					If Not rsSector.eof Then
						' agregado 14/6
						If session("EMPMUN") = "S" Then
							sSector = rsSector(1) & "&nbsp;-&nbsp;" & rsSector(0)
						Else
							sSector =  rsSector(0)
						End If

					End If

					Call ofv.cerrarconsulta(rsSector)

					sPerfil = " "
					If rs(6) <> "" Then
						Strsql = "Select descri From Tipoperfil Where tipoperfil = '" & rs(6) & "'"
						Set rsPerfil = ofv.crearconsultaEx(StrSql,cn,1,parametros)
						If Not rsPerfil.eof Then sPerfil = rsPerfil(0)
						Call ofv.cerrarconsulta(rsPerfil)
					End If

					sfrec = ""
					sfrecn = ""
					If Not IsNull(rs(11))  Then
						Strsql = "Select descri From frecuencia Where codfrec = '" & rs(11) & "' "
						Set rsfrec = ofv.crearconsultaEx(StrSql,cn,1,parametros)
						If Not rsfrec.eof Then
							sfrec = "<I>(" & rsfrec(0) & ")</I>"
							sfrecn = "(" & rsfrec(0) & ")"
						End If
						Call ofv.cerrarconsulta(rsfrec)
					End If

					Celdax = " " & rs(8) & "." & rs(7) & " " & sSector & " - " &  sPerfil & " - " & sfrec
					Celdan =  rs(8) & "." & rs(7) & " " & sSector & " - " &  sPerfil & " - " & sfrecn


					'Eliminar
					If  bModificar  And (rs(13) & "") <> "PL" Then
						sHref = "asp/EliminarEN.asp?ID=" & nID2 & "&Tabla=Objproced&volver=" & vuelta
						sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
						''    Celdax1 = ofv.BotonEliminar(ofv.convertircar(Celdan," ","_"),sHref)
					Else
						Celdax1 = ofv.GenCelda("","","","","","","","","si")
					End If
					If sHref <> "" Then
						Celdax1 = ofv.BotonEliminar("Eliminacion",sHref)
						Celdax1 = Replace(Celdax1,"<TD ","<TD title='Eliminacion' ")
					Else
						Celdax1 = ofv.GenCelda("","","","","","","","","si")
					End If

					Celda  = Celda & Celdax1

					sinput = "<a class=NT href=" & Chr(34) & "edProcedDET.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
					sinput = sinput & "&XEVEN=" & rs(0) & " "
					sinput = sinput &  Chr(34) & " title=" &  Chr(34) & "ver acciones de " & celdan &  Chr(34) &  " >" & Celdax & "</a>"

					Celda = Celda & ofv.GenCelda("","uti style='filter:shadow(color=#dddddd,direction=125);' ","","1%","","","",sicK,"si")



					Celda = Celda & ofv.GenCelda("","UT","","96%","","","",sinput,"si")

					' estado observacion

					Celdaobs  = ofv.GenCelda("","vt","","2","","","","&nbsp;","si")


					If (Nestado=5  Or Nestado=6) And (rs(14) & "") <> "" And bModificar Then

						stitle = "Hay observaciones para este grupo."
						sInput  = "<div style='border-style:solid;border-color:#993333;border-width:1;' >" & sImageobs & "</div>"
						Celdaobs  = ofv.GenCelda("","ut title='" & stitle & "'  valign=top ","","","1","","",sInput,"si")
					End If


					Celda = Celda & Celdaobs



					Formulario = Formulario & ofv.GenRow("","","","","","","",Celda,"si")

					rs.movenext
				Loop
				Call ofv.cerrarconsulta(rs)


				rs1.movenext
			Loop

			sHTML = sHTML & "<CENTER>"
			'sHTML = sHTML & ofv.GenTabla("","aaa","","90%","","0","","-1","2","","",Formulario,"si")
			if CONT = 0 then
			sHTML = sHTML & "<div class='Sininfo' align='center'><i class='las la-broadcast-tower' style='margin-right: 0.3rem;'></i>No tienes eventos agregados.</div>"
			       	else
			       	    sHTML = sHTML & ofv.GenTabla("","aaa","","90%","","0","","-1","2","","",Formulario,"si")
			       	    sHTML = sHTML & tabla
			end if

			sHTML = sHTML & "</CENTER>"


			'cambios
			Listar = ""


			If bVersion = "SI" Then
				If SESSION("SHWMOD") <> "N" Then

				ElseIf SESSION("VIEWMOD") <> "N" Then

					'xxxx    = CorGenBotonXp(3, "NO", "&#0113;","NO", "NO", "Lista de cambios", "NO", "Cambios")
					Listar  = genbotoncambios("") 'ofv.GenBoton(3, "cambBot", "BT", "N", " onclick='AbrirCambios();' ", "N", "S", "N", "", "")

					sasp = "CoGenCambios.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
					sHTML = sHTML & ConsCambios(sasp)

					Celda = Listar


				End If


			End If





			'--------------------------------------------------------------------------------
			sHTML = sHTML & "</span></div>"

			Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
			Fila = ofv.GenRow("","","","","","","",Celda,"si")
			sHTML = sHTML &  ofv.GenTabla("TCEX1"," onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

			'Estructura del Indice

			sHTML = sHTML  '& CorPnlIndx("N","")

			'--------------------------------------------------------------------------------
			sHTML = sHTML & ofv.botoneraESSX("0",rs1,CantRegAMover,CantRegAMostrar,"no","no",posi,"","","no","no","no",Listar)


			Call ofv.cerrarconsulta(rs1)

			Call ofv.cerrarconn(cn)

			sHTML = sHTML & "</body></html>"

		End If

	End If

End If

response.write sHTML


%>