<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	sestado = 1
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0

	XXXX = GENLOGEXT("EDIDET")

	ofv.ObtenerAtributos Session("Form2"),""
	bModificar = ofv.AModificar


	vuelta = ""
	aplcheck()
	nCodigo = request.querystring("Codobjeto")
	sTipo = request.querystring("Tipobjeto")
	upg=request.querystring("upg")

	WKOPCION = request.querystring("PASO")

	'''' session("PASOOBS") = WKOPCION
	''' RESPONSE.WRITE  session("PASOOBS")

	If NCODIGO <> "" Then
		session("NCliente") = NCODIGO
	Else
		NCODIGO = session("NCliente")
	End If

	If WKOPCION <> "" Then
		session("PASOOBS") = WKOPCION
	Else
		WKOPCION = session("PASOOBS")
	End If

	If sTipo <> "" Then
		session("TCliente") = sTipo
	Else
		sTipo = session("TCliente")
	End If

	vuelta = ofv.vueltaasp2("../") & vuelta
	vuelta2 = ofv.vueltaasp2("")
	'vuelta2 = ofv.vueltaasp2("../")
	sestado = 9
	'response.write vuelta
	Set cn = ofv.conectar(ofv.strconn2)
	Set cnW = ofv.conectar(ofv.strconn5)
	''response.write WKAPLIC & " WKAPLIC "

	If WKAPLIC = "OR" Or WKAPLIC = "PL" Then

		Strsql = "SELECT ESTADO FROM ORRSKCIRCUITO  "
		Strsql = Strsql & " Where OBJID = 0" & nCodigo
		Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		If Not rs.eof Then sestado = CLng(rs("estado"))
		Call ofv.cerrarconsulta(rs)

	Else


		Strsql = "SELECT a.objid,a.codobjeto,a.tipobjeto,a.origen,a.instrumen,a.instid,a.fecemis,a.fecvigen,a.fecbaja,a.estado,"
		Strsql = Strsql & " a.autor,a.aprobo,a.elimino,a.fecvto,a.version,a.nrover,"
		Strsql = Strsql & " a.fecacta,a.nroacta,a.prtylink,a.prtylink2,a.fecpublic,b.tipocpt FROM OBJprty  a "
		Strsql = Strsql & " inner join tipobjeto b on ( b.tipobjeto=a.tipobjeto)"
		Strsql = Strsql & " Where a.codobjeto=" & nCodigo
		Strsql = Strsql & " And a.tipobjeto = '" & stipo & "'"

		Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		tipocpt=""
		If Not rs.eof Then
			sestado = CLng(rs("estado"))
			tipocpt=rs("tipocpt")
		End If
		Call ofv.cerrarconsulta(rs)

	End If



	If CStr(sestado) > "2" And CStr(sestado) < "5" Then
		If session("Nomodificar") = "S" Then bModificar = False
		response.redirect "coobjwfldet.asp?cod=SI&codobjeto=" & nCodigo & "&tipobjeto=" & sTipo

	ElseIf (CStr(sestado) = "2" Or CStr(sestado) = "5") And Session("Form2") <> "DESARROLLO" And  Session("Form2") <>"CLIMS4681274004444386061790" And WKAPLIC <> "OR" And WKAPLIC <> "PL" Then
		response.redirect "edobjwflconf.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo

	Else
		If Not (session("segct") = "S") Then
			If session("Nomodificar") = "S" Then bModificar = False
		End If

		sHTML = ofv.FormHeader(session("FormN"))

		'----------------------------------------------

		shtml = shtml & CorPnlFnc("N","")

		sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=ddd>"

		'-------------------------------------------------

		Strsql = "SELECT * FROM OBJWFL  "
		Strsql = Strsql & " Where codobjeto=" & nCodigo
		Strsql = Strsql & " And tipobjeto = '" & stipo & "'"
		Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		If  rs.eof Then

			strsql = "Insert into OBJWFL (codobjeto, tipobjeto, ordenest, fecvto) "
			strsql = strsql & " Values (" & nCodigo & ",'" & sTipo & "','1','1')"
			Call ofv.crearconsultaEx(StrSql,cn,1,volver)

		Else
			''  response.write "v " & rs(6) & "<br>"
			If rs(6)  Then response.redirect "edObjWflHst.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo

		End If

		Call ofv.cerrarconsulta(rs)


		If WKOPCION = "OBS" Or WKOPCION = "APRO" Then
			WKOPCION = "APRO"
			WKOPCIONdsc = "OBSERVACION"
		Else
			WKOPCION = "AUTO"
			WKOPCIONdsc = "AUTORIZACION"
		End If


		session("CKTTYPE") = WKOPCION



		xActiva = 0
		xObliga = 0

		If WKAPLIC = "OR" Or WKAPLIC = "PL" Then

		Else

			Strsql = "SELECT * FROM Defwfl"
			Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

			If Not rs.eof Then

				xActiva = rs(1)
				xObliga = rs(2)

			End If

			Call ofv.cerrarconsulta(rs)

		End If

		If Not xActiva Then xObliga = 0

		If xObliga Then

			strsql = "update OBJWFL set esqauto = 1, ordenest = 1, chat = 1  "
			Strsql = Strsql & " Where codobjeto=" & nCodigo
			Strsql = Strsql & " And tipobjeto = '" & stipo & "'"

			Call ofv.crearconsultaEx(StrSql,cn,1,volver)

		End If

		'Estado
		Dim Aest(6)
		Aest(0) = 0
		Aest(1) = "Desarrollo"
		Aest(2) = "Para Aprobar"
		Aest(3) = "Aprobada"
		Aest(4) = "Eliminada"

		Strsql = "SELECT * from estados  "
		Set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		Do Until rsx.eof
			Aest(rsx(0)) = UCase(rsx(1))
			rsx.movenext
		Loop
		Call ofv.cerrarconsulta(rsx)



		Strsql = "SELECT * FROM OBJWFL  "
		Strsql = Strsql & " Where codobjeto=" & nCodigo
		Strsql = Strsql & " And tipobjeto = '" & stipo & "'"
		Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	' response.write strsql & " Strsql "


		If session ("base") ="K" Then
			Celda = ofv.gencelda("","ttt colspan=4 aaa ","center","100%","","","","<BIG>CIRCUITO ENVIO DE FORMULARIO  PARA " & WKOPCIONdsc & "</BIG>","si")
		Else

			Celda = ofv.gencelda("","ttt colspan=4 aaa ","center","100%","","","","<BIG>CIRCUITO DE " & WKOPCIONdsc & "</BIG>","si")
		End If
		Fila  = ofv.GenRow("","","","","","","",Celda,"si")

		Celda  = ofv.gencelda("","tt","","20%","","","","ESTADO&nbsp;ACTUAL:","si")
		Celda = Celda & ofv.gencelda("","vti colspan=3","","","","","",Aest(CLng(sestado)),"si")
		Fila   = FILA & ofv.GenRow("","","","","","","",Celda,"si")


		'IF PARA QUE LOS DIAS DE VENCIMIENTO SEAN SOLO PARA ORDEN ESTRICTO
		'if rs("ordenest") then

		Celda  = ofv.gencelda("","tt","","20%","","","","DIAS&nbsp;VENCIMIENTO:","si")
		'sImagenI = "<B><FONT FACE=Corflaticon COLOR=#F7AA22 SIZE=4>&#0166;</FONT></B>"
		'sinputE = "<a class=cc target=_new style='text-decoration:none;background-color:none;' Xonmouseover='this.style.textDecoration=" & chr(34)
		'sinputE = sinputE  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
		'sinputE = sinputE  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
		'sinputE = sinputE  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
		'sinputE = sinputE  & "none" & chr(34) & ";' "
		'sinputE = sinputE  & "  href='javascript:void(0);'  " & GTCSETCALR(sFormF,rs(8).name,GTCSMT,GTCFCH,GTCFCHDAT,"",10,2,0,0,0)
		'sinputE = sinputE  & "  TITLE='Modificar fecha de vencimiento.'  ><b>&nbsp;" & sImagenI & " " & "&nbsp;</b></a>"
		'if rs(8)=0 then
		'sInput = "---"
		'else
		'sInput = Fechador(rs(8))
		'end if
		'Celda = ofv.gencelda("","vti colspan=3","","","","","",sInputD,"si")
		'Fila   = FILA & ofv.GenRow("","","","","25","","",Celda,"si")
		'tablax= ofv.GenTabla("","","","80% ","","0","","0","1","","",fila,"si")
		sFormF = "sFormFx"
		sInputD = ofv.GenerarInput(rs(8).name,rs(8),"text",sFormF,"10","dtn","si")
		sAccionx = "asp/actualizaren.asp?tabla=objwfl&ID=" & rs(0) &"&volver=" & vuelta
		FormularioF =  ofv.genform(sFormF,"",sAccionx,"","post","",sInputD,"si")
		Celda  = Celda & ofv.gencelda("","uti","","20%","","","",FormulariOF,"si")

		Fila   = FILA & ofv.GenRow("","","","","","","",Celda,"si")

		'sql = "select CASE WHEN "
		'sql = sql & " isdate(CONVERT(nvarchar, substring(SEGCIRCUITO.COMIENZO, 1, charindex(' ', SEGCIRCUITO.COMIENZO)), 103)) = 1 "
		'sql = sql & " THEN dateadd(day," & rs(8) & ",CAST(substring(SEGCIRCUITO.COMIENZO, 1, charindex(' ', SEGCIRCUITO.COMIENZO)) AS date)) "
		'sql = sql & " WHEN isdate(CONVERT(nvarchar, substring(SEGCIRCUITO.COMIENZO, 1, charindex(' ', SEGCIRCUITO.COMIENZO)), 103)) = 0 "
		'sql = sql & " THEN dateadd(day," & rs(8) & ",CONVERT(date, substring(SEGCIRCUITO.COMIENZO, 1, charindex(' ', SEGCIRCUITO.COMIENZO)), 1)) "
		'sql = sql & " end AS fecha "
		'sql = sql & " from segcircuito where objcod='" & nCodigo & "' and objtipo='" & sTipo & "' "
		'response.write sql


		sql = "select substring(SEGCIRCUITO.COMIENZO, 1, charindex(' ', SEGCIRCUITO.COMIENZO))"
		sql = sql & " from segcircuito where objcod='" & nCodigo & "' and objtipo='" & sTipo & "' "

		Set rx = ofv.crearconsultaex(sql,cnW,1,parametros)

		If Not rx.eof Then

			comienzo = rx(0)
			If IsDate(comienzo) Then
				vFec = DateAdd("d",CDbl("0" & rs(8)),CDate(comienzo))
				sInput = FechadorCor(vFec,"")
			Else
				sInput="---"
			End If


		Else

			sInput = "CIRCUITO NO INICIADO"

		End If

		Call ofv.cerrarconsulta(rx)

		Celda  = ofv.gencelda("","tt","","20%","","","","PLAZO HASTA:","si")
		Celda  = Celda & ofv.gencelda("","uti colspan=2","","20%","","","","<b>" & sInput & "</b>","si")

		Fila   = FILA & ofv.GenRow("","","","","","","",Celda,"si")



		'end if

		If Not rs.eof Then

			sForm = "wfl"
			Celda = ofv.gencelda("","tt ","","20%","","","","AUTORIZANTES:","si")
			'  response.write  xobliga
			If bModificar And (CInt(sestado) = 1 Or CInt(sestado) = 6) And Not xobliga Then
				sInput = ofv.combosino(sForm & "1",rs(3),rs(3).name)

				sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objwfl&volver=" & vuelta
				sAccion = sAccion & "?upg=N&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
				sInput =  ofv.GenForm(sForm & "1","",sAccion,"","post","",sInput,"si")
				sInput = ofv.gencelda("","VTi colspan=2","","20%","","","",sInput,"si")
			Else
				sval = "NO"
				If rs(3) Then sval = "SI"
				sinput = ofv.gencelda("","VTi colspan=2","","20%","","","",sval,"si")
			End If

			Celda = Celda & sinput



			If rs(3) And (CInt(sestado) = 1 Or CInt(sestado) = 6)  Then
				sInput = ofv.GenerarInput("boton1","Lista","submit",sForm & "1a","5","bt","")

				sAccion = "edobjwfldet.asp?volver=" & vuelta2
				sAccion = sAccion & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
				sInput =  ofv.GenForm(sForm & "1a","",sAccion,"","post","",sInput,"si")
				sInput = ofv.gencelda("","ut","center","20%","","","",sInput,"si")
			Else

				sInput = ofv.gencelda("","ut","","20%","","","","&nbsp;","si")


			End If

			Celda = Celda & sinput
			Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

			''''
			'''' solo si esta activa la accion de lista de autorizantes '''''''''''''''''''''
			''''


			If rs(3) Then

				sval    = "NO"
				svaldsc = "INDISTINTO"
				If rs(4) Then
					sval = "SI"
					svaldsc = "ESTRICTO"
				End If

				Celda = ofv.gencelda("","tt ","","20%","","","","Orden:","si")

				If bModificar And (CInt(sestado) = 1 Or CInt(sestado) = 6) And Not xobliga Then
					sInput = ofv.combosino(sForm & "1b",rs(4),rs(4).name)
					sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objwfl&volver=" & vuelta
					sAccion = sAccion & "?upg=S&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
					sInput =  ofv.GenForm(sForm & "1b","",sAccion,"","post","",sInput,"si")
					sInput = ofv.gencelda("","VTi colspan=2","","10%","","","",sInput,"si")

				Else

					sinput = ofv.gencelda("","VTi colspan=2","","10%","","","",sval,"si")
				End If

				Celda = Celda & sinput

				sInput = ofv.gencelda("","TTT","","10%","","","",svaldsc,"si")
				Celda = Celda & sinput
				Fila = fila & ofv.GenRow("","","","","25","","",Celda,"si")


				''''ordenest = 1,


				If bModificar  And (CInt(sestado) = 1 Or CInt(sestado) = 6)   Then
					strsql = "update OBJWFL set esqauto = 1,  chat = 1  "
					Strsql = Strsql & " Where codobjeto=" & nCodigo
					Strsql = Strsql & " And tipobjeto = '" & stipo & "'"

					Call ofv.crearconsultaEx(StrSql,cn,1,volver)
				End If





				'Texto del Mensaje
				esttxt = " STYLE={font-family:tahoma;color:#225588;font-size:7pt;background-color:transparent;} "

				XMESAURL   =  Left(request.servervariables("SCRIPT_NAME"),InStr(Right(request.servervariables("SCRIPT_NAME"),Len(request.servervariables("SCRIPT_NAME"))-1),"/"))

				XMESAURL   = TipoSite("")  & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & XMESAURL

				Celda = ofv.gencelda("","ttr colspan=4","LEFT","","","","",XMESAURL ,"si")
				'  Fila = fila & ofv.GenRow("","","","","25","","",Celda,"si")



				Celda = ofv.gencelda("","ttr colspan=4","LEFT","","","","","MENSAJE:","si")
				Fila = fila & ofv.GenRow("","","","","25","","",Celda,"si")

				If bModificar  And (CInt(sestado) = 1 Or CInt(sestado) = 6) Then

					wkfmensaje = rs(7) & ""

					If upg  <> "S" Then

						If WKAPLIC = "OR" Or WKAPLIC = "PL" Then

							wkfmensaje = "Se encuentra Disponible en la Aplicación " & Session("CorTitApli") & " un CIRCUITO"
							wkfmensaje = wkfmensaje & " pendiente de su aprobación."
							wkfmensaje = wkfmensaje & Chr(13) & Chr(13) & " <APLURL>INGRESAR A LA APLICACION</APLURL>"
							wkfmensaje = UCase(wkfmensaje)
						ElseIf session("base")= "K" Then
							'' response.write session("base") & "3"
							' wkfmensaje = "Se encuentra Disponible en la Aplicación " & Session("CorTitApli") & " un formulario"
							' wkfmensaje = wkfmensaje & " pendiente de su aprobación."
							wkfmensaje = "Se encuentra disponible en la Aplicación  " & Session("CorTitApli") & " un Formulario"
							wkfmensaje = wkfmensaje & " Pendiente de su aprobación."
							wkfmensaje = wkfmensaje & " Agradeceremos nos hagan llegar vuestros comentarios y/o aprobación en un plazo de días"
							wkfmensaje = wkfmensaje & Chr(13) & Chr(13) & " <APLURL>INGRESAR A LA APLICACION</APLURL>"


						Else

							'wkfmensaje = "Se encuentra Disponible en la Aplicación " & Session("CorTitApli") & " un proyecto"
							'wkfmensaje = wkfmensaje & " normativo pendiente de su aprobación."
							'wkfmensaje = wkfmensaje & CHR(13) & CHR(13) & " <APLURL>INGRESAR A LA APLICACION</APLURL>"
							'wkfmensaje = UCASE(wkfmensaje)
							If WKOPCION = "AUTO" Then
								If session("APLMDCLI")="BPN" Then
									sUsuario = getinfouser(session("usuario"))
									wkfmensaje = " SE ENCUENTRA DISPONIBLE EN LA APLICACIÓN EMANAGER  UN PROYECTO NORMATIVO PENDIENTE DE SU <b><u>APROBACIÓN</u></b>. "
									wkfmensaje = wkfmensaje & " CONFORME LO ESTABLECIDO EN LA NORMATIVA INTERNA SOLICITAMOS DAR CONFORMIDAD DENTRO DE LOS <b><u>DOS (2) DÍAS</u></b> "
									wkfmensaje = wkfmensaje & " HÁBILES SIGUIENTES. TRANSCURRIDO DICHO PLAZO SIN QUE SE HAYA REGISTRADO VUESTRA INTERVENCIÓN, SE DARÁ  "
									wkfmensaje = wkfmensaje & " POR CONFORMADO TÁCITAMENTE Y SE PROCEDERÁ A LA APROBACIÓN DEL DOCUMENTO POR LA INSTANCIA CORRESPONDIENTE "
									wkfmensaje = wkfmensaje & " CONFORME LAS FACULTADES DELEGADAS POR EL DIRECTORIO DEL BANCO. <b><u>" & session("SGU_nombre") & "</u></b>"
									wkfmensaje = wkfmensaje & Chr(13) & Chr(13) & " <APLURL>INGRESAR A LA APLICACION</APLURL>"

									' wkfmensaje = UCASE(wkfmensaje)
								ElseIf session("APLMDCLI")="BGAL" Then

									wkfmensaje = "Se encuentra disponible en la Aplicación  " & Session("CorTitApli") & " un proyecto"
									wkfmensaje = wkfmensaje & " normativo pendiente de su aprobación."
									wkfmensaje = wkfmensaje & " Agradeceremos nos hagan llegar vuestros comentarios y/o aprobación en un plazo de días"
									wkfmensaje = wkfmensaje & Chr(13) & Chr(13) & " <APLURL>INGRESAR A LA APLICACION</APLURL>"


								Else
									wkfmensaje = "Se encuentra Disponible en la Aplicación " & Session("CorTitApli") & " un proyecto"
									wkfmensaje = wkfmensaje & " normativo pendiente de su aprobación."
									wkfmensaje = wkfmensaje & Chr(13) & Chr(13) & " <APLURL>INGRESAR A LA APLICACION</APLURL>"
									wkfmensaje = UCase(wkfmensaje)
								End If


							Else

								If session("APLMDCLI")="BPN" Then
									sUsuario = getinfouser(session("usuario"))
									wkfmensaje = "Se encuentra Disponible en la Aplicación " & Session("CorTitApli") & " un proyecto"
									wkfmensaje = wkfmensaje & " normativo pendiente de su observación. De no recibir respuesta en"
									wkfmensaje = wkfmensaje & " el plazo de 5 dias hábiles"
									wkfmensaje = wkfmensaje & " se entenderá que no existen observaciones por su parte, procediendo a"
									wkfmensaje = wkfmensaje & " cancelar el circuito de observación y a publicar el esquema normativo. "
									wkfmensaje = wkfmensaje & " Quedamos a disposición. "
									wkfmensaje = wkfmensaje &  session("SGU_nombre")
									wkfmensaje = wkfmensaje & Chr(13) & Chr(13) & " <APLURL>INGRESAR A LA APLICACION</APLURL>"
									wkfmensaje = UCase(wkfmensaje)
								Else
									wkfmensaje = "Se encuentra Disponible en la Aplicación " & Session("CorTitApli") & " un proyecto"
									wkfmensaje = wkfmensaje & " normativo pendiente de su aprobación."
									wkfmensaje = wkfmensaje & Chr(13) & Chr(13) & " <APLURL>INGRESAR A LA APLICACION</APLURL>"
									wkfmensaje = UCase(wkfmensaje)
								End If
							End If

						End If


						strsql = "update OBJWFL set MSGCHAT = '" & wkfmensaje & "'  "
						Strsql = Strsql & " Where codobjeto=" & nCodigo
						Strsql = Strsql & " And tipobjeto = '" & stipo & "'"

						Call ofv.crearconsultaEx(StrSql,cn,1,volver)
					End If

					nCol = 80

					t3b = "<font class=at><textarea CLASS=DT "
					t3b = t3b & "name=" & rs(7).name & " cols=" & nCol & " rows=6 "
					If bModificar Then
						t3b = t3b & " onblur=" & sForm & "3s" & "3.submit() "
						t3b = t3b & " onkeypress=" & Chr(34)
						t3b = t3b & " if (this.value.length > 2450) {alert('El texto es demasiado"
						t3b = t3b & " extenso'); return false;}" & Chr(34) & " >"
					Else
						t3b = t3b & " readonly>"
					End If
					t3b = t3b & wkfmensaje & "</textarea></font>"

					sAccion = "asp/ActualizarEN.asp?ID=" & rs(0) & "&Tabla=Objwfl&volver=" & vuelta
					sAccion = sAccion & "?upg=S&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
					sInput =  ofv.GenForm(sForm & "3s3","",sAccion,"","post","",t3b,"si")
					sInput = ofv.gencelda("","ut colspan=4","center","","","","",sInput,"si")
				Else
					sinput = ofv.gencelda("","ut colspan=4","center","","","","",rs(7),"si")
				End If
				'celda= ofv.gencelda("","ut colspan=4","center","","","","",t3b,"si")
				'Celda = sinput
				Fila = fila & ofv.GenRow("","","","","","","",sInput,"si")


			End If


			WKFSIG = "S"


			If rs(3) Then
				WKFSIG2 = "S"
			Else
				WKFSIG2 = "N"
			End If

			If rs(4) Then
				WKFMODAL = "S"
			Else
				WKFMODAL = "N"
			End If

		End If


		Call ofv.cerrarconsulta(rs)

		sHTML = sHTML & "<br><center>"
		sHTML = sHTML & ofv.GenTabla("","aaa  ","","70% style='border-style:solid;border-width:1;border-color:#999999;' ","","0","0","-2","1","","",fila,"si")
		sHTML = sHTML & "</center>"

		FILA = ""

		Celda = ofv.gencelda("","ttt colspan=4","","","","","","&nbsp;estado del circuito","si")
		Formulario2 = ofv.GenRow("","","","","30","","",Celda,"si")

		If WKFSIG = "S" Then

			If (CInt(sestado) = 1 Or CInt(sestado) = 6) And (bModificar   Or session("Nomodificarce") <> "S") And Session("Segce") = "S"  Then

				Strsql = "SELECT * FROM OBJwfldet  "
				Strsql = Strsql & " Where codobjeto=" & nCodigo
				Strsql = Strsql & " And tipobjeto = '" & stipo & "' And CKTTYPE = '" & session("CKTTYPE") & "' order by orden,objid"
				Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
				If RS.EOF Then
					WKFSIGOK = "N"
				Else
					WKFSIGOK = "S"
				End If


				If WKAPLIC = "OR" Or WKAPLIC = "PL" Then

					Strsql = "SELECT * FROM ORRSKCONFIG WHERE APLIC = '" & WKAPLIC & "'  "
					Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
					If RS.EOF Then
						WKFSIGOK = "N2"
					Else
						If WKOPCION = "AUTO" Then
							WKFCKT = RS(5)
						Else
							WKFCKT = RS(6)
						End If
					End If

				Else

					Strsql = "SELECT * FROM TIPOBJETO  "
					Strsql = Strsql & " Where "
					Strsql = Strsql & " tipobjeto = '" & stipo & "' "
					Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
					If RS.EOF Then
						WKFSIGOK = "N2"
					Else
						If WKOPCION = "AUTO" Then
							WKFCKT = RS(19)
						Else
							WKFCKT = RS(20)
						End If
					End If

				End If



				If WKFSIGOK = "S" Or (WKFSIGOK = "N" And WKFSIG2 = "N" And WKOPCION = "AUTO" ) Then

					Celda = ""
					If WKFSIGOK = "N" Then
						sInput2 = "<b>ATENCION!!</b><BR>EN CASO DE INICIAR, EL CIRCUITO CERRARA CON "
						SINPUT2 = SINPUT2 & " <u><b>" & SESSION("SEGUSUARIO") & "</b></u>  COMO UNICO AUTORIZANTE.<BR>"
					Else
						sInput2 = "<b>ATENCION!!</b><BR>EL CIRCUITO ESTA EN CONDICIONES DE SER INICIADO POR "
						SINPUT2 = SINPUT2 & " <u><b>" & SESSION("SEGUSUARIO") & "</b></u>  COMO ORIGINANTE.<BR>"
					End If
					Celda = ofv.gencelda("","uti colspan=3 style='text-align:justify;' ","","","","","",SINPUT2,"si")

					'' Agregado 20042015
					'''sInput = ofv.GenerarInput("boton1","INICIAR","submit","edPropiedades2","5","bt","")
					sInput = ofv.GenerarInput("boton1","INICIAR","submit","edPropiedades2","5","bt onclick=""this.style.cursor='wait';this.style.visibility='hidden';""","")
					'' Agregado 20042015 FIN

					sAccion = "EDWKFSTART.asp?OPCION=" & WKOPCION & "&MODAL=" & WKFMODAL &  "&Codobjeto=" & ncodigo & "&Tipobjeto=" & stipo
					sInput = ofv.GenForm("edPropiedades2","",sAccion,"","post","",sInput,"si")

					Celda = celda & ofv.gencelda("","uti","center","10%","","","",sInput,"si")

					Formulario2 = Formulario2 &  ofv.GenRow("","","","","","","",Celda,"si")



				Else

					Celda = ""
					If WKFSIGOK = "N" Then
						sInput = "<b>ATENCION!!</b><BR>NO PUEDE INICIAR EL CIRCUITO, NO INFORMO LA LISTA DE AUTORIZANTES.<BR>"
					Else
						sInput = "<b>ATENCION!!</b><BR>NO PUEDE INICIAR EL CIRCUITO, EL TIPO DE DOCUMENTO NO ESTA CONFIGURADO.<BR>"
					End If
					Celda = ofv.gencelda("","uti colspan=3 style='text-align:justify;' ","","","","","",SINPUT,"si")

					sInput = "&nbsp;"
					Celda = celda & ofv.gencelda("","ut","center","5%","","","",sInput,"si")

					Fila = ofv.GenRow("","","","","","","",Celda,"si")

					Formulario2 = Formulario2 &  Fila

				End If

			Else

				Celda = ""

				If (CInt(sestado) = 1 Or CInt(sestado) = 6)  Then
					sInput = "<b>ATENCION!!</b><BR>NO PUEDE INICIAR EL CIRCUITO DESDE AQUI O NO ESTA AUTORIZADO PARA INICIARLO.<BR>"
				Else
					sInput = "<FONT COLOR=#005500>EN CURSO.</FONT><BR>"
				End If
				Celda = ofv.gencelda("","uti colspan=3 style='text-align:justify;' ","","","","","",SINPUT,"si")

				sInput = "&nbsp;"
				Celda = celda & ofv.gencelda("","ut","center","5%","","","",sInput,"si")

				Fila = ofv.GenRow("","","","","","","",Celda,"si")

				Formulario2 = Formulario2 &  Fila


			End If

			StrSql = "Select * From SEGCIRCUITO WHERE OBJTIPO = '" & session("TCliente") & "' and  OBJCOD = '" & session("NCliente") & "' ORDER BY COMIENZO DESC "
			Set rscc = ofv.crearconsultaEx(StrSql,cnW,1,parametros)
			If Not rscc.eof Then
				xcircuito="S"
			Else
				xcircuito="N"
			End If
			Call ofv.cerrarconsulta(rscc)
			Call ofv.cerrarconn(cn)
			If  CInt(sestado) = 1 And   xcircuito ="S" And Session("Segce") = "S"     Then
				sInput = ofv.GenerarInput("boton1","Consultar historial","submit","edPropiedades3","5","bt","")
				sAccion = "CoHistWkf.asp"
				sInput = ofv.GenForm("edPropiedades3","",sAccion,"","post","",sInput,"si")
				Celda = ofv.gencelda("","Vt colspan=2","center","30%","","","",sInput,"si")
				' Celda = celda & ofv.gencelda("","Vt colspan=2 ","center","30%","","","",sInput,"si")
				Fila = ofv.GenRow("","","","","","","",Celda,"si")
				Formulario2 = Formulario2 & Fila
			ElseIf ( CInt(sestado) = 2 Or CInt(sestado) = 5)  And  Session("Segce") = "S"  Then

				sInput = ofv.GenerarInput("boton1","Consultar historial","submit","edPropiedades3","5","bt","")
				sAccion = "CoHistWkf.asp"
				sInput = ofv.GenForm("edPropiedades3","",sAccion,"","post","",sInput,"si")
				Celda = ofv.gencelda("","Vt colspan=2","center","30%","","","",sInput,"si")

				If (bModificar   Or session("Nomodificarce") <> "S") And Session("Segce") = "S" Then
					sInput = ofv.GenerarInput("boton1","Cancelar Circuito","submit","edPropiedades3","5","bt style='background-color:#993333;' ","")
					sAccion = "EDCANCELWkf.asp?OPCION=CONF"
					sInput = ofv.GenForm("edPropiedades3","",sAccion,"","post","",sInput,"si")
				Else
					sInput = "&nbsp;"
				End If

				Celda = celda & ofv.gencelda("","Vt colspan=2 ","center","30%","","","",sInput,"si")

				Fila = ofv.GenRow("","","","","","","",Celda,"si")


				Formulario2 = Formulario2 & Fila


			End If


		Else

			If WKAPLIC = "OR" Or WKAPLIC = "PL" Then

			Else


				Strsql = "SELECT OBJPRTY.objid,OBJPRTY.codobjeto,OBJPRTY.tipobjeto,OBJPRTY.origen,OBJPRTY.instrumen,OBJPRTY.instid,OBJPRTY.fecemis,"
				Strsql = Strsql & " OBJPRTY.fecvigen,OBJPRTY.fecbaja,OBJPRTY.estado,OBJPRTY.autor,OBJPRTY.aprobo,OBJPRTY.elimino,OBJPRTY.fecvto,OBJPRTY.version,"
				Strsql = Strsql & " OBJPRTY.nrover,OBJPRTY.fecacta,OBJPRTY.nroacta,OBJPRTY.prtylink,OBJPRTY.prtylink2,OBJPRTY.fecpublic,Estados.estado FROM OBJPRTY, Estados "
				Strsql = Strsql & "WHERE OBJPRTY.estado = Estados.objid AND codobjeto=" & nCodigo
				Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
				If Not rs.eof Then
					If session("segct") = "S" Then
						Celda = ofv.gencelda("","Vt colspan=2","center","30%","","","","CAMBIO DE ESTADO","si")
						Celda = celda & ofv.gencelda("","Vt ","","30%","","","","NUEVO ESTADO:","si")
						'xxxxxxxxxxxxx


						Select Case CStr(rs(9))
							Case "1"
							If (bModificar   Or session("Nomodificarce") <> "S") And Session("Segce") = "S"  Then
								sInput = "<font class=at><select name=" & rs(9).name
								sInput = sInput & " onchange=document.edPropiedades2.submit(); >"
								sInput = sInput & "<option selected value=1>" & Aest(1) & "</option>"
								sInput = sInput & "<option value=2>" & Aest(2) & "</option></select></font>"
							Else
								sInput = ofv.GenerarInput(rs(9).name,rs(16),"readonly","","20","dt","")
							End If
							Case "6"
							If (bModificar   Or session("Nomodificarce") <> "S") And Session("Segce") = "S"  Then
								sInput = "<font class=at><select name=" & rs(9).name
								sInput = sInput & " onchange=document.edPropiedades2.submit(); >"
								sInput = sInput & "<option selected value=1>" & Aest(1) & "</option>"
								sInput = sInput & "<option value=2>" & Aest(2) & "</option></select></font>"
							Else
								sInput = ofv.GenerarInput(rs(9).name,rs(16),"readonly","","20","dt","")
							End If
							Case "2"
							If (bModificar   Or session("Nomodificarce") <> "S") And Session("Segce") = "S" Then
								sInput = "<font class=at><select name=" & rs(9).name
								sInput = sInput & " onchange=document.edPropiedades2.submit(); >"
								sInput = sInput & "<option selected value=2>" & Aest(2) & "</option>"
								sInput = sInput & "<option value=1>" & Aest(1) & "</option>"
								sInput = sInput & "</select></font>"
							Else
								sInput = ofv.GenerarInput(rs(9).name,rs(16),"readonly","","20","dt","")
							End If
						End Select
						Celda = Celda & ofv.GenCelda("","VT","","","","","",sInput,"si")
						Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

						sAccion = ""
						If (bModificar   Or session("Nomodificarce") <> "S") And Session("Segce") = "S" Then
							sAccion = "asp/ActualizarEN.asp?opc=WFL&ID=" & rs(0) & "&Tabla=Objprty&volver=" & vuelta
							sAccion = sAccion & "?upg=S&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
						End If
						Formulario2 = ofv.GenForm("edPropiedades2","",sAccion,"","post","",Fila,"si")

					End If
				End If

			End If

			'xxxxxxxxxx

		End If

		If WKAPLIC = "OR" Or WKAPLIC = "PL" Then

			sAccion = Session("ASPSTR") & "&ACCIONEJEC=ABREESTAD"
			Retornar = ofv.BotonVolver(sAccion)


		End If


		sHTML = sHTML & "<br><center>"
		sHTML = sHTML & ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#999999;' ","","70% ","","0","0","1","0","","",Formulario2,"si")
		sHTML = sHTML & "</center>"



		'--------------------------------------------------------------------------------
		sHTML = sHTML & "</span></div>"

		Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
		Fila = ofv.GenRow("","","","","","","",Celda,"si")
		sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

		'Estructura del Indice

		sHTML = sHTML & CorPnlIndx("N","")

		'--------------------------------------------------------------------------------

		sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Retornar,"")

		Call ofv.cerrarconsulta(rs)
		Call ofv.cerrarconn(cn)
		sestado = 1
		sHTML = sHTML & "</body></html>"    ''& SESSION("WKF")
	End If
End If

If (CInt(sestado) = 1 Or CInt(sestado) = 6)  Then response.write sHTML



Function GETDH()

	XKD = DatePart("d", Date)
	XKM = DatePart("m", Date)
	XKY = DatePart("yyyy", Date)





	GETDH = XKD & "/" & XKM & "/" & XKY & " " & Time()


End Function

Function Fechador(wkfecha)

	'response.write wkfecha

	fec = CDate(wkfecha + CLng(CDate("01-01-1900")))
	ndia = ofv.cdia(Weekday(fec))
	nMes = ofv.cmes(Month(fec))

	Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)


End Function
%>