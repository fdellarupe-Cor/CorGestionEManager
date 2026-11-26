<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0
	ncaller = request.querystring("caller")
	nCodigo = request.querystring("Codobjeto")
	sTipo = request.querystring("Tipobjeto")
	sCod = request.querystring("Cod")
	INDICE=request.querystring("Indice")
	Volver = request.querystring("Volver")
	SACC=request.querystring("ACC")
	segckt=request.querystring("segckt")
	estricto=request.querystring("tipowfl")
	Volver = Volver & "?Cod=" & sCod &"_ACC=" & SACC & "_segckt=" & segckt & "_acp=S_estricto=" & estricto
	sxopc1 = ""

	If SACC="PEND" Then INDICE="NO"

	'''''''''''''''''''''''''''''''''''''''''''''''''''

	sHTML = ofv.FormHeader("Agregar Usuario")
	If Indice <>"NO" Then
		'----------------------------------------------

		shtml = shtml & CorPnlFnc("N","")

		sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"
	End If
	'-------------------------------------------------
	SHTML = SHTML & GenBuscaUsr("N","N","N","S",Chr(9),Chr(9),Chr(9),Chr(9))

	If SESSION("FBUSR") <> "NO" Then

		Set cn = ofv.conectar(ofv.strconn2)

		If SACC = "PEND" Then

			Set cn3 = ofv.conectar(ofv.strconn5)

			orden = 0

			puede = false
            if session("SEGCKT")="" then
             paso=segckt
            else
             paso=session("SEGCKT")
            end if

            puede=true

			sql = "select codigo from segcktresp where segckt='" & paso & "'"
			'response.write sql  & "       ----"& segckt & "  "  & request.form

			Set rz = ofv.crearconsultaex(sql,cn3,1,parametros)

			do while not rz.eof
			  if SESSION("FBUSR") = rz(0) then puede = false
			 rz.movenext
			loop


			If puede Then

				usr = SESSION("FBUSR")

				strsql = "select MAX(orden),segcktpsc from segcktresp where segckt='" & segckt & "' and modo='L' group by segcktpsc"
			'	Set rs = ofv.crearconsultaex(strsql,cn3,1,parametros)

				orden = rs(0) + 10
                segcktpsc=rs(1)
				email=""
				'sql = "select email from usuarios where ID='" & SESSION("FBUSR") & "'"
				'set em = ofv.crearconsultaex(sql,cn,1,parametros)
				'if not em.eof then
				'	email = em(0)
				'end if
				xx=getinfouser( SESSION("FBUSR"))
				email=session("sgu_mail")
				sql = "insert into segcktresp(segckt,segcktpsc,modo,tipo,codigo,orden,nivel,subcodigo, segemail) values "
				sql = sql & " ('" & segckt & "','" & rs(1) & "','L','USR','" & SESSION("FBUSR") & "','" & orden & "','5','39','" & email & "')"
			'	Call ofv.crearconsultaex(sql,cn3,1,parametros)


				If Not CBool(estricto) Then
					SEGCKTPSCSIGX = GTCKTPSC() & "IND" & orden

					strsql = "select ID from segcktresp where segckt='" & segckt & "' and codigo='" & SESSION("FBUSR") & "'"
					'response.write strsql
					Set rx = ofv.crearconsultaex(strsql,cn3,1,parametros)
                    segidresp= rx(0)
					sql = "insert into segpaso (segckt,segcktpsc,model,comienzo,fin,okrch,resptipo,responsable,respasg,pscdsc,observaciones,nivel,subcodigo,segidresp,segsalida) values "
					sql = sql & " ('" & segckt & "','" & SEGCKTPSCSIGX & "','" & rs(1) & "','" & GETDH() & "','NO','NO','USR','" & SESSION("FBUSR") & "','" & SESSION("FBUSR") & "','20: AUTORIZACION DE RESPONSABLES','','5','39','" & rx(0) & "','')"
				'	Call ofv.crearconsultaex(sql,cn3,1,parametros)

				End If

				mstdsc = "Se encuentra Disponible en la Aplicación " & Session("CorTitApli") & " un CIRCUITO"
				mstdsc = mstdsc & " pendiente de su aprobación."
				'XXXX = GENCKTMAIL("USR",usr,mstdsc & "<br>" & "<br>" & " POR @@@NORMREQ@@@",SESSION("ORGMAIL"),"3")

				Call ofv.cerrarconsulta(rs)
				Call ofv.cerrarconsulta(rx)
			End If
		End If

		Set cn = ofv.conectar(ofv.strconn2)

		strsql = "Select * From objwfldet "
		'response.write strsql
		Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

		f1a = "asp/AgregarEN.asp?Tabla=OBJwflDET"
		f1a = f1a & "&volver=" & volver & "&Codobjeto=" & nCodigo
		f1a = f1a & "&Tipobjeto=" & sTipo

		sHTML = sHTML & ofv.GenForm("AgreArchb","",f1a,"","post","","","no")

		If session("CKTTYPE")="" Then session("CKTTYPE")="AUTO"


		sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & SESSION("FBUSR") & "' NAME='" & RS(4).NAME & "' >"
		sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='0' NAME='" & RS(3).NAME & "' >"
		sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & session("CKTTYPE") & "' NAME='CKTTYPE' >"
   '     If session("EMPMUN") = "S" Then
		  sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & SESSION("UN") & "' NAME='codemp' >"
	'	end if


		'session("CKTTYPE")=""
		Call ofv.cerrarconsulta(rs)
		Call ofv.cerrarconn(cn)


	Else

		If ncaller = "AUT" Then
			sAccion = "AutDefaultdet.asp"
		Else
			If SACC="PEND" Then
				sAccion = "edobjwfldetPEND.asp?segckt=" & segckt & "&ACC=PEND&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Cod=" & sCod & "&volver=edobjwfl.asp"
			Else
				sAccion = "edobjwfldet.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Cod=" & sCod & "&volver=edobjwfl.asp"
			End If
		End If
		xHTML="<INPUT TYPE='HIDDEN' VALUE='" & usr & "' NAME='usr' >"


		sHTML = sHTML & ofv.GenForm("AgreArchb","",sAccion,"","post","",xhtml,"no")


	End If


	'Botones Aceptar y Cancelar

	If ncaller = "AUT" Then
		sAccion = "AutDefaultdet.asp"
	Else
		If SACC="PEND" Then
			sAccion = "edobjwfldetPEND.asp?segckt=" & segckt & "&ACC=PEND&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Cod=" & sCod & "&volver=edobjwfl.asp"
		Else
			sAccion = "edobjwfldet.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Cod=" & sCod & "&volver=edobjwfl.asp"
		End If
	End If
	zbtn =  ofv.AgregarBoton("AgreArchb",sAccion,"_self")
	zbtn=replace(zbtn,"document.","document.getElementById(""agreg"").style.display = ""none"";document.")
	shtml = shtml & zbtn


	Agregar = ""
	Buscar  = ""
	Celda   = ""


	sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)
	If Indice <>"NO" Then

		'--------------------------------------------------------------------------------
		sHTML = sHTML & "</span></div>"

		Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
		Fila = ofv.GenRow("","","","","","","",Celda,"si")
		sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

		'Estructura del Indice

		sHTML = sHTML & CorPnlIndx("N","")

		'--------------------------------------------------------------------------------
	End If
	sHTML = sHTML & "</body></html>"
End If

response.write sHTML

Function GETDH()

	XKD = DatePart("d",Date())
	XKM = DatePart("m",Date())
	XKY = DatePart("yyyy",Date())

	GETDH = XKD & "/" & XKM & "/" & XKY & " " & Time()


End Function

Function GTCKTPSC()

	sh = DatePart("h",Date())
	sm = DatePart("n",Date())
	ss = DatePart("s",Date())

	XCodobJ = sh & sm & ss

	GThms = DateDiff("d",Date(),CDate("12/12/2050")) & FormatDateTime(Now(),3)
	GTses = session.sessionid
	GTsecu = CStr(XCodobJ & GThms & GTses)

	GTsecuf = "CKTPSC"

	For GTi=1 To Len(GTsecu)
		GTcs = Mid(GTsecu,GTi,1)
		If GTcs => "0" And GTcs <= "9" Then
			GTsecuf = GTsecuf & GTcs
		End If
	Next

	GTCKTPSC = Trim(GTsecuf)

End Function

Function GETRESPOCKT(RESPCOD,RESPTIP)

	GETRESPOCKT = "SIN ASIGNAR"



	XRESPO = RESPCOD & ""

	If XRESPO <> "" And XRESPO <> "NO" Then

		Set cn2 = ofv.conectar(ofv.strconn0)

		Select Case RESPTIP

			Case "USR"
			' strsql = "Select * From USUARIOS where ID = '" & XRESPO & "' "
			' set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
			' if not rsniv.eof then sInputx = rsniv(1)
			' call ofv.cerrarconsulta(rsniv)

			xxxx = getInfoUser(XRESPO)
			If session("SGU_nombre") <> "" Then    SINPUTX= session("SGU_nombre")


			Case "GRP"
			strsql = "Select * From GRUPOS where ID = '" & XRESPO & "' "
			Set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
			If Not rsniv.eof Then sInputx = rsniv(1)
			Call ofv.cerrarconsulta(rsniv)

			Case "SEC"
			XRESPO = 0 & XRESPO
			strsql = "Select * From ECO where ECO = " & XRESPO
			'''     response.write strsql
			Set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
			If Not rsniv.eof Then sInputx = rsniv(1)
			Call ofv.cerrarconsulta(rsniv)

		End Select

		GETRESPOCKT = sInputx

	End If


End Function

Function GENCKTMAIL(WETIPO,WECOD,WEMESSAGE,WEORIGEN,WELVL)

	XRESPO = WECOD & ""

	If XRESPO <> "" And XRESPO <> "NO" Then

		Set cn2 = ofv.conectar(ofv.strconn0)

		XPARA = GETRESPOCKT(WECOD,WETIPO)
		STRSQL = "SELECT EMAIL,DESCRIPCION FROM USUARIOS WHERE "

		XGUMENSAGEw = WEMESSAGE


		XMESAURL  = ""

		If  SESSION("URLTYP")  = "V" Then

			XMESAURL  =  Left(request.servervariables("SCRIPT_NAME"),InStr(Right(request.servervariables("SCRIPT_NAME"),Len(request.servervariables("SCRIPT_NAME"))-1),"/"))

		End If

		If SESSION("URLPORT")  = "N" Then
			XMESAURL   = "http://" & request.servervariables("SERVER_NAME")  & XMESAURL
		Else
			XMESAURL   = "http://" & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & XMESAURL
		End If

		XGUMENSAGEw = Replace(XGUMENSAGEw,"<APLURL>","<U><A HREF='" & XMESAURL  & "' TARGET='_blank' >")
		XGUMENSAGEw = Replace(XGUMENSAGEw,"</APLURL>","</A></U>")

		WEMESSAGE = XGUMENSAGEw

		Select Case WETIPO

			Case "USR"
			' STRSQL = STRSQL & " ID = '" & XRESPO & "' "


			xxxx = getInfoUser(XRESPO)
			If session("SGU_mail")<> "" Then    SINPUTX= session("SGU_mail")


			stowho = SINPUTX & ""
			If Len(stowho) > 0 Then
				stowho = Trim(stowho)

				subject = "SOLICITUD DE AUTORIZACION"
				If WETIPO = "USR" Then
					XPARA2 = XPARA
				Else
					XPARA2 = XPARA & " / " & session("SGU_nombre")
				End If

				Xmessage = GenMDLMAIL (XPARA2,WEORIGEN,WEMESSAGE)
				''  message = "<u><font size=1 color=#000080 >" & XPARA2 & "</font></u><BR><BR><b><font size=1  >" & WEMESSAGE & "</font><b>"
				Call SendMail(Xmessage, stowho, WEORIGEN,subject, session("MailComponent"),"SI","1", "NO")
			End If

			Case "GRP"

			' STRSQL = STRSQL & " GRUPO = '" & XRESPO & "' and ( NOT nivel >  " & WELVL & " ) "

			Call  getUsrIngroup(XRESPO)
			A_SGU_UIGuserid =Split(session("SGU_UIGuserid"),Chr(9))
			A_SGU_UIGdescri =Split(session("SGU_UIGdescri") ,Chr(9))

			'   set RSXPA = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
			' DO UNTIL RSXPA.EOF
			For a_i=0 To UBound(SGU_UIGdescri)-1
				xxxx = getInfoUser(A_SGU_UIGuserid(a_i))
				stowho = session("SGU_mail") & ""
				If Len(stowho) > 0 Then
					stowho = Trim(stowho)

					subject = "SOLICITUD DE AUTORIZACION"
					If WETIPO = "USR" Then
						XPARA2 = XPARA
					Else
						XPARA2 = XPARA & " / " & A_SGU_UIGdescri(a_i)
					End If
					'' message = "<u><font size=1 color=#000080 >" & XPARA2 & "</font></u><BR><BR><b><font size=1  >" & WEMESSAGE & "</font><b>"
					Xmessage = GenMDLMAIL (XPARA2,WEORIGEN,WEMESSAGE)
					Call SendMail(Xmessage, stowho, WEORIGEN,subject, session("MailComponent"),"SI","1", "NO")
				End If

				'RSXPA.MOVENEXT
			Next

			Case "SEC"
			XRESPO = 0 & XRESPO

			STRSQL = STRSQL & " SECTOR = " & XRESPO & " and ( NOT nivel >  " & WELVL & " ) "

		End Select







	End If

End Function

%>