
<!--#INCLUDE FILE="IncFile/CorConect.inc" -->

<!-- #INCLUDE FILE="IncFile/CorInit.inc" -->
<!-- #INCLUDE FILE="IncFile/CorSDK.inc" -->
<%

Set oIni = server.CreateObject("CorSecurityTCDC.Inicio")

Set oIni.Request = Request
Set oIni.Server = Server
Set oIni.Response = Response
Set oIni.Session = Session
Set oIni.Ofv = Ofv
oIni.Modal = Session("Modal")
APLCHECK()
session("noved") = "NO"

Opcion = Request.querystring("Opcion")
USRRV=REQUEST.QUERYSTRING("USUARIO") &""


If session("WEBACC")="S"  Then
	If session("APLMDCLI")="BP" Then
		If OPCION="IX" Then OPCION="IXW"
	ElseIf session("APLMDCLI")="RIVA" Then
		If OPCION="IX" Then OPCION="IXWF"

	End If
End If


If Opcion = "ISET" Then
	Opcion = "I"
	Session("EXPLOG") = "N"
	Session("SETLOG") = "S"
End If


If Opcion = "I" Then
	Conset(session("CONSET"))
	ofv.wstyles = CorStyler()

	If Not (Session("Usuario") & "") <> "" Then

		xddat1 =  request.querystring("ddat")
		xddat = Replace(CStr(xddat1),"___",Chr(9))
		xddatV = Split(xddat,Chr(9),-1,1)

		Session("Usuario")     = Replace(Replace(xddatV(0),"_"," "),"*","")
		Session("UsuarNT")     = Replace(Replace(xddatV(1),"_"," "),"*","")
		Session("DOMINIO")     = Replace(Replace(xddatV(2),"_"," "),"*","")
		SESSION("LOGIDUSRDET") = Replace(Replace(xddatV(3),"_"," "),"*","")


	End If

	Set cn2 = ofv.conectar(ofv.strconn2)

	XXXX = GETUSRLOGUPD("LOGOFF")

	Call ofv.cerrarconn(cn2)

	sHTML = LogFinal()
	Response.write sHTML

ElseIf Opcion = "IX" Or Opcion = "IXIND" Then

	SESSION("UN") = ""

	SESSION("VIEWOPC") = Opcion

	xcinit = corinit()

	SESSION("CLIEMP") = 0

	SESSION("USREXPIR") = "D"

	If  Session("Modal") = "MSNTT" And  Session("EXPLOG") = "S" Then
		sHTML = DIRECTLOG()
	Else
		''  sHTML = oIni.Inicio("")
		sHTML = Inicio("")
	End If



	Response.write sHTML

ElseIf Opcion = "IXW" Then
	xcinit = corinit()

	SESSION("CLIEMP") = 0

	SESSION("USREXPIR") = "D"


	sHTML = webvalida()



	Response.write sHTML
ElseIf Opcion = "IXWF" Then
	xcinit = corinit()

	SESSION("CLIEMP") = 0

	SESSION("USREXPIR") = "D"


	sHTML = webvalidaRiva()



	Response.write sHTML

ElseIf Opcion = "IXWV" Then


	vu = webvalidabt()


	If vu = "OK" Then
		Call ChkGrp()
		response.redirect "inicio.asp?opcion=VD"
	Else
		response.redirect "inicio.asp?opcion=IXW"
	End If

	Response.write sHTML
ElseIf Opcion = "IXEX" Then


	vu = LOGUINEx(USRRV)

	'response.write VU & "222"
	If vu = "OK" Then
		Call ChkGrp()
		response.redirect "inicio.asp?opcion=VD"
	Else
		response.redirect "inicio.asp?opcion=IXWF"
	End If

	Response.write sHTML
ElseIf Opcion = "IXWVF" Then


	vu = webvalidaRivaex()


	If vu = "OK" Then
		Call ChkGrp()
		response.redirect "inicio.asp?opcion=VD"
	Else
		response.redirect "inicio.asp?opcion=IXWF"
	End If

	Response.write sHTML


ElseIf Opcion = "IXS" Then

	Session("EXPLOG") = "N"
	Session("DELLOG") = "N"
	Session("ESPLOG") = "S"

	xcinit = corinit()

	SESSION("CLIEMP") = 0

	''  sHTML = oIni.Inicio("")
	sHTML = Inicio("")


	Response.write sHTML
ElseIf Opcion = "V" Or Opcion = "VD"  Or Opcion = "VDIND" Then
	If Session("APLBGD") = "" Then
		Conset(session("CONSET"))
		ofv.wstyles = CorStyler()
	End If
	session("TINI") = "I"


	If  Opcion = "VD" Then
		SHTML = "0"
	Else
		sHTML = oIni.Valida
	End If

	If Len(shtml) > 2 Then
		Response.write sHTML
	Else
		If sHTML = "1" Or sHTML = "2"  Or sHTML = "3"  Or sHTML = "6"  Or sHTML = "7"  Or sHTML = "8"  Or sHTML = "9"  Or sHTML = "11"  Then
			''' sHTML = oIni.Inicio(sHTML)
			sHTML = Inicio(sHTML)
			Response.write sHTML
		ElseIf sHTML = "0" Then
			'''  response.buffer=true
			If Session("Usuario") = UCase(Request.Form("Usuarnt")) Then
				session("TINI") = "Z"
			End If



			''''''''''''''
			Set XFCN = ofv.conectar(ofv.strconn0)


			XFStrSql = "Select sessexp,descripcion From Grupos where id = '" & Session("SegGrp") & "' "

			''''  response.write XFStrSql
			Set XFrs = ofv.crearconsultaEx(XFStrSql,XFCN,1,parametros)

			If  Not XFrs.eof Then
				SESSION("USREXPIR") = "D"

				If XFrs(0) <> "D" Then SESSION("USREXPIR") = CInt("0" & XFrs(0)) * 60
				If Session("EXPLOG") <> "S" Then session("SegGrpdsc") = XFRS(1)
			Else
				SESSION("USREXPIR") = "D"
			End If
			Call ofv.cerrarconsulta(XFrs)
			Call ofv.cerrarconn(XFCN)

			''''''''''' call ChkGrp()




			'''''''''''''' CONFIGURA EMPRESA DEL USUARIO
			If SESSION("CLIEMP") = 0 Then
				SESSION("SETEMP") = "N"
				Set cn = ofv.conectar(ofv.strconn2)

				'  StrSql = "Select SITE From USUARIOS WHERE ID = '" & SESSION("USUARIO") & "' "
				' Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
				' IF NOT RS.EOF THEN
				'   SESSION("CLIEMP") = clng(RS(0))
				'  ELSE
				'    SESSION("CLIEMP") = 1
				'  END IF

				XX=GETINFOUSER(SESSION("USUARIO"))
				IF session("SGU_site")="" THEN session("SGU_site")=1
				SESSION("CLIEMP")=session("SGU_site")
                Session("UsrMail")=session("SGU_mail")
                Session("UsrSector") = session("SGU_idcostcenter")
                Session("UsrPuesto") = session("SGU_codigopuesto")
				StrSql = "Select * From Clienteplan WHERE ID = " & SESSION("CLIEMP")
				'RESPONSE.WRITE STRSQL
				Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
				If Not RS.EOF Then
					Session("Cliente") = rs(1)
					Session("CliLogo") = "../COREss20/Empresa/" & rs(5)
					If (rs(11) & "" ) <> "" Then Session("CliLogoP") = "../COREss20/Empresa/" & rs(11)
					Session("USRSET")  = RS(10)
				End If

				StrSql = "Select * From SYSDEFDSKTOP WHERE OBJID = 0" & Session("USRSET")
				Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
				If Not RS.EOF Then
					If rs(2) <> "SI" Then
						XCONFUSRSET = CONUSRSET()
					Else
						Conset(session("CONSET"))
					End If
				End If

				Call ofv.cerrarconn(cn)
			End If
			''''''''''''''

			'''''''' Log ingreso '''''
			Set cn2 = ofv.conectar(ofv.strconn2)

			XXXX = GENUSRLOG()
			XXXX = GETUSRLOGUPD("LOGON")

			Call ofv.cerrarconn(cn2)
			'''''''''''''

            SELECT CASE session("APLMDCLI")
            CASE "BGAL"
			 xshtml = "<html><body ><form name='ejec' id='ejec' method='post' action='CorinicioGl.asp?Opcion=IS' style='visibility:hidden;' ><input  type='submit' value='n'   ></form>"
			CASE ELSE
			 xshtml = "<html><body ><form name='ejec' id='ejec' method='post' action='Corinicio.asp?Opcion=IS' style='visibility:hidden;' ><input  type='submit' value='n'   ></form>"
			END  SELECT
			xshtml = xshtml & "<script language='javascript'>document.ejec.submit();</script></body></html>"


			Response.write xshtml



		ElseIf sHTML = "4" Or sHTML = "5" Then

			'''''''' Log ingreso '''''
			Set cn2 = ofv.conectar(ofv.strconn2)

			XXXX = GENUSRLOG()
			XXXX = GETUSRLOGUPD("LOGON")

			Call ofv.cerrarconn(cn2)
			'''''''''''''

			'''''''''''''' CONFIGURA EMPRESA DEL USUARIO
			If SESSION("CLIEMP") = 0 Then
				SESSION("SETEMP") = "N"
				Set cn = ofv.conectar(ofv.strconn2)

				StrSql = "Select SITE From USUARIOS WHERE ID = '" & SESSION("USUARIO") & "' "
				Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
				If Not RS.EOF Then
					SESSION("CLIEMP") = CLng(RS(0))

				Else
					SESSION("CLIEMP") = 1
				End If

				StrSql = "Select * From Clienteplan WHERE ID = " & SESSION("CLIEMP")
				Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
				If Not RS.EOF Then
					Session("Cliente") = rs(1)
					Session("CliLogo") = "../COREss20/Empresa/" & rs(5)
					If (rs(11) & "" ) <> "" Then Session("CliLogoP") = "../COREss20/Empresa/" & rs(11)
					Session("USRSET")  = RS(10)
				End If

				StrSql = "Select * From SYSDEFDSKTOP WHERE OBJID = 0" & Session("USRSET")
				Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
				If Not RS.EOF Then
					If rs(2) <> "SI" Then

						XCONFUSRSET = CONUSRSET()
					Else
						Conset(session("CONSET"))
					End If
				End If

				Call ofv.cerrarconn(cn)
			End If
			''''''''''''''
            IF session("CORMENUNW")="S" THEN
             Response.redirect "CorinicioGl.asp?Opcion=CP"
            ELSE
			 Response.redirect "Corinicio.asp?Opcion=CP"
			END IF

		Else
			Response.write sHTML
		End If
	End If
ElseIf Opcion = "IS" Then

	xcinit = UsrInit()

	''' sHTML = oIni.iniciosesion
	sHTML = iniciosesion

	'''''''''''''' CONFIGURA EMPRESA DEL USUARIO
	'  IF SESSION("CLIEMP") = 0 THEN
	SESSION("SETEMP") = "N"
	Set cn = ofv.conectar(ofv.strconn2)

	'  StrSql = "Select SITE From USUARIOS WHERE ID = '" & SESSION("USUARIO") & "' "
	'  Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
	'  IF NOT RS.EOF THEN
	'     SESSION("CLIEMP") = clng(RS(0))

	'   ELSE
	'     SESSION("CLIEMP") = 1


	'  END IF
	XXX=GETINFOUSER(SESSION("USUARIO"))
	if session("SGU_site")="" then session("SGU_site")=1
	SESSION("CLIEMP")=session("SGU_site")
	StrSql = "Select * From Clienteplan WHERE ID = " & SESSION("CLIEMP")
	Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
	If Not RS.EOF Then
		Session("Cliente") = rs(1)
		Session("CliLogo") = "../COREss20/Empresa/" & rs(5)
		If (rs(11) & "" ) <> "" Then Session("CliLogoP") = "../COREss20/Empresa/" & rs(11)
		Session("USRSET")  = RS(10)
	End If

	StrSql = "Select * From SYSDEFDSKTOP WHERE OBJID = 0" & Session("USRSET")
	Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
	If Not RS.EOF Then
		If rs(2) <> "SI" Then
			XCONFUSRSET = CONUSRSET()
		Else
			Conset(session("CONSET"))
		End If
	End If


	'  END IF
	''''''''''''''
	''
	'''''''' CONFIGURACION GLOBAL MULTIEMPRESA
	''
	If SESSION("EMPMUN") = "S" Then
		StrSql = "Select * From Clienteplan WHERE ID = 1"
		Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
		If Not RS.EOF Then
			SESSION("GLBLOG") = RS(8)
			SESSION("GLBMUN") = CLng(RS(9))
		Else
			SESSION("GLBLOG") = 1
			SESSION("GLBMUN") = 1
		End If

	Else
		SESSION("GLBLOG") = 1
		SESSION("GLBMUN") = 1
	End If
	Call ofv.cerrarconn(cn)

	Response.write sHTML

ElseIf Opcion = "CP" Then
	sHTML = oIni.CambiaPass("")
	Response.write sHTML
ElseIf Opcion = "CPOK" Then
	sHTML = oIni.CambiaPassOK
	If Len(shtml) > 1 Then
		Response.write sHTML
	Else
		If sHTML = "0" Then
			Response.redirect "inicio.asp?Opcion=IS"
		Else
			sHTML = oIni.CambiaPass(sHTML)
			Response.write sHTML
		End If
	End If
End If

Function DIRECTLOG()

	If  Session("DELLOG") = "S" Then

		vu = OFV.ValidarUsuario()

		vu = "OK"

		usuario=UserLogon()

		'usuario="AMENDEZ"
		Call getInfoUser(Usuario)

		If session("USRPRV")="LACAJA" Then

			If session("SGU_nombre")="NO" Then
				session("USRPRV")="I"
				USUARIO="USRCONSULTA"
				Call getInfoUser(USUARIO)
				session("USRPRV")="LACAJA"
			End If
		End If

		Session("Usuario")   = session("SGU_userid")
		Session("UsrOpcion") = session("SGU_nombre")
		Session("SkrDefT")   = "0"
		Session("SkrDefE")   = "0"
		Session("SkrDefI")   = "0"
		Session("UsrPmin")   = "0"
		Session("UsrPmax")   = "0"
		Session("UsrHist")   = "0"
		Session("SegUsuario") = session("SGU_nombre")
		Session("SegNivel")  = CDbl("8")
		Session("SegGrp")    = "DEF"
		Session("SegPerfil") = Session("Usuario")
		Session("UsuarNT")   = Session("Usuario")

		Session("UsrSector") = session("SGU_idcostcenter")
		Session("UsrSecDsc") = session("SGU_descrisector")
		Session("UsrPuesto") = session("SGU_codigopuesto")
		Session("UsrPstDsc") = session("SGU_descripuesto")
		Session("UsrMail")   = session("SGU_mail")
		Session("UsrMail")   = session("SGU_mail")
		SESSION("CLIEMP")=0
		'SESSION("CLIEMP")= session("SGU_SITE")
		Call getGroupsUser(Usuario)

		Session("SegGrpOpt")    = "S"

		Session("SegGrpLst")    = session("SGU_grupos")

		Call ChkGrp()

		If Opcion = "IXIND" Then
			XFURL = "Inicio.asp?Opcion=VDIND"

		Else
			XFURL = "Inicio.asp?Opcion=VD"

		End If

	Else
		vu    = OFV.ValidarUsuario()
		XFURL = "Inicio.asp?Opcion=V"
	End If


	If vu = "OK" Then

		fondo = " TOPMARGIN=1 LEFTMARGIN=1 ONLOAD='document.INIT.submit();'   "

		UserNt = UCase(Request.ServerVariables("LOGON_USER"))
		UserNt = Mid(UserNt, InStrRev(UserNt, "\") + 1, Len(UserNt))

		DOMAINNT = UCase(Request.ServerVariables("LOGON_USER"))
		DOMAINNT = Left(DOMAINNT, InStr(DOMAINNT, "\")) & "\"
		DOMAINNT = Replace(DOMAINNT, "\", "")

		XHTML = OFV.MenuHeader("", fondo) & Chr(13) & "<center><BR><BR><BR><BR><BR>"

		XHTML = XHTML & "<font face=verdana size=5 color=#6699cc ><b>E-SOLUTION SUITE</B></FONT><BR><BR><BR>"


		sInput = "<BIG>VERIFICANDO ACCESO PARA</BIG>"

		Celda = OFV.gencelda("", "ttt VALIGN=MIDDLE ", "center", "80%", "50%", "", "", sInput, "si")
		Fila = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

		sInput1 = OFV.GenerarInput("usuarNt", UserNt, "HIDDEN", "", "80", "at", "no")
		sInput2 = OFV.GenerarInput("DOMAINNt", DOMAINNT, "HIDDEN", "", "80", "at", "no")
		sInput3 = OFV.GenerarInput("usuar", "", "HIDDEN", "", "62", "at", "no")
		sInput4 = OFV.GenerarInput("clave", "", "HIDDEN", "", "62", "at", "no")

		sInput = "<BR><BR>USUARIO:&nbsp;" & UserNt & "<BR><BR>DOMINIO:&nbsp;" & DOMAINNT & "<BR><BR>"

		Celda = OFV.gencelda("", "uts", "center", "80%", "20", "", "", sInput & sInput1 &  sInput2 &  sInput3 &  sInput4, "si")
		Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

		Formulario = OFV.GenForm("INIT", "", XFURL, "", "post", "", Fila, "si")

		XHTML = XHTML & OFV.GenTabla("", "AAA STYLE='BORDER-STYLE:solid;border-width:2;border-color:#000080;' ", "", "", "", "0", "", "0", "0", "", "", Formulario, "si")
		XHTML = XHTML & "</center></BODY></html>"

	End If

	DIRECTLOG = XHTML


End Function


Function WebValida()

	vu    = OFV.ValidarUsuario()
	XFURL = "Inicio.asp?Opcion=IXWV"


	fondo = " TOPMARGIN=1 LEFTMARGIN=1    "




	Set cn = ofv.conectar(ofv.strconn2)
	StrSql = "Select * From Clienteplan WHERE ID =1 "
	Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
	If Not RS.EOF Then
		Session("Cliente") = rs(1)
		Session("CliLogo") = "../COREss20/Empresa/" & rs(5)
		If (rs(11) & "" ) <> "" Then Session("CliLogoP") = "../COREss20/Empresa/" & rs(11)
		Session("USRSET")  = RS(10)
	End If
	Call ofv.cerrarconn(cn)


	XHTML =  "<img src=" & Session("CliLogo") &">"
	XHTML = XHTML & OFV.MenuHeader("", fondo) & Chr(13) & "<center><BR><BR><BR><BR><BR>"
	XHTML = XHTML & "<font face=verdana size=5 color=#6699cc ><b>COR E-MANAGER </B></FONT><BR>"
	XHTML = XHTML & "<font face=verdana size=2 color=#6699cc ><b>BASE DE CONOCIMIENTOS </B></FONT><BR><BR><BR>"



	FILA=""
	Celda = OFV.gencelda("", "ttt ", "center", "80%", "20", "", "", "<B>USUARIO:</B>", "si")
	Fila = Fila & OFV.GenRow("u11", "aaa style='display:block;cursor:pointer;' ", "", "", "", "", "", Celda, "si")

	sInput = OFV.GenerarInput("usuar", "", "agregar", "", "40", "dt", "no")
	Celda = OFV.gencelda("", "at ", "center", "80%", "20", "", "", sInput, "si")
	Fila = Fila & OFV.GenRow("u12", "aaa style='display:block;cursor:pointer;' ", "", "", "", "", "", Celda, "si")

	Celda = OFV.gencelda("", "at ", "center", "80%", "20", "", "", "<B>PASSWORD:</B>", "si")
	Fila = Fila & OFV.GenRow("u13", "aaa style='display:block;cursor:pointer;' ", "", "", "", "", "", Celda, "si")

	sInput = OFV.GenerarInput("clave", "", "password", "", "40", "dt", "no")
	Celda = OFV.gencelda("", "at ", "center", "80%", "20", "", "", sInput, "si")
	Fila = Fila & OFV.GenRow("u14", "aaa style='display:block;cursor:pointer;' ", "", "", "", "", "", Celda, "si")

	sInput = OFV.GenerarInput("xxx", "login", "submit", "", "62", "bt", "no")
	Celda = OFV.gencelda("", "at ", "center", "80%", "20", "", "", sInput, "si")
	Fila = Fila & OFV.GenRow("u14", "aaa style='display:block;cursor:pointer;' ", "", "", "", "", "", Celda, "si")

	sInput =session("mensajeWeb")
	Celda = OFV.gencelda("", "at ", "center", "80%", "20", "", "", sInput, "si")
	Fila = Fila & OFV.GenRow("u14", "aaa style='display:block;cursor:pointer;' ", "", "", "", "", "", Celda, "si")


	Formulario =OFV.GenForm("INIT", "", XFURL, "", "post", "", Fila , "si")

	XHTML = XHTML & OFV.GenTabla("", "AAA STYLE='BORDER-STYLE:solid;border-width:2;border-color:#000080;' ", "", "", "", "0", "", "0", "0", "", "", Formulario , "si")
	XHTML = XHTML & "</center></BODY></html>"



	WebValida = XHTML


End Function

Function WebValidaRiva()

	vu    = OFV.ValidarUsuario()
	XFURL = "Inicio.asp?Opcion=IXWVF"


	fondo = " TOPMARGIN=1 LEFTMARGIN=1    "




	Set cn = ofv.conectar(ofv.strconn2)
	StrSql = "Select * From Clienteplan WHERE ID =1 "
	Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")
	If Not RS.EOF Then
		Session("Cliente") = rs(1)
		Session("CliLogo") = "Empresa/" & rs(5)
		If (rs(11) & "" ) <> "" Then Session("CliLogoP") = "/Empresa/" & rs(11)
		Session("USRSET")  = RS(10)
	End If
	Call ofv.cerrarconn(cn)


	XHTML =  "<img src='" & Session("CliLogo") &"'>"
	XHTML = XHTML & OFV.MenuHeader("", fondo) & Chr(13) & "<center><BR><BR><BR><BR><BR>"
	XHTML = XHTML & "<font face=verdana size=5 color=#6699cc ><b>COR E-MANAGER </B></FONT><BR>"
	XHTML = XHTML & "<font face=verdana size=2 color=#6699cc ><b>BASE DE CONOCIMIENTOS </B></FONT><BR><BR><BR>"



	FILA=""
	Celda = OFV.gencelda("", "ttt ", "center", "80%", "20", "", "", "<B>USUARIO:</B>", "si")
	Fila = Fila & OFV.GenRow("u11", "aaa style='display:block;cursor:pointer;' ", "", "", "", "", "", Celda, "si")

	sInput = OFV.GenerarInput("usuar", "", "agregar", "", "40", "dt", "no")
	Celda = OFV.gencelda("", "at ", "center", "80%", "20", "", "", sInput, "si")
	Fila = Fila & OFV.GenRow("u12", "aaa style='display:block;cursor:pointer;' ", "", "", "", "", "", Celda, "si")

	Celda = OFV.gencelda("", "at ", "center", "80%", "20", "", "", "<B>PASSWORD:</B>", "si")
	Fila = Fila & OFV.GenRow("u13", "aaa style='display:block;cursor:pointer;' ", "", "", "", "", "", Celda, "si")

	sInput = OFV.GenerarInput("clave", "", "password", "", "40", "dt", "no")
	Celda = OFV.gencelda("", "at ", "center", "80%", "20", "", "", sInput, "si")
	Fila = Fila & OFV.GenRow("u14", "aaa style='display:block;cursor:pointer;' ", "", "", "", "", "", Celda, "si")

	sInput = OFV.GenerarInput("xxx", "login", "submit", "", "62", "bt", "no")
	Celda = OFV.gencelda("", "at ", "center", "80%", "20", "", "", sInput, "si")
	Fila = Fila & OFV.GenRow("u14", "aaa style='display:block;cursor:pointer;' ", "", "", "", "", "", Celda, "si")

	sInput =session("mensajeWeb")
	Celda = OFV.gencelda("", "at ", "center", "80%", "20", "", "", sInput, "si")
	Fila = Fila & OFV.GenRow("u14", "aaa style='display:block;cursor:pointer;' ", "", "", "", "", "", Celda, "si")


	Formulario =OFV.GenForm("INIT", "", XFURL, "", "post", "", Fila , "si")

	XHTML = XHTML & OFV.GenTabla("", "AAA STYLE='BORDER-STYLE:solid;border-width:2;border-color:#000080;' ", "", "", "", "0", "", "0", "0", "", "", Formulario , "si")
	XHTML = XHTML & "</center></BODY></html>"



	WebValidaRiva = XHTML


End Function



Function ChkGrp()



	'    SESSION("USREXPIR") = 50

	SESSION("USREXPIR") = "0"

	Set XFCN = ofv.conectar(ofv.strconn0)

	XFStrSql = "Select id,sessexp,descripcion From Grupos where "

	session("SegGrpLst") = Session("SegGrp") & Chr(9) & session("SegGrpLst")
	session("SegGrpdsc") = Session("SegGrp") & Chr(9) & session("SegGrpLst")


	XMFTGRP = Split(session("SegGrpLst"),Chr(9))
	XMFSEP  = ""


	' E galicia

	If session("USRPRV")  = "I" Or session("USRPRV")  = "E" Then
		XFStrSql = XFStrSql & "  id in ("
	Else
		XFStrSql = XFStrSql & " mapgrp in ("
	End If

	For XMFI = 0 To UBound(XMFTGRP)
		XFStrSql  = XFStrSql  & XMFSEP & "'" & Replace(XMFTGRP(XMFI) & "","'","''") & "'"
		XMFSEP  = ","
	Next

	XFStrSql = XFStrSql & ") "
	XFStrSql = XFStrSql & " Order By id"
	' session("sss")=XFStrSql
	Set XFrs = ofv.crearconsultaEx(XFStrSql,XFCN,1,parametros)

	If  Not XFrs.eof Then
		XMFSEP               = ""
		session("SegGrpLst") = ""
		session("SegGrpdsc") = ""
		Do Until XFRS.EOF

			If XFrs(1) <> "D" And SESSION("USREXPIR") <> "D" Then
				If (CInt("0" & XFrs(1)) * 60) > CInt(SESSION("USREXPIR")) Then
					SESSION("USREXPIR") = CInt("0" & XFrs(1)) * 60
				End If
			Else
				SESSION("USREXPIR") = "D"
			End If

			session("SegGrpLst") = session("SegGrpLst") & XMFSEP & XFRS(0)
			session("SegGrpdsc") = session("SegGrpdsc") & XMFSEP & XFRS(2)
			XMFSEP  = Chr(9)
			XFRS.MOVENEXT
		Loop

		If  session("SegGrpLst") = "" Then session("SegGrpLst") = Session("SegGrp")

	Else
		session("SegGrpLst") = Session("SegGrp")
		session("SegGrpdsc") = Session("SegGrp")
	End If
	Call ofv.cerrarconsulta(XFrs)
	Call ofv.cerrarconn(XFCN)

End Function

%>
