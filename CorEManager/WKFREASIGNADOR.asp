<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0



	CantRegMoverx = request.querystring("CantRegMoverx")
	SWKFCOD     = REQUEST.QUERYSTRING("WKFCOD")
	xtip= REQUEST.QUERYSTRING("xtip")
	tipobjeto= REQUEST.QUERYSTRING("tipobjeto")
	ID=REQUEST.QUERYSTRING("ID")
	OPWK = request.querystring("OPWK")
	SOPCION     = REQUEST.QUERYSTRING("OPCION")
    aplcheck()
	If ID & "" <> "" Then
		session("NCliente") = ID
	Else
		ID = session("NCliente")
	End If

    if tipobjeto & "" <> "" then
		session("TCliente") = tipobjeto
	else
		tipobjeto = session("TCliente")
	end if


	Set cn = ofv.conectar(ofv.strconn2)
	Set cnx = ofv.conectar(ofv.strconn5)

	Strsql = "Select * From OBJETOS where CODOBJETO = " & ID
	'response.write strsql
	Set rs = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
	If Not rs.EOF Then
		Strsql = "Select * From TIPOBJETO where TIPOBJETO = '" & TIPOBJETO & "' "

		Set rsX1 = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
		If Not rsX1.EOF Then
			TIPOOBJ = rsX1(1)
		End If

		If tipocpt = "K" Then

			session("NOMREQ") = TIPOOBJ & ":&nbsp;" & rs(1) & "&nbsp;-&nbsp;" & rs(3)

		Else

			session("NOMREQ") = TIPOOBJ & ":&nbsp;" & rs(1) & "&nbsp;-&nbsp;" & rs(3)

		End If

	End If

	Call ofv.cerrarconsulta(rs)
	Call ofv.cerrarconsulta(rsX1)

	If SOPCION = "" Then SOPCION = "NO"


	If SWKFCOD <> "" Then
		session("SWKFCOD")     = SWKFCOD
		session("CantRegMoverx")     = CantRegMoverx

	Else
		SWKFCOD     = session("SWKFCOD")
		CantRegMoverx     = session("CantRegMoverx")

	End If

	vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)



	Set cn = ofv.conectar(ofv.strconn2)
	Set cnx = ofv.conectar(ofv.strconn5)



	If SOPCION = "REASIGNA" Then


		sInputx = "COR@MANAGER.COM"

		If SESSION("WKFETPU") = "D" Then               ' EMISOR DEFAULT (EL RESPONSABLE DEL PASO)

			xxxx = getInfoUserw(SESSION("ORGUSR"))
			If session("SGU_mail") <> "" Then    SINPUTX = SESSION("SGU_mail")

			SINPUTX =  SINPUTX & ""

		ElseIf SESSION("WKFETPU") = "U" Then           ' EMISOR PREDETERMINADO (PARAMETRIZADO EL LA BASE DE USUARIOS)

			WKFEUSR = SESSION("WKFEUSR")

			If WKFEUSR = "" Then   WKFEUSR = "COREMANAGER"


			strsql = "Select EMAIL From USUARIOS where ID = '" & WKFEUSR & "' "
			Set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not rsniv.eof Then sInputx = rsniv(0)
			Call ofv.cerrarconsulta(rsniv)

		Else                                           ' EMISOR FIJO (USUARIO Y MAIL FIJOS EN CORCONECT)

			sInputx = SESSION("WKFEMAI")


		End If



		If SINPUTX = "" Then SINPUTX = "COR@MANAGER.COM"

		SESSION("ORGMAIL") = SINPUTX

		STRSQL = "SELECT * FROM SEGPASO WHERE SEGCKTPSC = '" & SWKFCOD & "' "
		Set rs = ofv.crearconsultaEx(StrSql,cnX,1,parametros)
		If Not rs.eof Then

			USRANT = REQUEST.FORM("RSPANT") & ""
			USRNEW = REQUEST.FORM("RSPNEW") & ""
			CHGOPC = REQUEST.FORM("RSPOPC") & ""
			WKFPASCOD = RS(3) & ""
			SEGCKT    = RS(1) & ""


			If CHGOPC = "2" Then

				RSKCOD = "0"

				STRSQL = "SELECT * FROM CKTVIEW WHERE SEGCKTPSC = '" & SWKFCOD & "' "
				Set rs2 = ofv.crearconsultaEx(StrSql,cnX,1,parametros)
				If Not rs2.eof Then

					RSKCOD = RS2(4) & ""

				End If

				STRSQL = "SELECT * FROM WKFCKTMASTER WHERE WKFPASCOD = '" & WKFPASCOD & "' "
				Set rs2 = ofv.crearconsultaEx(StrSql,cnX,1,parametros)
				If Not rs2.eof Then

					WKFRESPONSABLE = RS2(16) & ""


					STRSQL = "SELECT * FROM OBJWFLDET WHERE TIPOBJETO = 'RSK' AND CODOBJETO = " & RSKCOD & " AND COMEN1 = '" & WKFRESPONSABLE & "' "
					Set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
					If Not rs3.eof Then

						strsql = "UPDATE OBJWFLDET SET USR = '" & USRNEW & "'"
						Strsql = Strsql & " WHERE TIPOBJETO = 'RSK' AND CODOBJETO = " & RSKCOD & " "
						Strsql = Strsql & " AND COMEN1 = '" & WKFRESPONSABLE & "' "
						Call ofv.crearconsultaEx(StrSql,cn,1,parametros)


						strsql = "UPDATE SEGCKTRESP SET CODIGO = '" & USRNEW & "'"
						Strsql = Strsql & " WHERE SEGCKT = " & SEGCKT & " "
						Strsql = Strsql & " AND MODO = '" & WKFRESPONSABLE & "' "
						Call ofv.crearconsultaEx(StrSql,cnx,1,parametros)


					End If




				End If

			End If



			Strsql = ""

			SEGCKTPSC = GTCKTPSC()

			'''  SEGCKTPSC = RS(2) & "RA"

			XCOM    = GETDH()
			XFIN    = "NO"
			XOBSERV = "REASIGNACION A OTRO RESPONSABLE"
			XNIVEL  = RS(12)
			XSUBCOD = RS(13)
			XRCOD   = USRNEW
			XRASG   = XRCOD

			strsql = "select orden,nivel,subcodigo from segcktresp where segckt='" & segckt & "' and modo='L' and codigo='" & rs(9) & "'"
			Set rg = ofv.crearconsultaex(strsql,cnX,1,parametros)

			xxx = getinfouser(XRCOD)

            if session("SGU_idcostcenter") & "" = "" or not isnumeric(session("SGU_idcostcenter")) then session("SGU_idcostcenter")="0"

			sql = "insert into segcktresp(segckt,segcktpsc,modo,tipo,codigo,orden,nivel,subcodigo,segemail,codempre,sector) values "
			sql = sql & " ('" & segckt & "','" & rs(3) & "','L','" & rs(7) & "','" & XRCOD & "','" & rg(0) & "','" & rg(1) & "','" & rg(2) & "','" & session("SGU_MAIL") & "','" & session("SGU_site") & "','" & session("SGU_idcostcenter") & "')"

			Call ofv.crearconsultaex(sql,cnX,1,parametros)

			strsql = "select ID from segcktresp where segckt='" & segckt & "' and modo='L' and codigo='" & XRCOD & "'"
			Set rx = ofv.crearconsultaex(strsql,cnX,1,parametros)

			strsql = "insert into SEGPASO (SEGCKT,SEGCKTPSC,MODEL,COMIENZO,FIN,OKRCH,RESPTIPO,RESPONSABLE,RESPASG,PSCDSC,OBSERVACIONES,NIVEL,SUBCODIGO,SEGIDRESP) values "
			STRSQL = STRSQL & " ('" & RS(1) & "','" & SEGCKTPSC & "','" & RS(3) & "','" & XCOM & "','" & XFIN & "','NO','USR','" & XRCOD & "','" & XRASG & "','" & RS(10) & "','','" & XNIVEL & "','" & XSUBCOD & "','" & rx(0) & "')"
			Call ofv.crearconsultaEx(StrSql,cnX,1,parametros)


			strsql = "UPDATE SEGPASO SET FIN = '" & XCOM & "', OKRCH = '" & SEGCKTPSC
			Strsql = Strsql & "', observaciones = '" & XOBSERV & "'"
			Strsql = Strsql & " WHERE SEGCKTPSC = '" & SWKFCOD & "' "
			Call ofv.crearconsultaEx(StrSql,cnX,1,parametros)


			sql = "select responsable from segcircuito where objcod='" & session("NCliente") & "' and OBJTIPO='" & session("TCliente") & "' order by ID desc"
			Set rsresp = ofv.crearconsultaex(sql,cnx,1,parametros)

			Dim usrcc

			If Not rsresp.eof  and session("APLMDCLI")<>"xBGAL" Then

				xxx = getinfouser(rsresp(0))

				usrcc = SESSION("SGU_mail")
                XXXX = GENCKTMAIL("USR",rsresp(0),"REASIGNACION DEL PASO A NUEVO RESPONSABLE<br><br>" & RS(10) & " POR @@@NORMREQ@@@",SESSION("ORGMAIL"),"8")
			End If

            Call ofv.cerrarconsulta(rsresp)

			'''''''''''''''''''
			' agregar envio de mail


			XXXX = GENCKTMAIL("USR",XRCOD,"REASIGNACION DEL PASO A NUEVO RESPONSABLE<br><br>" & RS(10) & " POR @@@NORMREQ@@@",SESSION("ORGMAIL"),"8")
			'''''''''''''''''''


		End If

		If OPWK = "REASIGNAR" Then

			'response.write "<script>window.top.close();</script>"

		Else

			'RESPONSE.REDIRECT "listapendientes.asp?CANTREGMOVER=" & CANTREGMOVERX
			'RESPONSE.REDIRECT "imgfiles.asp"
		End If

	End If









	sHTML = ofv.FormHeader(session("FormN"))

	If OPWK = "REASIGNAR" Then

		shtml = shtml & CorPnlFnc("N","")

		sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

	End If

	Fila = ""
	xTrColor = "#F9F9F9"
	WKFINSELECT = ""



	''RESPONSE.WRITE " QS:" & REQUEST.QUERYSTRING & " RF:" & REQUEST.FORM & " FB:" & SESSION("FBUSR") & "<BR>"


	sImagevMa    = "<font face='arial' size=1 xcolor='#aa0000' style='font-weight:900;background-color:#aa0000;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;MUY&nbsp;ALTO&nbsp;</font>"
	sImageva    = "<font face='arial' size=1 xcolor='#f7aa22' style='font-weight:900;background-color:#f7aa22;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;ALTO&nbsp;</font>"
	sImagevm    = "<font face='arial' size=1 xcolor='#f7EEDD' style='font-weight:900;background-color:#f7EEDD;color:#888888;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;MODERADO&nbsp;</font>"
	sImagevb    = "<font face='arial' size=1 xcolor='#6699CC' style='font-weight:900;background-color:#6699CC;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;ESCASO&nbsp;</font>"
	sImagevNUL    = "<font face='arial' size=1 xcolor='#00AA00' style='font-weight:900;background-color:#00AA00;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;NULO&nbsp;</font>"

	sImageSI    = "<font face='arial' size=1  style='font-weight:900;background-color:#00AA00;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;SI&nbsp;</font>"
	sImageNO    = "<font face='arial' size=1  style='font-weight:900;background-color:#AA0000;color:#ffffff;border-style:solid;border-width:1;border-color:#aaaaaa;' >&nbsp;NO&nbsp;</font>"

	sImagevn    = "<font face='arial' size=1 xcolor='#dddddd' style='font-weight:900;background-color:#dddddd;color:#ffffff' >----</font>"

	wkstylof = "AAA style='font-family:arial;font-size:6pt;color:#888888;border-style:solid;border-width:1;border-color:#aaaaaa;' "
	wkstylof2 = "UTI style='border-style:solid;border-width:1;border-color:#aaaaaa;' "

	''''b.codvinobj = 'VN' AND


	Strsql = ""


	Strsql = Strsql & " select a.* from ( "

	Strsql = Strsql & " SELECT b.*, cast('12/12/2050' as datetime) - (cast(substring(b.segcktpsc,10,5) as integer)  ) as FECCOMIENZO,  "
	Strsql = Strsql & " cast( ( getdate() - ( cast('12/12/2050' as datetime) - (cast(substring(b.segcktpsc,10,5) as integer)  ) )  ) as integer) as DIFER "
	Strsql = Strsql & " FROM CKTVIEW b "
	Strsql = Strsql & " ) A "

	Strsql = Strsql & " WHERE a.segcktpsc = '" & SWKFCOD & "' "


	''  response.write strsql & "<br>"
	Set rs = ofv.crearconsultaEx(StrSql,cnX,1,parametros)
	If rs.eof Then

		Celda = ""
		Celda = Celda & ofv.gencelda("","VTG colspan=2 style='background-color:#F0F0F0;color:#888888' " & wkstylo1,"CENTER","","","","","NO HAY CIRCUITOS PENDIENTES PARA<BR>LA BUSQUEDA SOLICITADA","si")
		Fila  = FILA & ofv.genrow("","","","","50","","",Celda,"si")


	Else
		posI = 0



		'wkstylo1 = " style='border-style:solid;border-width:1;border-color:#aaaaaa;COLOR:#000000;' "
		'wkstyloT = " style='border-style:solid;border-width:1;border-color:#aaaaaa;COLOR:#FFFFFF;BACKGROUND-COLOR:#888888;' "


		WKIPROCc = 0
		MAINPROC = 0
		MAINACTI = ""

		Celda = ""
		Celda = Celda & ofv.gencelda("","ttt colspan=2 ","","","","","","LISTA DE CIRCUITOS PENDIENTES","si")
		Fila  = FILA & ofv.genrow("","","","","25","","",Celda,"si")



		sInput2 = rs("feccomienzo")
		sinput3 = rs("difer")





		'''#F7DD88

		WKIACTIV = ""

		WKIPROCD = ""
		WKIPROCN = 0


		Celda = ""

		sInput = RS(4)

		Strsql = ""
		If xtip="PD" Then
			codobj=0
			SQL= "SELECT  TEXTOIMPRESION,'Solicitud nro:' + cast (a.objid as nvarchar) + '   ' + B.DESCRI + ':' + BA.NOMBRE FROM PEDIDOS A "
			SQL= SQL  & "  INNER JOIN PEDIDOSDET BA ON(BA.IDPEDIDO=A.OBJID)"
			SQL= SQL  & "  INNER JOIN TIPOBJETO B ON(B.TIPOBJETO=A.TIPOBJETO)"
			SQL= SQL  & " WHERE A.OBJID=0" & ID
			'RESPONSE.WRITE SQL
			'TIPOBJETO WHERE TIPOBJETO='" & TIPOBJETO & "'"
			Set rsx1 = ofv.crearconsultaEx(SQL,cn,1,parametros)
			SDOC=""
			If Not rsx1.EOF Then
				xxcodobj=Split(rsx1(0),"-")
				If IsArray(xxcodobj) And UBound(xxcodobj)=1 Then
					codobj=xxcodobj(1)
					SDOC=rsx1(1)
				End If
			End If

			Call ofv.cerrarconsulta(rsx1)

			StrSql="   SELECT  b.codprodserv , '<b>'+ a.descri + '</b> / ' + b.descri as descri2 FROM  tema a "
			StrSql=StrSql & "        inner join     PRODSERV b on(a.codtema=b.codtema)"
			StrSql=StrSql & "  WHERE CODPRODSERV=0" & codobj
			StrSql= StrSql & " order by a.descri,b.descri"
			Set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not rsx1.eof Then
				sinput = rsx1(1)
			Else
				sinput=""
			End If
			Call ofv.cerrarconsulta(rsx1)
		Else




			Strsql = ""
			Strsql = Strsql & " select * "
			Strsql = Strsql & " from csrelpathdescriex a   "
			Strsql = Strsql & " inner join objetos b on (a.cobj = b.codobjeto)   "

			Strsql = Strsql & " where  a.cobj = 0" & sInput & "   "

			Set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)



			If Not rsx1.eof Then
				sinput=rsx1("temadescri") & " / " & rsx1("stemadescri")
				SDOC=RSX1("TOBJDESCRI") & " : " & RSX1("EXPR7")
			Else
				sinput="---"
				SDOC=""
			End If

			Call ofv.cerrarconsulta(rsx1)
		End If
		''''     sInput = WKIPROCD & " - (" & rsX1(0) & ") - " & rs(4)





		Celda = ""
		Celda = Celda & ofv.gencelda("","tt" ,"","30%","","","","TEMA / SUBTEMA:","si")
		Celda = Celda & ofv.gencelda("","UTi style=border-style:solid;border-width:1;border-color:#aaaaaa;" & wkstylo1,"","","","","",sInput,"si")
		Fila  = FILA & ofv.genrow("","aaa Xbgcolor='#ddddff'  ","","","25","","",Celda,"si")

		Celda =  ofv.gencelda("","tt","","30%","","","","DOCUMENTO:","si")
		Celda = Celda & ofv.gencelda("","UTi style=border-style:solid;border-width:1;border-color:#aaaaaa;" & wkstylo1,"","","","","",SDOC,"si")
		Fila  = FILA & ofv.genrow("","aaa  ","","","25","","",Celda,"si")


		sInput = RS(14)

		xxxx = getInfoUser(RS(14))
		If session("SGU_nombre") <> "" Then
			sInput =  "<A HREF='WKFREASIGNADORUSR.ASP?" & request.querystring & "' CLASS='NT' style='color:" & Session("CMFPNBOT") & "' TITLE='REASIGNAR PASO A OTRO USUARIO.' ><U>" & session("SGU_nombre") & "</U></A>"
		Else
			sInput =  "<A HREF='WKFREASIGNADORUSR.ASP' CLASS='NT' TITLE='REASIGNAR PASO A OTRO USUARIO.' ><U>" & RS(14) & "</U></A>"
		End If


		Celda = ""
		Celda = Celda & ofv.gencelda("","tt","","","","","","USUARIO RESPONSABLE:","si")
		Celda = Celda & ofv.gencelda("","UTi style=border-style:solid;border-width:1;border-color:#aaaaaa;" & wkstylo1 & "  ","","","","","",sInput,"si")
		Fila  = FILA & ofv.genrow("","aaa   ","","","25","","",Celda,"si")



		Celda = ""
		Celda = Celda & ofv.gencelda("","tt","","","","","","PASO:","si")

		sInput = RS(7) & "&nbsp;/&nbsp;" & RS(16)
		Celda = Celda & ofv.gencelda("","UTi style=border-style:solid;border-width:1;border-color:#aaaaaa;" & wkstylo1,"","","","","",sInput,"si")
		Fila  = FILA & ofv.genrow("","aaa   ","","","25","","",Celda,"si")


		Celda = ""
		Celda = Celda & ofv.gencelda("","tt","","","","","","INICIO:","si")

		sinput = sInput2 & "&nbsp;"
		Celda = Celda & ofv.gencelda("", "UTI style=border-style:solid;border-width:1;border-color:#aaaaaa;" & wkstylo1,"","","","","",sInput,"si")
		Fila  = FILA & ofv.genrow("","aaa   ","","","25","","",Celda,"si")



		Celda = ""
		Celda = Celda & ofv.gencelda("","tt","","","","","","NUEVO RESPONSABLE:","si")


		sInput = "SIN DETERMINAR"
		SUSUARIO = "NO"

		If SOPCION = "SETUSR" Then

			If SESSION("FBUSR") <> "NO" Then
				SUSUARIOX = SESSION("FBUSR")
				SUSUARIO  = Mid(SUSUARIOX,5,Len(SUSUARIOX)-4)
			End If

		End If



		If SUSUARIO <> "NO" Then

			xxxx = getInfoUser(SUSUARIO)
			If session("SGU_nombre") <> "" Then
				sInput =   "<U><FONT COLOR='#993333' >" & session("SGU_nombre") & "</FONT></U>"
			End If

		End If

		Celda = Celda & ofv.gencelda("","UTi  style=border-style:solid;border-width:1;border-color:#aaaaaa;" & wkstylo1 & "  ","","","","","",sInput,"si")

		Fila  = FILA & ofv.genrow("","aaa   ","","","25","","",Celda,"si")








	End If
	sHTML = sHTML & "<CENTER>" & ofv.GenTabla("","aaa ","","80% ","","0","","0","-1","","",Fila,"si") & "</CENTER>"


	If SOPCION = "SETUSR" Then

		If SESSION("FBUSR") <> "NO" Then
			SUSUARIOX = SESSION("FBUSR")
			SUSUARIO  = Mid(SUSUARIOX,5,Len(SUSUARIOX)-4)

			sAccion = "WKFREASIGNADOR.ASP?OPCION=REASIGNA" & "&OPWK=" & OPWK & "&id=" & id & "&tipobjeto=" & tipobjeto
			sHTML = sHTML & "<CENTER>"
			sHTML = sHTML & ofv.GenForm("AgreArchb","",sAccion,"","post","","","no")

			sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & RS(14) & "' NAME='RSPANT' >"
			sHTML = sHTML & "<INPUT TYPE='HIDDEN' VALUE='" & SUSUARIO & "' NAME='RSPNEW' >"



			sHTML = sHTML & "<SELECT NAME='RSPOPC' >"
			sHTML = sHTML & "<OPTION VALUE='1' SELECTED >SOLO PARA ESTE PASO</OPTION>"
			'sHTML = sHTML & "<OPTION VALUE='2'          >CAMBIAR TAMBIEN EN AUTOEVALUACION</OPTION>"
			sHTML = sHTML & "</SELECT><BR><BR>"


			Celda = ofv.gencelda("","VT  " & wkstylo1 & "   ","CENTER","","","","","<FONT COLOR='#993333' >CONFIRMA LA REASIGNACION DEL RESPONSABLE PARA ESTE PASO???</FONT>","si")

			Fila  = ofv.genrow("","aaa   ","","","35","","",Celda,"si")


			sHTML = sHTML & ofv.GenTabla("","aaa ","","80% ","","0","","0","-1","","",Fila,"si") & "</CENTER>"

			'Botones Aceptar y Cancelar
			If OPWK <> "REASIGNAR" Then
				sAccion = "listapendientes.asp?CANTREGMOVER=" & CANTREGMOVERX
			Else
				sAccion = "wkfreasignador.asp?ID=" & nCodigo & "&tipobjeto=" & stipo & "&WKFCOD=" & session("SEGCKTPSC") & "&OPWK=REASIGNAR" & "&CantRegMoverx=" & CantRegMoverx
				'response.write sAccion
			End If
			sHTML = sHTML & ofv.AgregarBoton("AgreArchb",sAccion,"_self")


		End If

	End If





	'Botonera
	Agregar = ""
	Buscar  = ""
	If OPWK <> "REASIGNAR" Then
		sAccion = "listapendientes.asp?CANTREGMOVER=" & CANTREGMOVERX
		Celda = ofv.BotonVolver(sAccion)
	Else
		Celda = ""
	End If

	'  sAccion = "wkfreasignador.asp?CANTREGMOVER=" & CANTREGMOVERX & request.querystring




	sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)
	'  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

	Call ofv.cerrarconsulta(rs)
	Call ofv.cerrarconn(cn)
	Call ofv.cerrarconn(cnx)
	'--------------------------------------------------------------------------------
	If OPWK = "REASIGNAR" Then

		sHTML = sHTML & "</span></div>"

		Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
		Fila = ofv.GenRow("","","","","","","",Celda,"si")
		sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

		'Estructura del Indice

		sHTML = sHTML & CorPnlIndx("N","")

	End If
	'--------------------------------------------------------------------------------
	sHTML = sHTML & "</body></html>"
End If

Response.Write sHTML

Function GETDH()

	XKD = DatePart("d",Date())
	XKM = DatePart("m",Date())
	XKY = DatePart("yyyy",Date())

	GETDH = XKD & "/" & XKM & "/" & XKY & " " & Time()


End Function



Function GENCKTMAIL(WETIPO,WECOD,WEMESSAGE,WEORIGEN,WELVL)


	XRESPO = WECOD & ""

	If XRESPO <> "" And XRESPO <> "NO" Then

		Set cn2 = ofv.conectar(ofv.strconn0)

		XPARA = GETRESPOCKT(WECOD,WETIPO)
		STRSQL = "SELECT EMAIL,DESCRIPCION FROM USUARIOS WHERE "

		If OPWK = "REASIGNAR" Then
			XGUMENSAGEw = WEMESSAGE & "<br><br>" & "Se encuentra disponible un proyecto normativo en la Aplicación EMANAGER V4.8 pendiente de su aprobación. Agradeceremos nos hagan llegar vuestros comentarios y/o aprobación" & "<br><br>" & "<APLURL>INGRESE A LA APLICACION</APLURL>"
		Else
			XGUMENSAGEw = WEMESSAGE
		End If
		XMESAURL  = ""

		If  SESSION("URLTYP")  = "V" Then

			XMESAURL  =  Left(request.servervariables("SCRIPT_NAME"),InStr(Right(request.servervariables("SCRIPT_NAME"),Len(request.servervariables("SCRIPT_NAME"))-1),"/"))

		End If

		If SESSION("URLPORT")  = "N" Then
			XMESAURL   = TipoSite("") & request.servervariables("SERVER_NAME")  & XMESAURL
		Else
			XMESAURL   = TipoSite("") & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & XMESAURL
		End If

		' XMESAURL   = "http://" & request.servervariables("SERVER_NAME") & ":" & request.servervariables("SERVER_PORT") & XMESAURL

		XGUMENSAGEw = Replace(XGUMENSAGEw,"<APLURL>","<U><A HREF='" & XMESAURL  & "' TARGET='_blank' >")
		XGUMENSAGEw = Replace(XGUMENSAGEw,"</APLURL>","</A></U>")

		WEMESSAGE = XGUMENSAGEw


		Select Case WETIPO

			Case "USR"
			'STRSQL = STRSQL & " ID = '" & XRESPO & "' "
			xxxx = getInfoUserw(XRESPO)
			If session("SGU_mail")<> "" Then    SINPUTX= session("SGU_mail")


			stowho = SINPUTX & ""
			If Len(stowho) > 0 Then
				stowho = Trim(stowho)

				If OPWK = "REASIGNAR" Then
					subject = "REASIGNACION DE AUTORIZACION"
				Else
					If session("MODMAIN") = "CL" Then
						subject = "REASIGNACION DE AUTORIZACION"
					Else
						subject = "SOLICITUD DE AUTORIZACION"
					End If
				End If
				If WETIPO = "USR" Then
					XPARA2 = XPARA
				Else
					XPARA2 = XPARA & " / " & session("SGU_nombre")
				End If
				''  message = "<u><font size=1 color=#000080 >" & XPARA2 & "</font></u><BR><BR><b><font size=1  >" & WEMESSAGE & "</font><b>"
				Xmessage = GenMDLMAIL (XPARA2,WEORIGEN,WEMESSAGE)
				Call SendMail(Xmessage, stowho, WEORIGEN,subject, session("MailComponent"),"SI","1", "NO")
			End If


			Case "GRP"

			'STRSQL = STRSQL & " GRUPO = '" & XRESPO & "' and ( NOT nivel >  " & WELVL & " ) "

			Call  getUsrIngroup(XRESPO)
			A_SGU_UIGuserid =Split(session("SGU_UIGuserid"),Chr(9))
			A_SGU_UIGdescri =Split(session("SGU_UIGdescri") ,Chr(9))

			'   set RSXPA = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
			' DO UNTIL RSXPA.EOF
			For a_i=0 To UBound(SGU_UIGdescri)-1
				xxxx = getInfoUserw(A_SGU_UIGuserid(a_i))
				stowho = session("SGU_mail") & ""
				If Len(stowho) > 0 Then
					stowho = Trim(stowho)

					If OPWK = "REASIGNAR" Then
						subject = "REASIGNACION DE AUTORIZACION"
					Else
						subject = "SOLICITUD DE AUTORIZACION"
					End If
					If WETIPO = "USR" Then
						XPARA2 = XPARA
					Else
						XPARA2 = XPARA & " / " & A_SGU_UIGdescri(a_i)
					End If
					''' message = "<u><font size=1 color=#000080 >" & XPARA2 & "</font></u><BR><BR><b><font size=1  >" & WEMESSAGE & "</font><b>"
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

Function GETRESPOCKT(RESPCOD,RESPTIP)

	GETRESPOCKT = "SIN ASIGNAR"



	XRESPO = RESPCOD & ""

	If XRESPO <> "" And XRESPO <> "NO" Then

		Set cn2 = ofv.conectar(ofv.strconn0)
		Set cn3 = ofv.conectar(ofv.strconn2)

		Select Case RESPTIP

			Case "USR"

			xxxx = getInfoUserw(XRESPO)
			sInputx = session("SGU_nombre")



			Case "GRP"
			strsql = "Select * From GRUPOS where ID = '" & XRESPO & "' "
			Set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
			If Not rsniv.eof Then sInputx = rsniv(1)
			Call ofv.cerrarconsulta(rsniv)

			Case "SEC"
			XRESPO = 0 & XRESPO
			strsql = "Select * From ECO where ECO = " & XRESPO
			'''     response.write strsql
			Set rsniv = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
			If Not rsniv.eof Then sInputx = rsniv(1)
			Call ofv.cerrarconsulta(rsniv)

		End Select

		GETRESPOCKT = sInputx

	End If


End Function


%>