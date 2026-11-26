<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

xfecLCID=Session.LCID
Session.LCID=2057

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else

	'response.write request.querystring & "<br>" & request.form

	Dim bconsultar()

	APLCHECK()

	Generar = ""

	Addqry = ""
	Response.expires = 0

	Server.scripttimeout = 600
	Response.buffer = True
	smenu = Request.querystring("Menu")
	stipoC = Request.querystring("tipoC")
	SEXPORT=REQUEST.FORM("EXPORT")
	empre = request.querystring("empre")
	SHTMLEXPORT="<TABLE>"

	If stipoC <> "" Then
		SESSION("TIPOC") = stipoC
	End If

	sorden = Request.querystring("sorden")

	If Len(sorden) = 0 Then
		sorden = "N"
	End If

	Sfiltro = Request.querystring("sfiltro")

	If Len(Sfiltro) = 0 Then
		Sfiltro = "TODOS"
	End If

	Sfiltro = Replace(Sfiltro, "_", " ")

	strsql = ""



	If Request.Form.Count > 0 Then

		sID = Request.Form("Desde3") + CLng(CDate("01-01-1900"))
		sID2 = Request.Form("Hasta3") + CLng(CDate("01-01-1900"))
		sID3 = Request.Form("Fecha")
		Addqry = "&Desde=" & sID & "&Hasta=" & sID2 & "&Fecha=" & sID3
	Else
		sID = Request.querystring("Desde") + CLng(CDate("01-01-1900"))
		sID2 = Request.querystring("Hasta") + CLng(CDate("01-01-1900"))
		sID3 = Request.querystring("Fecha")
		Addqry = "&Desde=" & sID & "&Hasta=" & sID2 & "&Fecha=" & sID3
	End If

	Select Case SESSION("TIPOC")


		Case  "ACCXESQ"

		XFDES = REQUEST.FORM("DESDE3")  + CLng(CDate("01-01-1900"))
		XFHAS = REQUEST.FORM("HASTA3")  + CLng(CDate("01-01-1900"))

		XFSEL = REQUEST.FORM("ESEL")


		Select Case XFSEL

			Case "MAS"


			XORDER = " ORDER BY CANTF DESC, FULTACC DESC, A.USUARIO "


			Case "MEN"

			XORDER = " ORDER BY CANTF , FULTACC DESC, A.USUARIO"

			Case "NOL"


			XORDER = " ORDER BY FULTACC DESC, A.USUARIO"

		End Select



		If XFSEL="NOL" Then

			STRSQL = "select z.usuario, z.usuario AS USERACC, max(z.FINICIO) as FULTACC ,0, 0 AS PROMF,0 AS CANTF from objalogdet Z "
			STRSQL = STRSQL & "where z.tipobjeto='SKRLOG' AND "

			STRSQL = STRSQL & " (z.FINICIO > "&(XFDES - 1) &"  AND z.FINICIO < "&(XFHAS + 1)&") "
			STRSQL = STRSQL & " and z.usuario not IN("
			strsql = strsql & "SELECT A.USUARIO"
			strsql = strsql & " FROM  "
			strsql = strsql & "       (select codobjeto,tipobjeto,usuario, max(fultacc) as fultacc, sum(acon) as acon "
			strsql = strsql & "            from OBJALOG where codemp=" & session("UN")
			strsql = strsql & "            GROUP BY codobjeto,tipobjeto,USUARIO "
			strsql = strsql & "            ) A "
			strsql = strsql & " LEFT JOIN (select CODOBJETO, TIPOBJETO, FINICIO, (HFINAL - HINICIO) AS PROM, "
			strsql = strsql & "            1 AS CANT, USUARIO "
			strsql = strsql & "            from objalogdet "
			strsql = strsql & "            where tacc = 'CONSEN' "
			strsql = strsql & "           ) B "
			strsql = strsql & "    ON A.USUARIO = B.USUARIO "

			strsql = strsql & " WHERE A.CODOBJETO = 0 AND A.TIPOBJETO = 'SKRLOG' AND (FINICIO > " & (XFDES - 1) & " "
			strsql = strsql & " AND FINICIO < " & (XFHAS + 1) & ") "
			strsql = strsql & " GROUP BY A.USUARIO ) group by z.usuario"

			'' response.write strsql & " strsql linea 114 " & "<br>"




			STRSQL = "select z.ID, z.ID AS USERACC, 0 as FULTACC ,0, 0 AS PROMF,0 AS CANTF from USUARIOS Z "
			STRSQL = STRSQL & "where  "
			STRSQL = STRSQL & "  z.ID not IN("
			STRSQL = STRSQL & " select DISTINCT W.usuario from objalogdet W "
			STRSQL = STRSQL & " where W.tipobjeto='SKRLOG' AND "
			STRSQL = STRSQL & " (W.FINICIO > "&(XFDES - 1) &"  AND W.FINICIO < "&(XFHAS + 1)&") "
			STRSQL = STRSQL & " ) order by z.descripcion "

			'response.write strsql & " strsql linea 127 " & "<br>"


		Else
			'response.write "HOLA" & "<br>"
			strsql = strsql & "SELECT A.USUARIO, A.USUARIO AS USERACC, MAX(A.FULTACC) AS FULTACC, SUM(A.ACON), "
			strsql = strsql & " AVG(ISNULL(PROM,0)) AS PROMF, SUM(ISNULL(CANT,0)) AS CANTF,c.cliente "
			strsql = strsql & " FROM  "
			strsql = strsql & "       (select codobjeto,tipobjeto,usuario, max(fultacc) as fultacc, sum(acon) as acon,codemp "

			If session("UN") <> 1 Then
				strsql = strsql & "            from OBJALOG where codemp=" & session("UN")
			Else
				strsql = strsql & "            from OBJALOG "
			End If

			strsql = strsql & "            GROUP BY codobjeto,tipobjeto,USUARIO,codemp "
			strsql = strsql & "            ) A "
			strsql = strsql & " LEFT JOIN (select CODOBJETO, TIPOBJETO.TIPOBJETO, FINICIO, (HFINAL - HINICIO) AS PROM, "
			strsql = strsql & "            1 AS CANT, USUARIO "
			strsql = strsql & "            from objalogdet "
			If session("base")="K" Then

				strsql=strsql & " inner join TIPOBJETO on tipobjeto.tipobjeto=objalogdet.tipobjeto and tipobjeto.TIPOCPT ='K'  "
			Else
				strsql=strsql & " inner join TIPOBJETO on tipobjeto.tipobjeto=objalogdet.tipobjeto and tipobjeto.TIPOCPT <>'K'  "
			End If
			strsql = strsql & "            where tacc = 'CONSEN' "
			strsql = strsql & "           ) B "
			strsql = strsql & "    ON A.USUARIO = B.USUARIO "
			if session("EMPMUN") = "S" then
				strsql = strsql & " inner join clienteplan c on (a.codemp=c.id) "
			else
				strsql = strsql & " inner join clienteplan c on (a.codemp=c.id and a.codemp=" & session("UN") & ") "
			end if
			strsql = strsql & " WHERE A.CODOBJETO = 0 AND A.TIPOBJETO = 'SKRLOG' AND (FINICIO > " & (sID - 1) & " "
			strsql = strsql & " AND FINICIO < " & (sID2 + 1) & ") "
			strsql = strsql & " GROUP BY A.USUARIO,a.codemp,c.cliente "
			strsql = strsql & XORDER

			'response.write strsql & " strsql linea 151 " & "<br>"


		End If

		' RESPONSE.WRITE STRSQL & "<BR>"

		Case  "ACCXESQDEP"

		XFDES = REQUEST.FORM("DESDE3")  + CLng(CDate("01-01-1900"))
		XFHAS = REQUEST.FORM("HASTA3")  + CLng(CDate("01-01-1900"))



		XFSEL = REQUEST.FORM("ESEL")
		XFDEP = "0" & REQUEST.FORM("LISTA")

		Call getUsrInSect(XFDEP,"")

		TABUSRXSECT = Split(session("SGU_UIGuserid"),Chr(9))

		WKFUSRXSECT = " AND A.USUARIO IN ('USUARIOS'"

		For IIUSR=0 To UBound(TABUSRXSECT)
			WKFUSRXSECT = WKFUSRXSECT & ",'" & TABUSRXSECT(IIUSR) & "'"

		Next

		WKFUSRXSECT = WKFUSRXSECT & ") "



		Select Case XFSEL

			Case "MAS"


			XORDER = " ORDER BY CANTF DESC, FULTACC DESC, A.USUARIO "


			Case "MEN"

			XORDER = " ORDER BY CANTF , FULTACC DESC, A.USUARIO"

			Case "NOL"


			XORDER = " ORDER BY FULTACC DESC, A.USUARIO"

		End Select





		strsql = strsql & "SELECT A.USUARIO, A.USUARIO AS USERACC, MAX(A.FULTACC) AS FULTACC, SUM(A.ACON), "
		strsql = strsql & " AVG(ISNULL(PROM,0)) AS PROMF, SUM(ISNULL(CANT,0)) AS CANTF "
		strsql = strsql & " FROM  "
		strsql = strsql & "       (select codobjeto,tipobjeto,usuario, max(fultacc) as fultacc, sum(acon) as acon "
		strsql = strsql & "            from OBJALOG where codemp=" & session("UN")
		strsql = strsql & "            GROUP BY codobjeto,tipobjeto,USUARIO "
		strsql = strsql & "            ) A "
		strsql = strsql & " LEFT JOIN (select CODOBJETO, TIPOBJETO.TIPOBJETO, FINICIO, (HFINAL - HINICIO) AS PROM, "
		strsql = strsql & "            1 AS CANT, USUARIO "
		strsql = strsql & "            from objalogdet "
		If session("base")="K" Then

			strsql=strsql & " inner join TIPOBJETO on tipobjeto.tipobjeto=objalogdet.tipobjeto and tipobjeto.TIPOCPT ='K'  "
		Else
			strsql=strsql & " inner join TIPOBJETO on tipobjeto.tipobjeto=objalogdet.tipobjeto and tipobjeto.TIPOCPT <>'K'  "
		End If
		strsql = strsql & "            where tacc = 'CONSEN' "
		strsql = strsql & "           ) B "
		strsql = strsql & "    ON A.USUARIO = B.USUARIO "
		''''   strsql = strsql & " INNER JOIN USUARIOS E ON E.ID = A.USUARIO "
		strsql = strsql & " WHERE A.CODOBJETO = 0 AND A.TIPOBJETO = 'SKRLOG' AND (FINICIO > " & (XFDES - 1) & " "
		strsql = strsql & " AND FINICIO < " & (XFHAS + 1) & ") " & WKFUSRXSECT




		strsql = strsql & " GROUP BY A.USUARIO "
		strsql = strsql & XORDER



		'RESPONSE.WRITE STRSQL & " strsql linea 231 " & "<BR>"




		Case  "ACCXUSR"




	End Select


	CantRegMostrar = CDbl(Request.querystring("CantRegMostrar"))
	CantRegMover = CDbl(Request.querystring("CantRegMover"))





	sverestado = " and objnivel >= " & Session("SegNivel") & " And Estcod In (9,3)"

	'response.write "STRSQL: " & strsql



	sMenuAdd = ""
	sMenuMas = "&smenu=" & smenu & "&Desde=" & sID & "&Hasta="
	sMenuMas = sMenuMas & sID2 & "&Fecha=" & sID3


	xsid = OFV.convertircar(sID, "/", "")
	xsid = OFV.convertircar(xsid, "-", "")
	If (Not IsNumeric(xsid)) Or (sID = "") Then sID = "1-1-1950"
	xsid = OFV.convertircar(sID2, "/", "")
	xsid = OFV.convertircar(xsid, "-", "")
	If (Not IsNumeric(xsid)) Or (sID2 = "") Then sID2 = "31-12-2100"


	Addqry = "&Menu=" & smenu & Addqry
	starget = "estactalis.asp?cantregmover=0&cantregmostrar=" & CantRegMostrar & Addqry  &"&perfil="&Replace(sperfil," ","_")
	sVolver = "estactalis.asp?Menu=" & smenu & sMenuAdd &"&perfil="&Replace(sperfil," ","_")
	Addqry = Addqry & "&sorden=" & sorden & "&sfiltro=" & Replace(Sfiltro, " ", "_")
	Set cn = OFV.conectar(OFV.strconn2)

	sql = "select cliente from clienteplan where id=" & session("UN")
	Set rx = ofv.crearconsultaEx(sql,cn,1,parametros)
	empresa = ""
	If Not rx.eof Then empresa = rx(0)
	Call OFV.cerrarconsulta(rx)

	Select Case SESSION("TIPOC")


		Case  "ACCXESQ"
		'response.write "SID: " & sID & "<br>" & "SID2: " & Sid2
		sTitulo = "ACCESO&nbsp;DE&nbsp;USUARIOS &nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;Desde:&nbsp;" & Fechador(sID) & "&nbsp;|&nbsp;&nbsp;&nbsp;Hasta:&nbsp;" & Fechador(sID2) & "&nbsp;|&nbsp;&nbsp;&nbsp;Empresa:&nbsp;" & empresa

		Case  "ACCXESQDEP"

		sUsrSECDSC = "----"

		Strsql2 = "     select DESCrI FROM ECO  "
		Strsql2 = Strsql2 & " WHERE ECO = '" & XFDEP & "' "
		'''      response.write strsql & "<br>"
		Set rs1 = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
		If Not rs1.eof Then
			sUsrSECDSC = rs1(0)
		End If


		sTitulo = "ACCESO&nbsp;POR&nbsp;DEPARTAMENTO:&nbsp;" & sUsrSECDSC & " &nbsp;&nbsp;&nbsp;|&nbsp;Desde:&nbsp;" & Fechador(sID) & "&nbsp;|&nbsp;Hasta:&nbsp;" & Fechador(sID2) & "&nbsp;|&nbsp;&nbsp;&nbsp;Empresa:&nbsp;" & empresa

		Case  Else

		sTitulo = "ACCESO&nbsp;DE&nbsp;USUARIOS &nbsp;&nbsp;&nbsp;|&nbsp;Desde:&nbsp;" & Fechador(sID) & "&nbsp;|&nbsp;Hasta:&nbsp;" & Fechador(sID2) & "&nbsp;|&nbsp;&nbsp;&nbsp;Empresa:&nbsp;" & empresa

	End Select



	'response.write "sql: " & strsql & "<br>"
	Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")

	sHTML = OFV.FormHeader(sTitulo)

	xs = "<script language=JavaScript>" & Chr(13)
	xs = xs & "function AbrirCons(Datos) {" & Chr(13)
	xs = xs & " myWin= open(Datos"
	xs = xs & "," & Chr(34) & "Cons" & Chr(34) & "," & Chr(34)
	xs = xs & "scrollbars=yes,menubar=no,toolbar=no,height=570,width=790" & Chr(34)
	xs = xs & " ); " & Chr(13) & " } " & Chr(13) & "</script>" & Chr(13)

	sHTML = sHTML & xs

	sHTML = sHTML & "<center>"

	Celda = OFV.GenCelda("", "ttt colspan=8", "", "", "", "", "", sTitulo, "si")

	Fila = OFV.GenRow("", "", "", "", "25", "", "", Celda, "si")


	Celda = OFV.GenCelda("", "ttr", "center", "5%", "", "", "", "*", "si")
	Celda = Celda & OFV.GenCelda("", "ttr ", "", "30%", "", "", "", "USUARIO", "si")
	'If session("UN") = 1 And SESSION("TIPOC") = "ACCXESQ" Then
	If session("EMPMUN") = "S" And SESSION("TIPOC") = "ACCXESQ" Then
		Celda = Celda & OFV.GenCelda("", "ttr ", "", "30%", "", "", "", "EMPRESA", "si")
	End If
	If session("APLMDCLI") <> "BGAL" Then
		Celda = Celda & OFV.GenCelda("", "ttr", "", "15%", "", "", "", "DEPARTAMENTO", "si")
	End If
	Celda = Celda & OFV.GenCelda("", "ttr", "", "10%", "", "", "", "ULTIMO ACCESO", "si")
	Celda = Celda & OFV.GenCelda("", "ttr", "", "5%", "", "", "", "CANTIDAD", "si")
	Celda = Celda & OFV.GenCelda("", "Ttr", "", "5%", "", "", "", "PROMEDIO", "si")

	Fila = fila & OFV.GenRow("", "", "", "", "25", "", "", Celda, "si")

	If SEXPORT <> "XLS" Then
		Response.write sHTML
		sHTML = ""
		Response.flush
	End If

	If SEXPORT = "XLS" Then
		SHTMLEXPORT=SHTMLEXPORT & "<TR><TD>" & "*" & "</TD>"
		SHTMLEXPORT=SHTMLEXPORT & "<TD><U>" & "USUARIO" & "</U></TD>"
		If session("APLMDCLI") <> "BGAL" Then
			SHTMLEXPORT=SHTMLEXPORT & "<TD><U>" & "DEPARTAMENTO" & "</U></TD>"
		End If
		SHTMLEXPORT=SHTMLEXPORT & "<TD><U>" & "ULTIMO ACCESO" & "</U></TD>"
		SHTMLEXPORT=SHTMLEXPORT & "<TD><U>" & "CANTIDAD" & "</U></TD>"
		SHTMLEXPORT=SHTMLEXPORT & "<TD><U>" & "PROMEDIO" & "</U></TD>"
		'mauro 09/05/17
		SHTMLEXPORT=SHTMLEXPORT & "<TD><U>" & "FECHA DESDE" & "</U></TD>"
		SHTMLEXPORT=SHTMLEXPORT & "<TD><U>" & "FECHA HASTA" & "</U></TD></TR>"
	End If


	nLinea = 0
	' response.write
	fi = CLng(CDate("01-01-1900"))
	PosReg = 0

	Do While (Not rs.EOF)

		'sImagen = "<font face=corflaticon size=3 color=#993333>&#0128;</font>&nbsp;"
		sImagen = "<i class='las la-user';></i>&nbsp;"

		dVigencia = "SD"
		If Not IsNull(rs(2)) Then dVigencia = CDate(CLng(rs(2)))
		dVign = InStr(dVigencia, "1900")
		If dVign > 0 Then dVigencia = "SD"


		If XFSEL="NOL" Then
			dVigencia = "-----"
		End If

		xxxx    = getInfoUserEX(RS(1),session("UN"))
		sGrpUsr = session("SGU_nombre")
		sUsrSEC = "0" & session("SGU_idcostcenter")
		sUsrSECDSC = "----"

		Strsql = "     select DESCrI FROM ECO  "
		Strsql = Strsql & " WHERE ECO = '" & sUsrSEC & "' "
		''response.write strsql & "<br>"
		Set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		If Not rs1.eof Then
			sUsrSECDSC = rs1(0)
		End If


		sobjeto = sGrpUsr

		sActcon = "CoESTDLISUSRDET.asp?USR=" & rs(0) & "&COD=NO"
		sActcon = sActcon & "&FDES=" & sID & "&FHAS=" & sID2
		sHref = "&nbsp;"

		sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "'); return false;" & Chr(34)




		Celda = ""
		If  session("APLMDCLI")="CFA" Then
			sLink=sImagen
		Else
			sLink = OFV.GenLink("", "UT", sHref, "", "", sFuncion, sImagen, "si")
		End If

		If SEXPORT = "XLS" Then
			SHTMLEXPORT=SHTMLEXPORT & "<TR><TD>" & POSREG+1 & "</TD>"
			SHTMLEXPORT=SHTMLEXPORT & "<TD>" & sobjeto & "</TD>"
			If session("APLMDCLI") <> "BGAL" Then
				SHTMLEXPORT=SHTMLEXPORT & "<TD>" & sUsrSECDSC & "</TD>"
			End If
			SHTMLEXPORT=SHTMLEXPORT & "<TD>" & FechadorcOR(dVigencia,0) & "</TD>"
			SHTMLEXPORT=SHTMLEXPORT & "<TD>" & RS(5) & "</TD>"
			SHTMLEXPORT=SHTMLEXPORT & "<TD>" & GETGLBHMS(RS(4)) & "</TD>"
			'SHTMLEXPORT=SHTMLEXPORT & "<TD>" & RS(5) & "</TD>"
			'SHTMLEXPORT=SHTMLEXPORT & "<TD>" & RS(5) & "</TD>"
			SHTMLEXPORT=SHTMLEXPORT &  "<TD>" & fechadorcor(CDate(XFDES),"") & "</TD>"
			SHTMLEXPORT=SHTMLEXPORT &  "<TD>" & fechadorcor(CDate(XFHAS),"") & "</TD></TR>"
		End If
		Celda = Celda & OFV.GenCelda("", "UTS", "center", "", "", "", "",sLink, "si")


		If  session("APLMDCLI")="CFA" Then
			sLink=sobjeto
		Else
			sLink = OFV.GenLink("", "UT", sHref, "", "", sFuncion, sobjeto, "si")
		End If

		If Session("EMPMUN")  = "S" Then
			Celda = Celda & OFV.GenCelda("", "UT ", "", "30%", "", "", "", sLink, "si")

			If session("UN") = 1 And SESSION("TIPOC") = "ACCXESQ" Then
				Celda = Celda & OFV.GenCelda("", "UT ", "", "30%", "", "", "", rs(6), "si")
			End If
		Else
			if session("APLMDCLI") <> "BGAL" then
				Celda = Celda & OFV.GenCelda("", "UT xCOLSPAN=3 ", "", "", "", "", "", sLink, "si")
			else
				Celda = Celda & OFV.GenCelda("", "UT ", "", "", "", "", "", sLink, "si")
			end if
		End If


		If session("APLMDCLI") <> "BGAL" Then
			Celda = Celda & OFV.GenCelda("", "UT  ", "", "15%", "", "", "", sUsrSECDSC, "si")
		End If

		Celda = Celda & OFV.GenCelda("", "UT ", "", "10%", "", "", "", FechadorcOR(dVigencia,0), "si")
		Celda = Celda & OFV.GenCelda("", "UT ", "center", "5%", "", "", "", RS(5) & "&nbsp;&nbsp;&nbsp;", "si")

		Celda = Celda & OFV.GenCelda("", "UT ", "center", "5%", "", "", "", GETGLBHMS(RS(4)) & "&nbsp;", "si")

		Fila = Fila & OFV.GenRow("", "", "", "", "25", "", "", Celda, "si")

		PosReg = PosReg + 1
		rs.movenext
		nLinea = nLinea + 1
		If SEXPORT <> "XLS"  Then

			sHTML = ""
			If nLinea > 20 Then
				nLinea = 0
				Response.flush
			End If
		End If

	Loop

	If PosReg > 0 Then


		Celda = ofv.GenCelda("","ttt  colspan=7 ","","","","",""," TOTAL: " & posreg & " USUARIOS             " & sInput,"si")

	Else


		Celda = ofv.GenCelda("","ttt  colspan=7 ","","","","",""," NO HAY USUARIOS PARA LA PRESENTE SELECCION ","si")
	End If

	Fila = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

	XTABLA = OFV.GenTabla("", "aaa", "", "", "", "0", "", "-1", "0", "", "", Fila, "si")


	Formulario = XTABLA

	sHTML = sHTML & Formulario

	Select Case SESSION("TIPOC")


		Case  "ACCXESQ"

		sAccion = "COESTDXUSR.asp"

		Case  "ACCXESQDEP"

		sAccion = "COESTDXDEP.asp"

		Case  Else

		sAccion = "COESTDXUSR.asp"

	End Select


	Celda   = ofv.BotonVolver(sAccion)

	sHTML = sHTML & OFV.botoneraESSX("0", "", "", "", "", "", "", "", "", "", "", "", Celda)

	sHTML = sHTML & "</body></html>"

	session("sperfil")=""


	Call OFV.cerrarconsulta(rs)
	Call OFV.cerrarconn(cn)

	''     Generar = sHTML


	If SEXPORT ="XLS" Then
		Response.ContentType = "application/vnd.ms-excel"
		Response.AddHeader "content-disposition", " filename=CoEstUlis.xls"
		response.write SHTMLEXPORT & "</TABLE>"
	Else
		response.write sHTML
		response.flush
	End If


	Session.LCID=xfecLCID
End If




Function Fechador(wkfecha)



	fec = CDate(wkfecha) ' + clng(cdate("01-01-1900")))
	ndia = ofv.cdia(Weekday(fec))
	nMes = ofv.cmes(Month(fec))

	Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)


End Function

%>