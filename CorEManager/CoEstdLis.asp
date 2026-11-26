<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else



	Dim bconsultar()


	Generar = ""

	Addqry = ""
	Response.expires = 0

	Server.scripttimeout = 60*30
	Response.buffer = True
	smenu = Request.querystring("Menu")
	stipoC = Request.querystring("tipoC")

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
	SEXPORT=REQUEST.FORM("EXPORT")
	SESELESTADOS=REQUEST.FORM("ESELESTADOS")&""
	' RESPONSE.WRITE SESELESTADOS
	' IF SESELESTADOS="TODOS" THEN SESELESTADOS= "3"
	'RESPONSE.WRITE REQUEST.querystring
	SHTMLEXPORT="<TABLE>"

	empresa = request.querystring("empre")

	Select Case SESSION("TIPOC")


		Case  "ACCXESQ"

		XFDES = REQUEST.FORM("DESDE3")  + CLng(CDate("01-01-1900"))
		XFHAS = REQUEST.FORM("HASTA3")  + CLng(CDate("01-01-1900"))



		XFSEL = REQUEST.FORM("ESEL")

		'response.write xfsel

		Select Case XFSEL

			Case "MAS"

			XWHERE = "  AND (FINICIO > " & (XFDES - 1) & " AND FINICIO < " & (XFHAS + 1) & ") "
			XORDER = "ORDER BY CANTF DESC, FULTACC DESC "


			Case "MEN"

			XWHERE = "  AND (FINICIO > " & (XFDES - 1) & " AND FINICIO < " & (XFHAS + 1) & ") "
			XORDER = "ORDER BY CANTF, FULTACC DESC "

			Case "NOL"

			XWHERE = " AND FINICIO IS NULL "
			XORDER = " ORDER BY FULTACC DESC "

		End Select

		strsql = " select A.CODOBJETO,  A.FULTACC, SUM(ISNULL(CANT,0)) AS CANTF, MAX(ESTADO) AS ESTADO  "





		strsql = strsql & " FROM OBJALOG A LEFT JOIN "
		If session("UN") = 1 Then
			strsql = strsql & " (select CODOBJETO, TIPOBJETO, FINICIO, AVG(HFINAL - HINICIO) AS PROM, SUM(1) AS CANT "
			strsql = strsql & " from objalogdet "
			strsql = strsql & " where tacc = 'CONSEN' "
			strsql = strsql & " GROUP BY CODOBJETO,TIPOBJETO,FINICIO) B "
		Else
			strsql = strsql & " (select objalogdet.CODOBJETO, objalogdet.TIPOBJETO, objalogdet.FINICIO, AVG(HFINAL - HINICIO) AS PROM, SUM(1) AS CANT "
			strsql = strsql & " from objalogdet "
			strsql = strsql & " inner join objalog c on (c.tipobjeto='SKRLOG' and c.usuario=objalogdet.usuario and c.codemp=" & session("UN") & ") "
			strsql = strsql & " where tacc = 'CONSEN' "
			strsql = strsql & " GROUP BY objalogdet.CODOBJETO,objalogdet.TIPOBJETO,objalogdet.FINICIO) B "
		End If
		strsql = strsql & "ON A.CODOBJETO = B.CODOBJETO AND A.TIPOBJETO = B.TIPOBJETO "
		strsql = strsql & " INNER JOIN OBJETOS C ON A.CODOBJETO = C.CODOBJETO  "
		strsql = strsql & " INNER JOIN OBJPRTY D ON A.CODOBJETO = D.CODOBJETO  "

		If SESSION("BASE")="K" Then
			strsql = strsql & " INNER JOIN tipobjeto Df ON A.tipobjeto = Df.tipobjeto and df.TIPOCPT='K'  "
		Else
			strsql = strsql & " INNER JOIN tipobjeto Df ON A.tipobjeto = Df.tipobjeto and df.TIPOCPT <> 'K'  "
		End If


		If  SESELESTADOS <> "TODOS" Then

			strsql = strsql & " WHERE A.CODOBJETO > 0  AND ESTADO in ("& SESELESTADOS &") "

			If session("UN") <> 1 Then
				strsql = strsql & " AND C.OBJUN=" & session("UN")
			End If

			strsql = strsql & " " & XWHERE

		Else

			SESELESTADOS="1,2,3,4,5,6"

			strsql = strsql & " WHERE A.CODOBJETO > 0  AND ESTADO in ("& SESELESTADOS &") "

			If session("UN") <> 1 Then
				strsql = strsql & " AND C.OBJUN=" & session("UN")
			End If

			strsql = strsql & " " & XWHERE

		End If

		strsql = strsql & " GROUP BY A.CODOBJETO,A.TIPOBJETO,FULTACC, ACON "
		strsql = strsql & XORDER

		'response.write strsql & " strsql " & "<br>"




		Case  "ACCXUSR"




	End Select


	CantRegMostrar = CDbl(Request.querystring("CantRegMostrar"))
	CantRegMover = CDbl(Request.querystring("CantRegMover"))





	sverestado = " and objnivel >= " & Session("SegNivel") & " And Estcod In (9,3)"


	If Request.Form.Count > 0 Then

		sID = Request.Form("Desde3")
		sID2 = Request.Form("Hasta3")
		sID3 = Request.Form("Fecha")
		Addqry = "&Desde=" & sID & "&Hasta=" & sID2 & "&Fecha=" & sID3
	Else
		sID = Request.querystring("Desde")
		sID2 = Request.querystring("Hasta")
		sID3 = Request.querystring("Fecha")
		Addqry = "&Desde=" & sID & "&Hasta=" & sID2 & "&Fecha=" & sID3
	End If

	sMenuAdd = ""
	sMenuMas = "&smenu=" & smenu & "&Desde=" & sID & "&Hasta="
	sMenuMas = sMenuMas & sID2 & "&Fecha=" & sID3


	xsid = OFV.convertircar(sID, "/", "")
	xsid = OFV.convertircar(xsid, "-", "")
	If (Not IsNumeric(xsid)) Or (sID = "") Then sID = "1-1-1950"
	xsid = OFV.convertircar(sID2, "/", "")
	xsid = OFV.convertircar(xsid, "-", "")
	If (Not IsNumeric(xsid)) Or (sID2 = "") Then sID2 = "31-12-2100"


	Addqry  = "&Menu=" & smenu & Addqry
	starget = "estactalis.asp?cantregmover=0&cantregmostrar=" & CantRegMostrar & Addqry  &"&perfil="&Replace(sperfil," ","_")
	sVolver = "estactalis.asp?Menu=" & smenu & sMenuAdd &"&perfil="&Replace(sperfil," ","_")
	Addqry  = Addqry & "&sorden=" & sorden & "&sfiltro=" & Replace(Sfiltro, " ", "_")
	Set cn  = OFV.conectar(OFV.strconn2)
	Set cnx  = OFV.conectar(OFV.strconn2)

	sql = "select cliente from clienteplan where id=" & session("UN")
	Set rx = OFV.crearconsultaEx(sql, cn, 1, parametros)
	empresa = ""
	If Not rx.eof Then empresa = rx(0)
	Call OFV.CERRARCONSULTA(RX)

	sTitulo ="DOCUMENTOS&nbsp;Vigentes&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & Fechador(sID) & "&nbsp;&nbsp;al&nbsp;&nbsp;" & Fechador(sID2) & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Empresa:&nbsp;" & empresa

	' response.write "sql: " & strsql & "<br>"
	' response.write BuscarIds(2219)
	Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")

	sHTML = OFV.FormHeader(sTitulo)
	sAccionq = "COESTDXESQ.asp"
	Celdaz   = ofv.BotonVolver(sAccionq)
	xxx= OFV.botoneraESSX("0", "", "", "", "", "", "", "", "", "no", "no",  Celdaz,"no")

	sHTML =sHTML & xxx
	xs = "<script language=JavaScript>document.getElementById('btnas').style.display="&Chr(34)&"none"&Chr(34)&";" & Chr(13)
	xs = xs & "function AbrirCons(Datos) {" & Chr(13)
	xs = xs & " myWin= open(Datos"
	xs = xs & "," & Chr(34) & "Cons" & Chr(34) & "," & Chr(34)
	xs = xs & "menubar=no,toolbar=no,height=570,width=790" & Chr(34)
	xs = xs & " ); " & Chr(13) & " } " & Chr(13) & "</script>" & Chr(13)

	sHTML = sHTML & xs

	'' sHTML = sHTML & "<center>"

	Celda = OFV.GenCelda("", "ttt colspan=7", "", "", "", "", "", sTitulo, "si")

	Fila = OFV.GenRow("", "", "", "", "25", "", "", Celda, "si")


	Celda = OFV.GenCelda("", "ttr align=center", "", "3%", "", "", "", "*", "si")
	'Celda = Celda & OFV.GenCelda("", "ttr colspan=2", "", "55%", "", "", "", "ESQUEMA NORMATIVO", "si")
	If SESSION("BASE")="K" Then
		Celda = Celda & OFV.GenCelda("", "ttr colspan=2", "", "55%", "", "", "", "FORMULARIO", "si")
	Else
		Celda = Celda & OFV.GenCelda("", "ttr colspan=2", "", "55%", "", "", "", "ESQUEMA NORMATIVO", "si")
	End If
	Celda = Celda & OFV.GenCelda("", "ttr", "", "20%", "", "", "", "TIPO", "si")
	Celda = Celda & OFV.GenCelda("", "ttr", "", "10%", "", "", "", "Estado", "si")
	Celda = Celda & OFV.GenCelda("", "ttr", "", "10%", "", "", "", "VIGENCIA", "si")
	Celda = Celda & OFV.GenCelda("", "ttr COLSPAN=2 ", "CENTER", "30%", "", "", "", "ACCESOS", "si")

	Fila = fila & OFV.GenRow("", "", "", "", "25", "", "", Celda, "si")


	If SEXPORT <> "XLS" Then




		Response.write   sHTML
		sHTML = ""
		Response.flush

	End If

	If SEXPORT = "XLS" Then
		SHTMLEXPORT=SHTMLEXPORT & "<TR><TD>" & "*" & "</TD>"
		SHTMLEXPORT=SHTMLEXPORT & "<TD><U>" & "ESQUEMA NORMATIVO" & "</U></TD>"
		SHTMLEXPORT=SHTMLEXPORT & "<TD><U>" & "TIPO" & "</U></TD>"
		SHTMLEXPORT=SHTMLEXPORT & "<TD><U>" & "ESTADO" & "</U></TD>"
		SHTMLEXPORT=SHTMLEXPORT & "<TD><U>" & "VIGENCIA" & "</U></TD>"
		SHTMLEXPORT=SHTMLEXPORT & "<TD><U>" & "ACCESOS" & "</U></TD>"
		SHTMLEXPORT=SHTMLEXPORT & "<TD><U>" & "FECHA DESDE" & "</U></TD>"
		SHTMLEXPORT=SHTMLEXPORT & "<TD><U>" & "FECHA HASTA" & "</U></TD></TR>"
		SHTMLEXPORT=SHTMLEXPORT & "<TD></TD></TR>"
	End If

	nLinea = 0
	' response.write
	fi = CLng(CDate("01-01-1900"))
	PosReg = 0
	'Do Until rs.EOF
	' On Error Resume Next

	''   If Not rs.EOF Then rs.Move (CantRegMover)
	''   Do While (Not rs.EOF) And (PosReg < CantRegMostrar)

	Naccesos=0
	xcantar=0
	Do While (Not rs.EOF)
		If  SESELESTADOS= "TODOS" Then
			sidz=  BuscarIds(rs(0))
		Else
			sidz="(" & rs(0) & ")"
		End If

		strsql =   "SELECT A.CODOBJETO, A.TIPOBJETO, A.FULTACC, A.ACON, AVG(ISNULL(PROM,0)) AS PROMF, SUM(ISNULL(CANT,0)) AS CANTF, "
		strsql = strsql & " MAX(ESTADO) AS ESTADO , MAX(VERSION) AS VERSION, "
		strsql = strsql & " MAX(NROVER) AS NROVER, MAX(FECVIGEN) AS FECVIGEN, "
		strsql = strsql & " MAX(CODIDEN) AS CODIDEN, MAX(C.DESCRI) AS DESCRI  "
		strsql = strsql & " FROM OBJALOG A LEFT JOIN "
		If session("UN") = 1 Then
			strsql = strsql & " (select CODOBJETO, TIPOBJETO, FINICIO, AVG(HFINAL - HINICIO) AS PROM, SUM(1) AS CANT "
			strsql = strsql & " from objalogdet "
			strsql = strsql & " where tacc = 'CONSEN' "
			strsql = strsql & " GROUP BY CODOBJETO,TIPOBJETO,FINICIO) B "
		Else
			strsql = strsql & " (select objalogdet.CODOBJETO, objalogdet.TIPOBJETO, objalogdet.FINICIO, AVG(HFINAL - HINICIO) AS PROM, SUM(1) AS CANT "
			strsql = strsql & " from objalogdet "
			strsql = strsql & " inner join objalog c on c.tipobjeto='SKRLOG' and c.usuario=objalogdet.usuario and c.codemp=" & session("UN")
			strsql = strsql & " where tacc = 'CONSEN' "
			strsql = strsql & " GROUP BY objalogdet.CODOBJETO,objalogdet.TIPOBJETO,objalogdet.FINICIO) B "
		End If

		strsql = strsql & " ON A.CODOBJETO = B.CODOBJETO AND A.TIPOBJETO = B.TIPOBJETO "
		strsql = strsql & " INNER JOIN OBJETOS C ON A.CODOBJETO = C.CODOBJETO  "
		strsql = strsql & " INNER JOIN OBJPRTY D ON A.CODOBJETO = D.CODOBJETO  "
		If SESSION("BASE")="K" Then
			strsql = strsql & " INNER JOIN tipobjeto Df ON C.tipobjeto = Df.tipobjeto and df.TIPOCPT='K'  "
		Else
			strsql = strsql & " INNER JOIN tipobjeto Df ON C.tipobjeto = Df.tipobjeto and df.TIPOCPT <> 'K'  "
		End If
		If  SESELESTADOS <> "TODOS" Then
			strsql = strsql & "WHERE A.CODOBJETO > 0  AND ESTADO in ("& SESELESTADOS &")  " & XWHERE
		Else
			strsql = strsql & "WHERE A.CODOBJETO > 0  AND ESTADO in (3,4)  " & XWHERE
		End If
		strsql=  strsql & " and A.CODOBJETO in " & sidz
		strsql = strsql & " GROUP BY A.CODOBJETO,A.TIPOBJETO,FULTACC, ACON "
		strsql = strsql & XORDER

		'RESPONSE.WRITE STRSQL & "<BR>" & CHR(13)
		Set rs2 = OFV.crearconsultaEx(strsql, cnx, 1, "")
		miTabla = rs2.GetRows

		Call OFV.CERRARCONsULTA(RS2)



		' Do While (Not rs2.EOF)

		If   UBound(miTabla,2) > 0 Then
			Celda =OFV.GenCelda("", "UTS COLSPAN=7 ", "", "", "", "", "", "<hr>", "si")
			Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
		End If
		xcantaacc=0
		For I = 0 To UBound(miTabla,2)




			For J = 0 To UBound(miTabla, 1)




				'sImagen = "<font face=corflaticon size=4 color=#993333>2</font>"
				sImagen = "<i class='las la-plus'></i>"

				dVigencia = "SD"
				If Not IsNull(miTabla(9,i)) Then dVigencia = fechadorcor( CDate(CLng(miTabla(9,i)) + CLng(CDate("01-01-1900"))),"")
				dVign = InStr(dVigencia, "1900")
				If dVign > 0 Then dVigencia = "SD"

				sTipobjeto = miTabla(1,I)
				sCODobjeto = miTabla(0,I)
				sobjeto    = miTabla(0,I)

				Strsql  = "Select * From Tipobjeto Where tipobjeto='" & sTipobjeto & "'"
				Set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
				If  Not RSX.EOF Then sTipobjeto = RSX(1)

				sobjeto = miTabla(11,I) & "-" & miTabla(10,I)

				If  miTabla(7,I) = "O" Then
					sobjeto = sobjeto & "&nbsp;(O)"
				Else
					sobjeto = sobjeto & "&nbsp;(V&nbsp;" & miTabla(8,I) & ")"
				End If

				sActcon  = "CoESTDLISDET.asp?Codigo=" & miTabla(0,I) & "&Tipo=" & miTabla(1,I)
				sActcon  = sActcon & "&FDES=" & XFDES & "&FHAS=" & XFHAS
				sHref    = "&nbsp;"
				sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "'); return false;" & Chr(34)
				'  response.write sactcon

				Celda  = ""
				sLink  = OFV.GenLink("", "UT", sHref, "", "", sFuncion, sImagen, "si")
				Celda  = Celda & OFV.GenCelda("", "UTS COLSPAN=2 ", "", "", "", "", "", sLink, "si")

                fcodiden=""
                if SESSION("BASE")="K" then
				 sql = "select codigoform from objpropiedades where codobjeto=" & miTabla(0,I)
				 Set rg = ofv.crearconsultaex(sql,cn,1,parametros)
				  If Not rg.eof  Then
					fcodiden=rg(0)   & "&nbsp;-&nbsp;"
				   Else
					fcodiden=""
				   End If
				    Call ofv.cerrarconsulta (rg)
				  end if
				sLink  = OFV.GenLink("", "Ut", sHref, "", "", sFuncion, fcodiden & sobjeto, "si")
				Celda  = Celda & OFV.GenCelda("", "UT", "", "", "", "", "", sLink, "si")

				Celda  = Celda & OFV.GenCelda("", "UT ", "", "", "", "", "",sTipobjeto, "si")

				Select  Case miTabla(6,I)&""

					Case  "3" estado="VIGENTE"
					Case "4"  estado=" NO VIGENTE"
					Case Else estado="&nbsp;"
				End Select



				Celda  = Celda & OFV.GenCelda("", "UT ", "", "", "", "", "",estado, "si")

				Celda  = Celda & OFV.GenCelda("", "UT ", "", "", "", "", "", dVigencia, "si")




				Celda  = Celda & OFV.GenCelda("", "UT ", "RIGHT", "", "", "", "", miTabla(5,I) & "&nbsp;", "si")
				' Celda  = Celda & OFV.GenCelda("", "UT ", "RIGHT", "", "", "", "", fechadorcor(XFDES), "si")
				' Celda  = Celda & OFV.GenCelda("", "UT ", "RIGHT", "", "", "", "", fechadorcor(XFHAS), "si")





				Celda  = Celda & OFV.GenCelda("", "UT ", "", "", "", "", "", sAutor, "si")

			Next

			Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
			If SEXPORT = "XLS" Then
				SHTMLEXPORT=SHTMLEXPORT & "<TR><TD>" & POSREG+1 & "</TD>"
				SHTMLEXPORT=SHTMLEXPORT & "<TD>" & fcodiden & sobjeto & "</TD>"
				SHTMLEXPORT=SHTMLEXPORT & "<TD>"  & sTipobjeto & "</TD>"
				SHTMLEXPORT=SHTMLEXPORT & "<TD>"  & estado & "</TD>"

				SHTMLEXPORT=SHTMLEXPORT &  "<TD>" & dVigencia & "</TD>"
				SHTMLEXPORT=SHTMLEXPORT &  "<TD>"  & miTabla(5,I) & "</TD>"
				SHTMLEXPORT=SHTMLEXPORT &  "<TD>" & fechadorcor(CDate(XFDES),"") & "</TD>"
				SHTMLEXPORT=SHTMLEXPORT &  "<TD>" & fechadorcor(CDate(XFHAS),"") & "</TD></TR>"
			End If
			' Fila = Fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")





			naccesos=naccesos + CLng("0" &miTabla(5,I) )

			xcantaacc=xcantaacc + CLng("0" &miTabla(5,I) )
			PosReg = PosReg + 1

			nLinea = nLinea + 1
			If SEXPORT <> "XLS" Then
				Response.write sHTML
			End If
			sHTML = ""
			If nLinea > 20 Then
				nLinea = 0
				If SEXPORT <> "XLS" Then
					Response.flush
				End If
			End If


		Next
		If   UBound(miTabla,2) > 0 Then


			Celda = ofv.GenCelda("","ttt  colspan=6 ","","","","","", "TOTAL " ,"si")
			Celda = CELDA & 	ofv.GenCelda("","ttt  ALIGN=RIGHT ","","","","","", xcantaacc ,"si")
			Fila =Fila & ofv.GenRow("","","","","30","","",Celda,"si")


		End If

		' call ofv.cerrarconsulta(rs2)
		rs.movenext
	Loop
	If PosReg > 0 Then


		Celda = ofv.GenCelda("","ttt  colspan=5 ","","","","",""," TOTAL: " & posreg & " DOCUMENTOS             " & sInput,"si")
		Celda = celda & ofv.GenCelda("","ttt colspan=2 align=center","","","","",""," TOTAL ACCESOS&nbsp;:&nbsp;&nbsp;" & naccesos ,"si")

	Else


		Celda = ofv.GenCelda("","ttt  colspan=7 ","","","","",""," NO HAY DOCUMENTOS PARA LA PRESENTE SELECCION ","si")
	End If

	Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

	XTABLA =  ofv.GenTabla("","aaa","","","","0","","0","2","","",Fila,"si") & "<br>"


	'  Formulario = XTABLA

	sHTML = sHTML & XTABLA


	sHTML = sHTML & "</body></html><script language=JavaScript>document.getElementById('btnas').style.display="&Chr(34)&""&Chr(34)&";document.getElementById('btnas').style.left="&Chr(34)&"80%"&Chr(34)&";</script>"

	session("sperfil")=""


	Call OFV.cerrarconsulta(rs)
	Call OFV.cerrarconn(cn)
	Call OFV.cerrarconn(cnx)

	''     Generar = sHTML


	If SEXPORT ="XLS" Then
		Response.ContentType = "application/vnd.ms-excel"
		Response.AddHeader "content-disposition", " filename=CoEstdlis.xls"

		response.write SHTMLEXPORT & "</TABLE>"
	Else
		response.write sHTML
		response.flush

		'   sAccionq = "COESTDXESQ.asp"
		'   Celdaz   = ofv.BotonVolver(sAccionq)
		'    xxx= OFV.botoneraESSX("0", "", "", "", "", "", "", "", "", "no", "no",  Celdaz,"no")

		'   response.write xxx


	End If


End If




Function Fechador(wkfecha)



	fec = CDate(wkfecha + CLng(CDate("01-01-1900")))
	ndia = ofv.cdia(Weekday(fec))
	nMes = ofv.cmes(Month(fec))

	Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)


End Function


Function BuscarIds(id)
	Dim strsql,rsid

	Dim sids


	sids=id
	delim=","
	strsql="select objrem from objetos where codobjeto=0" &id
	'response.write strsql
	Set rsid = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If Not rsid.eof Then
		Orem= rsid(0)
		sids =sids & delim & Orem
	End If
	Call ofv.cerrarconsulta(rsid)

	strsql="select objrem  from objetos a inner join objprty b "
	strsql= strsql & " on ( a.codobjeto=b.codobjeto )"
	strsql= strsql & " where a.codobjeto=0" & Orem

	' response.write "<br>" & strsql &"<BR>"
	Set rsid = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If  Not rsid.eof Then
		Do While Not rsid.eof
			Orem=rsid(0)


			If (Orem=0) Then Exit Do
			sids =sids & delim & Orem
			strsql="select objrem  from objetos a inner join objprty b "
			strsql= strsql & " on ( a.codobjeto=b.codobjeto )"
			strsql= strsql & " where a.codobjeto=0" & Orem
			Call ofv.cerrarconsulta(rsid)
			Set rsid = ofv.crearconsultaEx(StrSql,cn,1,parametros)

		Loop
	End If
	BuscarIds="(" & sids & ")"
End Function


Function SPEXE(CN)

	Set objCommandSec = SERVER.CreateObject("ADODB.Command")
	objCommandSec.ActiveConnection = Conn
	objCommandSec.CommandType = 4
	objCommandSec.CommandText = "PRC_EXECSQLex "

	objCommandSec.Parameters.Refresh

	objCommandSec.parameters.append objCommandSec.createParameter("@inVar1", adVarChar, adParamInput, 255, "someValue1")
	objCommandSec.parameters.append objCommandSec.createParameter("@inVar2", adVarChar, adParamInput, 255, "someValue2")
	objCommandSec.parameters.append objCommandSec.createParameter("@outVar1", adVarChar, adParamReturnValue)
	objCommandSec.parameters.append objCommandSec.createParameter("@outVar2", adVarChar, adParamReturnValue)
	objCommandSec.Execute , , adExecuteNoRecords
	outVar1 = objCommandSec.parameters("@outVar1").value
	outVar2 = objCommandSec.parameters("@outVar2").value


End Function



%>