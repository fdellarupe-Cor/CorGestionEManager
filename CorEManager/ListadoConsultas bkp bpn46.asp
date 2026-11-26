<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	'response.write Session("Form")


	'''''''''''''''''''response.write Session("Form")

	'for each x in request.form
	'   response.write "-- " & x & " : " & request.form(x) & " --<br>" & chr(13)
	'next
	Generar = ""

	Addqry = ""
	Response.expires = 0

	Server.scripttimeout = 600
	Response.buffer = True

	' response.write Session("Form")
	XXXX    = NEWEMPRE("NO")
	MasterType = "SCR"

	Mtype = request.querystring("MType") & ""
	If Mtype <> "" Then MasterType = Mtype & ""

	smenu   = Request.querystring("Menu")
	sorden  = Request.querystring("sorden")
	sNombre = request.querystring("Nombre")

	Set cn = OFV.conectar(OFV.strconn2)
	Set cnx = OFV.conectar(OFV.strconn5)
	'response.write smenu & " CCCCCC smenu " & chr(13)


	'response.write "STRING: " & request.querystring & "<br>" & "FORM: " &  request.form

	If Session("CorApli") = "COREM20" Then
		wkaplic = "EM"
	ElseIf Session("CorApli") = "COREP20" Then
		wkaplic = "PL"
	Else
		wkaplic = "OR"
	End If


	If Len(sorden) = 0 Then
		sorden = "N"
	End If

	Sfiltro = Request.form("sfiltro")

	If Len(Sfiltro) = 0 Then
		Sfiltro = Request.querystring("sfiltro")
	End If

	If Len(Sfiltro) = 0 Then
		Sfiltro = "TODOS"
	End If

	Sfiltro = Replace(Sfiltro, "_", " ")
	CantRegMostrar = 10
	CantRegMover = CDbl("0" &Request.querystring("CantRegMover"))
	ReDim Preserve bconsultar(4)
	OFV.ObtenerAtributosNodo Session("Form"), "", "ESTADO", "4"
	bconsultar(4) = OFV.AConsultar
	OFV.ObtenerAtributosNodo Session("Form"), "", "ESTADO", "2"
	bconsultar(2) = OFV.AConsultar
	OFV.ObtenerAtributosNodo Session("Form"), "", "ESTADO", "1"
	bconsultar(1) = OFV.AConsultar
	OFV.ObtenerAtributosNodo Session("Form"), "", "ESTADO", "3"
	bconsultar(3) = OFV.AConsultar
	SQLESTADOS=""
	If smenu = "PORFECHA" Or smenu = "PORREVISION" Then
		sverestado = " and csporfechas.objnivel >= " & Session("SegNivel") & " And Estcod In (9,"
	ElseIf  smenu = "PORSECTOR" Then
		SQLESTADOS= " AND c.estado In(9,"
		SQLNIVEL=" x.objnivel >="  &  Session("SegNivel")

	ElseIf smenu = "PORLINEGO" Then
		sverestado = " and objetos.objnivel >= " & Session("SegNivel") & " And objprty.estado In (9,"
	Else
		'  SQLNIVEL=" x.objnivel >="  &  Session("SegNivel")
		sverestado = " and Objetos.objnivel >= " & Session("SegNivel") & " And Estados.objid In (9,"
	End If

	'if smenu <> "PORLINEGO" then

	For i = 1 To 4
		If bconsultar(i) Then
			sverestado = sverestado & i & ","
			SQLESTADOS = SQLESTADOS & i & ","
			If i = 1 Then sverestado = sverestado & "6" & ","
			If i = 2 Then sverestado = sverestado & "5" & ","

			If i = 1 Then SQLESTADOS = SQLESTADOS & "6" & ","
			If i = 2 Then SQLESTADOS = SQLESTADOS & "5" & ","

		End If
	Next
	sverestado = Left(sverestado, Len(sverestado) - 1) & ") "
	SQLESTADOS = Left(SQLESTADOS, Len(SQLESTADOS) - 1) & ") "
	'RESPONSE.WRITE sverestado &"<BR>"
	'RESPONSE.WRITE SQLESTADOS &"<BR>"
	''''''     FILTRO POR EMPRESAç
	'end if

	addfrom =  ""

	''response.write smenu & " smenu " & chr(13)



	If smenu = "PORFECHA" Or smenu = "PORREVISION" Or smenu = "PORACTA" Then


		If (SESSION("EMPMUN") = "S" And SESSION("GLBMUN") = 1) Or (SESSION("EMPMUN") = "N") Then
			addfrom =  " ,objetos , tipobjeto  "
			sverestado = sverestado & " and csporfechas.CODOBJETO = objetos.CODOBJETO "
			sverestado = sverestado & " and csporfechas.tipoBJETO = tipobjeto.tipobjeto "
			sverestado = sverestado & " and objetos.OBJUN = " & Session("UN") & " "
		Else
			addfrom =  " ,OBJMEMPALC ,objetos , tipobjeto "
			sverestado = sverestado & " and csporfechas.CODOBJETO = objetos.CODOBJETO "
			sverestado = sverestado & " and csporfechas.tipoBJETO = tipobjeto.tipobjeto "
			sverestado = sverestado & " and "
			sverestado = sverestado & " ( (objetos.OBJUN = " & Session("UN") & ") OR "
			sverestado = sverestado & "   ( objetos.OBJUN = 1 AND "
			sverestado = sverestado & "     csporfechas.CODOBJETO = OBJMEMPALC.CODOBJETO AND "
			sverestado = sverestado & "     OBJMEMPALC.CODEMPRE = " & Session("UN") & ") "
			sverestado = sverestado & " ) "
		End If

		sverestado = sverestado & " and TIPOBJETO.TIPOCPT = 'D' "
		sverestado = sverestado & " and (objetos.APLIC = '" & wkaplic & "' or objetos.APLIC = 'GLB') "

	Else

		SQLSECTORUN=""
		If (SESSION("EMPMUN") = "S" And SESSION("GLBMUN") = 1) Or (SESSION("EMPMUN") = "N") Then
			sverestado = sverestado & " and Objetos.OBJUN = " & Session("UN") & " "
			SQLSECTORUN =" x.OBJUN=" & Session("UN")
		Else
			addfrom    =  " ,OBJMEMPALC "
			sverestado = sverestado & " and "
			sverestado = sverestado & " ( (OBJETOS.OBJUN = " & Session("UN") & ") OR "
			sverestado = sverestado & "   ( OBJETOS.OBJUN = 1 AND "
			sverestado = sverestado & "     OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO AND "
			sverestado = sverestado & "     OBJMEMPALC.CODEMPRE = " & Session("UN") & ") "
			sverestado = sverestado & " ) "
			SQLSECTORUN= " ( (x.OBJUN=" & Session("UN") & ") OR "
			SQLSECTORUN=SQLSECTORUN & "( X.OBJUN = 1 AND "
			SQLSECTORUN = SQLSECTORUN & "     X.CODOBJETO IN (SELECT DISTINCT CODOBJETO FROM OBJMEMPALC WHERE CODEMPRE= "  &Session("UN")& ")))"
			'  RESPONSE.WRITE SQLSECTORUN & "<BR>"
		End If
		SQLTIPOCPT="e.TIPOCPT = 'D'"
		SQLAPLIC=" (x.APLIC = '" & wkaplic &"' or x.APLIC = 'GLB')"
		sverestado = sverestado & " and TIPOBJETO.TIPOCPT = 'D' "
		sverestado = sverestado & " and (OBJETOS.APLIC = '" & wkaplic & "' or OBJETOS.APLIC = 'GLB') "
	End If



	'''''''''''
	' If Request.Form.Count > 1 Then
	'    Robj = "F"
	' else
	'    Robj = "Q"
	' end if
	'response.write request.querystring("scal")
	scal=request.querystring("scal")
	If Request.Form.Count > 1  Then
		Robj = "F"
	Else
		Robj = "Q"
	End If


	'   response.write Request.Form.Count
	scal=request.querystring("scal")
	'response.write smenu
	If smenu = "PORACTA"  Then

		If  Robj = "F"   Then

			sID1  = Request.Form("Desde3") & ""
			'   sID1  = CLng(CDate(sID1)) - CLng(CDate("01-01-1900"))

			'   response.write request.form
			sID2  = Request.Form("Hasta3") & ""
			'   sID2  = CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))

			Addqry = "&Desde=" & sID1 & "&Hasta=" & sID2
		Else
			sID1   = Request.querystring("Desde") & ""
			sID2   = Request.querystring("Hasta") & ""
			Addqry = "&Desde=" & sID1 & "&Hasta=" & sID2
		End If

		sMenuAdd    = ""
		sMenuMas    = "&smenu=" & smenu & "&Desde=" & sID1 & "&Hasta="
		sMenuMas    = sMenuMas & sID2

		xsid        = Replace(sID1, "/", "")
		xsid        = Replace(xsid, "-", "")
		If (Not IsNumeric(xsid)) Or (sID1 = "") Then  sID1  = "1-1-1950"

		xsid        = Replace(sID2, "/", "")
		xsid        = Replace(xsid, "-", "")
		If (Not IsNumeric(xsid)) Or (sID2 = "") Then sID2 = "31-12-2100"

	ElseIf smenu = "PORREVISION" Then

		If  Robj = "F"   Then

			sID1   = Request.Form("Desde3") & ""

			sID2   = Request.Form("Hasta3") & ""
			Addqry = "&Desde=" & sID1 & "&Hasta=" & sID2
		Else
			sID1   = Request.querystring("Desde") & ""
			sID2   = Request.querystring("Hasta") & ""
			Addqry = "&Desde=" & sID1 & "&Hasta=" & sID2
		End If

		sMenuAdd    = ""
		sMenuMas    = "&smenu=" & smenu & "&Desde=" & sID1 & "&Hasta="
		sMenuMas    = sMenuMas & sID2

		xsid        = Replace(sID1, "/", "")
		xsid        = Replace(xsid, "-", "")
		If (Not IsNumeric(xsid)) Or (sID1  = "") Then  sID1  = "1-1-1950"

		xsid        = Replace(sID2, "/", "")
		xsid        = Replace(xsid, "-", "")
		If (Not IsNumeric(xsid)) Or (sID2 = "") Then sID2 = "31-12-2100"

	ElseIf smenu = "PORFECHA" Then

		If  Robj = "F"  Then

			sID1   = Request.Form("Desde3") & ""
			sID2   = Request.Form("Hasta3") & ""
			sID3   = Request.Form("Fecha") & ""
			Addqry = "&Desde=" & sID1 & "&Hasta=" & sID2 & "&Fecha=" & sID3
		Else
			sID1   = Request.querystring("Desde") & ""
			sID2   = Request.querystring("Hasta") & ""
			sID3   = Request.querystring("Fecha") & ""
			Addqry = "&Desde=" & sID1 & "&Hasta=" & sID2 & "&Fecha=" & sID3
		End If

		sMenuAdd    = ""
		sMenuMas    = "&smenu=" & smenu & "&Desde=" & sID1 & "&Hasta="
		sMenuMas    = sMenuMas & sID2 & "&Fecha=" & sID3

		xsid        = Replace(sID1, "/", "")
		xsid        = Replace(xsid, "-", "")
		If (Not IsNumeric(xsid)) Or (sID1 = "") Then  sID1  = "1-1-1950"

		xsid        = Replace(sID2, "/", "")
		xsid        = Replace(xsid, "-", "")
		If (Not IsNumeric(xsid)) Or (sID2 = "") Then sID2 = "31-12-2100"

	ElseIf smenu = "PORPRODUCTO" Then

		If  Robj = "F" Then

			sID1   = Request.Form("ListaS") & ""
			sID2   = Request.Form("ListaG") & ""
			sID3   = Request.Form("ListaP") & ""
			Addqry = "&ListaS=" & sID1 & "&ListaG=" & sID2 & "&ListaP=" & sID3
			sID1   = Replace(sID1, "_", " ")
			sID2   = Replace(sID2, "_", " ")
			sID3   = Replace(sID3, "_", " ")

		Else
			sID1  = Request.querystring("ListaS") & ""
			sID2   = Request.querystring("ListaG") & ""
			sID3   = Request.querystring("ListaP") & ""
			Addqry = "&ListaS=" & sID1 & "&ListaG=" & sID2 & "&ListaP=" & sID3
			sID1    = Replace(sID1, "_", " ")
			sID2   = Replace(sID2, "_", " ")
			sID3   = Replace(sID3, "_", " ")

		End If

	ElseIf smenu = "PORNROCIRCULAR" Then

		If  Robj = "F" Then

			sID1   = Request.Form("ListaG") & ""
			sID2   = Request.Form("ListaN") & ""
			Addqry = "&ListaG=" & sID1 & "&ListaN=" & sID2
			sID1   = Replace(sID1, "_", " ")
			sID2   = Replace(sID2, "_", " ")
		Else

			sID1   = Request.querystring("ListaG") & ""
			sID2   = Request.querystring("ListaN") & ""
			Addqry = "&ListaG=" & sID1 & "&ListaN=" & sID2
			sID1   = Replace(sID1, "_", " ")
			sID2   = Replace(sID2, "_", " ")

		End If

	ElseIf smenu = "PORINSTRUMENTO" Then

		If  Robj = "F" Then

			sID1   = Request.Form("Lista") & ""
			sID2   = Request.Form("buscar") & ""
			Addqry = "&Lista=" & sID1 & "&buscar=" & sID2
			sID2   = Replace(sID2, "_", " ")

		Else

			sID1   = Request.querystring("Lista") & ""
			sID2   = Request.querystring("buscar") & ""
			Addqry = "&Lista=" & sID1 & "&buscar=" & sID2
			sID2   = Replace(sID2, "_", " ")

		End If


	ElseIf smenu = "PORTEXTO" Then


		If  Robj = "F" Then


			sID1   = Replace(Request.Form("buscar") & "", "'", "")
			sID1   = Replace(sID1, Chr(34), "")

			sID2   = Request.Form("lista") & ""

			tsel   = Request.Form("tsel") & ""
			tgrp   = Request.Form("tgrp") & ""
			tcod   = Request.Form("tcod") & ""
			ttip   = Request.Form("ttip") & ""
			tnam   = Request.Form("tnam") & ""

			tdoc   = Request.Form("tdoc") & ""


			Addqry = "&buscar=" & Replace(sID1," ","_")

			Addqry = Addqry & "&Lista=" & sID2

			Addqry = Addqry & "&tsel=" & tsel & "&tgrp=" & tgrp & "&tcod=" & tcod & "&ttip=" & ttip & "&tnam=" & tnam

			Addqry = Addqry & "&tdoc=" & tdoc

			sID1    = Replace(sID1, "_", " ")

		Else

			sID1  = Request.querystring("buscar") & ""

			sID2   = Request.querystring("Lista") & ""

			tsel   = Request.querystring("tsel") & ""
			tgrp   = Request.querystring("tgrp") & ""
			tcod   = Request.querystring("tcod") & ""
			ttip   = Request.querystring("ttip") & ""
			tnam   = Request.querystring("tnam") & ""

			tdoc   = Request.querystring("tdoc") & ""

			Addqry = "&buscar=" & sID1

			Addqry = Addqry & "&Lista=" & sID2

			Addqry = Addqry & "&tsel=" & tsel & "&tgrp=" & tgrp & "&tcod=" & tcod & "&ttip=" & ttip & "&tnam=" & tnam

			Addqry = Addqry & "&tdoc=" & tdoc

			sID1    = Replace(sID1, "_", " ")

		End If



		' sMenuAdd = ""
		' sMenuMas = "&smenu=" & smenu & "&Instrumento=" & sID1 & "&buscar="
		' sMenuMas = SmenuMas & sID2

		sMenuAdd = ""
		sMenuMas = "&smenu=" & smenu & "&buscar=" & sID1 & "&Lista="
		sMenuMas = sMenuMas & sID2 & "&tsel=" & tsel & "&tgrp=" & tgrp & "&tcod=" & tcod & "&ttip=" & ttip & "&tnam=" & tnam
		sMenuMas = sMenuMas & "&tdoc=" & tdoc


	ElseIf smenu = "PORCATEGO" Then


		If  Robj = "F" Then


			sID1   = Replace(Request.Form("buscar") & "", "'", "")
			sID1   = Replace(sID1, Chr(34), "")

			sID2   = Request.Form("lista") & ""

			tsel   = Request.Form("tsel") & ""
			tgrp   = Request.Form("tgrp") & ""
			tcod   = Request.Form("tcod") & ""
			ttip   = Request.Form("ttip") & ""
			tnam   = Request.Form("tnam") & ""

			tdoc   = Request.Form("tdoc") & ""


			Addqry = "&buscar=" & Replace(sID1," ","_")

			Addqry = Addqry & "&Lista=" & sID2

			Addqry = Addqry & "&tsel=" & tsel & "&tgrp=" & tgrp & "&tcod=" & tcod & "&ttip=" & ttip & "&tnam=" & tnam

			Addqry = Addqry & "&tdoc=" & tdoc

			sID    = Replace(sID1, "_", " ")

		Else

			sID1   = Request.querystring("buscar") & ""

			sID2   = Request.querystring("Lista") & ""

			tsel   = Request.querystring("tsel") & ""
			tgrp   = Request.querystring("tgrp") & ""
			tcod   = Request.querystring("tcod") & ""
			ttip   = Request.querystring("ttip") & ""
			tnam   = Request.querystring("tnam") & ""

			tdoc   = Request.querystring("tdoc") & ""

			Addqry = "&buscar=" & sID1

			Addqry = Addqry & "&Lista=" & sID2

			Addqry = Addqry & "&tsel=" & tsel & "&tgrp=" & tgrp & "&tcod=" & tcod & "&ttip=" & ttip & "&tnam=" & tnam

			Addqry = Addqry & "&tdoc=" & tdoc

			sID    = Replace(sID1, "_", " ")

		End If



		' sMenuAdd = ""
		' sMenuMas = "&smenu=" & smenu & "&Instrumento=" & sID1 & "&buscar="
		' sMenuMas = sMenuMas & sID2

		sMenuAdd = ""
		sMenuMas = "&smenu=" & smenu & "&buscar=" & sID1 & "&Lista="
		sMenuMas = sMenuMas & sID2 & "&tsel=" & tsel & "&tgrp=" & tgrp & "&tcod=" & tcod & "&ttip=" & ttip & "&tnam=" & tnam
		sMenuMas = sMenuMas & "&tdoc=" & tdoc


	ElseIf smenu = "PORIDENTIFICACION" Then

		If  Robj = "F" Then

			sID1   = Request.Form("buscar") & ""
			Addqry = "&buscar=" & sID1

		Else
			sID1   = Request.querystring("buscar") & ""
			Addqry = "&buscar=" & sID1
		End If

		sMenuAdd = ""
		sMenuMas = "&smenu=" & smenu & "&IDENTIFICACION=" & sID

	ElseIf smenu = "PORESTADO"  Then



		If  Robj = "F"  Then
			sID  = request.form("lista")
			sID1   = Request.Form("Desde3") & ""

			sID2   = Request.Form("Hasta3") & ""
			cdesde = CLng(CDate(sID1)) - CLng(CDate("01-01-1900"))
			chasta = CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))

			Addqry = "&Desde=" & sID1 & "&Hasta=" & sID2 & "&lista=" & sID

			'	 response.write Addqry & " Addqry " & "<br>"

		Else

			sID    = request.querystring("lista")
			sID1   = Request.querystring("Desde") & ""
			sID2   = Request.querystring("Hasta") & ""
			cdesde = CLng(CDate(sID1)) - CLng(CDate("01-01-1900"))
			chasta = CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))

			Addqry = "&Desde=" & sID1 & "&Hasta=" & sID2 & "&lista=" & sID

		End If

		sMenuAdd    = ""
		sMenuMas    = "&smenu=" & smenu & "&Desde=" & sID1 & "&Hasta="
		sMenuMas    =  sMenuMas & sID2  & "&lista=" & sID


	ElseIf   smenu = "PORCOMPONENTE" Then

		If  Robj = "F" Then


			sID2    = Request.Form("lista") & ""
			sID3    = Request.Form("buscar") & ""
			If SID3="" Then SID3=sID2

			Addqry  = "&Lista=" & sID2 & "&buscar=" & sID3 '& "&perfil=" & SPerfil

			SPerfil = Replace(SPerfil, "_", " ")

		Else

			SPerfil = Request.querystring("perfil") & ""
			sID1    = Request.querystring("Lista") & ""
			sID2    = Request.querystring("Lista") & ""
			sID3    = Request.querystring("buscar") & ""
			Addqry  = "&Lista=" & sID2 & "&buscar=" & sID3 '& "&perfil=" & SPerfil
			SPerfil = Replace(SPerfil, "_", " ")

		End If

		sID1        = OFV.convertircar(sID1, "_", " ")
		If sID1     = "" Then sID1 = sID3

		sMenuAdd    = ""
		sMenuMas    = "&smenu=" & smenu & "&lista=" & sID2 & "&buscar="
		sMenuMas    = sMenuMas & sID3 & "&perfil=" & SPerfil

		' response.write sID2  & " lista " & "<br>"
		' response.write sID3  & " buscar " & "<br>"
		'  response.write SPerfil  & " perfil " & "<br>"
	ElseIf   smenu = "PORLINEGO" Then

		If  Robj = "F" Then
			' RESPONSE.WRITE REQUEST.FORM & "  F"
			' SPerfil = Request.Form("perfil") & ""
			sID1    = Request.Form("listaS") & ""
			sID2    = Request.Form("listaG") & ""
			sID3    = Request.Form("LISTAP") & ""
			sID4    = Request.Form("SFILTRO") & ""
			sfil = sID4

			Addqry  = "&ListaS=" & sID1 & "&LISTAG=" & sID2 & "&LISTAP=" & sID3 & "&SFILTRO=" & sID4 '& "&SMENU=" & SMENU

			SPerfil = Replace(SPerfil, "_", " ")

		Else
			'   RESPONSE.WRITE REQUEST.FORM & "  Q"
			' SPerfil = Request.querystring("perfil") & ""
			sID1    = Request.querystring("ListaS") & ""
			sID2    = Request.querystring("ListaG") & ""
			sID3    = Request.querystring("LISTAP") & ""
			sID4    = Request.Form("SFILTRO") & ""
			Addqry  = "&ListaS=" & sID1 & "&LISTAG=" & sID2 & "&LISTAP=" & sID3 & "&SFILTRO=" & sID4  '& "&SMENU=" & SMENU

		End If

		If MasterType <> "EXP" Then

			starget = "ListadoConsultas.asp?cantregmover=0&cantregmostrar=" & CantRegMostrar & Addqry & "&MENU=" & SMENU
			sAccionEXP = starget & "&Mtype=EXP"
			SFORMEXP = Formulario & ofv.GenForm("SUBEXP","",sAccionEXP,"_blank","post","","","si")

			fecHoy = fechadorCor(Now,"")

			sql = "select descri from tipoproducto where objid=" & sid1
			Set rz = OFV.crearconsultaEx(sql,cn,1,parametros)
			prodTit = ""
			If Not rz.eof Then prodTit = rz(0)

			sTitulo = "LINEA DE NEGOCIO:&nbsp;" & prodTit & "&nbsp;&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;&nbsp;FECHA DE CONSULTA:&nbsp;" & fecHoy
			sTitulo = sTitulo & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			sTitulo = sTitulo & "<input type=button onclick='document.all.SUBEXP.submit();' value='EXPORTAR' class=bt >"

		End If

	ElseIf   smenu = "PORRESUMEN" Then

		if Request.Form("Desde3") & "" <> "" and Request.Form("Hasta3") & "" <> "" then

			sID1   = Request.Form("Desde3") & ""
			sID2   = Request.Form("Hasta3") & ""

			cdesde = CLng(CDate(request.form("desde3"))) - CLng(CDate("01-01-1900"))
			chasta = CLng(CDate(request.form("hasta3"))) - CLng(CDate("01-01-1900"))

			Addqry  = "&Desde3=" & sID1 & "&Hasta3=" & sID2 & "&SMENU=" & SMENU

		else

			sID1   = Request.querystring("Desde3") & ""
			sID2   = Request.querystring("Hasta3") & ""

			cdesde = CLng(CDate(request.querystring("desde3"))) - CLng(CDate("01-01-1900"))
			chasta = CLng(CDate(request.querystring("hasta3"))) - CLng(CDate("01-01-1900"))

			Addqry  = "&Desde3=" & sID1 & "&Hasta3=" & sID2 & "&SMENU=" & SMENU

		end if

		If MasterType <> "EXP" Then

			starget = "ListadoConsultas.asp?cantregmover=0&cantregmostrar=" & CantRegMostrar & Addqry & "&MENU=" & SMENU
			sAccionEXP = starget & "&Mtype=EXP"
			SFORMEXP = Formulario & ofv.GenForm("SUBEXP","",sAccionEXP,"_blank","post","","","si")

			fi = CLng(CDate("01-01-1900"))
			fechad = CDate(CLng(cdesde) + fi)
			fDesde = fechadorCor(fechad,"")
			fechah = CDate(CLng(chasta) + fi)
			fHasta = fechadorCor(fechah,"")

			sTitulo = "NORMAS PUBLICADAS&nbsp;&nbsp;&nbsp;&nbsp;Desde:&nbsp;" & fDesde & "&nbsp;&nbsp;|&nbsp;&nbsp;Hasta:&nbsp;" & fHasta
			sTitulo = sTitulo & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
			sTitulo = sTitulo & "<input type=button onclick='document.all.SUBEXP.submit();' value='EXPORTAR' class=bt >"

		End If
		'Addqry  = "&desde=" & cdesde & "&hasta=" & chasta & "&MENU=" & SMENU & "&NOMBRE=" & snombre

	Else


		If  Robj = "F" Then

			SPerfil = Request.Form("perfil") & ""
			sID1    = Request.Form("lista") & ""
			sID2    = Request.Form("lista") & ""
			sID3    = Request.Form("buscar") & ""

			Addqry  = "&Lista=" & sID2 & "&buscar=" & sID3 & "&perfil=" & SPerfil

			SPerfil = Replace(SPerfil, "_", " ")

		Else

			SPerfil = Request.querystring("perfil") & ""
			sID1    = Request.querystring("Lista") & ""
			sID2    = Request.querystring("Lista") & ""
			sID3    = Request.querystring("buscar") & ""
			Addqry  = "&Lista=" & sID2 & "&buscar=" & sID3 & "&perfil=" & SPerfil
			SPerfil = Replace(SPerfil, "_", " ")

		End If

		sID1        = OFV.convertircar(sID1, "_", " ")
		If sID1     = "" Then sID1 = sID3

		sMenuAdd    = ""
		sMenuMas    = "&smenu=" & smenu & "&lista=" & sID2 & "&buscar="
		sMenuMas    = sMenuMas & sID3 & "&perfil=" & SPerfil

		' response.write sID2  & " lista " & "<br>"
		' response.write sID3  & " buscar " & "<br>"
		' response.write SPerfil  & " perfil " & "<br>"

	End If

	Addqry = "&Menu=" & smenu & Addqry & "&Nombre=" & sNombre

	starget = "ListadoConsultas.asp?cantregmover=0&cantregmostrar=" & CantRegMostrar & Addqry

	'response.write Addqry  & "<br>"
	'response.write starget & "<br>"

	sVolver = "ListadoConsultas.asp?Menu=" & smenu & sMenuAdd

	If smenu <> "PORLINEGO" Then
		Addqry = Addqry & "&sorden=" & sorden & "&sfiltro=" & Replace(Sfiltro, " ", "_")
	End If
	'response.write Addqry



	cn.CommandTimeout = 0
	strsql = "SELECT DISTINCT TIPOBJETO.descri, OBJETOS.codiden, OBJETOS.descri, "
	strsql = strsql & "OBJETOS.codobjeto, OBJETOS.tipobjeto, OBJPRTY.fecemis, "
	strsql = strsql & "OBJPRTY.fecvigen, OBJPRTY.fecbaja, Estados.estado as estdescri, "
	strsql = strsql & "OBJETOS.codiden, OBJETOS.descri, TIPOBJETO.archivo, OBJPRTY.autor, "
	strsql = strsql & "OBJPRTY.aprobo, OBJPRTY.elimino, TIPOBJETO.perfil, OBJPRTY.fecvto,"
	strsql = strsql & "Estados.estado as estcod,OBJETOS.objnivel,OBJPRTY.Version,OBJPRTY.nrover,"
	strsql = strsql & "OBJETOS.OBJUN,OBJPRTY.fecpublic "
	Strsql = Strsql & ", t.descri, t.descri, t.descri, t.descri, st.descri   "


	addfrom = addfrom & ", csrelpath a, tema t, prodserv  st "
	''''  addfrom = addfrom & ", objinfo g "

	wherefrom = " and (a.cobj = OBJETOS.codobjeto) "
	wherefrom = wherefrom & " and (a.tema = t.codtema) "
	wherefrom = wherefrom & " and (a.stema = st.codprodserv) "
	''response.write wherefrom

	''RESPONSE.WRITE smenu & " smenu " & chr(13)


	Select Case smenu

		Case "PORLINEGO"

		'prod = request.form("ListaS")
		'proc = request.form("ListaG")
		'subp = request.form("ListaP")
		'tobj = request.form("SFILTRO")


		prod = sid1
		proc = sID2 'request.form("ListaG")
		subp = sID3 'request.form("ListaP")
		'tobj = request.form("SFILTRO")
		tobj = SID4
		'response.write tobj
		'sID1    = Request.querystring("ListaS") & ""
		'ID2    = Request.querystring("ListaG") & ""
		'sID3    = Request.querystring("LISTAP") & ""
		strsql = ""
		strsql = strsql & "select  objetos.*,objprty.fecvigen,TIPOBJETO.PERFIL,TIPOBJETO.DESCRI AS TDESCRI,tipoproducto.descri as tp ,tipoagrupa.descri as ta,tiposegmento.descri as ts from objetos  "
		strsql = strsql & " inner join tipobjeto on (objetos.tipobjeto=tipobjeto.tipobjeto  and tipobjeto.tipocpt='D') "
		strsql = strsql & " inner join tipoproducto on (objetos.tsegmento=tipoproducto.objid)"
		strsql = strsql & " inner join objprty on (objetos.codobjeto = objprty.codobjeto and objetos.tipobjeto = objprty.tipobjeto and objprty.estado=3 )"
		strsql = strsql & " left join objtsegagr  on (objetos.codobjeto = objtsegagr.codobjeto and objetos.tipobjeto = objtsegagr.tipobjeto)"
		strsql = strsql & " left join tipoagrupa on (objtsegagr.tagrupa=tipoagrupa.objid)"
		strsql = strsql & " left join tiposegmento on (objtsegagr.tsegmento=tiposegmento.objid)"

		strsql = strsql & " where  objetos.tsegmento=" & prod
		If proc <> "nsnc" Then
			'Se le agregó el and a la línea de abajo porque faltaba
			strsql=strsql  & " and objtsegagr.tagrupa=" & proc
		End If

		If subp <> "nsnc" Then
			strsql=strsql  & " and objtsegagr.tsegmento=" & subp
		End If

		If tobj <> "" And tobj <> "0" Then
			strsql = strsql & " and objetos.tipobjeto='" & tobj & "'"
		End If



		'strsql = strsql & " and objetos.tsegmento=" & prod

		'strsql = strsql & sverestado

		'	strsql = strsql & " order by objtsegagr.tagrupa,objtsegagr.tsegmento"
		'response.write strsql
		'sql = "select descri from tipoproducto where objid=" & prod
		'Set rz = OFV.crearconsultaEx(sql,cn,1,parametros)

		'prodTit = ""

		'If Not rz.eof Then prodTit = rz(0)

		'fecHoy = fechadorCor(Now,"")

		Case "PORCOMPONENTE"

		'   strsql = strsql & "FROM OBJACCION, COMPLEMENTOS, OBJETOS, TIPOBJETO, OBJPRTY, Estados, objlnk, objestdet  " & addfrom
		'   strsql = strsql & " where   "
		'   strsql = strsql & "OBJETOS.codobjeto = OBJlnk.codobjeto and "
		'   strsql = strsql & "objlnk.ltipo = 'KL' and "
		'   strsql = strsql & "complementos.codcomp = objlnk.codobjetol and "
		'   strsql = strsql & " ((OBJACCION.codobjeto = OBJETOS.codobjeto and OBJACCION.codcomp = objlnk.objid and "
		'   strsql = strsql & "objlnk.iden2 = objaccion.codaccion) or "
		'   strsql = strsql & " (objestdet.codobjeto = OBJETOS.codobjeto and "
		'   strsql = strsql & " objestdet.detimagen = 'J' and "
		'   strsql = strsql & " objestdet.detimagen = objlnk.objid and "
		'   strsql = strsql & " objlnk.iden2 = objestdet.objid) ) and "
		' '  strsql = strsql & " objlnk.iden2 = CONVERT(nvarchar(50), dbo.OBJESTDET.objid)) ) and "
		'
		'   strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and "
		'   strsql = strsql & "OBJPRTY.estado = Estados.objid and "
		'   strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto "
		'   strsql = strsql & " and Complementos.descri = '" & sID1 & "'"

		'		 strsql = strsql & " FROM "
		' 		 strsql = strsql & " objetos "

		' strsql = strsql & " inner join  OBJPRTY on (OBJETOS.codobjeto = OBJPRTY.codobjeto)"
		' strsql = strsql & " inner join  Estados on (OBJPRTY.estado = Estados.objid)"
		' strsql = strsql & " inner join  TIPOBJETO on (OBJETOS.TIPOBJETO = TIPOBJETO.tipobjeto)"
		' strsql = strsql & " inner join  OBJlnk on (objetos.codobjeto=OBJlnk.codobjeto and objlnk.ltipo = 'KL') "
		' strsql = strsql & " inner join  COMPLEMENTOS on (complementos.codcomp = objlnk.codobjetol and Complementos.descri = '"&sid&"')"
		' strsql = strsql & " inner join  objestdet on(objestdet.codobjeto = OBJETOS.codobjeto)"
		' strsql = strsql & " inner join  OBJACCION on"
		' strsql = strsql & " ((OBJACCION.codobjeto = OBJETOS.codobjeto and OBJACCION.codcomp = objlnk.objid "
		' strsql = strsql & " and objlnk.iden2 = objaccion.codaccion) "
		' strsql = strsql & " or (objestdet.codobjeto = OBJETOS.codobjeto and objestdet.detimagen = 'J'"
		' strsql = strsql & " and objestdet.detimagen = objlnk.objid and objlnk.iden2 = objestdet.objid) )"
		' strsql = strsql & " inner join csrelpath a on(a.cobj = OBJETOS.codobjeto) "
		' strsql = strsql & " inner join tema t on (a.tema = t.codtema)"
		' strsql = strsql & " inner join prodserv st  on (a.stema = st.codprodserv)"


		strsql = strsql & " FROM "
		strsql = strsql & " objetos "

		strsql = strsql & " inner join  OBJPRTY      on (OBJETOS.codobjeto = OBJPRTY.codobjeto)"
		strsql = strsql & " inner join  Estados      on (OBJPRTY.estado = Estados.objid)"
		strsql = strsql & " inner join  TIPOBJETO    on (OBJETOS.TIPOBJETO = TIPOBJETO.tipobjeto)"
		strsql = strsql & " inner join  OBJlnk       on (objetos.codobjeto=OBJlnk.codobjeto and objlnk.ltipo = 'KL') "
		strsql = strsql & " inner join  COMPLEMENTOS on (complementos.codcomp = objlnk.codobjetol and Complementos.descri like '"& sid3& "')"

		strsql = strsql & " left join objestdet on (OBJETOS.codobjeto = objestdet.codobjeto and objestdet.dettipo = 'J' and objlnk.objid = objestdet.detimagen) "
		strsql = strsql & " left join OBJACCION on (OBJETOS.codobjeto = OBJACCION.codobjeto and objlnk.objid = OBJACCION.codcomp) "

		strsql = strsql & " inner join csrelpath a on(a.cobj = OBJETOS.codobjeto) "
		strsql = strsql & " inner join tema t on (a.tema = t.codtema)"
		strsql = strsql & " inner join prodserv st  on (a.stema = st.codprodserv)"



		sTitulo = "Componente: " & sID1




		Case "PORESTADO"

		'' response.write Robj & " Robj " & "<br>"

		sID  = request.form("lista")

		If  Robj = "F"  Then

			sID1   = Request.Form("Desde3") & ""

			sID2   = Request.Form("Hasta3") & ""
			cdesde = CLng(CDate(sID1)) - CLng(CDate("01-01-1900"))
			chasta = CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))

			Addqry = "&Menu=PORESTADO&NOMBRE=" & nEstado &"&Desde=" & sID1 & "&Hasta=" & sID2 & "&lista=" & sID

			'	 response.write Addqry & " Addqry " & "<br>"

		Else

			sID    = request.querystring("lista")
			sID1   = Request.querystring("Desde") & ""
			sID2   = Request.querystring("Hasta") & ""
			cdesde = CLng(CDate(sID1)) - CLng(CDate("01-01-1900"))
			chasta = CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))

			Addqry = "&Menu=PORESTADO&NOMBRE=" & nEstado &"&Desde=" & sID1 & "&Hasta=" & sID2 & "&lista=" & sID

		End If

		sMenuAdd    = ""
		sMenuMas    = "&smenu=" & smenu & "&Desde=" & sID1 & "&Hasta="
		sMenuMas    =  sMenuMas & sID2  & "&lista=" & sID
		'	response.write smenu & " smenu " & chr(13)

		ssEstado = ""

		strsql = strsql & "FROM OBJETOS, TIPOBJETO, OBJPRTY, Estados " & addfrom
		strsql = strsql & " where  "
		strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and "
		strsql = strsql & "OBJPRTY.estado = Estados.objid and "
		strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto "
		strsql = strsql & "and OBJPRTY.estado = " & sID
		strsql = strsql & " and (Objprty.fecemis between '" & cdesde & "' and '" & chasta & "' ) "

		sTitulo = "Estado: " & sID
		' response.write strsql & " strsql " & chr(13)
		' response.write sID & " sID " & chr(13)
		sqle = "Select * from estados where objid =  " & sID
		Set rsx = OFV.crearconsultaEx(sqle, cn, 1, "")
		If Not rsx.EOF Then
			sTitulo = "Estado: " & rsx(1)
			ssEstado = rsx(1)
		End If

		swDesde2 = CDate(CLng(cdesde) + CLng(CDate("01-01-1900")) )
		swHasta2 = CDate(CLng(chasta) + CLng(CDate("01-01-1900")) )


		sTitulo = "Por estado:  " & ssEstado &"&nbsp;&nbsp; Desde: "& swDesde2 & "&nbsp;&nbsp;|&nbsp; Hasta:&nbsp;" & swHasta2

		Case "PORTEXTO"

		htwhere = ""

		strsql = strsql & "FROM OBJETOS, TIPOBJETO, OBJPRTY, Estados "

		If sid2 <> "99999" Then

			strsql = strsql & ", eco, OBJSECTOR "

		End If

		If tsel <> "T" And tgrp <> "M" And (tgrp = "T" And tcod <> "99999") Then

			strsql = strsql & ", csrelpath rp "

			If tgrp = "T" Then

				htwhere = " and OBJETOS.codobjeto = rp.cobj and rp.tema = 0" & tcod & " "

			ElseIf tgrp = "S" Then

				htwhere = " and OBJETOS.codobjeto = rp.cobj and rp.stema = 0" & tcod & " "

			Else

				htwhere = " and OBJETOS.codobjeto = rp.cobj and ("
				htwhere = htwhere & " (rp.mc1 = 0" & tcod & " and rp.mt1 = '" & ttip & "' ) or "
				htwhere = htwhere & " (rp.mc2 = 0" & tcod & " and rp.mt2 = '" & ttip & "' ) or "
				htwhere = htwhere & " (rp.mc3 = 0" & tcod & " and rp.mt3 = '" & ttip & "' ) or "
				htwhere = htwhere & " (rp.mc4 = 0" & tcod & " and rp.mt4 = '" & ttip & "' ) or "
				htwhere = htwhere & " (rp.mc5 = 0" & tcod & " and rp.mt5 = '" & ttip & "' ) or "
				htwhere = htwhere & " (rp.mc6 = 0" & tcod & " and rp.mt6 = '" & ttip & "' ) or "
				htwhere = htwhere & " (rp.mc7 = 0" & tcod & " and rp.mt7 = '" & ttip & "' ) or "
				htwhere = htwhere & " (rp.cobj = 0" & tcod & " and rp.tobj = '" & ttip & "' )  "
				htwhere = htwhere & " ) "

			End If



		End If




		strsql = strsql & addfrom
		strsql = strsql & " where   "
		strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and "
		strsql = strsql & "OBJPRTY.estado = Estados.objid and "
		strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto "

		'   response.write sid2
		If sid2 <> "99999" Then

			strsql = strsql & " and ECO.eco = OBJSECTOR.codsector and "
			strsql = strsql & " OBJETOS.codobjeto = OBJSECTOR.codobjeto "
			strsql = strsql & " and eco.eco = '" & sID2 & "'"

		End If

		If tdoc <> "99999" Then

			strsql = strsql & " and OBJETOS.tipobjeto = '" & tdoc & "' "

		End If


		If tsel <> "T" And tgrp <> "M" Then

			strsql = strsql & htwhere

		End If

		nwsid = vertexto(sid1)


		strsql = strsql & " and (OBJETOS.descriabr  COLLATE SQL_Latin1_General_CP1_CI_AI like '%" & nwsID & "%' "
		strsql = strsql & " or OBJETOS.descri   COLLATE SQL_Latin1_General_CP1_CI_AI like '%" & nwsID & "%') "

		'Response.write nwsid
		sTitulo = "Texto: " & sID1

		If sid2 <> "99999" Then

			sqle = "Select descri from eco where eco =0" & sID2
			Set rsx = OFV.crearconsultaEx(sqle, cn, 1, "")
			If Not rsx.eof Then
				sTitulo = sTitulo & "&nbsp;|&nbsp;Sector: " & rsx(0)
			Else
				sTitulo = sTitulo & "&nbsp;|&nbsp;Sector: "
			End If
			Call ofv.cerrarconsulta(rsx)
		End If

		If tsel <> "T" And tgrp <> "M" Then

			sTitulo = sTitulo & "&nbsp;|&nbsp;Carpeta:&nbsp;" & Replace(tnam,"_","&nbsp;")


		End If

		Case "PORCATEGO"

		htwhere = ""

		strsql = strsql & "FROM OBJETOS, TIPOBJETO, OBJPRTY, Estados, OBJRSKREL "





		'''''EVENRSK



		strsql = strsql & addfrom
		strsql = strsql & " where  "
		strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and "
		strsql = strsql & "OBJPRTY.estado = Estados.objid and "
		strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto AND "
		strsql = strsql & "OBJETOS.codobjeto = OBJRSKREL.codobjeto AND "
		strsql = strsql & "OBJRSKREL.EVENRSK = 0" & sID2


		Sqlx = "SELEct descri FROM RSKsituacion WHERE id = 0" & tdoc
		Set rsx = OFV.crearconsultaEx(Sqlx, cn, 1, "")
		If Not rsx.EOF Then situnn = rsx(0) & ""
		Call OFV.cerrarconsulta(rsx)

		Sqlx = "SELEct descri FROM RSKcatego WHERE id = 0" & tcod
		Set rsx = OFV.crearconsultaEx(Sqlx, cn, 1, "")
		If Not rsx.EOF Then catenn = rsx(0) & ""
		Call OFV.cerrarconsulta(rsx)

		Sqlx = "SELEct descri FROM RSKevento WHERE id = 0" & sID2
		Set rsx = OFV.crearconsultaEx(Sqlx, cn, 1, "")
		If Not rsx.EOF Then evennn = rsx(0) & ""
		Call OFV.cerrarconsulta(rsx)

		'Response.write strsql & "<br>"
		sTitulo = ""
		sTitulo = sTitulo & "&nbsp;Evento:&nbsp;" & evennn & "<br>"
		sTitulo = sTitulo & "&nbsp;|&nbsp;Categoria:&nbsp;" & catenn
		sTitulo = sTitulo & "&nbsp;|&nbsp;Situacion:&nbsp;" & situnn







		Case "PORPRODUCTO"

		strsql = strsql & "FROM OBJETOS, TIPOBJETO, OBJPRTY, Estados, OBJTSEGAGR " & addfrom
		strsql = strsql & " where   "
		strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and "
		strsql = strsql & "OBJPRTY.estado = Estados.objid and "
		strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto "
		strsql = strsql & "   "

		If UCase(sID3 & "") <> UCase("nsnc") Then
			strsql = strsql & " AND OBJETOS.TSEGMENTO =" & CDbl("0" &sID3) & " "
		End If

		If UCase(sID1 & "") <> UCase("nsnc") Or UCase(sID2 & "") <> UCase("nsnc") Then
			strsql = strsql & " AND ((OBJETOS.codobjeto = OBJTSEGAGR.codobjeto)  AND ( "
		End If

		If UCase(sID1 & "") <> UCase("nsnc") Then
			strsql = strsql & " (OBJTSEGAGR.TSEGMENTO = " & CDbl("0"& sID1) & " ) "
		End If
		If UCase(sID2 & "") <> UCase("nsnc") Then
			If UCase(sID1 & "") <> UCase("nsnc") Then strsql = strsql & " AND "
			strsql = strsql & " (OBJTSEGAGR.TAGRUPA = " & CDbl("0" & sID2) & " ) "
		End If

		If UCase(sID1 & "") <> UCase("nsnc") Or UCase(sID2 & "") <> UCase("nsnc") Then
			strsql = strsql & " ) )"
		End If


		xinstrum = ""
		If UCase(sID1 & "") <> UCase("nsnc") Then
			Sqlx = "SELECT descri from tiposegmento where objid=" & CDbl("0" & sID1)
			Set rsx = OFV.crearconsultaEx(Sqlx, cn, 1, "")
			If Not rsx.EOF Then xinstrum = rsx(0) & ""
			Call OFV.cerrarconsulta(rsx)
		End If



		If UCase(sID2 & "") <> UCase("nsnc") Then
			Sqlx = "SELECT descri from tipoAGRUPA where objid=" & CDbl("0" & sID1)
			Set rsx = OFV.crearconsultaEx(Sqlx, cn, 1, "")
			If Not rsx.EOF Then xinstrum2 = rsx(0) & ""
			Call OFV.cerrarconsulta(rsx)
		End If
		sTitulo = "SEGMENTO :" & xinstrum & "&nbsp;|&nbsp;GRUPO :" & xinstrum2

		Case "PORNROCIRCULAR"

		strsql = strsql & "FROM OBJETOS, TIPOBJETO, OBJPRTY, Estados, OBJTSEGAGR " & addfrom
		strsql = strsql & " where  "
		strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and "
		strsql = strsql & "OBJPRTY.estado = Estados.objid and "
		strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto "

		If UCase(sID2 & "") <> UCase("nsnc") Then
			strsql = strsql & " AND OBJETOS.codiden = '" & CStr(sID2 & "") & "' "
		End If

		If UCase(sID1 & "") <> UCase("nsnc") Then
			strsql = strsql & " AND ((OBJETOS.codobjeto = OBJTSEGAGR.codobjeto)  AND (OBJTSEGAGR.TAGRUPA = " & CDbl("0" & sID1) & " ) ) "
		End If
		''' Response.write strsql & "<br>"

		If UCase(sID2 & "") <> UCase("nsnc") Then
			Sqlx = "SELECT descri from tipoAGRUPA where objid=" & CDbl("0" & sID1)
			Set rsx = OFV.crearconsultaEx(Sqlx, cn, 1, "")
			If Not rsx.EOF Then xinstrum2 = rsx(0) & ""
			Call OFV.cerrarconsulta(rsx)
		End If
		'' xinstrum2 = ""
		sTitulo = " GRUPO :" & xinstrum2
		'response.write sID1 & " sID1<br><br>"
		'response.write sID2 & " sID2<br><br>"
		'response.write xinstrum2 & " xinstrum2<br><br>"
		'response.write request.form & "<br><br>"
		'response.write request.form("agrm") & "aca"

		Case "PORINSTRUMENTO"



		strsql = strsql & " FROM OBJETOS  "
		strsql = strsql & " INNER JOIN OBJPRTY on "
		strsql = strsql & " OBJETOS.codobjeto = OBJPRTY.codobjeto "
		strsql = strsql & " INNER JOIN TIPOBJETO ON "
		strsql = strsql & " OBJETOS.tipobjeto = TIPOBJETO.tipobjeto "
		strsql = strsql & " INNER JOIN Estados ON "
		strsql = strsql & " OBJPRTY.estado = Estados.objid "
		strsql = strsql & " LEFT JOIN OBJPRTYADV ON "
		strsql = strsql & " OBJPRTY.codobjeto = OBJPRTYADV.codobjeto "

		'     if SESSION("EMPMUN") = "S" AND SESSION("GLBMUN") = 2 THEN
		strsql = strsql & " LEFT join OBJMEMPALC "
		strsql = strsql & " on OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO "
		'     END IF
		strsql = strsql & " inner join  csrelpath a on( a.cobj = OBJETOS.codobjeto)"
		strsql = strsql & " inner join   tema t on( a.tema = t.codtema)"
		strsql = strsql & " inner join   prodserv  st on(a.stema = st.codprodserv)"

		Strsql = strsql & " WHERE  "

		strsql = strsql & " ( "

		If ("0" & sID1)<>"0" Then


			If sID2 <> "" Then
				strsql = strsql & "  ((OBJPRTY.instrumen=" & CDbl("0" &sID1) & ")"
			Else
				strsql = strsql & "  (OBJPRTY.instrumen=" & CDbl("0" &sID1) & " or OBJPRTYADV.instrua=" & CDbl("0" &sID1) & ")  "
			End If

			If sID2 <> "" Then
				strsql = strsql & " AND (OBJPRTY.instid ='" & sID2 & "')) OR "
				strsql = strsql & " (OBJPRTYADV.instrua =" & sID1 & " and OBJPRTYADV.instid ='" & sID2 & "')"
			End If

		Else

			strsql = strsql & "(OBJPRTY.instid ='" & sID2 & "')"

		End If

		strsql = strsql & " ) "

		'' response.write strsql & "strsql"

		xinstrum = ""
		Sqlx = "SELECT descri from instrumen where objid=" & CDbl("0" &sID1)
		Set rsx = OFV.crearconsultaEx(Sqlx, cn, 1, "")
		If Not rsx.EOF Then xinstrum = rsx(0)
		Call OFV.cerrarconsulta(rsx)
		sTitulo = "INSTRUMENTO " & xinstrum & " " & sID2

		Case "PORIDENTIFICACION"

		strsql = strsql & "FROM OBJETOS, TIPOBJETO, OBJPRTY, Estados " & addfrom
		strsql = strsql & " where   "
		strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and "
		strsql = strsql & "OBJPRTY.estado = Estados.objid and "
		strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto "
		strsql = strsql & "and OBJETOS.Codiden = '" & sID1 & "' "
		sTitulo = "IDENTIFICACION " & sID1

		Case "PORSECTOR"
		strsql =   " SELECT dISTINCT e.descri, x.codiden, x.descri, x.codobjeto, x.tipobjeto,"
		strsql = strsql &  "c.fecemis, c.fecvigen, c.fecbaja, c1.estado as estdescri,"
		strsql = strsql & "  x.codiden, x.descri, e.archivo, c.autor, c.aprobo,"
		strsql = strsql & "   c.elimino, e.perfil, c.fecvto,c1.estado as estcod,"
		strsql = strsql & "   x.objnivel,c.Version,c.nrover,x.OBJUN,c.fecpublic ,"
		strsql = strsql & "   d.temadescri, d.temadescri, d.temadescri, d.temadescri, d.stemadescri "

		Select Case SPerfil
			Case "nsnc"

			strsql = strsql & "  FROM "
			strsql = strsql & "    ( "
			strsql = strsql & " 	select codobjeto,tipobjeto from objcontrol "

			strsql = strsql & " 	where codsector='0" & sID1 & "'"


			strsql = strsql & " 	union "

			strsql = strsql & " 	 select a.codobjeto,a.tipobjeto "

			strsql = strsql & " 	 from  "
			strsql = strsql & " 	 OBJPROCED a "
			strsql = strsql & " 	 inner join objsector b on ( b.objid=a.codsector)"
			strsql = strsql & " 	 inner join  tipoperfil c on (c.tipoperfil=a.codperfil)"

			strsql = strsql & " 	 where b.codsector='0" & sID1 &"'"
			strsql = strsql & " 	)  x2 "
			strsql = strsql & "INNER JOIN OBJETOS X "
			strsql = strsql & " on ("

			strsql = strsql & "  (x2.codobjeto=x.codobjeto and x2.tipobjeto=x.tipobjeto) "
			strsql = strsql & "  and "
			strsql = strsql &    SQLAPLIC      '   " (x.APLIC = 'EM' or x.APLIC = 'GLB')"
			strsql = strsql & "  and"

			strsql = strsql & SQLNIVEL '"  x.objnivel >= 3 "
			strsql = strsql & "  and"

			strsql = strsql &  SQLSECTORUN & " )"

			strsql = strsql & " inner join objprty c on (c.codobjeto=x.codobjeto "& SQLESTADOS & ") "
			strsql = strsql & " inner join estados c1 on (c1.objid= c.estado)"
			strsql = strsql & " inner join csrelpathdescriex d on(d.cobj=c.codobjeto)"
			strsql = strsql & " inner join TIPOBJETO e on(e.tipobjeto=x.tipobjeto and e.TIPOCPT = 'D')"

			Case Else
			strsql = strsql & "  FROM "
			strsql = strsql & "    ( "
			strsql = strsql & " 	select codobjeto,tipobjeto from objcontrol "

			strsql = strsql & " 	where codsector= " & sID1 & " AND  codperfil='" & SPERFIL &"'"


			strsql = strsql & " 	union "

			strsql = strsql & " 	 select a.codobjeto,a.tipobjeto "

			strsql = strsql & " 	 from  "
			strsql = strsql & " 	 OBJPROCED a "
			strsql = strsql & " 	 inner join objsector b on ( b.objid=a.codsector)"
			strsql = strsql & " 	 inner join  tipoperfil c on (c.tipoperfil=a.codperfil)"

			strsql = strsql & " 	 where b.codsector=" & sID1 & " AND c.tipoperfil='" & SPERFIL & "'"
			strsql = strsql & " 	)  x2 "
			strsql = strsql & "INNER JOIN OBJETOS X "
			strsql = strsql & " on ("

			strsql = strsql & "  (x2.codobjeto=x.codobjeto and x2.tipobjeto=x.tipobjeto) "
			strsql = strsql & "  and "
			'strsql = strsql & " (x.APLIC = 'EM' or x.APLIC = 'GLB')"
			strsql = strsql &    SQLAPLIC
			strsql = strsql & "  and"

			strsql = strsql & SQLNIVEL '"  x.objnivel >= 3 "
			strsql = strsql & "  and"

			strsql = strsql &  SQLSECTORUN & " )"

			strsql = strsql & " inner join objprty c on (c.codobjeto=x.codobjeto "& SQLESTADOS & ") "
			strsql = strsql & " inner join estados c1 on (c1.objid= c.estado)"
			strsql = strsql & " inner join csrelpathdescriex d on(d.cobj=c.codobjeto)"
			strsql = strsql & " inner join TIPOBJETO e on(e.tipobjeto=x.tipobjeto and e.TIPOCPT = 'D')"
		End Select
		'response.write strsql
		''     if  SPerfil <> "nsnc" then

		'sTitulo = "Sector:&nbsp;" & sID1
		sqle = "Select descri from eco where eco =0" & sID1
		Set rsx = OFV.crearconsultaEx(sqle, cn, 1, "")
		If Not rsx.eof Then
			sTitulo = "Sector:&nbsp;" & rsx(0)
		Else
			sTitulo = "Sector:&nbsp;"
		End If

		''	   Response.write rsx(0) & " descri " & "<br>"
		''	   Response.write sPerfil & " sPerfil " & "<br>"
		''	   Response.write sID1 & " sID1 " & "<br>"

		Call ofv.cerrarconsulta(rsx)

		''     if  SPerfil <> "nsnc" then sTitulo = sTitulo & "&nbsp;|&nbsp;Perfil:&nbsp;" & SPerfil
		If  SPerfil <> "nsnc" Then

			sTitulo = sTitulo & "&nbsp;"

			Strsqle = "Select descri from tipoperfil where Tipoperfil ='" & SPerfil & "' "
			Set rsxx = OFV.crearconsultaEx(Strsqle, cn, 1, "")
			If Not rsxx.eof Then
				sTitulo = sTitulo & "&nbsp;|&nbsp;Perfil:&nbsp;" & rsxx(0)
			Else
				sTitulo = sTitulo & "&nbsp;|&nbsp;Perfil:&nbsp;" & "TODOS"
			End If

			Call ofv.cerrarconsulta(rsxx)

		Else
			sTitulo = sTitulo & "&nbsp;|&nbsp;Perfil:&nbsp;" & "TODOS"
		End If

		Case "PORTEMA"
		sql = "Select descri From Objetos Where codobjeto = " & sID1
		Set rsObj = OFV.crearconsultaEx(sql, cn, 1, "")
		If Not rsObj.EOF Then sObj = rsObj(0)
		Call OFV.cerrarconsulta(rsObj)

		strsql = strsql & "FROM OBJETOS, TIPOBJETO, OBJPRTY, Estados " & addfrom
		strsql = strsql & " where    "
		strsql = strsql & "OBJETOS.tipobjeto=TIPOBJETO.tipobjeto and "
		strsql = strsql & "OBJPRTY.estado = Estados.objid and "
		strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto "
		strsql = strsql & "and Objetos.codobjeto = " & sID1
		sTitulo = "Tema: " & sObj


		Case "PORFECHA"
		strsql = "select * from csporfechas " & addfrom
		cdesde = CLng(CDate(sID1)) - CLng(CDate("01-01-1900"))
		chasta = CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))
		If sID3 = "Emision" Then
			strsql = strsql & "WHERE   xfecemis between '" & cdesde & "' and '" & chasta & "' "
			sTitulo = "Creacion&nbsp;&nbsp;Del:&nbsp;" & sID1 & "&nbsp;&nbsp;|&nbsp;al:&nbsp;" & sID2
		ElseIf sID3 = "Vigencia" Then
			strsql = strsql & "WHERE   xfecvigen between '" & cdesde & "' and '" & chasta & "' "
			sTitulo = "vigencia&nbsp;&nbsp;Del:&nbsp;" & sID1 & "&nbsp;&nbsp;|&nbsp;al:&nbsp;" & sID2
		ElseIf sID3 = "Eliminacion" Then
			strsql = strsql & "WHERE   xfecbaja between '" & cdesde & "' and '" & chasta & "' "
			sTitulo = "Eliminacion&nbsp;&nbsp;Del:&nbsp;" & sID1 & "&nbsp;&nbsp;|&nbsp;al:&nbsp;" & sID2
		ElseIf sID3 = "Vencimiento" Then
			strsql = strsql & "WHERE  xfecvto between '" & cdesde & "' and '" & chasta & "' "
			sTitulo = "Revision&nbsp;&nbsp;Del:&nbsp;" & sID1 & "&nbsp;&nbsp;|&nbsp;al:&nbsp;" & sID2
		ElseIf sID3 = "Publicacion" Then
			strsql = strsql & "WHERE   xfecpublic between '" & cdesde & "' and '" & chasta & "' "
			sTitulo = "Publicacion&nbsp;&nbsp;Del:&nbsp;" & sID1 & "&nbsp;&nbsp;|&nbsp;al:&nbsp;" & sID2
		End If

		Case "PORREVISION"
		strsql = "select * from csporfechas " & addfrom
		cdesde = CDate(sID1 + CLng(CDate("01-01-1900")) )
		chasta = CDate(sID2 + CLng(CDate("01-01-1900")) )

		strsql = strsql & "WHERE  xfecvto between '" & sID1 & "' and '" & sID2 & "' "
		sTitulo = "Revision&nbsp;&nbsp;Del:&nbsp;" & cdesde & "&nbsp;&nbsp;|&nbsp;al:&nbsp;" & chasta

		Case "PORACTA"

		'cdesde = CLng(CDate(sID1)) - CLng(CDate("01-01-1900"))
		'  chasta = CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))


		If  Robj = "F"  Then

			sID1   = Request.Form("Desde3") & ""
			' sID1  =  CLng(CDate(sID)) - CLng(CDate("01-01-1900"))
			sID2   = Request.Form("Hasta3") & ""
			' sID2  =  CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))
			sID3   = Request.Form("Fecha") & ""
			nacta  = Request.Form("nacta")


			Addqry = "&Menu=PORACTA&NOMBRE=ACTA&Desde=" & sID1 & "&Hasta=" & sID2 & "&Fecha=" & sID3

		Else
			sID1   = Request.querystring("Desde") & ""
			sID2   = Request.querystring("Hasta") & ""
			sID3   = Request.querystring("Fecha") & ""
			nacta  = Request.querystring("nacta") & ""
			'   response.write request.form

			Addqry = "&Menu=PORACTA&NOMBRE=ACTA&Desde=" & sID1 & "&Hasta=" & sID2 & "&Fecha=" & sID3


		End If


		sMenuAdd    = ""
		sMenuMas    = "&smenu=" & smenu & "&Desde=" & sID1 & "&Hasta="
		sMenuMas    =  sMenuMas & sID2 & "&Fecha=" & sID3


		If  nacta <> "" Then
			buscaacta = " Objprty.nroacta = '" & nacta & "'"
		Else
			buscaacta = "Objprty.nroacta is not null"
		End If


		cdesde = CLng(CDate(sID1)) - CLng(CDate("01-01-1900"))
		chasta = CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))


		strsql = "select DISTINCT  csporfechas.*, Objprty.fecacta,objetos.descri,tipobjeto.descri,objprty.fecemis "
		strsql = strsql & " from csporfechas join Objprty on csporfechas.codobjeto = Objprty.codobjeto, Estados  " & addfrom
		strsql = strsql & " where " & buscaacta & " and Objprty.fecacta between '" & cdesde & "' and '" & chasta & "' "

		sTitulo = "Por acta:  " & nacta &"&nbsp;&nbsp; Desde: "& sID1 & "&nbsp;&nbsp;|&nbsp; Hasta:&nbsp;" & sID2


	Case "PORRESUMEN"

			strsql = "select a.codiden,a.codobjeto,a.tipobjeto,a.descri,c.descri,b.fecemis,b.fecpublic,b.fecvigen,d.estado,b.autor,e.temadescri,e.stemadescri,f.objmensaje "
			strsql = strsql & " from objetos a "
			strsql = strsql & " inner join objprty b on (a.codobjeto = b.codobjeto and a.tipobjeto = b.tipobjeto) "
			strsql = strsql & " inner join tipobjeto c on (a.tipobjeto = c.tipobjeto) "
			strsql = strsql & " inner join estados d on (b.estado = d.objid) "
			strsql = strsql & " inner join csrelpathdescriex e on (a.codobjeto = e.cobj and a.tipobjeto = e.tobj) "
			strsql = strsql & " left join objinfo f on (a.codobjeto = f.codobjeto and a.tipobjeto = f.tipobjeto) "
			strsql = strsql & " where (b.fecpublic > " & cdesde - 1 & " and b.fecpublic < " & chasta + 1 & ")"

	End Select

	If smenu <> "PORSECTOR" And smenu <> "PORLINEGO" and smenu <> "PORRESUMEN" Then
		strsql = strsql & sverestado  & wherefrom
	End If

	sTitOrd  = ""
	sTitOrd2 = ""

	If smenu = "PORFECHA" Or smenu = "PORREVISION" Then

		If Sfiltro <> "TODOS" Then
			sTitOrd2 = "Solo&nbsp;" & Sfiltro & "&nbsp;&nbsp;"
			strsql  = strsql & " and csporfechas.descri='" & Sfiltro & "' "
		End If

		Select Case sorden

			Case "N"
			sTitOrd = sTitOrd & "Nombre"
			strsql  = strsql & " order by  objdescri, csporfechas.descri"
			Case "T"
			sTitOrd = sTitOrd & "Tipo"
			strsql  = strsql & " order by  csporfechas.descri,objdescri"
			Case "C"
			sTitOrd = sTitOrd & "Creacion"
			strsql  = strsql & " order by  fecemis desc, objdescri, csporfechas.descri"
			Case "E"
			sTitOrd = sTitOrd & "Estado"
			strsql  = strsql & " order by  fecemis, estdescri, objdescri, csporfechas.descri"
			Case "P"
			sTitOrd = sTitOrd & "Publica"
			strsql  = strsql & " order by  fecpublic desc, csporfechas.descri"
			Case "V"
			sTitOrd = sTitOrd & "Vigencia"
			strsql  = strsql & " order by  fecvigen desc, csporfechas.descri"
			Case "A"
			sTitOrd = sTitOrd & "Autor"
			strsql  = strsql & " order by  autor, csporfechas.descri"
		End Select
	ElseIf smenu = "PORSECTOR" Then

		Select Case sorden
			Case "N"
			sTitOrd = sTitOrd & "Nombre"
			strsql  = strsql & " order by X.descri,E.descri "
			Case "T"
			sTitOrd = sTitOrd & "Tipo"
			strsql  = strsql & " order by E.descri,X.descri "
			Case "C"
			sTitOrd = sTitOrd & "Creacion"
			strsql  = strsql & " order by C.fecemis desc,E.descri "
			Case "E"
			sTitOrd = sTitOrd & "Estado"
			strsql  = strsql & " order by C.fecemis, estdescri,X.descri,E.descri "
			Case "P"
			sTitOrd = sTitOrd & "Publica"
			strsql  = strsql & " order by C.fecpublic desc, E.descri"
			Case "V"
			sTitOrd = sTitOrd & "Vigencia"
			strsql  = strsql & " order by C.fecvigen desc, E.descri"
			Case "A"
			sTitOrd = sTitOrd & "Autor"
			strsql  = strsql & " order by C.autor, E	.descri"
		End Select

	ElseIf smenu = "PORLINEGO" Then
		Select Case sorden
			Case "N"
			sTitOrd = sTitOrd & "Nombre"
			strsql  = strsql & " order by OBJETOS.tsegmento,tipoagrupa.secuen,tiposegmento.secuen,tipobjeto.ARCHIVO,OBJETOS.INDOBJ, "
			strsql  = strsql & " OBJPRTY.FECVIGEN,OBJETOS.DESCRI"
		End Select

	ElseIf smenu = "PORRESUMEN" Then

			Select Case sorden
				Case "CI"
					sTitOrd = sTitOrd & "Codigo"
					strsql  = strsql & " order by a.codiden "
				Case "N"
					sTitOrd = sTitOrd & "Esquema Normativo"
					strsql  = strsql & " order by a.descri "
				Case "T"
					sTitOrd = sTitOrd & "Tipo"
					strsql  = strsql & " order by a.tipobjeto "
				Case "C"
					sTitOrd = sTitOrd & "Creacion"
					strsql  = strsql & " order by b.fecemis "
				Case "E"
					sTitOrd = sTitOrd & "Estado"
					strsql  = strsql & " order by d.estado "
				Case "P"
					sTitOrd = sTitOrd & "Publicacion"
					strsql  = strsql & " order by b.fecpublic "
				Case "V"
					sTitOrd = sTitOrd & "Vigencia"
					strsql  = strsql & " order by b.fecvigen "
				Case "A"
					sTitOrd = sTitOrd & "Autor"
					strsql  = strsql & " order by b.autor "
				Case "CA"
					sTitOrd = sTitOrd & "Carpeta"
					strsql  = strsql & " order by e.temadescri,e.stemadescri "
			End Select

	Else
		If Sfiltro <> "TODOS" Then
			sTitOrd2 = "Solo&nbsp;" & Sfiltro & "&nbsp;&nbsp;"
			strsql = strsql & " and tipobjeto.descri='" & Sfiltro & "' "
		End If

		Select Case sorden
			Case "N"
			sTitOrd = sTitOrd & "Nombre"
			strsql  = strsql & " order by objetos.descri,tipobjeto.descri "

			Case "P"
			sTitOrd = sTitOrd & "PROCESO"
			strsql  = strsql & " order by objprty.fecpublic desc, tipobjeto.descri"

		End Select



	End If

	'response.write "<br><br><br>" & strsql & "<br>"

	If request.querystring("scal")="S" Then
		strsql=session("sqlcons")
	Else
		session("sqlcons")=strsql
	End If
	'response.write session("sqlcons")
	'''Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")

	'response.write strsql & "<br>" & "ORDEN: " & sorden
	Set rs = OFV.crearconsultaEx(strsql, cn, 1, "")

	If MasterType <> "EXP" Then

		sHTML = OFV.FormHeader(Session("FormN")) & Chr(13)

	Else
		'    shtml=""
		Response.ContentType = "application/vnd.ms-excel"
		Response.AddHeader "content-disposition", " filename=salida.xls"
		' Response.AddHeader "Content-Type", "application/xls"
		'Response.AddHeader "Content-Disposition", "inline;filename=33.xls"
	End If

	sHTML = ofv.FormHeader(session("FormN")) & Chr(13)
	xs = "<script language=JavaScript>" & Chr(13)
	xs = xs & "function AbrirCons(Datos) {" & Chr(13)
	xs = xs & " var h=570;w=790 ;"
	If  session("APLMDCLI")="BC" Then
		xs = xs &  "if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;} "
	End If

	xs = xs & " myWin= open(Datos"
	xs = xs & "," & Chr(34) & "Cons" & Chr(34) & ","' & Chr(34)
	'xs = xs & "menubar=no,toolbar=no,height=570,width=790,resizable=yes,status=yes" & Chr(34)
	xs = xs &"  'menubar=no,toolbar=no,width='+w+',height='+h+',top=0,left=0,resizable=yes,status=yes'" '& CHR(34)
	xs = xs & " ); " & Chr(13) & " } " & Chr(13) & "</script>" & Chr(13)

	sHTML = sHTML & xs


	'   response.write Request.querystring & "<br>"
	'   response.write Request.form & "<br>"

	sImagenM = "<B><FONT FACE=""corflaticon"" COLOR=#005500 SIZE=3>2</FONT></B>"
	sImagenI2 = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=3>&#0128;</FONT></B>"
	sImagenI = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=2>u</FONT></B>"
	sImagenI3 = "<B><FONT FACE=""corflaticon"" COLOR=#6699CC SIZE=2>=</FONT></B>"
	sImagenEM = "<B><FONT FACE=corflaticon COLOR=#F7AA22 SIZE=3>&#0252;</FONT></B>"
	sImagen4si = "<B><FONT FACE=corflaticon COLOR=#993333 SIZE=3>4</FONT></B>"
	sImagenx = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=4>x</FONT></B>"
	sImagena = "<B><FONT FACE=corflaticon COLOR=#f7aa22 SIZE=2>4</FONT></B>"



	Fila    = ""
	Celda   = ""
	sHref   = ""

	If MasterType <> "EXP" Then
		If  Session("EMPMUN") = "S" Then

			sLink   = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(0,5),"0","")
			Celda   = Celda & ofv.GenCelda("","tt ","","8%","","","","EMPRESA:&nbsp;&nbsp;","si")
			Celda   = Celda & ofv.GenCelda("","ttt ","center","5%","","","",sLink,"si")
			Celda   = Celda & ofv.GenCelda("","ttt  ","","","","","","&nbsp;<BIG>" & ATREMP(0,4) & "</BIG>","si")

			Fila    = fila & ofv.GenRow("","","top","","","","",Celda,"si")

			Tabla   = ofv.GenTabla("","aaa  ","","100% ","","0","","-1","-2","","",Fila,"si")
			sHTML   = sHTML & Tabla
			sHTML   = sHTML & "<br>"
		End If
	End If

	Fila = ""

	'***************** LINK DEL CAMBIAR VISTA
	sNombre=Replace(sNombre," ","_")
	Select Case smenu

		Case "PORACTA"
		sAccion = "ConsdocumentoswebScons.asp?xx=" & Addqry ' "&id=" & sID1 & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N&Nombre=" & sNombre & "&Desde="& Request.querystring("Desde")  &"&Hasta="& Request.querystring("hasta") &"&Fecha="& si3
		sAccion = sAccion &"&nacta="& nacta

		Case "PORFECHA"
		sAccion = "ConsdocumentoswebScons.asp?xx=" & Addqry ' & "&id=" & sID1 & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N&Nombre=" & sNombre & "&Desde="& sid  &"&Hasta="& sid2 &"&Fecha="& sid3

		Case "PORESTADO"
		sAccion = "ConsdocumentoswebScons.asp?xx=" & Addqry '& sMenu & "&id=" & sID & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N&Nombre=" & sNombre & "&Desde="& cdesde &"&Hasta="& chasta &"&Fecha="& Request.querystring("fecha")

		Case "PORREVISION"
		sAccion = "ConsdocumentoswebScons.asp?xx=" & Addqry ' "&id=" & sID1 & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N&Nombre=" & sNombre & "&Desde="& sid &"&Hasta="& sid2 &"&Fecha="& Request.querystring("fecha")

		Case "PORSECTOR"
		sAccion = "ConsdocumentoswebScons.asp?xx=" & Addqry '"&id=" & sID1 & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N&Nombre=" & sNombre & "&perfil="& Sperfil &"&Hasta="& chasta &"&Fecha="& Request.querystring("fecha")

		Case "PORCOMPONENTE"
		sAccion = "ConsdocumentoswebScons.asp?xx=" & Addqry ' "&id=" & sID1 & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N&Nombre=" & sNombre

		Case "PORTEXTO"
		sAccion = "ConsdocumentoswebScons.asp?xx=" & Addqry '

		Case "PORPRODUCTO"
		sAccion = "ConsdocumentoswebScons.asp?xx=" & Addqry

		Case "PORIDENTIFICACION"
		sAccion = "ConsdocumentoswebScons.asp?xx=" & Addqry

		Case "PORNROCIRCULAR"
		sAccion = "ConsdocumentoswebScons.asp?xx=" & Addqry

		Case "PORINSTRUMENTO"
		sAccion = "ConsdocumentoswebScons.asp?xx=" & Addqry

		Case "PORLINEGO"
		sAccion = "ConsdocumentoswebScons.asp?xx=" & Addqry

		Case "PORRESUMEN"
		sAccion = "ConsdocumentoswebScons.asp?xx=" & Addqry

	End Select


	'response.write "<br> starget:"& starget

	'RESPONSE.WRITE smenu

	'*****************************************
	'''
	'sAccion = "ConsdocumentoswebScons.asp?Menu=" & sMenu & "&id=" & sID1 & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N&Nombre=" & sNombre& "&Desde="& cdesde &"&Hasta="& chasta &"&Fecha="& Request.querystring("fecha")
	if smenu <> "PORRESUMEN" then

		sInput   = ofv.GenerarInput("","Vista Web","submit","","10","bt","")

	' sInput   = sinput & ofv.GenerarInput("lista",sID1,"hidden","","1","dt","")
	' sInput   = sinput & ofv.GenerarInput("lista2",sID1,"hidden","","1","dt","")
		Formvista = ofv.GenForm("vistar","",sAccion,"","post"," style='display:none;' ",sInput,"si") & "Cambiar a Vista Web"

	end if


		Celda  = OFV.GenCelda("", "vt style='border-style:solid;border-width:1;border-color:#dddddd;'  ", "", "", "", "", "", sImagena & "&nbsp;" & sTitulo, "si")



	If smenu <> "PORACTA" Then
		Celda  = Celda & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#dddddd;cursor:pointer;' ","center onclick=""document.getElementById('vistar').submit();"" ","20%","","","",Formvista,"si")
	Else
		Celda  = Celda & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#dddddd;cursor:pointer;' ","center onclicka=""document.getElementById('vistar').submit();"" ","20%","","","","&nbsp;","si")
	End If
	If smenu <> "PORLINEGO" Then
		Celda  = Celda & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#dddddd;' ","right","20%","","","","&nbsp;" & sTitOrd2 & "Orden:&nbsp;" & sTitOrd & "&nbsp;&nbsp;","si")
	End If

	Fila   = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

	tablaz = OFV.GenTabla("", "aaa bgcolor='#eeeeee' style='border-style:solid;border-width:1;border-color:#dddddd;' ", "", "100% ", "", "0", "", "1", "2", "", "", Fila, "si")

	sHTML = sHTML & Tablaz

	Fila = ""

	'' AGREGADO 26052014

	If smenu ="PORESTADO" Then
		If ssEstado = "VIGENTE" Then                     'pregunto si es estado vigente para obtener los totales por tipo'
			strsql = " SELECT isnull(max(TIPOBJETO.descri),'SIN DATOS') as tipo, isnull(sum(1),0) as total "
			strsql = strsql & " FROM OBJETOS, TIPOBJETO, OBJPRTY, Estados , csrelpath a, tema t, prodserv  st "
			strsql = strsql & " where  OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and OBJPRTY.estado = Estados.objid and OBJETOS.codobjeto = OBJPRTY.codobjeto "
			strsql = strsql & " and OBJPRTY.estado = 3 and Objetos.objnivel >= 3 And Estados.objid In (9,1,6,2,5,3,4)  and Objetos.OBJUN = 1  and TIPOBJETO.TIPOCPT = 'D' "
			strsql = strsql & " and (OBJETOS.APLIC = 'EM' or OBJETOS.APLIC = 'GLB')  and (a.cobj = OBJETOS.codobjeto)  and (a.tema = t.codtema) "
			strsql = strsql & " and (a.stema = st.codprodserv) "
			strsql = strsql & " group by  TIPOBJETO.descri "


			Set rsx1 = OFV.crearconsultaEx(strsql, cn, 1, "")

			If Not rsx1.eof  Then

				'' Titulo de Totales por Tipo -unica vez-

				Celda = ""
				Celda  = OFV.GenCelda("", "vt xstyle='border-style:solid;border-width:5;border:none;'  ", "", "20%", "", "", "", "&nbsp;" & "TOTALES POR TIPO" & "", "si")
				Fila   = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
				tablac = OFV.GenTabla("", "aaa bgcolor='#eeeeee' border=1 xstyle='border-style:solid;border-width:1;border:none;' ", "", "60% ", "", "0", "", "1", "2", "", "", Fila, "si")
				sHTML = sHTML & Tablac
				Fila = ""

				'' Rutina para mostrar los totales
				'' response.write " count "    & rsx1.fields.count   &"<br>"

				iii = 0
				iii3 = 0
				For iii = 1 To 1         'to 9 step 3

					'        response.write " &rsx1 " & rsx1(0) &"<br>"
					'        response.write " &iii "  & iii     &"<br>"
					'        response.write " Xrsx1X " & Hxcod & " " & Hxtip & " " & iii &"<br>"

					Fila  = ""
					Do While Not rsx1.eof
						Celda = ""
						For iii3=1 To 3
							If  rsx1.eof Then
								Exit For 'RSX1.MOVENEXT

							End If
							Hxcod = rsx1(0) & ""  'descripcion
							Hxtip = rsx1(1) & ""  'cantidad

							''             '  response.write " *rsx1 " & rsx1(0) &"<br>"

							Celda  = Celda & OFV.GenCelda("", "vt xstyle='border-style:solid;border-width:5;border:none;'  ", "", "20%", "", "", "", "&nbsp;" & Hxcod & ": " & Hxtip & "", "si")

							RSX1.MOVENEXT
						Next

						Fila   = fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

					Loop

				Next

			End If
		End If

		tablac = OFV.GenTabla("", "aaa bgcolor='#eeeeee' border=1 xstyle='border-style:solid;border-width:1;border:none;' ", "", "60% ", "", "0", "", "1", "2", "", "", Fila, "si")
		sHTML = sHTML & Tablac
		Fila = ""

	End If

	'' AGREGADO 26052014 FIN

	'' sHTML = OFV.MENUHeader("","")

	'  response.write "<br><br><br><br>" & Request.form & "<br>"

	'  response.write " starget " & starget & "<br>"
	Celda = ""
	If smenu <> "PORLINEGO" And MasterType <> "EXP" Then

		if smenu = "PORRESUMEN" then
			Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=CI target=_self ><B><u>&nbsp;COD. IDEN&nbsp;</u></B></A>", "si")
		end if
		if smenu = "PORRESUMEN" then
			Celda = Celda & OFV.GenCelda("", "ttx colspan=2", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=N target=_self ><B><u>&nbsp;ESQUEMA NORMATIVO&nbsp;</u></B></A>", "si")
		else
			Celda = Celda & OFV.GenCelda("", "ttx colspan=3", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=N target=_self ><B><u>&nbsp;ESQUEMA NORMATIVO&nbsp;</u></B></A>", "si")
		end if
		Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=T target=_self ><B><u>&nbsp;TIPO&nbsp;</u></B></A>", "si")
		If smenu = "PORESTADO" Then
			If sID = 4 Then
				Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Eliminacion", "si")
			Else
				Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=C target=_self ><B><u>&nbsp;CREACION&nbsp;</u></B></A>", "si")
			End If
		Else
			Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=C target=_self ><B><u>&nbsp;CREACION&nbsp;</u></B></A>", "si")
		End If
		If smenu = "PORFECHA" Then
			If sID3 = "Eliminacion" Then
				Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Eliminacion", "si")
			ElseIf sID3 = "Vencimiento" Then
				Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "REVISION", "si")
			ElseIf sID3 = "Publicacion" Then
				'*********Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Publicacion", "si")
				Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=P target=_self ><B><u>&nbsp;PUBLICACION&nbsp;</u></B></A>", "si")
			Else
				'*********Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Vigencia", "si")
				Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=V target=_self ><B><u>&nbsp;VIGENCIA&nbsp;</u></B></A>", "si")
			End If
		ElseIf smenu = "PORREVISION" Then
			Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "REVISION", "si")
		Else
			'*******Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Publicacion", "si")
			Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=P target=_self ><B><u>&nbsp;PUBLICACION&nbsp;</u></B></A>", "si")
			'*******Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Vigencia", "si")
			Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=V target=_self ><B><u>&nbsp;VIGENCIA&nbsp;</u></B></A>", "si")
		End If
		if smenu = "PORRESUMEN" then
			Celda = Celda & OFV.GenCelda("", "tTx", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=CA target=_self ><B><u>&nbsp;CARPETA&nbsp;</u></B></A>", "si")
		end if

		If smenu <> "PORESTADO" Then
			Celda = Celda & OFV.GenCelda("", "tTx", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=E target=_self ><B><u>&nbsp;ESTADO&nbsp;</u></B></A>", "si")
			Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=A target=_self ><B><u>&nbsp;AUTOR&nbsp;</u></B></A>", "si")
		Else
			'*****Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Autor", "si")
			Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=A target=_self ><B><u>&nbsp;AUTOR&nbsp;</u></B></A>", "si")
			'  If sID = 3 Or sID = 4 Then
			Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Aprobo", "si")
			'  End If
			If sID = 4 Then
				Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Elimino", "si")
			End If
		End If

	Else

		if smenu = "PORLINEGO" then

	 		if mastertype <> "EXP" then
				Celda = OFV.GenCelda("", "ttx style='color:#FFFFFF;' ", "", "", "", "", "", "<B><u>&nbsp;PROCESO&nbsp;</u></B>", "si")
				Celda = Celda & OFV.GenCelda("", "ttx style='color:#FFFFFF;' ", "", "", "", "", "", "<B><u>&nbsp;SUBPROCESO&nbsp;</u></B>", "si")
				Celda = Celda & OFV.GenCelda("", "ttx style='color:#FFFFFF;' ", "", "", "", "", "", "<B><u>&nbsp;DESCRIPCION&nbsp;</u></B>", "si")
				Celda = Celda & OFV.GenCelda("", "ttx style='color:#FFFFFF;' ", "", "", "", "", "", "<B><u>&nbsp;TIPO&nbsp;</u></B>", "si")
				Celda = Celda & OFV.GenCelda("", "ttx style='color:#FFFFFF;' ", "", "", "", "", "", "<B><u>&nbsp;FECHA VIG&nbsp;</u></B>", "si")
	   		else
	 	    	Celda = OFV.GenCelda("", "ttx", "", "", "", "", "", "<B>&nbsp;PROCESO&nbsp;</B>", "si")
	    		Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "<B>&nbsp;SUBPROCESO&nbsp;</B>", "si")
	    		Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "<B>&nbsp;DESCRIPCION&nbsp;</B>", "si")
	   	 		Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "<B>&nbsp;TIPO&nbsp;</B>", "si")
	   	 		Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "<B>&nbsp;FECHA VIG&nbsp;</B>", "si")
	   		end if

	    else

			Celda = OFV.GenCelda("", "ttx", "", "", "", "", "", "<B>&nbsp;COD. IDEN&nbsp;</B>", "si")
    		Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "<B>&nbsp;ESQUEMA NORMATIVO&nbsp;</B>", "si")
    		Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "<B>&nbsp;TIPO&nbsp;</B>", "si")
   	 		Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "<B>&nbsp;CREACION&nbsp;</B>", "si")
	   	 	Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "<B>&nbsp;PUBLICACION&nbsp;</B>", "si")
	   	 	Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "<B>&nbsp;VIGENCIA&nbsp;</B>", "si")
	   	 	Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "<B>&nbsp;CARPETA&nbsp;</B>", "si")
	   	 	Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "<B>&nbsp;ESTADO&nbsp;</B>", "si")
	   	 	Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "<B>&nbsp;AUTOR&nbsp;</B>", "si")
	   	 	Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "<B>&nbsp;MENSAJE&nbsp;</B>", "si")

	    end if

	End If
	Fila = fila & OFV.GenRow("", "", "", "", "20", "", "", Celda, "si")

	If MasterType <> "EXP" Then
		response.flush
		response.write sHTML
		sHTML = ""
		response.flush
	Else

	End If
	' response.write strsql


	'''''''''''''''''''' FIN ENCABEZADOS ''''''''''''''''''''''''''''''''''''''

	nLinea = 0

	fi = CLng(CDate("01-01-1900"))
	PosReg = 0

	xTrColor = "#FFFFFF"

	On Error Resume Next
	If Not rs.EOF Then
		If MasterType <> "EXP" Then
			rs.Move (CantRegMover)
		End If
		Do While (Not rs.EOF) And (PosReg < CantRegMostrar)
			dEmision  = "..."
			dVigencia = "..."
			sArchivo = "Iconos/" & rs(11)
			If Not IsNull(rs(0)) Then stipo = rs(0)

			If smenu = "PORLINEGO" Then
				rwspan = ""
				If MasterType <> "EXP" Then rwspan = " rowspan=2 "
			End If
			If smenu = "PORESTADO" Then
				If sID = 4 Then
					If Not IsNull(rs(5)) Then dEmision = CDate(CLng(rs(7)) + fi)
				Else
					If Not IsNull(rs(5)) Then dEmision = CDate(CLng(rs(5)) + fi)
				End If
			Else
				If Not IsNull(rs(5)) Then dEmision = CDate(CLng(rs(5)) + fi)
			End If

			dEmi = InStr(dEmision, "1900")
			If dEmi > 0 Then dEmision = "..."

			If smenu = "PORFECHA" Then
				If sID3 = "Eliminacion" Then
					If Not IsNull(rs(7)) Then dVigencia = CDate(CLng(rs(7)) + fi)
				ElseIf sID3 = "Vencimiento" Then
					If Not IsNull(rs(16)) Then dVigencia = CDate(CLng(rs(16)) + fi)
				ElseIf sID3 = "Publicacion" Then
					If Not IsNull(rs(24)) Then dVigencia = CDate(CLng(rs(24)) + fi)
				Else
					If Not IsNull(rs(6)) Then dVigencia = CDate(CLng(rs(6)) + fi)
				End If
			ElseIf smenu = "PORREVISION" Then
				If Not IsNull(rs(16)) Then dVigencia = CDate(CLng(rs(16)) + fi)
			Else
				If Not IsNull(rs(6)) Then dVigencia = CDate(CLng(rs(6)) + fi)
			End If

			dpublic = "..."

			If smenu <> "PORFECHA" And smenu <> "PORREVISION"  Then
				If Not IsNull(rs(22)) Then dpublic = CDate(CLng(rs(22)) + fi)
				dVign = InStr(dpublic, "1900")
				If dVign > 0 Then dpublic = "..."
			End If

			dVign = InStr(dVigencia, "1900")
			If dVign > 0 Then dVigencia = "..."

			If smenu <> "PORESTADO" Then
				If Not IsNull(rs(8)) Then
					If (Not CDate(CLng(rs(16)) + fi) > Date) And (CDate(CLng(rs(16)) + fi) <> CDate("01-01-1900")) Then
						sEstado = "VENCIDA"
					Else
						sEstado = rs(8)
					End If
				End If
				sAutor   = "&nbsp;"
				If Not IsNull(rs(12)) Then sAutor   = rs(12)
			Else
				sAutor   = "&nbsp;"
				sAprobo  = "&nbsp;"
				sElimino = "&nbsp;"
				If Not IsNull(rs(12)) Then sAutor   = rs(12)
				If Not IsNull(rs(13)) Then sAprobo  = rs(13)
				If Not IsNull(rs(14)) Then sElimino = rs(14)
			End If


			'
			If Not IsNull(rs(9)) Then nIden = rs(9)
			If Not IsNull(rs(10)) Then sobjeto = rs(10)

			If smenu <> "PORFECHA" And smenu <> "PORREVISION"  Then

				If Not IsNull(rs(19)) Then sobjeto = sobjeto & " - " & rs(19)
				If Not IsNull(rs(20)) And rs(19)&"" ="V" Then sobjeto = sobjeto & " " & rs(20)

			Else

				If smenu="PORFECHA"   Then
					Call GETLISTOBJ("E",rS(3),0)
					sobjeto =  XGUTABLISDAT
				End If
			End If

			If smenu = "PORLINEGO" Then

				procTit = rs("ta")
				subpTit = rs("ts")
				sLink = EsqnomPx(rs("tipobjeto"), rs("codobjeto"), ofv)
				tobj = rs("tdescri")
				fecha = CDate(CLng(rs(20)) + fi)
				fecha = fechadorCor(fecha,"")

			End If

			If smenu = "PORRESUMEN" Then

				codIden = rs(0)

				tobj = rs(4)

				fechac = CDate(CLng(rs(5)) + fi)
				fCreac = fechadorCor(fechac,"")

				fechap = CDate(CLng(rs(6)) + fi)
				fPublic = fechadorCor(fechap,"")

				fechav = CDate(CLng(rs(7)) + fi)
				fVigen = fechadorCor(fechav,"")

				carpeta = rs(10) & "/" & rs(11) & "/" & rs(3)
				estado = rs(8)
				autor = rs(9)
				xmens = rs(12)

			End If

			If MasterType <> "EXP" Then

				If Session("USRES") <> "B" Then
					sImagen = OFV.GenImage("", "", "", "", sArchivo, "0", "")
				Else
					sImagen = "<font face=corflaticon size=4 color=#993333>2</font>"
				End If

				sImagen = "<font face='corflaticon' size=3 color='#993333'>4</font>"

				sActcon = "Consultar.asp?Codigo=" & rs("codobjeto") & "&Tipo=" & rs("tipobjeto")
				sActcon = sActcon & "&Perfil=" & rs("perfil") & "&Datos=" & OFV.convertircar(sMenuMas, "&", "*")
				sActcon = sActcon & "&Volver=" & sVolver
				sHref = "&nbsp;"
				sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "'); return false;" & Chr(34)

				if smenu <> "PORRESUMEN" then
					sLink = OFV.GenLink("", "", sHref, "", "", sFuncion, sImagen, "si")
				else
					sLinkx = OFV.GenLink("", "", sHref, "", "", sFuncion, sImagen, "si")
				end if

				If smenu <> "PORLINEGO"  Then
					If smenu = "PORRESUMEN"  Then
						Celda = OFV.GenCelda("", "ut", "", "3%", "", "", "", codIden, "si")
						Celda = Celda & OFV.GenCelda("", "ut", "", "1%", "", "", "", sLinkx, "si")
					else
						Celda = OFV.GenCelda("", "ut", "", "1%", "", "", "", sLink, "si")
					end if
				Else
					Celda = OFV.GenCelda("", "ut", "", "18%", "", "", "",procTit, "si")
				End If
				smasc = "&nbsp;"
				If smenu <> "PORFECHA"  And smenu <> "PORREVISION" and smenu <> "PORRESUMEN" Then
					If  rs(21) = 1 And  SESSION("UN") <> 1 And  SESSION("EMPMUN") = "S" Then
						smasc = "&nbsp;<font color=#993333  ><font class=btic title='Documento Corporativo' >G</font></font>"
					End If
				else

					smasc = ""

				End If
				If session("APLMDCLI")="BC" Then
					sLink = OFV.GenLink("", "nt", sHref, "", "", sFuncion, "<b>" & rs("codiden") & "&nbsp;-&nbsp;</b>" & sobjeto, "si")
				Else
					If smenu <> "PORLINEGO" Then
						if smenu <> "PORRESUMEN" then
							sLink = OFV.GenLink("", "nt", sHref, "", "", sFuncion, rs("descri") &"&nbsp; : &nbsp;</b>"& rs("codiden") & "&nbsp;-&nbsp;</b>" & sobjeto, "si")
						end if
					Else
						sobj=EsqnomPx(rs("tipobjeto"), rs("codobjeto"), ofv)
						sLink = OFV.GenLink("", "nt style='color:#0017FF;' ", sHref, "", "Consultar", sFuncion, sobj, "si")
					End If
				End If
				sTipoDsc = stipo
				sTipoDsc = "<a class=nt title='solo documentos de tipo " & stipo & "' href=" & starget & "&sorden=" & sorden & "&sfiltro=" & Replace(stipo, " ", "_") & " target=_self ><font color='" & Session("CMFPNIPFT") & "' >" & sTipoDsc & "</font></A>"
				If smenu <> "PORLINEGO" Then
					if smenu = "PORRESUMEN" then
						sLinkz = OFV.GenLink("", "nt", sHref, "", "", sFuncion, EsqnomPx(rs(2), rs(1), ofv), "si")
						Celda = Celda & OFV.GenCelda("", "ut ", "", "8%", "", "", "",sLinkz, "si")
					else
						Celda = Celda & OFV.GenCelda("", "ut ", "", "3%", "", "", "", sLink, "si")
					end if
				Else
					Celda = Celda & OFV.GenCelda("", "ut ", "", "18%", "", "", "",subpTit , "si")
				End If
				If smenu <> "PORLINEGO" Then
					if smenu = "PORRESUMEN" then
						Celda = Celda & OFV.GenCelda("", "ut ", "", "2%", "", "", "", tobj , "si")
					else
						Celda = Celda & OFV.GenCelda("", "ut ", "", "4%", "", "", "", smasc , "si")
					end if
				Else
					Celda = Celda & OFV.GenCelda("", "ut ", "", "36%", "", "", "", sLink, "si")
				End If
				If smenu <> "PORLINEGO" Then
					if smenu = "PORRESUMEN" then
						Celda = Celda & OFV.GenCelda("", "ut ", "", "3%", "", "", "",fCreac, "si")
					else
						Celda = Celda & OFV.GenCelda("", "ut ", "", "5%", "", "", "","" &  sTipoDsc, "si")
					end if
				Else
					Celda = Celda & OFV.GenCelda("", "ut ", "", "14%", "", "", "",rs("tdescri"), "si")
				End If
				If smenu <> "PORLINEGO" Then
					if smenu = "PORRESUMEN" then
						Celda = Celda & OFV.GenCelda("", "ut ", "", "3%", "", "", "",fPublic, "si")
					else
						Celda = Celda & OFV.GenCelda("", "ut ", "", "7%", "", "", "","" &   fechadorCor(dEmision,"") & "", "si")
					end if
				Else

					If rs(20) <> 0 Then
						fecha = CDate(CLng(rs(20)) + fi)
						fecha = fechadorCor(fecha,"")
					Else
						fecha = "..."
					End If

					Celda = Celda & OFV.GenCelda("", "ut ", "", "14%", "", "", "",fecha , "si")
				End If

				If smenu <> "PORFECHA"  And smenu <> "PORREVISION" And smenu <> "PORLINEGO" Then
					if smenu = "PORRESUMEN" then
						Celda = Celda & OFV.GenCelda("", "ut ", "", "3%", "", "", "",fVigen, "si")
					else
						Celda = Celda & OFV.GenCelda("", "ut ", "", "10%", "", "", "","" &   fechadorCor(dpublic,"") & "", "si")
					end if
				End If


				If smenu <> "PORLINEGO" Then
					if smenu = "PORRESUMEN" then
						Celda = Celda & OFV.GenCelda("", "ut ", "", "10%", "", "", "",carpeta, "si")
					else
						Celda = Celda & OFV.GenCelda("", "ut ", "", "10%", "", "", "","" &   fechadorCor(dVigencia,"") & "", "si")
					end if
				End If
				If smenu <> "PORESTADO" Then
					If sEstado = "VENCIDA" Then
						sInput = "<font color=#880000>" & sEstado & "</font>"
					Else
						sInput = sEstado
					End If
					If smenu <> "PORLINEGO" Then
						if smenu = "PORRESUMEN" then
							Celda = Celda & OFV.GenCelda("", "ut ", "", "2%", "", "", "",estado, "si")
							Celda = Celda & OFV.GenCelda("", "ut ", "", "2%", "", "", "",autor, "si")
						else
							Celda = Celda & OFV.GenCelda("", "ut ", "", "5%", "", "", "", sInput & "", "si")
							Celda = Celda & OFV.GenCelda("", "ut ", "", "5%", "", "", "", sAutor & "", "si")
						end if
					End If
				Else
					Celda = Celda & OFV.GenCelda("", "ut ", "", "5%", "", "", "", sAutor, "si")
					'   If sID = 3 Or sID = 4 Then
					sAprobo = "<a href=coobjwfldet.asp?Tipobjeto=" & rs("tipobjeto") & "&Codobjeto="
					sAprobo = sAprobo & rs("codobjeto") & " target=_blank><IMG SRC=Imagenes\ICorSInfoUser.jpg title=" & Chr(34)
					sAprobo = sAprobo & " Esquema&nbsp;de&nbsp;Autorizaci&oacute;n&nbsp;de&nbsp;" & Chr(34) & " border=0></a>"

					sRef     = "coobjwfldet.asp?Tipobjeto=" & rs("tipobjeto") & "&Codobjeto=" & rs("codobjeto")
					sImagenP = "&#0232;"
					sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
					sinput   = "<a class='ft1n' " & sFuncAnc & " href='" & sRef & "' "
					sinput   = sinput  & "' title='Esquema&nbsp;de&nbsp;Autorizaci&oacute;n&nbsp;de&nbsp;" & sobjeto & "' target='_blank' >" & sImagenP & "</a>"
					sAprobo  =  sinput

					Celda    = Celda & OFV.GenCelda("", "ft", "", "5%", "", "", "", sAprobo, "si")
					' End If
					If sID = 4 Then
						Celda = Celda & OFV.GenCelda("", "ut ", "", "5%", "", "", "", sElimino, "si")
					End If
				End If
			Else

			    if smenu = "PORLINEGO" then

					Celda = OFV.GenCelda("", "ut", "", "18%", "", "", "",procTit, "si")
					Celda = Celda & OFV.GenCelda("", "ut ", "", "18%", "", "", "",subpTit , "si")
					Celda = Celda & OFV.GenCelda("", "ut ", "", "36%", "", "", "", sLink, "si")
					Celda = Celda & OFV.GenCelda("", "ut ", "", "14%", "", "", "",tobj, "si")
					Celda = Celda & OFV.GenCelda("", "ut ", "", "14%", "", "", "",fecha , "si")

				elseif smenu = "PORRESUMEN" then

					sLink = EsqnomPx(rs("tipobjeto"), rs("codobjeto"), ofv)

					Celda = OFV.GenCelda("", "ut", "", "18%", "", "", "",codIden, "si")
					Celda = Celda & OFV.GenCelda("", "ut ", "", "18%", "", "", "",sLink, "si")
					Celda = Celda & OFV.GenCelda("", "ut ", "", "36%", "", "", "",tobj, "si")
					Celda = Celda & OFV.GenCelda("", "ut ", "", "14%", "", "", "",fCreac, "si")
					Celda = Celda & OFV.GenCelda("", "ut ", "", "14%", "", "", "",fPublic, "si")
					Celda = Celda & OFV.GenCelda("", "ut ", "", "18%", "", "", "",fVigen, "si")
					Celda = Celda & OFV.GenCelda("", "ut ", "", "18%", "", "", "",carpeta, "si")
					Celda = Celda & OFV.GenCelda("", "ut ", "", "36%", "", "", "",estado, "si")
					Celda = Celda & OFV.GenCelda("", "ut ", "", "14%", "", "", "",autor, "si")
					Celda = Celda & OFV.GenCelda("", "ut ", "", "14%", "", "", "",xmens, "si")

				end if

			End If

			Fila = Fila & OFV.GenRow("", "aaa bgcolor=" & xTrColor & " ", "", "", "22", "", "", Celda, "si")

			If smenu = "PORRESUMEN" and mastertype <> "EXP" Then

			    Celda = OFV.GenCelda("", "ut colspan=10 style='border-style:solid;border-width:1;border-color:#aaaaaa;' ", "", "5%", "", "", "",xmens, "si")
			    Fila = Fila & OFV.GenRow("", "aaa bgcolor=" & xTrColor & " ", "", "", "22", "", "", Celda, "si")

			    'SPATHDOCD = "<hr  style='border-style:solid;border-color:#aaaaaa;' >"
			    'Celda = OFV.GenCelda("", "ut colspan=10 ", "", "5%", "", "", "", SPATHDOCD, "si")
			    'Fila = Fila & OFV.GenRow("", "aaa bgcolor=" & xTrColor & " ", "", "", "", "", "", Celda, "si")

            end if

			If xTrColor = "#FFFFFF" Then
				xTrColor = "#F0F0F0"
			Else
				xTrColor = "#FFFFFF"
			End If

			if mastertype <> "EXP" then
				PosReg = PosReg + 1
			end if
			rs.movenext
			nLinea = nLinea + 1
			If MasterType <> "EXP" Then
				response.write sHTML
				sHTML = ""
				If nLinea > 20 Then
					nLinea = 0
					response.flush
				End If

			Else

			End If

		Loop

	Else

		XFRASE = "<BR><BR>&nbsp;<B><font color=#993333 SIZE=3 FACE='corflaticon' >i</font></B>"
		XFRASE = XFRASE &  "&nbsp;NO SE ENCONTRARON DOCUMENTOS PARA LA BUSQUEDA SOLICITADA<BR><BR>"

		Celda = OFV.GenCelda("", "ut colspan=2 ", "", "", "", "", "", XFRASE, "si")


		Fila = Fila & OFV.GenRow("", "aaa bgcolor=" & xTrColor & " ", "", "", "22", "", "", Celda, "si")


	End If


	If MasterType <> "EXP" Then
		sHTML = sHTML & OFV.GenTabla("", "", "", "100%  style={border-style:solid;border-width:1;border-color:#eeeeee;} ", "", "0", "", "-1", "0", "", "", Fila, "si") & SFORMEXP
	Else
		Response.Write OFV.GenTabla("", "", "", "100%  style={border-style:solid;border-width:1;border-color:#eeeeee;} ", "", "0", "", "-1", "0", "", "", Fila, "si")
	End If
	'response.write Addqry

	If SMENU = "PORREVISION" Then
		sAccion = "COFECREVISION.asp?Menu=" & sMenu & "&Nombre=" & sNombre

	ElseIf SMENU = "PORLINEGO" Then

		If sid2 = "nsnc" Then sid2 = 0
		If sid3 = "nsnc" Then sid3 = 0

		sAccion = "Subconsulta.asp?Menu=" & sMenu & "&Nombre=" & sNombre & "&opc=vuel&id=" & sid1 & "&proc=" & sid2 & "&subp=" & sid3 & "&sfiltro=" & sfiltro
	Else
		sAccion = "Subconsulta.asp?Menu=" & sMenu & "&Nombre=" & sNombre
	End If



	Celda   = ofv.BotonVolver(sAccion)


	If MasterType <> "EXP" Then
		sHTML = sHTML & OFV.botoneraESSX("c", rs, CantRegMover, CantRegMostrar, "no", "no", PosReg, Addqry, "", "no", "", "", Celda)
		sHTML = sHTML & "</body></html>"
	Else
		'   Response.Write "</body></html>"
	End If

	If smenu = "PORSECTOR" Then
		Session("sperfil") = Replace(SPerfil, " ", "_")
	Else
		Session("sperfil") = ""
	End If
	Call OFV.cerrarconsulta(rs)
	Call OFV.cerrarconn(cn)

	Generar = sHTML

	'''''''''''''''''''

	sHTML = Generar

	'' sHTML = oClc.Generar



End If

If MasterType <> "EXP" Then
	'response.write strsql
	response.write sHTML
	response.flush
Else

End If

Function vertexto(nwtexto)

	nwtextoF = nwtexto & ""

	If  nwtextoF <> "" Then

		nwtextoA = Split(nwtextoF," ")

		nwsep    = ""
		nwtextoF = ""

		For iii2 = 0 To UBound(nwtextoA)

			If Len(nwtextoA(iii2)) > 2 Then

				nwtextoF =  nwtextoF & nwsep & nwtextoA(iii2)

				nwsep = "%"
			End If
		Next


	End If



	vertexto = nwtextoF

End Function

%>