
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok = OFV.CheckUsuario()

If ok <> "Y" Then
	uv = OFV.uv
	If uv = "NO" Then ok = OFV.ValidarUsuario
	sHTML = ok
Else

	Generar = ""

	Addqry = ""
	response.expires = 0

	Server.scripttimeout = 60 * 6
	response.buffer = True



	XXXX = NEWEMPRE("NO")
	' SESSION("UN")=SEMPRE
	MasterType = "SCR"

	Mtype = request.querystring("MType") & ""
	If Mtype <> "" Then MasterType = Mtype & ""

	smenu = request.querystring("Menu")
	sorden = request.querystring("sorden")
	sNombre = request.querystring("Nombre")
	'response.write "STRING: " & SESSION("SETEMP") & "<br>" & "FORM: " & request.form & "<br>"

	SFORMEXP = ""


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

	Sfiltro = request.Form("sfiltro")

	If Len(Sfiltro) = 0 Then
		Sfiltro = request.querystring("sfiltro")
	End If

	If Len(Sfiltro) = 0 Then
		Sfiltro = "TODOS"
	End If

	Sfiltro = Replace(Sfiltro, "_", " ")
	CantRegMostrar = 10
	CantRegMover = CDbl("0" & request.querystring("CantRegMover"))
	ReDim Preserve bconsultar(4)
	OFV.ObtenerAtributosNodo Session("Form"), "", "ESTADO", "4"
	bconsultar(4) = OFV.AConsultar
	OFV.ObtenerAtributosNodo Session("Form"), "", "ESTADO", "2"
	bconsultar(2) = OFV.AConsultar
	OFV.ObtenerAtributosNodo Session("Form"), "", "ESTADO", "1"
	bconsultar(1) = OFV.AConsultar
	OFV.ObtenerAtributosNodo Session("Form"), "", "ESTADO", "3"
	bconsultar(3) = OFV.AConsultar
	SQLESTADOS = ""
	If smenu = "PORFECHA" Or smenu = "PORREVISION" Then
		sverestado = " and csporfechas.objnivel >= " & Session("SegNivel") & " And Estcod In (9,"
	ElseIf smenu = "PORSECTOR" Then
		SQLESTADOS = " AND c.estado In(9,"
		SQLNIVEL = " x.objnivel >=" & Session("SegNivel")


	Else
		'  SQLNIVEL=" x.objnivel >="  &  Session("SegNivel")
		sverestado = " and Objetos.objnivel >= " & Session("SegNivel") & " And Estados.objid In (9,"
	End If

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

	addfrom = ""
	'RESPONSE.WRITE sverestado & "<BR>"
	'RESPONSE.WRITE SQLESTADOS & "<BR>"
	''response.write smenu & " smenu " & chr(13)



	If smenu = "PORFECHA" Or smenu = "PORREVISION" Or smenu = "PORACTA"  Or smenu = "PORRESUMEN" Then


		If (Session("EMPMUN") = "S" And Session("GLBMUN") = 1) Or (Session("EMPMUN") = "N") Then
			addfrom = " ,objetos , tipobjeto  "
			sverestado = sverestado & " and csporfechas.CODOBJETO = objetos.CODOBJETO "
			sverestado = sverestado & " and csporfechas.tipoBJETO = tipobjeto.tipobjeto "
			sverestado = sverestado & " and objetos.OBJUN = " & Session("UN") & " "
		Else
			addfrom = " ,OBJMEMPALC ,objetos , tipobjeto "
			sverestado = sverestado & " and csporfechas.CODOBJETO = objetos.CODOBJETO "
			sverestado = sverestado & " and csporfechas.tipoBJETO = tipobjeto.tipobjeto "
			sverestado = sverestado & " and "
			sverestado = sverestado & " ( (objetos.OBJUN = " & Session("UN") & ") OR "
			sverestado = sverestado & "   ( objetos.OBJUN = 1 AND "
			sverestado = sverestado & "     csporfechas.CODOBJETO = OBJMEMPALC.CODOBJETO AND "
			sverestado = sverestado & "     OBJMEMPALC.CODEMPRE = " & Session("UN") & ") "
			sverestado = sverestado & " ) "
		End If
		If SESSION("BASE")="K" Then
			sverestado = sverestado & " and TIPOBJETO.TIPOCPT = 'K' "
		Else

			sverestado = sverestado & " and TIPOBJETO.TIPOCPT = 'D' "
		End If
		' response.write sverestado
		sverestado = sverestado & " and (objetos.APLIC = '" & wkaplic & "' or objetos.APLIC = 'GLB') "

	Else

		SQLSECTORUN = ""
		If (Session("EMPMUN") = "S" And Session("GLBMUN") = 1) Or (Session("EMPMUN") = "N") Then
			sverestado = sverestado & " and Objetos.OBJUN = " & Session("UN") & " "
			SQLSECTORUN = " x.OBJUN=" & Session("UN")
		Else
			addfrom = " ,OBJMEMPALC "
			sverestado = sverestado & " and "
			sverestado = sverestado & " ( (OBJETOS.OBJUN = " & Session("UN") & ") OR "
			sverestado = sverestado & "   ( "
			sverestado = sverestado & "     OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO AND "
			sverestado = sverestado & "     OBJMEMPALC.CODEMPRE = " & Session("UN") & ") "
			sverestado = sverestado & " ) "
			SQLSECTORUN = " ( (x.OBJUN=" & Session("UN") & ") OR "
			SQLSECTORUN = SQLSECTORUN & "( "
			SQLSECTORUN = SQLSECTORUN & "     X.CODOBJETO IN (SELECT DISTINCT CODOBJETO FROM OBJMEMPALC WHERE CODEMPRE= " & Session("UN") & ")))"
			'RESPONSE.WRITE SQLSECTORUN & "PPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPPP<BR>"
		End If
		If SESSION("BASE")="K" Then
			SQLTIPOCPT="e.TIPOCPT = 'K'"
		Else
			SQLTIPOCPT = "e.TIPOCPT = 'D'"
		End If
		SQLAPLIC = " (x.APLIC = '" & wkaplic & "' or x.APLIC = 'GLB')"
		If SESSION("BASE")="K" Then
			sverestado = sverestado & " and TIPOBJETO.TIPOCPT = 'K' "
		Else

			sverestado = sverestado & " and TIPOBJETO.TIPOCPT = 'D' "
		End If
		sverestado = sverestado & " and (OBJETOS.APLIC = '" & wkaplic & "' or OBJETOS.APLIC = 'GLB') "

	End If



	'''''''''''

	scal = request.querystring("scal")
	If request.Form.Count > 1 Then
		Robj = "F"
	Else
		Robj = "Q"
	End If


	'  response.write Request.Form
	scal = request.querystring("scal")
	'response.write smenu
	If smenu = "PORACTA"  Or smenu = "PORRESUMEN" Then

		If Robj = "F" Then

			sID1 = request.Form("Desde3") & ""
			'   sID1  = CLng(CDate(sID1)) - CLng(CDate("01-01-1900"))

			'   response.write request.form
			sID2 = request.Form("Hasta3") & ""
			'   sID2  = CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))

			Addqry = "&Desde=" & sID1 & "&Hasta=" & sID2
		Else
			sID1 = request.querystring("Desde") & ""
			sID2 = request.querystring("Hasta") & ""
			Addqry = "&Desde=" & sID1 & "&Hasta=" & sID2
		End If

		sMenuAdd = ""
		sMenuMas = "&smenu=" & smenu & "&Desde=" & sID1 & "&Hasta="
		sMenuMas = sMenuMas & sID2

		xsid = Replace(sID1, "/", "")
		xsid = Replace(xsid, "-", "")
		If (Not IsNumeric(xsid)) Or (sID1 = "") Then sID1 = "1-1-1950"

		xsid = Replace(sID2, "/", "")
		xsid = Replace(xsid, "-", "")
		If (Not IsNumeric(xsid)) Or (sID2 = "") Then sID2 = "31-12-2100"

	ElseIf smenu = "PORREVISION" Then

		If Robj = "F" Then

			sID1 = request.Form("Desde3") & ""

			sID2 = request.Form("Hasta3") & ""
			Addqry = "&Desde=" & sID1 & "&Hasta=" & sID2
		Else
			sID1 = request.querystring("Desde") & ""
			sID2 = request.querystring("Hasta") & ""
			Addqry = "&Desde=" & sID1 & "&Hasta=" & sID2
		End If

		sMenuAdd = ""
		sMenuMas = "&smenu=" & smenu & "&Desde=" & sID1 & "&Hasta="
		sMenuMas = sMenuMas & sID2

		xsid = Replace(sID1, "/", "")
		xsid = Replace(xsid, "-", "")
		If (Not IsNumeric(xsid)) Or (sID1 = "") Then sID1 = "1-1-1950"

		xsid = Replace(sID2, "/", "")
		xsid = Replace(xsid, "-", "")
		If (Not IsNumeric(xsid)) Or (sID2 = "") Then sID2 = "31-12-2100"
	ElseIf smenu = "CONSGRAL" Then

		'  Response.write Request.querystring

		FESTADO         = request.form("estado")
		FBANCA          = request.form("banca")
		FGUARDA         = request.form("TGUARDA")
		FPRODSERV       = request.form("PRODU-SERVI")
		FCATEGORIA      = request.form("CATEGORIAF")
		FCLASE          = request.form("CLASEF")
		FCODIGO         = request.form("CODIGOF")
		FFIRMAS         = request.form("FIRMASF")
		IMPRIME         = request.form("IMPRESION")
		TFECHA          = request.form("FECHA")
		FDESDE          = request.form("DESDE")
		FHASTA          = request.form("HASTA")
		FDESDE2         = request.form("DESDE2")
		FHASTA2         = request.form("HASTA2")

	ElseIf smenu = "PORFECHA" Then

		If Robj = "F" Then

			sID1 = request.Form("Desde3") & ""
			sID2 = request.Form("Hasta3") & ""
			sid3 = request.Form("Fecha") & ""
			Addqry = "&Desde=" & sID1 & "&Hasta=" & sID2 & "&Fecha=" & sid3
		Else
			sID1 = request.querystring("Desde") & ""
			sID2 = request.querystring("Hasta") & ""
			sid3 = request.querystring("Fecha") & ""
			Addqry = "&Desde=" & sID1 & "&Hasta=" & sID2 & "&Fecha=" & sid3
		End If

		sMenuAdd = ""
		sMenuMas = "&smenu=" & smenu & "&Desde=" & sID1 & "&Hasta="
		sMenuMas = sMenuMas & sID2 & "&Fecha=" & sid3

		xsid = Replace(sID1, "/", "")
		xsid = Replace(xsid, "-", "")
		If (Not IsNumeric(xsid)) Or (sID1 = "") Then sID1 = "1-1-1950"

		xsid = Replace(sID2, "/", "")
		xsid = Replace(xsid, "-", "")
		If (Not IsNumeric(xsid)) Or (sID2 = "") Then sID2 = "31-12-2100"

	ElseIf smenu = "PORPRODUCTO" Then

		If Robj = "F" Then

			sID1 = request.Form("ListaS") & ""
			sID2 = request.Form("ListaG") & ""
			sid3 = request.Form("ListaP") & ""
			Addqry = "&ListaS=" & sID1 & "&ListaG=" & sID2 & "&ListaP=" & sid3
			sID1 = Replace(sID1, "_", " ")
			sID2 = Replace(sID2, "_", " ")
			sid3 = Replace(sid3, "_", " ")

		Else
			sID1 = request.querystring("ListaS") & ""
			sID2 = request.querystring("ListaG") & ""
			sid3 = request.querystring("ListaP") & ""
			Addqry = "&ListaS=" & sID1 & "&ListaG=" & sID2 & "&ListaP=" & sid3
			sID1 = Replace(sID1, "_", " ")
			sID2 = Replace(sID2, "_", " ")
			sid3 = Replace(sid3, "_", " ")

		End If

	ElseIf smenu = "PORNROCIRCULAR" Then

		If Robj = "F" Then

			sID1 = request.Form("ListaG") & ""
			sID2 = request.Form("ListaN") & ""
			Addqry = "&ListaG=" & sID1 & "&ListaN=" & sID2
			sID1 = Replace(sID1, "_", " ")
			sID2 = Replace(sID2, "_", " ")
		Else

			sID1 = request.querystring("ListaG") & ""
			sID2 = request.querystring("ListaN") & ""
			Addqry = "&ListaG=" & sID1 & "&ListaN=" & sID2
			sID1 = Replace(sID1, "_", " ")
			sID2 = Replace(sID2, "_", " ")

		End If

	ElseIf smenu = "PORINSTRUMENTO" Then

		If Robj = "F" Then

			sID1 = request.Form("Lista") & ""
			sID2 = request.Form("buscar") & ""
			Addqry = "&Lista=" & sID1 & "&buscar=" & sID2
			sID2 = Replace(sID2, "_", " ")

		Else

			sID1 = request.querystring("Lista") & ""
			sID2 = request.querystring("buscar") & ""
			Addqry = "&Lista=" & sID1 & "&buscar=" & sID2
			sID2 = Replace(sID2, "_", " ")

		End If


	ElseIf smenu = "PORTEXTO" Then


		If Robj = "F" Then


			sID1 = Replace(request.Form("buscar") & "", "'", "")
			sID1 = Replace(sID1, Chr(34), "")

			sID2 = request.Form("lista") & ""

			tsel = request.Form("tsel") & ""
			tgrp = request.Form("tgrp") & ""
			tcod = request.Form("tcod") & ""
			ttip = request.Form("ttip") & ""
			tnam = request.Form("tnam") & ""

			tdoc = request.Form("tdoc") & ""


			Addqry = "&buscar=" & Replace(sID1, " ", "_")

			Addqry = Addqry & "&Lista=" & sID2

			Addqry = Addqry & "&tsel=" & tsel & "&tgrp=" & tgrp & "&tcod=" & tcod & "&ttip=" & ttip & "&tnam=" & tnam

			Addqry = Addqry & "&tdoc=" & tdoc

			sID1 = Replace(sID1, "_", " ")

		Else

			sID1 = request.querystring("buscar") & ""

			sID2 = request.querystring("Lista") & ""

			tsel = request.querystring("tsel") & ""
			tgrp = request.querystring("tgrp") & ""
			tcod = request.querystring("tcod") & ""
			ttip = request.querystring("ttip") & ""
			tnam = request.querystring("tnam") & ""

			tdoc = request.querystring("tdoc") & ""

			Addqry = "&buscar=" & sID1

			Addqry = Addqry & "&Lista=" & sID2

			Addqry = Addqry & "&tsel=" & tsel & "&tgrp=" & tgrp & "&tcod=" & tcod & "&ttip=" & ttip & "&tnam=" & tnam

			Addqry = Addqry & "&tdoc=" & tdoc

			sID1 = Replace(sID1, "_", " ")

		End If



		' sMenuAdd = ""
		' sMenuMas = "&smenu=" & smenu & "&Instrumento=" & sID1 & "&buscar="
		' sMenuMas = SmenuMas & sID2

		sMenuAdd = ""
		sMenuMas = "&smenu=" & smenu & "&buscar=" & sID1 & "&Lista="
		sMenuMas = sMenuMas & sID2 & "&tsel=" & tsel & "&tgrp=" & tgrp & "&tcod=" & tcod & "&ttip=" & ttip & "&tnam=" & tnam
		sMenuMas = sMenuMas & "&tdoc=" & tdoc


	ElseIf smenu = "PORCATEGO" Then


		If Robj = "F" Then


			sID1 = Replace(request.Form("buscar") & "", "'", "")
			sID1 = Replace(sID1, Chr(34), "")

			sID2 = request.Form("lista") & ""

			tsel = request.Form("tsel") & ""
			tgrp = request.Form("tgrp") & ""
			tcod = request.Form("tcod") & ""
			ttip = request.Form("ttip") & ""
			tnam = request.Form("tnam") & ""

			tdoc = request.Form("tdoc") & ""


			Addqry = "&buscar=" & Replace(sID1, " ", "_")

			Addqry = Addqry & "&Lista=" & sID2

			Addqry = Addqry & "&tsel=" & tsel & "&tgrp=" & tgrp & "&tcod=" & tcod & "&ttip=" & ttip & "&tnam=" & tnam

			Addqry = Addqry & "&tdoc=" & tdoc

			sID = Replace(sID1, "_", " ")

		Else

			sID1 = request.querystring("buscar") & ""

			sID2 = request.querystring("Lista") & ""

			tsel = request.querystring("tsel") & ""
			tgrp = request.querystring("tgrp") & ""
			tcod = request.querystring("tcod") & ""
			ttip = request.querystring("ttip") & ""
			tnam = request.querystring("tnam") & ""

			tdoc = request.querystring("tdoc") & ""

			Addqry = "&buscar=" & sID1

			Addqry = Addqry & "&Lista=" & sID2

			Addqry = Addqry & "&tsel=" & tsel & "&tgrp=" & tgrp & "&tcod=" & tcod & "&ttip=" & ttip & "&tnam=" & tnam

			Addqry = Addqry & "&tdoc=" & tdoc

			sID = Replace(sID1, "_", " ")

		End If



		' sMenuAdd = ""
		' sMenuMas = "&smenu=" & smenu & "&Instrumento=" & sID1 & "&buscar="
		' sMenuMas = sMenuMas & sID2

		sMenuAdd = ""
		sMenuMas = "&smenu=" & smenu & "&buscar=" & sID1 & "&Lista="
		sMenuMas = sMenuMas & sID2 & "&tsel=" & tsel & "&tgrp=" & tgrp & "&tcod=" & tcod & "&ttip=" & ttip & "&tnam=" & tnam
		sMenuMas = sMenuMas & "&tdoc=" & tdoc


	ElseIf smenu = "PORIDENTIFICACION" Then

		If Robj = "F" Then

			sID1 = request.Form("buscar") & ""
			Addqry = "&buscar=" & sID1

		Else
			sID1 = request.querystring("buscar") & ""
			Addqry = "&buscar=" & sID1
		End If

		sMenuAdd = ""
		sMenuMas = "&smenu=" & smenu & "&IDENTIFICACION=" & sID

	ElseIf smenu = "PORESTADO" Then



		If Robj = "F" Then
			sID = request.Form("lista")
			sID1 = request.Form("Desde3") & ""

			sID2 = request.Form("Hasta3") & ""
			'cdesde = CLng(CDate(sID1)) - CLng(CDate("01-01-1900"))
			'chasta = CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))

			Addqry = "&Desde=" & sID1 & "&Hasta=" & sID2 & "&lista=" & sID

			'        response.write Addqry & " Addqry " & "<br>"

		Else

			sID = request.querystring("lista")
			sID1 = request.querystring("Desde") & ""
			sID2 = request.querystring("Hasta") & ""
			'cdesde = CLng(CDate(sID1)) - CLng(CDate("01-01-1900"))
			'chasta = CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))

			Addqry = "&Desde=" & sID1 & "&Hasta=" & sID2 & "&lista=" & sID

		End If

		sMenuAdd = ""
		sMenuMas = "&smenu=" & smenu & "&Desde=" & sID1 & "&Hasta="
		sMenuMas = sMenuMas & sID2 & "&lista=" & sID

	ElseIf smenu = "PORAVENCER" Then



		If Robj = "F" Then
			sID = request.Form("lista")
			sID2 = Date()
			chasta = CLng(sID2) - CLng(CDate("01-01-1900"))

			Addqry = "Hasta=" & sID2 & "&lista=" & sID

			'        response.write Addqry & " Addqry " & "<br>"

		Else

			sID = request.querystring("lista")
			sID1 = request.querystring("Desde") & ""
			sID2 = request.querystring("Hasta") & ""
			'cdesde = CLng(CDate(sID1)) - CLng(CDate("01-01-1900"))
			'chasta = CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))

			Addqry = "&Desde=" & sID1 & "&Hasta=" & sID2 & "&lista=" & sID

		End If

		sMenuAdd = ""
		sMenuMas = "&smenu=" & smenu & "&Desde=" & sID1 & "&Hasta="
		sMenuMas = sMenuMas & sID2 & "&lista=" & sID



	ElseIf smenu = "PORCOMPONENTE" Then

		If Robj = "F" Then


			sID2 = request.Form("lista") & ""
			sid3 = request.Form("buscar") & ""
			If sid3 = "" Then sid3 = sID2

			Addqry = "&Lista=" & sID2 & "&buscar=" & sid3  '& "&perfil=" & SPerfil

			SPerfil = Replace(SPerfil, "_", " ")

		Else

			SPerfil = request.querystring("perfil") & ""
			sID1 = request.querystring("Lista") & ""
			sID2 = request.querystring("Lista") & ""
			sid3 = request.querystring("buscar") & ""
			Addqry = "&Lista=" & sID2 & "&buscar=" & sid3  '& "&perfil=" & SPerfil
			SPerfil = Replace(SPerfil, "_", " ")

		End If

		sID1 = OFV.convertircar(sID1, "_", " ")
		If sID1 = "" Then sID1 = sid3

		sMenuAdd = ""
		sMenuMas = "&smenu=" & smenu & "&lista=" & sID2 & "&buscar="
		sMenuMas = sMenuMas & sid3 & "&perfil=" & SPerfil

		' response.write sID2  & " lista " & "<br>"
		' response.write sID3  & " buscar " & "<br>"
		'  response.write SPerfil  & " perfil " & "<br>"

	Else


		If Robj = "F" Then

			SPerfil = request.Form("perfil") & ""
			sID1 = request.Form("lista") & ""
			sID2 = request.Form("lista") & ""
			sid3 = request.Form("buscar") & ""

			Addqry = "&Lista=" & sID2 & "&buscar=" & sid3 & "&perfil=" & SPerfil

			SPerfil = Replace(SPerfil, "_", " ")

		Else

			SPerfil = request.querystring("perfil") & ""
			sID1 = request.querystring("Lista") & ""
			sID2 = request.querystring("Lista") & ""
			sid3 = request.querystring("buscar") & ""
			Addqry = "&Lista=" & sID2 & "&buscar=" & sid3 & "&perfil=" & SPerfil
			SPerfil = Replace(SPerfil, "_", " ")

		End If

		sID1 = OFV.convertircar(sID1, "_", " ")
		If sID1 = "" Then sID1 = sid3

		sMenuAdd = ""
		sMenuMas = "&smenu=" & smenu & "&lista=" & sID2 & "&buscar="
		sMenuMas = sMenuMas & sid3 & "&perfil=" & SPerfil

		' response.write sID2  & " lista " & "<br>"
		' response.write sID3  & " buscar " & "<br>"
		' response.write SPerfil  & " perfil " & "<br>"

	End If

	Addqry = "&Menu=" & smenu & Addqry & "&Nombre=" & sNombre

	starget = "ListadoConsultas.asp?cantregmover=0&cantregmostrar=" & CantRegMostrar & Addqry

	'response.write Addqry  & "<br>"
	'response.write starget & "<br>"

	sVolver = "ListadoConsultas.asp?Menu=" & smenu & sMenuAdd

	Addqry = Addqry & "&sorden=" & sorden & "&sfiltro=" & Replace(Sfiltro, " ", "_")

	'response.write Addqry

	Set cn = OFV.conectar(OFV.strconn2)
	Set cnx = OFV.conectar(OFV.strconn5)

	cn.CommandTimeout = 0
	strsql = "SELECT DISTINCT TIPOBJETO.descri, OBJETOS.codiden, OBJETOS.descri, "
	strsql = strsql & "OBJETOS.codobjeto, OBJETOS.tipobjeto, OBJPRTY.fecemis, "
	strsql = strsql & "OBJPRTY.fecvigen, OBJPRTY.fecbaja, Estados.estado as estdescri, "
	strsql = strsql & "OBJETOS.codiden, OBJETOS.descri, TIPOBJETO.archivo, OBJPRTY.autor, "
	strsql = strsql & "OBJPRTY.aprobo, OBJPRTY.elimino, TIPOBJETO.perfil, OBJPRTY.fecvto,"
	strsql = strsql & "Estados.estado as estcod,OBJETOS.objnivel,OBJPRTY.Version,OBJPRTY.nrover,"
	strsql = strsql & "OBJETOS.OBJUN,OBJPRTY.fecpublic "
	strsql = strsql & ", t.descri, t.descri, t.descri, t.descri, st.descri   "


	addfrom = addfrom & ", csrelpath a, tema t, prodserv  st "

	wherefrom = " and (a.cobj = OBJETOS.codobjeto) "
	wherefrom = wherefrom & " and (a.tema = t.codtema) "
	wherefrom = wherefrom & " and (a.stema = st.codprodserv) "
	''response.write wherefrom

	'RESPONSE.WRITE strsql & " smenu " & chr(13)


	Select Case smenu
		Case "PORCOMPONENTE"



		strsql = strsql & " FROM "
		strsql = strsql & " objetos "

		strsql = strsql & " inner join  OBJPRTY      on (OBJETOS.codobjeto = OBJPRTY.codobjeto)"
		strsql = strsql & " inner join  Estados      on (OBJPRTY.estado = Estados.objid)"
		If SESSION("BASE")="K" Then
			strsql = strsql & " inner join  TIPOBJETO    on (OBJETOS.TIPOBJETO = TIPOBJETO.tipobjeto AND TIPOBJETO.TIPOCPT='K' )"
		Else
			strsql = strsql & " inner join  TIPOBJETO    on (OBJETOS.TIPOBJETO = TIPOBJETO.tipobjeto AND TIPOBJETO.TIPOCPT <> 'K' )"
		End If
		strsql = strsql & " inner join  OBJlnk       on (objetos.codobjeto=OBJlnk.codobjeto and objlnk.ltipo = 'KL') "
		strsql = strsql & " inner join  COMPLEMENTOS on (complementos.codcomp = objlnk.codobjetol and Complementos.descri like '"& sid3& "')"

		strsql = strsql & " left join objestdet on (OBJETOS.codobjeto = objestdet.codobjeto and objestdet.dettipo = 'J' and objlnk.objid = objestdet.detimagen) "
		strsql = strsql & " left join OBJACCION on (OBJETOS.codobjeto = OBJACCION.codobjeto and objlnk.objid = OBJACCION.codcomp) "

		strsql = strsql & " inner join csrelpath a on(a.cobj = OBJETOS.codobjeto) "
		strsql = strsql & " inner join tema t on (a.tema = t.codtema)"
		strsql = strsql & " inner join prodserv st  on (a.stema = st.codprodserv)"
		strsql = strsql & " LEFT join OBJMEMPALC "
		strsql = strsql & " on OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO "



		sTitulo = "Componente: " & sID1




		Case "PORESTADO"

		'' response.write Robj & " Robj " & "<br>"

		sID = request.Form("lista")

		If Robj = "F" Then

			sID1 = request.Form("Desde3") & ""

			sID2 = request.Form("Hasta3") & ""
			'cdesde = CLng(CDate(sID1)) - CLng(CDate("01-01-1900"))
			'chasta = CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))

			Addqry = "&Menu=PORESTADO&NOMBRE=" & nEstado &"&Desde=" & sID1 & "&Hasta=" & sID2 & "&lista=" & sID

			'        response.write Addqry & " Addqry " & "<br>"

		Else

			sID = request.querystring("lista")
			sID1 = request.querystring("Desde") & ""
			sID2 = request.querystring("Hasta") & ""
			' cdesde = CLng(CDate(sID1)) - CLng(CDate("01-01-1900"))
			' chasta = CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))

			Addqry = "&Menu=PORESTADO&NOMBRE=" & nEstado &"&Desde=" & sID1 & "&Hasta=" & sID2 & "&lista=" & sID & "&sorden=" & sorden
			'response.write addqry
		End If

		sMenuAdd = ""
		sMenuMas = "&smenu=" & smenu & "&Desde=" & sID1 & "&Hasta="
		sMenuMas = sMenuMas & sID2 & "&lista=" & sID
		'       response.write smenu & " smenu " & chr(13)

		ssEstado = ""

		strsql = strsql & "FROM OBJETOS, TIPOBJETO, OBJPRTY, Estados " & addfrom
		strsql = strsql & " where  "
		If SESSION("BASE")="K" Then
			strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and   TIPOBJETO.TIPOCPT='K' AND "
		Else
			strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto AND   TIPOBJETO.TIPOCPT <> 'K' and "
		End If
		strsql = strsql & "OBJPRTY.estado = Estados.objid and "
		strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto "
		strsql = strsql & "and OBJPRTY.estado = " & sID
		strsql = strsql & " and (Objprty.fecemis between '" & cdesde & "' and '" & chasta & "' ) "

		sTitulo = "Estado: " & sID
		' response.write strsql & " strsql QQQQQQQQQQQQQQQ" & chr(13)
		' response.write sID & " sID " & chr(13)
		sqle = "Select * from estados where objid =  " & sID
		Set rsx = OFV.crearconsultaEx(sqle, cn, 1, "")
		If Not rsx.EOF Then
			sTitulo = "Estado: " & rsx(1)
			ssEstado = rsx(1)
		End If

		swDesde2 = CDate(CLng(cdesde) + CLng(CDate("01-01-1900")))
		swHasta2 = CDate(CLng(chasta) + CLng(CDate("01-01-1900")))


		sTitulo = "Por estado:  " & ssEstado ' & "&nbsp;&nbsp; Fecha de creación Desde: " & swDesde2 & "&nbsp;&nbsp;|&nbsp; Fecha de creación Hasta:&nbsp;" & swHasta2

		Case "PORTEXTO"

		htwhere = ""

		strsql = strsql & "FROM OBJETOS, TIPOBJETO, OBJPRTY, Estados "

		If sID2 <> "99999" Then

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
		If SESSION("BASE")="K" Then
			strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and TIPOBJETO.TIPOCPT='K' AND "
		Else
			strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and TIPOBJETO.TIPOCPT <> 'K' AND "
		End If
		strsql = strsql & "OBJPRTY.estado = Estados.objid and "
		strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto "

		'   response.write sid2
		If sID2 <> "99999" Then

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

		nwsID = vertexto(sID1)


		strsql = strsql & " and (OBJETOS.descriabr  COLLATE SQL_Latin1_General_CP1_CI_AI like '%" & nwsID & "%' "
		strsql = strsql & " or OBJETOS.descri   COLLATE SQL_Latin1_General_CP1_CI_AI like '%" & nwsID & "%') "

		'Response.write nwsid
		sTitulo = "Texto: " & sID1

		If sID2 <> "99999" Then

			sqle = "Select descri from eco where eco =0" & sID2
			Set rsx = OFV.crearconsultaEx(sqle, cn, 1, "")
			If Not rsx.EOF Then
				sTitulo = sTitulo & "&nbsp;|&nbsp;Sector: " & rsx(0)
			Else
				sTitulo = sTitulo & "&nbsp;|&nbsp;Sector: "
			End If
			Call OFV.cerrarconsulta(rsx)
		End If

		If tsel <> "T" And tgrp <> "M" Then

			sTitulo = sTitulo & "&nbsp;|&nbsp;Carpeta:&nbsp;" & Replace(tnam, "_", "&nbsp;")


		End If

		Case "PORCATEGO"

		htwhere = ""

		strsql = strsql & "FROM OBJETOS, TIPOBJETO, OBJPRTY, Estados, OBJRSKREL "





		'''''EVENRSK



		strsql = strsql & addfrom
		strsql = strsql & " where  "
		If SESSION("BASE")="K" Then
			strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto AND  TIPOBJETO.TIPOCPT='K' and "
		Else
			strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto AND  TIPOBJETO.TIPOCPT <> 'K' and "
		End If
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
		'strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and "
		If SESSION("BASE")="K" Then
			strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto AND  TIPOBJETO.TIPOCPT='K' and "
		Else
			strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto AND  TIPOBJETO.TIPOCPT <> 'K' and "
		End If
		strsql = strsql & "OBJPRTY.estado = Estados.objid and "
		strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto "
		strsql = strsql & "   "

		If UCase(sid3 & "") <> UCase("nsnc") Then
			strsql = strsql & " AND OBJETOS.TSEGMENTO =" & CDbl("0" & sid3) & " "
		End If

		If UCase(sID1 & "") <> UCase("nsnc") Or UCase(sID2 & "") <> UCase("nsnc") Then
			strsql = strsql & " AND ((OBJETOS.codobjeto = OBJTSEGAGR.codobjeto)  AND ( "
		End If

		If UCase(sID1 & "") <> UCase("nsnc") Then
			strsql = strsql & " (OBJTSEGAGR.TSEGMENTO = " & CDbl("0" & sID1) & " ) "
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
		' strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and "
		If SESSION("BASE")="K" Then
			strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto AND  TIPOBJETO.TIPOCPT='K' and "
		Else
			strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto AND  TIPOBJETO.TIPOCPT <> 'K' and "
		End If
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
			'   response.write sID2
			If UCase(sID1 & "") <> UCase("nsnc") Then
				Sqlx = "SELECT descri from tipoAGRUPA where objid=" & CDbl("0" & sID1)
				Set rsx = OFV.crearconsultaEx(Sqlx, cn, 1, "")
				If Not rsx.EOF Then xinstrum2 = rsx(0) & ""
				Call OFV.cerrarconsulta(rsx)
			End If
		End If
		'' xinstrum2 = ""
		sTitulo = " GRUPO :" & xinstrum2


		Case "PORINSTRUMENTO"



		strsql = strsql & " FROM OBJETOS  "
		strsql = strsql & " INNER JOIN OBJPRTY on "
		strsql = strsql & " OBJETOS.codobjeto = OBJPRTY.codobjeto "
		strsql = strsql & " INNER JOIN TIPOBJETO ON "
		'strsql = strsql & " OBJETOS.tipobjeto = TIPOBJETO.tipobjeto "
		If SESSION("BASE")="K" Then
			strsql = strsql & " (OBJETOS.tipobjeto = TIPOBJETO.tipobjeto AND  TIPOBJETO.TIPOCPT='K') "
		Else
			strsql = strsql & "( OBJETOS.tipobjeto = TIPOBJETO.tipobjeto AND  TIPOBJETO.TIPOCPT <> 'K') "
		End If
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

		strsql = strsql & " WHERE  "

		strsql = strsql & " ( "

		If ("0" & sID1) <> "0" Then


			If sID2 <> "" Then
				strsql = strsql & "  ((OBJPRTY.instrumen=" & CDbl("0" & sID1) & ")"
			Else
				strsql = strsql & "  (OBJPRTY.instrumen=" & CDbl("0" & sID1) & " or OBJPRTYADV.instrua=" & CDbl("0" & sID1) & ")  "
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
		Sqlx = "SELECT descri from instrumen where objid=" & CDbl("0" & sID1)
		Set rsx = OFV.crearconsultaEx(Sqlx, cn, 1, "")
		If Not rsx.EOF Then xinstrum = rsx(0)
		Call OFV.cerrarconsulta(rsx)
		sTitulo = "INSTRUMENTO " & xinstrum & " " & sID2

		Case "PORIDENTIFICACION"

		strsql = strsql & "FROM OBJETOS, TIPOBJETO, OBJPRTY, Estados " & addfrom
		strsql = strsql & " where   "
		'strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and "

		If SESSION("BASE")="K" Then
			strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto AND  TIPOBJETO.TIPOCPT='K' and "
		Else
			strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto AND  TIPOBJETO.TIPOCPT <> 'K' and "
		End If
		strsql = strsql & "OBJPRTY.estado = Estados.objid and "
		strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto "
		strsql = strsql & "and OBJETOS.Codiden = '" & sID1 & "' "
		sTitulo = "IDENTIFICACION " & sID1

		Case "PORSECTOR"
		strsql = " SELECT dISTINCT e.descri, x.codiden, x.descri, x.codobjeto, x.tipobjeto,"
		strsql = strsql & "c.fecemis, c.fecvigen, c.fecbaja, c1.estado as estdescri,"
		strsql = strsql & "  x.codiden, x.descri, e.archivo, c.autor, c.aprobo,"
		strsql = strsql & "   c.elimino, e.perfil, c.fecvto,c1.objid as estcod,"
		strsql = strsql & "   x.objnivel,c.Version,c.nrover,x.OBJUN,c.fecpublic ,"
		strsql = strsql & "   d.temadescri, d.temadescri, d.temadescri, d.temadescri, d.stemadescri "

		Select Case SPerfil
			Case "nsnc"

			strsql = strsql & "  FROM "
			strsql = strsql & "    ( "
			strsql = strsql & "    select codobjeto,tipobjeto from objcontrol "

			strsql = strsql & "     where codsector='0" & sID1 & "'"


			strsql = strsql & "     union "

			strsql = strsql & "      select a.codobjeto,a.tipobjeto "

			strsql = strsql & "      from  "
			strsql = strsql & "      OBJPROCED a "
			strsql = strsql & "      inner join objsector b on ( b.objid=a.codsector)"
			strsql = strsql & "      inner join  tipoperfil c on (c.tipoperfil=a.codperfil)"

			strsql = strsql & "      where b.codsector='0" & sID1 & "'"
			strsql = strsql & " union select distinct codobjeto ,tipobjeto from objsector where objsector.codsector='" & sID1 & "' "
			strsql = strsql & "     )  x2 "
			strsql = strsql & "INNER JOIN OBJETOS X "
			strsql = strsql & " on ("

			strsql = strsql & "  (x2.codobjeto=x.codobjeto and x2.tipobjeto=x.tipobjeto) "
			strsql = strsql & "  and "
			strsql = strsql & SQLAPLIC         '   " (x.APLIC = 'EM' or x.APLIC = 'GLB')"
			strsql = strsql & "  and"

			strsql = strsql & SQLNIVEL    '"  x.objnivel >= 3 "
			strsql = strsql & "  and"

			strsql = strsql & SQLSECTORUN & " )"

			strsql = strsql & " inner join objprty c on (c.codobjeto=x.codobjeto " & SQLESTADOS & ") "
			strsql = strsql & " inner join estados c1 on (c1.objid= c.estado)"
			strsql = strsql & " inner join csrelpathdescriex d on(d.cobj=c.codobjeto)"

			'  IF SESSION("BASE")="K" THEN
			'	  strsql = strsql & " inner join  TIPOBJETO    on (x.TIPOBJETO = TIPOBJETO.tipobjeto AND TIPOBJETO.TIPOCPT='K' )"
			'	  ELSE
			'	   strsql = strsql & " inner join  TIPOBJETO    on (x.TIPOBJETO = TIPOBJETO.tipobjeto AND TIPOBJETO.TIPOCPT <> 'K' )"
			'     END IF
			If session("base")="K" Then
				strsql = strsql & " inner join TIPOBJETO e on(e.tipobjeto=x.tipobjeto and e.TIPOCPT = 'K')"
			Else
				strsql = strsql & " inner join TIPOBJETO e on(e.tipobjeto=x.tipobjeto and e.TIPOCPT = 'D')"
			End If


			Case Else
			strsql = strsql & "  FROM "
			strsql = strsql & "    ( "
			strsql = strsql & "    select codobjeto,tipobjeto from objcontrol "

			strsql = strsql & "     where codsector= " & sID1 & " AND  codperfil='" & SPerfil & "'"


			strsql = strsql & "     union "

			strsql = strsql & "      select a.codobjeto,a.tipobjeto "

			strsql = strsql & "      from  "
			strsql = strsql & "      OBJPROCED a "
			strsql = strsql & "      inner join objsector b on ( b.objid=a.codsector)"
			strsql = strsql & "      inner join  tipoperfil c on (c.tipoperfil=a.codperfil)"

			strsql = strsql & "      where b.codsector=" & sID1 & " AND c.tipoperfil='" & SPerfil & "'"
			strsql = strsql & "     )  x2 "
			strsql = strsql & "INNER JOIN OBJETOS X "
			strsql = strsql & " on ("

			strsql = strsql & "  (x2.codobjeto=x.codobjeto and x2.tipobjeto=x.tipobjeto) "
			strsql = strsql & "  and "
			'strsql = strsql & " (x.APLIC = 'EM' or x.APLIC = 'GLB')"
			strsql = strsql & SQLAPLIC
			strsql = strsql & "  and"

			strsql = strsql & SQLNIVEL    '"  x.objnivel >= 3 "
			strsql = strsql & "  and"

			strsql = strsql & SQLSECTORUN & " )"

			strsql = strsql & " inner join objprty c on (c.codobjeto=x.codobjeto " & SQLESTADOS & ") "
			strsql = strsql & " inner join estados c1 on (c1.objid= c.estado)"
			strsql = strsql & " inner join csrelpathdescriex d on(d.cobj=c.codobjeto)"

			If session("base")="K" Then
				strsql = strsql & " inner join TIPOBJETO e on(e.tipobjeto=x.tipobjeto and e.TIPOCPT = 'K')"
			Else
				strsql = strsql & " inner join TIPOBJETO e on(e.tipobjeto=x.tipobjeto and e.TIPOCPT = 'D')"
			End If

		End Select

		sqle = "Select descri from eco where eco =0" & sID1
		Set rsx = OFV.crearconsultaEx(sqle, cn, 1, "")
		If Not rsx.EOF Then
			sTitulo = "Sector:&nbsp;" & rsx(0)
		Else
			sTitulo = "Sector:&nbsp;"
		End If



		Call OFV.cerrarconsulta(rsx)

		''     if  SPerfil <> "nsnc" then sTitulo = sTitulo & "&nbsp;|&nbsp;Perfil:&nbsp;" & SPerfil
		If SPerfil <> "nsnc" Then

			'' AGREGADO 26052014
			sTitulo = sTitulo & "&nbsp;"

			Strsqle = "Select descri from tipoperfil where Tipoperfil ='" & SPerfil & "' "
			Set rsxx = OFV.crearconsultaEx(Strsqle, cn, 1, "")
			If Not rsxx.EOF Then
				sTitulo = sTitulo & "&nbsp;|&nbsp;Perfil:&nbsp;" & rsxx(0)
			Else
				sTitulo = sTitulo & "&nbsp;|&nbsp;Perfil:&nbsp;" & "TODOS"
			End If

			Call OFV.cerrarconsulta(rsxx)

		Else
			sTitulo = sTitulo & "&nbsp;|&nbsp;Perfil:&nbsp;" & "TODOS"
		End If

		'' AGREGADO 26052014 Fin

		Case "PORTEMA"
		sql = "Select descri From Objetos Where codobjeto = " & sID1
		Set rsObj = OFV.crearconsultaEx(sql, cn, 1, "")
		If Not rsObj.EOF Then sObj = rsObj(0)
		Call OFV.cerrarconsulta(rsObj)

		strsql = strsql & "FROM OBJETOS, TIPOBJETO, OBJPRTY, Estados " & addfrom
		strsql = strsql & " where    "
		' strsql = strsql & "OBJETOS.tipobjeto=TIPOBJETO.tipobjeto and "
		If SESSION("BASE")="K" Then
			strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto AND  TIPOBJETO.TIPOCPT='K' and "
		Else
			strsql = strsql & "OBJETOS.tipobjeto = TIPOBJETO.tipobjeto AND  TIPOBJETO.TIPOCPT <> 'K' and "
		End If
		strsql = strsql & "OBJPRTY.estado = Estados.objid and "
		strsql = strsql & "OBJETOS.codobjeto = OBJPRTY.codobjeto "
		strsql = strsql & "and Objetos.codobjeto = " & sID1
		sTitulo = "Tema: " & sObj
		Case "CONSGRAL"



		NEWSTRSQL = ""

		NEWSTRSQL= " SELECT * FROM OBJETOS A "
		NEWSTRSQL= NEWSTRSQL & " INNER JOIN OBJPROPIEDADES B ON "
		NEWSTRSQL= NEWSTRSQL & " (A.CODOBJETO=B.CODOBJETO AND tiempoguarda='"& FGUARDA & "'"
		NEWSTRSQL= NEWSTRSQL & " AND catForm='" & FCATEGORIA & "' AND claseform='"& FCLASE &"' and CODIGOFORM='"& FCODIGO &"' "
		NEWSTRSQL= NEWSTRSQL & " AND FIRMASFORM='"& FFIRMAS &"') "
		NEWSTRSQL= NEWSTRSQL & " INNER JOIN TIPOBJETO C ON (A.TIPOBJETO=C.TIPOBJETO AND TIPOCPT='K') "
		NEWSTRSQL= NEWSTRSQL & " INNER JOIN OBJPRTY D ON (D.CODOBJETO=A.CODOBJETO AND"


		If TFECHA = "Emision" Then

			NEWSTRSQL= NEWSTRSQL & " D.FECEMIS BETWEEN '"& FDESDE &"' AND '"& FHASTA &"' "

		ElseIf TFECHA = "Vigencia" Then

			NEWSTRSQL= NEWSTRSQL & " D.FECVIGEN BETWEEN '"& FDESDE &"' AND '"& FHASTA &"' "

		ElseIf TFECHA = "Eliminacion" Then

			NEWSTRSQL= NEWSTRSQL & " D.FECBAJA BETWEEN '"& FDESDE &"' AND '"& FHASTA &"' "

		ElseIf TFECHA = "Vencimiento" Then

			NEWSTRSQL= NEWSTRSQL & " D.FECVTO BETWEEN '"& FDESDE &"' AND '"& FHASTA &"' "

		ElseIf TFECHA = "Publicacion" Then

			NEWSTRSQL= NEWSTRSQL & " D.FECPUBLIC BETWEEN '"& FDESDE &"' AND '"& FHASTA &"' "

		End If
		NEWSTRSQL= NEWSTRSQL & "  and d.estado='"& FESTADO &"') "
		NEWSTRSQL= NEWSTRSQL & " INNER JOIN OBJTSEGAGR F ON (F.CODOBJETO=A.codobjeto AND F.TSEGMENTO='"& FBANCA &"')--BANCA "
		NEWSTRSQL= NEWSTRSQL & " WHERE A.TSEGMENTO='"& FPRODSERV &"'--PRODU-SERVI "

		Set rsx1 = OFV.crearconsultaEx(NEWSTRSQL, cn, 1, "")

		'Response.write NEWSTRSQL


		Case "PORFECHA"
		strsql = "select * from csporfechas " & addfrom
		cdesde = CLng(CDate(sID1)) - CLng(CDate("01-01-1900"))
		chasta = CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))
		If sid3 = "Emision" Then
			strsql = strsql & "WHERE   xfecemis between '" & cdesde & "' and '" & chasta & "' "
			sTitulo = "Creacion&nbsp;&nbsp;Del:&nbsp;" & sID1 & "&nbsp;&nbsp;|&nbsp;al:&nbsp;" & sID2
		ElseIf sid3 = "Vigencia" Then
			strsql = strsql & "WHERE   xfecvigen between '" & cdesde & "' and '" & chasta & "' "
			sTitulo = "vigencia&nbsp;&nbsp;Del:&nbsp;" & sID1 & "&nbsp;&nbsp;|&nbsp;al:&nbsp;" & sID2
		ElseIf sid3 = "Eliminacion" Then
			strsql = strsql & "WHERE   xfecbaja between '" & cdesde & "' and '" & chasta & "' "
			sTitulo = "Eliminacion&nbsp;&nbsp;Del:&nbsp;" & sID1 & "&nbsp;&nbsp;|&nbsp;al:&nbsp;" & sID2
		ElseIf sid3 = "Vencimiento" Then
			strsql = strsql & "WHERE  xfecvto between '" & cdesde & "' and '" & chasta & "' "
			sTitulo = "Revision&nbsp;&nbsp;Del:&nbsp;" & sID1 & "&nbsp;&nbsp;|&nbsp;al:&nbsp;" & sID2
		ElseIf sid3 = "Publicacion" Then
			strsql = strsql & "WHERE   xfecpublic between '" & cdesde & "' and '" & chasta & "' "
			sTitulo = "Publicacion&nbsp;&nbsp;Del:&nbsp;" & sID1 & "&nbsp;&nbsp;|&nbsp;al:&nbsp;" & sID2
		End If

		Case "PORREVISION"
		strsql = "select * from csporfechas " & addfrom
		cdesde = fechaconvert(sID1 + CLng(CDate("01-01-1900")) )
		chasta = fechaconvert(sID2 + CLng(CDate("01-01-1900")) )
        sID1=sID1 + CLng(CDate("01-01-1900"))
        sID2=sID2 + CLng(CDate("01-01-1900"))
		strsql = strsql & "WHERE  xfecvto between '" & sID1 & "' and '" & sID2 & "' "
		sTitulo = "Revision&nbsp;&nbsp;Del:&nbsp;" & cdesde & "&nbsp;&nbsp;|&nbsp;al:&nbsp;" & chasta

		Case "PORACTA"

		If Robj = "F" Then

			sID1 = request.Form("Desde3") & ""
			' sID1  =  CLng(CDate(sID)) - CLng(CDate("01-01-1900"))
			sID2 = request.Form("Hasta3") & ""
			' sID2  =  CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))
			sid3 = request.Form("Fecha") & ""
			nacta = request.Form("nacta")


			Addqry = "&Menu=PORACTA&NOMBRE=ACTA&Desde=" & sID1 & "&Hasta=" & sID2 & "&Fecha=" & sid3

		Else
			sID1 = request.querystring("Desde") & ""
			sID2 = request.querystring("Hasta") & ""
			sid3 = request.querystring("Fecha") & ""
			nacta = request.querystring("nacta") & ""
			'   response.write request.form

			Addqry = "&Menu=PORACTA&NOMBRE=ACTA&Desde=" & sID1 & "&Hasta=" & sID2 & "&Fecha=" & sid3


		End If


		sMenuAdd = ""
		sMenuMas = "&smenu=" & smenu & "&Desde=" & sID1 & "&Hasta="
		sMenuMas = sMenuMas & sID2 & "&Fecha=" & sid3


		If nacta <> "" Then
			buscaacta = " Objprty.nroacta = '" & nacta & "'"
		Else
			buscaacta = "Objprty.nroacta is not null"
		End If


		cdesde = CLng(CDate(sID1)) - CLng(CDate("01-01-1900"))
		chasta = CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))


		strsql = "select DISTINCT  csporfechas.*, Objprty.fecacta,objetos.descri,tipobjeto.descri,objprty.fecemis "
		strsql = strsql & " from csporfechas join Objprty on csporfechas.codobjeto = Objprty.codobjeto, Estados  " & addfrom
		strsql = strsql & " where " & buscaacta & " and Objprty.fecacta between '" & cdesde & "' and '" & chasta & "' "
		' response.write cdesde
		If cdesde=0 And chasta=0 Then
			sTitulo = "Por acta:  " & nacta '& "&nbsp;&nbsp; Desde: " & sID1 & "&nbsp;&nbsp;|&nbsp; Hasta:&nbsp;" & sID2
		Else

			sTitulo = "Por acta:  " & nacta & "&nbsp;&nbsp; Desde: " & sID1 & "&nbsp;&nbsp;|&nbsp; Hasta:&nbsp;" & sID2
		End If

		Case "PORRESUMEN"

		If Robj = "F" Then

			sID1 = request.Form("Desde3") & ""
			' sID1  =  CLng(CDate(sID)) - CLng(CDate("01-01-1900"))
			sID2 = request.Form("Hasta3") & ""
			' sID2  =  CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))
			sid3 = request.Form("Fecha") & ""
			nacta = request.Form("nacta")


			Addqry = "&Menu=PORRESUMEN&NOMBRE=RESUMEN&Desde=" & sID1 & "&Hasta=" & sID2 & "&Fecha=" & sid3

		Else
			sID1 = request.querystring("Desde") & ""
			sID2 = request.querystring("Hasta") & ""
			sid3 = request.querystring("Fecha") & ""
			nacta = request.querystring("nacta") & ""
			'   response.write request.form

			Addqry = "&Menu=PORRESUMEN&NOMBRE=RESUMEN&Desde=" & sID1 & "&Hasta=" & sID2 & "&Fecha=" & sid3


		End If


		sMenuAdd = ""
		sMenuMas = "&smenu=" & smenu & "&Desde=" & sID1 & "&Hasta="
		sMenuMas = sMenuMas & sID2 & "&Fecha=" & sid3


		'   If nacta <> "" Then
		'       buscaacta = " Objprty.nroacta = '" & nacta & "'"
		'   Else
		'       buscaacta = "Objprty.nroacta is not null"
		'   End If


		cdesde = CLng(CDate(sID1)) - CLng(CDate("01-01-1900"))
		chasta = CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))


		strsql = "select DISTINCT  csporfechas.*, Objprty.fecacta,objetos.descri,tipobjeto.descri,objprty.fecemis "
		strsql = strsql & " from csporfechas join Objprty on csporfechas.codobjeto = Objprty.codobjeto, Estados  " & addfrom
		strsql = strsql & " where " & buscaacta & " and Objprty.fecacta between '" & cdesde & "' and '" & chasta & "' "

		sTitulo = "Resumen publicados&nbsp; Desde: " & sID1 & "&nbsp;&nbsp;|&nbsp; Hasta:&nbsp;" & sID2

		If MasterType <> "EXP" Then

			sAccionEXP = starget & "&Mtype=EXP"
			SFORMEXP = Formulario & ofv.GenForm("SUBEXP","",sAccionEXP,"_blank","post","","","si")



			''' sTitulo = sTitulo &  ".........................<a class=tt href=" & starget & "&Mtype=EXP target=_BLANK ><B><u>&nbsp;EXPORTAR&nbsp;</u></B></A>"
			sTitulo = sTitulo &  ".........................<input type=button onclick='document.all.SUBEXP.submit();' value='EXPORTAR' class=bt >"
		End If



	End Select

	If smenu <> "PORSECTOR" Then
		If smenu <> "PORCOMPONENTE" Then strsql = strsql & sverestado & wherefrom
		If smenu = "PORCOMPONENTE" Then strsql = strsql & " WHERE OBJETOS.CODOBJETO > 0 " & sverestado

	End If

	sTitOrd = ""
	sTitOrd2 = ""

	If smenu = "PORFECHA" Or smenu = "PORREVISION" Then

		If Sfiltro <> "TODOS" Then
			sTitOrd2 = "Solo&nbsp;" & Sfiltro & "&nbsp;&nbsp;"
			strsql = strsql & " and csporfechas.descri='" & Sfiltro & "' "
		End If

		Select Case sorden
			Case "N"
			sTitOrd = sTitOrd & "Nombre"
			strsql = strsql & " order by  objdescri, csporfechas.descri"
			Case "T"
			sTitOrd = sTitOrd & "Tipo"
			strsql = strsql & " order by  csporfechas.descri,objdescri"
			Case "C"
			sTitOrd = sTitOrd & "Creacion"
			strsql = strsql & " order by  fecemis desc, objdescri, csporfechas.descri"
			Case "E"
			sTitOrd = sTitOrd & "Estado"
			strsql = strsql & " order by  fecemis, estdescri, objdescri, csporfechas.descri"
			Case "P"
			sTitOrd = sTitOrd & "Publica"
			strsql = strsql & " order by  fecpublic desc, csporfechas.descri"
			Case "V"
			sTitOrd = sTitOrd & "Vigencia"
			strsql = strsql & " order by  fecvigen desc, csporfechas.descri"
			Case "A"
			sTitOrd = sTitOrd & "Autor"
			strsql = strsql & " order by  autor, csporfechas.descri"
		End Select
	ElseIf smenu = "PORSECTOR" Then

		Select Case sorden
			Case "N"
			sTitOrd = sTitOrd & "Nombre"
			strsql = strsql & " order by X.descri,E.descri "
			Case "T"
			sTitOrd = sTitOrd & "Tipo"
			strsql = strsql & " order by E.descri,X.descri "
			Case "C"
			sTitOrd = sTitOrd & "Creacion"
			strsql = strsql & " order by C.fecemis desc,E.descri "
			Case "E"
			sTitOrd = sTitOrd & "Estado"
			strsql = strsql & " order by C.fecemis, estdescri,X.descri,E.descri "
			Case "P"
			sTitOrd = sTitOrd & "Publica"
			strsql = strsql & " order by C.fecpublic desc, E.descri"
			Case "V"
			sTitOrd = sTitOrd & "Vigencia"
			strsql = strsql & " order by C.fecvigen desc, E.descri"
			Case "A"
			sTitOrd = sTitOrd & "Autor"
			strsql = strsql & " order by C.autor, E        .descri"
		End Select


	Else
		If Sfiltro <> "TODOS" Then
			sTitOrd2 = "Solo&nbsp;" & Sfiltro & "&nbsp;&nbsp;"
			strsql = strsql & " and tipobjeto.descri='" & Sfiltro & "' "
		End If

		Select Case sorden
			Case "N"
			sTitOrd = sTitOrd & "Nombre"
			strsql = strsql & " order by objetos.descri,tipobjeto.descri "
			Case "T"
			sTitOrd = sTitOrd & "Tipo"
			strsql = strsql & " order by tipobjeto.descri,objetos.descri "
			Case "C"
			sTitOrd = sTitOrd & "Creacion"
			strsql = strsql & " order by objprty.fecemis desc,tipobjeto.descri "
			Case "E"
			sTitOrd = sTitOrd & "Estado"
			strsql = strsql & " order by fecemis, estdescri,objetos.descri,tipobjeto.descri "
			Case "P"
			sTitOrd = sTitOrd & "Publica"
			strsql = strsql & " order by objprty.fecpublic desc, tipobjeto.descri"
			Case "V"
			sTitOrd = sTitOrd & "Vigencia"
			strsql = strsql & " order by objprty.fecvigen desc, tipobjeto.descri"
			Case "A"
			sTitOrd = sTitOrd & "Autor"
			strsql = strsql & " order by objprty.autor, tipobjeto.descri"
		End Select
	End If











	'----------------------------------------------------------------------------------------------------------------------------------

	' NUEVOS SQL LOS DE ARRIBA QUEDARON ANULADOS -> AM

	'RESPONSE. WRITE "<BR><BR><BR>" & smenu & "<BR><BR><BR>"

	SQL_SELECT=""
	Select Case smenu

		Case "PORESTADO"
		' SQL_SELECT="SELECT DISTINCT TIPOBJETO.descri, OBJETOS.codiden, OBJETOS.descri, OBJETOS.codobjeto, OBJETOS.tipobjeto, OBJPRTY.fecemis, OBJPRTY.fecvigen, OBJPRTY.fecbaja, Estados.estado as estdescri, OBJETOS.codiden, OBJETOS.descri, TIPOBJETO.archivo, OBJPRTY.autor, OBJPRTY.aprobo, OBJPRTY.elimino, TIPOBJETO.perfil, OBJPRTY.fecvto,Estados.estado as estcod,OBJETOS.objnivel,OBJPRTY.Version,OBJPRTY.nrover,OBJETOS.OBJUN,OBJPRTY.fecpublic , TEMADESCRI,  TEMADESCRI,  TEMADESCRI, TEMADESCRI, STEMADESCRI "
		SQL_SELECT="SELECT DISTINCT TIPOBJETO.descri, OBJETOS.codiden, OBJETOS.descri, OBJETOS.codobjeto, OBJETOS.tipobjeto, OBJPRTY.fecemis, OBJPRTY.fecvigen, OBJPRTY.fecbaja, Estados.estado as estdescri, OBJETOS.codiden, OBJETOS.descri, TIPOBJETO.archivo, OBJPRTY.autor, OBJPRTY.aprobo, OBJPRTY.elimino, TIPOBJETO.perfil, OBJPRTY.fecvto,Estados.estado as estcod,OBJETOS.objnivel,OBJPRTY.Version,OBJPRTY.nrover,OBJETOS.OBJUN,OBJPRTY.fecpublic"

		Case "PORTEXTO"
		'   SQL_SELECT="SELECT DISTINCT TIPOBJETO.descri, OBJETOS.codiden, OBJETOS.descri, OBJETOS.codobjeto, OBJETOS.tipobjeto, OBJPRTY.fecemis, OBJPRTY.fecvigen, OBJPRTY.fecbaja, Estados.estado as estdescri, OBJETOS.codiden, OBJETOS.descri, TIPOBJETO.archivo, OBJPRTY.autor, OBJPRTY.aprobo, OBJPRTY.elimino, TIPOBJETO.perfil, OBJPRTY.fecvto,Estados.estado as estcod,OBJETOS.objnivel,OBJPRTY.Version,OBJPRTY.nrover,OBJETOS.OBJUN,OBJPRTY.fecpublic , TEMADESCRI, TEMADESCRI, TEMADESCRI, TEMADESCRI, STEMADESCRI "
		SQL_SELECT="SELECT DISTINCT TIPOBJETO.descri, OBJETOS.codiden, OBJETOS.descri, OBJETOS.codobjeto, OBJETOS.tipobjeto, OBJPRTY.fecemis, OBJPRTY.fecvigen, OBJPRTY.fecbaja, Estados.estado as estdescri, OBJETOS.codiden, OBJETOS.descri, TIPOBJETO.archivo, OBJPRTY.autor, OBJPRTY.aprobo, OBJPRTY.elimino, TIPOBJETO.perfil, OBJPRTY.fecvto,Estados.estado as estcod,OBJETOS.objnivel,OBJPRTY.Version,OBJPRTY.nrover,OBJETOS.OBJUN,OBJPRTY.fecpublic,OBJPRTY.ESTADO  "
		Case "PORIDENTIFICACION"
		'SQL_SELECT="SELECT DISTINCT TIPOBJETO.descri, OBJETOS.codiden, OBJETOS.descri, OBJETOS.codobjeto, OBJETOS.tipobjeto, OBJPRTY.fecemis, OBJPRTY.fecvigen, OBJPRTY.fecbaja, Estados.estado as estdescri, OBJETOS.codiden, OBJETOS.descri, TIPOBJETO.archivo, OBJPRTY.autor, OBJPRTY.aprobo, OBJPRTY.elimino, TIPOBJETO.perfil, OBJPRTY.fecvto,Estados.estado as estcod,OBJETOS.objnivel,OBJPRTY.Version,OBJPRTY.nrover,OBJETOS.OBJUN,OBJPRTY.fecpublic , TEMADESCRI, TEMADESCRI, TEMADESCRI, TEMADESCRI, STEMADESCRI"
		SQL_SELECT="SELECT DISTINCT TIPOBJETO.descri, OBJETOS.codiden, OBJETOS.descri, OBJETOS.codobjeto, OBJETOS.tipobjeto, OBJPRTY.fecemis, OBJPRTY.fecvigen, OBJPRTY.fecbaja, Estados.estado as estdescri, OBJETOS.codiden, OBJETOS.descri, TIPOBJETO.archivo, OBJPRTY.autor, OBJPRTY.aprobo, OBJPRTY.elimino, TIPOBJETO.perfil, OBJPRTY.fecvto,Estados.estado as estcod,OBJETOS.objnivel,OBJPRTY.Version,OBJPRTY.nrover,OBJETOS.OBJUN,OBJPRTY.fecpublic "

		Case "PORFECHA"
		'SQL_SELECT=" select csporfechas.* "
		SQL_SELECT="SELECT DISTINCT  "
		SQL_SELECT=  SQL_SELECT & "dbo.TIPOBJETO.descri, dbo.OBJETOS.codiden, dbo.OBJETOS.descri AS objdescri, dbo.OBJETOS.codobjeto, dbo.OBJETOS.tipobjeto, dbo.OBJPRTY.fecemis, dbo.OBJPRTY.fecvigen, "
		SQL_SELECT=  SQL_SELECT & "dbo.OBJPRTY.fecbaja, dbo.Estados.estado AS estdescri, dbo.OBJETOS.codiden AS objcodiden, dbo.OBJETOS.descri AS objdescri2, dbo.TIPOBJETO.Archivo, dbo.OBJPRTY.autor, "
		SQL_SELECT=  SQL_SELECT & "dbo.OBJPRTY.aprobo, dbo.OBJPRTY.elimino, dbo.TIPOBJETO.perfil, dbo.OBJPRTY.fecvto, CONVERT(numeric, dbo.OBJPRTY.fecemis) - CONVERT(numeric, CONVERT(datetime, '19000101')) "
		SQL_SELECT=  SQL_SELECT & " AS xfecemis, CONVERT(numeric, dbo.OBJPRTY.fecvigen) - CONVERT(numeric, CONVERT(datetime, '19000101')) AS xfecvigen, CONVERT(numeric, dbo.OBJPRTY.fecbaja) - CONVERT(numeric, "
		SQL_SELECT=  SQL_SELECT & " CONVERT(datetime, '19000101')) AS xfecbaja, CONVERT(numeric, dbo.OBJPRTY.fecvto) - CONVERT(numeric, CONVERT(datetime, '19000101')) AS xfecvto, dbo.OBJETOS.OBJNIVEL, "
		SQL_SELECT=  SQL_SELECT & " dbo.Estados.objid AS estcod, CONVERT(numeric, dbo.OBJPRTY.fecpublic) - CONVERT(numeric, CONVERT(datetime, '19000101')) AS xfecpublic, dbo.OBJPRTY.fecpublic, dbo.OBJPRTY.nroacta"


		Case "PORREVISION"
		'SQL_SELECT=" select csporfechas.* "
		SQL_SELECT="SELECT DISTINCT  "
		SQL_SELECT=  SQL_SELECT & "dbo.TIPOBJETO.descri, dbo.OBJETOS.codiden, dbo.OBJETOS.descri AS objdescri, dbo.OBJETOS.codobjeto, dbo.OBJETOS.tipobjeto, dbo.OBJPRTY.fecemis, dbo.OBJPRTY.fecvigen, "
		SQL_SELECT=  SQL_SELECT & "dbo.OBJPRTY.fecbaja, dbo.Estados.estado AS estdescri, dbo.OBJETOS.codiden AS objcodiden, dbo.OBJETOS.descri AS objdescri2, dbo.TIPOBJETO.Archivo, dbo.OBJPRTY.autor, "
		SQL_SELECT=  SQL_SELECT & "dbo.OBJPRTY.aprobo, dbo.OBJPRTY.elimino, dbo.TIPOBJETO.perfil, dbo.OBJPRTY.fecvto, CONVERT(numeric, dbo.OBJPRTY.fecemis) - CONVERT(numeric, CONVERT(datetime, '19000101')) "
		SQL_SELECT=  SQL_SELECT & " AS xfecemis, CONVERT(numeric, dbo.OBJPRTY.fecvigen) - CONVERT(numeric, CONVERT(datetime, '19000101')) AS xfecvigen, CONVERT(numeric, dbo.OBJPRTY.fecbaja) - CONVERT(numeric, "
		SQL_SELECT=  SQL_SELECT & " CONVERT(datetime, '19000101')) AS xfecbaja, CONVERT(numeric, dbo.OBJPRTY.fecvto) - CONVERT(numeric, CONVERT(datetime, '19000101')) AS xfecvto, dbo.OBJETOS.OBJNIVEL, "
		SQL_SELECT=  SQL_SELECT & " dbo.Estados.objid AS estcod, CONVERT(numeric, dbo.OBJPRTY.fecpublic) - CONVERT(numeric, CONVERT(datetime, '19000101')) AS xfecpublic, dbo.OBJPRTY.fecpublic, dbo.OBJPRTY.nroacta"


		Case "PORACTA"
		'SQL_SELECT=" select csporfechas.* "
		SQL_SELECT="SELECT DISTINCT  "
		SQL_SELECT=  SQL_SELECT & "dbo.TIPOBJETO.descri, dbo.OBJETOS.codiden, dbo.OBJETOS.descri AS objdescri, dbo.OBJETOS.codobjeto, dbo.OBJETOS.tipobjeto, dbo.OBJPRTY.fecemis, dbo.OBJPRTY.fecvigen, "
		SQL_SELECT=  SQL_SELECT & "dbo.OBJPRTY.fecbaja, dbo.Estados.estado AS estdescri, dbo.OBJETOS.codiden AS objcodiden, dbo.OBJETOS.descri AS objdescri2, dbo.TIPOBJETO.Archivo, dbo.OBJPRTY.autor, "
		SQL_SELECT=  SQL_SELECT & "dbo.OBJPRTY.aprobo, dbo.OBJPRTY.elimino, dbo.TIPOBJETO.perfil, dbo.OBJPRTY.fecvto, CONVERT(numeric, dbo.OBJPRTY.fecemis) - CONVERT(numeric, CONVERT(datetime, '19000101')) "
		SQL_SELECT=  SQL_SELECT & " AS xfecemis, CONVERT(numeric, dbo.OBJPRTY.fecvigen) - CONVERT(numeric, CONVERT(datetime, '19000101')) AS xfecvigen, CONVERT(numeric, dbo.OBJPRTY.fecbaja) - CONVERT(numeric, "
		SQL_SELECT=  SQL_SELECT & " CONVERT(datetime, '19000101')) AS xfecbaja, CONVERT(numeric, dbo.OBJPRTY.fecvto) - CONVERT(numeric, CONVERT(datetime, '19000101')) AS xfecvto, dbo.OBJETOS.OBJNIVEL, "
		SQL_SELECT=  SQL_SELECT & " dbo.Estados.objid AS estcod, CONVERT(numeric, dbo.OBJPRTY.fecpublic) - CONVERT(numeric, CONVERT(datetime, '19000101')) AS xfecpublic, dbo.OBJPRTY.fecpublic, dbo.OBJPRTY.nroacta"

		Case "PORRESUMEN"

		SQL_SELECT="SELECT DISTINCT  "
		SQL_SELECT=  SQL_SELECT & "dbo.TIPOBJETO.descri, dbo.OBJETOS.codiden, dbo.OBJETOS.descri AS objdescri, dbo.OBJETOS.codobjeto, dbo.OBJETOS.tipobjeto, dbo.OBJPRTY.fecemis, dbo.OBJPRTY.fecvigen, "
		SQL_SELECT=  SQL_SELECT & "dbo.OBJPRTY.fecbaja, dbo.Estados.estado AS estdescri, dbo.OBJETOS.codiden AS objcodiden, dbo.OBJETOS.descri AS objdescri2, dbo.TIPOBJETO.Archivo, dbo.OBJPRTY.autor, "
		SQL_SELECT=  SQL_SELECT & "dbo.OBJPRTY.aprobo, dbo.OBJPRTY.elimino, dbo.TIPOBJETO.perfil, dbo.OBJPRTY.fecvto, CONVERT(numeric, dbo.OBJPRTY.fecemis) - CONVERT(numeric, CONVERT(datetime, '19000101')) "
		SQL_SELECT=  SQL_SELECT & " AS xfecemis, CONVERT(numeric, dbo.OBJPRTY.fecvigen) - CONVERT(numeric, CONVERT(datetime, '19000101')) AS xfecvigen, CONVERT(numeric, dbo.OBJPRTY.fecbaja) - CONVERT(numeric, "
		SQL_SELECT=  SQL_SELECT & " CONVERT(datetime, '19000101')) AS xfecbaja, CONVERT(numeric, dbo.OBJPRTY.fecvto) - CONVERT(numeric, CONVERT(datetime, '19000101')) AS xfecvto, dbo.OBJETOS.OBJNIVEL, "
		SQL_SELECT=  SQL_SELECT & " dbo.Estados.objid AS estcod, CONVERT(numeric, dbo.OBJPRTY.fecpublic) - CONVERT(numeric, CONVERT(datetime, '19000101')) AS xfecpublic, dbo.OBJPRTY.fecpublic, "
		SQL_SELECT=  SQL_SELECT & " A.temadescri,A.stemadescri,isnull(A.expr1,'') AS C1,isnull(A.expr2,'') AS C2,isnull(A.expr3,'') AS C3,isnull(A.expr4,'') AS C4, "
		SQL_SELECT=  SQL_SELECT & " isnull(A.expr5,'') AS C5,isnull(A.expr6,'') AS C6,isnull(A.expr7,'') AS C7,isnull(A.expr0,'') AS C0, d.objtitulo, d.objcopete,dbo.OBJPRTY.version,dbo.OBJPRTY.nrover"

		Case "PORINSTRUMENTO"
		'SQL_SELECT="SELECT DISTINCT TIPOBJETO.descri, OBJETOS.codiden, OBJETOS.descri, OBJETOS.codobjeto, OBJETOS.tipobjeto, OBJPRTY.fecemis, OBJPRTY.fecvigen, OBJPRTY.fecbaja, Estados.estado as estdescri, OBJETOS.codiden, OBJETOS.descri, TIPOBJETO.archivo, OBJPRTY.autor, OBJPRTY.aprobo, OBJPRTY.elimino, TIPOBJETO.perfil, OBJPRTY.fecvto,Estados.estado as estcod,OBJETOS.objnivel,OBJPRTY.Version,OBJPRTY.nrover,OBJETOS.OBJUN,OBJPRTY.fecpublic , TEMADESCRI, TEMADESCRI, TEMADESCRI, TEMADESCRI, STEMADESCRI   "
		SQL_SELECT="SELECT DISTINCT TIPOBJETO.descri, OBJETOS.codiden, OBJETOS.descri, OBJETOS.codobjeto, OBJETOS.tipobjeto, OBJPRTY.fecemis, OBJPRTY.fecvigen, OBJPRTY.fecbaja, Estados.estado as estdescri, OBJETOS.codiden, OBJETOS.descri, TIPOBJETO.archivo, OBJPRTY.autor, OBJPRTY.aprobo, OBJPRTY.elimino, TIPOBJETO.perfil, OBJPRTY.fecvto,Estados.estado as estcod,OBJETOS.objnivel,OBJPRTY.Version,OBJPRTY.nrover,OBJETOS.OBJUN,OBJPRTY.fecpublic   "

		Case "PORPRODUCTO"
		'  SQL_SELECT="SELECT DISTINCT TIPOBJETO.descri, OBJETOS.codiden, OBJETOS.descri, OBJETOS.codobjeto, OBJETOS.tipobjeto, OBJPRTY.fecemis, OBJPRTY.fecvigen, OBJPRTY.fecbaja, Estados.estado as estdescri, OBJETOS.codiden, OBJETOS.descri, TIPOBJETO.archivo, OBJPRTY.autor, OBJPRTY.aprobo, OBJPRTY.elimino, TIPOBJETO.perfil, OBJPRTY.fecvto,Estados.estado as estcod,OBJETOS.objnivel,OBJPRTY.Version,OBJPRTY.nrover,OBJETOS.OBJUN,OBJPRTY.fecpublic , TEMADESCRI, TEMADESCRI, TEMADESCRI, TEMADESCRI, STEMADESCRI  "
		SQL_SELECT="SELECT DISTINCT TIPOBJETO.descri, OBJETOS.codiden, OBJETOS.descri, OBJETOS.codobjeto, OBJETOS.tipobjeto, OBJPRTY.fecemis, OBJPRTY.fecvigen, OBJPRTY.fecbaja, Estados.estado as estdescri, OBJETOS.codiden, OBJETOS.descri, TIPOBJETO.archivo, OBJPRTY.autor, OBJPRTY.aprobo, OBJPRTY.elimino, TIPOBJETO.perfil, OBJPRTY.fecvto,Estados.estado as estcod,OBJETOS.objnivel,OBJPRTY.Version,OBJPRTY.nrover,OBJETOS.OBJUN,OBJPRTY.fecpublic   "
		Case "PORNROCIRCULAR"
		'SQL_SELECT="SELECT DISTINCT TIPOBJETO.descri, OBJETOS.codiden, OBJETOS.descri, OBJETOS.codobjeto, OBJETOS.tipobjeto, OBJPRTY.fecemis, OBJPRTY.fecvigen, OBJPRTY.fecbaja, Estados.estado as estdescri, OBJETOS.codiden, OBJETOS.descri, TIPOBJETO.archivo, OBJPRTY.autor, OBJPRTY.aprobo, OBJPRTY.elimino, TIPOBJETO.perfil, OBJPRTY.fecvto,Estados.estado as estcod,OBJETOS.objnivel,OBJPRTY.Version,OBJPRTY.nrover,OBJETOS.OBJUN,OBJPRTY.fecpublic , TEMADESCRI, TEMADESCRI, TEMADESCRI, TEMADESCRI, STEMADESCRI"
		SQL_SELECT="SELECT DISTINCT TIPOBJETO.descri, OBJETOS.codiden, OBJETOS.descri, OBJETOS.codobjeto, OBJETOS.tipobjeto, OBJPRTY.fecemis, OBJPRTY.fecvigen, OBJPRTY.fecbaja, Estados.estado as estdescri, OBJETOS.codiden, OBJETOS.descri, TIPOBJETO.archivo, OBJPRTY.autor, OBJPRTY.aprobo, OBJPRTY.elimino, TIPOBJETO.perfil, OBJPRTY.fecvto,Estados.estado as estcod,OBJETOS.objnivel,OBJPRTY.Version,OBJPRTY.nrover,OBJETOS.OBJUN,OBJPRTY.fecpublic "
	End Select

	'-------------------------------------------------------

	SQL_JOIN=""

	Select Case smenu

		Case "PORESTADO"
		SQL_JOIN= " FROM OBJETOS "
		SQL_JOIN= SQL_JOIN  & " INNER JOIN OBJPRTY on  (OBJETOS.codobjeto = OBJPRTY.codobjeto  AND OBJPRTY.estado IN(9," & SQLESTADOS & " )"

		SQL_JOIN= SQL_JOIN  & " INNER JOIN TIPOBJETO ON  OBJETOS.tipobjeto = TIPOBJETO.tipobjeto"


		SQL_JOIN= SQL_JOIN  & " INNER JOIN Estados ON  Estados.OBJID=OBJPRTY.estado "
		' SQL_JOIN= SQL_JOIN  & " inner join  csrelpathDESCRIEX a on( a.cobj = OBJETOS.codobjeto) "
		' SQL_JOIN= SQL_JOIN  & " inner join relaobj on ( OBJETOS.codobjeto =relaobj.codojasoc and OBJETOS.tipobjeto =relaobj.tipobjasoc)"
		'SQL_JOIN= SQL_JOIN  & "  inner join prodserv on (codprodserv=relaobj.codobjmaster and relaobj.tipobjmaster=marcaps)"
		'SQL_JOIN= SQL_JOIN  & " inner join tema   on ( tema.codtema=prodserv.codtema)"


	'	If (Session("EMPMUN") = "S") Then
			SQL_JOIN= SQL_JOIN  & " LEFT join OBJMEMPALC  on OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO  "
	'	End If


		Case "PORTEXTO"
		SQL_JOIN= " FROM OBJETOS "
		SQL_JOIN= SQL_JOIN  & " INNER JOIN OBJPRTY on  (OBJETOS.codobjeto = OBJPRTY.codobjeto  AND OBJPRTY.estado IN(9," & SQLESTADOS & " )"
		SQL_JOIN= SQL_JOIN  & " INNER JOIN TIPOBJETO ON  OBJETOS.tipobjeto = TIPOBJETO.tipobjeto "
		SQL_JOIN= SQL_JOIN  & " INNER JOIN Estados ON  Estados.OBJID=OBJPRTY.estado "
		' SQL_JOIN= SQL_JOIN  & " inner join  csrelpathDESCRIEX a on( a.cobj = OBJETOS.codobjeto) "
		If tcod & "" <> "" And tcod & "" <> "99999" Then
			SQL_JOIN= SQL_JOIN  & "  INNER JOIN csrelpathdescriNV ON  (csrelpathdescriNV.tema='" & tcod & "'  and csrelpathdescriNV.cobj=OBJETOS.codobjeto and csrelpathdescriNV.tobj=OBJETOS.TIPOBJETO)"
		End If
	'	If (Session("EMPMUN") = "S") Then
			SQL_JOIN= SQL_JOIN  & " LEFT join OBJMEMPALC  on OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO  "
	'	End If

		Case "PORIDENTIFICACION"
		SQL_JOIN= " FROM OBJETOS "
		SQL_JOIN= SQL_JOIN  & " INNER JOIN OBJPRTY on  (OBJETOS.codobjeto = OBJPRTY.codobjeto  AND OBJPRTY.estado IN(9," & SQLESTADOS & " )"
		SQL_JOIN= SQL_JOIN  & " INNER JOIN TIPOBJETO ON  OBJETOS.tipobjeto = TIPOBJETO.tipobjeto "
		SQL_JOIN= SQL_JOIN  & " INNER JOIN Estados ON  Estados.OBJID=OBJPRTY.estado "
		' SQL_JOIN= SQL_JOIN  & " inner join  csrelpathDESCRIEX a on( a.cobj = OBJETOS.codobjeto) "
		'If Session("EMPMUN") = "S" Then
			SQL_JOIN= SQL_JOIN  & " LEFT join OBJMEMPALC  on OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO  "
		'End If

		Case "PORFECHA"
		SQL_JOIN= " FROM objetos  "
		' SQL_JOIN= SQL_JOIN  & " INNER JOIN objetos on  (objetos.codobjeto = objetos.codobjeto)"
		SQL_JOIN= SQL_JOIN  & " INNER JOIN OBJPRTY on  (objetos.codobjeto = OBJPRTY.codobjeto  AND OBJPRTY.estado IN(9," & SQLESTADOS & " )"
		SQL_JOIN= SQL_JOIN  & " INNER JOIN TIPOBJETO ON  objetos.tipobjeto = TIPOBJETO.tipobjeto "
		SQL_JOIN= SQL_JOIN  & " INNER JOIN Estados ON  Estados.OBJID=OBJPRTY.estado "
		' SQL_JOIN= SQL_JOIN  & " inner join  csrelpathDESCRIEX a on( a.cobj = objetos.codobjeto) "
		'If Session("EMPMUN") = "S"  Then
			SQL_JOIN= SQL_JOIN  & " LEFT join OBJMEMPALC  on objetos.CODOBJETO = OBJMEMPALC.CODOBJETO  "
		'End If


		Case "PORREVISION"
		SQL_JOIN= " FROM objetos  "
		' SQL_JOIN= SQL_JOIN  & " INNER JOIN objetos on  (objetos.codobjeto = objetos.codobjeto)"
		SQL_JOIN= SQL_JOIN  & " INNER JOIN OBJPRTY on  (objetos.codobjeto = OBJPRTY.codobjeto  AND OBJPRTY.estado IN(9," & SQLESTADOS & " )"
		SQL_JOIN= SQL_JOIN  & " INNER JOIN TIPOBJETO ON  objetos.tipobjeto = TIPOBJETO.tipobjeto "
		SQL_JOIN= SQL_JOIN  & " INNER JOIN Estados ON  Estados.OBJID=OBJPRTY.estado "
		' SQL_JOIN= SQL_JOIN  & " inner join  csrelpathDESCRIEX a on( a.cobj = objetos.codobjeto) "
		'If Session("EMPMUN") = "S"  Then
			SQL_JOIN= SQL_JOIN  & " LEFT join OBJMEMPALC  on objetos.CODOBJETO = OBJMEMPALC.CODOBJETO  "
		'End If


		Case "PORACTA"
		SQL_JOIN= " FROM objetos  "
		' SQL_JOIN= SQL_JOIN  & " INNER JOIN objetos on  (objetos.codobjeto = objetos.codobjeto)"
		SQL_JOIN= SQL_JOIN  & " INNER JOIN OBJPRTY on  (objetos.codobjeto = OBJPRTY.codobjeto  AND OBJPRTY.estado IN(9," & SQLESTADOS & " )"
		SQL_JOIN= SQL_JOIN  & " INNER JOIN TIPOBJETO ON  objetos.tipobjeto = TIPOBJETO.tipobjeto "
		SQL_JOIN= SQL_JOIN  & " INNER JOIN Estados ON  Estados.OBJID=OBJPRTY.estado "
		' SQL_JOIN= SQL_JOIN  & " inner join  csrelpathDESCRIEX a on( a.cobj = objetos.codobjeto) "
		'If Session("EMPMUN") = "S"  Then
			SQL_JOIN= SQL_JOIN  & " LEFT join OBJMEMPALC  on objetos.CODOBJETO = OBJMEMPALC.CODOBJETO  "
		'End If

		Case "PORRESUMEN"
		SQL_JOIN= " FROM objetos  "
		SQL_JOIN= SQL_JOIN  & " INNER JOIN OBJPRTY on  (objetos.codobjeto = OBJPRTY.codobjeto  AND OBJPRTY.estado IN(9," & SQLESTADOS & " )"
		SQL_JOIN= SQL_JOIN  & " INNER JOIN TIPOBJETO ON  objetos.tipobjeto = TIPOBJETO.tipobjeto "
		SQL_JOIN= SQL_JOIN  & " INNER JOIN Estados ON  Estados.OBJID=OBJPRTY.estado "
		SQL_JOIN= SQL_JOIN  & " inner join  csrelpathDESCRInv a on( a.cobj = objetos.codobjeto) "
		SQL_JOIN= SQL_JOIN  & " left join OBJINFO  d on ( objetos.codobjeto=d.codobjeto) "
		'  If Session("EMPMUN") = "S"  Then
		'      SQL_JOIN= SQL_JOIN  & " LEFT join OBJMEMPALC  on objetos.CODOBJETO = OBJMEMPALC.CODOBJETO  "
		' END IF



		Case "PORINSTRUMENTO"
		SQL_JOIN= " FROM OBJETOS "
		SQL_JOIN= SQL_JOIN  & " INNER JOIN OBJPRTY on  (OBJETOS.codobjeto = OBJPRTY.codobjeto  AND OBJPRTY.estado IN(9," & SQLESTADOS & " )"
		SQL_JOIN= SQL_JOIN  & " INNER JOIN TIPOBJETO ON  OBJETOS.tipobjeto = TIPOBJETO.tipobjeto "
		SQL_JOIN= SQL_JOIN  & " INNER JOIN Estados ON  Estados.OBJID=OBJPRTY.estado "
		' SQL_JOIN= SQL_JOIN  & " inner join  csrelpathDESCRIEX a on( a.cobj = OBJETOS.codobjeto) "
		SQL_JOIN= SQL_JOIN  & " LEFT JOIN OBJPRTYADV ON  OBJPRTY.codobjeto = OBJPRTYADV.codobjeto "
		'If (Session("EMPMUN") = "S" ) Then
			SQL_JOIN= SQL_JOIN  & " LEFT join OBJMEMPALC  on OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO  "
		'End If
		Case "PORPRODUCTO"

		SQL_JOIN= " FROM OBJETOS "
		SQL_JOIN= SQL_JOIN  & " INNER JOIN OBJPRTY on  (OBJETOS.codobjeto = OBJPRTY.codobjeto  AND OBJPRTY.estado IN(9," & SQLESTADOS & " )"
		SQL_JOIN= SQL_JOIN  & " INNER JOIN TIPOBJETO ON  OBJETOS.tipobjeto = TIPOBJETO.tipobjeto "
		SQL_JOIN= SQL_JOIN  & " INNER JOIN Estados ON  Estados.OBJID=OBJPRTY.estado "
		'  SQL_JOIN= SQL_JOIN  & " inner join  csrelpathDESCRIEX a on( a.cobj = OBJETOS.codobjeto) "
		'If (Session("EMPMUN") = "S")  Then
			SQL_JOIN= SQL_JOIN  & " LEFT join OBJMEMPALC  on OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO  "
		'End If

		SQL_JOIN= SQL_JOIN  & " LEFT join OBJTSEGAGR  on OBJETOS.CODOBJETO = OBJTSEGAGR.CODOBJETO  "

		Case "PORNROCIRCULAR"

		SQL_JOIN= " FROM OBJETOS "
		SQL_JOIN= SQL_JOIN  & " INNER JOIN OBJPRTY on  (OBJETOS.codobjeto = OBJPRTY.codobjeto  AND OBJPRTY.estado IN(9," & SQLESTADOS & " )"
		' SQL_JOIN= SQL_JOIN  & " INNER JOIN OBJTSEGAGR ON  OBJTSEGAGR.codobjeto = OBJETOS.codobjeto  "
		SQL_JOIN= SQL_JOIN  & " INNER JOIN TIPOBJETO ON  OBJETOS.tipobjeto = TIPOBJETO.tipobjeto "
		SQL_JOIN= SQL_JOIN  & " INNER JOIN Estados ON  Estados.OBJID=OBJPRTY.estado "
		' SQL_JOIN= SQL_JOIN  & " inner join  csrelpathDESCRIEX a on( a.cobj = OBJETOS.codobjeto) "


		'If (Session("EMPMUN") = "S") Then
			SQL_JOIN= SQL_JOIN  & " LEFT join OBJMEMPALC  on OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO  "
		'End If



		SQL_JOIN= SQL_JOIN  & " LEFT join OBJTSEGAGR  on OBJETOS.CODOBJETO = OBJTSEGAGR.CODOBJETO  "
	End Select



	'-----------------------------------------------------------------------------------------------------------------------------------------------------


	SQL_WHERE=" WHERE "

	if session("base") = "K" THEN
	 STPIPOF= " IN ('K') "
	ELSE
	 STPIPOF= " NOT IN ('K','PD') "
	END IF


	Select Case smenu

		Case "PORESTADO"
		SQL_WHERE = SQL_WHERE & "  OBJPRTY.estado = " & sID
		SQL_WHERE = SQL_WHERE & " AND Objetos.objnivel >= " & Session("SegNivel")
		SQL_WHERE= SQL_WHERE & " AND TIPOBJETO.TIPOCPT  " & STPIPOF & " and (OBJETOS.APLIC = '" & wkaplic & "' or OBJETOS.APLIC = 'GLB') "
		' SQL_WHERE= SQL_WHERE & " AND   (Objprty.fecemis between '" & cdesde & "' and '" & chasta & "' ) "

	'	If (Session("EMPMUN") = "S") Then
	'		SQL_WHERE= SQL_WHERE & "  AND  OBJETOS.OBJUN = " & Session("UN")
			SQL_WHERE= SQL_WHERE & "  AND  ( (OBJETOS.OBJUN = " & Session("UN") & ") OR    (   OBJMEMPALC.CODEMPRE = " & Session("UN") & ")  )"
		'	response.write SQL_WHERE
	'	End If
		Case "PORTEXTO"

		SQL_WHERE= SQL_WHERE &  " (OBJETOS.descriabr  COLLATE SQL_Latin1_General_CP1_CI_AI like '%" & nwsID & "%' "
		SQL_WHERE = SQL_WHERE & " or OBJETOS.descri   COLLATE SQL_Latin1_General_CP1_CI_AI like '%" & nwsID & "%' or"
		SQL_WHERE= SQL_WHERE &  " OBJETOS.codiden COLLATE SQL_Latin1_General_CP1_CI_AI like '%" & nwsID & "%' )"



		If tsel <> "T" And tgrp <> "M" And (tgrp = "T" And tcod <> "99999") Then

			'   strsql = strsql & ", csrelpath rp "

			If tgrp = "T" Then

				'  SQL_WHERE= SQL_WHERE & "  AND  a.TEMA = 0" & tcod & " "

			ElseIf tgrp = "S" Then

				SQL_WHERE= SQL_WHERE & " and   a.stema = 0" & tcod & " "

			Else

				'     htwhere = " and OBJETOS.codobjeto = rp.cobj and ("
				'     htwhere = htwhere & " (rp.mc1 = 0" & tcod & " and rp.mt1 = '" & ttip & "' ) or "
				'     htwhere = htwhere & " (rp.mc2 = 0" & tcod & " and rp.mt2 = '" & ttip & "' ) or "
				'     htwhere = htwhere & " (rp.mc3 = 0" & tcod & " and rp.mt3 = '" & ttip & "' ) or "
				'     htwhere = htwhere & " (rp.mc4 = 0" & tcod & " and rp.mt4 = '" & ttip & "' ) or "
				'     htwhere = htwhere & " (rp.mc5 = 0" & tcod & " and rp.mt5 = '" & ttip & "' ) or "
				'     htwhere = htwhere & " (rp.mc6 = 0" & tcod & " and rp.mt6 = '" & ttip & "' ) or "
				'     htwhere = htwhere & " (rp.mc7 = 0" & tcod & " and rp.mt7 = '" & ttip & "' ) or "
				'     htwhere = htwhere & " (rp.cobj = 0" & tcod & " and rp.tobj = '" & ttip & "' )  "
				'     htwhere = htwhere & " ) "

			End If



		End If

		SQL_WHERE = SQL_WHERE & " AND Objetos.objnivel >= " & Session("SegNivel")
		SQL_WHERE= SQL_WHERE & " AND TIPOBJETO.TIPOCPT " & STPIPOF & " and (OBJETOS.APLIC = '" & wkaplic & "' or OBJETOS.APLIC = 'GLB') "


	'	If (Session("EMPMUN") = "S") Then
			SQL_WHERE= SQL_WHERE & "  AND  ( (OBJETOS.OBJUN = " & Session("UN") &") OR    (   OBJMEMPALC.CODEMPRE = " & Session("UN") & ")  )"
	'	End If


		Case "PORIDENTIFICACION"

		SQL_WHERE = SQL_WHERE & " OBJETOS.codiden = '" & CStr(sID1 & "") & "' "
		SQL_WHERE = SQL_WHERE & " AND Objetos.objnivel >= " & Session("SegNivel")
		SQL_WHERE= SQL_WHERE & " AND TIPOBJETO.TIPOCPT " & STPIPOF & "  and (OBJETOS.APLIC = '" & wkaplic & "' or OBJETOS.APLIC = 'GLB') "

	'	If (Session("EMPMUN") = "S") Then
			SQL_WHERE= SQL_WHERE & "  AND  ( (OBJETOS.OBJUN = " & Session("UN") &") OR    (   OBJMEMPALC.CODEMPRE = " & Session("UN") & ")  )"
	'	End If

		'    strsql = strsql & " AND OBJETOS.codiden = '" & CStr(sID2 & "") & "' "
		Case "PORFECHA"


		cdesde = CLng(CDate(sID1)) - CLng(CDate("01-01-1900"))
		chasta = CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))
		If sid3 = "Emision" Then
			SQL_WHERE = SQL_WHERE &"   fecemis between '" & cdesde & "' and '" & chasta & "' "

		ElseIf sid3 = "Vigencia" Then
			SQL_WHERE = SQL_WHERE &"    fecvigen between '" & cdesde & "' and '" & chasta & "' "

		ElseIf sid3 = "Eliminacion" Then
			SQL_WHERE = SQL_WHERE &"   fecbaja between '" & cdesde & "' and '" & chasta & "' "

		ElseIf sid3 = "Vencimiento" Then
			SQL_WHERE = SQL_WHERE &"   fecvto between '" & cdesde & "' and '" & chasta & "' "

		ElseIf sid3 = "Publicacion" Then
			SQL_WHERE = SQL_WHERE &"    fecpublic between '" & cdesde & "' and '" & chasta & "' "

		End If

		SQL_WHERE = SQL_WHERE & " AND Objetos.objnivel >= " & Session("SegNivel")
		SQL_WHERE= SQL_WHERE & " AND TIPOBJETO.TIPOCPT " & STPIPOF & "  and (OBJETOS.APLIC = '" & wkaplic & "' or OBJETOS.APLIC = 'GLB') "

	'	If (Session("EMPMUN") = "S") Then
			SQL_WHERE= SQL_WHERE & "  AND  ( (OBJETOS.OBJUN = " & Session("UN") &") OR    (   OBJMEMPALC.CODEMPRE = " & Session("UN") & ")  )"
	'	End If



		Case "PORREVISION"

        'response.write sID1 & "  "   & sID2
		cdesde = CLng(CDate(sID1)) - CLng(CDate("01-01-1900"))
		chasta = CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))

		SQL_WHERE = SQL_WHERE &"   fecvto between '" & cdesde & "' and '" & chasta & "' "



		SQL_WHERE = SQL_WHERE & " AND Objetos.objnivel >= " & Session("SegNivel")
		SQL_WHERE= SQL_WHERE & " AND TIPOBJETO.TIPOCPT " & STPIPOF & "  and (OBJETOS.APLIC = '" & wkaplic & "' or OBJETOS.APLIC = 'GLB') "

	'	If (Session("EMPMUN") = "S") Then
			SQL_WHERE= SQL_WHERE & "  AND  ( (OBJETOS.OBJUN = " & Session("UN") &") OR    (   OBJMEMPALC.CODEMPRE = " & Session("UN") & ")  )"
	'	End If

		Case "PORACTA"


		cdesde = CLng(CDate(sID1)) - CLng(CDate("01-01-1900"))
		chasta = CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))
		sqland=""
		If cdesde  <> 0 And chasta <> 0 Then
			SQL_WHERE = SQL_WHERE &"   fecACTA between '" & cdesde & "' and '" & chasta & "' "
			sqland=" and "
		End If

		If nacta <> "" Then
			SQL_WHERE = SQL_WHERE & sqland & "  Objprty.nroacta = '" & nacta & "'"
		Else
			SQL_WHERE = SQL_WHERE & sqland & "   Objprty.nroacta is not null"
		End If

		SQL_WHERE = SQL_WHERE & " AND Objetos.objnivel >= " & Session("SegNivel")
		SQL_WHERE= SQL_WHERE & " AND TIPOBJETO.TIPOCPT " & STPIPOF & "  and (OBJETOS.APLIC = '" & wkaplic & "' or OBJETOS.APLIC = 'GLB') "

	'	If (Session("EMPMUN") = "S") Then
			SQL_WHERE= SQL_WHERE & "  AND  ( (OBJETOS.OBJUN = " & Session("UN") &") OR    (   OBJMEMPALC.CODEMPRE = " & Session("UN") & ")  )"
	'	End If


		Case "PORRESUMEN"


		cdesde = CLng(CDate(sID1)) - CLng(CDate("01-01-1900"))
		chasta = CLng(CDate(sID2)) - CLng(CDate("01-01-1900"))

		SQL_WHERE = SQL_WHERE & "   fecpublic between '" & cdesde & "' and '" & chasta & "' "



		SQL_WHERE = SQL_WHERE & " AND Objetos.objnivel >= " & Session("SegNivel")
		SQL_WHERE= SQL_WHERE & " AND TIPOBJETO.TIPOCPT " & STPIPOF & "  and (OBJETOS.APLIC = '" & wkaplic & "' or OBJETOS.APLIC = 'GLB') "

	'	If (Session("EMPMUN") = "S") Then
			SQL_WHERE= SQL_WHERE & "  AND  ( (OBJETOS.OBJUN = " & Session("UN") &") OR    (   OBJMEMPALC.CODEMPRE = " & Session("UN") & ")  )"
	'	End If

		Case "PORINSTRUMENTO"



		If ("0" & sID1) <> "0" Then


			If sID2 <> "" Then
				SQL_WHERE = SQL_WHERE & "  ((OBJPRTY.instrumen=" & CDbl("0" & sID1) & ")"
			Else
				SQL_WHERE = SQL_WHERE &  "  (OBJPRTY.instrumen=" & CDbl("0" & sID1) & " or OBJPRTYADV.instrua=" & CDbl("0" & sID1) & ")  "
			End If

			If sID2 <> "" Then
				SQL_WHERE = SQL_WHERE &  " AND (OBJPRTY.instid ='" & sID2 & "')) OR "
				SQL_WHERE = SQL_WHERE &  " (OBJPRTYADV.instrua =" & sID1 & " and OBJPRTYADV.instid ='" & sID2 & "')"
			End If

		Else

			SQL_WHERE = SQL_WHERE & "(OBJPRTY.instid ='" & sID2 & "')"

		End If

		SQL_WHERE = SQL_WHERE & " AND Objetos.objnivel >= " & Session("SegNivel")
		SQL_WHERE= SQL_WHERE & " AND TIPOBJETO.TIPOCPT " & STPIPOF & "  and (OBJETOS.APLIC = '" & wkaplic & "' or OBJETOS.APLIC = 'GLB') "

	'	If (Session("EMPMUN") = "S") Then
			SQL_WHERE= SQL_WHERE & "  AND  ( (OBJETOS.OBJUN = " & Session("UN") &") OR    (   OBJMEMPALC.CODEMPRE = " & Session("UN") & ")  )"
	'	End If


		Case "PORPRODUCTO"


		SQL_WHERE = SQL_WHERE & "  Objetos.objnivel >= " & Session("SegNivel")
		SQL_WHERE= SQL_WHERE & " AND TIPOBJETO.TIPOCPT" & STPIPOF & "  and (OBJETOS.APLIC = '" & wkaplic & "' or OBJETOS.APLIC = 'GLB') "

	'	If (Session("EMPMUN") = "S") Then
			SQL_WHERE= SQL_WHERE & "  AND  ( (OBJETOS.OBJUN = " & Session("UN") &") OR    (   OBJMEMPALC.CODEMPRE = " & Session("UN") & ")  )"
	'	End If

		If UCase(sid3 & "") <> UCase("nsnc") Then
			SQL_WHERE = SQL_WHERE & "   AND OBJETOS.TSEGMENTO =" & CDbl("0" & sid3) & " "
		End If

		If UCase(sID1 & "") <> UCase("nsnc") Or UCase(sID2 & "") <> UCase("nsnc") Then
			SQL_WHERE = SQL_WHERE & "    AND ( "
		End If

		If UCase(sID1 & "") <> UCase("nsnc") Then
			SQL_WHERE = SQL_WHERE & " (OBJTSEGAGR.TSEGMENTO = " & CDbl("0" & sID1) & " ) "
		End If

		If UCase(sID2 & "") <> UCase("nsnc") Then
			'   If UCase(sID1 & "") <> UCase("nsnc") Then strsql = strsql & " AND "
			SQL_WHERE = SQL_WHERE & " AND (OBJTSEGAGR.TAGRUPA = " & CDbl("0" & sID2) & " ) "
		End If

		If UCase(sID1 & "") <> UCase("nsnc") Or UCase(sID2 & "") <> UCase("nsnc") Then
			SQL_WHERE = SQL_WHERE & "  )"
		End If

		Case "PORNROCIRCULAR"

		SQL_WHERE = SQL_WHERE & "  Objetos.objnivel >= " & Session("SegNivel")
		SQL_WHERE= SQL_WHERE & " AND TIPOBJETO.TIPOCPT " & STPIPOF & "  and (OBJETOS.APLIC = '" & wkaplic & "' or OBJETOS.APLIC = 'GLB') "

	'	If (Session("EMPMUN") = "S") Then
			SQL_WHERE= SQL_WHERE & "  AND  ( (OBJETOS.OBJUN = " & Session("UN") &") OR    (   OBJMEMPALC.CODEMPRE = " & Session("UN") & ")  )"
	'	End If
		'RESPONSE.WRITE REQUEST.FORM
		If UCase(sID2 & "") <> UCase("nsnc") And UCase(sID2 & "") <> "" Then
			SQL_WHERE= SQL_WHERE & " AND OBJETOS.codiden = '" & CStr(sID2 & "") & "' "
		End If

		If UCase(sID1 & "") <> UCase("nsnc") Then
			SQL_WHERE= SQL_WHERE &  "    AND (OBJTSEGAGR.TAGRUPA = " & CDbl("0" & sID1) & " )  "
		End If

	End Select


'response.write  SQL_WHERE


	sTitOrd = ""
	sTitOrd2 = ""
	SQL_ORDEN=""
	If smenu = "PORFECHA" Or smenu = "PORREVISION" Then

		If Sfiltro <> "TODOS" Then
			sTitOrd2 = "Solo&nbsp;" & Sfiltro & "&nbsp;&nbsp;"
			SQL_ORDEN =  " and descri='" & Sfiltro & "' "
		End If

		Select Case sorden
			Case "N"
			sTitOrd = sTitOrd & "Nombre"
			SQL_ORDEN =  " order by  objdescri, descri"
			Case "T"
			sTitOrd = sTitOrd & "Tipo"
			SQL_ORDEN =  " order by  descri,objdescri"
			Case "C"
			sTitOrd = sTitOrd & "Creacion"
			SQL_ORDEN =  " order by  fecemis desc, objdescri, descri"
			Case "E"
			sTitOrd = sTitOrd & "Estado"
			SQL_ORDEN =  " order by  fecemis, estdescri, objdescri, descri"
			Case "P"
			sTitOrd = sTitOrd & "Publica"
			SQL_ORDEN =  " order by  fecpublic desc, descri"
			Case "V"
			sTitOrd = sTitOrd & "Vigencia"
			SQL_ORDEN = " order by  fecvigen desc, descri"
			Case "A"
			sTitOrd = sTitOrd & "Autor"
			SQL_ORDEN =  " order by  autor, descri"
		End Select
	ElseIf smenu = "PORSECTOR" Then

		Select Case sorden
			Case "N"
			sTitOrd = sTitOrd & "Nombre"
			SQL_ORDEN =  " order by X.descri,E.descri "
			Case "T"
			sTitOrd = sTitOrd & "Tipo"
			SQL_ORDEN = " order by E.descri,X.descri "
			Case "C"
			sTitOrd = sTitOrd & "Creacion"
			SQL_ORDEN = " order by C.fecemis desc,E.descri "
			Case "E"
			sTitOrd = sTitOrd & "Estado"
			strsql = strsql & " order by C.fecemis, estdescri,X.descri,E.descri "
			Case "P"
			sTitOrd = sTitOrd & "Publica"
			SQL_ORDEN =  " order by C.fecpublic desc, E.descri"
			Case "V"
			sTitOrd = sTitOrd & "Vigencia"
			SQL_ORDEN =  " order by C.fecvigen desc, E.descri"
			Case "A"
			sTitOrd = sTitOrd & "Autor"
			SQL_ORDEN =  " order by C.autor, E        .descri"
		End Select


	Else
		If Sfiltro <> "TODOS" Then
			sTitOrd2 = "Solo&nbsp;" & Sfiltro & "&nbsp;&nbsp;"
			'VER    strsql = strsql & " and tipobjeto.descri='" & Sfiltro & "' "
		End If

		Select Case sorden
			Case "N"
			sTitOrd = sTitOrd & "Nombre"
			SQL_ORDEN =  " order by objetos.descri,tipobjeto.descri "
			Case "T"
			sTitOrd = sTitOrd & "Tipo"
			SQL_ORDEN =  " order by tipobjeto.descri,objetos.descri "
			Case "C"
			sTitOrd = sTitOrd & "Creacion"
			SQL_ORDEN =   " order by objprty.fecemis desc,tipobjeto.descri "
			Case "E"
			sTitOrd = sTitOrd & "Estado"
			SQL_ORDEN = " order by fecemis, estdescri,objetos.descri,tipobjeto.descri "
			Case "P"
			sTitOrd = sTitOrd & "Publica"
			SQL_ORDEN =  " order by objprty.fecpublic desc, tipobjeto.descri"
			Case "V"
			sTitOrd = sTitOrd & "Vigencia"
			SQL_ORDEN =  " order by objprty.fecvigen desc, tipobjeto.descri"
			Case "A"
			sTitOrd = sTitOrd & "Autor"
			SQL_ORDEN =  " order by objprty.autor, tipobjeto.descri"
		End Select
	End If



	' FIN DE NUEVO SQL
	'-------------------------------------------------------------------------------------------------------------------------------------



	'  response.write SMENU & "  <BR>"   & STRSQL   '"<br><br><br>" & SQL_SELECT & "<br>" & SQL_JOIN  & "<BR>" & SQL_WHERE & "<BR>" & SQL_ORDEN
	Select Case smenu
		Case "PORSECTOR"

		Case "PORCOMPONENTE"
		'RESPONSE.WRITE tsel  & "    "  &  tgrp   & "   "   & tcod
		Case "PORACTAW"
		Case Else
		strsql=SQL_SELECT & SQL_JOIN &   SQL_WHERE & SQL_ORDEN

	End Select


	If request.querystring("scal") = "S" Then
		strsql = Session("sqlcons")
	Else
		Session("sqlcons") = strsql
	End If

	'response.write SMENU  & "<br>" & strsql  & Chr(13)
	'response.write strsql
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


	xs = "<script language=JavaScript>" & Chr(13)
	xs = xs & "function AbrirCons(Datos) {" & Chr(13)
	xs = xs & " var h=570;w=790 ;"
'	If Session("APLMDCLI") = "BC" Then
		xs = xs & "if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;} "
	'End If

	xs = xs & " myWin= open(Datos"
	xs = xs & "," & Chr(34) & "Cons" & Chr(34) & ","    ' & Chr(34)
	'xs = xs & "menubar=no,toolbar=no,height=570,width=790,resizable=yes,status=yes" & Chr(34)
	xs = xs & "  'menubar=no,toolbar=no,width='+w+',height='+h+',top=0,left=0,resizable=yes,status=yes'"    '& CHR(34)
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



	Fila = ""
	Celda = ""
	sHref = ""

	If MasterType <> "EXP" Then

		If Session("EMPMUN") = "S" Then

			sLink = OFV.GenImage("", "", "", "", "..\CORESS20\EMPRESA\" & ATREMP(0, 5), "0", "")
			Celda = Celda & OFV.GenCelda("", "tt ", "", "8%", "", "", "", "EMPRESA:&nbsp;&nbsp;", "si")
			Celda = Celda & OFV.GenCelda("", "ttt ", "center", "5%", "", "", "", sLink, "si")
			Celda = Celda & OFV.GenCelda("", "ttt  ", "", "", "", "", "", "&nbsp;<BIG>" & ATREMP(0, 4) & "</BIG>", "si")

			Fila = Fila & OFV.GenRow("", "", "top", "", "", "", "", Celda, "si")

			Tabla = OFV.GenTabla("", "aaa  ", "", "100% ", "", "0", "", "-1", "-2", "", "", Fila, "si")
			sHTML = sHTML & Tabla
			sHTML = sHTML & "<br>"
		End If

	End If
	Fila = ""

	'***************** LINK DEL CAMBIAR VISTA
	sNombre = Replace(sNombre, " ", "_")
	Select Case smenu

		Case "PORACTA"
		sAccion = "ConsdocumentoswebScons.asp?xx=" & Addqry    ' "&id=" & sID1 & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N&Nombre=" & sNombre & "&Desde="& Request.querystring("Desde")  &"&Hasta="& Request.querystring("hasta") &"&Fecha="& si3
		sAccion = sAccion & "&nacta=" & nacta

		Case "PORRESUMEN"
		sAccion = "ConsdocumentoswebScons.asp?xx=" & Addqry    ' "&id=" & sID1 & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N&Nombre=" & sNombre & "&Desde="& Request.querystring("Desde")  &"&Hasta="& Request.querystring("hasta") &"&Fecha="& si3
		sAccion = sAccion & "&nacta=" & nacta


		Case "PORFECHA"
		sAccion = "ConsdocumentoswebScons.asp?xx=" & Addqry    ' & "&id=" & sID1 & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N&Nombre=" & sNombre & "&Desde="& sid  &"&Hasta="& sid2 &"&Fecha="& sid3

		Case "PORESTADO"
		sAccion = "ConsdocumentoswebScons.asp?xx=" & Addqry    '& sMenu & "&id=" & sID & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N&Nombre=" & sNombre & "&Desde="& cdesde &"&Hasta="& chasta &"&Fecha="& Request.querystring("fecha")

		Case "PORREVISION"
		sAccion = "ConsdocumentoswebScons.asp?xx=" & Addqry    ' "&id=" & sID1 & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N&Nombre=" & sNombre & "&Desde="& sid &"&Hasta="& sid2 &"&Fecha="& Request.querystring("fecha")

		Case "PORSECTOR"
		sAccion = "ConsdocumentoswebScons.asp?xx=" & Addqry    '"&id=" & sID1 & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N&Nombre=" & sNombre & "&perfil="& Sperfil &"&Hasta="& chasta &"&Fecha="& Request.querystring("fecha")

		Case "PORCOMPONENTE"
		sAccion = "ConsdocumentoswebScons.asp?xx=" & Addqry    ' "&id=" & sID1 & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N&Nombre=" & sNombre

		Case "PORTEXTO"

		sAccion = "ConsdocumentoswebScons.asp?xx=" & Addqry    '

		Case "PORPRODUCTO"


		sAccion = "ConsdocumentoswebScons.asp?xx=" & Addqry

		Case "PORIDENTIFICACION"


		sAccion = "ConsdocumentoswebScons.asp?xx=" & Addqry

		Case "PORNROCIRCULAR"


		sAccion = "ConsdocumentoswebScons.asp?xx=" & Addqry


		Case "PORINSTRUMENTO"


		sAccion = "ConsdocumentoswebScons.asp?xx=" & Addqry

	End Select


	'response.write "<br> starget:"& starget


	sAccion = sAccion & "&sorden=" & sorden
	'*****************************************
	'''
	'sAccion = "ConsdocumentoswebScons.asp?Menu=" & sMenu & "&id=" & sID1 & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N&Nombre=" & sNombre& "&Desde="& cdesde &"&Hasta="& chasta &"&Fecha="& Request.querystring("fecha")
	sinput = OFV.GenerarInput("", "Vista Web", "submit", "", "10", "bt", "")

	' sInput   = sinput & ofv.GenerarInput("lista",sID1,"hidden","","1","dt","")
	' sInput   = sinput & ofv.GenerarInput("lista2",sID1,"hidden","","1","dt","")
	Formvista = OFV.GenForm("vistar", "", sAccion, "", "post", " style='display:none;' ", sinput, "si") & "Cambiar a Vista Web"

	Celda = OFV.GenCelda("", "vt style='border-style:solid;border-width:1;border-color:#dddddd;'  ", "", "", "", "", "", sImagena & "&nbsp;" & sTitulo & "", "si")

	If smenu <> "PORACTA" And smenu <> "PORRESUMEN" Then
		Celda = Celda & OFV.GenCelda("", "vt style='border-style:solid;border-width:1;border-color:#dddddd;cursor:pointer;' ", "center onclick=""document.getElementById('vistar').submit();"" ", "20%", "", "", "", Formvista, "si")
	Else
		Celda = Celda & OFV.GenCelda("", "vt style='border-style:solid;border-width:1;border-color:#dddddd;cursor:pointer;' ", "center onclicka=""document.getElementById('vistar').submit();"" ", "20%", "", "", "", "&nbsp;", "si")
	End If
	Celda = Celda & OFV.GenCelda("", "vt style='border-style:solid;border-width:1;border-color:#dddddd;' ", "right", "20%", "", "", "", "&nbsp;" & sTitOrd2 & "Orden:&nbsp;" & sTitOrd & "&nbsp;&nbsp;", "si")


	Fila = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")

	Tablaz = OFV.GenTabla("", "aaa bgcolor='#eeeeee' style='border-style:solid;border-width:1;border-color:#dddddd;' ", "", "100% ", "", "0", "", "1", "2", "", "", Fila, "si")

	sHTML = sHTML & Tablaz

	Fila = ""

	'' AGREGADO 26052014

	If smenu = "PORESTADO" Then
		'If ssEstado = "VIGENTE" Then                     'pregunto si es estado vigente para obtener los totales por tipo'
		'strsql = " SELECT isnull(max(TIPOBJETO.descri),'SIN DATOS') as tipo, isnull(sum(1),0) as total "
		'strsql = strsql & " FROM OBJETOS, TIPOBJETO, OBJPRTY, Estados , csrelpath a, tema t, prodserv  st "
		'strsql = strsql & " where  OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and OBJPRTY.estado = Estados.objid and OBJETOS.codobjeto = OBJPRTY.codobjeto "
		'strsql = strsql & " and OBJPRTY.estado = " & sID & " and Objetos.objnivel >= 3 And Estados.objid In (9,1,6,2,5,3,4)  and Objetos.OBJUN = 1  and TIPOBJETO.TIPOCPT = 'D' "
		'strsql = strsql & " and (OBJETOS.APLIC = 'EM' or OBJETOS.APLIC = 'GLB')  and (a.cobj = OBJETOS.codobjeto)  and (a.tema = t.codtema) "
		'strsql = strsql & " and (a.stema = st.codprodserv) "
		'strsql = strsql & " group by  TIPOBJETO.descri "

		strsql2 = "select count(*), max(tipobjeto.descri) from objetos"
		strsql2 = strsql2 & " inner join objprty on (objprty.tipobjeto=objetos.tipobjeto and objprty.codobjeto= objetos.codobjeto and estado =" & sid & ")"
		If SESSION("BASE")="K" Then
			strsql2 = strsql2 & " inner join tipobjeto on (tipobjeto.tipobjeto = objprty.tipobjeto  AND TIPOBJETO.TIPOCPT = 'K'  and (OBJETOS.APLIC = '" & wkaplic & "' or OBJETOS.APLIC = 'GLB'))"
		Else
			strsql2 = strsql2 & " inner join tipobjeto on (tipobjeto.tipobjeto = objprty.tipobjeto  AND TIPOBJETO.TIPOCPT = 'D'  and (OBJETOS.APLIC = '" & wkaplic & "' or OBJETOS.APLIC = 'GLB'))"
		End If

		'If (Session("EMPMUN") = "S") Then
		'         strsql2 = strsql2  & " LEFT join OBJMEMPALC  on OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO  "
		' END IF
		' SQL_WHERE= SQL_WHERE & " AND   (Objprty.fecemis between '" & cdesde & "' and '" & chasta & "' ) "

		If (Session("EMPMUN") = "S") Then
			strsql2= strsql2 & "  where  OBJPRTY.estado = " & sID & " AND OBJETOS.OBJUN = " & Session("UN")
		End If
		strsql2 = strsql2 & " group by objetos.tipobjeto"
		'response.write strsql2

		Set RSX1 = OFV.crearconsultaEx(strsql2, cn, 1, parametros)

		If Not RSX1.EOF Then
			'' Titulo de Totales por Tipo -unica vez-

			Celda = ""
			Celda = OFV.GenCelda("", "vt xstyle='border-style:solid;border-width:5;border:none;'  ", "", "10%", "", "", "", "&nbsp;" & "TOTALES POR TIPO" & "", "si")
			Fila = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
			Tablac = OFV.GenTabla("", "aaa bgcolor='#eeeeee' border=1 xstyle='border-style:solid;border-width:1;border:none;' ", "", "10% ", "", "0", "", "1", "2", "", "", Fila, "si")
			Fila = ""

			'' Rutina para mostrar los totales por tipo
			Celda1=""
			Do While Not rsx1.eof
				Celda1 = OFV.GenCelda("", "vt xstyle='border-style:solid;border-width:5;border:none;'  ", "", "10%", "", "", "",rsx1(1) & ": " & rsx1(0), "si")
				Fila1 = Fila1 & OFV.GenRow("", "", "", "", "", "", "", Celda1, "si")
				Fila=""
				rsx1.movenext
			Loop
			Tablac1 = OFV.GenTabla("", "aaa bgcolor='#eeeeee' border=1 xstyle='border-style:solid;border-width:1;border:none;' ", "", "10% ", "", "0", "", "1", "2", "", "", Fila1, "si")
			sHTML = sHTML & Tablac & Tablac1

		End If
	End If

	'' AGREGADO 26052014 FIN

	'' sHTML = OFV.MENUHeader("","")

	'  response.write "<br><br><br><br>" & Request.form & "<br>"

	'  response.write " starget " & starget & "<br>"

	If MasterType <> "EXP" Then

		'Celda = OFV.GenCelda("", "ttx colspan=3", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=N target=_self ><B><u>&nbsp;ESQUEMA NORMATIVO&nbsp;</u></B></A>", "si")
		If SESSION("BASE")="K" Then
			Celda = OFV.GenCelda("", "ttx colspan=3", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=N target=_self ><B><u>&nbsp;FORMULARIO&nbsp;</u></B></A>", "si")
		Else

			Celda = OFV.GenCelda("", "ttx colspan=3", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=N target=_self ><B><u>&nbsp;ESQUEMA NORMATIVO&nbsp;</u></B></A>", "si")
		End If
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
			If sid3 = "Eliminacion" Then
				Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Eliminacion", "si")
			ElseIf sid3 = "Vencimiento" Then
				Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "REVISION", "si")
			ElseIf sid3 = "Publicacion" Then
				'*********Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Publicacion", "si")
				Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=P target=_self ><B><u>&nbsp;PUBLICACION&nbsp;</u></B></A>", "si")
			Else
				'*********Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Vigencia", "si")
				Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=V target=_self ><B><u>&nbsp;VIGENCIA&nbsp;</u></B></A>", "si")
			End If
		ElseIf smenu = "PORREVISION" Then
			Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "REVISION", "si")
		ElseIf smenu = "PORRESUMEN" Then
			'*******Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Publicacion", "si")
			Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=P target=_self ><B><u>&nbsp;PUBLICACION&nbsp;</u></B></A>", "si")
			'*******Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Carpeta", "si")
			Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "CARPETA", "si")

		Else
			'*******Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Publicacion", "si")
			Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=P target=_self ><B><u>&nbsp;PUBLICACION&nbsp;</u></B></A>", "si")
			'*******Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Vigencia", "si")
			Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "<a class=tt href=" & starget & "&sorden=V target=_self ><B><u>&nbsp;VIGENCIA&nbsp;</u></B></A>", "si")
		End If
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


		Celda = OFV.GenCelda("", "ttx colspan=2", "", "", "", "", "", "ESQUEMA NORMATIVO", "si")
		Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "VERSION", "si")
		Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "TIPO", "si")
		If smenu = "PORESTADO" Then
			If sID = 4 Then
				Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Eliminacion", "si")
			Else
				Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "CREACION", "si")
			End If
		Else
			Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "CREACION", "si")
		End If
		If smenu = "PORFECHA" Then
			If sid3 = "Eliminacion" Then
				Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Eliminacion", "si")
			ElseIf sid3 = "Vencimiento" Then
				Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "REVISION", "si")
			ElseIf sid3 = "Publicacion" Then
				Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Publicacion", "si")

			Else
				Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Vigencia", "si")
			End If
		ElseIf smenu = "PORREVISION" Then
			Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "REVISION", "si")
		ElseIf smenu = "PORRESUMEN" Then
			Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Publicacion", "si")
			Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "CARPETA", "si")

		Else
			Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Publicacion", "si")
			Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Vigencia", "si")
		End If
		If smenu <> "PORESTADO" Then
			Celda = Celda & OFV.GenCelda("", "tTx", "", "", "", "", "", "ESTADO", "si")
			Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "AUTOR", "si")
		Else

			Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Autor", "si")
			'  If sID = 3 Or sID = 4 Then
			Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Aprobo", "si")
			'  End If
			If sID = 4 Then
				Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Elimino", "si")
			End If
		End If

		If smenu = "PORRESUMEN" Then
			Celda = Celda & OFV.GenCelda("", "ttx", "", "", "", "", "", "Resumen", "si")
		End If

	End If



	Fila = Fila & OFV.GenRow("", "", "", "", "20", "", "", Celda, "si")


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
	'response.write strsql
	On Error Resume Next
	If Not rs.EOF Then

		If MasterType <> "EXP" Then
			rs.Move (CantRegMover)
		End If

		Do While (Not rs.EOF)

			dEmision = "..."
			dVigencia = "..."
			sArchivo = "Iconos/" & rs(11)
			If Not IsNull(rs(0)) Then stipo = rs(0)

			rwspan = ""
			If smenu = "PORRESUMEN" Then
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
				If sid3 = "Eliminacion" Then
					If Not IsNull(rs(7)) Then dVigencia = CDate(CLng(rs(7)) + fi)
				ElseIf sid3 = "Vencimiento" Then
					If Not IsNull(rs(16)) Then dVigencia = CDate(CLng(rs(16)) + fi)
				ElseIf sid3 = "Publicacion" Then
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

			If smenu <> "PORFECHA" And smenu <> "PORREVISION" Then
				If Not IsNull(rs(22)) Then dpublic = CDate(CLng(rs(22)) + fi)
				dVign = InStr(dpublic, "1900")
				If dVign > 0 Then dpublic = "..."
			End If

			If smenu = "PORRESUMEN" Then
				If Not IsNull(rs("FECPUBLIC")) Then dpublic = CDate(CLng(rs("FECPUBLIC")) + fi)
				dVign = InStr(dpublic, "1900")
				If dVign > 0 Then dpublic = "..."
			End If

			dVign = InStr(dVigencia, "1900")
			If dVign > 0 Then dVigencia = "..."

			If smenu <> "PORESTADO" Then
				If Not IsNull(rs(8)) Then
					If ( (rs("estcod") = "3" OR rs("estcod") = "VIGENTE" )and session("APLMDCLI") = "CFA")  and ((Not CDate(CLng(rs(16)) + fi) > Date) And (CDate(CLng(rs(16)) + fi) <> CDate("01-01-1900"))) Then
						sEstado = "VIGENTE"
					ElseIf (Not CDate(CLng(rs(16)) + fi) > Date) And (CDate(CLng(rs(16)) + fi) <> CDate("01-01-1900")) Then
						sEstado = "VENCIDA"
					Else
						sEstado = rs(8)
					End If
				End If
				'sEstado =  sEstado & "  " & session("APLMDCLI") & " "  & rs("estcod")
				sAutor = "&nbsp;"
				'SESTADO= SESTADO & "  "  & rs("estcod")
				If Not IsNull(rs(12)) Then sAutor = rs(12)
			Else
				sAutor = "&nbsp;"
				sAprobo = "&nbsp;"
				sElimino = "&nbsp;"

				If Not IsNull(rs(12)) Then sAutor = rs(12)
				If Not IsNull(rs(13)) Then sAprobo = rs(13)
				If Not IsNull(rs(14)) Then sElimino = rs(14)
			End If


			'
			If Not IsNull(rs(9)) Then nIden = rs(9)
			If Not IsNull(rs(10)) Then sobjeto = rs(10)

			If smenu <> "PORFECHA" And smenu <> "PORREVISION" Then

				If Not IsNull(rs(19)) Then sobjeto = sobjeto & " - " & rs(19)
				If Not IsNull(rs(20)) And rs(19) & "" = "V" Then sobjeto = sobjeto & " " & rs(20)

			Else

				If smenu = "PORFECHA" Then
					Call GETLISTOBJ("E", rs(3), 0)
					sobjeto = XGUTABLISDAT
				End If
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

				sLink = OFV.GenLink("", "", sHref, "", "", sFuncion, sImagen, "si")

			Else

				sLink = "-"

			End If

			Celda = OFV.GenCelda("", "ut " & rwspan, "", "3%", "", "", "", sLink, "si")

			smasc = "&nbsp;"
			If smenu <> "PORFECHA" And smenu <> "PORREVISION" Then
				If rs(21) = 1 And Session("UN") <> 1 And Session("EMPMUN") = "S" Then
					smasc = "&nbsp;<font color=#993333  ><font class=btic title='Documento Corporativo' >G</font></font>"
				End If
			End If

			If Session("APLMDCLI") = "BC" Then
				If MasterType <> "EXP" Then
					sLink = OFV.GenLink("", "nt", sHref, "", "", sFuncion, "<b>" & rs("codiden") & "&nbsp;-&nbsp;</b>" & sobjeto, "si")
				Else
					sLink = "<b>" & rs("codiden") & "&nbsp;-&nbsp;</b>" & sobjeto
				End If
			Else
				If MasterType <> "EXP" Then
					If smenu = "PORACTA" Then
						sLink = OFV.GenLink("", "nt", sHref, "", "", sFuncion, EsqnomPx(rs("tipobjeto"), rs("codobjeto"), ofv) & "<b> -- Acta :" & rs("nroacta") & "</b>", "si")
					Else
						sLink = OFV.GenLink("", "nt", sHref, "", "", sFuncion, EsqnomPx(rs("tipobjeto"), rs("codobjeto"), ofv), "si")
					End If

				Else
					sLink = EsqnomPx(rs("tipobjeto"), rs("codobjeto"), ofv)
					If smenu = "PORRESUMEN" Then
						slink2 = rs("version") & ""
						If slink2 <> "O" Then   slink2 = slink2 & rs("nrover") & ""
						smasc = slink2
					End If
				End If

			End If

			sTipoDsc = stipo
			If MasterType <> "EXP" Then
				sTipoDscx = "<a class=nt title='solo documentos de tipo " & stipo & "' href=" & starget & "&sorden=" & sorden & "&sfiltro=" & Replace(stipo, " ", "_") & " target=_self ><font color='"
				sTipoDscx = sTipoDscx & Session("CMFPNIPFT") & "' >" & sTipoDsc & "</font></A>"
				sTipoDsc = sTipoDscx
			End If

			Celda = Celda & OFV.GenCelda("", "ut " & rwspan, "", "35%", "", "", "", sLink, "si")
			Celda = Celda & OFV.GenCelda("", "ut " & rwspan, "", "4%", "", "", "", smasc, "si")
			Celda = Celda & OFV.GenCelda("", "ut ", "", "15%", "", "", "", "" & sTipoDsc, "si")
			Celda = Celda & OFV.GenCelda("", "ut ", "", "10%", "", "", "", "" & fechadorCor(dEmision, ""), "si")
			If smenu <> "PORFECHA" And smenu <> "PORREVISION" Then
				Celda = Celda & OFV.GenCelda("", "ut ", "", "10%", "", "", "", "" & fechadorCor(dpublic, ""), "si")
			End If

			If smenu = "PORRESUMEN" Then
				SPATHDOC = rs("temadescri") & "/" & rs("stemadescri")
				If (RS("C0")&"") <> "" Then  SPATHDOC = SPATHDOC & "/" & rs("C0")
				If (RS("C1")&"") <> "" Then  SPATHDOC = SPATHDOC & "/" & rs("C1")
				If (RS("C2")&"") <> "" Then  SPATHDOC = SPATHDOC & "/" & rs("C2")
				If (RS("C3")&"") <> "" Then  SPATHDOC = SPATHDOC & "/" & rs("C3")
				If (RS("C4")&"") <> "" Then  SPATHDOC = SPATHDOC & "/" & rs("C4")
				If (RS("C5")&"") <> "" Then  SPATHDOC = SPATHDOC & "/" & rs("C5")
				If (RS("C6")&"") <> "" Then  SPATHDOC = SPATHDOC & "/" & rs("C6")
				If (RS("C7")&"") <> "" Then  SPATHDOC = SPATHDOC & "/" & rs("C7")



				If MasterType <> "EXP" Then
					If Len(SPATHDOC) > 12 Then
						SPATHDOCD = Left(SPATHDOC,10) & "..."
					Else
						SPATHDOCD = SPATHDOC
					End If
				Else
					SPATHDOCD = SPATHDOC

				End If



				Celda = Celda & OFV.GenCelda("", "ut TITLE='" & SPATHDOC & "' ", "", "10%", "", "", "", SPATHDOCD, "si")
			Else
				Celda = Celda & OFV.GenCelda("", "ut ", "", "10%", "", "", "", "" & fechadorCor(dVigencia, ""), "si")
			End If



			If smenu <> "PORESTADO" Then
				If sEstado = "VENCIDA" Then
					sinput = "<font color=#880000>" & sEstado & "</font>"
				Else
					sinput = sEstado
				End If
				Celda = Celda & OFV.GenCelda("", "ut ", "", "5%", "", "", "", sinput, "si")
				Celda = Celda & OFV.GenCelda("", "ut ", "", "5%", "", "", "", sAutor, "si")
			Else
				Celda = Celda & OFV.GenCelda("", "ut ", "", "5%", "", "", "", sAutor, "si")
				If session("base") <> "K" Then
					sAprobo = "<a href=coobjwfldet.asp?Tipobjeto=" & rs("tipobjeto") & "&Codobjeto="
					sAprobo = sAprobo & rs("codobjeto") & " target=_blank><IMG SRC=Imagenes\ICorSInfoUser.jpg title=" & Chr(34)
					sAprobo = sAprobo & " Esquema&nbsp;de&nbsp;Autorizaci&oacute;n&nbsp;de&nbsp;" & Chr(34) & " border=0></a>"

					sRef = "coobjwfldet.asp?Tipobjeto=" & rs("tipobjeto") & "&Codobjeto=" & rs("codobjeto")
					sImagenP = "&#0232;"
					sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
					sinput = "<a class='ft1n' " & sFuncAnc & " href='" & sRef & "' "
					sinput = sinput & "' title='Esquema&nbsp;de&nbsp;Autorizaci&oacute;n&nbsp;de&nbsp;" & sobjeto & "' target='_blank' >" & sImagenP & "</a>"
					sAprobo = sinput

					Celda = Celda & OFV.GenCelda("", "ft", "", "5%", "", "", "", sAprobo, "si")
				Else

					Celda    = Celda & OFV.GenCelda("", "ft", "", "5%", "", "", "", "...", "si")
				End If
				If sID = 4 Then
					Celda = Celda & OFV.GenCelda("", "ut ", "", "5%", "", "", "", sElimino, "si")
				End If
			End If


			If smenu = "PORRESUMEN" Then
				SPATHDOCD = rs("objcopete") & "&nbsp;"
				If MasterType <> "EXP" Then

				Else
					Celda = Celda & OFV.GenCelda("", "ut ", "", "5%", "", "", "", SPATHDOCD, "si")
				End If

			End If



			Fila = Fila & OFV.GenRow("", "aaa bgcolor=" & xTrColor & " ", "", "", "22", "", "", Celda, "si")

			If smenu = "PORRESUMEN" And (MasterType <> "EXP") Then
				Celda = OFV.GenCelda("", "ut colspan=6 style='border-style:solid;border-width:1;border-color:#aaaaaa;' ", "", "5%", "", "", "", SPATHDOCD, "si")
				Fila = Fila & OFV.GenRow("", "aaa bgcolor=" & xTrColor & " ", "", "", "22", "", "", Celda, "si")

				SPATHDOCD = "<hr  style='border-style:solid;border-color:#aaaaaa;' >"
				Celda = OFV.GenCelda("", "ut colspan=9  ", "", "5%", "", "", "", SPATHDOCD, "si")
				Fila = Fila & OFV.GenRow("", "aaa bgcolor=" & xTrColor & " ", "", "", "", "", "", Celda, "si")

			End If



			If xTrColor = "#FFFFFF" Then
				xTrColor = "#F0F0F0"
			Else
				xTrColor = "#FFFFFF"
			End If

			PosReg = PosReg + 1
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


			If MasterType <> "EXP" Then
				If Not (PosReg < CantRegMostrar)  Then Exit Do

			End If


		Loop

	Else

		XFRASE = "<BR><BR>&nbsp;<B><font color=#993333 SIZE=3 FACE='corflaticon' >i</font></B>"
		XFRASE = XFRASE & "&nbsp;NO SE ENCONTRARON DOCUMENTOS PARA LA BUSQUEDA SOLICITADA<BR><BR>"

		Celda = OFV.GenCelda("", "ut colspan=2 ", "", "", "", "", "", XFRASE, "si")


		Fila = Fila & OFV.GenRow("", "aaa bgcolor=" & xTrColor & " ", "", "", "22", "", "", Celda, "si")


	End If


	If MasterType <> "EXP" Then
		sHTML = sHTML & OFV.GenTabla("", "", "", "100%  style={border-style:solid;border-width:1;border-color:#eeeeee;} ", "", "0", "", "-1", "0", "", "", Fila, "si") & SFORMEXP

	Else
		Response.Write OFV.GenTabla("", "", "", "100%  style={border-style:solid;border-width:1;border-color:#eeeeee;} ", "", "0", "", "-1", "0", "", "", Fila, "si")
	End If



	'response.write Addqry

	If smenu = "PORREVISION" Then
		sAccion = "COFECREVISION.asp?Menu=" & smenu & "&Nombre=" & sNombre
	Else
		sAccion = "Subconsulta.asp?Menu=" & smenu & "&Nombre=" & sNombre
	End If

	Celda = OFV.BotonVolver(sAccion)




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


End If


If MasterType <> "EXP" Then
	'response.write strsql
	response.write sHTML
	response.flush
Else

End If

Function vertexto(nwtexto)

	nwtextoF = nwtexto & ""

	If nwtextoF <> "" Then

		nwtextoA = Split(nwtextoF, " ")

		nwsep = ""
		nwtextoF = ""

		For iii2 = 0 To UBound(nwtextoA)

			If Len(nwtextoA(iii2)) > 2 Then

				nwtextoF = nwtextoF & nwsep & nwtextoA(iii2)

				nwsep = "%"
			End If
		Next


	End If



	vertexto = nwtextoF
End Function

Function fechaConvert(wkfecha)
	Nfecha=CDate(wkfecha)
	Nfecha=DateSerial(Year(wkfecha),Month(wkfecha),Day(wkfecha))
	fechaConvert=Nfecha
End Function

Function fechatostr(wkfecha)
	fechatostr=CDate(wkfecha)
	fechatostr=Day(wkfecha) & "/" & Month(wkfecha) & "/" & Year(wkfecha)

End Function

%>
