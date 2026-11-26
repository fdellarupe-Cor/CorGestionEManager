

<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->



<%

ok = OFV.CheckUsuario()
If ok <> "Y" Then
	uv = OFV.uv
	If uv = "NO" Then ok = OFV.ValidarUsuario
	shtml = ok
Else
	response.expires = 0
	CantRegMostrarx = 8
	response.buffer = True

	' on error resume next
	vuelta = OFV.vueltaasp("../", CantRegMostrarx, CantRegMoverx)

	Set CN = OFV.conectar(OFV.strconn2)

	sUN = 1
	'response.write request.querystring & " <br>"

	XXXX = NEWEMPRE("NO")


	sUN = request.querystring("EMPRE")
	'response.write "sEMPRE: " & SESSION("UN")
	If sUN = ""  And SESSION("UN") = "" Then

		sUN = ATREMP(0, 0)
		SESSION("UN") = sUN & ""
		If (SESSION("UN") & "") = "" Then SESSION("UN") = 1
	ElseIf sUN <> "" Then
		XXXX = NEWEMPRE(sUN)
		SESSION("UN")=sUN
	Else
		XXXX = NEWEMPRE(SESSION("UN"))
		sUN=SESSION("UN")
	End If


	TablaH = ofv.GenTabla("","","","20%","","0","0","0","25","","",selecEmpre(),"si")

	sUN = session("UN")
	'	RESPONSE.WRITE sUN
	Dim atit
	anchopag = "98%"
	Dim TIPOCPT
	Dim tipodoc

	search_phrase = Replace(Request.querystring("YIDER")& "",Chr(34),"")
	search_phrase=Replace(search_phrase,"'","")
	If search_phrase = "" Then
		search_phrase = Replace(Request.Form("buscar") & "",Chr(34),"")
		search_phrase=Replace(search_phrase,"'","")
	End If

	tipobsq = Request.querystring("b")
	FORMULARIOS=REQUEST.QUERYSTRING("FORMULARIOS")

	If FORMULARIOS <> "" Then
		session("formularios")=FORMULARIOS
	Else
		FORMULARIOS=session("formularios")
	End If
	'response.write FORMULARIOS
	ingreso=REQUEST.QUERYSTRING("ing")
	'RESPONSE.WRITE REQUEST.QUERYSTRING
	search_strings = Split(search_phrase , " ", -1, 1)

	xsdelim=""
	sresult=""
	For Each search_string In search_strings

		If  Len(search_string & "") > 3 Then

			sresult= sresult & xsdelim & search_string
			xsdelim= " "

		End If

	Next

	'  search_phrase=sresult 'trim(search_string)

	Session("yider") = search_phrase

	filtroEsq = ""
	'RESPONSE.WRITE ingreso
	If tipobsq <> ""  And ingreso <> "P" Then
		delim = ""
		DELIM2 = ""
		SESQ = ""
		For i = 1 To Request.Form("ESQUEMAS").Count

			filtroEsq = filtroEsq & delim & "tipo='" & Request.Form("ESQUEMAS").Item(i) & "'"
			SESQ = SESQ & DELIM2 & "'" & Request.Form("ESQUEMAS").Item(i) & "'"
			DELIM2 = ","
			delim = " or "
		Next


		delim = ""

		filtrost = ""
		For i = 1 To Request.Form("subtemas").Count

			filtrost = filtrost & delim & Request.Form("subtemas").Item(i)

			delim = ","

		Next

	End If
	'xxx=armaglos()


	xclase1 = "bietx"

	xclase2 = "bietxG"
	m_search_style = "font-family:Arial; font-size:12pt; color:#000000;font-weight: bold"
	m_search_style2 = "font-family:Arial; font-size:18pt; color:#000000;font-weight: bold"
	m_style_you_searched_for = "font-family:Arial; font-size:10pt; color:#000000"
	m_style_title = "text-decoration: underline;"
	'm_style_title = "font-size:9pt; font-family:Arial;color:#B43104;font-weight: bold;text-decoration: underline;font-style: italic"

	m_style_titlex = "font-size:15pt; font-family:Arial;color:#000000;font-weight: bold;text-decoration: underline;font-style: italic"
	m_style_text = "font-size:8pt; font-family:verdana;xfont-style: italic"
	m_style_noise_word = "font-size:10pt; font-family:Arial; color:#6f6f6f"
	m_style_url = "font-size:10pt; font-family:Arial; color:green"
	m_style_br = "font-size:10pt; font-family:Arial"
	m_style_more_results_text = "font-size:10pt; font-family:Arial;color:#000000;"
	m_style_more_results_link = "font-size:10pt; font-family:Arial;color:#000000;text-decoration: underline"
	m_style_more_results_link_Next = "font-size:12pt; font-family:Arial; font-weight:bold;color:#000000;text-decoration: underline"


	shtml = OFV.FormHeader(Session("FormN"))
	shtml = shtml & CorPnlFnc("N", "")
	shtml = shtml & "<script languaje=javascript src='incfile/tooltip.js'></script>"
	shtml = shtml & "<script languaje=javascript > var posp=1;var cantdoc=0;var topepag=0;pagant=1;"


	shtml = shtml & "function mostrarp(dato){" & Chr(13)
	'  shtml = shtml & " for(var id=1;id <= topepag ;id++){"
	shtml = shtml & " var o=document.getElementById('pag' + pagant);" & Chr(13)
	shtml = shtml & " o.style.display='none';"
	';}"

	shtml = shtml & " var o=document.getElementById('pag' + dato);" & Chr(13)
	shtml = shtml & " o.style.display='';"
	shtml = shtml & " var o=document.getElementById('lnk' + dato);" & Chr(13)
	shtml = shtml & " o.style.fontSize='15pt';o.style.color='#000000';"
	shtml = shtml & " var o=document.getElementById('lnk' + pagant);" & Chr(13)
	shtml = shtml & " o.style.fontSize='10pt';pagant=dato;o.style.color='#B43104';"
	shtml = shtml & " var op=document.getElementById('infoP');" & Chr(13)
	shtml = shtml & " op.innerHTML='<h5>Pagina ' + dato + ' de ' + topepag + '  - Total Documentos encontrados : ' + cantdoc +'<h5>' ;" & Chr(13)



	shtml = shtml & "}"


	shtml = shtml & "function mostrarpav(dato){" & Chr(13)
	'  shtml = shtml & " for(var id=1;id <= topepag ;id++){"
	shtml = shtml & " var o=document.getElementById('pag' + pagant);" & Chr(13)
	shtml = shtml & " o.style.display='none';"
	';}"

	shtml = shtml & " var o=document.getElementById('pag' + dato);" & Chr(13)
	shtml = shtml & " o.style.display='';"
	shtml = shtml & " var o=document.getElementById('lnkav' + dato);" & Chr(13)
	shtml = shtml & " o.style.fontSize='15pt';o.style.color='#000000';"
	shtml = shtml & " var o=document.getElementById('lnkav' + pagant);" & Chr(13)
	shtml = shtml & " o.style.fontSize='10pt';pagant=dato;o.style.color='#B43104';"
	shtml = shtml & " var op=document.getElementById('infoPAV');" & Chr(13)
	shtml = shtml & " op.innerHTML='<h5>Pagina ' + dato + ' de ' + topepag + '  - Total Documentos encontrados : ' + cantdoc +'<h5>' ;" & Chr(13)



	shtml = shtml & "}"



	shtml = shtml & "function mostrarpox(dato){" & Chr(13)
	shtml = shtml & " if (dato == pagant){ return; }"
	shtml = shtml & " var o=document.getElementById('pag' + dato);" & Chr(13)
	shtml = shtml & " o.style.display='';"
	shtml = shtml & " var o=document.getElementById('pag' + pagant);" & Chr(13)
	shtml = shtml & " o.style.display='none';pagant=dato;"
	shtml = shtml & " var op=document.getElementById('infoP');" & Chr(13)
	shtml = shtml & " op.innerHTML='<h5>Pagina ' + dato + ' de ' + topepag + '  - Total Documentos encontrados : ' + cantdoc +'<h5>' ;" & Chr(13)



	shtml = shtml & "}"

	shtml = shtml & "function mostrarelem(dato,dato2){" & Chr(13)

	shtml = shtml & " var o=document.getElementById(dato);" & Chr(13)
	shtml = shtml & " o.style.display='';"
	shtml = shtml & " var o=document.getElementById(dato2);" & Chr(13)
	shtml = shtml & " o.style.display='none';"
	shtml = shtml & "}"


	shtml = shtml & "function mostrar(dato){" & Chr(13)
	shtml = shtml & "if(dato=='S'){if (posp==topepag) {  return;} " & Chr(13)
	shtml = shtml & " var o=document.getElementById('pag' + posp);" & Chr(13)

	shtml = shtml & "o.style.display='none';posp++;pagant=posp;" & Chr(13)
	shtml = shtml & " if (posp==topepag) {document.getElementById('btsig').className='bietxG';}"

	shtml = shtml & "var o=document.getElementById('pag' + posp);" & Chr(13)
	shtml = shtml & " var op=document.getElementById('infoP');" & Chr(13)
	shtml = shtml & " op.innerHTML='<h5>Pagina ' + posp + ' de ' + topepag + '  - Total Documentos encontrados : ' + cantdoc +'<h5>';" & Chr(13)
	shtml = shtml & "document.getElementById('btant').className='bietx';" & Chr(13)

	shtml = shtml & "o.style.display='';}"
	shtml = shtml & "else{ if (posp==1) { return;}" & Chr(13)
	shtml = shtml & " var o=document.getElementById('pag' + posp);" & Chr(13)

	shtml = shtml & "o.style.display='none';posp--;pagant=posp;" & Chr(13)
	shtml = shtml & " if (posp==1) {document.getElementById('btant').className='bietxG';}"
	shtml = shtml & " if (posp < topepag) {document.getElementById('btsig').className='bietx';}"
	shtml = shtml & " var op=document.getElementById('infoP');" & Chr(13)
	shtml = shtml & " op.innerHTML='<p><h5>Pagina ' + posp + ' de ' + topepag + '  - Total Documentos encontrados : ' + cantdoc +'<h5></p>';" & Chr(13)
	shtml = shtml & "var o=document.getElementById('pag' + posp);" & Chr(13)
	shtml = shtml & "o.style.display='';"

	shtml = shtml & "}}"
	shtml = shtml & "  function elim(str,len)"
	shtml = shtml & "  {"


	shtml = shtml & "  str = str.split(/[\s,;.?!:\n//]+/).filter(function ( str ) {"
	shtml = shtml & "      var word = str.match(/(\w+)/);"
	shtml = shtml & "      return word && word[0].length > len;"
	shtml = shtml & "   }).join(' ');"

	shtml = shtml & "   return str;"
	shtml = shtml & "  }"

	shtml = shtml & "</script>"


	sql = "select orden,b.valor1,a.descri from filtrobsq b  "
	sql = sql & " inner join tipobjeto a on(a.tipobjeto COLLATE DATABASE_DEFAULT =b.valor1 COLLATE DATABASE_DEFAULT )"

	sql = sql & " where b.tipo='ESQUEMA' ORDER BY b.ORDEN"

	Set rsESQ = OFV.crearconsultaEx(sql, CN, 1, volver)
	SESQUEMAS = "Esquemas normativos:  "

	delim = ""
	DELIM2 = ""
	orden_doc = ""
	nohay2=False
	If rsESQ.EOF Then
		nohay2=True
		orden_doc = orden_doc & " WHEN ''  THEN 0"  & Chr(13)
	End If
	sinputesquemas = "<select multiple name=esquemas style='display:none'>"
	aesq = Split(SESQ, ",")

	Do While Not rsESQ.EOF

		orden_doc = orden_doc & " WHEN '" & rsESQ(1) & "'  THEN " & rsESQ(0) & Chr(13)
		If Request.Form("ESQUEMAS").Count = 0 Then
			filtroEsq = filtroEsq & delim & "tipo='" & rsESQ(1) & "'"
			SESQ = SESQ & DELIM2 & "'" & rsESQ(1) & "'"
			DELIM2 = ","
			delim = " or "
			sinputesquemas = sinputesquemas & " <option value=" & rsESQ(1) & " >" & rsESQ(2) & "</option>"

		Else
			For i = 1 To Request.Form("ESQUEMAS").Count

				If Request.Form("ESQUEMAS").Item(i) = rsESQ(1) Then
					sinputesquemasZ = " <option value=" & rsESQ(1) & " selected >" & rsESQ(2) & "</option>"
					Exit For
				Else
					sinputesquemasZ = " <option value=" & rsESQ(1) & " >" & rsESQ(2) & "</option>"
				End If

			Next
			sinputesquemas = sinputesquemas & sinputesquemasZ
		End If


		rsESQ.movenext
	Loop
	sinputesquemas = sinputesquemas & "</select> "
	Call OFV.cerrarconsulta(rsESQ)



	If session("APLMDCLI")="BPN" Then
		sql = "select 0,tipobjeto from tipobjeto  "
	Else
		sql = "select archivo,tipobjeto from tipobjeto  "
	End If

	Set rsESQ = OFV.crearconsultaEx(sql, CN, 1, volver)

	delim = ""
	orden_docFull = ""


	Do While Not rsESQ.EOF
		'  SESQUEMAS=SESQUEMAS & delim & rsESQ("descri")
		Iorden=Trim(rsESQ(0) & "")
		If  Iorden ="" Or Not IsNumeric(Iorden) Then
			Iorden="0"
		End If
		orden_docFull = orden_docFull & " WHEN '" & rsESQ(1) & "'  THEN " & Iorden & Chr(13)

		delim = " or "
		rsESQ.movenext
	Loop


	Call OFV.cerrarconsulta(rsESQ)
	sql = "select b.codprodserv ,  c.descri  +'/' + b.descri as descri2"

	sql = sql & "  from filtrobsq a"

	sql = sql & "  inner join     PRODSERV b on(a.valor1=b.codPRODSERV)"
	sql = sql & "  inner join     tema c  on( c.codtema=b.codtema)"
	sql = sql & "   where a.tipo='ST'"
	sql = sql & " ORDER BY ORDEN "



	Set rsESQ = OFV.crearconsultaEx(sql, CN, 1, volver)


	delim = ""
	' filtrost=""
	If rsESQ.EOF Then
		sinputSUBTEMAS=""
		nohay=True
	Else
		nohay=False
		delim = ""
		sinputSUBTEMAS = "<select multiple name=SUBTEMAS >"
		Do While Not rsESQ.EOF
			If Request.Form("SUBTEMAS").Count = 0 Then
				filtrost = filtrost & delim & rsESQ(0)
				delim = ","

				sinputSUBTEMAS = sinputSUBTEMAS & " <option value=" & rsESQ(0) & ">" & rsESQ(1) & "</option>"
			Else
				For i = 1 To Request.Form("SUBTEMAS").Count

					If CStr(Request.Form("SUBTEMAS").Item(i) & "") = CStr(rsESQ(0) & "") Then
						sinputesquemasZ = " <option value=" & rsESQ(0) & " selected >" & rsESQ(1) & "</option>"
						Exit For
					Else
						sinputesquemasZ = " <option value=" & rsESQ(0) & " >" & rsESQ(1) & "</option>"
					End If

				Next
				sinputSUBTEMAS = sinputSUBTEMAS & sinputesquemasZ
			End If


			rsESQ.movenext
		Loop
		sinputSUBTEMAS = sinputSUBTEMAS & "</select> <HR>"
	End If
	Call OFV.cerrarconsulta(rsESQ)


	If tipobsq = "" Then
		tablab = "<DIV id='DTBUSCAR' STYLE='float: left; z-index:1;overflow:auto;width:98%;border-style:solid;border-width:2;border-color:#ffffff;' >"
	Else
		tablab = "<DIV id='DTBUSCAR' STYLE='float: left; z-index:1;overflow:auto;width:98%;border-style:solid;border-width:2;border-color:#ffffff;display:none;' >"
	End If
	sAccion = "searchex.asp?FORMULARIOS=" & FORMULARIOS & "&ing=" & ingreso
	sInputT = "<input type='text' class='dt' name='BUSCAR' id='textot' value=""" & search_phrase & """ size='80' xstyle='height:20px;border-style:solid;border-width:1;border-color:#666666;' >"
	'BtnAva = "<input  id=btsig type=BUTTON  size=10%  value='Ir a Busqueda avanzada' class=" & xclase1 & "   onclick=""mostrarelem('DTBUSCARav','DTBUSCAR');""   >"
	If nohay And nohay2 Then
		BtnAva=""
	Else
		BtnAva = "<input  id=btsig type=BUTTON  size=10%  value='Ir a Busqueda avanzada' class=" & xclase1 & "   onclick=""mostrarelem('DTBUSCARav','DTBUSCAR');""   >"
	End If
	sInput = OFV.GenerarInput("", "Buscar", "submit", "", "5", "bietx xSTYLE='color:#ffffff;font-weight:bold;background-color:#990000;border-style:solid;border-width:1;border-color:#dddddd;height:20px;cursor:hand;' ", "")
	If ingreso="P" Then
		sInput = sInput & "&nbsp;&nbsp;&nbsp;" & OFV.GenerarInput("", "Nueva Busqueda", "button", "", "5", "at   onclick=document.Subconsultaz.submit(); STYLE='color:#ffffff;font-weight:bold;background-color:#990000;border-style:solid;border-width:1;border-color:#dddddd;height:20px;cursor:hand;' ", "")
	Else

	End If
	tablas="<table><tr><td>" & sInputT & "</td><td>" & sInput &  "</td></tr></table>"
	formx = OFV.GenForm("Subconsulta", "", sAccion, sTarget, "post", "", tablas, "si")
	formx = formx & "<DIV ID=PAGINAS><h5>Calculando...<h5></DIV>"
	If ingreso="P" Then
		SHREF="SCONSUL.ASP?Ing=P&BASE=K&seg=CLIMN51812732031257124269088&MNN=CONSULTA_GENERAL_DE_FORMULARIOS"
		formxvuelve = OFV.GenForm("Subconsultaz style='display:none;'", "", SHREF, sTarget, "post", "", "", "si")
	Else
		formxvuelve=""
	End If
	celda = OFV.GenCelda("", "vt valign=top", "", "", "", "", "", formx & formxvuelve, "si")

	celda = celda & OFV.GenCelda("", "ss valign=top align=center", "", "70%", "", "", "", BtnAva & "<div style=" & m_search_style2 & "><b><p><div id=infoP ><h5>Calculando...<h5></div></div>", "si")

	siguiente = "<input  id=btsig type=BUTTON  size=10%  value='PAG. SIGUIENTE' class=" & xclase1 & "   onclick=""mostrar('S');""   >"
	ANTERIOR = "<input  id=btant type=BUTTON  size=10%  value='PAG. ANTERIOR' class=" & xclase2 & "   onclick=""mostrar('A');""   >"
	'  celda = celda & ofv.GenCelda("", "vt valign=top align=right", "", "70%", "", "", "", ANTERIOR & "&nbsp;&nbsp;&nbsp;" &siguiente, "si")
	'celda = celda & ofv.GenCelda("", "vt", "", "70%", "", "", "", siguiente, "si")
	sTarget = ""






	Fila = OFV.GenRow("", "", "", "", "", "", "", celda, "si")
	' celda=ofv.GenCelda("", "vt colspan=3 valign=top ", "", "", "", "", "",  "<DIV ID=PAGINAS><h5>Calculando...<h5></DIV>", "si")
	' Fila = fila &  ofv.GenRow("", "", "", "", "", "", "", celda, "si")

	tablab = tablab & OFV.GenTabla("", "aaa     STYLE='border-style:solid;border-width:1;border-color:#ffcc00;' ", "", "", "", "0", "", "0", "0", "", "", Fila, "si")

	tablab = tablab & "<br>" & "</DIV>" & Chr(13)
	If tipobsq = "" Then
		tablabavan = "<DIV id='DTBUSCARav' STYLE='float: left; z-index:1;overflow:auto;width:98%;border-style:solid;border-width:2;border-color:#ffffff;display:none;' >"
	Else
		tablabavan = "<DIV id='DTBUSCARav' STYLE='float: left; z-index:1;overflow:auto;width:98%;border-style:solid;border-width:2;border-color:#ffffff;' >"
	End If
	sAccion = "searchex.asp?B=BA&FORMULARIOS=" & FORMULARIOS & "&ing=" & ingreso
	sInputT = "<input type='text' class='dt' name='BUSCAR' id='textot' value=""" & search_phrase & """ size='80' style='height:20px;border-style:solid;border-width:1;border-color:#666666;' >"
	BtnAva = "<input  id=btsig type=BUTTON  size=10%  value='Ir a Busqueda Simple' class=" & xclase1 & "   onclick=""mostrarelem('DTBUSCAR','DTBUSCARav');""   >"
	sInput = OFV.GenerarInput("", "Buscar", "submit", "", "5", "bietx xSTYLE='color:#ffffff;font-weight:bold;background-color:#990000;border-style:solid;border-width:1;border-color:#dddddd;height:20px;cursor:hand;' ", "")
	tablas="<table><tr><td>" & sInputT & "</td><td>" & sInput &  "</td></tr><tr><td colspan=2>"& sinputesquemas & sinputSUBTEMAS &"</td></tr></table>"
	formx = OFV.GenForm("Subconsulta", "", sAccion, sTarget, "post", "", tablas , "si")
	'formx = OFV.GenForm("Subconsulta", "", sAccion, sTarget, "post", "", sInputT & sInput & "<br>" & sinputesquemas & sinputSUBTEMAS , "si")
	formx = formx & "<DIV ID=PAGINASAV><h5><h5></DIV>"
	celda = OFV.GenCelda("", "vt valign=top", "", "", "", "", "", formx, "si")

	celda = celda & OFV.GenCelda("", "ss valign=top align=center", "", "70%", "", "", "", BtnAva & "<div style=" & m_search_style2 & "><b><p><div id=infoPAV ><h5><h5></div></div>", "si")


	sTarget = ""






	Fila = OFV.GenRow("", "", "", "", "", "", "", celda, "si")





	celda = OFV.GenCelda("", "valign=top ", "", "", "", "", "", "temas", "si")




	celda = celda & OFV.GenCelda("", " valign=top ", "", "", "", "", "", "<b>" & UCase(SESQUEMAS) & "</b>", "si")

	'  Fila = fila &  ofv.GenRow("", "", "", "", "", "", "", celda, "si")

	tablabavan = tablabavan & OFV.GenTabla("", "aaa   STYLE='border-style:solid;border-width:1;border-color:#ffcc00;' ", "", "", "", "0", "", "0", "0", "", "", Fila, "si")

	tablabavan = tablabavan & "<br>" & "</DIV>" & Chr(13)
	If session("APLMDCLI") <> "GPa" Then
		response.write shtml  & tablah  & tablab & tablabavan  & ofv.botoneraESSX("0","","","","","","","","","no","","","")

		response.flush
	End If
	Formulario = ""    'ofv.Encabezados("info")




	Set findd = Server.CreateObject("CorsearchExXmlfiltrov5.Xmlado")

	Set cnbsq = OFV.conectar(strconn8)

	Set findd.OFV = OFV
	Set findd.Session = Session
	Set findd.response = response
	Set findd.Server = Server
	Set findd.Cnb = cnbsq


	findd.canttx = 300
	findd.cantreg = 100

	findd.search_phrase = EliminarAcentos(search_phrase & "")

	If tipobsq <> "" Then
		findd.orden_doc = " ,CASE y.tipo "

		findd.orden_doc = findd.orden_doc & orden_doc
		findd.orden_doc = findd.orden_doc & " Else "
		findd.orden_doc = findd.orden_doc & " 99 "
		findd.orden_doc = findd.orden_doc & " End "
		findd.orden_doc = findd.orden_doc & " AS ORDEN_DOC "
	Else
		findd.orden_doc = " ,CASE y.tipo "
		findd.orden_doc = findd.orden_doc & orden_docFull
		findd.orden_doc = findd.orden_doc & " Else "
		findd.orden_doc = findd.orden_doc & " 99 "
		findd.orden_doc = findd.orden_doc & " End "
		findd.orden_doc = findd.orden_doc & " AS ORDEN_DOC "
	End If


	findd.canttx = 300
	findd.cantreg = 50
	findd.listacar="{}[]!#$%&/()=?¡'¿|*+¨´:.;,<>-" ' sirve para fulltext con palabras con caracteres especiales ejemplo 123456789-0 guion es un caracter especial.le sacamos el * para buscar exacto


	' filtro si es formularios o documentos
	'
	If FORMULARIOS="K" Then

		FINDD.FILTROS=""

		If tipobsq ="BA" Then
			If filtrost <> "" Then
				FINDD.FILTROS="AND F.P3 in(" & filtrost & ")"
			End If
			If SESQ <> "" Then
				FINDD.FILTROS=FINDD.FILTROS & " AND F.tipobjeto in(" & SESQ & ")"
			End If

		End If
		FINDD.FILTROS= FINDD.FILTROS & " AND F.P1='K'"

		FESTADO         = CStr(request.form("estado") & "")

		FBANCA          = CStr(request.form("banca")& "")
		FGUARDA         = CStr(request.form("TGUARDA")& "")
		FPRODSERV       = CStr(request.form("PRODU-SERVI")& "")
		FCATEGORIA      = CStr(request.form("CATEGORIAF")& "")
		FCLASE          = CStr(request.form("CLASEF")& "")
		FCODIGO         = CStr(request.form("CODIGOF")& "")
		FFIRMAS         = CStr(request.form("FIRMASF")& "")
		IMPRIME         = CStr(request.form("IMPRESION")& "")
		TFECHA          = CStr(request.form("FECHA")& "")
		FDESDE          = CStr(request.form("DESDE2")& "")
		FHASTA          = CStr(request.form("HASTA2")& "")
		FPROCESO          = CStr(request.form("PROCESO")& "")



		'response.write request.form

		NEWSTRSQL=""

		If FPRODSERV ="" Or FPRODSERV="0" Then
		Else
			NEWSTRSQL= NEWSTRSQL & "  AND p9='"& FPRODSERV & "'"
		End If

		If FGUARDA ="" Or FGUARDA="0" Then
		Else
			NEWSTRSQL= NEWSTRSQL & "  AND p8='"& FGUARDA & "'"
		End If



		If FCATEGORIA ="" Or FCATEGORIA="0" Then
		Else
			NEWSTRSQL= NEWSTRSQL & " AND p11='" & FCATEGORIA & "' "
		End If

		If FPROCESO ="" Or FPROCESO="0" Then
		Else
			NEWSTRSQL= NEWSTRSQL & " AND p10='" & FPROCESO & "' "
		End If

		If FCLASE ="" Or FCLASE ="0" Then
		Else
			NEWSTRSQL= NEWSTRSQL & " AND p12='"& FCLASE &"' "
		End If

		If FCODIGO ="" Or FCODIGO ="0" Then
		Else
			NEWSTRSQL= NEWSTRSQL & " and p5='"& FCODIGO &"' "
		End If

		If FFIRMAS ="" Or FFIRMAS ="0" Then
			' NEWSTRSQL= NEWSTRSQL & " AND FIRMASFORM> '0'"
		Else
			NEWSTRSQL= NEWSTRSQL & " AND p13='"& FFIRMAS &"' "
		End If

		If IMPRIME ="" Or IMPRIME ="0" Then
		Else
			NEWSTRSQL= NEWSTRSQL & " AND p14='"& IMPRIME &"' "
		End If


		If FBANCA ="" Or FBANCA ="0" Then
		Else
			NEWSTRSQL= NEWSTRSQL & " AND p7='"& FBANCA &"' "
		End If


		If TFECHA = "NO" Or tfecha = "" Then
		Else

			NEWSTRSQL= NEWSTRSQL & "  AND p15  > = '"& FDESDE &"' AND p15 <= '"& FHASTA &"' "
		End If

		FINDD.FILTROS=FINDD.FILTROS &  NEWSTRSQL


		'response.write FINDD.FILTROS
	Else
		FINDD.FILTROS=""
		If tipobsq ="BA" Then
			If filtrost <> "" Then
				FINDD.FILTROS=" AND F.P3 in(" & filtrost & ")"
			End If

			If SESQ <> "" Then
				FINDD.FILTROS=FINDD.FILTROS & " AND F.tipobjeto in(" & SESQ & ")"
			End If
		End If
		FINDD.FILTROS= FINDD.FILTROS & " AND F.P1='D'"
	End If

	' findd.ADDINNER=" inner join (SELECT DISTINCT XX.CODOBJETO  ,XX.TIPOBJETO  FROM colsbsq XX WHERE  NOMCOL='EMP' AND TIPOC='U' AND VALORCOL='"& sun &"') Z "
	' findd.ADDINNER =findd.ADDINNER & " ON Z.CODOBJETO=F.CODOBJETO and Z.TIPOBJETO=F.TIPOBJETO"
	If Session("EMPMUN")="S" Then
		VALOR1=Chr(34) & "[[OBJUN_"  & sun & "]]" & Chr(34)

		VALOR2=" OR "  &  CHR(34)   & "[[CODEMPRE_" & sUN &"]]" & CHR(34)

		findd.ADDINNER=" inner join (select distinct codobjeto,tipobjeto "
		findd.ADDINNER =findd.ADDINNER & "from containstable(ObjMetadata, metadata, '" & VALOR1  & VALOR2 & "') as ct "

		findd.ADDINNER =findd.ADDINNER & "join ObjMetadata as y on ct.[key]=y.[objid] ) z "
		findd.ADDINNER =findd.ADDINNER & " ON Z.CODOBJETO=F.CODOBJETO and Z.TIPOBJETO=F.TIPOBJETO"

		findd.ADDINNER= " inner join (select  codobjeto,tipobjeto,metadata from containstable(ObjMetadata, metadata, '"& VALOR1  & VALOR2 & "') as ct2 "
		findd.ADDINNER= findd.ADDINNER & " inner join ObjMetadata as y on ct2.[key]=y.[objid] "
		findd.ADDINNER= findd.ADDINNER & " ) as ct3 on ct3.codobjeto=t1.codigo and ct3.tipobjeto=t1.tipo "

	End If

	'and AUTOR_ADM and CODSECTOR_132 and CODSECTOR_41


	'response.write  Session("vvv")
	If findd.IsNoiseWordex(findd.search_phrase) = 0 Then

		xinfo = "<h5>Sin Resultados..........</h5>"
		shtml = "<script>"


		shtml = shtml & " var op1=document.getElementById('infoP')" & Chr(13)

		shtml = shtml & " var op2=document.getElementById('infoPAV');" & Chr(13)

		shtml = shtml & " op1.innerHTML='"& xinfo&"';"& Chr(13)
		shtml = shtml & " op2.innerHTML='"& xinfo&"';"& Chr(13)

		shtml = shtml & "var ops1=document.getElementById('PAGINAS');ops1.innerHTML='"&xinfo&"'; "

		shtml = shtml & "var ops2=document.getElementById('PAGINASAV');ops2.innerHTML='"&xinfo&"'; "




		shtml = shtml & " </script> "

	Else

		' response.write time() & "  1<br>"
		'Call findd.findDoc
		Call findd.findDocex
		'response.write time() & " 2<br>"
		'response.write findd.search_phrase
		'RESPONSE.WRITE findd.sqlj
		'Set rsc = findd.orsFx
		Set rsc =  findd.rsext 'cnbsq.execute(findd.sqlj)
		'response.write time() & " 3<br>"
		If tipobsq <> "" Then
			'rsc.SORT = "ORDEN_DOC,rank desc,codigo,ranksol desc"



			'   If filtroEsq <> "" Then
			'       rsc.Filter = 0

			'      rsc.Filter = filtroEsq

			'  End If



			'   If filtrost <> "" Then

			'     rsc.Filter = filtrost
			' End If

			If sINDOC Then
				rsc.Filter = "tipo='ZXZZZZX'"
			End If

		Else
			'rsc.SORT = "ORDEN_DOC,codigo,rank desc ,ranksol desc"
			'	RSC.SORT="ORDEN_DOC,rank desc,codigo ,ranksol desc"
		End If

		xo = 0
		url = "#Foo"
		totaldoc = 1
		ipag = 1
		Formularioz = ""
		' rsc.move 11
		' rsc.movenext 11
		solapa10 = False
		solapa5 = False

		If session("APLMDCLI") = "BV" Then
			limitpag=500
		Else
			limitpag=11
		End If
		'	response.write time() & " 4<br>"
		Select  Case  session("APLMDCLI")

			Case "BPx" ' nueva visualizacion
			XXXX = NEWEMPRE("NO")

			session("ASP")="searchex.ASP"

			sEMPRE = request.querystring("EMPRE")

			If SEMPRE = ""  And SESSION("UN") = "" Then

				SEMPRE = ATREMP(0, 0)
				SESSION("UN") = SEMPRE & ""
				If (SESSION("UN") & "") = "" Then SESSION("UN") = 1
			ElseIf sEMPRE <> "" Then
				XXXX = NEWEMPRE(sEMPRE)
				SESSION("UN")=SEMPRE
			Else
				XXXX = NEWEMPRE(SESSION("UN"))
				sempre=SESSION("UN")
			End If


			shtml = OFV.FormHeader(Session("FormN"))

			If  Session("EMPMUN") = "S" Then

				Fila    = ""

				SINPUT3 = "&nbsp;<A class=tt title='SELECCIONE EMPRESA' HREF='javascript:void(0);' onclick='shwDem();' >EMPRESA:&nbsp;&nbsp;</a>"

				Fila    = ""
				Celda   = ""
				sHref   = ""

				sImage  = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(0,5),"0","")
				sLink   = ofv.GenLink("","ttt",sHref,"ftemas","","",sImage,"si")
				Celda   = Celda & ofv.GenCelda("","tt ","","8%","","","",SINPUT3,"si")
				Celda   = Celda & ofv.GenCelda("","ttt ","center","5%","","","",sLink,"si")
				Celda   = Celda & ofv.GenCelda("","ttt  ","","","","","","&nbsp;<BIG>" & ATREMP(0,4) & "</BIG>","si")

				Fila    = fila & ofv.GenRow("","","top","","","","",Celda,"si")

				Tabla   = ofv.GenTabla("","aaa style='"  & Session("CBRDABL") & "' ","","100% ","","0","","-1","-2","","",Fila,"si")





				shtml=shtml & "<script language=javascript>" & chr(13)
				shtml=shtml & "function shwDem(){" & chr(13)
				shtml=shtml & "document.all.DEM1.style.display='block';" & chr(13)
				shtml=shtml & "}" & chr(13)
				shtml=shtml & "</script>" & chr(13)
				'----------------------------------------------
				sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
				sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
				sinput2 = sinput2  & "none" & chr(34) & ";' href='javascript:void(0);' ONCLICK=" & chr(34) & "document.all.DEM1.style.display='none';" & chr(34)
				sinput2 = sinput2  & " TITLE='OCULTAR' ><b>" & sImagenx & "</b></a>"
				TABLA = Tabla & "<DIV id='DEM1' STYLE='display:none;z-index:1;margin-top:1;margin-left:1;background-color:#ffffff;position:absolute;left:1px;top:240px;left:20px;height:250px;width:350px;border-style:inset;border-color:#f7aa22;border-width:2;overflow:auto;'>"


				TABLAX = TablaX & OFV.gentabla("", "aaa style='filter:shadow(color=#dddddd,direction=125);' ", "", "100%", "", "0", "", "-2", "-1", "", "", "", "no")
				TABLAX = TablaX & "<tr><td  "
				TABLAX = TablaX & "class=vT align=left colspan=3 STYLE='border-style:outset;border-width:2;background-color:#eeeeee;' >" & sinput2 & "&nbsp;SELECCIONE EMPRESA</td></tr>"
				TABLAX = TablaX & "<tr><td colspan=3 "
				TABLAX = TablaX & " >&nbsp;</td></tr>"

				For IAEMP=1 To TAEMP

					'''  RESPONSE.WRITE  ATREMP(IAEMP,2)

					If ATREMP(IAEMP,2) = 1 Then

						If ATREMP(IAEMP,0) = ATREMP(0,0) Then
							TABLAX = TablaX & "<tr BGCOLOR=#EEEEEE >"
						Else
							TABLAX = TablaX & "<tr>"
						End If

						TABLAX = TablaX & "<td width=5 "
						If ATREMP(IAEMP,0) = ATREMP(0,0) Then
							TABLAX = TablaX & "class=vt  >" & sImagen4si & "</td>"
						Else
							TABLAX = TablaX & "class=vt  >" & sImagenI3 & "</td>"
						End If

						sImage = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(IAEMP,5),"0","")
						TABLAX = TablaX & "<td width=5 "
						TABLAX = TablaX & "class=vt  >" & sImage & "</td>"

						TABLAX = TablaX & "<td class=vt  >"
						'ing=&FORMULARIOS=D&YIDER=cuenta&
						'    search_phrase = Replace(Request.querystring("YIDER")& "",Chr(34),"")
						'	search_phrase=Replace(search_phrase,"'","")
						'	If search_phrase = "" Then
						'		search_phrase = Replace(Request.Form("buscar") & "",Chr(34),"")
						'		search_phrase=Replace(search_phrase,"'","")
						'	End If

						'	tipobsq = Request.querystring("b")
						'		FORMULARIOS=REQUEST.QUERYSTRING("FORMULARIOS")
						'ingreso=REQUEST.QUERYSTRING("ing")

						sHref = "searchex.asp?ing=" & ingreso & "&FORMULARIOS=" & FORMULARIOS & "&YIDER=" & search_phrase & "&EMPRE=" & ATREMP(IAEMP,0)

						'''  sHref = "OBJtreeTema.asp?Tipo=" & sTipo & "&Nombre=" & LNombre & "&Menu=" & SESSION("MENU") & "&EMPRE=" & ATREMP(IAEMP,0)
						sLink = ofv.GenLink("","cc",sHref,"","","","<b>" & ATREMP(IAEMP,4) & "</b>","si")
						TABLAX = TablaX & sLink
						TABLAX = TablaX & "</td>"
						TABLAX = TablaX & "</tr>"
					End If

				Next

				TABLAX = TablaX & OFV.genFtabla()

				TABLA = Tabla & TABLAX

				TABLA = Tabla & "</DIV>"

				sHTML = sHTML & TABLA

			End If
			If Session("EMPMUN") = "N" Then
				sHref2 = "searchex.asp?ing=" & ingreso & "&FORMULARIOS=" & FORMULARIOS & "&EMPRE=1"

			Else
				sHref2 = "searchex.asp?ing=" & ingreso & "&FORMULARIOS=" & FORMULARIOS & "&EMPRE=" & ATREMP(IAEMP,0)
			End If

			sHTML = sHTML & "<br><form method=post action='" & sHref2 & "'><input type=text name=buscar id=buscar value='" & search_phrase & "' class='ui-widget ui-state-default ui-corner-all'></input><input type=submit class='ui-widget ui-state-default ui-corner-all' value=Buscar></input></form>"
			shtml = shtml & CorPnlFnc("N", "")
			idf=0
			shtmldocs=""
			shtmtabD=""
			lipag=""
			tpag=1
			xentros=False
			Do While Not rsc.EOF
				iddoc=rsc("codigo")
				atit = titsolapa(rsc("codigo"), rsc("tipo"))
				shtmldocs= shtmldocs & "<h1><a class=nt>"  & EsqnomPx(rsc("tipo"), rsc("codigo"), ofv)  & "</a></h1><div>"
				'response.write iddoc & "<br>"
				Do While iddoc = rsc("codigo")
					rsc.movenext
					If rsc.eof Then
						xentros=True
						Exit Do
					Else
						xacceso = "codigo*" & rsc("codigo") & "_tipo*" & rsc("tipo") & "_perfil*OBJARC"
						xacceso = Replace(xacceso, "_", "&")
						If tipocpt="K" Then
							xacceso = Replace(xacceso, "*", "=") & "&NIV1=15&NIV2=" & rsc("index1")
						Else
							xacceso = Replace(xacceso, "*", "=") & "&NIV1=" & rsc("solapa") & "&NIV2=" & rsc("index1")
						End If
						sActcon = "Consultar.asp?Ysearch=S&DOCMODAL=CL&ACC=XP&" & xacceso & "&Datos=*smenu=PORTEXTO"
						clickn = Chr(34) & "AbrirCons('" & sActcon & "','" & rsc("tipo") & "');return false;" & Chr(34)
						shtmldocs = shtmldocs & "<button onclick=" & clickn & " class='ui-button ui-widget ui-corner-all'><a class=nt>"& GETSOLAPA(rsc("solapa"), rsc("index1")) &"</a></button>"
						'shtmldocs = shtmldocs & "<p>" &  GETSOLAPA(rsc("solapa"), rsc("index1"))  & "</p>"
						'iddoc=rsc("codigo")
					End If
				Loop

				shtmldocs=shtmldocs & "</div>"
				idf=idf +1
				If idf=10 Then
					lipag=lipag & "<li><a href='#tabs-"& tpag &"'>Pagina "& tpag &"</a></li>"


					shtmtabD = shtmtabD & " <div id='tabs-"& tpag &"'><div class='accordion'>"
					shtmtabD=shtmtabD & shtmldocs
					shtmtabD=shtmtabD &  "</div></div>"
					shtmldocs=""
					tpag=tpag +1
					idf=0
				End If

			Loop
			shtmtab="<div id='tabs'>"
			shtmtab=shtmtab & "  <ul>"
			shtmtab=shtmtab & lipag
			shtmtab=shtmtab & " </ul>"
			shtmtab=shtmtab & shtmtabD
			shtmtab=shtmtab &  "</div>"
			shtml=shtml & shtmtab
			'SHTML=SHTML & "222222222222222222222"
			Case Else 'todos
			Do While Not rsc.EOF And ipag < limitpag
                Formularioz=Formularioz & "<div  style='border-radius: 5px;background: #73AD21;'>"
				iddoc = rsc(0)
				If existeDoc(iddoc,rsc("tipo"),tipodocx) Then
					solapa10 = False
					solapa5 = False
					xo = xo + 1
					tips=armatip(rsc(0))
					xacceso = "codigo*" & rsc("codigo") & "_tipo*" & rsc("tipo") & "_perfil*OBJARC"
					xacceso = Replace(xacceso, "_", "&")
					If Portada(rsc("codigo")) = "SI" Then
						xacceso = Replace(xacceso, "*", "=") & "&NIV1=0&NIV2=" & rsc("index1")
					Else
						If  tipocpt = "K" Then
							xacceso = Replace(xacceso, "*", "=") & "&NIV1=15&NIV2=" & rsc("index1")
						Else
							xacceso = Replace(xacceso, "*", "=") & "&NIV1=" & rsc("solapa") & "&NIV2=" & rsc("index1")
						End If
					End If


					sActcon = "Consultar.asp?Ysearch=S&DOCMODAL=CL&ACC=XP&" & xacceso & "&Datos=*smenu=PORTEXTO"
					clickn = Chr(34) & "AbrirCons('" & sActcon & "','" & rsc("tipo") & "');return false;" & Chr(34)

					If session("APLMDCLI")="BGAL" Then
						Formularioz=  Formularioz & NewLine(0) & NewLine(0) & "<a  class=nt onMouseout=""hideddrivetip()"" onMouseover=""ddrivetip('"&armatip(rsc(0))&"',250,'#dddddd')"" onclick="&clickn&" href=""" & url & search_maintain_search_term & """ style=""" & m_style_title & """>" & tipodocx & " : " &  Replace(rsc("title")," ","&nbsp;") & "</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
					Else
						Formularioz=  Formularioz & NewLine(0) & NewLine(0) & "<a  class=nt onMouseout=""hideddrivetip()""  onMouseover=""ddrivetip('"&armatip(rsc(0))&"',250,'#dddddd')"" onclick="&clickn&" href=""" & url & search_maintain_search_term & """ style=""" & m_style_title & """>" &  Replace(rsc("title")," ","&nbsp;") & "</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
					End If

					texto = NewLine(0) & "<br><span style=""" & m_style_text & """>" & rsc(3) & "</span>" & "<br>"
					i = 1
					atit = titsolapa(rsc(0), rsc("tipo"))
					'   response.write atit(1)
					antsol = 0



					If rsc("solapa") = 5 Then
						solapa5 = True
					End If
					If   rsc("solapa") = 10 Then

						solapa10 = True
					End If

					rsc.movenext
					Do While Not rsc.EOF

						If iddoc <> rsc(0) Then
							iddoc = rsc(0)
							Formularioz = Formularioz & NewLine(0) & "<br>" & texto & "<hr></div>"
							totaldoc = totaldoc + 1
							Exit Do
						End If

						If rsc("solapa") = "5" Then

							If Not solapa5 Then
								'if rsc("solapa") <> antsol and not solapa10 then
								xacceso = "codigo*" & rsc("codigo") & "_tipo*" & rsc("tipo") & "_perfil*OBJARC"
								xacceso = Replace(xacceso, "_", "&")
								xacceso = Replace(xacceso, "*", "=") & "&NIV1=" & rsc("solapa") & "&NIV2=" & rsc("index1")
								sActcon = "Consultar.asp?Ysearch=S&DOCMODAL=CL&ACC=XP&" & xacceso & "&Datos=*smenu=PORTEXTO"
								clickn = Chr(34) & "AbrirCons('" & sActcon & "','" & rsc("tipo") & "');return false;" & Chr(34)
								Formularioz = Formularioz & "<a  onMouseout=""hideddrivetip()"" onMouseover=""ddrivetip('" & GETSOLAPA(rsc("solapa"), rsc("index1")) & "',450,'#dddddd')"" class=nt onclick=" & clickn & " href=""" & url & search_maintain_search_term & """ style=""" & m_style_title & """>&nbsp;" & i & "&nbsp;</a></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
								i = i + 1
								' antsol=rsc("solapa")
								solapa5 = True
							End If

						ElseIf rsc("solapa") = 10 Then
							If Not solapa10 Then
								'if rsc("solapa") <> antsol and not solapa10 then
								xacceso = "codigo*" & rsc("codigo") & "_tipo*" & rsc("tipo") & "_perfil*OBJARC"
								xacceso = Replace(xacceso, "_", "&")
								If tipocpt="K" Then
									xacceso = Replace(xacceso, "*", "=") & "&NIV1=15&NIV2=" & rsc("index1")
								Else
									xacceso = Replace(xacceso, "*", "=") & "&NIV1=" & rsc("solapa") & "&NIV2=" & rsc("index1")
								End If
								sActcon = "Consultar.asp?Ysearch=S&DOCMODAL=CL&ACC=XP&" & xacceso & "&Datos=*smenu=PORTEXTO"
								clickn = Chr(34) & "AbrirCons('" & sActcon & "','" & rsc("tipo") & "');return false;" & Chr(34)
								Formularioz = Formularioz & "<a  onMouseout=""hideddrivetip()"" onMouseover=""ddrivetip('" & GETSOLAPA(rsc("solapa"), rsc("index1")) & "',450,'#dddddd')"" class=nt onclick=" & clickn & " href=""" & url & search_maintain_search_term & """ style=""" & m_style_title & """>&nbsp;" & i & "&nbsp;</a></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
								i = i + 1
								' antsol=rsc("solapa")
								solapa10 = True
							End If
						Else
							xacceso = "codigo*" & rsc("codigo") & "_tipo*" & rsc("tipo") & "_perfil*OBJARC"
							xacceso = Replace(xacceso, "_", "&")
							If tipocpt="K" Then
								xacceso = Replace(xacceso, "*", "=") & "&NIV1=15&NIV2=" & rsc("index1")
							Else
								xacceso = Replace(xacceso, "*", "=") & "&NIV1=" & rsc("solapa") & "&NIV2=" & rsc("index1")
							End If
							sActcon = "Consultar.asp?Ysearch=S&DOCMODAL=CL&ACC=XP&" & xacceso & "&Datos=*smenu=PORTEXTO"
							clickn = Chr(34) & "AbrirCons('" & sActcon & "','" & rsc("tipo") & "');return false;" & Chr(34)
							Formularioz = Formularioz & "<a  onMouseout=""hideddrivetip()"" onMouseover=""ddrivetip('" & GETSOLAPA(rsc("solapa"), rsc("index1")) & "',450,'#dddddd')"" class=nt onclick=" & clickn & " href=""" & url & search_maintain_search_term & """ style=""" & m_style_title & """>&nbsp;" & i & "&nbsp;</a></span>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
							i = i + 1
						End If
						'end if

						rsc.movenext
						' 0 i=i+1
					Loop
					If totaldoc = 9 Then
						totaldoc = 1
						If ipag = 1 Then
							'Formularios=Formularios & "<div id=pag"&ipag&" >" & chr(13)  & formularioz & chr(13)  &"</div>" & chr(13)
							response.write "<div sTYLE='float: left; z-index:1;overflow:auto;width:" & anchopag & ";border-style:solid;border-width:2;border-color:#ffffff;' id=pag"&ipag&" >" & Chr(13)  & formularioz & Chr(13)  &"</div>" & Chr(13)
							Formularioz = ""
							response.flush
						Else
							'Formularios=Formularios & "<div id=pag"&ipag&"  style='display=none;'>" & chr(13)  & formularioz & chr(13)  &"</div>" & chr(13)
							response.write  "<div id=pag"&ipag&"  sTYLE='float: left; z-index:1;overflow:auto;width:" & anchopag & ";border-style:solid;border-width:2;border-color:#ffffff;display:none;'>" & Chr(13)  & formularioz & Chr(13)  &"</div>" & Chr(13)
							Formularioz = ""
							response.flush
						End If
						ipag = ipag + 1
						Formularioz = ""
					End If
				Else
					rsc.movenext
				End If
			Loop


			If totaldoc >= 1 Then
				If ipag = 1 And xo > 0 Then
					'     Formularios=Formularios & "<div id=pag"&ipag&" >" & chr(13)  & formularioz & chr(13)  &"</div>" & chr(13)
					response.write  "<div sTYLE='float: left; z-index:1;overflow:auto;width:" & anchopag & ";border-style:solid;border-width:2;border-color:#ffffff;' id=pag"&ipag&" ><br>" & Chr(13)  & formularioz & NewLine(0) &  "<br>" & texto & "<hr></div>" & Chr(13)  &"</div>" & Chr(13)
					response.flush
				Else
					'    Formularios=Formularios & "<div id=pag"&ipag&"  style='display=none;'>" & chr(13)  & formularioz & chr(13)  &"</div>" & chr(13)

					If ipag < limitpag Then
						response.write "<div id=pag"&ipag&"  sTYLE='float: left; z-index:1;overflow:auto;width:" & anchopag & " ;border-style:solid;border-width:2;border-color:#ffffff;display:none;'>" & Chr(13)  & formularioz & NewLine(0) &  "<br>" & texto & "<hr></div>" & Chr(13)  &"</div>" & Chr(13)
						response.flush
					Else
						ipag=ipag-1
					End If
				End If
			End If




			If xo > 0 Then
				If ipag = 1 Then

					xinfo = "<h5>Pagina 1" & " de " & ipag  & " - Total Documentos encontrados : " & xo & "<h5>"
				Else
					xinfo = "<h5>Pagina 1" & " de " & ipag  & " - Total Documentos encontrados : " & xo & "<h5>"
				End If
			Else
				xinfo = "<h5> Total Documentos encontrados : " & xo & "<h5>"
			End If

			shtml = "<script>"

			'  if tipobsq="" then
			shtml = shtml & " var op1=document.getElementById('infoP')" & Chr(13)
			'   else
			shtml = shtml & " var op2=document.getElementById('infoPAV');" & Chr(13)
			'   end if
			shtml = shtml & " op1.innerHTML='"& xinfo&"';"& Chr(13)
			shtml = shtml & " op2.innerHTML='"& xinfo&"';"& Chr(13)
			shtml = shtml & "topepag=" & ipag & ";" & Chr(13)
			shtml = shtml & "cantdoc=" & xo & ";" & Chr(13)
			If xo > 0 Then
				'  if tipobsq="" then
				shtml = shtml & "var ops1=document.getElementById('PAGINAS');ops1.innerHTML='PAGINAS :&nbsp;&nbsp;'; "
				'  ELSE
				shtml = shtml & "var ops2=document.getElementById('PAGINASAV');ops2.innerHTML='PAGINAS :&nbsp;&nbsp;'; "
				'  END IF


				shtml = shtml & " for(var id=1;id <= topepag ;id++)"
				shtml = shtml & "{ "
				shtml = shtml & " if (id==1 && topepag !=1){"
				shtml = shtml & " ops1.innerHTML=ops1.innerHTML + '<a  id=lnk' + id  + ' href=""#Foo""   onclick=""mostrarp(' + id +');"" style=""" & m_style_titlex & """ >&nbsp' + id + '&nbsp</a>&nbsp;&nbsp;&nbsp;';" & Chr(13)
				shtml = shtml & "}else{"
				shtml = shtml & " ops1.innerHTML=ops1.innerHTML + '<a  id=lnk' + id  + ' href=""#Foo""   onclick=""mostrarp(' + id +');"" style=""" & m_style_title & """ >&nbsp' + id + '&nbsp</a>&nbsp;&nbsp;&nbsp;';" & Chr(13)
				shtml = shtml & "}"
				shtml = shtml & " ops2.innerHTML=ops2.innerHTML + '<a  id=lnkav' + id  + ' href=""#Foo""   onclick=""mostrarpav(' + id +');"" style=""" & m_style_title & """ >&nbsp' + id + '&nbsp</a>&nbsp;&nbsp;&nbsp;';" & Chr(13)
				shtml = shtml & "}"
			Else

				'  if tipobsq="" then
				shtml = shtml & "var ops1=document.getElementById('PAGINAS');ops1.innerHTML='&nbsp;&nbsp;'; "
				'  ELSE
				shtml = shtml & "var ops2=document.getElementById('PAGINASAV');ops2.innerHTML='&nbsp;&nbsp;'; "
				'  END IF
			End If

			shtml = shtml & " </script> "
		End Select
	End If

	response.write shtml  &  ofv.botoneraESSX("0","","","","","","","","","no","","","") & "</body></html>"


End If

Function armatip(idC)
	texto = ""

	' on error resume next
	strsql = "select a.temadescri,a.stemadescri,b.*,c.estado AS ESTADOD,b1.codiden,b11.TIPOCPT from CSRELPATHDESCRINV a "
	strsql = strsql & " inner join objetos b1 on(b1.codobjeto= a.cobj)"
	strsql = strsql & " inner join TIPOBJETO b11 on(b11.TIPOBJETO= B1.TIPOBJETO)"
	strsql = strsql & " inner join objprty b on(b.codobjeto= a.cobj)"
	strsql = strsql & " inner join estados c on(c.objid=b.estado) where a.cobj=0" & idC
	'  RESPONSE.WRITE STRSQL
	Set recordsetip = CN.Execute(strsql)
	'response.write recordsetip("autor")
	'xxx=getinfouser(recordsetip("autor"))
	If Not recordsetip.eof Then
		TIPOCPT=recordsetip("TIPOCPT")
		texto = "<B><U>TEMA</U> :</b>&nbsp;&nbsp;" & recordsetip("TEMADESCRI") & "<BR><B><U>SUBTEMA</U> :</b>&nbsp;&nbsp;" & recordsetip("STEMADESCRI")    '& "<BR>"
		'TEXTO = TEXTO & "<B><U>AUTOR</U>:&nbsp;&nbsp;</B>" & Replace(Session("SGU_NOMBRE"), " ", "&nbsp;") '& "&nbsp;<B>SUBTEMA : </b>"& recordsetip("STEMADESCRI")& "<BR>"
		' TEXTO = TEXTO &  "&nbsp;<B>SUBTEMA : </b>"& recordsetip("STEMADESCRI")& "<BR>"
		texto = texto & "<br><B><U>ESTADO</U>&nbsp;&nbsp;:</B>" & recordsetip("ESTADOD") & "<br><U><b>VIGENCIA</U> : </b>" & fechadorcor(CDate(CLng("0" & recordsetip("fecvigen")) + CLng(CDate("01-01-1900"))), "")
		If recordsetip("VERSION") = "O" Then
			zVersion = "(O)"
		Else
			xx= recordsetip("NROVER")

			zVersion = "V (" &  xx & ")"
		End If

		If session("APLMDCLI") <> "BGAL" Then
			texto = texto & "<BR><B><U>COD. IDENT.</U> :&nbsp;&nbsp;</B>" & recordsetip("CODIDEN")
		End If
		texto = texto & "<BR><B><U>VERSIÓN</U>  :&nbsp;&nbsp;</B>" & zVersion & "  "    '& idC
		armatip = "<b>" & texto & "</b>"
	Else
		armatip=""
	End If
	Call OFV.cerrarconsulta(recordsetip)
End Function


Function Portada(idC)
	texto = ""

	If session("APLMDCLI")="BC" Then
		strsql = "select codobjeto from objportada where dettipo='S' and codobjeto=0" & idC

		Set recordseport = CN.Execute(strsql)

		If Not recordseport.EOF Then
			Portada = "SI"
		Else
			Portada = "NO"
		End If
		Call OFV.cerrarconsulta(recordseport)
	Else
		Portada = "NO"
	End If

End Function



Private Function NewLine(spaces)
	NewLine = vbCrLf & Space(spaces)
End Function


Public Function ReplaceRegExp(ByVal strString, strPattern, strReplace)
	Dim RE: Set RE = New RegExp

	With RE
		.Pattern = strPattern
		.Global = True
		.IgnoreCase = True
		ReplaceRegExp = .Replace(strString, strReplace)
	End With
End Function

Function titsolapa(sCodigo, stipo)

	strsql = "Select * From Objperfildoc Where codobjeto =" & sCodigo
	'response.write "sql " & strsql
	Set rsperfil = OFV.crearconsultaEx(strsql, CN, 1, volver)
	If rsperfil.EOF Then
		perfilsp = "N"
		'  response.write "no "

		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'''' VERIFICA EL PERFIL GENERAL DEL OBJETO ''''''
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


		strsql = "Select Objperfil.* From Objperfil, Tipobjeto where "
		strsql = strsql & "Objperfil.tipoperfil = Tipobjeto.perfil "
		strsql = strsql & "and Tipobjeto.tipobjeto = '" & stipo & "'"
		'     response.write "sql 1111" & strsql
		Set rsperfil = OFV.crearconsultaEx(strsql, CN, 1, volver)


	End If

	If Not rsperfil.EOF Then
		stit = "vacio"
		delim = "###"
		For t = 17 To 33
			stit = stit & delim & rsperfil(t)

		Next
		'  response.write stit
		titsolapa = Split(stit, "###")
	End If
	Call OFV.cerrarconsulta(rsperfil)
End Function
Function GETSOLAPA(id, evt)
	ID=CStr(ID&"")

	Select Case id

		Case "1"
		TIT = "SOLAPA " & atit(1)    '"Solapa propiedades"
		Case "2"
		TIT = "SOLAPA " & atit(2)    '"Solapa consideraciones"
		Case "3"
		TIT = "SOLAPA " & atit(3)    '"Solapa sectores"
		Case "4"
		TIT = "SOLAPA " & atit(4)    ' "Solapa eventos"

		Case "5"
		TIT = "SOLAPA " & atit(5)    '"Solapa procedimientos"
		TIT = TIT & "<b>" & bsqevt(evt) & "</b>"
		Case "6"
		TIT = "SOLAPA " & atit(6)    '"Solapa descripcion"
		Case "7"
		TIT = "SOLAPA " & atit(7)    '"Solapa controles"
		Case "8"
		TIT = "SOLAPA " & atit(8)    '"Solapa esquemas"
		Case "9"
		TIT = "SOLAPA " & atit(9)    '"Solapa tablas"
		Case "10"
		TIT = "SOLAPA " & atit(10)    '"Solapa documentos"
		Case "11"
		TIT = "SOLAPA " & atit(11)    '"Solapa contenido"
		Case "12"
		TIT = "SOLAPA " & atit(12)    '"Solapa cursograma"
		Case "13"
		TIT = "SOLAPA " & atit(13)    '"Solapa contratos"
		Case "14"
		TIT = "SOLAPA " & atit(14)    '"Solapa faq"
		Case "15"
		TIT = "SOLAPA " & atit(15)    '"Solapa general"
		Case Else
		TIT = "Sin descripción"




	End Select
	If session("APLMDCLI")="GP" Then TIT=REPLACE(TIT,"SOLAPA","")
	GETSOLAPA = "<B>" & TIT & "</B>"
End Function


Function bsqevt(id)

	sql = "select * from CSCEventoSectorPerfil Where codproced ='" & id & "'"

	Set recordevt = CN.Execute(sql)

	If Not recordevt.EOF Then
		EVSC = recordevt(2)
		EVPS = recordevt(6)
		EVDS = recordevt(4)
		secd = recordevt(3)
		perd = recordevt(5)
		FREC = recordevt(8)
		stexto1 = "<br><br>EVENTO: " & EVSC & "." & EVPS & "&nbsp;-&nbsp;" & EVDS
		stexto1 = stexto1 & "<br>SECTOR: " & secd
		stexto1 = stexto1 & "<br>PERFIL: " & perd
		stexto1 = stexto1 & "<br>FRECUENCIA: " & FREC
		stexto1 = Replace(stexto1, " ", "&nbsp;")
		bsqevt = stexto1

	End If

	Call OFV.cerrarconsulta(recordevt)


End Function


Function EliminarAcentos(texto)

	If Len(texto) <> 0 Then
		texto = reemplaza(texto, "[ÀÁÄáàä]", "A")
		texto = reemplaza(texto, "[ÈÉËèéë]", "E")
		texto = reemplaza(texto, "[ÌÍÏìíï]", "I")
		texto = reemplaza(texto, "[ÒÓÖòóö]", "O")
		texto = reemplaza(texto, "[ÙÚÜùúü]", "U")
	End If
	EliminarAcentos = texto
End Function

Function existeDoc(id,nCodobjeto,ByRef TIPODOC)
	Dim strsql
	Dim recordsetip
	existeDoc=False



	' select a.temadescri,a.stemadescri,b.*,c.estado AS ESTADOD,b1.codiden ,D.SEGOWN
	'from CSRELPATHDESCRIex a
	' inner join objetos b1 on(b1.codobjeto= a.cobj)
	' inner join objprty b on(b.codobjeto= a.cobj)
	' inner join estados c on(c.objid=b.estado)
	' inner join objsdk d on(d.codobjeto=b.codobjeto )
	' left join  OBJSDKDET e on(
	' E.CODOBJETO= D.CODOBJETO
	' )
	' where


	' a.cobj=02202 and b1.objnivel >= 3  And c.objid In (9,1,6,2,5,3)
	' AND
	'((E.grpusr = 'USR-CL') or (E.grpusr = 'GRP-CONS'))


	strsql = "select a.temadescri,a.stemadescri,b.*,c.estado AS ESTADOD,b1.codiden,E.DESCRI AS TIPODOC from CSRELPATHDESCRINV a "
	strsql = strsql & " inner join objetos b1 on(b1.codobjeto= a.cobj)"
	strsql = strsql & " inner join objprty b on(b.codobjeto= a.cobj)"
	strsql = strsql & " inner join tipobjeto e on(e.tipobjeto= b.tipobjeto)"
	strsql = strsql & " inner join estados c on(c.objid=b.estado) where a.cobj=0" & id
	If  session("APLMDCLI") = "BGAL" Then
		strsql=strsql & validadoc()
	End If
	'  RESPONSE.WRITE STRSQL
	Set recordsetip = CN.Execute(strsql)
	If Not recordsetip.eof Then
		TIPODOC=recordsetip("TIPODOC")
		If  session("APLMDCLI") <> "BGAL" Then
			TIPODOC=recordsetip("TIPODOC")
			If valpermisosdoc(nCodobjeto,id) Then
				existeDoc=True
			Else
				existeDoc=False
			End If
		Else
			existeDoc=True
		End If
	Else
		existeDoc=False

	End If

	Call ofv.cerrarconsulta(recordsetip)
End Function


Function valpermisosdoc(sTipobjeto,nCodobjeto)

	valpermisosdoc=False
	Strsql = "Select * From OBJSDK Where tipobjeto='" & sTipobjeto & "' and codobjeto = "  & nCodobjeto
	'response.write strsql
	Set rsval1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

	If Not rsval1.eof Then

		Segct = "S"
		Segmo = "S"
		Segco = "S"
		Segce = "S"

		If rsval1(3) Then Segct = "N"
		If rsval1(4) Then Segmo = "N"
		If rsval1(5) Then Segco = "N"
		If rsval1(6) Then Segce = "N"
		ownerDoc=rsval1(7)

		Call ofv.cerrarconsulta(rsval1)

		If Session("Usuario") <> ownerDoc And Segco = "N" Then


			Strsql = "Select * From OBJSDKDET Where tipobjeto='" & sTipobjeto
			Strsql = Strsql & "' and codobjeto = "  & nCodobjeto
			Strsql = Strsql & " and ("
			Strsql = Strsql & "(grpusr = 'USR-" & Session("Usuario") & "') "

			'''''''''''''''''''''''''''''''''''''''''''
			If  Session("DELLOG") = "S" Then

				XMFTGRP = Split(session("SegGrpLst"),Chr(9))
				xcount  = UBound(XMFTGRP) + 1
				XMFSEP  = ""
				StrSql = StrSql & " or ( grpusr in ("

				For XMFI = 0 To UBound(XMFTGRP)
					StrSql = StrSql  & XMFSEP & "'GRP-" & XMFTGRP(XMFI) & "'"
					XMFSEP  = ","
				Next

				StrSql = StrSql & ")) "

			Else
				Strsql = Strsql & " or (grpusr = 'GRP-" & Session("Seggrp") & "')"
			End If


			''''''''''''''''''''''''''''''''''''''''''


			Strsql = Strsql & ") "

			' response.write strsql
			Set rsval = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If  Not rsval.eof Then ' ok si esta en la lista  de consulta del doc
				valpermisosdoc=True
			Else

				valpermisosdoc=False


			End If

			Call ofv.cerrarconsulta(rsval)

		Else ' si es el owner del doc

			valpermisosdoc=True

		End If

	End If


End Function

Function validadoc()


	sFormid = "CLIMS157172171124281051096140"

	ReDim Preserve bconsultar(4)
	OFV.ObtenerAtributosNodo sFormid, "", "ESTADO", "4"
	bconsultar(4) = OFV.AConsultar
	OFV.ObtenerAtributosNodo sFormid, "", "ESTADO", "2"
	bconsultar(2) = OFV.AConsultar
	OFV.ObtenerAtributosNodo sFormid, "", "ESTADO", "1"
	bconsultar(1) = OFV.AConsultar
	OFV.ObtenerAtributosNodo sFormid, "", "ESTADO", "3"
	bconsultar(3) = OFV.AConsultar
	sverestado = " and b1.objnivel >= " & Session("SegNivel") & " And c.objid In (9,"


	For i = 1 To 4
		If bconsultar(i) Then
			sverestado = sverestado & i & ","
			If i = 1 Then sverestado = sverestado & "6" & ","
			If i = 2 Then sverestado = sverestado & "5" & ","

		End If
	Next
	sverestado = Left(sverestado, Len(sverestado) - 1) & ") "


	validadoc = sverestado
End Function


Function reemplaza(texto, patrn, replStr)
	Dim regEx
	Set regEx = New RegExp            ' Crea una expresión regular.
	regEx.Pattern = patrn            ' Establece el modelo.
	regEx.IgnoreCase = True            ' Hace distinción entre mayúsculas y minúsculas.
	regEx.Global = True
	reemplaza = regEx.Replace(texto, replStr)   ' Realiza el reemplazo.
	Set regEx = Nothing
End Function

Function selecEmpre()
	Dim TABLAX,Tabla
	sImagenx = "<B><FONT FACE=CorFlatIcon COLOR=#aa0000 SIZE=4>x</FONT></B>"

	If  Session("EMPMUN") = "S" Then

		Fila    = ""

		SINPUT3 = "&nbsp;<A class=tt style='color:#000000;' title='SELECCIONE EMPRESA' HREF='javascript:void(0);' onclick='shwDem();' >EMPRESA:&nbsp;&nbsp;</a>"

		Fila    = ""
		Celda   = ""
		sHref   = ""

		sImage  = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(0,5),"0","")
		sLink   = ofv.GenLink("","ttt",sHref,"ftemas","","",sImage,"si")
		Celda   = Celda & ofv.GenCelda("","tt ","","8%","","","",SINPUT3,"si")
		Celda   = Celda & ofv.GenCelda("","ttt ","center","5%","","","",sLink,"si")
		Celda   = Celda & ofv.GenCelda("","ttt  ","","","","","","&nbsp;<BIG>" & ATREMP(0,4) & "</BIG>","si")

		Fila    = fila & ofv.GenRow("","","top","","","","",Celda,"si")

		Tabla   = ofv.GenTabla("","aaa style='"  & Session("CBRDABL") & "' ","","100% ","","0","","-1","-2","","",Fila,"si")





		shtmlx=  "<script language=javascript>" & Chr(13)
		shtmlx=shtmlx & "function shwDem(){" & Chr(13)
		shtmlx=shtmlx & "document.all.DEM1.style.display='block';" & Chr(13)
		shtmlx=shtmlx & "}" & Chr(13)
		shtmlx=shtmlx & "</script>" & Chr(13)
		'----------------------------------------------
		sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & Chr(34)
		sinput2 = sinput2  & "underline" & Chr(34) & ";' onmouseout='this.style.textDecoration=" & Chr(34)
		sinput2 = sinput2  & "none" & Chr(34) & ";' href='javascript:void(0);' ONCLICK=" & Chr(34) & "document.all.DEM1.style.display='none';" & Chr(34)
		sinput2 = sinput2  & " TITLE='OCULTAR' ><b>" & sImagenx & "</b></a>"
		TABLA = Tabla & "<DIV id='DEM1' STYLE='display:none;z-index:1;margin-top:1;margin-left:1;background-color:#ffffff;position:absolute;left:1px;top:240px;left:20px;height:250px;width:350px;border-style:inset;border-color:#f7aa22;border-width:2;overflow:auto;'>"


		TABLAX = TablaX & OFV.gentabla("", "aaa style='filter:shadow(color=#dddddd,direction=125);' ", "", "100%", "", "0", "", "-2", "-1", "", "", "", "no")
		TABLAX = TablaX & "<tr><td  "
		TABLAX = TablaX & "class=vT align=left colspan=3 STYLE='border-style:outset;border-width:2;background-color:#eeeeee;' >" & sinput2 & "&nbsp;SELECCIONE EMPRESA</td></tr>"
		TABLAX = TablaX & "<tr><td colspan=3 "
		TABLAX = TablaX & " >&nbsp;</td></tr>"

		For IAEMP=1 To TAEMP

			'''  RESPONSE.WRITE  ATREMP(IAEMP,2)

			If ATREMP(IAEMP,2) = 1 Then

				If ATREMP(IAEMP,0) = ATREMP(0,0) Then
					TABLAX = TablaX & "<tr BGCOLOR=#EEEEEE >"
				Else
					TABLAX = TablaX & "<tr>"
				End If

				TABLAX = TablaX & "<td width=5 "
				If ATREMP(IAEMP,0) = ATREMP(0,0) Then
					TABLAX = TablaX & "class=vt  >" & sImagen4si & "</td>"
				Else
					TABLAX = TablaX & "class=vt  >" & sImagenI3 & "</td>"
				End If

				sImage = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(IAEMP,5),"0","")
				TABLAX = TablaX & "<td width=5 "
				TABLAX = TablaX & "class=vt  >" & sImage & "</td>"

				TABLAX = TablaX & "<td class=vt  >"

				'sHref = "subconsulta.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0&EMPRE=" & ATREMP(IAEMP,0)
				sHref = "searchex.asp?EMPRE=" & ATREMP(IAEMP,0)
				'''  sHref = "OBJtreeTema.asp?Tipo=" & sTipo & "&Nombre=" & LNombre & "&Menu=" & SESSION("MENU") & "&EMPRE=" & ATREMP(IAEMP,0)
				sLink = ofv.GenLink("","cc",sHref,"","","","<b>" & ATREMP(IAEMP,4) & "</b>","si")
				TABLAX = TablaX & sLink
				TABLAX = TablaX & "</td>"
				TABLAX = TablaX & "</tr>"
			End If

		Next

		TABLAX = TablaX & OFV.genFtabla()

		TABLA = Tabla & TABLAX

		TABLA = Tabla & "</DIV>"

		selecEmpre=tabla & shtmlx

	End If

End Function
%>
