

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
	datofiltro=request.form("filtros")& ""
'		response.write datofiltro

	if datofiltro<>"" then

    arrf=split(datofiltro,"####")
    TIPOF=" AND ("
    autorF=" AND ("
    TEMAF=" AND ("
    SECTORF=" AND ("
    DELIM=""
    DELIMA=""
    DELIMT=""
    DELIMS=""
    for i= 0 to ubound(arrf)

     ColF=split(arrf(i)  ,":")

     select case UCASE(ColF(0)&"")
     case "TIPO"
         TIPOF=TIPOF & DELIM &  CHR(34)   & "[[TIPO_" & ColF(1) &"]]" & CHR(34)
         DELIM=" OR "

      case "AUTOR"
	    autorF=autorF & DELIMA &  CHR(34)   & "[[AUTOR_" & ColF(1) &"]]" & CHR(34)
        DELIMa=" OR "

      case "TEMA"
	  	TEMAF=TEMAF & DELIMT &  CHR(34)   & "[[TEMA_" & ColF(1) &"]]" & CHR(34)
        DELIMT=" OR "


      case "SECTOR"
	  	SECTORF=SECTORF & DELIMS &  CHR(34)   & "[[CODSECTOR_" & ColF(1) &"]]" & CHR(34)
	    DELIMS=" AND "

     end select

    next
    TIPOF=TIPOF & ")"
    autorF=autorF & ")"
    TEMAF=TEMAF & ")"
    SECTORF=SECTORF & ")"
    if delim="" then tipof=""
    if delimA="" then autorF=""
    if DELIMT="" then TEMAF=""
    if DELIMS="" then SECTORF=""

    FILTROSF=TIPOF &  autorF & TEMAF & SECTORF
	end if
	XXXX = NEWEMPRE("NO")
   RESPONSE.WRITE FILTROSF

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
	xs =  "function AbrirCons(Datos,nom) {" & Chr(13)
	xs = xs & " var h=570;w=790 ;"
	If SESSION("APLMDCLI") = "BC" Then
		xs = xs & "if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;} "
	End If
	xs = xs & " myWin= open(Datos"

	xs = xs & "," & Chr(34) & "ConsLnk" & Chr(34) & " + nom,"    '& Chr(34)
	' xs = xs & "'menubar=yes,toolbar=yes,height=h,width=w,resizable=yes'" '& Chr(34)
	xs = xs & "  'toolbar=no,width='+w+',height='+h+',top=0,left=0,resizable=yes'"    '& CHR(34)
	xs = xs & " ); " & Chr(13) & " } " & Chr(13)
	'shtml = shtml	& xs
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

	shtml = shtml & "document.getElementById('buscar').value='" & search_phrase &"';</script>"


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
	If session("menuex") <> "S" Then
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


	findd.canttx = 0  ' -1 no trae texto, 0 trae el toda la columna texto,cualquier otro valor > 0 trae esa cantidad
	findd.cantreg = 150
	findd.listacar="{}[]!#$%&/()=?°'ø|*+®¥:.;,<>-" ' sirve para fulltext con palabras con caracteres especiales ejemplo 123456789-0 guion es un caracter especial.le sacamos el * para buscar exacto


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
		FINDD.FILTROS= FINDD.FILTROS & " AND F.P1='D' "
		'NUEVO FILTROS
		IF TIPOF <> "" THEN FINDD.FILTROS=FINDD.FILTROS '& " AND F.TIPOBJETO " & TIPOF
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
	Else

		VALOR1=Chr(34) & "[[OBJUN_"  & sun & "]]" & Chr(34)

		VALOR2=" OR "  &  CHR(34)   & "[[CODEMPRE_" & sUN &"]]" & CHR(34)
        VALOR3=" AND "  &  CHR(34)   & "[[AUTOR_ADM]]" & CHR(34)
		findd.ADDINNER=" inner join (select distinct codobjeto,tipobjeto "
		findd.ADDINNER =findd.ADDINNER & "from containstable(ObjMetadata, metadata, '" & VALOR1	 & FILTROSF	 & "') as ct "

		findd.ADDINNER =findd.ADDINNER & "join ObjMetadata as y on ct.[key]=y.[objid] ) z "
		findd.ADDINNER =findd.ADDINNER & " ON Z.CODOBJETO=F.CODOBJETO and Z.TIPOBJETO=F.TIPOBJETO"

		findd.ADDINNER= " inner join (select  codobjeto,tipobjeto,metadata from containstable(ObjMetadata, metadata, '"& VALOR1  & FILTROSF  & "') as ct2 "
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
		'			RESPONSE.WRITE findd.sqlj
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
		Select  Case  session("menuex")
		Case "S" ' nueva visualizacion
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


					'shtml = OFV.FormHeader(Session("FormN"))

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

					'sHTML = sHTML & "<br><form method=post action='" & sHref2 & "'><input type=text name=buscar id=buscar value='" & search_phrase & "' class=dt xclass='ui-widget ui-state-default ui-corner-all'></input><input type=submit class=bietx xclass='ui-widget ui-state-default ui-corner-all' value=Buscar></input></form>"
					'shtml = shtml & CorPnlFnc("N", "")
					celda=""
					fila=""
					idf=0
					shtmldocs=""
					shtmtabD=""
					lipag=""
					tpag=1
					xentros=False
					sdataDoc=""

					shtmldocsF= "<div class='em-app' id='app'>"
					shtmldocsF=shtmldocsF & "<aside class='em-side' id='sidebar' aria-label='Filtros'>"
					shtmldocsF=shtmldocsF & "      <div class='em-side-head'>"
					shtmldocsF=shtmldocsF & "        <button class='em-toggle' id='btnToggle' title='Ocultar/Mostrar'><i class='la la-angle-left' aria-hidden='true'></i></button>"
					shtmldocsF=shtmldocsF & "      </div>"

					shtmldocsF=shtmldocsF & "      <div id='sidebarResults' aria-live='polite'>"
					shtmldocsF=shtmldocsF & "        <h3 class='em-title' id='resultsTitle'>Resultados</h3>"
					shtmldocsF=shtmldocsF & "        <div class='em-count'><span id='resultsCount'>0</span> resultados</div>"
      				shtmldocsF=shtmldocsF & "</div>"
                    shtmldocsF=  shtmldocsF & armafiltros()




                   ' shtmldocsF=  shtmldocsF &  "<main>"
					'shtmldocsF=  shtmldocsF &  "      <h3 class='em-section-title'>Filtros seleccionados</h3><form id=filtrosbsq>"
					'shtmldocsF=  shtmldocsF &  "     <div id='chipsPanel' class='em-chips' aria-live='polite'>"
					'shtmldocsF=  shtmldocsF &  "       <span class='em-chips-label'>Filtros:</span>"
					'shtmldocsF=  shtmldocsF &  "       <button id='btnClearAll' class='em-clear-all' type='button'>Quitar todos</button>"
      				'shtmldocsF=  shtmldocsF &  "</div></form>"


                    shtmldocs=" <div class='swiper mySwiperInd'>"
					shtmldocs= shtmldocs & " <div class='swiper-wrapper'>"
					shtmldocs=""
					cant=1
					sul=""
				    sli=""
					Do While RETORNAVALOR(RSC,0) <> "-1"
						Dim 	Asolpas
						Dim 	Aadjuntos
						ReDim 	Asolpas(100)

						ix=0

						iddoc=rsc("codigo") & ""
						atit = titsolapa(rsc("codigo"), rsc("tipo"))

						strsql = "select a.temadescri,a.stemadescri,b.*,c.estado AS ESTADOD,b1.codiden,b11.TIPOCPT from CSRELPATHDESCRINV a "
						strsql = strsql & " inner join objetos b1 on(b1.codobjeto= a.cobj)"
						strsql = strsql & " inner join TIPOBJETO b11 on(b11.TIPOBJETO= B1.TIPOBJETO)"
						strsql = strsql & " inner join objprty b on(b.codobjeto= a.cobj)"
						strsql = strsql & " inner join estados c on(c.objid=b.estado) where a.cobj=0" & iddoc
						'	response.write strsql & "<br>"
						Set recordsetip = CN.Execute(strsql)
						tarjeta=""
						Sesquema=""
						If Not recordsetip.eof Then
							Sesquema=EsqnomPx(rsc("tipo"), rsc("codigo"), ofv)

							xacceso = "codigo*" & rsc("codigo") & "_tipo*" & rsc("tipo") & "_perfil*OBJARC"
							xacceso = Replace(xacceso, "_", "&")
							If tipocpt="K" Then
								xacceso = Replace(xacceso, "*", "=") & "&NIV1=15&NIV2=" & rsc("index1")
							Else
								xacceso = Replace(xacceso, "*", "=") & "&NIV1=" & rsc("solapa") & "&NIV2=" & rsc("index1")
							End If
							sActcon = "Consultar.asp?Ysearch=S&DOCMODAL=CL&ACC=XP&" & xacceso & "&Datos=*smenu=PORTEXTO"
							clickn = Chr(34) & "AbrirCons('" & sActcon & "','" & rsc("tipo") & "');return false;" & Chr(34)
							tarjeta = tarjeta & " <div class='result'>"
							tarjeta = tarjeta & "    <img class='thumbnail' src='https://cdn-icons-png.flaticon.com/512/3095/3095583.png' alt='doc1'>"
							tarjeta = tarjeta & "    <div class='info'>"
							tarjeta = tarjeta & "      <h3><A CLASS=nt onclick="& clickn & "><div class=titbsq>" & Sesquema & "</div></a></h3>"
							tarjeta = tarjeta & "  <div class='solapas-row'>"
							tarjeta = tarjeta & "        <div class='solapas'>"

							If recordsetip("VERSION") = "O" Then
								zVersion = "(O)"
							Else
								xx= recordsetip("NROVER")
								zVersion = "V (" &  xx & ")"
							End If

							'tooltip doc
							textot =   " 	<div class='tooltip-icon'><span class='dots'><i class='la la-ellipsis-v'></i></span></div>"
							textot = textot & "    <div class='tooltip-content'>"
							textot = textot & "      <div>"
							textot = textot & "        <p><strong>Tema:</strong> " & recordsetip("TEMADESCRI") & "</p>"
							textot = textot & "        <p><strong>SubTema:</strong>"& recordsetip("STEMADESCRI") &"</p>"
							textot = textot & "        <p><strong>Estado:</strong> " & recordsetip("ESTADOD") & "</p>"
							textot = textot & "        <p><strong>VersiÛn:</strong>" & zversion &"</p>"
							textot = textot & "     </div>"
							textot = textot & "      <div>"
							textot = textot & "        <p><strong>CÛdigo:</strong>"& recordsetip("CODIDEN")& "</p>"
							textot = textot & "        <p><strong>Vigencia:</strong> " & fechadorcor(CDate(CLng("0" & recordsetip("fecvigen")) + CLng(CDate("01-01-1900"))), "")& "</p>"
							textot = textot & "      </div>"
							textot = textot & "	</div>"

						End If
						Call ofv.cerrarconsulta(recordsetip)
						Do While  iddoc = RETORNAVALOR(RSC,0)
							entro=False
							Ssolapa=GETSOLAPA(rsc("solapa"), rsc("index1"))
							textodoc=left(rsc("text"),350)

							xacceso = "codigo*" & rsc("codigo") & "_tipo*" & rsc("tipo") & "_perfil*OBJARC"
							xacceso = Replace(xacceso, "_", "&")
							If tipocpt="K" Then
								xacceso = Replace(xacceso, "*", "=") & "&NIV1=15&NIV2=" & rsc("index1")
							Else
								xacceso = Replace(xacceso, "*", "=") & "&NIV1=" & rsc("solapa") & "&NIV2=" & rsc("index1")
							End If
							sActcon = "Consultar.asp?Ysearch=S&DOCMODAL=CL&ACC=XP&" & xacceso & "&Datos=*smenu=PORTEXTO"
							clickn = Chr(34) & "AbrirCons('" & sActcon & "','" & rsc("tipo") & "');return false;" & Chr(34)
							'solapa documentos

							If rsc("solapa")=10 And Not entro Then
								strsql="SELECT archivo, directorio,descri  FROM Objhtm where codobjeto=0" & rsc("codigo") & " and tipobjeto='" & rsc("tipo") & "' "
								' strsql=strsql & " and descri='" & rsc("title") & "'"
								ixarch=0
								ReDim 	Aadjuntos(100,2)
								Set recordsetadj = CN.Execute(strsql)
								Do While Not recordsetadj.eof
									sdirec = recordsetadj(1) & ""
									If sdirec = "" Then sdirec = "Archivos"
									sarch = sdirec & "/" & recordsetadj(0)

									Aadjuntos(ixarch,0)=replace(sarch &"" ,"//","/")
									Aadjuntos(ixarch,1)=recordsetadj("descri") 'rsc("title")
									' Aadjuntos(ixarch,2)=rsc("title")
									ixarch=ixarch+1
									recordsetadj.movenext
								Loop
								entro=True
							End If
							'armo solapas doc se puede repetir
							If Not InArray(Asolpas,Ssolapa) Then
								'
								tarjeta = tarjeta & "          <a onclick="& clickn & " href='#'>" & Ssolapa & "</a>"
								Asolpas(ix)=Ssolapa
								ix=ix+1

							End If

							rsc.MOVENEXT
						Loop
						tarjeta = tarjeta & "        </div>"
						tarjeta = tarjeta & "        <button class='ver-mas-btn' onclick='toggleDescription(this)'>Ver m·s</button>"
						tarjeta = tarjeta & " </div>"
						tarjeta = tarjeta & "      <div class='description'>"
						tarjeta = tarjeta & left(textodoc,350)
						tarjeta = tarjeta & "      </div>"
						tarjeta = tarjeta & "    </div>"
						textoADJ= "<div class='attachments-columns' id='adjuntos'" & idc &">"

						If entro Then
							For i=0 To ixarch-1

								textoADJ = textoADJ & "      <a target=_new href='" &  Aadjuntos(i,0) & "' title='" & Aadjuntos(i,1) &"' ><i class='la la-file-alt'></i>" & Aadjuntos(i,1) & "</a>"
							Next
							If ixarch > 6 Then
								textoADJ = textoADJ & "      <button class='ver-todos-btn' onclick='mostrarTodos('adjuntos" &idc &"')'>Ver todos</button>"

							End If
						End If
						textoADJ = textoADJ & "</div>"

						tarjeta = tarjeta  & textoADJ &  textot & "</div>"
                        sli= sli & "<li xclass=documento-link><i class=Ciclass></i>"  & tarjeta & "</li>"
						 if cant mod 3 = 0 and false then
							 sul=sul & "<div class='swiper-slide'>"& "<ul>" & sli & "</ul>" & "</div>"
						 	sli=""
						 	cant=1
						else
						 	cant=cant +1
						end if
                    shtmldocs=shtmldocs & tarjeta
					Loop
                         sul="<u>" &  sli & "</u>"
					'    if sli <> "" then  sul=sul & "<div class='swiper-slide'>"& "<ul>" & sli & "</ul>" & "</div>"

						'shtmldocs= shtmldocs & sul & "</div>"
						'shtmldocs= shtmldocs & "    <div class='swiper-pagination'></div>"
						'shtmldocs= shtmldocs & " </div>"
						sul=sul & tarjeta

          '  <!-- ?? NUEVO: DIV contenedor de Resultados -->
              shtmldocsF=  shtmldocsF &   " <main>"
			  shtmldocsF=  shtmldocsF &   "	      <h3 class='em-section-title'>Filtros seleccionados</h3>"
			  shtmldocsF=  shtmldocsF &   "      <div id='chipsPanel' class='em-chips' aria-live='polite'>"
			  shtmldocsF=  shtmldocsF &   "	        <span class='em-chips-label'>Filtros:</span>"
			  shtmldocsF=  shtmldocsF &   "        <button id='btnClearAll' class='em-clear-all' type='button'>Quitar todos</button>"
			  shtmldocsF=  shtmldocsF &   " <form target=info id=sx action='searchex.asp?ing=&FORMULARIOS=D&YIDER="&search_phrase&"' method=post><button id='ver'>ver</button><input type=hidden name=filtros id=filtros value=''><form>"
			  shtmldocsF=  shtmldocsF &   "	      </div>"

			 shtmldocsF=  shtmldocsF &   "	      <div id='emResults' class='em-results' aria-live='polite'>"
			shtmldocsF=  shtmldocsF &   "	        <div class='em-results-head'>"
			shtmldocsF=  shtmldocsF &   "	          <h4 class='em-results-title'>" & shtmldocs & "</h4>"
			shtmldocsF=  shtmldocsF &   "	        </div>"
			shtmldocsF=  shtmldocsF &   "	        <div class='em-results-empty' id='emResultsEmpty los filtros para ver resultados.</div>"
			shtmldocsF=  shtmldocsF &   "	        <div class='em-results-grid' id='emResultsGrid' hidden></div>"
			shtmldocsF=  shtmldocsF &   "	      </div>"
			shtmldocsF=  shtmldocsF &   "	    </main>"
  			shtmldocsF=  shtmldocsF &   "	</div>"

			'    shtmldocsF=  shtmldocsF &    "<section class='em-results'>"
			'	shtmldocsF=  shtmldocsF &     " <div id='chipsPanel' class='em-chips' aria-live='polite'>"
			'	shtmldocsF=  shtmldocsF &        "<span class='em-chips-label'>Filtros:</span>"
			'	shtmldocsF=  shtmldocsF &        "<button id='btnClearAll' class='em-clear-all' type='button'>Quitar todos</button>"
			'	shtmldocsF=  shtmldocsF &      "</div>"

			'	shtmldocsF=  shtmldocsF &     " <div class='em-list' id='resultsList'>"
			'	shtmldocsF=  shtmldocsF &     "        <div class='em-item'>" & shtmldocs & "</div>"
			'	shtmldocsF=  shtmldocsF &     " </div>"
             ''       shtmldocsF=  shtmldocsF &   " </section>"
 			'	 shtmldocsF=shtmldocsf & "  </div>"
				shtml=shtml & shtmldocsF  ' &  sdataDoc

			Case "S00" ' nueva visualizacion
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


			'shtml = OFV.FormHeader(Session("FormN"))

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

			'sHTML = sHTML & "<br><form method=post action='" & sHref2 & "'><input type=text name=buscar id=buscar value='" & search_phrase & "' class=dt xclass='ui-widget ui-state-default ui-corner-all'></input><input type=submit class=bietx xclass='ui-widget ui-state-default ui-corner-all' value=Buscar></input></form>"
			'shtml = shtml & CorPnlFnc("N", "")
			celda=""
			fila=""
			idf=0
			shtmldocs=""
			shtmtabD=""
			lipag=""
			tpag=1
			xentros=False
			sdataDoc=""

			Do While RETORNAVALOR(RSC,0) <> "-1"
			'	Dim 	Asolpas
			'	Dim 	Aadjuntos
			'	ReDim 	Asolpas(100)

				ix=0

				iddoc=rsc("codigo") & ""
				atit = titsolapa(rsc("codigo"), rsc("tipo"))

				strsql = "select a.temadescri,a.stemadescri,b.*,c.estado AS ESTADOD,b1.codiden,b11.TIPOCPT from CSRELPATHDESCRINV a "
				strsql = strsql & " inner join objetos b1 on(b1.codobjeto= a.cobj)"
				strsql = strsql & " inner join TIPOBJETO b11 on(b11.TIPOBJETO= B1.TIPOBJETO)"
				strsql = strsql & " inner join objprty b on(b.codobjeto= a.cobj)"
				strsql = strsql & " inner join estados c on(c.objid=b.estado) where a.cobj=0" & iddoc
				'	response.write strsql & "<br>"
				Set recordsetip = CN.Execute(strsql)
				tarjeta=""
				Sesquema=""
				If Not recordsetip.eof Then
					Sesquema=EsqnomPx(rsc("tipo"), rsc("codigo"), ofv)

					xacceso = "codigo*" & rsc("codigo") & "_tipo*" & rsc("tipo") & "_perfil*OBJARC"
					xacceso = Replace(xacceso, "_", "&")
					If tipocpt="K" Then
						xacceso = Replace(xacceso, "*", "=") & "&NIV1=15&NIV2=" & rsc("index1")
					Else
						xacceso = Replace(xacceso, "*", "=") & "&NIV1=" & rsc("solapa") & "&NIV2=" & rsc("index1")
					End If
					sActcon = "Consultar.asp?Ysearch=S&DOCMODAL=CL&ACC=XP&" & xacceso & "&Datos=*smenu=PORTEXTO"
					clickn = Chr(34) & "AbrirCons('" & sActcon & "','" & rsc("tipo") & "');return false;" & Chr(34)
					tarjeta = tarjeta & " <div class='result'>"
					tarjeta = tarjeta & "    <img class='thumbnail' src='https://cdn-icons-png.flaticon.com/512/3095/3095583.png' alt='doc1'>"
					tarjeta = tarjeta & "    <div class='info'>"
					tarjeta = tarjeta & "      <h3><A CLASS=nt onclick="& clickn & "><div class=titbsq>" & Sesquema & "</div></a></h3>"
					tarjeta = tarjeta & "  <div class='solapas-row'>"
					tarjeta = tarjeta & "        <div class='solapas'>"

					If recordsetip("VERSION") = "O" Then
						zVersion = "(O)"
					Else
						xx= recordsetip("NROVER")
						zVersion = "V (" &  xx & ")"
					End If

					'tooltip doc
					textot =   " 	<div class='tooltip-icon'><span class='dots'><i class='la la-ellipsis-v'></i></span></div>"
					textot = textot & "    <div class='tooltip-content'>"
					textot = textot & "      <div>"
					textot = textot & "        <p><strong>Tema:</strong> " & recordsetip("TEMADESCRI") & "</p>"
					textot = textot & "        <p><strong>SubTema:</strong>"& recordsetip("STEMADESCRI") &"</p>"
					textot = textot & "        <p><strong>Estado:</strong> " & recordsetip("ESTADOD") & "</p>"
					textot = textot & "        <p><strong>VersiÛn:</strong>" & zversion &"</p>"
					textot = textot & "     </div>"
					textot = textot & "      <div>"
					textot = textot & "        <p><strong>CÛdigo:</strong>"& recordsetip("CODIDEN")& "</p>"
					textot = textot & "        <p><strong>Vigencia:</strong> " & fechadorcor(CDate(CLng("0" & recordsetip("fecvigen")) + CLng(CDate("01-01-1900"))), "")& "</p>"
					textot = textot & "      </div>"
					textot = textot & "	</div>"

				End If
				Call ofv.cerrarconsulta(recordsetip)
				Do While  iddoc = RETORNAVALOR(RSC,0)
					entro=False
					Ssolapa=GETSOLAPA(rsc("solapa"), rsc("index1"))
					textodoc=left(rsc("text"),350)

					xacceso = "codigo*" & rsc("codigo") & "_tipo*" & rsc("tipo") & "_perfil*OBJARC"
					xacceso = Replace(xacceso, "_", "&")
					If tipocpt="K" Then
						xacceso = Replace(xacceso, "*", "=") & "&NIV1=15&NIV2=" & rsc("index1")
					Else
						xacceso = Replace(xacceso, "*", "=") & "&NIV1=" & rsc("solapa") & "&NIV2=" & rsc("index1")
					End If
					sActcon = "Consultar.asp?Ysearch=S&DOCMODAL=CL&ACC=XP&" & xacceso & "&Datos=*smenu=PORTEXTO"
					clickn = Chr(34) & "AbrirCons('" & sActcon & "','" & rsc("tipo") & "');return false;" & Chr(34)
					'solapa documentos

					If rsc("solapa")=10 And Not entro Then
						strsql="SELECT archivo, directorio,descri  FROM Objhtm where codobjeto=0" & rsc("codigo") & " and tipobjeto='" & rsc("tipo") & "' "
						' strsql=strsql & " and descri='" & rsc("title") & "'"
						ixarch=0
						ReDim 	Aadjuntos(100,2)
						Set recordsetadj = CN.Execute(strsql)
						Do While Not recordsetadj.eof
							sdirec = recordsetadj(1) & ""
							If sdirec = "" Then sdirec = "Archivos"
							sarch = sdirec & "/" & recordsetadj(0)

							Aadjuntos(ixarch,0)=replace(sarch &"" ,"//","/")
							Aadjuntos(ixarch,1)=recordsetadj("descri") 'rsc("title")
							' Aadjuntos(ixarch,2)=rsc("title")
							ixarch=ixarch+1
							recordsetadj.movenext
						Loop
						entro=True
					End If
					'armo solapas doc se puede repetir
					If Not InArray(Asolpas,Ssolapa) Then
						'
						tarjeta = tarjeta & "          <a onclick="& clickn & " href='#'>" & Ssolapa & "</a>"
						Asolpas(ix)=Ssolapa
						ix=ix+1

					End If

					rsc.MOVENEXT
				Loop
				tarjeta = tarjeta & "        </div>"
				tarjeta = tarjeta & "        <button class='ver-mas-btn' onclick='toggleDescription(this)'>Ver m·s</button>"
				tarjeta = tarjeta & " </div>"
				tarjeta = tarjeta & "      <div class='description'>"
				tarjeta = tarjeta & left(textodoc,350)
				tarjeta = tarjeta & "      </div>"
				tarjeta = tarjeta & "    </div>"
				textoADJ= "<div class='attachments-columns' id='adjuntos'" & idc &">"

				If entro Then
					For i=0 To ixarch-1

						textoADJ = textoADJ & "      <a target=_new href='" &  Aadjuntos(i,0) & "' title='" & Aadjuntos(i,1) &"' ><i class='la la-file-alt'></i>" & Aadjuntos(i,1) & "</a>"
					Next
					If ixarch > 6 Then
						textoADJ = textoADJ & "      <button class='ver-todos-btn' onclick='mostrarTodos('adjuntos" &idc &"')'>Ver todos</button>"

					End If
				End If
				textoADJ = textoADJ & "</div>"

				tarjeta = tarjeta  & textoADJ &  textot & "</div>"
				sdataDoc=sdataDoc & tarjeta
				idf=idf +1
				If idf=10 Or iddoc =-1  Then
					lipag=lipag & "<li class=lipag ><a href='#tabs-"& tpag &"'>Pagina "& tpag &"</a></li>"


					'shtmtabD = shtmtabD & " <div id='tabs-"& tpag &"'><div xclass='accordion'>"
					'shtmtabD=shtmtabD & "<table  width=100% ><tr><td width=100% >" & sdataDoc & "</td><td width=0% class=tdview id='dataview" & tpag & "' ></td></tr></table>"
					'shtmtabD=shtmtabD &  "</div></div>"
					shtmtabD = shtmtabD & " <div id='tabs-" & tpag &"'>" & sdataDoc &  "</div>"
					sdataDoc=""
					tpag=tpag +1
					idf=0

				End If

			Loop

			If lipag="" Then
				lipag=lipag & "<li class=lipag ><a href='#tabs-1'>Pagina 1</a></li>"
				shtmtabD = shtmtabD & " <div id='tabs-" & tpag &"'>" & sdataDoc &  "</div>"
			End If
			shtmtab="<div id='tabs'>"
			shtmtab=shtmtab & "  <ul>"
			shtmtab=shtmtab & lipag
			shtmtab=shtmtab & " </ul>"
			shtmtab=shtmtab & shtmtabD
			shtmtab=shtmtab &  "</div>"
			shtml=shtml & shtmtab  ' &  sdataDoc




			Case "S1" ' nueva visualizacion
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


			'shtml = OFV.FormHeader(Session("FormN"))

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

			'sHTML = sHTML & "<br><form method=post action='" & sHref2 & "'><input type=text name=buscar id=buscar value='" & search_phrase & "' class=dt xclass='ui-widget ui-state-default ui-corner-all'></input><input type=submit class=bietx xclass='ui-widget ui-state-default ui-corner-all' value=Buscar></input></form>"
			'shtml = shtml & CorPnlFnc("N", "")
			celda=""
			fila=""
			idf=0
			shtmldocs=""
			shtmtabD=""
			lipag=""
			tpag=1
			xentros=False
			sdataDoc=""
			Do While RETORNAVALOR(RSC,0) <> "-1"
				iddoc=rsc("codigo") & ""
				atit = titsolapa(rsc("codigo"), rsc("tipo"))
				shtmldocs= shtmldocs & "<div class='infodoc' data-view='dataview" & tpag & "' data-id='datadoc" & iddoc & "' xstyle='border-radius: 5px;background: #73AD21;'><a class=nt>"  & EsqnomPx(rsc("tipo"), rsc("codigo"), ofv)  & "</a>"
				sh="<a class='bietx' href='Reportes.asp?ID=" & iddoc &"&Titulo=Esquemas_Normativos&perfil=" & RSC("tipo")& "&amp;tipobj=&GLBSOLN=ESPREP' target='_new'>Impresion</a>"

				sdataDoc= sdataDoc &  "<div id='datadoc" & iddoc & "' style='display:none'><div class=viewdoc style='z-index:9999'> " & armatip(iddoc,rsc("tipo"))
				Do While  iddoc = RETORNAVALOR(RSC,0)

					xacceso = "codigo*" & rsc("codigo") & "_tipo*" & rsc("tipo") & "_perfil*OBJARC"
					xacceso = Replace(xacceso, "_", "&")
					If tipocpt="K" Then
						xacceso = Replace(xacceso, "*", "=") & "&NIV1=15&NIV2=" & rsc("index1")
					Else
						xacceso = Replace(xacceso, "*", "=") & "&NIV1=" & rsc("solapa") & "&NIV2=" & rsc("index1")
					End If
					sActcon = "Consultar.asp?Ysearch=S&DOCMODAL=CL&ACC=XP&" & xacceso & "&Datos=*smenu=PORTEXTO"
					clickn = Chr(34) & "AbrirCons('" & sActcon & "','" & rsc("tipo") & "');return false;" & Chr(34)
					sdataDoc = sdataDoc & "<button onclick=" & clickn & " class='ui-button ui-widget ui-corner-all'><a class=nt>"& GETSOLAPA(rsc("solapa"), rsc("index1")) &"</a></button>"


					rsc.MOVENEXT
				Loop
				sdataDoc = sdataDoc & "<button  class='ui-button ui-widget ui-corner-all'>" & sh &"</button>"

				sdataDoc = sdataDoc & "</div></div>"
				shtmldocs=shtmldocs  & "</div>"
				idf=idf +1
				If idf=10 Then
					lipag=lipag & "<li class=lipag ><a href='#tabs-"& tpag &"'>Pagina "& tpag &"</a></li>"


					shtmtabD = shtmtabD & " <div id='tabs-"& tpag &"'><div xclass='accordion'>"
					shtmtabD=shtmtabD & "<table  width=100% ><tr><td width=60% >" & shtmldocs & "</td><td width=40% class=tdview id='dataview" & tpag & "' ></td></tr></table>"
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
			shtml=shtml & shtmtab &  sdataDoc
			'SHTML=SHTML & ofv.botoneraESSX("0","","","","","","","","","no","","","") & "</body></html>"
			Case Else
			Do While Not rsc.EOF And ipag < limitpag

				iddoc = rsc(0)
				If existeDoc(iddoc,rsc("tipo"),tipodocx) Then
					solapa10 = False
					solapa5 = False
					xo = xo + 1
					tips=armatip(rsc(0),rsc("tipo"))
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
						Formularioz=  Formularioz & NewLine(0) & NewLine(0) & "<a  class=nt onMouseout=""hideddrivetip()"" onMouseover=""ddrivetip('"&armatip(rsc(0),rsc("tipo"))&"',250,'#dddddd')"" onclick="&clickn&" href=""" & url & search_maintain_search_term & """ style=""" & m_style_title & """>" & tipodocx & " : " &  Replace(rsc("title")," ","&nbsp;") & "</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
					Else
						Formularioz=  Formularioz & NewLine(0) & NewLine(0) & "<a  class=nt onMouseout=""hideddrivetip()""  onMouseover=""ddrivetip('"&armatip(rsc(0),rsc("tipo"))&"',250,'#dddddd')"" onclick="&clickn&" href=""" & url & search_maintain_search_term & """ style=""" & m_style_title & """>" &  Replace(rsc("title")," ","&nbsp;") & "</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
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
							Formularioz = Formularioz & NewLine(0) & "<br>" & texto & "<hr>"
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
					response.write  "<div sTYLE='float: left; z-index:1;overflow:auto;width:" & anchopag & ";border-style:solid;border-width:2;border-color:#ffffff;' id=pag"&ipag&" ><br>" & Chr(13)  & formularioz & NewLine(0) &  "<br>" & texto & "<hr>" & Chr(13)  &"</div>" & Chr(13)
					response.flush
				Else
					'    Formularios=Formularios & "<div id=pag"&ipag&"  style='display=none;'>" & chr(13)  & formularioz & chr(13)  &"</div>" & chr(13)

					If ipag < limitpag Then
						response.write "<div id=pag"&ipag&"  sTYLE='float: left; z-index:1;overflow:auto;width:" & anchopag & " ;border-style:solid;border-width:2;border-color:#ffffff;display:none;'>" & Chr(13)  & formularioz & NewLine(0) &  "<br>" & texto & "<hr>" & Chr(13)  &"</div>" & Chr(13)
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
			shtml =   shtml & " var op1=document.getElementById('infoP')" & Chr(13)
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
	shtmlx=    "<script>  function toggleDescription(btn) {"

	shtmlx=shtmlx &   "var description = btn.parentNode.nextElementSibling;"
	shtmlx=shtmlx &   "description.classList.toggle('expanded');"
	shtmlx= shtmlx &  "btn.textContent = description.classList.contains('expanded') ? 'Ver menos' : 'Ver m·s';"
	shtmlx= shtmlx & "}"
	shtmlx=shtmlx &    "   function mostrarTodos(id) {alert('1');"
	shtmlx=shtmlx &    " const contenedor = document.getElementById(id);"
	shtmlx=shtmlx &     " contenedor.classList.add('show-all');"
	shtmlx=shtmlx & "    } </script>"



	response.write shtml   &  ofv.botoneraESSX("0","","","","","","","","","no","","","") &  shtmlx  & "</body></html>"


End If


function armafiltros()



'      <!-- Facet tipo SELECT con optgroups -->
smbsq=   "      <section class='em-facet' data-facet='empresa' aria-expanded='true'>"
smbsq= smbsq & "        <div class='em-facet-head'>"
smbsq= smbsq & "          <h4 class='em-facet-title'>Empresa</h4>"
smbsq= smbsq & "          <button class='em-facet-toggle' type='button'><span class='em-caret'><i class='la la-angle-up em-caret' aria-hidden='true'></i></span></button>"
smbsq= smbsq & "        </div>"
smbsq= smbsq & "        <div class='em-facet-body'>"
smbsq= smbsq & "          <select class='em-select' data-facet='empresa'>"
smbsq= smbsq & "            <option value=''>Todas</option>"
'smbsq= smbsq & "            <optgroup label='' data-group-id='PV'>"
smbsq= smbsq & "              <option value='BSF' data-item-id='BSF'>Banco Santa Fe</option>"
smbsq= smbsq & "              <option value='BSJ' data-item-id='BSJ'>Banco San Juan</option>"
smbsq= smbsq & "              <option value='BER' data-item-id='BER'>Banco Entre Rios</option>"
smbsq= smbsq & "              <option value='BSC' data-item-id='BSC'>Banco Santa Cruz</option>"
smbsq= smbsq & "              <option value='GP'  data-item-id='GP'>Grupo Petersen</option>"
'smbsq= smbsq & "            </optgroup>"
smbsq= smbsq & "          </select>"
smbsq= smbsq & "        </div>"
smbsq= smbsq & "      </section>"


'      <!-- Facet CHECKBOX con grupos -->
smbsq= smbsq & "      <section class='em-facet' data-facet='TEMA' aria-expanded='true'>"
smbsq= smbsq & "        <div class='em-facet-head'>"
smbsq= smbsq & "          <h4 class='em-facet-title'>TEMAS</h4>"
smbsq= smbsq & "          <button class='em-facet-toggle' type='button'><span class='em-caret'><i class='la la-angle-up em-caret' aria-hidden='true'></i></span></button>"
smbsq= smbsq & "        </div>"
smbsq= smbsq & "        <div class='em-facet-body'><div class=contg>"
'          <!-- Grupo -->

SQL="SELECT TOP 20 codtema,DESCRI FROM tema ORDER BY DESCRI"
Set rssmbsq = ofv.crearconsultaEx(SQL,cn,1,parametros)

DO WHILE NOT rssmbsq.EOF
  smbsq= smbsq & "          <label class='em-chk'>"
  smbsq= smbsq & "            <input type='checkbox' name='TEMA' value='" & rssmbsq("DESCRI") & "'"
smbsq= smbsq & "                   data-item-id='" & rssmbsq("CODTEMA") & "'  data-group-id='TEMA'>"
  smbsq= smbsq &             rssmbsq("DESCRI")
  smbsq= smbsq & "          </label>"

rssmbsq.MOVENEXT
LOOP
Call OFV.cerrarconsulta(rssmbsq)



smbsq= smbsq & "        </div></div>"
smbsq= smbsq & "      </section>"





'      <!-- Facet CHECKBOX con grupos -->
smbsq= smbsq & "      <section class='em-facet' data-facet='SECTOR' aria-expanded='true'>"
smbsq= smbsq & "        <div class='em-facet-head'>"
smbsq= smbsq & "          <h4 class='em-facet-title'>SECTOR</h4>"
smbsq= smbsq & "          <button class='em-facet-toggle' type='button'><span class='em-caret'><i class='la la-angle-up em-caret' aria-hidden='true'></i></span></button>"
smbsq= smbsq & "        </div>"
smbsq= smbsq & "        <div class='em-facet-body'><div class=contg>"
'          <!-- Grupo -->

SQL="SELECT DISTINCT eco.eco,ECO.descri "
SQL=SQL & " FROM OBJSECTOR "
SQL=SQL & "inner join ECO on (OBJSECTOR.codsector = ECO.eco) "
SQL=SQL & "inner join OBJETOS on (OBJSECTOR.CODOBJETO = OBJETOS.CODOBJETO) "
SQL=SQL & "WHERE eco.eco > 0 AND OBJETOS.OBJUN=1"
Set rssmbsq = ofv.crearconsultaEx(SQL,cn,1,parametros)

DO WHILE NOT rssmbsq.EOF
  smbsq= smbsq & "          <label class='em-chk'>"
  smbsq= smbsq & "            <input type='checkbox' name='SECTOR' value='" & rssmbsq("DESCRI") & "'"
  smbsq= smbsq & "                   data-item-id='" & rssmbsq("eco") & "'  data-group-id='SECTOR'>"
  smbsq= smbsq &             rssmbsq("DESCRI")
  smbsq= smbsq & "          </label>"

rssmbsq.MOVENEXT
LOOP
Call OFV.cerrarconsulta(rssmbsq)



smbsq= smbsq & "        </div></div>"
smbsq= smbsq & "      </section>"



'      <!-- Facet CHECKBOX con grupos -->
smbsq= smbsq & "      <section class='em-facet' data-facet='autor' aria-expanded='true'>"
smbsq= smbsq & "        <div class='em-facet-head'>"
smbsq= smbsq & "          <h4 class='em-facet-title'>AUTOR</h4>"
smbsq= smbsq & "          <button class='em-facet-toggle' type='button'><span class='em-caret'><i class='la la-angle-up em-caret' aria-hidden='true'></i></span></button>"
smbsq= smbsq & "        </div>"
smbsq= smbsq & "        <div class='em-facet-body'><div class=contg>"
'          <!-- Grupo -->

SQL="SELECT distinct AUTOR FROM OBJPRTY  WHERE ESTADO=3"
Set rssmbsq = ofv.crearconsultaEx(SQL,cn,1,parametros)

DO WHILE NOT rssmbsq.EOF
  smbsq= smbsq & "          <label class='em-chk'>"
  XXX=GETINFOUSER(rssmbsq("AUTOR"))
  smbsq= smbsq & "            <input type='checkbox' name='autor' value='" & SESSION("SGU_NOMBRE") & "'"
  smbsq= smbsq & "                   data-item-id='" & rssmbsq("AUTOR") & "'>"
  smbsq= smbsq &             SESSION("SGU_NOMBRE")
  smbsq= smbsq & "          </label>"

rssmbsq.MOVENEXT
LOOP
Call OFV.cerrarconsulta(rssmbsq)



smbsq= smbsq & "        </div></div>"
smbsq= smbsq & "      </section>"





  '    <!-- TIP: para sumar un facet nuevo copi√° un <section.em-facet> y ajust√° data-facet, labels y data-* -->
 smbsq= smbsq & "   </aside>"







armafiltros=smbsq
end function
function armafiltros2()
smbsq="	 <div class='app' id='app'>"

smbsq=smbsq & "	    <aside class='sidebar' id='sidebar' aria-label='Filtros'>"
smbsq=smbsq & "	      <div class='side-head'>"
smbsq=smbsq & "	        <button class='btn-toggle' id='btnToggle' title='Ocultar/Mostrar'>‚ùÆ</button>"
smbsq=smbsq & "	      </div>"

'	      <!-- Encabezado de resultados en el men√∫ -->
smbsq=smbsq & "	      <div id='sidebarResults' aria-live='polite'>"
smbsq=smbsq & "	        <h3 class='title' id='resultsTitle'>Resultados</h3>"
smbsq=smbsq & "	        <div class='count'><span id='resultsCount'>10</span> resultados</div>"
smbsq=smbsq & "	      </div>"

smbsq=smbsq & "	      <section class='facet'>"
smbsq=smbsq & "	        <h4>Tipo de esquema</h4>"
smbsq=smbsq & "	        <div class='content'>"
smbsq=smbsq & "	          <label class='chk'><input type='checkbox' name='tipo' value='NORMA'> Norma</label>"
smbsq=smbsq & "	          <label class='chk'><input type='checkbox' name='tipo' value='PROCEDIMIENTO'> Procedimiento</label>"
smbsq=smbsq & "	          <label class='chk'><input type='checkbox' name='tipo' value='INSTRUCTIVO'> Instructivo</label>"
smbsq=smbsq & "	          <label class='chk'><input type='checkbox' name='tipo' value='MANUAL'> Manual</label>"
smbsq=smbsq & "	          <label class='chk'><input type='checkbox' name='tipo' value='OTROS'> Otros</label>"
smbsq=smbsq & "	        </div>"
smbsq=smbsq & "	      </section>"

smbsq=smbsq & "	      <section class='facet'>"
smbsq=smbsq & "	        <h4>Sector</h4>"
smbsq=smbsq & "	        <div class='content'>"
smbsq=smbsq & "	          <label class='chk'><input type='checkbox' name='sector' value='Comercial'> Comercial</label>"
smbsq=smbsq & "	          <label class='chk'><input type='checkbox' name='sector' value='Operaciones'> Operaciones</label>"
smbsq=smbsq & "	          <label class='chk'><input type='checkbox' name='sector' value='Legales'> Legales</label>"
smbsq=smbsq & "	          <label class='chk'><input type='checkbox' name='sector' value='Riesgo'> Riesgo</label>"
smbsq=smbsq & "	          <label class='chk'><input type='checkbox' name='sector' value='TI'> TI</label>"
smbsq=smbsq & "	        </div>"
smbsq=smbsq & "	      </section>"
smbsq=smbsq & "	    </aside>"

'	    <!-- ===== Chips seleccionados ===== -->
smbsq=smbsq & "	    <main>"
smbsq=smbsq & "	      <h3 class='section-title'>Filtros seleccionados</h3><form id=filtrosbsq>"
smbsq=smbsq & "	      <div id='chipsPanel' class='chips-panel' aria-live='polite'>"
smbsq=smbsq & "	        <span class='chips-label'>Filtros:</span>"
smbsq=smbsq & "	        <button id='btnClearAll' class='btn-clear-all' type='button'>Quitar todos</button>"
smbsq=smbsq & "	      </div></form>"
smbsq=smbsq & "	    </main>"
smbsq=smbsq & " </div>"

armafiltros2=smbsq
end function


Function armatip(idC,tipo)
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
		Sesquema=EsqnomPx(tipo, idC, ofv)
		TIPOCPT=recordsetip("TIPOCPT")
		texto = "<u><i>" & Sesquema &"</i></u> <br><br><B><U>TEMA</U> :</b>&nbsp;&nbsp;" & recordsetip("TEMADESCRI") & "<BR><B><U>SUBTEMA</U> :</b>&nbsp;&nbsp;" & recordsetip("STEMADESCRI")    '& "<BR>"
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
		texto = texto & "<BR><B><U>VERSI”N</U>  :&nbsp;&nbsp;</B>" & zVersion & "  "    '& idC
		armatip = "<b>" & texto & "</b>"
	Else
		armatip=""
	End If
	Call OFV.cerrarconsulta(recordsetip)
End Function


Function armatip2(idC,tipo,rsc,Asolapas,Adocadj)
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
		Sesquema=EsqnomPx(tipo, idC, ofv)

		xacceso = "codigo*" & rsc("codigo") & "_tipo*" & rsc("tipo") & "_perfil*OBJARC"
		xacceso = Replace(xacceso, "_", "&")
		If tipocpt="K" Then
			xacceso = Replace(xacceso, "*", "=") & "&NIV1=15&NIV2=" & rsc("index1")
		Else
			xacceso = Replace(xacceso, "*", "=") & "&NIV1=" & rsc("solapa") & "&NIV2=" & rsc("index1")
		End If
		sActcon = "Consultar.asp?Ysearch=S&DOCMODAL=CL&ACC=XP&" & xacceso & "&Datos=*smenu=PORTEXTO"
		clickn = Chr(34) & "AbrirCons('" & sActcon & "','" & rsc("tipo") & "');return false;" & Chr(34)
		texto = texto & " <div class='result'>"
		texto = texto & "    <img class='thumbnail' src='https://cdn-icons-png.flaticon.com/512/3095/3095583.png' alt='doc1'>"
		texto = texto & "    <div class='info'>"
		texto = texto & "      <h3><A CLASS=nt onclick="& clickn & ">" & Sesquema & "</a></h3>"
		texto = texto & "  <div class='solapas-row'>"
		texto = texto & "        <div class='solapas'>"
		For i=0 To ubound(Asolapas) -1
			texto = texto & "          <a href='#'>" & Asolapas(i) & "</a>"
		Next
		texto = texto & "        </div>"
		texto = texto & "        <button class='ver-mas-btn' onclick='toggleDescription(this)'>Ver m·s</button>"
		texto = texto & " </div>"
		texto = texto & "      <div class='description'>"
		texto = texto & left(rsc("text"),350) '"        Procedimiento interno para el manejo de cuentas corrientes mediante cheque(O). Se establece la normativa aplicable para validaciÛn y endoso."
		texto = texto & "      </div>"
		texto = texto & "    </div>"
		texto = texto & "<div class='attachments-columns' id='adjuntos'" & idc &">"
		texto = texto & "      <a href='#'><i class='la la-file-alt'></i> Cheque_formulario.pdf</a>"
		texto = texto & "      <a href='#'><i class='la la-file-alt'></i> Manual_operativo.pdf</a>"
		texto = texto & "      <a href='#'><i class='la la-file-alt'></i> Instructivo_2024.pdf</a>"
		texto = texto & "      <a href='#'><i class='la la-file-alt'></i> Circular_145.pdf</a>"
		texto = texto & "     <a href='#'><i class='la la-file-alt'></i> Modelo_endoso.pdf</a>"
		texto = texto & "     <a href='#'><i class='la la-file-alt'></i> Reglamento_bancario.pdf</a>"
		texto = texto & "      <a href='#' class='extra'><i class='la la-file-alt'></i> Anexo_legal.pdf</a>"
		texto = texto & "      <a href='#' class='extra'><i class='la la-file-alt'></i> Referencias_externas.pdf</a>"
		texto = texto & "      <button class='ver-todos-btn' onclick='mostrarTodos('adjuntos" &idc &"')'>Ver todos los archivos adjuntos</button>"
		texto = texto & "</div>"


		If recordsetip("VERSION") = "O" Then
			zVersion = "(O)"
		Else
			xx= recordsetip("NROVER")
			zVersion = "V (" &  xx & ")"
		End If

		texto = texto & " 	<div class='tooltip-icon'><span class='dots'><i class='la la-ellipsis-v'></i></span></div>"
		texto = texto & "    <div class='tooltip-content'>"
		texto = texto & "      <div>"
		texto = texto & "        <p><strong>Tema:</strong> COMUNICADOS</p>"
		texto = texto & "        <p><strong>SubTema:</strong> MP CTAS CTES</p>"
		texto = texto & "        <p><strong>Estado:</strong> Vigente</p>"
		texto = texto & "        <p><strong>VersiÛn:</strong> 1</p>"
		texto = texto & "     </div>"
		texto = texto & "      <div>"
		texto = texto & "        <p><strong>CÛdigo:</strong> RE-0002</p>"
		texto = texto & "        <p><strong>Vigencia:</strong> 27/11/2024</p>"
		texto = texto & "      </div>"
		texto = texto & "	</div>"

		texto = texto & "</div>"


		armatip2 = texto
	Else
		armatip2=""
	End If
	Call OFV.cerrarconsulta(recordsetip)
End Function

Function armatip3(idC,tipo,rsc)
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
		Sesquema=EsqnomPx(tipo, idC, ofv)

		xacceso = "codigo*" & rsc("codigo") & "_tipo*" & rsc("tipo") & "_perfil*OBJARC"
		xacceso = Replace(xacceso, "_", "&")
		If tipocpt="K" Then
			xacceso = Replace(xacceso, "*", "=") & "&NIV1=15&NIV2=" & rsc("index1")
		Else
			xacceso = Replace(xacceso, "*", "=") & "&NIV1=" & rsc("solapa") & "&NIV2=" & rsc("index1")
		End If
		sActcon = "Consultar.asp?Ysearch=S&DOCMODAL=CL&ACC=XP&" & xacceso & "&Datos=*smenu=PORTEXTO"
		clickn = Chr(34) & "AbrirCons('" & sActcon & "','" & rsc("tipo") & "');return false;" & Chr(34)
		texto = texto & " <div class='result'>"
		texto = texto & "    <img class='thumbnail' src='https://cdn-icons-png.flaticon.com/512/3095/3095583.png' alt='doc1'>"
		texto = texto & "    <div class='info'>"
		texto = texto & "      <h3><A CLASS=nt onclick="& clickn & ">" & Sesquema & "</a></h3>"
		texto = texto & "      <small>Sitio correspondiente: Banca minorista</small>"
		texto = texto & "      <div class='attachment'><i>*</i> Anexo.pdf</div>"
		texto = texto & "      <div class='description'>"
		texto = texto & left(rsc("text"),150) '"        Procedimiento interno para el manejo de cuentas corrientes mediante cheque(O). Se establece la normativa aplicable para validaciÛn y endoso."
		texto = texto & "      </div>"
		texto = texto & "    </div>"
		texto = texto & "    <div class='tags'>"
		texto = texto & "      <span class='tag'>Tema: " & recordsetip("TEMADESCRI") & "</span>"
		texto = texto & "      <span class='tag'>SubTema: " &recordsetip("STEMADESCRI") & "</span>"
		texto = texto & "      <span class='tag'>Estado: " & recordsetip("ESTADOD") & "</span>"
		texto = texto & "      <span class='tag'>Vigencia: " & fechadorcor(CDate(CLng("0" & recordsetip("fecvigen")) + CLng(CDate("01-01-1900"))), "") & "</span>"
		texto = texto & "     <span class='tag'>Cod Ident: " & recordsetip("CODIDEN") & "</span>"

		If recordsetip("VERSION") = "O" Then
			zVersion = "(O)"
		Else
			xx= recordsetip("NROVER")
			zVersion = "V (" &  xx & ")"
		End If

		texto = texto & "<span class='tag'>VersiÛn:" &zVersion & "</span>"
		texto = texto & "    </div>"
		texto = texto & "</div>"


		armatip3 = texto
	Else
		armatip3=""
	End If
	Call OFV.cerrarconsulta(recordsetip)
End Function

Function armatip3(idC,tipo)
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
		Sesquema=EsqnomPx(tipo, idC, ofv)
		TIPOCPT=recordsetip("TIPOCPT")
		texto =""' "<div>" & Sesquema &"</div>"
		texto = texto &  "<div>Tema: </i></u>" & recordsetip("TEMADESCRI") & "</div>"
		texto = texto &  "<div>SubTema: " & recordsetip("STEMADESCRI") & "</div>"
		texto = texto &  "<div>Estado: " &recordsetip("ESTADOD") & "</div>"
		texto = texto &  "<div>VIGENCIA: " & fechadorcor(CDate(CLng("0" & recordsetip("fecvigen")) + CLng(CDate("01-01-1900"))), "") & "</div>"
		texto = texto &  "<div>COD. IDENT.: " &  recordsetip("CODIDEN") & "</div>"

		If recordsetip("VERSION") = "O" Then
			zVersion = "(O)"
		Else
			xx= recordsetip("NROVER")

			zVersion = "V (" &  xx & ")"
		End If

		texto = texto &  "<div>VERSI”N: " &  zVersion & "</div>"


		armatip3 = texto
	Else
		armatip3=""
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
		TIT = "" & atit(1)    '"Solapa propiedades"
		Case "2"
		TIT = "" & atit(2)    '"Solapa consideraciones"
		Case "3"
		TIT = "" & atit(3)    '"Solapa sectores"
		Case "4"
		TIT = "" & atit(4)    ' "Solapa eventos"

		Case "5"
		TIT = "" & atit(5)    '"Solapa procedimientos"
		'TIT = TIT & "<b>" & bsqevt(evt) & "</b>"
		Case "6"
		TIT = "" & atit(6)    '"Solapa descripcion"
		Case "7"
		TIT = "" & atit(7)    '"Solapa controles"
		Case "8"
		TIT = "" & atit(8)    '"Solapa esquemas"
		Case "9"
		TIT = "" & atit(9)    '"Solapa tablas"
		Case "10"
		TIT = "" & atit(10)    '"Solapa documentos"
		Case "11"
		TIT = "" & atit(11)    '"Solapa contenido"
		Case "12"
		TIT = "" & atit(12)    '"Solapa cursograma"
		Case "13"
		TIT = "" & atit(13)    '"Solapa contratos"
		Case "14"
		TIT = "" & atit(14)    '"Solapa faq"
		Case "15"
		TIT = "" & atit(15)    '"Solapa general"
		Case Else
		TIT = "Sin descripciÛn"




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
		texto = reemplaza(texto, "[¿¡ƒ·‡‰]", "A")
		texto = reemplaza(texto, "[»…ÀËÈÎ]", "E")
		texto = reemplaza(texto, "[ÃÕœÏÌÔ]", "I")
		texto = reemplaza(texto, "[“”÷ÚÛˆ]", "O")
		texto = reemplaza(texto, "[Ÿ⁄‹˘˙¸]", "U")
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
	Set regEx = New RegExp            ' Crea una expresiÛn regular.
	regEx.Pattern = patrn            ' Establece el modelo.
	regEx.IgnoreCase = True            ' Hace distinciÛn entre may˙sculas y min˙sculas.
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

Function InArray(theArray,theValue)
	Dim i, fnd
	fnd = False
	For i = 0 To UBound(theArray)
		If theArray(i) = theValue Then
			fnd = True
			Exit For
		End If
	Next
	InArray = fnd
End Function


Function RETORNAVALOR(RST,POS)

	If Not RST.EOF Then
		RETORNAVALOR=CSTR(RST(POS)& "")
	Else
		RETORNAVALOR="-1"
	End If
End Function
%>
