

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
	verDocs=request.querystring("verDocs")& ""
	Response.AddHeader "Content-Type", "text/html; charset=iso-8859-1"
	' response.write verDocs
	' on error resume next
	vuelta = OFV.vueltaasp("../", CantRegMostrarx, CantRegMoverx)
    CantRegMoverx = cdbl("0" & request.querystring("CantRegMover"))
    'response.write CantRegMoverx
	Set CN = OFV.conectar(OFV.strconn2)

	sUN = 1
	datofiltro=request.form("filtros")& ""
	'		response.write datofiltro

	If datofiltro<>"" Then

		arrf=split(datofiltro,"####")
		TIPOF=" AND ("
		autorF=" AND ("
		TEMAF=" AND ("
		SECTORF=" AND ("
		DELIM=""
		DELIMA=""
		DELIMT=""
		DELIMS=""
		For i= 0 To ubound(arrf)

			ColF=split(arrf(i)  ,":")

			Select Case UCASE(ColF(0)&"")
				Case "TIPO"
				TIPOF=TIPOF & DELIM &  CHR(34)   & "[[TIPO_" & FiltroVal(ColF(1)) &"]]" & CHR(34)
				DELIM=" OR "

				Case "AUTOR"
				autorF=autorF & DELIMA &  CHR(34)   & "[[AUTOR_" & FiltroVal(ColF(1)) &"]]" & CHR(34)
				DELIMa=" OR "

				Case "TEMA"
				TEMAF=TEMAF & DELIMT &  CHR(34)   & "[[TEMA_" & FiltroVal(ColF(1)) &"]]" & CHR(34)
				DELIMT=" OR "


				Case "SECTOR"
				SECTORF=SECTORF & DELIMS &  CHR(34)   & "[[CODSECTOR_" & FiltroVal(ColF(1)) &"]]" & CHR(34)
				DELIMS=" AND "

			End Select

		Next
		TIPOF=TIPOF & ")"
		autorF=autorF & ")"
		TEMAF=TEMAF & ")"
		SECTORF=SECTORF & ")"
		If delim="" Then tipof=""
		If delimA="" Then autorF=""
		If DELIMT="" Then TEMAF=""
		If DELIMS="" Then SECTORF=""

		FILTROSF=TIPOF &  autorF & TEMAF & SECTORF '& " and " & chr(34) & "[[dato_ale]]" & chr(34)
	End If
	XXXX = NEWEMPRE("NO")
	'RESPONSE.WRITE FILTROSF

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



	sUN = session("UN")
	Dim atit
	anchopag = "98%"
	Dim TIPOCPT
	Dim tipodoc

	search_phrase = Replace(Request.querystring("YIDER")& "",Chr(34),"")
	search_phrase=Replace(search_phrase,"'","")
	'loader="<div id='loader-wrapper' STYLE='display:block'>"
		'  loader= loader & "        <div class='loader'></div>"
		   loader= "        <p id='loader-wrapper'>Procesando busqueda....</p>"
		'   loader= loader & "    </div>"
     response.write loader
     response.flush
	If search_phrase = "" Then

		'search_phrase = Replace(Request.Form("buscar") & "",Chr(34),"")
		'search_phrase=Replace(search_phrase,"'","")
		'response.write("No se encontro documentos, por favor ingrese un documento a buscar")

	Response.Write "<script type='text/javascript'>"

	Response.Write "alert('Debe ingresar mÌnimo 3 caracteres para la busqueda. No se encontraron documentos que coincidan');"

	'Response.Write "window.location.reload();"

    Response.Write "window.location.href = 'novedmenu.asp';"
    Response.Write "</script>"
    Response.End '''agrego response.end para que no cargue nada m·s. 05-01 - CG
	End If

	tipobsq = Request.querystring("b")
	FORMULARIOS=REQUEST.QUERYSTRING("FORMULARIOS")

	If FORMULARIOS <> "" Then
		session("formularios")=FORMULARIOS
	Else
		FORMULARIOS=session("formularios")
	End If
	ingreso=REQUEST.QUERYSTRING("ing")
	search_strings = Split(search_phrase , " ", -1, 1)

	xsdelim=""
	sresult=""
	For Each search_string In search_strings

		If  Len(search_string & "") > 3 Then

			sresult= sresult & xsdelim & search_string
			xsdelim= " "

		End If

	Next


	Session("yider") = search_phrase




	shtml = OFV.FormHeader(Session("FormN"))
	shtml = shtml & CorPnlFnc("N", "")
	xs=  " function AbrirCons(Datos,nom) {" & Chr(13)
	'xs = xs & " var h=570;w=790 ;"
	'If SESSION("APLMDCLI") = "BC" Then
		xs = xs & "if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;} "
	'End If
	xs = xs & " myWin= open(Datos"

	xs = xs & "," & Chr(34) & "ConsLnk" & Chr(34) & " + nom,"    '& Chr(34)
	xs = xs & "  'toolbar=no,width='+w+',height='+h+',top=0,left=0,resizable=yes'"    '& CHR(34)
	xs = xs & " ); " & Chr(13) & " } " & Chr(13)

	shtml = shtml & "<script>  function elim(str,len)"
	shtml = shtml & "  {"


	shtml = shtml & "  str = str.split(/[\s,;.?!:\n//]+/).filter(function ( str ) {"
	shtml = shtml & "      var word = str.match(/(\w+)/);"
	shtml = shtml & "      return word && word[0].length > len;"
	shtml = shtml & "   }).join(' ');"

	shtml = shtml & "   return str;"
	shtml = shtml & "  }" & xs
	shtml = shtml & "document.getElementById('buscar').value='" & search_phrase & "';"
	' [EXACTA] Casilla "Frase exacta" a continuacion del cuadro de busqueda; el estado se guarda en la cookie exacta (la lee CfgBusqueda)
	shtml = shtml & "(function(){var b=document.getElementById('buscar');if(!b||document.getElementById('exacta'))return;var l=document.createElement('label');l.style.cssText='font-size:12px;margin-left:6px;white-space:nowrap;cursor:pointer;';l.innerHTML='<input type=checkbox id=exacta style=vertical-align:middle> Frase exacta';var t=document.getElementById('btnBusc')||b;t.parentNode.insertBefore(l,t.nextSibling);var c=document.getElementById('exacta');c.checked=(document.cookie.indexOf('exacta=S')>=0);c.onchange=function(){document.cookie='exacta='+(c.checked?'S':'N')+'; path=/';};})();"

	shtmlx=    " function toggleDescription(btn) {"

	shtmlx=shtmlx &   "var description = btn.parentNode.nextElementSibling;"
	shtmlx=shtmlx &   "description.classList.toggle('expanded');"
	shtmlx= shtmlx &  "btn.textContent = description.classList.contains('expanded') ? 'Ver menos' : 'Ver m·s';"
	shtmlx= shtmlx & "}"
	shtmlx=shtmlx &    "   function mostrarTodos(id) {alert('1');"
	shtmlx=shtmlx &    " const contenedor = document.getElementById(id);"
	shtmlx=shtmlx &     " contenedor.classList.add('show-all');"
	shtmlx=shtmlx & "    } "


	shtml = shtml & "const myCustomEvent = new Event('myCustomEvent');"


	shtml = shtml &    "document.addEventListener('myCustomEvent', function() {"
	shtml = shtml &    "    alert('Custom event triggered!');"
	shtml = shtml & "}); function ale(){alert('1');}"

	shtml = shtml & shtmlx &"</script>"





	If session("APLMDCLI")="BPN" Then
		sql = "select 0,tipobjeto from tipobjeto  "
	Else
		sql = "select archivo,tipobjeto from tipobjeto  "
	End If

	Set rsESQ = OFV.crearconsultaEx(sql, CN, 1, volver)

	delim = ""
	orden_docFull = ""


	Do While Not rsESQ.EOF
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



	Formulario = ""    'ofv.Encabezados("info")




	' [LIMITE] techo de filas que se piden al componente; se muestran solo limitpag documentos y se informa el total
	techoFilas = 1000
	codTemaLista = ""
	diagFilas = ""
	Set docTipo = Server.CreateObject("Scripting.Dictionary")   ' codigo de documento -> tipo (esquema), para contar por esquema
	cantFilas = 0
	cantTotDoc = 0
	totTxt = 0
	' [FACETAS] verDocs=F: solo devuelve (texto) los valores de ESQUEMA / TEMA / AUTOR que quedan disponibles con los filtros aplicados
	If verDocs = "F" Then
		Call FacetDisponibles(datofiltro)
		Response.End
	End If
	Set cnbsq = Nothing   ' [PORT] fija cnbsq como variable global
	errBusq = 0
	Set findd = CfgBusqueda(FILTROSF)
	'and AUTOR_ADM and CODSECTOR_132 and CODSECTOR_41



	' [ERRORES] Frase con palabras reservadas / de ruido: no se ejecuta la busqueda pero se sigue por la misma rama de "sin documentos" (mantiene cabecera, buscador, filtros y "0 Documento/s encontrado/s")
	esRuido = (findd.IsNoiseWordex(findd.search_phrase) = 0)
	If False Then

	Else

		' response.write time() & "  1<br>"
		'Call findd.findDoc
		if session("bsqnew")="S" And Not esRuido Then
		Set objCmd = Server.CreateObject("ADODB.Command")

		' 2. CONFIGURACI”N INDISPENSABLE PARA SQLOLEDB: Cursor del lado del cliente
		cnbsq.CursorLocation = 3 ' 3 equivale a adUseClient

		' 3. Configuramos el objeto Command de forma tradicional
		With objCmd
		    Set .ActiveConnection = cnbsq
		    .CommandText = "dbo.BuscarDocumentosYider"
		    .CommandType = 4 ' 4 = adCmdStoredProc (Procedimiento Almacenado)

		    ' Mapeamos los par·metros estrictamente en el orden del SP
		    .Parameters.Append .CreateParameter("@return_value", 3, 4)
		    .Parameters.Append .CreateParameter("@BusquedaUsuario", 202, 1, 500, "cheque")
		    .Parameters.Append .CreateParameter("@MargenCaracteres", 3, 1, , 100)
		    .Parameters.Append .CreateParameter("@FiltroMetadata", 202, 1, 1000, Null)
		End With

		' 4. SOLUCCI”N AL ERROR '800a01a8': Instanciamos rsc y abrimos directamente desde el Recordset
		' Esto salta los mensajes intermedios autom·ticamente y puebla la variable con el SELECT final
		Dim rsc
		Set rsc = Server.CreateObject("ADODB.Recordset")
		rsc.Open objCmd

		else
		' [ERRORES] un error de la libreria (p. ej. palabra reservada en la frase) no debe mostrarse al usuario
		On Error Resume Next
		If Not esRuido Then Call findd.findDocex
		If Err.Number <> 0 Then errBusq = Err.Number
		Err.Clear
		On Error GoTo 0
		end if
		'response.write time() & " 2<br>"
		'response.write findd.search_phrase
		'			RESPONSE.WRITE findd.sqlj
		'Set rsc = findd.orsFx
		On Error Resume Next
		If session("bsqnew") <> "S" And Not esRuido Then Set rsc =  findd.rsext 'cnbsq.execute(findd.sqlj)
		If Err.Number <> 0 Then errBusq = Err.Number
		Err.Clear
		On Error GoTo 0
		' [ERRORES] sin recordset valido (error o sin objeto): se usa uno vacio -> "0 Documento/s encontrado/s"
		If Not IsObject(rsc) Then
			Set rsc = RsVacio()
		ElseIf rsc Is Nothing Then
			Set rsc = RsVacio()
		End If
		'response.write time() & " 3<br>"
		' [PAGINADO] antes: rsc.move CantRegMoverx (movia FILAS, dentro del mismo documento solo cambiaba el texto). Ahora se saltan documentos completos (ver mas abajo)
		url = "#Foo"
		totaldoc = 1
		ipag = 1
		Formularioz = ""
		' rsc.move 11
		' rsc.movenext 11
		solapa10 = False
		solapa5 = False

		If session("APLMDCLI") = "BV" Then
			limitpag=200
		Else
			limitpag=5
		End If
		' [PAGINADO] pagina = CantRegMoverx (0 = primera): se saltan CantRegMoverx * limitpag documentos completos; los saltados igual cuentan para las facetas
		docsSaltar = CLng(CantRegMoverx) * limitpag
		Do While docsSaltar > 0 And RETORNAVALOR(RSC,0) <> "-1"
			iddocS = rsc("codigo") & ""
			ConteoAgrega iddocS, rsc("tipo")
			Do While iddocS = RETORNAVALOR(RSC,0)
				rsc.MOVENEXT
				cantFilas = cantFilas + 1
			Loop
			docsSaltar = docsSaltar - 1
		Loop
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
			'shtmldocsF=shtmldocsF & "        <button class='em-toggle' id='btnToggle' title='Ocultar/Mostrar'><i class='la la-angle-left' aria-hidden='true'></i></button>"
			shtmldocsF=shtmldocsF & "      </div>"

			shtmldocsF=shtmldocsF & "      <div id='sidebarResults' aria-live='polite'>"
			shtmldocsF=shtmldocsF & "        <div class='em-title' id='resultsTitle'>Resultados </div> "
			shtmldocsF=shtmldocsF & "<li><font color=#000000 >Documento/s encontrado/s con la palabra :<i><span class='noved-doc-tag' style='font size:20px'>" &search_phrase&"</b></i></font></i>" 'agregado 12/12 CG
			shtmldocsF=shtmldocsF & "        <div class='em-count' style='font-size:15px;'><i><b><span id='resultsCount' ></span></b></i>&nbsp;&nbsp;Documento/s encontrado/s</div>"
			shtmldocsF=shtmldocsF & "</div>"
			shtmldocsF=  shtmldocsF &   "      <div id='chipsPanel' class='em-chips' aria-live='polite'>"
			shtmldocsF=  shtmldocsF &   "	        <span class='em-chips-label'>Filtros:</span>"
			shtmldocsF=  shtmldocsF &   "        <button id='btnClearAll' class='em-clear-all' type='button'>Quitar todos</button>"
			'shtmldocsF=  shtmldocsF &   " <form target=info id=sx action='searchex.asp?verDocs=S&ing=&FORMULARIOS=D&YIDER="&search_phrase&"' method=post><button class='em-filtro-all' id='ver'>Aplicar Filtros</button><input type=hidden name=filtros id=filtros value=''><form>"
			shtmldocsF=  shtmldocsF &   "	      </div>"


			'''ordenar
			'shtmldocsF=shtmldocsF & "<div class='sort-toolbar'>"
			'shtmldocsF=shtmldocsF & "<label for='sortResults' class='sort-label'>Ordenar por:</label>"
			'shtmldocsF=shtmldocsF & "<select id='sortResults' class='sort-select'>"
			'shtmldocsF=shtmldocsF & "<option value='relevance' selected>Relevancia (Original)</option>"
			'shtmldocsF=shtmldocsF & "<option value='asc'>Alfabeticamente (A - Z)</option>"
			'shtmldocsF=shtmldocsF & "<option value='desc'>Alfabeticamente (Z - A)</option>"
			'shtmldocsF=shtmldocsF & "</select>"
            'shtmldocsF=shtmldocsF  & "</div> "

            ' --- CONTENEDOR PRINCIPAL DE RESULTADOS ---
			' IMPORTANTE: Agregamos el ID 'results-container' para que el JS lo encuentre
			shtmldocsF = shtmldocsF & "<div id='results-container' class='content-scroll-area' style='padding: 20px;'>"
			shtmldocsF=  shtmldocsF & armafiltros()


			''''fin ordernar 19/01


			shtmldocs=""
			cant=1
			sul=""
			sli=""
			cantdoc=0
			cantTotDoc=0
			if RSC.eof then ' sin datos
			shtmldocs=" <div id=swiperbsp class='swiper mySwiper2' data-cant='0'>"
			            sli="<li><font color=#000000 >sin documentos que contenga :<i><span class='noved-doc-tag'>" &search_phrase&"</span></i></font></i>"
						shtmldocs= shtmldocs & "<div class='swiper-slide'>"& "<ul>" & sli & "</ul>" & "</div>"
						shtmldocs= shtmldocs & "    <div class='swiper-pagination'></div>"
						shtmldocs= shtmldocs & " </div>"
						'sul=sul & tarjeta


						'  <!-- ?? NUEVO: DIV contenedor de Resultados -->
						shtmldocsF=  shtmldocsF &   " <main>"
						'shtmldocsF=  shtmldocsF &   "	      <div class='em-section-title'>Filtros seleccionados</div>"
						shtmldocsF=  shtmldocsF &   "      <div id='chipsPanel' class='em-chips' aria-live='polite'>"
						shtmldocsF=  shtmldocsF &   "	        <span class='em-chips-label'>Filtros:</span>"
						shtmldocsF=  shtmldocsF &   "        <button id='btnClearAll' class='em-clear-all' type='button'>Quitar todos</button>"
						shtmldocsF=  shtmldocsF &   " <form target=info id=sx action='searchex.asp?cantregmover=0&verDocs=S&ing=&FORMULARIOS=D&YIDER="&search_phrase&"' method=post><button type='button' class='bietx' id='ver'>Aplicar Filtros</button><input type=hidden name=filtros id=filtros value=''><form>" 'button class='em-filtro-all'
						shtmldocsF=  shtmldocsF &   "	      </div>"

						shtmldocsF=  shtmldocsF &   "	      <div id='emResults' class='em-results' aria-live='polite'>"
						shtmldocsF=  shtmldocsF &   "	        <div class='em-results-head' id=rst name=rst >"
						'shtmldocsF=  shtmldocsF &   "	          <h4 id=rst name=rst class='em-results-title'><div>" & ahtmldocs & "</div></h4>"
						shtmldocsF=  shtmldocsF & shtmldocs
						shtmldocsF=  shtmldocsF &   "	        </div>"
						shtmldocsF=  shtmldocsF &   "	        <div class='em-results-empty' id='emResultsEmpty los filtros para ver resultados.</div>"
						shtmldocsF=  shtmldocsF &   "	        <div class='em-results-grid' id='emResultsGrid' hidden></div>"
						shtmldocsF=  shtmldocsF &   "	      </div>"
						shtmldocsF=  shtmldocsF &   "	    </main>"
						shtmldocsF=  shtmldocsF &   "	</div>"

						shtml=shtml & shtmldocsF  ' &  sdataDoc
			else
			tarjeta=""
			Do While RETORNAVALOR(RSC,0) <> "-1"  and cantdoc < limitpag
				Dim 	Asolpas
				Dim 	Aadjuntos
				ReDim 	Asolpas(100)

				ix=0

				iddoc=rsc("codigo") & ""
				ConteoAgrega iddoc, rsc("tipo")
				atit = titsolapa(rsc("codigo"), rsc("tipo"))

				strsql = "select a.temadescri,a.stemadescri,b.*,c.estado AS ESTADOD,b1.codiden,b11.TIPOCPT,b.autor as Autor from CSRELPATHDESCRINV a "
				strsql = strsql & " inner join objetos b1 on(b1.codobjeto= a.cobj)"
				strsql = strsql & " inner join TIPOBJETO b11 on(b11.TIPOBJETO= B1.TIPOBJETO)"
				strsql = strsql & " inner join objprty b on(b.codobjeto= a.cobj)"
				strsql = strsql & " inner join estados c on(c.objid=b.estado) where a.cobj=0" & iddoc
				'	response.write strsql & "<br>"
				Set recordsetip = CN.Execute(strsql)
				'tarjeta=""
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
					tarjeta=tarjeta & "<div class='infodoc'>" 'class='susc-item'
				'	tarjeta=tarjeta & "    <div class='media'>"
				'	tarjeta=tarjeta & "      <img src='https://images.unsplash.com/photo-1529101091764-c3526daf38fe?q=80&w=800&auto=format&fit=crop' alt=''>"
				'	tarjeta=tarjeta & "    </div>"

				'	tarjeta=tarjeta & "    <div class='meta'>"
				'	tarjeta=tarjeta & "      <span class='badge'><span class='label'>CÛdigo:</span> "&  recordsetip("CODIDEN") & "</span>"
				'	tarjeta=tarjeta & "      <span class='badge'><span class='label'>Vigencia:</span> " & fechadorcor(CDate(CLng("0" & recordsetip("fecvigen")) + CLng(CDate("01-01-1900"))), "")&"</span>"
				'	tarjeta=tarjeta & "    </div>"

					tarjeta=tarjeta & "    <div class='content2'><div class='title-row ' >"
					tarjeta=tarjeta & "      <div class='title'>"
					' tarjeta=tarjeta & Sesquema
					 tarjeta = tarjeta & "      <div CLASS='enlace-con-ellipsis' title='" & Sesquema & "' onclick="& clickn & " >" & Sesquema & "</div>"
					' tarjeta = tarjeta & "  <A xCLASS=nt title='" & Sesquema & "' onclick="& clickn & " style='word-wrap: break-word; overflow-wrap: break-word; word-break: break-word;'>" & Sesquema & "</a>"



					'tarjeta=tarjeta & "        <span class='authb-pill' >" & recordsetip("ESTADOD") & "</span>"
					tarjeta=tarjeta & "      </div>"
                    tarjeta=tarjeta & " <div class='doc-info'><i class='las la-info-circle la'></i>"
				'	tarjeta=tarjeta & "          <button class='info-btn' aria-label='VersiÛn del documento'>"
				'	tarjeta=tarjeta & "            <wsvg class=ttttt width='16' height='16' viewBox='0 0 24 24' fill='none' stroke='currentColor'>"
				'	tarjeta=tarjeta & "              <circle cx='12' cy='12' r='10' stroke-width='1.6'></circle>"
				'	tarjeta=tarjeta & "              <path d='M12 10v7' stroke-width='1.6' stroke-linecap='round'></path>"
				'	tarjeta=tarjeta & "              <circle cx='12' cy='7' r='1.2' fill='currentColor'></circle>"
				'	 tarjeta=tarjeta & "           </wsvg>"
				'	tarjeta=tarjeta & "          </button>"
					tarjeta=tarjeta & "          <div class='info-tip' id='docTooltip'>"
					IF session("APLMDCLI")<> "BPN" and session("APLMDCLI")<> "BC" THEN
					 tarjeta=tarjeta & "           <div class='info-row'><span class='label'>Autor:</span><span>" &recordsetip("AUTOR") & "</span></div>"
					END IF
					IF  session("APLMDCLI")<> "BC" THEN
					 tarjeta=tarjeta & "           <div class='info-row'><span class='label'>CÛdigo:</span><span>" &recordsetip("CODIDEN") & "</span></div>"
					END IF
					tarjeta=tarjeta & "           <div class='info-row'><span class='label'>Vigencia:</span><span>" & fechadorcor(CDate(CLng("0" & recordsetip("fecvigen")) + CLng(CDate("01-01-1900"))), "")& "</span></div>"
					  IF recordsetip("version") ="V" THEN
					   VERSION=recordsetip("version")  & " (" & recordsetip("NROVER") & ")"
					  ELSE
					   VERSION=recordsetip("version")
					  END IF
					 tarjeta=tarjeta & "           <div class='info-row'><span class='label'>VersiÛn:</span><span>" & VERSION & "</span></div>"

                  IF session("APLMDCLI")<> "BPN" and session("APLMDCLI")<> "BC" THEN
					if Recordsetip("fecvto")=0 then
					 tarjeta=tarjeta & "           <div class='info-row'><span class='label'>Vencimiento:</span><span>SIN FECHA</span></div>"
					ELSE
					 tarjeta=tarjeta & "           <div class='info-row'><span class='label'>Vencimiento:</span><span>" & fechadorcor(CDate(CLng("0" & recordsetip("fecvto")) + CLng(CDate("01-01-1900"))), "")& "</span></div>"

					end if
                  END IF
					' tarjeta=tarjeta & "           <div class='info-row'><span class='label'>Emision:</span><span>" & fechadorcor(CDate(CLng("0" & recordsetip("fecemis")) + CLng(CDate("01-01-1900"))), "")& "</span></div>"
					' tarjeta=tarjeta & "           <div class='info-row'><span class='label'>Publicacion:</span><span>" & fechadorcor(CDate(CLng("0" & recordsetip("fecpublic")) + CLng(CDate("01-01-1900"))), "")& "</span></div>"
					' tarjeta=tarjeta & "           <div class='info-row'><span class='label'>VersiÛn:</span><span>(v) 12</span></div>"
					 tarjeta=tarjeta & "         </div>"
        				tarjeta=tarjeta & "</div></div></div>"
					tarjeta=tarjeta & "      <nav class='noved-doc-path'>"
					tarjeta=tarjeta & "        <span i class='las la-folder'></i></span> " & recordsetip("TEMADESCRI") & "<i class='las la-angle-right'></i> <span> "&  recordsetip("STEMADESCRI")& "</span>"
					tarjeta=tarjeta & "      </nav>"

					      <!-- Texto con Ö M·s -->
					tarjeta=tarjeta & "      <div class='text-box' id='tb-1'>"
					tarjeta=tarjeta & "        <span class='excerpt'>"
					PALB=SPLIT(search_phrase," ")
					textodoc=EliminarAcentos(rsc("text").value)
					POSP=INSTR(ucase(textodoc),ucase(palb(0)))
					'response.write posp & "<br>"
					if POSP  > 0 then
					  textodoc=mid(textodoc & "", POSP ,350)
					  for i =0 to ubound(PALB)
					  textodoc=replace(lcase(textodoc),lcase(palb(i)),"<b><i>" & ucase(palb(i)) &"</i></b>")
                      next

					end if

					tarjeta=tarjeta & left(textodoc,150)& "<span class='dots'>...</span><span class='hidden'>"
					tarjeta=tarjeta & mid(textodoc,151,300) & "</span>"
					tarjeta=tarjeta & "        </span>"
					tarjeta=tarjeta & "        <a href='#' class='more'>M·s</a>"
					tarjeta=tarjeta & "      </div>"

					      <!-- Adjuntos (muestra hasta 4; resto al expandir) -->
					tarjeta=tarjeta & "      <div class='attachments' id='atts-1' data-max='4'>"




				End If
				hayCard = Not recordsetip.eof   ' [CUENTA] solo es documento (tarjeta) si tiene ruta / objeto / estado; si no, se salta
				Call ofv.cerrarconsulta(recordsetip)
				If hayCard Then cantdoc=cantdoc + 1
				lenTarj = Len(tarjeta)
				entro=False
				'response.write iddoc & "<br>"
				'if cantdoc= 6 then exit while
				Do While  iddoc = RETORNAVALOR(RSC,0) and cantdoc < limitpag

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

					If rsc("solapa")=10  Then
						strsql="SELECT archivo, directorio,descri  FROM Objhtm where OBJID=" & rsc("INDEX1") & " AND codobjeto=0" & rsc("codigo") & " and tipobjeto='" & rsc("tipo") & "' "
						' strsql=strsql & " and descri='" & rsc("title") & "'"
						ixarch=0
						ReDim 	Aadjuntos(100,2)
						Set recordsetadj = CN.Execute(strsql)
						Do While Not recordsetadj.eof
							sdirec = recordsetadj(1) & ""
							If sdirec = "" Then sdirec = "Archivos"
							sarch = sdirec & "/" & recordsetadj(0)

							Aadjuntos(ixarch,0)=replace(sarch &"" ,"//","/")
							sdoc=replace(sarch &"" ,"//","/")


							Aadjuntos(ixarch,1)=recordsetadj("descri") 'rsc("title")
							DOCTIT=recordsetadj("descri")
							IF TRIM(DOCTIT&"") ="" THEN DOCTIT="SIN TITULO"
							' Aadjuntos(ixarch,2)=rsc("title")
							ixarch=ixarch+1
                            if InStrRev(sdoc,".") > 0 then
							 tarjeta=tarjeta & "        <a class='chip' TARGET=_NEW href='" &sdoc & "'>"
							 tarjeta=tarjeta & "          <svg viewBox='0 0 24 24' fill='none' stroke='currentColor'><path d='M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z' stroke-width='1.6'/><path d='M14 2v6h6' stroke-width='1.6'/></svg>"
							 tarjeta=tarjeta & "          <span class='name'>" & DOCTIT &"</span><span class='tag'>"& UCASE(mid(sdoc,InStrRev(sdoc,"."),len(sdoc)) &"" )&"</span>"
							 tarjeta=tarjeta & "        </a>"
							end if


							recordsetadj.movenext
						Loop
						entro=True
					End If
					'armo solapas doc se puede repetir
					If Not InArray(Asolpas,Ssolapa) Then

					End If

					diagFilas = diagFilas & rsc("codigo") & "/" & rsc("tipo") & "/" & rsc("solapa") & ";"
					rsc.MOVENEXT
					cantFilas = cantFilas + 1
				Loop



                tarjeta=tarjeta & "      </div>"
				if entro then
				tarjeta=tarjeta & "      <span  class='btn-files' data-target='#atts-1' data-more='M·s archivos' data-less='Menos archivos' style='display:none;'>M·s archivos</span>"
				end if
				tarjeta=tarjeta & "    </div>"
  				tarjeta=tarjeta & "</article>"
				If Not hayCard Then tarjeta = Left(tarjeta, lenTarj)   ' [CUENTA] sin tarjeta: no deja restos
				sli= sli & "<li ><i class=Ciclass></i>"  & tarjeta & "</li>"
				If cant mod 1115 = 0  Then
					sul=sul & "<div class='swiper-slide'>"& "<ul>" & sli & "</ul>" & "</div>"
					sli=""
					cant=1
				Else
					cant=cant +1
				End If
				'shtmldocs=shtmldocs & sul

			Loop

			' [LIMITE] Documentos que no se dibujan (pasaron limitpag): solo se cuentan y se juntan sus codigos para el conteo por tema
			Do While RETORNAVALOR(RSC,0) <> "-1"
				iddocR = rsc("codigo") & ""
				ConteoAgrega iddocR, rsc("tipo")
				Do While iddocR = RETORNAVALOR(RSC,0)
					diagFilas = diagFilas & rsc("codigo") & "/" & rsc("tipo") & "/" & rsc("solapa") & ";"
					rsc.MOVENEXT
					cantFilas = cantFilas + 1
				Loop
			Loop
			' [CUENTA] Total = documentos recorridos (saltados por la pagina, dibujados y resto) que cumplen lo mismo que exige la tarjeta
			cantTotDoc = ContarValidos()
			totTxt = cantTotDoc
			If cantFilas >= 200 Then totTxt = "m&aacute;s de " & cantTotDoc   ' 200 = findd.cantreg: la libreria trae como maximo esas filas

			'sul="<u>" &  sli & "</u>"
			If sli <> "" Then  sul=sul & "<div class='swiper-slide'>"& "<ul>" & sli & "</ul>" & "</div>"
			shtmldocs=" <div id=swiperbsp xclass='swipere mySwiper2' data-cant='" & totTxt & "'>"
			shtmldocs= shtmldocs & " <div xclass='swiper-wrapper'>"
			shtmldocs= shtmldocs & tarjeta & "</div>"
			shtmldocs= shtmldocs & "    <div class='swiper-pagination'></div>"
			shtmldocs= shtmldocs & " </div>"
			'sul=sul & tarjeta
			' [PAGINADO] Links de pagina 1-3 (moverRegistros) dentro de los resultados, con estilo propio
			If cantdoc > 0 Or CantRegMoverx > 0 Then
				pagHtml = "<div class='em-pag' style='margin:8px 0;font-size:13px;'>P&aacute;ginas: "
				For ipg = 1 To 3
					If CLng(CantRegMoverx) = ipg - 1 Then
						pagHtml = pagHtml & "<span data-pos='" & (ipg - 1) & "' style='display:inline-block;min-width:24px;margin:0 2px;padding:2px 6px;text-align:center;border:1px solid #666;font-weight:bold;background:#e8e8e8;'>" & ipg & "</span>"
					Else
						pagHtml = pagHtml & "<a href='#' data-pos='" & (ipg - 1) & "' onclick='moverRegistros(this);return false;' style='display:inline-block;min-width:24px;margin:0 2px;padding:2px 6px;text-align:center;border:1px solid #bbb;text-decoration:none;cursor:pointer;'>" & ipg & "</a>"
					End If
				Next
				pagHtml = pagHtml & "</div>"
				shtmldocs = shtmldocs & pagHtml
			End If
			' [LIMITE] Se muestran limitpag documentos; si hay mas, se informa el total (con "mas de" si se alcanzo el techo de busqueda)
			totTxt = cantTotDoc
			If cantFilas >= 200 Then totTxt = "m&aacute;s de " & cantTotDoc




			'  <!-- ?? NUEVO: DIV contenedor de Resultados -->
			shtmldocsF=  shtmldocsF &   " <main>"
			'shtmldocsF=  shtmldocsF &   "	      <h3 class='em-section-title'>Filtros seleccionados</h3>"
			'shtmldocsF=  shtmldocsF &   "      <div id='chipsPanel' class='em-chips' aria-live='polite'>"
			'shtmldocsF=  shtmldocsF &   "	        <span class='em-chips-label'>Filtros:</span>"
			'shtmldocsF=  shtmldocsF &   "        <button id='btnClearAll' class='em-clear-all' type='button'>Quitar todos</button>"
			'shtmldocsF=  shtmldocsF &   " <form target=info id=sx action='searchex.asp?verDocs=S&ing=&FORMULARIOS=D&YIDER="&search_phrase&"' method=post><button class='em-filtro-all' id='ver'>Aplicar Filtros</button><input type=hidden name=filtros id=filtros value=''><form>"
			'shtmldocsF=  shtmldocsF &   "	      </div>"

			shtmldocsF=  shtmldocsF &   "	      <div id='emResults' class='em-results' aria-live='polite'>"
			shtmldocsF=  shtmldocsF &   "	        <div class='em-results-head' id=rst name=rst >"
			'shtmldocsF=  shtmldocsF &   "	          <h4 id=rst name=rst class='em-results-title'><div>" & ahtmldocs & "</div></h4>"
			shtmldocsF=  shtmldocsF & shtmldocs
			shtmldocsF=  shtmldocsF &   "	        </div>"
			shtmldocsF=  shtmldocsF &   "	        <div class='em-results-empty' id='emResultsEmpty los filtros para ver resultados.</div>"
			shtmldocsF=  shtmldocsF &   "	        <div class='em-results-grid' id='emResultsGrid' hidden></div>"
			shtmldocsF=  shtmldocsF &   "	      </div>"
			shtmldocsF=  shtmldocsF &   "	    </main>"
			shtmldocsF=  shtmldocsF &   "	</div>"

			shtml=shtml & shtmldocsF  ' &  sdataDoc

            end if

		End Select

	End If


	shtml = Replace(shtml, "@@TEMA_FACET@@", ArmaFacetTema())   ' [TEMA] facetas segun los documentos encontrados
	shtml = Replace(shtml, "@@TIPO_FACET@@", ArmaFacetTipo())   ' [ESQUEMA]
	shtml = Replace(shtml, "@@AUTOR_FACET@@", ArmaFacetAutor())   ' [AUTOR]
	If verDocs="S" Then
	 if cantdoc > 0 then
		response.write shtmldocs '"<html><body onload=""window.top.getIFD('info').getElementById('rst').innerHTML=document.body.innerHTML ; window.top.getIFD('info').getElementById('resultsCount').innerHTML='"& cantdoc &"';"">" & shtmldocs & "</body></html>"
		else
		response.write shtmldocs '"<html><body onload=""window.top.getIFD('info').getElementById('rst').innerHTML=document.body.innerHTML ; window.top.getIFD('info').getElementById('resultsCount').innerHTML='"& cantdoc &"';"">" & shtmldocs & "</body></html>"

		end if
	Else
	'loader-wrapper
	     if cantdoc > 0 then
		   shtml=shtml & "<script>document.getElementById('resultsCount').innerHTML='"& totTxt &"';"
		else
		    shtml=shtml & "<script>document.getElementById('resultsCount').innerHTML=' "& cantdoc &"';"
		end if
		shtml=shtml & "document.getElementById('loader-wrapper').style.display = 'none';"
		shtml=shtml & "</script>"
		'paginado
		sul = ""   ' [PAGINADO] los links ahora van dentro de los resultados

         sce="<script> function moverRegistros(elemento) {"

		  sce=sce & "  const valorPos = elemento.getAttribute('data-pos');"


		  sce=sce & "  const form = document.getElementById('sx'); "


		 sce=sce & "    const url = new URL(form.action, window.location.origin); "
		 sce=sce & "   url.searchParams.set('cantregmover', valorPos);"


		 sce= sce & "   form.action = url.pathname + url.search;"
		 sce=sce & "   window.__pgSolo=1;document.getElementById('ver').click();"
         sce=sce & "}</script>"
         ' fin del paginado
		response.write shtml   & sce &  ofv.botoneraESSX("0","","","","","","","","","no","","","")  &sul  & "<iframe WIDTH=600PX HEIGHT=400PX  style='display:none' id=enviaf name=enviaf></iframe></body></html>"
	End If

	' [ERRORES] limpieza protegida: si no hubo busqueda (rsc / findd sin objeto) no debe fallar
	On Error Resume Next
	rsc.close()
	Set rsc =nothing
	call ofv.cerrarconn(cnbsq)
	set findd=nothing

	call ofv.cerrarconn(CN)
	Err.Clear
	On Error GoTo 0
End If



Function stripAccent(Text )

    Const AccChars = "äéöûü¿¡¬√ƒ≈«»… ÀÃÕŒœ–—“”‘’÷Ÿ⁄€‹›‡·‚„‰ÂÁËÈÍÎÏÌÓÔÒÚÛÙıˆ˘˙˚¸˝ˇ"
    Const RegChars = "SZszYAAAAAACEEEEIIIIDNOOOOOUUUUYaaaaaaceeeeiiiidnooooouuuuyy"

    Dim A
    Dim B
    Dim i

    For i = 1 To Len(AccChars)
        A = Mid(AccChars, i, 1)
        B = Mid(RegChars, i, 1)
        Text = Replace(Text, A, B)
    Next

    stripAccent = Text

End Function


' [LIMITE] Configura el componente de busqueda (antes estaba en linea). filtrosX = filtros de facetas (AND ...)
Function CfgBusqueda(filtrosX)
	Dim findd, VALOR1, VALOR2, VALOR3   ' cnbsq es global (lo usa el codigo principal)
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
	findd.cantreg = 200

	'---------------------------------------------------------------------------
	'13/11/2026 -> ALE
	' nuevo pondera la solapa 15 sumando 10000
	IF session("APLMDCLI")="BC" or true THEN
	 findd.solapa=15
	 findd.pondera=10000
	 findd.ModoBsq="B" ' B agrega * al final de cada palabra ,cualquier otro valor es sin * al final de cada palabra ES PARA BUSQUEDA EXACTA
	 If Request.Cookies("exacta") = "S" Then findd.ModoBsq="X"   ' [EXACTA] casilla "Frase exacta" marcada
	END IF

	'---------------------------------------------------------------------------
	findd.listacar="{}[]!#$%&/()=?°'ø|*+®¥:.;,<>-" ' sirve para fulltext con palabras con caracteres especiales ejemplo 123456789-0 guion es un caracter especial.le sacamos el * para buscar exacto


	' filtro si es formularios o documentos
	'

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
		findd.ADDINNER =findd.ADDINNER & "from containstable(ObjMetadata, metadata, '" & VALOR1	 & filtrosX	 & "') as ct "

		findd.ADDINNER =findd.ADDINNER & "join ObjMetadata as y on ct.[key]=y.[objid] ) z "
		findd.ADDINNER =findd.ADDINNER & " ON Z.CODOBJETO=F.CODOBJETO and Z.TIPOBJETO=F.TIPOBJETO"

		findd.ADDINNER= " inner join (select  codobjeto,tipobjeto,metadata from containstable(ObjMetadata, metadata, '"& VALOR1  & filtrosX  & "') as ct2 "
		findd.ADDINNER= findd.ADDINNER & " inner join ObjMetadata as y on ct2.[key]=y.[objid] "
		findd.ADDINNER= findd.ADDINNER & " ) as ct3 on ct3.codobjeto=t1.codigo and ct3.tipobjeto=t1.tipo "
	End If

	Set CfgBusqueda = findd
End Function

' [LIMITE] Junta los codigos de documento encontrados (para el conteo por tema)
Sub ConteoAgrega(c, t)
	If IsNumeric(c & "") And Len(codTemaLista) < 12000 Then codTemaLista = codTemaLista & "," & CLng(c)
	If Not docTipo.Exists(CStr(c)) Then docTipo(CStr(c)) = Trim(t & "")
End Sub

' [TEMA] Facet TEMA con la cantidad de documentos encontrados por tema (sin filtro de tema). Los temas sin documentos se ocultan.
' Si no se puede contar (hay filtro de tema, no hay resultados o falla la consulta) se muestran todos los temas, como antes.
Function ArmaFacetTema()
	Dim h, rsT, rsC, dic, kT, nT, usarConteo, sqlT
	h = ""
	Set dic = Server.CreateObject("Scripting.Dictionary")
	usarConteo = False
	If TEMAF = "" And Trim(codTemaLista & "") <> "" Then
		On Error Resume Next
		Set rsC = CN.Execute("select temadescri, count(distinct cobj) as cant from CSRELPATHDESCRINV where cobj in (0" & codTemaLista & ") group by temadescri")
		If Err.Number = 0 Then
			Do While Not rsC.EOF
				dic(UCase(Trim(rsC(0) & ""))) = CLng(rsC(1))
				rsC.MoveNext
			Loop
			rsC.Close
		End If
		Err.Clear
		On Error Goto 0
		usarConteo = (dic.Count > 0)
	End If
	sqlT = "SELECT  codtema,DESCRI FROM tema ORDER BY DESCRI"
	Set rsT = ofv.crearconsultaEx(sqlT, cn, 1, parametros)
	Do While Not rsT.EOF
		kT = UCase(Trim(rsT("DESCRI") & ""))
		nT = 0
		If usarConteo Then
			If dic.Exists(kT) Then nT = dic(kT)
		End If
		If (Not usarConteo) Or nT > 0 Then
			h = h & "          <label class='em-chk' title=""" & rsT("DESCRI") & """>"
			h = h & "            <input type='checkbox' name='TEMA' value='" & rsT("DESCRI") & "'"
			h = h & "                   data-item-id='" & rsT("CODTEMA") & "'  data-group-id='TEMA'>"
			' [VISUAL] sin contador: solo se ocultan los que no tienen documentos
			h = h & rsT("DESCRI")
			h = h & "          </label>"
		End If
		rsT.MOVENEXT
	Loop
	Call OFV.cerrarconsulta(rsT)
	Set dic = Nothing
	ArmaFacetTema = h
End Function


' [ESQUEMA] Facet ESQUEMA con la cantidad de documentos encontrados por esquema (tipo de objeto). Los esquemas sin documentos se ocultan.
' Si no se puede contar (hay filtros aplicados o no hay resultados) se muestran todos, como antes.
Function ArmaFacetTipo()
	Dim h, rsT, dicT, kD, nT, usarConteo, tp
	h = ""
	Set dicT = Server.CreateObject("Scripting.Dictionary")
	usarConteo = False
	If FILTROSF = "" And docTipo.Count > 0 Then
		For Each kD In docTipo.Keys
			tp = UCase(Trim(docTipo(kD) & ""))
			dicT(tp) = CLng("0" & dicT(tp)) + 1
		Next
		usarConteo = (dicT.Count > 0)
	End If
	Set rsT = ofv.crearconsultaEx("select tipobjeto,descri from tipobjeto  where aplic in('EM','GLB') AND TIPOCPT='D' order by descri", cn, 1, parametros)
	Do While Not rsT.EOF
		nT = 0
		If usarConteo Then
			If dicT.Exists(UCase(Trim(rsT("TIPOBJETO") & ""))) Then nT = dicT(UCase(Trim(rsT("TIPOBJETO") & "")))
		End If
		If (Not usarConteo) Or nT > 0 Then
			h = h & "          <label class='em-chk' title=""" & rsT("DESCRI") & """>"
			h = h & "            <input type='checkbox' name='TIPO' value='" & rsT("DESCRI") & "'"
			h = h & "                   data-item-id='" & rsT("TIPOBJETO") & "'  data-group-id='TIPO'>"
			' [VISUAL] sin contador: solo se ocultan los que no tienen documentos
			h = h & rsT("DESCRI")
			h = h & "          </label>"
		End If
		rsT.MOVENEXT
	Loop
	Call OFV.cerrarconsulta(rsT)
	Set dicT = Nothing
	ArmaFacetTipo = h
End Function

' [AUTOR] Facet AUTOR con la cantidad de documentos encontrados por autor. Los autores sin documentos se ocultan.
' Si no se puede contar (hay filtros aplicados, no hay resultados o falla la consulta) se muestran todos, como antes.
Function ArmaFacetAutor()
	Dim h, rsA, rsC, dicA, nA, usarConteo, xxx, kA
	h = ""
	Set dicA = Server.CreateObject("Scripting.Dictionary")
	usarConteo = False
	If FILTROSF = "" And Trim(codTemaLista & "") <> "" Then
		On Error Resume Next
		Set rsC = CN.Execute("select autor, count(distinct codobjeto) as cant from objprty where estado=3 and codobjeto in (0" & codTemaLista & ") group by autor")
		If Err.Number = 0 Then
			Do While Not rsC.EOF
				dicA(UCase(Trim(rsC(0) & ""))) = CLng(rsC(1))
				rsC.MoveNext
			Loop
			rsC.Close
		End If
		Err.Clear
		On Error Goto 0
		usarConteo = (dicA.Count > 0)
	End If
	Set rsA = ofv.crearconsultaEx("SELECT distinct AUTOR FROM OBJPRTY  WHERE ESTADO=3 ", cn, 1, parametros)
	Do While Not rsA.EOF
		kA = UCase(Trim(rsA("AUTOR") & ""))
		nA = 0
		If usarConteo Then
			If dicA.Exists(kA) Then nA = dicA(kA)
		End If
		If (Not usarConteo) Or nA > 0 Then
			xxx = getinfouser(rsA("AUTOR"))
			h = h & "          <label class='em-chk' title=""" & SESSION("SGU_NOMBRE") & """>"
			h = h & "            <input type='checkbox' name='autor' value='" & SESSION("SGU_NOMBRE") & "'"
			h = h & "                   data-item-id='" & rsA("AUTOR") & "'>"
			' [VISUAL] sin contador: solo se ocultan los que no tienen documentos
			h = h & SESSION("SGU_NOMBRE")
			h = h & "          </label>"
		End If
		rsA.MOVENEXT
	Loop
	Call OFV.cerrarconsulta(rsA)
	Set dicA = Nothing
	ArmaFacetAutor = h
End Function

' [FILTROS] El valor de un filtro (esquema / tema / autor / sector) va entre comillas dobles dentro de la frase de CONTAINSTABLE, que a su
' vez va entre comillas simples dentro del SQL. Una comilla doble en el valor (hay temas cuyo nombre termina con ") rompe la frase y el
' filtro deja pasar documentos que no corresponden; una comilla simple rompe el SQL. La busqueda de texto completo ignora la puntuacion,
' asi que quitar las dobles no cambia la coincidencia; las simples se duplican.
' [CUENTA] Cuenta los documentos recorridos (docTipo) que tienen tarjeta: misma consulta que valida cada tarjeta (ruta, objeto, tipo, estado).
' Si la consulta falla, devuelve la cantidad de documentos recorridos.
Function ContarValidos()
	Dim lista, k, rsV, n
	lista = ""
	For Each k In docTipo.Keys
		If IsNumeric(k) Then lista = lista & "," & CLng(k)
	Next
	n = docTipo.Count
	If lista <> "" Then
		On Error Resume Next
		Set rsV = CN.Execute("select count(distinct a.cobj) from CSRELPATHDESCRINV a inner join objetos b1 on (b1.codobjeto = a.cobj) inner join TIPOBJETO b11 on (b11.TIPOBJETO = b1.TIPOBJETO) inner join objprty b on (b.codobjeto = a.cobj) inner join estados c on (c.objid = b.estado) where a.cobj in (0" & lista & ")")
		If Err.Number = 0 Then
			If Not rsV.EOF Then n = CLng(rsV(0))
			rsV.Close
		End If
		Err.Clear
		On Error Goto 0
	End If
	ContarValidos = n
End Function

' [ERRORES] Recordset vacio (desconectado) con los campos que usa la pagina: se usa cuando la busqueda falla o no devuelve objeto
Function RsVacio()
	Dim r
	Set r = Server.CreateObject("ADODB.Recordset")
	r.CursorLocation = 3
	r.Fields.Append "codigo", 200, 20
	r.Fields.Append "tipo", 200, 20
	r.Fields.Append "solapa", 3
	r.Fields.Append "index1", 3
	r.Fields.Append "text", 200, 8000
	r.Open
	Set RsVacio = r
End Function

Function FiltroVal(v)
	FiltroVal = Replace(Replace(v & "", Chr(34), ""), "'", "''")
End Function

' [FACETAS] Arma el filtro de facetas (mismo formato que arma FILTROSF al inicio) SIN las selecciones del grupo grpEx.
' grpEx = "" no excluye ningun grupo (criterio vigente: los combos reflejan los documentos encontrados con todos los filtros).
Function ArmaFiltrosSin(dato, grpEx)
	Dim arr, i, cf, g, tF, aF, mF, sF, dT, dA, dM, dS
	tF = " AND (" : aF = " AND (" : mF = " AND (" : sF = " AND ("
	dT = "" : dA = "" : dM = "" : dS = ""
	If Trim(dato & "") <> "" Then
		arr = Split(dato, "####")
		For i = 0 To UBound(arr)
			cf = Split(arr(i), ":")
			If UBound(cf) >= 1 Then
				g = UCase(cf(0) & "")
				If g <> UCase(grpEx & "") Then
					Select Case g
						Case "TIPO"
							tF = tF & dT & Chr(34) & "[[TIPO_" & FiltroVal(cf(1)) & "]]" & Chr(34)
							dT = " OR "
						Case "AUTOR"
							aF = aF & dA & Chr(34) & "[[AUTOR_" & FiltroVal(cf(1)) & "]]" & Chr(34)
							dA = " OR "
						Case "TEMA"
							mF = mF & dM & Chr(34) & "[[TEMA_" & FiltroVal(cf(1)) & "]]" & Chr(34)
							dM = " OR "
						Case "SECTOR"
							sF = sF & dS & Chr(34) & "[[CODSECTOR_" & FiltroVal(cf(1)) & "]]" & Chr(34)
							dS = " AND "
					End Select
				End If
			End If
		Next
	End If
	tF = tF & ")" : aF = aF & ")" : mF = mF & ")" : sF = sF & ")"
	If dT = "" Then tF = ""
	If dA = "" Then aF = ""
	If dM = "" Then mF = ""
	If dS = "" Then sF = ""
	ArmaFiltrosSin = tF & aF & mF & sF
End Function

' [FACETAS] Ejecuta la busqueda con el filtro fx y devuelve un Dictionary codigo de documento -> tipo
Function BuscaDocTipos(fx)
	Dim fb, rsx, d, c
	Set d = Server.CreateObject("Scripting.Dictionary")
	Set fb = CfgBusqueda(fx)
	If fb.IsNoiseWordex(fb.search_phrase) <> 0 Then
		On Error Resume Next
		Call fb.findDocex
		Set rsx = fb.rsext
		okB = (Err.Number = 0 And IsObject(rsx))
		Err.Clear
		On Error GoTo 0
		If okB Then
		Do While RETORNAVALOR(rsx, 0) <> "-1"
			c = rsx("codigo") & ""
			If Not d.Exists(c) Then d(c) = Trim(rsx("tipo") & "")
			rsx.MOVENEXT
		Loop
		End If
	End If
	Set BuscaDocTipos = d
End Function

' [FACETAS] Respuesta de verDocs=F: una linea por grupo con los valores presentes en los documentos encontrados (todos los filtros aplicados).
'   TIPO=<codigos de esquema separados por coma>   TEMA=<descripciones separadas por ||>   AUTOR=<codigos de autor separados por coma>
Sub FacetDisponibles(dato)
	Dim cache, g, fx, d, k, lista, sal, rsq, dv, v
	Set cache = Server.CreateObject("Scripting.Dictionary")
	sal = ""
	For Each g In Array("TIPO", "TEMA", "AUTOR")
		' Se usan TODOS los filtros aplicados (incluido el grupo propio): los combos reflejan solo los documentos encontrados
		fx = ArmaFiltrosSin(dato, "")
		If cache.Exists(fx) Then
			Set d = cache(fx)
		Else
			Set d = BuscaDocTipos(fx)
			cache.Add fx, d
		End If
		lista = ""
		For Each k In d.Keys
			If IsNumeric(k) Then lista = lista & "," & CLng(k)
		Next
		Set dv = Server.CreateObject("Scripting.Dictionary")
		If g = "TIPO" Then
			For Each k In d.Keys
				If Not dv.Exists(d(k)) Then dv.Add d(k), 1
			Next
			sal = sal & "TIPO=" & Join(dv.Keys, ",") & vbLf
		Else
			On Error Resume Next
			If lista <> "" Then
				If g = "TEMA" Then
					Set rsq = CN.Execute("select distinct temadescri from CSRELPATHDESCRINV where cobj in (0" & lista & ")")
				Else
					Set rsq = CN.Execute("select distinct autor from objprty where estado=3 and codobjeto in (0" & lista & ")")
				End If
				If Err.Number = 0 Then
					Do While Not rsq.EOF
						v = Trim(rsq(0) & "")
						If v <> "" And Not dv.Exists(v) Then dv.Add v, 1
						rsq.MoveNext
					Loop
					rsq.Close
				End If
			End If
			Err.Clear
			On Error Goto 0
			If g = "TEMA" Then
				sal = sal & "TEMA=" & Join(dv.Keys, "||") & vbLf
			Else
				sal = sal & "AUTOR=" & Join(dv.Keys, ",") & vbLf
			End If
		End If
		Set dv = Nothing
	Next
	Response.Write sal
End Sub

Function armafiltros()



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

     smbsq=""
    IF  Session("APLMDCLI") <> "BC" THEN
	'      <!-- Facet CHECKBOX con grupos -->
	smbsq= smbsq & "      <section class='em-facet' data-facet='TIPO' aria-expanded='true'>"
	smbsq= smbsq & "        <div class='em-facet-head'>"
	smbsq= smbsq & "          <h4 class='em-facet-title'>ESQUEMA</h4>"
	smbsq= smbsq & "          <button class='em-facet-toggle' type='button'><span class='em-caret'><i class='la la-angle-up em-caret' aria-hidden='true'></i></span></button>"
	smbsq= smbsq & "        </div>"
	smbsq= smbsq & "        <div class='em-facet-body'><div class=contg>"
	'          <!-- Grupo -->

	' [ESQUEMA] Los esquemas (con su cantidad) se arman despues de la busqueda: ArmaFacetTipo()
	smbsq= smbsq & "@@TIPO_FACET@@"



	smbsq= smbsq & "        </div></div>"
	smbsq= smbsq & "      </section>"
    END IF
	'      <!-- Facet CHECKBOX con grupos -->
	smbsq= smbsq & "      <section class='em-facet' data-facet='TEMA' aria-expanded='true'>"
	smbsq= smbsq & "        <div class='em-facet-head'>"
	smbsq= smbsq & "          <h4 class='em-facet-title'>TEMA</h4>"
	smbsq= smbsq & "          <button class='em-facet-toggle' type='button'><span class='em-caret'><i class='la la-angle-up em-caret' aria-hidden='true'></i></span></button>"
	smbsq= smbsq & "        </div>"
	smbsq= smbsq & "        <div class='em-facet-body'><div class=contg>"
	'          <!-- Grupo -->

	' [TEMA] Los temas (con su cantidad) se arman despues de la busqueda: ArmaFacetTema()
	smbsq= smbsq & "@@TEMA_FACET@@"



	smbsq= smbsq & "        </div></div>"
	smbsq= smbsq & "      </section>"



   IF Session("APLMDCLI") <> "BPN" and Session("APLMDCLI") <> "BC" THEN

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
	SQL=SQL & "WHERE eco.eco > 0 AND OBJETOS.OBJUN=1 order by eco.descri"
	'response.write sql
	Set rssmbsq = ofv.crearconsultaEx(SQL,cn,1,parametros)

	Do While Not rssmbsq.EOF
		smbsq= smbsq & "          <label class='em-chk' title=""" & rssmbsq("DESCRI") & """>"
		smbsq= smbsq & "            <input type='checkbox' name='SECTOR' value='" & rssmbsq("DESCRI") & "'"
		smbsq= smbsq & "                   data-item-id='" & rssmbsq("eco") & "'  data-group-id='SECTOR'>"
		smbsq= smbsq &             rssmbsq("DESCRI")
		smbsq= smbsq & "          </label>"

		rssmbsq.MOVENEXT
	Loop
	Call OFV.cerrarconsulta(rssmbsq)



	smbsq= smbsq & "        </div></div>"
	smbsq= smbsq & "      </section>"
  END IF
 if session("APLMDCLI") <> "BP" AND Session("APLMDCLI") <> "BPN" AND Session("APLMDCLI") <> "BC" then

	'      <!-- Facet CHECKBOX con grupos -->
	smbsq= smbsq & "      <section class='em-facet' data-facet='autor' aria-expanded='true'>"
	smbsq= smbsq & "        <div class='em-facet-head'>"
	smbsq= smbsq & "          <h4 class='em-facet-title'>AUTOR</h4>"
	smbsq= smbsq & "          <button class='em-facet-toggle' type='button'><span class='em-caret'><i class='la la-angle-up em-caret' aria-hidden='true'></i></span></button>"
	smbsq= smbsq & "        </div>"
	smbsq= smbsq & "        <div class='em-facet-body'><div class=contg>"
	'          <!-- Grupo -->

	' [AUTOR] Los autores (con su cantidad) se arman despues de la busqueda: ArmaFacetAutor()
	smbsq= smbsq & "@@AUTOR_FACET@@"



	smbsq= smbsq & "        </div></div>"
	smbsq= smbsq & "      </section>"

  end if


	'smbsq= smbsq &  "      <section class='em-facet' wdata-facet='empresa' aria-expanded='true'>"
	smbsq= smbsq & "        <div class='em-facet-head'>"
	smbsq= smbsq & "          <h4 class='em-facet-title'></h4>"

	smbsq=  smbsq &   "<span class=filtro><br><form target=info id=sx action='searchex.asp?cantregmover=0&verDocs=S&ing=&FORMULARIOS=D&YIDER="&search_phrase&"' method=post><button type='button' class='bietx' id='ver'>Aplicar Filtros</button><input type=hidden name=filtros id=filtros value=''><form></span>" 'button class='em-filtro-all'

	smbsq= smbsq & "        </div>"
'	smbsq= smbsq & "      </section>"



	' [FACETAS] Al aplicar filtros, los combos ESQUEMA / TEMA / AUTOR muestran solo los valores presentes en los resultados (solo los valores que aplican a los documentos encontrados; lo tildado se mantiene)
	smbsq= smbsq & "<script>"
	smbsq= smbsq & "(function(){"
	smbsq= smbsq & "var f=document.getElementById('sx');if(!f)return;var ult=null;"
	smbsq= smbsq & "function det(){var h=document.getElementById('filtros');return h?(h.value||''):'';}"
	smbsq= smbsq & "function aplicar(t){"
	smbsq= smbsq & " var g={};var L=t.split('\n');"
	smbsq= smbsq & " for(var i=0;i<L.length;i++){var p=L[i].indexOf('=');if(p>0){g[L[i].substr(0,p).replace(/\s/g,'')]=L[i].substr(p+1).replace(/\r/g,'');}}"
	smbsq= smbsq & " var cfg=[['TIPO','input[type=checkbox][name=TIPO]','id',','],['TEMA','input[type=checkbox][name=TEMA]','val','||'],['AUTOR','input[type=checkbox][name=autor]','id',',']];"
	smbsq= smbsq & " for(var c=0;c<cfg.length;c++){"
	smbsq= smbsq & "  if(g[cfg[c][0]]===undefined)continue;"
	smbsq= smbsq & "  var set={};var arr=g[cfg[c][0]].split(cfg[c][3]);"
	smbsq= smbsq & "  for(var k=0;k<arr.length;k++){if(arr[k]!=='')set[arr[k].toUpperCase()]=1;}"
	smbsq= smbsq & "  var ins=document.querySelectorAll(cfg[c][1]);"
	smbsq= smbsq & "  for(var j=0;j<ins.length;j++){var el=ins[j];var key=(cfg[c][2]==='id'?el.getAttribute('data-item-id'):el.value)||'';var ok=el.checked||set[String(key).toUpperCase()]===1;el.parentNode.style.display=ok?'':'none';}"
	smbsq= smbsq & " }"
	smbsq= smbsq & "}"
	smbsq= smbsq & "function sync(){var v=det();if(v===ult)return;ult=v;var x=new XMLHttpRequest();var u=(f.getAttribute('action')||'').replace('verDocs=S','verDocs=F');"
	smbsq= smbsq & " x.open('POST',u,true);x.setRequestHeader('Content-Type','application/x-www-form-urlencoded');"
	smbsq= smbsq & " x.onreadystatechange=function(){if(x.readyState===4&&x.status===200){try{aplicar(x.responseText);}catch(e){}}};"
	smbsq= smbsq & " x.send('filtros='+encodeURIComponent(v));}"
	smbsq= smbsq & "f.addEventListener('submit',function(){sync();});"
	smbsq= smbsq & "var b=document.getElementById('ver');if(b)b.addEventListener('click',function(){setTimeout(sync,150);});"
	' [PAGINADO] Cada filtro aplicado (o Quitar todos) vuelve a la primera pagina; solo moverRegistros conserva cantregmover
	smbsq= smbsq & "document.addEventListener('click',function(e){var t=e.target;while(t&&t!==document){if(t.id==='ver'||t.id==='btnClearAll'){if(window.__pgSolo){window.__pgSolo=0;}else{f.setAttribute('action',(f.getAttribute('action')||'').replace(/cantregmover=\d+/,'cantregmover=0'));}break;}t=t.parentNode;}},true);"
	smbsq= smbsq & "})();"
	smbsq= smbsq & "</script>"
	smbsq= smbsq & "   </aside>"







	armafiltros=smbsq
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
