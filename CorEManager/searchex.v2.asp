

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
				TIPOF=TIPOF & DELIM &  CHR(34)   & "[[TIPO_" & ColF(1) &"]]" & CHR(34)
				DELIM=" OR "

				Case "AUTOR"
				autorF=autorF & DELIMA &  CHR(34)   & "[[AUTOR_" & ColF(1) &"]]" & CHR(34)
				DELIMa=" OR "

				Case "TEMA"
				TEMAF=TEMAF & DELIMT &  CHR(34)   & "[[TEMA_" & ColF(1) &"]]" & CHR(34)
				DELIMT=" OR "


				Case "SECTOR"
				SECTORF=SECTORF & DELIMS &  CHR(34)   & "[[CODSECTOR_" & ColF(1) &"]]" & CHR(34)
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
	cantFilas = 0
	cantTotDoc = 0
	totTxt = 0
	Set findd = CfgBusqueda(FILTROSF)
	'and AUTOR_ADM and CODSECTOR_132 and CODSECTOR_41


	'response.write  Session("vvv")
	If findd.IsNoiseWordex(findd.search_phrase) = 0 Then

		xinfo = "<h5>Sin Resultados..........</h5>"
		shtml = "<script>"






		shtml = shtml & " </script> "

	Else

		' response.write time() & "  1<br>"
		'Call findd.findDoc
		Call findd.findDocex
		'response.write time() & " 2<br>"
		'response.write findd.search_phrase
		'				RESPONSE.WRITE findd.sqlj
		'Set rsc = findd.orsFx
		Set rsc =  findd.rsext 'cnbsq.execute(findd.sqlj)
		'response.write time() & " 3<br>"


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
						shtmldocsF=  shtmldocsF &   " <form target=info id=sx action='searchex.asp?verDocs=S&ing=&FORMULARIOS=D&YIDER="&search_phrase&"' method=post><button class='bietx' id='ver'>Aplicar Filtros</button><input type=hidden name=filtros id=filtros value=''><form>" 'button class='em-filtro-all'
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
			Do While RETORNAVALOR(RSC,0) <> "-1"
				If cantdoc >= limitpag Then Exit Do   ' [LIMITE] el resto solo se cuenta (mas abajo)
				Dim 	Asolpas
				Dim 	Aadjuntos
				ReDim 	Asolpas(100)

				ix=0

				iddoc=rsc("codigo") & ""
				ConteoAgrega iddoc
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
					 tarjeta=tarjeta & "           <div class='info-row'><span class='label'>Autor:</span><span>" &recordsetip("AUTOR") & "</span></div>"
					 tarjeta=tarjeta & "           <div class='info-row'><span class='label'>CÛdigo:</span><span>" &recordsetip("CODIDEN") & "</span></div>"
					 tarjeta=tarjeta & "           <div class='info-row'><span class='label'>Vigencia:</span><span>" & fechadorcor(CDate(CLng("0" & recordsetip("fecvigen")) + CLng(CDate("01-01-1900"))), "")& "</span></div>"
					 tarjeta=tarjeta & "           <div class='info-row'><span class='label'>VersiÛn:</span><span>" &recordsetip("version") & "</span></div>"
					 tarjeta=tarjeta & "           <div class='info-row'><span class='label'>Vencimiento:</span><span>" & fechadorcor(CDate(CLng("0" & recordsetip("fecvto")) + CLng(CDate("01-01-1900"))), "")& "</span></div>"
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
					textodoc=EliminarAcentos(rsc("text"))
					POSP=INSTR(ucase(textodoc),ucase(palb(0)))
					'response.write posp & "<br>"
					if POSP  > 0 then
					  textodoc=mid(textodoc & "", POSP ,350)
					  for i =0 to ubound(PALB)
					  textodoc=replace(lcase(textodoc),lcase(palb(i)),"<b><i>" & ucase(palb(i)) &"</i></b>")
                      next

					end if

					tarjeta=tarjeta & left(textodoc,150)& "<span class='dots'>...</span><span class='hidden'>"
					tarjeta=tarjeta & mid(textodoc,150,300) & "</span>"
					tarjeta=tarjeta & "        </span>"
					tarjeta=tarjeta & "        <a href='#' class='more'>M·s</a>"
					tarjeta=tarjeta & "      </div>"

					      <!-- Adjuntos (muestra hasta 4; resto al expandir) -->
					tarjeta=tarjeta & "      <div class='attachments' id='atts-1' data-max='4'>"




				End If
				Call ofv.cerrarconsulta(recordsetip)
				cantdoc=cantdoc + 1
				entro=False
				Do While  iddoc = RETORNAVALOR(RSC,0)

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

					rsc.MOVENEXT
					cantFilas = cantFilas + 1
				Loop



                tarjeta=tarjeta & "      </div>"
				tarjeta=tarjeta & "      <span  class='btn-files' data-target='#atts-1' data-more='M·s archivos' data-less='Menos archivos' style='display:none;'>M·s archivos</span>"
				tarjeta=tarjeta & "    </div>"
  				tarjeta=tarjeta & "</article>"
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
			cantTotDoc = cantdoc
			Do While RETORNAVALOR(RSC,0) <> "-1"
				iddocR = rsc("codigo") & ""
				ConteoAgrega iddocR
				cantTotDoc = cantTotDoc + 1
				Do While iddocR = RETORNAVALOR(RSC,0)
					rsc.MOVENEXT
					cantFilas = cantFilas + 1
				Loop
			Loop

			'sul="<u>" &  sli & "</u>"
			If sli <> "" Then  sul=sul & "<div class='swiper-slide'>"& "<ul>" & sli & "</ul>" & "</div>"
			shtmldocs=" <div id=swiperbsp xclass='swipere mySwiper2' data-cant='" & cantdoc & "'>"
			shtmldocs= shtmldocs & " <div xclass='swiper-wrapper'>"
			shtmldocs= shtmldocs & tarjeta & "</div>"
			shtmldocs= shtmldocs & "    <div class='swiper-pagination'></div>"
			shtmldocs= shtmldocs & " </div>"
			'sul=sul & tarjeta
			' [LIMITE] Se muestran limitpag documentos; si hay mas, se informa el total (con "mas de" si se alcanzo el techo de busqueda)
			totTxt = cantTotDoc
			If cantFilas >= techoFilas Then totTxt = "m&aacute;s de " & cantTotDoc
			If cantTotDoc > cantdoc Then shtmldocs = "<div class='em-limite' style='font-size:13px;margin:6px 0;'>Mostrando <b>" & cantdoc & "</b> de <b>" & totTxt & "</b> documentos. Use los filtros o afine la b&uacute;squeda para ver el resto.</div>" & shtmldocs




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



	shtml = Replace(shtml, "@@TEMA_FACET@@", ArmaFacetTema())   ' [TEMA] cantidades por tema
	If verDocs="S" Then
		response.write shtmldocs '"<html><body onload=""window.top.getIFD('info').getElementById('rst').innerHTML=document.body.innerHTML ; window.top.getIFD('info').getElementById('resultsCount').innerHTML='"& cantdoc &"';"">" & shtmldocs & "</body></html>"
	Else
		shtml=shtml & "<script>document.getElementById('resultsCount').innerHTML='"& totTxt &"'</script>"
		response.write shtml    &  ofv.botoneraESSX("0","","","","","","","","","no","","","")   & "<iframe WIDTH=600PX HEIGHT=400PX  style='display:none' id=enviaf name=enviaf></iframe></body></html>"
	End If

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
	Dim findd, cnbsq, VALOR1, VALOR2, VALOR3
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
	findd.cantreg = techoFilas  ' [LIMITE] techo de filas (antes 150)
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
Sub ConteoAgrega(c)
	If IsNumeric(c & "") And Len(codTemaLista) < 12000 Then codTemaLista = codTemaLista & "," & CLng(c)
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
			If usarConteo Then h = h & "(" & nT & ") "
			h = h & rsT("DESCRI")
			h = h & "          </label>"
		End If
		rsT.MOVENEXT
	Loop
	Call OFV.cerrarconsulta(rsT)
	Set dic = Nothing
	ArmaFacetTema = h
End Function


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
	'      <!-- Facet CHECKBOX con grupos -->
	smbsq= smbsq & "      <section class='em-facet' data-facet='TIPO' aria-expanded='true'>"
	smbsq= smbsq & "        <div class='em-facet-head'>"
	smbsq= smbsq & "          <h4 class='em-facet-title'>ESQUEMA</h4>"
	smbsq= smbsq & "          <button class='em-facet-toggle' type='button'><span class='em-caret'><i class='la la-angle-up em-caret' aria-hidden='true'></i></span></button>"
	smbsq= smbsq & "        </div>"
	smbsq= smbsq & "        <div class='em-facet-body'><div class=contg>"
	'          <!-- Grupo -->

	SQL="select tipobjeto,descri from tipobjeto order by descri"
	Set rssmbsq = ofv.crearconsultaEx(SQL,cn,1,parametros)

	Do While Not rssmbsq.EOF
		smbsq= smbsq & "          <label class='em-chk' title=""" & rssmbsq("DESCRI") & """>"
		smbsq= smbsq & "            <input type='checkbox' name='TIPO' value='" & rssmbsq("DESCRI") & "'"
		smbsq= smbsq & "                   data-item-id='" & rssmbsq("TIPOBJETO") & "'  data-group-id='TIPO'>"
		smbsq= smbsq &             rssmbsq("DESCRI")
		smbsq= smbsq & "          </label>"

		rssmbsq.MOVENEXT
	Loop
	Call OFV.cerrarconsulta(rssmbsq)



	smbsq= smbsq & "        </div></div>"
	smbsq= smbsq & "      </section>"

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
	SQL=SQL & "WHERE eco.eco > 0 AND OBJETOS.OBJUN=1 order by descri"
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



	'      <!-- Facet CHECKBOX con grupos -->
	smbsq= smbsq & "      <section class='em-facet' data-facet='autor' aria-expanded='true'>"
	smbsq= smbsq & "        <div class='em-facet-head'>"
	smbsq= smbsq & "          <h4 class='em-facet-title'>AUTOR</h4>"
	smbsq= smbsq & "          <button class='em-facet-toggle' type='button'><span class='em-caret'><i class='la la-angle-up em-caret' aria-hidden='true'></i></span></button>"
	smbsq= smbsq & "        </div>"
	smbsq= smbsq & "        <div class='em-facet-body'><div class=contg>"
	'          <!-- Grupo -->

	SQL="SELECT distinct AUTOR FROM OBJPRTY  WHERE ESTADO=3 "
	Set rssmbsq = ofv.crearconsultaEx(SQL,cn,1,parametros)

	Do While Not rssmbsq.EOF
		xxx=getinfouser(rssmbsq("AUTOR"))
		smbsq= smbsq & "          <label class='em-chk' title=""" & SESSION("SGU_NOMBRE") & """>"
		smbsq= smbsq & "            <input type='checkbox' name='autor' value='" & SESSION("SGU_NOMBRE") & "'"
		smbsq= smbsq & "                   data-item-id='" & rssmbsq("AUTOR") & "'>"
		smbsq= smbsq &             SESSION("SGU_NOMBRE")
		smbsq= smbsq & "          </label>"

		rssmbsq.MOVENEXT
	Loop
	Call OFV.cerrarconsulta(rssmbsq)



	smbsq= smbsq & "        </div></div>"
	smbsq= smbsq & "      </section>"




	'smbsq= smbsq &  "      <section class='em-facet' wdata-facet='empresa' aria-expanded='true'>"
	smbsq= smbsq & "        <div class='em-facet-head'>"
	smbsq= smbsq & "          <h4 class='em-facet-title'></h4>"

	smbsq=  smbsq &   "<span class=filtro><br><form target=info id=sx action='searchex.asp?verDocs=S&ing=&FORMULARIOS=D&YIDER="&search_phrase&"' method=post><button class='bietx' id='ver'>Aplicar Filtros</button><input type=hidden name=filtros id=filtros value=''><form></span>" 'button class='em-filtro-all'

	smbsq= smbsq & "        </div>"
'	smbsq= smbsq & "      </section>"



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
