<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/datosDetalles.inc" -->
<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<%
' [CG] Control de usuario E-Manager (igual que CorEManager/Acciones.asp)
cgOk = ofv.CheckUsuario()
If cgOk <> "Y" Then
	cgUv = ofv.uv
	If cgUv = "NO" Then cgOk = ofv.ValidarUsuario
	Response.Write cgOk
	Response.End
End If


html = headhtml()
html = html & "<body>"

'fini = clng(Date())-clng(CDate("01-01-1900"))
'response.write(fini)
'response.write(GTFechadorsh(fini))
Set Cnn = Server.CreateObject("ADODB.Connection")
On Error Resume Next
Cnn.Open scon1
If Err.Number <> 0 Then
	Response.Write "<div style='color:red;font-family:monospace;padding:10px;border:1px solid red;margin:10px;'>"
	Response.Write "<b>ERROR DE CONEXION A BASE DE DATOS</b><br>"
	Response.Write "Codigo: " & Err.Number & "<br>"
	Response.Write "Detalle: " & Server.HTMLEncode(Err.Description) & "</div>"
	Response.End
End If
On Error Goto 0
'response.write request.querystring



id = request.querystring("id")
opc = request.querystring("opc")
sol = request.querystring("sol")
tipo = request.querystring("tipo")
idproy = request.querystring("idproy")
vue = request.querystring("vue")
cons = request.querystring("cons")
xtipo = request.querystring("xtipo")
cli = request.querystring("cli")
relainc = request.querystring("relainc")
estado = request.querystring("estado")
aplic = request.querystring("aplic")
xmenu = request.querystring("xmenu")
xtmenu = request.querystring("xtmenu")

titP = getName(cnn,"titulo","proyectos",id)
tit = getName(cnn,"titulo","incidentes",id)

Dim st0,st1,st3
'st0 = "font-family:verdana;font-size:7.5pt;padding-left:5px;background-color:#ececec;border-radius:5px;border-style:solid;border-color:fff;height:27.5px;border-width:2px;"
'st1 = "font-family:arial;font-size:7.75pt;padding-left:5px;letter-spacing:0.25px;"
'st3 = "font-family:arial;font-size:8pt;border-style:solid;border-radius:2.5px;border-color:#ececec;width:90%;height:25px;letter-spacing:0.5px;"

nFecini = 0

If nFecini < CLng(CDate("01-01-1900")) Then
	nFecini = CLng(CDate("01-01-1900"))
Else
	nFecini = 0
End If

' [LOG] Consulta de un incidente / proyecto (CG_LOG)
If opc = "veri" Then CgLog "INCIDENTE", request.querystring("id"), "CONSULTA"
If opc = "verp" Then CgLog "PROYECTO", request.querystring("id"), "CONSULTA"

If ( opc = "acti" Or opc = "actp" Or opc = "veri" Or opc = "verp" ) Then
    html = html & "<div>"
	html = html & cabecera()
	'html = html & "<div style='height:8%;width:80%;border-top:solid 5px #FF3D3D;border-bottom:double 6px #001BFF;margin: 0 auto;'>"
	'[CG-ANT] html = html & "<div class='titulo-seccion page-toolbar'>"
	' [CG] La barra superior necesita cliente, prioridad y fechas, que se leen mas abajo.
	' [CG] Por eso su fila de titulo y botones se guarda en cgTopFila y se inserta con CgTopbar().
	cgHtmlPrevio = html ' [CG]
	html = "" ' [CG]
	cgTopEmitido = False ' [CG]
	cgTopEdit = False ' [CG] True en las pantallas de edicion
	'[CG-ANT] html = html & "<div class='cg-app cg-topbar'><div class='cg-topbar__inner'><div class='cg-topbar__top'>"
	html = html & "<div class='ui-topbar__top'>"

	If opc = "veri" Then

		'[CG-ANT] html = html & "<div><h2>'" & Server.HTMLEncode(LEFT(tit,40)) & "'</h2></div>"
		'[CG-ANT] html = html & "<div class='cg-topbar__ident'><h2 class='cg-topbar__title'>'" & Server.HTMLEncode(LEFT(tit,40)) & "'</h2></div>"
		html = html & "<div class='ui-topbar__ident'><span class='ui-topbar__id'>INC-" & id & "</span><h2 class='ui-topbar__title'>" & Server.HTMLEncode(tit & "") & "</h2></div>"

	ElseIf opc = "acti" Then

		'[CG-ANT] html = html & "<div><h2>'" & Server.HTMLEncode(LEFT(tit,40)) & "'</h2></div>"
		'[CG-ANT] html = html & "<div class='cg-topbar__ident'><h2 class='cg-topbar__title'>'" & Server.HTMLEncode(LEFT(tit,40)) & "'</h2></div>"
		html = html & "<div class='ui-topbar__ident'><span class='ui-topbar__id'>INC-" & id & "</span><!--CG_TITULO--></div>" ' [CG] el titulo editable se inserta al leer el registro
		'cliente = request.querystring("cli")
		cliente = request.form("cli")
		form5 = "<form id='form5' xstyle='height:70%;' method='post' action='incidentes.asp?opc=listar&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
		'[CG-ANT] boton5 = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('form5').submit();"">"
		boton5 = "<input class='ui-button ui-button--secondary ui-button--sm' type='button' value='Volver' onclick=""document.getElementById('form5').submit();"">"
		'[CG-ANT] html = html & "<div class='page-toolbar-action'>" & boton5 & "</div>"
		html = html & "<div class='ui-topbar__actions'>" & boton5 & "</div>"

	ElseIf opc = "actp" Then

		If tipo = "cons" Then

			'[CG-ANT] html = html & "<div><h2>'" & Server.HTMLEncode(LEFT(titP,40)) & "'</h2></div>"
			'[CG-ANT] html = html & "<div class='cg-topbar__ident'><h2 class='cg-topbar__title'>'" & Server.HTMLEncode(LEFT(titP,40)) & "'</h2></div>"
			html = html & "<div class='ui-topbar__ident'><span class='ui-topbar__id'>PRY-" & id & "</span><!--CG_TITULO--></div>" ' [CG] el titulo editable se inserta al leer el registro

		Else

			sol = request.querystring("sol")
			'[CG-ANT] html = html & "<div><h2>" & Server.HTMLEncode(LEFT(titP,40)) & "</h2></div>"
			'[CG-ANT] html = html & "<div class='cg-topbar__ident'><h2 class='cg-topbar__title'>" & Server.HTMLEncode(LEFT(titP,40)) & "</h2></div>"
			html = html & "<div class='ui-topbar__ident'><span class='ui-topbar__id'>PRY-" & id & "</span><!--CG_TITULO--></div>" ' [CG] el titulo editable se inserta al leer el registro

		End If

		'cliente = request.querystring("cli")
		cliente = request.form("cli")
		formz1 = "<form id='formz1' xstyle='height:70%;' method='post' action='proyectos.asp?opc=listar&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
		'[CG-ANT] botonz1 = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('formz1').submit();"">"
		botonz1 = "<input class='ui-button ui-button--secondary ui-button--sm' type='button' value='Volver' onclick=""document.getElementById('formz1').submit();"">"
		'[CG-ANT] html = html & "<div class='page-toolbar-action'>" & botonz1 & "</div>"
		html = html & "<div class='ui-topbar__actions'>" & botonz1 & "</div>"

	ElseIf opc = "verp" Then

		If tipo = "cons" Then

			'[CG-ANT] html = html & "<div><h2>'" & Server.HTMLEncode(LEFT(titP,40)) & "'</h2></div>"
			'[CG-ANT] html = html & "<div class='cg-topbar__ident'><h2 class='cg-topbar__title'>'" & Server.HTMLEncode(LEFT(titP,40)) & "'</h2></div>"
			html = html & "<div class='ui-topbar__ident'><span class='ui-topbar__id'>PRY-" & id & "</span><h2 class='ui-topbar__title'>" & Server.HTMLEncode(titP & "") & "</h2></div>"

		Else

			'[CG-ANT] html = html & "<div><h2>" & Server.HTMLEncode(ucase(LEFT(titP,40))) & "</h2></div>"
			'[CG-ANT] html = html & "<div class='cg-topbar__ident'><h2 class='cg-topbar__title'>" & Server.HTMLEncode(ucase(LEFT(titP,40))) & "</h2></div>"
			html = html & "<div class='ui-topbar__ident'><span class='ui-topbar__id'>PRY-" & id & "</span><h2 class='ui-topbar__title'>" & Server.HTMLEncode(titP & "") & "</h2></div>"

		End If

	End If

	If opc = "veri" Then
		'response.write "Fase2"
		If tipo = "cons" Then

			cli = request.querystring("cli")
			est = request.querystring("est")
			aplic = request.querystring("apli")
			relainc = request.querystring("relainc")
			fec = request.querystring("fec")
			descri = request.querystring("desc")
			form7 = "<form id='form7' xstyle='height:70%;' method='post' action='incidentes.asp?opc=consi&cli=" & cli & "&est=" & est & "&apli=" & aplic & "&fec=" & fec & "&desc=" & descri & "&relainc=" & relainc& "&xmenu=12&xtmenu=INICIO'>"
			'[CG-ANT] boton7 = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('form7').submit();""></form>"
			boton7 = "<input class='ui-button ui-button--secondary ui-button--sm' type='button' value='Volver' onclick=""document.getElementById('form7').submit();""></form>"

			'[CG-ANT] html = html & "<div class='page-toolbar-action'>" & form7 & boton7 & "</div>"
			html = html & "<div class='ui-topbar__actions'>" & form7 & boton7 & "</div>"

		ElseIf tipo = "relac" Then

			form7 = "<form id='form7' xstyle='height:70%;' method='post' action='detalles.asp?opc=verp&id=" & idproy & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
			'[CG-ANT] boton7 = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('form7').submit();""></form>"
			boton7 = "<input class='ui-button ui-button--secondary ui-button--sm' type='button' value='Volver' onclick=""document.getElementById('form7').submit();""></form>"
			'[CG-ANT] html = html & "<div class='page-toolbar-action'>" & form7 & boton7 & "</div>"
			html = html & "<div class='ui-topbar__actions'>" & form7 & boton7 & "</div>"

		ElseIf tipo = "incxp" Then

			form7 = "<form id='form7' xstyle='height:70%;' method='post' action='proyRelac.asp?" & request.querystring & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
			'[CG-ANT] boton7 = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('form7').submit();""></form>"
			boton7 = "<input class='ui-button ui-button--secondary ui-button--sm' type='button' value='Volver' onclick=""document.getElementById('form7').submit();""></form>"
			'[CG-ANT] html = html & "<div class='page-toolbar-action'>" & form7 & boton7 & "</div>"
			html = html & "<div class='ui-topbar__actions'>" & form7 & boton7 & "</div>"

		Else
			cli = request.querystring("cli")
			tit = request.querystring("tit")
			aplic = request.querystring("apli")  'agregado 4/11/2025

			ACTpuede = puedeAct(cnn,id)
			'response.write "Puede Actualizar:" & ACTpuede
			If ACTpuede = "SI" Then
				xform7 = "<form id='xform7' xstyle='height:70%;' method='post' action='detalles.asp?opc=acti&id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
				If getPermisosAtrm("CG_INCIDENTES","MODIF") = 1  Then
					'[CG-ANT] xboton7 = "<input class='bot ui-button' type='button' value='ACTUALIZAR INCIDENTE' onclick=""document.getElementById('xform7').submit();"">"
					'[CG-ANT] xboton7 = "<input class='cg-btn cg-btn--light cg-btn--sm' type='button' value='Actualizar incidente' onclick=""document.getElementById('xform7').submit();"">"
					xboton7 = "<input class='ui-button ui-button--light ui-button--sm' type='button' value='Editar' onclick=""document.getElementById('xform7').submit();"">"
				Else
					'[CG-ANT] xboton7 = "<input class='botsp ui-button' type='button' value='ACTUALIZAR INCIDENTE' onclick=""alert('Sin permiso para modificar.');"">"
					'[CG-ANT] xboton7 = "<input class='cg-btn cg-btn--light cg-btn--sm cg-btn--disabled' type='button' value='Actualizar incidente' onclick=""alert('Sin permiso para modificar.');"">"
					xboton7 = "<input class='ui-button ui-button--light ui-button--sm ui-button--disabled' type='button' value='Editar' onclick=""alert('Sin permiso para modificar.');"">"
				End If
				'[CG-ANT] html = html & "<div class='page-toolbar-action'>" & xboton7 & "</div>"
				html = html & "<div class='ui-topbar__actions'>" & xboton7 & "</div>"
			End If

			form7 = "<form id='form7' xstyle='height:70%;' method='post' action='incidentes.asp?opc=listar&cli=" & cli & "&tit=" & tit & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
			' [CONSULTA] Si se llego desde vistaconsulta.asp (vc=1), Volver regresa a esa consulta con los mismos filtros
			If Request.QueryString("vc") = "1" Then form7 = "<form id='form7' method='post' action='" & DetUrlVueltaConsulta("I") & "'></form>"
			'[CG-ANT] boton7 = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('form7').submit();"">"
			boton7 = "<input class='ui-button ui-button--secondary ui-button--sm' type='button' value='Volver' onclick=""document.getElementById('form7').submit();"">"
			'[CG-ANT] html = html & "<div class='page-toolbar-action'>" & boton7 & "</div>"
			html = html & "<div class='ui-topbar__actions'>" & boton7 & "</div>"

		End If

	ElseIf opc = "verp" Then
	
		'response.write "tipo :" & tipo

		If tipo = "cons" Then

			cli = request.querystring("cli")
			est = request.querystring("est")
			apli = request.querystring("apli")
			fecha = request.querystring("fec")
			descri = request.querystring("desc")
			form8 = "<form id='form8' xstyle='height:70%;' method='post' action='proyectos.asp?sel=1&cli=" & cli & "&est=" & est & "&apli=" & apli & "&fec=" & fecha & "&desc=" & descri & "&opc=consi&xmenu=12&xtmenu=INICIO'>"
			'[CG-ANT] boton8 = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('form8').submit();""></form>"
			boton8 = "<input class='ui-button ui-button--secondary ui-button--sm' type='button' value='Volver' onclick=""document.getElementById('form8').submit();""></form>"
			'[CG-ANT] html = html & "<div class='page-toolbar-action'>" & form8 & boton8 & "</div>"
			html = html & "<div class='ui-topbar__actions'>" & form8 & boton8 & "</div>"
			'html = html & "recordset(2) & recordset (3) "
		Else
			cli = request.querystring("cli")
			tit = request.querystring("tit")
			aplic = request.querystring("apli")

			xform7 = "<form id='xform7' xstyle='height:70%;' method='post' action='detalles.asp?opc=actp&id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"

			If getPermisosAtrm("CG_PROYECTOS","MODIF") = 1 Then
				'[CG-ANT] xboton7 = "<input class='bot ui-button' type='button' value='ACTUALIZAR PROYECTO' onclick=""document.getElementById('xform7').submit();"">"
				'[CG-ANT] xboton7 = "<input class='cg-btn cg-btn--light cg-btn--sm' type='button' value='Actualizar proyecto' onclick=""document.getElementById('xform7').submit();"">"
				xboton7 = "<input class='ui-button ui-button--light ui-button--sm' type='button' value='Editar' onclick=""document.getElementById('xform7').submit();"">"
			Else
				'[CG-ANT] xboton7 = "<input class='botsp ui-button' type='button' value='ACTUALIZAR PROYECTO' onclick=""alert('Sin permiso para modificar.');"">"
				'[CG-ANT] xboton7 = "<input class='cg-btn cg-btn--light cg-btn--sm cg-btn--disabled' type='button' value='Actualizar proyecto' onclick=""alert('Sin permiso para modificar.');"">"
				xboton7 = "<input class='ui-button ui-button--light ui-button--sm ui-button--disabled' type='button' value='Editar' onclick=""alert('Sin permiso para modificar.');"">"
			End If
			'[CG-ANT] html = html & "<div class='page-toolbar-action'>" & xboton7 & "</div>"
			html = html & "<div class='ui-topbar__actions'>" & xboton7 & "</div>"

			form8 = "<form id='form8' xstyle='height:70%;' method='post' action='proyectos.asp?cli=" & cli & "&tit=" & tit & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
			' [CONSULTA] Si se llego desde vistaconsulta.asp (vc=1), Volver regresa a esa consulta con los mismos filtros
			If Request.QueryString("vc") = "1" Then form8 = "<form id='form8' method='post' action='" & DetUrlVueltaConsulta("P") & "'></form>"
			'[CG-ANT] boton8 = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('form8').submit();"">"
			boton8 = "<input class='ui-button ui-button--secondary ui-button--sm' type='button' value='Volver' onclick=""document.getElementById('form8').submit();"">"
		'[CG-ANT] html = html & "<div class='page-toolbar-action'>" & boton8 & "</div>"
		html = html & "<div class='ui-topbar__actions'>" & boton8 & "</div>"

		End If


	End If

	'[CG-ANT] html = html & "</div>"
	'[CG-ANT] html = html & "</div></div></div>"
	html = html & "</div>"
	cgTopFila = html ' [CG] fila de titulo y botones lista para CgTopbar()
	html = cgHtmlPrevio ' [CG]

	html = html & "</div>"

	'[CG-ANT] ' =====================================================================
	'[CG-ANT] ' BLOQUE ANTERIOR COMPLETO (ver y editar incidente y proyecto).
	'[CG-ANT] ' Reemplazado por el bloque [CG] que sigue. Se borra entero buscando [CG-ANT].
	'[CG-ANT] ' =====================================================================
	'[CG-ANT] If opc = "veri" Then

		'[CG-ANT] zzz = ObjAcc(id,"I",cnn,"D")
		'[CG-ANT] XXX = SYSLOGON(cnn,"CONSULTA_INC_" & id)

			'[CG-ANT] Set rs = ObtenerDetalleIncidente(Cnn, id)

		'[CG-ANT] If Not rs.eof Then

			'[CG-ANT] 'html = html & "<div style='height:80%;width:100%;float:left;'>"
			'[CG-ANT] '[CG-ANT] html = html & "<div style='width:100%;'>" ' Limpiamos el float y el height forzado
			'[CG-ANT] html = html & "<div class='cg-app cg-page'>"

			'[CG-ANT] 'html = html & "<div id='tabs' style='border-style: none;'>"
			'[CG-ANT] '[CG-ANT] html = html & "<div id='tabs' style='border-style: none;'>"
			'[CG-ANT] html = html & "<div id='tabs' class='cg-container'>"

			'[CG-ANT] ' --- PESTANAS (Mantenemos la estructura ul/li para no romper el JS, pero inyectamos el CSS moderno) ---
			'[CG-ANT] 'html = html & "<ul>"
			'[CG-ANT] 'html = html & "<li><a href='#fragment-1'>DATOS</a></li>"
			'[CG-ANT] 'html = html & "<li><a href='#fragment-2'>DETALLE</a></li>"
			'[CG-ANT] 'html = html & "<li><a href='#fragment-3'>AVANCE</a></li>"
			'[CG-ANT] 'html = html & "<li><a href='#fragment-4'>APLs RELACIONADOS</a></li>"
			'[CG-ANT] 'html = html & "</ul>"
			'[CG-ANT] ' DEPURADO [CG]: la pantalla nueva no usa pestanas; cada seccion es una tarjeta (cg-card)
			'[CG-ANT] '[CG-ANT] html = html & "<ul class='nav-tabs ui-tablist' style='list-style: none; padding: 0; margin: 0; margin-bottom: 1.5rem;'>"
			'[CG-ANT] '[CG-ANT] html = html & "<li><a class='nav-tab active ui-tablist__link is-active' style='display: block;' href='#fragment-1'>DATOS</a></li>"
			'[CG-ANT] '[CG-ANT] html = html & "<li><a class='nav-tab ui-tablist__link' style='display: block;' href='#fragment-2'>DETALLE</a></li>"
			'[CG-ANT] '[CG-ANT] html = html & "<li><a class='nav-tab ui-tablist__link' style='display: block;' href='#fragment-3'>AVANCE</a></li>"
			'[CG-ANT] '[CG-ANT] html = html & "<li><a class='nav-tab ui-tablist__link' style='display: block;' href='#fragment-4'>APLs RELACIONADOS</a></li>"
			'[CG-ANT] '[CG-ANT] html = html & "</ul>"


			'[CG-ANT] ' ==========================================
			'[CG-ANT] ' PESTANA 1: DATOS
			'[CG-ANT] ' ==========================================
			'[CG-ANT] '[CG-ANT] html = html & "<div id='fragment-1'>"
			'[CG-ANT] html = html & "<div id='fragment-1' class='cg-card'>"
			'[CG-ANT] html = html & "<div class='cg-card__title'>Informaci&oacute;n</div>" ' [CG]

			'[CG-ANT] ' Envolvemos en nuestro bloque moderno y aplicamos la grilla doble
			'[CG-ANT] '[CG-ANT] html = html & "<div class='bloque-datos ui-panel'>"
			'[CG-ANT] html = html & "<div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='bloque-body ui-panel__body grid-datos-doble ui-field-grid--dual'>"
			'[CG-ANT] html = html & "<div class='cg-fields'>"

			'[CG-ANT] 'html = html & "<table class='tablDet' cellspacing=0 cellpadding=3 align=center>"

			'[CG-ANT] ' FILA 1
			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td class='fil'><b>NRO</b></td><td class='fil'>" & rs(0) & "</td>"
			'[CG-ANT] 'html = html & "<td class='fil'><b>TITULO</b></td><td class='fil'>" & rs(22) & "</td>"
			'[CG-ANT] 'html = html & "</tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>NRO</div><div class='dato-value'>" & rs(0) & "</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Nro</div><div class='cg-field__value'>" & rs(0) & "</div></div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>TITULO</div><div class='dato-value'>" & Server.HTMLEncode(rs(22) & "") & "</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>T&iacute;tulo</div><div class='cg-field__value'>" & Server.HTMLEncode(rs(22) & "") & "</div></div>"

			'[CG-ANT] ' FILA 2
			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td class='fil'><b>CLIENTE</b></td><td class='fil'>" & rs(1) & "</td>"
			'[CG-ANT] 'html = html & "<td class='fil'><b>CONTACTO</b></td>"
			'[CG-ANT] 'html = html & "<td class='fil'><a class='red' href='detalleCont.asp?idcon=" & rs(15) & "&id=" & rs(0) & "&opc=veri&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'><font style='color:#FF0000;'>" & rs(2) & "</font></a></td>"
			'[CG-ANT] 'html = html & "</tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>CLIENTE</div><div class='dato-value'>" & Server.HTMLEncode(rs(1) & "") & "</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Cliente</div><div class='cg-field__value'>" & Server.HTMLEncode(rs(1) & "") & "</div></div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>CONTACTO</div><div class='dato-value'><a class='col-link' style='text-decoration: none;' href='detalleCont.asp?idcon=" & rs(15) & "&id=" & rs(0) & "&opc=veri&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>" & Server.HTMLEncode(rs(2) & "") & "</a></div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Contacto</div><div class='cg-field__value'><a class='cg-link' href='detalleCont.asp?idcon=" & rs(15) & "&id=" & rs(0) & "&opc=veri&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>" & Server.HTMLEncode(rs(2) & "") & "</a></div></div>"

			'[CG-ANT] sRelPro = ""
			'[CG-ANT] If rs(14) <> 0 Then sRelPro = "ID: " & rs(14)

			'[CG-ANT] ' FILA 3
			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td class='fil'><b>APLICACION</b></td><td class='fil'>" & rs(3) & "</td>"
			'[CG-ANT] 'html = html & "<td class='fil'><b>RELACIONADO A PROY</b></td><td class='fil'><a class='red' href='detalles.asp?id=" & rs(14) & "&opc=verp&sol=" & rs(25) & "&xmenu=10&xtmenu=INICIO'><font style='color:#FF0000;'>" & sRelPro & "</font></a></td>"
			'[CG-ANT] 'html = html & "</tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>APLICACION</div><div class='dato-value'>" & rs(3) & "</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Aplicaci&oacute;n</div><div class='cg-field__value'>" & rs(3) & "</div></div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>RELACIONADO A PROY</div><div class='dato-value'><a class='col-link' style='text-decoration: none;' href='detalles.asp?id=" & rs(14) & "&opc=verp&sol=" & rs(25) & "&xmenu=10&xtmenu=INICIO'>" & sRelPro & "</a></div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Relacionado a proy</div><div class='cg-field__value'><a class='cg-link' href='detalles.asp?id=" & rs(14) & "&opc=verp&sol=" & rs(25) & "&xmenu=10&xtmenu=INICIO'>" & sRelPro & "</a></div></div>"

			'[CG-ANT] ' FILA 4
			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td class='fil'><b>PRIORIDAD</b></td><td class='fil'>" & rs(5) & "</td>"
			'[CG-ANT] 'html = html & "<td class='fil'><b>RESPONSABLE GESTION</b></td><td class='fil'>" & rs(6) & "</td>"
			'[CG-ANT] 'html = html & "</tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>PRIORIDAD</div><div class='dato-value'>" & rs(5) & "</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Prioridad</div><div class='cg-field__value'>" & rs(5) & "</div></div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>RESPONSABLE GESTION</div><div class='dato-value'>" & rs(6) & "</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Responsable gesti&oacute;n</div><div class='cg-field__value'>" & rs(6) & "</div></div>"

			'[CG-ANT] rp = ""
			'[CG-ANT] If rs(7) & "" = "" Then
				'[CG-ANT] rp = "---"
			'[CG-ANT] Else
				'[CG-ANT] rp = rs(7)
			'[CG-ANT] End If

			'[CG-ANT] ' FILA 5
			'[CG-ANT] 'html = html & "<td class='fil'><b>RESPONSABLE TECNICO</b></td><td class='fil'>" & rp & "</td>"
			'[CG-ANT] 'html = html & "<td class='fil'><b>REQUIERE PRESUPUESTO</b></td><td class='fil'>" & rs(8) & "</td>"
			'[CG-ANT] 'html = html & "</tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>RESPONSABLE TECNICO</div><div class='dato-value'>" & rp & "</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Responsable t&eacute;cnico</div><div class='cg-field__value'>" & rp & "</div></div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>REQUIERE PRESUPUESTO</div><div class='dato-value'>" & rs(8) & "</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Requiere presupuesto</div><div class='cg-field__value'>" & rs(8) & "</div></div>"

			'[CG-ANT] ' FILA 6
			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td class='fil'><b>ESTADO DESARROLLO</b></td><td class='fil'>" & rs(9) & "</td>"
			'[CG-ANT] 'html = html & "<td class='fil'><b>ESTADO FUNCIONAL</b></td><td class='fil'>" & rs(10) & "</td>"
			'[CG-ANT] 'html = html & "</tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>ESTADO DESARROLLO</div><div class='dato-value'>" & rs(9) & "</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Estado desarrollo</div><div class='cg-field__value'>" & rs(9) & "</div></div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>ESTADO FUNCIONAL</div><div class='dato-value'>" & rs(10) & "</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Estado funcional</div><div class='cg-field__value'>" & rs(10) & "</div></div>"

			'[CG-ANT] sFec = Fechador(rs(11) + CLng(CDate("01-01-1900")) )
			
			'[CG-ANT] 'response.write "datos:" & rs(11) & "ini" & nFecini & "fechador:" & sFec& "IsDate:" & IsDate(fec) & "FormatDateTime:"& FormatDateTime(fec, 2)
'[CG-ANT] 'IsDate(fec) Then
        '[CG-ANT] 'Fechador = FormatDateTime(fec, 2)
			'[CG-ANT] ' FILA 7
			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td class='fil'><b>FECHA CREACION</b></td><td class='fil'>" & sFec & "</td>"
			'[CG-ANT] 'html = html & "<td class='fil'><b>BROWSER</b></td><td class='fil'>" & rs(12) & "</td>"
			'[CG-ANT] 'html = html & "</tr>"
			
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>FECHA CREACION</div><div class='dato-value'>" & sFec & "</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Fecha creaci&oacute;n</div><div class='cg-field__value'>" & sFec & "</div></div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>BROWSER</div><div class='dato-value'>" & rs(12) & "</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Browser</div><div class='cg-field__value'>" & rs(12) & "</div></div>"

			'[CG-ANT] od = ""
			'[CG-ANT] If rs(16) & "" = "0" Then
				'[CG-ANT] od = "---"
			'[CG-ANT] Else
				'[CG-ANT] od = Fechador(rs(16) + nFecini)
			'[CG-ANT] End If

			'[CG-ANT] ' FILA 8
			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td class='fil'><b>VIA PEDIDO</b></td><td class='fil'>" & rs(13) & "</td>"
			'[CG-ANT] 'html = html & "<td class='fil'><b>OK DESARROLLO</b></td><td class='fil'>" & od & "</td>"
			'[CG-ANT] 'html = html & "</tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>VIA PEDIDO</div><div class='dato-value'>" & rs(13) & "</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>V&iacute;a pedido</div><div class='cg-field__value'>" & rs(13) & "</div></div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>OK DESARROLLO</div><div class='dato-value'>" & od & "</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>OK desarrollo</div><div class='cg-field__value'>" & od & "</div></div>"

			'[CG-ANT] of = ""
			'[CG-ANT] If rs(17) & "" = "0" Then
				'[CG-ANT] of = "---"
			'[CG-ANT] Else
				'[CG-ANT] of = Fechador(rs(17) + nFecini)
			'[CG-ANT] End If

			'[CG-ANT] ' FILA 9
			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td class='fil'><b>OK FUNCIONAL</b></td><td class='fil'>" & of & "</td>"
			'[CG-ANT] 'html = html & "<td class='fil'><b>OK CLIENTE</b></td><td class='fil'>" & rs(18) & "</td>"
			'[CG-ANT] 'html = html & "</tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>OK FUNCIONAL</div><div class='dato-value'>" & of & "</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>OK funcional</div><div class='cg-field__value'>" & of & "</div></div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>OK CLIENTE</div><div class='dato-value'>" & rs(18) & "</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>OK cliente</div><div class='cg-field__value'>" & rs(18) & "</div></div>"

			'[CG-ANT] ' FILA 10 (Solo tiene un Label-Value, cerramos con dos divs vacNos para mantener la grilla perfecta)
			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td class='fil'><b>LINK DOCUMENTO</b></td>"   & "<td class='fil'>"
			'[CG-ANT] xruta="http://" & replace(rs("ruta") & "","\\","")
			'[CG-ANT] xruta=replace(xruta,"\","/")
			'[CG-ANT] 'html = html & "<a href='" & xruta & "' target='_new'><font style='color:#FF0000;'>" & Left(rs("ruta"),30) & "</font></a>"
			'[CG-ANT] 'html = html & "</td>"
			'[CG-ANT] 'html = html & "<td class='fil'>&nbsp;</td>"
			'[CG-ANT] 'html = html & "<td class='fil'>&nbsp;</td>"
			'[CG-ANT] 'html = html & "</tr>"

			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>LINK DOCUMENTO</div><div class='dato-value' style='word-break: break-all;'>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Link documento</div><div class='cg-field__value cg-field__value--break'>"
			'[CG-ANT] '[CG-ANT] html = html & "<a class='col-link' style='text-decoration: none;' href='" & Server.HTMLEncode(xruta) & "' target='_new'>" & Server.HTMLEncode(Left(rs("ruta") & "",30)) & "</a>"
			'[CG-ANT] html = html & "<a class='cg-link' href='" & Server.HTMLEncode(xruta) & "' target='_new'>" & Server.HTMLEncode(Left(rs("ruta") & "",30)) & "</a>"
			'[CG-ANT] '[CG-ANT] html = html & "</div>"
			'[CG-ANT] html = html & "</div></div>"
			'[CG-ANT] ' DEPURADO [CG]: relleno para alinear la grilla vieja de 4 columnas; cg-fields se acomoda solo
			'[CG-ANT] '[CG-ANT] html = html & "<div></div><div></div>" ' Estos reemplazan a los "&nbsp;" viejos

			'[CG-ANT] 'html = html & "</table>"

			'[CG-ANT] ' Cerramos bloque-body y bloque-datos
			'[CG-ANT] html = html & "</div>"
			'[CG-ANT] html = html & "</div>"

			'[CG-ANT] html = html & "</div>" ' Cierre de fragment-1

			'[CG-ANT] ' ==========================================
			'[CG-ANT] ' PESTANA 2: DETALLES
			'[CG-ANT] ' ==========================================
			'[CG-ANT] '[CG-ANT] html = html & "<div id='fragment-2'>"
			'[CG-ANT] html = html & "<div id='fragment-2' class='cg-card'>"
			'[CG-ANT] html = html & "<div class='cg-card__title'>Detalle</div>" ' [CG]

			'[CG-ANT] 'html = html & "<table class='tabldet' style='border-style: none;' cellspacing=0 align=left>"
			'[CG-ANT] ' Usamos nuestro envase estNndar y flex-column para apilar los elementos
			'[CG-ANT] '[CG-ANT] html = html & "<div class='bloque-datos ui-panel'>"
			'[CG-ANT] html = html & "<div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='bloque-body ui-panel__body' style='display: flex; flex-direction: column; gap: 1.5rem;'>"
			'[CG-ANT] html = html & "<div>"

			'[CG-ANT] ' --- BLOQUE DE TEXTO (DETALLE) ---
			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td class='fil' style='font-size:11pt;'>" & Replace(rs(4) & "",Chr(13),"<br>") & "</td>"
			'[CG-ANT] 'html = html & "</tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div>"
			'[CG-ANT] html = html & "<div class='cg-field'>"
			'[CG-ANT] html = html & "<div class='cg-field__label'>Descripci&oacute;n</div>" ' [CG]
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value' style='line-height: 1.6; color: #334155; font-size: 0.95rem;'>" & Replace(Server.HTMLEncode(rs(4) & ""),Chr(13),"<br>") & "</div>"
			'[CG-ANT] html = html & "<div class='cg-prose'>" & Replace(Server.HTMLEncode(rs(4) & ""),Chr(13),"<br>") & "</div>"
			'[CG-ANT] html = html & "</div>"


			'[CG-ANT] ' --- BLOQUE DE PATH ---
			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td class='fil' style='font-size:12pt;'><BR><u><b>PATH</b></u></td>"
			'[CG-ANT] 'html = html & "</tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div>"
			'[CG-ANT] html = html & "<div class='cg-field'>"
			'[CG-ANT] '[CG-ANT] html = html & "<div style='font-size: 1rem; font-weight: 700; color: #0f172a; margin-bottom: 0.5rem; letter-spacing: 0.05em;'>PATH</div>"
			'[CG-ANT] html = html & "<div class='cg-field__label'>Path</div>"

			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td class='fil'><BR>"
			'[CG-ANT] XPATH="\\" & request.servervariables("SERVER_NAME") &  RIGHT(RUTA,LEN(RUTA)-2)
			'[CG-ANT] 'html = html & XPATH & "INCIDENTES\" & Rs(0)
			'[CG-ANT] 'html = html  & "<BR></td></tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value' style='word-break: break-all; color: #475569;'>"
			'[CG-ANT] html = html & "<div class='cg-path'><span class='cg-path__text'>"
			'[CG-ANT] html = html & XPATH & "INCIDENTES\" & Rs(0)
			'[CG-ANT] '[CG-ANT] html = html & "</div>"
			'[CG-ANT] html = html & "</span></div>"
			'[CG-ANT] html = html & "</div>"


			'[CG-ANT] ' --- BLOQUE DE ARCHIVOS ---
			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td class='fil' style='font-size:12pt;'><BR><u><b>ARCHIVOS</b></u></td>"
			'[CG-ANT] 'html = html & "</tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div>"
			'[CG-ANT] html = html & "<div class='cg-field'>"
			'[CG-ANT] '[CG-ANT] html = html & "<div style='font-size: 1rem; font-weight: 700; color: #0f172a; margin-bottom: 0.5rem; letter-spacing: 0.05em;'>ARCHIVOS</div>"
			'[CG-ANT] html = html & "<div class='cg-field__label'>Archivos</div>"

			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td class='fil'>"
			'[CG-ANT] RUTA=RUTA & "\INCIDENTES\" & Rs(0)

			'[CG-ANT] 'html = html & LISTARDIRTORIO(RUTA,Rs(0),"INCIDENTES")
			'[CG-ANT] 'html = html & "</td></tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>"
			'[CG-ANT] html = html & "<ul class='cg-files'>"
			'[CG-ANT] html = html & LISTARDIRTORIO(RUTA,Rs(0),"INCIDENTES")
			'[CG-ANT] '[CG-ANT] html = html & "</div>"
			'[CG-ANT] html = html & "</ul>"
			'[CG-ANT] html = html & "</div>"

			'[CG-ANT] 'html = html & "</table>"

			'[CG-ANT] ' Cerramos el bloque-body y bloque-datos
			'[CG-ANT] html = html & "</div>"
			'[CG-ANT] html = html & "</div>"

			'[CG-ANT] html = html & "</div>" ' Cierre de fragment-2

			'[CG-ANT] ' ==========================================
			'[CG-ANT] ' PESTANA 3: AVANCES
			'[CG-ANT] ' ==========================================
			'[CG-ANT] '[CG-ANT] html = html & "<div id='fragment-3'>"
			'[CG-ANT] html = html & "<div id='fragment-3' class='cg-card'>"
			'[CG-ANT] html = html & "<div class='cg-card__title'>Historial de avances</div>" ' [CG]

			'[CG-ANT] 'response.write(UserLogon())

			'[CG-ANT] ' Envolvemos en nuestra tarjeta moderna
			'[CG-ANT] '[CG-ANT] html = html & "<div class='bloque-datos ui-panel'>"
			'[CG-ANT] html = html & "<div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='bloque-body ui-panel__body'>"
			'[CG-ANT] html = html & "<div>"

			'[CG-ANT] 'html = html & "<table class='tablDet' style='border-style: none;' cellspacing=0 align=left>"

			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td class='fil' style='font-size:11pt;'>" & Replace(rs(21) & "",Chr(13),"<br>") & "</td>"

			'[CG-ANT] ' Texto de avances con buen interlineado y color de lectura cNmodo
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value' style='line-height: 1.6; color: #334155; font-size: 0.95rem;'>"
			'[CG-ANT] html = html & "<div class='cg-prose'>"
			'[CG-ANT] 'html = html & Replace(Server.HTMLEncode(rs(21) & ""),Chr(13),"<br>")
			'[CG-ANT] html = html & "</div>"
			'[CG-ANT] 'response.write "Avances db rs21"& rs(21)
	'[CG-ANT] '		response.write "ID db rs22"& rs(0)
'[CG-ANT] ' rs(14)
		

		'[CG-ANT] ' html = html & "<div style='font-size: 1rem; font-weight: 700; color: #0f172a; margin-bottom: 1rem; letter-spacing: 0.05em;'>AVANCES</div>"
		
		'[CG-ANT] ' '-- Diferencia entre CorGestion Base y 5

		'[CG-ANT] ' '----------------------------CONEXION A COMENTARIOS---------------------------------------------------
				 '[CG-ANT] html = html & HtmlHistorialRelainc(Cnn, rs(0), "inc")
				
				'[CG-ANT] ' html = html & "<textarea class='txtar' id='avances' name='avances' type='text' style='height:200px;width:100%;'>" & rx(27) & "</textarea>"


			'[CG-ANT] ' --- Restos del cNdigo viejo comentados tal cual los pasaste ---
			'[CG-ANT] 'response.write(request.querystring(rs3))
			'[CG-ANT] '		html = html & "</tbody>"
			'[CG-ANT] '		html = html & "</table>"

			'[CG-ANT] 'html = html & "</tr>"

			'[CG-ANT] 'html = html & "</table>"

			'[CG-ANT] ' Cerramos bloque-body y bloque-datos
			'[CG-ANT] html = html & "</div>"
			'[CG-ANT] html = html & "</div>"

			'[CG-ANT] html = html & "</div>" ' Cierre de fragment-3

			'[CG-ANT] ' ==========================================
			'[CG-ANT] ' PESTANA 4: APLs RELACIONADOS
			'[CG-ANT] ' ==========================================
			'[CG-ANT] '[CG-ANT] html = html & "<div id='fragment-4'>"
			'[CG-ANT] html = html & "<div id='fragment-4' class='cg-card'>"
			'[CG-ANT] html = html & "<div class='cg-card__title'>APLs relacionados</div>" ' [CG]

			'[CG-ANT] 'html = html & "<table style='border:solid 1px #aaaaaa;' cellspacing=0 cellpadding=3 align=center>"
			'[CG-ANT] ' Abrimos nuestra tarjeta estNndar
			'[CG-ANT] '[CG-ANT] html = html & "<div class='bloque-datos ui-panel'>"
			'[CG-ANT] html = html & "<div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='bloque-body ui-panel__body'>"
			'[CG-ANT] html = html & "<div>"

			'[CG-ANT] Set rz = ObtenerAplRelac(Cnn, id, "I")

			'[CG-ANT] ' --- TNTULO DE LA SECCINN ---
			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td style='text-align:center'><font style='font-size:20pt;color:#000000;'>APLs</font></td>"
			'[CG-ANT] 'html = html & "</tr>"
			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td> </td>"
			'[CG-ANT] 'html = html & "</tr>"
			'[CG-ANT] ' DEPURADO [CG]: el titulo ahora es el header de la tarjeta (cg-card__title)
			'[CG-ANT] '[CG-ANT] html = html & "<div style='font-size: 1.25rem; font-weight: 700; color: #0f172a; margin-bottom: 1rem; border-bottom: 1px solid #e2e8f0; padding-bottom: 0.5rem;'>APLs Relacionados</div>"

			'[CG-ANT] ' Contenedor flexible para apilar los links como si fuera una lista
			'[CG-ANT] '[CG-ANT] html = html & "<div style='display: flex; flex-direction: column;'>"
			'[CG-ANT] html = html & "<div class='cg-list'>"

			'[CG-ANT] If Not rz.eof Then

				'[CG-ANT] Do While Not rz.eof

					'[CG-ANT] 'html = html & "<tr>"
					'[CG-ANT] 'ref = "<a class='red2' style='font-size:12pt;' href='detalles2.asp?opc=ver&id=" & rz(1) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>" & rz(4) & "</a>"
					'[CG-ANT] 'html = html & "<td align=left>" & ref & "</td>"
					'[CG-ANT] 'html = html & "</tr>"

					'[CG-ANT] ' Aplicamos .col-link y armamos una "fila" div limpia con un borde sutil abajo
					'[CG-ANT] '[CG-ANT] ref = "<a class='col-link' style='text-decoration: none; font-size: 1rem;' href='detalles2.asp?opc=ver&id=" & rz(1) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>" & rz(4) & "</a>"
					'[CG-ANT] ref = "<a class='cg-link' href='detalles2.asp?opc=ver&id=" & rz(1) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>" & rz(4) & "</a>"
					'[CG-ANT] '[CG-ANT] html = html & "<div style='padding: 0.75rem 0; border-bottom: 1px solid #f1f5f9;'>" & ref & "</div>"
					'[CG-ANT] html = html & "<div class='cg-list__item'>" & ref & "</div>"

					'[CG-ANT] rz.movenext
				'[CG-ANT] Loop

			'[CG-ANT] Else

				'[CG-ANT] 'html = html & "<tr>"
				'[CG-ANT] 'html = html & "<td>-----</td>"
				'[CG-ANT] 'html = html & "</tr>"
				'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value' style='color: #94a3b8; font-style: italic; padding: 1rem 0;'>-----</div>"
				'[CG-ANT] html = html & "<div class='cg-empty'>No hay APLs relacionados.</div>"

			'[CG-ANT] End If

			'[CG-ANT] 'html = html & "</table>"
			'[CG-ANT] html = html & "</div>" ' Cierra el contenedor de la lista

			'[CG-ANT] html = html & "</div>" ' Cierra bloque-body
			'[CG-ANT] html = html & "</div>" ' Cierra bloque-datos

			'[CG-ANT] html = html & "</div>" ' Cierra fragment-4

			'[CG-ANT] html = html & "</div>" ' Cierra el div general de los tabs (el float:left original del primer fragmento)

		'[CG-ANT] End If

		'[CG-ANT] rs.close

	'[CG-ANT] ElseIf opc = "verp" Then

		'[CG-ANT] zzz = ObjAcc(id,"P",cnn,"D")
		'[CG-ANT] XXX = SYSLOGON(cnn,"CONSULTA_PROY_" & id)

		'[CG-ANT] Set rs = ObtenerDetalleProyecto(Cnn, id)
					'[CG-ANT] 'response.write(sql)
					

					 '[CG-ANT] 'response.write("OPC = VERP")
		'[CG-ANT] If Not rs.eof Then

			'[CG-ANT] 'html = html & "<div style='height:80%;width:80%;xfloat:left;margin:0 auto;'>"
			'[CG-ANT] '[CG-ANT] html = html & "<div style='width:100%;'>"
			'[CG-ANT] html = html & "<div class='cg-app cg-page'>"

			'[CG-ANT] '[CG-ANT] html = html & "<div id='tabs' style='border-style: none;'>"
			'[CG-ANT] html = html & "<div id='tabs' class='cg-container'>"

			'[CG-ANT] ' DEPURADO [CG]: la pantalla nueva no usa pestanas; cada seccion es una tarjeta (cg-card)
			'[CG-ANT] '[CG-ANT] html = html & "<ul class='nav-tabs ui-tablist' style='list-style: none; padding: 0; margin: 0; margin-bottom: 1.5rem;'>"
			'[CG-ANT] '[CG-ANT] html = html & "<li><a class='nav-tab active ui-tablist__link is-active' style='display: block;' href='#fragment-1'>DATOS</a></li>"
			'[CG-ANT] '[CG-ANT] html = html & "<li><a class='nav-tab ui-tablist__link' style='display: block;' href='#fragment-2'>DETALLE</a></li>"
			'[CG-ANT] '[CG-ANT] html = html & "<li><a class='nav-tab ui-tablist__link' style='display: block;' href='#fragment-3'>AVANCE</a></li>"
			'[CG-ANT] '[CG-ANT] html = html & "<li><a class='nav-tab ui-tablist__link' style='display: block;' href='#fragment-4'>INCIDENTES Y APLs RELACIONADOS</a></li>"
			'[CG-ANT] '[CG-ANT] html = html & "</ul>"

			'[CG-ANT] 'DATOS

			'[CG-ANT] '[CG-ANT] html = html & "<div id='fragment-1'>"
			'[CG-ANT] html = html & "<div id='fragment-1' class='cg-card'>"
			'[CG-ANT] html = html & "<div class='cg-card__title'>Informaci&oacute;n</div>" ' [CG]

			'[CG-ANT] ' --- INICIO NORMALIZACION: CONTENEDOR PRINCIPAL ---
			'[CG-ANT] 'html = html & "<table class='tablDet' cellspacing=0 cellpadding=0 align=center style='border-style:none;'>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='bloque-datos ui-panel'>"
			'[CG-ANT] html = html & "<div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='bloque-body ui-panel__body grid-datos-doble ui-field-grid--dual'>"
			'[CG-ANT] html = html & "<div class='cg-fields'>"

			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td class='fil' style='" & st0 & "' >ID</td><td class='fil' style='" & st1 & "' >" & rs(0) & "</td>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>ID</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>ID</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value' style='font-weight: 600;'>" & rs(0) & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value'>" & rs(0) & "</div></div>"

			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td class='fil' style='" & st0 & "'>TITULO PROYECTO</td><td class='fil' style='" & st1 & "' >" & UCASE(rs(2)) & "</td>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>TITULO PROYECTO</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>T&iacute;tulo proyecto</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>" & Server.HTMLEncode(UCASE(rs(2) & "")) & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value'>" & Server.HTMLEncode(UCASE(rs(2) & "")) & "</div></div>"

			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td class='fil' style='" & st0 & "'>CONTACTO</td>"
			'[CG-ANT] 'html = html & "<td class='fil' style='" & st1 & "' ><a class='red' href='detalleCont.asp?id=" & id & "&opc=" & opc & "&sol=" & sol & "&idcon=" & rs(22) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'><font style='color:#FF0000;'>" & rs(5) & "</font></a></td>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>CONTACTO</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Contacto</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>"
			'[CG-ANT] html = html & "<div class='cg-field__value'>"
			'[CG-ANT] '[CG-ANT] html = html & "<a class='red' href='detalleCont.asp?id=" & id & "&opc=" & opc & "&sol=" & sol & "&idcon=" & rs(22) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "' style='text-decoration: none;'>"
			'[CG-ANT] html = html & "<a class='cg-link' href='detalleCont.asp?id=" & id & "&opc=" & opc & "&sol=" & sol & "&idcon=" & rs(22) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
			'[CG-ANT] '[CG-ANT] html = html & "<span style='color:#FF0000;cursor:pointer;'>" & Server.HTMLEncode(rs(5) & "") & "</span></a>"
			'[CG-ANT] html = html & Server.HTMLEncode(rs(5) & "") & "</a>"
			'[CG-ANT] '[CG-ANT] html = html & "</div>"
			'[CG-ANT] html = html & "</div></div>"

			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td class='fil' style='" & st0 & "'>RESPON GESTION</td><td class='fil' style='" & st1 & "' >" & rs(7) & "</td>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>RESPON GESTION</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Respon gesti&oacute;n</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>" & rs(7) & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value'>" & rs(7) & "</div></div>"

			'[CG-ANT] 'html = html & "<td class='fil' style='" & st0 & "'>SOLICITUD NUMERO</td><td class='fil' style='" & st1 & "' >" & rs(1) & "</td>"
			'[CG-ANT] 'html = html & "</tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>SOLICITUD NUMERO</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Solicitud n&uacute;mero</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>" & rs(1) & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value'>" & rs(1) & "</div></div>"

			'[CG-ANT] 'html = html & "<td class='fil' style='" & st0 & "'>CLIENTE</td><td class='fil' style='" & st1 & "' >" & rs(4) & "</td>"
			'[CG-ANT] 'html = html & "</tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>CLIENTE</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Cliente</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>" & Server.HTMLEncode(rs(4) & "") & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value'>" & Server.HTMLEncode(rs(4) & "") & "</div></div>"

			'[CG-ANT] 'html = html & "<td class='fil' style='" & st0 & "'>APLICACION</td><td class='fil' style='" & st1 & "' >" & rs(6) & "</td>"
			'[CG-ANT] 'html = html & "</tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>APLICACION</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Aplicaci&oacute;n</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>" & rs(6) & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value'>" & rs(6) & "</div></div>"

			'[CG-ANT] 'html = html & "<td class='fil' style='" & st0 & "'>RESPON TEC</td><td class='fil' style='" & st1 & "' >" & rs(8) & "</td>"
			'[CG-ANT] 'html = html & "</tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>RESPON TEC</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Respon t&eacute;cnico</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>" & rs(8) & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value'>" & rs(8) & "</div></div>"

			'[CG-ANT] sFec = ""
			'[CG-ANT] If rs(9) <> 0 Then
				'[CG-ANT] sFec = Fechador(rs(9) + nFecini)
			'[CG-ANT] End If

			'[CG-ANT] sFecEn = ""
			'[CG-ANT] If rs(10) <> 0 Then
				'[CG-ANT] sFecEn = Fechador(rs(10) + nFecini)
			'[CG-ANT] End If
			
			'[CG-ANT] 'response.write "fechador (" & rs(9) & ") (" & rs(10) & ")"

			'[CG-ANT] ' --- SECCION: FILA FECHAS ---
			'[CG-ANT] 'html = html & "<tr>"

			'[CG-ANT] 'html = html & "<td class='fil' style='" & st0 & "'>FECHA CREACION</td><td class='fil' style='" & st1 & "' >" & sFec & "</td>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>FECHA CREACION</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Fecha creaci&oacute;n</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>" & sFec & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value'>" & sFec & "</div></div>"

			'[CG-ANT] 'html = html & "<td class='fil' style='" & st0 & "'>FECHA ENTREGA</td><td class='fil' style='" & st1 & "' >" & sFecEn & "</td>"
			'[CG-ANT] 'html = html & "</tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>FECHA ENTREGA</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Fecha entrega</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>" & sFecEn & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value'>" & sFecEn & "</div></div>"

			'[CG-ANT] ' --- SECCION: FILA ESTADOS ---
			'[CG-ANT] 'html = html & "<tr>"

			'[CG-ANT] 'html = html & "<td class='fil' style='" & st0 & "'>ESTADO FUNCIONAL</td><td class='fil' style='" & st1 & "' >" & rs(12) & "</td>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>ESTADO FUNCIONAL</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Estado funcional</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>" & rs(12) & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value'>" & rs(12) & "</div></div>"

			'[CG-ANT] 'html = html & "<td class='fil' style='" & st0 & "'>ESTADO DESARROLLO</td><td class='fil' style='" & st1 & "' >" & rs(11) & "</td>"
			'[CG-ANT] 'html = html & "</tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>ESTADO DESARROLLO</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Estado desarrollo</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>" & rs(11) & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value'>" & rs(11) & "</div></div>"

			'[CG-ANT] ' --- SECCION: FILA PRIORIDAD Y RELEVAMIENTO ---
			'[CG-ANT] 'html = html & "<tr>"

			'[CG-ANT] 'html = html & "<td class='fil' style='" & st0 & "'>PRIORIDAD</td><td class='fil' style='" & st1 & "' >" & rs(13) & "</td>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>PRIORIDAD</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Prioridad</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>" & rs(13) & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value'>" & rs(13) & "</div></div>"

			'[CG-ANT] 'html = html & "<td class='fil' style='" & st0 & "'>RELEV Y DEFINICION (hs)</td><td class='fil' style='" & st1 & "' >" & rs(14) & "</td>"
			'[CG-ANT] 'html = html & "</tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>RELEV Y DEFINICION (hs)</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Relev y definici&oacute;n (hs)</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>" & rs(14) & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value'>" & rs(14) & "</div></div>"

			'[CG-ANT] ' --- SECCION: FILA DISENO Y DESARROLLO ---
			'[CG-ANT] 'html = html & "<tr>"

			'[CG-ANT] 'html = html & "<td class='fil' style='" & st0 & "'>DISE&Ntilde;O (hs)</td><td class='fil' style='" & st1 & "' >" & rs(15) & "</td>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>DISE&Ntilde;O (hs)</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Dise&ntilde;o (hs)</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>" & rs(15) & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value'>" & rs(15) & "</div></div>"

			'[CG-ANT] 'html = html & "<td class='fil' style='" & st0 & "'>DESARROLLO (hs)</td><td class='fil' style='" & st1 & "' >" & rs(16) & "</td>"
			'[CG-ANT] 'html = html & "</tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>DESARROLLO (hs)</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Desarrollo (hs)</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>" & rs(16) & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value'>" & rs(16) & "</div></div>"

			'[CG-ANT] ' --- SECCION: FILA TESTEO Y TOTAL ---
			'[CG-ANT] 'html = html & "<tr>"

			'[CG-ANT] 'html = html & "<td class='fil' style='" & st0 & "'>TESTEO (hs)</td><td class='fil' style='" & st1 & "' >" & rs(17) & "</td>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>TESTEO (hs)</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Testeo (hs)</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>" & rs(17) & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value'>" & rs(17) & "</div></div>"

			'[CG-ANT] 'html = html & "<td class='fil' style='" & st0 & "'>TOTAL (hs)</td><td class='fil' style='" & st1 & "' >" & rs(21) & "</td>"
			'[CG-ANT] 'html = html & "</tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>TOTAL (hs)</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Total (hs)</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>" & rs(21) & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value cg-hours__total'>" & rs(21) & "</div></div>"

			'[CG-ANT] ' --- SECCION: FILA BROWSER Y VIA PEDIDO ---
			'[CG-ANT] 'html = html & "<tr>"

			'[CG-ANT] 'html = html & "<td class='fil' style='" & st0 & "'>BROWSER</td><td class='fil' style='" & st1 & "' >" & rs(18) & "</td>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>BROWSER</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Browser</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>" & rs(18) & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value'>" & rs(18) & "</div></div>"

			'[CG-ANT] 'html = html & "<td class='fil' style='" & st0 & "'>VIA PEDIDO</td><td class='fil' style='" & st1 & "' >" & rs(19) & "</td>"
			'[CG-ANT] 'html = html & "</tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>VIA PEDIDO</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>V&iacute;a pedido</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>" & rs(19) & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value'>" & rs(19) & "</div></div>"

			'[CG-ANT] od = ""
			'[CG-ANT] If rs(23) & "" = "0" Then
				'[CG-ANT] od = "---"
			'[CG-ANT] Else
				'[CG-ANT] od = Fechador(rs(23) + nFecini)
			'[CG-ANT] End If

			'[CG-ANT] ' --- SECCION: FILA REQ PRESUPUESTO Y OK DESARROLLO ---
			'[CG-ANT] 'html = html & "<tr>"

			'[CG-ANT] 'html = html & "<td class='fil' style='" & st0 & "'>REQ PRESUPUESTO</td><td class='fil' style='" & st1 & "' >" & rs(20) & "</td>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>REQ PRESUPUESTO</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Req presupuesto</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>" & rs(20) & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value'>" & rs(20) & "</div></div>"

			'[CG-ANT] 'html = html & "<td class='fil' style='" & st0 & "'>OK DESARROLLO</td><td class='fil' style='" & st1 & "' >" & od & "</td>"
			'[CG-ANT] 'html = html & "</tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>OK DESARROLLO</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>OK desarrollo</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>" & od & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value'>" & od & "</div></div>"

			'[CG-ANT] of = ""
			'[CG-ANT] If rs(24) & "" = "0" Then
				'[CG-ANT] of = "---"
			'[CG-ANT] Else
				'[CG-ANT] of = Fechador(rs(24) + nFecini)
			'[CG-ANT] End If

			'[CG-ANT] ' --- SECCION: FILA OK FUNCIONAL Y OK CLIENTE ---
			'[CG-ANT] 'html = html & "<tr>"

			'[CG-ANT] 'html = html & "<td class='fil' style='" & st0 & "'>OK FUNCIONAL</td><td class='fil' style='" & st1 & "' >" & of & "</td>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>OK FUNCIONAL</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>OK funcional</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>" & of & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value'>" & of & "</div></div>"

			'[CG-ANT] 'html = html & "<td class='fil' style='" & st0 & "'>OK CLIENTE</td><td class='fil' style='" & st1 & "' >" & rs(25) & "</td>"
			'[CG-ANT] 'html = html & "</tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>OK CLIENTE</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>OK cliente</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>" & rs(25) & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value'>" & rs(25) & "</div></div>"

			'[CG-ANT] ' --- SECCION: FILA LINK DOCUMENTO ---
			'[CG-ANT] 'html = html & "<tr>"

			'[CG-ANT] 'html = html & "<td class='fil' style='" & st0 & "'>LINK DOCUMENTO</td><td class='fil' style='" & st1 & "' >"
			'[CG-ANT] 'html = html & "<a href='" & rs("ruta") & "' target='_new'><font style='color:#FF0000;'>" & Left(rs("ruta"),30) & "</font></a>"
			'[CG-ANT] 'html = html & "</td>"
			'[CG-ANT] 'html = html & "</tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>LINK DOCUMENTO</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Link documento</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value' style='word-break:break-all;'>"
			'[CG-ANT] html = html & "<div class='cg-field__value cg-field__value--break'>"
			'[CG-ANT] '[CG-ANT] html = html & "<a href='" & Server.HTMLEncode(rs("ruta") & "") & "' target='_new' style='text-decoration: none;'>"
			'[CG-ANT] html = html & "<a class='cg-link' href='" & Server.HTMLEncode(rs("ruta") & "") & "' target='_new'>"
			'[CG-ANT] '[CG-ANT] html = html & "<span style='color:#FF0000; cursor:pointer;'>" & Server.HTMLEncode(Left(rs("ruta") & "",30)) & "</span></a>"
			'[CG-ANT] html = html & Server.HTMLEncode(Left(rs("ruta") & "",30)) & "</a>"
			'[CG-ANT] '[CG-ANT] html = html & "</div>"
			'[CG-ANT] html = html & "</div></div>"

			'[CG-ANT] ' --- SECCION: CIERRE DE CONTENEDOR PRINCIPAL ---
			'[CG-ANT] 'html = html & "</table>"
			'[CG-ANT] 'html = html & "</div>"
			'[CG-ANT] html = html & "</div>" ' Fin padding interno (2rem)
			'[CG-ANT] html = html & "</div>" ' Fin .bloque-datos

			'[CG-ANT] 'html = html & "</div>"
			'[CG-ANT] html = html & "</div>" ' Fin #fragment-1

			'[CG-ANT] 'DETALLES
			'[CG-ANT] '[CG-ANT] html = html & "<div id='fragment-2'>"
			'[CG-ANT] html = html & "<div id='fragment-2' class='cg-card'>"
			'[CG-ANT] html = html & "<div class='cg-card__title'>Detalle</div>" ' [CG]

			'[CG-ANT] ' --- INICIO NORMALIZACION: CONTENEDOR PRINCIPAL DETALLES ---
			'[CG-ANT] 'html = html & "<table cellspacing=0 align=left>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='bloque-datos ui-panel'>"
			'[CG-ANT] html = html & "<div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='bloque-body ui-panel__body' style='display: flex; flex-direction: column; gap: 1.5rem;'>"
			'[CG-ANT] html = html & "<div>"

			'[CG-ANT] ' --- SECCION: TEXTO DESCRIPTIVO ---
			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td class='fil' style='font-size:9.25pt;font-family:arial;letter-spacing:0.25px;'>" & Replace(rs(3) & "",Chr(13),"<br>") & "</td>"
			'[CG-ANT] '[CG-ANT] html = html & "<div>"
			'[CG-ANT] html = html & "<div class='cg-field'>"
			'[CG-ANT] html = html & "<div class='cg-field__label'>Descripci&oacute;n</div>" ' [CG]
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value' style='width: 100%;'>" & Replace(Server.HTMLEncode(rs(3) & ""),Chr(13),"<br>") & "</div>"
			'[CG-ANT] html = html & "<div class='cg-prose'>" & Replace(Server.HTMLEncode(rs(3) & ""),Chr(13),"<br>") & "</div>"
			'[CG-ANT] html = html & "</div>"

			'[CG-ANT] ' --- SECCION: CIERRE DE CONTENEDOR ---
			'[CG-ANT] 'html = html & "</tr>"
			'[CG-ANT] 'html = html & "</table>"
			'[CG-ANT] html = html & "</div>" ' Fin padding interno
			'[CG-ANT] html = html & "</div>" ' Fin .bloque-datos

			'[CG-ANT] html = html & "</div>" ' Fin #fragment-2

			'[CG-ANT] 'AVANCES
			'[CG-ANT] '[CG-ANT] html = html & "<div id='fragment-3'>"
			'[CG-ANT] html = html & "<div id='fragment-3' class='cg-card'>"
			'[CG-ANT] html = html & "<div class='cg-card__title'>Historial de avances</div>" ' [CG]

			'[CG-ANT] ' --- INICIO NORMALIZACION: CONTENEDOR PRINCIPAL AVANCES ---
			'[CG-ANT] 'html = html & "<table cellspacing=0 align=left>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='bloque-datos ui-panel'>"
			'[CG-ANT] html = html & "<div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='bloque-body ui-panel__body' style='display: flex; flex-direction: column; gap: 2rem;'>"
			'[CG-ANT] html = html & "<div>"

			'[CG-ANT] ' --- SECCION: TEXTO DE AVANCES ---
			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td class='fil' style='font-size:9.25pt;font-family:arial;letter-spacing:0.25px;'>" & Replace(rs(27) & "",Chr(13),"<br>") & "</td>"
			'[CG-ANT] 'html = html & "</tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div>"
			'[CG-ANT] html = html & "<div class='cg-field'>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>DESCRIPCION DE AVANCES</div>"
			'[CG-ANT] html = html & "<div class='cg-field__label'>Descripci&oacute;n de avances</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value' style='width: 100%; margin-top:0.5rem;'>" & Replace(Server.HTMLEncode(rs(27) & ""),Chr(13),"<br>") & "</div>"
			'[CG-ANT] html = html & "<div class='cg-prose'>" & Replace(Server.HTMLEncode(rs(27) & ""),Chr(13),"<br>") & "</div>"
			'[CG-ANT] html = html & "</div>"

			'[CG-ANT] ' --- SECCION: COMENTARIOS HISTORICOS ---
			'[CG-ANT] 'html = html & "<table class='tablDet' style='margin-top:50px;width:80%;' cellspacing=0 align=left>"
			'[CG-ANT] 'html = html & "<tr><td class='fil' style='font-size:12pt;'><b><u>AVANCES / COMENTARIOS</u></b></td></tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div>"
			'[CG-ANT] html = html & "<div class='cg-field'>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label' style='border-bottom: 1px solid #ccc; padding-bottom: 0.5rem; margin-bottom: 1rem;'>HISTORIAL DE RESPONSABLES</div>"
			'[CG-ANT] html = html & "<div class='cg-field__label'>Historial de responsables</div>"
			'[CG-ANT] html = html & HtmlHistorialRelainc(Cnn, id, "proy")

			'[CG-ANT] ' --- SECCION: NUEVO COMENTARIO ---
			'[CG-ANT] 'html = html & "<br><br><b>Nuevo comentario:</b><br>"
			'[CG-ANT] 'html = html & "<textarea class='txtar' id='avances' name='avances' style='height:100px;width:100%;'></textarea></td></tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div style='margin-top: 1rem;'>"
			'[CG-ANT] html = html & "<div class='cg-field cg-field--top'>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>NUEVO COMENTARIO:</div>"
			'[CG-ANT] html = html & "<div class='cg-field__label'>Nuevo comentario</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<textarea class='txtar' id='avances' name='avances' style='height:100px; width:100%; margin-top:0.5rem; border: 1px solid #ccc; border-radius: 4px; padding: 0.5rem;'></textarea>"
			'[CG-ANT] html = html & "<textarea class='cg-input' id='avances' name='avances' rows='4'></textarea>"
			'[CG-ANT] html = html & "</div>"
			
			'[CG-ANT] html = html & "</div>" ' Fin seccion comentarios

			'[CG-ANT] ' --- CIERRE CORRECTO DEL CONTENEDOR PRINCIPAL AVANCES ---
			'[CG-ANT] html = html & "</div>" ' Fin .bloque-body
			'[CG-ANT] html = html & "</div>" ' Fin .bloque-datos
			'[CG-ANT] html = html & "</div>" ' Fin #fragment-3

			'[CG-ANT] 'este es el apartado de avances que muestra los comentarios hechos previamente
			'[CG-ANT] 'se deberia recorrer el recordset y asi podes leerlos

			'[CG-ANT] 'INCIDENTES Y APLs RELACIONADOS
			'[CG-ANT] '[CG-ANT] html = html & "<div id='fragment-4'>"
			'[CG-ANT] html = html & "<div id='fragment-4' class='cg-card'>"
			'[CG-ANT] html = html & "<div class='cg-card__title'>Incidentes y APLs relacionados</div>" ' [CG]

			'[CG-ANT] ' --- INICIO NORMALIZACION: GRID DOS COLUMNAS APLs e INCIDENTES ---
			'[CG-ANT] '[CG-ANT] html = html & "<div class='bloque-datos ui-panel'>"
			'[CG-ANT] html = html & "<div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='bloque-body ui-panel__body grid-datos-doble ui-field-grid--dual'>"
			'[CG-ANT] html = html & "<div class='cg-grid-2'>"

			'[CG-ANT] ' --- SECCION: APLs ---
			'[CG-ANT] 'html = html & "<table style='border:solid 1px #aaaaaa;float:left;' cellspacing=0 cellpadding=3 align=center>"
			'[CG-ANT] 'html = html & "<tr><td><font style='font-size:20pt;color:#000000;'>APLs</font></td></tr>"
			'[CG-ANT] html = html & "<div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>APLs RELACIONADOS</div>"
			'[CG-ANT] html = html & "<div class='cg-field__label'>APLs relacionados</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div style='display: flex; flex-direction: column; gap: 0.8rem; margin-top: 1rem;'>"
			'[CG-ANT] html = html & "<div class='cg-list'>"

			'[CG-ANT] Set rz = ObtenerAplRelac(Cnn, id, "P")

			'[CG-ANT] If Not rz.eof Then
				 '[CG-ANT] Do While Not rz.eof
					'[CG-ANT] 'html = html & "<tr><td align=center>"  & ref & "</td></tr>"
					 '[CG-ANT] '[CG-ANT] html = html & "<div>"
					 '[CG-ANT] html = html & "<div class='cg-list__item'>"
					'[CG-ANT] '[CG-ANT] html = html & "<a href='detalles2.asp?opc=ver&id=" & rz(1) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "' style='text-decoration:none;'>"
					'[CG-ANT] html = html & "<a class='cg-link' href='detalles2.asp?opc=ver&id=" & rz(1) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
					'[CG-ANT] '[CG-ANT] html = html & "<span style='cursor:pointer; text-decoration:underline; color:#FF0000; font-weight:600;'>" & rz(5) & "</span>"
					'[CG-ANT] html = html & rz(5)
					'[CG-ANT] html = html & "</a>"
					 '[CG-ANT] html = html & "</div>"
					 '[CG-ANT] rz.movenext
				 '[CG-ANT] Loop
			'[CG-ANT] Else
				'[CG-ANT] 'html = html & "<tr><td>-----</td></tr>"
				'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>-----</div>"
				'[CG-ANT] html = html & "<div class='cg-empty'>No hay APLs relacionados.</div>"
			'[CG-ANT] End If
			'[CG-ANT] 'html = html & "</table>"
			'[CG-ANT] html = html & "</div>" ' Fin flex col APLs
			'[CG-ANT] html = html & "</div>" ' Fin columna 1 de la grid

			'[CG-ANT] ' --- SECCION: INCIDENTES ---
			'[CG-ANT] 'html = html & "<table style='border:solid 1px #aaaaaa;float:right;' cellspacing=0 cellpadding=3 align=center>"
			'[CG-ANT] 'html = html & "<tr><td><font style='font-size:20pt;color:#000000;'>INCIDENTES</font></td></tr>"
			'[CG-ANT] html = html & "<div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>INCIDENTES RELACIONADOS</div>"
			'[CG-ANT] html = html & "<div class='cg-field__label'>Incidentes relacionados</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div style='display: flex; flex-direction: column; gap: 0.8rem; margin-top: 1rem;'>"
			'[CG-ANT] html = html & "<div class='cg-list'>"

			'[CG-ANT] Set rz = ObtenerIncidentesPorProyecto(Cnn, id)

			'[CG-ANT] If Not rz.eof Then
				'[CG-ANT] Do While Not rz.eof
					'[CG-ANT] 'html = html & "<tr><td align=center>" & ref & "</td></tr>"
					'[CG-ANT] '[CG-ANT] html = html & "<div>"
					'[CG-ANT] html = html & "<div class='cg-list__item'>"
					'[CG-ANT] '[CG-ANT] html = html & "<a href='detalles.asp?opc=veri&id=" & rz(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "' style='text-decoration:none;'>"
					'[CG-ANT] html = html & "<a class='cg-link' href='detalles.asp?opc=veri&id=" & rz(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
					'[CG-ANT] '[CG-ANT] html = html & "<span style='cursor:pointer; text-decoration:underline; color:#FF0000; font-weight:600;'>" & rz(0) & "</span>"
					'[CG-ANT] html = html & rz(0)
					'[CG-ANT] html = html & "</a>"
					'[CG-ANT] html = html & "</div>"
					'[CG-ANT] rz.movenext
				'[CG-ANT] Loop
			'[CG-ANT] Else
				'[CG-ANT] 'html = html & "<tr><td>-----</td></tr>"
				'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>-----</div>"
				'[CG-ANT] html = html & "<div class='cg-empty'>No hay incidentes relacionados.</div>"
			'[CG-ANT] End If
			'[CG-ANT] 'html = html & "</table>"
			'[CG-ANT] html = html & "</div>" ' Fin flex col Incidentes
			'[CG-ANT] html = html & "</div>" ' Fin columna 2 de la grid

			'[CG-ANT] ' --- CIERRE FINAL DEL FRAGMENT 4 ---
			'[CG-ANT] html = html & "</div>" ' Fin .bloque-body
			'[CG-ANT] html = html & "</div>" ' Fin .bloque-datos
			'[CG-ANT] html = html & "</div>" ' Fin #fragment-4


			'[CG-ANT] '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

			'[CG-ANT] ' --- AHORA SN: CERRAMOS LOS WRAPPERS PRINCIPALES ---
			'[CG-ANT] html = html & "</div>" ' Cierra el <div id='tabs'>
			'[CG-ANT] html = html & "</div>" ' Cierra el <div style='width:100%;'> principal


			'[CG-ANT] 'html = html & "</div>"

			'[CG-ANT] 'html = html & "<div style='height:100%;width:50%;float:right'><br>"






		'[CG-ANT] End If

		'[CG-ANT] ' DEPURADO [CG]: cierre de </table> sin tabla abierta (resto de la maquetacion vieja)
		'[CG-ANT] '[CG-ANT] html = html & "</table>"

	'[CG-ANT] ElseIf opc = "acti" Then

		'[CG-ANT] zzz = ObjAcc(id,"I",cnn,"A")
		'[CG-ANT] XXX = SYSLOGON(cnn,"ACTUALIZA_INC_" & id)

		'[CG-ANT] '[CG-ANT] html = html & "<div class='layout-principal'>"
		'[CG-ANT] html = html & "<div class='cg-app cg-edit'>"
		'[CG-ANT] 'html = html & "<div style='height:80%;width:50%;float:left;'><br>"
		'[CG-ANT] '[CG-ANT] html = html & "<div class='columna-izquierda'><br>"
		'[CG-ANT] html = html & "<div class='cg-edit__col'>"

		'[CG-ANT] 'html = html & "<form id='sForm' method='post' action='actualizar_v2.asp?tabla=incidentes&id=" & id & "&opt=actinc&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=detalles.asp?opc=acti' >" 'agregado por update relainc
		'[CG-ANT] ' Se agregN &relainc=relainc para que la siguiente pNgina sepa quN tabla usar
		'[CG-ANT] html = html & "<form id='sForm' method='post' action='actualizar_v2.asp?tabla=incidentes&relainc=relainc&id=" & id & "&opt=actinc&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=detalles.asp?opc=acti' >"
		'[CG-ANT] 'html = html & "<table class='tablDet' style='width:70%;' cellspacing=0 align=center>"
		'[CG-ANT] '[CG-ANT] html = html & "<div class='bloque-datos ui-panel'>"
		'[CG-ANT] html = html & "<div class='cg-card'>"
		
		'[CG-ANT] Dim wPermisos, editaFunc, editaTec, editaFuncEstado
		'[CG-ANT] wPermisos = getPermisos(xmenu,xtmenu,"MODIF")
		'[CG-ANT] editaFunc = (CStr(wPermisos) = "1" And session("FUNC") = "FUNC")
		'[CG-ANT] editaTec = (CStr(wPermisos) = "1" And session("FUNC") = "TEC")
		'[CG-ANT] Set rx = ObtenerIncidenteEdicion(Cnn, id) ' TODO: revisar/crear SP sugerido sp_ObtenerIncidenteEdicion

		'[CG-ANT] 'html = html & "<tr>"
		'[CG-ANT] 'html = html & "<td class='fil' colspan=2 style='font-size:12pt;'><b><u>DATOS</u></b></td>"
		'[CG-ANT] 'html = html & "</tr>"

		'[CG-ANT] '[CG-ANT] html = html & "<div class='bloque-header'>"
		'[CG-ANT] html = html & "<div class='cg-card__title'>"
		'[CG-ANT] '[CG-ANT] html = html & "DATOS"
		'[CG-ANT] html = html & "Informaci&oacute;n"
		'[CG-ANT] html = html & "</div>"
		'[CG-ANT] '[CG-ANT] html = html & "<div class='bloque-body ui-panel__body'>"
		'[CG-ANT] html = html & "<div>"
		'[CG-ANT] '[CG-ANT] html = html & "<div class='grid-datos'>"
		'[CG-ANT] html = html & "<div class='cg-fields'>"

		'[CG-ANT] 'html = html & "<tr>"
		'[CG-ANT] 'html = html & "<td class='fil'>NRO</td><td id='nid' class='fil'>" & id & "</td>"
		'[CG-ANT] 'html = html & "</tr>"
		'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>NRO</div><div id='nid' class='dato-value'>" & id & "</div>"
		'[CG-ANT] html = html & "<div><div class='cg-field__label'>Nro</div><div id='nid' class='cg-field__value'>" & id & "</div></div>"

		'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>CLIENTE</div><div class='dato-value'>"
		'[CG-ANT] html = html & "<div><div class='cg-field__label'>Cliente</div><div class='cg-field__value'>"
		'[CG-ANT] html = html & ComboCampo(Cnn, rx(1), "id_cliente", "id_cliente", editaFunc)
		'[CG-ANT] '[CG-ANT] html = html & "</div>"
		'[CG-ANT] html = html & "</div></div>"

		'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>TITULO</div><div class='dato-value'>"
		'[CG-ANT] html = html & "<div><div class='cg-field__label'>T&iacute;tulo</div><div class='cg-field__value'>"
		'[CG-ANT] html = html & HtmlInputCampo("titulo", rx(28), editaFunc)
		'[CG-ANT] '[CG-ANT] html = html & "</div>"
		'[CG-ANT] html = html & "</div></div>"

		'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>CONTACTO</div><div class='dato-value'>"
		'[CG-ANT] html = html & "<div><div class='cg-field__label'>Contacto</div><div class='cg-field__value'>"
		'[CG-ANT] html = html & ComboCampo(Cnn, rx(3), "contacto", "contacto", editaFunc)
		'[CG-ANT] '[CG-ANT] html = html & "</div>"
		'[CG-ANT] html = html & "</div></div>"

		'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>APLICACION</div><div class='dato-value'>"
		'[CG-ANT] html = html & "<div><div class='cg-field__label'>Aplicaci&oacute;n</div><div class='cg-field__value'>"
		'[CG-ANT] html = html & ComboCampo(Cnn, rx(5), "id_aplic", "id_aplic", editaFunc)
		'[CG-ANT] '[CG-ANT] html = html & "</div>"
		'[CG-ANT] html = html & "</div></div>"

		'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>RELACIONADO A PROYECTO</div><div class='dato-value'>"
		'[CG-ANT] html = html & "<div><div class='cg-field__label'>Relacionado a proyecto</div><div class='cg-field__value'>"
 
		'[CG-ANT] html = html & ComboProyectoCampo(Cnn, rx(1), rx(1), editaFunc)
		'[CG-ANT] 'html = html & ComboCampo(Cnn, rx(1), "id_proyecto", "id_proyecto", editaFunc)
		'[CG-ANT] '[CG-ANT] html = html & "</div>"
		'[CG-ANT] html = html & "</div></div>"

		'[CG-ANT] ' ' html = html & "<div class='dato-label'>PRIORIDAD</div><div class='dato-value'>"
		'[CG-ANT] ' ' html = html & ComboCampo(Cnn, rx(7), "prioridad", "prioridad", editaFunc)
		'[CG-ANT] ' ' html = html & "</div>"

		'[CG-ANT] ' ' html = html & "<div class='dato-label'>RESPONSABLE GESTION</div><div class='dato-value'>"
		'[CG-ANT] ' ' html = html & ComboCampo(Cnn, rx(10), "respon_gestion", "respon_gestion", editaFunc)
		'[CG-ANT] ' ' html = html & "</div>"

		'[CG-ANT] ' ' html = html & "<div class='dato-label'>RESPONSABLE TECNICO</div><div class='dato-value'>"
		'[CG-ANT] ' ' html = html & ComboCampo(Cnn, rx(12), "respon_tec", "respon_tec", editaTec)
		'[CG-ANT] ' ' html = html & "</div>"

		'[CG-ANT] sFec = Fechador(rx(19) + nFecini)

		'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>REQUIERE PRESUPUESTO</div><div class='dato-value'>"
		'[CG-ANT] html = html & "<div><div class='cg-field__label'>Requiere presupuesto</div><div class='cg-field__value'>"
		'[CG-ANT] html = html & HtmlSelectSiNo("req_presup", rx(14), editaFunc)
		'[CG-ANT] '[CG-ANT] html = html & "</div>"
		'[CG-ANT] html = html & "</div></div>"

		'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>FECHA INGRESO</div><div class='dato-value'>"
		'[CG-ANT] html = html & "<div><div class='cg-field__label'>Fecha ingreso</div><div class='cg-field__value'>"
		'[CG-ANT] html = html & sFec
		'[CG-ANT] '[CG-ANT] html = html & "</div>"
		'[CG-ANT] html = html & "</div></div>"

		'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>BROWSER</div><div class='dato-value'>"
		'[CG-ANT] html = html & "<div><div class='cg-field__label'>Browser</div><div class='cg-field__value'>"
		'[CG-ANT] html = html & ComboCampo(Cnn, rx(20), "browser", "browser", editaFunc)
		'[CG-ANT] '[CG-ANT] html = html & "</div>"
		'[CG-ANT] html = html & "</div></div>"

		'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>VIA PEDIDO</div><div class='dato-value'>"
		'[CG-ANT] html = html & "<div><div class='cg-field__label'>V&iacute;a pedido</div><div class='cg-field__value'>"
		'[CG-ANT] html = html & ComboCampo(Cnn, rx(22), "via_pedido", "via_pedido", editaFunc)
		'[CG-ANT] '[CG-ANT] html = html & "</div>"
		'[CG-ANT] html = html & "</div></div>"

		'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>ESTADO DESARROLLO</div><div class='dato-value'>"
		'[CG-ANT] html = html & "<div><div class='cg-field__label'>Estado desarrollo</div><div class='cg-field__value'>"
		'[CG-ANT] html = html & ComboCampo(Cnn, rx(15), "estado_desa", "estado_desa", editaTec)
		'[CG-ANT] '[CG-ANT] html = html & "</div>"
		'[CG-ANT] html = html & "</div></div>"

		'[CG-ANT] editaFuncEstado = (editaFunc And (rx(16) & "" = "OK" Or rx(16) & "" = "BAJA" Or rx(16) & "" = "---"))
		'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>ESTADO FUNCIONAL</div><div class='dato-value'>"
		'[CG-ANT] html = html & "<div><div class='cg-field__label'>Estado funcional</div><div class='cg-field__value'>"
		'[CG-ANT] html = html & ComboCampo(Cnn, rx(17), "estado_func", "estado_func", editaFuncEstado)
		'[CG-ANT] '[CG-ANT] html = html & "</div>"
		'[CG-ANT] html = html & "</div></div>"

		'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>FECHA OK CLIENTE</div><div class='dato-value'>"
		'[CG-ANT] html = html & "<div><div class='cg-field__label'>Fecha OK cliente</div><div class='cg-field__value'>"
		'[CG-ANT] html = html & HtmlInputCampo("ok_cliente", rx(25), editaFunc And (rx(16) & "" = "OK" And rx(18) & "" = "OK"))
		'[CG-ANT] '[CG-ANT] html = html & "</div>"
		'[CG-ANT] html = html & "</div></div>"

		'[CG-ANT] 'html = html & "<tr>"
		'[CG-ANT] 'html = html & "<td class='fil'>LINK DOCUMENTO</td><td class='fil'>"
		'[CG-ANT] '[CG-ANT] test = "<a href='" & Server.HTMLEncode(rx("ruta") & "") & "'><span style='color:#FF0000;'>Test</span></a>"
		'[CG-ANT] test = "<a class='cg-link' href='" & Server.HTMLEncode(rx("ruta") & "") & "'>Test</a>"
		'[CG-ANT] 'html = html & "<input class='inpu' style='width:90%;' id='ruta' name='ruta' size=50 type='text' value='" & rx("ruta") & "'>" & test
		'[CG-ANT] 'html = html & "</td>"
		'[CG-ANT] 'html = html & "</tr>"
		'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>LINK DOCUMENTO</div><div class='dato-value'>"
		'[CG-ANT] html = html & "<div><div class='cg-field__label'>Link documento</div><div class='cg-field__value cg-field__value--inline'>"
		'[CG-ANT] '[CG-ANT] html = html & "<input class='inpu ui-input' style='width:90%;' id='ruta' name='ruta' size=50 type='text' value='" & Server.HTMLEncode(rx("ruta") & "") & "'>" & test
		'[CG-ANT] html = html & "<input class='cg-input' id='ruta' name='ruta' size=50 type='text' value='" & Server.HTMLEncode(rx("ruta") & "") & "'>" & test
		'[CG-ANT] '[CG-ANT] html = html & "</div>"
		'[CG-ANT] html = html & "</div></div>"


		'[CG-ANT] ' --- LA MAGIA DE LOS CIERRES Y EL BOTNN ---

		'[CG-ANT] 'html = html & "<tr>"
		'[CG-ANT] 'html = html & "<td colspan=2>&nbsp;</td>"
		'[CG-ANT] 'html = html & "</tr>"

		'[CG-ANT] ' 1. Cierra el grid-datos para salir de las dos columnas
		'[CG-ANT] html = html & "</div>"

		'[CG-ANT] 'html = html & "<tr>"
		'[CG-ANT] 'html = html & "<td id='sAgre' class='fil' colspan=2 align=center><font style='color:#FF0000;'>RELACIONAR APL</font></td>"
		'[CG-ANT] 'html = html & "</tr>"

		'[CG-ANT] ' Creamos el botNn centrado (con un margin-top de 1.5rem reemplazamos ese &nbsp; viejo para que respire)
		'[CG-ANT] '[CG-ANT] html = html & "<div id='sAgre' style='text-align: center; margin-top: 1.5rem; cursor: pointer;'>"
		'[CG-ANT] html = html & "<div id='sAgre' class='cg-actions-center'>"
		'[CG-ANT] '[CG-ANT] html = html & "  <span class='ui-button ui-button--quiet' style='color:#FF0000;'>RELACIONAR APL</span>"
		'[CG-ANT] html = html & "<span class='cg-btn cg-btn--secondary cg-btn--sm'>Relacionar APL</span>"
		'[CG-ANT] html = html & "</div>"

		'[CG-ANT] ' 2 y 3. Cerramos el cuerpo y la tarjeta principal
		'[CG-ANT] html = html & "</div>" ' Cierra bloque-body
		'[CG-ANT] html = html & "</div>" ' Cierra bloque-datos

		'[CG-ANT] '[CG-ANT] html = html & "<br><span style='color:#000000;font-size:12pt;margin-left:100px;'>"
		'[CG-ANT] html = html & "<div class='cg-field__hint'>"
		'[CG-ANT] '[CG-ANT] html = html & "<u>ejemplo link:</u> 'http://server-40/registroincidente/' + nombre documento"
		'[CG-ANT] html = html & "Ejemplo de link: 'http://server-40/registroincidente/' + nombre documento"
		'[CG-ANT] '[CG-ANT] html = html & "</span>"
		'[CG-ANT] html = html & "</div>"

		'[CG-ANT] html = html & "</div>"

		'[CG-ANT] 'html = html & "<div style='height:100%;width:50%;float:right;'><br>"
		'[CG-ANT] '[CG-ANT] html = html & "<div class='columna-derecha'><br>"
		'[CG-ANT] html = html & "<div class='cg-edit__col'>"

		'[CG-ANT] 'html = html & "<table class='tablDet' style='width:80%;' cellspacing=0 align=left>"
		'[CG-ANT] ' Reutilizamos nuestra tarjeta base
		'[CG-ANT] '[CG-ANT] html = html & "<div class='bloque-datos ui-panel'>"
		'[CG-ANT] html = html & "<div class='cg-card'>"
		'[CG-ANT] html = html & "<div class='cg-card__title'>Detalle</div>" ' [CG]
		'[CG-ANT] ' Usamos flex-column con un gap (espacio) para separar los 3 bloques sin usar <br>
		'[CG-ANT] '[CG-ANT] html = html & "<div class='bloque-body ui-panel__body' style='display: flex; flex-direction: column; gap: 1.5rem;'>"
		'[CG-ANT] html = html & "<div>"

		'[CG-ANT] 'html = html & "<tr>"
		'[CG-ANT] 'html = html & "<td class='fil' style='font-size:12pt;'><u><b>DETALLES</b></u></td>"
		'[CG-ANT] 'html = html & "</tr>"
		'[CG-ANT] '[CG-ANT] html = html & "<div>" ' Abrimos el contenedor de DETALLES
		'[CG-ANT] html = html & "<div class='cg-field'>"
		'[CG-ANT] '[CG-ANT] html = html & "<div style='font-size: 1rem; font-weight: 700; color: #0f172a; margin-bottom: 0.5rem; letter-spacing: 0.05em;'>DETALLES</div>"
		'[CG-ANT] html = html & "<div class='cg-field__label'>Descripci&oacute;n</div>"

		'[CG-ANT] 'If rx(10) = session("USUARIO") And session("FUNC") = "FUNC" Then
		'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>"
		'[CG-ANT] html = html & "<div>"
		'[CG-ANT] html = html & HtmlTextAreaCampo("detalles", rx(8), "200px", editaFunc)
		'[CG-ANT] html = html & "</div>"

		'[CG-ANT] html = html & "</div>" ' Cerramos el contenedor de DETALLES


		'[CG-ANT] 'html = html & "<tr>"
		'[CG-ANT] 'html = html & "<td class='fil' style='font-size:12pt;'><BR><u><b>PATH</b></u></td>"
		'[CG-ANT] 'html = html & "</tr>"
		'[CG-ANT] '[CG-ANT] html = html & "<div>" ' Abrimos el contenedor de PATH
		'[CG-ANT] html = html & "<div class='cg-field'>"
		'[CG-ANT] '[CG-ANT] html = html & "<div style='font-size: 1rem; font-weight: 700; color: #0f172a; margin-bottom: 0.5rem; letter-spacing: 0.05em;'>PATH</div>"
		'[CG-ANT] html = html & "<div class='cg-field__label'>Path</div>"

		'[CG-ANT] 'html = html & "<tr>"
		'[CG-ANT] 'html = html & "<td class='fil'><BR>"
		'[CG-ANT] XPATH="\\" & request.servervariables("SERVER_NAME") &  RIGHT(RUTA,LEN(RUTA)-2)
		'[CG-ANT] 'html = html & XPATH & "INCIDENTES\" & RX(0)
		'[CG-ANT] 'html = html  & "<BR></td></tr>"
		'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value' style='word-break: break-all; color: #475569;'>"
		'[CG-ANT] html = html & "<div class='cg-path'><span class='cg-path__text'>"
		'[CG-ANT] html = html & XPATH & "INCIDENTES\" & RX(0)
		'[CG-ANT] '[CG-ANT] html = html & "</div>"
		'[CG-ANT] html = html & "</span></div>"

		'[CG-ANT] html = html & "</div>" ' Cerramos el contenedor de PATH


		'[CG-ANT] 'html = html & "<tr>"
		'[CG-ANT] 'html = html & "<td class='fil' style='font-size:12pt;'><BR><u><b>ARCHIVOS</b></u></td>"
		'[CG-ANT] 'html = html & "</tr>"
		'[CG-ANT] '[CG-ANT] html = html & "<div>" ' Abrimos el contenedor de ARCHIVOS
		'[CG-ANT] html = html & "<div class='cg-field'>"
		'[CG-ANT] '[CG-ANT] html = html & "<div style='font-size: 1rem; font-weight: 700; color: #0f172a; margin-bottom: 0.5rem; letter-spacing: 0.05em;'>ARCHIVOS</div>"
		'[CG-ANT] html = html & "<div class='cg-field__label'>Archivos</div>"

		'[CG-ANT] 'html = html & "<tr>"
		'[CG-ANT] 'html = html & "<td class='fil'>"
		'[CG-ANT] RUTA=RUTA & "\INCIDENTES\" & RX(0)

		'[CG-ANT] 'html = html & LISTARDIRTORIO(RUTA,RX(0),"INCIDENTES")
		'[CG-ANT] 'html = html & "</td></tr>"
		'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>"
		'[CG-ANT] html = html & "<ul class='cg-files'>"
		'[CG-ANT] html = html & LISTARDIRTORIO(RUTA,RX(0),"INCIDENTES")
		'[CG-ANT] '[CG-ANT] html = html & "</div>"
		'[CG-ANT] html = html & "</ul>"

		'[CG-ANT] html = html & "</div>" ' Cerramos el contenedor de ARCHIVOS

		'[CG-ANT] 'html = html & "</table>"
		'[CG-ANT] html = html & "</div>" ' Cerramos el bloque-body
		'[CG-ANT] html = html & "</div>" ' Cerramos el bloque-datos principal

		'[CG-ANT] 'html = html & "<table class='tablDet' style='margin-top:50px;width:80%;' cellspacing=0 align=left>"
		'[CG-ANT] '[CG-ANT] html = html & "<div class='bloque-datos ui-panel' >"
		'[CG-ANT] html = html & "<div class='cg-card'>"
		'[CG-ANT] html = html & "<div class='cg-card__title'>Historial de responsables</div>" ' [CG]
		'[CG-ANT] '[CG-ANT] html = html & "<div class='bloque-body ui-panel__body' style='display: flex; flex-direction: column; gap: 1.5rem;'>"
		'[CG-ANT] html = html & "<div>"

		'[CG-ANT] 'html = html & "<tr>"
		'[CG-ANT] 'html = html & "<td class='fil' style='font-size:12pt;'><b><u>AVANCES </u></b></td>"
		'[CG-ANT] 'html = html & "</tr>"
		'[CG-ANT] html = html & "<div>" ' Contenedor general del tNtulo y contenido
		'[CG-ANT] ' DEPURADO [CG]: el titulo ahora es el header de la tarjeta (cg-card__title)
		'[CG-ANT] '[CG-ANT] html = html & "<div style='font-size: 1rem; font-weight: 700; color: #0f172a; margin-bottom: 1rem; letter-spacing: 0.05em;'>HISTORIAL RESPONSABLES</div>"
		'[CG-ANT] html = html & HtmlHistorialRelainc(Cnn, id, "inc")

		'[CG-ANT] If editaFunc Or editaTec Then
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value' style='width: 100%; overflow-x: auto;'>"
			'[CG-ANT] html = html & "<div class='cg-table-scroll cg-field--top'>"
			'[CG-ANT] html = html & HtmlCamposAvance(RolResponsableActual())
			'[CG-ANT] html = html & "</div>"
		'[CG-ANT] End If

		'[CG-ANT] 'html = html & "<tr>"
		'[CG-ANT] 'html = html & "<td>&nbsp;</td>"
		'[CG-ANT] 'html = html & "</tr>"

		'[CG-ANT] 'html = html & "<tr>"
		'[CG-ANT] 'html = html & "<td>"
		'[CG-ANT] 'html = html & "<input class='bot ui-button' type='button' xstyle='border-radius:10px;border:solid 2px;border-color:#FF0000;height:100%;width:100%;' value='ACTUALIZAR' onclick=""actualizar('sForm');"">"
		'[CG-ANT] 'html = html & "</td></tr>"

		'[CG-ANT] ' BotNn de Actualizar alineado y limpio
		'[CG-ANT] '[CG-ANT] html = html & "<div class='ui-edit-actions' style='margin-top: 0.5rem;'>"
		'[CG-ANT] html = html & "<div class='cg-savebar'>"
		'[CG-ANT] '[CG-ANT] html = html & "<input class='bot ui-button ui-button--primary' type='button' style='padding: 0.6rem 1.5rem; background-color: #3b82f6; color: white; border: none; border-radius: 6px; font-weight: 600; cursor: pointer;' value='ACTUALIZAR' onclick=""actualizar('sForm');"">"
		'[CG-ANT] html = html & "<input class='cg-btn cg-btn--primary cg-btn--lg' type='button' value='Actualizar informaci&oacute;n' onclick=""actualizar('sForm');"">"
		'[CG-ANT] html = html & "</div>"

		'[CG-ANT] 'html = html & "</table>"

		'[CG-ANT] html = html & "</div>" ' Cierra el contenedor general que abrimos arriba del tNtulo
		'[CG-ANT] html = html & "</div>" ' Cierra bloque-body
		'[CG-ANT] html = html & "</div>" ' Cierra bloque-datos
		'[CG-ANT] html = html & "</div>" 'Cierra layout-principal

		'[CG-ANT] html = html & "</form>"


		'[CG-ANT] rx.close

	'[CG-ANT] ElseIf opc = "actp" Then

		'[CG-ANT] zzz = ObjAcc(id,"P",cnn,"A")
		'[CG-ANT] XXX = SYSLOGON(cnn,"ACTUALIZA_PROY_" & id)
		'[CG-ANT] wPermisos = getPermisos(xmenu,xtmenu,"MODIF")
		'[CG-ANT] editaFunc = (CStr(wPermisos) = "1" And session("FUNC") = "FUNC")
		'[CG-ANT] editaTec = (CStr(wPermisos) = "1" And session("FUNC") = "TEC")
		

		'[CG-ANT] Set rs = ObtenerProyectoEdicion(Cnn, id)
			'[CG-ANT] 'response.write(sql) ' comentado 6/11/2025
		'[CG-ANT] If Not rs.eof Then

			'[CG-ANT] ' ' ' --- COLUMNA IZQUIERDA (CONTENEDOR Y FORMULARIO) ---

			'[CG-ANT] ' ' Abrimos el contenedor global con padding
			'[CG-ANT] '[CG-ANT] html = html & "<div class='inner' style='padding: 2rem;'>"
			'[CG-ANT] html = html & "<div class='cg-app cg-edit'>"

			'[CG-ANT] 'html = html & "<div style='height:85%;width:50%;float:left;'><br>"
			'[CG-ANT] ' Mantenemos el float por ahora para no romper tu maquetaciNn general, pero le damos padding derecho para que respire con la otra columna
			'[CG-ANT] '[CG-ANT] html = html & "<div style='width: 50%; float: left; padding-right: 1.5rem;'>"
			'[CG-ANT] html = html & "<div class='cg-edit__col'>"

			'[CG-ANT] '[CG-ANT] html = html & "<form id='sForm' method='post' action='actualizar_v2.asp?tabla=proyectos&id=" & id & "&opt=actproy&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=detalles.asp?opc=actp&sol=" & rs(1) & "' >" 'agregado para probar update relainc
			'[CG-ANT] html = html & "<form id='sForm' class='cg-card' method='post' action='actualizar_v2.asp?tabla=proyectos&id=" & id & "&opt=actproy&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=detalles.asp?opc=actp&sol=" & rs(1) & "' >" 'agregado para probar update relainc


			'[CG-ANT] '[CG-ANT] html = html & "<div class='bloque-header'>"
			'[CG-ANT] html = html & "<div class='cg-card__title'>"
			'[CG-ANT] '[CG-ANT] html = html & "DATOS"
			'[CG-ANT] html = html & "Informaci&oacute;n"
			'[CG-ANT] html = html & "</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='bloque-body ui-panel__body'>"
			'[CG-ANT] html = html & "<div>"
		
			'[CG-ANT] 'html = html & "<table class='tablDet' style='width:70%;border-style:none;' cellspacing=0 align=center>"
			'[CG-ANT] ' Abrimos nuestra tarjeta y la grilla
			'[CG-ANT] '[CG-ANT] html = html & "<div class='bloque-datos ui-panel'>"
			'[CG-ANT] html = html & "<div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='bloque-body ui-panel__body'>"
			'[CG-ANT] html = html & "<div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='grid-datos-doble'>"
			'[CG-ANT] html = html & "<div class='cg-fields'>"

			'[CG-ANT] ' html = html & "<tr>"
			'[CG-ANT] ' html = html & "<td class='fil' style='font-size:12pt;' colspan=2><b><u>DATOS</u></b></td>"
			'[CG-ANT] ' html = html & "</tr>"

			'[CG-ANT] '--- DATOS DE LA GRILLA ---

			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td class='fil' style='" & st0 & "width:40%;'>ID</td><td id='nid' class='fil' style='" & st1 & "'>" & rs(0) & "</td>"
			'[CG-ANT] 'html = html & "</tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>ID</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>ID</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value' id='nid' style='font-weight: 600;'>" & rs(0) & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value' id='nid'>" & rs(0) & "</div></div>"

			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>SOLICITUD NUMERO</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Solicitud n&uacute;mero</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>" & HtmlInputCampo("solic_nro", rs(1), editaFunc) & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value'>" & HtmlInputCampo("solic_nro", rs(1), editaFunc) & "</div></div>"

			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>TITULO PROYECTO</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>T&iacute;tulo proyecto</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>" & HtmlInputCampo("titulo", rs(2), editaFunc) & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value'>" & HtmlInputCampo("titulo", rs(2), editaFunc) & "</div></div>"

			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>CLIENTE</div><div class='dato-value'>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Cliente</div><div class='cg-field__value'>"
			'[CG-ANT] html = html & ComboCampo(Cnn, rs(21), "id_cliente", "id_cliente", editaFunc)
			'[CG-ANT] '[CG-ANT] html = html & "</div>"
			'[CG-ANT] html = html & "</div></div>"

			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>CONTACTO</div><div class='dato-value'>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Contacto</div><div class='cg-field__value'>"
			'[CG-ANT] html = html & ComboCampo(Cnn, rs(22), "contacto", "contacto", editaFunc)
			'[CG-ANT] '[CG-ANT] html = html & "</div>"
			'[CG-ANT] html = html & "</div></div>"

			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>APLICACION</div><div class='dato-value'>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Aplicaci&oacute;n</div><div class='cg-field__value'>"
			'[CG-ANT] html = html & ComboCampo(Cnn, rs(23), "id_aplic", "id_aplic", editaFunc)
			'[CG-ANT] '[CG-ANT] html = html & "</div>"
			'[CG-ANT] html = html & "</div></div>"

			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>RESPONSABLE GESTION</div><div class='dato-value'>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Responsable gesti&oacute;n</div><div class='cg-field__value'>"
			'[CG-ANT] html = html & ComboCampo(Cnn, rs(24), "respon_gestion", "respon_gestion", editaFunc)
			'[CG-ANT] '[CG-ANT] html = html & "</div>"
			'[CG-ANT] html = html & "</div></div>"

			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>RESPONSABLE TECNICO</div><div class='dato-value'>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Responsable t&eacute;cnico</div><div class='cg-field__value'>"
			'[CG-ANT] html = html & ComboCampo(Cnn, rs(25), "respon_tec", "respon_tec", editaTec)
			'[CG-ANT] '[CG-ANT] html = html & "</div>"
			'[CG-ANT] html = html & "</div></div>"

			'[CG-ANT] sFec = ""
			'[CG-ANT] If rs(9) <> 0 Then
				'[CG-ANT] sFec = FormatDateTime((Fechador(rs(9) + nFecini)), 2)
			'[CG-ANT] End If
			
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>FECHA INGRESO</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Fecha ingreso</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>" & sFec  & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value'>" & sFec & "</div></div>"

			'[CG-ANT] sFecEn = ""
			'[CG-ANT] If rs(10) <> 0 Then
				'[CG-ANT] sFecEn = FormatDateTime((Fechador(rs(10) + nFecini)), 2)
			'[CG-ANT] End If

			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>FECHA ENTREGA</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Fecha entrega</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>" & HtmlInputCampo("fecha_entrega", sFecEn, editaFunc) & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value'>" & HtmlInputCampo("fecha_entrega", sFecEn, editaFunc) & "</div></div>"

			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>PRIORIDAD</div><div class='dato-value'>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Prioridad</div><div class='cg-field__value'>"
			'[CG-ANT] html = html & ComboCampo(Cnn, rs(28), "prioridad", "prioridad", editaFunc)
			'[CG-ANT] '[CG-ANT] html = html & "</div>"
			'[CG-ANT] html = html & "</div></div>"

			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>RELEV Y DEFINICION (hs)</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Relev y definici&oacute;n (hs)</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>" & HtmlInputCampo("tiempo_relevamiento", rs(14), editaFunc) & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value'>" & HtmlInputCampo("tiempo_relevamiento", rs(14), editaFunc) & "</div></div>"

			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>DISENO (hs)</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Dise&ntilde;o (hs)</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>" & HtmlInputCampo("tiempo_diseno", rs(15), editaFunc) & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value'>" & HtmlInputCampo("tiempo_diseno", rs(15), editaFunc) & "</div></div>"

			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>DESARROLLO (hs)</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Desarrollo (hs)</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>" & HtmlInputCampo("tiempo_desarr", rs(16), editaTec) & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value'>" & HtmlInputCampo("tiempo_desarr", rs(16), editaTec) & "</div></div>"

			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>TESTEO (hs)</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Testeo (hs)</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>" & HtmlInputCampo("tiempo_test", rs(17), editaFunc) & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value'>" & HtmlInputCampo("tiempo_test", rs(17), editaFunc) & "</div></div>"

			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>TOTAL (hs)</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Total (hs)</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value' style='font-weight: 600; padding-left: 0.5rem;'>" & rs(32) & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value cg-hours__total'>" & rs(32) & "</div></div>"

			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>BROWSER</div><div class='dato-value'>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Browser</div><div class='cg-field__value'>"
			'[CG-ANT] html = html & ComboCampo(Cnn, rs(29), "browser", "browser", editaFunc)
			'[CG-ANT] '[CG-ANT] html = html & "</div>"
			'[CG-ANT] html = html & "</div></div>"

			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>VIA PEDIDO</div><div class='dato-value'>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>V&iacute;a pedido</div><div class='cg-field__value'>"
			'[CG-ANT] html = html & ComboCampo(Cnn, rs(30), "via_pedido", "via_pedido", editaFunc)
			'[CG-ANT] '[CG-ANT] html = html & "</div>"
			'[CG-ANT] html = html & "</div></div>"

			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>REQUIERE PRESUPUESTO</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Requiere presupuesto</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>" & HtmlSelectSiNo("req_presup", rs(20), editaFunc) & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value'>" & HtmlSelectSiNo("req_presup", rs(20), editaFunc) & "</div></div>"

			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>ESTADO DESARROLLO</div><div class='dato-value'>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Estado desarrollo</div><div class='cg-field__value'>"
			'[CG-ANT] html = html & ComboCampo(Cnn, rs(27), "estado_desa", "estado_desa", editaTec)
			'[CG-ANT] '[CG-ANT] html = html & "</div>"
			'[CG-ANT] html = html & "</div></div>"

			'[CG-ANT] 'editaFuncEstado = (editaFunc And (rx(16) & "" = "OK" Or rx(16) & "" = "BAJA" Or rx(16) & "" = "---"))
		
			'[CG-ANT] editaFuncEstado = (editaFunc And (rs(11) & "" = "OK" Or rs(11) & "" = "BAJA"))
			 '[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>ESTADO FUNCIONAL</div>"
			 '[CG-ANT] html = html & "<div><div class='cg-field__label'>Estado funcional</div>"
			 '[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>" & ComboCampo(Cnn, rs(26), "estado_func", "estado_func", editaFuncEstado) & "</div>"
			 '[CG-ANT] html = html & "<div class='cg-field__value'>" & ComboCampo(Cnn, rs(26), "estado_func", "estado_func", editaFuncEstado) & "</div></div>"

			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>FECHA OK CLIENTE</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Fecha OK cliente</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value'>" & HtmlInputCampo("ok_cliente", rs(33), editaFunc And (rs(11) & "" = "OK" And rs(12) & "" = "OK")) & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value'>" & HtmlInputCampo("ok_cliente", rs(33), editaFunc And (rs(11) & "" = "OK" And rs(12) & "" = "OK")) & "</div></div>"

			'[CG-ANT] If Cstr(wPermisos) = "1" Then
			'[CG-ANT] '[CG-ANT] test = "<a class='col-link' style='margin-left: 10px; font-weight: 600; text-decoration: none;' href='" & Server.HTMLEncode(rs("ruta") & "") & "'>Test</a>"
			'[CG-ANT] test = "<a class='cg-link' href='" & Server.HTMLEncode(rs("ruta") & "") & "'>Test</a>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>LINK DOCUMENTO</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Link documento</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value' style='display: flex; align-items: center;'><input class='input-moderno ui-input' id='ruta' name='ruta' type='text' value='" & Server.HTMLEncode(rs("ruta") & "") & "'>" & test & "</div>"
			'[CG-ANT] html = html & "<div class='cg-field__value cg-field__value--inline'><input class='cg-input' id='ruta' name='ruta' type='text' value='" & Server.HTMLEncode(rs("ruta") & "") & "'>" & test & "</div></div>"
			'[CG-ANT] Else
			'[CG-ANT] xruta="http://" & replace(rs("ruta") & "","\\","")
			'[CG-ANT] xruta=replace(xruta,"\","/")
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>LINK DOCUMENTO</div>"
			'[CG-ANT] html = html & "<div><div class='cg-field__label'>Link documento</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value' style='word-break:break-all;'><a href='" & Server.HTMLEncode(xruta) & "' target='_new' style='text-decoration:none;'><span style='color:#FF0000; cursor:pointer;'>" & Server.HTMLEncode(Left(rs("ruta") & "",30)) & "</span></a></div>"
			'[CG-ANT] html = html & "<div class='cg-field__value cg-field__value--break'><a class='cg-link' href='" & Server.HTMLEncode(xruta) & "' target='_new'>" & Server.HTMLEncode(Left(rs("ruta") & "",30)) & "</a></div></div>"
			'[CG-ANT] End If

			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td colspan=2>&nbsp;</td>"
			'[CG-ANT] 'html = html & "</tr>"

			'[CG-ANT] ' Cerramos la grilla de datos acN porque "RELACIONAR APL" rompe las 2 columnas
			'[CG-ANT] html = html & "</div>"

			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td id='sAgre' class='fil' colspan=2 align=center><font style='color:#FF0000;'>RELACIONAR APL</font></td>"
			'[CG-ANT] 'html = html & "</tr>"
			'[CG-ANT] ' Sacamos el <font> obsoleto y usamos estilos CSS modernos
			'[CG-ANT] '[CG-ANT] html = html & "<div style='text-align:center;margin-top:1.5rem;'><span id='sAgre' style='text-decoration:underline;cursor:pointer;color:blue;'>RELACIONAR APL</span></div>"
			'[CG-ANT] html = html & "<div class='cg-actions-center'><span id='sAgre' class='cg-btn cg-btn--secondary cg-btn--sm'>Relacionar APL</span></div>"

			'[CG-ANT] 'html = html & "</table>"
			'[CG-ANT] 'html = html & "</div>"

			'[CG-ANT] ' Cerramos los contenedores de la tarjeta y la columna izquierda
			'[CG-ANT] html = html & "</div>" ' Cierra bloque-body
			'[CG-ANT] html = html & "</div>" ' Cierra bloque-datos
			'[CG-ANT] html = html & "</div>" ' Cierra columna izquierda (width: 50%)
			'[CG-ANT] html = html & "</div>" ' Cierra columna izquierda (width: 50%)

			'[CG-ANT] ' --- COLUMNA DERECHA ---

			'[CG-ANT] 'html = html & "<div style='height:100%;width:50%;float:right'><br>"
			'[CG-ANT] '[CG-ANT] html = html & "<div style='width: 50%; float: right; padding-left: 1.5rem;'>"
			'[CG-ANT] html = html & "<div class='cg-edit__col'>"

			'[CG-ANT] 'html = html & "<table class='tablDet' style='width:80%;border-style:none;' cellspacing=0 align=left>"
			'[CG-ANT] ' Envolvemos en nuestra tarjeta, pero le sacamos el borde y la sombra para que se vea mNs limpio como en tu captura
			'[CG-ANT] '[CG-ANT] html = html & "<div class='bloque-datos ui-panel'>"
			'[CG-ANT] html = html & "<div class='cg-edit__col'>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='bloque-body ui-panel__body'>"
			'[CG-ANT] html = html & "<div class='cg-card'>"

			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td class='fil' ystyle='font-size:12pt;' style='" & st0 & "padding:5px;height:30px;' ><b><u>DETALLES</u></b></td>"
			'[CG-ANT] 'html = html & "</tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='bloque-header' style='margin-bottom: 0.5rem;'>DETALLES</div>"
			'[CG-ANT] html = html & "<div class='cg-card__title'>Detalle</div>"

			'[CG-ANT] 'if rs(24) = session("USUARIO") and session("FUNC") = "FUNC" then

			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td class='fil' >"
			'[CG-ANT] 'html = html & "<textarea class='txtar' id='descripcion' name='descripcion' type='text' style='height:200px;width:100%;font-family:arial;font-size:9pt;padding:5px;border-style:none;' >" & rs(3) & "</textarea>"
			'[CG-ANT] 'html = html & "</td></tr>"
			'[CG-ANT] ' DEPURADO [CG]: LINK DOCUMENTO duplicado. Ya se muestra en la tarjeta Informacion con la misma condicion. Con los dos activos el form enviaba dos campos "ruta" (id y name repetidos).
			'[CG-ANT] '[CG-ANT] If Cstr(wPermisos) = "1" Then
			'[CG-ANT] '[CG-ANT] test = "<a class='col-link' style='margin-left: 10px; font-weight: 600; text-decoration: none;' href='" & Server.HTMLEncode(rs("ruta") & "") & "'>Test</a>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>LINK DOCUMENTO</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value' style='display: flex; align-items: center;'><input class='input-moderno ui-input' id='ruta' name='ruta' type='text' value='" & Server.HTMLEncode(rs("ruta") & "") & "'>" & test & "</div>"
			'[CG-ANT] '[CG-ANT] Else
			'[CG-ANT] '[CG-ANT] xruta="http://" & replace(rs("ruta") & "","\\","")
			'[CG-ANT] '[CG-ANT] xruta=replace(xruta,"\","/")
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label'>LINK DOCUMENTO</div>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value' style='word-break:break-all;'><a href='" & Server.HTMLEncode(xruta) & "' target='_new' style='text-decoration:none;'><span style='color:#FF0000; cursor:pointer;'>" & Server.HTMLEncode(Left(rs("ruta") & "",30)) & "</span></a></div>"
			'[CG-ANT] '[CG-ANT] End If

			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td class='fil' style='font-size:12pt;'><BR><u><b>PATH</b></u></td>"
			'[CG-ANT] 'html = html & "</tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label' style='font-size: 1.1rem; margin-bottom: 0.5rem;'>PATH</div>"
			'[CG-ANT] html = html & "<div class='cg-field'><div class='cg-field__label'>Path</div>"

			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td class='fil'><BR>"

			'[CG-ANT] ' --- LOGICA VBSCRIPT INTACTA ---
			'[CG-ANT] XPATH="\\" & request.servervariables("SERVER_NAME") &  RIGHT(RUTA,LEN(RUTA)-2)

			'[CG-ANT] 'html = html & XPATH & "proyectos\" & Rs(0)
			'[CG-ANT] 'html = html  & "<BR></td></tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value' style='word-break: break-all; margin-bottom: 2rem;'>" & XPATH & "proyectos\" & Rs(0) & "</div>"
			'[CG-ANT] html = html & "<div class='cg-path'><span class='cg-path__text'>" & XPATH & "proyectos\" & Rs(0) & "</span></div></div>"

			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td class='fil' style='font-size:12pt;'><BR><u><b>ARCHIVOS</b></u></td>"
			'[CG-ANT] 'html = html & "</tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label' style='font-size: 1.1rem; margin-bottom: 0.5rem;'>ARCHIVOS</div>"
			'[CG-ANT] html = html & "<div class='cg-field'><div class='cg-field__label'>Archivos</div>"

			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td class='fil'>"

			'[CG-ANT] ' --- LOGICA VBSCRIPT INTACTA ---
			'[CG-ANT] RUTA=RUTA & "\proyectos\" & Rs(0)

			'[CG-ANT] 'html = html & LISTARDIRTORIO(RUTA,Rs(0),"proyectos")
			'[CG-ANT] 'html = html & "</td></tr>"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-value lista-archivos-moderna'>" & LISTARDIRTORIO(RUTA,Rs(0),"proyectos") & "</div>"
			'[CG-ANT] html = html & "<ul class='cg-files'>" & LISTARDIRTORIO(RUTA,Rs(0),"proyectos") & "</ul></div>"

			'[CG-ANT] ' else

			'[CG-ANT] ' html = html & "<tr>"
			'[CG-ANT] ' html = html & "<td class='fil'>"
			'[CG-ANT] ' html = html & rs(3)
			'[CG-ANT] ' html = html & "</td></tr>"

			'[CG-ANT] ' end if

			'[CG-ANT] 'html = html & "</table><br><br><br><br><br><br><br>"
			'[CG-ANT] ' Cerramos los contenedores en lugar de la tabla y sacamos los <br> innecesarios
			'[CG-ANT] 'html = html & "</div>" ' Cierra bloque-body
			'[CG-ANT] 'html = html & "</div>" ' Cierra bloque-datos
			'[CG-ANT] 'html = html & "</div>" ' Cierra columna derecha (width: 50%)


			'[CG-ANT] ' --- ESPACIADOR FINAL ---

			'[CG-ANT] 'html = html & "<table class='tablDet' style='margin-top:50px;width:80%;border-style:none;' cellspacing=0 align=left>"
			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td class='fil' style='" & st0 & "border-style:none;height:30px;' ><b><u></u></b></td>"
			'[CG-ANT] 'html = html & "</tr>"

			'[CG-ANT] ' Reemplazamos toda la tabla estructural por un simple div espaciador moderno
			'[CG-ANT] '[CG-ANT] html = html & "<div style='margin-top: 3rem; width: 100%;'></div>"
			'[CG-ANT] html = html & "</div><div class='cg-card'><div class='cg-card__title'>Historial de responsables</div>"


			'[CG-ANT] 'if (rs(24) = session("USUARIO") and session("FUNC") = "FUNC") or (rs(25) = session("USUARIO") and session("FUNC") = "TEC") then

			'[CG-ANT] ' --- CAMPO AVANCES Y BOTNN ACTUALIZAR ---

			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td class='fil'>"
			'[CG-ANT] 'html = html & "<textarea class='txtar' id='avances' name='avances' type='text' style='height:50px;width:100%;font-family:arial;font-size:9pt;border-style:none;padding:5px;'>" & rs(35) & "</textarea>"
			'[CG-ANT] 'html = html & "</td>"
			'[CG-ANT] 'html = html & "</tr>"

			'[CG-ANT] ' DEPURADO [CG]: el titulo ahora es el header de la tarjeta (cg-card__title)
			'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label' style='font-size: 1.1rem; margin-bottom: 0.5rem;'>HISTORIAL RESPONSABLES</div>"
			'[CG-ANT] html = html & HtmlHistorialRelainc(Cnn, id, "proy")
			'[CG-ANT] If editaFunc Or editaTec Then
				'[CG-ANT] '[CG-ANT] html = html & "<div class='dato-label' style='font-size: 1.1rem; margin-bottom: 0.5rem;'>AVANCES</div>"
				'[CG-ANT] html = html & "<div class='cg-field cg-field--top'><div class='cg-field__label'>Avances</div>"
				'[CG-ANT] '[CG-ANT] html = html & "<div style='margin-bottom: 2rem;'>" & HtmlCamposAvance(RolResponsableActual()) & "</div>"
				'[CG-ANT] html = html & "<div class='cg-table-scroll'>" & HtmlCamposAvance(RolResponsableActual()) & "</div></div>"
			'[CG-ANT] End If

			'[CG-ANT] ' else

			'[CG-ANT] ' html = html & "<tr>"
			'[CG-ANT] ' html = html & "<td class='fil'>" & rs(35) & "</td>"
			'[CG-ANT] ' html = html & "</tr>"

			'[CG-ANT] ' end if

			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td>&nbsp;</td>"
			'[CG-ANT] 'html = html & "</tr>"

			'[CG-ANT] 'html = html & "<tr>"
			'[CG-ANT] 'html = html & "<td align=left>"
			'[CG-ANT] 'html = html & "<input class='bot ui-button' type='button' value='ACTUALIZAR' onclick=""actualizar('sForm');"">"
			'[CG-ANT] 'html = html & "</td></tr>"

			'[CG-ANT] ' Contenedor del botNn con el mismo estilo dinNmico (hover/click) que el botNn "Agregar"
			'[CG-ANT] '[CG-ANT] html = html & "<div class='ui-edit-actions'>"
			'[CG-ANT] html = html & "<div class='cg-savebar'>"
			'[CG-ANT] '[CG-ANT] html = html & "<input class='ui-button ui-button--primary' type='button' style='padding: 0.6rem 2.5rem; background-color: #3b82f6; color: white; border: none; border-radius: 6px; font-weight: 600; cursor: pointer; transition: all 0.2s ease;' value='ACTUALIZAR' onclick=""actualizar('sForm');"" onmouseover=""this.style.backgroundColor='#2563eb'"" onmouseout=""this.style.backgroundColor='#3b82f6'"" onmousedown=""this.style.transform='scale(0.95)'"" onmouseup=""this.style.transform='scale(1)'"">"
			'[CG-ANT] html = html & "<input class='cg-btn cg-btn--primary cg-btn--lg' type='button' value='Actualizar informaci&oacute;n' onclick=""actualizar('sForm');"">"
			'[CG-ANT] html = html & "</div>"

		'[CG-ANT] End If

		'[CG-ANT] ' DEPURADO [CG]: cierre de </table> sin tabla abierta (resto de la maquetacion vieja)
		'[CG-ANT] '[CG-ANT] html = html & "</table>"
		'[CG-ANT] ' ==========================================
		'[CG-ANT] ' AHORA SN: CERRAMOS LA COLUMNA DERECHA
		'[CG-ANT] ' ==========================================
		'[CG-ANT] html = html & "</div>" ' Cierra bloque-body
		'[CG-ANT] html = html & "</div>" ' Cierra bloque-datos
		'[CG-ANT] html = html & "</div>" ' Cierra columna derecha (width: 50%)

		'[CG-ANT] ' Y finalmente limpiamos los floats para que el contenedor general no colapse
		'[CG-ANT] ' DEPURADO [CG]: ya no hay floats que limpiar; las columnas son una grilla (cg-edit)
		'[CG-ANT] '[CG-ANT] html = html & "<div style='clear: both;'></div>"

		'[CG-ANT] ' Cerramos el contenedor global .inner
		'[CG-ANT] html = html & "</div>"

		'[CG-ANT] html = html & "</form>"

		'[CG-ANT] 'rs.close

	'[CG-ANT] End If

	' =====================================================================
	' [CG] CONTENIDO SEGUN LA PROPUESTA (vista-incidente / vista-proyecto)
	' El bloque anterior completo queda arriba, comentado con '[CG-ANT].
	' Las llamadas de datos, permisos y guardado son las mismas de antes.
	' =====================================================================
	If opc = "veri" Then

		zzz = ObjAcc(id,"I",cnn,"D")
		XXX = SYSLOGON(cnn,"CONSULTA_INC_" & id)

		Set rs = ObtenerDetalleIncidente(Cnn, id)

		If Not rs.eof Then

			' --- Datos calculados igual que en la pantalla anterior ---
			sFec = Fechador(CgFechaDb(rs(11)))

			sRelPro = ""
			If rs(14) <> 0 Then sRelPro = "ID: " & rs(14)

			rp = ""
			If rs(7) & "" = "" Then
				rp = "---"
			Else
				rp = rs(7)
			End If

			od = ""
			If rs(16) & "" = "0" Or rs(16) & "" = "" Then
				od = "---"
			Else
				od = Fechador(CgFechaDb(rs(16)))
			End If

			of = ""
			If rs(17) & "" = "0" Or rs(17) & "" = "" Then
				of = "---"
			Else
				of = Fechador(CgFechaDb(rs(17)))
			End If

			xruta="http://" & replace(rs("ruta") & "","\\","")
			xruta=replace(xruta,"\","/")

			' --- Barra superior: identificacion, datos clave y semaforo ---
			cgFIni = CgFechaDb(rs(11))
			cgFLim = Empty
			If Not IsEmpty(cgFIni) And CgPlazoDias(rs(5)) > 0 Then cgFLim = DateAdd("d", CgPlazoDias(rs(5)), cgFIni)
			cgMetas = CgMeta("Cliente", CgTexto(rs(1))) & CgMeta("Fecha creaci&oacute;n", CgTexto(sFec)) & CgMeta("Prioridad", CgPrioridadPill(rs(5)))
			html = html & CgTopbar(cgTopFila, cgMetas, CgSemaforo(cgFIni, cgFLim, rs(18), "plazo de la prioridad"))

			html = html & "<div class='ui-app ui-page'><div class='ui-container'>"
			' [UI] Vista en dos columnas: Informacion y Detalle a la izquierda, historial a la derecha
			'[UI-ANT] html = html & "<div class='ui-det-grid'><div class='ui-det-grid__principal'>"
			' [UI] Lo construido hasta aca se guarda; las tarjetas se arman y despues se ordenan
			detPrevio = html
			html = ""

			' --- Tarjeta Informacion (lo que no esta en la barra superior) ---
			html = html & "<div class='ui-card'>" & CgCardTitulo("Informaci&oacute;n") & "<div class='ui-fields'>"
			html = html & CgCampo("Contacto", "<a class='ui-link' href='detalleCont.asp?idcon=" & rs(15) & "&id=" & rs(0) & "&opc=veri&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>" & Server.HTMLEncode(rs(2) & "") & "</a>")
			html = html & CgCampo("Aplicaci&oacute;n", CgTexto(rs(3)))
			html = html & CgCampo("Ambiente", AmbienteSelectHtml("", AmbienteValor(rs), False))
			html = html & CgCampo("V&iacute;a pedido", CgTexto(rs(13)))
			html = html & CgCampo("Browser", CgTexto(rs(12)))
			If sRelPro = "" Then
				html = html & CgCampo("Relacionado a proy", "<span class='ui-field__value--muted'>Sin relaci&oacute;n</span>")
			Else
				html = html & CgCampo("Relacionado a proy", "<a class='ui-link' href='detalles.asp?id=" & rs(14) & "&opc=verp&sol=" & rs(25) & "&xmenu=10&xtmenu=INICIO'>" & sRelPro & "</a>")
			End If
			html = html & CgCampo("Requiere presupuesto", CgTexto(rs(8)))
			html = html & CgCampo("Estado funcional", CgBadgeEstado(rs(10)))
			html = html & CgCampo("Estado desarrollo", CgBadgeEstado(rs(9)))
			html = html & CgCampo("Responsable gesti&oacute;n", CgTexto(rs(6)))
			html = html & CgCampo("Responsable t&eacute;cnico", CgTexto(rp))
			html = html & CgCampo("OK desarrollo", CgTexto(od))
			html = html & CgCampo("OK funcional", CgTexto(of))
			sFecOkCli = "---"
			If rs(18) & "" <> "0" And rs(18) & "" <> "" Then sFecOkCli = Fechador(CgFechaDb(rs(18)))
			html = html & CgCampo("Fecha OK cliente", CgTexto(sFecOkCli))
			' [UI] Estado del cliente quitado: se decidio no sumar este dato (pedido de Luis)
			'[UI-ANT] html = html & CgCampo("Estado cliente",  CgBadgeEstado(rs(27))) 'CgFalta("estado del cliente"))
			html = html & CgCampo("Tiempo funcional", CgTexto(rs(28))) 'CgFalta("tiempo funcional"))
			html = html & CgCampo("Tiempo desarrollo", CgTexto(rs(29))) 'CgFalta("tiempo de desarrollo"))
			If CgHayDato(rs("ruta")) Then
				html = html & CgCampoMod("Link documento", "<a class='ui-link' href='" & Server.HTMLEncode(xruta) & "' target='_new'>" & Server.HTMLEncode(Left(rs("ruta") & "",30)) & "</a>", "ui-field__value--break")
			Else
				html = html & CgCampo("Link documento", "<span class='ui-field__value--muted'>Sin documento asociado</span>")
			End If
			html = html & "</div></div>"

			' [UI] Informacion se guarda aparte: va debajo, a todo el ancho
			detInfo = html
			html = ""
			' --- Tarjeta Detalle ---
			XPATH="\\" & request.servervariables("SERVER_NAME") &  RIGHT(RUTA,LEN(RUTA)-2)
			RUTA=RUTA & "\INCIDENTES\" & Rs(0)
			html = html & "<div class='ui-card'>" & CgCardTitulo("Detalle")
			'[UI-ANT] html = html & "<div class='ui-field'><div class='ui-field__label'>Descripci&oacute;n</div><div class='ui-prose'>" & Replace(Server.HTMLEncode(rs(4) & ""),Chr(13),"<br>") & "</div></div>"
			' [UI] Descripcion recortada (14 renglones); si queda recortada aparece 'Ver detalle completo', que abre el pop-up de lectura
			html = html & "<div class='ui-field'><div class='ui-field__label'>Descripci&oacute;n</div><div class='ui-prose ui-recorte' id='detTexto'>" & Replace(Server.HTMLEncode(rs(4) & ""),Chr(13),"<br>") & "</div>" & _
				"<button type='button' class='ui-recorte__ver' data-recorte='detTexto' hidden onclick=""UiVerTexto('Detalle', 'detTexto', this)"">Ver detalle completo</button></div>" & HtmlDialogoLectura()
			html = html & CgPath(XPATH & "INCIDENTES\" & Rs(0))
			html = html & "<div class='ui-field'><div class='ui-field__label'>Archivos</div><ul class='ui-files'>" & LISTARDIRTORIO(RUTA,Rs(0),"INCIDENTES") & "</ul></div>"
			html = html & "</div>"
			' [UI] Columna derecha: historial de avances (con el pop-up para anadir un avance)
			'[UI-ANT] html = html & "</div><aside class='ui-det-grid__lateral' aria-label='Historial de avances'>" & DetTarjetaHistorial(rs(0), "inc", "incidentes", "avainc") & "</aside></div>"
			' [UI] Orden de la vista: historial (izquierda) y Detalle (derecha) con la misma altura; debajo, Informacion a todo el ancho
			detDetalle = html
			'[UI-ANT] html = detPrevio & "<div class='ui-det-grid'><div class='ui-det-grid__principal'>" & DetTarjetaHistorial(rs(0), "inc", "incidentes", "avainc") & "</div><div class='ui-det-grid__lateral'>" & detDetalle & "</div></div>" & detInfo
			' [UI] Detalle a la izquierda (columna angosta) y Avances a la derecha (columna ancha), cada uno con su ancho de antes
			html = detPrevio & "<div class='ui-det-grid'><div class='ui-det-grid__principal'>" & detDetalle & "</div><div class='ui-det-grid__lateral'>" & DetTarjetaHistorial(rs(0), "inc", "incidentes", "avainc") & "</div></div>" & detInfo

			' --- Tarjeta Historial de avances ---
			' (El texto libre de avances viejo, rs(21), quedo sin mostrar a pedido)
'[UI-ANT] html = html & "<div class='ui-card'>" & CgCardHeader("Historial de avances", CgBotonPendiente("A&ntilde;adir", "Falta conectar: alta de avances desde esta pantalla. Hoy se cargan desde Editar."))
' [UI] Anadir avance desde la vista (pop-up). Mismos campos y nombres que la edicion (HtmlCamposAvance):
' texto, fecha y hora (por defecto, el momento actual), responsable (por defecto, el rol del usuario) y tipo.
' Lo guarda actualizar_v2.asp (opt=avainc) con RegistrarAvance, igual que la edicion.
'[UI-ANT] Dim avPuede, avHoy, avHora, avRol, rsTA, avTipos
'[UI-ANT] avPuede = (getPermisos(xmenu, xtmenu, "MODIF") = 1)
'[UI-ANT] If avPuede Then
	'[UI-ANT] html = html & "<div class='ui-card'>" & CgCardHeader("Historial de avances", "<button type='button' class='ui-button ui-button--light ui-button--sm' id='btnAvance' aria-haspopup='dialog' onclick='abrirAvance()'>+ A&ntilde;adir avance</button>")
'[UI-ANT] Else
	' (detalles.asp no incluye ui_componentes.inc: el boton deshabilitado va escrito aca)
	'[UI-ANT] html = html & "<div class='ui-card'>" & CgCardHeader("Historial de avances", "<button type='button' class='ui-button ui-button--light ui-button--sm ui-button--disabled' aria-disabled='true' title='Sin permiso para modificar el incidente.'>+ A&ntilde;adir avance</button>")
'[UI-ANT] End If
'[UI-ANT] If Request.QueryString("av") = "ok" Then html = html & "<div class='ui-alert ui-alert--success' role='status'>Avance agregado.</div>"
'[UI-ANT] If Request.QueryString("av") = "error" Then html = html & "<div class='ui-alert ui-alert--danger' role='alert'>No se pudo agregar el avance: complet&aacute; el texto y verific&aacute; tu permiso para modificar el incidente.</div>"
'[UI-ANT] If avPuede Then
	'[UI-ANT] avHoy = Year(Date()) & "-" & Right("0" & Month(Date()), 2) & "-" & Right("0" & Day(Date()), 2)
	'[UI-ANT] avHora = Right("0" & Hour(Time()), 2) & ":" & Right("0" & Minute(Time()), 2)
	'[UI-ANT] avRol = RolResponsableActual()
	'[UI-ANT] avTipos = "<option value=''>Seleccion&aacute; un tipo</option>"
	'[UI-ANT] Set rsTA = GetListaComboExtended(Cnn, "tipo_avance")
	'[UI-ANT] Do While Not rsTA.EOF
		'[UI-ANT] avTipos = avTipos & "<option value='" & rsTA(0) & "'>" & Server.HTMLEncode(rsTA(1) & "") & "</option>"
		'[UI-ANT] rsTA.MoveNext
	'[UI-ANT] Loop
	'[UI-ANT] rsTA.Close
	'[UI-ANT] html = html & "<dialog id='dlgAvance' class='ui-dialogo' aria-labelledby='dlgAvanceTitulo'>"
	'[UI-ANT] html = html & "<form method='post' action='actualizar_v2.asp?tabla=incidentes&opt=avainc&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&cantregmover=0'>"
	'[UI-ANT] html = html & "<div class='ui-dialogo__head'><h2 class='ui-dialogo__titulo' id='dlgAvanceTitulo'>Nuevo avance</h2><button type='button' class='ui-dialogo__cerrar' aria-label='Cerrar' onclick='cerrarAvance()'>&times;</button></div>"
	'[UI-ANT] html = html & "<div class='ui-dialogo__body'>"
	'[UI-ANT] html = html & "<div class='ui-field'><label class='ui-field__label' for='avance'>Detalle del avance <span class='ui-field__required' aria-hidden='true'> *</span></label><textarea class='ui-input ui-textarea' id='avance' name='avance' rows='5' maxlength='4000' required aria-required='true'></textarea></div>"
	'[UI-ANT] html = html & "<div class='ui-dialogo__grid'>"
	'[UI-ANT] html = html & "<div class='ui-field'><label class='ui-field__label' for='avance_fecha'>Fecha</label><input class='ui-input' id='avance_fecha' name='avance_fecha' type='date' value='" & avHoy & "'></div>"
	'[UI-ANT] html = html & "<div class='ui-field'><label class='ui-field__label' for='avance_hora'>Hora</label><input class='ui-input' id='avance_hora' name='avance_hora' type='time' value='" & avHora & "'></div>"
	'[UI-ANT] html = html & "<div class='ui-field'><label class='ui-field__label' for='responsable_avance'>Responsable</label><select class='ui-select' id='responsable_avance' name='responsable_avance'>"
	'[UI-ANT] html = html & "<option value='FUNCIONAL'" & IIf(avRol = "FUNCIONAL", " selected", "") & ">Funcional</option>"
	'[UI-ANT] html = html & "<option value='TECNICO'" & IIf(avRol = "TECNICO", " selected", "") & ">T&eacute;cnico</option>"
	'[UI-ANT] html = html & "<option value='CLIENTE'" & IIf(avRol = "CLIENTE", " selected", "") & ">Cliente</option></select></div>"
	'[UI-ANT] html = html & "<div class='ui-field'><label class='ui-field__label' for='tipo_avance'>Tipo de avance</label><select class='ui-select' id='tipo_avance' name='tipo_avance'>" & avTipos & "</select></div>"
	'[UI-ANT] html = html & "</div></div>"
	'[UI-ANT] html = html & "<div class='ui-dialogo__acciones'><button type='button' class='ui-button ui-button--secondary' onclick='cerrarAvance()'>Cancelar</button><button type='submit' class='ui-button ui-button--primary'>Guardar avance</button></div>"
	'[UI-ANT] html = html & "</form></dialog>"
	'[UI-ANT] html = html & "<script>function abrirAvance(){var d=document.getElementById('dlgAvance');d.showModal();document.getElementById('avance').focus();}"
	'[UI-ANT] html = html & "function cerrarAvance(){document.getElementById('dlgAvance').close();}"
	'[UI-ANT] html = html & "document.getElementById('dlgAvance').addEventListener('close',function(){var b=document.getElementById('btnAvance');if(b)b.focus();});"
	'[UI-ANT] html = html & "document.getElementById('dlgAvance').addEventListener('click',function(e){if(e.target===this)cerrarAvance();});</script>"
'[UI-ANT] End If
			'[UI-ANT] html = html & CgFaltaBloque("tipo de avance (Anotaci&oacute;n, Solicitud, Insistencia), estado, testing y recepci&oacute;n de solicitudes. Abajo se muestra el historial que existe hoy.")
			'[UI-ANT] html = html & HtmlHistorialRelainc(Cnn, rs(0), "inc")
			'[UI-ANT] html = html & "</div>"

			' --- Tarjeta APLs relacionados ---
			'[UI-ANT] Set rz = ObtenerAplRelac(Cnn, id, "I")
			'[UI-ANT] html = html & "<div class='ui-card'>" & CgCardHeader("APLs relacionados", CgBotonPendiente("Relacionar APL", "Por ahora los APLs se relacionan desde Editar."))
			'[UI-ANT] html = html & CgTablaAplAbre()
			'[UI-ANT] If Not rz.eof Then
				'[UI-ANT] Do While Not rz.eof
					'[UI-ANT] html = html & CgFilaApl(rz)
					'[UI-ANT] rz.movenext
				'[UI-ANT] Loop
			'[UI-ANT] Else
				'[UI-ANT] html = html & "<tr><td colspan='6'><div class='ui-empty'>No hay APLs relacionados.</div></td></tr>"
			'[UI-ANT] End If
			'[UI-ANT] html = html & "</tbody></table></div></div>"
		' [UI] APLs relacionados: tarjeta con el pop-up 'Relacionar APL' (solo pide el ID del APL)
		html = html & DetTarjetaApl(id, "I", "incidentes")

			html = html & "</div></div>"

		End If

		rs.close

	ElseIf opc = "verp" Then

		zzz = ObjAcc(id,"P",cnn,"D")
		XXX = SYSLOGON(cnn,"CONSULTA_PROY_" & id)

		Set rs = ObtenerDetalleProyecto(Cnn, id)

		If Not rs.eof Then

			' --- Datos calculados igual que en la pantalla anterior ---
			sFec = ""
			If rs(9) <> 0 Then sFec = Fechador(CgFechaDb(rs(9)))
			sFecEn = ""
			If rs(10) <> 0 Then sFecEn = Fechador(CgFechaDb(rs(10)))

			od = ""
			If rs(23) & "" = "0" Or rs(23) & "" = "" Then
				od = "---"
			Else
				od = Fechador(CgFechaDb(rs(23)))
			End If

			of = ""
			If rs(24) & "" = "0" Or rs(24) & "" = "" Then
				of = "---"
			Else
				of = Fechador(CgFechaDb(rs(24)))
			End If

			' --- Barra superior: en proyectos el semaforo se rige por la fecha de entrega ---
			cgFIni = CgFechaDb(rs(9))
			cgFLim = CgFechaDb(rs(10))
			If CgHayDato(sFecEn) Then
				cgEntrega = Server.HTMLEncode(sFecEn)
			Else
				cgEntrega = CgFalta("fecha de entrega")
			End If
			cgMetas = CgMeta("Cliente", CgTexto(rs(4))) & CgMeta("Fecha creaci&oacute;n", CgTexto(sFec)) & CgMeta("Fecha entrega", cgEntrega) & CgMeta("Prioridad", CgPrioridadPill(rs(13)))
			html = html & CgTopbar(cgTopFila, cgMetas, CgSemaforo(cgFIni, cgFLim, rs(25), "fecha de entrega"))

			html = html & "<div class='ui-app ui-page'><div class='ui-container'>"
			' [UI] Vista en dos columnas: Informacion y Detalle a la izquierda, historial a la derecha
			'[UI-ANT] html = html & "<div class='ui-det-grid'><div class='ui-det-grid__principal'>"
			' [UI] Lo construido hasta aca se guarda; las tarjetas se arman y despues se ordenan
			detPrevio = html
			html = ""

			' --- Tarjeta Informacion ---
			html = html & "<div class='ui-card'>" & CgCardTitulo("Informaci&oacute;n") & "<div class='ui-fields'>"
			html = html & CgCampo("Solicitud n&uacute;mero", CgTexto(rs(1)))
			html = html & CgCampo("Contacto", "<a class='ui-link' href='detalleCont.asp?id=" & id & "&opc=" & opc & "&sol=" & sol & "&idcon=" & rs(22) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>" & Server.HTMLEncode(rs(5) & "") & "</a>")
			html = html & CgCampo("Aplicaci&oacute;n", CgTexto(rs(6)))
			html = html & CgCampo("V&iacute;a pedido", CgTexto(rs(19)))
			html = html & CgCampo("Browser", CgTexto(rs(18)))
			html = html & CgCampo("Req presupuesto", CgTexto(rs(20)))
			html = html & CgCampo("Estado funcional", CgBadgeEstado(rs(12)))
			html = html & CgCampo("Estado desarrollo", CgBadgeEstado(rs(11)))
			html = html & CgCampo("Respon gesti&oacute;n", CgTexto(rs(7)))
			html = html & CgCampo("Respon t&eacute;cnico", CgTexto(rs(8)))
			html = html & CgCampo("OK desarrollo", CgTexto(od))
			html = html & CgCampo("OK funcional", CgTexto(of))
			sFecOkCli = "---"
			If rs(25) & "" <> "0" And rs(25) & "" <> "" Then sFecOkCli = Fechador(CgFechaDb(rs(25)))
			html = html & CgCampo("OK cliente", CgTexto(sFecOkCli))
			' [UI] Estado del cliente quitado: se decidio no sumar este dato (pedido de Luis)
			'[UI-ANT] html = html & CgCampo("Estado cliente", CgFalta("estado del cliente"))
			If CgHayDato(rs("ruta")) Then
				html = html & CgCampoMod("Link documento", "<a class='ui-link' href='" & Server.HTMLEncode(rs("ruta") & "") & "' target='_new'>" & Server.HTMLEncode(Left(rs("ruta") & "",30)) & "</a>", "ui-field__value--break")
			Else
				html = html & CgCampo("Link documento", "<span class='ui-field__value--muted'>Sin documento asociado</span>")
			End If
			html = html & "</div></div>"

			' --- Tarjeta Carga de horas ---
			html = html & "<div class='ui-card'>" & CgCardTitulo("Carga de horas") & "<div class='ui-fields'>"
			html = html & CgCampo("Relev y definici&oacute;n", CgHoras(rs(14)))
			html = html & CgCampo("Dise&ntilde;o", CgHoras(rs(15)))
			html = html & CgCampo("Desarrollo", CgHoras(rs(16)))
			html = html & CgCampo("Testeo", CgHoras(rs(17)))
			html = html & CgCampoMod("Total", CgHoras(rs(21)), "ui-hours__total")
			html = html & "</div>" & CgBarraHoras(rs(14), rs(15), rs(16), rs(17)) & "</div>"

			' [UI] Informacion se guarda aparte: va debajo, a todo el ancho
			detInfo = html
			html = ""
			' --- Tarjeta Detalle ---
			XPATH="\\" & request.servervariables("SERVER_NAME") &  RIGHT(RUTA,LEN(RUTA)-2)
			RUTA=RUTA & "\proyectos\" & Rs(0)
			html = html & "<div class='ui-card'>" & CgCardTitulo("Detalle")
			'[UI-ANT] html = html & "<div class='ui-field'><div class='ui-field__label'>Descripci&oacute;n</div><div class='ui-prose'>" & Replace(Server.HTMLEncode(rs(3) & ""),Chr(13),"<br>") & "</div></div>"
			' [UI] Descripcion recortada (14 renglones); si queda recortada aparece 'Ver detalle completo', que abre el pop-up de lectura
			html = html & "<div class='ui-field'><div class='ui-field__label'>Descripci&oacute;n</div><div class='ui-prose ui-recorte' id='detTexto'>" & Replace(Server.HTMLEncode(rs(3) & ""),Chr(13),"<br>") & "</div>" & _
				"<button type='button' class='ui-recorte__ver' data-recorte='detTexto' hidden onclick=""UiVerTexto('Detalle', 'detTexto', this)"">Ver detalle completo</button></div>" & HtmlDialogoLectura()
			html = html & CgPath(XPATH & "proyectos\" & Rs(0))
			html = html & "<div class='ui-field'><div class='ui-field__label'>Archivos</div><ul class='ui-files'>" & LISTARDIRTORIO(RUTA,Rs(0),"proyectos") & "</ul></div>"
			html = html & "</div>"
			' [UI] Columna derecha: historial de avances (con el pop-up para anadir un avance)
			'[UI-ANT] html = html & "</div><aside class='ui-det-grid__lateral' aria-label='Historial de avances'>" & DetTarjetaHistorial(id, "proy", "proyectos", "avaproy") & "</aside></div>"
			' [UI] Orden de la vista: historial (izquierda) y Detalle (derecha) con la misma altura; debajo, Informacion a todo el ancho
			detDetalle = html
			'[UI-ANT] html = detPrevio & "<div class='ui-det-grid'><div class='ui-det-grid__principal'>" & DetTarjetaHistorial(id, "proy", "proyectos", "avaproy") & "</div><div class='ui-det-grid__lateral'>" & detDetalle & "</div></div>" & detInfo
			' [UI] Detalle a la izquierda (columna angosta) y Avances a la derecha (columna ancha), cada uno con su ancho de antes
			html = detPrevio & "<div class='ui-det-grid'><div class='ui-det-grid__principal'>" & detDetalle & "</div><div class='ui-det-grid__lateral'>" & DetTarjetaHistorial(id, "proy", "proyectos", "avaproy") & "</div></div>" & detInfo

			' --- Tarjeta Historial de avances ---
			' (El texto libre de avances viejo, rs(27), no se muestra: mismo criterio que en incidentes)
			'[UI-ANT] html = html & "<div class='ui-card'>" & CgCardHeader("Historial de avances", CgBotonPendiente("A&ntilde;adir", "Falta conectar: alta de avances desde esta pantalla. Hoy se cargan desde Editar."))
			'[UI-ANT] html = html & CgFaltaBloque("tipo de avance (Anotaci&oacute;n, Solicitud, Insistencia), estado, testing y recepci&oacute;n de solicitudes. Abajo se muestra el historial que existe hoy.")
			'[UI-ANT] html = html & HtmlHistorialRelainc(Cnn, id, "proy")
			'[UI-ANT] html = html & "</div>"

			' --- Tarjeta APLs relacionados ---
			'[UI-ANT] Set rz = ObtenerAplRelac(Cnn, id, "P")
			'[UI-ANT] html = html & "<div class='ui-card'>" & CgCardHeader("APLs relacionados", CgBotonPendiente("Relacionar APL", "Por ahora los APLs se relacionan desde Editar."))
			'[UI-ANT] html = html & CgTablaAplAbre()
			'[UI-ANT] If Not rz.eof Then
				'[UI-ANT] Do While Not rz.eof
					'[UI-ANT] html = html & CgFilaApl(rz)
					'[UI-ANT] rz.movenext
				'[UI-ANT] Loop
			'[UI-ANT] Else
				'[UI-ANT] html = html & "<tr><td colspan='6'><div class='ui-empty'>No hay APLs relacionados.</div></td></tr>"
			'[UI-ANT] End If
			'[UI-ANT] html = html & "</tbody></table></div></div>"
		' [UI] APLs relacionados: tarjeta con el pop-up 'Relacionar APL' (solo pide el ID del APL)
		html = html & DetTarjetaApl(id, "P", "proyectos")

			' --- Tarjeta Incidentes relacionados ---
			Set rz = ObtenerIncidentesPorProyecto(Cnn, id)
			html = html & "<div class='ui-card'>" & CgCardTitulo("Incidentes relacionados") & "<div class='ui-list'>"
			If Not rz.eof Then
				Do While Not rz.eof
					html = html & "<div class='ui-list__item'><a class='ui-link' href='detalles.asp?opc=veri&id=" & rz(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>INC-" & rz(0) & "</a></div>"
					rz.movenext
				Loop
			Else
				html = html & "<div class='ui-empty'>No hay incidentes relacionados.</div>"
			End If
			html = html & "</div></div>"

			html = html & "</div></div>"

		End If

	ElseIf opc = "acti" Then

		zzz = ObjAcc(id,"I",cnn,"A")
		XXX = SYSLOGON(cnn,"ACTUALIZA_INC_" & id)

		Dim wPermisos, editaFunc, editaTec, editaFuncEstado
		wPermisos = getPermisosAtrm("CG_INCIDENTES","MODIF")
		editaFunc = (CStr(wPermisos) = "1" And session("FUNC") = "FUNC")
		editaTec = (CStr(wPermisos) = "1" And session("FUNC") = "TEC")
		Set rx = ObtenerIncidenteEdicion(Cnn, id) ' TODO: revisar/crear SP sugerido sp_ObtenerIncidenteEdicion

		' --- Datos calculados igual que en la pantalla anterior ---
		sFec = Fechador(CgFechaDb(rx(19)))
		editaFuncEstado = (editaFunc And (rx(16) & "" = "OK" Or rx(16) & "" = "BAJA" Or rx(16) & "" = "---"))
		test = "<a class='ui-link' href='" & Server.HTMLEncode(rx("ruta") & "") & "'>Test</a>"

		' rx(25) (ok_cliente) viene en formato numerico (dias desde 1900), igual que fecha_entrega
		' en la edicion de proyectos: se formatea con Fechador antes de usarlo en el input editable.
		sFecOkCli = ""
		If rx(25) & "" <> "0" And rx(25) & "" <> "" Then sFecOkCli = Fechador(CgFechaDb(rx(25)))

		' rx (ObtenerIncidenteEdicion) no tiene mapeados ok_desarrollo/ok_funcional en un indice
		' confirmado; se reconsulta con ObtenerDetalleIncidente (misma fuente que usa la vista de
		' solo lectura en rs(16)/rs(17)) solo para sacar esas dos fechas, mismo criterio que ahi:
		' "0" = sin fecha cargada -> "---", si no Fechador(valor + nFecini).
		Set rzFechasOk = ObtenerDetalleIncidente(Cnn, id)
		od = "---"
		of = "---"
		If Not rzFechasOk.eof Then
			If rzFechasOk(16) & "" <> "0" And rzFechasOk(16) & "" <> "" Then od = Fechador(CgFechaDb(rzFechasOk(16)))
			If rzFechasOk(17) & "" <> "0" And rzFechasOk(17) & "" <> "" Then of = Fechador(CgFechaDb(rzFechasOk(17)))
		End If
		rzFechasOk.Close

		' El form abre antes de la barra superior: titulo y cliente se editan ahi y tienen que viajar en el form
		html = html & "<form id='sForm' method='post' action='actualizar_v2.asp?tabla=incidentes&relainc=relainc&id=" & id & "&opt=actinc&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=detalles.asp?opc=acti' >"

		' --- Barra superior editable ---
		' Prioridad: rx(7) es el id y rx(9) la descripcion (Alta/Media/Baja); el pill y el plazo usan la descripcion
		cgTopFila = Replace(cgTopFila, "<!--CG_TITULO-->", "<div class='ui-topbar__title-field'>" & HtmlInputCampo("titulo", rx(28), editaFunc) & "</div>")
		cgFIni = CgFechaDb(rx(19))
		cgFLim = Empty
		If Not IsEmpty(cgFIni) And CgPlazoDias(rx(9)) > 0 Then cgFLim = DateAdd("d", CgPlazoDias(rx(9)), cgFIni)
		cgTopEdit = True
		cgMetas = CgMeta("Cliente", ComboCampo(Cnn, rx(1), "id_cliente", "id_cliente", editaFunc)) & CgMeta("Fecha creaci&oacute;n", CgTexto(sFec)) & CgMeta("Prioridad", CgPrioridadPill(rx(9)))
		html = html & CgTopbar(cgTopFila, cgMetas, CgSemaforo(cgFIni, cgFLim, rx(25), "plazo de la prioridad"))

		html = html & "<div class='ui-app ui-page ui-page--edit'><div class='ui-container'>"

		' --- Tarjeta Informacion ---
		html = html & "<div class='ui-card'>" & CgCardTitulo("Informaci&oacute;n") & "<div class='ui-fields'>"
		html = html & CgCampo("Contacto", ComboCampo(Cnn, rx(3), "contacto", "contacto", editaFunc))
		html = html & CgCampo("Aplicaci&oacute;n", ComboCampo(Cnn, rx(5), "id_aplic", "id_aplic", editaFunc))
		html = html & CgCampo("Ambiente", AmbienteSelectHtml("ambiente", AmbienteValor(rx), (editaFunc Or editaTec)))
		html = html & CgCampo("V&iacute;a pedido", ComboCampo(Cnn, rx(22), "via_pedido", "via_pedido", editaFunc))
		html = html & CgCampo("Browser", ComboCampo(Cnn, rx(20), "browser", "browser", editaFunc))
		html = html & CgCampo("Relacionado a proyecto", ComboProyectoCampo(Cnn, rx(24), rx(1), editaFunc))   ' rx(24) = proyecto del incidente (puede estar vacio); rx(1) = cliente (filtro)
		html = html & CgCampo("Requiere presupuesto", HtmlSelectSiNo("req_presup", rx(14), editaFunc))
		html = html & CgCampo("Estado funcional", ComboCampo(Cnn, rx(17), "estado_func", "estado_func", editaFuncEstado))
		html = html & CgCampo("Estado desarrollo", ComboCampo(Cnn, rx(15), "estado_desa", "estado_desa", editaTec))
		html = html & CgCampo("Responsable gesti&oacute;n", ComboCampo(Cnn, rx(10), "respon_gestion", "respon_gestion", editaFunc))
		html = html & CgCampo("Responsable t&eacute;cnico", ComboCampo(Cnn, rx(12), "respon_tec", "respon_tec", editaTec))
		'[UI-ANT] html = html & CgCampo("Tiempo funcional", CgTexto(of))
		'[UI-ANT] html = html & CgCampo("Tiempo desarrollo", CgTexto(od))
		' [UI] Las fechas de OK desarrollo y OK funcional van con su propio nombre, igual que en la vista.
		' [UI] Los tiempos funcional y de desarrollo (minutos) siguen pendientes de conectar.
		html = html & CgCampo("OK desarrollo", CgTexto(od)) ' [UI]
		html = html & CgCampo("OK funcional", CgTexto(of)) ' [UI]
		html = html & CgCampo("Fecha OK cliente", HtmlInputCampo("ok_cliente", sFecOkCli, editaFunc And (rx(16) & "" = "OK" And rx(18) & "" = "OK")))
		' [UI] Estado del cliente quitado: se decidio no sumar este dato (pedido de Luis)
		'[UI-ANT] html = html & CgCampo("Estado cliente", CgFalta("estado del cliente"))
		html = html & CgCampo("Tiempo funcional", CgTexto(of))
		html = html & CgCampo("Tiempo desarrollo", CgTexto(od))
		html = html & "<div class='ui-field'><div class='ui-field__label'>Link documento</div><div class='ui-field__value'><input class='ui-input' id='ruta' name='ruta' type='text' value='" & Server.HTMLEncode(rx("ruta") & "") & "'>" & test & "</div><div class='ui-field__hint'>Ejemplo: 'http://server-40/registroincidente/' + nombre del documento</div></div>"
		html = html & "</div></div>"

		' --- Tarjeta Detalle ---
		XPATH="\\" & request.servervariables("SERVER_NAME") &  RIGHT(RUTA,LEN(RUTA)-2)
		RUTA=RUTA & "\INCIDENTES\" & RX(0)
		html = html & "<div class='ui-card'>" & CgCardTitulo("Detalle")
		html = html & "<div class='ui-field'><div class='ui-field__label'>Descripci&oacute;n</div>" & HtmlTextAreaCampo("detalles", rx(8), "200px", editaFunc) & "</div>"
		html = html & CgPath(XPATH & "INCIDENTES\" & RX(0))
		html = html & "<div class='ui-field'><div class='ui-field__label'>Archivos</div><ul class='ui-files'>" & LISTARDIRTORIO(RUTA,RX(0),"INCIDENTES") & "</ul></div>"
		html = html & "</div>"

		' [UI] Edicion: solo Informacion y Detalle. El historial (con el alta de avances) y los APLs se ven y se cargan desde la vista.
		' --- Tarjeta Historial de avances ---
		'[UI-ANT] html = html & "<div class='ui-card'>" & CgCardTitulo("Historial de avances")
		'[UI-ANT] html = html & CgFaltaBloque("tipo de avance (Anotaci&oacute;n, Solicitud, Insistencia), estado, testing y recepci&oacute;n de solicitudes. Por ahora el avance se carga en los campos de abajo y se guarda con Actualizar informaci&oacute;n.")
		'[UI-ANT] html = html & HtmlHistorialRelainc(Cnn, id, "inc")
		'[UI-ANT] If editaFunc Or editaTec Then
			'[UI-ANT] html = html & "<div class='ui-field ui-field--top'><div class='ui-field__label'>Nuevo avance</div><div class='ui-table-wrap'>" & HtmlCamposAvance(RolResponsableActual()) & "</div></div>"
		'[UI-ANT] End If
		'[UI-ANT] html = html & "</div>"

		' --- Tarjeta APLs relacionados (el id sAgre abre el dialogo de siempre) ---
		'[UI-ANT] Set rz = ObtenerAplRelac(Cnn, id, "I")
		'[UI-ANT] html = html & "<div class='ui-card'>" & CgCardHeader("APLs relacionados", "<span id='sAgre' class='ui-button ui-button--secondary ui-button--sm'>Relacionar APL</span>")
		'[UI-ANT] html = html & CgTablaAplAbre()
		'[UI-ANT] If Not rz.eof Then
			'[UI-ANT] Do While Not rz.eof
				'[UI-ANT] html = html & CgFilaApl(rz)
				'[UI-ANT] rz.movenext
			'[UI-ANT] Loop
		'[UI-ANT] Else
			'[UI-ANT] html = html & "<tr><td colspan='6'><div class='ui-empty'>No hay APLs relacionados.</div></td></tr>"
		'[UI-ANT] End If
		'[UI-ANT] html = html & "</tbody></table></div></div>"

		' [CG] APLs relacionados tambien en edicion (funcionalidad de P2)
		html = html & DetTarjetaApl(id, "I", "incidentes")

		' --- Guardado: al final, porque tambien envia el nuevo avance del historial ---
		html = html & CgSavebar("detalles.asp?opc=veri&id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu)

		html = html & "</div></div>"
		html = html & "</form>"

		rx.close

	ElseIf opc = "actp" Then

		zzz = ObjAcc(id,"P",cnn,"A")
		XXX = SYSLOGON(cnn,"ACTUALIZA_PROY_" & id)
		wPermisos = getPermisosAtrm("CG_PROYECTOS","MODIF")
		editaFunc = (CStr(wPermisos) = "1" And session("FUNC") = "FUNC")
		editaTec = (CStr(wPermisos) = "1" And session("FUNC") = "TEC")

		Set rs = ObtenerProyectoEdicion(Cnn, id)
		If Not rs.eof Then

			' --- Datos calculados igual que en la pantalla anterior ---
			sFec = ""
			If rs(9) <> 0 Then sFec = FormatDateTime((Fechador(rs(9) + nFecini)), 2)
			sFecEn = ""
			If rs(10) <> 0 Then sFecEn = FormatDateTime((Fechador(rs(10) + nFecini)), 2)
			sFecOkCli = ""
			If rs(33) & "" <> "0" And rs(33) & "" <> "" Then sFecOkCli = FormatDateTime((Fechador(rs(33) + nFecini)), 2)
			editaFuncEstado = (editaFunc And (rs(11) & "" = "OK" Or rs(11) & "" = "BAJA"))

			' El form abre antes de la barra superior: titulo, cliente, entrega y prioridad se editan ahi
			html = html & "<form id='sForm' method='post' action='actualizar_v2.asp?tabla=proyectos&id=" & id & "&opt=actproy&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=detalles.asp?opc=actp&sol=" & rs(1) & "' >"

			' --- Barra superior editable ---
			cgTopFila = Replace(cgTopFila, "<!--CG_TITULO-->", "<div class='ui-topbar__title-field'>" & HtmlInputCampo("titulo", rs(2), (editaFunc Or editaTec)) & "</div>")
			cgFIni = CgFechaDb(rs(9))
			cgFLim = CgFechaDb(rs(10))
			cgTopEdit = True
			cgMetas = CgMeta("Cliente", ComboCampo(Cnn, rs(21), "id_cliente", "id_cliente", (editaFunc Or editaTec))) & CgMeta("Fecha creaci&oacute;n", CgTexto(sFec)) & CgMeta("Fecha entrega", HtmlInputCampo("fecha_entrega", sFecEn, editaFunc)) & CgMeta("Prioridad", ComboCampo(Cnn, rs(28), "prioridad", "prioridad", editaFunc))
			html = html & CgTopbar(cgTopFila, cgMetas, CgSemaforo(cgFIni, cgFLim, rs(33), "fecha de entrega"))

			html = html & "<div class='ui-app ui-page ui-page--edit'><div class='ui-container'>"

			' --- Tarjeta Informacion ---
			html = html & "<div class='ui-card'>" & CgCardTitulo("Informaci&oacute;n") & "<div class='ui-fields'>"
			html = html & CgCampo("Solicitud n&uacute;mero", HtmlInputCampo("solic_nro", rs(1), (editaFunc Or editaTec)))
			html = html & CgCampo("Contacto", ComboCampo(Cnn, rs(22), "contacto", "contacto", (editaFunc Or editaTec)))
			html = html & CgCampo("Aplicaci&oacute;n", ComboCampo(Cnn, rs(23), "id_aplic", "id_aplic", (editaFunc Or editaTec)))
			html = html & CgCampo("V&iacute;a pedido", ComboCampo(Cnn, rs(30), "via_pedido", "via_pedido", (editaFunc Or editaTec)))
			html = html & CgCampo("Browser", ComboCampo(Cnn, rs(29), "browser", "browser", (editaFunc Or editaTec)))
			html = html & CgCampo("Requiere presupuesto", HtmlSelectSiNo("req_presup", rs(20), (editaFunc Or editaTec)))
			html = html & CgCampo("Estado funcional", ComboCampo(Cnn, rs(26), "estado_func", "estado_func", editaFuncEstado))
			html = html & CgCampo("Estado desarrollo", ComboCampo(Cnn, rs(27), "estado_desa", "estado_desa", editaTec))
			html = html & CgCampo("Responsable gesti&oacute;n", ComboCampo(Cnn, rs(24), "respon_gestion", "respon_gestion", editaFunc))
			html = html & CgCampo("Responsable t&eacute;cnico", ComboCampo(Cnn, rs(25), "respon_tec", "respon_tec", editaTec))
			html = html & CgCampo("Fecha OK cliente", HtmlInputCampo("ok_cliente", sFecOkCli, (editaFunc Or editaTec) And (rs(11) & "" = "OK" And rs(12) & "" = "OK")))
			' [UI] Estado del cliente quitado: se decidio no sumar este dato (pedido de Luis)
			'[UI-ANT] html = html & CgCampo("Estado cliente", CgFalta("estado del cliente"))
			If Cstr(wPermisos) = "1" Then
				test = "<a class='ui-link' href='" & Server.HTMLEncode(rs("ruta") & "") & "'>Test</a>"
				html = html & "<div><div class='ui-field__label'>Link documento</div><div class='ui-field__value ui-field__value--inline'><input class='ui-input' id='ruta' name='ruta' type='text' value='" & Server.HTMLEncode(rs("ruta") & "") & "'>" & test & "</div></div>"
			Else
				xruta="http://" & replace(rs("ruta") & "","\\","")
				xruta=replace(xruta,"\","/")
				html = html & CgCampoMod("Link documento", "<a class='ui-link' href='" & Server.HTMLEncode(xruta) & "' target='_new'>" & Server.HTMLEncode(Left(rs("ruta") & "",30)) & "</a>", "ui-field__value--break")
			End If
			html = html & "</div></div>"

			' --- Tarjeta Carga de horas ---
			html = html & "<div class='ui-card'>" & CgCardTitulo("Carga de horas") & "<div class='ui-fields'>"
			html = html & CgCampo("Relev y definici&oacute;n (hs)", HtmlInputCampo("tiempo_relevamiento", rs(14), editaFunc))
			html = html & CgCampo("Dise&ntilde;o (hs)", HtmlInputCampo("tiempo_diseno", rs(15), editaFunc))
			html = html & CgCampo("Desarrollo (hs)", HtmlInputCampo("tiempo_desarr", rs(16), editaTec))
			html = html & CgCampo("Testeo (hs)", HtmlInputCampo("tiempo_test", rs(17), editaFunc))
			html = html & CgCampoMod("Total", CgHoras(rs(32)), "ui-hours__total")
			html = html & "</div>" & CgBarraHoras(rs(14), rs(15), rs(16), rs(17)) & "</div>"

			' --- Tarjeta Detalle (la descripcion no se edita desde aca, igual que antes) ---
			XPATH="\\" & request.servervariables("SERVER_NAME") &  RIGHT(RUTA,LEN(RUTA)-2)
			RUTA=RUTA & "\proyectos\" & Rs(0)
			html = html & "<div class='ui-card'>" & CgCardTitulo("Detalle")
			html = html & "<div class='ui-field'><div class='ui-field__label'>Descripci&oacute;n</div><div class='ui-prose'>" & Replace(Server.HTMLEncode(rs(3) & ""),Chr(13),"<br>") & "</div></div>"
			html = html & CgPath(XPATH & "proyectos\" & Rs(0))
			html = html & "<div class='ui-field'><div class='ui-field__label'>Archivos</div><ul class='ui-files'>" & LISTARDIRTORIO(RUTA,Rs(0),"proyectos") & "</ul></div>"
			html = html & "</div>"

		' [UI] Edicion: solo Informacion, Carga de horas y Detalle. Historial, APLs e incidentes relacionados se ven y se cargan desde la vista.
			' --- Tarjeta Historial de avances ---
			'[UI-ANT] html = html & "<div class='ui-card'>" & CgCardTitulo("Historial de avances")
			'[UI-ANT] html = html & CgFaltaBloque("tipo de avance (Anotaci&oacute;n, Solicitud, Insistencia), estado, testing y recepci&oacute;n de solicitudes. Por ahora el avance se carga en los campos de abajo y se guarda con Actualizar informaci&oacute;n.")
			'[UI-ANT] html = html & HtmlHistorialRelainc(Cnn, id, "proy")
			'[UI-ANT] If editaFunc Or editaTec Then
				'[UI-ANT] html = html & "<div class='ui-field ui-field--top'><div class='ui-field__label'>Nuevo avance</div><div class='ui-table-wrap'>" & HtmlCamposAvance(RolResponsableActual()) & "</div></div>"
			'[UI-ANT] End If
			'[UI-ANT] html = html & "</div>"

			' --- Tarjeta APLs relacionados (el id sAgre abre el dialogo de siempre) ---
			'[UI-ANT] Set rz = ObtenerAplRelac(Cnn, id, "P")
			'[UI-ANT] html = html & "<div class='ui-card'>" & CgCardHeader("APLs relacionados", "<span id='sAgre' class='ui-button ui-button--secondary ui-button--sm'>Relacionar APL</span>")
			'[UI-ANT] html = html & CgTablaAplAbre()
			'[UI-ANT] If Not rz.eof Then
				'[UI-ANT] Do While Not rz.eof
					'[UI-ANT] html = html & CgFilaApl(rz)
					'[UI-ANT] rz.movenext
				'[UI-ANT] Loop
			'[UI-ANT] Else
				'[UI-ANT] html = html & "<tr><td colspan='6'><div class='ui-empty'>No hay APLs relacionados.</div></td></tr>"
			'[UI-ANT] End If
			'[UI-ANT] html = html & "</tbody></table></div></div>"

			' --- Tarjeta Incidentes relacionados ---
			'[UI-ANT] Set rz = ObtenerIncidentesPorProyecto(Cnn, id)
			'[UI-ANT] html = html & "<div class='ui-card'>" & CgCardTitulo("Incidentes relacionados") & "<div class='ui-list'>"
			'[UI-ANT] If Not rz.eof Then
				'[UI-ANT] Do While Not rz.eof
					'[UI-ANT] html = html & "<div class='ui-list__item'><a class='ui-link' href='detalles.asp?opc=veri&id=" & rz(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>INC-" & rz(0) & "</a></div>"
					'[UI-ANT] rz.movenext
				'[UI-ANT] Loop
			'[UI-ANT] Else
				'[UI-ANT] html = html & "<div class='ui-empty'>No hay incidentes relacionados.</div>"
			'[UI-ANT] End If
			'[UI-ANT] html = html & "</div></div>"

			' [CG] APLs relacionados tambien en edicion (funcionalidad de P2)
			html = html & DetTarjetaApl(id, "P", "proyectos")

			' --- Guardado ---
			html = html & CgSavebar("detalles.asp?opc=verp&id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu)

			html = html & "</div></div>"
			html = html & "</form>"

		End If

	End If

	' Si no hubo datos (registro inexistente), igual se muestra la barra con titulo y botones
	If Not cgTopEmitido Then html = html & CgTopbar(Replace(cgTopFila, "<!--CG_TITULO-->", ""), "", "")

	html = html & "</div>"

	html = html & "<script type='text/javascript'>"
	html = html & " function actualizar(formu){ "
	html = html & " var r = confirm('Desea actualizar ?'); "
	html = html & " if (r){ "
	html = html & " document.getElementById(formu).submit(); "
	html = html & " }else{ "
	html = html & " return false;"
	html = html & " }} "
	html = html & "</script>"

	' [CG] Copiar el path. En la intranet por HTTP el portapapeles moderno no esta disponible:
	' [CG] se usa la alternativa execCommand y, si falla, se muestra la ruta para copiarla a mano.
	html = html & "<script type='text/javascript'>"
	html = html & "function cgCopiar(idElem){"
	html = html & " var el = document.getElementById(idElem); if (!el) { return; }"
	html = html & " var texto = el.innerText || el.textContent;"
	html = html & " if (navigator.clipboard && window.isSecureContext) { navigator.clipboard.writeText(texto).then(function(){ cgAviso('Path copiado'); }); return; }"
	html = html & " var area = document.createElement('textarea'); area.value = texto; area.style.position = 'fixed'; area.style.opacity = '0';"
	html = html & " document.body.appendChild(area); area.select();"
	html = html & " try { document.execCommand('copy'); cgAviso('Path copiado'); } catch (e) { window.prompt('Copi\u00e1 la ruta:', texto); }"
	html = html & " document.body.removeChild(area);"
	html = html & "}"
	html = html & "function cgAviso(msg){"
	html = html & " var t = document.createElement('div'); t.className = 'ui-toast ui-toast--visible'; t.innerHTML = msg;"
	html = html & " document.body.appendChild(t); setTimeout(function(){ document.body.removeChild(t); }, 2200);"
	html = html & "}"
	' [CG] En edicion: avisa si hay cambios sin guardar
	html = html & "(function(){ var f = document.getElementById('sForm'); var e = document.getElementById('cgEstadoCambios'); if (!f || !e) { return; }"
	html = html & " function marcar(){ e.innerHTML = 'Hay cambios sin guardar'; e.className = 'ui-savebar__status ui-savebar__status--pending'; }"
	html = html & " if (f.addEventListener) { f.addEventListener('input', marcar); f.addEventListener('change', marcar); } })();"
	html = html & "</script>"

' Else

	' If opc = "veri" Then

		' vuel = request.querystring("vuel")

		' html = html & "<div>"

		' html = html & "<header id='header' style='height:16%'>"

		' html = html & "<div>"

		' html = html & "<a href='index.asp' class='logo'>"

		' html = html & "<img src='images/logocorP4.gif' width=10% >"
		' html = html & "<font style='font-size:40pt;align:center;valign:center;font-family:Serif;color:#000000;'>Cor Consulting</font>"

		' html = html & "</a>"

		' html = html & "</div>"

		' html = html & "</header>"

		' html = html & "</div>"

		' html = html & "<div><br><br>"

		' html = html & "<table style='height:50%;width:90%;' align=center>"

		' html = html & "<tr>"

		' html = html & "<td style='height:100%;width:100%;' ><font style='color:#000000;font-size:30pt;'>INCIDENTE NO ENCONTRADO</font></td>"

		' html = html & "</tr>"

		' html = html & "<tr>"

		' html = html & "<td style='height:100%;width:100%;' >"
		' html = html & "<form id='xfrm' method='post' action='consultas.asp?sel=" & vuel & "' >"
		' html = html & "<input style='border-radius:10px;border:solid 2px;border-color:blue;height:100%;width:50%;' value='VOLVER' type='button' onclick=""document.getElementById('xfrm').submit();"">"
		' html = html & "</form>"
		' html = html & "</td>"

		' html = html & "</tr>"

		' html = html & "</table>"

		' html = html & "</div>"

	' ElseIf opc = "verp" Then

		' vuel = request.querystring("vuel")

		' html = html & "<div>"

		' html = html & "<header id='header' style='height:16%'>"

		' html = html & "<div>"

		' html = html & "<a href='index.asp' class='logo'>"

		' html = html & "<img src='images/logocorP4.gif' width=10% >"
		' html = html & "<font style='font-size:40pt;align:center;valign:center;font-family:Serif;color:#000000;'>Cor Consulting</font>"

		' html = html & "</a>"

		' html = html & "</div>"

		' html = html & "</header>"

		' html = html & "</div>"

		' html = html & "<div><br><br>"

		' html = html & "<table style='height:50%;width:90%;' align=center>"

		' html = html & "<tr>"

		' html = html & "<td style='height:100%;width:100%;' ><font style='color:#000000;font-size:30pt;'>PROYECTO NO ENCONTRADO</font></td>"

		' html = html & "</tr>"

		' html = html & "<tr>"

		' html = html & "<td style='height:100%;width:100%;' >"
		' html = html & "<form id='xfrm' method='post' action='consultas.asp?sel=" & vuel & "' >"
		' html = html & "<input style='border-radius:10px;border:solid 2px;border-color:blue;height:100%;width:50%;' value='VOLVER' type='button' onclick=""document.getElementById('xfrm').submit();"">"
		' html = html & "</form>"
		' html = html & "</td>"

		' html = html & "</tr>"

		' html = html & "</table>"

		' html = html & "</div>"

	' End If


End If

html = html & "<script src='assets/js/relaciones.js'></script>"

html = html & "<div id='dialog' title='' >"
html = html & "</div>"

html = html & "<script>"
html = html & "  $( function() {"

html = html & "$.datepicker.regional['es'] = {"
html = html & "closeText: 'Cerrar',"
html = html & "prevText: '< Ant',"
html = html & "nextText: 'Sig >',"
html = html & "currentText: 'Hoy',"
html = html & "monthNames: ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'],"
html = html & "monthNamesShort: ['Ene','Feb','Mar','Abr', 'May','Jun','Jul','Ago','Sep', 'Oct','Nov','Dic'],"
'[CG-ANT] html = html & "dayNames: ['Domingo', 'Lunes', 'Martes', 'Miercoles', 'Jueves', 'Viernes', 'Sabado'],"
html = html & "dayNames: ['Domingo', 'Lunes', 'Martes', 'Mi\u00e9rcoles', 'Jueves', 'Viernes', 'S\u00e1bado'],"
'[CG-ANT] html = html & "dayNamesShort: ['Dom','Lun','Mar','MiN','Juv','Vie','SNb'],"
html = html & "dayNamesShort: ['Dom','Lun','Mar','Mi\u00e9','Jue','Vie','S\u00e1b'],"
html = html & "dayNamesMin: ['Do','Lu','Ma','Mi','Ju','Vi','Sa'],"
html = html & "weekHeader: 'Sm',"
html = html & "dateFormat: 'dd/mm/yy',"
html = html & "firstDay: 1,"
html = html & "isRTL: false,"
html = html & "showMonthAfterYear: false,"
html = html & "yearSuffix: ''"
html = html & "};"
html = html & "$.datepicker.setDefaults($.datepicker.regional['es']);"

html = html & "    $( ""#fecha_ingreso"" ).datepicker();"
html = html & "    $( ""#fecha_entrega"" ).datepicker();"
html = html & "    $( ""#ok_cliente"" ).datepicker();"
html = html & "  } );"
html = html & "  </script>"

html = html & "<script>"

' DEPURADO [CG]: sin pestanas no hace falta inicializar jQuery UI Tabs
'[CG-ANT] html = html & "$( '#tabs' ).tabs();"

Dim tipoRelacDialog
If opc = "actp" Then
	tipoRelacDialog = "P"
Else
	tipoRelacDialog = "I"
End If

html = html & "$( '#dialog' ).dialog({"
html = html & "title: 'Agregar relacion',"
html = html & "autoOpen: false,"
html = html & "resizable: false,"
html = html & "buttons: {"
html = html & "'Aceptar': function () {"
html = html & "capturarYProcesarApl('" & tipoRelacDialog & "');"
html = html & "$(this).dialog('close');"
html = html & "},"
html = html & "'Cerrar': function () {"
html = html & "$(this).dialog('close');"
html = html & "}"
html = html & "},"
html = html & "show: {"
html = html & "effect: 'blind',"
html = html & "duration: 100"
html = html & "      },"
html = html & "hide: {"
html = html & "effect: 'blind',"
html = html & "duration: 100"
html = html & "}"
html = html & "});"

html = html & "$( '#sAgre' ).click(function() {"
html = html & "$.get('armaCargaParam.asp?param=APL', function( data ) {"
html = html & "$( '#dialog' ).html(data);"
html = html & "});"
html = html & "$( '#dialog' ).dialog( 'open' );"
html = html & "});"

html = html & "</script>"

html = html & sformz & form7 & xform7 & form5 & xform7 & form8 & formz1 & "</body>"
html = html & "</html>"


Function LISTARDIRTORIO(DIR,id,d)
	Set FSO = server.createObject("Scripting.FileSystemObject")
	'RESPONSE.WRITE DIR
	If Not fso.FolderExists(DIR ) Then
		fso.CreateFolder (DIR)
	End If

	Set carpeta= FSO.GetFolder(DIR)




	Set archivos = carpeta.Files

	For Each nombre_archivo In archivos
		'[CG-ANT] surl="<a class='col-link' target=_new href='" &  shttp & d& "/" & id & "/"  & Server.HTMLEncode(nombre_archivo.name) & "'>" & Server.HTMLEncode(nombre_archivo.name) & "</a><br>"
		surl="<li class='ui-files__item'><a class='ui-link' target=_new href='" &  shttp & d& "/" & id & "/"  & Server.HTMLEncode(nombre_archivo.name) & "'>" & Server.HTMLEncode(nombre_archivo.name) & "</a></li>"
		LISTARDIRTORIO=LISTARDIRTORIO & surl
	Next
End Function

' =====================================================================
' [CG] FUNCIONES DE PRESENTACION DEL SISTEMA VISUAL (cor-gestion.css)
' Solo arman HTML: no consultan ni modifican datos.
' =====================================================================

' Etiqueta + valor dentro de la grilla de campos
Function CgCampo(etiqueta, valorHtml)
	CgCampo = "<div><div class='ui-field__label'>" & etiqueta & "</div><div class='ui-field__value'>" & valorHtml & "</div></div>"
End Function

' Igual que CgCampo, con un modificador para el valor (ej. cg-field__value--break)
Function CgCampoMod(etiqueta, valorHtml, modificador)
	CgCampoMod = "<div><div class='ui-field__label'>" & etiqueta & "</div><div class='ui-field__value " & modificador & "'>" & valorHtml & "</div></div>"
End Function

' Dato que la propuesta necesita y todavia no esta conectado
Function CgFalta(que)
	CgFalta = "<span class='ui-falta' title='Dato pendiente de conectar'>Falta dato: " & que & "</span>"
End Function

' Aviso de bloque para funcionalidades pendientes de conectar
Function CgFaltaBloque(texto)
	CgFaltaBloque = "<div class='ui-falta-bloque'><b>Falta conectar:</b> " & texto & "</div>"
End Function

' True si el valor trae un dato real (no vacio, no "---", no "0")
Function CgHayDato(v)
	Dim t
	t = Trim(v & "")
	CgHayDato = Not (t = "" Or t = "---" Or t = "0")
End Function

' Texto del sistema, o "Pendiente" cuando viene vacio
Function CgTexto(v)
	If CgHayDato(v) Then
		CgTexto = Server.HTMLEncode(Trim(v & ""))
	Else
		CgTexto = "<span class='ui-field__value--muted'>Pendiente</span>"
	End If
End Function

' "PENDIENTE" -> "Pendiente". Las siglas cortas (OK, SI, NO) quedan igual
Function CgCapitalizar(v)
	Dim t
	t = Trim(v & "")
	If Len(t) > 2 Then
		CgCapitalizar = UCase(Left(t, 1)) & LCase(Mid(t, 2))
	Else
		CgCapitalizar = t
	End If
End Function

' Pill de estado con color segun el valor del catalogo
' CONFIRMAR: los valores reales del catalogo de estados, para completar los colores
Function CgBadgeEstado(v)
	Dim t, c
	t = UCase(Trim(v & ""))
	If t = "" Or t = "---" Then
		CgBadgeEstado = "<span class='ui-field__value--muted'>Sin estado</span>"
		Exit Function
	End If
	Select Case t
		Case "OK", "FINALIZADO", "CERRADO", "APLICADO", "CONFORME", "RESUELTO"
			c = "green"
		Case "BAJA", "RECHAZADO", "CANCELADO"
			c = "red"
		Case "EN CURSO", "EN DESARROLLO", "EN PROCESO"
			c = "amber"
		Case "TESTING", "EN TESTING", "EN TESTEO"
			c = "purple"
		Case "EN ANALISIS", "EN RELEVAMIENTO", "EN DEFINICION"
			c = "blue"
		Case Else
			c = "gray"
	End Select
	CgBadgeEstado = "<span class='ui-badge ui-badge--" & c & "'>" & Server.HTMLEncode(CgCapitalizar(v)) & "</span>"
End Function

' Pill de prioridad: Alta rojo, Media ambar, Baja verde
Function CgPrioridadPill(v)
	Dim t, c
	t = UCase(Trim(v & ""))
	If t = "" Then
		CgPrioridadPill = CgFalta("prioridad")
		Exit Function
	End If
	Select Case t
		Case "ALTA", "URGENTE", "CRITICA"
			c = "red"
		Case "MEDIA"
			c = "amber"
		Case "BAJA"
			c = "green"
		Case Else
			c = "gray"
	End Select
	CgPrioridadPill = "<span class='ui-badge ui-badge--" & c & "'>" & Server.HTMLEncode(CgCapitalizar(v)) & "</span>"
End Function

' Dias de compromiso segun prioridad (definido en la propuesta)
' PENDIENTE DEFINICION: hoy se cuentan dias corridos. Falta confirmar si son habiles,
' si se descuentan feriados y si el plazo se pausa mientras hay una solicitud al cliente.
Function CgPlazoDias(v)
	Select Case UCase(Trim(v & ""))
		Case "ALTA"
			CgPlazoDias = 5
		Case "MEDIA"
			CgPlazoDias = 10
		Case "BAJA"
			CgPlazoDias = 15
		Case Else
			CgPlazoDias = 0
	End Select
End Function

' Fecha del sistema (dias desde 01-01-1900, mismo criterio que Fechador) a Date. Empty si no hay fecha
Function CgFechaDb(v)
	CgFechaDb = Empty
	If IsNumeric(v & "") Then
		If CLng(v) > 0 Then CgFechaDb = CDate(CLng(v) + CLng(CDate("01-01-1900")))
	End If
End Function

' Semaforo de compromiso.
' fIni, fLim: Date o Empty. okCliente: si trae dato, el caso esta resuelto y no corre el plazo.
' faltaQue: que dato falta cuando no se puede calcular.
Function CgSemaforo(fIni, fLim, okCliente, faltaQue)
	Dim plazo, rest, dias, cls, num, etiqueta, pct
	If CgHayDato(okCliente) Then
		CgSemaforo = "<div class='ui-semaforo ui-semaforo--done'><div class='ui-semaforo__cifra'><div class='ui-semaforo__num'>&#10003;</div><div class='ui-semaforo__label'>Resuelto</div></div><div class='ui-semaforo__graf'><div class='ui-semaforo__scale'><span>OK cliente: " & Server.HTMLEncode(Trim(okCliente & "")) & "</span></div></div></div>"
		Exit Function
	End If
	If IsEmpty(fIni) Or IsEmpty(fLim) Then
		CgSemaforo = "<div class='ui-semaforo ui-semaforo--gray'><div class='ui-semaforo__cifra'><div class='ui-semaforo__num'>?</div><div class='ui-semaforo__label'>Sin plazo</div></div><div class='ui-semaforo__graf'>" & CgFalta(faltaQue) & "</div></div>"
		Exit Function
	End If
	plazo = DateDiff("d", fIni, fLim)
	If plazo < 1 Then plazo = 1
	rest = DateDiff("d", Date(), fLim)
	dias = Abs(rest)
	If rest <= 2 Then
		cls = "red"
	ElseIf rest <= 5 Then
		cls = "amber"
	Else
		cls = "green"
	End If
	If rest = 0 Then
		num = "&iexcl;Hoy!"
		etiqueta = "vence"
	ElseIf rest < 0 Then
		num = dias
		If dias = 1 Then
			etiqueta = "d&iacute;a vencido"
		Else
			etiqueta = "d&iacute;as vencido"
		End If
	Else
		num = dias
		If dias = 1 Then
			etiqueta = "d&iacute;a restante"
		Else
			etiqueta = "d&iacute;as restantes"
		End If
	End If
	pct = Round(((plazo - rest) / plazo) * 100)
	If pct < 0 Then pct = 0
	If pct > 100 Then pct = 100
	CgSemaforo = "<div class='ui-semaforo ui-semaforo--" & cls & "'><div class='ui-semaforo__cifra'><div class='ui-semaforo__num'>" & num & "</div><div class='ui-semaforo__label'>" & etiqueta & "</div></div><div class='ui-semaforo__graf'><div class='ui-semaforo__bar'><span style='width:" & pct & "%'></span></div><div class='ui-semaforo__scale'><span>Inicio</span><span>" & pct & "% transcurrido</span><span>L&iacute;mite</span></div></div></div>"
End Function

' Un dato de la barra superior (cliente, fecha, prioridad...)
Function CgMeta(etiqueta, valorHtml)
	CgMeta = "<div><span class='ui-topbar__label'>" & etiqueta & "</span><span class='ui-topbar__value'>" & valorHtml & "</span></div>"
End Function

' Barra superior completa: la fila de titulo y botones (cgTopFila) + datos clave + semaforo.
' Con cgTopEdit = True usa la variante de edicion (controles mas compactos).
Function CgTopbar(fila, metas, semaforo)
	Dim modificador
	modificador = ""
	If cgTopEdit Then modificador = " ui-topbar--edit"
	CgTopbar = "<div class='ui-app ui-topbar" & modificador & "'><div class='ui-topbar__inner'>" & fila & "<div class='ui-topbar__bottom'><div class='ui-topbar__meta'>" & metas & "</div><div>" & semaforo & "</div></div></div></div>"
	cgTopEmitido = True
End Function

' Titulo simple de tarjeta
Function CgCardTitulo(titulo)
	CgCardTitulo = "<div class='ui-card__title'>" & titulo & "</div>"
End Function

' Encabezado de tarjeta con una accion a la derecha
Function CgCardHeader(titulo, accionHtml)
	CgCardHeader = "<div class='ui-card__header'><div class='ui-card__title'>" & titulo & "</div>" & accionHtml & "</div>"
End Function

' Boton visible pero todavia sin conectar: explica por que al tocarlo
Function CgBotonPendiente(texto, motivo)
	CgBotonPendiente = "<button type='button' class='ui-button ui-button--secondary ui-button--sm ui-button--disabled' title='" & motivo & "' onclick=""alert('" & motivo & "');"">" & texto & "</button>"
End Function

' Path con boton para copiar (la funcion JS cgCopiar esta al final de la pagina)
Function CgPath(ruta)
	CgPath = "<div class='ui-field'><div class='ui-field__label'>Path</div><div class='ui-path'><span class='ui-path__text' id='cgPath'>" & Server.HTMLEncode(ruta & "") & "</span><button type='button' class='ui-path__copy' onclick=""cgCopiar('cgPath');"">Copiar</button></div></div>"
End Function

' Apertura de la tabla de APLs con las columnas de la propuesta
Function CgTablaAplAbre()
	CgTablaAplAbre = "<div class='ui-table-wrap'><table class='ui-table'><thead><tr><th>APL</th><th>Tipo</th><th>Fecha de env&iacute;o</th><th>Responsable</th><th>Estado del cliente</th><th>Descripci&oacute;n / alcance</th></tr></thead><tbody>"
End Function

' Fila de APL relacionado. Recibe el recordset de ObtenerAplRelac
' (sp_ObtenerAplRelac -> sp_RelacionAPL @Modo='relac') y lee las columnas
' POR NOMBRE: el SP devuelve dos columnas "id"/"ID", y por posicion rz(4)
' es b.ID y rz(5) es b.NOMBRE, lo que antes hacia que incidentes mostrara
' el numero de APL y proyectos el nombre.
Function CgFilaApl(rz)
	Dim idApl, fecEnvio
	idApl = rz("id_apl") & ""
	fecEnvio = ""
	If IsNumeric(rz("FECHA_ENVIO") & "") Then
		If CLng(rz("FECHA_ENVIO")) <> 0 Then fecEnvio = FormatDateTime(CDate(CLng(rz("FECHA_ENVIO")) + CLng(CDate("01-01-1900"))), 2) ' FECHA_ENVIO = dias desde 1900-01-01
	End If
	'[UI-ANT] CgFilaApl = "<tr>" & _
	'[UI-ANT] "<td><a class='ui-link' href='detalles2.asp?opc=ver&id=" & Server.URLEncode(idApl) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>" & Server.HTMLEncode(idApl) & "</a></td>" & _
	'[UI-ANT] "<td>" & Server.HTMLEncode(rz("tipo_relac") & "") & "</td>" & _
	'[UI-ANT] "<td>" & fecEnvio & "</td>" & _
	'[UI-ANT] "<td>" & Server.HTMLEncode(rz("usuario") & "") & "</td>" & _
	'[UI-ANT] "<td>" & Server.HTMLEncode(rz("DESCESTADO") & "") & "</td>" & _
	'[UI-ANT] "<td>" & Server.HTMLEncode(rz("NOMBRE") & "") & "</td>" & _
	'[UI-ANT] "</tr>"
	' [UI] Los datos vacios se muestran con un guion atenuado (CgCelda), como en el resto del sistema
	CgFilaApl = "<tr>" & _
		"<td><a class='ui-link' href='detalles2.asp?opc=ver&id=" & Server.URLEncode(idApl) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>" & Server.HTMLEncode(idApl) & "</a></td>" & _
		"<td>" & CgCelda(rz("tipo_relac")) & "</td>" & _
		"<td>" & CgCelda(fecEnvio) & "</td>" & _
		"<td>" & CgCelda(rz("usuario")) & "</td>" & _
		"<td>" & CgCelda(rz("DESCESTADO")) & "</td>" & _
		"<td>" & CgCelda(rz("NOMBRE")) & "</td>" & _
		"</tr>"
End Function

' [UI] Celda de tabla: el dato, o un guion atenuado si viene vacio
Function CgCelda(v)
	If Trim(v & "") = "" Then
		CgCelda = "<span class='ui-status-empty'>&mdash;</span>"
	Else
		CgCelda = Server.HTMLEncode(Trim(v & ""))
	End If
End Function

' Barra de guardado de las pantallas de edicion
Function CgSavebar(urlCancelar)
	CgSavebar = "<div class='ui-savebar'><span class='ui-savebar__status' id='cgEstadoCambios'>Sin cambios pendientes</span><div class='ui-savebar__actions'><button type='button' class='ui-button ui-button--secondary ui-button--lg' onclick=""location.href='" & urlCancelar & "';"">Cancelar</button><input class='ui-button ui-button--primary ui-button--lg' type='button' value='Actualizar informaci&oacute;n' onclick=""actualizar('sForm');""></div></div>"
End Function

' Numero seguro para calculos (vacio o texto -> 0)
Function CgNum(v)
	If IsNumeric(v & "") Then
		CgNum = CDbl(v)
	Else
		CgNum = 0
	End If
End Function

' Horas para mostrar: "3 h" o "Pendiente"
Function CgHoras(v)
	If IsNumeric(v & "") Then
		CgHoras = v & " h"
	Else
		CgHoras = "<span class='ui-field__value--muted'>Pendiente</span>"
	End If
End Function

' Un tramo de la barra de horas. El Replace evita la coma decimal de la configuracion regional en el CSS
Function CgTramo(v, total, color, nombre)
	CgTramo = "<span style='width:" & Replace(CStr(Round(CgNum(v) / total * 100, 2)), ",", ".") & "%;background:" & color & "' title='" & nombre & "'></span>"
End Function

' Barra proporcional de la carga de horas del proyecto + leyenda
Function CgBarraHoras(hRel, hDis, hDes, hTes)
	Dim total
	total = CgNum(hRel) + CgNum(hDis) + CgNum(hDes) + CgNum(hTes)
	CgBarraHoras = "<div class='ui-hours__bar'>"
	If total > 0 Then
		CgBarraHoras = CgBarraHoras & CgTramo(hRel, total, "#7995FF", "Relev y definicion") & CgTramo(hDis, total, "#6D28D9", "Diseno") & CgTramo(hDes, total, "#132574", "Desarrollo") & CgTramo(hTes, total, "#0E7490", "Testeo")
	End If
	CgBarraHoras = CgBarraHoras & "</div><div class='ui-hours__legend'><span><span class='ui-dot' style='background:#7995FF'></span>Relev y definici&oacute;n</span><span><span class='ui-dot' style='background:#6D28D9'></span>Dise&ntilde;o</span><span><span class='ui-dot' style='background:#132574'></span>Desarrollo</span><span><span class='ui-dot' style='background:#0E7490'></span>Testeo</span></div>"
End Function

response.write html


' =============================================================================
' [UI] Tarjetas de la vista con pop-up (incidente y proyecto)
' =============================================================================

' Script para abrir y cerrar los pop-ups de la vista (se define una sola vez aunque haya dos tarjetas)
Function DetScriptDialogos()
	DetScriptDialogos = "<script>if(!window.DetAbrir){window.DetAbrir=function(id,foco){var d=document.getElementById(id);d.__origen=document.activeElement;d.showModal();var f=document.getElementById(foco);if(f)f.focus();};" & _
		"window.DetCerrar=function(id){document.getElementById(id).close();};" & _
		"document.addEventListener('close',function(e){var d=e.target;if(d.classList&&d.classList.contains('ui-dialogo')&&d.__origen)d.__origen.focus();},true);" & _
		"document.addEventListener('click',function(e){if(e.target.tagName==='DIALOG'&&e.target.classList.contains('ui-dialogo'))e.target.close();});}</script>"
End Function

' Historial de avances con el pop-up para anadir un avance. Mismos campos y nombres que la edicion
' (HtmlCamposAvance): texto, fecha y hora (por defecto, el momento actual), responsable (por defecto,
' el rol del usuario) y tipo. Lo guarda actualizar_v2.asp con RegistrarAvance.
' [CONSULTA] Direccion de la vista de consulta de este caso con los filtros con que se llego (para el boton Volver). tipoC = I | P
Function DetUrlVueltaConsulta(tipoC)
	Dim u, n
	u = "vistaconsulta.asp?tipo=" & tipoC & "&id=" & CLng(id) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu
	For Each n In Array("cli", "est", "apli", "fec", "desc", "busc", "aplic")
		If Trim(Request.QueryString(n) & "") <> "" Then u = u & "&" & n & "=" & Server.URLEncode(Request.QueryString(n) & "")
	Next
	DetUrlVueltaConsulta = u
End Function

' tipoHist: "inc" | "proy" - tablaAct: "incidentes" | "proyectos" - optAv: "avainc" | "avaproy"
Function DetTarjetaHistorial(idCaso, tipoHist, tablaAct, optAv)
	Dim h, puede, hoy, hora, rol, rsT, tipos
	puede = (getPermisosAtrm(IIf(tipoHist = "inc", "CG_INCIDENTES", "CG_PROYECTOS"), "MODIF") = 1)
	If puede Then
		h = "<div class='ui-card'>" & CgCardHeader("Historial de avances", "<button type='button' class='ui-button ui-button--light ui-button--sm' id='btnAvance' aria-haspopup='dialog' onclick=""DetNuevoAvance()"">+ A&ntilde;adir avance</button>")
	Else
		h = "<div class='ui-card'>" & CgCardHeader("Historial de avances", "<button type='button' class='ui-button ui-button--light ui-button--sm ui-button--disabled' aria-disabled='true' title='Sin permiso para modificar.'>+ A&ntilde;adir avance</button>")
	End If
	If Request.QueryString("av") = "ok" Then h = h & "<div class='ui-alert ui-alert--success' role='status'>Avance agregado.</div>"
	If Request.QueryString("av") = "error" Then h = h & "<div class='ui-alert ui-alert--danger' role='alert'>No se pudo agregar el avance: complet&aacute; el texto y verific&aacute; tu permiso para modificar.</div>"
	If Request.QueryString("av") = "estado" Then h = h & "<div class='ui-alert ui-alert--success' role='status'>Estado del avance actualizado.</div>"
	If Request.QueryString("av") = "estadoerror" Then h = h & "<div class='ui-alert ui-alert--danger' role='alert'>No se pudo cambiar el estado del avance: verific&aacute; tu permiso para modificar.</div>"
	If puede Then
		hoy = Year(Date()) & "-" & Right("0" & Month(Date()), 2) & "-" & Right("0" & Day(Date()), 2)
		hora = Right("0" & Hour(Time()), 2) & ":" & Right("0" & Minute(Time()), 2)
		rol = RolResponsableActual()
		tipos = "<option value=''>Seleccion&aacute; un tipo</option>"
		Set rsT = GetListaComboExtended(Cnn, "tipo_avance")
		Do While Not rsT.EOF
			tipos = tipos & "<option value='" & rsT(0) & "'>" & Server.HTMLEncode(rsT(1) & "") & "</option>"
			rsT.MoveNext
		Loop
		rsT.Close
		h = h & "<dialog id='dlgAvance' class='ui-dialogo' aria-labelledby='dlgAvanceTitulo'>"
		h = h & "<form method='post' action='actualizar_v2.asp?tabla=" & tablaAct & "&opt=" & optAv & "&id=" & idCaso & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&cantregmover=0'>"
		h = h & "<div class='ui-dialogo__head'><h2 class='ui-dialogo__titulo' id='dlgAvanceTitulo'>Nuevo avance</h2><button type='button' class='ui-dialogo__cerrar' aria-label='Cerrar' onclick=""DetCerrar('dlgAvance')"">&times;</button></div>"
		h = h & "<div class='ui-dialogo__body'>"
		' [INSISTIR] Avance de origen (solo cuando el alta nace de 'Insistir'): el vinculo real es el ID en avance_origen (se guarda en ID_ANT)
		h = h & "<input type='hidden' id='avance_origen' name='avance_origen' value=''>"
		h = h & "<div class='ui-field' id='campo-origen' role='group' aria-labelledby='origen-etiqueta' style='display:none'><span class='ui-field__label' id='origen-etiqueta'>Avance de origen</span><div class='ui-dialogo__valor ui-dialogo__valor--bloqueado' id='origen-texto'></div><p class='ui-dialogo__fijo'>La nueva entrada queda vinculada a este avance.</p></div>"
		h = h & "<div class='ui-field'><label class='ui-field__label' for='avance'>Detalle del avance <span class='ui-field__required' aria-hidden='true'> *</span></label><textarea class='ui-input ui-textarea' id='avance' name='avance' rows='5' maxlength='4000' required aria-required='true'></textarea></div>"
		h = h & "<div class='ui-dialogo__grid'>"
		h = h & "<div class='ui-field'><label class='ui-field__label' for='avance_fecha'>Fecha</label><input class='ui-input' id='avance_fecha' name='avance_fecha' type='date' value='" & hoy & "'></div>"
		h = h & "<div class='ui-field'><label class='ui-field__label' for='avance_hora'>Hora</label><input class='ui-input' id='avance_hora' name='avance_hora' type='time' value='" & hora & "'></div>"
		h = h & "<div class='ui-field'><label class='ui-field__label' for='responsable_avance'>Responsable</label><select class='ui-select' id='responsable_avance' name='responsable_avance'>"
		h = h & "<option value='FUNCIONAL'" & IIf(rol = "FUNCIONAL", " selected", "") & ">Funcional</option>"
		h = h & "<option value='TECNICO'" & IIf(rol = "TECNICO", " selected", "") & ">T&eacute;cnico</option>"
		h = h & "<option value='CLIENTE'" & IIf(rol = "CLIENTE", " selected", "") & ">Cliente</option></select></div>"
		h = h & "<div class='ui-field'><label class='ui-field__label' for='tipo_avance'>Tipo de avance</label><select class='ui-select' id='tipo_avance' name='tipo_avance'>" & tipos & "</select></div>"
		h = h & "</div></div>"
		h = h & "<div class='ui-dialogo__acciones'><button type='button' class='ui-button ui-button--secondary' onclick=""DetCerrar('dlgAvance')"">Cancelar</button><button type='submit' class='ui-button ui-button--primary'>Guardar avance</button></div>"
		h = h & "</form></dialog>" & DetScriptDialogos()
		' [INSISTIR] "Insistir" = alta NUEVA sobre un tema existente: copia el detalle y el responsable del avance (editables), fecha y hora
		' actuales, tipo "Reclamo" preseleccionado (editable) y guarda el avance de origen (avance_origen -> ID_ANT).
		' "+ Anadir avance" (DetNuevoAvance) vuelve al alta normal.
		h = h & "<script>(function(){var rd=document.getElementById('responsable_avance');window.__avDefResp=rd?rd.value:'';"
		h = h & "window.DetNuevoAvance=function(){var d=document;if(window.__avIns){window.__avIns=false;d.getElementById('avance').value='';d.getElementById('tipo_avance').value='';d.getElementById('responsable_avance').value=window.__avDefResp;}"
		h = h & "d.getElementById('avance_origen').value='';d.getElementById('campo-origen').style.display='none';d.getElementById('dlgAvanceTitulo').textContent='Nuevo avance';DetAbrir('dlgAvance','avance');};"
		h = h & "window.HistInsistir=function(b){var d=document,n=new Date(),p=function(x){return (x<10?'0':'')+x;},s=d.getElementById('tipo_avance'),rs=d.getElementById('responsable_avance'),r=(b.getAttribute('data-responsable')||'').toUpperCase(),i;"
		h = h & "d.getElementById('avance').value=b.getAttribute('data-texto')||'';d.getElementById('avance_origen').value=b.getAttribute('data-id')||'';"
		h = h & "for(i=0;i<rs.options.length;i++){if(rs.options[i].value===r){rs.selectedIndex=i;break;}}"
		h = h & "d.getElementById('avance_fecha').value=n.getFullYear()+'-'+p(n.getMonth()+1)+'-'+p(n.getDate());d.getElementById('avance_hora').value=p(n.getHours())+':'+p(n.getMinutes());"
		h = h & "for(i=0;i<s.options.length;i++){if(s.options[i].text.trim().toLowerCase()==='reclamo'){s.selectedIndex=i;break;}}"
		h = h & "var tp=b.getAttribute('data-tipo');d.getElementById('origen-texto').textContent=b.getAttribute('data-fecha')+(tp?' \u00b7 '+tp:'');"
		h = h & "d.getElementById('campo-origen').style.display='';d.getElementById('dlgAvanceTitulo').textContent='Insistir sobre un avance';window.__avIns=true;DetAbrir('dlgAvance','avance');};})();</script>"
	End If
	' [UI] Anotacion quitada: el tipo, el estado y la via del avance ya estan conectados (pedido de Luis)
	'[UI-ANT] h = h & CgFaltaBloque("estado, testing y recepci&oacute;n de solicitudes. Abajo se muestra el historial que existe hoy.")
	' [INSISTIR] El historial muestra 'Insistir' en cada avance solo a quien puede modificar
	uiHistPermiteInsistir = puede
	h = h & HtmlHistorialRelainc(Cnn, idCaso, tipoHist)
	uiHistPermiteInsistir = False
	' [ESTADO] El pop-up de detalle del avance ofrece cambiar su estado solo a quien puede modificar
	If puede Then h = h & HistEstadoCfgScript(idCaso, tablaAct)
	' [INICIO] Desde inicio.asp (Completar testing / Recibi la respuesta / Registrar insistencia): ?avance=<id>&accion=estado|insistir.
	' Al cargar, abre el avance elegido: su detalle (solo cambia el estado) o el alta de Insistir. Si el avance no esta en este caso, no hace nada.
	Dim avAuto, avAccion
	avAuto = Trim(Request.QueryString("avance") & "")
	If IsNumeric(avAuto) And avAuto <> "" Then
		avAccion = LCase(Trim(Request.QueryString("accion") & ""))
		If avAccion <> "insistir" Then avAccion = "estado"
		h = h & "<script>window.addEventListener('load',function(){var id='" & CLng(avAuto) & "',b=null;"
		If avAccion = "insistir" Then h = h & "b=document.querySelector('.ui-hist__insistir[data-id=\x22'+id+'\x22]');"
		h = h & "if(!b)b=document.querySelector('.ui-hist__ver[data-id=\x22'+id+'\x22]');if(b){b.scrollIntoView({block:'center'});b.click();}});</script>"
	End If
	h = h & "</div>"
	DetTarjetaHistorial = h
End Function

' APLs relacionados con el pop-up para relacionar uno: solo pide el ID del APL.
' Lo guarda actualizar_v2.asp (opt=relapl), que valida que el APL exista y que no este ya relacionado.
' tipoRel: "I" | "P" - tablaAct: "incidentes" | "proyectos"
Function DetTarjetaApl(idCaso, tipoRel, tablaAct)
	Dim h, puede, rzA, msg, caso
	puede = (getPermisosAtrm(IIf(tipoRel = "P", "CG_PROYECTOS", "CG_INCIDENTES"), "MODIF") = 1)
	If tipoRel = "P" Then caso = "proyecto" Else caso = "incidente"
	If puede Then
		h = "<div class='ui-card'>" & CgCardHeader("APLs relacionados", "<button type='button' class='ui-button ui-button--light ui-button--sm' id='btnApl' aria-haspopup='dialog' onclick=""DetAbrir('dlgApl', 'valor1')"">+ Relacionar APL</button>")
	Else
		h = "<div class='ui-card'>" & CgCardHeader("APLs relacionados", "<button type='button' class='ui-button ui-button--light ui-button--sm ui-button--disabled' aria-disabled='true' title='Sin permiso para modificar.'>+ Relacionar APL</button>")
	End If
	msg = Request.QueryString("apl") & ""
	If msg = "ok" Then h = h & "<div class='ui-alert ui-alert--success' role='status'>APL " & Server.HTMLEncode(Request.QueryString("idapl") & "") & " relacionado.</div>"
	If msg = "ya" Then h = h & "<div class='ui-alert ui-alert--warning' role='status'>Ese APL ya estaba relacionado con este " & caso & ".</div>"
	If msg = "noexiste" Then h = h & "<div class='ui-alert ui-alert--danger' role='alert'>No existe un APL con el n&uacute;mero " & Server.HTMLEncode(Request.QueryString("idapl") & "") & ".</div>"
	If msg = "error" Then h = h & "<div class='ui-alert ui-alert--danger' role='alert'>No se pudo relacionar el APL: verific&aacute; el n&uacute;mero y tu permiso para modificar.</div>"
	h = h & CgTablaAplAbre()
	Set rzA = ObtenerAplRelac(Cnn, idCaso, tipoRel)
	If Not rzA.eof Then
		Do While Not rzA.eof
			h = h & CgFilaApl(rzA)
			rzA.movenext
		Loop
	Else
		h = h & "<tr><td colspan='6'><div class='ui-empty'>No hay APLs relacionados.</div></td></tr>"
	End If
	h = h & "</tbody></table></div>"
	If puede Then
		' [CG] Misma funcionalidad que en P2: relaciones.js (capturarYProcesarApl) lee el APL desde #valor1 y el nro del caso
		' desde #nid, muestra los datos del APL para confirmar y guarda con agreParam.asp. Sin <form> ni name: si la pantalla
		' esta dentro del formulario de edicion, el campo no se envia con el guardado.
		h = h & "<span id='nid' hidden>" & idCaso & "</span>"
		h = h & "<dialog id='dlgApl' class='ui-dialogo ui-dialogo--chico' aria-labelledby='dlgAplTitulo'>"
		h = h & "<div class='ui-dialogo__head'><h2 class='ui-dialogo__titulo' id='dlgAplTitulo'>Relacionar APL</h2><button type='button' class='ui-dialogo__cerrar' aria-label='Cerrar' onclick=""DetCerrar('dlgApl')"">&times;</button></div>"
		h = h & "<div class='ui-dialogo__body'><div class='ui-field'><label class='ui-field__label' for='valor1'>ID de APL <span class='ui-field__required' aria-hidden='true'> *</span></label>"
		h = h & "<input class='ui-input' id='valor1' type='number' min='1' step='1' inputmode='numeric' required aria-required='true' aria-describedby='id_apl_ayuda'>"
		h = h & "<p class='ui-dialogo__ayuda' id='id_apl_ayuda'>Ingrese el ID del APL (primera columna del listado de APLs, no la secuencia). El APL se relaciona con este " & caso & ".</p></div></div>"
		h = h & "<div class='ui-dialogo__acciones'><button type='button' class='ui-button ui-button--secondary' onclick=""DetCerrar('dlgApl')"">Cancelar</button><button type='button' class='ui-button ui-button--primary' onclick=""DetCerrar('dlgApl');capturarYProcesarApl('" & tipoRel & "');"">Relacionar</button></div>"
		h = h & "</dialog>" & DetScriptDialogos()
	End If
	h = h & "</div>"
	DetTarjetaApl = h
End Function
%>