<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/ui_componentes.inc" -->
<!-- #INCLUDE FILE="IncFile/datosDetalles.inc" -->
<!-- #INCLUDE FILE="IncFile/datosApl.inc" -->
<%
'Option Explicit

' Inicializacion de variables y conexion
Dim Cnn, cons, opc, xmenu, xtmenu, ord, par, cliente, idAplic, idtip, asp, fec, nFecini
Dim rs, rsCombos, CantRegMostrar, CantRegMover, PosReg, env

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

cons = Request.QueryString("cons")
opc = Request.QueryString("opc")
' [UI] Desde Consultas llega opc=consi: es la vista de consulta (antes se mostraba el listado normal)
' Confirmado con la base: sp_cg_ConsultarEntidad usa @Cons vacio = listado, 'cli'/'apli' = por cliente o
' aplicativo, y cualquier otro valor = consulta (filtra por motivo y fecha de creacion)
If cons = "" And opc = "consi" Then cons = "1"
xmenu = Request.QueryString("xmenu")
xtmenu = Request.QueryString("xtmenu")
ord = Request.QueryString("ord")
par = Request.QueryString("par") : If par = "" Then par = "asc"

' Parametrizacion segura de enteros
cliente = CLng("0" & Request.Form("cli"))
If cliente = 0 Then cliente = CLng("0" & Request.QueryString("cli"))

idAplic = CLng("0" & Request.Form("aplic"))
If idAplic = 0 Then idAplic = CLng("0" & Request.QueryString("aplic"))

idtip = CLng("0" & Request.Form("tip"))
If idtip = 0 Then idtip = CLng("0" & Request.QueryString("tip"))

asp = Request.Form("busc") : If asp = "" Then asp = Request.QueryString("busc")
fec = Request.Form("fec") : If fec = "" Then fec = Request.QueryString("fec")

CantRegMostrar = 10
CantRegMover = CLng("0" & Request.QueryString("CantRegMover"))
If CantRegMover < 0 Then CantRegMover = 0

nFecini = 0
If nFecini < CLng(CDate("01-01-1900")) Then nFecini = CLng(CDate("01-01-1900"))

' Llamada a la capa de datos
'[UI-ANT] Set rs = GetAPLs(Cnn, cons, cliente, idAplic, idtip, asp, CLng("0" & fec), ord)
' [UI] La fecha llega como dd/mm/aaaa: se pasa al formato de la base (CLng("0" & fec) daba error de tipo)
Dim fecSerial
fecSerial = 0
If IsNumeric(fec & "") Then
	fecSerial = CLng(fec)
ElseIf IsDate(fec & "") Then
	fecSerial = CLng(CDate(fec)) - CLng(CDate("01-01-1900"))
End If
Set rs = GetAPLs(Cnn, cons, cliente, idAplic, idtip, asp, fecSerial, ord)
If Not rs.EOF Then rs.Move (CantRegMover)

' ==============================================================================
' INICIO DE RENDERIZADO (HTML)
' ==============================================================================
'[UI-ANT] ' ===== Vista anterior completa: reemplazada por el bloque [UI] que sigue. Se borra entera buscando [UI-ANT] =====
'[UI-ANT] html = headhtml()
'[UI-ANT] html = html & "<body>"

'[UI-ANT] html = html & cabecera()

'[UI-ANT] formsAccion = ""

'[UI-ANT] If cons = "" Then

	'[UI-ANT] ' --- HEADER DE FILTROS Y CONTROLES (titulo + filtros + acciones en una sola fila) ---
	'[UI-ANT] html = html & "<div class='titulo-seccion ui-page-header' style='padding: 1rem 2rem;'>"
	'[UI-ANT] html = html & "<div class='ui-toolbar' style='display:flex; justify-content:space-between; align-items:center; gap: 1rem;'>"

	'[UI-ANT] html = html & "<div><h2>APLs</h2></div>"

	'[UI-ANT] html = html & "<form id='xForm' method='post' action='apl.asp?ord=" & ord & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
	'[UI-ANT] html = html & "<div class='ui-toolbar__filters' style='display:flex; gap: 1rem; align-items:center;'>"

	'[UI-ANT] ' COMBO CLIENTES USANDO LA FUNCION CENTRALIZADA
	'[UI-ANT] 'html = html & "<strong>CLIENTE:</strong>"
	'[UI-ANT] html = html & "<select class='ui-select' onchange=""document.getElementById('xForm').submit();"" id='cli' name='cli'>"
	'[UI-ANT] html = html & "                            <option value=""0"">CLIENTE: TODOS</option>" & vbCrLf
	

	'[UI-ANT] Set rsCombos = GetListaComboExtended(Cnn, "CLIENTES")
	'[UI-ANT] Do While Not rsCombos.EOF
		'[UI-ANT] selCli = "" : If CStr(rsCombos(0)) = CStr(cliente) Then selCli = "selected"
		'[UI-ANT] html = html & "<option value='" & rsCombos(0) & "' " & selCli & ">" & Server.HTMLEncode(rsCombos(1) & "") & "</option>"
		'[UI-ANT] rsCombos.MoveNext
	'[UI-ANT] Loop

	'[UI-ANT] html = html & "</select>"

	'[UI-ANT] ' COMBO APLICATIVOS USANDO LA FUNCION CENTRALIZADA
	'[UI-ANT] 'html = html & "<strong>APLICATIVO:</strong>"
	'[UI-ANT] html = html & "<select class='ui-select' onchange=""document.getElementById('xForm').submit();"" id='aplic' name='aplic'>"
	'[UI-ANT] html = html & "<option value='0'>APLICATIVO:TODOS</option>"

	'[UI-ANT] Set rsCombos = GetListaComboExtended(Cnn, "APLICATIVOS")
	'[UI-ANT] Do While Not rsCombos.EOF
		'[UI-ANT] selApl = "" : If CStr(rsCombos(0)) = CStr(idAplic) Then selApl = "selected"
		'[UI-ANT] html = html & "<option value='" & rsCombos(0) & "' " & selApl & ">" & Server.HTMLEncode(rsCombos(1) & "") & "</option>"
		'[UI-ANT] rsCombos.MoveNext
	'[UI-ANT] Loop

	'[UI-ANT] html = html & "</select>"

	'[UI-ANT] ' COMBO TIPO APL USANDO LA FUNCION CENTRALIZADA
	'[UI-ANT] 'html = html & "<strong>TIPO:</strong>"
	'[UI-ANT] html = html & "<select class='ui-select' onchange=""document.getElementById('xForm').submit();"" id='tip' name='tip'>"
	'[UI-ANT] html = html & "<option value='0'>TIPO:TODOS</option>"

	'[UI-ANT] Set rsCombos = GetListaComboExtended(Cnn, "TIPO_APL")
	'[UI-ANT] Do While Not rsCombos.EOF
		'[UI-ANT] selTip = "" : If CStr(rsCombos(0)) = CStr(idtip) Then selTip = "selected"
		'[UI-ANT] html = html & "<option value='" & rsCombos(0) & "' " & selTip & ">" & Server.HTMLEncode(rsCombos(1) & "") & "</option>"
		'[UI-ANT] rsCombos.MoveNext
	'[UI-ANT] Loop

	'[UI-ANT] html = html & "</select>"

	'[UI-ANT] html = html & "</div>" ' Fin ui-toolbar__filters
	'[UI-ANT] html = html & "</form>"

	'[UI-ANT] html = html & "<div class='ui-toolbar__actions' style='display:flex; gap: 1rem;'>"

	'[UI-ANT] If CStr(getPermisos(xmenu, xtmenu, "ALTA")) = "1" Then
		'[UI-ANT] formsAccion = formsAccion & "<form id='formNuevoApl' class='form-oculto' method='post' action='agregar3.asp?tipo=APL&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
		'[UI-ANT] html = html & "<input class='bot ui-button ui-button--primary' type='button' value='NUEVO APL' onclick=""document.getElementById('formNuevoApl').submit();"">"
	'[UI-ANT] Else
		'[UI-ANT] html = html & "<input class='botsp ui-button' type='button' value='CARGAR NUEVO APL' disabled>"
	'[UI-ANT] End If

	'[UI-ANT] formsAccion = formsAccion & "<form id='formVolver' class='form-oculto' method='post' action='index.asp'></form>"
	'[UI-ANT] html = html & "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('formVolver').submit();"">"

	'[UI-ANT] html = html & "</div>" ' Fin ui-toolbar__actions

	'[UI-ANT] html = html & "</div>" ' Fin ui-toolbar
	'[UI-ANT] html = html & "</div>" ' Fin titulo-seccion

'[UI-ANT] End If

'[UI-ANT] ' ==============================================================================
'[UI-ANT] ' GRILLA DE RESULTADOS
'[UI-ANT] ' ==============================================================================
'[UI-ANT] html = html & "<div class='inner ui-main' style='padding:2rem'>"
'[UI-ANT] html = html & "<div class='bloque-datos table-responsive ui-table-wrap'>"
'[UI-ANT] html = html & "<table class='tabla-moderna ui-table' style='width: 100%;'>"

'[UI-ANT] html = html & "<thead><tr>"
'[UI-ANT] html = html & "<th style='width: 40px;'></th>"
'[UI-ANT] html = html & "<th style='width: 40px;'></th>"
'[UI-ANT] html = html & "<th><a class='col-link' style='text-decoration: none;' href='apl.asp?ord=sec&cli=" & cliente & "&par=asc&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&aplic=" & idAplic & "'>SEC</a></th>"
'[UI-ANT] html = html & "<th><a class='col-link' style='text-decoration: none;' href='apl.asp?ord=nom&cli=" & cliente & "&par=asc&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&aplic=" & idAplic & "'>NOMBRE</a></th>"
'[UI-ANT] html = html & "<th>CLIENTE</th>"
'[UI-ANT] html = html & "<th>APLICATIVO</th>"
'[UI-ANT] html = html & "<th style='text-align: center;'>FECHA ENVIO</th>"
'[UI-ANT] If cons = "" Then html = html & "<th>TIPO</th>"
'[UI-ANT] html = html & "</tr></thead>"

'[UI-ANT] html = html & "<tbody>"

'[UI-ANT] PosReg = 0
'[UI-ANT] Do While (Not rs.EOF) And (PosReg < CantRegMostrar)

	'[UI-ANT] ' Logica visual del badge de fecha
	'[UI-ANT] If IsNull(rs(4)) Or rs(4) = 0 Then
		'[UI-ANT] env = "<span class='badge badge-peligro ui-badge ui-badge--danger'>NO ENVIADO</span>"
	'[UI-ANT] Else
		'[UI-ANT] env = "<span class='badge badge-ok ui-badge ui-badge--ok'>" & Fechador(rs(4) + nFecini) & "</span>"
	'[UI-ANT] End If

	'[UI-ANT] UrlDetalle = "detalles2.asp?id=" & rs(0) & "&opc=ver&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&idAplic=" & idAplic
	'[UI-ANT] JsClickFila = "onclick=""document.location='" & UrlDetalle & "'"""

	'[UI-ANT] html = html & "<tr>"

	'[UI-ANT] ' CONTROLES DE ACCION
	'[UI-ANT] formsAccion = formsAccion & "<form id='formDel" & rs(0) & "' class='form-oculto' method='post' action='eliminar.asp?opc=apl&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=apl.asp'></form>"
	'[UI-ANT] html = html & "<td style='text-align:center;'>"
	'[UI-ANT] If CStr(getPermisos(xmenu, xtmenu, "ELIM")) = "1" Then
		'[UI-ANT] html = html & "<input class='btn-accion btn-accion-peligro ui-button ui-button--icon ui-button--danger' type='button' value='X' onclick=""borrar('formDel" & rs(0) & "');"">"
	'[UI-ANT] Else
		'[UI-ANT] html = html & "<input class='btn-accion btn-accion-deshabilitado ui-button ui-button--icon' type='button' value='X' disabled>"
	'[UI-ANT] End If
	'[UI-ANT] html = html & "</td>"

	'[UI-ANT] formsAccion = formsAccion & "<form id='formAct" & rs(0) & "' class='form-oculto' method='post' action='detalles2.asp?id=" & rs(0) & "&opc=act&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&idaplic=" & idAplic & "&idtip=" & idtip & "&cantregmover=" & CantRegMover & "'></form>"
	'[UI-ANT] html = html & "<td style='text-align:center;'>"
	'[UI-ANT] If CStr(getPermisos(xmenu, xtmenu, "MODIF")) = "1" Then
		'[UI-ANT] html = html & "<input class='btn-accion btn-accion-primario ui-button ui-button--icon ui-button--primary' type='button' value='A' onclick=""document.getElementById('formAct" & rs(0) & "').submit();"">"
	'[UI-ANT] Else
		'[UI-ANT] html = html & "<input class='btn-accion btn-accion-deshabilitado ui-button ui-button--icon' type='button' value='A' disabled>"
	'[UI-ANT] End If
	'[UI-ANT] html = html & "</td>"

	'[UI-ANT] ' DATOS DEL REGISTRO
	'[UI-ANT] html = html & "<td style='text-align: center; cursor:pointer;' " & JsClickFila & ">" & rs("sec") & "</td>"
	'[UI-ANT] html = html & "<td style='cursor:pointer;' " & JsClickFila & ">" & UCase(rs("nombre")) & "</td>"
	'[UI-ANT] html = html & "<td style='cursor:pointer;' " & JsClickFila & ">" & rs("cliente") & "</td>"
	'[UI-ANT] html = html & "<td style='cursor:pointer;' " & JsClickFila & ">" & rs("aplicativo") & "</td>"
	'[UI-ANT] html = html & "<td style='text-align: center; cursor:pointer;' " & JsClickFila & ">" & env & "</td>"

	'[UI-ANT] If cons = "" Then
		'[UI-ANT] html = html & "<td style='cursor:pointer;' " & JsClickFila & ">" & rs("tipo_apl") & "</td>"
	'[UI-ANT] End If

	'[UI-ANT] html = html & "</tr>"

	'[UI-ANT] PosReg = PosReg + 1
	'[UI-ANT] rs.MoveNext
'[UI-ANT] Loop

'[UI-ANT] html = html & "</tbody>"
'[UI-ANT] html = html & "</table>"

'[UI-ANT] ' PAGINACION
'[UI-ANT] html = html & "<div class='ui-pagination' style='display: flex; gap: 0.5rem; margin-top: 1rem; justify-content: flex-end;'>"

'[UI-ANT] If CantRegMover <> 0 Then
	'[UI-ANT] formsAccion = formsAccion & "<form id='formAtras' class='form-oculto' method='post' action='apl.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&aplic=" & idAplic & "&cantregmover=" & (CantRegMover - CantRegMostrar) & "&ord=" & ord & "&cli=" & cliente & "&par=" & par & "&cons=" & cons & "'></form>"
	'[UI-ANT] html = html & "<input class='bot ui-button' type='button' value='ATRAS' onclick=""document.getElementById('formAtras').submit();"">"
'[UI-ANT] Else
	'[UI-ANT] html = html & "<input class='botsp ui-button' type='button' value='ATRAS' disabled>"
'[UI-ANT] End If

'[UI-ANT] If Not rs.EOF Then
	'[UI-ANT] formsAccion = formsAccion & "<form id='formSig' class='form-oculto' method='post' action='apl.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&aplic=" & idAplic & "&cantregmover=" & (CantRegMover + CantRegMostrar) & "&ord=" & ord & "&cli=" & cliente & "&par=" & par & "&cons=" & cons & "'></form>"
	'[UI-ANT] html = html & "<input class='bot ui-button' type='button' value='SIGUIENTE' onclick=""document.getElementById('formSig').submit();"">"
'[UI-ANT] Else
	'[UI-ANT] html = html & "<input class='botsp ui-button' type='button' value='SIGUIENTE' disabled>"
'[UI-ANT] End If

'[UI-ANT] html = html & "</div>" ' Fin paginacion

'[UI-ANT] html = html & "</div>" ' Fin bloque-datos
'[UI-ANT] html = html & "</div>" ' Fin inner

'[UI-ANT] ' --- SCRIPTS ---
'[UI-ANT] html = html & "<script type='text/javascript'>"
'[UI-ANT] html = html & "function borrar(formu) {"
'[UI-ANT] html = html & "    if (confirm('?Desea eliminar este APL?')) {"
'[UI-ANT] html = html & "        document.getElementById(formu).submit();"
'[UI-ANT] html = html & "    }"
'[UI-ANT] html = html & "}"
'[UI-ANT] html = html & "</script>"

' =========================================================================
' [UI] VISTA con el sistema visual (IncFile/ui_componentes.inc)
' El bloque anterior completo queda arriba, comentado con '[UI-ANT].
' Carga de datos, formularios, URLs y permisos: los mismos de antes.
' =========================================================================
html = headhtml()
html = html & "<body>"

html = html & cabecera()

formsAccion = ""

' Filtros que se conservan al ordenar y al paginar
' (antes se perdia el tipo, y en la consulta tambien la busqueda, la fecha y la propia vista de consulta)
Dim qsFiltros, accTop, filtrosHtml, selHtml, hayMas, pagAnt, pagSig, irDetalle, envio, accFila, colsApl
qsFiltros = "&cli=" & cliente & "&aplic=" & idAplic & "&tip=" & idtip & "&cons=" & cons & "&busc=" & Server.URLEncode(asp & "") & "&fec=" & Server.URLEncode(fec & "") & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu

If cons = "" Then

	' --- Barra superior: Cargar y Volver ---
	Dim puedeAlta
	puedeAlta = (CStr(getPermisosAtrm("CG_APL", "ALTA")) = "1")
	If puedeAlta Then
		formsAccion = formsAccion & "<form id='formNuevoApl' class='form-oculto' method='post' action='agregar3.asp?tipo=APL&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
	End If
	accTop = UiBoton("Cargar APL", "light", "document.getElementById('formNuevoApl').submit();", puedeAlta, "Sin permiso para agregar.")
	formsAccion = formsAccion & "<form id='formVolver' class='form-oculto' method='post' action='index.asp'></form>"
	accTop = accTop & UiBoton("Volver", "secondary", "document.getElementById('formVolver').submit();", True, "")

	' --- Filtros en el encabezado de la tarjeta: se envian solos al cambiar, como antes ---
	filtrosHtml = "<form id='xForm' class='ui-card__filtros' method='post' action='apl.asp?ord=" & ord & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"

	selHtml = "<select class='ui-select' onchange=""document.getElementById('xForm').submit();"" id='cli' name='cli'><option value='0'>Todos los clientes</option>"
	Set rsCombos = GetListaComboExtended(Cnn, "CLIENTES")
	Do While Not rsCombos.EOF
		selCli = ""
		If CStr(rsCombos(0)) = CStr(cliente) Then selCli = "selected"
		selHtml = selHtml & "<option value='" & rsCombos(0) & "' " & selCli & ">" & Server.HTMLEncode(rsCombos(1) & "") & "</option>"
		rsCombos.MoveNext
	Loop
	filtrosHtml = filtrosHtml & UiFiltroCompacto("cli", "Filtrar por cliente", selHtml & "</select>")

	selHtml = "<select class='ui-select' onchange=""document.getElementById('xForm').submit();"" id='aplic' name='aplic'><option value='0'>Todos los aplicativos</option>"
	Set rsCombos = GetListaComboExtended(Cnn, "APLICATIVOS")
	Do While Not rsCombos.EOF
		selApl = ""
		If CStr(rsCombos(0)) = CStr(idAplic) Then selApl = "selected"
		selHtml = selHtml & "<option value='" & rsCombos(0) & "' " & selApl & ">" & Server.HTMLEncode(rsCombos(1) & "") & "</option>"
		rsCombos.MoveNext
	Loop
	filtrosHtml = filtrosHtml & UiFiltroCompacto("aplic", "Filtrar por aplicativo", selHtml & "</select>")

	selHtml = "<select class='ui-select' onchange=""document.getElementById('xForm').submit();"" id='tip' name='tip'><option value='0'>Todos los tipos</option>"
	Set rsCombos = GetListaComboExtended(Cnn, "TIPO_APL")
	Do While Not rsCombos.EOF
		selTip = ""
		If CStr(rsCombos(0)) = CStr(idtip) Then selTip = "selected"
		selHtml = selHtml & "<option value='" & rsCombos(0) & "' " & selTip & ">" & Server.HTMLEncode(rsCombos(1) & "") & "</option>"
		rsCombos.MoveNext
	Loop
	filtrosHtml = filtrosHtml & UiFiltroCompacto("tip", "Filtrar por tipo", selHtml & "</select>")

	filtrosHtml = filtrosHtml & "</form>"

	html = html & UiPantallaAbre("APLs", accTop, "wide")
	html = html & UiCardAbre("Listado de APLs", filtrosHtml)

Else
	' Consulta: sin filtros propios (llegan desde la pantalla de consulta). Volver regresa a Consultas, pestana APL
	formsAccion = formsAccion & "<form id='formVolver' class='form-oculto' method='post' action='consultas.asp?sel=2&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
	accTop = UiBoton("Volver", "secondary", "document.getElementById('formVolver').submit();", True, "")
	html = html & UiPantallaAbre("Consulta de APLs", accTop, "wide")
	html = html & UiCardAbre("Resultados", "")
End If

' --- Tabla ---
html = html & UiTablaAbre("Listado de APLs. Se puede ordenar por secuencia o por nombre.")
html = html & "<thead><tr>"
html = html & UiTh("<span class='ui-sr-only'>Acciones</span>", "ui-table__col--actions")
html = html & UiTh("ID", "")
html = html & UiThOrden("Sec.", "apl.asp?ord=sec&par=asc" & qsFiltros, ord = "sec")
html = html & UiThOrden("Nombre", "apl.asp?ord=nom&par=asc" & qsFiltros, ord = "nom")
html = html & UiTh("Cliente", "") & UiTh("Aplicativo", "") & UiTh("Env&iacute;o", "")
If cons = "" Then html = html & UiTh("Tipo", "")
html = html & "</tr></thead><tbody>"

PosReg = 0
Do While (Not rs.EOF) And (PosReg < CantRegMostrar)

	' Fecha de envio legible (Fechador(fecha + nFecini) mostraba el numero crudo)
	If IsNull(rs(4)) Or rs(4) = 0 Then
		envio = UiBadge("No enviado", "red")
	Else
		envio = "<span class='ui-table__nowrap'>" & UiFecha(rs(4)) & "</span>"
	End If

	UrlDetalle = "detalles2.asp?id=" & rs(0) & "&opc=ver&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&idAplic=" & idAplic
	' [UI] Desde Consultas: abre la vista de consulta de APL (solo lectura); Volver regresa con los filtros
	If cons <> "" Then UrlDetalle = "vistaconsulta.asp?tipo=A&id=" & rs(0) & "&cli=" & cliente & "&aplic=" & idAplic & "&busc=" & Server.URLEncode(asp & "") & "&fec=" & Server.URLEncode(fec & "") & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu
	irDetalle = " class='ui-table__cell--link' onclick=""document.location='" & UrlDetalle & "'"""

	formsAccion = formsAccion & "<form id='formDel" & rs(0) & "' class='form-oculto' method='post' action='eliminar.asp?opc=apl&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=apl.asp'></form>"
	formsAccion = formsAccion & "<form id='formAct" & rs(0) & "' class='form-oculto' method='post' action='detalles2.asp?id=" & rs(0) & "&opc=act&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&idaplic=" & idAplic & "&idtip=" & idtip & "&cantregmover=" & CantRegMover & "'></form>"

	accFila = UiBotonIcono("&times;", "Eliminar APL " & rs(0), "danger", "borrar('formDel" & rs(0) & "');", CStr(getPermisosAtrm("CG_APL", "ELIM")) = "1", "Sin permiso para eliminar")
	accFila = accFila & UiBotonIcono("&#9998;", "Editar APL " & rs(0), "secondary", "document.getElementById('formAct" & rs(0) & "').submit();", CStr(getPermisosAtrm("CG_APL", "MODIF")) = "1", "Sin permiso para modificar")

	html = html & "<tr>"
	html = html & "<td><div class='ui-table__actions'>" & accFila & "</div></td>"
	html = html & "<td" & irDetalle & ">" & UiDato(rs(0)) & "</td>"
	html = html & "<td" & irDetalle & ">" & UiDato(rs("sec")) & "</td>"
	html = html & "<td>" & UiCeldaEnlace(UiTexto(rs("nombre")), UrlDetalle, UiTexto(rs("nombre"))) & "</td>"
	html = html & "<td" & irDetalle & ">" & UiTexto(rs("cliente")) & "</td>"
	html = html & "<td" & irDetalle & ">" & UiTexto(rs("aplicativo")) & "</td>"
	html = html & "<td" & irDetalle & ">" & envio & "</td>"
	If cons = "" Then html = html & "<td" & irDetalle & ">" & UiDato(rs("tipo_apl")) & "</td>"
	html = html & "</tr>"

	PosReg = PosReg + 1
	rs.MoveNext
Loop

colsApl = 7
If cons <> "" Then colsApl = 6
If PosReg = 0 Then html = html & UiFilaVacia(colsApl, "No hay APLs con estos filtros.")
html = html & UiTablaCierra()

' --- Paginacion (mismas URLs de antes, mas los filtros que se perdian) ---
hayMas = Not rs.EOF
If CantRegMover <> 0 Then
	formsAccion = formsAccion & "<form id='formAtras' class='form-oculto' method='post' action='apl.asp?cantregmover=" & (CantRegMover - CantRegMostrar) & "&ord=" & ord & "&par=" & par & qsFiltros & "'></form>"
	pagAnt = UiBoton("Anterior", "secondary", "document.getElementById('formAtras').submit();", True, "")
Else
	pagAnt = UiBoton("Anterior", "secondary", "", False, "Primera p&aacute;gina")
End If
If hayMas Then
	formsAccion = formsAccion & "<form id='formSig' class='form-oculto' method='post' action='apl.asp?cantregmover=" & (CantRegMover + CantRegMostrar) & "&ord=" & ord & "&par=" & par & qsFiltros & "'></form>"
	pagSig = UiBoton("Siguiente", "secondary", "document.getElementById('formSig').submit();", True, "")
Else
	pagSig = UiBoton("Siguiente", "secondary", "", False, "&Uacute;ltima p&aacute;gina")
End If
html = html & UiPaginacion(CantRegMover + 1, CantRegMover + PosReg, pagAnt, pagSig)

html = html & UiCardCierra()
html = html & UiPantallaCierra()

html = html & "<script type='text/javascript'>"
html = html & "function borrar(formu) {"
html = html & "    if (confirm('\u00bfDesea eliminar este APL?')) {"
html = html & "        document.getElementById(formu).submit();"
html = html & "    }"
html = html & "}"
html = html & "</script>"

html = html & "<script src='assets/js/jquery.min.js'></script>"
html = html & "<script src='assets/js/skel.min.js'></script>"
html = html & "<script src='assets/js/util.js'></script>"
html = html & "<script src='assets/js/main.js'></script>"

' Inyeccion de formularios ocultos acumulados
html = html & formsAccion
html = html & "</body>"
html = html & "</html>"

' --- CIERRE DE CONEXION ---
Cnn.Close
Set Cnn = Nothing

Response.Write html
%>
