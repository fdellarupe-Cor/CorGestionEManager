<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/datosDetalles.inc" -->
<!-- #INCLUDE FILE="IncFile/datosApl.inc" -->
<%
'Option Explicit

' Inicialización de variables y conexión
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
xmenu = Request.QueryString("xmenu")
xtmenu = Request.QueryString("xtmenu")
ord = Request.QueryString("ord")
par = Request.QueryString("par") : If par = "" Then par = "asc"

' Parametrización segura de enteros
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
Set rs = GetAPLs(Cnn, cons, cliente, idAplic, idtip, asp, CLng("0" & fec), ord)
If Not rs.EOF Then rs.Move (CantRegMover)

' ==============================================================================
' INICIO DE RENDERIZADO (HTML)
' ==============================================================================
html = headhtml()
html = html & "<body>"

html = html & cabecera()

formsAccion = ""

If cons = "" Then

	' --- HEADER DE FILTROS Y CONTROLES (titulo + filtros + acciones en una sola fila) ---
	html = html & "<div class='titulo-seccion ui-page-header' style='padding: 1rem 2rem;'>"
	html = html & "<div class='ui-toolbar' style='display:flex; justify-content:space-between; align-items:center; gap: 1rem;'>"

	html = html & "<div><h2>APLs</h2></div>"

	html = html & "<form id='xForm' method='post' action='apl.asp?ord=" & ord & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
	html = html & "<div class='ui-toolbar__filters' style='display:flex; gap: 1rem; align-items:center;'>"

	' COMBO CLIENTES USANDO LA FUNCIÓN CENTRALIZADA
	'html = html & "<strong>CLIENTE:</strong>"
	html = html & "<select class='ui-select' onchange=""document.getElementById('xForm').submit();"" id='cli' name='cli'>"
	html = html & "                            <option value=""0"">CLIENTE: TODOS</option>" & vbCrLf
	

	Set rsCombos = GetListaComboExtended(Cnn, "CLIENTES")
	Do While Not rsCombos.EOF
		selCli = "" : If CStr(rsCombos(0)) = CStr(cliente) Then selCli = "selected"
		html = html & "<option value='" & rsCombos(0) & "' " & selCli & ">" & Server.HTMLEncode(rsCombos(1) & "") & "</option>"
		rsCombos.MoveNext
	Loop

	html = html & "</select>"

	' COMBO APLICATIVOS USANDO LA FUNCIÓN CENTRALIZADA
	'html = html & "<strong>APLICATIVO:</strong>"
	html = html & "<select class='ui-select' onchange=""document.getElementById('xForm').submit();"" id='aplic' name='aplic'>"
	html = html & "<option value='0'>APLICATIVO:TODOS</option>"

	Set rsCombos = GetListaComboExtended(Cnn, "APLICATIVOS")
	Do While Not rsCombos.EOF
		selApl = "" : If CStr(rsCombos(0)) = CStr(idAplic) Then selApl = "selected"
		html = html & "<option value='" & rsCombos(0) & "' " & selApl & ">" & Server.HTMLEncode(rsCombos(1) & "") & "</option>"
		rsCombos.MoveNext
	Loop

	html = html & "</select>"

	' COMBO TIPO APL USANDO LA FUNCIÓN CENTRALIZADA
	'html = html & "<strong>TIPO:</strong>"
	html = html & "<select class='ui-select' onchange=""document.getElementById('xForm').submit();"" id='tip' name='tip'>"
	html = html & "<option value='0'>TIPO:TODOS</option>"

	Set rsCombos = GetListaComboExtended(Cnn, "TIPO_APL")
	Do While Not rsCombos.EOF
		selTip = "" : If CStr(rsCombos(0)) = CStr(idtip) Then selTip = "selected"
		html = html & "<option value='" & rsCombos(0) & "' " & selTip & ">" & Server.HTMLEncode(rsCombos(1) & "") & "</option>"
		rsCombos.MoveNext
	Loop

	html = html & "</select>"

	html = html & "</div>" ' Fin ui-toolbar__filters
	html = html & "</form>"

	html = html & "<div class='ui-toolbar__actions' style='display:flex; gap: 1rem;'>"

	If CStr(getPermisos(xmenu, xtmenu, "ALTA")) = "1" Then
		formsAccion = formsAccion & "<form id='formNuevoApl' class='form-oculto' method='post' action='agregar3.asp?tipo=APL&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
		html = html & "<input class='bot ui-button ui-button--primary' type='button' value='NUEVO APL' onclick=""document.getElementById('formNuevoApl').submit();"">"
	Else
		html = html & "<input class='botsp ui-button' type='button' value='CARGAR NUEVO APL' disabled>"
	End If

	formsAccion = formsAccion & "<form id='formVolver' class='form-oculto' method='post' action='index.asp'></form>"
	html = html & "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('formVolver').submit();"">"

	html = html & "</div>" ' Fin ui-toolbar__actions

	html = html & "</div>" ' Fin ui-toolbar
	html = html & "</div>" ' Fin titulo-seccion

End If

' ==============================================================================
' GRILLA DE RESULTADOS
' ==============================================================================
html = html & "<div class='inner ui-main' style='padding:2rem'>"
html = html & "<div class='bloque-datos table-responsive ui-table-wrap'>"
html = html & "<table class='tabla-moderna ui-table' style='width: 100%;'>"

html = html & "<thead><tr>"
html = html & "<th style='width: 40px;'></th>"
html = html & "<th style='width: 40px;'></th>"
html = html & "<th><a class='col-link' style='text-decoration: none;' href='apl.asp?ord=sec&cli=" & cliente & "&par=asc&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&aplic=" & idAplic & "'>SEC</a></th>"
html = html & "<th><a class='col-link' style='text-decoration: none;' href='apl.asp?ord=nom&cli=" & cliente & "&par=asc&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&aplic=" & idAplic & "'>NOMBRE</a></th>"
html = html & "<th>CLIENTE</th>"
html = html & "<th>APLICATIVO</th>"
html = html & "<th style='text-align: center;'>FECHA ENVIO</th>"
If cons = "" Then html = html & "<th>TIPO</th>"
html = html & "</tr></thead>"

html = html & "<tbody>"

PosReg = 0
Do While (Not rs.EOF) And (PosReg < CantRegMostrar)

	' Lógica visual del badge de fecha
	If IsNull(rs(4)) Or rs(4) = 0 Then
		env = "<span class='badge badge-peligro ui-badge ui-badge--danger'>NO ENVIADO</span>"
	Else
		env = "<span class='badge badge-ok ui-badge ui-badge--ok'>" & Fechador(rs(4) + nFecini) & "</span>"
	End If

	UrlDetalle = "detalles2.asp?id=" & rs(0) & "&opc=ver&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&idAplic=" & idAplic
	JsClickFila = "onclick=""document.location='" & UrlDetalle & "'"""

	html = html & "<tr>"

	' CONTROLES DE ACCIÓN
	formsAccion = formsAccion & "<form id='formDel" & rs(0) & "' class='form-oculto' method='post' action='eliminar.asp?opc=apl&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=apl.asp'></form>"
	html = html & "<td style='text-align:center;'>"
	If CStr(getPermisos(xmenu, xtmenu, "ELIM")) = "1" Then
		html = html & "<input class='btn-accion btn-accion-peligro ui-button ui-button--icon ui-button--danger' type='button' value='X' onclick=""borrar('formDel" & rs(0) & "');"">"
	Else
		html = html & "<input class='btn-accion btn-accion-deshabilitado ui-button ui-button--icon' type='button' value='X' disabled>"
	End If
	html = html & "</td>"

	formsAccion = formsAccion & "<form id='formAct" & rs(0) & "' class='form-oculto' method='post' action='detalles2.asp?id=" & rs(0) & "&opc=act&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&idaplic=" & idAplic & "&idtip=" & idtip & "&cantregmover=" & CantRegMover & "'></form>"
	html = html & "<td style='text-align:center;'>"
	If CStr(getPermisos(xmenu, xtmenu, "MODIF")) = "1" Then
		html = html & "<input class='btn-accion btn-accion-primario ui-button ui-button--icon ui-button--primary' type='button' value='A' onclick=""document.getElementById('formAct" & rs(0) & "').submit();"">"
	Else
		html = html & "<input class='btn-accion btn-accion-deshabilitado ui-button ui-button--icon' type='button' value='A' disabled>"
	End If
	html = html & "</td>"

	' DATOS DEL REGISTRO
	html = html & "<td style='text-align: center; cursor:pointer;' " & JsClickFila & ">" & rs("sec") & "</td>"
	html = html & "<td style='cursor:pointer;' " & JsClickFila & ">" & UCase(rs("nombre")) & "</td>"
	html = html & "<td style='cursor:pointer;' " & JsClickFila & ">" & rs("cliente") & "</td>"
	html = html & "<td style='cursor:pointer;' " & JsClickFila & ">" & rs("aplicativo") & "</td>"
	html = html & "<td style='text-align: center; cursor:pointer;' " & JsClickFila & ">" & env & "</td>"

	If cons = "" Then
		html = html & "<td style='cursor:pointer;' " & JsClickFila & ">" & rs("tipo_apl") & "</td>"
	End If

	html = html & "</tr>"

	PosReg = PosReg + 1
	rs.MoveNext
Loop

html = html & "</tbody>"
html = html & "</table>"

' PAGINACIÓN
html = html & "<div class='ui-pagination' style='display: flex; gap: 0.5rem; margin-top: 1rem; justify-content: flex-end;'>"

If CantRegMover <> 0 Then
	formsAccion = formsAccion & "<form id='formAtras' class='form-oculto' method='post' action='apl.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&aplic=" & idAplic & "&cantregmover=" & (CantRegMover - CantRegMostrar) & "&ord=" & ord & "&cli=" & cliente & "&par=" & par & "&cons=" & cons & "'></form>"
	html = html & "<input class='bot ui-button' type='button' value='ATRAS' onclick=""document.getElementById('formAtras').submit();"">"
Else
	html = html & "<input class='botsp ui-button' type='button' value='ATRAS' disabled>"
End If

If Not rs.EOF Then
	formsAccion = formsAccion & "<form id='formSig' class='form-oculto' method='post' action='apl.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&aplic=" & idAplic & "&cantregmover=" & (CantRegMover + CantRegMostrar) & "&ord=" & ord & "&cli=" & cliente & "&par=" & par & "&cons=" & cons & "'></form>"
	html = html & "<input class='bot ui-button' type='button' value='SIGUIENTE' onclick=""document.getElementById('formSig').submit();"">"
Else
	html = html & "<input class='botsp ui-button' type='button' value='SIGUIENTE' disabled>"
End If

html = html & "</div>" ' Fin paginación

html = html & "</div>" ' Fin bloque-datos
html = html & "</div>" ' Fin inner

' --- SCRIPTS ---
html = html & "<script type='text/javascript'>"
html = html & "function borrar(formu) {"
html = html & "    if (confirm('¿Desea eliminar este APL?')) {"
html = html & "        document.getElementById(formu).submit();"
html = html & "    }"
html = html & "}"
html = html & "</script>"

html = html & "<script src='assets/js/jquery.min.js'></script>"
html = html & "<script src='assets/js/skel.min.js'></script>"
html = html & "<script src='assets/js/util.js'></script>"
html = html & "<script src='assets/js/main.js'></script>"

' Inyección de formularios ocultos acumulados
html = html & formsAccion
html = html & "</body>"
html = html & "</html>"

' --- CIERRE DE CONEXION ---
Cnn.Close
Set Cnn = Nothing

Response.Write html
%>
