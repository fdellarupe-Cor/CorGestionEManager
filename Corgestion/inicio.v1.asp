<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/ui_componentes.inc" -->
<%
' =============================================================================
' inicio.asp - Pantalla de inicio de COR Gestion
'
' Responde lo que un trabajador necesita al entrar: que se le vence (Pendientes),
' que espera algo de el (Testing y Solicitudes) y que tiene hoy (Agenda).
'
' DATOS PENDIENTES (se ven en pantalla como "Falta dato"; buscar FALTA DATO en el codigo):
'   1. Casos asignados al usuario, con fecha de ingreso y de compromiso (SLA)
'   2. Anotaciones con testing pendiente (tablas avances / avances_testing)
'   3. Solicitudes al cliente sin respuesta (tabla avances)
' Cuando cada dato exista, se reemplaza "Nothing" por la consulta: el armado ya esta hecho.
'
' TARJETAS OCULTAS hasta tener los logs: "Ultima actividad" y "Ultimos consultados".
' Para activarlas: quitar el atributo hidden de cada <section> (buscar OCULTA).
'
' A CONFIRMAR (buscar CONFIRMAR): clave de log, xmenu/xtmenu de los botones de alta
' y si la agenda debe incluir los recordatorios en copia.
' =============================================================================

Dim Cnn, html, formsAccion, xmenu, xtmenu, usuario, funcionUsr, rsFun
Dim nombreUsuario, primerNombre, saludo, perfil, hoySerial, textoHoy
Dim DIAS, MESES, DIAS_CORTOS
Dim rsCasos, rsTesting, rsSolicitudes, rsAg
Dim porcAmbar, porcRojo
Dim cuentaRojo, cuentaAmbar, cuentaVerde, cuentaInc, cuentaPry, cantCasos
Dim filasCasos, itemsTesting, itemsSolicitudes, itemsAgenda, cantAgenda

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

xmenu = Request.QueryString("xmenu")
xtmenu = Request.QueryString("xtmenu")
' Usuario: la sesion puede venir vacia; si falta, UserLogon() (mismo criterio que index.asp)
usuario = Trim(Session("usuario") & "")
If usuario = "" Then usuario = Trim(UserLogon() & "")
funcionUsr = Trim(Session("FUNC") & "")
If funcionUsr = "" And usuario <> "" Then
	Set rsFun = Exec(Cnn, "sp_cg_Responsables @Modo='id_funcion_id', @IdSeg=" & SqlStr(usuario))
	If Not rsFun.eof Then funcionUsr = Trim(rsFun(1) & "")
	rsFun.Close
End If

' CONFIRMAR clave de log de la pantalla de inicio
'XXX = SYSLOGON(Cnn, "MENU_INICIO")

DIAS = Array("Domingo", "Lunes", "Martes", "Mi&eacute;rcoles", "Jueves", "Viernes", "S&aacute;bado")
DIAS_CORTOS = Array("Dom", "Lun", "Mar", "Mi&eacute;", "Jue", "Vie", "S&aacute;b")
MESES = Array("enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre")

' Hoy en el formato de la base (dias desde 01-01-1900)
hoySerial = CLng(Date) - CLng(CDate("01-01-1900"))
textoHoy = DIAS(Weekday(Date) - 1) & " " & Day(Date) & " de " & MESES(Month(Date) - 1)

' --- Saludo ---
nombreUsuario = Trim(getName(Cnn, "", "cor_responsables", usuario) & "")
primerNombre = ""
Dim partesNombre
If nombreUsuario <> "" Then
	partesNombre = Split(nombreUsuario, " ")
	primerNombre = UiCapitalizar(partesNombre(0))
End If
If Hour(Now) < 12 Then
	saludo = "Buenos d&iacute;as"
ElseIf Hour(Now) < 20 Then
	saludo = "Buenas tardes"
Else
	saludo = "Buenas noches"
End If
If primerNombre <> "" Then saludo = saludo & ", " & UiTexto(primerNombre)

Select Case funcionUsr
	Case "TEC"
		perfil = "T&eacute;cnico"
	Case "FUNC"
		perfil = "Funcional"
	Case Else
		perfil = "&mdash;"
End Select


' =============================================================================
' 1. PENDIENTES: casos asignados al usuario
' =============================================================================
' FALTA DATO: procedimiento con los incidentes y proyectos asignados al usuario. Columnas esperadas:
'   tipo ("I" incidente / "P" proyecto), id, titulo, cliente, prioridad, estado,
'   fecha_ingreso y fecha_compromiso (dias desde 01-01-1900)
' Ejemplo: Set rsCasos = Exec(Cnn, "sp_cg_Inicio @Modo='casos', @IdResp=" & SqlStr(usuario))
Set rsCasos = Nothing

' Semaforo: porcentaje del plazo del SLA consumido para pasar a ambar y a rojo (definido: 50 / 80)
porcAmbar = 50
porcRojo = 80

' Color del semaforo segun el porcentaje del plazo consumido
Function InicioColor(ingreso, compromiso)
	Dim total, consumido
	InicioColor = "gray"
	If porcAmbar = 0 Or porcRojo = 0 Then Exit Function
	total = CLng("0" & compromiso) - CLng("0" & ingreso)
	If total <= 0 Then Exit Function
	consumido = (hoySerial - CLng(ingreso)) * 100 / total
	If consumido >= porcRojo Then
		InicioColor = "red"
	ElseIf consumido >= porcAmbar Then
		InicioColor = "amber"
	Else
		InicioColor = "green"
	End If
End Function

' Texto del vencimiento: "Vencido hace 2 dias", "Vence hoy", "Faltan 4 dias"
Function InicioVence(faltan)
	If faltan < -1 Then
		InicioVence = "Vencido hace " & (-faltan) & " d&iacute;as"
	ElseIf faltan = -1 Then
		InicioVence = "Vencido ayer"
	ElseIf faltan = 0 Then
		InicioVence = "Vence hoy"
	ElseIf faltan = 1 Then
		InicioVence = "Vence ma&ntilde;ana"
	Else
		InicioVence = "Faltan " & faltan & " d&iacute;as"
	End If
End Function

cuentaRojo = 0 : cuentaAmbar = 0 : cuentaVerde = 0 : cuentaInc = 0 : cuentaPry = 0 : cantCasos = 0
filasCasos = ""
If Not rsCasos Is Nothing Then
	Dim colorCaso, tipoCaso, urlCaso, faltanCaso, etiqueta
	Do While Not rsCasos.EOF
		tipoCaso = UCase(rsCasos("tipo") & "")
		colorCaso = InicioColor(rsCasos("fecha_ingreso"), rsCasos("fecha_compromiso"))
		faltanCaso = CLng("0" & rsCasos("fecha_compromiso")) - hoySerial
		If tipoCaso = "P" Then
			urlCaso = "detalles.asp?id=" & rsCasos("id") & "&opc=verp"
			etiqueta = "<span class='ui-tag-id ui-tag-id--pry'>PRY-" & rsCasos("id") & "</span> "
			cuentaPry = cuentaPry + 1
		Else
			urlCaso = "detalles.asp?id=" & rsCasos("id") & "&opc=veri"
			etiqueta = "<span class='ui-tag-id'>INC-" & rsCasos("id") & "</span> "
			cuentaInc = cuentaInc + 1
		End If
		If colorCaso = "red" Then cuentaRojo = cuentaRojo + 1
		If colorCaso = "amber" Then cuentaAmbar = cuentaAmbar + 1
		If colorCaso = "green" Then cuentaVerde = cuentaVerde + 1

		filasCasos = filasCasos & "<tr data-tipo='" & tipoCaso & "' data-color='" & colorCaso & "'>" & _
			"<td><div class='ui-caso'><a class='ui-caso__link' href='" & urlCaso & "'>" & etiqueta & UiTexto(rsCasos("titulo")) & "</a>" & _
			"<span class='ui-caso__meta'>" & UiTexto(rsCasos("cliente")) & "</span></div></td>" & _
			"<td>" & UiBadgePrioridad(rsCasos("prioridad")) & "</td>" & _
			"<td>" & UiBadgeEstado(rsCasos("estado")) & "</td>" & _
			"<td><span class='ui-sem ui-sem--" & colorCaso & "'><span class='ui-sem__punto' aria-hidden='true'></span>" & _
			"<span>" & InicioVence(faltanCaso) & "<span class='ui-sem__fecha'>" & UiFecha(rsCasos("fecha_compromiso")) & "</span></span></span></td>" & _
			"</tr>"
		cantCasos = cantCasos + 1
		rsCasos.MoveNext
	Loop
End If


' =============================================================================
' 2. TESTING PENDIENTES: anotaciones que esperan la validacion del usuario
' =============================================================================
' FALTA DATO: anotaciones con testing pendiente (tablas avances / avances_testing). Columnas esperadas:
'   tipo_entidad ("I"/"P"), id_entidad, texto, responsable, fecha_hora, reintento (1 si el ultimo testing salio incorrecto)
Set rsTesting = Nothing

itemsTesting = ""
If Not rsTesting Is Nothing Then
	Dim urlTesting, tagTesting
	Do While Not rsTesting.EOF
		If UCase(rsTesting("tipo_entidad") & "") = "P" Then
			tagTesting = "<span class='ui-tag-id ui-tag-id--pry'>PRY-" & rsTesting("id_entidad") & "</span> "
			urlTesting = "detalles.asp?id=" & rsTesting("id_entidad") & "&opc=verp"
		Else
			tagTesting = "<span class='ui-tag-id'>INC-" & rsTesting("id_entidad") & "</span> "
			urlTesting = "detalles.asp?id=" & rsTesting("id_entidad") & "&opc=veri"
		End If
		itemsTesting = itemsTesting & "<li class='ui-pendiente'>" & _
			"<p class='ui-pendiente__texto'>" & tagTesting & UiTexto(rsTesting("texto")) & "</p>" & _
			"<span class='ui-pendiente__meta'>Anotaci&oacute;n de " & UiTexto(rsTesting("responsable")) & " &middot; " & UiFecha(rsTesting("fecha_hora")) & "</span>"
		If CLng("0" & rsTesting("reintento")) = 1 Then itemsTesting = itemsTesting & "<span>" & UiBadge("Reintento", "red") & " <span class='ui-pendiente__meta'>el testing anterior sali&oacute; incorrecto</span></span>"
		itemsTesting = itemsTesting & "<div class='ui-pendiente__acciones'><a class='ui-button ui-button--primary ui-button--sm' href='" & urlTesting & "'>Completar testing</a></div></li>"
		rsTesting.MoveNext
	Loop
End If


' =============================================================================
' 3. SOLICITUDES PENDIENTES: pedidos al cliente sin respuesta
' =============================================================================
' FALTA DATO: solicitudes sin respuesta del cliente (tabla avances). Columnas esperadas:
'   tipo_entidad ("I"/"P"), id_entidad, texto, cliente, fecha_hora, insistencias
Set rsSolicitudes = Nothing

itemsSolicitudes = ""
If Not rsSolicitudes Is Nothing Then
	Dim urlSol, tagSol, diasSol
	Do While Not rsSolicitudes.EOF
		If UCase(rsSolicitudes("tipo_entidad") & "") = "P" Then
			tagSol = "<span class='ui-tag-id ui-tag-id--pry'>PRY-" & rsSolicitudes("id_entidad") & "</span> "
			urlSol = "detalles.asp?id=" & rsSolicitudes("id_entidad") & "&opc=verp"
		Else
			tagSol = "<span class='ui-tag-id'>INC-" & rsSolicitudes("id_entidad") & "</span> "
			urlSol = "detalles.asp?id=" & rsSolicitudes("id_entidad") & "&opc=veri"
		End If
		diasSol = hoySerial - CLng("0" & rsSolicitudes("fecha_hora"))
		itemsSolicitudes = itemsSolicitudes & "<li class='ui-pendiente'>" & _
			"<p class='ui-pendiente__texto'>" & tagSol & UiTexto(rsSolicitudes("texto")) & "</p>" & _
			"<span class='ui-pendiente__meta'>" & UiTexto(rsSolicitudes("cliente")) & " &middot; pedido hace " & diasSol & " d&iacute;as"
		If CLng("0" & rsSolicitudes("insistencias")) > 0 Then itemsSolicitudes = itemsSolicitudes & " &middot; " & rsSolicitudes("insistencias") & " insistencias"
		itemsSolicitudes = itemsSolicitudes & "</span><div class='ui-pendiente__acciones'>" & _
			"<a class='ui-button ui-button--secondary ui-button--sm' href='" & urlSol & "'>Registrar insistencia</a>" & _
			"<a class='ui-button ui-button--primary ui-button--sm' href='" & urlSol & "'>Recib&iacute; la respuesta</a></div></li>"
		rsSolicitudes.MoveNext
	Loop
End If


' =============================================================================
' 4. AGENDA: recordatorios del usuario para hoy y los proximos 7 dias (dato existente)
' =============================================================================
' Columnas de sp_cg_Agenda @Modo='listar': 0 id, 1 asunto, 2 dia de aviso, 3 responsable, 4 avisado
' CONFIRMAR: si tambien deben verse los recordatorios en los que el usuario esta en copia
Set rsAg = Exec(Cnn, "sp_cg_Agenda @Modo='listar'")
itemsAgenda = ""
cantAgenda = 0
Dim diaAg, chipAg
Do While Not rsAg.EOF
	diaAg = CLng("0" & rsAg(2))
	If CStr(rsAg(3) & "") = CStr(usuario & "") And diaAg >= hoySerial And diaAg <= hoySerial + 7 And cantAgenda < 5 Then
		If diaAg = hoySerial Then
			chipAg = "<span class='ui-agenda__dia ui-agenda__dia--hoy'>Hoy</span>"
		Else
			chipAg = "<span class='ui-agenda__dia'>" & DIAS_CORTOS(Weekday(CDate(diaAg + CLng(CDate("01-01-1900")))) - 1) & "</span>"
		End If
		itemsAgenda = itemsAgenda & "<li class='ui-agenda__item'>" & chipAg & _
			"<div><p class='ui-agenda__texto'><a class='ui-caso__link' href='agenda.asp?opc=det&id=" & rsAg(0) & "'>" & UiTexto(rsAg(1)) & "</a></p>" & _
			"<div class='ui-agenda__meta'>" & UiFecha(rsAg(2)) & "</div></div></li>"
		cantAgenda = cantAgenda + 1
	End If
	rsAg.MoveNext
Loop


' =============================================================================
' VISTA
' =============================================================================
formsAccion = ""
html = headhtml()
html = html & "<body>"
html = html & "<div id='wrapper' role='banner'>" & cabecera() & "</div>"

html = html & "<main id='contenido'>"

' --- Barra superior: saludo, fecha, perfil y altas rapidas ---
Dim accTop
accTop = ""
If funcionUsr = "FUNC" Or funcionUsr = "TEC" Then
	' CONFIRMAR xmenu/xtmenu que corresponden a las altas desde el inicio
	accTop = accTop & "<a class='ui-button ui-button--light ui-button--sm' href='agregar2.asp?tipo=INC&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>Nuevo incidente</a>"
	accTop = accTop & "<a class='ui-button ui-button--secondary ui-button--sm' href='agregar2.asp?tipo=PROY&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>Nuevo proyecto</a>"
End If
accTop = accTop & "<a class='ui-button ui-button--secondary ui-button--sm' href='agenda.asp?opc=cargar&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>Nuevo recordatorio</a>"

Dim casosAsignados
If rsCasos Is Nothing Then
	casosAsignados = UiFalta("casos asignados")
Else
	casosAsignados = cantCasos
End If

html = html & "<div class='ui-app ui-topbar'><div class='ui-topbar__inner ui-topbar__inner--inicio'>" & _
	"<div class='ui-topbar__top'>" & _
		"<div class='ui-topbar__ident'><span class='ui-topbar__id'>Inicio</span><h1 class='ui-topbar__title'>" & saludo & "</h1></div>" & _
		"<nav class='ui-topbar__actions' aria-label='Acciones r&aacute;pidas'>" & accTop & "</nav>" & _
	"</div>" & _
	"<div class='ui-topbar__bottom'><div class='ui-topbar__meta'>" & _
		"<div><span class='ui-topbar__label'>Hoy</span><span class='ui-topbar__value'>" & textoHoy & "</span></div>" & _
		"<div><span class='ui-topbar__label'>Perfil</span><span class='ui-topbar__value'>" & perfil & "</span></div>" & _
		"<div><span class='ui-topbar__label'>Casos asignados</span><span class='ui-topbar__value'>" & casosAsignados & "</span></div>" & _
	"</div><div></div></div>" & _
	"</div></div>"

html = html & "<div class='ui-app ui-page'><div class='ui-container ui-container--inicio'>"
html = html & "<div class='ui-home'>"

' ===================== Columna principal =====================
html = html & "<div class='ui-home__col'>"

' --- Pendientes ---
html = html & "<section class='ui-card' aria-labelledby='tituloCasos'>"
html = html & "<div class='ui-card__header'><h2 class='ui-card__title' id='tituloCasos' tabindex='-1'>Pendientes</h2>"
html = html & "<div class='ui-semfiltro' role='group' aria-label='Filtrar por vencimiento'>" & _
	"<button type='button' class='ui-semfiltro__op' data-filtro='red' aria-pressed='false'><span class='ui-semfiltro__punto' style='background:#EF4444' aria-hidden='true'></span>Cr&iacute;ticos <span class='ui-semfiltro__cant'>" & cuentaRojo & "</span></button>" & _
	"<button type='button' class='ui-semfiltro__op' data-filtro='amber' aria-pressed='false'><span class='ui-semfiltro__punto' style='background:#F59E0B' aria-hidden='true'></span>Por vencer <span class='ui-semfiltro__cant'>" & cuentaAmbar & "</span></button>" & _
	"<button type='button' class='ui-semfiltro__op' data-filtro='green' aria-pressed='false'><span class='ui-semfiltro__punto' style='background:#22C55E' aria-hidden='true'></span>En plazo <span class='ui-semfiltro__cant'>" & cuentaVerde & "</span></button>" & _
	"</div>"
html = html & "<a class='ui-button ui-button--light ui-button--sm' href='incidentes.asp?opc=listar'>Ver todos</a></div>"

html = html & "<div class='ui-tabbar' role='tablist' aria-label='Tipo de caso'>" & _
	"<button type='button' class='ui-tabbar__link is-active' role='tab' aria-selected='true' data-tipo='todos'>Todos <span class='ui-tabbar__cant'>(" & cantCasos & ")</span></button>" & _
	"<button type='button' class='ui-tabbar__link' role='tab' aria-selected='false' tabindex='-1' data-tipo='I'>Incidentes <span class='ui-tabbar__cant'>(" & cuentaInc & ")</span></button>" & _
	"<button type='button' class='ui-tabbar__link' role='tab' aria-selected='false' tabindex='-1' data-tipo='P'>Proyectos <span class='ui-tabbar__cant'>(" & cuentaPry & ")</span></button>" & _
	"</div>"
html = html & "<div class='ui-filtro-activo' id='filtroActivo' hidden><span id='filtroTexto'></span>" & _
	"<button type='button' class='ui-button ui-button--quiet ui-button--sm' id='quitarFiltro'>Quitar filtros</button></div>"

If rsCasos Is Nothing Then
	html = html & UiFaltaBloque("incidentes y proyectos asignados al usuario, con fecha de ingreso y de compromiso (SLA).")
Else
	html = html & UiTablaAbre("Mis casos asignados, con su vencimiento")
	html = html & "<thead><tr>" & UiTh("Caso", "") & UiTh("Prioridad", "") & UiTh("Estado", "") & UiTh("Vencimiento", "") & "</tr></thead><tbody id='tablaCasos'>"
	If cantCasos = 0 Then
		html = html & UiFilaVacia(4, "No ten&eacute;s casos asignados.")
	Else
		html = html & filasCasos
	End If
	html = html & UiTablaCierra()
	html = html & "<p class='ui-sr-only' aria-live='polite' id='anuncioTabla'></p>"
End If
html = html & "</section>"

' --- Testing y Solicitudes ---
html = html & "<div class='ui-home__dupla'>"

html = html & "<section class='ui-card' aria-labelledby='tituloTesting'><div class='ui-card__header'><h2 class='ui-card__title' id='tituloTesting'>Testing pendientes</h2></div>"
If rsTesting Is Nothing Then
	html = html & UiFaltaBloque("anotaciones que esperan tu testing (historial de avances).")
ElseIf itemsTesting = "" Then
	html = html & "<div class='ui-empty'>No ten&eacute;s testing pendiente.</div>"
Else
	html = html & "<ul class='ui-pendientes'>" & itemsTesting & "</ul>"
End If
html = html & "</section>"

html = html & "<section class='ui-card' aria-labelledby='tituloSolicitudes'><div class='ui-card__header'><h2 class='ui-card__title' id='tituloSolicitudes'>Solicitudes pendientes</h2></div>"
If rsSolicitudes Is Nothing Then
	html = html & UiFaltaBloque("solicitudes al cliente sin respuesta (historial de avances).")
ElseIf itemsSolicitudes = "" Then
	html = html & "<div class='ui-empty'>No hay solicitudes esperando respuesta.</div>"
Else
	html = html & "<ul class='ui-pendientes'>" & itemsSolicitudes & "</ul>"
End If
html = html & "</section>"

html = html & "</div>" ' Fin ui-home__dupla
html = html & "</div>" ' Fin columna principal

' ===================== Columna lateral =====================
html = html & "<aside class='ui-home__col' aria-label='Agenda'>"

' --- Agenda ---
html = html & "<section class='ui-widget' aria-labelledby='tituloAgenda'>" & _
	"<div class='ui-widget__head'><h2 class='ui-widget__title' id='tituloAgenda'>Agenda</h2>" & _
	"<a class='ui-widget__link' href='agenda.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>Ver agenda</a></div>"
If cantAgenda = 0 Then
	html = html & "<div class='ui-empty'>No ten&eacute;s recordatorios para los pr&oacute;ximos 7 d&iacute;as.</div>"
Else
	html = html & "<ul class='ui-agenda'>" & itemsAgenda & "</ul>"
End If
html = html & "</section>"

' --- OCULTA: Ultima actividad (activar cuando exista el registro de eventos) ---
html = html & "<section class='ui-widget' aria-labelledby='tituloActividad' hidden>" & _
	"<div class='ui-widget__head'><h2 class='ui-widget__title' id='tituloActividad'>&Uacute;ltima actividad</h2></div>" & _
	UiFaltaBloque("registro de eventos de los casos del usuario (avances, asignaciones, cambios de fecha).") & _
	"</section>"

' --- OCULTA: Ultimos consultados (activar cuando exista el registro de accesos; ObjAcc podria servir) ---
html = html & "<section class='ui-widget' aria-labelledby='tituloUltimos' hidden>" & _
	"<div class='ui-widget__head'><h2 class='ui-widget__title' id='tituloUltimos'>&Uacute;ltimos consultados</h2></div>" & _
	UiFaltaBloque("registro de los &uacute;ltimos incidentes, proyectos y APLs abiertos por el usuario.") & _
	"</section>"

html = html & "</aside>"

html = html & "</div>" ' Fin ui-home
html = html & "</div></div>" ' Fin contenedor y pagina
html = html & "</main>"

' --- Filtros de la tabla de pendientes: selector de vencimiento, pestanas y Quitar filtros ---
html = html & "<script type='text/javascript'>"
html = html & "(function () {"
html = html & "  var tabla = document.getElementById('tablaCasos');"
html = html & "  var estado = { tipo: 'todos', color: null };"
html = html & "  var NOMBRE = { red: 'Cr\u00edticos', amber: 'Por vencer', green: 'En plazo' };"
html = html & "  var tabs = Array.prototype.slice.call(document.querySelectorAll('[role=tab]'));"
html = html & "  var ops = Array.prototype.slice.call(document.querySelectorAll('[data-filtro]'));"
html = html & "  function aplicar() {"
html = html & "    var visibles = 0, total = 0;"
html = html & "    if (tabla) { Array.prototype.forEach.call(tabla.querySelectorAll('tr[data-tipo]'), function (f) {"
html = html & "      total++; var ok = (estado.tipo === 'todos' || f.getAttribute('data-tipo') === estado.tipo) && (!estado.color || f.getAttribute('data-color') === estado.color);"
html = html & "      f.hidden = !ok; if (ok) visibles++; }); }"
html = html & "    var partes = [];"
html = html & "    if (estado.tipo !== 'todos') partes.push(estado.tipo === 'I' ? 'Incidentes' : 'Proyectos');"
html = html & "    if (estado.color) partes.push(NOMBRE[estado.color]);"
html = html & "    document.getElementById('filtroActivo').hidden = !partes.length;"
html = html & "    document.getElementById('filtroTexto').textContent = 'Filtrando: ' + partes.join(' \u00b7 ') + ' \u00b7 ' + visibles + ' de ' + total;"
html = html & "    var a = document.getElementById('anuncioTabla'); if (a) a.textContent = visibles + ' casos en la lista';"
html = html & "  }"
html = html & "  ops.forEach(function (b) { b.addEventListener('click', function () {"
html = html & "    var c = b.getAttribute('data-filtro'); estado.color = estado.color === c ? null : c;"
html = html & "    ops.forEach(function (x) { x.setAttribute('aria-pressed', String(x.getAttribute('data-filtro') === estado.color)); }); aplicar(); }); });"
html = html & "  function elegir(t, foco) { tabs.forEach(function (x) { var on = x === t; x.classList.toggle('is-active', on); x.setAttribute('aria-selected', String(on)); x.tabIndex = on ? 0 : -1; });"
html = html & "    estado.tipo = t.getAttribute('data-tipo'); aplicar(); if (foco) t.focus(); }"
html = html & "  tabs.forEach(function (t, i) { t.addEventListener('click', function () { elegir(t); });"
html = html & "    t.addEventListener('keydown', function (e) { var j = { ArrowRight: i + 1, ArrowLeft: i - 1, Home: 0, End: tabs.length - 1 }[e.key];"
html = html & "      if (j === undefined) return; e.preventDefault(); elegir(tabs[(j + tabs.length) % tabs.length], true); }); });"
html = html & "  document.getElementById('quitarFiltro').addEventListener('click', function () { estado.color = null;"
html = html & "    ops.forEach(function (x) { x.setAttribute('aria-pressed', 'false'); }); elegir(tabs[0]); document.getElementById('tituloCasos').focus(); });"
html = html & "})();"
html = html & "</script>"

html = html & "<script src='assets/js/jquery.min.js'></script>"
html = html & "<script src='assets/js/skel.min.js'></script>"
html = html & "<script src='assets/js/util.js'></script>"
html = html & "<script src='assets/js/main.js'></script>"

html = html & formsAccion
html = html & "</body></html>"

Cnn.Close
Set Cnn = Nothing

Response.Write html
%>
