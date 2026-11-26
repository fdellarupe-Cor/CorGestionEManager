<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/datosDetalles.inc" -->
<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/ui_componentes.inc" -->
<%
' [CG] Control de usuario E-Manager (igual que CorEManager/Acciones.asp)
cgOk = ofv.CheckUsuario()
If cgOk <> "Y" Then
	cgUv = ofv.uv
	If cgUv = "NO" Then cgOk = ofv.ValidarUsuario
	Response.Write cgOk
	Response.End
End If

' =============================================================================
' vistaconsulta.asp - Vista de consulta (solo lectura) de un incidente o proyecto
'
' Para quien entra por Consultas: un resumen general del caso, sin nada editable.
'   vistaconsulta.asp?tipo=I&id=4521   (incidente)
'   vistaconsulta.asp?tipo=P&id=153    (proyecto)
'   vistaconsulta.asp?tipo=A&id=2240   (APL: ciclo, arreglos y relacionados)
' Volver regresa a los resultados de la consulta con los mismos filtros
' (cli, est, apli, fec, desc), que llegan en la URL como hoy llegan a detalles.asp.
'
' Tarjetas: Estado actual -> Ciclo de vida + Tiempo consumido (incidente) u Horas (proyecto)
'           -> Detalle del proyecto (solo proyecto) -> Historial de avances
'
' NOTAS:
'   - El historial (tipo de avance, estado, via y avance de origen) es el de detalles.asp, en solo lectura: aqui no se ofrece
'     Insistir ni cambiar el estado de un avance (eso se activa solo en detalles.asp, con permiso de modificar).
'   - Tiempo funcional y de desarrollo del incidente: CG_INCIDENTES.TIEMPO_FUNC y TIEMPO_DESA todavia no se cargan; la tarjeta
'     "Tiempo consumido" muestra un texto neutro hasta que existan.
'   - Plazo del SLA del cliente (el SLA se trabaja aparte): mientras tanto, los plazos fijos de hoy.
' El ultimo avance sale de CG_AVANCES (sp_cg_ObtenerComentarios): el ultimo avance REAL (las filas iniciales no tienen ID).
' =============================================================================

Dim Cnn, html, id, tipo, xmenu, xtmenu, nFecini, porcAmbar, porcRojo
Dim rs, titulo, tabla, cliente, prioridad, estFun, estDes, respG, respT, aplic, via
Dim fCrea, fEntrega, okDes, okFun, okCli, solicitud, reqPres, descripcion
Dim hRel, hDis, hDes, hTes, accTop, formVolver, metas, rutaArchivos
Dim idProyRel, solicProy, tituloProy, ambVal, urlAbrir, qsFiltros, nomFiltro

html = headhtml()
html = html & "<body>"

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

id = Request.QueryString("id")
tipo = UCase(Request.QueryString("tipo") & "")
If tipo <> "P" And tipo <> "A" Then tipo = "I"
xmenu = Request.QueryString("xmenu")
xtmenu = Request.QueryString("xtmenu")

nFecini = CLng(CDate("01-01-1900"))
' Semaforo: porcentaje del plazo consumido para pasar a ambar y a rojo (definido: 50 / 80)
porcAmbar = 50
porcRojo = 80

' --- Volver: a los resultados de la consulta, con los mismos filtros ---
If tipo = "P" Then
	formVolver = "<form id='formVolver' class='form-oculto' method='post' action='proyectos.asp?opc=consi&sel=1&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>" & _
		ConsOculto("cli") & ConsOculto("est") & ConsOculto("apli") & ConsOculto("fec") & ConsOculto("desc") & "</form>"
ElseIf tipo = "A" Then
	formVolver = "<form id='formVolver' class='form-oculto' method='post' action='apl.asp?opc=consi&sel=2&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>" & _
		ConsOculto("cli") & ConsOculto("aplic") & ConsOculto("busc") & ConsOculto("fec") & "</form>"
Else
	formVolver = "<form id='formVolver' class='form-oculto' method='post' action='incidentes.asp?opc=consi&sel=0&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>" & _
		ConsOculto("cli") & ConsOculto("est") & ConsOculto("apli") & ConsOculto("fec") & ConsOculto("desc") & "</form>"
End If
accTop = "<span class='ui-topbar__view'>Vista de consulta</span>" & UiBoton("Volver", "secondary", "document.getElementById('formVolver').submit();", True, "")

html = html & "<div id='wrapper' role='banner'>" & cabecera() & "</div>"

' --- El id llega en la URL: se valida antes de consultar ---
If Not IsNumeric(id & "") Or Trim(id & "") = "" Then
	ConsSinCaso "El enlace no indica qu&eacute; caso mostrar."
End If
id = CLng(id)
' [LOG] Consulta del caso (CG_LOG)
CgLog IIf(tipo = "P", "PROYECTO", IIf(tipo = "A", "APL", "INCIDENTE")), id, "CONSULTA"

' [CONSULTA] Enlace al caso (pantalla de trabajo): detalles.asp (incidente / proyecto) o detalles2.asp (APL). Lleva vc=1 y los filtros de la
' consulta para que "Volver" del detalle regrese a esta vista de consulta (con los mismos filtros).
qsFiltros = ""
For Each nomFiltro In Array("cli", "est", "apli", "fec", "desc", "busc", "aplic")
	If Trim(Request.QueryString(nomFiltro) & "") <> "" Then qsFiltros = qsFiltros & "&" & nomFiltro & "=" & Server.URLEncode(Request.QueryString(nomFiltro) & "")
Next
If tipo = "P" Then
	urlAbrir = "detalles.asp?id=" & id & "&opc=verp&vc=1&xmenu=" & xmenu & "&xtmenu=" & xtmenu & qsFiltros
	accTop = UiBoton("Abrir proyecto", "light", "document.location='" & urlAbrir & "';", True, "") & accTop
ElseIf tipo = "A" Then
	urlAbrir = "detalles2.asp?id=" & id & "&opc=ver&vc=1&xmenu=" & xmenu & "&xtmenu=" & xtmenu & qsFiltros
	accTop = UiBoton("Abrir APL", "light", "document.location='" & urlAbrir & "';", True, "") & accTop
Else
	urlAbrir = "detalles.asp?id=" & id & "&opc=veri&vc=1&xmenu=" & xmenu & "&xtmenu=" & xtmenu & qsFiltros
	accTop = UiBoton("Abrir incidente", "light", "document.location='" & urlAbrir & "';", True, "") & accTop
End If

' APL: vista de consulta propia (se arma completa en ConsPaginaApl)
If tipo = "A" Then ConsPaginaApl id

If tipo = "P" Then
	zzz = ObjAcc(id,"P",cnn,"D")
	XXX = SYSLOGON(cnn,"CONSULTA_PROY_" & id)
	' [CG] Permiso de consulta de E-Manager (Atrm), igual que las listas de Consultas
	If getPermisosAtrm("CG_PROYECTOS", "CONS") <> 1 Then ConsSinCaso "No ten&eacute;s permiso para consultar proyectos."
	Set rs = ObtenerDetalleProyecto(Cnn, id)
	tabla = "proyectos"
Else
	zzz = ObjAcc(id,"I",cnn,"D")
	XXX = SYSLOGON(cnn,"CONSULTA_INC_" & id)
	' [CG] Permiso de consulta de E-Manager (Atrm), igual que las listas de Consultas
	If getPermisosAtrm("CG_INCIDENTES", "CONS") <> 1 Then ConsSinCaso "No ten&eacute;s permiso para consultar incidentes."
	Set rs = ObtenerDetalleIncidente(Cnn, id)
	tabla = "incidentes"
End If

If rs.eof Then
	If tipo = "P" Then
		ConsSinCaso "No se encontr&oacute; el proyecto " & id & "."
	Else
		ConsSinCaso "No se encontr&oacute; el incidente " & id & "."
	End If
End If

titulo = getName(cnn, "titulo", tabla, id)

' --- Datos del caso (mismas columnas que usa detalles.asp) ---
If tipo = "P" Then
	solicitud = rs(1) : descripcion = rs(3) : cliente = rs(4) : aplic = rs(6)
	respG = rs(7) : respT = rs(8) : fCrea = rs(9) : fEntrega = rs(10)
	estDes = rs(11) : estFun = rs(12) : prioridad = rs(13)
	hRel = rs(14) : hDis = rs(15) : hDes = rs(16) : hTes = rs(17)
	via = rs(19) : reqPres = rs(20) : okDes = rs(23) : okFun = rs(24) : okCli = rs(25)

Else
	cliente = rs(1) : aplic = rs(3) : prioridad = rs(5) : respG = rs(6) : respT = rs(7)
	estDes = rs(9) : estFun = rs(10) : fCrea = rs(11) : via = rs(13)
	okDes = rs(16) : okFun = rs(17) : okCli = rs(18)
	hRel = rs(28)
	hDes = rs(29)
	' Ambiente y proyecto relacionado (columnas del detalle: 14 id del proyecto, 25 solicitud, 26 titulo; AMBIENTE por nombre)
	idProyRel = rs(14) : solicProy = rs(25) : tituloProy = rs(26) : ambVal = AmbienteValor(rs)
	
End If

' =============================================================================
' BARRA SUPERIOR: identificacion, datos clave y semaforo
' =============================================================================
Dim fIni, fLim
fIni = ConsFechaDb(fCrea)
fLim = Empty
metas = ConsMeta("Cliente", UiDato(cliente)) & ConsMeta("Fecha creaci&oacute;n", UiFecha(fCrea))
If tipo = "P" Then
	fLim = ConsFechaDb(fEntrega)
	metas = metas & ConsMeta("Fecha entrega", UiFecha(fEntrega))
Else
	' FALTA DATO: plazo del SLA del cliente segun la prioridad (tabla SLA). Mientras tanto, los plazos fijos de hoy
	If Not IsEmpty(fIni) And ConsPlazoDias(prioridad) > 0 Then fLim = DateAdd("d", ConsPlazoDias(prioridad), fIni)
End If
metas = metas & ConsMeta("Prioridad", UiBadgePrioridad(prioridad))

html = html & "<main id='contenido'>"
html = html & "<div class='ui-app ui-topbar'><div class='ui-topbar__inner'>" & _
	"<div class='ui-topbar__top'>" & _
		"<div class='ui-topbar__ident'><span class='ui-topbar__id'><a href='" & urlAbrir & "' style='color:inherit' title='Abrir en la pantalla de trabajo'>" & ConsEtiqueta(tipo, id) & "</a></span><h1 class='ui-topbar__title'>" & UiTexto(titulo) & "</h1></div>" & _
		"<div class='ui-topbar__actions'>" & accTop & "</div>" & _
	"</div>" & _
	"<div class='ui-topbar__bottom'><div class='ui-topbar__meta'>" & metas & "</div>" & _
		ConsSemaforo(fIni, fLim, okCli, ConsTextoPlazo(tipo)) & "</div>" & _
	"</div></div>"

html = html & "<div class='ui-app ui-page'><div class='ui-container'>"

' =============================================================================
' ESTADO ACTUAL
' =============================================================================
html = html & "<section class='ui-card' aria-labelledby='tituloEstado'><h2 class='ui-card__title' id='tituloEstado'>Estado actual</h2><div class='ui-fields'>"
html = html & UiCampoLectura("Estado funcional", UiBadgeEstado(estFun), False)
html = html & UiCampoLectura("Estado desarrollo", UiBadgeEstado(estDes), False)
html = html & UiCampoLectura("Responsable de gesti&oacute;n", UiDato(respG), False)
html = html & UiCampoLectura("Responsable t&eacute;cnico", UiDato(respT), False)
html = html & UiCampoLectura("Aplicaci&oacute;n", UiDato(aplic), False)
If tipo = "I" Then
	html = html & UiCampoLectura("Ambiente", AmbienteSelectHtml("", ambVal, False), False)
	html = html & UiCampoLectura("Proyecto relacionado", ConsProyectoRel(idProyRel, solicProy, tituloProy), False)
End If
If tipo = "P" Then html = html & UiCampoLectura("Solicitud n&uacute;mero", UiDato(solicitud), False)
html = html & UiCampoLectura("V&iacute;a de pedido", UiDato(via), False)
If tipo = "P" Then html = html & UiCampoLectura("Requiere presupuesto", UiDato(reqPres), False)
html = html & "</div></section>"

' =============================================================================
' CICLO DE VIDA  +  TIEMPO CONSUMIDO (incidente) / HORAS (proyecto)
' =============================================================================
html = html & "<div class='ui-grid-2'>"

html = html & "<section class='ui-card' aria-labelledby='tituloCiclo'><h2 class='ui-card__title' id='tituloCiclo'>Ciclo de vida del " & ConsNombre(tipo) & "</h2><div class='ui-lifecycle'>"
html = html & ConsHito("Creaci&oacute;n", fCrea)
html = html & ConsUltimoAvance(id, tipo)
html = html & ConsHito("Desarrollo finalizado", okDes)
html = html & ConsHito("Funcional finalizado", okFun)
If tipo = "P" Then
	' La entrega comprometida es una meta: muestra su fecha, pero se cumple recien con la entrega (OK cliente)
	html = html & ConsHitoMeta("Entrega comprometida", fEntrega, ConsNum(okCli) > 0)
	html = html & ConsHito("Proyecto entregado", okCli)
Else
	html = html & ConsHito("Incidente solucionado", okCli)
End If
html = html & "</div></section>"

If tipo = "P" Then
	html = html & "<section class='ui-card' aria-labelledby='tituloHoras'><h2 class='ui-card__title' id='tituloHoras'>Horas del proyecto</h2>"
	html = html & ConsDonut(Array("Relevamiento y definici&oacute;n", "Dise&ntilde;o", "Desarrollo", "Testeo"), _
		Array(ConsNum(hRel), ConsNum(hDis), ConsNum(hDes), ConsNum(hTes)), _
		Array("#7995FF", "#6D28D9", "#132574", "#0E7490"), "Horas del proyecto")
	html = html & "</section>"
Else
	' Tiempo funcional y de desarrollo del incidente (CG_INCIDENTES.TIEMPO_FUNC / TIEMPO_DESA): todavia no se cargan; mientras tanto
	' se muestra un texto neutro. Cuando existan, aparece el grafico.
	html = html & "<section class='ui-card' aria-labelledby='tituloTiempo'><h2 class='ui-card__title' id='tituloTiempo'>Tiempo consumido</h2>"
	If ConsNum(hRel) + ConsNum(hDes) > 0 Then
		html = html & ConsDonut(Array("Tiempo funcional", "Tiempo de desarrollo"), Array(ConsNum(hRel), ConsNum(hDes)), Array("#7995FF", "#6D28D9"), "Tiempo consumido del incidente")
	Else
		html = html & "<div class='ui-empty'>Todav&iacute;a no se registra el tiempo funcional y de desarrollo de los incidentes.</div>"
	End If
	html = html & "</section>"
End If

html = html & "</div>" ' Fin ui-grid-2

' =============================================================================
' DETALLE DEL PROYECTO (descripcion y archivos)
' =============================================================================
If tipo = "P" Then
	rutaArchivos = RUTA & "\proyectos\" & id
	html = html & "<section class='ui-card' aria-labelledby='tituloDetalle'><h2 class='ui-card__title' id='tituloDetalle'>Detalle del proyecto</h2>"
	If Trim(descripcion & "") = "" Then
		html = html & "<div class='ui-empty'>Sin descripci&oacute;n cargada.</div>"
	Else
		html = html & "<div class='ui-prose'>" & Replace(Server.HTMLEncode(descripcion & ""), Chr(13), "<br>") & "</div>"
	End If
	html = html & "<div class='ui-field' style='margin-top:18px'><div class='ui-field__label'>Archivos</div><ul class='ui-files'>" & ConsListarArchivos(rutaArchivos, id, "proyectos") & "</ul></div>"
	html = html & "</section>"
End If

' =============================================================================
' HISTORIAL DE AVANCES
' =============================================================================
html = html & "<section class='ui-card' aria-labelledby='tituloHistorial'><h2 class='ui-card__title' id='tituloHistorial'>Historial de avances</h2>"
' Historial de avances: el mismo de detalles.asp (fecha, detalle, estado, tipo, responsable, usuario, via y, en el detalle de cada avance,
' de que avance nace). Solo lectura: no se activa Insistir ni el cambio de estado.
'[UI-ANT] html = html & UiFaltaBloque("tipo de avance'[UI-ANT] html = html & UiFaltaBloque("tipo de avance (anotaci&oacute;n, solicitud, insistencia) y testing: la columna TIPO de CG_AVANCES existe pero no se completa. Abajo, el historial actual.")
If tipo = "P" Then
	html = html & HtmlHistorialRelainc(Cnn, id, "proy")
Else
	html = html & HtmlHistorialRelainc(Cnn, id, "inc")
End If
html = html & "</section>"

html = html & "</div></div>" ' Fin contenedor y pagina
html = html & "</main>"
rs.close

html = html & "<script src='assets/js/jquery.min.js'></script>"
html = html & "<script src='assets/js/skel.min.js'></script>"
html = html & "<script src='assets/js/util.js'></script>"
html = html & "<script src='assets/js/main.js'></script>"
html = html & formVolver & "</body></html>"

Cnn.Close
Set Cnn = Nothing
Response.Write html


' =============================================================================
' FUNCIONES DE ESTA PANTALLA
' =============================================================================

' Pagina corta cuando el caso no se puede mostrar (id invalido o inexistente)
Sub ConsSinCaso(mensaje)
	html = html & "<main id='contenido'>" & UiTopbar("Vista de consulta", accTop) & _
		"<div class='ui-app ui-page'><div class='ui-container'><section class='ui-card'><div class='ui-empty'>" & mensaje & "</div></section></div></div></main>"
	html = html & formVolver & "</body></html>"
	Response.Write html
	Response.End
End Sub

' --- Vista de consulta de APL: pagina completa ---
Sub ConsPaginaApl(nId)
	Dim ra, metasApl
	If Not ExisteApl(Cnn, nId) Then ConsSinCaso "No se encontr&oacute; el APL " & nId & "."
	zzz = ObjAcc(nId, "A", cnn, "D")
	XXX = SYSLOGON(cnn, "CONSULTA_APL_" & nId)
	Set ra = ObtenerDetalleApl(Cnn, nId)
	If ra.eof Then ConsSinCaso "No se encontr&oacute; el APL " & nId & "."

	' Columnas del detalle: 0 id, 1 nombre, 2 cliente, 3 aplicativo, 4 responsable del envio, 5 envio,
	' 6 instalacion test, 7 arreglos, 8 instalacion produccion, 9 creacion, 12 secuencia
	metasApl = ConsMeta("Cliente", UiDato(ra(2))) & ConsMeta("Aplicativo", UiDato(ra(3))) & ConsMeta("Responsable del env&iacute;o", UiDato(ra(4))) & ConsMeta("Secuencia", UiDato(ra(12)))
	html = html & "<main id='contenido'>"
	html = html & "<div class='ui-app ui-topbar'><div class='ui-topbar__inner'>" & _
		"<div class='ui-topbar__top'>" & _
			"<div class='ui-topbar__ident'><span class='ui-topbar__id'><a href='" & urlAbrir & "' style='color:inherit' title='Abrir en la pantalla de trabajo'>APL-" & ra(0) & "</a></span><h1 class='ui-topbar__title'>" & UiTexto(ra(1)) & "</h1></div>" & _
			"<div class='ui-topbar__actions'>" & accTop & "</div>" & _
		"</div>" & _
		"<div class='ui-topbar__bottom'><div class='ui-topbar__meta'>" & metasApl & "</div>" & ConsEstadoApl(ra(5), ra(6), ra(8)) & "</div>" & _
		"</div></div>"
	html = html & "<div class='ui-app ui-page'><div class='ui-container'>"

	html = html & "<div class='ui-grid-2'>"
	html = html & "<section class='ui-card' aria-labelledby='tituloCiclo'><h2 class='ui-card__title' id='tituloCiclo'>Ciclo del APL</h2><div class='ui-lifecycle'>"
	html = html & ConsHito("Creaci&oacute;n", ra(9)) & ConsHito("Env&iacute;o al cliente", ra(5)) & ConsHito("Instalaci&oacute;n en test", ra(6)) & ConsHito("Instalaci&oacute;n en producci&oacute;n", ra(8))
	html = html & "</div></section>"
	html = html & "<section class='ui-card' aria-labelledby='tituloArreglos'><h2 class='ui-card__title' id='tituloArreglos'>Arreglos y m&oacute;dulos</h2>"
	If Trim(ra(7) & "") = "" Then
		html = html & "<div class='ui-empty'>Sin detalle de arreglos cargado.</div>"
	Else
		html = html & "<div class='ui-prose'>" & ConsTextoLargo(ra(7)) & "</div>"
	End If
	html = html & "</section></div>"

	html = html & "<section class='ui-card' aria-labelledby='tituloRel'><h2 class='ui-card__title' id='tituloRel'>Incidentes y proyectos relacionados</h2>"
	html = html & "<div class='ui-grid-2 ui-grid-2--plano'>"
	html = html & "<div><h3 class='ui-apl-rel__titulo'>Incidentes</h3>" & ConsRelacionadosApl(nId, "I") & "</div>"
	html = html & "<div><h3 class='ui-apl-rel__titulo'>Proyectos</h3>" & ConsRelacionadosApl(nId, "P") & "</div>"
	html = html & "</div></section>"
	html = html & "</div></div></main>"
	ra.Close

	html = html & "<script src='assets/js/jquery.min.js'></script>"
	html = html & "<script src='assets/js/skel.min.js'></script>"
	html = html & "<script src='assets/js/util.js'></script>"
	html = html & "<script src='assets/js/main.js'></script>"
	html = html & formVolver & "</body></html>"
	Response.Write html
	Response.End
End Sub

' Estado del APL segun su avance (en el lugar del semaforo de incidentes y proyectos)
Function ConsEstadoApl(fEnvio, fTest, fProd)
	Dim texto, color, detalle
	If ConsNum(fProd) > 0 Then
		texto = "En producci&oacute;n"
		color = "green"
		detalle = "Instalado el " & UiFecha(fProd)
	ElseIf ConsNum(fTest) > 0 Then
		texto = "En test"
		color = "purple"
		detalle = "Instalado en test el " & UiFecha(fTest)
	ElseIf ConsNum(fEnvio) > 0 Then
		texto = "Enviado"
		color = "blue"
		detalle = "Enviado el " & UiFecha(fEnvio)
	Else
		texto = "Sin enviar"
		color = "amber"
		detalle = "Todav&iacute;a no se envi&oacute; al cliente"
	End If
	ConsEstadoApl = "<div class='ui-apl-estado'><span class='ui-topbar__label'>Estado</span>" & UiBadge(texto, color) & "<span class='ui-apl-estado__detalle'>" & detalle & "</span></div>"
End Function

' Texto largo de la base: escapado, con sus saltos de linea (respeta los <br> que traiga)
Function ConsTextoLargo(v)
	Dim t
	t = Server.HTMLEncode(Trim(v & ""))
	t = Replace(Replace(Replace(t, "&lt;br&gt;", "<br>", 1, -1, 1), "&lt;br/&gt;", "<br>", 1, -1, 1), "&lt;br /&gt;", "<br>", 1, -1, 1)
	ConsTextoLargo = Replace(Replace(Replace(t, vbCrLf, "<br>"), vbLf, "<br>"), vbCr, "<br>")
End Function

' Incidentes (I) o proyectos (P) relacionados con el APL: abren su vista de consulta
Function ConsRelacionadosApl(nId, t)
	Dim r, h, n, url, etiqueta
	Set r = ObtenerRelacionesApl(Cnn, nId, t)
	h = ""
	n = 0
	Do While Not r.eof
		If t = "I" Then
			etiqueta = "<span class='ui-tag-id'>INC-" & r(2) & "</span>"
		Else
			etiqueta = "<span class='ui-tag-id ui-tag-id--pry'>PRY-" & r(2) & "</span>"
		End If
		url = "vistaconsulta.asp?tipo=" & t & "&id=" & r(2) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu
		h = h & "<li><a class='ui-relac__link' href='" & url & "'>" & etiqueta & " " & UiTexto(r(6)) & "</a></li>"
		n = n + 1
		r.movenext
	Loop
	If n = 0 Then
		If t = "I" Then ConsRelacionadosApl = "<p class='ui-status-empty'>Sin incidentes relacionados.</p>" Else ConsRelacionadosApl = "<p class='ui-status-empty'>Sin proyectos relacionados.</p>"
	Else
		ConsRelacionadosApl = "<ul class='ui-apl-rel'>" & h & "</ul>"
	End If
End Function

' Campo oculto para Volver: repite el filtro que llego en la URL
Function ConsOculto(nombre)
	ConsOculto = "<input type='hidden' name='" & nombre & "' value='" & Server.HTMLEncode(Request.QueryString(nombre) & "") & "'>"
End Function

Function ConsEtiqueta(t, n)
	If t = "P" Then
		ConsEtiqueta = "PRY-" & n
	Else
		ConsEtiqueta = "INC-" & n
	End If
End Function

Function ConsNombre(t)
	If t = "P" Then
		ConsNombre = "proyecto"
	Else
		ConsNombre = "incidente"
	End If
End Function

Function ConsTextoPlazo(t)
	If t = "P" Then
		ConsTextoPlazo = "fecha de entrega"
	Else
		ConsTextoPlazo = "plazo de la prioridad"
	End If
End Function

' Dato de la barra superior (etiqueta arriba, valor abajo)
Function ConsMeta(etiqueta, valorHtml)
	ConsMeta = "<div><span class='ui-topbar__label'>" & etiqueta & "</span><span class='ui-topbar__value'>" & valorHtml & "</span></div>"
End Function

' Fecha del sistema (dias desde 01-01-1900) como Date; vacia o en cero: Empty
Function ConsFechaDb(v)
	ConsFechaDb = Empty
	If IsNumeric(v & "") Then
		If CLng(v) > 0 Then ConsFechaDb = CDate(CLng(v) + CLng(CDate("01-01-1900")))
	End If
End Function

' Plazos fijos de hoy por prioridad (hasta que exista la tabla SLA)
Function ConsPlazoDias(v)
	Select Case UCase(Trim(v & ""))
		Case "ALTA"
			ConsPlazoDias = 5
		Case "MEDIA"
			ConsPlazoDias = 10
		Case "BAJA"
			ConsPlazoDias = 15
		Case Else
			ConsPlazoDias = 0
	End Select
End Function

Function ConsNum(v)
	If IsNumeric(v & "") Then
		ConsNum = CDbl(v)
	Else
		ConsNum = 0
	End If
End Function

' Numero para CSS: siempre con punto decimal (la configuracion regional usa coma)
Function ConsCss(n)
	ConsCss = Replace(CStr(Round(n, 2)), ",", ".")
End Function

' Semaforo con la regla definida: ambar al consumir el 50% del plazo, rojo al 80%
Function ConsSemaforo(fI, fL, okCliente, faltaQue)
	Dim plazo, rest, dias, cls, num, etiqueta, pct
	If ConsNum(okCliente) > 0 Then
		ConsSemaforo = "<div class='ui-semaforo ui-semaforo--done'><div class='ui-semaforo__cifra'><div class='ui-semaforo__num'>&#10003;</div><div class='ui-semaforo__label'>Resuelto</div></div><div class='ui-semaforo__graf'><div class='ui-semaforo__scale'><span>OK cliente: " & UiFecha(okCliente) & "</span></div></div></div>"
		Exit Function
	End If
	If IsEmpty(fI) Or IsEmpty(fL) Then
		ConsSemaforo = "<div class='ui-semaforo ui-semaforo--gray'><div class='ui-semaforo__cifra'><div class='ui-semaforo__num'>?</div><div class='ui-semaforo__label'>Sin plazo</div></div><div class='ui-semaforo__graf'>" & UiFalta(faltaQue) & "</div></div>"
		Exit Function
	End If
	plazo = DateDiff("d", fI, fL)
	If plazo < 1 Then plazo = 1
	rest = DateDiff("d", Date(), fL)
	dias = Abs(rest)
	pct = Round(((plazo - rest) / plazo) * 100)
	If pct >= porcRojo Then
		cls = "red"
	ElseIf pct >= porcAmbar Then
		cls = "amber"
	Else
		cls = "green"
	End If
	If pct < 0 Then pct = 0
	If pct > 100 Then pct = 100
	If rest = 0 Then
		num = "&iexcl;Hoy!"
		etiqueta = "vence"
	ElseIf rest < 0 Then
		num = dias
		If dias = 1 Then etiqueta = "d&iacute;a vencido" Else etiqueta = "d&iacute;as vencido"
	Else
		num = dias
		If dias = 1 Then etiqueta = "d&iacute;a restante" Else etiqueta = "d&iacute;as restantes"
	End If
	ConsSemaforo = "<div class='ui-semaforo ui-semaforo--" & cls & "'><div class='ui-semaforo__cifra'><div class='ui-semaforo__num'>" & num & "</div><div class='ui-semaforo__label'>" & etiqueta & "</div></div>" & _
		"<div class='ui-semaforo__graf'><div class='ui-semaforo__bar'><span style='width:" & pct & "%'></span></div>" & _
		"<div class='ui-semaforo__scale'><span>Inicio</span><span>" & pct & "% transcurrido</span><span>L&iacute;mite</span></div></div></div>"
End Function

' Ultimo avance real del caso (CG_AVANCES). sp_cg_ObtenerComentarios devuelve el historial del mas nuevo
' al mas viejo (descripcion, fecha, hora, usuario, responsable) e incluye el texto de avances anterior
' con la fecha de ingreso: se toma la primera fila con rol FUNCIONAL, TECNICO o CLIENTE, que son los avances cargados.
Function ConsUltimoAvance(nId, t)
	Dim rh, tipoObj, valor, hecho, rol, extra, tav, est
	If t = "P" Then tipoObj = "proy" Else tipoObj = "inc"
	Set rh = Exec(Cnn, "sp_cg_ObtenerComentarios @codobjeto=" & CLng(nId) & ", @tipoobjeto='" & tipoObj & "'")
	hecho = False
	valor = "Sin avances cargados"
	Do While Not rh.eof
		' Solo avances reales (los de CG_AVANCES tienen ID; las filas iniciales del caso no). El historial viene del mas nuevo al mas viejo.
		If (Not hecho) And HistCol(rh, "ID") <> "" Then
			rol = UCase(Trim(HistCol(rh, "responsable")))
			Select Case rol
				Case "FUNCIONAL"
					rol = "Funcional"
				Case "TECNICO"
					rol = "T&eacute;cnico"
				Case Else
					rol = "Cliente"
			End Select
			tav = Trim(HistCol(rh, "TAV_DESCRIPCION"))
			est = Trim(HistCol(rh, "AV_ESTADO"))
			extra = ""
			If tav <> "" Then extra = extra & " &middot; " & UiTexto(tav)
			If est <> "" Then extra = extra & " &middot; " & UiTexto(est)
			valor = UiFecha(rh("fecha")) & UiSub(rol & " &middot; " & UiTexto(HistCol(rh, "usuario")) & extra)
			hecho = True
		End If
		rh.movenext
	Loop
	rh.Close
	If hecho Then
		ConsUltimoAvance = "<div class='ui-lifecycle__item ui-lifecycle__item--done'><span class='ui-lifecycle__dot'></span><div><div class='ui-lifecycle__label'>&Uacute;ltimo avance</div><div class='ui-lifecycle__value'>" & valor & "</div></div></div>"
	Else
		ConsUltimoAvance = "<div class='ui-lifecycle__item ui-lifecycle__item--pend'><span class='ui-lifecycle__dot'></span><div><div class='ui-lifecycle__label'>&Uacute;ltimo avance</div><div class='ui-lifecycle__value'>" & valor & "</div></div></div>"
	End If
End Function

' Proyecto relacionado de un incidente: enlace a su vista de consulta (o un guion si no tiene)
Function ConsProyectoRel(idP, solic, titP)
	Dim texto
	ConsProyectoRel = UiDato("")
	If IsNumeric(idP & "") And Trim(idP & "") <> "" Then
		If CLng(idP) > 0 Then
			texto = Trim(solic & "")
			If texto <> "" And Trim(titP & "") <> "" Then texto = texto & " - "
			texto = texto & Trim(titP & "")
			ConsProyectoRel = "<a class='ui-relac__link' href='vistaconsulta.asp?tipo=P&id=" & CLng(idP) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'><span class='ui-tag-id ui-tag-id--pry'>PRY-" & CLng(idP) & "</span> " & UiTexto(texto) & "</a>"
		End If
	End If
End Function

' Hito del ciclo de vida: cumplido si tiene fecha
Function ConsHito(etiqueta, valorFecha)
	If ConsNum(valorFecha) > 0 Then
		ConsHito = "<div class='ui-lifecycle__item ui-lifecycle__item--done'><span class='ui-lifecycle__dot'></span><div><div class='ui-lifecycle__label'>" & etiqueta & "</div><div class='ui-lifecycle__value'>" & UiFecha(valorFecha) & "</div></div></div>"
	Else
		ConsHito = "<div class='ui-lifecycle__item ui-lifecycle__item--pend'><span class='ui-lifecycle__dot'></span><div><div class='ui-lifecycle__label'>" & etiqueta & "</div><div class='ui-lifecycle__value'>Pendiente</div></div></div>"
	End If
End Function

' Hito con fecha objetivo: siempre muestra la fecha; cumplido solo si se alcanzo
Function ConsHitoMeta(etiqueta, valorFecha, cumplido)
	Dim estado, valor
	If cumplido Then estado = "done" Else estado = "pend"
	If ConsNum(valorFecha) > 0 Then
		valor = UiFecha(valorFecha)
	Else
		valor = UiFalta("fecha de entrega")
	End If
	ConsHitoMeta = "<div class='ui-lifecycle__item ui-lifecycle__item--" & estado & "'><span class='ui-lifecycle__dot'></span><div><div class='ui-lifecycle__label'>" & etiqueta & "</div><div class='ui-lifecycle__value'>" & valor & "</div></div></div>"
End Function

' Grafico circular con leyenda (horas). nombres, valores y colores: arreglos del mismo largo
Function ConsDonut(nombres, valores, colores, descripcion)
	Dim total, i, acum, desde, tramos, leyenda, textoAccesible, pctItem
	total = 0
	For i = 0 To UBound(valores)
		total = total + valores(i)
	Next
	If total <= 0 Then
		ConsDonut = "<div class='ui-empty'>Todav&iacute;a no hay horas cargadas.</div>"
		Exit Function
	End If
	acum = 0
	tramos = ""
	leyenda = ""
	textoAccesible = descripcion & ": " & total & " h en total"
	For i = 0 To UBound(valores)
		If valores(i) > 0 Then
			desde = acum
			acum = acum + valores(i) / total * 100
			If tramos <> "" Then tramos = tramos & ", "
			tramos = tramos & colores(i) & " " & ConsCss(desde) & "% " & ConsCss(acum) & "%"
		End If
		pctItem = Round(valores(i) / total * 100)
		leyenda = leyenda & "<div class='ui-donut__item'><span class='ui-dot' style='background:" & colores(i) & "' aria-hidden='true'></span>" & _
			"<span class='ui-donut__name'>" & nombres(i) & "</span><span class='ui-donut__value'>" & valores(i) & " h</span><span class='ui-donut__pct'>" & pctItem & "%</span></div>"
		textoAccesible = textoAccesible & "; " & nombres(i) & " " & valores(i) & " h"
	Next
	ConsDonut = "<div class='ui-donut'><div class='ui-donut__chart' role='img' aria-label='" & textoAccesible & "' style='background:conic-gradient(" & tramos & ")'>" & _
		"<div class='ui-donut__center'><div class='ui-donut__total'>" & total & " h</div><div class='ui-donut__label'>Total</div></div></div>" & _
		"<div class='ui-donut__legend'>" & leyenda & "</div></div>"
End Function

' Archivos del caso. Solo lectura: a diferencia de LISTARDIRTORIO (detalles.asp), no crea la carpeta si no existe
Function ConsListarArchivos(dir, nId, d)
	Dim fso, carpeta, archivo, lista
	lista = ""
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	If fso.FolderExists(dir) Then
		Set carpeta = fso.GetFolder(dir)
		For Each archivo In carpeta.Files
			lista = lista & "<li class='ui-files__item'><a class='ui-link' target='_blank' href='" & shttp & d & "/" & nId & "/" & Server.HTMLEncode(archivo.name) & "'>" & Server.HTMLEncode(archivo.name) & "</a></li>"
		Next
	End If
	If lista = "" Then lista = "<li class='ui-files__item ui-files__item--empty'>Sin archivos adjuntos</li>"
	ConsListarArchivos = lista
End Function
%>
