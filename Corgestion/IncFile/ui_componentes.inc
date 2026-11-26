<%
' =========================================================================
' ui_componentes.inc - Componentes de presentacion del sistema visual
' (corgestion-ui.css). Solo arman HTML como string: no consultan ni
' modifican datos.
'
' Uso: incluir despues de funciones.inc
'   <!-- #INCLUDE FILE="IncFile/ui_componentes.inc" -->
'
' Pantalla de listado tipica:
'   sPage = sPage & UiPantallaAbre("Titulo", botonesHtml, "wide")
'   sPage = sPage & UiCardAbre("Listado de ...", "")
'   sPage = sPage & UiTablaAbre("Descripcion para lectores de pantalla")
'   sPage = sPage & "<thead><tr>" & UiTh(...) & UiThOrden(...) & "</tr></thead><tbody>"
'   ... filas ...
'   sPage = sPage & UiTablaCierra()
'   sPage = sPage & UiPaginacion(desde, hasta, botonAnterior, botonSiguiente)
'   sPage = sPage & UiCardCierra() & UiPantallaCierra()
'
' Textos con tildes: usar entidades HTML (&aacute;, &iacute;...) para que
' se vean bien con cualquier codepage.
' =========================================================================

' ---------- Texto ----------

' Dato de la base listo para mostrar (tolera Null)
Function UiTexto(v)
	UiTexto = Server.HTMLEncode(v & "")
End Function

' Dato de la base para una celda: el texto, o un guion atenuado si viene vacio
Function UiDato(v)
	If Trim(v & "") = "" Then
		UiDato = "<span class='ui-status-empty'>&mdash;</span>"
	Else
		UiDato = Server.HTMLEncode(v & "")
	End If
End Function

' "PENDIENTE" -> "Pendiente". Las siglas cortas (OK, SI, NO) quedan igual
Function UiCapitalizar(v)
	Dim t
	t = Trim(v & "")
	If Len(t) > 2 Then
		UiCapitalizar = UCase(Left(t, 1)) & LCase(Mid(t, 2))
	Else
		UiCapitalizar = t
	End If
End Function

' ---------- Estructura de la pagina ----------

' Barra superior: titulo de la pantalla y botones ya armados (normalmente se usa via UiPantallaAbre)
Function UiTopbar(titulo, accionesHtml)
	UiTopbar = "<div class='ui-app ui-topbar'><div class='ui-topbar__inner'><div class='ui-topbar__top'>" & _
		"<div class='ui-topbar__ident'><h1 class='ui-topbar__title'>" & titulo & "</h1></div>" & _
		"<div class='ui-topbar__actions'>" & accionesHtml & "</div>" & _
		"</div></div></div>"
End Function

' Apertura de la pantalla: region principal + barra superior + contenedor, en el orden correcto.
' ancho: "" (1040px), "wide" (1240px, listados) o "narrow" (720px, formularios)
Function UiPantallaAbre(titulo, accionesHtml, ancho)
	Dim clase
	clase = "ui-container"
	If ancho & "" <> "" Then clase = clase & " ui-container--" & ancho
	' La barra superior toma el mismo ancho que el contenido, para que el titulo quede alineado
	Dim barra
	barra = UiTopbar(titulo, accionesHtml)
	If ancho & "" <> "" Then barra = Replace(barra, "class='ui-topbar__inner'", "class='ui-topbar__inner ui-topbar__inner--" & ancho & "'")
	UiPantallaAbre = "<main id='contenido'>" & barra & _
		"<div class='ui-app ui-page'><div class='" & clase & "'>"
End Function

' Pantalla con una etiqueta sobre el titulo, que indica a que area pertenece (Parametros, Seguridad)
Function UiPantallaEtiquetaAbre(etiqueta, titulo, accionesHtml, ancho)
	UiPantallaEtiquetaAbre = Replace(UiPantallaAbre(titulo, accionesHtml, ancho), "<div class='ui-topbar__ident'>", _
		"<div class='ui-topbar__ident'><span class='ui-topbar__id'>" & etiqueta & "</span>", 1, 1)
End Function

' Pantalla de parametrizacion: etiqueta "Parametros" (todas las pantallas de configuracion del sistema)
Function UiPantallaParametrosAbre(titulo, accionesHtml, ancho)
	UiPantallaParametrosAbre = UiPantallaEtiquetaAbre("Par&aacute;metros", titulo, accionesHtml, ancho)
End Function

' Pantalla de seguridad: etiqueta "Seguridad" (usuarios y permisos)
Function UiPantallaSeguridadAbre(titulo, accionesHtml, ancho)
	UiPantallaSeguridadAbre = UiPantallaEtiquetaAbre("Seguridad", titulo, accionesHtml, ancho)
End Function

Function UiPantallaCierra()
	UiPantallaCierra = "</div></div></main>"
End Function

' Tarjeta con titulo; accionesHtml opcional va a la derecha del encabezado
Function UiCardAbre(titulo, accionesHtml)
	If accionesHtml & "" = "" Then
		UiCardAbre = "<section class='ui-card'><h2 class='ui-card__title'>" & titulo & "</h2>"
	Else
		UiCardAbre = "<section class='ui-card'><div class='ui-card__header'><h2 class='ui-card__title'>" & titulo & "</h2>" & accionesHtml & "</div>"
	End If
End Function

Function UiCardCierra()
	UiCardCierra = "</section>"
End Function

' Pie de tarjeta: texto informativo a la izquierda y acciones a la derecha
Function UiPie(info, accionesHtml)
	UiPie = "<div class='ui-card__footer ui-card__footer--split'><span class='ui-card__footer-info'>" & info & "</span>" & accionesHtml & "</div>"
End Function

' ---------- Botones ----------
' variante: "primary", "secondary", "light", "danger", "quiet"
' Sin permiso: queda visible, anunciado como deshabilitado y con el motivo (title)

Function UiBoton(texto, variante, onclick, habilitado, motivo)
	If habilitado Then
		UiBoton = "<button type='button' class='ui-button ui-button--" & variante & " ui-button--sm' onclick=""" & onclick & """>" & texto & "</button>"
	Else
		UiBoton = "<button type='button' class='ui-button ui-button--" & variante & " ui-button--sm ui-button--disabled' aria-disabled='true' title='" & motivo & "'>" & texto & "</button>"
	End If
End Function

' Boton redondo con simbolo. etiqueta: lo que hace (lo leen los lectores de pantalla)
Function UiBotonIcono(simbolo, etiqueta, variante, onclick, habilitado, motivo)
	If habilitado Then
		UiBotonIcono = "<button type='button' class='ui-button ui-button--icon ui-button--" & variante & "' aria-label='" & etiqueta & "' title='" & etiqueta & "' onclick=""" & onclick & """>" & simbolo & "</button>"
	Else
		UiBotonIcono = "<button type='button' class='ui-button ui-button--icon ui-button--disabled' aria-disabled='true' aria-label='" & etiqueta & ": " & motivo & "' title='" & motivo & "'>" & simbolo & "</button>"
	End If
End Function

' ---------- Filtros ----------

' Control de filtro con etiqueta visible. idControl = id del select/input
Function UiFiltro(idControl, etiqueta, controlHtml)
	UiFiltro = "<div class='ui-filtro-campo'><label class='ui-field__label' for='" & idControl & "'>" & etiqueta & "</label>" & controlHtml & "</div>"
End Function

' Filtro compacto para el encabezado de una tarjeta: sin etiqueta visible (la lee el lector de pantalla).
' La primera opcion del select tiene que explicar el filtro: "Todos los clientes".
Function UiFiltroCompacto(idControl, etiqueta, controlHtml)
	UiFiltroCompacto = "<label class='ui-sr-only' for='" & idControl & "'>" & etiqueta & "</label>" & controlHtml
End Function

' ---------- Fechas ----------

' Fecha como dd/mm/aaaa, sin depender de la configuracion regional. Acepta:
'   - el formato numerico del sistema (dias desde 01-01-1900), ej. 45931
'   - un campo de fecha (datetime) que devuelva el procedimiento
' Vacia, nula o en cero: guion.
Function UiFecha(v)
	Dim d
	UiFecha = "<span class='ui-status-empty'>&mdash;</span>"
	If IsNull(v) Or IsEmpty(v) Then Exit Function
	If VarType(v) = 7 Then
		d = v
	ElseIf IsNumeric(v & "") Then
		If CLng(v) <= 0 Then Exit Function
		d = CDate(CLng(v) + CLng(CDate("01-01-1900")))
	ElseIf IsDate(v & "") Then
		d = CDate(v & "")
	Else
		Exit Function
	End If
	UiFecha = Right("0" & Day(d), 2) & "/" & Right("0" & Month(d), 2) & "/" & Year(d)
End Function

' ---------- Tablas ----------

' descripcion: titulo de la tabla para lectores de pantalla
Function UiTablaAbre(descripcion)
	UiTablaAbre = "<div class='ui-table-wrap'><table class='ui-table ui-table--middle'><caption class='ui-sr-only'>" & descripcion & "</caption>"
End Function

Function UiTablaCierra()
	UiTablaCierra = "</tbody></table></div>"
End Function

' Encabezado de columna. clase opcional: "ui-table__col--actions", "ui-table__cell--center"...
Function UiTh(texto, clase)
	If clase & "" = "" Then
		UiTh = "<th scope='col'>" & texto & "</th>"
	Else
		UiTh = "<th scope='col' class='" & clase & "'>" & texto & "</th>"
	End If
End Function

' Encabezado ordenable. activo = True marca la columna por la que esta ordenado el listado
Function UiThOrden(texto, url, activo)
	Dim orden
	orden = ""
	If activo Then orden = " aria-sort='ascending'"
	UiThOrden = "<th scope='col'" & orden & "><a class='ui-table__sort' href='" & url & "'>" & texto & "</a></th>"
End Function

' Fila unica para listados sin resultados
Function UiFilaVacia(columnas, texto)
	UiFilaVacia = "<tr><td colspan='" & columnas & "'><div class='ui-empty'>" & texto & "</div></td></tr>"
End Function

' Dato principal de la fila como enlace real (se puede abrir con teclado)
Function UiCeldaEnlace(texto, url, tituloCompleto)
	UiCeldaEnlace = "<a class='ui-table__titulo' href='" & url & "' title='" & tituloCompleto & "'>" & texto & "</a>"
End Function

' Texto secundario debajo del dato de una celda (ej. responsable debajo del estado)
Function UiSub(texto)
	If texto & "" = "" Then
		UiSub = ""
	Else
		UiSub = "<span class='ui-table__sub'>" & texto & "</span>"
	End If
End Function

' Pie con paginacion: "Mostrando X a Y" + botones Anterior / Siguiente ya armados
Function UiPaginacion(desde, hasta, anteriorHtml, siguienteHtml)
	Dim info
	If hasta < desde Then
		info = "Sin resultados"
	Else
		info = "Mostrando " & desde & " a " & hasta
	End If
	UiPaginacion = UiPie(info, "<nav class='ui-pagination' aria-label='Paginaci&oacute;n'>" & anteriorHtml & siguienteHtml & "</nav>")
End Function

' Texto del rango mostrado: "Mostrando 1 a 12 de 57", "1 resultado" o "Sin resultados"
Function UiInfoRango(desde, hasta, total)
	If total = 0 Or hasta < desde Then
		UiInfoRango = "Sin resultados"
	ElseIf total = 1 Then
		UiInfoRango = "1 resultado"
	Else
		UiInfoRango = "Mostrando " & desde & " a " & hasta & " de " & total
	End If
End Function

' Pie con paginacion que informa el total de registros
Function UiPaginacionDe(desde, hasta, total, anteriorHtml, siguienteHtml)
	UiPaginacionDe = UiPie(UiInfoRango(desde, hasta, total), "<nav class='ui-pagination' aria-label='Paginaci&oacute;n'>" & anteriorHtml & siguienteHtml & "</nav>")
End Function

' ---------- Formularios de alta ----------
' Uso tipico:
'   html = html & "<form id='agre' ...>" & UiCardAbre("Datos del ...", "") & UiNotaObligatorios() & UiFormGridAbre()
'   html = html & UiCampo("tit", "T&iacute;tulo", UiInputTexto("tit", True, ""), True, True)
'   html = html & UiCampo("cli", "Cliente", UiSelectRs("cli", rsClientes, "", "", "", ""), False, False)
'   html = html & UiFormGridCierra() & UiCardCierra() & UiFormAcciones(UiBotonPrincipal("Agregar", "agregar();")) & "</form>"

' Grilla de dos columnas para formularios (una sola en pantallas angostas)
Function UiFormGridAbre()
	UiFormGridAbre = "<div class='ui-form-grid'>"
End Function

Function UiFormGridCierra()
	UiFormGridCierra = "</div>"
End Function

' Campo con su etiqueta asociada al control (for = id).
' requerido: muestra el asterisco. ancho: True ocupa toda la fila.
Function UiCampo(idControl, etiqueta, controlHtml, requerido, ancho)
	Dim clase, marca
	clase = ""
	If ancho Then clase = " class='ui-form-grid__ancho'"
	marca = ""
	If requerido Then marca = "<span class='ui-field__required' aria-hidden='true'> *</span>"
	UiCampo = "<div" & clase & "><label class='ui-field__label' for='" & idControl & "'>" & etiqueta & marca & "</label>" & controlHtml & "</div>"
End Function

' Dato de solo lectura con el mismo formato que un campo
Function UiCampoLectura(etiqueta, valorHtml, ancho)
	Dim clase
	clase = ""
	If ancho Then clase = " class='ui-form-grid__ancho'"
	UiCampoLectura = "<div" & clase & "><span class='ui-field__label'>" & etiqueta & "</span><div class='ui-field__value'>" & valorHtml & "</div></div>"
End Function

' Campo de texto. requerido lo anuncia a los lectores de pantalla. extra: atributos adicionales
Function UiInputTexto(id, requerido, extra)
	Dim req
	req = ""
	If requerido Then req = " aria-required='true'"
	UiInputTexto = "<input class='ui-input' type='text' id='" & id & "' name='" & id & "'" & req & " " & extra & ">"
End Function

Function UiTextarea(id, filas, requerido)
	Dim req
	req = ""
	If requerido Then req = " aria-required='true'"
	UiTextarea = "<textarea class='ui-input' id='" & id & "' name='" & id & "' rows='" & filas & "'" & req & "></textarea>"
End Function

' Select desde un recordset (rs(0) = valor, rs(1) = texto). Deja el recordset recorrido.
' valorSel: valor o texto de la opcion que va seleccionada ("" = ninguna).
' textoVacio / valorVacio: primera opcion opcional (ej. "Seleccionar", "0"); textoVacio "" = sin esa opcion.
Function UiSelectRs(id, rs, valorSel, textoVacio, valorVacio, extra)
	Dim h, marca
	h = "<select class='ui-select' id='" & id & "' name='" & id & "' " & extra & ">"
	If textoVacio & "" <> "" Then h = h & "<option value='" & valorVacio & "'>" & textoVacio & "</option>"
	Do While Not rs.EOF
		marca = ""
		If valorSel & "" <> "" Then
			If CStr(rs(0) & "") = CStr(valorSel & "") Or UCase(rs(1) & "") = UCase(valorSel & "") Then marca = " selected"
		End If
		h = h & "<option value='" & rs(0) & "'" & marca & ">" & Server.HTMLEncode(rs(1) & "") & "</option>"
		rs.MoveNext
	Loop
	UiSelectRs = h & "</select>"
End Function

' Select de opciones fijas: "valor|Texto||valor|Texto" (|| separa opciones: el ; no sirve porque
' lo usan las entidades como &iacute;). valorSel marca la elegida
Function UiSelectOpciones(id, opciones, valorSel, extra)
	Dim h, partes, i, par, marca
	h = "<select class='ui-select' id='" & id & "' name='" & id & "' " & extra & ">"
	partes = Split(opciones, "||")
	For i = 0 To UBound(partes)
		par = Split(partes(i), "|")
		marca = ""
		If CStr(par(0)) = CStr(valorSel & "") Then marca = " selected"
		h = h & "<option value='" & par(0) & "'" & marca & ">" & par(1) & "</option>"
	Next
	UiSelectOpciones = h & "</select>"
End Function

' Grupo de campos con titulo (ej. "Horas estimadas"). Abre la grilla interna de 4 columnas
Function UiGrupoAbre(titulo)
	UiGrupoAbre = "<fieldset class='ui-fieldset ui-form-grid__ancho'><legend class='ui-fieldset__legend'>" & titulo & "</legend><div class='ui-form-grid ui-form-grid--4'>"
End Function

Function UiGrupoCierra()
	UiGrupoCierra = "</div></fieldset>"
End Function

Function UiNotaObligatorios()
	UiNotaObligatorios = "<p class='ui-form-nota'>Los campos marcados con <span class='ui-field__required'>*</span> son obligatorios.</p>"
End Function

' Boton principal grande del formulario
Function UiBotonPrincipal(texto, onclick)
	UiBotonPrincipal = "<button type='button' class='ui-button ui-button--primary ui-button--lg' onclick=""" & onclick & """>" & texto & "</button>"
End Function

' Enlace con forma de boton (para encabezados de tarjeta y barras). Sin permiso: boton deshabilitado con el motivo
Function UiBotonEnlaceSi(texto, url, habilitado, motivo)
	If habilitado Then
		UiBotonEnlaceSi = "<a class='ui-button ui-button--light ui-button--sm' href='" & url & "'>" & texto & "</a>"
	Else
		UiBotonEnlaceSi = UiBoton(texto, "light", "", False, motivo)
	End If
End Function

' Boton principal que envia el formulario (type=submit: tambien envia con Enter).
' Sin permiso: queda visible y deshabilitado, con el motivo.
Function UiBotonEnviar(texto, habilitado, motivo)
	If habilitado Then
		UiBotonEnviar = "<button type='submit' class='ui-button ui-button--primary ui-button--lg'>" & texto & "</button>"
	Else
		UiBotonEnviar = "<button type='button' class='ui-button ui-button--primary ui-button--lg ui-button--disabled' aria-disabled='true' title='" & motivo & "'>" & texto & "</button>"
	End If
End Function

' Barra al pie del formulario con el boton principal centrado
Function UiFormAcciones(botonesHtml)
	UiFormAcciones = "<div class='ui-savebar'>" & botonesHtml & "</div>"
End Function

' ---------- Datos pendientes de conectar ----------
' Marca en pantalla un dato que todavia no existe en el sistema (borde ambar punteado).
' Se reemplaza por el dato real cuando este disponible.

Function UiFalta(texto)
	UiFalta = "<span class='ui-falta'>Falta dato: " & texto & "</span>"
End Function

Function UiFaltaBloque(texto)
	UiFaltaBloque = "<div class='ui-falta-bloque'>Falta dato: " & texto & "</div>"
End Function

' ---------- Badges ----------

Function UiBadge(texto, color)
	UiBadge = "<span class='ui-badge ui-badge--" & color & "'>" & texto & "</span>"
End Function

' Estado de catalogo (desarrollo, funcional...). Mismos colores que detalles.asp
Function UiBadgeEstado(valor)
	Dim t, c
	t = UCase(Trim(valor & ""))
	If t = "" Or t = "---" Then
		UiBadgeEstado = "<span class='ui-status-empty'>&mdash;</span>"
		Exit Function
	End If
	Select Case t
		Case "OK", "FINALIZADO", "CERRADO", "APLICADO", "CONFORME", "RESUELTO"
			c = "green"
		Case "BAJA", "RECHAZADO", "CANCELADO"
			c = "red"
		Case "EN CURSO", "EN DESARROLLO", "EN PROCESO"
			c = "amber"
		Case "TESTING", "EN TESTING", "EN TESTEO", "TESTING EN CLIENTE"
			c = "purple"
		Case "EN ANALISIS", "EN RELEVAMIENTO", "EN DEFINICION"
			c = "blue"
		Case Else
			c = "gray"
	End Select
	UiBadgeEstado = UiBadge(UiTexto(UiCapitalizar(valor)), c)
End Function

' Prioridad: Alta rojo, Media ambar, Baja verde
Function UiBadgePrioridad(valor)
	Dim t, c
	t = UCase(Trim(valor & ""))
	If t = "" Then
		UiBadgePrioridad = "<span class='ui-status-empty'>&mdash;</span>"
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
	UiBadgePrioridad = UiBadge(UiTexto(UiCapitalizar(valor)), c)
End Function
%>