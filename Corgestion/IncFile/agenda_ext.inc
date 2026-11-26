<%
' =========================================================================
' agenda_ext.inc - Reuniones de la agenda (consulta, alta, edicion, postergar, seguimiento,
' participantes y temario). Lo incluye agenda.asp y resuelve los modos de AgendaExtDebeManejar.
'
' [UI] Normalizado con el sistema visual (IncFile/ui_componentes.inc, que incluye agenda.asp):
'   - Cabecera comun y componentes (antes un documento propio con CSS inline y marca "CG").
'   - Mensajes en la pagina despues de guardar (antes alert); los participantes no agregados
'     por cruce de horario se listan en un aviso.
'   - Participantes del sistema por LOGIN: la base guarda ID_PERSONA_SEG (texto) y corresp_listar
'     lo devuelve como ID_PERSONA. Antes el guardado exigia un numero: los participantes del sistema
'     no se guardaban, no se validaban sus cruces y se perdian al editar la reunion.
'   - Responsable por login (@IdRespon es texto) y usuario con respaldo en UserLogon().
'   - JavaScript que arma los campos con el DOM (antes atributos sin comillas: un nombre manual
'     con espacios se cortaba).
' Se conservan los nombres de funciones, campos, procedimientos y destinos.
' El archivo anterior completo queda al final, comentado con '[UI-ANT].
' =========================================================================

Function AgendaExtDebeManejar(opc)
	AgendaExtDebeManejar = (opc = "consulta" Or opc = "cargar" Or opc = "det" Or opc = "agenda_ext_guardar" Or opc = "agenda_ext_invitar" Or opc = "agenda_ext_postergar" Or opc = "agenda_ext_seguimiento" Or opc = "agenda_ext_eliminar_corresp")
End Function

Sub AgendaExtRender(cn, opc, xmenu, xtmenu)
	Dim id, rs, avisosCruce, accionLog
	If opc = "consulta" Then
		AgendaExtConsulta cn, xmenu, xtmenu
		Exit Sub
	End If
	id = Request.QueryString("id")
	If opc = "agenda_ext_guardar" Then
		avisosCruce = ""
		accionLog = "ALTA"
		If AgendaExtEsIdValido(id) Then accionLog = "ACTUALIZA"
		id = AgendaExtGuardar(cn, id, avisosCruce)
		CgLog "AGENDA", id, accionLog
		Session("agendaExtCruces") = avisosCruce
		AgendaExtRedir "guardada", id, xmenu, xtmenu
		Exit Sub
	End If
	' [INVITACION] Envia el mail con el .ics (agenda_invitacion.inc) y vuelve a la reunion con el resultado
	If opc = "agenda_ext_invitar" Then
		AgendaExtInvitar cn, id
		AgendaExtRedir "invitacion", id, xmenu, xtmenu
		Exit Sub
	End If
	If opc = "agenda_ext_postergar" Then
		id = AgendaExtPostergar(cn, id)
		CgLog "AGENDA", id, "ACTUALIZA"
		AgendaExtRedir "postergada", id, xmenu, xtmenu
		Exit Sub
	End If
	If opc = "agenda_ext_seguimiento" Then
		id = AgendaExtSeguimiento(cn, id)
		CgLog "AGENDA", id, "ALTA"
		AgendaExtRedir "seguimiento", id, xmenu, xtmenu
		Exit Sub
	End If
	If opc = "agenda_ext_eliminar_corresp" Then
		AgendaExtEliminarCorresp cn, xmenu, xtmenu
		Exit Sub
	End If
	If AgendaExtEsIdValido(id) Then
		Set rs = Exec(cn, "sp_cg_AgendaExt @Modo='obtener', @Id=" & CLng(id))
		CgLog "AGENDA", id, "CONSULTA"
	Else
		Set rs = Nothing
	End If
	AgendaExtPagina cn, rs, id, xmenu, xtmenu
End Sub

' Usuario actual: la sesion puede venir vacia; si falta, UserLogon()
Function AgendaExtUsuario()
	AgendaExtUsuario = Trim(Session("usuario") & "")
	If AgendaExtUsuario = "" Then AgendaExtUsuario = Trim(UserLogon() & "")
End Function

' =============================================================================
' CONSULTA DE REUNIONES (opc=consulta)
' =============================================================================
Sub AgendaExtConsulta(cn, xmenu, xtmenu)
	Dim rs, mover, mostrar, pos, volver, editar, estado, padreLista, filas, h, acc, pagAnt, pagSig, base
	mostrar = 15
	mover = AgendaExtSqlInt(Request.QueryString("CantRegMover"), 0)
	If mover < 0 Then mover = 0
	Set rs = Exec(cn, "sp_cg_AgendaExt @Modo='consulta', @CantRegMover=" & mover & ", @CantRegMostrar=" & mostrar)
	base = "&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu)
	volver = "agenda.asp?" & Mid(base, 2)

	filas = ""
	pos = 0
	Do While Not rs.EOF And pos < mostrar
		estado = AgendaExtVal(rs, "ESTADO", "PENDIENTE")
		editar = "agenda.asp?opc=edicion&id=" & rs("ID") & base
		padreLista = AgendaExtVal(rs, "ID_AGENDA_PADRE", "")
		If AgendaExtEsIdValido(padreLista) Then
			padreLista = "<a class='ui-link' href='agenda.asp?opc=det&id=" & CLng(padreLista) & base & "'>Reuni&oacute;n #" & CLng(padreLista) & "</a>"
		Else
			padreLista = "<span class='ui-status-empty'>&mdash;</span>"
		End If
		filas = filas & "<tr>" & _
			"<td><a class='ui-tag-id ui-tag-id--link' href='" & editar & "' aria-label='Abrir reuni&oacute;n " & rs("ID") & "'>#" & rs("ID") & "</a></td>" & _
			"<td><span class='ui-table__titulo'>" & UiTexto(Left(AgendaExtVal(rs, "ASUNTO", ""), 80)) & "</span></td>" & _
			"<td><span class='ui-table__nowrap'>" & UiFecha(AgendaExtVal(rs, "FECHA_REUNION", "")) & "</span></td>" & _
			"<td><span class='ui-table__nowrap'>" & AgendaExtHtmlTime(AgendaExtVal(rs, "HORA_INICIO", "")) & " a " & AgendaExtHtmlTime(AgendaExtVal(rs, "HORA_FIN", "")) & "</span></td>" & _
			"<td>" & UiDato(AgendaExtVal(rs, "RESPONSABLE", "")) & "</td>" & _
			"<td>" & AgendaExtBadgeEstado(estado) & "</td>" & _
			"<td>" & padreLista & "</td></tr>"
		pos = pos + 1
		rs.MoveNext
	Loop

	acc = "<a class='ui-button ui-button--light ui-button--sm' href='agenda.asp?opc=edicion" & base & "'>+ Nueva reuni&oacute;n</a>" & _
		"<a class='ui-button ui-button--secondary ui-button--sm' href='" & volver & "'>Volver</a>"
	h = headhtml() & "<body><div id='wrapper' role='banner'>" & cabecera() & "</div>"
	h = h & UiPantallaAbre("Reuniones", acc, "wide")
	h = h & UiCardAbre("Reuniones", "") & UiTablaAbre("Reuniones: n&uacute;mero, asunto, fecha, horario, responsable, estado y reuni&oacute;n de origen")
	h = h & "<thead><tr>" & UiTh("N&ordm;", "") & UiTh("Asunto", "") & UiTh("Fecha", "") & UiTh("Horario", "") & UiTh("Responsable", "") & UiTh("Estado", "") & UiTh("Seguimiento de", "") & "</tr></thead><tbody>"
	If pos = 0 Then
		h = h & UiFilaVacia(7, "No hay reuniones cargadas.")
	Else
		h = h & filas
	End If
	h = h & UiTablaCierra()
	If mover > 0 Then
		pagAnt = "<a class='ui-button ui-button--secondary ui-button--sm' href='agenda.asp?opc=consulta&CantRegMover=" & (mover - mostrar) & base & "'>Anterior</a>"
	Else
		pagAnt = UiBoton("Anterior", "secondary", "", False, "Primera p&aacute;gina")
	End If
	If Not rs.EOF Then
		pagSig = "<a class='ui-button ui-button--secondary ui-button--sm' href='agenda.asp?opc=consulta&CantRegMover=" & (mover + mostrar) & base & "'>Siguiente</a>"
	Else
		pagSig = UiBoton("Siguiente", "secondary", "", False, "&Uacute;ltima p&aacute;gina")
	End If
	h = h & UiPaginacion(mover + 1, mover + pos, pagAnt, pagSig)
	h = h & UiCardCierra() & UiPantallaCierra()
	h = h & "<script src='assets/js/jquery.min.js'></script><script src='assets/js/skel.min.js'></script><script src='assets/js/util.js'></script><script src='assets/js/main.js'></script>"
	h = h & "</body></html>"
	rs.Close
	Response.Write h
End Sub

' =============================================================================
' GUARDAR (alta o edicion), con validacion de cruces de cada participante del sistema
' =============================================================================
Function AgendaExtGuardar(cn, idActual, ByRef avisosCruceOut)
	Dim q, rs, id, i, n, persona, nombre, email, copia, tipo, relacionId, proyecto, incidente, tema, conclusion, accion
	Dim fechaReunionSql, horaInicioSql, horaFinSql, avisosCruce, descripcionCruce, vistos, claveP
	vistos = "|"
	fechaReunionSql = AgendaExtSqlDate(Request.Form("fecha_reunion"))
	horaInicioSql = AgendaExtSqlTime(Request.Form("hora_inicio"))
	horaFinSql = AgendaExtSqlTime(Request.Form("hora_fin"))
	avisosCruce = ""
	q = "sp_cg_AgendaExt @Modo='insertar'"
	If AgendaExtEsIdValido(idActual) Then q = "sp_cg_AgendaExt @Modo='actualizar', @Id=" & CLng(idActual)
	q = q & ", @Asunto=" & AgendaExtSqlText(Request.Form("asunto"))
	q = q & ", @FechaReunion=" & fechaReunionSql
	q = q & ", @HoraInicio=" & horaInicioSql
	q = q & ", @HoraFin=" & horaFinSql
	q = q & ", @FechaAviso=" & AgendaExtSqlDate(Request.Form("fecha_aviso"))
	' [UI] El responsable es un login (texto): antes iba por AgendaExtSqlInt, que esperaba un numero
	q = q & ", @IdRespon=" & AgendaExtSqlStr(Request.Form("id_respon"), AgendaExtUsuario())
	q = q & ", @IdHerramienta=" & AgendaExtSqlNullableInt(Request.Form("id_herramienta"))
	q = q & ", @Mensaje=" & AgendaExtSqlText(Request.Form("mensaje"))
	q = q & ", @Conclusion=" & AgendaExtSqlText(Request.Form("conclusion"))
	' Descripcion breve (CG_AGENDA.DESCRIPCION, hasta 50): siempre se envia, '' la borra (DB/Cambios_Agenda_Descripcion.sql)
	q = q & ", @Descripcion=N'" & Replace(Left(Trim(Request.Form("descripcion") & ""), 50), "'", "''") & "'"
	' Enlace de la reunion (CG_AGENDA.ENLACE_REUNION, hasta 500): siempre se envia, '' lo borra
	q = q & ", @Enlace=N'" & Replace(Left(Trim(Request.Form("enlace") & ""), 500), "'", "''") & "'"
	q = q & ", @Estado=" & AgendaExtSqlText(Request.Form("estado"))
	If AgendaExtEsIdValido(idActual) Then
		ExecNonQuery cn, q
		id = CLng(idActual)
	Else
		Set rs = Exec(cn, q & ", @TipoOrigen='N'")
		id = CLng(rs("ID"))
	End If

	If AgendaExtEsIdValido(idActual) Then
		ExecNonQuery cn, "sp_cg_AgendaExt @Modo='corresp_limpiar', @Id=" & id
		ExecNonQuery cn, "sp_cg_AgendaExt @Modo='subtema_limpiar', @Id=" & id
	End If

	' --- Participantes ---
	n = AgendaExtSqlInt(Request.Form("corresponsable_count"), 0)
	For i = 1 To n
		persona = Trim(Request.Form("cor_persona_" & i) & "")
		nombre = Request.Form("cor_nombre_" & i)
		email = Request.Form("cor_email_" & i)
		copia = AgendaExtSqlBit(Request.Form("cor_copia_" & i))
		' Sin duplicados: la misma persona (o el mismo contacto manual) se agrega una sola vez por guardado
		If persona <> "" Then claveP = "p:" & LCase(persona) Else claveP = "m:" & LCase(Trim(nombre & "")) & "/" & LCase(Trim(email & ""))
		If InStr(vistos, "|" & claveP & "|") > 0 Then
			persona = ""
			nombre = ""
		Else
			vistos = vistos & claveP & "|"
		End If
		If persona <> "" Then
			' [UI] Persona del sistema por login: se guarda siempre (ID_PERSONA_SEG); si tiene otra reunion en el mismo horario, solo se avisa (Alerta)
			descripcionCruce = AgendaExtDescribirCruceParticipante(cn, id, persona, fechaReunionSql, horaInicioSql, horaFinSql)
			If descripcionCruce <> "" Then
				avisosCruce = avisosCruce & AgendaExtNombrePersona(cn, persona) & ": " & descripcionCruce & vbLf
			End If
			ExecNonQuery cn, "sp_cg_AgendaExt @Modo='corresp_agregar', @Id=" & id & ", @IdPersona=" & AgendaExtSqlText(persona) & ", @EnviarCopia=" & copia
		ElseIf Trim(nombre & "") <> "" Then
			ExecNonQuery cn, "sp_cg_AgendaExt @Modo='corresp_agregar', @Id=" & id & ", @NombreManual=" & AgendaExtSqlText(nombre) & ", @EmailManual=" & AgendaExtSqlText(email) & ", @EnviarCopia=" & copia
		End If
	Next

	' --- Temario ---
	n = AgendaExtSqlInt(Request.Form("subtema_count"), 0)
	For i = 1 To n
		tipo = AgendaExtSqlInt(Request.Form("sub_tipo_" & i), 0)
		' Las filas quitadas en pantalla no se envian: se saltean
		If tipo > 0 Then
			relacionId = AgendaExtSqlInt(Request.Form("sub_relacion_id_" & i), 0)
			proyecto = 0
			incidente = 0
			If tipo = 1 Then proyecto = relacionId
			If tipo = 2 Then incidente = relacionId
			tema = Request.Form("sub_tema_" & i)
			conclusion = Request.Form("sub_conclusion_" & i)
			accion = AgendaExtSqlInt(Request.Form("sub_accion_" & i), 0)
			ExecNonQuery cn, "sp_cg_AgendaExt @Modo='subtema_agregar', @Id=" & id & ", @IdTipoRelacion=" & tipo & ", @IdProyecto=" & AgendaExtSqlNullableInt(proyecto) & ", @IdIncidente=" & AgendaExtSqlNullableInt(incidente) & ", @Tema=" & AgendaExtSqlText(tema) & ", @Conclusion=" & AgendaExtSqlText(conclusion) & ", @TiempoUsado=" & AgendaExtSqlInt(Request.Form("sub_tiempo_" & i), 0) & ", @Orden=" & i & ", @Seguimiento=" & AgendaExtSqlBit(Request.Form("sub_seguimiento_" & i)) & ", @IdAccion=" & AgendaExtSqlNullableInt(accion)
		End If
	Next
	avisosCruceOut = avisosCruce
	AgendaExtGuardar = id
End Function

' Otra reunion del participante en el mismo horario (vacio si no hay). idPersona: login
Function AgendaExtDescribirCruceParticipante(cn, idAgendaActual, idPersona, fechaReunionSql, horaInicioSql, horaFinSql)
	Dim rsCruce, detalle, idsVistos
	idsVistos = ","
	AgendaExtDescribirCruceParticipante = ""
	If fechaReunionSql = "NULL" Or horaInicioSql = "NULL" Or horaFinSql = "NULL" Then Exit Function
	Set rsCruce = Exec(cn, "sp_cg_Agenda_ValidarCruceHorario @IdResponsable=" & AgendaExtSqlText(idPersona) & ", @FechaReunion=" & fechaReunionSql & ", @HoraInicio=" & horaInicioSql & ", @HoraFin=" & horaFinSql & ", @IdAgendaExcluir=" & CLng(idAgendaActual))
	detalle = ""
	Do While Not rsCruce.EOF
	  If InStr(idsVistos, "," & rsCruce("ID_AGENDA_CRUCE") & ",") = 0 Then
		idsVistos = idsVistos & rsCruce("ID_AGENDA_CRUCE") & ","
		If detalle <> "" Then detalle = detalle & "; "
		detalle = detalle & Chr(34) & AgendaExtVal(rsCruce, "ASUNTO", "(sin asunto)") & Chr(34) & " (reuni" & Chr(243) & "n #" & rsCruce("ID_AGENDA_CRUCE") & ", " & AgendaExtHtmlTime(rsCruce("HORA_INICIO")) & " a " & AgendaExtHtmlTime(rsCruce("HORA_FIN")) & ")"
	  End If
		rsCruce.MoveNext
	Loop
	rsCruce.Close
	AgendaExtDescribirCruceParticipante = detalle
End Function

' Nombre de una persona del sistema por su login
Function AgendaExtNombrePersona(cn, idPersona)
	Dim rsPersona, nombre
	nombre = idPersona & ""
	Set rsPersona = Exec(cn, "SELECT MAX(USRNAME) AS NOMBRE FROM dbo.vw_UserGrupos WHERE USRNET=" & AgendaExtSqlText(idPersona))
	If Not rsPersona.EOF Then
		If Trim(AgendaExtVal(rsPersona, "NOMBRE", "")) <> "" Then nombre = Trim(rsPersona("NOMBRE"))
	End If
	rsPersona.Close
	AgendaExtNombrePersona = nombre
End Function

Function AgendaExtPostergar(cn, id)
	Dim rs
	Set rs = Exec(cn, "sp_cg_AgendaExt @Modo='postergar', @Id=" & CLng(id) & ", @FechaReunion=" & AgendaExtSqlDate(Request.Form("fecha_reunion")) & ", @HoraInicio=" & AgendaExtSqlTime(Request.Form("hora_inicio")) & ", @HoraFin=" & AgendaExtSqlTime(Request.Form("hora_fin")))
	AgendaExtPostergar = CLng(rs("ID"))
End Function

Function AgendaExtSeguimiento(cn, id)
	Dim rs
	Set rs = Exec(cn, "sp_cg_AgendaExt @Modo='seguimiento', @Id=" & CLng(id) & ", @FechaReunion=" & AgendaExtSqlDate(Request.Form("fecha_reunion")) & ", @HoraInicio=" & AgendaExtSqlTime(Request.Form("hora_inicio")) & ", @HoraFin=" & AgendaExtSqlTime(Request.Form("hora_fin")) & ", @Conclusion=" & AgendaExtSqlText(Request.Form("conclusion")))
	AgendaExtSeguimiento = CLng(rs("ID"))
End Function

Sub AgendaExtEliminarCorresp(cn, xmenu, xtmenu)
	Dim id, correspId
	id = Request.QueryString("id")
	correspId = Request.QueryString("corresp_id")
	If AgendaExtEsIdValido(id) And AgendaExtEsIdValido(correspId) Then
		ExecNonQuery cn, "sp_cg_AgendaExt @Modo='corresp_eliminar', @Id=" & CLng(id) & ", @IdCorresponsable=" & CLng(correspId)
		Response.Redirect "agenda.asp?opc=edicion&id=" & CLng(id) & "&msg=participante&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu) & AgendaExtVuelta()
	Else
		Response.Redirect AgendaExtUrlVolver(xmenu, xtmenu)
	End If
End Sub

' =============================================================================
' PAGINA DE LA REUNION (alta o edicion)
' =============================================================================
Sub AgendaExtPagina(cn, rs, id, xmenu, xtmenu)
	Dim editar, asunto, descripcionBreve, enlaceReunion, fecha, aviso, hi, hf, mensaje, conclusion, resp, herramienta, estado, agendaId, agendaPadreId, volver, base
	Dim h, acc, tituloPagina, guardarAction, msg, cruces
	editar = False
	If IsObject(rs) Then
		If Not rs Is Nothing Then
			If Not rs.EOF Then editar = True
		End If
	End If
	If editar Then
		agendaId = AgendaExtVal(rs, "ID", id)
		agendaPadreId = AgendaExtVal(rs, "ID_AGENDA_PADRE", "")
		asunto = AgendaExtVal(rs, "ASUNTO", "")
		descripcionBreve = AgendaExtVal(rs, "DESCRIPCION", "")
		enlaceReunion = AgendaExtVal(rs, "ENLACE_REUNION", "")
		fecha = AgendaExtVal(rs, "FECHA_REUNION", "")
		aviso = AgendaExtVal(rs, "FECHA_AVISO", "")
		hi = AgendaExtVal(rs, "HORA_INICIO", "")
		hf = AgendaExtVal(rs, "HORA_FIN", "")
		mensaje = AgendaExtVal(rs, "MENSAJE", "")
		conclusion = AgendaExtVal(rs, "CONCLUSION", "")
		resp = AgendaExtVal(rs, "ID_RESPON", AgendaExtUsuario())
		herramienta = AgendaExtVal(rs, "ID_HERRAMIENTA", 0)
		estado = AgendaExtVal(rs, "ESTADO", "PENDIENTE")
	Else
		agendaId = ""
		agendaPadreId = ""
		asunto = ""
		descripcionBreve = ""
		enlaceReunion = ""
		fecha = AgendaExtHoy()
		aviso = AgendaExtHoy()
		hi = 900
		hf = 1000
		mensaje = ""
		conclusion = ""
		resp = AgendaExtUsuario()
		herramienta = 0
		estado = "PENDIENTE"
	End If
	' [UI] Si la reunion se abrio desde la vista semanal o mensual, Volver y Cancelar llevan ahi (y se mantiene al guardar)
	base = "&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu) & AgendaExtVuelta()
	volver = AgendaExtUrlVolver(xmenu, xtmenu)

	If editar Then tituloPagina = "Reuni&oacute;n #" & UiTexto(agendaId) Else tituloPagina = "Nueva reuni&oacute;n"
	'[UI-ANT] acc = "<a class='ui-button ui-button--light ui-button--sm' href='agenda_semana.asp?" & Mid(base, 2) & "'>Vista semanal</a>" & _
	'[UI-ANT] "<a class='ui-button ui-button--secondary ui-button--sm' href='" & volver & "'>Volver</a>"
	' [UI] 'Vista semanal' solo si no se vino de la agenda semanal o mensual (en ese caso Volver ya lleva ahi)
	acc = ""
	If AgendaExtVuelta() = "" Then acc = "<a class='ui-button ui-button--light ui-button--sm' href='agenda_semana.asp?xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu) & "'>Vista semanal</a>"
	acc = acc & "<a class='ui-button ui-button--secondary ui-button--sm' href='" & volver & "'>Volver</a>"

	h = headhtml() & "<body><div id='wrapper' role='banner'>" & cabecera() & "</div>"
	' Ancho de 1240px: el temario tiene ocho columnas editables (y es el mismo ancho de la vista semanal)
	h = h & UiPantallaAbre(tituloPagina, acc, "wide")

	' --- Mensajes despues de guardar (antes alert) ---
	msg = Request.QueryString("msg") & ""
	If msg = "guardada" Then h = h & "<div class='ui-alert ui-alert--success' role='status'>Reuni&oacute;n guardada.</div>"
	If msg = "postergada" Then h = h & "<div class='ui-alert ui-alert--success' role='status'>La reuni&oacute;n fue postergada. Esta es la nueva reuni&oacute;n.</div>"
	If msg = "seguimiento" Then h = h & "<div class='ui-alert ui-alert--success' role='status'>Se gener&oacute; la reuni&oacute;n de seguimiento. Esta es la nueva reuni&oacute;n.</div>"
	If msg = "participante" Then h = h & "<div class='ui-alert ui-alert--success' role='status'>Participante quitado.</div>"
	If msg = "invitacion" Then h = h & AgInvMensajeHtml()
	cruces = Trim(Session("agendaExtCruces") & "")
	If cruces <> "" And msg = "guardada" Then
		h = h & "<div class='ui-alert ui-alert--warning' role='alert'><strong>Alerta</strong>: algunos participantes ya tienen otra reuni&oacute;n en ese horario:<ul class='ui-alert__lista'>"
		Dim lineaCruce
		For Each lineaCruce In Split(cruces, vbLf)
			If Trim(lineaCruce) <> "" Then h = h & "<li>" & UiTexto(lineaCruce) & "</li>"
		Next
		h = h & "</ul></div>"
	End If
	Session("agendaExtCruces") = ""
	If AgendaExtEsIdValido(agendaPadreId) Then
		h = h & "<div class='ui-alert' role='note'>Es un seguimiento de la <a class='ui-link' href='agenda.asp?opc=det&id=" & CLng(agendaPadreId) & base & "'>reuni&oacute;n #" & CLng(agendaPadreId) & "</a>.</div>"
	End If

	guardarAction = "agenda.asp?opc=agenda_ext_guardar"
	If editar Then guardarAction = guardarAction & "&id=" & CLng(id)
	guardarAction = guardarAction & base
	' [INVITACION] Tarjeta de invitacion por mail y calendarios (solo con la reunion ya guardada); va FUERA del formulario de la reunion
	If editar Then h = h & AgInvCardHtml(cn, id, base)
	h = h & "<form id='agendaExt' method='post' action='" & guardarAction & "' onsubmit='return agendaExtValidar();'>"

	' --- Datos de la reunion ---
	h = h & UiCardAbre("Datos de la reuni&oacute;n", "") & UiNotaObligatorios() & UiFormGridAbre()
	h = h & UiCampo("asunto", "Asunto", AgendaExtControl("asunto", asunto, "text", True, "maxlength='1000'"), True, True)
	h = h & UiCampo("descripcion", "Descripci&oacute;n breve", AgendaExtControl("descripcion", descripcionBreve, "text", False, "maxlength='50'"), False, True)
	h = h & UiCampo("id_respon", "Responsable", AgendaExtSelectResponsables(cn, resp), False, False)
	h = h & UiCampo("estado", "Estado", AgendaExtSelectEstados(cn, estado), True, False)
	h = h & UiCampo("fecha_reunion", "Fecha de la reuni&oacute;n", AgendaExtControl("fecha_reunion", AgendaExtHtmlDate(fecha), "date", True, ""), True, False)
	h = h & UiCampo("fecha_aviso", "Fecha de aviso", AgendaExtControl("fecha_aviso", AgendaExtHtmlDate(aviso), "date", True, ""), True, False)
	h = h & UiCampo("hora_inicio", "Hora de inicio", AgendaExtControl("hora_inicio", AgendaExtHtmlTime(hi), "time", True, ""), True, False)
	h = h & UiCampo("hora_fin", "Hora de fin", AgendaExtControl("hora_fin", AgendaExtHtmlTime(hf), "time", True, ""), True, False)
	h = h & UiCampo("id_herramienta", "Herramienta", AgendaExtSelectHerramientas(cn, herramienta), False, False)
	' [INVITACION] Enlace de Teams / Meet: se incluye en el mail y en el .ics de la invitacion (DB/Cambios_Agenda_Invitacion.sql)
	h = h & UiCampo("enlace", "Enlace de la reuni&oacute;n (Teams / Meet)", AgendaExtControl("enlace", enlaceReunion, "text", False, "maxlength='500' placeholder='https://teams.microsoft.com/...'"), False, True)
	h = h & UiCampo("mensaje", "Mensaje", "<textarea class='ui-input ui-textarea' id='mensaje' name='mensaje' rows='3'>" & Server.HTMLEncode(mensaje & "") & "</textarea>", False, True)
	h = h & UiFormGridCierra() & UiCardCierra()

	h = h & AgendaExtCorresponsables(cn, id, xmenu, xtmenu)
	h = h & AgendaExtSubtemas(cn, id, xmenu, xtmenu)

	' --- Conclusion ---
	h = h & UiCardAbre("Conclusi&oacute;n general", "")
	h = h & "<label class='ui-sr-only' for='conclusion'>Conclusi&oacute;n general de la reuni&oacute;n</label>"
	h = h & "<textarea class='ui-input ui-textarea' id='conclusion' name='conclusion' rows='4'>" & Server.HTMLEncode(conclusion & "") & "</textarea>"
	h = h & UiCardCierra()

	' --- Acciones (mismos destinos: postergar y seguimiento usan la fecha y el horario del formulario) ---
	' Botones sobre el fondo de la pagina: variante blanca (la secundaria se confunde con el fondo)
	h = h & "<div class='ui-agenda-acciones'><a class='ui-button ui-button--light' href='" & volver & "'>Cancelar</a><div class='ui-agenda-acciones__der'>"
	If editar Then
		h = h & "<button type='submit' class='ui-button ui-button--light' formaction='agenda.asp?opc=agenda_ext_postergar&id=" & CLng(id) & base & "' title='Usa la fecha y el horario cargados arriba'>Postergar</button>"
		h = h & "<button type='submit' class='ui-button ui-button--light' formaction='agenda.asp?opc=agenda_ext_seguimiento&id=" & CLng(id) & base & "' title='Crea una reuni&oacute;n nueva con la fecha y el horario cargados arriba'>Generar seguimiento</button>"
	End If
	h = h & "<button type='submit' class='ui-button ui-button--primary'>Guardar reuni&oacute;n</button></div></div>"
	If editar Then h = h & "<p class='ui-table__nota ui-agenda-acciones__nota'>Para postergar o generar un seguimiento, cambi&aacute; la fecha y el horario de arriba antes de tocar el bot&oacute;n.</p>"
	h = h & "</form>"
	h = h & UiPantallaCierra()

	h = h & "<script>" & AgendaExtJs() & "</script>"
	h = h & "<script src='assets/js/jquery.min.js'></script><script src='assets/js/skel.min.js'></script><script src='assets/js/util.js'></script><script src='assets/js/main.js'></script>"
	h = h & "</body></html>"
	Response.Write h
End Sub

' --- Participantes ---
Function AgendaExtCorresponsables(cn, id, xmenu, xtmenu)
	Dim rs, detalle, nombre, persona, manual, email, copia, n, correspId, h, eliminarUrl, chips, mailOpc
	h = UiCardAbre("Participantes", "")
	h = h & "<div class='ui-agenda-alta'>"
	h = h & "<div class='ui-field'><label class='ui-field__label' for='corPersona'>Persona del sistema</label><select class='ui-select' id='corPersona'><option value=''>Elegir persona</option>"
	Set rs = Exec(cn, "sp_cg_Responsables @Modo='listar_activos_sin', @ExcluirId=" & AgendaExtSqlText(AgendaExtUsuario()))
	Do While Not rs.EOF
		' el mail de la persona (columna MAIL del SP) viaja en data-mail para mostrarlo al agregarla
		If rs.Fields.Count > 5 Then mailOpc = rs(5) & "" Else mailOpc = ""
		h = h & "<option value='" & Server.HTMLEncode(rs(0) & "") & "' data-mail='" & Server.HTMLEncode(mailOpc) & "'>" & Server.HTMLEncode(Trim(rs(1) & " " & rs(2))) & "</option>"
		rs.MoveNext
	Loop
	h = h & "</select></div>"
	h = h & "<p class='ui-agenda-alta__o' aria-hidden='true'>o</p>"
	h = h & "<div class='ui-field'><label class='ui-field__label' for='corNombre'>Contacto externo: nombre</label><input class='ui-input' id='corNombre' type='text' autocomplete='off'></div>"
	h = h & "<div class='ui-field'><label class='ui-field__label' for='corEmail'>Mail</label><input class='ui-input' id='corEmail' type='email' autocomplete='off'></div>"
	h = h & "<label class='ui-opcion ui-agenda-alta__copia'><input class='ui-check' id='corCopia' type='checkbox' checked><span>Enviar copia</span></label>"
	h = h & "<button type='button' class='ui-button ui-button--secondary' onclick='agendaExtAgregarCorresponsable()'>+ Agregar</button>"
	h = h & "</div>"

	chips = ""
	n = 0
	If AgendaExtEsIdValido(id) Then
		Set detalle = Exec(cn, "sp_cg_AgendaExt @Modo='corresp_listar', @Id=" & CLng(id))
		Do While Not detalle.EOF
			n = n + 1
			correspId = AgendaExtVal(detalle, "ID_AGENDA_CORRESP", "")
			persona = AgendaExtVal(detalle, "ID_PERSONA", "")
			manual = AgendaExtVal(detalle, "NOMBRE_MANUAL", "")
			email = AgendaExtVal(detalle, "EMAIL_MANUAL", "")
			copia = AgendaExtVal(detalle, "ENVIAR_COPIA", 1)
			nombre = AgendaExtVal(detalle, "NOMBRE_PERSONA", manual)
			eliminarUrl = "agenda.asp?opc=agenda_ext_eliminar_corresp&id=" & CLng(id) & "&corresp_id=" & correspId & "&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu) & AgendaExtVuelta()   ' [UI] mantiene la vista de origen al quitar un participante
			chips = chips & "<li class='ui-copia'><span class='ui-participante'>" & UiTexto(nombre)
			If Trim(email & "") <> "" Then chips = chips & " <span class='ui-participante__mail'>" & UiTexto(email) & "</span>"
			If CStr(copia) = "1" Or copia = True Then chips = chips & " <span class='ui-badge ui-badge--blue'>Copia</span>"
			chips = chips & "</span>" & _
				"<input type='hidden' name='cor_persona_" & n & "' value='" & Server.HTMLEncode(persona & "") & "'>" & _
				"<input type='hidden' name='cor_nombre_" & n & "' value='" & Server.HTMLEncode(manual & "") & "'>" & _
				"<input type='hidden' name='cor_email_" & n & "' value='" & Server.HTMLEncode(email & "") & "'>" & _
				"<input type='hidden' name='cor_copia_" & n & "' value='" & AgendaExtSqlBit(copia) & "'>" & _
				"<a class='ui-button ui-button--icon ui-button--danger' href='" & eliminarUrl & "' aria-label='Quitar a " & UiTexto(nombre) & "' onclick=""return confirm('\u00bfQuitar este participante?');"">&times;</a></li>"
			detalle.MoveNext
		Loop
	End If
	h = h & "<ul class='ui-copias ui-agenda-participantes' id='corresponsables'>" & chips & "</ul>"
	h = h & "<p class='ui-status-empty' id='corVacio'" & AgendaExtOculto(n > 0) & ">Todav&iacute;a no hay participantes.</p>"
	h = h & "<input type='hidden' id='corresponsable_count' name='corresponsable_count' value='" & n & "'>"
	h = h & UiCardCierra()
	AgendaExtCorresponsables = h
End Function

' --- Temario ---
Function AgendaExtSubtemas(cn, id, xmenu, xtmenu)
	Dim detalle, n, tipo, proyecto, incidente, relacionId, tema, conclusion, minutos, seguimiento, accion, h, filas
	filas = ""
	n = 0
	If AgendaExtEsIdValido(id) Then
		Set detalle = Exec(cn, "sp_cg_AgendaExt @Modo='subtema_listar', @Id=" & CLng(id))
		Do While Not detalle.EOF
			n = n + 1
			tipo = AgendaExtVal(detalle, "ID_TIPO_RELACION", 0)
			proyecto = AgendaExtVal(detalle, "ID_PROYECTO", "")
			incidente = AgendaExtVal(detalle, "ID_INCIDENTE", "")
			tema = AgendaExtVal(detalle, "TEMA_A_TRATAR", "")
			conclusion = AgendaExtVal(detalle, "CONCLUSION", "")
			minutos = AgendaExtVal(detalle, "TIEMPO_USADO", 0)
			seguimiento = AgendaExtVal(detalle, "SEGUIMIENTO", 0)
			accion = AgendaExtVal(detalle, "ID_ACCION", "")
			relacionId = ""
			If tipo = 1 Then relacionId = proyecto
			If tipo = 2 Then relacionId = incidente
			filas = filas & AgendaExtFilaTema(cn, n, tipo, relacionId, tema, conclusion, minutos, seguimiento, accion)
			detalle.MoveNext
		Loop
	End If
	h = UiCardAbre("Temario", "<button type='button' class='ui-button ui-button--light ui-button--sm' onclick='agendaExtAgregarSubtema()'>+ Agregar tema</button>")
	h = h & UiTablaAbre("Temas de la reuni&oacute;n")
	h = h & "<thead><tr>" & UiTh("Tipo", "") & UiTh("Proyecto o incidente", "") & UiTh("Tema", "") & UiTh("Seguimiento", "ui-table__cell--center") & UiTh("Conclusi&oacute;n", "") & UiTh("Min.", "") & UiTh("Acci&oacute;n", "") & UiTh("<span class='ui-sr-only'>Quitar</span>", "") & "</tr></thead>"
	h = h & "<tbody id='subtemasBody'>" & filas & UiTablaCierra()
	h = h & "<p class='ui-status-empty' id='subVacio'" & AgendaExtOculto(n > 0) & ">Todav&iacute;a no hay temas. Sum&aacute; el primero con &laquo;Agregar tema&raquo;.</p>"
	h = h & "<input type='hidden' id='subtema_count' name='subtema_count' value='" & n & "'>"
	h = h & "<template id='agendaTemaTemplate'>" & AgendaExtFilaTema(cn, "__N__", 1, "", "", "", 0, 0, 0) & "</template>"
	h = h & UiCardCierra()
	AgendaExtSubtemas = h
End Function

' Una fila del temario (la misma para las existentes y las nuevas; __N__ se reemplaza en el navegador)
Function AgendaExtFilaTema(cn, n, tipo, relacionId, tema, conclusion, minutos, seguimiento, accion)
	Dim h, etiquetaRel
	If CStr(tipo) = "2" Then etiquetaRel = "N&ordm; de incidente" Else etiquetaRel = "N&ordm; de proyecto"
	h = "<tr>"
	h = h & "<td><select class='ui-select ui-select--compacta' name='sub_tipo_" & n & "' aria-label='Tipo del tema' onchange='agendaExtCambiarRelacion(this)'>" & _
		"<option value='1'" & AgendaExtSelected(1, tipo) & ">Proyecto</option><option value='2'" & AgendaExtSelected(2, tipo) & ">Incidente</option>" & _
		"<option value='3'" & AgendaExtSelected(3, tipo) & ">Consultor&iacute;a</option><option value='4'" & AgendaExtSelected(4, tipo) & ">Otro</option></select></td>"
	h = h & "<td><div class='relation-fields ui-agenda-relacion'><input class='ui-input relation-id' name='sub_relacion_id_" & n & "' type='number' min='1' value='" & Server.HTMLEncode(relacionId & "") & "' placeholder='N&ordm;' aria-label='" & etiquetaRel & "'>" & _
		"<button type='button' class='ui-button ui-button--secondary ui-button--sm btn-buscar-entidad' onclick='agendaExtBuscarEntidad(this)'>Buscar</button></div></td>"
	h = h & "<td><input class='ui-input ui-agenda-tema' name='sub_tema_" & n & "' value='" & Server.HTMLEncode(tema & "") & "' required maxlength='500' aria-label='Tema a tratar'></td>"
	h = h & "<td class='ui-table__cell--center'><input class='ui-check' name='sub_seguimiento_" & n & "' type='checkbox' value='1'" & AgendaExtChecked(seguimiento) & " aria-label='Requiere seguimiento'></td>"
	h = h & "<td><textarea class='ui-input ui-textarea ui-agenda-conclusion' name='sub_conclusion_" & n & "' rows='2' maxlength='4000' aria-label='Conclusi&oacute;n del tema'>" & Server.HTMLEncode(conclusion & "") & "</textarea></td>"
	h = h & "<td><input class='ui-input ui-agenda-minutos' name='sub_tiempo_" & n & "' type='number' min='0' value='" & Server.HTMLEncode(minutos & "") & "' aria-label='Minutos usados'></td>"
	h = h & "<td><select class='ui-select ui-select--compacta' name='sub_accion_" & n & "' aria-label='Acci&oacute;n'>" & AgendaExtOpcionesAccion(cn, accion) & "</select></td>"
	h = h & "<td><button type='button' class='ui-button ui-button--icon ui-button--danger' aria-label='Quitar tema' onclick='agendaExtQuitarFila(this)'>&times;</button></td>"
	h = h & "</tr>"
	AgendaExtFilaTema = h
End Function

Function AgendaExtSelectResponsables(cn, resp)
	Dim rs, h
	h = "<select class='ui-select' id='id_respon' name='id_respon'>"
	Set rs = Exec(cn, "sp_cg_Responsables @Modo='listar_activos_sin', @ExcluirId=''")
	Do While Not rs.EOF
		h = h & "<option value='" & Server.HTMLEncode(rs(0) & "") & "'" & AgendaExtSelected(rs(0), resp) & ">" & Server.HTMLEncode(Trim(rs(1) & " " & rs(2))) & "</option>"
		rs.MoveNext
	Loop
	AgendaExtSelectResponsables = h & "</select>"
End Function

Function AgendaExtSelectHerramientas(cn, sel)
	Dim rs, h
	h = "<select class='ui-select' id='id_herramienta' name='id_herramienta'><option value='0'>Ninguna</option>"
	Set rs = Exec(cn, "sp_cg_Catalogo_Listar @Catalogo='aplicativos'")
	Do While Not rs.EOF
		h = h & "<option value='" & rs(0) & "'" & AgendaExtSelected(rs(0), sel) & ">" & Server.HTMLEncode(rs(1) & "") & "</option>"
		rs.MoveNext
	Loop
	AgendaExtSelectHerramientas = h & "</select>"
End Function

Function AgendaExtSelectEstados(cn, sel)
	Dim rs, h
	h = "<select class='ui-select' id='estado' name='estado' required aria-required='true'>"
	Set rs = Exec(cn, "sp_cg_AgendaExt @Modo='estados_listar'")
	Do While Not rs.EOF
		h = h & "<option value='" & Server.HTMLEncode(rs(1) & "") & "'" & AgendaExtSelected(rs(1), sel) & ">" & Server.HTMLEncode(rs(1) & "") & "</option>"
		rs.MoveNext
	Loop
	AgendaExtSelectEstados = h & "</select>"
End Function

' Campo de texto, fecha u hora con el estilo del sistema (mismo nombre de campo)
Function AgendaExtControl(fieldName, fieldValue, fieldType, required, extra)
	Dim h
	h = "<input class='ui-input' id='" & fieldName & "' name='" & fieldName & "' type='" & fieldType & "' value='" & Server.HTMLEncode(fieldValue & "") & "'"
	If required Then h = h & " required aria-required='true'"
	If extra <> "" Then h = h & " " & extra
	AgendaExtControl = h & ">"
End Function

Function AgendaExtBadgeEstado(estado)
	Select Case UCase(Trim(estado & ""))
		Case "PENDIENTE"
			AgendaExtBadgeEstado = UiBadge("Pendiente", "blue")
		Case "POSTERGADA"
			AgendaExtBadgeEstado = UiBadge("Postergada", "amber")
		Case "REALIZADA"
			AgendaExtBadgeEstado = UiBadge("Realizada", "gray")
		Case Else
			AgendaExtBadgeEstado = UiBadge(UiTexto(UiCapitalizar(estado)), "gray")
	End Select
End Function

Function AgendaExtOculto(ocultar)
	If ocultar Then AgendaExtOculto = " hidden" Else AgendaExtOculto = ""
End Function

Function AgendaExtSelected(value, selectedValue)
	If CStr(value & "") = CStr(selectedValue & "") Then AgendaExtSelected = " selected" Else AgendaExtSelected = ""
End Function
Function AgendaExtChecked(value)
	If AgendaExtSqlBit(value) = "1" Then AgendaExtChecked = " checked" Else AgendaExtChecked = ""
End Function
Function AgendaExtOpcionesAccion(cn, selectedValue)
	Dim rs, h
	h = "<option value='0'>Ninguna</option>"
	Set rs = Exec(cn, "sp_cg_AgendaExt @Modo='acciones_listar'")
	If Not rs Is Nothing Then
		If Not rs.State = 0 Then
			Do While Not rs.EOF
				h = h & "<option value='" & rs(0) & "'" & AgendaExtSelected(rs(0), selectedValue) & ">" & Server.HTMLEncode(rs(1) & "") & "</option>"
				rs.MoveNext
			Loop
		End If
	End If
	AgendaExtOpcionesAccion = h
End Function
Function AgendaExtSqlText(value)
	If Trim(value & "") = "" Then AgendaExtSqlText = "NULL" Else AgendaExtSqlText = "N'" & Replace(value & "", "'", "''") & "'"
End Function

Function AgendaExtSoloDigitos(value)
	Dim i, c, out
	out = ""
	For i = 1 To Len(value & "")
		c = Mid(value & "", i, 1)
		If c >= "0" And c <= "9" Then out = out & c
	Next
	If out = "" Then out = "0"
	AgendaExtSoloDigitos = out
End Function
Function AgendaExtHoy()
	AgendaExtHoy = CLng(CDate(Date())) - CLng(CDate("01-01-1900"))
End Function

Function AgendaExtSqlDate(value)
	If IsDate(value) Then AgendaExtSqlDate = CStr(CLng(CDate(value)) - CLng(CDate("01-01-1900"))) Else AgendaExtSqlDate = "NULL"
End Function

Function AgendaExtSqlTime(value)
	Dim t, hh, mm
	AgendaExtSqlTime = "NULL"
	t = Trim(value & "")
	If t = "" Then Exit Function
	t = Replace(t, ".", ":")
	If InStr(t, ":") > 0 Then
		hh = CLng(AgendaExtSoloDigitos(Split(t, ":")(0)))
		mm = CLng(AgendaExtSoloDigitos(Split(t, ":")(1)))
		If hh >= 0 And hh <= 23 And mm >= 0 And mm <= 59 Then AgendaExtSqlTime = CStr(hh * 100 + mm)
	ElseIf IsNumeric(t) Then
		AgendaExtSqlTime = CStr(CLng(t))
	End If
End Function

Function AgendaExtSqlStr(value, fallback)
	Dim v
	v = Trim(value & "")
	If v = "" Then v = Trim(fallback & "")
	AgendaExtSqlStr = "N'" & Replace(v, "'", "''") & "'"
End Function
Function AgendaExtSqlInt(value, fallback)
	If IsNumeric(value) Then AgendaExtSqlInt = CLng(value) Else AgendaExtSqlInt = fallback
End Function
Function AgendaExtSqlNullableInt(value)
	If IsNumeric(value) Then
		If CLng(value) > 0 Then AgendaExtSqlNullableInt = CStr(CLng(value)) Else AgendaExtSqlNullableInt = "NULL"
	Else
		AgendaExtSqlNullableInt = "NULL"
	End If
End Function
Function AgendaExtEsIdValido(value)
	AgendaExtEsIdValido = False
	If IsNumeric(value) Then
		If CLng(value) > 0 Then AgendaExtEsIdValido = True
	End If
End Function
Function AgendaExtSqlBit(value)
	If value & "" = "1" Or UCase(value & "") = "SI" Or UCase(value & "") = "TRUE" Or UCase(value & "") = "VERDADERO" Then AgendaExtSqlBit = "1" Else AgendaExtSqlBit = "0"
End Function
Function AgendaExtVal(rs, fieldName, fallback)
	If IsNull(rs(fieldName)) Then AgendaExtVal = fallback Else AgendaExtVal = rs(fieldName)
End Function

Function AgendaExtHtmlDate(value)
	Dim d
	AgendaExtHtmlDate = ""
	If IsNumeric(value) Then
		If CLng(value) > 0 Then
			d = CDate(CLng(value) + CLng(CDate("01-01-1900")))
			AgendaExtHtmlDate = Year(d) & "-" & Right("0" & Month(d), 2) & "-" & Right("0" & Day(d), 2)
		End If
	End If
End Function

Function AgendaExtHtmlTime(value)
	Dim n
	AgendaExtHtmlTime = ""
	If IsNumeric(value) Then
		n = CLng(value)
		If n >= 0 Then AgendaExtHtmlTime = Right("0" & CStr(n \ 100), 2) & ":" & Right("0" & CStr(n Mod 100), 2)
	End If
End Function

' Despues de guardar, postergar o generar seguimiento: a la reunion con un mensaje en la pagina (antes alert)
' [UI] Vista de la agenda desde la que se abrio la reunion: vuelta = semana | mes, vd = fecha de la semana o mes que se miraba.
' Devuelve los parametros listos para sumar a los enlaces (vacio si no vino de ahi o si los datos no son validos).
Function AgendaExtVuelta()
	Dim v, vd
	v = LCase(Trim(Request.QueryString("vuelta") & ""))
	vd = Trim(Request.QueryString("vd") & "")
	AgendaExtVuelta = ""
	If (v = "semana" Or v = "mes") And IsNumeric(vd) And vd <> "" Then AgendaExtVuelta = "&vuelta=" & v & "&vd=" & CLng(vd)
End Function

' [UI] Destino de Volver y Cancelar: la vista semanal o mensual de origen (en la misma semana o mes); si no, la lista
Function AgendaExtUrlVolver(xmenu, xtmenu)
	If AgendaExtVuelta() <> "" Then
		AgendaExtUrlVolver = "agenda_semana.asp?vista=" & LCase(Trim(Request.QueryString("vuelta") & "")) & "&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu) & "&d=" & CLng(Request.QueryString("vd"))
	Else
		AgendaExtUrlVolver = "agenda.asp?xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu)
	End If
End Function

Sub AgendaExtRedir(codigo, id, xmenu, xtmenu)
	Response.Redirect "agenda.asp?opc=det&id=" & CLng(id) & "&msg=" & codigo & "&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu) & AgendaExtVuelta()
End Sub

' JavaScript de la pagina. Mismos nombres de funciones (buscar_entidades.asp llama a agendaExtSeleccionarEntidad).
' Los campos se arman con el DOM: valores con espacios o comillas llegan completos.
Function AgendaExtJs()
	Dim j
	j = "var corN=parseInt(document.getElementById('corresponsable_count').value||0,10),subN=parseInt(document.getElementById('subtema_count').value||0,10);"
	j = j & "function agendaExtOculto(nombre,valor){var i=document.createElement('input');i.type='hidden';i.name=nombre;i.value=valor;return i;}"
	j = j & "function agendaExtActualizarVacios(){document.getElementById('corVacio').hidden=document.querySelectorAll('#corresponsables li').length>0;document.getElementById('subVacio').hidden=document.querySelectorAll('#subtemasBody tr').length>0;}"
	j = j & "function agendaExtAgregarCorresponsable(){"
	j = j & "var p=document.getElementById('corPersona'),n=document.getElementById('corNombre'),e=document.getElementById('corEmail'),c=document.getElementById('corCopia');"
	j = j & "if(!p.value&&!n.value.trim()){alert('Eleg\u00ed una persona del sistema o complet\u00e1 el nombre del contacto externo.');p.focus();return;}"
	j = j & "if(!p.value&&!e.value.trim()&&c.checked){alert('Para enviar copia a un contacto externo, complet\u00e1 su mail.');e.focus();return;}"
	j = j & "corN++;document.getElementById('corresponsable_count').value=corN;"
	j = j & "var li=document.createElement('li');li.className='ui-copia';var s=document.createElement('span');s.className='ui-participante';"
	j = j & "s.textContent=p.value?p.options[p.selectedIndex].text:n.value.trim();"
	j = j & "var mail=e.value.trim();if(p.value&&!mail){mail=(p.options[p.selectedIndex].getAttribute('data-mail')||'').trim();}if(mail){var m=document.createElement('span');m.className='ui-participante__mail';m.textContent=' '+mail;s.appendChild(m);}"
	j = j & "if(c.checked){var b=document.createElement('span');b.className='ui-badge ui-badge--blue';b.textContent='Copia';s.appendChild(document.createTextNode(' '));s.appendChild(b);}"
	j = j & "li.appendChild(s);li.appendChild(agendaExtOculto('cor_persona_'+corN,p.value));li.appendChild(agendaExtOculto('cor_nombre_'+corN,p.value?'':n.value.trim()));"
	j = j & "li.appendChild(agendaExtOculto('cor_email_'+corN,e.value.trim()));li.appendChild(agendaExtOculto('cor_copia_'+corN,c.checked?'1':'0'));"
	j = j & "var q=document.createElement('button');q.type='button';q.className='ui-button ui-button--icon ui-button--danger';q.setAttribute('aria-label','Quitar a '+s.firstChild.textContent);q.innerHTML='&times;';"
	j = j & "q.onclick=function(){li.parentNode.removeChild(li);agendaExtActualizarVacios();};li.appendChild(q);"
	j = j & "document.getElementById('corresponsables').appendChild(li);p.value='';n.value='';e.value='';agendaExtActualizarVacios();}"
	j = j & "function agendaExtAgregarSubtema(){subN++;document.getElementById('subtema_count').value=subN;"
	j = j & "var t=document.getElementById('agendaTemaTemplate').innerHTML.split('__N__').join(subN);var cuerpo=document.getElementById('subtemasBody');"
	j = j & "var tmp=document.createElement('tbody');tmp.innerHTML=t;var tr=tmp.firstElementChild;cuerpo.appendChild(tr);agendaExtCambiarRelacion(tr.querySelector('select'));agendaExtActualizarVacios();tr.querySelector('input[name^=sub_tema_]').focus();}"
	j = j & "function agendaExtQuitarFila(boton){var tr=boton.closest('tr');tr.parentNode.removeChild(tr);agendaExtActualizarVacios();}"
	j = j & "function agendaExtCambiarRelacion(sel){var box=sel.closest('tr').querySelector('.relation-fields'),input=box.querySelector('.relation-id');var visible=sel.value==='1'||sel.value==='2';"
	j = j & "box.hidden=!visible;input.setAttribute('aria-label',sel.value==='2'?'N\u00ba de incidente':'N\u00ba de proyecto');if(!visible){input.value='';}}"
	j = j & "function agendaExtBuscarEntidad(boton){var input=boton.parentNode.querySelector('.relation-id'),tipo=boton.closest('tr').querySelector('select[name^=sub_tipo_]').value,tabla=tipo==='1'?'PROYECTO':'INCIDENTE';"
	j = j & "window.open('buscar_entidades.asp?tabla='+tabla+'&campo='+encodeURIComponent(input.name),'buscar_entidad','width=900,height=650,resizable=yes,scrollbars=yes');}"
	j = j & "function agendaExtSeleccionarEntidad(campo,id,nombre){var input=document.querySelector('input[name='+JSON.stringify(campo)+']');if(input){input.value=id;input.title=nombre;}}"
	j = j & "function agendaExtValidar(){var a=document.getElementById('hora_inicio').value,b=document.getElementById('hora_fin').value;if(a&&b&&b<=a){alert('La hora de fin tiene que ser posterior a la de inicio.');document.getElementById('hora_fin').focus();return false;}return true;}"
	j = j & "document.querySelectorAll('#subtemasBody select[name^=sub_tipo_]').forEach(agendaExtCambiarRelacion);"
	AgendaExtJs = j
End Function

'[UI-ANT] ' ===== Archivo anterior completo (documento propio con CSS inline y participantes por numero). Se borra entero buscando [UI-ANT] =====
'[UI-ANT] Function AgendaExtDebeManejar(opc)
	'[UI-ANT] AgendaExtDebeManejar = (opc = "consulta" Or opc = "cargar" Or opc = "det" Or opc = "agenda_ext_guardar" Or opc = "agenda_ext_postergar" Or opc = "agenda_ext_seguimiento" Or opc = "agenda_ext_eliminar_corresp")
'[UI-ANT] End Function

'[UI-ANT] Sub AgendaExtRender(cn, opc, xmenu, xtmenu)
	'[UI-ANT] Dim id, rs, avisosCruce, mensajeGuardado
	'[UI-ANT] If opc = "consulta" Then
		'[UI-ANT] AgendaExtConsulta cn, xmenu, xtmenu
		'[UI-ANT] Exit Sub
	'[UI-ANT] End If
	'[UI-ANT] id = Request.QueryString("id")
	'[UI-ANT] If opc = "agenda_ext_guardar" Then
		'[UI-ANT] avisosCruce = ""
		'[UI-ANT] id = AgendaExtGuardar(cn, id, avisosCruce)
		'[UI-ANT] mensajeGuardado = "Agenda guardada correctamente."
		'[UI-ANT] If Trim(avisosCruce) <> "" Then
			'[UI-ANT] mensajeGuardado = mensajeGuardado & vbCrLf & vbCrLf & "ATENCION: no se agregaron los siguientes participantes por tener otra reunion cruzada en ese horario:" & vbCrLf & avisosCruce
		'[UI-ANT] End If
		'[UI-ANT] AgendaExtRedir mensajeGuardado, id, xmenu, xtmenu
		'[UI-ANT] Exit Sub
	'[UI-ANT] End If
	'[UI-ANT] If opc = "agenda_ext_postergar" Then
		'[UI-ANT] id = AgendaExtPostergar(cn, id)
		'[UI-ANT] AgendaExtRedir "La reunion fue postergada.", id, xmenu, xtmenu
		'[UI-ANT] Exit Sub
	'[UI-ANT] End If
	'[UI-ANT] If opc = "agenda_ext_seguimiento" Then
		'[UI-ANT] id = AgendaExtSeguimiento(cn, id)
		'[UI-ANT] AgendaExtRedir "El seguimiento fue generado.", id, xmenu, xtmenu
		'[UI-ANT] Exit Sub
	'[UI-ANT] End If
	'[UI-ANT] If opc = "agenda_ext_eliminar_corresp" Then
		'[UI-ANT] AgendaExtEliminarCorresp cn, xmenu, xtmenu
		'[UI-ANT] Exit Sub
	'[UI-ANT] End If
	'[UI-ANT] If AgendaExtEsIdValido(id) Then
		'[UI-ANT] Set rs = Exec(cn, "sp_cg_AgendaExt @Modo='obtener', @Id=" & CLng(id))
	'[UI-ANT] Else
		'[UI-ANT] Set rs = Nothing
	'[UI-ANT] End If
	'[UI-ANT] AgendaExtPagina cn, rs, id, xmenu, xtmenu
'[UI-ANT] End Sub

'[UI-ANT] Sub AgendaExtConsulta(cn, xmenu, xtmenu)
	'[UI-ANT] Dim rs, mover, mostrar, pos, volver, editar, estado, responsable, fecha, padreLista, padreListaUrl
	'[UI-ANT] mostrar = 15
	'[UI-ANT] mover = AgendaExtSqlInt(Request.QueryString("CantRegMover"), 0)
	'[UI-ANT] If mover < 0 Then mover = 0
	'[UI-ANT] Set rs = Exec(cn, "sp_cg_AgendaExt @Modo='consulta', @CantRegMover=" & mover & ", @CantRegMostrar=" & mostrar)
	'[UI-ANT] volver = "agenda.asp?xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu)
	'[UI-ANT] Response.Write "<!doctype html><html lang='es'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width, initial-scale=1'><title>Agenda | Cor Gestion</title><style>" & AgendaExtCss() & "</style></head><body>"
	'[UI-ANT] Response.Write "<header class='app-header'><div class='brand'><div class='brand-mark'>CG</div><div><strong>Cor Gestion</strong><small>Sistema Administrativo</small></div></div><a class='btn btn-secondary' href='" & volver & "'>Volver</a></header><main class='shell'><div class='breadcrumb'>Agenda / Reuniones operativas</div><div class='page-head'><div><h1>Agenda</h1><p>Consulta y administra reuniones, seguimientos y temas relacionados.</p></div><div class='toolbar'><a class='btn btn-primary' href='agenda.asp?opc=edicion&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu) & "'>+ Cargar reunion</a><button type='button' class='btn btn-secondary' onclick=""window.open('agenda_semana.asp?xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu) & "','agenda_semana','width=1520,height=900,resizable=yes,scrollbars=yes')"">Vista semanal</button></div></div><section class='panel'><div class='panel-head'><div><h2>Reuniones agendadas</h2><span class='muted'>Agenda operativa</span></div></div><div class='panel-body table-wrap'><table><thead><tr><th>ID agenda</th><th>Asunto</th><th>Fecha</th><th>Horario</th><th>Responsable</th><th>Estado</th><th>Registro padre</th><th>Relaciones</th><th></th></tr></thead><tbody>"
	'[UI-ANT] pos = 0
	'[UI-ANT] Do While Not rs.EOF And pos < mostrar
		'[UI-ANT] fecha = AgendaExtHtmlDate(AgendaExtVal(rs, "FECHA_REUNION", ""))
		'[UI-ANT] estado = AgendaExtVal(rs, "ESTADO", "PENDIENTE")
		'[UI-ANT] responsable = AgendaExtVal(rs, "RESPONSABLE", "Sin responsable")
		'[UI-ANT] editar = "agenda.asp?opc=edicion&id=" & rs("ID") & "&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu)
		'[UI-ANT] padreLista = AgendaExtVal(rs, "ID_AGENDA_PADRE", "")
		'[UI-ANT] padreListaUrl = "agenda.asp?opc=det&id=" & padreLista & "&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu)
		'[UI-ANT] If AgendaExtEsIdValido(padreLista) Then padreLista = "<a href='" & padreListaUrl & "'>Agenda #" & Server.HTMLEncode(padreLista & "") & "</a>" Else padreLista = "-"
		'[UI-ANT] Response.Write "<tr><td><a href='" & editar & "'>#" & rs("ID") & "</a></td><td>" & Server.HTMLEncode(Left(AgendaExtVal(rs, "ASUNTO", ""), 80)) & "</td><td>" & fecha & "</td><td>" & AgendaExtHtmlTime(AgendaExtVal(rs, "HORA_INICIO", "")) & " - " & AgendaExtHtmlTime(AgendaExtVal(rs, "HORA_FIN", "")) & "</td><td>" & Server.HTMLEncode(responsable) & "</td><td><span class='status'>" & Server.HTMLEncode(estado) & "</span></td><td>" & padreLista & "</td><td>" & rs("CANT_CORRESP") & " corresponsable(s), " & rs("CANT_SUBTEMAS") & " subtema(s)</td><td><a href='" & editar & "'>Ver / editar</a></td></tr>"
		'[UI-ANT] pos = pos + 1
		'[UI-ANT] rs.MoveNext
	'[UI-ANT] Loop
	'[UI-ANT] Response.Write "</tbody></table></div></section><div class='pagination'>"
	'[UI-ANT] If mover > 0 Then Response.Write "<a href='agenda.asp?opc=consulta&CantRegMover=" & (mover - mostrar) & "&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu) & "'>Anterior</a>"
	'[UI-ANT] If Not rs.EOF Then Response.Write "<a href='agenda.asp?opc=consulta&CantRegMover=" & (mover + mostrar) & "&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu) & "'>Siguiente</a>"
	'[UI-ANT] Response.Write "</div></main></body></html>"
	'[UI-ANT] rs.Close
'[UI-ANT] End Sub

'[UI-ANT] Function AgendaExtGuardar(cn, idActual, ByRef avisosCruceOut)
	'[UI-ANT] Dim q, rs, id, i, n, persona, nombre, email, copia, tipo, relacionId, proyecto, incidente, tema, conclusion, accion
	'[UI-ANT] Dim fechaReunionSql, horaInicioSql, horaFinSql, avisosCruce, descripcionCruce
	'[UI-ANT] fechaReunionSql = AgendaExtSqlDate(Request.Form("fecha_reunion"))
	'[UI-ANT] horaInicioSql = AgendaExtSqlTime(Request.Form("hora_inicio"))
	'[UI-ANT] horaFinSql = AgendaExtSqlTime(Request.Form("hora_fin"))
	'[UI-ANT] avisosCruce = ""
	'[UI-ANT] q = "sp_cg_AgendaExt @Modo='insertar'"
	'[UI-ANT] If AgendaExtEsIdValido(idActual) Then q = "sp_cg_AgendaExt @Modo='actualizar', @Id=" & CLng(idActual)
	'[UI-ANT] q = q & ", @Asunto=" & AgendaExtSqlText(Request.Form("asunto"))
	'[UI-ANT] q = q & ", @FechaReunion=" & fechaReunionSql
	'[UI-ANT] q = q & ", @HoraInicio=" & horaInicioSql
	'[UI-ANT] q = q & ", @HoraFin=" & horaFinSql
	'[UI-ANT] q = q & ", @FechaAviso=" & AgendaExtSqlDate(Request.Form("fecha_aviso"))
	'[UI-ANT] q = q & ", @IdRespon=" & AgendaExtSqlInt(Request.Form("id_respon"), Session("usuario"))
	'[UI-ANT] q = q & ", @IdHerramienta=" & AgendaExtSqlNullableInt(Request.Form("id_herramienta"))
	'[UI-ANT] q = q & ", @Mensaje=" & AgendaExtSqlText(Request.Form("mensaje"))
	'[UI-ANT] q = q & ", @Conclusion=" & AgendaExtSqlText(Request.Form("conclusion"))
	'[UI-ANT] q = q & ", @Estado=" & AgendaExtSqlText(Request.Form("estado"))
	'[UI-ANT] If AgendaExtEsIdValido(idActual) Then
		'[UI-ANT] ExecNonQuery cn, q
		'[UI-ANT] id = CLng(idActual)
	'[UI-ANT] Else
		'[UI-ANT] Set rs = Exec(cn, q & ", @TipoOrigen='N'")
		'[UI-ANT] id = CLng(rs("ID"))
	'[UI-ANT] End If

	'[UI-ANT] If AgendaExtEsIdValido(idActual) Then
		'[UI-ANT] ExecNonQuery cn, "sp_cg_AgendaExt @Modo='corresp_limpiar', @Id=" & id
		'[UI-ANT] ExecNonQuery cn, "sp_cg_AgendaExt @Modo='subtema_limpiar', @Id=" & id
	'[UI-ANT] End If
	'[UI-ANT] n = AgendaExtSqlInt(Request.Form("corresponsable_count"), 0)
		'[UI-ANT] For i = 1 To n
			'[UI-ANT] persona = Request.Form("cor_persona_" & i)
			'[UI-ANT] nombre = Request.Form("cor_nombre_" & i)
			'[UI-ANT] email = Request.Form("cor_email_" & i)
			'[UI-ANT] copia = AgendaExtSqlBit(Request.Form("cor_copia_" & i))
			'[UI-ANT] If IsNumeric(persona) Then
				'[UI-ANT] If CLng(persona) > 0 Then
					' Antes de sumar un participante del sistema, se valida si tiene
					' otra reunion cruzada en el mismo horario (ya sea como
					' responsable o como corresponsable de otra agenda).
					'[UI-ANT] descripcionCruce = AgendaExtDescribirCruceParticipante(cn, id, CLng(persona), fechaReunionSql, horaInicioSql, horaFinSql)
					'[UI-ANT] If descripcionCruce <> "" Then
						'[UI-ANT] avisosCruce = avisosCruce & "- " & AgendaExtNombrePersona(cn, CLng(persona)) & ": " & descripcionCruce & vbCrLf
					'[UI-ANT] Else
						'[UI-ANT] ExecNonQuery cn, "sp_cg_AgendaExt @Modo='corresp_agregar', @Id=" & id & ", @IdPersona=" & CLng(persona) & ", @EnviarCopia=" & copia
					'[UI-ANT] End If
				'[UI-ANT] ElseIf Trim(nombre) <> "" Then
					'[UI-ANT] ExecNonQuery cn, "sp_cg_AgendaExt @Modo='corresp_agregar', @Id=" & id & ", @NombreManual=" & AgendaExtSqlText(nombre) & ", @EmailManual=" & AgendaExtSqlText(email) & ", @EnviarCopia=" & copia
				'[UI-ANT] End If
			'[UI-ANT] ElseIf Trim(nombre) <> "" Then
				'[UI-ANT] ExecNonQuery cn, "sp_cg_AgendaExt @Modo='corresp_agregar', @Id=" & id & ", @NombreManual=" & AgendaExtSqlText(nombre) & ", @EmailManual=" & AgendaExtSqlText(email) & ", @EnviarCopia=" & copia
			'[UI-ANT] End If
		'[UI-ANT] Next

		'[UI-ANT] n = AgendaExtSqlInt(Request.Form("subtema_count"), 0)
		'[UI-ANT] For i = 1 To n
			'[UI-ANT] tipo = AgendaExtSqlInt(Request.Form("sub_tipo_" & i), 0)
			'[UI-ANT] relacionId = AgendaExtSqlInt(Request.Form("sub_relacion_id_" & i), 0)
			'[UI-ANT] proyecto = 0
			'[UI-ANT] incidente = 0
			'[UI-ANT] If tipo = 1 Then proyecto = relacionId
			'[UI-ANT] If tipo = 2 Then incidente = relacionId
			'[UI-ANT] tema = Request.Form("sub_tema_" & i)
			'[UI-ANT] conclusion = Request.Form("sub_conclusion_" & i)
			'[UI-ANT] accion = AgendaExtSqlInt(Request.Form("sub_accion_" & i), 0)
			'[UI-ANT] ExecNonQuery cn, "sp_cg_AgendaExt @Modo='subtema_agregar', @Id=" & id & ", @IdTipoRelacion=" & tipo & ", @IdProyecto=" & AgendaExtSqlNullableInt(proyecto) & ", @IdIncidente=" & AgendaExtSqlNullableInt(incidente) & ", @Tema=" & AgendaExtSqlText(tema) & ", @Conclusion=" & AgendaExtSqlText(conclusion) & ", @TiempoUsado=" & AgendaExtSqlInt(Request.Form("sub_tiempo_" & i), 0) & ", @Orden=" & i & ", @Seguimiento=" & AgendaExtSqlBit(Request.Form("sub_seguimiento_" & i)) & ", @IdAccion=" & AgendaExtSqlNullableInt(accion)
		'[UI-ANT] Next
	'[UI-ANT] avisosCruceOut = avisosCruce
	'[UI-ANT] AgendaExtGuardar = id
'[UI-ANT] End Function

' Chequea si @IdPersona tiene otra reunion (PENDIENTE/POSTERGADA) cuyo
' horario se solape con [horaInicioSql, horaFinSql) el dia fechaReunionSql,
' ya sea como responsable de esa reunion o como corresponsable de ella.
' idAgendaActual se excluye del chequeo (es la reunion que se esta guardando).
' Devuelve "" si no hay cruce, o un texto con el detalle de los cruces.
'[UI-ANT] Function AgendaExtDescribirCruceParticipante(cn, idAgendaActual, idPersona, fechaReunionSql, horaInicioSql, horaFinSql)
	'[UI-ANT] Dim rsCruce, detalle
	'[UI-ANT] AgendaExtDescribirCruceParticipante = ""
	'[UI-ANT] If fechaReunionSql = "NULL" Or horaInicioSql = "NULL" Or horaFinSql = "NULL" Then Exit Function
	'[UI-ANT] Set rsCruce = Exec(cn, "sp_cg_Agenda_ValidarCruceHorario @IdResponsable=" & idPersona & ", @FechaReunion=" & fechaReunionSql & ", @HoraInicio=" & horaInicioSql & ", @HoraFin=" & horaFinSql & ", @IdAgendaExcluir=" & idAgendaActual)
	'[UI-ANT] detalle = ""
	'[UI-ANT] Do While Not rsCruce.EOF
		'[UI-ANT] If detalle <> "" Then detalle = detalle & "; "
		'[UI-ANT] detalle = detalle & """" & AgendaExtVal(rsCruce, "ASUNTO", "(sin asunto)") & """ (#" & rsCruce("ID_AGENDA_CRUCE") & ") " & AgendaExtHtmlTime(rsCruce("HORA_INICIO")) & "-" & AgendaExtHtmlTime(rsCruce("HORA_FIN"))
		'[UI-ANT] rsCruce.MoveNext
	'[UI-ANT] Loop
	'[UI-ANT] rsCruce.Close
	'[UI-ANT] AgendaExtDescribirCruceParticipante = detalle
'[UI-ANT] End Function

'[UI-ANT] Function AgendaExtNombrePersona(cn, idPersona)
	'[UI-ANT] Dim rsPersona, nombre
	'[UI-ANT] nombre = "Persona #" & idPersona
	'[UI-ANT] Set rsPersona = Exec(cn, "SELECT RTRIM(ISNULL(NOMBRE,'') + ' ' + ISNULL(APELLIDO,'')) AS NOMBRE FROM dbo.CG_COR_RESPONSABLES WHERE ID=" & idPersona)
	'[UI-ANT] If Not rsPersona.EOF Then
		'[UI-ANT] If Trim(AgendaExtVal(rsPersona, "NOMBRE", "")) <> "" Then nombre = Trim(rsPersona("NOMBRE"))
	'[UI-ANT] End If
	'[UI-ANT] rsPersona.Close
	'[UI-ANT] AgendaExtNombrePersona = nombre
'[UI-ANT] End Function

'[UI-ANT] Function AgendaExtPostergar(cn, id)
	'[UI-ANT] Dim rs
	'[UI-ANT] Set rs = Exec(cn, "sp_cg_AgendaExt @Modo='postergar', @Id=" & CLng(id) & ", @FechaReunion=" & AgendaExtSqlDate(Request.Form("fecha_reunion")) & ", @HoraInicio=" & AgendaExtSqlTime(Request.Form("hora_inicio")) & ", @HoraFin=" & AgendaExtSqlTime(Request.Form("hora_fin")))
	'[UI-ANT] AgendaExtPostergar = CLng(rs("ID"))
'[UI-ANT] End Function

'[UI-ANT] Function AgendaExtSeguimiento(cn, id)
	'[UI-ANT] Dim rs
	'[UI-ANT] Set rs = Exec(cn, "sp_cg_AgendaExt @Modo='seguimiento', @Id=" & CLng(id) & ", @FechaReunion=" & AgendaExtSqlDate(Request.Form("fecha_reunion")) & ", @HoraInicio=" & AgendaExtSqlTime(Request.Form("hora_inicio")) & ", @HoraFin=" & AgendaExtSqlTime(Request.Form("hora_fin")) & ", @Conclusion=" & AgendaExtSqlText(Request.Form("conclusion")))
	'[UI-ANT] AgendaExtSeguimiento = CLng(rs("ID"))
'[UI-ANT] End Function

'[UI-ANT] Sub AgendaExtEliminarCorresp(cn, xmenu, xtmenu)
	'[UI-ANT] Dim id, correspId
	'[UI-ANT] id = Request.QueryString("id")
	'[UI-ANT] correspId = Request.QueryString("corresp_id")
	'[UI-ANT] If AgendaExtEsIdValido(id) And AgendaExtEsIdValido(correspId) Then
		'[UI-ANT] ExecNonQuery cn, "sp_cg_AgendaExt @Modo='corresp_eliminar', @Id=" & CLng(id) & ", @IdCorresponsable=" & CLng(correspId)
	'[UI-ANT] End If
	'[UI-ANT] Response.Redirect "agenda.asp?opc=edicion&id=" & id & "&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu)
'[UI-ANT] End Sub

'[UI-ANT] Sub AgendaExtPagina(cn, rs, id, xmenu, xtmenu)
	'[UI-ANT] Dim editar, asunto, fecha, aviso, hi, hf, mensaje, conclusion, resp, herramienta, estado, agendaId, agendaPadreId, volver, padreUrl
	'[UI-ANT] editar = False
	'[UI-ANT] If IsObject(rs) Then
		'[UI-ANT] If Not rs Is Nothing Then
			'[UI-ANT] If Not rs.EOF Then editar = True
		'[UI-ANT] End If
	'[UI-ANT] End If
	'[UI-ANT] If editar Then
		'[UI-ANT] agendaId = AgendaExtVal(rs, "ID", id)
		'[UI-ANT] agendaPadreId = AgendaExtVal(rs, "ID_AGENDA_PADRE", "")
		'[UI-ANT] asunto = AgendaExtVal(rs, "ASUNTO", "") : fecha = AgendaExtVal(rs, "FECHA_REUNION", "") : aviso = AgendaExtVal(rs, "FECHA_AVISO", "")
		'[UI-ANT] hi = AgendaExtVal(rs, "HORA_INICIO", "") : hf = AgendaExtVal(rs, "HORA_FIN", "") : mensaje = AgendaExtVal(rs, "MENSAJE", "")
		'[UI-ANT] conclusion = AgendaExtVal(rs, "CONCLUSION", "") : resp = AgendaExtVal(rs, "ID_RESPON", Session("usuario")) : herramienta = AgendaExtVal(rs, "ID_HERRAMIENTA", 0) : estado = AgendaExtVal(rs, "ESTADO", "PENDIENTE")
	'[UI-ANT] Else
		'[UI-ANT] agendaId = "" : agendaPadreId = ""
		'[UI-ANT] asunto = "" : fecha = AgendaExtHoy() : aviso = AgendaExtHoy() : hi = 900 : hf = 1000 : mensaje = "" : conclusion = "" : resp = Session("usuario") : herramienta = 0 : estado = "PENDIENTE"
	'[UI-ANT] End If
	'[UI-ANT] padreUrl = ""
	'[UI-ANT] If AgendaExtEsIdValido(agendaPadreId) Then padreUrl = "agenda.asp?opc=det&id=" & CLng(agendaPadreId) & "&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu)
	'[UI-ANT] Dim tituloPagina
	'[UI-ANT] If editar Then tituloPagina = "Editar reunion" Else tituloPagina = "Nueva reunion"
	'[UI-ANT] volver = "agenda.asp?xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu)
	'[UI-ANT] Response.Write "<!doctype html><html lang='es'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width, initial-scale=1'><title>Agenda | Cor Gestion</title><style>" & AgendaExtCss() & "</style></head><body><header class='app-header'><div class='brand'><div class='brand-mark'>CG</div><div><strong>Cor Gestion</strong><small>Sistema Administrativo</small></div></div><a class='btn btn-secondary' href='" & volver & "'>Volver al listado</a></header><main class='shell'><div class='breadcrumb'>Agenda / Reuniones operativas</div><div class='page-head'><div><h1>" & tituloPagina & "</h1><p>Cabecera, participantes y temas relacionados.</p></div></div>"
	'[UI-ANT] Dim guardarAction
	'[UI-ANT] guardarAction = "agenda.asp?opc=agenda_ext_guardar"
	'[UI-ANT] If editar Then guardarAction = guardarAction & "&id=" & id
	'[UI-ANT] guardarAction = guardarAction & "&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu)
	'[UI-ANT] Response.Write "<form id='agendaExt' method='post' action='" & guardarAction & "' onsubmit='return agendaExtValidar();'><section class='panel'><div class='panel-head'><div><h2>Datos generales</h2><span class='muted'>Cabecera de AGENDA</span></div><div class='record-meta'><span>ID agenda: <strong>" & Server.HTMLEncode(agendaId & "") & "</strong></span>"
	'[UI-ANT] If padreUrl <> "" Then Response.Write "<span>Registro padre: <a href='" & padreUrl & "'>Agenda #" & Server.HTMLEncode(agendaPadreId & "") & "</a></span>"
	'[UI-ANT] Response.Write "</div></div><div class='panel-body'><div class='form-grid'>"
	'[UI-ANT] AgendaExtInput "Asunto", "asunto", asunto, "text", True
	'[UI-ANT] AgendaExtSelectResponsables cn, resp
	'[UI-ANT] AgendaExtInput "Fecha reunion", "fecha_reunion", AgendaExtHtmlDate(fecha), "date", True
	'[UI-ANT] AgendaExtInput "Fecha aviso", "fecha_aviso", AgendaExtHtmlDate(aviso), "date", True
	'[UI-ANT] AgendaExtInput "Hora inicio", "hora_inicio", AgendaExtHtmlTime(hi), "time", True
	'[UI-ANT] AgendaExtInput "Hora fin", "hora_fin", AgendaExtHtmlTime(hf), "time", True
	'[UI-ANT] AgendaExtSelectHerramientas cn, herramienta
	'[UI-ANT] AgendaExtSelectEstados cn, estado
	'[UI-ANT] Response.Write "<label class='full'>Mensaje<textarea name='mensaje' rows='3'>" & Server.HTMLEncode(mensaje & "") & "</textarea></label></div></div></section>"
	'[UI-ANT] AgendaExtCorresponsables cn, id, xmenu, xtmenu
	'[UI-ANT] AgendaExtSubtemas cn, id, xmenu, xtmenu
	'[UI-ANT] Response.Write "<section class='panel'><div class='panel-body'><label class='full'>Conclusion general de la reunion<textarea name='conclusion' rows='4'>" & Server.HTMLEncode(conclusion & "") & "</textarea></label></div></section><div class='action-bar'><a class='btn' href='" & volver & "'>Cancelar</a><div class='right'><button class='btn btn-primary' type='submit'>Guardar reunion</button>"
	'[UI-ANT] If editar Then Response.Write "<button type='submit' formaction='agenda.asp?opc=agenda_ext_postergar&id=" & id & "&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu) & "'>Postergar</button><button type='submit' formaction='agenda.asp?opc=agenda_ext_seguimiento&id=" & id & "&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu) & "'>Generar seguimiento</button>"
	'[UI-ANT] Response.Write "</div></div></form></main><script>" & AgendaExtJs() & AgendaExtRelacionJs() & "</script></body></html>"
'[UI-ANT] End Sub

'[UI-ANT] Sub AgendaExtCorresponsables(cn, id, xmenu, xtmenu)
	'[UI-ANT] Dim rs, detalle, nombre, persona, manual, email, copia, n, correspId, chipText, eliminarUrl
	'[UI-ANT] Response.Write "<section class='panel'><h2>Participantes (corresponsables)</h2><div class='grid'><label>Persona del sistema<select id='corPersona'><option value=''>Seleccionar</option>"
	'[UI-ANT] Set rs = Exec(cn, "sp_cg_Responsables @Modo='listar_activos_sin', @ExcluirId=" & Session("usuario"))
	'[UI-ANT] Do While Not rs.EOF
		'[UI-ANT] Response.Write "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>" : rs.MoveNext
	'[UI-ANT] Loop
	'[UI-ANT] Response.Write "</select></label><label>Nombre manual<input id='corNombre' type='text'></label><label>Email manual<input id='corEmail' type='email'></label><label>Enviar copia<select id='corCopia'><option value='1'>SI</option><option value='0'>NO</option></select></label></div><button type='button' onclick='agendaExtAgregarCorresponsable()'>Agregar corresponsable</button><div id='corresponsables'>"
	'[UI-ANT] n = 0
	'[UI-ANT] If AgendaExtEsIdValido(id) Then
		'[UI-ANT] Set detalle = Exec(cn, "sp_cg_AgendaExt @Modo='corresp_listar', @Id=" & CLng(id))
		'[UI-ANT] Do While Not detalle.EOF
			'[UI-ANT] n = n + 1
			'[UI-ANT] correspId = AgendaExtVal(detalle, "ID_AGENDA_CORRESP", "")
			'[UI-ANT] persona = AgendaExtVal(detalle, "ID_PERSONA", "")
			'[UI-ANT] manual = AgendaExtVal(detalle, "NOMBRE_MANUAL", "")
			'[UI-ANT] email = AgendaExtVal(detalle, "EMAIL_MANUAL", "")
			'[UI-ANT] copia = AgendaExtVal(detalle, "ENVIAR_COPIA", 1)
			'[UI-ANT] nombre = AgendaExtVal(detalle, "NOMBRE_PERSONA", manual)
			'[UI-ANT] chipText = Server.HTMLEncode(nombre & "")
			'[UI-ANT] If Trim(email & "") <> "" Then chipText = chipText & " &lt;" & Server.HTMLEncode(email & "") & "&gt;"
			'[UI-ANT] If copia = 1 Then chipText = chipText & " <small>(copia)</small>"
			'[UI-ANT] eliminarUrl = "agenda.asp?opc=agenda_ext_eliminar_corresp&id=" & id & "&corresp_id=" & correspId & "&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu)
			'[UI-ANT] Response.Write "<div class='chip'>" & chipText & "<input type='hidden' name='cor_persona_" & n & "' value='" & Server.HTMLEncode(persona & "") & "'><input type='hidden' name='cor_nombre_" & n & "' value='" & Server.HTMLEncode(manual & "") & "'><input type='hidden' name='cor_email_" & n & "' value='" & Server.HTMLEncode(email & "") & "'><input type='hidden' name='cor_copia_" & n & "' value='" & Server.HTMLEncode(copia & "") & "'><a href='" & eliminarUrl & "' class='chip-remove' title='Eliminar participante' onclick=""return confirm('?Eliminar participante?');"">x</a></div>"
			'[UI-ANT] detalle.MoveNext
		'[UI-ANT] Loop
	'[UI-ANT] End If
	'[UI-ANT] Response.Write "</div><input type='hidden' id='corresponsable_count' name='corresponsable_count' value='" & n & "'></section>"
'[UI-ANT] End Sub

'[UI-ANT] Sub AgendaExtSubtemas(cn, id, xmenu, xtmenu)
	'[UI-ANT] Dim detalle, n, tipo, proyecto, incidente, relacionId, tema, conclusion, minutos, seguimiento, accion, opcionesAccion
	'[UI-ANT] opcionesAccion = AgendaExtOpcionesAccion(cn, 0)
	'[UI-ANT] Response.Write "<section class='panel'><h2>Temario y subtemas</h2><table id='subtemas'><thead><tr><th>Tipo</th><th>Proyecto / incidente</th><th>Tema a tratar</th><th>Conclusion</th><th>Min.</th><th>Seguimiento</th><th>Accion</th><th></th></tr></thead><tbody>"
	'[UI-ANT] n = 0
	'[UI-ANT] If AgendaExtEsIdValido(id) Then
		'[UI-ANT] Set detalle = Exec(cn, "sp_cg_AgendaExt @Modo='subtema_listar', @Id=" & CLng(id))
		'[UI-ANT] Do While Not detalle.EOF
			'[UI-ANT] n = n + 1 : tipo = AgendaExtVal(detalle, "ID_TIPO_RELACION", 0) : proyecto = AgendaExtVal(detalle, "ID_PROYECTO", "") : incidente = AgendaExtVal(detalle, "ID_INCIDENTE", "") : tema = AgendaExtVal(detalle, "TEMA_A_TRATAR", "") : conclusion = AgendaExtVal(detalle, "CONCLUSION", "") : minutos = AgendaExtVal(detalle, "TIEMPO_USADO", 0) : seguimiento = AgendaExtVal(detalle, "SEGUIMIENTO", 0) : accion = AgendaExtVal(detalle, "ID_ACCION", "")
			'[UI-ANT] opcionesAccion = AgendaExtOpcionesAccion(cn, accion)
			'[UI-ANT] relacionId = ""
			'[UI-ANT] If tipo = 1 Then relacionId = proyecto
			'[UI-ANT] If tipo = 2 Then relacionId = incidente
			'[UI-ANT] Response.Write "<tr><td><select name='sub_tipo_" & n & "' onchange='agendaExtCambiarRelacion(this)'><option value='1'" & AgendaExtSelected(1, tipo) & ">PROYECTO</option><option value='2'" & AgendaExtSelected(2, tipo) & ">INCIDENTE</option><option value='3'" & AgendaExtSelected(3, tipo) & ">CONSULTORIA</option><option value='4'" & AgendaExtSelected(4, tipo) & ">OTRO</option></select></td><td><div class='relation-fields' style='display:block'><label class='relation-id-label'>ID proyecto/incidente<input class='relation-id' name='sub_relacion_id_" & n & "' type='number' min='1' value='" & Server.HTMLEncode(relacionId & "") & "' placeholder='ID'><button type='button' class='btn-buscar-entidad' onclick='agendaExtBuscarEntidad(this)'>Buscar</button></label></div></td><td><input name='sub_tema_" & n & "' value='" & Server.HTMLEncode(tema & "") & "' required maxlength='500'></td><td><textarea name='sub_conclusion_" & n & "' rows='2' maxlength='4000'>" & Server.HTMLEncode(conclusion & "") & "</textarea></td><td><input name='sub_tiempo_" & n & "' type='number' min='0' value='" & minutos & "'></td><td><input name='sub_seguimiento_" & n & "' type='checkbox' value='1'" & AgendaExtChecked(seguimiento) & "></td><td><select name='sub_accion_" & n & "'>" & opcionesAccion & "</select></td><td><button type='button' onclick='this.parentNode.parentNode.remove()'>Quitar</button></td></tr>"
			'[UI-ANT] detalle.MoveNext
		'[UI-ANT] Loop
	'[UI-ANT] End If
	'[UI-ANT] Response.Write "</tbody></table><button type='button' onclick='agendaExtAgregarSubtema()'>Agregar subtema</button><input type='hidden' id='subtema_count' name='subtema_count' value='" & n & "'><select id='agendaAccionTemplate' style='display:none'>" & AgendaExtOpcionesAccion(cn, 0) & "</select></section>"
'[UI-ANT] End Sub

'[UI-ANT] Sub AgendaExtSelectResponsables(cn, resp)
	'[UI-ANT] Dim rs
	'[UI-ANT] Response.Write "<label>Responsable<select name='id_respon'>"
	'[UI-ANT] Set rs = Exec(cn, "sp_cg_Responsables @Modo='listar_activos_sin', @ExcluirId=0")
	'[UI-ANT] Do While Not rs.EOF
		'[UI-ANT] Response.Write "<option value='" & rs(0) & "'" & AgendaExtSelected(rs(0), resp) & ">" & Server.HTMLEncode(rs(1) & "") & "</option>" : rs.MoveNext
	'[UI-ANT] Loop
	'[UI-ANT] Response.Write "</select></label>"
'[UI-ANT] End Sub

'[UI-ANT] Sub AgendaExtSelectHerramientas(cn, sel)
	'[UI-ANT] Dim rs
	'[UI-ANT] Response.Write "<label>Herramienta<select name='id_herramienta'><option value='0'>Ninguna</option>"
	'[UI-ANT] Set rs = Exec(cn, "sp_cg_Catalogo_Listar @Catalogo='aplicativos'")
	'[UI-ANT] Do While Not rs.EOF
		'[UI-ANT] Response.Write "<option value='" & rs(0) & "'" & AgendaExtSelected(rs(0), sel) & ">" & Server.HTMLEncode(rs(1) & "") & "</option>" : rs.MoveNext
	'[UI-ANT] Loop
	'[UI-ANT] Response.Write "</select></label>"
'[UI-ANT] End Sub

'[UI-ANT] Sub AgendaExtSelectEstados(cn, sel)
	'[UI-ANT] Dim rs
	'[UI-ANT] Response.Write "<label>Estado<select name='estado' required>"
	'[UI-ANT] Set rs = Exec(cn, "sp_cg_AgendaExt @Modo='estados_listar'")
	'[UI-ANT] Do While Not rs.EOF
		'[UI-ANT] Response.Write "<option value='" & Server.HTMLEncode(rs(1) & "") & "'" & AgendaExtSelected(rs(1), sel) & ">" & Server.HTMLEncode(rs(1) & "") & "</option>"
		'[UI-ANT] rs.MoveNext
	'[UI-ANT] Loop
	'[UI-ANT] Response.Write "</select></label>"
'[UI-ANT] End Sub

'[UI-ANT] Sub AgendaExtInput(labelText, fieldName, fieldValue, fieldType, required)
	'[UI-ANT] Response.Write "<label>" & labelText & "<input name='" & fieldName & "' type='" & fieldType & "' value='" & Server.HTMLEncode(fieldValue & "") & "'"
	'[UI-ANT] If required Then Response.Write " required"
	'[UI-ANT] Response.Write "></label>"
'[UI-ANT] End Sub

'[UI-ANT] Function AgendaExtSelected(value, selectedValue)
	'[UI-ANT] If CStr(value) = CStr(selectedValue) Then AgendaExtSelected = " selected" Else AgendaExtSelected = ""
'[UI-ANT] End Function
'[UI-ANT] Function AgendaExtChecked(value)
	'[UI-ANT] If AgendaExtSqlBit(value) = "1" Then AgendaExtChecked = " checked" Else AgendaExtChecked = ""
'[UI-ANT] End Function
'[UI-ANT] Function AgendaExtOpcionesAccion(cn, selectedValue)
	'[UI-ANT] Dim rs, html
	'[UI-ANT] html = "<option value='0'>Seleccionar</option>"
	'[UI-ANT] Set rs = Exec(cn, "sp_cg_AgendaExt @Modo='acciones_listar'")
	'[UI-ANT] If Not rs Is Nothing Then
		'[UI-ANT] If Not rs.State = 0 Then
			'[UI-ANT] Do While Not rs.EOF
				'[UI-ANT] html = html & "<option value='" & rs(0) & "'" & AgendaExtSelected(rs(0), selectedValue) & ">" & Server.HTMLEncode(rs(1) & "") & "</option>"
				'[UI-ANT] rs.MoveNext
			'[UI-ANT] Loop
		'[UI-ANT] End If
	'[UI-ANT] End If
	'[UI-ANT] AgendaExtOpcionesAccion = html
'[UI-ANT] End Function
'[UI-ANT] Function AgendaExtSqlText(value)
	'[UI-ANT] If Trim(value & "") = "" Then AgendaExtSqlText = "NULL" Else AgendaExtSqlText = "N'" & Replace(value & "", "'", "''") & "'"
'[UI-ANT] End Function
' --- EPOCH: la base guarda fechas como DIAS desde 1900-01-01 (INT) y horas como HHMM (INT) ---
'[UI-ANT] Function AgendaExtSoloDigitos(value)
	'[UI-ANT] Dim i, c, out
	'[UI-ANT] out = ""
	'[UI-ANT] For i = 1 To Len(value & "")
		'[UI-ANT] c = Mid(value & "", i, 1)
		'[UI-ANT] If c >= "0" And c <= "9" Then out = out & c
	'[UI-ANT] Next
	'[UI-ANT] If out = "" Then out = "0"
	'[UI-ANT] AgendaExtSoloDigitos = out
'[UI-ANT] End Function
'[UI-ANT] Function AgendaExtHoy()
	'[UI-ANT] AgendaExtHoy = CLng(CDate(Date())) - CLng(CDate("01-01-1900"))
'[UI-ANT] End Function
' Recibe "yyyy-mm-dd" del form y devuelve DIAS (int) para el SP
'[UI-ANT] Function AgendaExtSqlDate(value)
	'[UI-ANT] If IsDate(value) Then AgendaExtSqlDate = CStr(CLng(CDate(value)) - CLng(CDate("01-01-1900"))) Else AgendaExtSqlDate = "NULL"
'[UI-ANT] End Function
' Recibe "hh:mm" del form y devuelve HHMM (int) para el SP
'[UI-ANT] Function AgendaExtSqlTime(value)
	'[UI-ANT] Dim t, hh, mm
	'[UI-ANT] AgendaExtSqlTime = "NULL"
	'[UI-ANT] t = Trim(value & "")
	'[UI-ANT] If t = "" Then Exit Function
	'[UI-ANT] t = Replace(t, ".", ":")
	'[UI-ANT] If InStr(t, ":") > 0 Then
		'[UI-ANT] hh = CLng(AgendaExtSoloDigitos(Split(t, ":")(0)))
		'[UI-ANT] mm = CLng(AgendaExtSoloDigitos(Split(t, ":")(1)))
		'[UI-ANT] If hh >= 0 And hh <= 23 And mm >= 0 And mm <= 59 Then AgendaExtSqlTime = CStr(hh * 100 + mm)
	'[UI-ANT] ElseIf IsNumeric(t) Then
		'[UI-ANT] AgendaExtSqlTime = CStr(CLng(t))
	'[UI-ANT] End If
'[UI-ANT] End Function
' Texto SQL (N'...') con comillas escapadas; si value viene vacio usa fallback. Para ID_SEG.
'[UI-ANT] Function AgendaExtSqlStr(value, fallback)
	'[UI-ANT] Dim v
	'[UI-ANT] v = Trim(value & "")
	'[UI-ANT] If v = "" Then v = Trim(fallback & "")
	'[UI-ANT] AgendaExtSqlStr = "N'" & Replace(v, "'", "''") & "'"
'[UI-ANT] End Function
'[UI-ANT] Function AgendaExtSqlInt(value, fallback)
	'[UI-ANT] If IsNumeric(value) Then AgendaExtSqlInt = CLng(value) Else AgendaExtSqlInt = fallback
'[UI-ANT] End Function
'[UI-ANT] Function AgendaExtSqlNullableInt(value)
	'[UI-ANT] If IsNumeric(value) Then
		'[UI-ANT] If CLng(value) > 0 Then AgendaExtSqlNullableInt = CStr(CLng(value)) Else AgendaExtSqlNullableInt = "NULL"
	'[UI-ANT] Else
		'[UI-ANT] AgendaExtSqlNullableInt = "NULL"
	'[UI-ANT] End If
'[UI-ANT] End Function
'[UI-ANT] Function AgendaExtEsIdValido(value)
	'[UI-ANT] AgendaExtEsIdValido = False
	'[UI-ANT] If IsNumeric(value) Then
		'[UI-ANT] If CLng(value) > 0 Then AgendaExtEsIdValido = True
	'[UI-ANT] End If
'[UI-ANT] End Function
'[UI-ANT] Function AgendaExtSqlBit(value)
	'[UI-ANT] If value = "1" Or UCase(value & "") = "SI" Or value = True Then AgendaExtSqlBit = "1" Else AgendaExtSqlBit = "0"
'[UI-ANT] End Function
'[UI-ANT] Function AgendaExtVal(rs, fieldName, fallback)
	'[UI-ANT] If IsNull(rs(fieldName)) Then AgendaExtVal = fallback Else AgendaExtVal = rs(fieldName)
'[UI-ANT] End Function
' DIAS (int) -> "yyyy-mm-dd" para <input type="date">
'[UI-ANT] Function AgendaExtHtmlDate(value)
	'[UI-ANT] Dim d
	'[UI-ANT] AgendaExtHtmlDate = ""
	'[UI-ANT] If IsNumeric(value) Then
		'[UI-ANT] If CLng(value) > 0 Then
			'[UI-ANT] d = CDate(CLng(value) + CLng(CDate("01-01-1900")))
			'[UI-ANT] AgendaExtHtmlDate = Year(d) & "-" & Right("0" & Month(d), 2) & "-" & Right("0" & Day(d), 2)
		'[UI-ANT] End If
	'[UI-ANT] End If
'[UI-ANT] End Function
' HHMM (int) -> "hh:mm"
'[UI-ANT] Function AgendaExtHtmlTime(value)
	'[UI-ANT] Dim n
	'[UI-ANT] AgendaExtHtmlTime = ""
	'[UI-ANT] If IsNumeric(value) Then
		'[UI-ANT] n = CLng(value)
		'[UI-ANT] If n >= 0 Then AgendaExtHtmlTime = Right("0" & CStr(n \ 100), 2) & ":" & Right("0" & CStr(n Mod 100), 2)
	'[UI-ANT] End If
'[UI-ANT] End Function
'[UI-ANT] Sub AgendaExtRedir(message, id, xmenu, xtmenu)
	'[UI-ANT] Dim msgJs
	'[UI-ANT] msgJs = Replace(message, "\", "\\")
	'[UI-ANT] msgJs = Replace(msgJs, "'", "\'")
	'[UI-ANT] msgJs = Replace(msgJs, vbCrLf, "\n")
	'[UI-ANT] msgJs = Replace(msgJs, vbCr, "\n")
	'[UI-ANT] msgJs = Replace(msgJs, vbLf, "\n")
	'[UI-ANT] Response.Write "<script>alert('" & msgJs & "');window.location='agenda.asp?opc=det&id=" & id & "&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu) & "';</script>"
'[UI-ANT] End Sub
'[UI-ANT] Function AgendaExtCss()
	'[UI-ANT] AgendaExtCss = ":root{--navy:#0e1949;--blue:#165dff;--blue-soft:#eef4ff;--ink:#24324a;--muted:#6d7b91;--line:#d8e0ec;--surface:#fff;--canvas:#f4f7fb;--shadow:0 8px 22px rgba(30,52,91,.07)}*{box-sizing:border-box}body{margin:0;background:var(--canvas);color:var(--ink);font:14px Segoe UI,Tahoma,sans-serif}button,input,select,textarea{font:inherit}a{color:#165dff}button{cursor:pointer}.app-header{background:var(--navy);color:#fff;padding:14px 28px;display:flex;align-items:center;justify-content:space-between;gap:20px}.brand{display:flex;align-items:center;gap:12px;font-size:16px}.brand-mark{width:36px;height:36px;border:2px solid #8ea9ff;border-radius:8px;display:grid;place-items:center;font-weight:800}.brand small{display:block;color:#bdc9ee;font-size:11px;margin-top:2px}.shell{max-width:1260px;margin:0 auto;padding:26px 28px 48px}.breadcrumb{color:var(--muted);font-size:12px;margin-bottom:8px}.page-head{display:flex;align-items:end;justify-content:space-between;gap:20px;margin-bottom:22px}h1{color:var(--navy);font-size:25px;margin:0}h2{color:var(--navy);font-size:15px;margin:0}.page-head p{color:var(--muted);margin:6px 0 0}.toolbar{display:flex;gap:8px;align-items:center;flex-wrap:wrap}.btn{display:inline-block;border:1px solid var(--line);border-radius:5px;padding:9px 13px;background:#fff;color:var(--ink);font-weight:600;text-decoration:none}.btn-primary{color:#fff;background:var(--blue);border-color:var(--blue)}.btn-secondary{background:var(--blue-soft);border-color:#b9ccff;color:#174cb9}.panel{background:var(--surface);border:1px solid var(--line);border-radius:7px;box-shadow:var(--shadow);margin-bottom:18px;overflow:hidden}.panel-head{padding:15px 18px;border-bottom:1px solid var(--line);display:flex;justify-content:space-between;align-items:center;gap:14px}.panel-head span,.muted{color:var(--muted);font-size:12px}.panel-body{padding:18px}.form-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px}.form-grid .full{grid-column:1 / -1}label{display:flex;flex-direction:column;gap:6px;color:var(--muted);font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.04em}input,select,textarea{width:100%;border:1px solid #c8d3e3;border-radius:5px;background:#fff;color:var(--ink);padding:9px 10px;outline:0;font-weight:400;text-transform:none;letter-spacing:0}textarea{min-height:80px;resize:vertical}input:focus,select:focus,textarea:focus{border-color:var(--blue);box-shadow:0 0 0 3px rgba(22,93,255,.12)}.grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px;padding:18px}.grid+button,.panel>button{margin:0 18px 14px}.panel>h2{padding:15px 18px;border-bottom:1px solid var(--line)}table{width:100%;border-collapse:collapse}.table-wrap{overflow-x:auto}th{background:#f1f5fa;color:#60708a;font-size:11px;text-transform:uppercase;letter-spacing:.04em;text-align:left;white-space:nowrap}th,td{padding:12px 14px;border-bottom:1px solid #e7ecf3;vertical-align:middle}tbody tr:hover{background:#f8fbff}.status{display:inline-flex;align-items:center;border-radius:4px;padding:4px 8px;background:#fff1cf;color:#a96800;font-size:11px;font-weight:700}.pagination{display:flex;justify-content:center;gap:8px;padding-top:14px}.chip{display:inline-flex;gap:8px;align-items:center;background:var(--blue-soft);color:#174cb9;border-radius:5px;padding:7px 9px;margin:3px}.chip small{color:var(--muted);font-size:11px;font-weight:400}.chip-remove{margin-left:4px;color:#8b0000;font-weight:700;text-decoration:none;line-height:1;border:none;background:transparent;cursor:pointer;font-size:12px}.chip-remove:hover{color:#d32f2f}.action-bar{display:flex;justify-content:space-between;gap:10px;align-items:center;padding-top:4px}.action-bar .right{display:flex;gap:8px}.actions{display:flex;justify-content:flex-end;gap:8px;margin:18px 0}.actions button{margin:0}@media(max-width:900px){.shell{padding:20px 14px 40px}.app-header{padding:12px 16px}.page-head{align-items:flex-start;flex-direction:column}.form-grid,.grid{grid-template-columns:1fr 1fr}}@media(max-width:560px){.form-grid,.grid{grid-template-columns:1fr}.form-grid .full{grid-column:auto}.action-bar,.action-bar .right{width:100%;flex-wrap:wrap}.action-bar .btn,.action-bar button{flex:1}.brand small{display:none}}"
'[UI-ANT] End Function
'[UI-ANT] Function AgendaExtJs()
	'[UI-ANT] AgendaExtJs = "var corN=parseInt(document.getElementById('corresponsable_count').value||0),subN=parseInt(document.getElementById('subtema_count').value||0);function agendaExtAgregarCorresponsable(){var p=document.getElementById('corPersona'),n=document.getElementById('corNombre'),e=document.getElementById('corEmail'),c=document.getElementById('corCopia'),name=p.options[p.selectedIndex].text;if(!p.value&&!n.value.trim()){alert('Seleccione una persona o complete el nombre manual.');return}if(!p.value&&!e.value.trim()&&c.value==='1'){alert('El email es obligatorio para contactos manuales con copia.');return}corN++;document.getElementById('corresponsable_count').value=corN;var d=document.createElement('div');d.className='chip';var label=name==='Seleccionar'?n.value+' &lt;'+e.value+'&gt;':name;if(e.value.trim()&&name!=='Seleccionar')label+=' &lt;'+e.value+'&gt;';if(c.value==='1')label+=' <small>(copia)</small>';d.innerHTML=label;d.innerHTML+='<input type=hidden name=cor_persona_'+corN+' value='+p.value+'><input type=hidden name=cor_nombre_'+corN+' value='+n.value+'><input type=hidden name=cor_email_'+corN+' value='+e.value+'><input type=hidden name=cor_copia_'+corN+' value='+c.value+'><button type=button class=chip-remove onclick=this.parentNode.remove()>x</button>';document.getElementById('corresponsables').appendChild(d);p.value='';n.value='';e.value=''}function agendaExtAgregarSubtema(){subN++;document.getElementById('subtema_count').value=subN;var tr=document.createElement('tr');tr.innerHTML='<td><select name=sub_tipo_'+subN+'><option value=1>PROYECTO</option><option value=2>INCIDENTE</option><option value=3>CONSULTORIA</option><option value=4>OTRO</option></select></td><td><input name=sub_proyecto_'+subN+' type=number placeholder=ID proyecto><input name=sub_incidente_'+subN+' type=number placeholder=ID incidente></td><td><input name=sub_tema_'+subN+' required maxlength=500></td><td><input name=sub_tiempo_'+subN+' type=number min=0 value=0></td><td><input name=sub_seguimiento_'+subN+' type=checkbox value=1></td><td><select name=sub_accion_'+subN+'>'+document.getElementById('agendaAccionTemplate').innerHTML+'</select></td><td><button type=button onclick=this.parentNode.parentNode.remove()>Quitar</button></td>';document.querySelector('#subtemas tbody').appendChild(tr)}function agendaExtValidar(){var a=document.querySelector('[name=hora_inicio]').value,b=document.querySelector('[name=hora_fin]').value;if(a&&b&&b<=a){alert('La hora de fin debe ser posterior a la hora de inicio.');return false}return true}"
'[UI-ANT] End Function
'[UI-ANT] Function AgendaExtRelacionJs()
	'[UI-ANT] AgendaExtRelacionJs = "function agendaExtCambiarRelacion(sel){var box=sel.parentNode.parentNode.querySelector('.relation-fields'),input=box.querySelector('.relation-id'),label=box.querySelector('.relation-id-label'),buscar=box.querySelector('.btn-buscar-entidad');var visible=sel.value==='1'||sel.value==='2';box.style.display=visible?'block':'none';label.firstChild.nodeValue=sel.value==='2'?'ID incidente':'ID proyecto';if(buscar){buscar.style.display=visible?'inline-block':'none';}if(!visible){input.value='';}}function agendaExtBuscarEntidad(boton){var input=boton.parentNode.querySelector('.relation-id'),tipo=boton.closest('tr').querySelector('select[name^=sub_tipo_]').value,tabla=tipo==='1'?'PROYECTO':'INCIDENTE';window.open('buscar_entidades.asp?tabla='+tabla+'&campo='+encodeURIComponent(input.name),'buscar_entidad','width=900,height=650,resizable=yes,scrollbars=yes');}function agendaExtSeleccionarEntidad(campo,id,nombre){var input=document.querySelector('input[name='+JSON.stringify(campo)+']');if(input){input.value=id;input.title=nombre;}}function agendaExtAgregarSubtema(){subN++;document.getElementById('subtema_count').value=subN;var tr=document.createElement('tr');tr.innerHTML='<td><select name=sub_tipo_'+subN+' onchange=agendaExtCambiarRelacion(this)><option value=1>PROYECTO</option><option value=2>INCIDENTE</option><option value=3>CONSULTORIA</option><option value=4>OTRO</option></select></td><td><div class=relation-fields><label class=relation-id-label>ID proyecto<input class=relation-id name=sub_relacion_id_'+subN+' type=number min=1 placeholder=ID><button type=button class=btn-buscar-entidad onclick=agendaExtBuscarEntidad(this)>Buscar</button></label></div></td><td><input name=sub_tema_'+subN+' required maxlength=500></td><td><textarea name=sub_conclusion_'+subN+' rows=2 maxlength=4000></textarea></td><td><input name=sub_tiempo_'+subN+' type=number min=0 value=0></td><td><input name=sub_seguimiento_'+subN+' type=checkbox value=1></td><td><select name=sub_accion_'+subN+'>'+document.getElementById('agendaAccionTemplate').innerHTML+'</select></td><td><button type=button onclick=this.parentNode.parentNode.remove()>Quitar</button></td>';document.querySelector('#subtemas tbody').appendChild(tr);agendaExtCambiarRelacion(tr.querySelector('select'));}document.querySelectorAll('#subtemas select[name^=sub_tipo_]').forEach(agendaExtCambiarRelacion);"
'[UI-ANT] End Function
%>