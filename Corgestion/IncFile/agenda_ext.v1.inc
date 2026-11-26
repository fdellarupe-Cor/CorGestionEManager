<%
Function AgendaExtDebeManejar(opc)
	AgendaExtDebeManejar = (opc = "consulta" Or opc = "cargar" Or opc = "det" Or opc = "agenda_ext_guardar" Or opc = "agenda_ext_postergar" Or opc = "agenda_ext_seguimiento" Or opc = "agenda_ext_eliminar_corresp")
End Function

Sub AgendaExtRender(cn, opc, xmenu, xtmenu)
	Dim id, rs, avisosCruce, mensajeGuardado
	If opc = "consulta" Then
		AgendaExtConsulta cn, xmenu, xtmenu
		Exit Sub
	End If
	id = Request.QueryString("id")
	If opc = "agenda_ext_guardar" Then
		avisosCruce = ""
		id = AgendaExtGuardar(cn, id, avisosCruce)
		mensajeGuardado = "Agenda guardada correctamente."
		If Trim(avisosCruce) <> "" Then
			mensajeGuardado = mensajeGuardado & vbCrLf & vbCrLf & "ATENCION: no se agregaron los siguientes participantes por tener otra reunion cruzada en ese horario:" & vbCrLf & avisosCruce
		End If
		AgendaExtRedir mensajeGuardado, id, xmenu, xtmenu
		Exit Sub
	End If
	If opc = "agenda_ext_postergar" Then
		id = AgendaExtPostergar(cn, id)
		AgendaExtRedir "La reunión fue postergada.", id, xmenu, xtmenu
		Exit Sub
	End If
	If opc = "agenda_ext_seguimiento" Then
		id = AgendaExtSeguimiento(cn, id)
		AgendaExtRedir "El seguimiento fue generado.", id, xmenu, xtmenu
		Exit Sub
	End If
	If opc = "agenda_ext_eliminar_corresp" Then
		AgendaExtEliminarCorresp cn, xmenu, xtmenu
		Exit Sub
	End If
	If AgendaExtEsIdValido(id) Then
		Set rs = Exec(cn, "sp_cg_AgendaExt @Modo='obtener', @Id=" & CLng(id))
	Else
		Set rs = Nothing
	End If
	AgendaExtPagina cn, rs, id, xmenu, xtmenu
End Sub

Sub AgendaExtConsulta(cn, xmenu, xtmenu)
	Dim rs, mover, mostrar, pos, volver, editar, estado, responsable, fecha, padreLista, padreListaUrl
	mostrar = 15
	mover = AgendaExtSqlInt(Request.QueryString("CantRegMover"), 0)
	If mover < 0 Then mover = 0
	Set rs = Exec(cn, "sp_cg_AgendaExt @Modo='consulta', @CantRegMover=" & mover & ", @CantRegMostrar=" & mostrar)
	volver = "agenda.asp?xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu)
	Response.Write "<!doctype html><html lang='es'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width, initial-scale=1'><title>Agenda | Cor Gestión</title><style>" & AgendaExtCss() & "</style></head><body>"
	Response.Write "<header class='app-header'><div class='brand'><div class='brand-mark'>CG</div><div><strong>Cor Gestión</strong><small>Sistema Administrativo</small></div></div><a class='btn btn-secondary' href='" & volver & "'>Volver</a></header><main class='shell'><div class='breadcrumb'>Agenda / Reuniones operativas</div><div class='page-head'><div><h1>Agenda</h1><p>Consulta y administra reuniones, seguimientos y temas relacionados.</p></div><div class='toolbar'><a class='btn btn-primary' href='agenda.asp?opc=edicion&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu) & "'>+ Cargar reunión</a><button type='button' class='btn btn-secondary' onclick=""window.open('agenda_semana.asp?xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu) & "','agenda_semana','width=1520,height=900,resizable=yes,scrollbars=yes')"">Vista semanal</button></div></div><section class='panel'><div class='panel-head'><div><h2>Reuniones agendadas</h2><span class='muted'>Agenda operativa</span></div></div><div class='panel-body table-wrap'><table><thead><tr><th>ID agenda</th><th>Asunto</th><th>Fecha</th><th>Horario</th><th>Responsable</th><th>Estado</th><th>Registro padre</th><th>Relaciones</th><th></th></tr></thead><tbody>"
	pos = 0
	Do While Not rs.EOF And pos < mostrar
		fecha = AgendaExtHtmlDate(AgendaExtVal(rs, "FECHA_REUNION", ""))
		estado = AgendaExtVal(rs, "ESTADO", "PENDIENTE")
		responsable = AgendaExtVal(rs, "RESPONSABLE", "Sin responsable")
		editar = "agenda.asp?opc=edicion&id=" & rs("ID") & "&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu)
		padreLista = AgendaExtVal(rs, "ID_AGENDA_PADRE", "")
		padreListaUrl = "agenda.asp?opc=det&id=" & padreLista & "&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu)
		If AgendaExtEsIdValido(padreLista) Then padreLista = "<a href='" & padreListaUrl & "'>Agenda #" & Server.HTMLEncode(padreLista & "") & "</a>" Else padreLista = "-"
		Response.Write "<tr><td><a href='" & editar & "'>#" & rs("ID") & "</a></td><td>" & Server.HTMLEncode(Left(AgendaExtVal(rs, "ASUNTO", ""), 80)) & "</td><td>" & fecha & "</td><td>" & AgendaExtHtmlTime(AgendaExtVal(rs, "HORA_INICIO", "")) & " - " & AgendaExtHtmlTime(AgendaExtVal(rs, "HORA_FIN", "")) & "</td><td>" & Server.HTMLEncode(responsable) & "</td><td><span class='status'>" & Server.HTMLEncode(estado) & "</span></td><td>" & padreLista & "</td><td>" & rs("CANT_CORRESP") & " corresponsable(s), " & rs("CANT_SUBTEMAS") & " subtema(s)</td><td><a href='" & editar & "'>Ver / editar</a></td></tr>"
		pos = pos + 1
		rs.MoveNext
	Loop
	Response.Write "</tbody></table></div></section><div class='pagination'>"
	If mover > 0 Then Response.Write "<a href='agenda.asp?opc=consulta&CantRegMover=" & (mover - mostrar) & "&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu) & "'>Anterior</a>"
	If Not rs.EOF Then Response.Write "<a href='agenda.asp?opc=consulta&CantRegMover=" & (mover + mostrar) & "&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu) & "'>Siguiente</a>"
	Response.Write "</div></main></body></html>"
	rs.Close
End Sub

Function AgendaExtGuardar(cn, idActual, ByRef avisosCruceOut)
	Dim q, rs, id, i, n, persona, nombre, email, copia, tipo, relacionId, proyecto, incidente, tema, conclusion, accion
	Dim fechaReunionSql, horaInicioSql, horaFinSql, avisosCruce, descripcionCruce
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
	q = q & ", @IdRespon=" & AgendaExtSqlInt(Request.Form("id_respon"), Session("usuario"))
	q = q & ", @IdHerramienta=" & AgendaExtSqlNullableInt(Request.Form("id_herramienta"))
	q = q & ", @Mensaje=" & AgendaExtSqlText(Request.Form("mensaje"))
	q = q & ", @Conclusion=" & AgendaExtSqlText(Request.Form("conclusion"))
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
	n = AgendaExtSqlInt(Request.Form("corresponsable_count"), 0)
		For i = 1 To n
			persona = Request.Form("cor_persona_" & i)
			nombre = Request.Form("cor_nombre_" & i)
			email = Request.Form("cor_email_" & i)
			copia = AgendaExtSqlBit(Request.Form("cor_copia_" & i))
			If IsNumeric(persona) Then
				If CLng(persona) > 0 Then
					' Antes de sumar un participante del sistema, se valida si tiene
					' otra reunion cruzada en el mismo horario (ya sea como
					' responsable o como corresponsable de otra agenda).
					descripcionCruce = AgendaExtDescribirCruceParticipante(cn, id, CLng(persona), fechaReunionSql, horaInicioSql, horaFinSql)
					If descripcionCruce <> "" Then
						avisosCruce = avisosCruce & "- " & AgendaExtNombrePersona(cn, CLng(persona)) & ": " & descripcionCruce & vbCrLf
					Else
						ExecNonQuery cn, "sp_cg_AgendaExt @Modo='corresp_agregar', @Id=" & id & ", @IdPersona=" & CLng(persona) & ", @EnviarCopia=" & copia
					End If
				ElseIf Trim(nombre) <> "" Then
					ExecNonQuery cn, "sp_cg_AgendaExt @Modo='corresp_agregar', @Id=" & id & ", @NombreManual=" & AgendaExtSqlText(nombre) & ", @EmailManual=" & AgendaExtSqlText(email) & ", @EnviarCopia=" & copia
				End If
			ElseIf Trim(nombre) <> "" Then
				ExecNonQuery cn, "sp_cg_AgendaExt @Modo='corresp_agregar', @Id=" & id & ", @NombreManual=" & AgendaExtSqlText(nombre) & ", @EmailManual=" & AgendaExtSqlText(email) & ", @EnviarCopia=" & copia
			End If
		Next

		n = AgendaExtSqlInt(Request.Form("subtema_count"), 0)
		For i = 1 To n
			tipo = AgendaExtSqlInt(Request.Form("sub_tipo_" & i), 0)
			relacionId = AgendaExtSqlInt(Request.Form("sub_relacion_id_" & i), 0)
			proyecto = 0
			incidente = 0
			If tipo = 1 Then proyecto = relacionId
			If tipo = 2 Then incidente = relacionId
			tema = Request.Form("sub_tema_" & i)
			conclusion = Request.Form("sub_conclusion_" & i)
			accion = AgendaExtSqlInt(Request.Form("sub_accion_" & i), 0)
			ExecNonQuery cn, "sp_cg_AgendaExt @Modo='subtema_agregar', @Id=" & id & ", @IdTipoRelacion=" & tipo & ", @IdProyecto=" & AgendaExtSqlNullableInt(proyecto) & ", @IdIncidente=" & AgendaExtSqlNullableInt(incidente) & ", @Tema=" & AgendaExtSqlText(tema) & ", @Conclusion=" & AgendaExtSqlText(conclusion) & ", @TiempoUsado=" & AgendaExtSqlInt(Request.Form("sub_tiempo_" & i), 0) & ", @Orden=" & i & ", @Seguimiento=" & AgendaExtSqlBit(Request.Form("sub_seguimiento_" & i)) & ", @IdAccion=" & AgendaExtSqlNullableInt(accion)
		Next
	avisosCruceOut = avisosCruce
	AgendaExtGuardar = id
End Function

' Chequea si @IdPersona tiene otra reunion (PENDIENTE/POSTERGADA) cuyo
' horario se solape con [horaInicioSql, horaFinSql) el dia fechaReunionSql,
' ya sea como responsable de esa reunion o como corresponsable de ella.
' idAgendaActual se excluye del chequeo (es la reunion que se esta guardando).
' Devuelve "" si no hay cruce, o un texto con el detalle de los cruces.
Function AgendaExtDescribirCruceParticipante(cn, idAgendaActual, idPersona, fechaReunionSql, horaInicioSql, horaFinSql)
	Dim rsCruce, detalle
	AgendaExtDescribirCruceParticipante = ""
	If fechaReunionSql = "NULL" Or horaInicioSql = "NULL" Or horaFinSql = "NULL" Then Exit Function
	Set rsCruce = Exec(cn, "sp_cg_Agenda_ValidarCruceHorario @IdResponsable=" & idPersona & ", @FechaReunion=" & fechaReunionSql & ", @HoraInicio=" & horaInicioSql & ", @HoraFin=" & horaFinSql & ", @IdAgendaExcluir=" & idAgendaActual)
	detalle = ""
	Do While Not rsCruce.EOF
		If detalle <> "" Then detalle = detalle & "; "
		detalle = detalle & """" & AgendaExtVal(rsCruce, "ASUNTO", "(sin asunto)") & """ (#" & rsCruce("ID_AGENDA_CRUCE") & ") " & AgendaExtHtmlTime(rsCruce("HORA_INICIO")) & "-" & AgendaExtHtmlTime(rsCruce("HORA_FIN"))
		rsCruce.MoveNext
	Loop
	rsCruce.Close
	AgendaExtDescribirCruceParticipante = detalle
End Function

Function AgendaExtNombrePersona(cn, idPersona)
	Dim rsPersona, nombre
	nombre = "Persona #" & idPersona
	Set rsPersona = Exec(cn, "SELECT RTRIM(ISNULL(NOMBRE,'') + ' ' + ISNULL(APELLIDO,'')) AS NOMBRE FROM dbo.CG_COR_RESPONSABLES WHERE ID=" & idPersona)
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
	End If
	Response.Redirect "agenda.asp?opc=edicion&id=" & id & "&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu)
End Sub

Sub AgendaExtPagina(cn, rs, id, xmenu, xtmenu)
	Dim editar, asunto, fecha, aviso, hi, hf, mensaje, conclusion, resp, herramienta, estado, agendaId, agendaPadreId, volver, padreUrl
	editar = False
	If IsObject(rs) Then
		If Not rs Is Nothing Then
			If Not rs.EOF Then editar = True
		End If
	End If
	If editar Then
		agendaId = AgendaExtVal(rs, "ID", id)
		agendaPadreId = AgendaExtVal(rs, "ID_AGENDA_PADRE", "")
		asunto = AgendaExtVal(rs, "ASUNTO", "") : fecha = AgendaExtVal(rs, "FECHA_REUNION", "") : aviso = AgendaExtVal(rs, "FECHA_AVISO", "")
		hi = AgendaExtVal(rs, "HORA_INICIO", "") : hf = AgendaExtVal(rs, "HORA_FIN", "") : mensaje = AgendaExtVal(rs, "MENSAJE", "")
		conclusion = AgendaExtVal(rs, "CONCLUSION", "") : resp = AgendaExtVal(rs, "ID_RESPON", Session("usuario")) : herramienta = AgendaExtVal(rs, "ID_HERRAMIENTA", 0) : estado = AgendaExtVal(rs, "ESTADO", "PENDIENTE")
	Else
		agendaId = "" : agendaPadreId = ""
		asunto = "" : fecha = AgendaExtHoy() : aviso = AgendaExtHoy() : hi = 900 : hf = 1000 : mensaje = "" : conclusion = "" : resp = Session("usuario") : herramienta = 0 : estado = "PENDIENTE"
	End If
	padreUrl = ""
	If AgendaExtEsIdValido(agendaPadreId) Then padreUrl = "agenda.asp?opc=det&id=" & CLng(agendaPadreId) & "&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu)
	Dim tituloPagina
	If editar Then tituloPagina = "Editar reunión" Else tituloPagina = "Nueva reunión"
	volver = "agenda.asp?xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu)
	Response.Write "<!doctype html><html lang='es'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width, initial-scale=1'><title>Agenda | Cor Gestión</title><style>" & AgendaExtCss() & "</style></head><body><header class='app-header'><div class='brand'><div class='brand-mark'>CG</div><div><strong>Cor Gestión</strong><small>Sistema Administrativo</small></div></div><a class='btn btn-secondary' href='" & volver & "'>Volver al listado</a></header><main class='shell'><div class='breadcrumb'>Agenda / Reuniones operativas</div><div class='page-head'><div><h1>" & tituloPagina & "</h1><p>Cabecera, participantes y temas relacionados.</p></div></div>"
	Dim guardarAction
	guardarAction = "agenda.asp?opc=agenda_ext_guardar"
	If editar Then guardarAction = guardarAction & "&id=" & id
	guardarAction = guardarAction & "&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu)
	Response.Write "<form id='agendaExt' method='post' action='" & guardarAction & "' onsubmit='return agendaExtValidar();'><section class='panel'><div class='panel-head'><div><h2>Datos generales</h2><span class='muted'>Cabecera de AGENDA</span></div><div class='record-meta'><span>ID agenda: <strong>" & Server.HTMLEncode(agendaId & "") & "</strong></span>"
	If padreUrl <> "" Then Response.Write "<span>Registro padre: <a href='" & padreUrl & "'>Agenda #" & Server.HTMLEncode(agendaPadreId & "") & "</a></span>"
	Response.Write "</div></div><div class='panel-body'><div class='form-grid'>"
	AgendaExtInput "Asunto", "asunto", asunto, "text", True
	AgendaExtSelectResponsables cn, resp
	AgendaExtInput "Fecha reunión", "fecha_reunion", AgendaExtHtmlDate(fecha), "date", True
	AgendaExtInput "Fecha aviso", "fecha_aviso", AgendaExtHtmlDate(aviso), "date", True
	AgendaExtInput "Hora inicio", "hora_inicio", AgendaExtHtmlTime(hi), "time", True
	AgendaExtInput "Hora fin", "hora_fin", AgendaExtHtmlTime(hf), "time", True
	AgendaExtSelectHerramientas cn, herramienta
	AgendaExtSelectEstados cn, estado
	Response.Write "<label class='full'>Mensaje<textarea name='mensaje' rows='3'>" & Server.HTMLEncode(mensaje & "") & "</textarea></label></div></div></section>"
	AgendaExtCorresponsables cn, id, xmenu, xtmenu
	AgendaExtSubtemas cn, id, xmenu, xtmenu
	Response.Write "<section class='panel'><div class='panel-body'><label class='full'>Conclusión general de la reunión<textarea name='conclusion' rows='4'>" & Server.HTMLEncode(conclusion & "") & "</textarea></label></div></section><div class='action-bar'><a class='btn' href='" & volver & "'>Cancelar</a><div class='right'><button class='btn btn-primary' type='submit'>Guardar reunión</button>"
	If editar Then Response.Write "<button type='submit' formaction='agenda.asp?opc=agenda_ext_postergar&id=" & id & "&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu) & "'>Postergar</button><button type='submit' formaction='agenda.asp?opc=agenda_ext_seguimiento&id=" & id & "&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu) & "'>Generar seguimiento</button>"
	Response.Write "</div></div></form></main><script>" & AgendaExtJs() & AgendaExtRelacionJs() & "</script></body></html>"
End Sub

Sub AgendaExtCorresponsables(cn, id, xmenu, xtmenu)
	Dim rs, detalle, nombre, persona, manual, email, copia, n, correspId, chipText, eliminarUrl
	Response.Write "<section class='panel'><h2>Participantes (corresponsables)</h2><div class='grid'><label>Persona del sistema<select id='corPersona'><option value=''>Seleccionar</option>"
	Set rs = Exec(cn, "sp_cg_Responsables @Modo='listar_activos_sin', @ExcluirId=" & Session("usuario"))
	Do While Not rs.EOF
		Response.Write "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>" : rs.MoveNext
	Loop
	Response.Write "</select></label><label>Nombre manual<input id='corNombre' type='text'></label><label>Email manual<input id='corEmail' type='email'></label><label>Enviar copia<select id='corCopia'><option value='1'>SI</option><option value='0'>NO</option></select></label></div><button type='button' onclick='agendaExtAgregarCorresponsable()'>Agregar corresponsable</button><div id='corresponsables'>"
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
			chipText = Server.HTMLEncode(nombre & "")
			If Trim(email & "") <> "" Then chipText = chipText & " &lt;" & Server.HTMLEncode(email & "") & "&gt;"
			If copia = 1 Then chipText = chipText & " <small>(copia)</small>"
			eliminarUrl = "agenda.asp?opc=agenda_ext_eliminar_corresp&id=" & id & "&corresp_id=" & correspId & "&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu)
			Response.Write "<div class='chip'>" & chipText & "<input type='hidden' name='cor_persona_" & n & "' value='" & Server.HTMLEncode(persona & "") & "'><input type='hidden' name='cor_nombre_" & n & "' value='" & Server.HTMLEncode(manual & "") & "'><input type='hidden' name='cor_email_" & n & "' value='" & Server.HTMLEncode(email & "") & "'><input type='hidden' name='cor_copia_" & n & "' value='" & Server.HTMLEncode(copia & "") & "'><a href='" & eliminarUrl & "' class='chip-remove' title='Eliminar participante' onclick=""return confirm('¿Eliminar participante?');"">x</a></div>"
			detalle.MoveNext
		Loop
	End If
	Response.Write "</div><input type='hidden' id='corresponsable_count' name='corresponsable_count' value='" & n & "'></section>"
End Sub

Sub AgendaExtSubtemas(cn, id, xmenu, xtmenu)
	Dim detalle, n, tipo, proyecto, incidente, relacionId, tema, conclusion, minutos, seguimiento, accion, opcionesAccion
	opcionesAccion = AgendaExtOpcionesAccion(cn, 0)
	Response.Write "<section class='panel'><h2>Temario y subtemas</h2><table id='subtemas'><thead><tr><th>Tipo</th><th>Proyecto / incidente</th><th>Tema a tratar</th><th>Conclusión</th><th>Min.</th><th>Seguimiento</th><th>Acción</th><th></th></tr></thead><tbody>"
	n = 0
	If AgendaExtEsIdValido(id) Then
		Set detalle = Exec(cn, "sp_cg_AgendaExt @Modo='subtema_listar', @Id=" & CLng(id))
		Do While Not detalle.EOF
			n = n + 1 : tipo = AgendaExtVal(detalle, "ID_TIPO_RELACION", 0) : proyecto = AgendaExtVal(detalle, "ID_PROYECTO", "") : incidente = AgendaExtVal(detalle, "ID_INCIDENTE", "") : tema = AgendaExtVal(detalle, "TEMA_A_TRATAR", "") : conclusion = AgendaExtVal(detalle, "CONCLUSION", "") : minutos = AgendaExtVal(detalle, "TIEMPO_USADO", 0) : seguimiento = AgendaExtVal(detalle, "SEGUIMIENTO", 0) : accion = AgendaExtVal(detalle, "ID_ACCION", "")
			opcionesAccion = AgendaExtOpcionesAccion(cn, accion)
			relacionId = ""
			If tipo = 1 Then relacionId = proyecto
			If tipo = 2 Then relacionId = incidente
			Response.Write "<tr><td><select name='sub_tipo_" & n & "' onchange='agendaExtCambiarRelacion(this)'><option value='1'" & AgendaExtSelected(1, tipo) & ">PROYECTO</option><option value='2'" & AgendaExtSelected(2, tipo) & ">INCIDENTE</option><option value='3'" & AgendaExtSelected(3, tipo) & ">CONSULTORIA</option><option value='4'" & AgendaExtSelected(4, tipo) & ">OTRO</option></select></td><td><div class='relation-fields' style='display:block'><label class='relation-id-label'>ID proyecto/incidente<input class='relation-id' name='sub_relacion_id_" & n & "' type='number' min='1' value='" & Server.HTMLEncode(relacionId & "") & "' placeholder='ID'><button type='button' class='btn-buscar-entidad' onclick='agendaExtBuscarEntidad(this)'>Buscar</button></label></div></td><td><input name='sub_tema_" & n & "' value='" & Server.HTMLEncode(tema & "") & "' required maxlength='500'></td><td><textarea name='sub_conclusion_" & n & "' rows='2' maxlength='4000'>" & Server.HTMLEncode(conclusion & "") & "</textarea></td><td><input name='sub_tiempo_" & n & "' type='number' min='0' value='" & minutos & "'></td><td><input name='sub_seguimiento_" & n & "' type='checkbox' value='1'" & AgendaExtChecked(seguimiento) & "></td><td><select name='sub_accion_" & n & "'>" & opcionesAccion & "</select></td><td><button type='button' onclick='this.parentNode.parentNode.remove()'>Quitar</button></td></tr>"
			detalle.MoveNext
		Loop
	End If
	Response.Write "</tbody></table><button type='button' onclick='agendaExtAgregarSubtema()'>Agregar subtema</button><input type='hidden' id='subtema_count' name='subtema_count' value='" & n & "'><select id='agendaAccionTemplate' style='display:none'>" & AgendaExtOpcionesAccion(cn, 0) & "</select></section>"
End Sub

Sub AgendaExtSelectResponsables(cn, resp)
	Dim rs
	Response.Write "<label>Responsable<select name='id_respon'>"
	Set rs = Exec(cn, "sp_cg_Responsables @Modo='listar_activos_sin', @ExcluirId=0")
	Do While Not rs.EOF
		Response.Write "<option value='" & rs(0) & "'" & AgendaExtSelected(rs(0), resp) & ">" & Server.HTMLEncode(rs(1) & "") & "</option>" : rs.MoveNext
	Loop
	Response.Write "</select></label>"
End Sub

Sub AgendaExtSelectHerramientas(cn, sel)
	Dim rs
	Response.Write "<label>Herramienta<select name='id_herramienta'><option value='0'>Ninguna</option>"
	Set rs = Exec(cn, "sp_cg_Catalogo_Listar @Catalogo='aplicativos'")
	Do While Not rs.EOF
		Response.Write "<option value='" & rs(0) & "'" & AgendaExtSelected(rs(0), sel) & ">" & Server.HTMLEncode(rs(1) & "") & "</option>" : rs.MoveNext
	Loop
	Response.Write "</select></label>"
End Sub

Sub AgendaExtSelectEstados(cn, sel)
	Dim rs
	Response.Write "<label>Estado<select name='estado' required>"
	Set rs = Exec(cn, "sp_cg_AgendaExt @Modo='estados_listar'")
	Do While Not rs.EOF
		Response.Write "<option value='" & Server.HTMLEncode(rs(1) & "") & "'" & AgendaExtSelected(rs(1), sel) & ">" & Server.HTMLEncode(rs(1) & "") & "</option>"
		rs.MoveNext
	Loop
	Response.Write "</select></label>"
End Sub

Sub AgendaExtInput(labelText, fieldName, fieldValue, fieldType, required)
	Response.Write "<label>" & labelText & "<input name='" & fieldName & "' type='" & fieldType & "' value='" & Server.HTMLEncode(fieldValue & "") & "'"
	If required Then Response.Write " required"
	Response.Write "></label>"
End Sub

Function AgendaExtSelected(value, selectedValue)
	If CStr(value) = CStr(selectedValue) Then AgendaExtSelected = " selected" Else AgendaExtSelected = ""
End Function
Function AgendaExtChecked(value)
	If AgendaExtSqlBit(value) = "1" Then AgendaExtChecked = " checked" Else AgendaExtChecked = ""
End Function
Function AgendaExtOpcionesAccion(cn, selectedValue)
	Dim rs, html
	html = "<option value='0'>Seleccionar</option>"
	Set rs = Exec(cn, "sp_cg_AgendaExt @Modo='acciones_listar'")
	If Not rs Is Nothing Then
		If Not rs.State = 0 Then
			Do While Not rs.EOF
				html = html & "<option value='" & rs(0) & "'" & AgendaExtSelected(rs(0), selectedValue) & ">" & Server.HTMLEncode(rs(1) & "") & "</option>"
				rs.MoveNext
			Loop
		End If
	End If
	AgendaExtOpcionesAccion = html
End Function
Function AgendaExtSqlText(value)
	If Trim(value & "") = "" Then AgendaExtSqlText = "NULL" Else AgendaExtSqlText = "N'" & Replace(value & "", "'", "''") & "'"
End Function
' --- EPOCH: la base guarda fechas como DIAS desde 1900-01-01 (INT) y horas como HHMM (INT) ---
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
' Recibe "yyyy-mm-dd" del form y devuelve DIAS (int) para el SP
Function AgendaExtSqlDate(value)
	If IsDate(value) Then AgendaExtSqlDate = CStr(CLng(CDate(value)) - CLng(CDate("01-01-1900"))) Else AgendaExtSqlDate = "NULL"
End Function
' Recibe "hh:mm" del form y devuelve HHMM (int) para el SP
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
' Texto SQL (N'...') con comillas escapadas; si value viene vacio usa fallback. Para ID_SEG.
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
	If value = "1" Or UCase(value & "") = "SI" Or value = True Then AgendaExtSqlBit = "1" Else AgendaExtSqlBit = "0"
End Function
Function AgendaExtVal(rs, fieldName, fallback)
	If IsNull(rs(fieldName)) Then AgendaExtVal = fallback Else AgendaExtVal = rs(fieldName)
End Function
' DIAS (int) -> "yyyy-mm-dd" para <input type="date">
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
' HHMM (int) -> "hh:mm"
Function AgendaExtHtmlTime(value)
	Dim n
	AgendaExtHtmlTime = ""
	If IsNumeric(value) Then
		n = CLng(value)
		If n >= 0 Then AgendaExtHtmlTime = Right("0" & CStr(n \ 100), 2) & ":" & Right("0" & CStr(n Mod 100), 2)
	End If
End Function
Sub AgendaExtRedir(message, id, xmenu, xtmenu)
	Dim msgJs
	msgJs = Replace(message, "\", "\\")
	msgJs = Replace(msgJs, "'", "\'")
	msgJs = Replace(msgJs, vbCrLf, "\n")
	msgJs = Replace(msgJs, vbCr, "\n")
	msgJs = Replace(msgJs, vbLf, "\n")
	Response.Write "<script>alert('" & msgJs & "');window.location='agenda.asp?opc=det&id=" & id & "&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu) & "';</script>"
End Sub
Function AgendaExtCss()
	AgendaExtCss = ":root{--navy:#0e1949;--blue:#165dff;--blue-soft:#eef4ff;--ink:#24324a;--muted:#6d7b91;--line:#d8e0ec;--surface:#fff;--canvas:#f4f7fb;--shadow:0 8px 22px rgba(30,52,91,.07)}*{box-sizing:border-box}body{margin:0;background:var(--canvas);color:var(--ink);font:14px Segoe UI,Tahoma,sans-serif}button,input,select,textarea{font:inherit}a{color:#165dff}button{cursor:pointer}.app-header{background:var(--navy);color:#fff;padding:14px 28px;display:flex;align-items:center;justify-content:space-between;gap:20px}.brand{display:flex;align-items:center;gap:12px;font-size:16px}.brand-mark{width:36px;height:36px;border:2px solid #8ea9ff;border-radius:8px;display:grid;place-items:center;font-weight:800}.brand small{display:block;color:#bdc9ee;font-size:11px;margin-top:2px}.shell{max-width:1260px;margin:0 auto;padding:26px 28px 48px}.breadcrumb{color:var(--muted);font-size:12px;margin-bottom:8px}.page-head{display:flex;align-items:end;justify-content:space-between;gap:20px;margin-bottom:22px}h1{color:var(--navy);font-size:25px;margin:0}h2{color:var(--navy);font-size:15px;margin:0}.page-head p{color:var(--muted);margin:6px 0 0}.toolbar{display:flex;gap:8px;align-items:center;flex-wrap:wrap}.btn{display:inline-block;border:1px solid var(--line);border-radius:5px;padding:9px 13px;background:#fff;color:var(--ink);font-weight:600;text-decoration:none}.btn-primary{color:#fff;background:var(--blue);border-color:var(--blue)}.btn-secondary{background:var(--blue-soft);border-color:#b9ccff;color:#174cb9}.panel{background:var(--surface);border:1px solid var(--line);border-radius:7px;box-shadow:var(--shadow);margin-bottom:18px;overflow:hidden}.panel-head{padding:15px 18px;border-bottom:1px solid var(--line);display:flex;justify-content:space-between;align-items:center;gap:14px}.panel-head span,.muted{color:var(--muted);font-size:12px}.panel-body{padding:18px}.form-grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px}.form-grid .full{grid-column:1 / -1}label{display:flex;flex-direction:column;gap:6px;color:var(--muted);font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.04em}input,select,textarea{width:100%;border:1px solid #c8d3e3;border-radius:5px;background:#fff;color:var(--ink);padding:9px 10px;outline:0;font-weight:400;text-transform:none;letter-spacing:0}textarea{min-height:80px;resize:vertical}input:focus,select:focus,textarea:focus{border-color:var(--blue);box-shadow:0 0 0 3px rgba(22,93,255,.12)}.grid{display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:10px;padding:18px}.grid+button,.panel>button{margin:0 18px 14px}.panel>h2{padding:15px 18px;border-bottom:1px solid var(--line)}table{width:100%;border-collapse:collapse}.table-wrap{overflow-x:auto}th{background:#f1f5fa;color:#60708a;font-size:11px;text-transform:uppercase;letter-spacing:.04em;text-align:left;white-space:nowrap}th,td{padding:12px 14px;border-bottom:1px solid #e7ecf3;vertical-align:middle}tbody tr:hover{background:#f8fbff}.status{display:inline-flex;align-items:center;border-radius:4px;padding:4px 8px;background:#fff1cf;color:#a96800;font-size:11px;font-weight:700}.pagination{display:flex;justify-content:center;gap:8px;padding-top:14px}.chip{display:inline-flex;gap:8px;align-items:center;background:var(--blue-soft);color:#174cb9;border-radius:5px;padding:7px 9px;margin:3px}.chip small{color:var(--muted);font-size:11px;font-weight:400}.chip-remove{margin-left:4px;color:#8b0000;font-weight:700;text-decoration:none;line-height:1;border:none;background:transparent;cursor:pointer;font-size:12px}.chip-remove:hover{color:#d32f2f}.action-bar{display:flex;justify-content:space-between;gap:10px;align-items:center;padding-top:4px}.action-bar .right{display:flex;gap:8px}.actions{display:flex;justify-content:flex-end;gap:8px;margin:18px 0}.actions button{margin:0}@media(max-width:900px){.shell{padding:20px 14px 40px}.app-header{padding:12px 16px}.page-head{align-items:flex-start;flex-direction:column}.form-grid,.grid{grid-template-columns:1fr 1fr}}@media(max-width:560px){.form-grid,.grid{grid-template-columns:1fr}.form-grid .full{grid-column:auto}.action-bar,.action-bar .right{width:100%;flex-wrap:wrap}.action-bar .btn,.action-bar button{flex:1}.brand small{display:none}}"
End Function
Function AgendaExtJs()
	AgendaExtJs = "var corN=parseInt(document.getElementById('corresponsable_count').value||0),subN=parseInt(document.getElementById('subtema_count').value||0);function agendaExtAgregarCorresponsable(){var p=document.getElementById('corPersona'),n=document.getElementById('corNombre'),e=document.getElementById('corEmail'),c=document.getElementById('corCopia'),name=p.options[p.selectedIndex].text;if(!p.value&&!n.value.trim()){alert('Seleccione una persona o complete el nombre manual.');return}if(!p.value&&!e.value.trim()&&c.value==='1'){alert('El email es obligatorio para contactos manuales con copia.');return}corN++;document.getElementById('corresponsable_count').value=corN;var d=document.createElement('div');d.className='chip';var label=name==='Seleccionar'?n.value+' &lt;'+e.value+'&gt;':name;if(e.value.trim()&&name!=='Seleccionar')label+=' &lt;'+e.value+'&gt;';if(c.value==='1')label+=' <small>(copia)</small>';d.innerHTML=label;d.innerHTML+='<input type=hidden name=cor_persona_'+corN+' value='+p.value+'><input type=hidden name=cor_nombre_'+corN+' value='+n.value+'><input type=hidden name=cor_email_'+corN+' value='+e.value+'><input type=hidden name=cor_copia_'+corN+' value='+c.value+'><button type=button class=chip-remove onclick=this.parentNode.remove()>x</button>';document.getElementById('corresponsables').appendChild(d);p.value='';n.value='';e.value=''}function agendaExtAgregarSubtema(){subN++;document.getElementById('subtema_count').value=subN;var tr=document.createElement('tr');tr.innerHTML='<td><select name=sub_tipo_'+subN+'><option value=1>PROYECTO</option><option value=2>INCIDENTE</option><option value=3>CONSULTORIA</option><option value=4>OTRO</option></select></td><td><input name=sub_proyecto_'+subN+' type=number placeholder=ID proyecto><input name=sub_incidente_'+subN+' type=number placeholder=ID incidente></td><td><input name=sub_tema_'+subN+' required maxlength=500></td><td><input name=sub_tiempo_'+subN+' type=number min=0 value=0></td><td><input name=sub_seguimiento_'+subN+' type=checkbox value=1></td><td><select name=sub_accion_'+subN+'>'+document.getElementById('agendaAccionTemplate').innerHTML+'</select></td><td><button type=button onclick=this.parentNode.parentNode.remove()>Quitar</button></td>';document.querySelector('#subtemas tbody').appendChild(tr)}function agendaExtValidar(){var a=document.querySelector('[name=hora_inicio]').value,b=document.querySelector('[name=hora_fin]').value;if(a&&b&&b<=a){alert('La hora de fin debe ser posterior a la hora de inicio.');return false}return true}"
End Function
Function AgendaExtRelacionJs()
	AgendaExtRelacionJs = "function agendaExtCambiarRelacion(sel){var box=sel.parentNode.parentNode.querySelector('.relation-fields'),input=box.querySelector('.relation-id'),label=box.querySelector('.relation-id-label'),buscar=box.querySelector('.btn-buscar-entidad');var visible=sel.value==='1'||sel.value==='2';box.style.display=visible?'block':'none';label.firstChild.nodeValue=sel.value==='2'?'ID incidente':'ID proyecto';if(buscar){buscar.style.display=visible?'inline-block':'none';}if(!visible){input.value='';}}function agendaExtBuscarEntidad(boton){var input=boton.parentNode.querySelector('.relation-id'),tipo=boton.closest('tr').querySelector('select[name^=sub_tipo_]').value,tabla=tipo==='1'?'PROYECTO':'INCIDENTE';window.open('buscar_entidades.asp?tabla='+tabla+'&campo='+encodeURIComponent(input.name),'buscar_entidad','width=900,height=650,resizable=yes,scrollbars=yes');}function agendaExtSeleccionarEntidad(campo,id,nombre){var input=document.querySelector('input[name='+JSON.stringify(campo)+']');if(input){input.value=id;input.title=nombre;}}function agendaExtAgregarSubtema(){subN++;document.getElementById('subtema_count').value=subN;var tr=document.createElement('tr');tr.innerHTML='<td><select name=sub_tipo_'+subN+' onchange=agendaExtCambiarRelacion(this)><option value=1>PROYECTO</option><option value=2>INCIDENTE</option><option value=3>CONSULTORIA</option><option value=4>OTRO</option></select></td><td><div class=relation-fields><label class=relation-id-label>ID proyecto<input class=relation-id name=sub_relacion_id_'+subN+' type=number min=1 placeholder=ID><button type=button class=btn-buscar-entidad onclick=agendaExtBuscarEntidad(this)>Buscar</button></label></div></td><td><input name=sub_tema_'+subN+' required maxlength=500></td><td><textarea name=sub_conclusion_'+subN+' rows=2 maxlength=4000></textarea></td><td><input name=sub_tiempo_'+subN+' type=number min=0 value=0></td><td><input name=sub_seguimiento_'+subN+' type=checkbox value=1></td><td><select name=sub_accion_'+subN+'>'+document.getElementById('agendaAccionTemplate').innerHTML+'</select></td><td><button type=button onclick=this.parentNode.parentNode.remove()>Quitar</button></td>';document.querySelector('#subtemas tbody').appendChild(tr);agendaExtCambiarRelacion(tr.querySelector('select'));}document.querySelectorAll('#subtemas select[name^=sub_tipo_]').forEach(agendaExtCambiarRelacion);"
End Function
%>