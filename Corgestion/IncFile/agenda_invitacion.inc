<%
' =============================================================================
' agenda_invitacion.inc - Invitacion de la reunion por mail (.ics) y enlaces a calendarios
'
' Lo incluye agenda.asp. Lo usa agenda_ext.inc (pantalla de la reunion, opc=det):
'   AgInvCardHtml(cn, id, base)   tarjeta "Invitacion": estado, destinatarios, enlaces a calendarios y boton de envio
'   AgendaExtInvitar(cn, id)      opc=agenda_ext_invitar: arma y envia el mail con el .ics adjunto
' Requiere DB/Cambios_Agenda_Invitacion.sql (sp_cg_AgendaExt modos invitacion_datos / invitacion_registrar).
'
' Reglas:
'   - Destinatarios: TODOS los participantes con mail (del sistema o manuales, tengan o no marcada la copia) y el responsable.
'   - Remitente y organizador: el responsable de la reunion; si no tiene mail (OBJALOG.MAIL), AGINV_REMITENTE_DEFECTO.
'   - Hora local de Argentina (UTC-3, sin horario de verano): el .ics va en UTC.
'   - Reenvio: mismo UID y SEQUENCE + 1, el calendario actualiza el evento en lugar de duplicarlo.
'   - Prueba: si se completa "Enviar solo a", todo va a esa direccion y NO se registra el envio.
'   - SMTP: Session("SMTPServer") / Session("SMTPPort"), igual que SendMail.
' =============================================================================

Const AGINV_REMITENTE_DEFECTO = "soptec@corconsulting.com.ar"
Const AGINV_UTC_HORAS = 3
Dim gAgInvPaso
gAgInvPaso = ""

Function AgInvPad(n, l)
	AgInvPad = Right(String(l, "0") & CStr(n), l)
End Function

' Fecha (dias desde 1900-01-01) + hora (HHMM) -> Date local
Function AgInvFechaHora(fechaInt, horaInt)
	Dim d
	d = DateAdd("d", CLng(fechaInt), DateSerial(1900, 1, 1))
	AgInvFechaHora = DateAdd("n", (CLng(horaInt) \ 100) * 60 + (CLng(horaInt) Mod 100), d)
End Function

Function AgInvUtc(dt)
	AgInvUtc = DateAdd("h", AGINV_UTC_HORAS, dt)
End Function

' yyyymmddThhmmssZ (dt ya en UTC)
Function AgInvFmtIcs(dt)
	AgInvFmtIcs = Year(dt) & AgInvPad(Month(dt), 2) & AgInvPad(Day(dt), 2) & "T" & AgInvPad(Hour(dt), 2) & AgInvPad(Minute(dt), 2) & AgInvPad(Second(dt), 2) & "Z"
End Function

' yyyy-mm-ddThh:mm:ssZ (dt ya en UTC)
Function AgInvFmtIso(dt)
	AgInvFmtIso = Year(dt) & "-" & AgInvPad(Month(dt), 2) & "-" & AgInvPad(Day(dt), 2) & "T" & AgInvPad(Hour(dt), 2) & ":" & AgInvPad(Minute(dt), 2) & ":" & AgInvPad(Second(dt), 2) & "Z"
End Function

Function AgInvEnc(s)
	AgInvEnc = Replace(Server.URLEncode(s & ""), "+", "%20")
End Function

Function AgInvEsMail(m)
	Dim p
	AgInvEsMail = False
	m = Trim(m & "")
	p = InStrRev(m, "@")
	If p < 2 Then Exit Function
	If InStr(m, " ") > 0 Then Exit Function
	If InStr(p, m, ".") = 0 Then Exit Function
	AgInvEsMail = True
End Function

' Texto de una propiedad iCalendar: escapa \ ; , y saltos de linea
Function AgInvEscIcs(s)
	s = Replace(s & "", "\", "\\")
	s = Replace(s, ";", "\;")
	s = Replace(s, ",", "\,")
	s = Replace(s, vbCrLf, "\n")
	s = Replace(s, vbCr, "\n")
	s = Replace(s, vbLf, "\n")
	AgInvEscIcs = s
End Function

' Lineas de mas de 60 caracteres: se continuan en la siguiente con un espacio inicial (RFC 5545)
Function AgInvFold(linea)
	Dim r, resto
	If Len(linea) <= 60 Then
		AgInvFold = linea
		Exit Function
	End If
	r = Left(linea, 60)
	resto = Mid(linea, 61)
	Do While Len(resto) > 0
		r = r & vbCrLf & " " & Left(resto, 59)
		resto = Mid(resto, 60)
	Loop
	AgInvFold = r
End Function

' Datos de la reunion (sp_cg_AgendaExt invitacion_datos) en un Dictionary; Nothing si no existe
Function AgInvCargar(cn, id)
	Dim rs, d, f, hi, hf
	Set AgInvCargar = Nothing
	gAgInvPaso = "cargar: sp invitacion_datos"
	Set rs = Exec(cn, "sp_cg_AgendaExt @Modo='invitacion_datos', @Id=" & CLng(id))
	If rs.EOF Then
		rs.Close
		Exit Function
	End If
	Set d = Server.CreateObject("Scripting.Dictionary")
	gAgInvPaso = "cargar: leer campos"
	d("asunto") = Trim(rs("ASUNTO") & "")
	d("enlace") = Trim(rs("ENLACE_REUNION") & "")
	d("resp") = Trim(rs("RESPONSABLE") & "")
	d("respMail") = Trim(rs("MAIL_RESPONSABLE") & "")
	d("usuario") = Trim(rs("ID_RESPON") & "")
	d("seq") = CLng("0" & rs("SEQ_INVITACION"))
	d("ult") = ""
	If Not IsNull(rs("FEC_ULT_ENVIO")) Then d("ult") = Right("0" & Day(rs("FEC_ULT_ENVIO")), 2) & "/" & Right("0" & Month(rs("FEC_ULT_ENVIO")), 2) & "/" & Year(rs("FEC_ULT_ENVIO")) & " " & Right("0" & Hour(rs("FEC_ULT_ENVIO")), 2) & ":" & Right("0" & Minute(rs("FEC_ULT_ENVIO")), 2)
	d("ultUsu") = Trim(rs("USU_ULT_ENVIO") & "")
	gAgInvPaso = "cargar: fechas"
	d("valido") = False
	f = rs("FECHA_REUNION")
	hi = rs("HORA_INICIO")
	hf = rs("HORA_FIN")
	If Not IsNull(f) And Not IsNull(hi) And Not IsNull(hf) Then
		d("ini") = AgInvFechaHora(f, hi)
		d("fin") = AgInvFechaHora(f, hf)
		If d("fin") > d("ini") Then d("valido") = True
	End If
	d("mensaje") = Trim(rs("MENSAJE") & "")   ' campo largo (nvarchar(max)): se lee al final
	rs.Close
	Set AgInvCargar = d
End Function

' Destinatarios: Dictionary mail(minusculas) -> nombre. 'faltan' (ByRef) lista a quienes no tienen mail cargado
Function AgInvDestinatarios(cn, id, d, faltan)
	Dim dest, rs, nombre, mail, k
	gAgInvPaso = "destinatarios"
	Set dest = Server.CreateObject("Scripting.Dictionary")
	faltan = ""
	If AgInvEsMail(d("respMail")) Then dest(LCase(d("respMail"))) = d("resp")
	Set rs = Exec(cn, "sp_cg_AgendaExt @Modo='corresp_listar', @Id=" & CLng(id))
	Do While Not rs.EOF
		nombre = Trim(rs("NOMBRE_PERSONA") & "")
		If nombre = "" Then nombre = Trim(rs("NOMBRE_MANUAL") & "")
		mail = Trim(rs("EMAIL_MANUAL") & "")
		If AgInvEsMail(mail) Then
			k = LCase(mail)
			If Not dest.Exists(k) Then dest(k) = nombre
		Else
			If nombre <> "" Then faltan = faltan & nombre & "; "
		End If
		rs.MoveNext
	Loop
	rs.Close
	Set AgInvDestinatarios = dest
End Function

Function AgInvRemitenteMail(d)
	If AgInvEsMail(d("respMail")) Then AgInvRemitenteMail = d("respMail") Else AgInvRemitenteMail = AGINV_REMITENTE_DEFECTO
End Function

Function AgInvRemitenteNombre(d)
	If d("resp") <> "" Then AgInvRemitenteNombre = Replace(d("resp"), """", "") Else AgInvRemitenteNombre = "Corgestion"
End Function

' Detalle de texto plano (para calendarios): mensaje y enlace
Function AgInvDetalle(d)
	Dim r
	r = d("mensaje")
	If d("enlace") <> "" Then
		If r <> "" Then r = r & vbLf & vbLf
		r = r & d("enlace")
	End If
	AgInvDetalle = r
End Function

' Contenido del archivo .ics (METHOD:REQUEST). seqNuevo: SEQUENCE del evento
Function AgInvIcs(id, d, dest, seqNuevo)
	Dim s, k, mailOrg, nomOrg, ahora
	mailOrg = AgInvRemitenteMail(d)
	nomOrg = AgInvRemitenteNombre(d)
	ahora = AgInvUtc(Now())
	s = "BEGIN:VCALENDAR" & vbCrLf & "PRODID:-//Corgestion//Agenda//ES" & vbCrLf & "VERSION:2.0" & vbCrLf & "CALSCALE:GREGORIAN" & vbCrLf & "METHOD:REQUEST" & vbCrLf
	s = s & "BEGIN:VEVENT" & vbCrLf
	s = s & "UID:cg-agenda-" & CLng(id) & "@corgestion" & vbCrLf
	s = s & "DTSTAMP:" & AgInvFmtIcs(ahora) & vbCrLf
	s = s & "SEQUENCE:" & CLng(seqNuevo) & vbCrLf
	s = s & "DTSTART:" & AgInvFmtIcs(AgInvUtc(d("ini"))) & vbCrLf
	s = s & "DTEND:" & AgInvFmtIcs(AgInvUtc(d("fin"))) & vbCrLf
	s = s & AgInvFold("SUMMARY:" & AgInvEscIcs(d("asunto"))) & vbCrLf
	s = s & AgInvFold("DESCRIPTION:" & AgInvEscIcs(AgInvDetalle(d))) & vbCrLf
	If d("enlace") <> "" Then
		s = s & AgInvFold("LOCATION:" & AgInvEscIcs(d("enlace"))) & vbCrLf
		s = s & AgInvFold("URL:" & d("enlace")) & vbCrLf
	End If
	s = s & AgInvFold("ORGANIZER;CN=" & AgInvEscIcs(nomOrg) & ":mailto:" & mailOrg) & vbCrLf
	For Each k In dest.Keys
		If k <> LCase(mailOrg) Then
			s = s & AgInvFold("ATTENDEE;CUTYPE=INDIVIDUAL;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE;CN=" & AgInvEscIcs(dest(k)) & ":mailto:" & k) & vbCrLf
		End If
	Next
	s = s & "STATUS:CONFIRMED" & vbCrLf & "TRANSP:OPAQUE" & vbCrLf & "END:VEVENT" & vbCrLf & "END:VCALENDAR" & vbCrLf
	AgInvIcs = s
End Function

' Cuerpo HTML del mail (los textos fijos usan entidades HTML)
Function AgInvHtml(d, esActualizacion)
	Dim h, ini, fin
	ini = d("ini")
	fin = d("fin")
	h = "<div style='font-family:Arial,sans-serif;font-size:14px;color:#222'>"
	If esActualizacion Then h = h & "<p><b>Actualizaci&oacute;n de la reuni&oacute;n:</b> cambi&oacute; alg&uacute;n dato; el evento de su calendario se actualiza.</p>"
	h = h & "<p style='font-size:16px'><b>" & Server.HTMLEncode(d("asunto")) & "</b></p>"
	h = h & "<p><b>Fecha:</b> " & Right("0" & Day(ini), 2) & "/" & Right("0" & Month(ini), 2) & "/" & Year(ini) & "<br>"
	h = h & "<b>Horario:</b> " & AgInvPad(Hour(ini), 2) & ":" & AgInvPad(Minute(ini), 2) & " a " & AgInvPad(Hour(fin), 2) & ":" & AgInvPad(Minute(fin), 2) & " (hora de Argentina)<br>"
	h = h & "<b>Organiza:</b> " & Server.HTMLEncode(AgInvRemitenteNombre(d)) & "</p>"
	If d("enlace") <> "" Then h = h & "<p><b>Enlace de la reuni&oacute;n:</b> <a href='" & Server.HTMLEncode(d("enlace")) & "'>" & Server.HTMLEncode(d("enlace")) & "</a></p>"
	If d("mensaje") <> "" Then h = h & "<p>" & Replace(Server.HTMLEncode(d("mensaje")), vbLf, "<br>") & "</p>"
	h = h & "<p style='color:#666;font-size:12px'>Este mensaje incluye un archivo de invitaci&oacute;n (.ics): abralo para aceptar o rechazar la reuni&oacute;n en su calendario.</p>"
	h = h & "</div>"
	AgInvHtml = h
End Function

' Envia la invitacion. mailPrueba <> "" => todo a esa direccion y sin registrar el envio.
' Devuelve el texto del resultado (texto plano); ok (ByRef) indica si se envio.
Function AgInvEnviar(cn, id, mailPrueba, ok)
	Dim d, dest, faltan, seqNuevo, ics, fso, carpeta, ruta, st, st2, msg, att, para, k, asunto, esAct, errTxt, destPrueba
	ok = False
	Set d = AgInvCargar(cn, id)
	If d Is Nothing Then
		AgInvEnviar = "La reuni" & ChrW(243) & "n no existe o est" & ChrW(225) & " dada de baja."
		Exit Function
	End If
	If Not d("valido") Then
		AgInvEnviar = "La reuni" & ChrW(243) & "n no tiene fecha y horario v" & ChrW(225) & "lidos (la hora de fin debe ser posterior a la de inicio)."
		Exit Function
	End If
	If Trim(Session("SMTPServer") & "") = "" Then
		AgInvEnviar = "No hay servidor de correo configurado (Session SMTPServer). Consulte con soporte."
		Exit Function
	End If

	Set dest = AgInvDestinatarios(cn, id, d, faltan)
	If Trim(mailPrueba & "") <> "" Then
		If Not AgInvEsMail(mailPrueba) Then
			AgInvEnviar = "La direcci" & ChrW(243) & "n de prueba no es v" & ChrW(225) & "lida."
			Exit Function
		End If
		Set destPrueba = Server.CreateObject("Scripting.Dictionary")
		destPrueba(LCase(Trim(mailPrueba))) = "Prueba"
		Set dest = destPrueba
	End If
	If dest.Count = 0 Then
		AgInvEnviar = "No hay destinatarios con mail. Cargue el mail de los participantes."
		Exit Function
	End If

	seqNuevo = d("seq") + 1
	esAct = (d("seq") > 0)
	ics = AgInvIcs(id, d, dest, seqNuevo)
	asunto = "Invitaci" & ChrW(243) & "n: " & d("asunto")
	If esAct Then asunto = "Actualizaci" & ChrW(243) & "n: " & d("asunto")
	If Trim(mailPrueba & "") <> "" Then asunto = "[PRUEBA] " & asunto

	para = ""
	For Each k In dest.Keys
		If para <> "" Then para = para & ","
		para = para & k
	Next

	On Error Resume Next
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	carpeta = fso.GetSpecialFolder(2) & "\cgInv" & CLng(id) & "_" & CLng(Timer() * 100)
	fso.CreateFolder carpeta
	ruta = carpeta & "\invitacion.ics"
	Set st = Server.CreateObject("ADODB.Stream")
	st.Type = 2
	st.Charset = "utf-8"
	st.Open
	st.WriteText ics
	st.Position = 0
	st.Type = 1
	st.Position = 3
	Set st2 = Server.CreateObject("ADODB.Stream")
	st2.Type = 1
	st2.Open
	st.CopyTo st2
	st2.SaveToFile ruta, 2
	st2.Close
	st.Close

	Set msg = Server.CreateObject("CDO.Message")
	Dim cfg
	Set cfg = Server.CreateObject("CDO.Configuration")
	cfg.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = Session("SMTPServer")
	cfg.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = Session("SMTPPort")
	cfg.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	cfg.Fields("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
	cfg.Fields.Update
	Set msg.Configuration = cfg
	msg.From = """" & AgInvRemitenteNombre(d) & """ <" & AgInvRemitenteMail(d) & ">"
	msg.To = para
	msg.Subject = asunto
	msg.HTMLBody = AgInvHtml(d, esAct)
	msg.BodyPart.Charset = "utf-8"
	msg.HTMLBodyPart.Charset = "utf-8"
	Set att = msg.AddAttachment(ruta)
	att.Fields("urn:schemas:mailheader:content-type") = "text/calendar; method=REQUEST; charset=utf-8; name=""invitacion.ics"""
	att.Fields.Update
	msg.Send
	errTxt = ""
	If Err.Number <> 0 Then errTxt = Err.Description
	Err.Clear
	fso.DeleteFolder carpeta, True
	Err.Clear
	On Error Goto 0

	If errTxt <> "" Then
		AgInvEnviar = "No se pudo enviar el mail: " & errTxt
		Exit Function
	End If

	If Trim(mailPrueba & "") = "" Then
		ExecNonQuery cn, "sp_cg_AgendaExt @Modo='invitacion_registrar', @Id=" & CLng(id) & ", @Usuario=N'" & Replace(Left(Trim(Session("usuario") & ""), 100), "'", "''") & "'"
	End If
	ok = True
	If Trim(mailPrueba & "") <> "" Then
		AgInvEnviar = "Invitaci" & ChrW(243) & "n de PRUEBA enviada a " & Trim(mailPrueba) & ". No se registr" & ChrW(243) & " como enviada."
	Else
		If esAct Then errTxt = "Actualizaci" & ChrW(243) & "n enviada a " & dest.Count & " destinatario(s)." Else errTxt = "Invitaci" & ChrW(243) & "n enviada a " & dest.Count & " destinatario(s)."
		If faltan <> "" Then errTxt = errTxt & " Sin mail cargado (no recibieron): " & faltan
		AgInvEnviar = errTxt
	End If
End Function

' opc=agenda_ext_invitar (POST): permiso de modificar, envio y mensaje para la pantalla
Sub AgendaExtInvitar(cn, id)
	Dim ok, txt
	ok = False
	If getPermisosAtrm("CG_AGENDA", "MODIF") <> 1 Then
		txt = "Sin permiso para enviar invitaciones."
	Else
		txt = AgInvEnviar(cn, CLng(id), Request.Form("prueba"), ok)
	End If
	Session("agendaExtInvTxt") = txt
	If ok Then Session("agendaExtInvOk") = "1" Else Session("agendaExtInvOk") = "0"
End Sub

' Mensaje del ultimo envio (se muestra una vez)
Function AgInvMensajeHtml()
	Dim txt
	txt = Session("agendaExtInvTxt") & ""
	If txt = "" Then
		AgInvMensajeHtml = ""
		Exit Function
	End If
	If Session("agendaExtInvOk") = "1" Then
		AgInvMensajeHtml = "<div class='ui-alert ui-alert--success' role='status'>" & Server.HTMLEncode(txt) & "</div>"
	Else
		AgInvMensajeHtml = "<div class='ui-alert ui-alert--danger' role='alert'>" & Server.HTMLEncode(txt) & "</div>"
	End If
	Session("agendaExtInvTxt") = ""
End Function

' Tarjeta "Invitacion" de la pantalla de la reunion (solo con la reunion ya guardada).
' Si algo falla NO se cae la pantalla de la reunion: se muestra el aviso con el paso en que fallo.
Function AgInvCardHtml(cn, id, base)
	Dim r
	gAgInvPaso = "inicio"
	On Error Resume Next
	r = AgInvCardHtmlX(cn, id, base)
	If Err.Number <> 0 Then
		r = "<div class='ui-alert ui-alert--warning' role='alert'>No se pudo preparar la invitaci&oacute;n (paso: " & Server.HTMLEncode(gAgInvPaso) & "). Error " & Err.Number & ": " & Server.HTMLEncode(Err.Description) & "</div>"
		Err.Clear
	End If
	On Error Goto 0
	AgInvCardHtml = r
End Function

Function AgInvCardHtmlX(cn, id, base)
	Dim d, dest, faltan, h, k, lista, mails, urlG, urlO, urlT, detalle, puede, n, confirma, ini, fin
	AgInvCardHtmlX = ""
	Set d = AgInvCargar(cn, id)
	If d Is Nothing Then Exit Function
	gAgInvPaso = "permiso"
	puede = (getPermisosAtrm("CG_AGENDA", "MODIF") = 1)
	Set dest = AgInvDestinatarios(cn, id, d, faltan)

	h = UiCardAbre("Invitaci&oacute;n por mail y calendario", "")
	If Not d("valido") Then
		h = h & "<div class='ui-alert ui-alert--warning' role='note'>La reuni&oacute;n necesita fecha y horario v&aacute;lidos (hora de fin posterior a la de inicio) para enviar la invitaci&oacute;n.</div>"
		AgInvCardHtmlX = h & UiCardCierra()
		Exit Function
	End If

	If d("ult") <> "" Then
		h = h & "<p>Invitaci&oacute;n enviada el <strong>" & UiTexto(d("ult")) & "</strong>"
		If d("ultUsu") <> "" Then h = h & " por " & UiTexto(d("ultUsu"))
		h = h & ". Si vuelve a enviarla, se actualiza el mismo evento en los calendarios.</p>"
	Else
		h = h & "<p>La invitaci&oacute;n todav&iacute;a no se envi&oacute;. Guarde los cambios de la reuni&oacute;n antes de enviarla.</p>"
	End If

	gAgInvPaso = "lista de destinatarios"
	lista = ""
	mails = ""
	For Each k In dest.Keys
		lista = lista & "<li>" & UiTexto(dest(k)) & " <span class='ui-participante__mail'>" & UiTexto(k) & "</span></li>"
		If mails <> "" Then mails = mails & ","
		mails = mails & k
	Next
	If lista <> "" Then h = h & "<p>Recibir&aacute;n la invitaci&oacute;n (" & dest.Count & "):</p><ul class='ui-alert__lista'>" & lista & "</ul>" Else h = h & "<div class='ui-alert ui-alert--warning' role='note'>No hay destinatarios con mail cargado.</div>"
	If faltan <> "" Then h = h & "<div class='ui-alert ui-alert--warning' role='note'>Sin mail cargado (no recibir&aacute;n la invitaci&oacute;n): " & UiTexto(faltan) & "</div>"
	h = h & "<p>Remitente: " & UiTexto(AgInvRemitenteMail(d)) & "</p>"

	' Enlaces a calendarios (cada usuario agrega el evento al suyo)
	gAgInvPaso = "enlaces de calendario"
	ini = AgInvUtc(d("ini"))
	fin = AgInvUtc(d("fin"))
	detalle = AgInvDetalle(d)
	urlG = "https://calendar.google.com/calendar/render?action=TEMPLATE&text=" & AgInvEnc(d("asunto")) & "&dates=" & AgInvFmtIcs(ini) & "/" & AgInvFmtIcs(fin) & "&details=" & AgInvEnc(detalle) & "&location=" & AgInvEnc(d("enlace")) & "&add=" & AgInvEnc(mails)
	urlO = "https://outlook.office.com/calendar/0/deeplink/compose?path=%2Fcalendar%2Faction%2Fcompose&rru=addevent&subject=" & AgInvEnc(d("asunto")) & "&startdt=" & AgInvEnc(AgInvFmtIso(ini)) & "&enddt=" & AgInvEnc(AgInvFmtIso(fin)) & "&body=" & AgInvEnc(detalle) & "&location=" & AgInvEnc(d("enlace")) & "&to=" & AgInvEnc(Replace(mails, ",", ";"))
	urlT = "https://teams.microsoft.com/l/meeting/new?subject=" & AgInvEnc(d("asunto")) & "&startTime=" & AgInvEnc(AgInvFmtIso(ini)) & "&endTime=" & AgInvEnc(AgInvFmtIso(fin)) & "&attendees=" & AgInvEnc(mails) & "&content=" & AgInvEnc(detalle)
	h = h & "<p>Agregar al calendario propio:</p><div class='ui-form__acciones'>"
	h = h & "<a class='ui-button ui-button--light ui-button--sm' href='" & Server.HTMLEncode(urlG) & "' target='_blank' rel='noopener'>Google Calendar</a> "
	h = h & "<a class='ui-button ui-button--light ui-button--sm' href='" & Server.HTMLEncode(urlO) & "' target='_blank' rel='noopener'>Outlook</a> "
	h = h & "<a class='ui-button ui-button--light ui-button--sm' href='" & Server.HTMLEncode(urlT) & "' target='_blank' rel='noopener'>Teams</a></div>"
	h = h & "<p class='ui-dialogo__ayuda'>Con Teams se crea la reuni&oacute;n all&iacute;; luego pegue el enlace generado en el campo &quot;Enlace de la reuni&oacute;n&quot; y guarde.</p>"

	' Envio
	gAgInvPaso = "formulario de envio"
	n = dest.Count
	If d("ult") <> "" Then
		confirma = "Ya se envi\u00f3 antes. Se enviar\u00e1 una actualizaci\u00f3n a " & n & " destinatario(s). \u00bfContinuar?"
	Else
		confirma = "Se enviar\u00e1 la invitaci\u00f3n a " & n & " destinatario(s). \u00bfContinuar?"
	End If
	h = h & "<form method='post' action='agenda.asp?opc=agenda_ext_invitar&id=" & CLng(id) & base & "' onsubmit=""return confirm('" & confirma & "');"">"
	h = h & UiFormGridAbre()
	h = h & UiCampo("prueba", "Enviar solo a (prueba, opcional)", "<input class='ui-input' id='prueba' name='prueba' type='email' maxlength='150' placeholder='su.mail@empresa.com'>", False, True)
	h = h & UiFormGridCierra()
	h = h & UiFormAcciones(UiBotonEnviar("Enviar invitaci&oacute;n", puede And n > 0, "Sin permiso de modificaci&oacute;n o sin destinatarios con mail."))
	h = h & "</form>"
	AgInvCardHtmlX = h & UiCardCierra()
End Function
%>
