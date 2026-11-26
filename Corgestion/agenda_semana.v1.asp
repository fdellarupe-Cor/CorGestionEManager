<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/agenda_ext.inc" -->
<%
' ============================================================================
' agenda_semana.asp
' Pantalla flotante (popup) "Agenda - Vista semanal": calendario tipo
' Outlook/Teams con las reuniones de la semana y sus participantes.
' Para cada reunion PENDIENTE/POSTERGADA valida cruces de horario llamando a
' dbo.sp_Agenda_ValidarCruceHorario por cada responsable/participante del
' sistema.
' Parametros: d = fecha de referencia (DIAS desde 1900-01-01, mismo formato
' que AGENDA.FECHA_REUNION); xmenu/xtmenu = contexto de menu para volver a
' agenda.asp. Se abre con window.open() desde agenda_ext.inc.
' ============================================================================

Dim i, j
Dim listaPart(), nPart, dicP
Dim personas(), nPersonas, k, listaP
Dim xmenu, xtmenu
xmenu = Request.QueryString("xmenu") & ""
xtmenu = Request.QueryString("xtmenu") & ""

Dim Cnn
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

' --- Semana a mostrar (lunes a viernes) a partir del parametro "d" ---
Dim EPOCH, epochRef, fechaRefDate, diaSemana, epochLunes, epochDesde, epochHasta, hoyEpoch
EPOCH = CDate("01-01-1900")
hoyEpoch = AgendaExtHoy()
' OJO: IsNumeric(Empty) = True en VBScript. Sin el "& """ un "d" ausente
' daba epochRef = 0 y la vista mostraba la semana del 1/1/1900.
Dim dParam
dParam = Trim(Request.QueryString("d") & "")
If dParam <> "" And IsNumeric(dParam) Then
	epochRef = CLng(dParam)
Else
	epochRef = hoyEpoch
End If
fechaRefDate = CDate(CLng(epochRef) + CLng(EPOCH))
diaSemana = Weekday(fechaRefDate, vbMonday) ' 1=lunes .. 7=domingo
epochLunes = epochRef - (diaSemana - 1)
epochDesde = epochLunes
epochHasta = epochLunes + 4 ' lunes..viernes

Dim NOMBRES_DIA(4)
NOMBRES_DIA(0) = "LUN" : NOMBRES_DIA(1) = "MAR" : NOMBRES_DIA(2) = "MI&Eacute;" : NOMBRES_DIA(3) = "JUE" : NOMBRES_DIA(4) = "VIE"

' --- Helpers ---
Function HHMMaDecimal(v)
	If Not IsNumeric(v) Then
		HHMMaDecimal = 0
	Else
		HHMMaDecimal = Int(CLng(v) / 100) + (CLng(v) Mod 100) / 60
	End If
End Function

Function FormatHoraDecimal(v)
	Dim hh, mm
	hh = Int(v)
	mm = Round((v - hh) * 60)
	FormatHoraDecimal = Right("0" & hh, 2) & ":" & Right("0" & mm, 2)
End Function

' Numero para CSS con punto decimal. CStr() usa el separador del idioma del
' servidor (coma en espanol) y "18,43px" es CSS invalido.
Function CssNum(v)
	CssNum = Replace(CStr(Round(CDbl(v), 2)), ",", ".")
End Function

Function JsStr(s)
	Dim t
	t = s & ""
	t = Replace(t, "\", "\\")
	t = Replace(t, "'", "\'")
	t = Replace(t, vbCrLf, "\n")
	t = Replace(t, vbCr, "\n")
	t = Replace(t, vbLf, "\n")
	JsStr = t
End Function

Dim PALETA(7)
PALETA(0) = "#0f6cbd" : PALETA(1) = "#8764b8" : PALETA(2) = "#107c10" : PALETA(3) = "#ca5010"
PALETA(4) = "#498205" : PALETA(5) = "#c239b3" : PALETA(6) = "#986f0b" : PALETA(7) = "#005e50"

Function ColorPersona(idPersona)
	Dim s, h, j
	s = Trim(idPersona & "")
	If s = "" Or s = "0" Then
		ColorPersona = "#8a8886"
	Else
		h = 0
		For j = 1 To Len(s)
			h = (h + Asc(Mid(s, j, 1))) Mod 8
		Next
		ColorPersona = PALETA(h)
	End If
End Function

Function Iniciales(nombre)
	Dim partes, ini, i
	partes = Split(Trim(nombre & ""), " ")
	ini = ""
	For i = 0 To UBound(partes)
		If Trim(partes(i)) <> "" And Len(ini) < 2 Then ini = ini & UCase(Left(partes(i), 1))
	Next
	If ini = "" Then ini = "?"
	Iniciales = ini
End Function

' Clase CSS de estado (definida en assets/css/corgestion-ui.css): resuelve
' fondo, texto y borde izquierdo de la reunion en la grilla.
Function EstadoClase(estado)
	Select Case UCase(estado & "")
		Case "PENDIENTE": EstadoClase = "ui-week-reunion--pendiente"
		Case "POSTERGADA": EstadoClase = "ui-week-reunion--postergada"
		Case "REALIZADA": EstadoClase = "ui-week-reunion--realizada"
		Case Else: EstadoClase = "ui-week-reunion--realizada"
	End Select
End Function

' Color de texto (hex) del estado, usado solo en el panel lateral (JS).
Function EstadoColorTextoJs(estado)
	Select Case UCase(estado & "")
		Case "PENDIENTE": EstadoColorTextoJs = "#0f4778"
		Case "POSTERGADA": EstadoColorTextoJs = "#6b4e00"
		Case "REALIZADA": EstadoColorTextoJs = "#484644"
		Case Else: EstadoColorTextoJs = "#484644"
	End Select
End Function

' --- Trae reuniones + corresponsables de la semana (2 resultsets) ---
' IMPORTANTE: hay que agotar el primer resultset (reuniones) ANTES de pedir
' NextRecordset(), porque ADO descarta las filas no leidas del resultset
' actual al avanzar al siguiente.
Dim rsReu, rsCorr
Set rsReu = Exec(Cnn, "sp_cg_Agenda_VistaSemanal @FechaDesde=" & epochDesde & ", @FechaHasta=" & epochHasta)

' --- Arma la coleccion de reuniones (Dictionary por reunion), sin participantes todavia ---
Dim dicReuniones, ordenIds, totalReuniones
Set dicReuniones = Server.CreateObject("Scripting.Dictionary")
totalReuniones = 0
ReDim ordenIds(2000)

Dim HORA_START, HORA_END
HORA_START = 8
HORA_END = 20

Do While Not rsReu.EOF
	Dim m, idReu
	Set m = Server.CreateObject("Scripting.Dictionary")
	idReu = CLng(rsReu("ID"))
	m.Add "id", idReu
	m.Add "asunto", rsReu("ASUNTO") & ""
	m.Add "fecha", CLng(rsReu("FECHA_REUNION"))
	m.Add "hi", HHMMaDecimal(rsReu("HORA_INICIO"))
	m.Add "hf", HHMMaDecimal(rsReu("HORA_FIN"))
	m.Add "estado", rsReu("ESTADO_AGENDA") & ""
	m.Add "respId", rsReu("ID_RESPONSABLE") & ""
	m.Add "respNombre", rsReu("RESPONSABLE") & ""
	If Trim(m("respNombre")) = "" Then m("respNombre") = "(sin responsable)"
	m.Add "hiInt", CLng(rsReu("HORA_INICIO"))
	m.Add "hfInt", CLng(rsReu("HORA_FIN"))

	If m("hi") < HORA_START Then HORA_START = Int(m("hi"))
	Dim techo
	techo = Int(m("hf"))
	If m("hf") > techo Then techo = techo + 1
	If techo > HORA_END Then HORA_END = techo

	dicReuniones.Add CStr(idReu), m
	ordenIds(totalReuniones) = CStr(idReu)
	totalReuniones = totalReuniones + 1
	rsReu.MoveNext
Loop
ReDim Preserve ordenIds(totalReuniones)

' --- Ahora si, el segundo resultset: corresponsables de esas reuniones ---
Set rsCorr = rsReu.NextRecordset()

' --- Agrupa corresponsables por ID_AGENDA: dicCorr(idAgenda) = "idPersona|nombre" Chr(9) ... ---
Dim dicCorr, claveAgenda
Set dicCorr = Server.CreateObject("Scripting.Dictionary")
If Not rsCorr Is Nothing Then
	If Not rsCorr.State = 0 Then
		Do While Not rsCorr.EOF
			claveAgenda = CStr(rsCorr("ID_AGENDA"))
			If Not dicCorr.Exists(claveAgenda) Then dicCorr.Add claveAgenda, ""
			dicCorr(claveAgenda) = dicCorr(claveAgenda) & (rsCorr("ID_PERSONA") & "") & "|" & (rsCorr("NOMBRE") & "") & Chr(9)
			rsCorr.MoveNext
		Loop
	End If
End If

' --- Segunda pasada: adjunta la lista de participantes a cada reunion ---
For i = 0 To totalReuniones - 1
	Dim mAttach, idReuAttach, partesCorr, p, nombreP, idP
	Set mAttach = dicReuniones(ordenIds(i))
	idReuAttach = mAttach("id")

	nPart = 0
	ReDim listaPart(200)
	If dicCorr.Exists(CStr(idReuAttach)) Then
		partesCorr = Split(dicCorr(CStr(idReuAttach)), Chr(9))
		For Each p In partesCorr
			If Trim(p) <> "" Then
				idP = Split(p, "|")(0)
				nombreP = Split(p, "|")(1)
				Set dicP = Server.CreateObject("Scripting.Dictionary")
				dicP.Add "id", idP
				dicP.Add "nombre", nombreP
				Set listaPart(nPart) = dicP
				nPart = nPart + 1
			End If
		Next
	End If
	If nPart = 0 Then
		listaPart = Array()
	Else
		ReDim Preserve listaPart(nPart - 1)
	End If
	mAttach.Add "participantes", listaPart
	mAttach.Add "nParticipantes", nPart
Next

If Not rsReu Is Nothing Then If rsReu.State <> 0 Then rsReu.Close
If Not rsCorr Is Nothing Then If rsCorr.State <> 0 Then rsCorr.Close

' --- Valida cruces de horario con sp_Agenda_ValidarCruceHorario ---
' Para cada reunion PENDIENTE/POSTERGADA, valida al responsable y a cada
' participante del sistema (con ID_PERSONA numerico) contra el resto de la
' agenda, excluyendo la propia reunion.
For i = 0 To totalReuniones - 1
	Dim mm
	Set mm = dicReuniones(ordenIds(i))
	Dim dicCruces
	Set dicCruces = Server.CreateObject("Scripting.Dictionary") ' clave = idPersona, valor = array de detalles (texto con separadores)
	mm.Add "cruces", dicCruces

	If UCase(mm("estado")) = "PENDIENTE" Or UCase(mm("estado")) = "POSTERGADA" Then
		Dim horaIniInt, horaFinInt, fechaInt
		fechaInt = mm("fecha")
		horaIniInt = mm("hiInt")
		horaFinInt = mm("hfInt")

		nPersonas = 0
		ReDim personas(mm("nParticipantes") + 1)
		If Trim(mm("respId") & "") <> "" Then
			If Trim(mm("respId") & "") <> "0" Then
				personas(nPersonas) = mm("respId")
				nPersonas = nPersonas + 1
			End If
		End If
		listaP = mm("participantes")
		For k = 0 To UBound(listaP)
			If Trim(listaP(k)("id") & "") <> "" Then
				If Trim(listaP(k)("id") & "") <> "0" Then
					personas(nPersonas) = listaP(k)("id")
					nPersonas = nPersonas + 1
				End If
			End If
		Next

		For k = 0 To nPersonas - 1
			Dim idPersonaChk, rsCruce, detalleCruce
			idPersonaChk = CStr(personas(k))
			Set rsCruce = Exec(Cnn, "sp_cg_Agenda_ValidarCruceHorario @IdResponsable=" & SqlStr(idPersonaChk) & ", @FechaReunion=" & fechaInt & ", @HoraInicio=" & horaIniInt & ", @HoraFin=" & horaFinInt & ", @IdAgendaExcluir=" & mm("id"))
			detalleCruce = ""
			Do While Not rsCruce.EOF
				If detalleCruce <> "" Then detalleCruce = detalleCruce & Chr(29)
				detalleCruce = detalleCruce & (rsCruce("ID_AGENDA_CRUCE") & "") & Chr(28) & (rsCruce("ASUNTO") & "") & Chr(28) & (rsCruce("HORA_INICIO") & "") & Chr(28) & (rsCruce("HORA_FIN") & "") & Chr(28) & (rsCruce("ROL") & "")
				rsCruce.MoveNext
			Loop
			rsCruce.Close
			If detalleCruce <> "" Then dicCruces.Add CStr(idPersonaChk), detalleCruce
		Next
	End If
Next

' --- Agrupa por dia de semana (0=lunes..4=viernes), ordena por hora y asigna columnas para solapados ---
Sub AsignarColumnas(arr, n)
	Dim ii, kk, jj, colAsignada, activeColsUsed, clusterCount, clusterEnd
	Dim activeCols(60), clusterMembers(60)
	activeColsUsed = 0 : clusterCount = 0 : clusterEnd = -1
	For ii = 0 To n - 1
		Dim mi
		Set mi = arr(ii)
		If clusterCount > 0 And mi("hi") >= clusterEnd Then
			For jj = 0 To clusterCount - 1
				clusterMembers(jj)("cols") = activeColsUsed
			Next
			clusterCount = 0 : activeColsUsed = 0 : clusterEnd = -1
		End If
		colAsignada = -1
		For kk = 0 To activeColsUsed - 1
			If activeCols(kk) <= mi("hi") Then
				colAsignada = kk
				Exit For
			End If
		Next
		If colAsignada = -1 Then
			colAsignada = activeColsUsed
			activeColsUsed = activeColsUsed + 1
		End If
		activeCols(colAsignada) = mi("hf")
		mi("col") = colAsignada
		Set clusterMembers(clusterCount) = mi
		clusterCount = clusterCount + 1
		If mi("hf") > clusterEnd Then clusterEnd = mi("hf")
	Next
	If clusterCount > 0 Then
		For jj = 0 To clusterCount - 1
			clusterMembers(jj)("cols") = activeColsUsed
		Next
	End If
End Sub

' Se agrupa en 5 arrays independientes (uno por dia Lun..Vie)
Dim diaArr0(), diaArr1(), diaArr2(), diaArr3(), diaArr4()
Dim nDia0, nDia1, nDia2, nDia3, nDia4
nDia0 = 0 : nDia1 = 0 : nDia2 = 0 : nDia3 = 0 : nDia4 = 0
ReDim diaArr0(500) : ReDim diaArr1(500) : ReDim diaArr2(500) : ReDim diaArr3(500) : ReDim diaArr4(500)

For i = 0 To totalReuniones - 1
	Dim mrec, dow
	Set mrec = dicReuniones(ordenIds(i))
	dow = mrec("fecha") - epochDesde
	Select Case dow
		Case 0: Set diaArr0(nDia0) = mrec : nDia0 = nDia0 + 1
		Case 1: Set diaArr1(nDia1) = mrec : nDia1 = nDia1 + 1
		Case 2: Set diaArr2(nDia2) = mrec : nDia2 = nDia2 + 1
		Case 3: Set diaArr3(nDia3) = mrec : nDia3 = nDia3 + 1
		Case 4: Set diaArr4(nDia4) = mrec : nDia4 = nDia4 + 1
	End Select
Next

' Ordena cada dia por hora de inicio (insertion sort simple)
Sub OrdenarPorHora(arr, n)
	Dim ii, jj
	For ii = 1 To n - 1
		Dim actual
		Set actual = arr(ii)
		jj = ii - 1
		Do While jj >= 0
			If arr(jj)("hi") > actual("hi") Then
				Set arr(jj + 1) = arr(jj)
				jj = jj - 1
			Else
				Exit Do
			End If
		Loop
		Set arr(jj + 1) = actual
	Next
End Sub

OrdenarPorHora diaArr0, nDia0 : AsignarColumnas diaArr0, nDia0
OrdenarPorHora diaArr1, nDia1 : AsignarColumnas diaArr1, nDia1
OrdenarPorHora diaArr2, nDia2 : AsignarColumnas diaArr2, nDia2
OrdenarPorHora diaArr3, nDia3 : AsignarColumnas diaArr3, nDia3
OrdenarPorHora diaArr4, nDia4 : AsignarColumnas diaArr4, nDia4

Dim ROW_H
ROW_H = 58
Dim alturaGrilla
alturaGrilla = (HORA_END - HORA_START) * ROW_H

Dim ahoraDecimal, mostrandoSemanaActual
ahoraDecimal = Hour(Now()) + Minute(Now()) / 60
mostrandoSemanaActual = (hoyEpoch >= epochDesde And hoyEpoch <= epochHasta)

Dim volverUrl, nuevaUrl
volverUrl = "agenda.asp?xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu)
nuevaUrl = "agenda.asp?opc=edicion&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu)

Function CruceExiste(dic, idPersona)
	If dic.Exists(CStr(idPersona)) Then
		CruceExiste = "true"
	Else
		CruceExiste = "false"
	End If
End Function

Function NombrePorId(rm, idBuscado)
	If CStr(rm("respId")) = CStr(idBuscado) Then
		NombrePorId = rm("respNombre")
		Exit Function
	End If
	Dim lst, pp2
	lst = rm("participantes")
	For Each pp2 In lst
		If CStr(pp2("id")) = CStr(idBuscado) Then
			NombrePorId = pp2("nombre")
			Exit Function
		End If
	Next
	NombrePorId = "Persona #" & idBuscado
End Function

' ============================================================================
' Armado del HTML: todo se concatena en xhtml (mismo patron que agenda.asp /
' agenda_ext.inc) y se emite con un unico Response.Write al final.
' ============================================================================
Dim xhtml, rangoLbl, dLunes, dViernes

dLunes = CDate(epochLunes + CLng(EPOCH))
dViernes = CDate(epochLunes + 4 + CLng(EPOCH))
If Month(dLunes) = Month(dViernes) Then
	rangoLbl = Day(dLunes) & " &ndash; " & Day(dViernes) & " de " & MonthName(Month(dViernes)) & ", " & Year(dViernes)
Else
	rangoLbl = Day(dLunes) & " " & MonthName(Month(dLunes), True) & " &ndash; " & Day(dViernes) & " " & MonthName(Month(dViernes), True) & ", " & Year(dViernes)
End If

xhtml = "<!doctype html><html lang='es'><head>"
xhtml = xhtml & "<meta charset='utf-8'><meta name='viewport' content='width=device-width, initial-scale=1'>"
xhtml = xhtml & "<title>Agenda | Vista semanal</title>"
xhtml = xhtml & "<link rel='stylesheet' href='assets/css/main.css' />"
xhtml = xhtml & "<link rel='stylesheet' href='assets/css/corgestion-ui.css?v=20260924c' />"
xhtml = xhtml & "</head><body class='ui-week-shell'>"

xhtml = xhtml & "<div class='ui-week-header'>"
xhtml = xhtml & "<div class='ui-week-brand'><span class='ui-week-mark'>CG</span> Agenda &mdash; Vista semanal</div>"
xhtml = xhtml & "<div class='ui-week-nav'>"
xhtml = xhtml & "<a class='ui-button ui-button--quiet' href='agenda_semana.asp?d=" & (epochLunes - 7) & "&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu) & "' title='Semana anterior'>&#8592;</a>"
xhtml = xhtml & "<a class='ui-button ui-button--quiet' href='agenda_semana.asp?d=" & hoyEpoch & "&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu) & "'>Hoy</a>"
xhtml = xhtml & "<a class='ui-button ui-button--quiet' href='agenda_semana.asp?d=" & (epochLunes + 7) & "&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu) & "' title='Semana siguiente'>&#8594;</a>"
xhtml = xhtml & "<div class='ui-week-rango'>" & rangoLbl & "</div>"
xhtml = xhtml & "</div>"
xhtml = xhtml & "<div class='ui-cluster'>"
xhtml = xhtml & "<button type='button' class='ui-button ui-button--primary' onclick='irANueva()'>+ Nueva reuni&oacute;n</button>"
xhtml = xhtml & "<button type='button' class='ui-button' onclick='cerrarVentana()'>Cerrar</button>"
xhtml = xhtml & "</div>"
xhtml = xhtml & "</div>"

xhtml = xhtml & "<div class='ui-week-toolbar'>"
xhtml = xhtml & "<div class='ui-week-item'><span class='ui-week-dot ui-week-reunion--pendiente'></span>Pendiente</div>"
xhtml = xhtml & "<div class='ui-week-item'><span class='ui-week-dot ui-week-reunion--postergada'></span>Postergada</div>"
xhtml = xhtml & "<div class='ui-week-item'><span class='ui-week-dot ui-week-reunion--realizada'></span>Realizada</div>"
xhtml = xhtml & "<div class='ui-week-item' style='color:var(--ui-danger);font-weight:600;'><span class='ui-week-dot ui-week-dot--round' style='background:var(--ui-danger);'></span>Cruce de horario detectado (sp_cg_Agenda_ValidarCruceHorario)</div>"
xhtml = xhtml & "</div>"

xhtml = xhtml & "<div class='ui-week-body'>"
xhtml = xhtml & "<div class='ui-week-calendario'>"
xhtml = xhtml & "<div class='ui-week-cabecera-dias'><div class='ui-week-gutter'></div>"

For i = 0 To 4
	Dim esHoy, diaEpoch, claseDia
	diaEpoch = epochLunes + i
	esHoy = (diaEpoch = hoyEpoch)
	If esHoy Then claseDia = "ui-week-dia ui-week-dia--hoy" Else claseDia = "ui-week-dia"
	xhtml = xhtml & "<div class='" & claseDia & "'><div class='ui-week-lbl'>" & NOMBRES_DIA(i) & "</div><div class='ui-week-num'>" & Day(CDate(diaEpoch + CLng(EPOCH))) & "</div></div>"
Next

xhtml = xhtml & "</div>"
xhtml = xhtml & "<div class='ui-week-grilla-wrap' id='grillaWrap'>"
xhtml = xhtml & "<div class='ui-week-gutter-horas' style='height:" & alturaGrilla & "px;'>"

Dim hIdx
For hIdx = HORA_START To HORA_END - 1
	xhtml = xhtml & "<div class='ui-week-hora' style='top:" & ((hIdx - HORA_START) * ROW_H) & "px;'>" & Right("0" & hIdx, 2) & ":00</div>"
Next

xhtml = xhtml & "</div>"
xhtml = xhtml & "<div class='ui-week-columnas'>"

Dim jsMeetings
jsMeetings = ""

For i = 0 To 4
	xhtml = xhtml & "<div class='ui-week-dia-col' style='height:" & alturaGrilla & "px;'>"

	Dim hy
	For hy = 0 To (HORA_END - HORA_START)
		xhtml = xhtml & "<div class='ui-week-linea-hora' style='top:" & (hy * ROW_H) & "px;'></div>"
	Next

	If (epochLunes + i) = hoyEpoch And ahoraDecimal >= HORA_START And ahoraDecimal < HORA_END Then
		xhtml = xhtml & "<div class='ui-week-linea-ahora' id='lineaAhora' style='top:" & CssNum((ahoraDecimal - HORA_START) * ROW_H) & "px;'><span class='ui-week-punto'></span></div>"
	End If

	Dim arrDia, nArrDia
	Select Case i
		Case 0: arrDia = diaArr0 : nArrDia = nDia0
		Case 1: arrDia = diaArr1 : nArrDia = nDia1
		Case 2: arrDia = diaArr2 : nArrDia = nDia2
		Case 3: arrDia = diaArr3 : nArrDia = nDia3
		Case 4: arrDia = diaArr4 : nArrDia = nDia4
	End Select

	For j = 0 To nArrDia - 1
		Dim rm, top, height, widthPct, leftPct, claseEstado, hasCruce, cruceDic, claseBorde
		Set rm = arrDia(j)
		top = (rm("hi") - HORA_START) * ROW_H + 2
		height = (rm("hf") - rm("hi")) * ROW_H - 4
		widthPct = 100 / rm("cols")
		leftPct = rm("col") * widthPct
		claseEstado = EstadoClase(rm("estado"))

		Set cruceDic = rm("cruces")
		hasCruce = (cruceDic.Count > 0)
		claseBorde = ""
		If hasCruce Then claseBorde = " ui-week-reunion--cruce"

		xhtml = xhtml & "<button type='button' class='ui-week-reunion " & claseEstado & claseBorde & "' onclick='seleccionarReunion(" & rm("id") & ")' style='top:" & CssNum(top) & "px;height:" & CssNum(height) & "px;left:calc(" & CssNum(leftPct) & "% + 3px);width:calc(" & CssNum(widthPct) & "% - 7px);'>"
		If hasCruce Then xhtml = xhtml & "<span class='ui-week-flag' aria-label='Cruce de horario detectado'>!</span>"
		xhtml = xhtml & "<div class='ui-week-titulo'>" & Server.HTMLEncode(rm("asunto")) & "</div>"
		xhtml = xhtml & "<div class='ui-week-hora-lbl'>" & FormatHoraDecimal(rm("hi")) & "&ndash;" & FormatHoraDecimal(rm("hf")) & "</div>"
		xhtml = xhtml & "<div class='ui-week-avatares'>"
		If Trim(rm("respId") & "") <> "" And CStr(rm("respId")) <> "0" Then
			xhtml = xhtml & "<span class='ui-week-avatar' style='background:" & ColorPersona(rm("respId")) & ";'>" & Iniciales(rm("respNombre")) & "</span>"
		End If
		Dim listaPp, pp
		listaPp = rm("participantes")
		For Each pp In listaPp
			xhtml = xhtml & "<span class='ui-week-avatar' style='background:" & ColorPersona(pp("id")) & ";'>" & Iniciales(pp("nombre")) & "</span>"
		Next
		xhtml = xhtml & "</div></button>"

		' --- arma el registro JS para esta reunion (usado por el panel lateral) ---
		Dim jsParticipantes, cruceLista, idKey
		jsParticipantes = ""
		If Trim(rm("respId") & "") <> "" And CStr(rm("respId")) <> "0" Then
			jsParticipantes = jsParticipantes & "{id:'" & JsStr(rm("respId")) & "',nombre:'" & JsStr(rm("respNombre")) & "',rol:'RESPONSABLE',cruce:" & CruceExiste(cruceDic, rm("respId")) & "},"
		End If
		For Each pp In listaPp
			jsParticipantes = jsParticipantes & "{id:'" & JsStr(pp("id")) & "',nombre:'" & JsStr(pp("nombre")) & "',rol:'PARTICIPANTE',cruce:" & CruceExiste(cruceDic, pp("id")) & "},"
		Next

		cruceLista = ""
		For Each idKey In cruceDic.Keys
			Dim detalles, det, nombrePersona
			nombrePersona = NombrePorId(rm, idKey)
			detalles = Split(cruceDic(idKey), Chr(29))
			For Each det In detalles
				Dim campos
				campos = Split(det, Chr(28))
				cruceLista = cruceLista & "{persona:'" & JsStr(nombrePersona) & "',otroId:'" & JsStr(campos(0)) & "',otroTitulo:'" & JsStr(campos(1)) & "',otroHorario:'" & JsStr(FormatHoraDecimal(HHMMaDecimal(campos(2)))) & "\u2013" & JsStr(FormatHoraDecimal(HHMMaDecimal(campos(3)))) & "',otroRol:'" & JsStr(campos(4)) & "'},"
			Next
		Next

		jsMeetings = jsMeetings & "'" & rm("id") & "':{id:" & rm("id") & ",titulo:'" & JsStr(rm("asunto")) & "',estado:'" & JsStr(rm("estado")) & "',fechaLabel:'" & JsStr(NOMBRES_DIA(i) & " " & Day(CDate((epochLunes+i) + CLng(EPOCH)))) & "',horaLabel:'" & JsStr(FormatHoraDecimal(rm("hi"))) & "\u2013" & JsStr(FormatHoraDecimal(rm("hf"))) & "',personas:[" & jsParticipantes & "],cruces:[" & cruceLista & "]},"
	Next

	xhtml = xhtml & "</div>"
Next

xhtml = xhtml & "</div></div></div>"

xhtml = xhtml & "<div class='ui-week-panel' id='panel'>"
xhtml = xhtml & "<div class='ui-week-vacio' id='panelVacio'>"
xhtml = xhtml & "<div style='font-size:32px;'>&#128197;</div>"
xhtml = xhtml & "<div style='font-size:13px;'>Seleccion&aacute; una reuni&oacute;n del calendario<br>para ver el detalle y la validaci&oacute;n de cruces.</div>"
xhtml = xhtml & "</div>"
xhtml = xhtml & "<div id='panelContenido' style='display:none;'></div>"
xhtml = xhtml & "</div>"
xhtml = xhtml & "</div>"

xhtml = xhtml & "<script>"
xhtml = xhtml & "var XMENU = '" & JsStr(xmenu) & "';"
xhtml = xhtml & "var XTMENU = '" & JsStr(xtmenu) & "';"
xhtml = xhtml & "var REUNIONES = {" & jsMeetings & "};"
' OJO: CStr(True) depende del idioma del servidor ("Verdadero" en espanol),
' y eso rompe el JS. Se escribe el literal true/false a mano.
If mostrandoSemanaActual Then
	xhtml = xhtml & "var MOSTRANDO_HOY = true;"
Else
	xhtml = xhtml & "var MOSTRANDO_HOY = false;"
End If

xhtml = xhtml & "function cerrarVentana(){ window.close(); }"

xhtml = xhtml & "function irANueva(){"
xhtml = xhtml & "  var url = 'agenda.asp?opc=edicion&xmenu=' + encodeURIComponent(XMENU) + '&xtmenu=' + encodeURIComponent(XTMENU);"
xhtml = xhtml & "  if (window.opener && !window.opener.closed) { window.opener.location.href = url; window.close(); }"
xhtml = xhtml & "  else { window.location.href = url; }"
xhtml = xhtml & "}"

xhtml = xhtml & "function abrirReunion(id){"
xhtml = xhtml & "  var url = 'agenda.asp?opc=det&id=' + id + '&xmenu=' + encodeURIComponent(XMENU) + '&xtmenu=' + encodeURIComponent(XTMENU);"
xhtml = xhtml & "  if (window.opener && !window.opener.closed) { window.opener.location.href = url; window.close(); }"
xhtml = xhtml & "  else { window.location.href = url; }"
xhtml = xhtml & "}"

xhtml = xhtml & "function seleccionarReunion(id){"
xhtml = xhtml & "  var r = REUNIONES[id];"
xhtml = xhtml & "  if (!r) return;"
xhtml = xhtml & "  document.getElementById('panelVacio').style.display = 'none';"
xhtml = xhtml & "  var cont = document.getElementById('panelContenido');"
xhtml = xhtml & "  cont.style.display = 'block';"
xhtml = xhtml & "  var html = '';"
xhtml = xhtml & "  html += '<div class=""ui-week-estado"" style=""color:' + estadoTexto(r.estado) + '"">' + r.estado + '</div>';"
xhtml = xhtml & "  html += '<h3>' + escapeHtml(r.titulo) + '</h3>';"
xhtml = xhtml & "  html += '<div class=""ui-week-subtit"">' + escapeHtml(r.fechaLabel) + ' &middot; ' + escapeHtml(r.horaLabel) + '</div>';"
xhtml = xhtml & "  html += '<div class=""ui-week-seccion""><div class=""ui-week-seccion-titulo"">Participantes</div>';"
xhtml = xhtml & "  for (var i = 0; i < r.personas.length; i++) {"
xhtml = xhtml & "    var p = r.personas[i];"
xhtml = xhtml & "    html += '<div class=""ui-week-persona-row""><span class=""ui-week-avatar-lg"" style=""background:' + colorPersona(p.id) + '"">' + iniciales(p.nombre) + '</span>';"
xhtml = xhtml & "    html += '<span style=""flex:1 1 auto;font-size:13px;"">' + escapeHtml(p.nombre) + ' <span style=""color:var(--ui-text-muted);font-size:11px;"">(' + p.rol.toLowerCase() + ')</span></span>';"
xhtml = xhtml & "    if (p.cruce) html += '<span class=""ui-week-tag-cruce"">Cruce</span>';"
xhtml = xhtml & "    html += '</div>';"
xhtml = xhtml & "  }"
xhtml = xhtml & "  html += '</div>';"
xhtml = xhtml & "  html += '<div class=""ui-week-seccion""><div class=""ui-week-seccion-titulo"">Validaci&oacute;n de cruce de horario</div>';"
xhtml = xhtml & "  if (r.cruces.length > 0) {"
xhtml = xhtml & "    for (var j = 0; j < r.cruces.length; j++) {"
xhtml = xhtml & "      var c = r.cruces[j];"
xhtml = xhtml & "      html += '<div class=""ui-week-cruce-item""><div class=""ui-week-p"">' + escapeHtml(c.persona) + ' tiene otra reuni&oacute;n cruzada</div>';"
xhtml = xhtml & "      html += '<div class=""ui-week-d"">&ldquo;' + escapeHtml(c.otroTitulo) + '&rdquo; &middot; ' + escapeHtml(c.otroHorario) + ' &middot; rol: ' + escapeHtml(c.otroRol) + '</div></div>';"
xhtml = xhtml & "    }"
xhtml = xhtml & "  } else {"
xhtml = xhtml & "    html += '<div class=""ui-week-sin-cruce"">&#10003; Sin cruces detectados para los participantes de esta reuni&oacute;n.</div>';"
xhtml = xhtml & "  }"
xhtml = xhtml & "  html += '<div class=""ui-week-sp-call"">EXEC sp_cg_Agenda_ValidarCruceHorario<br>&nbsp;&nbsp;@IdResponsable = &lt;persona&gt;, @FechaReunion, @HoraInicio, @HoraFin,<br>&nbsp;&nbsp;@IdAgendaExcluir = ' + r.id + '</div>';"
xhtml = xhtml & "  html += '</div>';"
xhtml = xhtml & "  html += '<a class=""ui-button ui-button--primary ui-week-abrir"" href=""javascript:void(0)"" onclick=""abrirReunion(' + r.id + ')"">Abrir reuni&oacute;n &rarr;</a>';"
xhtml = xhtml & "  cont.innerHTML = html;"
xhtml = xhtml & "}"

xhtml = xhtml & "function estadoTexto(estado){"
xhtml = xhtml & "  if (estado === 'PENDIENTE') return '" & EstadoColorTextoJs("PENDIENTE") & "';"
xhtml = xhtml & "  if (estado === 'POSTERGADA') return '" & EstadoColorTextoJs("POSTERGADA") & "';"
xhtml = xhtml & "  return '" & EstadoColorTextoJs("REALIZADA") & "';"
xhtml = xhtml & "}"
xhtml = xhtml & "var PALETA_JS = ['#0f6cbd','#8764b8','#107c10','#ca5010','#498205','#c239b3','#986f0b','#005e50'];"
xhtml = xhtml & "function colorPersona(id){"
xhtml = xhtml & "  var s = String(id == null ? '' : id);"
xhtml = xhtml & "  if (!s || s === '0') return '#8a8886';"
xhtml = xhtml & "  var h = 0; for (var i = 0; i < s.length; i++) { h = (h + s.charCodeAt(i)) % 8; }"
xhtml = xhtml & "  return PALETA_JS[h];"
xhtml = xhtml & "}"
xhtml = xhtml & "function iniciales(nombre){"
xhtml = xhtml & "  var partes = (nombre || '').trim().split(' ');"
xhtml = xhtml & "  var out = '';"
xhtml = xhtml & "  for (var i = 0; i < partes.length && out.length < 2; i++) { if (partes[i]) out += partes[i].charAt(0).toUpperCase(); }"
xhtml = xhtml & "  return out || '?';"
xhtml = xhtml & "}"
xhtml = xhtml & "function escapeHtml(s){"
xhtml = xhtml & "  var d = document.createElement('div');"
xhtml = xhtml & "  d.appendChild(document.createTextNode(s || ''));"
xhtml = xhtml & "  return d.innerHTML;"
xhtml = xhtml & "}"
' Al entrar, si se esta mostrando la semana de hoy, centra la grilla en la
' hora actual (linea roja) para que el usuario vea "hoy" de entrada.
xhtml = xhtml & "window.addEventListener('load', function(){"
xhtml = xhtml & "  if (!MOSTRANDO_HOY) return;"
xhtml = xhtml & "  var linea = document.getElementById('lineaAhora'), wrap = document.getElementById('grillaWrap');"
xhtml = xhtml & "  if (linea && wrap) { wrap.scrollTop = Math.max(0, linea.offsetTop - (wrap.clientHeight / 2)); }"
xhtml = xhtml & "});"
xhtml = xhtml & "</script>"
xhtml = xhtml & "</body></html>"

Response.Write xhtml

Cnn.Close
Set Cnn = Nothing
%>
