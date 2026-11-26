<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/ui_componentes.inc" -->
<%
' =========================================================================
' reportesAgenda.asp - REPORTE DE AGENDA POR USUARIO
' Lo abre reportes.asp ("Agenda por usuario"). Datos: sp_cg_AgendaReporte (DB/Cambios_Agenda_Reporte.sql).
'
' Filtros (POST desde reportes.asp; los links de paginacion los repiten en la URL):
'   agusu   login del usuario (CG_COR_RESPONSABLES.ID_SEG); vacio = todos los usuarios
'   agrol   A = responsable o participante (indistinto) | R = solo responsable | P = solo participante
'   agest   T = todas | PENDIENTE | POSTERGADA | REALIZADA
'   agfecd  desde (dd/mm/aaaa) - fecha de la reunion
'   agfech  hasta (dd/mm/aaaa)
'   pd / pa pagina del detalle / pagina del arbol (URL)
' (agrec, recordatorios, ya no se usa en este reporte)
'
' Secciones: Parametros del reporte, Resumen (tabla), Detalle de reuniones (paginado) y Arbol de seguimientos (paginado).
' Arbol: raiz = tema marcado "Requiere seguimiento"; ramas = el mismo tema (mismo tipo y texto) en las reuniones
' generadas con "Generar seguimiento". Las ramas se muestran aunque queden fuera del periodo o del usuario filtrado.
' =========================================================================
Const REP_POR_PAG_DET = 15   ' reuniones por pagina en el detalle
Const REP_POR_PAG_ARB = 8    ' arboles (temas raiz) por pagina

Dim html, Cnn, xmenu, xtmenu, accTop, fUsu, fRol, fEst, fDesde, fHasta, argSp, pagDet, pagArb

xmenu = request.querystring("xmenu")
xtmenu = request.querystring("xtmenu")
fUsu = Trim(Request("agusu") & "")
fRol = UCase(Trim(Request("agrol") & ""))
fEst = UCase(Trim(Request("agest") & ""))
fDesde = Trim(Request("agfecd") & "")
fHasta = Trim(Request("agfech") & "")
If fRol <> "R" And fRol <> "P" Then fRol = "A"
If fEst <> "PENDIENTE" And fEst <> "POSTERGADA" And fEst <> "REALIZADA" Then fEst = "T"
If Not IsDate(fDesde) Then fDesde = ""
If Not IsDate(fHasta) Then fHasta = ""
pagDet = RepPag(Request.QueryString("pd"))
pagArb = RepPag(Request.QueryString("pa"))

Set Cnn = Server.CreateObject("ADODB.Connection")
Cnn.open scon1
XXX = SYSLOGON(cnn, "REPORTE_AGENDA")

' Filtros para el SP (el usuario y las fechas vacios viajan como NULL; las fechas, en dias desde 1900-01-01)
argSp = ", @Usuario=" & RepSqlUsuario(fUsu) & ", @Rol=" & SqlStr(fRol) & ", @Estado=" & RepSqlEstado(fEst) & _
	", @Desde=" & RepSqlDias(fDesde) & ", @Hasta=" & RepSqlDias(fHasta)

html = headhtml() & "<body><div id='wrapper' role='banner'>" & cabecera() & "</div>"
accTop = UiBoton("Imprimir", "light", "window.print();", True, "") & _
	"<a class='ui-button ui-button--secondary ui-button--sm' href='reportes.asp?rep=agenda&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu) & "'>Volver</a>"
html = html & UiPantallaAbre("Reporte de agenda", accTop, "wide")
html = html & RepAgendaCuerpo(Cnn, argSp, fUsu, fRol, fEst, fDesde, fHasta)
html = html & UiPantallaCierra()
html = html & "<script src='assets/js/jquery.min.js'></script><script src='assets/js/skel.min.js'></script><script src='assets/js/util.js'></script><script src='assets/js/main.js'></script>"
html = html & "</body></html>"
Cnn.Close
Response.Write html

' --- Cuerpo del reporte (mismos componentes que el reporte por cliente, reportesDet.asp) ---
Function RepAgendaCuerpo(cn, argSp, fUsu, fRol, fEst, fDesde, fHasta)
	Dim h, usuTxt, rs, base, qs
	base = "&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu)
	qs = RepFiltroQs()

	' --- Parametros del reporte ---
	If fUsu = "" Then
		usuTxt = "Todos los usuarios"
	Else
		usuTxt = UiTexto(fUsu)
		Set rs = Exec(cn, "SELECT RTRIM(ISNULL(NOMBRE, '') + ' ' + ISNULL(APELLIDO, '')) AS NOM FROM CG_COR_RESPONSABLES WHERE ID_SEG = " & SqlStr(fUsu))
		If Not rs.EOF Then
			If Trim(rs("NOM") & "") <> "" Then usuTxt = UiTexto(rs("NOM")) & " (" & UiTexto(fUsu) & ")"
		End If
		rs.Close
	End If
	h = UiCardAbre("Par&aacute;metros del reporte", "") & "<div class='ui-fields'>"
	h = h & UiCampoLectura("Usuario", usuTxt, False)
	h = h & UiCampoLectura("Participaci&oacute;n", RepRolTexto(fRol), False)
	h = h & UiCampoLectura("Estado", RepEstadoTexto(fEst), False)
	h = h & UiCampoLectura("Per&iacute;odo", RepPeriodo(fDesde, fHasta), False)
	h = h & "</div>" & UiCardCierra()

	' --- Resumen (tabla) ---
	Dim total, pend, post, real, conSeg
	Set rs = Exec(cn, "sp_cg_AgendaReporte @Modo='resumen'" & argSp)
	total = 0 : pend = 0 : post = 0 : real = 0 : conSeg = 0
	If Not rs.EOF Then
		total = CLng(rs("TOTAL")) : pend = CLng(rs("PENDIENTES")) : post = CLng(rs("POSTERGADAS"))
		real = CLng(rs("REALIZADAS")) : conSeg = CLng(rs("CON_SEGUIMIENTO"))
	End If
	rs.Close
	h = h & UiCardAbre("Resumen", "")
	h = h & UiTablaAbre("Cantidad de reuniones por estado")
	h = h & "<thead><tr>" & UiTh("Concepto", "") & UiTh("Cantidad", "ui-table__cell--center") & UiTh("% del total", "ui-table__cell--center") & "</tr></thead><tbody>"
	h = h & RepFilaResumen("Reuniones en el per&iacute;odo", total, total, False)
	h = h & RepFilaResumen("Pendientes", pend, total, True)
	h = h & RepFilaResumen("Postergadas", post, total, True)
	h = h & RepFilaResumen("Realizadas", real, total, True)
	h = h & RepFilaResumen("Reuniones con seguimiento (cada una cuenta 1)", conSeg, total, True)
	h = h & UiTablaCierra() & UiCardCierra()

	' --- Detalle de reuniones (paginado) ---
	Dim n, filas, ini, fin, ant, sig
	ini = (pagDet - 1) * REP_POR_PAG_DET + 1
	fin = ini + REP_POR_PAG_DET - 1
	Set rs = Exec(cn, "sp_cg_AgendaReporte @Modo='detalle'" & argSp)
	n = 0 : filas = ""
	Do While Not rs.EOF
		n = n + 1
		If n >= ini And n <= fin Then
			filas = filas & "<tr><td>" & RepFecha((rs("FECHA_REUNION"))) & "</td><td>" & RepHorario((rs("HORA_INICIO")), (rs("HORA_FIN"))) & "</td><td>" & UiDato(rs("ASUNTO")) & "</td>" & _
				"<td>" & UiDato(rs("RESPONSABLE")) & "</td><td>" & UiDato(rs("PARTICIPACION")) & "</td><td>" & RepBadgeEstado(rs("ESTADO")) & "</td><td>" & RepEnlaceReunion((rs("ID_AGENDA_PADRE")), base) & "</td>" & _
				"<td class='ui-table__cell--center'>" & CLng(rs("TEMAS_SEG")) & "</td><td>" & RepEnlaceReunion((rs("ID")), base) & "</td></tr>"
		End If
		rs.MoveNext
	Loop
	rs.Close
	h = h & UiCardAbre("Detalle de reuniones", "")
	h = h & UiTablaAbre("Reuniones del per&iacute;odo")
	h = h & "<thead><tr>" & UiTh("Fecha", "") & UiTh("Horario", "") & UiTh("Asunto", "") & UiTh("Responsable", "") & UiTh("Participaci&oacute;n", "") & UiTh("Estado", "") & UiTh("Seguimiento de", "") & UiTh("Temas con seg.", "ui-table__cell--center") & UiTh("<span class='ui-sr-only'>Abrir</span>", "") & "</tr></thead><tbody>"
	If filas = "" Then filas = UiFilaVacia(9, "No hay reuniones para esta p&aacute;gina o para los filtros elegidos.")
	h = h & filas & UiTablaCierra()
	If fin > n Then fin = n
	RepBotonesPag pagDet, n, REP_POR_PAG_DET, "pd", "pa", pagArb, qs, base, ant, sig
	h = h & UiPaginacionDe(ini, fin, n, ant, sig)
	h = h & UiCardCierra()

	' --- Arbol de seguimientos (paginado) ---
	h = h & RepArbol(cn, argSp, base, qs)
	RepAgendaCuerpo = h
End Function

' Arbol: filas ordenadas por camino (RUTA); NIVEL 0 = tema raiz. Se pagina por arbol (nunca se corta uno al medio)
Function RepArbol(cn, argSp, base, qs)
	Dim rs, h, prev, nivel, k, raices, cls, rel, ini, fin, mostrar, ant, sig, ultima
	ini = (pagArb - 1) * REP_POR_PAG_ARB + 1
	fin = ini + REP_POR_PAG_ARB - 1
	mostrar = False
	Set rs = Exec(cn, "sp_cg_AgendaReporte @Modo='arbol'" & argSp)
	h = ""
	prev = -1 : raices = 0
	Do While Not rs.EOF
		nivel = CLng(rs("NIVEL"))
		If nivel = 0 Then
			raices = raices + 1
			mostrar = (raices >= ini And raices <= fin)
		End If
		If mostrar Then
			If prev < 0 Then
				h = h & "<ul class='ui-arbol'>"
			ElseIf nivel > prev Then
				h = h & "<ul>"
			ElseIf nivel = prev Then
				h = h & "</li>"
			Else
				For k = 1 To prev - nivel
					h = h & "</li></ul>"
				Next
				h = h & "</li>"
			End If
			cls = "ui-nodo"
			If nivel = 0 Then cls = cls & " ui-nodo--raiz"
			If CLng(rs("EN_FILTRO")) = 0 Then cls = cls & " ui-nodo--fuera"
			rel = ""
			Select Case UCase(rs("TIPO_RELACION") & "")
				Case "PROYECTO"
					If Not IsNull(rs("ID_PROYECTO")) Then rel = "Proyecto #" & CLng(rs("ID_PROYECTO"))
				Case "INCIDENTE"
					If Not IsNull(rs("ID_INCIDENTE")) Then rel = "Incidente #" & CLng(rs("ID_INCIDENTE"))
				Case "CONSULTORIA"
					rel = "Consultor&iacute;a"
			End Select
			h = h & "<li><div class='" & cls & "'>"
			h = h & "<span class='ui-nodo__tema'>" & UiTexto(rs("TEMA_A_TRATAR")) & "</span>"
			If rel <> "" Then h = h & "<span class='ui-nodo__meta'>" & rel & "</span>"
			h = h & "<span class='ui-nodo__meta'>"
			If nivel = 0 Then h = h & "Ra&iacute;z &middot; "
			h = h & RepEnlaceReunion((rs("ID_AGENDA")), base) & " &middot; " & RepFecha((rs("FECHA_REUNION"))) & " &middot; " & UiTexto(rs("RESPONSABLE")) & "</span>"
			h = h & RepBadgeEstado(rs("ESTADO"))
			If CLng(rs("EN_FILTRO")) = 0 Then h = h & UiBadge("Fuera del filtro", "purple")
			If Trim(rs("CONCLUSION") & "") <> "" Then h = h & "<span class='ui-nodo__conclusion'>" & UiTexto(rs("CONCLUSION")) & "</span>"
			h = h & "</div>"
			prev = nivel
		End If
		rs.MoveNext
	Loop
	rs.Close
	If prev >= 0 Then
		h = h & "</li>"
		For k = 1 To prev
			h = h & "</ul></li>"
		Next
		h = h & "</ul>"
	End If
	ultima = fin
	If ultima > raices Then ultima = raices
	RepArbol = UiCardAbre("&Aacute;rbol de seguimientos", "") & _
		"<p class='ui-table__nota'>Cada &aacute;rbol parte del tema ra&iacute;z (marcado &laquo;Requiere seguimiento&raquo;) y sigue por el mismo tema en las reuniones de seguimiento, aunque est&eacute;n fuera del per&iacute;odo o del usuario elegido.</p>"
	If raices = 0 Then
		RepArbol = RepArbol & "<div class='ui-empty'>No hay temas con seguimiento para los filtros elegidos.</div>"
	ElseIf h = "" Then
		RepArbol = RepArbol & "<div class='ui-empty'>No hay &aacute;rboles en esta p&aacute;gina.</div>"
	Else
		RepArbol = RepArbol & h
	End If
	RepBotonesPag pagArb, raices, REP_POR_PAG_ARB, "pa", "pd", pagDet, qs, base, ant, sig
	RepArbol = RepArbol & UiPaginacionDe(ini, ultima, raices, ant, sig) & UiCardCierra()
End Function

' --- Ayudas de presentacion ---
' Numero de pagina (>= 1)
Function RepPag(v)
	RepPag = 1
	If Trim(v & "") <> "" Then
		If IsNumeric(v & "") Then
			If CLng(v) > 1 Then RepPag = CLng(v)
		End If
	End If
End Function
' Filtros del reporte en la URL (los links de paginacion los conservan; la pagina lee POST o URL)
Function RepFiltroQs()
	RepFiltroQs = "&agusu=" & Server.URLEncode(fUsu) & "&agrol=" & Server.URLEncode(fRol) & "&agest=" & Server.URLEncode(fEst) & "&agfecd=" & Server.URLEncode(fDesde) & "&agfech=" & Server.URLEncode(fHasta)
End Function
' Botones Anterior / Siguiente de una seccion. pn = su parametro de pagina; po = el de la otra seccion (se conserva con valor pov)
Sub RepBotonesPag(pag, total, porPag, pn, po, pov, qs, base, ant, sig)
	Dim url
	url = "reportesAgenda.asp?" & po & "=" & pov & qs & base & "&" & pn & "="
	If pag > 1 Then
		ant = "<a class='ui-button ui-button--secondary ui-button--sm' href='" & url & (pag - 1) & "'>Anterior</a>"
	Else
		ant = UiBoton("Anterior", "secondary", "", False, "Primera p&aacute;gina")
	End If
	If pag * porPag < total Then
		sig = "<a class='ui-button ui-button--secondary ui-button--sm' href='" & url & (pag + 1) & "'>Siguiente</a>"
	Else
		sig = UiBoton("Siguiente", "secondary", "", False, "&Uacute;ltima p&aacute;gina")
	End If
End Sub
' Fila de la tabla de resumen
Function RepFilaResumen(texto, n, total, conPorc)
	Dim p
	If conPorc And total > 0 Then p = Round(n * 100 / total, 0) & " %" Else p = "&mdash;"
	RepFilaResumen = "<tr><td>" & texto & "</td><td class='ui-table__cell--center'>" & n & "</td><td class='ui-table__cell--center'>" & p & "</td></tr>"
End Function
Function RepBadgeEstado(e)
	Select Case UCase(e & "")
		Case "POSTERGADA"
			RepBadgeEstado = UiBadge("Postergada", "amber")
		Case "REALIZADA"
			RepBadgeEstado = UiBadge("Realizada", "gray")
		Case Else
			RepBadgeEstado = UiBadge("Pendiente", "blue")
	End Select
End Function
Function RepEnlaceReunion(id, base)
	If IsNull(id) Or Trim(id & "") = "" Then
		RepEnlaceReunion = "<span class='ui-status-empty'>&mdash;</span>"
	Else
		RepEnlaceReunion = "<a class='ui-link' href='agenda.asp?opc=det&id=" & CLng(id) & base & "'>#" & CLng(id) & "</a>"
	End If
End Function
' Dias desde 1900-01-01 -> dd/mm/aaaa
Function RepFecha(v)
	Dim d
	RepFecha = "<span class='ui-status-empty'>&mdash;</span>"
	If Not IsNull(v) Then
		If IsNumeric(v) Then
			If CLng(v) > 0 Then
				d = CDate(CLng(v) + CLng(DateSerial(1900, 1, 1)))
				RepFecha = Right("0" & Day(d), 2) & "/" & Right("0" & Month(d), 2) & "/" & Year(d)
			End If
		End If
	End If
End Function
' HHMM -> hh:mm
Function RepHora(v)
	RepHora = ""
	If Not IsNull(v) Then
		If IsNumeric(v) Then RepHora = Right("0" & (CLng(v) \ 100), 2) & ":" & Right("0" & (CLng(v) Mod 100), 2)
	End If
End Function
Function RepHorario(a, b)
	If RepHora(a) = "" Then
		RepHorario = "<span class='ui-status-empty'>&mdash;</span>"
	ElseIf RepHora(b) = "" Then
		RepHorario = RepHora(a)
	Else
		RepHorario = RepHora(a) & " a " & RepHora(b)
	End If
End Function
' Filtros hacia el SP
Function RepSqlUsuario(u)
	If u = "" Then RepSqlUsuario = "NULL" Else RepSqlUsuario = SqlStr(u)
End Function
Function RepSqlEstado(e)
	If e = "T" Then RepSqlEstado = "NULL" Else RepSqlEstado = SqlStr(e)
End Function
Function RepSqlDias(f)
	If f = "" Then RepSqlDias = "NULL" Else RepSqlDias = CStr(CLng(CDate(f)) - CLng(DateSerial(1900, 1, 1)))
End Function
Function RepPeriodo(d, hh)
	If d = "" And hh = "" Then
		RepPeriodo = "Sin l&iacute;mite"
	ElseIf d = "" Then
		RepPeriodo = "Hasta el " & Server.HTMLEncode(hh)
	ElseIf hh = "" Then
		RepPeriodo = "Desde el " & Server.HTMLEncode(d)
	Else
		RepPeriodo = Server.HTMLEncode(d) & " al " & Server.HTMLEncode(hh)
	End If
End Function
Function RepRolTexto(r)
	Select Case r
		Case "R"
			RepRolTexto = "Solo como responsable"
		Case "P"
			RepRolTexto = "Solo como participante"
		Case Else
			RepRolTexto = "Responsable o participante (indistinto)"
	End Select
End Function
Function RepEstadoTexto(e)
	Select Case e
		Case "PENDIENTE"
			RepEstadoTexto = "Pendientes"
		Case "POSTERGADA"
			RepEstadoTexto = "Postergadas"
		Case "REALIZADA"
			RepEstadoTexto = "Realizadas"
		Case Else
			RepEstadoTexto = "Todas"
	End Select
End Function
%>
