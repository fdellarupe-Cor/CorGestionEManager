<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/ui_componentes.inc" -->
<%
' =========================================================================
' reportesPendientes.asp - REPORTE DE PENDIENTES POR USUARIO (pantalla nueva, maqueta para conectar)
' Lo abre reportes.asp ("Pendientes por usuario"). Cada recuadro "Dato: ..." es un valor a completar;
' los comentarios [LOGICA] dicen que dato va y de donde puede salir. Las filas de ejemplo estan en
' bloques "repetir por cada ...". Los filtros aplicados ya muestran lo que llega del formulario.
'
' "Pendiente" = incidente o proyecto abierto (sin FEC_BAJA y sin cerrar). [LOGICA] Definir con el
' criterio del sistema (por ejemplo ESTADO activo y OK_CLIENTE vacio).
'
' Filtros (POST desde reportes.asp):
'   pusu    login del usuario (CG_COR_RESPONSABLES.ID_SEG); vacio = todos los usuarios
'   prol    A = cualquier rol | G = responsable de gestion (RESPON_GESTION_SEG) | T = responsable tecnico (RESPON_TEC_SEG)
'   ptipo   A = incidentes y proyectos | I = solo incidentes | P = solo proyectos
'   pcli    ID_CLIENTE; 0 = todos los clientes
'   pmanos  T = todos | FUNCIONAL | TECNICO | CLIENTE  (en manos de: columna RESPONSABLE)
'   pdias   numero: solo los que no tienen avances hace mas de N dias; vacio = todos
' =========================================================================
Dim html, Cnn, xmenu, xtmenu, accTop, fUsu, fRol, fTipo, fCli, fManos, fDias

xmenu = request.querystring("xmenu")
xtmenu = request.querystring("xtmenu")
fUsu = Trim(request.form("pusu") & "")
fRol = UCase(Trim(request.form("prol") & ""))
fTipo = UCase(Trim(request.form("ptipo") & ""))
fCli = Trim(request.form("pcli") & "")
fManos = UCase(Trim(request.form("pmanos") & ""))
fDias = Trim(request.form("pdias") & "")
If fRol <> "G" And fRol <> "T" Then fRol = "A"
If fTipo <> "I" And fTipo <> "P" Then fTipo = "A"
If Not IsNumeric(fCli) Or fCli = "" Then fCli = "0"
If fManos <> "FUNCIONAL" And fManos <> "TECNICO" And fManos <> "CLIENTE" Then fManos = "T"
If Not IsNumeric(fDias) Then fDias = ""

Set Cnn = Server.CreateObject("ADODB.Connection")
Cnn.open scon1
XXX = SYSLOGON(cnn, "REPORTE_PENDIENTES")

html = headhtml() & "<body><div id='wrapper' role='banner'>" & cabecera() & "</div>"
accTop = UiBoton("Imprimir", "light", "window.print();", True, "") & _
	"<a class='ui-button ui-button--secondary ui-button--sm' href='reportes.asp?rep=pendientes&xmenu=" & Server.URLEncode(xmenu) & "&xtmenu=" & Server.URLEncode(xtmenu) & "'>Volver</a>"
html = html & UiPantallaAbre("Reporte de pendientes", accTop, "")
html = html & RepPendCuerpo(fUsu, fRol, fTipo, fCli, fManos, fDias)
html = html & UiPantallaCierra()
html = html & "<script src='assets/js/jquery.min.js'></script><script src='assets/js/skel.min.js'></script><script src='assets/js/util.js'></script><script src='assets/js/main.js'></script>"
html = html & "</body></html>"
Cnn.Close
Response.Write html

' --- Cuerpo del reporte (mismos componentes que el reporte por cliente, reportesDet.asp) ---
Function RepPendCuerpo(fUsu, fRol, fTipo, fCli, fManos, fDias)
	Dim h, usuTxt, cliTxt, diasTxt
	' --- Parametros del reporte (ya funciona: muestra lo que llego del formulario) ---
	If fUsu = "" Then
		usuTxt = "Todos los usuarios"
	Else
		' [LOGICA] Nombre y apellido del usuario elegido: CG_COR_RESPONSABLES (NOMBRE, APELLIDO) por ID_SEG = pusu
		usuTxt = RepDato("nombre y apellido del usuario")
	End If
	If fCli = "0" Then
		cliTxt = "Todos los clientes"
	Else
		' [LOGICA] Nombre del cliente: catalogo de clientes por ID_CLIENTE = pcli
		cliTxt = RepDato("nombre del cliente")
	End If
	If fDias = "" Then diasTxt = "Cualquier antig&uuml;edad" Else diasTxt = "Hace m&aacute;s de " & CLng(fDias) & " d&iacute;as"
	h = UiCardAbre("Par&aacute;metros del reporte", "") & "<div class='ui-fields'>"
	h = h & UiCampoLectura("Usuario", usuTxt, False)
	h = h & UiCampoLectura("Rol", RepRolPend(fRol), False)
	h = h & UiCampoLectura("Tipo", RepTipoPend(fTipo), False)
	h = h & UiCampoLectura("Cliente", cliTxt, False)
	h = h & UiCampoLectura("En manos de", RepManosTexto(fManos), False)
	h = h & UiCampoLectura("Sin avance", diasTxt, False)
	h = h & "</div>" & UiCardCierra()

	' --- Resumen (todos los usuarios del resultado) ---
	h = h & UiCardAbre("Resumen", "") & "<div class='ui-fields'>"
	' [LOGICA] Totales de pendientes con los filtros aplicados (CG_INCIDENTES y CG_PROYECTOS abiertos)
	h = h & UiCampoLectura("Pendientes", RepDato("total"), False)
	If fTipo <> "P" Then h = h & UiCampoLectura("Incidentes", RepDato("cantidad"), False)
	If fTipo <> "I" Then h = h & UiCampoLectura("Proyectos", RepDato("cantidad"), False)
	' [LOGICA] En manos de: columna RESPONSABLE (FUNCIONAL / TECNICO / CLIENTE) de cada caso
	h = h & UiCampoLectura("En manos del cliente", RepDato("cantidad"), False)
	h = h & UiCampoLectura("En manos internas", RepDato("cantidad"), False)
	' [LOGICA] Sin avance reciente: casos cuyo ultimo CG_AVANCES.FECHA (CODOBJETO = ID, TIPOOBJETO = inc/proy, sin FEC_BAJA)
	'          es anterior a hoy - pdias (si pdias esta vacio, usar 15 dias como referencia)
	h = h & UiCampoLectura("Sin avance reciente", RepDato("cantidad"), False)
	' [LOGICA] Antiguedad promedio: dias desde FECHA_INGRESO hasta hoy, promedio de los pendientes
	h = h & UiCampoLectura("Antig&uuml;edad promedio", RepDato("N d&iacute;as"), False)
	h = h & "</div>" & UiCardCierra()

	' ===== [LOGICA] REPETIR POR CADA USUARIO DEL RESULTADO (si pusu tiene valor, un solo usuario) =====
	h = h & UiCardAbre("Pendientes de " & RepDato("nombre y apellido del usuario"), "")
	h = h & UiTablaAbre("Pendientes del usuario")
	h = h & "<thead><tr>" & UiTh("N&ordm;", "") & UiTh("Cliente", "") & UiTh("T&iacute;tulo", "") & UiTh("Prioridad", "") & UiTh("Estado funcional", "") & UiTh("Estado desarrollo", "") & _
		UiTh("En manos de", "") & UiTh("Rol del usuario", "") & UiTh("Ingreso", "") & UiTh("&Uacute;ltimo avance", "") & UiTh("D&iacute;as sin avance", "") & UiTh("<span class='ui-sr-only'>Abrir</span>", "") & "</tr></thead><tbody>"
	' ----- [LOGICA] REPETIR POR CADA CASO PENDIENTE (orden sugerido: dias sin avance, de mayor a menor) -----
	'   N: "INC-" & ID (CG_INCIDENTES) o "PRY-" & ID (CG_PROYECTOS), con la clase ui-tag-id (proyecto: ui-tag-id ui-tag-id--pry)
	'   Cliente: ID_CLIENTE -> nombre | Titulo: TITULO | Prioridad: PRIORIDAD -> descripcion (UiBadge, como en los listados)
	'   Estado funcional / desarrollo: ESTADO_FUNC / ESTADO_DESA -> descripcion
	'   En manos de: RESPONSABLE (Funcional / Tecnico / Cliente) | Rol del usuario: Gestion (RESPON_GESTION_SEG) / Tecnico (RESPON_TEC_SEG)
	'   Ingreso: FECHA_INGRESO | Ultimo avance: MAX(CG_AVANCES.FECHA) del caso ("Sin avances" si no tiene)
	'   Dias sin avance: hoy - ultimo avance (si supera pdias, o 15 por defecto, mostrarlo con UiBadge(..., "red"))
	'   Abrir: enlace ui-link a detalles.asp?opc=veri&id=ID (incidente) o detalles.asp?opc=verp&id=ID (proyecto)
	h = h & "<tr><td>" & RepDato("INC-/PRY- n&uacute;mero") & "</td><td>" & RepDato("cliente") & "</td><td>" & RepDato("t&iacute;tulo") & "</td><td>" & RepDato("prioridad") & "</td>" & _
		"<td>" & RepDato("estado") & "</td><td>" & RepDato("estado") & "</td><td>" & RepDato("Funcional / T&eacute;cnico / Cliente") & "</td><td>" & RepDato("Gesti&oacute;n / T&eacute;cnico") & "</td>" & _
		"<td>" & RepDato("fecha") & "</td><td>" & RepDato("fecha") & "</td><td>" & RepDato("d&iacute;as") & "</td><td>" & RepDato("Abrir") & "</td></tr>"
	' ----- fin de la repeticion -----
	' [LOGICA] Si el usuario no tiene pendientes: h = h & UiFilaVacia(12, "Sin pendientes.")
	h = h & UiTablaCierra()
	h = h & UiPie(RepDato("N pendientes") & " &middot; " & RepDato("N sin avance reciente"), "")
	h = h & UiCardCierra()
	' ===== fin de la repeticion por usuario =====

	' [LOGICA] Si no hay ningun pendiente para los filtros elegidos, mostrar en lugar de las tarjetas de usuarios:
	'   UiCardAbre("Sin resultados", "") & "<div class='ui-empty'>No hay pendientes para los filtros elegidos.</div>" & UiCardCierra()
	RepPendCuerpo = h
End Function

' --- Ayudas de presentacion ---
' Recuadro de la maqueta: marca el valor que falta conectar (desaparece al conectar los datos)
Function RepDato(texto)
	RepDato = "<span class='ui-dato'>Dato: " & texto & "</span>"
End Function
Function RepRolPend(r)
	Select Case r
		Case "G"
			RepRolPend = "Responsable de gesti&oacute;n"
		Case "T"
			RepRolPend = "Responsable t&eacute;cnico"
		Case Else
			RepRolPend = "Cualquier rol"
	End Select
End Function
Function RepTipoPend(t)
	Select Case t
		Case "I"
			RepTipoPend = "Solo incidentes"
		Case "P"
			RepTipoPend = "Solo proyectos"
		Case Else
			RepTipoPend = "Incidentes y proyectos"
	End Select
End Function
Function RepManosTexto(m)
	Select Case m
		Case "FUNCIONAL"
			RepManosTexto = "Funcional"
		Case "TECNICO"
			RepManosTexto = "T&eacute;cnico"
		Case "CLIENTE"
			RepManosTexto = "Cliente"
		Case Else
			RepManosTexto = "Todos"
	End Select
End Function
%>