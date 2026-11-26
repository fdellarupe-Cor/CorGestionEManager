	<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/ui_componentes.inc" -->
<%

'[UI-ANT] ' ===== Bloque anterior: reemplazado por el bloque [UI] que sigue. Se borra entero buscando [UI-ANT] =====
'[UI-ANT] html = "<html>"
'[UI-ANT] html = html & "<head>"

'[UI-ANT] html = html & "<title>Cor Gestion</title>"
'[UI-ANT] html = html & "<meta charset='utf-8' />"
'[UI-ANT] html = html & "<meta name='viewport' content='width=device-width, initial-scale=1' />"
'[UI-ANT] html = html & "<link rel='stylesheet' href='assets/css/main.css' />"
'[UI-ANT] html = html & "<style style='text/css'>"
'[UI-ANT] html = html & ".tablX{"
'[UI-ANT] html = html & "width:100%; "
'[UI-ANT] html = html & "border-collapse:collapse; "
'[UI-ANT] html = html & "}"
'[UI-ANT] html = html & ".tablX td{ "
'[UI-ANT] html = html & "padding:7px; "
'[UI-ANT] html = html & "}"
'[UI-ANT] html = html & ".tablX tr{"
'[UI-ANT] html = html & "background: #FFFFFF;"
'[UI-ANT] html = html & "}"
'[UI-ANT] html = html & ".tablX tr:hover {"
'[UI-ANT] html = html & "background-color: #D4D4D4;"
'[UI-ANT] html = html & "}"
'[UI-ANT] html = html & "</style>"

'[UI-ANT] html = html & "</head>"
'[UI-ANT] html = html & "<body>"

' [UI] Encabezado comun del sistema (el propio declaraba UTF-8 y no cargaba el CSS del sistema)
html = headhtml()
html = html & "<body>"

Set Cnn = Server.CreateObject("ADODB.Connection")
'[UI-ANT] Cnn.open "[credenciales eliminadas]"
' [UI] Conexion con la configuracion del sistema (scon1), como el resto de las pantallas.
' [UI] Antes se conectaba con el usuario administrador de SQL Server y su contrasena escritos en el codigo.
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

nFecini = 0

If nFecini < CLng(CDate("01-01-1900")) Then
	nFecini = CLng(CDate("01-01-1900"))
Else
	nFecini = 0
End If

opc = request.querystring("opc")
xmenu = request.querystring("xmenu")
xtmenu = request.querystring("xtmenu")

if opc <> "consi" then

XXX = SYSLOGON(cnn,"MENU_INCIDENTES")

cliente = request.form("cli")
if cliente & "" = "" then cliente = request.querystring("cli")
if cliente & "" = "" then cliente = 0

tit = request.form("tit")
if tit & "" = "" then tit = request.querystring("tit")
if tit & "" = "" then tit = ""

ord = request.querystring("ord")

CantRegMostrar = 10
'[UI-ANT] CantRegMover = CDbl(Request.querystring("CantRegMover"))
' [UI] Sin CantRegMover en la URL, CDbl("") daba error de tipo
CantRegMover = 0
If IsNumeric(Request.querystring("CantRegMover") & "") Then CantRegMover = CDbl(Request.querystring("CantRegMover"))
if CantRegMover < 0 then CantRegMover = 0


'[UI-ANT] ' ===== Bloque anterior: reemplazado por el bloque [UI] que sigue. Se borra entero buscando [UI-ANT] =====
'[UI-ANT] html = html & "<div>"

'[UI-ANT] html = html & "<header id='header' style='background:#0e1949;padding:1rem 0;'>"

'[UI-ANT] html = html & "<div >"

'[UI-ANT] html = html & "<table style='width:100%;'><tr style='display:flex;padding: 0 1rem;'>"

'[UI-ANT] html = html & "<td style='width:75%;' valign=center><a href='index.asp' class='logo' >"
'[UI-ANT] html = html & "<img src='images/iso_cor.png' style='width:6rem;' >"
'[UI-ANT] html = html & "<font style='font-size:15pt;color:white;'><b>Cor Gestion</b> | Sistema Administrativo</font></a></td>"
'[UI-ANT] html = html & "<td style='width:25%;display:flex;align-items:center;flex-direction:row-reverse;' align='right'>"
'[UI-ANT] html = html & "<nav>"
'[UI-ANT] html = html & "<ul>"
'[UI-ANT] 'html = html & "<li><a href='#menu'></a></li>"
'[UI-ANT] html = html & "</ul>"
'[UI-ANT] html = html & "</nav>"
'[UI-ANT] html = html & "</td>"

'[UI-ANT] html = html & "</tr></table>"

'[UI-ANT] html = html & "</div>"

'[UI-ANT] html = html & "</header>"

'[UI-ANT] response.write html
'[UI-ANT] ''''''''''''''''''''''''''
'[UI-ANT] html = ""
	'[UI-ANT] 'html = html & "<div style='height:77%;width:100%;'><br>"

	'[UI-ANT] ' --- INICIO NORMALIZACION: TABLA TO-DO (INCIDENTES) ---
	'[UI-ANT] ' Contenedor Inner
	'[UI-ANT] html = html & "<div class='inner' style='padding: 2rem;'>"
	
	'[UI-ANT] ' !CORRECCION!: Clases unificadas en un solo DIV para respetar el border-radius
	'[UI-ANT] html = html & "<div class='bloque-datos table-responsive'>"
	
	'[UI-ANT] 'html = html & "<table class='tablX' style='width:99%;color:#000000;font-family:Serif;text-align:center;' cellspacing=0 align=center>"
	'[UI-ANT] html = html & "<table class='tabla-moderna'>"
	'[UI-ANT] html = html & "<thead>"
	
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] html = html & "<tr>"
	
	'[UI-ANT] 'html = html & "<td class='encabe' style='width:1%;'>*</td>"
	'[UI-ANT] html = html & "<th style='width:1%;'></th>"
	
	'[UI-ANT] 'html = html & "<td class='encabe' style='width:1%;'>*</td>"
	'[UI-ANT] html = html & "<th style='width:1%;'></th>"
	
	'[UI-ANT] 'html = html & "<td class='encabe' style='width:1%;'>ID</td>"
	'[UI-ANT] html = html & "<th style='width:1%;'>ID</th>"
	
	'[UI-ANT] 'html = html & "<td class='encabe' style='width:28%;'><a class='red' href='incidentes.asp?ord=titulo&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>TITULO</a></td>"
	'[UI-ANT] html = html & "<th style='width:28%;'><a class='col-link' style='text-decoration:none;' href='incidentes.asp?ord=titulo&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>TITULO</a></th>"
	
	'[UI-ANT] 'html = html & "<td class='encabe' style='width:15%;'><a class='red' href='incidentes.asp?ord=cliente&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>CLIENTE</a></td>"
	'[UI-ANT] html = html & "<th style='width:15%;'><a class='col-link' style='text-decoration:none;' href='incidentes.asp?ord=cliente&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>CLIENTE</a></th>"
	
	'[UI-ANT] 'html = html & "<td class='encabe' style='width:10%;'>PRIORIDAD</td>"
	'[UI-ANT] html = html & "<th style='width:10%;'>PRIORIDAD</th>"
	
	'[UI-ANT] 'html = html & "<td class='encabe' style='width:15%;'>DESARROLLO</td>"
	'[UI-ANT] html = html & "<th style='width:15%;'>DESARROLLO</th>"
	
	'[UI-ANT] 'html = html & "<td class='encabe' style='width:15%;'>FUNCIONAL</td>"
	'[UI-ANT] html = html & "<th style='width:15%;'>FUNCIONAL</th>"
	
	'[UI-ANT] 'html = html & "<td class='encabe' style='width:12%;'>FEC INGRESO</td>"
	'[UI-ANT] html = html & "<th style='width:12%;'>FEC INGRESO</th>"
	
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "</tr>"
	'[UI-ANT] html = html & "</thead>"
	'[UI-ANT] html = html & "<tbody>"

	'[UI-ANT] cmd = "sp_cg_ConsultarEntidad @Entidad='incidente', @Modo='todo', @IdCliente=" & cliente & ", @Titulo='" & tit & "', @Orden='" & ord & "'"

	'[UI-ANT] Set rs = Cnn.Execute(cmd)

	'[UI-ANT] if not rs.eof then rs.Move (CantRegMover)

	'[UI-ANT] PosReg = 0
	'[UI-ANT] sColor = "#FFFFFF"
	'[UI-ANT] formx1 = ""
	'[UI-ANT] formx2 = ""
	'[UI-ANT] formx3 = ""

	'[UI-ANT] Do While (Not rs.eof) 'And (PosReg < CantRegMostrar)

		'[UI-ANT] ' Logica de color de fondo especifico por cliente
		'[UI-ANT] sColor = ""
		'[UI-ANT] if rs(1) & "" = "GRUPO PETERSEN" or rs(1) & "" = "BANCO DE ENTRE RIOS" or rs(1) & "" = "BANCO DE SAN JUAN" or rs(1) & "" = "BANCO DE SANTA CRUZ" or rs(1) & "" = "BANCO DE SANTA FE" then
			'[UI-ANT] sColor = "background-color: #F3E5D2;"
		'[UI-ANT] end if

		'[UI-ANT] est1 = ""
		'[UI-ANT] est2 = ""
		'[UI-ANT] prio = ""

		'[UI-ANT] ' --- BADGES DE ESTADO (DESARROLLO) ---
		'[UI-ANT] If rs(3) = "PENDIENTE" Then
			'[UI-ANT] est1 = "<span class='badge badge-media'>PEND</span> (" & rs(9) & ")"
		'[UI-ANT] ElseIf rs(3) = "OK" Then
			'[UI-ANT] est1 = "<span class='badge badge-ok'>OK</span> (" & rs(9) & ")"
		'[UI-ANT] ElseIf rs(3) = "BAJA" Then
			'[UI-ANT] est1 = "<span class='badge badge-peligro'>BAJA</span> (" & rs(9) & ")"
		'[UI-ANT] ElseIf rs(3) = "TESTING EN CLIENTE" Then
			'[UI-ANT] est1 = "<span class='badge badge-media'>TEST</span> (" & rs(9) & ")"
		'[UI-ANT] End If

		'[UI-ANT] ' --- BADGES DE ESTADO (FUNCIONAL) ---
		'[UI-ANT] If rs(4) = "PENDIENTE" Then
			'[UI-ANT] est2 = "<span class='badge badge-media'>PEND</span> (" & rs(8) & ")"
		'[UI-ANT] ElseIf rs(4) = "OK" Then
			'[UI-ANT] est2 = "<span class='badge badge-ok'>OK</span> (" & rs(8) & ")"
		'[UI-ANT] ElseIf rs(4) = "BAJA" Then
			'[UI-ANT] est2 = "<span class='badge badge-peligro'>BAJA</span> (" & rs(8) & ")"
		'[UI-ANT] ElseIf rs(4) = "TESTING EN CLIENTE" Then
			'[UI-ANT] est2 = "<span class='badge badge-media'>TEST</span> (" & rs(8) & ")"
		'[UI-ANT] End If

		'[UI-ANT] ' --- BADGES DE PRIORIDAD ---
		'[UI-ANT] if rs(2) = "ALTA" then
			'[UI-ANT] prio = "<span class='badge badge-peligro'>ALTA</span>"
		'[UI-ANT] elseif rs(2) = "MEDIA" then
			'[UI-ANT] prio = "<span class='badge badge-media'>MEDIA</span>"
		'[UI-ANT] else
			'[UI-ANT] prio = "<span class='badge badge-pendiente'>" & rs(2) & "</span>"
		'[UI-ANT] end if

		'[UI-ANT] ' --- BOTON ELIMINAR ---
		'[UI-ANT] formx1 = formx1 & "<form id='formx" & rs(0) & "' style='height:1px;' method='post' action='eliminar.asp?opc=inc&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=incidentes.asp'></form>"
		'[UI-ANT] if getPermisos(xmenu,xtmenu,"ELIM") = 1 then
			'[UI-ANT] botonx1 = "<input class='btn-accion btn-accion-x' type='button' value='X' onclick=""borrar('formx" & rs(0) & "');"">"
		'[UI-ANT] else
			'[UI-ANT] botonx1 = "<input class='btn-accion btn-accion-deshabilitado' type='button' value='X' onclick=""alert('Sin permiso para eliminar.');"">"
		'[UI-ANT] end if
		'[UI-ANT] botonx1="" ' Mantenemos la logica original que anula el boton

		'[UI-ANT] ' --- BOTON ACTUALIZAR ---
		'[UI-ANT] formx2 = formx2 & "<form id='formx2" & rs(0) & "' style='height:1px;' method='post' action='detalles.asp?id=" & rs(0) & "&opc=acti&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
		'[UI-ANT] if getPermisos(xmenu,xtmenu,"MODIF") = 1 then
			'[UI-ANT] botonx2 = "<input class='btn-accion btn-accion-a' type='button' value='A' onclick=""document.getElementById('formx2" & rs(0) & "').submit();"">"
		'[UI-ANT] else
			'[UI-ANT] botonx2 = "<input class='btn-accion btn-accion-deshabilitado' type='button' value='A' onclick=""alert('Sin permiso para modificar.');"">"
		'[UI-ANT] end if
		'[UI-ANT] botonx2="" ' Mantenemos la logica original que anula el boton

		'[UI-ANT] ' URL unificada para el evento onclick de las filas
		'[UI-ANT] urlDetalle = "document.location='detalles.asp?id=" & rs(0) & "&opc=veri&cli=" & cliente & "&tit=" & tit & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'"

		'[UI-ANT] ' --- INICIO FILA DE TABLA ---
		'[UI-ANT] 'html = html & "<tr>"
		'[UI-ANT] html = html & "<tr style='" & sColor & "'>"
		
		'[UI-ANT] 'html = html & "<td class='regis'>" & botonx1s & "</td>"
		'[UI-ANT] html = html & "<td>" & botonx1 & "</td>"

		'[UI-ANT] if session("USUARIO") = rs(6) or session("USUARIO") = rs(7) then
			'[UI-ANT] 'html = html & "<td class='regis' >" & botonx2 & "</td>"
			'[UI-ANT] html = html & "<td>" & botonx2 & "</td>"
		'[UI-ANT] else
			'[UI-ANT] 'html = html & "<td class='regis' ></td>"
			'[UI-ANT] html = html & "<td></td>"
		'[UI-ANT] end if

		'[UI-ANT] 'html = html & "<td class='regis' xonclick=""document.location='http://server-2025/corgestion/detalles.asp?id=" & rs(0) & "&opc=veri&cli=" & cliente & "&tit=" & tit & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'"">" & rs(0) & "</td>"
		'[UI-ANT] html = html & "<td onclick=""" & urlDetalle & """>" & rs(0) & "</td>"

		'[UI-ANT] if Trim(rs(7) & "") <> "" And Trim(rs(7) & "") <> "0" then
			'[UI-ANT] 'html = html & "<td class='regis' title='" & rs(11) & "'xonclick=""document.location='http://server-2025/corgestion/detalles.asp?id=" & rs(0) & "&opc=veri&cli=" & cliente & "&tit=" & tit & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'"">" & Left(UCASE(rs(11)),40) & "</td>"
			'[UI-ANT] html = html & "<td title='" & rs(11) & "' onclick=""" & urlDetalle & """>" & Left(UCASE(rs(11)),40) & "</td>"
		'[UI-ANT] else
			'[UI-ANT] 'html = html & "<td class='regis' xonclick=""document.location='http://server-2025/corgestion/detalles.asp?id=" & rs(0) & "&opc=veri&cli=" & cliente & "&tit=" & tit & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'"">" & Left(UCASE(rs(11)),40) & " <font style='color:#000FFF;'><b>(C)</b></font></td>"
			'[UI-ANT] html = html & "<td onclick=""" & urlDetalle & """>" & Left(UCASE(rs(11)),40) & " <span style='color:#2563eb; font-weight:bold;'>(C)</span></td>"
		'[UI-ANT] end if

		'[UI-ANT] 'html = html & "<td class='regis' xonclick=""document.location='http://server-2025/corgestion/detalles.asp?id=" & rs(0) & "&opc=veri&cli=" & cliente & "&tit=" & tit & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'"">" & rs(1) & "</td>"
		'[UI-ANT] html = html & "<td onclick=""" & urlDetalle & """>" & rs(1) & "</td>"
		
		'[UI-ANT] 'html = html & "<td class='regis' xonclick=""document.location='http://server-2025/corgestion/detalles.asp?id=" & rs(0) & "&opc=veri&cli=" & cliente & "&tit=" & tit & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'"">" & prio & "</td>"
		'[UI-ANT] html = html & "<td onclick=""" & urlDetalle & """>" & prio & "</td>"

		'[UI-ANT] if Trim(rs(7) & "") <> "" And Trim(rs(7) & "") <> "0" then
			'[UI-ANT] 'html = html & "<td class='regis' xonclick=""document.location='http://server-2025/corgestion/detalles.asp?id=" & rs(0) & "&opc=veri&cli=" & cliente & "&tit=" & tit & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'""><font style='font-size:12pt;'>" & est1 & "</font></td>"
			'[UI-ANT] html = html & "<td onclick=""" & urlDetalle & """>" & est1 & "</td>"
		'[UI-ANT] else
			'[UI-ANT] 'html = html & "<td class='regis' xonclick=""document.location='http://server-2025/corgestion/detalles.asp?id=" & rs(0) & "&opc=veri&cli=" & cliente & "&tit=" & tit & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'""><font style='font-size:12pt;'>-------------</td>"
			'[UI-ANT] html = html & "<td onclick=""" & urlDetalle & """><span style='color:#94a3b8;'>-------------</span></td>"
		'[UI-ANT] end if

		'[UI-ANT] 'html = html & "<td class='regis' xonclick=""document.location='http://server-2025/corgestion/detalles.asp?id=" & rs(0) & "&opc=veri&cli=" & cliente & "&tit=" & tit & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'""><font style='font-size:12pt;'>" & est2 & "</font></td>"
		'[UI-ANT] html = html & "<td onclick=""" & urlDetalle & """>" & est2 & "</td>"
		
		'[UI-ANT] sFec = Fechador(rs(5) + nFecini)
		
		'[UI-ANT] 'html = html & "<td class='regis' xonclick=""document.location='http://server-2025/corgestion/detalles.asp?id=" & rs(0) & "&opc=veri&cli=" & cliente & "&tit=" & tit & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'"">" & sFec & "</td>"
		'[UI-ANT] html = html & "<td onclick=""" & urlDetalle & """>" & sFec & "</td>"

		'[UI-ANT] 'html = html & "</tr>"
		'[UI-ANT] html = html & "</tr>"

		'[UI-ANT] PosReg = PosReg + 1
		'[UI-ANT] rs.movenext
	'[UI-ANT] Loop

	'[UI-ANT] 'html = html & "</table>"
	'[UI-ANT] html = html & "</tbody>"
	'[UI-ANT] html = html & "</table>"
	
	'[UI-ANT] ' --- CIERRES DE CONTENEDORES ---
	'[UI-ANT] html = html & "</div>" ' Fin .bloque-datos.table-responsive (UNIFICADO)
	'[UI-ANT] html = html & "</div>" ' Fin .inner

'[UI-ANT] 'xhtml = xhtml & "<div style='height:8%;width:100%;border-top:solid 5px #FF3D3D;border-bottom:double 6px #001BFF;'>"

'[UI-ANT] xhtml = xhtml & "<table class='titulo-seccion' >"
'[UI-ANT] xhtml = xhtml & "<tr style='display:flex;align-items:center;justify-content:space-between;'>"

'[UI-ANT] xhtml = xhtml & "<td><h2>INCIDENTES</h2></td>"



'[UI-ANT] if session("FUNC") = "FUNCw" or session("FUNC") = "TECw" then

	'[UI-ANT] formx4 = "<form id='formx4' style='height:70%;' method='post' action='agregar2.asp?tipo=INC&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
	'[UI-ANT] if getPermisos(xmenu,xtmenu,"ALTA") = 1 then
		'[UI-ANT] botonx4 = "<input class='bot' type='button' value='CARGAR NUEVO INCIDENTE' onclick=""document.getElementById('formx4').submit();"">"
	'[UI-ANT] else
		'[UI-ANT] botonx4 = "<input class='botsp' type='button' value='CARGAR NUEVO INCIDENTE' onclick=""alert('Sin permiso para agregar.');"">"
	'[UI-ANT] end if
	'[UI-ANT] xhtml = xhtml & "<td align=right valign=center>" & botonx4 & "</td>"

'[UI-ANT] end if

'[UI-ANT] formx5 = "<form id='formx5' style='height:70%;' method='post' action='index.asp'></form>"
'[UI-ANT] botonx5 = "<input class='bot' type='button' value='VOLVER' onclick=""document.getElementById('formx5').submit();"">"


'[UI-ANT] if cantregmover <> 0 then
	'[UI-ANT] formx22 = "<form id='formx22' style='height:1px;' method='post' action='todo.asp?opc=listar&cantregmover=" & (cantregmover - cantregmostrar) & "&ord=" & ord & "&cli=" & cli & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
	'[UI-ANT] botonx22 = "<input class='bot' type='button' value='ATRAS' onclick=""document.getElementById('formx22').submit();"">"
'[UI-ANT] else
	'[UI-ANT] botonx22 = "<input class='botsp' type='button' value='ATRAS'>"
'[UI-ANT] end if

'[UI-ANT] if not rs.eof then
	'[UI-ANT] formx23 = "<form id='formx23' style='height:1px;' method='post' action='todo.asp?opc=listar&cantregmover=" & (cantregmover + cantregmostrar) & "&ord=" & ord & "&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
	'[UI-ANT] botonx23 = "<input class='bot' type='button' value='SIGUIENTE' onclick=""document.getElementById('formx23').submit();"">"
'[UI-ANT] else
	'[UI-ANT] botonx23 = "<input class='botsp' type='button' value='SIGUIENTE'>"
'[UI-ANT] end if

	'[UI-ANT] 'xhtml = xhtml & "<td>" & botonx22 & "</td>"
	'[UI-ANT] 'xhtml = xhtml & "<td>" & botonx23 & "</td>"
	'[UI-ANT] xhtml = xhtml & "<td><div style='display: flex; gap: 0.5rem;'>" & botonx22 & botonx23 & "</div></td>"
	'[UI-ANT] xhtml = xhtml & "<td>" & botonx5 & "</td>"
	'[UI-ANT] xhtml = xhtml & "</tr></table>"

	'[UI-ANT] xhtml = xhtml & "</div>"

'[UI-ANT] html = xhtml & html

'[UI-ANT] 'html = html & "<br>"

'[UI-ANT] html = html & "<table>"

'[UI-ANT] html = html & "<tr>"
'[UI-ANT] html = html & "<td><font style='color:#000000;'><b>*</b></font><font style='color:#000FFF;'><b>(C)</b></font></td>"
'[UI-ANT] html = html & "<td><font style='color:#000000;'>&nbsp;&nbsp;(Consulta) resuelta sin incumbencia tecnica.</font></td>"
'[UI-ANT] html = html & "</tr>"

'[UI-ANT] html = html & "</table>"

'[UI-ANT] html = html & "</div>"

'[UI-ANT] 'CUANDO CONSULTAS INCIDENTES ENTRA POR ACA
'[UI-ANT] elseif opc = "consi" then

'[UI-ANT] sel = request.querystring("sel")
'[UI-ANT] xmenu = request.querystring("xmenu")
'[UI-ANT] xtmenu = request.querystring("xtmenu")
'[UI-ANT] XXX = SYSLOGON(cnn,"CONSULTA_INCIDENTES")


'[UI-ANT] html = html & "<div>"

'[UI-ANT] html = html & "<header id='header' style='background:#0e1949;padding:1rem 0;'>"

'[UI-ANT] html = html & "<div>"

'[UI-ANT] html = html & "<table style='width:100%;'><tr>"

'[UI-ANT] html = html & "<td style='width:75%;' valign=center><a href='index.asp' class='logo' >"
'[UI-ANT] html = html & "<img src='images/logocorP4.gif' width='10%' >"
'[UI-ANT] html = html & "<font style='margin-left:100px;font-size:30pt;align:center;valign:center;font-family:Serif;color:#000000;'>Cor Consulting</font>"
'[UI-ANT] html = html & "</a></td>"

'[UI-ANT] html = html & "</tr></table>"

'[UI-ANT] html = html & "</div>"

'[UI-ANT] html = html & "</header>"

'[UI-ANT] html = html & "</div>"

'[UI-ANT] 'html = html & "<div style='height:8%;width:100%;border-top:solid 5px #FF3D3D;border-bottom:double 6px #001BFF;'>"

'[UI-ANT] html = html & "<table class='titulo-seccion'>"

'[UI-ANT] html = html & "<tr style='display:flex;align-items:center;justify-content:space-between;'>"

'[UI-ANT] html = html & "<td><h2>CONSULTA INCIDENTES</h2></td>"

'[UI-ANT] formx10 = "<form id='formx10' style='height:70%;' method='post' action='consultas.asp?sel=" & sel & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
'[UI-ANT] botonx10 = "<input class='bot' type='button' value='VOLVER' onclick=""document.getElementById('formx10').submit();"">"

'[UI-ANT] html = html & "<td>" & botonx10 & "</td>"

'[UI-ANT] html = html & "</tr>"

'[UI-ANT] html = html & "</table>"

'[UI-ANT] html = html & "</div>"

'[UI-ANT] html = html & "<div style='height:77%;width:100%;'>"

'[UI-ANT] html = html & "<table style='width:85%;color:#000000;font-family:Serif;text-align:center;' cellspacing=0 align=center>"

'[UI-ANT] html = html & "<tr STYLE='height:40px;line-height:40px;background:#DAE8FF;'>"


'[UI-ANT] html = html & "<td class='encabe' style='font-family:tahoma;font-size:10pt;border-color:white;width:5%;'>ID</td>"
'[UI-ANT] html = html & "<td class='encabe' style='font-family:tahoma;font-size:10pt;border-color:white;width:40%;'>TITULO</td>"
'[UI-ANT] html = html & "<td class='encabe' style='border-color:white;width:1%;'>*</td>"
'[UI-ANT] html = html & "<td class='encabe' style='font-family:tahoma;font-size:10pt;border-color:white;width:18%;'>CLIENTE</td>"
'[UI-ANT] html = html & "<td class='encabe' style='font-family:tahoma;font-size:10pt;border-color:white;width:10.5%;'>DESARROLLO</td>"
'[UI-ANT] html = html & "<td class='encabe' style='font-family:tahoma;font-size:10pt;border-color:white;width:10.5%;'>FUNCIONAL</td>"
'[UI-ANT] html = html & "<td class='encabe' style='font-family:tahoma;font-size:10pt;border-color:white;width:15%;'>FEC INGRESO</td>"


'[UI-ANT] html = html & "</tr>"

'[UI-ANT] cmd = "sp_cg_ConsultarEntidad @Entidad='incidente', @Modo='consulta'"

	'[UI-ANT] cliente = request.form("cli")
	'[UI-ANT] if cliente & "" = "" then cliente = request.querystring("cli")
	'[UI-ANT] descri = request.form("desc")
	'[UI-ANT] if descri & "" = "" then descri = request.querystring("desc")
	'[UI-ANT] estado = request.form("est")
	'[UI-ANT] if estado & "" = "" then estado = request.querystring("est")
	'[UI-ANT] aplic = request.form("apli")
	'[UI-ANT] if aplic & "" = "" then aplic = request.querystring("apli")
	'[UI-ANT] fecha = request.form("fec")
	'[UI-ANT] if fecha & "" = "" then fecha = request.querystring("fec")


	'[UI-ANT] 'where cliente
	'[UI-ANT] if cliente <> 0 then
		'[UI-ANT] cmd = cmd & ", @IdCliente=" & cliente
	'[UI-ANT] end if

	'[UI-ANT] 'where descripcion
	'[UI-ANT] if descri & "" <> "" then
		'[UI-ANT] cmd = cmd & ", @Titulo='" & descri & "'"
	'[UI-ANT] end if

	'[UI-ANT] 'where estado
	'[UI-ANT] if estado <> 0 then
		'[UI-ANT] if estado = 1 then 'OK
			'[UI-ANT] cmd = cmd & ", @Estado=1"
		'[UI-ANT] elseif estado = 2 then 'PEND
			'[UI-ANT] cmd = cmd & ", @Estado=2"
		'[UI-ANT] elseif estado = 10 then 'BAJA
			'[UI-ANT] cmd = cmd & ", @Estado=10"
		'[UI-ANT] end if
	'[UI-ANT] end if

	'[UI-ANT] 'where aplicativo
	'[UI-ANT] if aplic <> 0 then
		'[UI-ANT] cmd = cmd & ", @IdAplicativo=" & aplic
	'[UI-ANT] end if

	'[UI-ANT] 'where fecha
	'[UI-ANT] if fecha & "" <> "" then
		'[UI-ANT] fecha = clng(cdate(fecha)) - nFecini
		'[UI-ANT] cmd = cmd & ", @Fecha=" & fecha
	'[UI-ANT] end if

'[UI-ANT] Set rs = Cnn.Execute(cmd)

'[UI-ANT] xform = ""
'[UI-ANT] sColor = "#FFFFFF"

'[UI-ANT] Do While Not rs.eof

'[UI-ANT] estD = ""
'[UI-ANT] estF = ""

'[UI-ANT] If rs(3) = "PENDIENTE" Then
	'[UI-ANT] estD = "<font style='color:#FF0000;'>PENDIENTE</font>"
'[UI-ANT] ElseIf rs(3) = "OK" Then
	'[UI-ANT] estD = "<font style='color:#088A08;'>OK</font>"
'[UI-ANT] ElseIf rs(3) = "BAJA" Then
	'[UI-ANT] estD = "<font style='color:#681415;'>BAJA</font>"
'[UI-ANT] End If

'[UI-ANT] If rs(4) = "PENDIENTE" Then
	'[UI-ANT] estF = "<font style='color:#FF0000;'>PENDIENTE</font>"
'[UI-ANT] ElseIf rs(4) = "OK" Then
	'[UI-ANT] estF = "<font style='color:#088A08;'>OK</font>"
'[UI-ANT] ElseIf rs(4) = "BAJA" Then
	'[UI-ANT] estF = "<font style='color:#681415;'>BAJA</font>"
'[UI-ANT] End If

'[UI-ANT] xform = xform & "<form id='xform" & rs(0) & "' method='post' action='detalles.asp?id=" & rs(0) & "&opc=veri&tipo=cons&" & request.form & "'></form>"
'[UI-ANT] if getPermisos(9,"INICIO","CONS") = 1 then
	'[UI-ANT] xboton = "<input class='bot' style='border-radius:5px;background:#DAE8FF;color:#008520;' type='button' value='D' onclick=""document.getElementById('xform" & rs(0) & "').submit();"">"
'[UI-ANT] else
	'[UI-ANT] xboton = "<input class='botsp' type='button' value='D' onclick=""alert('Sin permiso para consultar.');"">"
'[UI-ANT] end if

'[UI-ANT] html = html & "<tr>"

'[UI-ANT] html = html & "<td xclass='regis' style='font-family:tahoma;font-size:9pt;xpadding:8px;text-align:center;background-color:" & sColor & ";'>" & rs(0) & "</td>"
'[UI-ANT] html = html & "<td xclass='regis' style='font-family:tahoma;font-size:9pt;padding:8px;text-align:left;background-color:" & sColor & ";'>" & mid(ucase(rs(1)),1,58) & "</td>" 'UCASE(rs(1))'
'[UI-ANT] 'html = html & "<td xclass='regis' >" & xboton & "</td>"
'[UI-ANT] html = html & "<td xclass='regis' style='font-family:tahoma;font-size:9pt;padding-left:25px;text-align:left;background-color:" & sColor & ";'>" & rs(2) & "</td>"
'[UI-ANT] html = html & "<td xclass='regis' style='font-family:tahoma;font-size:9pt;xpadding-left:25px;text-align:center;background-color:" & sColor & ";'>" & estD & "</td>"
'[UI-ANT] html = html & "<td class='regis' style='font-family:tahoma;font-size:9pt;xpadding-left:25px;text-align:center;background-color:" & sColor & ";'>" & estF & "</td>"
'[UI-ANT] html = html & "<td class='regis' style='font-family:tahoma;font-size:9pt;xpadding-left:25px;text-align:center;background-color:" & sColor & ";'>" &  Fechador(rs(5) + nFecini) & "</td>"

'[UI-ANT] html = html & "</tr>"

'[UI-ANT] if sColor = "#FFFFFF" then
	'[UI-ANT] sColor = "#ECECEC"
'[UI-ANT] elseif sColor = "#ECECEC" then
	'[UI-ANT] sColor = "#FFFFFF"
'[UI-ANT] end if

'[UI-ANT] rs.movenext
'[UI-ANT] loop

'[UI-ANT] html = html & "</table>"


'[UI-ANT] end if



'[UI-ANT] html = html & "<script type='text/javascript'>"
'[UI-ANT] html = html & " function borrar(formu){ "
'[UI-ANT] html = html & " var r = confirm('Desea eliminar incidente?'); "
'[UI-ANT] html = html & " if (r){ "
'[UI-ANT] html = html & " document.getElementById(formu).submit(); "
'[UI-ANT] html = html & " }else{ "
'[UI-ANT] html = html & " return false;"
'[UI-ANT] html = html & " }} "
'[UI-ANT] html = html & "</script>"



'[UI-ANT] html = html & "<script src='assets/js/jquery.min.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/skel.min.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/util.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/main.js'></script>"

'[UI-ANT] html = html & formx22 & formx23 & formx4 & formx5 & xform & formx10 & formx1 & formx2 & formx3 & "</body>"
'[UI-ANT] html = html & "</html>"

	' =========================================================================
	' [UI] VISTA con el sistema visual (IncFile/ui_componentes.inc)
	' El bloque anterior completo queda arriba, comentado con '[UI-ANT].
	' Procedimientos, formularios y permisos: los mismos de antes.
	' =========================================================================
	Dim accTop, urlDetalle, irDetalle, cantTodo, marcaConsulta, celdaDesa, claseFila, esConsulta, qsOrden

	html = html & "<div id='wrapper' role='banner'>" & cabecera() & "</div>"

	' --- Barra superior ---
	' "Cargar incidente" conserva la condicion de antes (perfiles FUNCw / TECw): hoy no se muestra
	accTop = ""
	if session("FUNC") = "FUNCw" or session("FUNC") = "TECw" then
		formx4 = "<form id='formx4' class='form-oculto' method='post' action='agregar2.asp?tipo=INC&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
		accTop = accTop & UiBoton("Cargar incidente", "light", "document.getElementById('formx4').submit();", getPermisosAtrm("CG_INCIDENTES","ALTA") = 1, "Sin permiso para agregar.")
	end if
	formx5 = "<form id='formx5' class='form-oculto' method='post' action='index.asp'></form>"
	accTop = accTop & UiBoton("Volver", "secondary", "document.getElementById('formx5').submit();", True, "")
	html = html & UiPantallaAbre("Para hacer", accTop, "wide")
	html = html & UiCardAbre("Incidentes por atender", "")

	' [UI] Titulo y orden escapados, cliente validado como numero: antes se pegaban tal cual en la consulta
	If IsNumeric(cliente & "") Then cliente = CLng(cliente) Else cliente = 0
	cmd = "sp_cg_ConsultarEntidad @Entidad='incidente', @Modo='todo', @IdCliente=" & cliente & ", @Titulo=" & SqlStr(tit) & ", @Orden=" & SqlStr(ord)
	Set rs = Cnn.Execute(cmd)
	if not rs.eof then rs.Move (CantRegMover)

	' Ordenar se queda en esta pantalla (antes los encabezados llevaban a incidentes.asp)
	qsOrden = "&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu
	html = html & UiTablaAbre("Incidentes por atender. Se puede ordenar por t&iacute;tulo o por cliente.")
	html = html & "<thead><tr>" & UiTh("ID", "") & _
		UiThOrden("T&iacute;tulo", "todo.asp?ord=titulo" & qsOrden, ord = "titulo") & _
		UiThOrden("Cliente", "todo.asp?ord=cliente" & qsOrden, ord = "cliente") & _
		UiTh("Prioridad", "") & UiTh("Desarrollo", "") & UiTh("Funcional", "") & UiTh("Ingreso", "") & _
		"</tr></thead><tbody>"

	' Se muestran todos los registros, como antes (el limite por pagina estaba desactivado)
	cantTodo = 0
	Do While Not rs.eof

		urlDetalle = "detalles.asp?id=" & rs(0) & "&opc=veri&cli=" & cliente & "&tit=" & Server.URLEncode(tit & "") & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu
		irDetalle = " class='ui-table__cell--link' onclick=""document.location='" & urlDetalle & "'"""

		' Bancos del Grupo Petersen: fila resaltada (el mismo criterio de antes)
		claseFila = ""
		if rs(1) & "" = "GRUPO PETERSEN" or rs(1) & "" = "BANCO DE ENTRE RIOS" or rs(1) & "" = "BANCO DE SAN JUAN" or rs(1) & "" = "BANCO DE SANTA CRUZ" or rs(1) & "" = "BANCO DE SANTA FE" then
			claseFila = " class='ui-table__fila--grupo'"
		end if

		' Consulta: sin responsable tecnico
		esConsulta = (Trim(rs(7) & "") = "" Or Trim(rs(7) & "") = "0")
		marcaConsulta = ""
		If esConsulta Then marcaConsulta = " <span class='ui-badge ui-badge--blue' title='Resuelta sin incumbencia t&eacute;cnica'>Consulta</span>"
		If esConsulta Then
			celdaDesa = "<span class='ui-status-empty'>&mdash;</span>"
		Else
			celdaDesa = UiBadgeEstado(rs(3)) & UiSub(UiTexto(rs(9)))
		End If

		html = html & "<tr" & claseFila & ">" & _
			"<td" & irDetalle & ">" & rs(0) & "</td>" & _
			"<td>" & UiCeldaEnlace(UiTexto(rs(11)), urlDetalle, UiTexto(rs(11))) & marcaConsulta & "</td>" & _
			"<td" & irDetalle & ">" & UiTexto(rs(1)) & "</td>" & _
			"<td" & irDetalle & ">" & UiBadgePrioridad(rs(2)) & "</td>" & _
			"<td" & irDetalle & ">" & celdaDesa & "</td>" & _
			"<td" & irDetalle & ">" & UiBadgeEstado(rs(4)) & UiSub(UiTexto(rs(8))) & "</td>" & _
			"<td" & irDetalle & "><span class='ui-table__nowrap'>" & UiFecha(rs(5)) & "</span></td>" & _
			"</tr>"

		cantTodo = cantTodo + 1
		rs.movenext
	Loop

	If cantTodo = 0 Then html = html & UiFilaVacia(7, "No hay incidentes por atender.")
	html = html & UiTablaCierra()
	html = html & "<p class='ui-table__nota'><span class='ui-badge ui-badge--blue'>Consulta</span> Resuelta sin incumbencia t&eacute;cnica. &nbsp; <span class='ui-ref-grupo' aria-hidden='true'></span>Bancos del Grupo Petersen.</p>"
	If cantTodo = 1 Then
		html = html & UiPie("1 incidente", "")
	ElseIf cantTodo > 1 Then
		html = html & UiPie(cantTodo & " incidentes", "")
	End If
	html = html & UiCardCierra()
	html = html & UiPantallaCierra()

elseif opc = "consi" then

	sel = request.querystring("sel")
	xmenu = request.querystring("xmenu")
	xtmenu = request.querystring("xtmenu")
	XXX = SYSLOGON(cnn,"CONSULTA_INCIDENTES")

	Dim accCons, estCons, cantCons, puedeConsultar

	html = html & "<div id='wrapper' role='banner'>" & cabecera() & "</div>"

	formx10 = "<form id='formx10' class='form-oculto' method='post' action='consultas.asp?sel=" & sel & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
	accCons = UiBoton("Volver", "secondary", "document.getElementById('formx10').submit();", True, "")
	html = html & UiPantallaAbre("Consulta de incidentes", accCons, "wide")
	html = html & UiCardAbre("Resultados", "")

	' --- Filtros (los mismos de antes), validados antes de ir a la consulta ---
	cmd = "sp_cg_ConsultarEntidad @Entidad='incidente', @Modo='consulta'"
	cliente = request.form("cli")
	if cliente & "" = "" then cliente = request.querystring("cli")
	descri = request.form("desc")
	if descri & "" = "" then descri = request.querystring("desc")
	estado = request.form("est")
	if estado & "" = "" then estado = request.querystring("est")
	aplic = request.form("apli")
	if aplic & "" = "" then aplic = request.querystring("apli")
	fecha = request.form("fec")
	if fecha & "" = "" then fecha = request.querystring("fec")

	' [UI] Numeros validados, texto escapado y fecha verificada: antes se pegaban tal cual en la consulta
	If IsNumeric(cliente & "") Then cliente = CLng(cliente) Else cliente = 0
	If IsNumeric(estado & "") Then estado = CLng(estado) Else estado = 0
	If IsNumeric(aplic & "") Then aplic = CLng(aplic) Else aplic = 0
	if cliente <> 0 then cmd = cmd & ", @IdCliente=" & cliente
	if descri & "" <> "" then cmd = cmd & ", @Titulo=" & SqlStr(descri)
	if estado = 1 or estado = 2 or estado = 10 then cmd = cmd & ", @Estado=" & estado
	if aplic <> 0 then cmd = cmd & ", @IdAplicativo=" & aplic
	if IsDate(fecha & "") then cmd = cmd & ", @Fecha=" & (clng(cdate(fecha)) - nFecini)

	Set rs = Cnn.Execute(cmd)

	html = html & UiTablaAbre("Resultados de la consulta de incidentes")
	html = html & "<thead><tr>" & UiTh("ID", "") & UiTh("T&iacute;tulo", "") & UiTh("Cliente", "") & UiTh("Desarrollo", "") & _
		UiTh("Funcional", "") & UiTh("Ingreso", "") & UiTh("<span class='ui-sr-only'>Acciones</span>", "ui-table__col--actions") & "</tr></thead><tbody>"

	xform = ""
	cantCons = 0
	puedeConsultar = (getPermisosAtrm("CG_INCIDENTES","CONS") = 1)
	Do While Not rs.eof
		xform = xform & "<form id='xform" & rs(0) & "' class='form-oculto' method='post' action='detalles.asp?id=" & rs(0) & "&opc=veri&tipo=cons&" & request.form & "'></form>"
		html = html & "<tr>" & _
			"<td>" & rs(0) & "</td>" & _
			"<td><span class='ui-table__titulo'>" & UiTexto(rs(1)) & "</span></td>" & _
			"<td>" & UiTexto(rs(2)) & "</td>" & _
			"<td>" & UiBadgeEstado(rs(3)) & "</td>" & _
			"<td>" & UiBadgeEstado(rs(4)) & "</td>" & _
			"<td><span class='ui-table__nowrap'>" & UiFecha(rs(5)) & "</span></td>" & _
			"<td>" & UiBoton("Ver detalle", "secondary", "document.getElementById('xform" & rs(0) & "').submit();", puedeConsultar, "Sin permiso para consultar.") & "</td>" & _
			"</tr>"
		cantCons = cantCons + 1
		rs.movenext
	Loop
	If cantCons = 0 Then html = html & UiFilaVacia(7, "No hay incidentes que coincidan con la consulta.")
	html = html & UiTablaCierra()
	If cantCons = 1 Then
		html = html & UiPie("1 resultado", "")
	ElseIf cantCons > 1 Then
		html = html & UiPie(cantCons & " resultados", "")
	End If
	html = html & UiCardCierra()
	html = html & UiPantallaCierra()

end if

html = html & "<script src='assets/js/jquery.min.js'></script>"
html = html & "<script src='assets/js/skel.min.js'></script>"
html = html & "<script src='assets/js/util.js'></script>"
html = html & "<script src='assets/js/main.js'></script>"

html = html & formx4 & formx5 & xform & formx10 & "</body>"
html = html & "</html>"

response.write html

Function armaCombo(sql,cn,clave,valor,texto,comboid)

	Set nrs = cn.Execute(sql)

	html = html & "<select style='font-size:12pt;height:80%;color:#000000;border:solid 1px #000000;' id='" & comboid & "' name='" & comboid & "'>"



	Do While Not nrs.eof

		If clave = nrs(valor) Then

			html = html & "<option value='" & nrs(valor) & "' selected>" & nrs(texto) & "</option>"

		Else

			html = html & "<option value='" & nrs(valor) & "'>" & nrs(texto) & "</option>"

		End If

		nrs.movenext
	Loop

	armaCombo = html & "</select>"

End Function


%>