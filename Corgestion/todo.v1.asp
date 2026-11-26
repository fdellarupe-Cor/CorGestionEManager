	<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<%

html = "<html>"
html = html & "<head>"

html = html & "<title>Cor Gestion</title>"
html = html & "<meta charset='utf-8' />"
html = html & "<meta name='viewport' content='width=device-width, initial-scale=1' />"
html = html & "<link rel='stylesheet' href='assets/css/main.css' />"
html = html & "<style style='text/css'>"
html = html & ".tablX{"
html = html & "width:100%; "
html = html & "border-collapse:collapse; "
html = html & "}"
html = html & ".tablX td{ "
html = html & "padding:7px; "
html = html & "}"
html = html & ".tablX tr{"
html = html & "background: #FFFFFF;"
html = html & "}"
html = html & ".tablX tr:hover {"
html = html & "background-color: #D4D4D4;"
html = html & "}"
html = html & "</style>"

html = html & "</head>"
html = html & "<body>"

Set Cnn = Server.CreateObject("ADODB.Connection")
Cnn.open "Provider=SQLOLEDB;Data Source=server-2025;Initial Catalog=dbcorgestionv5; user id=sa;password=bb4515$;"

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
CantRegMover = CDbl(Request.querystring("CantRegMover"))
if CantRegMover < 0 then CantRegMover = 0


html = html & "<div>"

html = html & "<header id='header' style='background:#0e1949;padding:1rem 0;'>"

html = html & "<div >"

html = html & "<table style='width:100%;'><tr style='display:flex;padding: 0 1rem;'>"

html = html & "<td style='width:75%;' valign=center><a href='index.asp' class='logo' >"
html = html & "<img src='images/iso_cor.png' style='width:6rem;' >"
html = html & "<font style='font-size:15pt;color:white;'><b>Cor Gestion</b> | Sistema Administrativo</font></a></td>"
html = html & "<td style='width:25%;display:flex;align-items:center;flex-direction:row-reverse;' align='right'>"
html = html & "<nav>"
html = html & "<ul>"
'html = html & "<li><a href='#menu'></a></li>"
html = html & "</ul>"
html = html & "</nav>"
html = html & "</td>"

html = html & "</tr></table>"

html = html & "</div>"

html = html & "</header>"

response.write html
''''''''''''''''''''''''''
html = ""
	'html = html & "<div style='height:77%;width:100%;'><br>"

	' --- INICIO NORMALIZACION: TABLA TO-DO (INCIDENTES) ---
	' Contenedor Inner
	html = html & "<div class='inner' style='padding: 2rem;'>"
	
	' ¡CORRECCIÓN!: Clases unificadas en un solo DIV para respetar el border-radius
	html = html & "<div class='bloque-datos table-responsive'>"
	
	'html = html & "<table class='tablX' style='width:99%;color:#000000;font-family:Serif;text-align:center;' cellspacing=0 align=center>"
	html = html & "<table class='tabla-moderna'>"
	html = html & "<thead>"
	
	'html = html & "<tr>"
	html = html & "<tr>"
	
	'html = html & "<td class='encabe' style='width:1%;'>*</td>"
	html = html & "<th style='width:1%;'></th>"
	
	'html = html & "<td class='encabe' style='width:1%;'>*</td>"
	html = html & "<th style='width:1%;'></th>"
	
	'html = html & "<td class='encabe' style='width:1%;'>ID</td>"
	html = html & "<th style='width:1%;'>ID</th>"
	
	'html = html & "<td class='encabe' style='width:28%;'><a class='red' href='incidentes.asp?ord=titulo&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>TITULO</a></td>"
	html = html & "<th style='width:28%;'><a class='col-link' style='text-decoration:none;' href='incidentes.asp?ord=titulo&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>TITULO</a></th>"
	
	'html = html & "<td class='encabe' style='width:15%;'><a class='red' href='incidentes.asp?ord=cliente&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>CLIENTE</a></td>"
	html = html & "<th style='width:15%;'><a class='col-link' style='text-decoration:none;' href='incidentes.asp?ord=cliente&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>CLIENTE</a></th>"
	
	'html = html & "<td class='encabe' style='width:10%;'>PRIORIDAD</td>"
	html = html & "<th style='width:10%;'>PRIORIDAD</th>"
	
	'html = html & "<td class='encabe' style='width:15%;'>DESARROLLO</td>"
	html = html & "<th style='width:15%;'>DESARROLLO</th>"
	
	'html = html & "<td class='encabe' style='width:15%;'>FUNCIONAL</td>"
	html = html & "<th style='width:15%;'>FUNCIONAL</th>"
	
	'html = html & "<td class='encabe' style='width:12%;'>FEC INGRESO</td>"
	html = html & "<th style='width:12%;'>FEC INGRESO</th>"
	
	'html = html & "</tr>"
	html = html & "</tr>"
	html = html & "</thead>"
	html = html & "<tbody>"

	cmd = "sp_cg_ConsultarEntidad @Entidad='incidente', @Modo='todo', @IdCliente=" & cliente & ", @Titulo='" & tit & "', @Orden='" & ord & "'"

	Set rs = Cnn.Execute(cmd)

	if not rs.eof then rs.Move (CantRegMover)

	PosReg = 0
	sColor = "#FFFFFF"
	formx1 = ""
	formx2 = ""
	formx3 = ""

	Do While (Not rs.eof) 'And (PosReg < CantRegMostrar)

		' Lógica de color de fondo específico por cliente
		sColor = ""
		if rs(1) & "" = "GRUPO PETERSEN" or rs(1) & "" = "BANCO DE ENTRE RIOS" or rs(1) & "" = "BANCO DE SAN JUAN" or rs(1) & "" = "BANCO DE SANTA CRUZ" or rs(1) & "" = "BANCO DE SANTA FE" then
			sColor = "background-color: #F3E5D2;"
		end if

		est1 = ""
		est2 = ""
		prio = ""

		' --- BADGES DE ESTADO (DESARROLLO) ---
		If rs(3) = "PENDIENTE" Then
			est1 = "<span class='badge badge-media'>PEND</span> (" & rs(9) & ")"
		ElseIf rs(3) = "OK" Then
			est1 = "<span class='badge badge-ok'>OK</span> (" & rs(9) & ")"
		ElseIf rs(3) = "BAJA" Then
			est1 = "<span class='badge badge-peligro'>BAJA</span> (" & rs(9) & ")"
		ElseIf rs(3) = "TESTING EN CLIENTE" Then
			est1 = "<span class='badge badge-media'>TEST</span> (" & rs(9) & ")"
		End If

		' --- BADGES DE ESTADO (FUNCIONAL) ---
		If rs(4) = "PENDIENTE" Then
			est2 = "<span class='badge badge-media'>PEND</span> (" & rs(8) & ")"
		ElseIf rs(4) = "OK" Then
			est2 = "<span class='badge badge-ok'>OK</span> (" & rs(8) & ")"
		ElseIf rs(4) = "BAJA" Then
			est2 = "<span class='badge badge-peligro'>BAJA</span> (" & rs(8) & ")"
		ElseIf rs(4) = "TESTING EN CLIENTE" Then
			est2 = "<span class='badge badge-media'>TEST</span> (" & rs(8) & ")"
		End If

		' --- BADGES DE PRIORIDAD ---
		if rs(2) = "ALTA" then
			prio = "<span class='badge badge-peligro'>ALTA</span>"
		elseif rs(2) = "MEDIA" then
			prio = "<span class='badge badge-media'>MEDIA</span>"
		else
			prio = "<span class='badge badge-pendiente'>" & rs(2) & "</span>"
		end if

		' --- BOTON ELIMINAR ---
		formx1 = formx1 & "<form id='formx" & rs(0) & "' style='height:1px;' method='post' action='eliminar.asp?opc=inc&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=incidentes.asp'></form>"
		if getPermisos(xmenu,xtmenu,"ELIM") = 1 then
			botonx1 = "<input class='btn-accion btn-accion-x' type='button' value='X' onclick=""borrar('formx" & rs(0) & "');"">"
		else
			botonx1 = "<input class='btn-accion btn-accion-deshabilitado' type='button' value='X' onclick=""alert('Sin permiso para eliminar.');"">"
		end if
		botonx1="" ' Mantenemos la lógica original que anula el botón

		' --- BOTON ACTUALIZAR ---
		formx2 = formx2 & "<form id='formx2" & rs(0) & "' style='height:1px;' method='post' action='detalles.asp?id=" & rs(0) & "&opc=acti&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
		if getPermisos(xmenu,xtmenu,"MODIF") = 1 then
			botonx2 = "<input class='btn-accion btn-accion-a' type='button' value='A' onclick=""document.getElementById('formx2" & rs(0) & "').submit();"">"
		else
			botonx2 = "<input class='btn-accion btn-accion-deshabilitado' type='button' value='A' onclick=""alert('Sin permiso para modificar.');"">"
		end if
		botonx2="" ' Mantenemos la lógica original que anula el botón

		' URL unificada para el evento onclick de las filas
		urlDetalle = "document.location='detalles.asp?id=" & rs(0) & "&opc=veri&cli=" & cliente & "&tit=" & tit & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'"

		' --- INICIO FILA DE TABLA ---
		'html = html & "<tr>"
		html = html & "<tr style='" & sColor & "'>"
		
		'html = html & "<td class='regis'>" & botonx1s & "</td>"
		html = html & "<td>" & botonx1 & "</td>"

		if session("USUARIO") = rs(6) or session("USUARIO") = rs(7) then
			'html = html & "<td class='regis' >" & botonx2 & "</td>"
			html = html & "<td>" & botonx2 & "</td>"
		else
			'html = html & "<td class='regis' ></td>"
			html = html & "<td></td>"
		end if

		'html = html & "<td class='regis' xonclick=""document.location='http://server-2025/corgestion/detalles.asp?id=" & rs(0) & "&opc=veri&cli=" & cliente & "&tit=" & tit & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'"">" & rs(0) & "</td>"
		html = html & "<td onclick=""" & urlDetalle & """>" & rs(0) & "</td>"

		if Trim(rs(7) & "") <> "" And Trim(rs(7) & "") <> "0" then
			'html = html & "<td class='regis' title='" & rs(11) & "'xonclick=""document.location='http://server-2025/corgestion/detalles.asp?id=" & rs(0) & "&opc=veri&cli=" & cliente & "&tit=" & tit & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'"">" & Left(UCASE(rs(11)),40) & "</td>"
			html = html & "<td title='" & rs(11) & "' onclick=""" & urlDetalle & """>" & Left(UCASE(rs(11)),40) & "</td>"
		else
			'html = html & "<td class='regis' xonclick=""document.location='http://server-2025/corgestion/detalles.asp?id=" & rs(0) & "&opc=veri&cli=" & cliente & "&tit=" & tit & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'"">" & Left(UCASE(rs(11)),40) & " <font style='color:#000FFF;'><b>(C)</b></font></td>"
			html = html & "<td onclick=""" & urlDetalle & """>" & Left(UCASE(rs(11)),40) & " <span style='color:#2563eb; font-weight:bold;'>(C)</span></td>"
		end if

		'html = html & "<td class='regis' xonclick=""document.location='http://server-2025/corgestion/detalles.asp?id=" & rs(0) & "&opc=veri&cli=" & cliente & "&tit=" & tit & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'"">" & rs(1) & "</td>"
		html = html & "<td onclick=""" & urlDetalle & """>" & rs(1) & "</td>"
		
		'html = html & "<td class='regis' xonclick=""document.location='http://server-2025/corgestion/detalles.asp?id=" & rs(0) & "&opc=veri&cli=" & cliente & "&tit=" & tit & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'"">" & prio & "</td>"
		html = html & "<td onclick=""" & urlDetalle & """>" & prio & "</td>"

		if Trim(rs(7) & "") <> "" And Trim(rs(7) & "") <> "0" then
			'html = html & "<td class='regis' xonclick=""document.location='http://server-2025/corgestion/detalles.asp?id=" & rs(0) & "&opc=veri&cli=" & cliente & "&tit=" & tit & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'""><font style='font-size:12pt;'>" & est1 & "</font></td>"
			html = html & "<td onclick=""" & urlDetalle & """>" & est1 & "</td>"
		else
			'html = html & "<td class='regis' xonclick=""document.location='http://server-2025/corgestion/detalles.asp?id=" & rs(0) & "&opc=veri&cli=" & cliente & "&tit=" & tit & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'""><font style='font-size:12pt;'>-------------</td>"
			html = html & "<td onclick=""" & urlDetalle & """><span style='color:#94a3b8;'>-------------</span></td>"
		end if

		'html = html & "<td class='regis' xonclick=""document.location='http://server-2025/corgestion/detalles.asp?id=" & rs(0) & "&opc=veri&cli=" & cliente & "&tit=" & tit & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'""><font style='font-size:12pt;'>" & est2 & "</font></td>"
		html = html & "<td onclick=""" & urlDetalle & """>" & est2 & "</td>"
		
		sFec = Fechador(rs(5) + nFecini)
		
		'html = html & "<td class='regis' xonclick=""document.location='http://server-2025/corgestion/detalles.asp?id=" & rs(0) & "&opc=veri&cli=" & cliente & "&tit=" & tit & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'"">" & sFec & "</td>"
		html = html & "<td onclick=""" & urlDetalle & """>" & sFec & "</td>"

		'html = html & "</tr>"
		html = html & "</tr>"

		PosReg = PosReg + 1
		rs.movenext
	Loop

	'html = html & "</table>"
	html = html & "</tbody>"
	html = html & "</table>"
	
	' --- CIERRES DE CONTENEDORES ---
	html = html & "</div>" ' Fin .bloque-datos.table-responsive (UNIFICADO)
	html = html & "</div>" ' Fin .inner

'xhtml = xhtml & "<div style='height:8%;width:100%;border-top:solid 5px #FF3D3D;border-bottom:double 6px #001BFF;'>"

xhtml = xhtml & "<table class='titulo-seccion' >"
xhtml = xhtml & "<tr style='display:flex;align-items:center;justify-content:space-between;'>"

xhtml = xhtml & "<td><h2>INCIDENTES</h2></td>"



if session("FUNC") = "FUNCw" or session("FUNC") = "TECw" then

	formx4 = "<form id='formx4' style='height:70%;' method='post' action='agregar2.asp?tipo=INC&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
	if getPermisos(xmenu,xtmenu,"ALTA") = 1 then
		botonx4 = "<input class='bot' type='button' value='CARGAR NUEVO INCIDENTE' onclick=""document.getElementById('formx4').submit();"">"
	else
		botonx4 = "<input class='botsp' type='button' value='CARGAR NUEVO INCIDENTE' onclick=""alert('Sin permiso para agregar.');"">"
	end if
	xhtml = xhtml & "<td align=right valign=center>" & botonx4 & "</td>"

end if

formx5 = "<form id='formx5' style='height:70%;' method='post' action='index.asp'></form>"
botonx5 = "<input class='bot' type='button' value='VOLVER' onclick=""document.getElementById('formx5').submit();"">"


if cantregmover <> 0 then
	formx22 = "<form id='formx22' style='height:1px;' method='post' action='todo.asp?opc=listar&cantregmover=" & (cantregmover - cantregmostrar) & "&ord=" & ord & "&cli=" & cli & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
	botonx22 = "<input class='bot' type='button' value='ATRAS' onclick=""document.getElementById('formx22').submit();"">"
else
	botonx22 = "<input class='botsp' type='button' value='ATRAS'>"
end if

if not rs.eof then
	formx23 = "<form id='formx23' style='height:1px;' method='post' action='todo.asp?opc=listar&cantregmover=" & (cantregmover + cantregmostrar) & "&ord=" & ord & "&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
	botonx23 = "<input class='bot' type='button' value='SIGUIENTE' onclick=""document.getElementById('formx23').submit();"">"
else
	botonx23 = "<input class='botsp' type='button' value='SIGUIENTE'>"
end if

	'xhtml = xhtml & "<td>" & botonx22 & "</td>"
	'xhtml = xhtml & "<td>" & botonx23 & "</td>"
	xhtml = xhtml & "<td><div style='display: flex; gap: 0.5rem;'>" & botonx22 & botonx23 & "</div></td>"
	xhtml = xhtml & "<td>" & botonx5 & "</td>"
	xhtml = xhtml & "</tr></table>"

	xhtml = xhtml & "</div>"

html = xhtml & html

'html = html & "<br>"

html = html & "<table>"

html = html & "<tr>"
html = html & "<td><font style='color:#000000;'><b>*</b></font><font style='color:#000FFF;'><b>(C)</b></font></td>"
html = html & "<td><font style='color:#000000;'>&nbsp;&nbsp;(Consulta) resuelta sin incumbencia tecnica.</font></td>"
html = html & "</tr>"

html = html & "</table>"

html = html & "</div>"

'CUANDO CONSULTAS INCIDENTES ENTRA POR ACA
elseif opc = "consi" then

sel = request.querystring("sel")
xmenu = request.querystring("xmenu")
xtmenu = request.querystring("xtmenu")
XXX = SYSLOGON(cnn,"CONSULTA_INCIDENTES")


html = html & "<div>"

html = html & "<header id='header' style='background:#0e1949;padding:1rem 0;'>"

html = html & "<div>"

html = html & "<table style='width:100%;'><tr>"

html = html & "<td style='width:75%;' valign=center><a href='index.asp' class='logo' >"
html = html & "<img src='images/logocorP4.gif' width='10%' >"
html = html & "<font style='margin-left:100px;font-size:30pt;align:center;valign:center;font-family:Serif;color:#000000;'>Cor Consulting</font>"
html = html & "</a></td>"

html = html & "</tr></table>"

html = html & "</div>"

html = html & "</header>"

html = html & "</div>"

'html = html & "<div style='height:8%;width:100%;border-top:solid 5px #FF3D3D;border-bottom:double 6px #001BFF;'>"

html = html & "<table class='titulo-seccion'>"

html = html & "<tr style='display:flex;align-items:center;justify-content:space-between;'>"

html = html & "<td><h2>CONSULTA INCIDENTES</h2></td>"

formx10 = "<form id='formx10' style='height:70%;' method='post' action='consultas.asp?sel=" & sel & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
botonx10 = "<input class='bot' type='button' value='VOLVER' onclick=""document.getElementById('formx10').submit();"">"

html = html & "<td>" & botonx10 & "</td>"

html = html & "</tr>"

html = html & "</table>"

html = html & "</div>"

html = html & "<div style='height:77%;width:100%;'>"

html = html & "<table style='width:85%;color:#000000;font-family:Serif;text-align:center;' cellspacing=0 align=center>"

html = html & "<tr STYLE='height:40px;line-height:40px;background:#DAE8FF;'>"


html = html & "<td class='encabe' style='font-family:tahoma;font-size:10pt;border-color:white;width:5%;'>ID</td>"
html = html & "<td class='encabe' style='font-family:tahoma;font-size:10pt;border-color:white;width:40%;'>TITULO</td>"
html = html & "<td class='encabe' style='border-color:white;width:1%;'>*</td>"
html = html & "<td class='encabe' style='font-family:tahoma;font-size:10pt;border-color:white;width:18%;'>CLIENTE</td>"
html = html & "<td class='encabe' style='font-family:tahoma;font-size:10pt;border-color:white;width:10.5%;'>DESARROLLO</td>"
html = html & "<td class='encabe' style='font-family:tahoma;font-size:10pt;border-color:white;width:10.5%;'>FUNCIONAL</td>"
html = html & "<td class='encabe' style='font-family:tahoma;font-size:10pt;border-color:white;width:15%;'>FEC INGRESO</td>"


html = html & "</tr>"

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


	'where cliente
	if cliente <> 0 then
		cmd = cmd & ", @IdCliente=" & cliente
	end if

	'where descripcion
	if descri & "" <> "" then
		cmd = cmd & ", @Titulo='" & descri & "'"
	end if

	'where estado
	if estado <> 0 then
		if estado = 1 then 'OK
			cmd = cmd & ", @Estado=1"
		elseif estado = 2 then 'PEND
			cmd = cmd & ", @Estado=2"
		elseif estado = 10 then 'BAJA
			cmd = cmd & ", @Estado=10"
		end if
	end if

	'where aplicativo
	if aplic <> 0 then
		cmd = cmd & ", @IdAplicativo=" & aplic
	end if

	'where fecha
	if fecha & "" <> "" then
		fecha = clng(cdate(fecha)) - nFecini
		cmd = cmd & ", @Fecha=" & fecha
	end if

Set rs = Cnn.Execute(cmd)

xform = ""
sColor = "#FFFFFF"

Do While Not rs.eof

estD = ""
estF = ""

If rs(3) = "PENDIENTE" Then
	estD = "<font style='color:#FF0000;'>PENDIENTE</font>"
ElseIf rs(3) = "OK" Then
	estD = "<font style='color:#088A08;'>OK</font>"
ElseIf rs(3) = "BAJA" Then
	estD = "<font style='color:#681415;'>BAJA</font>"
End If

If rs(4) = "PENDIENTE" Then
	estF = "<font style='color:#FF0000;'>PENDIENTE</font>"
ElseIf rs(4) = "OK" Then
	estF = "<font style='color:#088A08;'>OK</font>"
ElseIf rs(4) = "BAJA" Then
	estF = "<font style='color:#681415;'>BAJA</font>"
End If

xform = xform & "<form id='xform" & rs(0) & "' method='post' action='detalles.asp?id=" & rs(0) & "&opc=veri&tipo=cons&" & request.form & "'></form>"
if getPermisos(9,"INICIO","CONS") = 1 then
	xboton = "<input class='bot' style='border-radius:5px;background:#DAE8FF;color:#008520;' type='button' value='D' onclick=""document.getElementById('xform" & rs(0) & "').submit();"">"
else
	xboton = "<input class='botsp' type='button' value='D' onclick=""alert('Sin permiso para consultar.');"">"
end if

html = html & "<tr>"

html = html & "<td xclass='regis' style='font-family:tahoma;font-size:9pt;xpadding:8px;text-align:center;background-color:" & sColor & ";'>" & rs(0) & "</td>"
html = html & "<td xclass='regis' style='font-family:tahoma;font-size:9pt;padding:8px;text-align:left;background-color:" & sColor & ";'>" & mid(ucase(rs(1)),1,58) & "</td>" 'UCASE(rs(1))'
'html = html & "<td xclass='regis' >" & xboton & "</td>"
html = html & "<td xclass='regis' style='font-family:tahoma;font-size:9pt;padding-left:25px;text-align:left;background-color:" & sColor & ";'>" & rs(2) & "</td>"
html = html & "<td xclass='regis' style='font-family:tahoma;font-size:9pt;xpadding-left:25px;text-align:center;background-color:" & sColor & ";'>" & estD & "</td>"
html = html & "<td class='regis' style='font-family:tahoma;font-size:9pt;xpadding-left:25px;text-align:center;background-color:" & sColor & ";'>" & estF & "</td>"
html = html & "<td class='regis' style='font-family:tahoma;font-size:9pt;xpadding-left:25px;text-align:center;background-color:" & sColor & ";'>" &  Fechador(rs(5) + nFecini) & "</td>"

html = html & "</tr>"

if sColor = "#FFFFFF" then
	sColor = "#ECECEC"
elseif sColor = "#ECECEC" then
	sColor = "#FFFFFF"
end if

rs.movenext
loop

html = html & "</table>"


end if



html = html & "<script type='text/javascript'>"
html = html & " function borrar(formu){ "
html = html & " var r = confirm('Desea eliminar incidente?'); "
html = html & " if (r){ "
html = html & " document.getElementById(formu).submit(); "
html = html & " }else{ "
html = html & " return false;"
html = html & " }} "
html = html & "</script>"



html = html & "<script src='assets/js/jquery.min.js'></script>"
html = html & "<script src='assets/js/skel.min.js'></script>"
html = html & "<script src='assets/js/util.js'></script>"
html = html & "<script src='assets/js/main.js'></script>"

html = html & formx22 & formx23 & formx4 & formx5 & xform & formx10 & formx1 & formx2 & formx3 & "</body>"
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