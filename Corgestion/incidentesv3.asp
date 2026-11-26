	<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<%

xhtml = headhtml()
 
xhtml = xhtml & "<body>"

Set Cnn = Server.CreateObject("ADODB.Connection")
On Error Resume Next
Cnn.open scon1
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

' --- MANEJO DE FILTROS --- modif
' Cliente
cliente = request.form("cli")
if cliente & "" = "" then cliente = request.querystring("cli")
if cliente & "" = "" then cliente = 0

' Titulo (Texto)
tit = request.form("tit")
if tit & "" = "" then tit = request.querystring("tit")
if tit & "" = "" then tit = ""

' Aplicativo
aplicativo = request.form("apli")
if aplicativo & "" = "" then aplicativo = request.querystring("apli")
if aplicativo & "" = "" then aplicativo = 0

' NUEVO: ID
filtro_id = request.form("filtro_id")
if filtro_id & "" = "" then filtro_id = request.querystring("filtro_id")

' NUEVO: Estado
filtro_estado = request.form("filtro_estado")
if filtro_estado & "" = "" then filtro_estado = request.querystring("filtro_estado")
if filtro_estado & "" = "" then filtro_estado = 0

' NUEVO: Referente
filtro_referente = request.form("filtro_referente")
if filtro_referente & "" = "" then filtro_referente = request.querystring("filtro_referente")
if filtro_referente & "" = "" then filtro_referente = 0
' --- FIN MANEJO DE FILTROS ---

' String para propagar filtros en todos los links
Dim sFilterParams
sFilterParams = "&xmenu=" & xmenu & "&xtmenu=" & xtmenu
sFilterParams = sFilterParams & "&cli=" & cliente
sFilterParams = sFilterParams & "&tit=" & tit
sFilterParams = sFilterParams & "&apli=" & aplicativo
sFilterParams = sFilterParams & "&filtro_id=" & filtro_id
sFilterParams = sFilterParams & "&filtro_estado=" & filtro_estado
sFilterParams = sFilterParams & "&filtro_referente=" & filtro_referente 'modificado 13/11/2025 CG


ord = request.querystring("ord")
orden = ""

select case ord

	case "titulo"

		 orden = "order by a.titulo"

	case "cliente"

		 orden = "order by a.id_cliente"

	case "aplicativos"

		orden = "order by a.id_aplic"
end select

CantRegMostrar = 10
CantRegMover = CDbl(Request.querystring("CantRegMover"))
if CantRegMover < 0 then CantRegMover = 0


html = html & "<div>"

html = html & "<header id='header' style='height:11%'>"

html = html & "<div>"

html = html & "<table style='width:100%;'><tr>"

html = html & "<td style='width:75%;' valign=center><a href='index.asp' class='logo' >"
html = html & "<img src='images/logocorP4.gif' width='10%' >"
html = html & "<font style='margin-left:100px;font-size:30pt;align:center;valign:center;font-family:Serif;color:#000000;'>Cor Consulting</font>"
html = html & "</a></td>"

html = html & "</tr></table>"

html = html & "</div>"

html = html & "</header>"

response.write html
''''''''''''''''''''''''''
html = ""
	html = html & "<div style='height:77%;width:100%;'><br>"

	html = html & "<table class='tablX' style='width:99%;color:#000000;font-family:Serif;text-align:center;' cellspacing=0 align=center>"

	html = html & "<tr>"

	html = html & "<td class='encabe' style='width:1%;'>*</td>"
	html = html & "<td class='encabe' style='width:1%;'>*</td>"
	html = html & "<td class='encabe' style='width:1%;'>*</td>"
	html = html & "<td class='encabe' style='width:1%;'>ID</td>"
	'html = html & "<td class='encabe' style='width:28%;'><a class='red' href='incidentes.asp?ord=titulo&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>TITULO</a></td>"
	'html = html & "<td class='encabe' style='width:15%;'><a class='red' href='incidentes.asp?ord=cliente&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>CLIENTE</a></td>"
	html = html & "<td class='encabe' style='width:28%;'><a class='red' href='incidentes.asp?ord=titulo" & sFilterParams & "'>TITULO</a></td>"
    html = html & "<td class='encabe' style='width:15%;'><a class='red' href='incidentes.asp?ord=cliente" & sFilterParams & "'>CLIENTE</a></td>"
    html = html & "<td class='encabe' style='width:15%;'>PRIORIDAD</td>"
	html = html & "<td class='encabe' style='width:5%;'>APLICATIVO</td>"
	html = html & "<td class='encabe' style='width:15%;'>DESARROLLO</td>"
	html = html & "<td class='encabe' style='width:15%;'>FUNCIONAL</td>"
	html = html & "<td class='encabe' style='width:12%;'>FEC INGRESO</td>"

	html = html & "</tr>"

	'LISTADO CENTRALIZADO EN sp_cg_Incidentes_Listar (sp_cg_ConsultarEntidad incidente/lista)
	fid = "NULL"
	if filtro_id & "" <> "" and IsNumeric(filtro_id) then fid = CLng(filtro_id)
	fres = "NULL"
	if filtro_referente & "" <> "" and CInt(filtro_referente) <> 0 then fres = CInt(filtro_referente)
	Set rs = Exec(Cnn, "sp_cg_Incidentes_Listar @IdCliente=" & cliente & ", @IdAplicativo=" & aplicativo & ", @Titulo='" & tit & "', @Orden='" & ord & "', @Id=" & fid & ", @Estado=" & filtro_estado & ", @Respon=" & fres)

	if not rs.eof then rs.Move (CantRegMover)

	PosReg = 0

	sColor = "#FFFFFF"
	formx1 = ""
	formx2 = ""
	formx3 = ""

	Do While (Not rs.eof) And (PosReg < CantRegMostrar)

		if rs(1) & "" = "GRUPO PETERSEN" or rs(1) & "" = "BANCO DE ENTRE RIOS" or rs(1) & "" = "BANCO DE SAN JUAN" or rs(1) & "" = "BANCO DE SANTA CRUZ" or rs(1) & "" = "BANCO DE SANTA FE" then
			sColor = "#F3E5D2"
		end if

		es1 = ""
		est2 = ""
		prio = ""

		If rs(3) = "PENDIENTE" Then
			est1 = "<font style='color:#FF0000;'>PEND</font>(" & rs(9) & ")"
		ElseIf rs(3) = "OK" Then
			est1 = "<font style='color:#088A08;'>OK</font>(" & rs(9) & ")"
		ElseIf rs(3) = "BAJA" Then
			est1 = "<font style='color:#681415;'>BAJA</font>(" & rs(9) & ")"
		ElseIf rs(3) = "TESTING EN CLIENTE" Then
			est1 = "<font style='color:#FF0000;'>TEST</font>(" & rs(9) & ")"
		End If

		If rs(4) = "PENDIENTE" Then
			est2 = "<font style='color:#FF0000;'>PEND</font>(" & rs(8) & ")"
		ElseIf rs(4) = "OK" Then
			est2 = "<font style='color:#088A08;'>OK</font>(" & rs(8) & ")"
		ElseIf rs(4) = "BAJA" Then
			est2 = "<font style='color:#681415;'>BAJA</font>(" & rs(8) & ")"
		ElseIf rs(4) = "TESTING EN CLIENTE" Then
			est2 = "<font style='color:#FF0000;'>TEST</font>(" & rs(8) & ")"
		End If

		if rs(2) = "ALTA" then
			prio = "ALTA <font style='color:#FF0000;'><b>( ! )</b></font>"
		else
			prio = rs(2)
		end if

		'BOTON ELIMINAR
		formx1 = formx1 & "<form id='formx" & rs(0) & "' style='height:1px;' method='post' action='eliminar.asp?opc=inc&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=incidentes.asp'></form>"
		if getPermisos(xmenu,xtmenu,"ELIM") = 1 then
			botonx1 = "<input class='bot' style='background:#DAE8FF;color:#FF0000;' type='button' value='X' onclick=""borrar('formx" & rs(0) & "');"">"
		else
			botonx1 = "<input class='botsp' type='button' value='X' onclick=""alert('Sin permiso para eliminar.');"">"
		end if

		'BOTON ACTUALIZAR
		formx2 = formx2 & "<form id='formx2" & rs(0) & "' style='height:1px;' method='post' action='detalles.asp?id=" & rs(0) & "&opc=acti&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
		'if getPermisos(xmenu,xtmenu,"MODIF") = 1 then
			botonx2 = "<input class='bot' style='background:#DAE8FF;color:#0036FF;' type='button' value='A' onclick=""document.getElementById('formx2" & rs(0) & "').submit();"">"
		'else
			'botonx2 = "<input class='botsp' type='button' value='A' onclick=""alert('Sin permiso para modificar.');"">"
		'end if

		'BOTON DETALLE
		'formx3 = formx3 & "<form id='formx3" & rs(0) & "' style='height:1px;' method='post' action='detalles.asp?id=" & rs(0) & "&opc=veri&cli=" & cliente & "&tit=" & tit & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
		'if getPermisos(xmenu,xtmenu,"CONS") = 1 then
		'	botonx3 = "<input class='bot' style='background:#DAE8FF;color:#008520;' type='button' value='D' onclick=""document.getElementById('formx3" & rs(0) & "').submit();"">"
		'else
		'	botonx3 = "<input class='botsp' type='button' value='D' onclick=""alert('Sin permiso para consultar.');"">"
		'end if

		html = html & "<tr>"
		html = html & "<td class='regis'>" & botonx1 & "</td>"

		'if session("USUARIO") = rs(6) or session("USUARIO") = rs(7) then ' comento para poder visualizar boton actualizar CG
			html = html & "<td class='regis' >" & botonx2 & "</td>"
		'else
			html = html & "<td class='regis' ></td>"
		'end if

		'html = html & "<td class='regis' >" & botonx3 & "</td>"
		html = html & "<td class='regis' onclick=""document.location='detalles.asp?id=" & rs(0) & "&opc=veri" & sFilterParams & "'"">" & rs(0) & "</td>"

		if rs(7) <> 0 then
			html = html & "<td class='regis' title='" & rs(11) & "' onclick=""document.location='http://server-2025/corgestionv5/detalles.asp?id=" & rs(0) & "&opc=veri&cli=" & cliente & "&tit=" & tit & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'"">" & Left(UCASE(rs(11)),40) & "</td>"
		else
			html = html & "<td class='regis' onclick=""document.location='detalles.asp?id=" & rs(0) & "&opc=veri" & sFilterParams & "'"">" & Left(UCASE(rs(11)),40) & " <font style='color:#000FFF;'><b>(C)</b></font></td>"
		end if

		html = html & "<td class='regis' onclick=""document.location='detalles.asp?id=" & rs(0) & "&opc=veri" & sFilterParams & "'"">" & rs(1) & "</td>"

		html = html & "<td class='regis' onclick=""document.location='detalles.asp?id=" & rs(0) & "&opc=veri" & sFilterParams & "'"">" & prio & "</td>"

		html = html & "<td class='regis' onclick=""document.location='detalles.asp?id=" & rs(0) & "&opc=veri" & sFilterParams & "'"">" & rs(12) & "</td>"

		if rs(5) <> 0 then
			html = html & "<td class='regis' onclick=""document.location='detalles.asp?id=" & rs(0) & "&opc=veri" & sFilterParams & "'""><font style='font-size:12pt;'>" & est1 & "</font></td>"
		else
			html = html & "<td class='regis' onclick=""document.location='detalles.asp?id=" & rs(0) & "&opc=veri" & sFilterParams & "'""><font style='font-size:12pt;'>-------------</td>"
		end if

		html = html & "<td class='regis' onclick=""document.location='detalles.asp?id=" & rs(0) & "&opc=veri" & sFilterParams & "'""><font style='font-size:12pt;'>" & est2 & "</font></td>"
		sFec = Fechador(rs(5) + nFecini)
		html = html & "<td class='regis' onclick=""document.location='detalles.asp?id=" & rs(0) & "&opc=veri" & sFilterParams & "'"">" & sFec & "</td>"



		html = html & "</tr>"

		PosReg = PosReg + 1

		'if sColor = "#FFFFFF" then
		'	sColor = "#ECECEC"
		'elseif sColor = "#ECECEC" then
		'	sColor = "#FFFFFF"
		'else
		'	sColor = "#FFFFFF"
		'end if

		rs.movenext
	Loop

	html = html & "</table>"

' Cargar variables para los botones de paginacion ANTES de construir el xhtml
' (Este código ya existe, solo nos aseguramos que se ejecute)
if cantregmover <> 0 then
    formx22 = "<form id='formx22' style='height:1px;' method='post' action='incidentes.asp?opc=listar&cantregmover=" & (cantregmover - cantregmostrar) & "&ord=" & ord & sFilterParams & "'></form>"
    botonx22 = "<input class='bot' type='button' value='ATRAS' onclick=""document.getElementById('formx22').submit();"">"
else
    botonx22 = "<input class='bot' type='button' value='ATRAS'>"
end if

if not rs.eof then
    formx23 = "<form id='formx23' style='height:1px;' method='post' action='incidentes.asp?opc=listar&cantregmover=" & (cantregmover + cantregmostrar) & "&ord=" & ord & sFilterParams & "'></form>"
    botonx23 = "<input class='bot' type='button' value='SIGUIENTE' onclick=""document.getElementById('formx23').submit();"">"
else
    botonx23 = "<input class='bot' type='button' value='SIGUIENTE'>"
end if

' Boton Volver
formx5 = "<form id='formx5' style='height:70%;' method='post' action='index.asp'></form>"
botonx5 = "<input class='bot' type='button' value='VOLVER' onclick=""document.getElementById('formx5').submit();"">"

' Boton Cargar Nuevo
if session("FUNC") = "FUNC" or session("FUNC") = "TEC" then
    formx4 = "<form id='formx4' style='height:70%;' method='post' action='agregar2.asp?tipo=INC&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
    if getPermisos(xmenu,xtmenu,"ALTA") = 1 then
        botonx4 = "<input class='bot' type='button' value='CARGAR NUEVO INCIDENTE' onclick=""document.getElementById('formx4').submit();"">"
    else
        botonx4 = "<input class='botsp' type='button' value='CARGAR NUEVO INCIDENTE' onclick=""alert('Sin permiso para agregar.');"">"
    end if
end if

' --- INICIO NUEVA BARRA DE FILTROS ---
xhtml = xhtml & "<div style='width:100%; padding: 10px 5px; border-top:solid 5px #FF3D3D; border-bottom:double 6px #001BFF;'>"

' Usamos un solo formulario para todos los filtros
xhtml = xhtml & "<form id='xForm' method='post' action='incidentes.asp?ord=" & ord & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"

' Fila 1: Titulo y Botones de acción
xhtml = xhtml & "<table style='width:100%;'><tr>"
xhtml = xhtml & "<td style='width:30%;'><font style='font-size:20pt;font-family:Serif;color:#000000;letter-spacing:0.2em;'><b>INCIDENTES</b></font></td>"
xhtml = xhtml & "<td style='width:70%; text-align:right;'>" & botonx4 & " " & botonx5 & "</td>"
xhtml = xhtml & "</tr></table>"

' Fila 2: Menú desplegable de filtros
xhtml = xhtml & "<details open style='margin-top: 10px;'>"
xhtml = xhtml & "  <summary style='font-size: 14pt; font-family: Serif; cursor: pointer; color: #001BFF;'><b>Filtros de Búsqueda</b></summary>"
xhtml = xhtml & "  <div style='padding: 10px; border: 1px solid #ccc; background: #f0f0f0; margin-top: 5px;'>"

' Fila de Filtros
xhtml = xhtml & "    <span style='margin-right: 10px; font-family:Serif;'>ID: <input type='text' name='filtro_id' value='" & filtro_id & "' style='width:60px;'></span>"
xhtml = xhtml & "    <span style='margin-right: 10px; font-family:Serif;'>Texto: <input type='text' name='tit' value='" & tit & "' style='width:150px;'></span>"
 
' Dropdown Cliente
xhtml = xhtml & "    <span style='margin-right: 10px; font-family:Serif;'>Cliente: <select class='selec' style='height:34px;' name='cli'>"
Set rh = Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='clientes'")
xhtml = xhtml & "<option value='0' >TODOS</option>"
do while not rh.eof
    if cstr(rh(0)) = cstr(cliente) then
        xhtml = xhtml & "<option value='" & rh(0) & "' selected >" & rh(1) & "</option>"
    else
        xhtml = xhtml & "<option value='" & rh(0) & "' >" & rh(1) & "</option>"
    end if
    rh.movenext
loop
xhtml = xhtml & "</select></span>"
rh.Close
Set rh = Nothing

' Dropdown Aplicativo
xhtml = xhtml & "    <span style='margin-right: 10px; font-family:Serif;'>Aplicativo: <select class='selec' style='height:34px;' name='apli'>"
Set rh = Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='aplicativos'")
xhtml = xhtml & "<option value='0' >TODOS</option>"
do while not rh.eof
    if cstr(rh(0)) = cstr(aplicativo) then
        xhtml = xhtml & "<option value='" & rh(0) & "' selected >" & rh(1) & "</option>"
    else
        xhtml = xhtml & "<option value='" & rh(0) & "' >" & rh(1) & "</option>"
    end if
    rh.movenext
loop
xhtml = xhtml & "</select></span>"
rh.Close
Set rh = Nothing

' NUEVO: Dropdown Estado
xhtml = xhtml & "    <span style='margin-right: 10px; font-family:Serif;'>Estado: <select class='selec' style='height:34px;' name='filtro_estado'>"
Set rh = Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='estados'")
xhtml = xhtml & "<option value='0' >TODOS</option>"
do while not rh.eof
    if cstr(rh(0)) = cstr(filtro_estado) then
        xhtml = xhtml & "<option value='" & rh(0) & "' selected >" & rh(1) & "</option>"
    else
        xhtml = xhtml & "<option value='" & rh(0) & "' >" & rh(1) & "</option>"
    end if
    rh.movenext
loop
xhtml = xhtml & "</select></span>"
rh.Close
Set rh = Nothing

' NUEVO: Dropdown Referente
xhtml = xhtml & "    <span style='margin-right: 10px; font-family:Serif;'>Referente: <select class='selec' style='height:34px;' name='filtro_referente'>"
Set rh = Exec(Cnn, "sp_cg_Responsables @Modo='combo_todos'")
xhtml = xhtml & "<option value='0' >TODOS</option>"
do while not rh.eof
    if cstr(rh(0)) = cstr(filtro_referente) then
        xhtml = xhtml & "<option value='" & rh(0) & "' selected >" & rh(1) & "</option>"
    else
        xhtml = xhtml & "<option value='" & rh(0) & "' >" & rh(1) & "</option>"
    end if
    rh.movenext
loop
xhtml = xhtml & "</select></span>"
rh.Close
Set rh = Nothing

' Boton de Búsqueda
xhtml = xhtml & "    <input class='bot' type='submit' value='Buscar'>"
xhtml = xhtml & "  </div>"
xhtml = xhtml & "</details>"

' Fila 3: Paginación
xhtml = xhtml & "<div style='text-align:right; margin-top: 10px;'>" & botonx22 & " " & botonx23 & "</div>"

xhtml = xhtml & "</form>" ' Cierre del formulario
xhtml = xhtml & "</div>"
' --- FIN NUEVA BARRA DE FILTROS ---

html = html & "</header>"

html = html & "</div>"

html = html & "<div style='height:8%;width:100%;border-top:solid 5px #FF3D3D;border-bottom:double 6px #001BFF;'>"

html = html & "<table style='height:100%;width:100%;'>"

html = html & "<tr>"

html = html & "<td><font style='margin-left:5%;font-size:16pt;font-family:Serif;color:#000000;letter-spacing:0.2em;'><b>CONSULTA INCIDENTES</b></font></td>"

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

'CONSULTA CENTRALIZADA EN sp_cg_Incidentes_Consultar (sp_cg_ConsultarEntidad incidente/consulta)
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

if fecha & "" <> "" then
	fecha = clng(cdate(fecha)) - nFecini
else
	fecha = "NULL"
end if

Set rs = Exec(Cnn, "sp_cg_Incidentes_Consultar @IdCliente=" & cliente & ", @Titulo='" & descri & "', @Estado=" & estado & ", @IdAplicativo=" & aplic & ", @Fecha=" & fecha)

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
html = html & "<td xclass='regis' >" & xboton & "</td>"
html = html & "<td xclass='regis' style='font-family:tahoma;font-size:9pt;padding-left:25px;text-align:left;background-color:" & sColor & ";'>" & rs(2) & "</td>"
html = html & "<td xclass='regis' style='font-family:tahoma;font-size:9pt;xpadding-left:25px;text-align:center;background-color:" & sColor & ";'>" & estD & "</td>"
html = html & "<td class='regis' style='font-family:tahoma;font-size:9pt;xpadding-left:25px;text-align:center;background-color:" & sColor & ";'>" & estF & "</td>"
html = html & "<td class='regis' style='font-family:tahoma;font-size:9pt;xpadding-left:25px;text-align:center;background-color:" & sColor & ";'>" &  Fechador(rs(5) + nFecini) & "</td>"
'html = html & "<td class='regis' style='font-family:tahoma;font-size:9pt;xpadding-left:25px;text-align:center;background-color:" & sColor & ";'>" &  rs(8) & "</td>"

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

%>