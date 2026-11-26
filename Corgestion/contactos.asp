<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/ui_componentes.inc" -->
<%

html = headhtml()
html = html & "<body>"

xmenu = request.querystring("xmenu")
xtmenu = request.querystring("xtmenu")
opc = request.querystring("opc")
id = request.querystring("id")
cli = request.form("cli")
if cli & "" = "" then cli = request.querystring("cli")
if cli = "" then cli = 0
' [UI] El cliente se usa en la consulta SQL: se valida como numero (antes se pegaba tal cual)
If IsNumeric(cli & "") Then cli = CLng(cli) Else cli = 0

CantRegMostrar = 10
'[UI-ANT] CantRegMover = CDbl(Request.querystring("CantRegMover"))
' [UI] Sin CantRegMover en la URL, CDbl("") daba error de tipo
CantRegMover = 0
If IsNumeric(Request.querystring("CantRegMover") & "") Then CantRegMover = CDbl(Request.querystring("CantRegMover"))
if CantRegMover < 0 then CantRegMover = 0

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

XXX = SYSLOGON(cnn,"PARAM_CONTACT")

If opc = "act" Then

	id = request.querystring("id")

	Dim cmdUpd
	Set cmdUpd = Server.CreateObject("ADODB.Command")
	With cmdUpd
		.ActiveConnection = Cnn
		.CommandText = "sp_cg_Contactos"
		.CommandType = 4
		.Parameters.Append .CreateParameter("@Modo", 200, 1, 20, "actualizar")
		.Parameters.Append .CreateParameter("@Id", 3, 1, 4, CLng(id))
		.Parameters.Append .CreateParameter("@IdCliente", 3, 1, 4, CLng(request.form("cli")))
		.Parameters.Append .CreateParameter("@Nombre", 200, 1, 100, CStr(request.form("nom")))
		.Parameters.Append .CreateParameter("@Apellido", 200, 1, 100, CStr(request.form("ape")))
		.Parameters.Append .CreateParameter("@Telefono", 200, 1, 50, CStr(request.form("tel")))
		.Parameters.Append .CreateParameter("@Mail", 200, 1, 100, CStr(request.form("mail")))
		.Parameters.Append .CreateParameter("@Tipo", 3, 1, 4, NULL )
	End With

	On Error Resume Next
	cmdUpd.Execute
	If Err.Number <> 0 Then
		html = html & "<div style='background:#fee2e2; color:#991b1b; border:1px solid #fecaca; padding:1rem; margin:1rem 2rem; border-radius:6px; font-family:monospace;'>"
		html = html & "<strong>ERROR al guardar contacto (ID " & Server.HTMLEncode(id & "") & "):</strong><br>"
		html = html & "Nro: " & Err.Number & " - " & Server.HTMLEncode(Err.Description & "") & "<br>"
		html = html & "<small>SP: sp_cg_Contactos @Modo='actualizar' @Id=" & Server.HTMLEncode(id & "") & "</small>"
		html = html & "</div>"
		html = html & "<script>alert('Error al guardar: " & Replace(Err.Description, "'", "\'") & " (" & Err.Number & ")');</script>"
		Err.Clear
	Else
		CgLog "CONTACTO", id, "ACTUALIZA"
		html = html & "<script>alert('Registro actualizado!');</script>"
	End If
	On Error Goto 0
	Set cmdUpd = Nothing

end if

'[UI-ANT] ' ===== Vista anterior completa: reemplazada por el bloque [UI] que sigue. Se borra entera buscando [UI-ANT] =====
'[UI-ANT] html = html & cabecera()

'[UI-ANT] html = html & "<div class='titulo-seccion ui-page-header' style='padding: 1rem 2rem;'>"
'[UI-ANT] html = html & "<div class='ui-toolbar' style='display:flex;align-items:center;justify-content:space-between;'>"

'[UI-ANT] html = html & "<div><h2>CONTACTOS</h2></div>"

'[UI-ANT] html = html & "<form id='xfrm' method='post' action='contactos.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
'[UI-ANT] html = html & "<div class='ui-toolbar__filters'>"

'[UI-ANT] html = html & "<select class='ui-select' id='cli' name='cli' onchange=""document.getElementById('xfrm').submit();"">"

'[UI-ANT] set rs = Cnn.Execute("sp_cg_Catalogo_Listar @Catalogo='clientes'")

'[UI-ANT] html = html & "<option value='0' >FILTRO CLIENTE</option>"

'[UI-ANT] do while not rs.eof

	'[UI-ANT] if cstr(cli) = cstr(rs(0)) then
		'[UI-ANT] html = html & "<option value='" & rs(0) & "' selected>" & Server.HTMLEncode(rs(1) & "") & "</option>"
	'[UI-ANT] else
		'[UI-ANT] html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
	'[UI-ANT] end if

'[UI-ANT] rs.movenext
'[UI-ANT] loop

'[UI-ANT] html = html & "</select>"
'[UI-ANT] html = html & "</div>" ' Fin ui-toolbar__filters
'[UI-ANT] html = html & "</form>"

'[UI-ANT] formx = "<form id='formx' class='form-oculto' method='post' action='agregar1.asp?tipo=CO&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
'[UI-ANT] if getPermisos(xmenu,xtmenu,"ALTA") = 1 then
	'[UI-ANT] botonx = "<input class='bot ui-button ui-button--primary' type='button' value='AGREGAR CONTACTO' onclick=""document.getElementById('formx').submit();"">"
'[UI-ANT] else
	'[UI-ANT] botonx = "<input class='botsp ui-button' type='button' value='AGREGAR CONTACTO' onclick=""alert('Sin permiso para agregar.');"">"
'[UI-ANT] end if

'[UI-ANT] html = html & "<div class='ui-toolbar__actions'>[BOTONERA_PAGINACION]" & botonx & "</div>"

'[UI-ANT] html = html & "</div>" ' Fin ui-toolbar
'[UI-ANT] html = html & "</div>" ' Fin titulo-seccion

'[UI-ANT] 'html1 = html1 & "<div style='height:57%;width:100%;'>"
'[UI-ANT] html1 = html1 & "<div class='inner ui-main' style='padding:2rem;'>"
'[UI-ANT] html1 = html1 & "<div class='bloque-datos table-responsive ui-table-wrap'>"

'[UI-ANT] 'html1 = html1 & "<table style='width:90%;heigth:80%;font-family:Serif;text-align:center;' cellspacing=0 cellpadding=3 align=center>"
'[UI-ANT] html1 = html1 & "<table class='tabla-moderna ui-table'>"

'[UI-ANT] html1 = html1 & "<thead>"
'[UI-ANT] 'html1 = html1 & "<tr>"
'[UI-ANT] html1 = html1 & "<tr>"

'[UI-ANT] 'html1 = html1 & "<td class='encabe' style='width:1%;' align=center>*</td>"
'[UI-ANT] html1 = html1 & "<th style='width: 40px;'>&nbsp;</th>"

'[UI-ANT] 'html1 = html1 & "<td class='encabe' style='width:1%;' align=center>ID</td>"
'[UI-ANT] html1 = html1 & "<th style='width: 5%;'>ID</th>"

'[UI-ANT] 'html1 = html1 & "<td class='encabe' style='width:18%;' align=center>NOMBRE</td>"
'[UI-ANT] html1 = html1 & "<th style='width: 18%;'>NOMBRE</th>"

'[UI-ANT] 'html1 = html1 & "<td class='encabe' style='width:18%;' align=center>APELLIDO</td>"
'[UI-ANT] html1 = html1 & "<th style='width: 18%;'>APELLIDO</th>"

'[UI-ANT] 'html1 = html1 & "<td class='encabe' style='width:18%;' align=center>TELEFONO</td>"
'[UI-ANT] html1 = html1 & "<th style='width: 15%;'>TELEFONO</th>"

'[UI-ANT] 'html1 = html1 & "<td class='encabe' style='width:22%;' align=center>MAIL</td>"
'[UI-ANT] html1 = html1 & "<th style='width: 22%;'>MAIL</th>"

'[UI-ANT] 'html1 = html1 & "<td class='encabe' style='width:17%;' align=center>CLIENTE</td>"
'[UI-ANT] html1 = html1 & "<th style='width: 17%;'>CLIENTE</th>"
'[UI-ANT] html1 = html1 & "<th style='width: 8%;'></th>"

'[UI-ANT] 'html1 = html1 & "</tr>"
'[UI-ANT] html1 = html1 & "</tr>"
'[UI-ANT] html1 = html1 & "</thead>"

'[UI-ANT] if cli <> 0 and opc <> "act" then
	'[UI-ANT] Set rs = Exec(Cnn, "sp_cg_Contactos @Modo='listar', @IdCliente=" & cli)
'[UI-ANT] else
	'[UI-ANT] Set rs = Exec(Cnn, "sp_cg_Contactos @Modo='listar'")
'[UI-ANT] end if

'[UI-ANT] formx1 = ""

'[UI-ANT] if not rs.eof then rs.Move (CantRegMover)

'[UI-ANT] PosReg = 0

'[UI-ANT] html1 = html1 & "<tbody>"

'[UI-ANT] Do While (Not rs.eof) And (PosReg < CantRegMostrar)

	'[UI-ANT] xFlg = ""
	'[UI-ANT] ' Generamos las pills dinamicas modernas con sus colores especificos por cliente
	'[UI-ANT] if rs(5) = 1 then
		'[UI-ANT] xFlg = "<span style='font-size: 0.75rem; font-weight: bold; color: #ea580c; background-color: #ffedd5; padding: 0.2rem 0.4rem; border-radius: 4px; white-space: nowrap;'>GAL</span>"
	'[UI-ANT] elseif rs(5) = 2 then
		'[UI-ANT] xFlg = "<span style='font-size: 0.75rem; font-weight: bold; color: #16a34a; background-color: #dcfce7; padding: 0.2rem 0.4rem; border-radius: 4px; white-space: nowrap;'>PAT</span>"
	'[UI-ANT] elseif rs(5) = 33 then
		'[UI-ANT] xFlg = "<span style='font-size: 0.75rem; font-weight: bold; color: #2563eb; background-color: #dbeafe; padding: 0.2rem 0.4rem; border-radius: 4px; white-space: nowrap;'>CIU</span>"
	'[UI-ANT] elseif rs(5) = 4 then
		'[UI-ANT] xFlg = "<span style='font-size: 0.75rem; font-weight: bold; color: #92400e; background-color: #fef3c7; padding: 0.2rem 0.4rem; border-radius: 4px; white-space: nowrap;'>PET</span>"
	'[UI-ANT] elseif rs(5) = 35 then
		'[UI-ANT] xFlg = "<span style='font-size: 0.75rem; font-weight: bold; color: #15803d; background-color: #dcfce7; padding: 0.2rem 0.4rem; border-radius: 4px; white-space: nowrap;'>VAL</span>"
	'[UI-ANT] elseif rs(5) = 32 then
		'[UI-ANT] xFlg = "<span style='font-size: 0.75rem; font-weight: bold; color: #854d0e; background-color: #fef08a; padding: 0.2rem 0.4rem; border-radius: 4px; white-space: nowrap;'>BPN</span>"
	'[UI-ANT] elseif rs(5) = 36 then
		'[UI-ANT] xFlg = "<span style='font-size: 0.75rem; font-weight: bold; color: #dc2626; background-color: #fee2e2; padding: 0.2rem 0.4rem; border-radius: 4px; white-space: nowrap;'>CFA</span>"
	'[UI-ANT] elseif rs(5) = 37 then
		'[UI-ANT] xFlg = "<span style='font-size: 0.75rem; font-weight: bold; color: #475569; background-color: #f1f5f9; padding: 0.2rem 0.4rem; border-radius: 4px; white-space: nowrap;'>COR</span>"
	'[UI-ANT] elseif rs(5) = 38 then
		'[UI-ANT] xFlg = "<span style='font-size: 0.75rem; font-weight: bold; color: #92400e; background-color: #fef3c7; padding: 0.2rem 0.4rem; border-radius: 4px; white-space: nowrap;'>BER</span>"
	'[UI-ANT] elseif rs(5) = 39 then
		'[UI-ANT] xFlg = "<span style='font-size: 0.75rem; font-weight: bold; color: #92400e; background-color: #fef3c7; padding: 0.2rem 0.4rem; border-radius: 4px; white-space: nowrap;'>BSJ</span>"
	'[UI-ANT] elseif rs(5) = 40 then
		'[UI-ANT] xFlg = "<span style='font-size: 0.75rem; font-weight: bold; color: #92400e; background-color: #fef3c7; padding: 0.2rem 0.4rem; border-radius: 4px; white-space: nowrap;'>BSC</span>"
	'[UI-ANT] elseif rs(5) = 41 then
		'[UI-ANT] xFlg = "<span style='font-size: 0.75rem; font-weight: bold; color: #92400e; background-color: #fef3c7; padding: 0.2rem 0.4rem; border-radius: 4px; white-space: nowrap;'>BSF</span>"
	'[UI-ANT] end if

	'[UI-ANT] formx1 = formx1 & "<form id='form" & rs(0) & "' class='form-oculto' method='post' action='eliminar.asp?opc=con&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=contactos.asp'></form>"
	
	'[UI-ANT] if getPermisos(xmenu,xtmenu,"ELIM") = 1 then
		'[UI-ANT] botonx1 = "<input class='btn-accion btn-accion-peligro ui-button ui-button--icon ui-button--danger' type='button' value='X' onclick=""borrar('form" & rs(0) & "');"">"
	'[UI-ANT] else
		'[UI-ANT] botonx1 = "<input class='btn-accion btn-accion-deshabilitado ui-button ui-button--icon' type='button' value='X' onclick=""alert('Sin permiso para eliminar.');"">"
	'[UI-ANT] end if

	'[UI-ANT] html1 = html1 & "<tr>"
	'[UI-ANT] html1 = html1 & "<td>" & botonx1 & "</td>"
	'[UI-ANT] html1 = html1 & "<td>" & rs(0) & "</td>"

	'[UI-ANT] html1 = html1 & "<form id='iform" & rs(0) & "' class='form-linea' method='post' action='contactos.asp?opc=act&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"

	'[UI-ANT] ' --- CELDA DEL NOMBRE CON LA PILL ALINEADA ---
	'[UI-ANT] html1 = html1 & "<td>"
	'[UI-ANT] if getPermisos(xmenu,xtmenu,"MODIF") = 1 then
		'[UI-ANT] ' Envolvemos la pill y el input en un flex container con alineacion centrada y una separacion (gap)
		'[UI-ANT] html1 = html1 & "<div style='display: flex; align-items: center; gap: 0.5rem;'>" & xFlg & "<input class='input-moderno ui-input' style='width: 100%;' name='nom' type='text' value='" & Server.HTMLEncode(rs(1) & "") & "' readonly></div>"
	'[UI-ANT] else
		'[UI-ANT] html1 = html1 & "<div style='display: flex; align-items: center; gap: 0.5rem;'>" & xFlg & "<input class='input-moderno input-deshabilitado ui-input' style='width: 100%;' name='nom' type='text' value='" & Server.HTMLEncode(rs(1) & "") & "' onblur=""alert('Sin permiso para modificar.');"" readonly></div>"
	'[UI-ANT] end if
	'[UI-ANT] html1 = html1 & "</td>"

	'[UI-ANT] 'html1 = html1 & "<td class='fil' >"
	'[UI-ANT] html1 = html1 & "<td>"
	'[UI-ANT] if getPermisos(xmenu,xtmenu,"MODIF") = 1 then
		'[UI-ANT] 'html1 = html1 & "<input class='inpu' style='width:90%;font-size:10pt;' id='ape' name='ape' type='text' value='" & rs(2) & "' onblur=""document.getElementById('iform" & rs(0) & "').submit();"" >"
		'[UI-ANT] html1 = html1 & "<input class='input-moderno ui-input' name='ape' type='text' value='" & Server.HTMLEncode(rs(2) & "") & "' readonly>"
	'[UI-ANT] else
		'[UI-ANT] 'html1 = html1 & "<input class='inpu' style='width:90%;font-size:10pt;' id='ape' name='ape' type='text' value='" & rs(2) & "' onblur=""alert('Sin permiso para modificar.');"" >"
		'[UI-ANT] html1 = html1 & "<input class='input-moderno input-deshabilitado ui-input' name='ape' type='text' value='" & Server.HTMLEncode(rs(2) & "") & "' onblur=""alert('Sin permiso para modificar.');"" readonly>"
	'[UI-ANT] end if
	'[UI-ANT] html1 = html1 & "</td>"

	'[UI-ANT] 'html1 = html1 & "<td class='fil' >"
	'[UI-ANT] html1 = html1 & "<td>"
	'[UI-ANT] if getPermisos(xmenu,xtmenu,"MODIF") = 1 then
		'[UI-ANT] 'html1 = html1 & "<input class='inpu' style='width:90%;font-size:10pt;' id='tel' name='tel' type='text' value='" & rs(3) & "' onblur=""document.getElementById('iform" & rs(0) & "').submit();"" >"
		'[UI-ANT] html1 = html1 & "<input class='input-moderno ui-input' name='tel' type='text' value='" & Server.HTMLEncode(rs(3) & "") & "' readonly>"
	'[UI-ANT] else
		'[UI-ANT] 'html1 = html1 & "<input class='inpu' style='width:90%;font-size:10pt;' id='tel' name='tel' type='text' value='" & rs(3) & "' onblur=""alert('Sin permiso para modificar.');"" >"
		'[UI-ANT] html1 = html1 & "<input class='input-moderno input-deshabilitado ui-input' name='tel' type='text' value='" & Server.HTMLEncode(rs(3) & "") & "' onblur=""alert('Sin permiso para modificar.');"" readonly>"
	'[UI-ANT] end if
	'[UI-ANT] html1 = html1 & "</td>"

	'[UI-ANT] 'html1 = html1 & "<td class='fil' >"
	'[UI-ANT] html1 = html1 & "<td>"
	'[UI-ANT] if getPermisos(xmenu,xtmenu,"MODIF") = 1 then
		'[UI-ANT] 'html1 = html1 & "<input class='inpu' style='width:100%;font-size:10pt;' id='mail' name='mail' type='text' value='" & rs(4) & "' onblur=""document.getElementById('iform" & rs(0) & "').submit();"" >"
		'[UI-ANT] html1 = html1 & "<input class='input-moderno ui-input' name='mail' type='text' value='" & Server.HTMLEncode(rs(4) & "") & "' readonly>"
	'[UI-ANT] else
		'[UI-ANT] 'html1 = html1 & "<input class='inpu' style='width:100%;font-size:10pt;' id='mail' name='mail' type='text' value='" & rs(4) & "' onblur=""alert('Sin permiso para modificar.');"" >"
		'[UI-ANT] html1 = html1 & "<input class='input-moderno input-deshabilitado ui-input' name='mail' type='text' value='" & Server.HTMLEncode(rs(4) & "") & "' onblur=""alert('Sin permiso para modificar.');"" readonly>"
	'[UI-ANT] end if
	'[UI-ANT] html1 = html1 & "</td>"

	'[UI-ANT] 'html1 = html1 & "<td class='fil' align=center>"
	'[UI-ANT] html1 = html1 & "<td>"
	'[UI-ANT] if getPermisos(xmenu,xtmenu,"MODIF") = 1 then
		'[UI-ANT] 'html1 = html1 & "<select class='selec' style='width:95%;font-size:10pt;' id='cli' name='cli' onchange=""document.getElementById('iform" & rs(0) & "').submit();"" >"
		'[UI-ANT] html1 = html1 & "<select class='input-moderno ui-select' name='cli' disabled>"
	'[UI-ANT] else
		'[UI-ANT] 'html1 = html1 & "<select class='selec' style='width:95%;font-size:10pt;' id='cli' name='cli' onchange=""alert('Sin permiso para modificar.');"" >"
		'[UI-ANT] html1 = html1 & "<select class='input-moderno input-deshabilitado ui-select' name='cli' onchange=""alert('Sin permiso para modificar.');"" disabled>"
	'[UI-ANT] end if

	'[UI-ANT] Set rx = Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='clientes'")

	'[UI-ANT] Do While Not rx.eof

		'[UI-ANT] If rx(0) = rs(5) Then
			'[UI-ANT] html1 = html1 & "<option value='" & rx(0) & "' selected>" & Server.HTMLEncode(rx(1) & "") & "</option>"
		'[UI-ANT] Else
			'[UI-ANT] html1 = html1 & "<option value='" & rx(0) & "'>" & Server.HTMLEncode(rx(1) & "") & "</option>"
		'[UI-ANT] End If

		'[UI-ANT] rx.movenext
	'[UI-ANT] Loop

	'[UI-ANT] html1 = html1 & "</select>"
	'[UI-ANT] html1 = html1 & "</td>"
	'[UI-ANT] html1 = html1 & "<td>" & HtmlBotonEditarFila("iform" & rs(0), getPermisos(xmenu,xtmenu,"MODIF") = 1) & "</td>"

	'[UI-ANT] html1 = html1 & "</form>"

	'[UI-ANT] 'html1 = html1 & "</tr>"
	'[UI-ANT] html1 = html1 & "</tr>"

	'[UI-ANT] PosReg = PosReg + 1

	'[UI-ANT] rs.movenext
'[UI-ANT] Loop

'[UI-ANT] html1 = html1 & "</tbody>"

'[UI-ANT] 'html1 = html1 & "</table>"
'[UI-ANT] html1 = html1 & "</table>"

'[UI-ANT] 'html1 = html1 & "</div>"
'[UI-ANT] html1 = html1 & "</div>" ' Cierra .bloque-datos table-responsive
'[UI-ANT] 'html1 = html1 & "</div>"
'[UI-ANT] html1 = html1 & "</div>" ' Cierra .inner


'[UI-ANT] 'xdiv =  "<td><div id=btn xstyle='top:85px;left:750px;position:absolute;'><table align=right><tr>"
'[UI-ANT] ' Transformamos la marana de tablas en un simple contenedor con clase para la paginacion
'[UI-ANT] misBotones = ""

'[UI-ANT] if cantregmover <> 0 then
	'[UI-ANT] formx22 = "<form id='formx22' class='form-oculto' method='post' action='contactos.asp?cantregmover=" & (cantregmover - cantregmostrar) & "&cli=" & cli & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
	'[UI-ANT] 'botonx22 = "<input class='bot' type='button' value='ATRAS' onclick=""document.getElementById('formx22').submit();"">"
	'[UI-ANT] botonx22 = "<input class='bot ui-button' type='button' value='ATRAS' onclick=""document.getElementById('formx22').submit();"">"
'[UI-ANT] else
	'[UI-ANT] formx22 = ""
	'[UI-ANT] 'botonx22 = "<input style='border-radius:7px;' type='button' value='ATRAS'>"
	'[UI-ANT] botonx22 = "<input class='botsp ui-button' type='button' value='ATRAS' disabled>"
'[UI-ANT] end if

'[UI-ANT] if not rs.eof then
	'[UI-ANT] formx23 = "<form id='formx23' class='form-oculto' method='post' action='contactos.asp?cantregmover=" & (cantregmover + cantregmostrar) & "&cli=" & cli & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
	'[UI-ANT] 'botonx23 = "<input class='bot' type='button' value='SIGUIENTE' onclick=""document.getElementById('formx23').submit();"">"
	'[UI-ANT] botonx23 = "<input class='bot ui-button' type='button' value='SIGUIENTE' onclick=""document.getElementById('formx23').submit();"">"
'[UI-ANT] else
	'[UI-ANT] formx23 = ""
	'[UI-ANT] 'botonx23 = "<input style='border-radius:7px;' type='button' value='SIGUIENTE'>"
	'[UI-ANT] botonx23 = "<input class='botsp ui-button' type='button' value='SIGUIENTE' disabled>"
'[UI-ANT] end if

'[UI-ANT] 'xdiv = xdiv & "<td>" & botonx22 & "</td>"
'[UI-ANT] 'xdiv = xdiv & "<td>" & botonx23 & "</td>"
'[UI-ANT] 'xdiv = xdiv & "</tr></table></div></td>"
'[UI-ANT] misBotones = "<div class='ui-pagination' style='display: flex; gap: 0.5rem;'>" & formx22 & botonx22 & formx23 & botonx23 & "</div>"

'[UI-ANT] html = Replace(html, "[BOTONERA_PAGINACION]", misBotones)

'[UI-ANT] html = html & "<script type='text/javascript'>"
'[UI-ANT] html = html & " function borrar(formu){ "
'[UI-ANT] html = html & " var r = confirm('Desea eliminar contacto?'); "
'[UI-ANT] html = html & " if (r){ "
'[UI-ANT] html = html & " document.getElementById(formu).submit(); "
'[UI-ANT] html = html & " }else{ "
'[UI-ANT] html = html & " return false;"
'[UI-ANT] html = html & " }} "
'[UI-ANT] html = html & "</script>"

'[UI-ANT] html = html & ScriptEdicionCatalogo()

'[UI-ANT] html = html & "<script src='assets/js/jquery.min.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/skel.min.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/util.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/main.js'></script>"

'[UI-ANT] html = html & xdiv & html1 & formx1 & shtml & formx22 & formx23 & formx & "</body>"
'[UI-ANT] html = html & "</html>"

' =========================================================================
' [UI] VISTA con el sistema visual (IncFile/ui_componentes.inc)
' Pantalla de parametros: etiqueta "Parametros" sobre el titulo.
' El bloque anterior completo queda arriba, comentado con '[UI-ANT].
' Consulta, filtro, paginacion, formularios de edicion y permisos: los mismos de antes.
' =========================================================================
Dim accTop, filtrosHtml, selHtml, puedeModif, sinPermiso, pagAnt, pagSig, siglaCli

html = html & "<div id='wrapper' role='banner'>" & cabecera() & "</div>"

' --- Barra superior: Agregar ---
formx = "<form id='formx' class='form-oculto' method='post' action='agregar1.asp?tipo=CO&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
accTop = UiBoton("Agregar contacto", "light", "document.getElementById('formx').submit();", getPermisosAtrm("CG_CONTACTOS","ALTA") = 1, "Sin permiso para agregar.")

' --- Filtro de cliente en el encabezado de la tarjeta (se envia solo al cambiar, como antes) ---
filtrosHtml = "<form id='xfrm' class='ui-card__filtros' method='post' action='contactos.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
selHtml = "<select class='ui-select' id='cli' name='cli' onchange=""document.getElementById('xfrm').submit();""><option value='0'>Todos los clientes</option>"
set rs = Cnn.Execute("sp_cg_Catalogo_Listar @Catalogo='clientes'")
do while not rs.eof
	if cstr(cli) = cstr(rs(0)) then
		selHtml = selHtml & "<option value='" & rs(0) & "' selected>" & Server.HTMLEncode(rs(1) & "") & "</option>"
	else
		selHtml = selHtml & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
	end if
	rs.movenext
loop
filtrosHtml = filtrosHtml & UiFiltroCompacto("cli", "Filtrar por cliente", selHtml & "</select>") & "</form>"

html = html & UiPantallaParametrosAbre("Contactos", accTop, "wide")
html = html & UiCardAbre("Listado de contactos", filtrosHtml)

html = html & UiTablaAbre("Contactos de los clientes")
html = html & "<thead><tr>" & UiTh("<span class='ui-sr-only'>Acciones</span>", "ui-table__col--action") & UiTh("ID", "") & _
	UiTh("Nombre", "") & UiTh("Apellido", "") & UiTh("Tel&eacute;fono", "") & UiTh("Mail", "") & UiTh("Cliente", "") & _
	UiTh("<span class='ui-sr-only'>Editar</span>", "ui-table__col--action") & "</tr></thead><tbody>"

if cli <> 0 and opc <> "act" then
	Set rs = Exec(Cnn, "sp_cg_Contactos @Modo='listar', @IdCliente=" & cli)
else
	Set rs = Exec(Cnn, "sp_cg_Contactos @Modo='listar'")
end if

formx1 = ""
if not rs.eof then rs.Move (CantRegMover)
PosReg = 0
puedeModif = (getPermisosAtrm("CG_CONTACTOS","MODIF") = 1)
' Sin permiso: los campos avisan al salir de ellos, como antes
sinPermiso = ""
If Not puedeModif Then sinPermiso = " onblur=""alert('Sin permiso para modificar.');"""

Do While (Not rs.eof) And (PosReg < CantRegMostrar)

	' Sigla del cliente junto al nombre (los mismos clientes de antes)
	siglaCli = ""
	Select Case CStr(rs(5) & "")
		Case "1"
			siglaCli = "GAL"
		Case "2"
			siglaCli = "PAT"
		Case "33"
			siglaCli = "CIU"
		Case "4"
			siglaCli = "PET"
		Case "35"
			siglaCli = "VAL"
		Case "32"
			siglaCli = "BPN"
		Case "36"
			siglaCli = "CFA"
		Case "37"
			siglaCli = "COR"
		Case "38"
			siglaCli = "BER"
		Case "39"
			siglaCli = "BSJ"
		Case "40"
			siglaCli = "BSC"
		Case "41"
			siglaCli = "BSF"
	End Select
	xFlg = ""
	If siglaCli <> "" Then xFlg = "<span class='ui-tag-id'>" & siglaCli & "</span>"

	formx1 = formx1 & "<form id='form" & rs(0) & "' class='form-oculto' method='post' action='eliminar.asp?opc=con&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=contactos.asp'></form>"

	html = html & "<tr>"
	html = html & "<td>" & UiBotonIcono("&times;", "Eliminar contacto " & rs(0), "danger", "borrar('form" & rs(0) & "');", getPermisosAtrm("CG_CONTACTOS","ELIM") = 1, "Sin permiso para eliminar") & "</td>"
	html = html & "<td>" & rs(0) & "</td>"

	' Formulario de edicion de la fila (misma estructura de antes: lo usan HtmlBotonEditarFila y ScriptEdicionCatalogo)
	html = html & "<form id='iform" & rs(0) & "' class='form-linea' method='post' action='contactos.asp?opc=act&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
	html = html & "<td><div class='ui-catalogo'>" & xFlg & "<input class='ui-input' name='nom' type='text' value='" & Server.HTMLEncode(rs(1) & "") & "' aria-label='Nombre'" & sinPermiso & " readonly></div></td>"
	html = html & "<td><input class='ui-input' name='ape' type='text' value='" & Server.HTMLEncode(rs(2) & "") & "' aria-label='Apellido'" & sinPermiso & " readonly></td>"
	html = html & "<td><input class='ui-input' name='tel' type='text' value='" & Server.HTMLEncode(rs(3) & "") & "' aria-label='Tel&eacute;fono'" & sinPermiso & " readonly></td>"
	html = html & "<td><input class='ui-input' name='mail' type='text' value='" & Server.HTMLEncode(rs(4) & "") & "' aria-label='Mail'" & sinPermiso & " readonly></td>"

	html = html & "<td><select class='ui-select' name='cli' aria-label='Cliente' disabled>"
	Set rx = Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='clientes'")
	Do While Not rx.eof
		If rx(0) = rs(5) Then
			html = html & "<option value='" & rx(0) & "' selected>" & Server.HTMLEncode(rx(1) & "") & "</option>"
		Else
			html = html & "<option value='" & rx(0) & "'>" & Server.HTMLEncode(rx(1) & "") & "</option>"
		End If
		rx.movenext
	Loop
	html = html & "</select></td>"

	html = html & "<td>" & HtmlBotonEditarFila("iform" & rs(0), puedeModif) & "</td>"
	html = html & "</form>"
	html = html & "</tr>"

	PosReg = PosReg + 1
	rs.movenext
Loop

If PosReg = 0 Then html = html & UiFilaVacia(8, "No hay contactos para este filtro.")
html = html & UiTablaCierra()

' --- Paginacion al pie (mismos formularios de antes) ---
if cantregmover <> 0 then
	formx22 = "<form id='formx22' class='form-oculto' method='post' action='contactos.asp?cantregmover=" & (cantregmover - cantregmostrar) & "&cli=" & cli & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
	pagAnt = UiBoton("Anterior", "secondary", "document.getElementById('formx22').submit();", True, "")
else
	formx22 = ""
	pagAnt = UiBoton("Anterior", "secondary", "", False, "Primera p&aacute;gina")
end if
if not rs.eof then
	formx23 = "<form id='formx23' class='form-oculto' method='post' action='contactos.asp?cantregmover=" & (cantregmover + cantregmostrar) & "&cli=" & cli & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
	pagSig = UiBoton("Siguiente", "secondary", "document.getElementById('formx23').submit();", True, "")
else
	formx23 = ""
	pagSig = UiBoton("Siguiente", "secondary", "", False, "&Uacute;ltima p&aacute;gina")
end if
html = html & UiPaginacion(CantRegMover + 1, CantRegMover + PosReg, pagAnt, pagSig)

html = html & UiCardCierra()
html = html & UiPantallaCierra()

html = html & "<script type='text/javascript'>"
html = html & " function borrar(formu){ "
html = html & " var r = confirm('\u00bfDesea eliminar el contacto?'); "
html = html & " if (r){ "
html = html & " document.getElementById(formu).submit(); "
html = html & " }else{ "
html = html & " return false;"
html = html & " }} "
html = html & "</script>"

html = html & ScriptEdicionCatalogo()

html = html & "<script src='assets/js/jquery.min.js'></script>"
html = html & "<script src='assets/js/skel.min.js'></script>"
html = html & "<script src='assets/js/util.js'></script>"
html = html & "<script src='assets/js/main.js'></script>"

html = html & formx1 & formx22 & formx23 & formx & "</body>"
html = html & "</html>"

response.write html
%>