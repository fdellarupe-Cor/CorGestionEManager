<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<%

html = headhtml()
html = html & "<body>"

opc = request.querystring("opc")
id = request.querystring("id")
xmenu = request.querystring("xmenu")
xtmenu = request.querystring("xtmenu")
'response.write request.form
Set Cnn = Server.CreateObject("ADODB.Connection")
Cnn.open scon1

XXX = SYSLOGON(cnn,"PARAM_CORRESP")

If opc = "act" Then

	id = request.querystring("id")

	Dim cmdUpd
	Set cmdUpd = Server.CreateObject("ADODB.Command")
	With cmdUpd
		.ActiveConnection = Cnn
		.CommandText = "sp_cg_Responsables"
		.CommandType = 4
		.Parameters.Append .CreateParameter("@Modo", 200, 1, 20, "actualizar") 						'@Modo      VARCHAR(25),
		.Parameters.Append .CreateParameter("@Id", 3, 1, 4, CLng(id))								'@Id        INT          = NULL,
		.Parameters.Append .CreateParameter("@Funcion", 200, 1, 10, CStr(request.form("fun")))		'@Funcion   VARCHAR(50)  = NULL,
		.Parameters.Append .CreateParameter("@IdSeg", 200, 1, 100, CStr(request.form("ID_SEG")))	'@IdSeg     NVARCHAR(100)= NULL,
		.Parameters.Append .CreateParameter("@ExcluirId", 200, 1, 100, Null)							'@ExcluirId INT          = NULL,
		.Parameters.Append .CreateParameter("@Ids", 200, 1, 100, Null)								'@Ids       NVARCHAR(2000) = NULL,
		.Parameters.Append .CreateParameter("@Nombre", 200, 1, 100, CStr(request.form("nom")))		'@Nombre    NVARCHAR(100)= NULL,
		.Parameters.Append .CreateParameter("@Apellido", 200, 1, 100, CStr(request.form("ape")))	'@Apellido  NVARCHAR(100)= NULL,
		.Parameters.Append .CreateParameter("@Tel", 200, 1, 50, CStr(request.form("tel")))			'@Tel       NVARCHAR(100)= NULL,
		.Parameters.Append .CreateParameter("@Mail", 200, 1, 100, CStr(request.form("mail")))		'@Mail      NVARCHAR(300)= NULL,
		.Parameters.Append .CreateParameter("@Activo", 200, 1, 10, CStr(request.form("activo")))	'@Activo    NVARCHAR(10) = NULL

		'.Execute
	End With
	
	 On Error Resume Next
    cmdUpd.Execute()
    If Err.Number <> 0 Then
        Response.Write "ERROR " & Err.Number & ": " & Err.Description
        Response.End
    End If
    On Error Goto 0
 
	
	Set cmdUpd = Nothing

	html = html & "<script>"
	html = html & "alert('Registro actualizado !');"
	html = html & "</script>"

End If

html = html & cabecera()

html = html & "<div class='titulo-seccion ui-page-header' style='padding: 1rem 2rem;'>"
html = html & "<div class='ui-toolbar' style='display:flex;align-items:center;justify-content:space-between;'>"

html = html & "<div><h2>COR RESPONSABLES</h2></div>"

formx = "<form id='formx' class='form-oculto' method='post' action='agregar1.asp?tipo=CR&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
if getPermisos(xmenu,xtmenu,"ALTA") = 1 then
	botonx = "<input class='bot ui-button ui-button--primary' type='button' value='AGREGAR RESPONSABLE' onclick=""document.getElementById('formx').submit();"">"
else
	botonx = "<input class='botsp ui-button' type='button' value='AGREGAR RESPONSABLE' onclick=""alert('Sin permiso para agregar.');"">"
end if
html = html & "<div class='ui-toolbar__actions'>" & botonx & "</div>"

html = html & "</div>" ' Fin ui-toolbar
html = html & "</div>" ' Fin titulo-seccion

'html = html & "<div style='height:70%;width:100%;'>"
html = html & "<div class='inner ui-main' style='padding:2rem;'>"
html = html & "<div class='bloque-datos table-responsive ui-table-wrap'>"

'html = html & "<table style='width:64%;heigth:80%;font-family:Serif;text-align:center;' cellspacing=0 cellpadding=3 align=center>"
html = html & "<table class='tabla-moderna ui-table'>"

html = html & "<thead>"
'html = html & "<tr>"
html = html & "<tr>"

'html = html & "<td class='encabe' style='width:1%;' align=center>*</td>"
html = html & "<th style='width: 40px;'>&nbsp;</th>"

'html = html & "<td class='encabe' style='width:1%;' align=center>ID</td>"
html = html & "<th style='width: 5%;'>ID</th>"

'html = html & "<td class='encabe' style='width:14%;' align=center>NOMBRE</td>"
html = html & "<th style='width: 14%;'>NOMBRE</th>"

'html = html & "<td class='encabe' style='width:10%;' align=center>APELLIDO</td>"
html = html & "<th style='width: 10%;'>APELLIDO</th>"

'html = html & "<td class='encabe' style='width:10%;' align=center>TELEFONO</td>"
html = html & "<th style='width: 10%;'>TELEFONO</th>"

'html = html & "<td class='encabe' style='width:18%;' align=center>MAIL</td>"
html = html & "<th style='width: 18%;'>MAIL</th>"

'html = html & "<td class='encabe' style='width:10%;' align=center>FUNCION</td>"
html = html & "<th style='width: 10%;'>FUNCION</th>"

'html = html & "<td class='encabe' style='width:10%;' align=center>ACTIVO</td>"
html = html & "<th style='width: 8%;'>ACTIVO</th>"

'html = html & "<td class='encabe' style='width:10%;' align=center>LOGUIN</td>"
html = html & "<th style='width: 15%;'>LOGUIN</th>"

'html = html & "</tr>"
html = html & "</tr>"
html = html & "</thead>"

Set rs = Cnn.Execute("sp_cg_Responsables @Modo='listar'")

formx1 = ""

html = html & "<tbody>"

Do While Not rs.eof

	xFlg = ""
	' Reemplazamos la etiqueta font y los espacios fijos por nuestras pills dinámicas
	if rs(3) = "TEC" then
		'xFlg = "<font style='font-size:8pt;color:#FF0000;'>TEC&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>"
		xFlg = "<span style='font-size: 0.75rem; font-weight: bold; color: #dc2626; background-color: #fee2e2; padding: 0.2rem 0.4rem; border-radius: 4px; white-space: nowrap;'>TEC</span>"
	elseif rs(3) = "FUNC" then
		'xFlg = "<font style='font-size:8pt;color:#004DFF;'>FUNC&nbsp;&nbsp;&nbsp;&nbsp;</font>"
		xFlg = "<span style='font-size: 0.75rem; font-weight: bold; color: #2563eb; background-color: #dbeafe; padding: 0.2rem 0.4rem; border-radius: 4px; white-space: nowrap;'>FUNC</span>"
	end if

	'formx1 = formx1 & "<form id='form" & rs(0) & "' style='height:1px;' method='post' action='eliminar.asp?opc=corr&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=corresp.asp'></form>"
	formx1 = formx1 & "<form id='form" & rs(0) & "' class='form-oculto' method='post' action='eliminar.asp?opc=corr&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=corresp.asp'></form>"
	
	if getPermisos(xmenu,xtmenu,"ELIM") = 1 then
		'botonx1 = "<input class='bot' type='button' value='X' onclick=""borrar('form" & rs(0) & "');"">"
		botonx1 = "<input class='btn-accion btn-accion-peligro ui-button ui-button--icon ui-button--danger' type='button' value='X' onclick=""borrar('form" & rs(0) & "');"">"
	else
		'botonx1 = "<input class='botsp' type='button' value='X' onclick=""alert('Sin permiso para eliminar.');"">"
		botonx1 = "<input class='btn-accion btn-accion-deshabilitado ui-button ui-button--icon' type='button' value='X' onclick=""alert('Sin permiso para eliminar.');"">"
	end if

	'html = html & "<tr>"
	html = html & "<tr>"
	
	'html = html & "<td class='fil' >" & botonx1 & "</td>"
	html = html & "<td>" & botonx1 & "</td>"
	
	'html = html & "<td class='fil' >" & rs(0) & "</td>"
	html = html & "<td>" & rs(0) & "</td>"

	'html = html & "<form id='iform" & rs(0) & "' style='display:none;' method='post' action='corresp.asp?opc=act&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
	html = html & "<form id='iform" & rs(0) & "' style='display: contents;' method='post' action='corresp.asp?opc=act&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"

	' --- NOMBRE CON LA PILL ---
	'html = html & "<td class='fil' >"
	html = html & "<td>"
	if getPermisos(xmenu,xtmenu,"MODIF") = 1 then
		'html = html & xFlg & "<input class='inpu' style='width:65%;font-size:10pt;' id='nom' name='nom' type='text' value='" & rs(1) & "' readonly >"
		html = html & "<div style='display: flex; align-items: center; gap: 0.5rem;'>" & xFlg & "<input class='input-moderno ui-input' style='width: 100%;' name='nom' type='text' value='" & Server.HTMLEncode(rs(1) & "") & "' readonly></div>"
	else
		'html = html & xFlg & "<input class='inpu' style='width:65%;font-size:10pt;' id='nom' name='nom' type='text' value='" & rs(1) & "' onblur=""alert('Sin permiso para modificar.');"" >"
		html = html & "<div style='display: flex; align-items: center; gap: 0.5rem;'>" & xFlg & "<input class='input-moderno input-deshabilitado ui-input' style='width: 100%;' name='nom' type='text' value='" & Server.HTMLEncode(rs(1) & "") & "' onblur=""alert('Sin permiso para modificar.');"" readonly></div>"
	end if
	html = html & "</td>"

	' --- APELLIDO ---
	'html = html & "<td class='fil' >"
	html = html & "<td>"
	if getPermisos(xmenu,xtmenu,"MODIF") = 1 then
		'html = html & "<input class='inpu' style='width:90%;font-size:10pt;' id='ape' name='ape' type='text' value='" & rs(2) & "' readonly >"
		html = html & "<input class='input-moderno ui-input' name='ape' type='text' value='" & Server.HTMLEncode(rs(2) & "") & "' readonly >"
	else
		'html = html & "<input class='inpu' style='width:90%;font-size:10pt;' id='ape' name='ape' type='text' value='" & rs(2) & "' onblur=""alert('Sin permiso para modificar.');"" >"
		html = html & "<input class='input-moderno input-deshabilitado ui-input' name='ape' type='text' value='" & Server.HTMLEncode(rs(2) & "") & "' onblur=""alert('Sin permiso para modificar.');"" readonly>"
	end if
	html = html & "</td>"

	' --- TELEFONO ---
	'html = html & "<td class='fil' >"
	html = html & "<td>"
	if getPermisos(xmenu,xtmenu,"MODIF") = 1 then
		'html = html & "<input class='inpu' style='width:90%;font-size:10pt;' id='tel' name='tel' type='text' value='" & rs(6) & "' readonly >"
		html = html & "<input class='input-moderno ui-input' name='tel' type='text' value='" & Server.HTMLEncode(rs(6) & "") & "' readonly >"
	else
		'html = html & "<input class='inpu' style='width:90%;font-size:10pt;' id='tel' name='tel' type='text' value='" & rs(6) & "' onblur=""alert('Sin permiso para modificar.');"" >"
		html = html & "<input class='input-moderno input-deshabilitado ui-input' name='tel' type='text' value='" & Server.HTMLEncode(rs(6) & "") & "' onblur=""alert('Sin permiso para modificar.');"" readonly>"
	end if
	html = html & "</td>"

	' --- MAIL ---
	'html = html & "<td class='fil' >"
	html = html & "<td>"
	if getPermisos(xmenu,xtmenu,"MODIF") = 1 then
		'html = html & "<input class='inpu' style='width:95%;font-size:10pt;' id='mail' name='mail' type='text' value='" & rs(5) & "' readonly >"
		html = html & "<input class='input-moderno ui-input' name='mail' type='text' value='" & Server.HTMLEncode(rs(5) & "") & "' readonly >"
	else
		'html = html & "<input class='inpu' style='width:95%;font-size:10pt;' id='mail' name='mail' type='text' value='" & rs(5) & "' onblur=""alert('Sin permiso para modificar.');"" >"
		html = html & "<input class='input-moderno input-deshabilitado ui-input' name='mail' type='text' value='" & Server.HTMLEncode(rs(5) & "") & "' onblur=""alert('Sin permiso para modificar.');"" readonly>"
	end if
	html = html & "</td>"

	' --- FUNCION ---
	'html = html & "<td class='fil' style='height:100%;' >"
	'html = html & "<div>"
	html = html & "<td>"
	if getPermisos(xmenu,xtmenu,"MODIF") = 1 then
		'html = html & "<select class='selec' id='fun' name='fun' onchange=""document.getElementById('iform" & rs(0) & "').submit();"" >"
		html = html & "<select class='input-moderno ui-select' name='fun' onchange=""document.getElementById('iform" & rs(0) & "').submit();"" >"
	else
		'html = html & "<select class='selec' id='fun' name='fun' onchange=""alert('Sin permiso para modificar.');"" >"
		html = html & "<select class='input-moderno input-deshabilitado ui-select' name='fun' onchange=""alert('Sin permiso para modificar.');"" disabled>"
	end if

	If rs(3) = "FUNC" Then
		html = html & "<option value='FUNC' selected >FUNCIONAL</option>"
		html = html & "<option value='TEC' >TECNICO</option>"
	ElseIf rs(3) = "TEC" Then
		html = html & "<option value='TEC' selected >TECNICO</option>"
		html = html & "<option value='FUNC' >FUNCIONAL</option>"
	End If

	html = html & "</select>"
	'html = html & "</div>"
	html = html & "</td>"


	' --- ACTIVO ---
	'html = html & "<td class='fil' style='height:100%;' >"
	'html = html & "<div>"
	html = html & "<td>"
	if getPermisos(xmenu,xtmenu,"MODIF") = 1 then
		'html = html & "<select class='selec' id='activo' name='activo' onchange=""document.getElementById('iform" & rs(0) & "').submit();"" >"
		html = html & "<select class='input-moderno ui-select' name='activo' onchange=""document.getElementById('iform" & rs(0) & "').submit();"" >"
	else
		'html = html & "<select class='selec' id='activo' name='activo' onchange=""alert('Sin permiso para modificar.');"" >"
		html = html & "<select class='input-moderno input-deshabilitado ui-select' name='activo' onchange=""alert('Sin permiso para modificar.');"" disabled>"
	end if

	sselec1=""
	if rs("activo")="S" then
		sselec1="SELECTED"
	end if

	sselec2=""
	if rs("activo")="N" then
		sselec2="SELECTED"
	end if

	html = html & "<option value='S' " & sselec1 & ">SI</option>"
	html = html & "<option value='N' " & sselec2 & ">NO</option>"

	html = html & "</select>"
	'html = html & "</div>"
	html = html & "</td>"


	' --- LOGUIN ---
	'html = html & "<td class='fil' >"
	html = html & "<td>"
	if getPermisos(xmenu,xtmenu,"MODIF") = 1 then
		'html = html & "<input class='inpu' style='width:110PX;font-size:10pt;' id='ID_SEG' name='ID_SEG' type='text' SIZE=50 value='" & rs("ID_SEG") & "' readonly >"
		html = html & "<input class='input-moderno ui-input' name='ID_SEG' type='text' value='" & Server.HTMLEncode(rs("ID_SEG") & "") & "' readonly >"
	else
		'html = html & "<input class='inpu' style='width:95%;font-size:10pt;' id='ID_SEG' name='ID_SEG' type='text' SIZE=50 value='" & rs("ID_SEG") & "' onblur=""alert('Sin permiso para modificar.');"" >"
		html = html & "<input class='input-moderno input-deshabilitado ui-input' name='ID_SEG' type='text' value='" & Server.HTMLEncode(rs("ID_SEG") & "") & "' onblur=""alert('Sin permiso para modificar.');"" readonly>"
	end if
	html = html & "</td>"

	html = html & "</form>"

	'html = html & "</tr>"
	html = html & "</tr>"


	rs.movenext
Loop

html = html & "</tbody>"

'html = html & "</table>"
html = html & "</table>"

'html = html & "</div>"
html = html & "</div>" ' Cierra .bloque-datos
'html = html & "</div>"
html = html & "</div>" ' Cierra .inner

html = html & "<script type='text/javascript'>"
html = html & " function borrar(formu){ "
html = html & " var r = confirm('Desea eliminar responsable?'); "
html = html & " if (r){ "
html = html & " document.getElementById(formu).submit(); "
html = html & " }else{ "
html = html & " return false;"
html = html & " }} "
html = html & "</script>"

html = html & "<script type='text/javascript'>"
html = html & " function modificar(formu){ "
html = html & " document.getElementById(formu).submit();"
html = html & " alert('Registro actualizado');} "
html = html & "</script>"

html = html & "<script src='assets/js/jquery.min.js'></script>"
html = html & "<script src='assets/js/skel.min.js'></script>"
html = html & "<script src='assets/js/util.js'></script>"
html = html & "<script src='assets/js/main.js'></script>"

html = html & formx1 & formx & "</body>"
html = html & "</html>"

response.write html
%>