<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/ui_componentes.inc" -->
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
	CgLog "RESPONSABLE", id, "ACTUALIZA"

	html = html & "<script>"
	html = html & "alert('Registro actualizado !');"
	html = html & "</script>"

End If

'[UI-ANT] ' ===== Vista anterior completa: reemplazada por el bloque [UI] que sigue. Se borra entera buscando [UI-ANT] =====
'[UI-ANT] html = html & cabecera()

'[UI-ANT] html = html & "<div class='titulo-seccion ui-page-header' style='padding: 1rem 2rem;'>"
'[UI-ANT] html = html & "<div class='ui-toolbar' style='display:flex;align-items:center;justify-content:space-between;'>"

'[UI-ANT] html = html & "<div><h2>COR RESPONSABLES</h2></div>"

'[UI-ANT] formx = "<form id='formx' class='form-oculto' method='post' action='agregar1.asp?tipo=CR&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
'[UI-ANT] if getPermisos(xmenu,xtmenu,"ALTA") = 1 then
	'[UI-ANT] botonx = "<input class='bot ui-button ui-button--primary' type='button' value='AGREGAR RESPONSABLE' onclick=""document.getElementById('formx').submit();"">"
'[UI-ANT] else
	'[UI-ANT] botonx = "<input class='botsp ui-button' type='button' value='AGREGAR RESPONSABLE' onclick=""alert('Sin permiso para agregar.');"">"
'[UI-ANT] end if
'[UI-ANT] html = html & "<div class='ui-toolbar__actions'>" & botonx & "</div>"

'[UI-ANT] html = html & "</div>" ' Fin ui-toolbar
'[UI-ANT] html = html & "</div>" ' Fin titulo-seccion

'[UI-ANT] 'html = html & "<div style='height:70%;width:100%;'>"
'[UI-ANT] html = html & "<div class='inner ui-main' style='padding:2rem;'>"
'[UI-ANT] html = html & "<div class='bloque-datos table-responsive ui-table-wrap'>"

'[UI-ANT] 'html = html & "<table style='width:64%;heigth:80%;font-family:Serif;text-align:center;' cellspacing=0 cellpadding=3 align=center>"
'[UI-ANT] html = html & "<table class='tabla-moderna ui-table'>"

'[UI-ANT] html = html & "<thead>"
'[UI-ANT] 'html = html & "<tr>"
'[UI-ANT] html = html & "<tr>"

'[UI-ANT] 'html = html & "<td class='encabe' style='width:1%;' align=center>*</td>"
'[UI-ANT] html = html & "<th style='width: 40px;'>&nbsp;</th>"

'[UI-ANT] 'html = html & "<td class='encabe' style='width:1%;' align=center>ID</td>"
'[UI-ANT] html = html & "<th style='width: 5%;'>ID</th>"

'[UI-ANT] 'html = html & "<td class='encabe' style='width:14%;' align=center>NOMBRE</td>"
'[UI-ANT] html = html & "<th style='width: 14%;'>NOMBRE</th>"

'[UI-ANT] 'html = html & "<td class='encabe' style='width:10%;' align=center>APELLIDO</td>"
'[UI-ANT] html = html & "<th style='width: 10%;'>APELLIDO</th>"

'[UI-ANT] 'html = html & "<td class='encabe' style='width:10%;' align=center>TELEFONO</td>"
'[UI-ANT] html = html & "<th style='width: 10%;'>TELEFONO</th>"

'[UI-ANT] 'html = html & "<td class='encabe' style='width:18%;' align=center>MAIL</td>"
'[UI-ANT] html = html & "<th style='width: 18%;'>MAIL</th>"

'[UI-ANT] 'html = html & "<td class='encabe' style='width:10%;' align=center>FUNCION</td>"
'[UI-ANT] html = html & "<th style='width: 10%;'>FUNCION</th>"

'[UI-ANT] 'html = html & "<td class='encabe' style='width:10%;' align=center>ACTIVO</td>"
'[UI-ANT] html = html & "<th style='width: 8%;'>ACTIVO</th>"

'[UI-ANT] 'html = html & "<td class='encabe' style='width:10%;' align=center>LOGUIN</td>"
'[UI-ANT] html = html & "<th style='width: 15%;'>LOGUIN</th>"

'[UI-ANT] 'html = html & "</tr>"
'[UI-ANT] html = html & "</tr>"
'[UI-ANT] html = html & "</thead>"

'[UI-ANT] Set rs = Cnn.Execute("sp_cg_Responsables @Modo='listar'")

'[UI-ANT] formx1 = ""

'[UI-ANT] html = html & "<tbody>"

'[UI-ANT] Do While Not rs.eof

	'[UI-ANT] xFlg = ""
	'[UI-ANT] ' Reemplazamos la etiqueta font y los espacios fijos por nuestras pills dinamicas
	'[UI-ANT] if rs(3) = "TEC" then
		'[UI-ANT] 'xFlg = "<font style='font-size:8pt;color:#FF0000;'>TEC&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>"
		'[UI-ANT] xFlg = "<span style='font-size: 0.75rem; font-weight: bold; color: #dc2626; background-color: #fee2e2; padding: 0.2rem 0.4rem; border-radius: 4px; white-space: nowrap;'>TEC</span>"
	'[UI-ANT] elseif rs(3) = "FUNC" then
		'[UI-ANT] 'xFlg = "<font style='font-size:8pt;color:#004DFF;'>FUNC&nbsp;&nbsp;&nbsp;&nbsp;</font>"
		'[UI-ANT] xFlg = "<span style='font-size: 0.75rem; font-weight: bold; color: #2563eb; background-color: #dbeafe; padding: 0.2rem 0.4rem; border-radius: 4px; white-space: nowrap;'>FUNC</span>"
	'[UI-ANT] end if

	'[UI-ANT] 'formx1 = formx1 & "<form id='form" & rs(0) & "' style='height:1px;' method='post' action='eliminar.asp?opc=corr&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=corresp.asp'></form>"
	'[UI-ANT] formx1 = formx1 & "<form id='form" & rs(0) & "' class='form-oculto' method='post' action='eliminar.asp?opc=corr&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=corresp.asp'></form>"
	
	'[UI-ANT] if getPermisos(xmenu,xtmenu,"ELIM") = 1 then
		'[UI-ANT] 'botonx1 = "<input class='bot' type='button' value='X' onclick=""borrar('form" & rs(0) & "');"">"
		'[UI-ANT] botonx1 = "<input class='btn-accion btn-accion-peligro ui-button ui-button--icon ui-button--danger' type='button' value='X' onclick=""borrar('form" & rs(0) & "');"">"
	'[UI-ANT] else
		'[UI-ANT] 'botonx1 = "<input class='botsp' type='button' value='X' onclick=""alert('Sin permiso para eliminar.');"">"
		'[UI-ANT] botonx1 = "<input class='btn-accion btn-accion-deshabilitado ui-button ui-button--icon' type='button' value='X' onclick=""alert('Sin permiso para eliminar.');"">"
	'[UI-ANT] end if

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] html = html & "<tr>"
	
	'[UI-ANT] 'html = html & "<td class='fil' >" & botonx1 & "</td>"
	'[UI-ANT] html = html & "<td>" & botonx1 & "</td>"
	
	'[UI-ANT] 'html = html & "<td class='fil' >" & rs(0) & "</td>"
	'[UI-ANT] html = html & "<td>" & rs(0) & "</td>"

	'[UI-ANT] 'html = html & "<form id='iform" & rs(0) & "' style='display:none;' method='post' action='corresp.asp?opc=act&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
	'[UI-ANT] html = html & "<form id='iform" & rs(0) & "' style='display: contents;' method='post' action='corresp.asp?opc=act&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"

	'[UI-ANT] ' --- NOMBRE CON LA PILL ---
	'[UI-ANT] 'html = html & "<td class='fil' >"
	'[UI-ANT] html = html & "<td>"
	'[UI-ANT] if getPermisos(xmenu,xtmenu,"MODIF") = 1 then
		'[UI-ANT] 'html = html & xFlg & "<input class='inpu' style='width:65%;font-size:10pt;' id='nom' name='nom' type='text' value='" & rs(1) & "' readonly >"
		'[UI-ANT] html = html & "<div style='display: flex; align-items: center; gap: 0.5rem;'>" & xFlg & "<input class='input-moderno ui-input' style='width: 100%;' name='nom' type='text' value='" & Server.HTMLEncode(rs(1) & "") & "' readonly></div>"
	'[UI-ANT] else
		'[UI-ANT] 'html = html & xFlg & "<input class='inpu' style='width:65%;font-size:10pt;' id='nom' name='nom' type='text' value='" & rs(1) & "' onblur=""alert('Sin permiso para modificar.');"" >"
		'[UI-ANT] html = html & "<div style='display: flex; align-items: center; gap: 0.5rem;'>" & xFlg & "<input class='input-moderno input-deshabilitado ui-input' style='width: 100%;' name='nom' type='text' value='" & Server.HTMLEncode(rs(1) & "") & "' onblur=""alert('Sin permiso para modificar.');"" readonly></div>"
	'[UI-ANT] end if
	'[UI-ANT] html = html & "</td>"

	'[UI-ANT] ' --- APELLIDO ---
	'[UI-ANT] 'html = html & "<td class='fil' >"
	'[UI-ANT] html = html & "<td>"
	'[UI-ANT] if getPermisos(xmenu,xtmenu,"MODIF") = 1 then
		'[UI-ANT] 'html = html & "<input class='inpu' style='width:90%;font-size:10pt;' id='ape' name='ape' type='text' value='" & rs(2) & "' readonly >"
		'[UI-ANT] html = html & "<input class='input-moderno ui-input' name='ape' type='text' value='" & Server.HTMLEncode(rs(2) & "") & "' readonly >"
	'[UI-ANT] else
		'[UI-ANT] 'html = html & "<input class='inpu' style='width:90%;font-size:10pt;' id='ape' name='ape' type='text' value='" & rs(2) & "' onblur=""alert('Sin permiso para modificar.');"" >"
		'[UI-ANT] html = html & "<input class='input-moderno input-deshabilitado ui-input' name='ape' type='text' value='" & Server.HTMLEncode(rs(2) & "") & "' onblur=""alert('Sin permiso para modificar.');"" readonly>"
	'[UI-ANT] end if
	'[UI-ANT] html = html & "</td>"

	'[UI-ANT] ' --- TELEFONO ---
	'[UI-ANT] 'html = html & "<td class='fil' >"
	'[UI-ANT] html = html & "<td>"
	'[UI-ANT] if getPermisos(xmenu,xtmenu,"MODIF") = 1 then
		'[UI-ANT] 'html = html & "<input class='inpu' style='width:90%;font-size:10pt;' id='tel' name='tel' type='text' value='" & rs(6) & "' readonly >"
		'[UI-ANT] html = html & "<input class='input-moderno ui-input' name='tel' type='text' value='" & Server.HTMLEncode(rs(6) & "") & "' readonly >"
	'[UI-ANT] else
		'[UI-ANT] 'html = html & "<input class='inpu' style='width:90%;font-size:10pt;' id='tel' name='tel' type='text' value='" & rs(6) & "' onblur=""alert('Sin permiso para modificar.');"" >"
		'[UI-ANT] html = html & "<input class='input-moderno input-deshabilitado ui-input' name='tel' type='text' value='" & Server.HTMLEncode(rs(6) & "") & "' onblur=""alert('Sin permiso para modificar.');"" readonly>"
	'[UI-ANT] end if
	'[UI-ANT] html = html & "</td>"

	'[UI-ANT] ' --- MAIL ---
	'[UI-ANT] 'html = html & "<td class='fil' >"
	'[UI-ANT] html = html & "<td>"
	'[UI-ANT] if getPermisos(xmenu,xtmenu,"MODIF") = 1 then
		'[UI-ANT] 'html = html & "<input class='inpu' style='width:95%;font-size:10pt;' id='mail' name='mail' type='text' value='" & rs(5) & "' readonly >"
		'[UI-ANT] html = html & "<input class='input-moderno ui-input' name='mail' type='text' value='" & Server.HTMLEncode(rs(5) & "") & "' readonly >"
	'[UI-ANT] else
		'[UI-ANT] 'html = html & "<input class='inpu' style='width:95%;font-size:10pt;' id='mail' name='mail' type='text' value='" & rs(5) & "' onblur=""alert('Sin permiso para modificar.');"" >"
		'[UI-ANT] html = html & "<input class='input-moderno input-deshabilitado ui-input' name='mail' type='text' value='" & Server.HTMLEncode(rs(5) & "") & "' onblur=""alert('Sin permiso para modificar.');"" readonly>"
	'[UI-ANT] end if
	'[UI-ANT] html = html & "</td>"

	'[UI-ANT] ' --- FUNCION ---
	'[UI-ANT] 'html = html & "<td class='fil' style='height:100%;' >"
	'[UI-ANT] 'html = html & "<div>"
	'[UI-ANT] html = html & "<td>"
	'[UI-ANT] if getPermisos(xmenu,xtmenu,"MODIF") = 1 then
		'[UI-ANT] 'html = html & "<select class='selec' id='fun' name='fun' onchange=""document.getElementById('iform" & rs(0) & "').submit();"" >"
		'[UI-ANT] html = html & "<select class='input-moderno ui-select' name='fun' onchange=""document.getElementById('iform" & rs(0) & "').submit();"" >"
	'[UI-ANT] else
		'[UI-ANT] 'html = html & "<select class='selec' id='fun' name='fun' onchange=""alert('Sin permiso para modificar.');"" >"
		'[UI-ANT] html = html & "<select class='input-moderno input-deshabilitado ui-select' name='fun' onchange=""alert('Sin permiso para modificar.');"" disabled>"
	'[UI-ANT] end if

	'[UI-ANT] If rs(3) = "FUNC" Then
		'[UI-ANT] html = html & "<option value='FUNC' selected >FUNCIONAL</option>"
		'[UI-ANT] html = html & "<option value='TEC' >TECNICO</option>"
	'[UI-ANT] ElseIf rs(3) = "TEC" Then
		'[UI-ANT] html = html & "<option value='TEC' selected >TECNICO</option>"
		'[UI-ANT] html = html & "<option value='FUNC' >FUNCIONAL</option>"
	'[UI-ANT] End If

	'[UI-ANT] html = html & "</select>"
	'[UI-ANT] 'html = html & "</div>"
	'[UI-ANT] html = html & "</td>"


	'[UI-ANT] ' --- ACTIVO ---
	'[UI-ANT] 'html = html & "<td class='fil' style='height:100%;' >"
	'[UI-ANT] 'html = html & "<div>"
	'[UI-ANT] html = html & "<td>"
	'[UI-ANT] if getPermisos(xmenu,xtmenu,"MODIF") = 1 then
		'[UI-ANT] 'html = html & "<select class='selec' id='activo' name='activo' onchange=""document.getElementById('iform" & rs(0) & "').submit();"" >"
		'[UI-ANT] html = html & "<select class='input-moderno ui-select' name='activo' onchange=""document.getElementById('iform" & rs(0) & "').submit();"" >"
	'[UI-ANT] else
		'[UI-ANT] 'html = html & "<select class='selec' id='activo' name='activo' onchange=""alert('Sin permiso para modificar.');"" >"
		'[UI-ANT] html = html & "<select class='input-moderno input-deshabilitado ui-select' name='activo' onchange=""alert('Sin permiso para modificar.');"" disabled>"
	'[UI-ANT] end if

	'[UI-ANT] sselec1=""
	'[UI-ANT] if rs("activo")="S" then
		'[UI-ANT] sselec1="SELECTED"
	'[UI-ANT] end if

	'[UI-ANT] sselec2=""
	'[UI-ANT] if rs("activo")="N" then
		'[UI-ANT] sselec2="SELECTED"
	'[UI-ANT] end if

	'[UI-ANT] html = html & "<option value='S' " & sselec1 & ">SI</option>"
	'[UI-ANT] html = html & "<option value='N' " & sselec2 & ">NO</option>"

	'[UI-ANT] html = html & "</select>"
	'[UI-ANT] 'html = html & "</div>"
	'[UI-ANT] html = html & "</td>"


	'[UI-ANT] ' --- LOGUIN ---
	'[UI-ANT] 'html = html & "<td class='fil' >"
	'[UI-ANT] html = html & "<td>"
	'[UI-ANT] if getPermisos(xmenu,xtmenu,"MODIF") = 1 then
		'[UI-ANT] 'html = html & "<input class='inpu' style='width:110PX;font-size:10pt;' id='ID_SEG' name='ID_SEG' type='text' SIZE=50 value='" & rs("ID_SEG") & "' readonly >"
		'[UI-ANT] html = html & "<input class='input-moderno ui-input' name='ID_SEG' type='text' value='" & Server.HTMLEncode(rs("ID_SEG") & "") & "' readonly >"
	'[UI-ANT] else
		'[UI-ANT] 'html = html & "<input class='inpu' style='width:95%;font-size:10pt;' id='ID_SEG' name='ID_SEG' type='text' SIZE=50 value='" & rs("ID_SEG") & "' onblur=""alert('Sin permiso para modificar.');"" >"
		'[UI-ANT] html = html & "<input class='input-moderno input-deshabilitado ui-input' name='ID_SEG' type='text' value='" & Server.HTMLEncode(rs("ID_SEG") & "") & "' onblur=""alert('Sin permiso para modificar.');"" readonly>"
	'[UI-ANT] end if
	'[UI-ANT] html = html & "</td>"

	'[UI-ANT] html = html & "</form>"

	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "</tr>"


	'[UI-ANT] rs.movenext
'[UI-ANT] Loop

'[UI-ANT] html = html & "</tbody>"

'[UI-ANT] 'html = html & "</table>"
'[UI-ANT] html = html & "</table>"

'[UI-ANT] 'html = html & "</div>"
'[UI-ANT] html = html & "</div>" ' Cierra .bloque-datos
'[UI-ANT] 'html = html & "</div>"
'[UI-ANT] html = html & "</div>" ' Cierra .inner

'[UI-ANT] html = html & "<script type='text/javascript'>"
'[UI-ANT] html = html & " function borrar(formu){ "
'[UI-ANT] html = html & " var r = confirm('Desea eliminar responsable?'); "
'[UI-ANT] html = html & " if (r){ "
'[UI-ANT] html = html & " document.getElementById(formu).submit(); "
'[UI-ANT] html = html & " }else{ "
'[UI-ANT] html = html & " return false;"
'[UI-ANT] html = html & " }} "
'[UI-ANT] html = html & "</script>"

'[UI-ANT] html = html & "<script type='text/javascript'>"
'[UI-ANT] html = html & " function modificar(formu){ "
'[UI-ANT] html = html & " document.getElementById(formu).submit();"
'[UI-ANT] html = html & " alert('Registro actualizado');} "
'[UI-ANT] html = html & "</script>"

'[UI-ANT] html = html & "<script src='assets/js/jquery.min.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/skel.min.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/util.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/main.js'></script>"

'[UI-ANT] html = html & formx1 & formx & "</body>"
'[UI-ANT] html = html & "</html>"

' =========================================================================
' [UI] VISTA con el sistema visual (IncFile/ui_componentes.inc)
' Pantalla de parametros: etiqueta "Parametros" sobre el titulo.
' El bloque anterior completo queda arriba, comentado con '[UI-ANT].
' Consulta, formularios de edicion y permisos: los mismos de antes.
' =========================================================================
Dim accTop, puedeModif, sinPermiso, cantResp, alCambiar

html = html & "<div id='wrapper' role='banner'>" & cabecera() & "</div>"

' --- Barra superior: Agregar ---
formx = "<form id='formx' class='form-oculto' method='post' action='agregar1.asp?tipo=CR&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
accTop = UiBoton("Agregar responsable", "light", "document.getElementById('formx').submit();", getPermisosAtrm("CG_RESPONSABLES","ALTA") = 1, "Sin permiso para agregar.")
html = html & UiPantallaParametrosAbre("Responsables", accTop, "wide")
html = html & UiCardAbre("Responsables de Cor", "")

html = html & UiTablaAbre("Responsables de Cor: datos, funci&oacute;n y estado")
html = html & "<thead><tr>" & UiTh("<span class='ui-sr-only'>Acciones</span>", "ui-table__col--action") & UiTh("ID", "") & _
	UiTh("Nombre", "") & UiTh("Apellido", "") & UiTh("Tel&eacute;fono", "") & UiTh("Mail", "") & _
	UiTh("Funci&oacute;n", "") & UiTh("Activo", "") & UiTh("Usuario", "") & "</tr></thead><tbody>"

Set rs = Cnn.Execute("sp_cg_Responsables @Modo='listar'")

formx1 = ""
cantResp = 0
puedeModif = (getPermisosAtrm("CG_RESPONSABLES","MODIF") = 1)
' Sin permiso: los campos avisan, como antes
sinPermiso = ""
If Not puedeModif Then sinPermiso = " onblur=""alert('Sin permiso para modificar.');"""

Do While Not rs.eof

	' Funcion junto al nombre (TEC / FUNC)
	xFlg = ""
	if rs(3) = "TEC" or rs(3) = "FUNC" then xFlg = "<span class='ui-tag-id'>" & rs(3) & "</span>"

	formx1 = formx1 & "<form id='form" & rs(0) & "' class='form-oculto' method='post' action='eliminar.asp?opc=corr&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=corresp.asp'></form>"

	html = html & "<tr>"
	html = html & "<td>" & UiBotonIcono("&times;", "Eliminar responsable " & rs(0), "danger", "borrar('form" & rs(0) & "');", getPermisosAtrm("CG_RESPONSABLES","ELIM") = 1, "Sin permiso para eliminar") & "</td>"
	html = html & "<td>" & rs(0) & "</td>"

	' Formulario de la fila (misma estructura de antes): funcion y activo lo envian al cambiar
	html = html & "<form id='iform" & rs(0) & "' style='display: contents;' method='post' action='corresp.asp?opc=act&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
	html = html & "<td><div class='ui-catalogo'>" & xFlg & "<input class='ui-input' name='nom' type='text' value='" & Server.HTMLEncode(rs(1) & "") & "' aria-label='Nombre'" & sinPermiso & " readonly></div></td>"
	html = html & "<td><input class='ui-input' name='ape' type='text' value='" & Server.HTMLEncode(rs(2) & "") & "' aria-label='Apellido'" & sinPermiso & " readonly></td>"
	html = html & "<td><input class='ui-input' name='tel' type='text' value='" & Server.HTMLEncode(rs(6) & "") & "' aria-label='Tel&eacute;fono'" & sinPermiso & " readonly></td>"
	html = html & "<td><input class='ui-input' name='mail' type='text' value='" & Server.HTMLEncode(rs(5) & "") & "' aria-label='Mail'" & sinPermiso & " readonly></td>"

	If puedeModif Then
		alCambiar = " onchange=""document.getElementById('iform" & rs(0) & "').submit();"""
	Else
		alCambiar = " onchange=""alert('Sin permiso para modificar.');"" disabled"
	End If

	html = html & "<td><select class='ui-select' name='fun' aria-label='Funci&oacute;n'" & alCambiar & ">"
	If rs(3) = "FUNC" Then
		html = html & "<option value='FUNC' selected>Funcional</option><option value='TEC'>T&eacute;cnico</option>"
	ElseIf rs(3) = "TEC" Then
		html = html & "<option value='TEC' selected>T&eacute;cnico</option><option value='FUNC'>Funcional</option>"
	End If
	html = html & "</select></td>"

	html = html & "<td><select class='ui-select' name='activo' aria-label='Activo'" & alCambiar & ">"
	If rs("activo") = "S" Then
		html = html & "<option value='S' selected>S&iacute;</option><option value='N'>No</option>"
	ElseIf rs("activo") = "N" Then
		html = html & "<option value='S'>S&iacute;</option><option value='N' selected>No</option>"
	Else
		html = html & "<option value='S'>S&iacute;</option><option value='N'>No</option>"
	End If
	html = html & "</select></td>"

	html = html & "<td><input class='ui-input' name='ID_SEG' type='text' value='" & Server.HTMLEncode(rs("ID_SEG") & "") & "' aria-label='Usuario'" & sinPermiso & " readonly></td>"
	html = html & "</form>"
	html = html & "</tr>"

	cantResp = cantResp + 1
	rs.movenext
Loop

If cantResp = 0 Then html = html & UiFilaVacia(9, "No hay responsables cargados.")
html = html & UiTablaCierra()
If cantResp = 1 Then
	html = html & UiPie("1 responsable", "")
ElseIf cantResp > 1 Then
	html = html & UiPie(cantResp & " responsables", "")
End If

html = html & UiCardCierra()
html = html & UiPantallaCierra()

html = html & "<script type='text/javascript'>"
html = html & " function borrar(formu){ "
html = html & " var r = confirm('\u00bfDesea eliminar el responsable?'); "
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