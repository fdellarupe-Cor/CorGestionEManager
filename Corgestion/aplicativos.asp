<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/ui_componentes.inc" -->
<%
html = headhtml()
html = html & "<body>"

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

opc = request.querystring("opc")
id = request.querystring("id")
xmenu = request.querystring("xmenu")
xtmenu = request.querystring("xtmenu")

XXX = SYSLOGON(cnn,"PARAM_APLIC")

If opc = "modif" Then

	Dim cmdUpd
	Set cmdUpd = Server.CreateObject("ADODB.Command")
	With cmdUpd
		.ActiveConnection = Cnn
		.CommandText = "sp_cg_Catalogo_Actualizar"
		.CommandType = 4
		.Parameters.Append .CreateParameter("@Tabla", 200, 1, 20, "aplicativos")
		.Parameters.Append .CreateParameter("@Id", 3, 1, 4, CLng(id))
		.Parameters.Append .CreateParameter("@Descripcion", 200, 1, 100, CStr(request.form("dsc")))
		.Execute
	End With
	Set cmdUpd = Nothing
	CgLog "APLICATIVO", id, "ACTUALIZA"

End If

'[UI-ANT] ' ===== Vista anterior completa: reemplazada por el bloque [UI] que sigue. Se borra entera buscando [UI-ANT] =====
'[UI-ANT] html = html & cabecera()

'[UI-ANT] html = html & "<div class='titulo-seccion ui-page-header' style='padding: 1rem 2rem;'>"
'[UI-ANT] html = html & "<div class='ui-toolbar' style='display:flex;align-items:center;justify-content:space-between;'>"

'[UI-ANT] html = html & "<div><h2>APLICATIVOS</h2></div>"

'[UI-ANT] formx = "<form id='formx' class='form-oculto' method='post' action='agregar1.asp?tipo=AP&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"

'[UI-ANT] if getPermisos(xmenu,xtmenu,"ALTA") = 1 then
	'[UI-ANT] botonx = "<input class='bot ui-button ui-button--primary' type='button' value='AGREGAR APLICATIVO' onclick=""document.getElementById('formx').submit();"">"
'[UI-ANT] else
	'[UI-ANT] botonx = "<input class='botsp ui-button' type='button' value='AGREGAR APLICATIVO' onclick=""alert('Sin permiso para agregar.');"">"
'[UI-ANT] end if

'[UI-ANT] html = html & "<div class='ui-toolbar__actions'>" & botonx & "</div>"

'[UI-ANT] html = html & "</div>" ' Fin ui-toolbar
'[UI-ANT] html = html & "</div>" ' Fin titulo-seccion

	'[UI-ANT] ' --- INICIO NORMALIZACION: TABLA DE APLICATIVOS (EDICION INLINE) ---

	'[UI-ANT] 'html = html & "<div style='height:57%;width:100%;float:left;overflow:auto;'>"
	'[UI-ANT] html = html & "<div class='inner ui-main' style='padding: 2rem;'>"
	'[UI-ANT] html = html & "<div class='bloque-datos table-responsive ui-table-wrap'>"

	'[UI-ANT] 'html = html & "<table style='width:35%;heigth:80%;font-family:Serif;text-align:center;' cellspacing=0 cellpadding=3 align=center>"
	'[UI-ANT] html = html & "<table class='tabla-moderna ui-table'>"

	'[UI-ANT] html = html & "<thead>"
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] html = html & "<tr>"
	
	'[UI-ANT] 'html = html & "<td class='encabe' style='width:1%;' align=center>*</td>"
	'[UI-ANT] html = html & "<th style='width: 50px; text-align: center;'>&nbsp;</th>"
	
	'[UI-ANT] 'html = html & "<td class='encabe' style='width:1%;' align=center>ID</td>"
	'[UI-ANT] html = html & "<th style='width: 60px; text-align: center;'>ID</th>"
	
	'[UI-ANT] 'html = html & "<td class='encabe' style='width:18%;' align=center>DESCRIPCION</td>"
	'[UI-ANT] html = html & "<th>DESCRIPCION</th>"
	
	'[UI-ANT] 'html = html & "<td class='encabe' style='width:15%;' align=center>CLIENTES</td>"
	'[UI-ANT] html = html & "<th style='text-align: center;'>CLIENTES</th>"
	
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "</tr>"
	'[UI-ANT] html = html & "</thead>"

	'[UI-ANT] Set rs = Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='aplicativos'")

	'[UI-ANT] formx1 = ""
	'[UI-ANT] formx2 = ""

	'[UI-ANT] html = html & "<tbody>"

	'[UI-ANT] Do While Not rs.eof

		'[UI-ANT] if getPermisos(xmenu,xtmenu,"MODIF") = 1 then
			'[UI-ANT] 'link = "<a class='red' href='detalleAplic.asp?id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>VER / MODIF</a>"
			'[UI-ANT] link = "<a class='col-link' style='color: blue; text-decoration: none; font-weight: bold; cursor: pointer;' href='detalleAplic.asp?id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>VER / MODIF</a>"
		'[UI-ANT] else
			'[UI-ANT] 'link = "<a class='red' onclick=""alert('Sin permiso para modificar.');"">VER / MODIF</a>"
			'[UI-ANT] ' Lo pasamos a un color grisaceo para indicar visualmente que no esta disponible
			'[UI-ANT] link = "<a style='color: #94a3b8; text-decoration: none; font-weight: bold; cursor: not-allowed;' onclick=""alert('Sin permiso para modificar.');"">VER / MODIF</a>"
		'[UI-ANT] end if

		'[UI-ANT] 'boton eliminar
		'[UI-ANT] 'formx1 = formx1 & "<form id='form" & rs(0) & "' style='height:1px;' method='post' action='eliminar.asp?opc=apli&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=aplicativos.asp'></form>"
		'[UI-ANT] formx1 = formx1 & "<form id='form" & rs(0) & "' class='form-oculto' method='post' action='eliminar.asp?opc=apli&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=aplicativos.asp'></form>"

		'[UI-ANT] if getPermisos(xmenu,xtmenu,"ELIM") = 1 then
			'[UI-ANT] 'botonx1 = "<input class='bot' type='button' value='X' onclick=""borrar('form" & rs(0) & "');"">"
			'[UI-ANT] botonx1 = "<input class='btn-accion btn-accion-peligro ui-button ui-button--icon ui-button--danger' type='button' value='X' onclick=""borrar('form" & rs(0) & "');"">"
		'[UI-ANT] else
			'[UI-ANT] 'botonx1 = "<input class='botsp' type='button' value='X' onclick=""alert('Sin permiso para eliminar.');"">"
			'[UI-ANT] botonx1 = "<input class='btn-accion btn-accion-deshabilitado ui-button ui-button--icon' type='button' value='X' onclick=""alert('Sin permiso para eliminar.');"">"
		'[UI-ANT] end if

		'[UI-ANT] 'descripcion
		'[UI-ANT] 'inputx10 = "<form id='sformx" & rs(0) & "' style='height:1px;' method='post' action='aplicativos.asp?opc=modif&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
		'[UI-ANT] inputx10 = "<form id='sformx" & rs(0) & "' class='form-oculto' method='post' action='aplicativos.asp?opc=modif&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"

		'[UI-ANT] if getPermisos(xmenu,xtmenu,"MODIF") = 1 then
			'[UI-ANT] inputx10 = inputx10 & HtmlCampoCatalogo("dsc", rs(1), "sformx" & rs(0), True) & "</form>"
		'[UI-ANT] else
			'[UI-ANT] inputx10 = inputx10 & "<input class='input-moderno ui-input' name='dsc' type='text' value='" & Server.HTMLEncode(rs(1) & "") & "' readonly></form>"
		'[UI-ANT] end if

		'[UI-ANT] 'html = html & "<tr>"
		'[UI-ANT] html = html & "<tr>"
		
		'[UI-ANT] 'html = html & "<td class='regis' >" & botonx1 & "</td>"
		'[UI-ANT] html = html & "<td style='text-align: center; vertical-align: middle;'>" & botonx1 & "</td>"
		
		'[UI-ANT] 'html = html & "<td class='regis' >" & rs(0) & "</td>"
		'[UI-ANT] html = html & "<td style='text-align: center; vertical-align: middle; color: #475569; font-weight: 500;'>" & rs(0) & "</td>"
		
		'[UI-ANT] 'html = html & "<td class='regis' >" & inputx10 & "</td>"
		'[UI-ANT] html = html & "<td style='vertical-align: middle;'>" & inputx10 & "</td>"
		
		'[UI-ANT] 'html = html & "<td class='regis' >" & link & "</td>"
		'[UI-ANT] html = html & "<td style='text-align: center; vertical-align: middle;'>" & link & "</td>"
		
		'[UI-ANT] 'html = html & "</tr>"
		'[UI-ANT] html = html & "</tr>"

		'[UI-ANT] rs.movenext
	'[UI-ANT] Loop

	'[UI-ANT] 'html = html & "</table>"
	'[UI-ANT] html = html & "</tbody>"
	'[UI-ANT] html = html & "</table>"
	'[UI-ANT] html = html & "</div>"
	
	'[UI-ANT] 'html = html & "</div>"
	'[UI-ANT] html = html & "</div>" ' Cierre del inner

'[UI-ANT] html = html & "<script type='text/javascript'>"
'[UI-ANT] html = html & " function borrar(formu){ "
'[UI-ANT] html = html & " var r = confirm('Desea eliminar aplicativo?'); "
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


' =========================================================================
' [UI] VISTA con el sistema visual (IncFile/ui_componentes.inc)
' Pantalla de parametros: etiqueta "Parametros" sobre el titulo.
' El bloque anterior completo queda arriba, comentado con '[UI-ANT].
' Consulta, formularios, campo de edicion y permisos: los mismos de antes.
' =========================================================================
Dim accTop, campoCat, celdasExtra, cantCat, colsCat

html = html & "<div id='wrapper' role='banner'>" & cabecera() & "</div>"

' --- Barra superior: Agregar ---
formx = "<form id='formx' class='form-oculto' method='post' action='agregar1.asp?tipo=AP&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
accTop = UiBoton("Agregar aplicativo", "light", "document.getElementById('formx').submit();", getPermisosAtrm("CG_APLICATIVOS","ALTA") = 1, "Sin permiso para agregar.")
html = html & UiPantallaParametrosAbre("Aplicativos", accTop, "")
html = html & UiCardAbre("Listado de aplicativos", "")

html = html & UiTablaAbre("Listado de aplicativos")
html = html & "<thead><tr>" & UiTh("<span class='ui-sr-only'>Acciones</span>", "ui-table__col--action") & UiTh("ID", "") & UiTh("Descripci&oacute;n", "") & UiTh("Clientes", "ui-table__cell--center") & "</tr></thead><tbody>"

Set rs = Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='aplicativos'")

formx1 = ""
cantCat = 0
celdasExtra = ""

Do While Not rs.eof

	xRef = ""

	' Eliminar (mismo formulario de antes)
	formx1 = formx1 & "<form id='form" & rs(0) & "' class='form-oculto' method='post' action='eliminar.asp?opc=apli&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=aplicativos.asp'></form>"

	' Descripcion editable en linea: el campo y el boton los arma HtmlCampoCatalogo (funciones.inc).
	' ui-catalogo hace que el texto ocupe el ancho disponible sin aplastar el boton.
	campoCat = "<form id='sformx" & rs(0) & "' class='form-oculto' method='post' action='aplicativos.asp?opc=modif&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
	if getPermisosAtrm("CG_APLICATIVOS","MODIF") = 1 then
		' El label (solo para lectores de pantalla) le pone nombre al campo que arma HtmlCampoCatalogo
		campoCat = campoCat & "<label class='ui-catalogo'><span class='ui-sr-only'>Descripci&oacute;n</span>" & HtmlCampoCatalogo("dsc", rs(1), "sformx" & rs(0), True) & xRef & "</label></form>"
	else
		campoCat = campoCat & "<div class='ui-catalogo'><input class='ui-input' name='dsc' type='text' value='" & Server.HTMLEncode(rs(1) & "") & "' aria-label='Descripci&oacute;n' readonly>" & xRef & "</div></form>"
	end if

	' Clientes que tienen el aplicativo (mismo enlace y permiso de antes)
	If getPermisosAtrm("CG_APLICATIVOS","MODIF") = 1 Then
		celdasExtra = "<td class='ui-table__cell--center'><a class='ui-button ui-button--secondary ui-button--sm' href='detalleAplic.asp?id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>Ver / modificar</a></td>"
	Else
		celdasExtra = "<td class='ui-table__cell--center'>" & UiBoton("Ver / modificar", "secondary", "", False, "Sin permiso para modificar.") & "</td>"
	End If

	html = html & "<tr>" & _
		"<td>" & UiBotonIcono("&times;", "Eliminar aplicativo " & rs(0), "danger", "borrar('form" & rs(0) & "');", getPermisosAtrm("CG_APLICATIVOS","ELIM") = 1, "Sin permiso para eliminar") & "</td>" & _
		"<td>" & rs(0) & "</td>" & _
		"<td>" & campoCat & "</td>" & _
		celdasExtra & _
		"</tr>"

	cantCat = cantCat + 1
	rs.movenext
Loop

colsCat = 4
If cantCat = 0 Then html = html & UiFilaVacia(colsCat, "No hay aplicativos cargados.")
html = html & UiTablaCierra()
If cantCat = 1 Then
	html = html & UiPie("1 aplicativo", "")
ElseIf cantCat > 1 Then
	html = html & UiPie(cantCat & " aplicativos", "")
End If

html = html & UiCardCierra()
html = html & UiPantallaCierra()

html = html & "<script type='text/javascript'>"
html = html & " function borrar(formu){ "
html = html & " var r = confirm('\u00bfDesea eliminar el aplicativo?'); "
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

html = html & formx1 & formx2 & formx & "</body>"
html = html & "</html>"

response.write html
%>