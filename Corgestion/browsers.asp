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

XXX = SYSLOGON(cnn,"PARAM_BROWSER")

id = request.querystring("id")
opc = request.querystring("opc")
xmenu = request.querystring("xmenu")
xtmenu = request.querystring("xtmenu")

If opc = "modif" Then

	Dim cmdUpd
	Set cmdUpd = Server.CreateObject("ADODB.Command")
	With cmdUpd
		.ActiveConnection = Cnn
		.CommandText = "sp_cg_Catalogo_Actualizar"
		.CommandType = 4
		.Parameters.Append .CreateParameter("@Tabla", 200, 1, 20, "browsers")
		.Parameters.Append .CreateParameter("@Id", 3, 1, 4, CLng(id))
		.Parameters.Append .CreateParameter("@Descripcion", 200, 1, 100, CStr(request.form("inp")))
		.Execute
	End With
	Set cmdUpd = Nothing
	CgLog "BROWSER", id, "ACTUALIZA"

End If

'[UI-ANT] ' ===== Vista anterior completa: reemplazada por el bloque [UI] que sigue. Se borra entera buscando [UI-ANT] =====
'[UI-ANT] html = html & cabecera()

'[UI-ANT] html = html & "<div class='titulo-seccion ui-page-header' style='padding: 1rem 2rem;'>"
'[UI-ANT] html = html & "<div class='ui-toolbar' style='display:flex;align-items:center;justify-content:space-between;'>"

'[UI-ANT] html = html & "<div><h2>BROWSERS</h2></div>"

'[UI-ANT] formx = "<form id='formx' class='form-oculto' method='post' action='agregar1.asp?tipo=BR&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"

'[UI-ANT] if getPermisos(xmenu,xtmenu,"ALTA") = 1 then
	'[UI-ANT] botonx = "<input class='bot ui-button ui-button--primary' type='button' value='AGREGAR BROWSER' onclick=""document.getElementById('formx').submit();"">"
'[UI-ANT] else
	'[UI-ANT] botonx = "<input class='botsp ui-button' type='button' value='AGREGAR BROWSER' onclick=""alert('Sin permiso para agregar.');"">"
'[UI-ANT] end if

'[UI-ANT] html = html & "<div class='ui-toolbar__actions'>" & botonx & "</div>"

'[UI-ANT] html = html & "</div>" ' Fin ui-toolbar
'[UI-ANT] html = html & "</div>" ' Fin titulo-seccion

'[UI-ANT] ' --- INICIO NORMALIZACION: TABLA DE BROWSERS ---

'[UI-ANT] 'html = html & "<div style='height:57%;width:100%;'>"
'[UI-ANT] html = html & "<div class='inner ui-main' style='padding: 2rem;'>"
'[UI-ANT] html = html & "<div class='bloque-datos table-responsive ui-table-wrap' style='max-width: 600px; margin: 0 auto;'>"

'[UI-ANT] 'html = html & "<table style='width:20%;heigth:80%;font-family:Serif;text-align:center;' cellspacing=0 cellpadding=3 align=center>"
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
	
'[UI-ANT] 'html = html & "</tr>"
'[UI-ANT] html = html & "</tr>"
'[UI-ANT] html = html & "</thead>"

'[UI-ANT] Set rs = Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='browsers'")

'[UI-ANT] formx1 = ""
'[UI-ANT] formux = ""

'[UI-ANT] Do While Not rs.eof

		'[UI-ANT] 'boton eliminar
		'[UI-ANT] 'formx1 = formx1 & "<form id='form" & rs(0) & "' style='height:1px;' method='post' action='eliminar.asp?opc=brow&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=browsers.asp'></form>"
		'[UI-ANT] formx1 = formx1 & "<form id='form" & rs(0) & "' class='form-oculto' method='post' action='eliminar.asp?opc=brow&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=browsers.asp'></form>"
	'[UI-ANT] if getPermisos(xmenu,xtmenu,"ELIM") = 1 then
		'[UI-ANT] 'botonx1 = "<input class='bot' type='button' value='X' onclick=""borrar('form" & rs(0) & "');"">"
		'[UI-ANT] botonx1 = "<input class='btn-accion btn-accion-peligro ui-button ui-button--icon ui-button--danger' type='button' value='X' onclick=""borrar('form" & rs(0) & "');"">"
	'[UI-ANT] else
		'[UI-ANT] 'botonx1 = "<input class='botsp' type='button' value='X' onclick=""alert('Sin permiso para eliminar.');"">"
		'[UI-ANT] botonx1 = "<input class='btn-accion btn-accion-deshabilitado ui-button ui-button--icon' type='button' value='X' onclick=""alert('Sin permiso para eliminar.');"">"
	'[UI-ANT] end if

	'[UI-ANT] 'descripcion
	'[UI-ANT] 'xinput = "<form style='height:0px;' id='formx" & rs(0) & "' method='post' action='browsers.asp?opc=modif&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
	'[UI-ANT] xinput = "<form id='formx" & rs(0) & "' class='form-oculto' method='post' action='browsers.asp?opc=modif&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
	'[UI-ANT] if getPermisos(xmenu,xtmenu,"MODIF") = 1 then
		'[UI-ANT] xinput = xinput & HtmlCampoCatalogo("inp", rs(1), "formx" & rs(0), True) & "</form>"
	'[UI-ANT] else
		'[UI-ANT] xinput = xinput & "<input class='input-moderno ui-input' id='inp' name='inp' type='text' value='" & Server.HTMLEncode(rs(1) & "") & "' readonly></form>"
	'[UI-ANT] end if

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] html = html & "<tr>"
		
	'[UI-ANT] 'html = html & "<td class='regis' >" & botonx1 & "</td>"
	'[UI-ANT] html = html & "<td style='text-align: center; vertical-align: middle;'>" & botonx1 & "</td>"
		
	'[UI-ANT] 'html = html & "<td class='regis' >" & rs(0) & "</td>"
	'[UI-ANT] html = html & "<td style='text-align: center; vertical-align: middle; color: #475569; font-weight: 500;'>" & rs(0) & "</td>"
		
	'[UI-ANT] 'html = html & "<td class='regis' >" & xinput & "</td>"
	'[UI-ANT] html = html & "<td style='vertical-align: middle;'>" & xinput & "</td>"
		
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "</tr>"

	'[UI-ANT] rs.movenext
'[UI-ANT] Loop

'[UI-ANT] 'html = html & "</table>"
'[UI-ANT] html = html & "</tbody>"
'[UI-ANT] html = html & "</table>"
'[UI-ANT] html = html & "</div>"
	
'[UI-ANT] 'html = html & "</div>"
'[UI-ANT] html = html & "</div>" ' Cierre de .inner

'[UI-ANT] 'html = html & "</div>"
'[UI-ANT] html = html & "</div>" ' Cierre del div contenedor extra de tu codigo original

'[UI-ANT] html = html & "<script type='text/javascript'>"
'[UI-ANT] html = html & " function borrar(formu){ "
'[UI-ANT] html = html & " var r = confirm('Desea eliminar browser?'); "
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
formx = "<form id='formx' class='form-oculto' method='post' action='agregar1.asp?tipo=BR&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
accTop = UiBoton("Agregar browser", "light", "document.getElementById('formx').submit();", getPermisosAtrm("CG_BROWSERS","ALTA") = 1, "Sin permiso para agregar.")
html = html & UiPantallaParametrosAbre("Browsers", accTop, "narrow")
html = html & UiCardAbre("Listado de browsers", "")

html = html & UiTablaAbre("Listado de browsers")
html = html & "<thead><tr>" & UiTh("<span class='ui-sr-only'>Acciones</span>", "ui-table__col--action") & UiTh("ID", "") & UiTh("Descripci&oacute;n", "") & "</tr></thead><tbody>"

Set rs = Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='browsers'")

formx1 = ""
cantCat = 0
celdasExtra = ""

Do While Not rs.eof

	xRef = ""

	' Eliminar (mismo formulario de antes)
	formx1 = formx1 & "<form id='form" & rs(0) & "' class='form-oculto' method='post' action='eliminar.asp?opc=brow&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=browsers.asp'></form>"

	' Descripcion editable en linea: el campo y el boton los arma HtmlCampoCatalogo (funciones.inc).
	' ui-catalogo hace que el texto ocupe el ancho disponible sin aplastar el boton.
	campoCat = "<form id='formx" & rs(0) & "' class='form-oculto' method='post' action='browsers.asp?opc=modif&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
	if getPermisosAtrm("CG_BROWSERS","MODIF") = 1 then
		' El label (solo para lectores de pantalla) le pone nombre al campo que arma HtmlCampoCatalogo
		campoCat = campoCat & "<label class='ui-catalogo'><span class='ui-sr-only'>Descripci&oacute;n</span>" & HtmlCampoCatalogo("inp", rs(1), "formx" & rs(0), True) & xRef & "</label></form>"
	else
		campoCat = campoCat & "<div class='ui-catalogo'><input class='ui-input' name='inp' type='text' value='" & Server.HTMLEncode(rs(1) & "") & "' aria-label='Descripci&oacute;n' readonly>" & xRef & "</div></form>"
	end if



	html = html & "<tr>" & _
		"<td>" & UiBotonIcono("&times;", "Eliminar browser " & rs(0), "danger", "borrar('form" & rs(0) & "');", getPermisosAtrm("CG_BROWSERS","ELIM") = 1, "Sin permiso para eliminar") & "</td>" & _
		"<td>" & rs(0) & "</td>" & _
		"<td>" & campoCat & "</td>" & _
		celdasExtra & _
		"</tr>"

	cantCat = cantCat + 1
	rs.movenext
Loop

colsCat = 3
If cantCat = 0 Then html = html & UiFilaVacia(colsCat, "No hay browsers cargados.")
html = html & UiTablaCierra()
If cantCat = 1 Then
	html = html & UiPie("1 browser", "")
ElseIf cantCat > 1 Then
	html = html & UiPie(cantCat & " browsers", "")
End If

html = html & UiCardCierra()
html = html & UiPantallaCierra()

html = html & "<script type='text/javascript'>"
html = html & " function borrar(formu){ "
html = html & " var r = confirm('\u00bfDesea eliminar el browser?'); "
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

html = html & formx1 & formux & formx & "</body>"
html = html & "</html>"

response.write html
%>