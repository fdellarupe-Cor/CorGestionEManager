<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
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

XXX = SYSLOGON(cnn,"PARAM_PRIORI")

opc = request.querystring("opc")
xmenu = request.querystring("xmenu")
xtmenu = request.querystring("xtmenu")

If opc = "modif" Then

	id = request.querystring("id")

	Dim cmdUpd
	Set cmdUpd = Server.CreateObject("ADODB.Command")
	With cmdUpd
		.ActiveConnection = Cnn
		.CommandText = "sp_cg_Catalogo_Actualizar"
		.CommandType = 4
		.Parameters.Append .CreateParameter("@Tabla", 200, 1, 20, "prioridades")
		.Parameters.Append .CreateParameter("@Id", 3, 1, 4, CLng(id))
		.Parameters.Append .CreateParameter("@Descripcion", 200, 1, 100, CStr(request.form("dsc")))
		.Execute
	End With
	Set cmdUpd = Nothing

End If

html = html & cabecera()

html = html & "<div class='titulo-seccion ui-page-header' style='padding: 1rem 2rem;'>"
html = html & "<div class='ui-toolbar' style='display:flex;align-items:center;justify-content:space-between;'>"

html = html & "<div><h2>PRIORIDADES</h2></div>"

formx = "<form id='formx' class='form-oculto' method='post' action='agregar1.asp?tipo=PR&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
if getPermisos(xmenu,xtmenu,"ALTA") = 1 then
	botonx = "<input class='bot ui-button ui-button--primary' type='button' value='AGREGAR PRIORIDAD' onclick=""document.getElementById('formx').submit();"">"
else
	botonx = "<input class='botsp ui-button' type='button' value='AGREGAR PRIORIDAD' onclick=""alert('Sin permiso para agregar.');"">"
end if

html = html & "<div class='ui-toolbar__actions'>" & botonx & "</div>"

html = html & "</div>" ' Fin ui-toolbar
html = html & "</div>" ' Fin titulo-seccion

'html = html & "<div style='height:57%;width:100%;'>"
html = html & "<div class='inner ui-main' style='padding:2rem;'>"
html = html & "<div class='bloque-datos table-responsive ui-table-wrap'>"

'html = html & "<table style='width:20%;heigth:80%;font-family:Serif;text-align:center;' cellspacing=0 cellpadding=3 align=center>"
html = html & "<table class='tabla-moderna ui-table'>"

html = html & "<thead>"
'html = html & "<tr>"
html = html & "<tr>"

'html = html & "<td class='encabe' style='width:1%;' align=center>*</td>"
html = html & "<th style='width: 40px;'>&nbsp;</th>"

'html = html & "<td class='encabe' style='width:1%;' align=center>ID</td>"
html = html & "<th style='width: 10%;'>ID</th>"

'html = html & "<td class='encabe' style='width:18%;' align=center>DESCRIPCION</td>"
html = html & "<th style='width: 90%;'>DESCRIPCION</th>"

'html = html & "</tr>"
html = html & "</tr>"
html = html & "</thead>"

Set rs = Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='prioridades'")

formx1 = ""

html = html & "<tbody>"

Do While Not rs.eof

	xRef = ""
	' Reemplazamos el font viejo por un span moderno para el indicador visual
	if rs(1) = "ALTA" then
		'xRef = "<font style='font-size:12pt;color:#FF0000;'><b>&nbsp;&nbsp;&nbsp;(!)</b></font>"
		xRef = "<span style='color: #dc2626; font-weight: bold;'>(!)</span>"
	else
		'xRef = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
		xRef = "" ' Ya no hace falta empujar con espacios gracias a flexbox
	end if

	'formx1 = formx1 & "<form id='form" & rs(0) & "' style='height:1px;' method='post' action='eliminar.asp?opc=prio&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=prioridades.asp'></form>"
	formx1 = formx1 & "<form id='form" & rs(0) & "' class='form-oculto' method='post' action='eliminar.asp?opc=prio&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=prioridades.asp'></form>"
	
	if getPermisos(xmenu,xtmenu,"ELIM") = 1 then
		'botonx1 = "<input class='bot' type='button' value='X' onclick=""borrar('form" & rs(0) & "');"">"
		botonx1 = "<input class='btn-accion btn-accion-peligro ui-button ui-button--icon ui-button--danger' type='button' value='X' onclick=""borrar('form" & rs(0) & "');"">"
	else
		'botonx1 = "<input class='botsp' type='button' value='X' onclick=""alert('Sin permiso para eliminar.');"">"
		botonx1 = "<input class='btn-accion btn-accion-deshabilitado ui-button ui-button--icon' type='button' value='X' onclick=""alert('Sin permiso para eliminar.');"">"
	end if

	'html = html & "<tr>"
	html = html & "<tr>"
	
	'html = html & "<td class='regis' >" & botonx1 & "</td>"
	html = html & "<td>" & botonx1 & "</td>"
	
	'html = html & "<td class='regis' >" & rs(0) & "</td>"
	html = html & "<td>" & rs(0) & "</td>"
	
	'html = html & "<td class='regis' >"
	html = html & "<td>"
	
	'html = html & "<form style='height:1px;' id='sformx" & rs(0) & "' method='post' action='prioridades.asp?opc=modif&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
	html = html & "<form id='sformx" & rs(0) & "' class='form-linea' method='post' action='prioridades.asp?opc=modif&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
	
	if getPermisos(xmenu,xtmenu,"MODIF") = 1 then
		html = html & "<div style='display: flex; align-items: center; gap: 0.5rem;'>" & HtmlCampoCatalogo("dsc", rs(1), "sformx" & rs(0), True) & xRef & "</div>"
	else
		html = html & "<div style='display: flex; align-items: center; gap: 0.5rem;'><input class='input-moderno input-deshabilitado ui-input' name='dsc' type='text' value='" & Server.HTMLEncode(rs(1) & "") & "' readonly>" & xRef & "</div>"
	end if
	
	'html = html & "</form>"
	html = html & "</form>"
	
	'html = html & "</td>"
	html = html & "</td>"
	
	'html = html & "</tr>"
	html = html & "</tr>"

	'html = html & "</form>"
	' Etiqueta huérfana de cierre comentada por seguridad del DOM
	'html = html & "</form>"

	rs.movenext
Loop

html = html & "</tbody>"

'html = html & "</table>"
html = html & "</table>"

'html = html & "</div>"
html = html & "</div>" ' Cierra .bloque-datos table-responsive

'html = html & "</div>"
html = html & "</div>" ' Cierra .inner

html = html & "<script type='text/javascript'>"
html = html & " function borrar(formu){ "
html = html & " var r = confirm('Desea eliminar prioridad?'); "
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

html = html & formx1 & formx & "</body>"
html = html & "</html>"

response.write html
%>