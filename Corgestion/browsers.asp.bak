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

End If

html = html & cabecera()

html = html & "<div class='titulo-seccion ui-page-header' style='padding: 1rem 2rem;'>"
html = html & "<div class='ui-toolbar' style='display:flex;align-items:center;justify-content:space-between;'>"

html = html & "<div><h2>BROWSERS</h2></div>"

formx = "<form id='formx' class='form-oculto' method='post' action='agregar1.asp?tipo=BR&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"

if getPermisos(xmenu,xtmenu,"ALTA") = 1 then
	botonx = "<input class='bot ui-button ui-button--primary' type='button' value='AGREGAR BROWSER' onclick=""document.getElementById('formx').submit();"">"
else
	botonx = "<input class='botsp ui-button' type='button' value='AGREGAR BROWSER' onclick=""alert('Sin permiso para agregar.');"">"
end if

html = html & "<div class='ui-toolbar__actions'>" & botonx & "</div>"

html = html & "</div>" ' Fin ui-toolbar
html = html & "</div>" ' Fin titulo-seccion

' --- INICIO NORMALIZACION: TABLA DE BROWSERS ---

'html = html & "<div style='height:57%;width:100%;'>"
html = html & "<div class='inner ui-main' style='padding: 2rem;'>"
html = html & "<div class='bloque-datos table-responsive ui-table-wrap' style='max-width: 600px; margin: 0 auto;'>"

'html = html & "<table style='width:20%;heigth:80%;font-family:Serif;text-align:center;' cellspacing=0 cellpadding=3 align=center>"
html = html & "<table class='tabla-moderna ui-table'>"

html = html & "<thead>"
'html = html & "<tr>"
html = html & "<tr>"
	
'html = html & "<td class='encabe' style='width:1%;' align=center>*</td>"
html = html & "<th style='width: 50px; text-align: center;'>&nbsp;</th>"
	
'html = html & "<td class='encabe' style='width:1%;' align=center>ID</td>"
html = html & "<th style='width: 60px; text-align: center;'>ID</th>"
	
'html = html & "<td class='encabe' style='width:18%;' align=center>DESCRIPCION</td>"
html = html & "<th>DESCRIPCION</th>"
	
'html = html & "</tr>"
html = html & "</tr>"
html = html & "</thead>"

Set rs = Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='browsers'")

formx1 = ""
formux = ""

Do While Not rs.eof

		'boton eliminar
		'formx1 = formx1 & "<form id='form" & rs(0) & "' style='height:1px;' method='post' action='eliminar.asp?opc=brow&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=browsers.asp'></form>"
		formx1 = formx1 & "<form id='form" & rs(0) & "' class='form-oculto' method='post' action='eliminar.asp?opc=brow&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=browsers.asp'></form>"
	if getPermisos(xmenu,xtmenu,"ELIM") = 1 then
		'botonx1 = "<input class='bot' type='button' value='X' onclick=""borrar('form" & rs(0) & "');"">"
		botonx1 = "<input class='btn-accion btn-accion-peligro ui-button ui-button--icon ui-button--danger' type='button' value='X' onclick=""borrar('form" & rs(0) & "');"">"
	else
		'botonx1 = "<input class='botsp' type='button' value='X' onclick=""alert('Sin permiso para eliminar.');"">"
		botonx1 = "<input class='btn-accion btn-accion-deshabilitado ui-button ui-button--icon' type='button' value='X' onclick=""alert('Sin permiso para eliminar.');"">"
	end if

	'descripcion
	'xinput = "<form style='height:0px;' id='formx" & rs(0) & "' method='post' action='browsers.asp?opc=modif&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
	xinput = "<form id='formx" & rs(0) & "' class='form-oculto' method='post' action='browsers.asp?opc=modif&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
	if getPermisos(xmenu,xtmenu,"MODIF") = 1 then
		xinput = xinput & HtmlCampoCatalogo("inp", rs(1), "formx" & rs(0), True) & "</form>"
	else
		xinput = xinput & "<input class='input-moderno ui-input' id='inp' name='inp' type='text' value='" & Server.HTMLEncode(rs(1) & "") & "' readonly></form>"
	end if

	'html = html & "<tr>"
	html = html & "<tr>"
		
	'html = html & "<td class='regis' >" & botonx1 & "</td>"
	html = html & "<td style='text-align: center; vertical-align: middle;'>" & botonx1 & "</td>"
		
	'html = html & "<td class='regis' >" & rs(0) & "</td>"
	html = html & "<td style='text-align: center; vertical-align: middle; color: #475569; font-weight: 500;'>" & rs(0) & "</td>"
		
	'html = html & "<td class='regis' >" & xinput & "</td>"
	html = html & "<td style='vertical-align: middle;'>" & xinput & "</td>"
		
	'html = html & "</tr>"
	html = html & "</tr>"

	rs.movenext
Loop

'html = html & "</table>"
html = html & "</tbody>"
html = html & "</table>"
html = html & "</div>"
	
'html = html & "</div>"
html = html & "</div>" ' Cierre de .inner

'html = html & "</div>"
html = html & "</div>" ' Cierre del div contenedor extra de tu código original

html = html & "<script type='text/javascript'>"
html = html & " function borrar(formu){ "
html = html & " var r = confirm('Desea eliminar browser?'); "
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