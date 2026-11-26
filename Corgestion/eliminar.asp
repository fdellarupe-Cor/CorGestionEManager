<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/datosEliminar.inc" -->
<%
' [CG] Control de usuario E-Manager (igual que CorEManager/Acciones.asp)
cgOk = ofv.CheckUsuario()
If cgOk <> "Y" Then
	cgUv = ofv.uv
	If cgUv = "NO" Then cgOk = ofv.ValidarUsuario
	Response.Write cgOk
	Response.End
End If


ruta=server.mappath(request.servervariables("script_name"))

pos=InStrRev(ruta,"\")
ruta=Mid(ruta,1,pos)

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

response.write "1.Eliminar"
If opc = "inc" Then

	Exito = EliminarRegistro(Cnn, opc, id, session("usuario"))

	set fso = SERVER.CreateObject("Scripting.FileSystemObject")
	ruta=ruta & "incidentes\"  &   id
			  ' response.write ruta
    If  fso.FolderExists(ruta ) Then
	 ' revisar si corresponde borrar
	'  fso.DeleteFolder(ruta)
   end if
   set fso=nothing
	response.write "2.Eliminar"
ElseIf opc = "proy" Then

	Exito = EliminarRegistro(Cnn, opc, id, session("usuario"))

ElseIf opc = "rec" Then

	Exito = EliminarRegistro(Cnn, opc, id, session("usuario"))

ElseIf opc = "apli" Then

	if EliminarRegistro(Cnn, opc, id, session("usuario")) then

	html = html & "<script type='text/javascript'>"
	html = html & " alert('Borrado exitosamente');"
	html = html & "</script>"

	else

	html = html & "<script type='text/javascript'>"
	html = html & " alert('El aplicativo no pudo ser borrado');"
	html = html & "</script>"

	end if

ElseIf opc = "brow" Then

	if EliminarRegistro(Cnn, opc, id, session("usuario")) then

	html = html & "<script type='text/javascript'>"
	html = html & " alert('Borrado exitosamente');"
	html = html & "</script>"

	else

	html = html & "<script type='text/javascript'>"
	html = html & " alert('El browser no pudo ser borrado');"
	html = html & "</script>"

	end if

ElseIf opc = "cli" Then
	response.write "envia orden eliminar"
	if EliminarRegistro(Cnn, opc, id, session("usuario")) then
		response.write " eliminar ok"
	html = html & "<script type='text/javascript'>"
	html = html & " alert('Borrado exitosamente');"
	html = html & "</script>"

	else
		response.write " eliminar error"
	html = html & "<script type='text/javascript'>"
	html = html & " alert('El cliente no pudo ser borrado');"
	html = html & "</script>"

	end if

ElseIf opc = "con" Then

	if EliminarRegistro(Cnn, opc, id, session("usuario")) then

	html = html & "<script type='text/javascript'>"
	html = html & " alert('Borrado exitosamente');"
	html = html & "</script>"

	else

	html = html & "<script type='text/javascript'>"
	html = html & " alert('El contacto no pudo ser borrado');"
	html = html & "</script>"

	end if

ElseIf opc = "corr" Then

	'controla para poder borrar que el usuario no este cargado en ningun objeto del sistema (ahora lo valida el SP)
	if EliminarRegistro(Cnn, opc, id, session("usuario")) then

	html = html & "<script type='text/javascript'>"
	html = html & " alert('Borrado exitosamente');"
	html = html & "</script>"

	else

	html = html & "<script type='text/javascript'>"
	html = html & " alert('El responsable no pudo ser borrado');"
	html = html & "</script>"

	end if

ElseIf opc = "est" Then

	if EliminarRegistro(Cnn, opc, id, session("usuario")) then

	html = html & "<script type='text/javascript'>"
	html = html & " alert('Borrado exitosamente');"
	html = html & "</script>"

	else

	html = html & "<script type='text/javascript'>"
	html = html & " alert('El estado no pudo ser borrado');"
	html = html & "</script>"

	end if

ElseIf opc = "prio" Then

	if EliminarRegistro(Cnn, opc, id, session("usuario")) then

	html = html & "<script type='text/javascript'>"
	html = html & " alert('Borrado exitosamente');"
	html = html & "</script>"

	else

	html = html & "<script type='text/javascript'>"
	html = html & " alert('La prioridad no pudo ser borrado');"
	html = html & "</script>"

	end if

ElseIf opc = "viap" Then

	if EliminarRegistro(Cnn, opc, id, session("usuario")) then

	html = html & "<script type='text/javascript'>"
	html = html & " alert('Borrado exitosamente');"
	html = html & "</script>"

	else

	html = html & "<script type='text/javascript'>"
	html = html & " alert('La via pedido no pudo ser borrado');"
	html = html & "</script>"

	end if

ElseIf opc = "apl" Then

	Exito = EliminarRegistro(Cnn, opc, id, session("usuario"))

ElseIf opc = "lic" Then

	Exito = EliminarRegistro(Cnn, opc, id, session("usuario"))

End If

vuelta = request.querystring("vuelta")

html = html & "<FORM name='init' method='post'  ACTION='" & vuelta & "?opc=listar&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "' >"
html = html & "<INPUT TYPE='SUBMIT' VALUE='INICIAR' style='display:none;' >"
html = html & "</FORM>"

html = html & "<script type='text/javascript'>"
html = html & " document.all.init.submit();"
html = html & "</script>"

html = html & "</body>"
html = html & "</html>"

response.write html
%>