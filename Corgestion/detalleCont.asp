<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/datosDetalles.inc" -->
<!-- #INCLUDE FILE="IncFile/ui_componentes.inc" -->
<%
'[UI-ANT] ' ===== Cuerpo anterior completo: reemplazado por el bloque [UI] que sigue. Se borra entero buscando [UI-ANT] =====
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->

 '[UI-ANT] html = headhtml()

'[UI-ANT] html = html & "<body>"

'[UI-ANT] idcon = request.querystring("idcon")
'[UI-ANT] id = request.querystring("id")
'[UI-ANT] opc = request.querystring("opc")
'[UI-ANT] sol = request.querystring("sol")
'[UI-ANT] xmenu = request.querystring("xmenu")
'[UI-ANT] xtmenu = request.querystring("xtmenu")

'[UI-ANT] Set Cnn = Server.CreateObject("ADODB.Connection")
'[UI-ANT] On Error Resume Next
'[UI-ANT] Cnn.open scon1
'[UI-ANT] If Err.Number <> 0 Then
	'[UI-ANT] Response.Write "<div style='color:red;font-family:monospace;padding:10px;border:1px solid red;margin:10px;'>"
	'[UI-ANT] Response.Write "<b>ERROR DE CONEXION A BASE DE DATOS</b><br>"
	'[UI-ANT] Response.Write "Codigo: " & Err.Number & "<br>"
	'[UI-ANT] Response.Write "Detalle: " & Server.HTMLEncode(Err.Description) & "</div>"
	'[UI-ANT] Response.End
'[UI-ANT] End If
'[UI-ANT] On Error Goto 0

'[UI-ANT] html = html & "<div>"

'[UI-ANT] html = html & "<header id='header' style='height:11%'>"

'[UI-ANT] html = html & "<div>"

'[UI-ANT] html = html & "<table style='width:100%;'><tr>"

'[UI-ANT] html = html & "<td style='width:75%;' valign=center><a href='index.asp' class='logo' >"
'[UI-ANT] html = html & "<img src='images/logocorP4.gif' width='10%' >"
'[UI-ANT] html = html & "<font style='margin-left:100px;font-size:30pt;align:center;valign:center;font-family:Serif;color:#000000;'>Cor Consulting</font>"
'[UI-ANT] html = html & "</a></td>"

'[UI-ANT] html = html & "</tr></table>"

'[UI-ANT] html = html & "</div>"

'[UI-ANT] html = html & "</header>"

'[UI-ANT] html = html & "<div style='height:9%;width:100%;border-top:solid 5px #FF3D3D;border-bottom:double 6px #001BFF;'>"

'[UI-ANT] html = html & "<table style='width:100%;'><tr>"

'[UI-ANT] html = html & "<td><font style='font-size:20pt;font-family:Serif;color:#000000;letter-spacing:0.2em;'><b>DETALLE CONTACTO</b></font></td>"

'[UI-ANT] formx5 = "<form id='formx5' style='height:70%;' method='post' action='detalles.asp?sol=" & sol & "&id=" & id & "&opc=" & opc & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
'[UI-ANT] botonx5 = "<input class='bot' style='height:80%;font-size:13pt;' type='button' value='VOLVER' onclick=""document.getElementById('formx5').submit();""></form>"
'[UI-ANT] html = html & "<td align=center>" & formx5 & botonx5 & "</td>"

'[UI-ANT] html = html & "</tr></table>"

'[UI-ANT] html = html & "</div>"


'[UI-ANT] html = html & "<div style='height:71%;width:100%;'><br>"

'[UI-ANT] html = html & "<table class='tabl' style='height:30%;width:40%;' align=center cellspacing=0>"


	
'[UI-ANT] Set rs = ObtenerContacto(Cnn, idcon )


'[UI-ANT] if not rs.eof then
 
	'[UI-ANT] nombre = rs(1) & " " & rs(2)

	'[UI-ANT] html = html & "<tr>"
	'[UI-ANT] html = html & "<td xclass='encabe' style='font-size:30pt;letter-spacing:3px;' colspan=2 align=center>" & nombre & "</td>"
	'[UI-ANT] html = html & "</tr>"
	'[UI-ANT] html = html & "<tr>"
	'[UI-ANT] html = html & "<td xclass='encabe' style='width:15%;' align=center>TELEFONO:</td>"
	'[UI-ANT] html = html & "<td xstyle='width:25%;font-size:18pt;border-bottom:solid 3px #001BFF;color:#000000;' align=center>" & rs(3) & "</td>"
	'[UI-ANT] html = html & "</tr>"
	'[UI-ANT] html = html & "<td xclass='encabe' style='width:15%;' align=center>MAIL:</td>"
	'[UI-ANT] html = html & "<td xstyle='width:25%;font-size:18pt;border-bottom:solid 3px #001BFF;color:#000000;' align=center>" & rs(4) & "</td>"
	'[UI-ANT] html = html & "</tr>"
	'[UI-ANT] html = html & "<td xclass='encabe' style='width:15%;' align=center>CLIENTE:</td>"
	'[UI-ANT] html = html & "<td xstyle='width:25%;font-size:18pt;border-bottom:solid 3px #001BFF;color:#000000;' align=center>" & rs(6) & "</td>"
	'[UI-ANT] html = html & "</tr>"

	'[UI-ANT] html = html & "</table>"

'[UI-ANT] end if

'[UI-ANT] html = html & "</div>"


'[UI-ANT] html = html & "</div>"

'[UI-ANT] html = html & "<script src='assets/js/jquery.min.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/skel.min.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/util.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/main.js'></script>"

'[UI-ANT] html = html & "</body>"
'[UI-ANT] html = html & "</html>"

'[UI-ANT] response.write html

' =========================================================================
' [UI] PANTALLA con el sistema visual (IncFile/ui_componentes.inc)
' Ficha de un contacto del cliente (se llega desde el detalle de un incidente o proyecto).
' El cuerpo anterior completo queda arriba, comentado con '[UI-ANT].
' Cambios: cabecera comun (antes una propia, con logo GIF), datos escapados, telefono y mail
' clicables, e id validado.
' =========================================================================
Dim accTop, nombreCon, telCon, mailCon, cliCon, telLimpio

html = headhtml()
html = html & "<body>"

idcon = request.querystring("idcon")
id = request.querystring("id")
opc = request.querystring("opc")
sol = request.querystring("sol")
xmenu = request.querystring("xmenu")
xtmenu = request.querystring("xtmenu")

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

html = html & "<div id='wrapper' role='banner'>" & cabecera() & "</div>"

' Volver al detalle del caso (el mismo destino de antes)
formx5 = "<form id='formx5' class='form-oculto' method='post' action='detalles.asp?sol=" & Server.URLEncode(sol & "") & "&id=" & Server.URLEncode(id & "") & "&opc=" & Server.URLEncode(opc & "") & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
accTop = UiBoton("Volver", "secondary", "document.getElementById('formx5').submit();", True, "")

' [UI] El id se valida como numero antes de consultar
If IsNumeric(idcon & "") And Trim(idcon & "") <> "" Then
	Set rs = ObtenerContacto(Cnn, CLng(idcon))
Else
	Set rs = Nothing
End If

If rs Is Nothing Then
	html = html & UiPantallaAbre("Contacto", accTop, "narrow")
	html = html & UiCardAbre("Contacto", "") & "<div class='ui-empty'>El enlace no indica qu&eacute; contacto mostrar.</div>" & UiCardCierra()
ElseIf rs.eof Then
	html = html & UiPantallaAbre("Contacto", accTop, "narrow")
	html = html & UiCardAbre("Contacto", "") & "<div class='ui-empty'>No se encontr&oacute; el contacto.</div>" & UiCardCierra()
Else
	nombreCon = Trim(rs(1) & " " & rs(2))
	telCon = Trim(rs(3) & "")
	mailCon = Trim(rs(4) & "")
	cliCon = Trim(rs(6) & "")

	html = html & UiPantallaAbre(UiTexto(nombreCon), accTop, "narrow")
	html = html & UiCardAbre("Datos de contacto", "") & "<div class='ui-fields'>"

	' Telefono y mail clicables: llamar o escribir directamente
	If telCon = "" Then
		html = html & UiCampoLectura("Tel&eacute;fono", UiDato(""), False)
	Else
		telLimpio = ConTelefono(telCon)
		html = html & UiCampoLectura("Tel&eacute;fono", "<a class='ui-link' href='tel:" & telLimpio & "'>" & UiTexto(telCon) & "</a>", False)
	End If
	If mailCon = "" Then
		html = html & UiCampoLectura("Mail", UiDato(""), False)
	Else
		html = html & UiCampoLectura("Mail", "<a class='ui-link' href='mailto:" & Server.HTMLEncode(mailCon) & "'>" & UiTexto(mailCon) & "</a>", False)
	End If
	html = html & UiCampoLectura("Cliente", UiDato(cliCon), False)
	html = html & "</div>" & UiCardCierra()
End If

html = html & UiPantallaCierra()

html = html & "<script src='assets/js/jquery.min.js'></script>"
html = html & "<script src='assets/js/skel.min.js'></script>"
html = html & "<script src='assets/js/util.js'></script>"
html = html & "<script src='assets/js/main.js'></script>"

html = html & formx5 & "</body>"
html = html & "</html>"

response.write html


' Telefono para el enlace tel: solo digitos y un + inicial
Function ConTelefono(t)
	Dim s, i, c
	s = ""
	For i = 1 To Len(t)
		c = Mid(t, i, 1)
		If (c >= "0" And c <= "9") Or (c = "+" And s = "") Then s = s & c
	Next
	ConTelefono = s
End Function
%>