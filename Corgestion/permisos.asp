<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/ui_componentes.inc" -->
<%
'[UI-ANT] ' ===== Cuerpo anterior completo: reemplazado por el bloque [UI] que sigue. Se borra entero buscando [UI-ANT] =====
'[UI-ANT] Response.Write headhtml()
'[UI-ANT] html = html & "<body>"

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

'[UI-ANT] xmenu = request.querystring("xmenu")
'[UI-ANT] xtmenu = request.querystring("xtmenu")

'[UI-ANT] 'XXX = SYSLOGON(cnn,"PARAM_APLIC")


'[UI-ANT] html = html & "<div>"

'[UI-ANT] html = html & "<header id='header' style='background:#0e1949;padding:1rem 0;'>"

'[UI-ANT] html = html & "<div >"

'[UI-ANT] html = html & "<table style='width:100%;'><tr style='display:flex;padding: 0 1rem;'>"

'[UI-ANT] html = html & "<td style='width:75%;' valign=center><a href='index.asp' class='logo' >"
'[UI-ANT] html = html & "<img src='images/iso_cor.png' style='width:6rem;' >"
'[UI-ANT] html = html & "<font style='font-size:15pt;color:white;'><b>Cor Gestion</b> | Sistema Administrativo</font></a></td>"
'[UI-ANT] html = html & "<td style='width:25%;display:flex;align-items:center;flex-direction:row-reverse;' align='right'>"
'[UI-ANT] html = html & "<nav>"
'[UI-ANT] html = html & "<ul>"
'[UI-ANT] 'html = html & "<li><a href='#menu'></a></li>"
'[UI-ANT] html = html & "</ul>"
'[UI-ANT] html = html & "</nav>"
'[UI-ANT] html = html & "</td>"

'[UI-ANT] html = html & "</tr></table>"

'[UI-ANT] html = html & "</div>"

'[UI-ANT] html = html & "</header>"

'[UI-ANT] html = html & "<table class='titulo-seccion' >"
'[UI-ANT] html = html & "<tr style='display:flex;align-items:center;justify-content:space-between;'>"

'[UI-ANT] html = html & "<td><h2>SEGURIDAD</h2></td>"

'[UI-ANT] html = html & "</tr>"

'[UI-ANT] html = html & "</table>"

'[UI-ANT] html = html & "</div>"

'[UI-ANT] 'html = html & "<div style='height:5%;width:100%;'>"

'[UI-ANT] 'html = html & "</div>"

'[UI-ANT] 'html = html & "<div style='height:57%;width:100%;float:left;overflow:auto;'>"
'[UI-ANT] html = html & "<div class='inner' style='padding: 2rem;'>"

'[UI-ANT] ' Fijate que en tu codigo original ya tenia una "x" adelante (xclass='tabl'), o sea que ya estaba anulada. La volamos.
'[UI-ANT] 'html = html & "<table xclass='tabl' style='width:30%;heigth:80%;font-family:Serif;text-align:center;' cellspacing=0 cellpadding=5 align=center border>"
'[UI-ANT] html = html & "<div class='bloque-datos' style='max-width: 600px; margin: 0 auto;'>"

'[UI-ANT] Set rs = Exec(Cnn, "sp_cg_Responsables @Modo='listar'")

'[UI-ANT] 'sImg = "<img src='images/llave.png' alt='Los Tejos' width='13' height='13' />"
'[UI-ANT] sImg = "<img src='images/llave.png' alt='Llave' width='13' height='13'>"

'[UI-ANT] Do While Not rs.eof

	'[UI-ANT] 'sLink = "<font style='font-size:15pt;'><b>" & rs(1) & "&nbsp;" & rs(2) & "&nbsp;</b></font><font style='font-size:12pt;'>(" & rs(3) & ")</font>"
	'[UI-ANT] sLink = "<span>" & rs(1) & " " & rs(2) & " (" & rs(3) & ")</span>"
	
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] html = html & "<div class='bloque-header'>"
	
	'[UI-ANT] ' Reemplazamos regis y el color fijo por nuestra clase moderna
	'[UI-ANT] 'html = html & "<td class='regis' style='background-color:#BCBDCA;' colspan=2>" & sLink & "</td>"
	'[UI-ANT] html = html & "<a colspan='2'>" & sLink & "</a>"
	
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "</div>"

	'[UI-ANT] ' Volamos la clase 'red' y los espacios manuales. Le aplicamos 'bot' para que los enlaces se vean como botones.
	'[UI-ANT] 'sLinkIni = "<a class='red' href='permisosDet.asp?idusr=" & Server.URLEncode(rs(4) & "") & "&tipousr=" & rs(3) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&tipomenu=INICIO'>INICIO&nbsp;&nbsp;&nbsp;" & sImg & "</a>"
	'[UI-ANT] sLinkIni = "<a class='col-link' style='display:flex;gap:1rem;align-items:center;' href='permisosDet.asp?idusr=" & Server.URLEncode(rs(4) & "") & "&tipousr=" & rs(3) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&tipomenu=INICIO'>" & sImg & "INICIO</a>"
	
	'[UI-ANT] 'sLinkPar = "<a class='red' href='permisosDet.asp?idusr=" & rs(0) & "&tipousr=" & rs(3) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&tipomenu=PARAM'>PARAMETROS&nbsp;&nbsp;&nbsp;" & sImg & "</a>"
	'[UI-ANT] sLinkPar = "<a class='col-link' style='display:flex;gap:1rem;align-items:center;' href='permisosDet.asp?idusr=" & rs(0) & "&tipousr=" & rs(3) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&tipomenu=PARAM'>" & sImg & "PARAMETROS</a>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] html = html & "<div class='bloque-body'>"
	'[UI-ANT] html = html & "<div style='display:flex;'>"
	
	'[UI-ANT] ' Limpiamos los td de la clase vieja regis y los estilos fijos
	'[UI-ANT] 'html = html & "<td class='regis' style='width:50%;' >" & sLinkIni & "</td>"
	'[UI-ANT] html = html & "<div style='width:50%'>" & sLinkIni & "</div>"
	
	'[UI-ANT] 'html = html & "<td class='regis' style='width:50%;' >" & sLinkPar & "</td>"
	'[UI-ANT] html = html & "<div style='width:50%'>" & sLinkPar & "</div>"
	
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "</div>"
	'[UI-ANT] html = html & "</div>"

'[UI-ANT] rs.movenext
'[UI-ANT] Loop

'[UI-ANT] 'html = html & "</table>"
'[UI-ANT] html = html & "</div>"

'[UI-ANT] 'html = html & "</div>"
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
' Seguridad: usuarios y acceso a sus permisos (menu principal y parametros).
' El cuerpo anterior completo queda arriba, comentado con '[UI-ANT].
' Consulta y enlaces: los mismos de antes.
' =========================================================================
Dim accTop, filas, cantUsr, funcion, estadoUsr

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

xmenu = request.querystring("xmenu")
xtmenu = request.querystring("xtmenu")

html = html & "<div id='wrapper' role='banner'>" & cabecera() & "</div>"

' [UI] Antes no tenia forma de volver
formVolver = "<form id='formVolver' class='form-oculto' method='post' action='index.asp'></form>"
accTop = UiBoton("Volver", "secondary", "document.getElementById('formVolver').submit();", True, "")

Set rs = Exec(Cnn, "sp_cg_Responsables @Modo='listar'")

filas = ""
cantUsr = 0
Do While Not rs.eof
	If rs(3) & "" = "TEC" Then
		funcion = "T&eacute;cnico"
	ElseIf rs(3) & "" = "FUNC" Then
		funcion = "Funcional"
	Else
		funcion = UiDato(rs(3))
	End If
	estadoUsr = ""
	If UCase(rs("activo") & "") = "N" Then estadoUsr = " " & UiBadge("Inactivo", "gray")

	' Los dos enlaces mandan el login (rs(4)): sp_cg_MenuUsuario filtra por idusr_seg.
	' [UI] Antes Parametros mandaba el id numerico (rs(0)) y sus permisos salian siempre vacios
	filas = filas & "<tr>" & _
		"<td><span class='ui-table__titulo'>" & UiTexto(rs(1) & " " & rs(2)) & "</span>" & estadoUsr & UiSub(UiTexto(rs(4))) & "</td>" & _
		"<td>" & funcion & "</td>" & _
		"<td><div class='ui-table__actions'>" & _
			"<a class='ui-button ui-button--secondary ui-button--sm' href='permisosDet.asp?idusr=" & Server.URLEncode(rs(4) & "") & "&tipousr=" & rs(3) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&tipomenu=INICIO' aria-label='Permisos del men&uacute; principal de " & UiTexto(rs(1) & " " & rs(2)) & "'>Men&uacute; principal</a>" & _
			"<a class='ui-button ui-button--secondary ui-button--sm' href='permisosDet.asp?idusr=" & Server.URLEncode(rs(4) & "") & "&tipousr=" & rs(3) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&tipomenu=PARAM' aria-label='Permisos de par&aacute;metros de " & UiTexto(rs(1) & " " & rs(2)) & "'>Par&aacute;metros</a>" & _
		"</div></td>" & _
		"</tr>"
	cantUsr = cantUsr + 1
	rs.movenext
Loop

html = html & UiPantallaSeguridadAbre("Usuarios y permisos", accTop, "")
html = html & UiCardAbre("Usuarios", "")
html = html & UiTablaAbre("Usuarios del sistema y acceso a sus permisos")
html = html & "<thead><tr>" & UiTh("Usuario", "") & UiTh("Funci&oacute;n", "") & UiTh("Permisos", "") & "</tr></thead><tbody>"
If cantUsr = 0 Then
	html = html & UiFilaVacia(3, "No hay usuarios cargados.")
Else
	html = html & filas
End If
html = html & UiTablaCierra()
If cantUsr = 1 Then
	html = html & UiPie("1 usuario", "")
ElseIf cantUsr > 1 Then
	html = html & UiPie(cantUsr & " usuarios", "")
End If
html = html & UiCardCierra()
html = html & UiPantallaCierra()

html = html & "<script src='assets/js/jquery.min.js'></script>"
html = html & "<script src='assets/js/skel.min.js'></script>"
html = html & "<script src='assets/js/util.js'></script>"
html = html & "<script src='assets/js/main.js'></script>"

html = html & formVolver & "</body>"
html = html & "</html>"

response.write html
%>