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

XXX = SYSLOGON(cnn,"MENU_ESTADISTICAS")

If request.form("sel") & "" = "" Then
	sel = request.querystring("sel")
Else
	sel = request.form("sel")
End If

If sel & "" = "" Then sel = 0

xmenu = request.querystring("xmenu")
xtmenu = request.querystring("xtmenu")

'[UI-ANT] ' ===== Bloque anterior: reemplazado por el bloque [UI] que sigue. Se borra entero buscando [UI-ANT] =====
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

'[UI-ANT] 'html = html & "<div style='height:9%;width:100%;border-top:solid 5px #FF3D3D;border-bottom:double 6px #001BFF;'>"

'[UI-ANT] html = html & "<table class='titulo-seccion' >"
'[UI-ANT] html = html & "<tr style='display:flex;align-items:center;justify-content:space-between;'>"

'[UI-ANT] html = html & "<td><h2>ESTADISTICAS</h2></td>"

'[UI-ANT] html = html & "<td>"

'[UI-ANT] html = html & "<div>"

'[UI-ANT] html = html & "<form id='sform' style='height:100%;width:100%;' method='post' action='estadisticas.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "' >"

'[UI-ANT] html = html & "<select id='sel' name='sel' onchange=""document.getElementById('sform').submit();"" >"

'[UI-ANT] If sel = 0 Then
	'[UI-ANT] html = html & "<option value='0' selected >INCIDENTES</option>"
'[UI-ANT] Else
	'[UI-ANT] html = html & "<option value='0' >INCIDENTES</option>"
'[UI-ANT] End If

'[UI-ANT] If sel = 1 Then
	'[UI-ANT] html = html & "<option value='1' selected >PROYECTOS</option>"
'[UI-ANT] Else
	'[UI-ANT] html = html & "<option value='1'>PROYECTOS</option>"
'[UI-ANT] End If

'[UI-ANT] If sel = 2 Then
	'[UI-ANT] html = html & "<option value='2' selected >APL</option>"
'[UI-ANT] Else
	'[UI-ANT] html = html & "<option value='2'>APL</option>"
'[UI-ANT] End If

'[UI-ANT] If sel = 3 Then
	'[UI-ANT] html = html & "<option value='3' selected >USUARIOS</option>"
'[UI-ANT] Else
	'[UI-ANT] html = html & "<option value='3'>USUARIOS</option>"
'[UI-ANT] End If

'[UI-ANT] html = html & "</select>"

'[UI-ANT] html = html & "</form>"

'[UI-ANT] html = html & "</div>"

'[UI-ANT] html = html & "</td>"

'[UI-ANT] formx04 = "<form id='formx04' method='post' action='index.asp'></form>"
'[UI-ANT] botonx04 = "<input class='bot' type='button' value='VOLVER' onclick=""document.getElementById('formx04').submit();"">"
'[UI-ANT] html = html & "<td>" & botonx04 & "</td>"

'[UI-ANT] html = html & "</tr>"

'[UI-ANT] html = html & "</table>"

'[UI-ANT] html = html & "</div>"

'[UI-ANT] 'DIV QUE SE RELLENA DEPENDE LO QUE QUIERA CONSULTAR ELEGIDO EN EL COMBO

'[UI-ANT] If sel = 0 Then 'INCIDENTES

	'[UI-ANT] ' --- INICIO NORMALIZACION: FORMULARIO ESTADISTICAS (INCIDENTES) ---
	'[UI-ANT] 'html = html & "<div style='height:75%;width:100%;'><br>"
	'[UI-ANT] html = html & "<div style='max-width: 900px; margin: 2rem auto;'>"

	'[UI-ANT] html = html & "<form id='xForm' method='post' action='estadisticasDet.asp?tipo=I&sel=" & sel & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"

	'[UI-ANT] 'html = html & "<table class='tabl' align=center cellspacing=0 cellpadding=3>"
	'[UI-ANT] html = html & "<div class='bloque-datos'>"

	'[UI-ANT] 'TITULO
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' style='font-size:12pt;' align=center valign=center colspan=4><b><u>INCIDENTES</u></b></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='bloque-header'>INCIDENTES</div>"

	'[UI-ANT] html = html & "<div class='bloque-body'>"

	'[UI-ANT] 'ESPACIO
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td colspan=4>&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"

	'[UI-ANT] ' --- INICIO GRILLA ---
	'[UI-ANT] html = html & "<div class='grid-datos-doble'>"

	'[UI-ANT] 'ID
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td >&nbsp;</td>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center >ID</td>"
	'[UI-ANT] 'html = html & "<td class='fil' >"
	'[UI-ANT] 'html = html & "<input class='inpu' id='id' name='id' type='text'>"
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "<td>&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>ID</div>"
	'[UI-ANT] html = html & "<div><input type='text' id='id' name='id' class='input-moderno'></div>"

	'[UI-ANT] 'CLIENTE
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td >&nbsp;</td>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center >CLIENTE</td>"
	'[UI-ANT] 'html = html & "<td class='fil' valign=center align=center>"
	'[UI-ANT] 'html = html & armaCombo("select id,descripcion from clientes",cnn,request.form("cli"),0,1,"cli")
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "<td>&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>CLIENTE</div>"
	'[UI-ANT] html = html & "<div class='dato-value'>"
	'[UI-ANT] html = html & armaCombo("sp_cg_Catalogo_Listar @Catalogo='clientes'",cnn,request.form("cli"),0,1,"cli")
	'[UI-ANT] html = html & "</select></div>" ' <--- PARCHE: Forzamos el cierre del combo

	'[UI-ANT] 'ULT USUARIO QUE ACCEDIO
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td >&nbsp;</td>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center >ULTIMO ACC USUARIO</td>"
	'[UI-ANT] 'html = html & "<td class='fil' valign=center align=center>"
	'[UI-ANT] 'html = html & armaCombo("select id,nombre + ' ' + apellido from CG_COR_RESPONSABLES",Cnn,request.form("uau"),0,1,"uau")
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "<td>&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>ULTIMO ACC USUARIO</div>"
	'[UI-ANT] html = html & "<div class='dato-value'>"
	'[UI-ANT] html = html & armaCombo("sp_cg_Responsables @Modo='combo_todos'",Cnn,request.form("uau"),0,1,"uau")
	'[UI-ANT] html = html & "</select></div>" ' <--- PARCHE: Forzamos el cierre del combo

	'[UI-ANT] 'ULT TIPO ACCESO
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td >&nbsp;</td>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center >ULTIMO TIPO ACC</td>"
	'[UI-ANT] 'html = html & "<td class='fil' valign=center align=center>"
	'[UI-ANT] 'html = html & "<select id='uta' name='uta' class='selec' >"
	'[UI-ANT] 'html = html & "<option value=''>S/D</option>"
	'[UI-ANT] 'html = html & "<option value='CONSULTA'>CONSULTA</option>"
	'[UI-ANT] 'html = html & "<option value='ACTUALIZA'>ACTUALIZA</option>"
	'[UI-ANT] 'html = html & "</select>"
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "<td>&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>ULTIMO TIPO ACC</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><select id='uta' name='uta'>"
	'[UI-ANT] html = html & "<option value=''>S/D</option>"
	'[UI-ANT] html = html & "<option value='CONSULTA'>CONSULTA</option>"
	'[UI-ANT] html = html & "<option value='ACTUALIZA'>ACTUALIZA</option>"
	'[UI-ANT] html = html & "</select></div>"

	'[UI-ANT] 'FECHA CREACION
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' valign=center>FEC CREAC DESDE</td>"
	'[UI-ANT] 'html = html & "<td class='fil' valign=center align=left>"
	'[UI-ANT] 'html = html & "<input id='fcd' name='fcd' class='inpu' type='text' />"
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "<td class='fil' >FEC CREAC HASTA</td>"
	'[UI-ANT] 'html = html & "<td class='fil' valign=center>"
	'[UI-ANT] 'html = html & "<input id='fch' name='fch' class='inpu' type='text' />"
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>FEC CREAC DESDE</div>"
	'[UI-ANT] html = html & "<div><input type='text' id='fcd' name='fcd' class='input-moderno'></div>"
	'[UI-ANT] html = html & "<div class='dato-label'>FEC CREAC HASTA</div>"
	'[UI-ANT] html = html & "<div><input type='text' id='fch' name='fch' class='input-moderno'></div>"

	'[UI-ANT] 'FECHA ULT ACC
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' valign=center>FEC ULT ACC DESDE</td>"
	'[UI-ANT] 'html = html & "<td class='fil' valign=center align=left>"
	'[UI-ANT] 'html = html & "<input id='fad' name='fad' class='inpu' type='text' />"
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "<td class='fil' valign=center>FEC ULT ACC HASTA</td>"
	'[UI-ANT] 'html = html & "<td class='fil' valign=center>"
	'[UI-ANT] 'html = html & "<input id='fah' name='fah' class='inpu' type='text' />"
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>FEC ULT ACC DESDE</div>"
	'[UI-ANT] html = html & "<div><input type='text' id='fad' name='fad' class='input-moderno'></div>"
	'[UI-ANT] html = html & "<div class='dato-label'>FEC ULT ACC HASTA</div>"
	'[UI-ANT] html = html & "<div><input type='text' id='fah' name='fah' class='input-moderno'></div>"

	'[UI-ANT] ' --- FIN GRILLA ---
	'[UI-ANT] html = html & "</div>"

	'[UI-ANT] 'ESPACIO
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td colspan=4>"
	'[UI-ANT] 'html = html & "&nbsp;"
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"

	'[UI-ANT] 'BOTON CONSULTAR
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' colspan=4 align=center>"
	'[UI-ANT] html = html & "<hr style='border: 0; border-top: 1px solid #e2e8f0; margin: 1.5rem 0;'>"
	'[UI-ANT] html = html & "<div style='text-align: center; margin-bottom: 0.5rem;'>"

	'[UI-ANT] if getPermisos(xmenu,xtmenu,"CONS") = 1 then
		'[UI-ANT] 'html = html & "<input type='button' class='bot' value='Consultar' onclick=""document.getElementById('xForm').submit();"">"
		'[UI-ANT] html = html & "<input type='button' class='btn-principal' value='Consultar' onclick=""document.getElementById('xForm').submit();"">"
	'[UI-ANT] else
		'[UI-ANT] 'html = html & "<input type='button' class='botsp' value='Consultar' onclick=""alert('Sin permiso para consultar.');"">"
		'[UI-ANT] html = html & "<input type='button' class='btn-principal btn-accion-deshabilitado' value='Consultar' onclick=""alert('Sin permiso para consultar.');"">"
	'[UI-ANT] end if

	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "</div>"

	'[UI-ANT] 'html = html & "</table>"
	'[UI-ANT] html = html & "</div>" ' Fin .bloque-body
	'[UI-ANT] html = html & "</div>" ' Fin .bloque-datos

	'[UI-ANT] html = html & "</form>"

	'[UI-ANT] 'html = html & "</div>"
	'[UI-ANT] html = html & "</div>" ' Fin max-width

'[UI-ANT] ElseIf sel = 1 Then 'PROYECTOS

	'[UI-ANT] ' --- INICIO NORMALIZACION: FORMULARIO ESTADISTICAS (PROYECTOS) ---
	'[UI-ANT] 'html = html & "<div style='height:75%;width:100%;'><br>"
	'[UI-ANT] html = html & "<div style='max-width: 900px; margin: 2rem auto;'>"

	'[UI-ANT] html = html & "<form id='xForm' method='post' action='estadisticasDet.asp?tipo=P&sel=" & sel & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"

	'[UI-ANT] 'html = html & "<table class='tabl' align=center cellspacing=0 cellpadding=3>"
	'[UI-ANT] html = html & "<div class='bloque-datos'>"

	'[UI-ANT] 'TITULO
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' style='font-size:12pt;' align=center valign=center colspan=4><b><u>PROYECTOS</u></b></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='bloque-header'>PROYECTOS</div>"

	'[UI-ANT] html = html & "<div class='bloque-body'>"

	'[UI-ANT] 'ESPACIO
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td colspan=4>&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"

	'[UI-ANT] ' --- INICIO GRILLA ---
	'[UI-ANT] html = html & "<div class='grid-datos-doble'>"

	'[UI-ANT] 'ID
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td >&nbsp;</td>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center >ID</td>"
	'[UI-ANT] 'html = html & "<td class='fil'>"
	'[UI-ANT] 'html = html & "<input id='id' name='id' type='text' class='inpu'>"
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "<td>&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>ID</div>"
	'[UI-ANT] html = html & "<div><input type='text' id='id' name='id' class='input-moderno'></div>"

	'[UI-ANT] 'CLIENTE
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td >&nbsp;</td>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center >CLIENTE</td>"
	'[UI-ANT] 'html = html & "<td class='fil' valign=center align=center>"
	'[UI-ANT] 'html = html & armaCombo("select id,descripcion from clientes",cnn,request.form("cli"),0,1,"cli")
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "<td>&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>CLIENTE</div>"
	'[UI-ANT] html = html & "<div class='dato-value'>"
	'[UI-ANT] html = html & armaCombo("sp_cg_Catalogo_Listar @Catalogo='clientes'",cnn,request.form("cli"),0,1,"cli")
	'[UI-ANT] html = html & "</select></div>" ' <--- PARCHE: Forzamos el cierre del combo

	'[UI-ANT] 'ULT USUARIO QUE ACCEDIO
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td >&nbsp;</td>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center >ULTIMO ACC USUARIO</td>"
	'[UI-ANT] 'html = html & "<td class='fil' valign=center align=center>"
	'[UI-ANT] 'html = html & armaCombo("select id,nombre + ' ' + apellido from CG_COR_RESPONSABLES",Cnn,request.form("uau"),0,1,"uau")
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "<td>&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>ULTIMO ACC USUARIO</div>"
	'[UI-ANT] html = html & "<div class='dato-value'>"
	'[UI-ANT] html = html & armaCombo("sp_cg_Responsables @Modo='combo_todos'",Cnn,request.form("uau"),0,1,"uau")
	'[UI-ANT] html = html & "</select></div>" ' <--- PARCHE: Forzamos el cierre del combo

	'[UI-ANT] 'ULT TIPO ACCESO
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td >&nbsp;</td>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center >ULTIMO TIPO ACC</td>"
	'[UI-ANT] 'html = html & "<td class='fil' valign=center align=center>"
	'[UI-ANT] 'html = html & "<select id='uta' name='uta' class='selec'>"
	'[UI-ANT] 'html = html & "<option value=''>S/D</option>"
	'[UI-ANT] 'html = html & "<option value='CONSULTA'>CONSULTA</option>"
	'[UI-ANT] 'html = html & "<option value='ACTUALIZA'>ACTUALIZA</option>"
	'[UI-ANT] 'html = html & "</select>"
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "<td>&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>ULTIMO TIPO ACC</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><select id='uta' name='uta'>"
	'[UI-ANT] html = html & "<option value=''>S/D</option>"
	'[UI-ANT] html = html & "<option value='CONSULTA'>CONSULTA</option>"
	'[UI-ANT] html = html & "<option value='ACTUALIZA'>ACTUALIZA</option>"
	'[UI-ANT] html = html & "</select></div>"

	'[UI-ANT] 'FECHA CREACION
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' align=right valign=center>FEC CREAC DESDE</td>"
	'[UI-ANT] 'html = html & "<td class='fil' valign=center align=left>"
	'[UI-ANT] 'html = html & "<input id='fcd' name='fcd' class='inpu' type='text' />"
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "<td class='fil' align=right valign=center>FEC CREAC HASTA</td>"
	'[UI-ANT] 'html = html & "<td class='fil' valign=center>"
	'[UI-ANT] 'html = html & "<input id='fch' name='fch' class='inpu' type='text' />"
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>FEC CREAC DESDE</div>"
	'[UI-ANT] html = html & "<div><input type='text' id='fcd' name='fcd' class='input-moderno'></div>"
	'[UI-ANT] html = html & "<div class='dato-label'>FEC CREAC HASTA</div>"
	'[UI-ANT] html = html & "<div><input type='text' id='fch' name='fch' class='input-moderno'></div>"

	'[UI-ANT] 'FECHA ULT ACC
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' align=right valign=center>FEC ULT ACC DESDE</td>"
	'[UI-ANT] 'html = html & "<td class='fil' valign=center align=left>"
	'[UI-ANT] 'html = html & "<input id='fad' name='fad' class='inpu' type='text' />"
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "<td class='fil' align=right valign=center>FEC ULT ACC HASTA</td>"
	'[UI-ANT] 'html = html & "<td class='fil' valign=center>"
	'[UI-ANT] 'html = html & "<input id='fah' name='fah' class='inpu' type='text' />"
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>FEC ULT ACC DESDE</div>"
	'[UI-ANT] html = html & "<div><input type='text' id='fad' name='fad' class='input-moderno'></div>"
	'[UI-ANT] html = html & "<div class='dato-label'>FEC ULT ACC HASTA</div>"
	'[UI-ANT] html = html & "<div><input type='text' id='fah' name='fah' class='input-moderno'></div>"

	'[UI-ANT] ' --- FIN GRILLA ---
	'[UI-ANT] html = html & "</div>"

	'[UI-ANT] 'ESPACIO
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td colspan=4>&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"

	'[UI-ANT] 'BOTON CONSULTAR
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' colspan=4 align=center>"
	'[UI-ANT] html = html & "<hr style='border: 0; border-top: 1px solid #e2e8f0; margin: 1.5rem 0;'>"
	'[UI-ANT] html = html & "<div style='text-align: center; margin-bottom: 0.5rem;'>"

	'[UI-ANT] if getPermisos(xmenu,xtmenu,"CONS") = 1 then
		'[UI-ANT] 'html = html & "<input type='button' class='bot' value='Consultar' onclick=""document.getElementById('xForm').submit();"">"
		'[UI-ANT] html = html & "<input type='button' class='btn-principal' value='Consultar' onclick=""document.getElementById('xForm').submit();"">"
	'[UI-ANT] else
		'[UI-ANT] 'html = html & "<input type='button' class='botsp' value='Consultar' onclick=""alert('Sin permiso para consultar.');"">"
		'[UI-ANT] html = html & "<input type='button' class='btn-principal btn-accion-deshabilitado' value='Consultar' onclick=""alert('Sin permiso para consultar.');"">"
	'[UI-ANT] end if

	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "</div>"

	'[UI-ANT] 'html = html & "</table>"
	'[UI-ANT] html = html & "</div>" ' Fin .bloque-body
	'[UI-ANT] html = html & "</div>" ' Fin .bloque-datos

	'[UI-ANT] html = html & "</form>"

	'[UI-ANT] 'html = html & "</div>"
	'[UI-ANT] html = html & "</div>" ' Fin max-width

'[UI-ANT] ElseIf sel = 2 Then 'APL

	'[UI-ANT] ' --- INICIO NORMALIZACION: FORMULARIO ESTADISTICAS (APLs) ---
	'[UI-ANT] 'html = html & "<div style='height:75%;width:100%;'><br>"
	'[UI-ANT] html = html & "<div style='max-width: 900px; margin: 2rem auto;'>"

	'[UI-ANT] html = html & "<form id='xForm' method='post' action='estadisticasDet.asp?tipo=A&sel=" & sel & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"

	'[UI-ANT] 'html = html & "<table class='tabl' align=center cellspacing=0 cellpadding=3>"
	'[UI-ANT] html = html & "<div class='bloque-datos'>"

	'[UI-ANT] 'TITULO
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' style='font-size:12pt;' align=center valign=center colspan=4><b><u>APLs</u></b></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='bloque-header'>APLs</div>"

	'[UI-ANT] html = html & "<div class='bloque-body'>"

	'[UI-ANT] 'ESPACIO
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td colspan=4>&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"

	'[UI-ANT] ' --- INICIO GRILLA ---
	'[UI-ANT] html = html & "<div class='grid-datos-doble'>"

	'[UI-ANT] 'ID
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td >&nbsp;</td>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center >ID</td>"
	'[UI-ANT] 'html = html & "<td class='fil' >"
	'[UI-ANT] 'html = html & "<input id='id' name='id' type='text' class='inpu'>"
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "<td>&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>ID</div>"
	'[UI-ANT] html = html & "<div><input type='text' id='id' name='id' class='input-moderno'></div>"

	'[UI-ANT] 'CLIENTE
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td >&nbsp;</td>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center >CLIENTE</td>"
	'[UI-ANT] 'html = html & "<td class='fil' valign=center align=center>"
	'[UI-ANT] 'html = html & armaCombo("select id,descripcion from clientes",cnn,request.form("cli"),0,1,"cli")
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "<td>&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>CLIENTE</div>"
	'[UI-ANT] html = html & "<div class='dato-value'>"
	'[UI-ANT] html = html & armaCombo("sp_cg_Catalogo_Listar @Catalogo='clientes'",cnn,request.form("cli"),0,1,"cli")
	'[UI-ANT] html = html & "</select></div>" ' <--- PARCHE: Forzamos el cierre del combo

	'[UI-ANT] 'ULT USUARIO QUE ACCEDIO
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td >&nbsp;</td>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center >ULTIMO ACC USUARIO</td>"
	'[UI-ANT] 'html = html & "<td class='fil' valign=center align=center>"
	'[UI-ANT] 'html = html & armaCombo("select id,nombre + ' ' + apellido from CG_COR_RESPONSABLES",Cnn,request.form("uau"),0,1,"uau")
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "<td>&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>ULTIMO ACC USUARIO</div>"
	'[UI-ANT] html = html & "<div class='dato-value'>"
	'[UI-ANT] html = html & armaCombo("sp_cg_Responsables @Modo='combo_todos'",Cnn,request.form("uau"),0,1,"uau")
	'[UI-ANT] html = html & "</select></div>" ' <--- PARCHE: Forzamos el cierre del combo

	'[UI-ANT] 'ULT TIPO ACCESO
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td >&nbsp;</td>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center >ULTIMO TIPO ACC</td>"
	'[UI-ANT] 'html = html & "<td class='fil' valign=center align=center>"
	'[UI-ANT] 'html = html & "<select id='uta' name='uta' class='selec'>"
	'[UI-ANT] 'html = html & "<option value=''>S/D</option>"
	'[UI-ANT] 'html = html & "<option value='CONSULTA'>CONSULTA</option>"
	'[UI-ANT] 'html = html & "<option value='ACTUALIZA'>ACTUALIZA</option>"
	'[UI-ANT] 'html = html & "</select>"
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "<td>&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>ULTIMO TIPO ACC</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><select id='uta' name='uta'>"
	'[UI-ANT] html = html & "<option value=''>S/D</option>"
	'[UI-ANT] html = html & "<option value='CONSULTA'>CONSULTA</option>"
	'[UI-ANT] html = html & "<option value='ACTUALIZA'>ACTUALIZA</option>"
	'[UI-ANT] html = html & "</select></div>"

	'[UI-ANT] 'FECHA CREACION
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' align=right valign=center>FEC CREAC DESDE</td>"
	'[UI-ANT] 'html = html & "<td class='fil' valign=center align=left>"
	'[UI-ANT] 'html = html & "<input id='fcd' name='fcd' class='inpu' type='text' />"
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "<td class='fil' align=right valign=center>FEC CREAC HASTA</td>"
	'[UI-ANT] 'html = html & "<td class='fil' valign=center>"
	'[UI-ANT] 'html = html & "<input id='fch' name='fch' class='inpu' type='text' />"
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>FEC CREAC DESDE</div>"
	'[UI-ANT] html = html & "<div><input type='text' id='fcd' name='fcd' class='input-moderno'></div>"
	'[UI-ANT] html = html & "<div class='dato-label'>FEC CREAC HASTA</div>"
	'[UI-ANT] html = html & "<div><input type='text' id='fch' name='fch' class='input-moderno'></div>"

	'[UI-ANT] 'FECHA ULT ACC
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' align=right valign=center>FEC ULT ACC DESDE</td>"
	'[UI-ANT] 'html = html & "<td class='fil' valign=center align=left>"
	'[UI-ANT] 'html = html & "<input id='fad' name='fad' class='inpu' type='text' />"
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "<td class='fil' align=right valign=center>FEC ULT ACC HASTA</td>"
	'[UI-ANT] 'html = html & "<td class='fil' valign=center>"
	'[UI-ANT] 'html = html & "<input id='fah' name='fah' class='inpu' type='text' />"
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>FEC ULT ACC DESDE</div>"
	'[UI-ANT] html = html & "<div><input type='text' id='fad' name='fad' class='input-moderno'></div>"
	'[UI-ANT] html = html & "<div class='dato-label'>FEC ULT ACC HASTA</div>"
	'[UI-ANT] html = html & "<div><input type='text' id='fah' name='fah' class='input-moderno'></div>"

	'[UI-ANT] ' --- FIN GRILLA ---
	'[UI-ANT] html = html & "</div>"

	'[UI-ANT] 'ESPACIO
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td colspan=4>&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"

	'[UI-ANT] 'BOTON CONSULTAR
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td colspan=4 align=center>"
	'[UI-ANT] html = html & "<hr style='border: 0; border-top: 1px solid #e2e8f0; margin: 1.5rem 0;'>"
	'[UI-ANT] html = html & "<div style='text-align: center; margin-bottom: 0.5rem;'>"

	'[UI-ANT] if getPermisos(xmenu,xtmenu,"CONS") = 1 then
		'[UI-ANT] 'html = html & "<input type='button' class='bot' value='Consultar' onclick=""document.getElementById('xForm').submit();"">"
		'[UI-ANT] html = html & "<input type='button' class='btn-principal' value='Consultar' onclick=""document.getElementById('xForm').submit();"">"
	'[UI-ANT] else
		'[UI-ANT] 'html = html & "<input type='button' class='botsp' value='Consultar' onclick=""alert('Sin permiso para consultar.');"">"
		'[UI-ANT] html = html & "<input type='button' class='btn-principal btn-accion-deshabilitado' value='Consultar' onclick=""alert('Sin permiso para consultar.');"">"
	'[UI-ANT] end if

	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "</div>"

	'[UI-ANT] 'html = html & "</table>"
	'[UI-ANT] html = html & "</div>" ' Fin .bloque-body
	'[UI-ANT] html = html & "</div>" ' Fin .bloque-datos

	'[UI-ANT] html = html & "</form>"

	'[UI-ANT] 'html = html & "</div>"
	'[UI-ANT] html = html & "</div>" ' Fin max-width

'[UI-ANT] ElseIf sel = 3 Then 'USUARIOS

	'[UI-ANT] ' --- INICIO NORMALIZACION: FORMULARIO ESTADISTICAS (USUARIOS) ---
	'[UI-ANT] 'html = html & "<div style='height:75%;width:100%;'><br>"
	'[UI-ANT] html = html & "<div style='max-width: 900px; margin: 2rem auto;'>"

	'[UI-ANT] html = html & "<form id='xForm' method='post' action='estadisticasDet.asp?tipo=U&sel=" & sel & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"

	'[UI-ANT] 'html = html & "<table class='tabl' style='width:40%;' align=center cellspacing=0 cellpadding=3>"
	'[UI-ANT] html = html & "<div class='bloque-datos'>"

	'[UI-ANT] 'TITULO
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' style='font-size:12pt;' align=center valign=center colspan=4><b><u>USUARIOS</u></b></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='bloque-header'>USUARIOS</div>"
	
	'[UI-ANT] html = html & "<div class='bloque-body'>"

	'[UI-ANT] 'ESPACIO
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td colspan=4>&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"

	'[UI-ANT] ' --- INICIO GRILLA ---
	'[UI-ANT] html = html & "<div class='grid-datos-doble'>"

	'[UI-ANT] 'USUARIOS
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td >&nbsp;</td>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center >ULTIMO ACC USUARIO</td>"
	'[UI-ANT] 'html = html & "<td class='fil' valign=center align=center>"
	'[UI-ANT] 'html = html & armaCombo("select id,nombre + ' ' + apellido from CG_COR_RESPONSABLES",Cnn,request.form("uau"),0,1,"uau")
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "<td>&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>ULTIMO ACC USUARIO</div>"
	'[UI-ANT] html = html & "<div class='dato-value'>"
	'[UI-ANT] html = html & armaCombo("sp_cg_Responsables @Modo='combo_todos'",Cnn,request.form("uau"),0,1,"uau")
	'[UI-ANT] html = html & "</select></div>" ' <--- PARCHE: Forzamos el cierre del combo

	'[UI-ANT] ' --- FIN GRILLA ---
	'[UI-ANT] html = html & "</div>"

	'[UI-ANT] 'ESPACIO
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td colspan=4>&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"

	'[UI-ANT] 'BOTON CONSULTAR
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' colspan=4 align=center>"
	'[UI-ANT] html = html & "<hr style='border: 0; border-top: 1px solid #e2e8f0; margin: 1.5rem 0;'>"
	'[UI-ANT] html = html & "<div style='text-align: center; margin-bottom: 0.5rem;'>"

	'[UI-ANT] if getPermisos(xmenu,xtmenu,"CONS") = 1 then
		'[UI-ANT] 'html = html & "<input type='button' class='bot' value='Consultar' onclick=""document.getElementById('xForm').submit();"">"
		'[UI-ANT] html = html & "<input type='button' class='btn-principal' value='Consultar' onclick=""document.getElementById('xForm').submit();"">"
	'[UI-ANT] else
		'[UI-ANT] 'html = html & "<input type='button' class='botsp' value='Consultar' onclick=""alert('Sin permiso para consultar.');"">"
		'[UI-ANT] html = html & "<input type='button' class='btn-principal btn-accion-deshabilitado' value='Consultar' onclick=""alert('Sin permiso para consultar.');"">"
	'[UI-ANT] end if

	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "</div>"

	'[UI-ANT] 'html = html & "</table>"
	'[UI-ANT] html = html & "</div>" ' Fin .bloque-body
	'[UI-ANT] html = html & "</div>" ' Fin .bloque-datos

	'[UI-ANT] html = html & "</form>"

	'[UI-ANT] 'html = html & "</div>"
	'[UI-ANT] html = html & "</div>" ' Fin max-width

'[UI-ANT] End If

'[UI-ANT] html = html & "<script type='text/javascript'>"
'[UI-ANT] html = html & " function modificar(formu){ "
'[UI-ANT] html = html & " document.getElementById(formu).submit();"
'[UI-ANT] html = html & " alert('Descripcion actualizada');} "
'[UI-ANT] html = html & "</script>"

'[UI-ANT] html = html & "<script src='assets/js/jquery.min.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/skel.min.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/util.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/main.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/jquery-ui.js'></script>"

'[UI-ANT] html = html & "<script>"
'[UI-ANT] html = html & "  $( function() {"

'[UI-ANT] html = html & "$.datepicker.regional['es'] = {"
'[UI-ANT] html = html & "closeText: 'Cerrar',"
'[UI-ANT] html = html & "prevText: '< Ant',"
'[UI-ANT] html = html & "nextText: 'Sig >',"
'[UI-ANT] html = html & "currentText: 'Hoy',"
'[UI-ANT] html = html & "monthNames: ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'],"
'[UI-ANT] html = html & "monthNamesShort: ['Ene','Feb','Mar','Abr', 'May','Jun','Jul','Ago','Sep', 'Oct','Nov','Dic'],"
'[UI-ANT] html = html & "dayNames: ['Domingo', 'Lunes', 'Martes', 'Miercoles', 'Jueves', 'Viernes', 'Sabado'],"
'[UI-ANT] html = html & "dayNamesShort: ['Dom','Lun','Mar','Mi?','Juv','Vie','S?b'],"
'[UI-ANT] html = html & "dayNamesMin: ['Do','Lu','Ma','Mi','Ju','Vi','Sa'],"
'[UI-ANT] html = html & "weekHeader: 'Sm',"
'[UI-ANT] html = html & "dateFormat: 'dd/mm/yy',"
'[UI-ANT] html = html & "firstDay: 1,"
'[UI-ANT] html = html & "isRTL: false,"
'[UI-ANT] html = html & "showMonthAfterYear: false,"
'[UI-ANT] html = html & "yearSuffix: ''"
'[UI-ANT] html = html & "};"
'[UI-ANT] html = html & "$.datepicker.setDefaults($.datepicker.regional['es']);"

'[UI-ANT] html = html & "    $( ""#fcd"" ).datepicker();"
'[UI-ANT] html = html & "    $( ""#fch"" ).datepicker();"
'[UI-ANT] html = html & "    $( ""#fad"" ).datepicker();"
'[UI-ANT] html = html & "    $( ""#fah"" ).datepicker();"

'[UI-ANT] html = html & "  } );"
'[UI-ANT] html = html & "  </script>"

'[UI-ANT] html = html & formx04 & "</body>"
'[UI-ANT] html = html & "</html>"

' =========================================================================
' [UI] VISTA con el sistema visual (IncFile/ui_componentes.inc)
' El bloque anterior completo queda arriba, comentado con '[UI-ANT].
' Formularios (destino y nombres de campos) y permiso: los mismos de antes.
' =========================================================================
Dim accTop, tipos, t, selTab, tabsHtml, tituloEst, tipoEst, puedeConsultar

' DEPURADO [UI]: la cabecera propia (logo e imagen armados a mano) se reemplaza por cabecera()
html = html & "<div id='wrapper' role='banner'>" & cabecera() & "</div>"

formx04 = "<form id='formx04' class='form-oculto' method='post' action='index.asp'></form>"
accTop = UiBoton("Volver", "secondary", "document.getElementById('formx04').submit();", True, "")
html = html & UiPantallaAbre("Estad&iacute;sticas", accTop, "")

' --- Pestanas: tipo de estadistica (antes un desplegable) ---
selTab = CInt("0" & sel)
tipos = Array("Incidentes", "Proyectos", "APL", "Usuarios")
tabsHtml = "<nav class='ui-tabbar ui-tabbar--pildoras' aria-label='Tipo de estad&iacute;stica'>"
For t = 0 To UBound(tipos)
	If t = selTab Then
		tabsHtml = tabsHtml & "<a class='ui-tabbar__link is-active' aria-current='page' href='estadisticas.asp?sel=" & t & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>" & tipos(t) & "</a>"
	Else
		tabsHtml = tabsHtml & "<a class='ui-tabbar__link' href='estadisticas.asp?sel=" & t & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>" & tipos(t) & "</a>"
	End If
Next
html = html & tabsHtml & "</nav>"

' --- Formulario segun el tipo (destino de antes: estadisticasDet.asp con su tipo) ---
If sel = 0 Then
	tituloEst = "Estad&iacute;sticas de incidentes"
	tipoEst = "I"
ElseIf sel = 1 Then
	tituloEst = "Estad&iacute;sticas de proyectos"
	tipoEst = "P"
ElseIf sel = 2 Then
	tituloEst = "Estad&iacute;sticas de APL"
	tipoEst = "A"
Else
	tituloEst = "Estad&iacute;sticas de usuarios"
	tipoEst = "U"
End If
puedeConsultar = (getPermisosAtrm("CG_ESTADISTICAS","CONS") = 1)

html = html & "<form id='xForm' method='post' action='estadisticasDet.asp?tipo=" & tipoEst & "&sel=" & sel & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
html = html & UiCardAbre(tituloEst, "") & UiFormGridAbre()

If tipoEst <> "U" Then
	html = html & UiCampo("id", "ID", UiInputTexto("id", False, "inputmode='numeric' autocomplete='off'"), False, False)
	html = html & UiCampo("cli", "Cliente", UiSelectRs("cli", Exec(cnn, "sp_cg_Catalogo_Listar @Catalogo='clientes'"), request.form("cli"), "Todos los clientes", "0", ""), False, False)
	html = html & UiCampo("uau", "&Uacute;ltimo usuario que accedi&oacute;", UiSelectRs("uau", Exec(Cnn, "sp_cg_Responsables @Modo='combo_todos'"), request.form("uau"), "Todos", "0", ""), False, False)
	html = html & UiCampo("uta", "&Uacute;ltimo tipo de acceso", UiSelectOpciones("uta", "|Todos||CONSULTA|Consulta||ACTUALIZA|Actualizaci&oacute;n", request.form("uta"), ""), False, False)
	html = html & UiGrupoAbre("Fechas")
	html = html & UiCampo("fcd", "Creaci&oacute;n desde", UiInputTexto("fcd", False, "autocomplete='off' placeholder='dd/mm/aaaa'"), False, False)
	html = html & UiCampo("fch", "Creaci&oacute;n hasta", UiInputTexto("fch", False, "autocomplete='off' placeholder='dd/mm/aaaa'"), False, False)
	html = html & UiCampo("fad", "&Uacute;ltimo acceso desde", UiInputTexto("fad", False, "autocomplete='off' placeholder='dd/mm/aaaa'"), False, False)
	html = html & UiCampo("fah", "&Uacute;ltimo acceso hasta", UiInputTexto("fah", False, "autocomplete='off' placeholder='dd/mm/aaaa'"), False, False)
	html = html & UiGrupoCierra()
Else
	html = html & UiCampo("uau", "&Uacute;ltimo usuario que accedi&oacute;", UiSelectRs("uau", Exec(Cnn, "sp_cg_Responsables @Modo='combo_todos'"), request.form("uau"), "Todos", "0", ""), False, True)
End If

html = html & UiFormGridCierra() & UiCardCierra()
html = html & UiFormAcciones(UiBotonEnviar("Consultar", puedeConsultar, "Sin permiso para consultar."))
html = html & "</form>"

html = html & UiPantallaCierra()

html = html & "<script src='assets/js/jquery.min.js'></script>"
html = html & "<script src='assets/js/skel.min.js'></script>"
html = html & "<script src='assets/js/util.js'></script>"
html = html & "<script src='assets/js/main.js'></script>"
html = html & "<script src='assets/js/jquery-ui.js'></script>"

html = html & "<script>"
html = html & "  $( function() {"
html = html & "$.datepicker.regional['es'] = {"
html = html & "closeText: 'Cerrar',"
html = html & "prevText: '< Ant',"
html = html & "nextText: 'Sig >',"
html = html & "currentText: 'Hoy',"
html = html & "monthNames: ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'],"
html = html & "monthNamesShort: ['Ene','Feb','Mar','Abr', 'May','Jun','Jul','Ago','Sep', 'Oct','Nov','Dic'],"
html = html & "dayNames: ['Domingo', 'Lunes', 'Martes', 'Mi\u00e9rcoles', 'Jueves', 'Viernes', 'S\u00e1bado'],"
html = html & "dayNamesShort: ['Dom','Lun','Mar','Mi\u00e9','Jue','Vie','S\u00e1b'],"
html = html & "dayNamesMin: ['Do','Lu','Ma','Mi','Ju','Vi','Sa'],"
html = html & "weekHeader: 'Sm',"
html = html & "dateFormat: 'dd/mm/yy',"
html = html & "firstDay: 1,"
html = html & "isRTL: false,"
html = html & "showMonthAfterYear: false,"
html = html & "yearSuffix: ''"
html = html & "};"
html = html & "$.datepicker.setDefaults($.datepicker.regional['es']);"
html = html & "    $( ""#fcd"" ).datepicker();"
html = html & "    $( ""#fch"" ).datepicker();"
html = html & "    $( ""#fad"" ).datepicker();"
html = html & "    $( ""#fah"" ).datepicker();"
html = html & "  } );"
html = html & "  </script>"

html = html & formx04 & "</body>"
html = html & "</html>"

response.write html

Function armaCombo(sql,cn,clave,valor,texto,comboid)

	Set nrs = Exec(cn, sql)

	xhtml = xhtml & "<select class='selec' xstyle='font-size:12pt;height:25px;color:#000000;border:solid 1px #000000;' id='" & comboid & "' name='" & comboid & "'>"

	if comboid = "uau" then
		xhtml = xhtml & "<option value='0' selected>TODOS</option>"
	else
		xhtml = xhtml & "<option value='0' selected>SIN DATOS</option>"
	end if


	Do While Not nrs.eof

		If clave = nrs(valor) Then

			xhtml = xhtml & "<option value='" & nrs(valor) & "' selected>" & nrs(texto) & "</option>"

		Else

			xhtml = xhtml & "<option value='" & nrs(valor) & "'>" & nrs(texto) & "</option>"

		End If

		nrs.movenext
	Loop

	armaCombo = xhtml & "</select>"

End Function

%>