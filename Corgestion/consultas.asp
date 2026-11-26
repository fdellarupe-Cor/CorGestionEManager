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

XXX = SYSLOGON(cnn,"MENU_CONSULTA")

If request.form("sel") & "" = "" Then
	sel = request.querystring("sel")
Else
	sel = request.form("sel")
End If

If sel & "" = "" Then sel = 0

xmenu = request.querystring("xmenu")
xtmenu = request.querystring("xtmenu")

'[UI-ANT] ' ===== Vista anterior completa: reemplazada por el bloque [UI] que sigue. Se borra entera buscando [UI-ANT] =====
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

'[UI-ANT] html = html & "<td><h2>CONSULTAS</h2></td>"

'[UI-ANT] html = html & "<td>"

'[UI-ANT] html = html & "<div>"

'[UI-ANT] html = html & "<form id='sform' style='height:100%;width:100%;' method='post' action='consultas.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "' >"

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
	'[UI-ANT] html = html & "<option value='3' selected >PROYECTOS C/ INCIDENTES</option>"
'[UI-ANT] Else
	'[UI-ANT] html = html & "<option value='3'>PROYECTOS C/ INCIDENTES</option>"
'[UI-ANT] End If

'[UI-ANT] If sel = 4 Then
	'[UI-ANT] html = html & "<option value='4' selected >INC Y PROY POR RESPONSABLE</option>"
'[UI-ANT] Else
	'[UI-ANT] html = html & "<option value='4'>INC Y PROY POR RESPONSABLE</option>"
'[UI-ANT] End If

'[UI-ANT] html = html & "</select>"

'[UI-ANT] html = html & "</form>"

'[UI-ANT] html = html & "</div>"

'[UI-ANT] html = html & "</td>"

'[UI-ANT] formx04 = "<form id='formx04' method='post' action='index.asp'></form>"
'[UI-ANT] botonx04 = "<input class='bot' type='button' value='VOLVER' onclick=""document.getElementById('formx04').submit();"">"
'[UI-ANT] 'html = html & "<td class='fil' align=center >" & botonx04 & "</td>"
'[UI-ANT] html = html & "<td>" & botonx04 & "</td>"

'[UI-ANT] html = html & "</tr>"

'[UI-ANT] html = html & "</table>"

'[UI-ANT] html = html & "</div><br>"

'[UI-ANT] 'DIV QUE SE RELLENA DEPENDE LO QUE QUIERA CONSULTAR ELEGIDO EN EL COMBO

'[UI-ANT] If sel = 0 Then 'INCIDENTES

	'[UI-ANT] ' --- INICIO NORMALIZACION: FILTRO DE INCIDENTES ---

	'[UI-ANT] 'html = html & "<div style='height:75%;width:100%;'><br>"
	'[UI-ANT] html = html & "<div class='inner' style='padding: 2rem;'>"

	'[UI-ANT] 'html = html & "<form id='xForm' method='post' action='incidentes.asp?opc=consi&sel=" & sel & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
	'[UI-ANT] html = html & "<form id='xForm' method='post' action='incidentes.asp?opc=consi&sel=" & sel & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"

	'[UI-ANT] 'html = html & "<table class='tabl' style='height:50%;width:40%;' align=center cellspacing=0>"
	'[UI-ANT] html = html & "<div class='bloque-datos' style='max-width: 600px; margin: 0 auto;'>"

	'[UI-ANT] 'ref = "<a class='red' href='incidentes.asp'>INCIDENTES</a>"
	'[UI-ANT] ' Ajustamos el enlace para que siga siendo rojo pero sin romper la est?tica moderna del header
	'[UI-ANT] ref = "<a class='col-link' href='incidentes.asp'>INCIDENTES</a>"
	
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' style='font-size:12pt;' colspan=2 align=center valign=center>FILTRAR " & ref & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='bloque-header' style='text-align: center; text-transform: uppercase;'>FILTRAR " & ref & "</div>"

	'[UI-ANT] html = html & "<div class='bloque-body'>"
	'[UI-ANT] html = html & "<div class='grid-datos'>"

	'[UI-ANT] ''CLIENTE
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center>CLIENTE</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>CLIENTE</div>"

	'[UI-ANT] 'html = html & "<td class='fil' valign=center>"
	'[UI-ANT] 'html = html & armaCombo("select id,descripcion from clientes",cnn,request.form("cli"),0,1,"cli")
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-value'>" & armaCombo("sp_cg_Catalogo_Listar @Catalogo='clientes'",cnn,request.form("cli"),0,1,"cli") & "</div>"

	'[UI-ANT] ''DESCRI
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center>DESCRIPCION</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>DESCRIPCION</div>"

	'[UI-ANT] 'html = html & "<td class='fil' valign=center>"
	'[UI-ANT] 'html = html & "<input class='inpu' id='desc' name='desc' style='width:70%;' type='text' />"
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div><input class='input-moderno' id='desc' name='desc' type='text' /></div>"

	'[UI-ANT] ''ESTADO
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center>ESTADO</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>ESTADO</div>"

	'[UI-ANT] 'html = html & "<td class='fil' valign=center>"
	'[UI-ANT] 'html = html & armaCombo("select id,descripcion from CG_ESTADOS",cnn,request.form("est"),0,1,"est")
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-value'>" & armaCombo("sp_cg_Catalogo_Listar @Catalogo='estados'",cnn,request.form("est"),0,1,"est") & "</div>"

	'[UI-ANT] ''APLIC
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center>APLICACION</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>APLICACION</div>"

	'[UI-ANT] 'html = html & "<td class='fil' valign=center>"
	'[UI-ANT] 'html = html & armaCombo("select id,descripcion from aplicativos",cnn,request.form("apli"),0,1,"apli")
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-value'>" & armaCombo("sp_cg_Catalogo_Listar @Catalogo='aplicativos'",cnn,request.form("apli"),0,1,"apli") & "</div>"

	'[UI-ANT] ''FECHA
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center>FECHA INGRESO</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>FECHA INGRESO</div>"

	'[UI-ANT] 'html = html & "<td class='fil' valign=center>"
	'[UI-ANT] 'html = html & "<input class='inpu' id='fec' name='fec' style='width:70%;' type='text' />"
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div><input class='input-moderno' id='fec' name='fec' type='text' /></div>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td style='height:10%;width:50%;' colspan=2 align=center>"
	'[UI-ANT] ' Abrimos el contenedor Flexbox para el bot?n centrado rompiendo la grilla
	'[UI-ANT] html = html & "<div style='grid-column: 1 / -1; display: flex; justify-content: center; margin-top: 1.5rem;'>"

	'[UI-ANT] if getPermisos(xmenu,xtmenu,"CONS") = 1 then
		'[UI-ANT] 'html = html & "<input class='bot' type='button' value='Consultar' onclick=""document.getElementById('xForm').submit();"">"
		'[UI-ANT] html = html & "<input class='btn-principal' type='button' value='Consultar' onclick=""document.getElementById('xForm').submit();"">"
	'[UI-ANT] else
		'[UI-ANT] 'html = html & "<input class='botsp' type='button' value='Consultar' onclick=""alert('Sin permiso para consultar.');"">"
		'[UI-ANT] html = html & "<input class='btn-principal btn-accion-deshabilitado' type='button' value='Consultar' onclick=""alert('Sin permiso para consultar.');"">"
	'[UI-ANT] end if

	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "</div>" ' Cierre de contenedor flexbox del bot?n

	'[UI-ANT] 'html = html & "</table>"
	'[UI-ANT] html = html & "</div>" ' Cierre de .grid-datos
	'[UI-ANT] html = html & "</div>" ' Cierre de .bloque-body
	'[UI-ANT] html = html & "</div>" ' Cierre de .bloque-datos

	'[UI-ANT] 'html = html & "</form>"
	'[UI-ANT] html = html & "</form>"

	'[UI-ANT] 'html = html & "</div>"
	'[UI-ANT] html = html & "</div>" ' Cierre de .inner

'[UI-ANT] ElseIf sel = 1 Then 'PROYECTOS

	'[UI-ANT] ' --- INICIO NORMALIZACION: FILTRO DE PROYECTOS ---

	'[UI-ANT] 'html = html & "<div style='height:75%;width:100%;'><br>"
	'[UI-ANT] html = html & "<div class='inner' style='padding: 2rem;'>"

	'[UI-ANT] 'html = html & "<form id='xForm' method='post' action='proyectos.asp?opc=consi&sel=" & sel & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
	'[UI-ANT] html = html & "<form id='xForm' method='post' action='proyectos.asp?opc=consi&sel=" & sel & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"

	'[UI-ANT] 'html = html & "<table class='tabl' style='height:50%;width:40%;' align=center cellspacing=0>"
	'[UI-ANT] html = html & "<div class='bloque-datos' style='max-width: 600px; margin: 0 auto;'>"

	'[UI-ANT] 'ref = "<a class='red' href='proyectos.asp'>PROYECTO</a>"
	'[UI-ANT] ref = "<a class='col-link' href='proyectos.asp'>PROYECTO</a>"
	
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' style='font-size:12pt;' colspan=2 align=center valign=center>FILTRAR " & ref & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='bloque-header' style='text-align: center; text-transform: uppercase;'>FILTRAR " & ref & "</div>"

	'[UI-ANT] html = html & "<div class='bloque-body'>"
	'[UI-ANT] html = html & "<div class='grid-datos'>"

	'[UI-ANT] ''CLIENTE
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center>CLIENTE</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>CLIENTE</div>"

	'[UI-ANT] 'html = html & "<td class='fil' valign=center>"
	'[UI-ANT] 'html = html & armaCombo("select id,descripcion from clientes",cnn,request.form("cli"),0,1,"cli")
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-value'>" & armaCombo("sp_cg_Catalogo_Listar @Catalogo='clientes'",cnn,request.form("cli"),0,1,"cli") & "</div>"

	'[UI-ANT] ''DESCRI
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center>DESCRIPCION</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>DESCRIPCION</div>"

	'[UI-ANT] 'html = html & "<td class='fil' valign=center>"
	'[UI-ANT] 'html = html & "<input class='inpu' id='desc' name='desc' style='width:70%;' type='text' />"
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div><input class='input-moderno' id='desc' name='desc' type='text' /></div>"

	'[UI-ANT] ''ESTADO
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center>ESTADO</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>ESTADO</div>"

	'[UI-ANT] 'html = html & "<td class='fil' valign=center>"
	'[UI-ANT] 'html = html & armaCombo("select id,descripcion from CG_ESTADOS",cnn,request.form("est"),0,1,"est")
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-value'>" & armaCombo("sp_cg_Catalogo_Listar @Catalogo='estados'",cnn,request.form("est"),0,1,"est") & "</div>"

	'[UI-ANT] ''APLIC
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center>APLICACION</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>APLICACION</div>"

	'[UI-ANT] 'html = html & "<td class='fil' valign=center>"
	'[UI-ANT] 'html = html & armaCombo("select id,descripcion from aplicativos",cnn,request.form("apli"),0,1,"apli")
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-value'>" & armaCombo("sp_cg_Catalogo_Listar @Catalogo='aplicativos'",cnn,request.form("apli"),0,1,"apli") & "</div>"

	'[UI-ANT] ''FECHA
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center>FECHA INGRESO</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>FECHA INGRESO</div>"

	'[UI-ANT] 'html = html & "<td class='fil' valign=center>"
	'[UI-ANT] 'html = html & "<input class='inpu' id='fec' name='fec' style='width:70%;' type='text' />"
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div><input class='input-moderno' id='fec' name='fec' type='text' /></div>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td style='height:10%;width:50%;' colspan=2 align=center>"
	'[UI-ANT] ' Abrimos el contenedor Flexbox para el bot?n centrado
	'[UI-ANT] html = html & "<div style='grid-column: 1 / -1; display: flex; justify-content: center; margin-top: 1.5rem;'>"

	'[UI-ANT] if getPermisos(xmenu,xtmenu,"CONS") = 1 then
		'[UI-ANT] 'html = html & "<input type='button' class='bot' value='Consultar' onclick=""document.getElementById('xForm').submit();"">"
		'[UI-ANT] html = html & "<input class='btn-principal' type='button' value='Consultar' onclick=""document.getElementById('xForm').submit();"">"
	'[UI-ANT] else
		'[UI-ANT] 'html = html & "<input type='button' class='botsp' value='Consultar' onclick=""alert('Sin permiso para consultar.');"">"
		'[UI-ANT] html = html & "<input class='btn-principal btn-accion-deshabilitado' type='button' value='Consultar' onclick=""alert('Sin permiso para consultar.');"">"
	'[UI-ANT] end if

	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "</div>" ' Cierre de contenedor flexbox del bot?n

	'[UI-ANT] 'html = html & "</table>"
	'[UI-ANT] html = html & "</div>" ' Cierre de .grid-datos
	'[UI-ANT] html = html & "</div>" ' Cierre de .bloque-body
	'[UI-ANT] html = html & "</div>" ' Cierre de .bloque-datos

	'[UI-ANT] 'html = html & "</form>"
	'[UI-ANT] html = html & "</form>"

	'[UI-ANT] 'html = html & "</div>"
	'[UI-ANT] html = html & "</div>" ' Cierre de .inner

'[UI-ANT] ElseIf sel = 2 Then 'APL

	'[UI-ANT] ' --- INICIO NORMALIZACION: FILTRO DE APLs ---

	'[UI-ANT] 'html = html & "<div style='height:75%;width:100%;'><br>"
	'[UI-ANT] html = html & "<div class='inner' style='padding: 2rem;'>"

	'[UI-ANT] 'html = html & "<form id='xForm' method='post' action='apl.asp?opc=consi&sel=" & sel & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
	'[UI-ANT] html = html & "<form id='xForm' method='post' action='apl.asp?opc=consi&sel=" & sel & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"

	'[UI-ANT] 'html = html & "<table class='tabl' style='height:50%;width:40%;' align=center cellspacing=0>"
	'[UI-ANT] html = html & "<div class='bloque-datos' style='max-width: 600px; margin: 0 auto;'>"

	'[UI-ANT] 'ref = "<a class='red' href='apl.asp'>APLs</a>"
	'[UI-ANT] ' Mantenemos tu directiva del color azul para la referencia
	'[UI-ANT] ref = "<a class='col-link' href='apl.asp'>APLs</a>"
	
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' style='font-size:12pt;' colspan=2 align=center valign=center>FILTRAR " & ref & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='bloque-header' style='text-align: center; text-transform: uppercase;'>FILTRAR " & ref & "</div>"

	'[UI-ANT] html = html & "<div class='bloque-body'>"
	'[UI-ANT] html = html & "<div class='grid-datos'>"

	'[UI-ANT] ''CLIENTE
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center>CLIENTE</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>CLIENTE</div>"

	'[UI-ANT] 'html = html & "<td class='fil' valign=center>"
	'[UI-ANT] 'html = html & armaCombo("select id,descripcion from clientes",cnn,request.form("cli"),0,1,"cli")
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-value'>" & armaCombo("sp_cg_Catalogo_Listar @Catalogo='clientes'",cnn,request.form("cli"),0,1,"cli") & "</div>"

	'[UI-ANT] ''APLIC
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center>APLICACION</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>APLICACION</div>"

	'[UI-ANT] 'html = html & "<td class='fil' valign=center>"
	'[UI-ANT] 'html = html & armaCombo("select id,descripcion from aplicativos",cnn,request.form("apli"),0,1,"apli")
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-value'>" & armaCombo("sp_cg_Catalogo_Listar @Catalogo='aplicativos'",cnn,request.form("apli"),0,1,"apli") & "</div>"

	'[UI-ANT] ''FECHA
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center>FECHA CREACION</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>FECHA CREACION</div>"

	'[UI-ANT] 'html = html & "<td class='fil' valign=center>"
	'[UI-ANT] 'html = html & "<input class='inpu' id='fec' name='fec' style='width:70%;' type='text' />"
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div><input class='input-moderno' id='fec' name='fec' type='text' /></div>"

	'[UI-ANT] ''BUSCADOR POR TEXTO
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center>NOMBRE ASP</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>NOMBRE ASP</div>"

	'[UI-ANT] 'html = html & "<td class='fil' valign=center>"
	'[UI-ANT] 'html = html & "<input class='inpu' style='width:70%;' id='busc' name='busc' type='text'>"
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div><input class='input-moderno' id='busc' name='busc' type='text'></div>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td style='height:10%;width:50%;' colspan=2 align=center>"
	'[UI-ANT] ' Contenedor Flexbox para centrar el bot?n rompiendo la grilla
	'[UI-ANT] html = html & "<div style='grid-column: 1 / -1; display: flex; justify-content: center; margin-top: 1.5rem;'>"

	'[UI-ANT] if getPermisos(xmenu,xtmenu,"CONS") = 1 then
		'[UI-ANT] 'html = html & "<input class='bot' type='button' value='Consultar' onclick=""document.getElementById('xForm').submit();"">"
		'[UI-ANT] html = html & "<input class='btn-principal' type='button' value='Consultar' onclick=""document.getElementById('xForm').submit();"">"
	'[UI-ANT] else
		'[UI-ANT] 'html = html & "<input class='botsp' type='button' value='Consultar' onclick=""alert('Sin permiso para consultar.');"">"
		'[UI-ANT] html = html & "<input class='btn-principal btn-accion-deshabilitado' type='button' value='Consultar' onclick=""alert('Sin permiso para consultar.');"">"
	'[UI-ANT] end if

	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "</div>" ' Cierre de contenedor flexbox

	'[UI-ANT] 'html = html & "</table>"
	'[UI-ANT] html = html & "</div>" ' Cierre de .grid-datos
	'[UI-ANT] html = html & "</div>" ' Cierre de .bloque-body
	'[UI-ANT] html = html & "</div>" ' Cierre de .bloque-datos

	'[UI-ANT] 'html = html & "</form>"
	'[UI-ANT] html = html & "</form>"

	'[UI-ANT] 'html = html & "</div>"
	'[UI-ANT] html = html & "</div>" ' Cierre de .inner

'[UI-ANT] ElseIf sel = 3 Then 'PROYECTOS C/ INCIDENTES

	'[UI-ANT] ' --- INICIO NORMALIZACION: FILTRO DE PROYECTOS C/ INCIDENTES ---

	'[UI-ANT] 'html = html & "<div style='height:75%;width:100%;'><br>"
	'[UI-ANT] html = html & "<div class='inner' style='padding: 2rem;'>"

	'[UI-ANT] 'html = html & "<form id='xForm' method='post' action='proyRelac.asp?sel=" & sel & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
	'[UI-ANT] html = html & "<form id='xForm' method='post' action='proyRelac.asp?sel=" & sel & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"

	'[UI-ANT] 'html = html & "<table class='tabl' style='height:50%;width:40%;' align=center cellspacing=0>"
	'[UI-ANT] html = html & "<div class='bloque-datos' style='max-width: 600px; margin: 0 auto;'>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' style='font-size:12pt;' colspan=2 align=center valign=center><b>PROYECTOS C/ INCIDENTES</b></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='bloque-header' style='text-align: center; text-transform: uppercase;'>PROYECTOS C/ INCIDENTES</div>"

	'[UI-ANT] html = html & "<div class='bloque-body'>"
	'[UI-ANT] html = html & "<div class='grid-datos'>"

	'[UI-ANT] ''CLIENTE
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center>CLIENTE</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>CLIENTE</div>"

	'[UI-ANT] 'html = html & "<td class='fil' valign=center>"
	'[UI-ANT] 'html = html & armaCombo("select id,descripcion from clientes",cnn,request.form("cli"),0,1,"cli")
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-value'>" & armaCombo("sp_cg_Catalogo_Listar @Catalogo='clientes'",cnn,request.form("cli"),0,1,"cli") & "</div>"

	'[UI-ANT] ''ESTADO
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center>ESTADO</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>ESTADO</div>"

	'[UI-ANT] 'html = html & "<td class='fil' valign=center>"
	'[UI-ANT] 'html = html & armaCombo("select id,descripcion from CG_ESTADOS",cnn,request.form("est"),0,1,"est")
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-value'>" & armaCombo("sp_cg_Catalogo_Listar @Catalogo='estados'",cnn,request.form("est"),0,1,"est") & "</div>"

	'[UI-ANT] ''APLIC
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center>APLICACION</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>APLICACION</div>"

	'[UI-ANT] 'html = html & "<td class='fil' valign=center>"
	'[UI-ANT] 'html = html & armaCombo("select id,descripcion from aplicativos",cnn,request.form("apli"),0,1,"apli")
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-value'>" & armaCombo("sp_cg_Catalogo_Listar @Catalogo='aplicativos'",cnn,request.form("apli"),0,1,"apli") & "</div>"

	'[UI-ANT] ''FECHA
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center>FECHA INGRESO</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>FECHA INGRESO</div>"

	'[UI-ANT] 'html = html & "<td class='fil' valign=center>"
	'[UI-ANT] 'html = html & "<input class='inpu' id='fec' name='fec' style='width:70%;' type='text' />"
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div><input class='input-moderno' id='fec' name='fec' type='text' /></div>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td style='height:10%;width:50%;' colspan=2 align=center>"
	'[UI-ANT] ' Contenedor Flexbox para centrar el bot?n rompiendo la grilla
	'[UI-ANT] html = html & "<div style='grid-column: 1 / -1; display: flex; justify-content: center; margin-top: 1.5rem;'>"

	'[UI-ANT] if getPermisos(xmenu,xtmenu,"CONS") = 1 then
		'[UI-ANT] 'html = html & "<input class='bot' type='button' value='Consultar' onclick=""document.getElementById('xForm').submit();"">"
		'[UI-ANT] html = html & "<input class='btn-principal' type='button' value='Consultar' onclick=""document.getElementById('xForm').submit();"">"
	'[UI-ANT] else
		'[UI-ANT] 'html = html & "<input class='botsp' type='button' value='Consultar' onclick=""alert('Sin permiso para consultar.');"">"
		'[UI-ANT] html = html & "<input class='btn-principal btn-accion-deshabilitado' type='button' value='Consultar' onclick=""alert('Sin permiso para consultar.');"">"
	'[UI-ANT] end if

	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "</div>" ' Cierre de contenedor flexbox

	'[UI-ANT] 'html = html & "</table>"
	'[UI-ANT] html = html & "</div>" ' Cierre de .grid-datos
	'[UI-ANT] html = html & "</div>" ' Cierre de .bloque-body
	'[UI-ANT] html = html & "</div>" ' Cierre de .bloque-datos

	'[UI-ANT] 'html = html & "</form>"
	'[UI-ANT] html = html & "</form>"

	'[UI-ANT] 'html = html & "</div>"
	'[UI-ANT] html = html & "</div>" ' Cierre de .inner

'[UI-ANT] ElseIf sel = 4 Then 'INC Y PROY POR RESPONSABLE

	'[UI-ANT] ' --- INICIO NORMALIZACION: FILTRO DE INC Y PROY POR RESPONSABLE ---

	'[UI-ANT] 'html = html & "<div style='height:75%;width:100%;'><br>"
	'[UI-ANT] html = html & "<div class='inner' style='padding: 2rem;'>"

	'[UI-ANT] 'html = html & "<form id='xForm' method='post' action='consultas.asp?sel=5&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
	'[UI-ANT] html = html & "<form id='xForm' method='post' action='consultas.asp?sel=5&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"

	'[UI-ANT] 'html = html & "<table class='tabl' style='height:30%;width:30%;' align=center cellspacing=0>"
	'[UI-ANT] html = html & "<div class='bloque-datos' style='max-width: 600px; margin: 0 auto;'>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' style='font-size:12pt;' colspan=2 align=center valign=center><b>INC Y PROY POR RESPONSABLE</b></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='bloque-header' style='text-align: center; text-transform: uppercase;'>INC Y PROY POR RESPONSABLE</div>"

	'[UI-ANT] html = html & "<div class='bloque-body'>"
	'[UI-ANT] html = html & "<div class='grid-datos'>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' valign=center align=center>RESPONSABLE</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>RESPONSABLE</div>"

	'[UI-ANT] 'html = html & "<td class='fil' valign=center align=center>"
	'[UI-ANT] 'html = html & armaCombo("select id,nombre + ' ' + apellido,funcion from CG_COR_RESPONSABLES",Cnn,request.form("uau"),0,1,"uau")
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-value'>" & armaCombo("sp_cg_Responsables @Modo='combo_todos'",Cnn,request.form("uau"),0,1,"uau") & "</div>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' colspan=2 align=center>"
	'[UI-ANT] ' Contenedor Flexbox para centrar el bot?n rompiendo la grilla
	'[UI-ANT] html = html & "<div style='grid-column: 1 / -1; display: flex; justify-content: center; margin-top: 1.5rem;'>"

	'[UI-ANT] if getPermisos(xmenu,xtmenu,"CONS") = 1 then
		'[UI-ANT] 'html = html & "<input class='bot' type='submit' value='Consultar' onclick=""document.getElementById('xForm').submit();"">"
		'[UI-ANT] html = html & "<input class='btn-principal' type='submit' value='Consultar' onclick=""document.getElementById('xForm').submit();"">"
	'[UI-ANT] else
		'[UI-ANT] 'html = html & "<input class='botsp' type='submit' value='Consultar' onclick=""alert('Sin permiso para consultar.');"">"
		'[UI-ANT] html = html & "<input class='btn-principal btn-accion-deshabilitado' type='submit' value='Consultar' onclick=""alert('Sin permiso para consultar.');"">"
	'[UI-ANT] end if

	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "</div>" ' Cierre de contenedor flexbox

	'[UI-ANT] 'html = html & "</table>"
	'[UI-ANT] html = html & "</div>" ' Cierre de .grid-datos
	'[UI-ANT] html = html & "</div>" ' Cierre de .bloque-body
	'[UI-ANT] html = html & "</div>" ' Cierre de .bloque-datos

	'[UI-ANT] 'html = html & "</form>"
	'[UI-ANT] html = html & "</form>"

	'[UI-ANT] 'html = html & "</div>"
	'[UI-ANT] html = html & "</div>" ' Cierre de .inner

'[UI-ANT] ElseIf sel = 5 Then

'[UI-ANT] idResp = request.form("uau")

'[UI-ANT] if Trim(idResp & "") = "0" Or Trim(idResp & "") = "" then 'POR TODOS

'[UI-ANT] Set rt = Exec(Cnn, "sp_cg_Responsables @Modo='id_funcion'")

'[UI-ANT] ' --- INICIO NORMALIZACION: TABLA DE RESULTADOS (TODOS) ---
'[UI-ANT] html = html & "<div class='inner' style='padding: 2rem;'>"
'[UI-ANT] html = html & "<div class='bloque-datos table-responsive'>"

'[UI-ANT] 'html = html & "<table class='tabl' style='width:40%;color:#000000;' align=center cellspacing=0 cellpadding=0>"
'[UI-ANT] html = html & "<table class='tabla-moderna'>"

'[UI-ANT] html = html & "<thead>"
'[UI-ANT] 'html = html & "<tr>"
'[UI-ANT] html = html & "<tr>"
	
'[UI-ANT] 'html = html & "<td style='border-bottom: solid 1px #aaaaaa;' align=center>*</td>"
'[UI-ANT] html = html & "<th>RESPONSABLE</th>"
	
'[UI-ANT] 'html = html & "<td style='border-bottom: solid 1px #aaaaaa;' align=center>INCIDENTES</td>"
'[UI-ANT] html = html & "<th style='text-align: center;'>INCIDENTES</th>"
	
'[UI-ANT] 'html = html & "<td style='border-bottom: solid 1px #aaaaaa;' align=center>PROYECTOS</td>"
'[UI-ANT] html = html & "<th style='text-align: center;'>PROYECTOS</th>"
	
'[UI-ANT] 'html = html & "</tr>"
'[UI-ANT] html = html & "</tr>"
'[UI-ANT] html = html & "</thead>"

'[UI-ANT] html = html & "<tbody>"

'[UI-ANT] totInc = 0
'[UI-ANT] totProy = 0

'[UI-ANT] do while not rt.eof

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] html = html & "<tr>"
		
	'[UI-ANT] 'html = html & "<td align=center>" & getResp(cnn,rt(0)) & "</td>"
	'[UI-ANT] html = html & "<td>" & getResp(cnn,rt(0)) & "</td>"
		
	'[UI-ANT] 'html = html & "<td align=center>" & getCantInc(cnn,rt(0),rt(1)) & "</td>"
	'[UI-ANT] html = html & "<td style='text-align: center;'>" & getCantInc(cnn,rt(0),rt(1)) & "</td>"
		
	'[UI-ANT] 'html = html & "<td align=center>" & getCantProy(cnn,rt(0),rt(1)) & "</td>"
	'[UI-ANT] html = html & "<td style='text-align: center;'>" & getCantProy(cnn,rt(0),rt(1)) & "</td>"
		
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "</tr>"

	'[UI-ANT] totInc = totInc + getCantInc(cnn,rt(0),rt(1))
	'[UI-ANT] totProy = totProy + getCantProy(cnn,rt(0),rt(1))

'[UI-ANT] rt.movenext
'[UI-ANT] loop

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] ' Agregamos un fondo sutil y negrita a la fila de totales para destacarla sin usar bordes quemados
	'[UI-ANT] html = html & "<tr style='font-weight: bold; background-color: #f8fafc;'>"
	
	'[UI-ANT] 'html = html & "<td style='border-top: solid 1px #aaaaaa;' align=center>TOTAL</td>"
	'[UI-ANT] html = html & "<td style='text-align: right; padding-right: 1rem;'>TOTAL</td>"
	
	'[UI-ANT] 'html = html & "<td style='border-top: solid 1px #aaaaaa;' align=center>" & totInc & "</td>"
	'[UI-ANT] html = html & "<td style='text-align: center;'>" & totInc & "</td>"
	
	'[UI-ANT] 'html = html & "<td style='border-top: solid 1px #aaaaaa;' align=center>" & totProy & "</td>"
	'[UI-ANT] html = html & "<td style='text-align: center;'>" & totProy & "</td>"
	
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "</tr>"

	'[UI-ANT] 'html = html & "</table>"
	'[UI-ANT] html = html & "</tbody>"
	'[UI-ANT] html = html & "</table>"
	'[UI-ANT] html = html & "</div>"
	'[UI-ANT] html = html & "</div>"

'[UI-ANT] else 'POR UN RESP

'[UI-ANT] Set rj = Exec(Cnn, "sp_cg_Responsables @Modo='id_funcion_id', @IdSeg=" & SqlStr(idResp))

'[UI-ANT] if not rj.eof then

	'[UI-ANT] ' --- INICIO NORMALIZACION: TABLA DE RESULTADOS (UN RESPONSABLE) ---
	'[UI-ANT] html = html & "<div class='inner' style='padding: 2rem;'>"
	'[UI-ANT] html = html & "<div class='bloque-datos table-responsive'>"

	'[UI-ANT] 'html = html & "<table class='tabl' style='width:40%;color:#000000;' align=center cellspacing=0 cellpadding=0>"
	'[UI-ANT] html = html & "<table class='tabla-moderna'>"

	'[UI-ANT] html = html & "<thead>"
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] html = html & "<tr>"
		
	'[UI-ANT] 'html = html & "<td style='border-bottom: solid 1px #aaaaaa;' align=center>*</td>"
	'[UI-ANT] html = html & "<th>RESPONSABLE</th>"
		
	'[UI-ANT] 'html = html & "<td style='border-bottom: solid 1px #aaaaaa;' align=center>INCIDENTES</td>"
	'[UI-ANT] html = html & "<th style='text-align: center;'>INCIDENTES</th>"
		
	'[UI-ANT] 'html = html & "<td style='border-bottom: solid 1px #aaaaaa;' align=center>PROYECTOS</td>"
	'[UI-ANT] html = html & "<th style='text-align: center;'>PROYECTOS</th>"
		
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "</tr>"
	'[UI-ANT] html = html & "</thead>"

	'[UI-ANT] html = html & "<tbody>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] html = html & "<tr>"
		
	'[UI-ANT] 'html = html & "<td align=center>" & getResp(cnn,rj(0)) & "</td>"
	'[UI-ANT] html = html & "<td>" & getResp(cnn,rj(0)) & "</td>"
		
	'[UI-ANT] 'html = html & "<td align=center>" & getCantInc(cnn,rj(0),rj(1)) & "</td>"
	'[UI-ANT] html = html & "<td style='text-align: center;'>" & getCantInc(cnn,rj(0),rj(1)) & "</td>"
	
	'[UI-ANT] 'html = html & "<td align=center>" & getCantProy(cnn,rj(0),rj(1)) & "</td>"
	'[UI-ANT] html = html & "<td style='text-align: center;'>" & getCantProy(cnn,rj(0),rj(1)) & "</td>"
		
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "</tr>"

'[UI-ANT] sTotal = getCantInc(cnn,rj(0),rj(1)) + getCantProy(cnn,rj(0),rj(1))

'[UI-ANT] 'html = html & "<tr>"
'[UI-ANT] html = html & "<tr style='font-weight: bold; background-color: #f8fafc;'>"
		
'[UI-ANT] 'html = html & "<td style='border-top: solid 1px #aaaaaa;' align=center>TOTAL</td>"
'[UI-ANT] html = html & "<td style='text-align: right; padding-right: 1rem;'>TOTAL GENERAL</td>"
		
'[UI-ANT] 'html = html & "<td align=center style='border-top: solid 1px #aaaaaa;' colspan=2>" & sTotal & "</td>"
'[UI-ANT] html = html & "<td style='text-align: center;' colspan='2'>" & sTotal & "</td>"
		
'[UI-ANT] 'html = html & "</tr>"
'[UI-ANT] html = html & "</tr>"

'[UI-ANT] 'html = html & "</table>"
'[UI-ANT] html = html & "</tbody>"
'[UI-ANT] html = html & "</table>"
'[UI-ANT] html = html & "</div>"
'[UI-ANT] html = html & "</div>"

'[UI-ANT] end if

'[UI-ANT] end if

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

'[UI-ANT] html = html & "    $( ""#fec"" ).datepicker();"

'[UI-ANT] html = html & "  } );"
'[UI-ANT] html = html & "  </script>"

'[UI-ANT] html = html & formx04 & "</body>"
'[UI-ANT] html = html & "</html>"

' =========================================================================
' [UI] VISTA con el sistema visual (IncFile/ui_componentes.inc)
' El bloque anterior completo queda arriba, comentado con '[UI-ANT].
' Formularios (destinos y nombres de campos), permisos y consultas: los mismos de antes.
' =========================================================================
Dim accTop, tabsHtml, selTab, tituloCons, accionCons, irListado, puedeConsultar, tipos, t
Dim cantInc, cantProy, totalUno

' DEPURADO [UI]: la cabecera propia de esta pantalla (logo e imagen armados a mano) se reemplaza
' por cabecera(), la misma de todo el sistema
html = html & "<div id='wrapper' role='banner'>" & cabecera() & "</div>"

' --- Barra superior ---
formx04 = "<form id='formx04' class='form-oculto' method='post' action='index.asp'></form>"
accTop = UiBoton("Volver", "secondary", "document.getElementById('formx04').submit();", True, "")
html = html & UiPantallaAbre("Consultas", accTop, "")

' --- Pestanas: tipo de consulta (antes un desplegable). Los resultados por responsable (sel 5) marcan su pestana ---
selTab = CInt("0" & sel)
If selTab = 5 Then selTab = 4
tipos = Array("Incidentes", "Proyectos", "APL", "Proyectos con incidentes", "Por responsable")
tabsHtml = "<nav class='ui-tabbar ui-tabbar--consultas' aria-label='Tipo de consulta'>"
For t = 0 To UBound(tipos)
	If t = selTab Then
		tabsHtml = tabsHtml & "<a class='ui-tabbar__link is-active' aria-current='page' href='consultas.asp?sel=" & t & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>" & tipos(t) & "</a>"
	Else
		tabsHtml = tabsHtml & "<a class='ui-tabbar__link' href='consultas.asp?sel=" & t & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>" & tipos(t) & "</a>"
	End If
Next
tabsHtml = tabsHtml & "</nav>"
html = html & tabsHtml

puedeConsultar = (getPermisosAtrm("CG_CONSULTAS","CONS") = 1)

If sel <> 5 Then

	' --- Formulario de busqueda: destino y titulo segun el tipo ---
	irListado = ""
	If sel = 0 Then
		tituloCons = "Consultar incidentes"
		accionCons = "incidentes.asp?opc=consi&sel=" & sel & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu
		irListado = "<a class='ui-button ui-button--light ui-button--sm' href='incidentes.asp'>Ir al listado</a>"
	ElseIf sel = 1 Then
		tituloCons = "Consultar proyectos"
		accionCons = "proyectos.asp?opc=consi&sel=" & sel & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu
		irListado = "<a class='ui-button ui-button--light ui-button--sm' href='proyectos.asp'>Ir al listado</a>"
	ElseIf sel = 2 Then
		tituloCons = "Consultar APLs"
		accionCons = "apl.asp?opc=consi&sel=" & sel & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu
		irListado = "<a class='ui-button ui-button--light ui-button--sm' href='apl.asp'>Ir al listado</a>"
	ElseIf sel = 3 Then
		tituloCons = "Proyectos con incidentes relacionados"
		accionCons = "proyRelac.asp?sel=" & sel & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu
	Else
		tituloCons = "Incidentes y proyectos por responsable"
		accionCons = "consultas.asp?sel=5&xmenu=" & xmenu & "&xtmenu=" & xtmenu
	End If

	html = html & "<form id='xForm' method='post' action='" & accionCons & "'>"
	html = html & UiCardAbre(tituloCons, irListado) & UiFormGridAbre()

	If sel = 4 Then
		html = html & UiCampo("uau", "Responsable", UiSelectRs("uau", Exec(Cnn, "sp_cg_Responsables @Modo='combo_todos'"), request.form("uau"), "Todos", "0", ""), False, True)
	Else
		html = html & UiCampo("cli", "Cliente", UiSelectRs("cli", Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='clientes'"), request.form("cli"), "Todos los clientes", "0", ""), False, False)
		' apl.asp lee la aplicacion como 'aplic' (antes se mandaba 'apli' y el filtro se ignoraba)
		If sel = 2 Then
			html = html & UiCampo("aplic", "Aplicaci&oacute;n", UiSelectRs("aplic", Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='aplicativos'"), request.form("aplic"), "Todas las aplicaciones", "0", ""), False, False)
		Else
			html = html & UiCampo("apli", "Aplicaci&oacute;n", UiSelectRs("apli", Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='aplicativos'"), request.form("apli"), "Todas las aplicaciones", "0", ""), False, False)
		End If
		If sel <> 2 Then
			html = html & UiCampo("est", "Estado", UiSelectRs("est", Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='estados'"), request.form("est"), "Todos los estados", "0", ""), False, False)
		End If
		If sel = 2 Then
			html = html & UiCampo("fec", "Fecha de creaci&oacute;n", UiInputTexto("fec", False, "autocomplete='off' placeholder='dd/mm/aaaa'"), False, False)
			html = html & UiCampo("busc", "Motivo del APL", UiInputTexto("busc", False, "autocomplete='off'"), False, False)
		Else
			html = html & UiCampo("fec", "Fecha de ingreso", UiInputTexto("fec", False, "autocomplete='off' placeholder='dd/mm/aaaa'"), False, False)
		End If
		If sel = 0 Or sel = 1 Then
			html = html & UiCampo("desc", "Descripci&oacute;n", UiInputTexto("desc", False, "autocomplete='off'"), False, True)
		End If
	End If

	html = html & UiFormGridCierra() & UiCardCierra()
	html = html & UiFormAcciones(UiBotonEnviar("Consultar", puedeConsultar, "Sin permiso para consultar."))
	html = html & "</form>"

Else
	' --- Resultado: incidentes y proyectos por responsable (sel 5) ---
	idResp = request.form("uau")

	html = html & UiCardAbre("Incidentes y proyectos por responsable", "<a class='ui-button ui-button--light ui-button--sm' href='consultas.asp?sel=4&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>Nueva consulta</a>")
	html = html & UiTablaAbre("Cantidad de incidentes y proyectos de cada responsable")
	html = html & "<thead><tr>" & UiTh("Responsable", "") & UiTh("Incidentes", "ui-table__cell--center") & UiTh("Proyectos", "ui-table__cell--center") & "</tr></thead><tbody>"

	if Trim(idResp & "") = "0" Or Trim(idResp & "") = "" then 'POR TODOS

		' [CONSULTA] Paginado de a 10 responsables; los totales siguen siendo los de TODOS los responsables
		Dim cantRegMoverP, cantRegMostrarP, idxFilaP, idxMostradasP
		cantRegMostrarP = 10
		cantRegMoverP = Request.QueryString("cantregmover")
		If Not IsNumeric(cantRegMoverP & "") Or Trim(cantRegMoverP & "") = "" Then cantRegMoverP = 0
		cantRegMoverP = CLng(cantRegMoverP)
		If cantRegMoverP < 0 Then cantRegMoverP = 0
		idxFilaP = 0
		idxMostradasP = 0
		Set rt = Exec(Cnn, "sp_cg_Responsables @Modo='id_funcion'")
		totInc = 0
		totProy = 0
		do while not rt.eof
			' Cada conteo se consulta una sola vez (antes se pedia de nuevo para el total)
			cantInc = getCantInc(cnn,rt(0),rt(1))
			cantProy = getCantProy(cnn,rt(0),rt(1))
			If idxFilaP >= cantRegMoverP And idxMostradasP < cantRegMostrarP Then
				html = html & "<tr><td>" & UiTexto(getResp(cnn,rt(0))) & "</td>" & _
					"<td class='ui-table__cell--center'>" & UiDato(cantInc) & "</td>" & _
					"<td class='ui-table__cell--center'>" & UiDato(cantProy) & "</td></tr>"
				idxMostradasP = idxMostradasP + 1
			End If
			idxFilaP = idxFilaP + 1
			totInc = totInc + CLng("0" & cantInc)
			totProy = totProy + CLng("0" & cantProy)
			rt.movenext
		loop
		html = html & "</tbody><tfoot><tr class='ui-table__total'><th scope='row'>Total</th>" & _
			"<td class='ui-table__cell--center'>" & totInc & "</td><td class='ui-table__cell--center'>" & totProy & "</td></tr></tfoot>"

	else 'POR UN RESP

		Set rj = Exec(Cnn, "sp_cg_Responsables @Modo='id_funcion_id', @IdSeg=" & SqlStr(idResp))
		if not rj.eof then
			cantInc = getCantInc(cnn,rj(0),rj(1))
			cantProy = getCantProy(cnn,rj(0),rj(1))
			html = html & "<tr><td>" & UiTexto(getResp(cnn,rj(0))) & "</td>" & _
				"<td class='ui-table__cell--center'>" & UiDato(cantInc) & "</td>" & _
				"<td class='ui-table__cell--center'>" & UiDato(cantProy) & "</td></tr>"
			totalUno = CLng("0" & cantInc) + CLng("0" & cantProy)
			html = html & "</tbody><tfoot><tr class='ui-table__total'><th scope='row'>Total general</th>" & _
				"<td class='ui-table__cell--center' colspan='2'>" & totalUno & "</td></tr></tfoot>"
		else
			html = html & UiFilaVacia(3, "No se encontr&oacute; el responsable seleccionado.") & "</tbody>"
		end if

	end if

	' La tabla ya cerro su tbody (y tfoot): se cierra solo el resto
	html = html & "</table></div>"
	' [CONSULTA] Pie con el paginado (solo cuando se listan todos los responsables)
	If Trim(idResp & "") = "0" Or Trim(idResp & "") = "" Then
		Dim urlPagP, botonAntP, botonSigP
		urlPagP = "consultas.asp?sel=5&xmenu=" & xmenu & "&xtmenu=" & xtmenu
		html = html & "<form id='formFiltros' class='form-oculto' method='post' action='" & urlPagP & "'><input type='hidden' name='uau' value='" & Server.HTMLEncode(idResp & "") & "'></form>"
		If cantRegMoverP > 0 Then
			botonAntP = UiBoton("Anterior", "secondary", "irConsulta('" & urlPagP & "&cantregmover=" & (cantRegMoverP - cantRegMostrarP) & "');", True, "")
		Else
			botonAntP = UiBoton("Anterior", "secondary", "", False, "Primera p&aacute;gina")
		End If
		If cantRegMoverP + idxMostradasP < idxFilaP Then
			botonSigP = UiBoton("Siguiente", "secondary", "irConsulta('" & urlPagP & "&cantregmover=" & (cantRegMoverP + cantRegMostrarP) & "');", True, "")
		Else
			botonSigP = UiBoton("Siguiente", "secondary", "", False, "&Uacute;ltima p&aacute;gina")
		End If
		html = html & UiPaginacionDe(cantRegMoverP + 1, cantRegMoverP + idxMostradasP, idxFilaP, botonAntP, botonSigP)
	End If
	html = html & UiCardCierra()

End If

html = html & UiPantallaCierra()

html = html & "<script type='text/javascript'>"
html = html & " function irConsulta(url){ var f = document.getElementById('formFiltros'); if (!f) return false; f.action = url; f.submit(); return false; } "
html = html & " function modificar(formu){ "
html = html & " document.getElementById(formu).submit();"
html = html & " alert('Descripci\u00f3n actualizada');} "
html = html & "</script>"

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
html = html & "    $( ""#fec"" ).datepicker();"
html = html & "  } );"
html = html & "  </script>"

html = html & formx04 & "</body>"
html = html & "</html>"

response.write html

Function armaCombo(sql,cn,clave,valor,texto,comboid)

	Set nrs = Exec(cn, sql)

	xhtml = xhtml & "<select class='selec' style='width:70%;' id='" & comboid & "' name='" & comboid & "'>"

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

	xhtml = xhtml & "</select>"

	armaCombo = xhtml

End Function

Function getResp(cn,xid)

	Set irs = Exec(cn, "sp_cg_Responsables @Modo='nombre_id', @IdSeg=" & SqlStr(xid))

	xResp = ""
	if not irs.eof then
		xResp = irs(0)
	else
		xResp = "TODOS"
	end if
	getResp = xResp

End Function

Function getCantInc(cn,sid,stipo)

	if stipo = "TEC" then
		Set xi = Exec(cn, "sp_cg_Registro_Contar @Tabla='incidentes', @Campo='respon_tec', @IdRespon=" & SqlStr(sid))
	elseif stipo = "FUNC" then
		Set xi = Exec(cn, "sp_cg_Registro_Contar @Tabla='incidentes', @Campo='respon_gestion', @IdRespon=" & SqlStr(sid))
	end if

	sCantInc = ""
	if not xi.eof then sCantInc = xi(0)

	getCantInc = sCantInc

End Function

Function getCantProy(cn,sid,stipo)

	if stipo = "TEC" then
		Set xp = Exec(cn, "sp_cg_Registro_Contar @Tabla='proyectos', @Campo='respon_tec', @IdRespon=" & SqlStr(sid))
	elseif stipo = "FUNC" then
		Set xp = Exec(cn, "sp_cg_Registro_Contar @Tabla='proyectos', @Campo='respon_gestion', @IdRespon=" & SqlStr(sid))
	end if

	sCantProy = ""
	if not xp.eof then sCantProy = xp(0)

	getCantProy = sCantProy

End Function

%>