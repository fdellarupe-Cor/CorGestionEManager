<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/ui_componentes.inc" -->
<%

'[UI-ANT] ' ===== Bloque anterior: reemplazado por el bloque [UI] que sigue. Se borra entero buscando [UI-ANT] =====
'[UI-ANT] html = "<html>"
'[UI-ANT] html = html & "<head>"

'[UI-ANT] html = html & "<title>Cor Gestion</title>"
'[UI-ANT] html = html & "<meta charset='utf-8' />"
'[UI-ANT] html = html & "<meta name='viewport' content='width=device-width, initial-scale=1' />"
'[UI-ANT] html = html & "<link rel='stylesheet' href='assets/css/main.css' />"
'[UI-ANT] html = html & "<link rel='stylesheet' href='assets/js/jquery-ui.css'>"

'[UI-ANT] html = html & "</head>"
'[UI-ANT] html = html & "<body>"

'[UI-ANT] Set Cnn = Server.CreateObject("ADODB.Connection")
'[UI-ANT] Cnn.open scon1

'[UI-ANT] XXX = SYSLOGON(cnn,"MENU_CONSULTA")

'[UI-ANT] If request.form("sel") & "" = "" Then
	'[UI-ANT] sel = request.querystring("sel")
'[UI-ANT] Else
	'[UI-ANT] sel = request.form("sel")
'[UI-ANT] End If

'[UI-ANT] If sel & "" = "" Then sel = 0

'[UI-ANT] xmenu = request.querystring("xmenu")
'[UI-ANT] xtmenu = request.querystring("xtmenu")

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

'[UI-ANT] html = html & "<td><h2>REPORTES</h2></td>"

'[UI-ANT] formx04 = "<form id='formx04' method='post' action='index.asp'></form>"
'[UI-ANT] botonx04 = "<input class='bot' type='button' value='VOLVER' onclick=""document.getElementById('formx04').submit();"">"
'[UI-ANT] html = html & "<td>" & botonx04 & "</td>"

'[UI-ANT] html = html & "</tr>"

'[UI-ANT] html = html & "</table>"

'[UI-ANT] html = html & "</div><br>"

	'[UI-ANT] ' --- INICIO NORMALIZACION: FORMULARIO GENERAR REPORTE ---
	'[UI-ANT] ' Contenedor centrado y compacto
	'[UI-ANT] 'html = html & "<div style='height:75%;width:100%;'><br>"
	'[UI-ANT] html = html & "<div style='max-width: 600px; margin: 2rem auto;'>"

	'[UI-ANT] html = html & "<form id='xForm' method='post' action='reportesDet.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"

	'[UI-ANT] 'html = html & "<table class='tabl' style='height:50%;width:40%;' align=center cellspacing=0>"
	'[UI-ANT] html = html & "<div class='bloque-datos'>"

	'[UI-ANT] ' --- ENCABEZADO ---
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' style='font-size:12pt;' colspan=2 align=center valign=center><b>GENERAR REPORTE DE INCIDENTES</b></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='bloque-header'>GENERAR REPORTE DE INCIDENTES</div>"
	
	'[UI-ANT] html = html & "<div class='bloque-body'>"

	'[UI-ANT] 'ESPACIO
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' colspan=2 >&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"
	
	'[UI-ANT] ' --- INICIO GRILLA ---
	'[UI-ANT] html = html & "<div class='grid-datos-doble'>"

	'[UI-ANT] 'CLIENTE
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center>CLIENTE</td>"
	'[UI-ANT] 'html = html & "<td class='fil' valign=center>"
	'[UI-ANT] 'html = html & armaCombo("select id,descripcion from clientes",cnn,request.form("cli"),0,1,"cli")
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>CLIENTE</div>"
	'[UI-ANT] html = html & "<div class='dato-value'>"
	'[UI-ANT] html = html & armaCombo("sp_cg_Catalogo_Listar @Catalogo='clientes'",cnn,request.form("cli"),0,1,"cli")
	'[UI-ANT] html = html & "</select></div>" ' <--- PARCHE: Forzamos el cierre del combo

	'[UI-ANT] 'INCIDENTES (Mantenemos la secci?n comentada original)
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center>INCIDENTES</td>"
	'[UI-ANT] 'html = html & "<td class='fil' valign=center>"
	'[UI-ANT] 'html = html & "<select class='selec' id='inc' name='inc' style='width:20%;'>"
	'[UI-ANT] 'html = html & "<option value='0'>NO</option>"
	'[UI-ANT] 'html = html & "<option value='1'>SI</option>"
	'[UI-ANT] 'html = html & "</select>"
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"

	'[UI-ANT] 'PROYECTOS (Mantenemos la secci?n comentada original)
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center>PROYECTOS</td>"
	'[UI-ANT] 'html = html & "<td class='fil' valign=center>"
	'[UI-ANT] 'html = html & "<select class='selec' id='proy' name='proy' style='width:20%;'>"
	'[UI-ANT] 'html = html & "<option value='0'>NO</option>"
	'[UI-ANT] 'html = html & "<option value='1'>SI</option>"
	'[UI-ANT] 'html = html & "</select>"
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"

	'[UI-ANT] 'APLs (Mantenemos la secci?n comentada original)
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center>APLs</td>"
	'[UI-ANT] 'html = html & "<td class='fil' valign=center>"
	'[UI-ANT] 'html = html & "<select class='selec' id='apl' name='apl' style='width:20%;'>"
	'[UI-ANT] 'html = html & "<option value='0'>NO</option>"
	'[UI-ANT] 'html = html & "<option value='1'>SI</option>"
	'[UI-ANT] 'html = html & "</select>"
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"

	'[UI-ANT] 'FECHA DESDE
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center>FECHA DESDE</td>"
	'[UI-ANT] 'html = html & "<td class='fil' valign=center>"
	'[UI-ANT] 'html = html & "<input class='inpu' id='fecd' name='fecd' style='width:40%;' type='text' />"
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>FECHA DESDE</div>"
	'[UI-ANT] html = html & "<div><input type='text' id='fecd' name='fecd' class='input-moderno'></div>"

	'[UI-ANT] 'FECHA HASTA
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' align=center valign=center>FECHA HASTA</td>"
	'[UI-ANT] 'html = html & "<td class='fil' valign=center>"
	'[UI-ANT] 'html = html & "<input class='inpu' id='fech' name='fech' style='width:40%;' type='text' />"
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>FECHA HASTA</div>"
	'[UI-ANT] html = html & "<div><input type='text' id='fech' name='fech' class='input-moderno'></div>"

	'[UI-ANT] 'ESPACIO
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' colspan=2 >&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"

	'[UI-ANT] ' --- FIN GRILLA ---
	'[UI-ANT] html = html & "</div>"

	'[UI-ANT] ' --- BOTON CONSULTAR ---
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td style='height:10%;width:50%;' colspan=2 align=center>"
	'[UI-ANT] html = html & "<hr style='border: 0; border-top: 1px solid #e2e8f0; margin: 1.5rem 0;'>"
	'[UI-ANT] html = html & "<div style='text-align: center; margin-bottom: 0.5rem;'>"

	'[UI-ANT] if getPermisos(xmenu,xtmenu,"CONS") = 1 then
		'[UI-ANT] 'html = html & "<input class='bot' type='button' value='Consultar' onclick=""document.getElementById('xForm').submit();"">"
		'[UI-ANT] html = html & "<input type='button' class='btn-principal' value='Consultar' onclick=""document.getElementById('xForm').submit();"">"
	'[UI-ANT] else
		'[UI-ANT] 'html = html & "<input class='botsp' type='button' value='Consultar' onclick=""alert('No tiene permiso para consultar.');"">"
		'[UI-ANT] ' Usamos la clase de deshabilitado para que mantenga la coherencia visual
		'[UI-ANT] html = html & "<input type='button' class='btn-principal btn-accion-deshabilitado' value='Consultar' onclick=""alert('No tiene permiso para consultar.');"">"
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

'[UI-ANT] html = html & "    $( ""#fecd"" ).datepicker();"
'[UI-ANT] html = html & "    $( ""#fech"" ).datepicker();"

'[UI-ANT] html = html & "  } );"
'[UI-ANT] html = html & "  </script>"

'[UI-ANT] html = html & formx04 & "</body>"
'[UI-ANT] html = html & "</html>"

' =========================================================================
' [UI] VISTA con el sistema visual (IncFile/ui_componentes.inc)
' El bloque anterior completo queda arriba, comentado con '[UI-ANT].
' Formulario (destino y nombres de campos) y permiso: los mismos de antes.
' =========================================================================
Dim accTop

' Encabezado comun del sistema: el propio no cargaba el CSS ni las tipografias, y declaraba
' UTF-8 aunque la pagina sale en ANSI (las tildes de la base se podian ver rotas)
html = headhtml()
html = html & "<body>"

Set Cnn = Server.CreateObject("ADODB.Connection")
Cnn.open scon1

XXX = SYSLOGON(cnn,"MENU_CONSULTA")

If request.form("sel") & "" = "" Then
	sel = request.querystring("sel")
Else
	sel = request.form("sel")
End If
If sel & "" = "" Then sel = 0

xmenu = request.querystring("xmenu")
xtmenu = request.querystring("xtmenu")

' DEPURADO [UI]: la cabecera propia (logo e imagen armados a mano) se reemplaza por cabecera()
html = html & "<div id='wrapper' role='banner'>" & cabecera() & "</div>"

formx04 = "<form id='formx04' class='form-oculto' method='post' action='index.asp'></form>"
accTop = UiBoton("Volver", "secondary", "document.getElementById('formx04').submit();", True, "")
html = html & UiPantallaAbre("Reportes", accTop, "")

' [UI] Tipos de reporte, con la misma barra de pestanas que Consultas y Estadisticas (ui-tabbar--pildoras):
' cada pestana es un enlace (reportes.asp?rep=...) y el servidor dibuja solo el formulario elegido, como hijo
' directo del contenedor (misma distancia entre la tarjeta y el boton que en Consultas).
' Los dos nuevos abren reportesAgenda.asp y reportesPendientes.asp (pantallas a conectar por el companero).
repTipo = LCase(request.querystring("rep") & "")
If repTipo <> "agenda" And repTipo <> "pendientes" Then repTipo = "cliente"
puedeCons = (getPermisosAtrm("CG_REPORTES", "CONS") = 1)
html = html & "<nav class='ui-tabbar ui-tabbar--pildoras' aria-label='Tipo de reporte'>" & RepTab("cliente", "Por cliente", repTipo) & RepTab("agenda", "Agenda por usuario", repTipo) & RepTab("pendientes", "Pendientes por usuario", repTipo) & "</nav>"

If repTipo = "cliente" Then
	' --- Por cliente (el formulario de siempre, sin cambios) ---
	html = html & "<form id='xForm' method='post' action='reportesDet.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
	html = html & UiCardAbre("Reporte de incidentes", "") & UiFormGridAbre()
	html = html & UiCampo("cli", "Cliente", UiSelectRs("cli", Cnn.Execute("sp_cg_Catalogo_Listar @Catalogo='clientes'"), request.form("cli"), "Todos los clientes", "0", ""), False, True)
	html = html & UiCampo("fecd", "Desde", UiInputTexto("fecd", False, "autocomplete='off' placeholder='dd/mm/aaaa'"), False, False)
	html = html & UiCampo("fech", "Hasta", UiInputTexto("fech", False, "autocomplete='off' placeholder='dd/mm/aaaa'"), False, False)
	html = html & UiFormGridCierra() & UiCardCierra()
	html = html & UiFormAcciones(UiBotonEnviar("Generar reporte", getPermisosAtrm("CG_REPORTES","CONS") = 1, "No tiene permiso para consultar."))
	html = html & "</form>"
ElseIf repTipo = "agenda" Then
	' --- Agenda por usuario ---
	html = html & "<form id='formAgenda' method='post' action='reportesAgenda.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
	html = html & UiCardAbre("Reporte de agenda por usuario", "") & UiFormGridAbre()
	html = html & UiCampo("agusu", "Usuario", RepSelectUsuarios("agusu"), False, True)
	html = html & UiCampo("agrol", "Participaci&oacute;n", "<select class='ui-select' id='agrol' name='agrol'><option value='A'>Responsable o participante</option><option value='R'>Solo como responsable</option><option value='P'>Solo como participante</option></select>", False, False)
	html = html & UiCampo("agest", "Estado de la reuni&oacute;n", "<select class='ui-select' id='agest' name='agest'><option value='T'>Todas</option><option value='PENDIENTE'>Pendientes</option><option value='POSTERGADA'>Postergadas</option><option value='REALIZADA'>Realizadas</option></select>", False, False)
	html = html & UiCampo("agfecd", "Desde", UiInputTexto("agfecd", False, "autocomplete='off' placeholder='dd/mm/aaaa'"), False, False)
	html = html & UiCampo("agfech", "Hasta", UiInputTexto("agfech", False, "autocomplete='off' placeholder='dd/mm/aaaa'"), False, False)
	html = html & UiCampo("agrec", "Recordatorios", "<label class='ui-opcion'><input class='ui-check' type='checkbox' id='agrec' name='agrec' value='1' checked><span>Incluir recordatorios pendientes</span></label>", False, True)
	html = html & UiFormGridCierra() & UiCardCierra()
	html = html & UiFormAcciones(UiBotonEnviar("Generar reporte", puedeCons, "No tiene permiso para consultar."))
	html = html & "</form>"
Else
	' --- Pendientes por usuario ---
	html = html & "<form id='formPendientes' method='post' action='reportesPendientes.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
	html = html & UiCardAbre("Reporte de pendientes por usuario", "") & UiFormGridAbre()
	html = html & UiCampo("pusu", "Usuario", RepSelectUsuarios("pusu"), False, True)
	html = html & UiCampo("prol", "Rol del usuario", "<select class='ui-select' id='prol' name='prol'><option value='A'>Cualquier rol</option><option value='G'>Responsable de gesti&oacute;n</option><option value='T'>Responsable t&eacute;cnico</option></select>", False, False)
	html = html & UiCampo("ptipo", "Tipo", "<select class='ui-select' id='ptipo' name='ptipo'><option value='A'>Incidentes y proyectos</option><option value='I'>Solo incidentes</option><option value='P'>Solo proyectos</option></select>", False, False)
	html = html & UiCampo("pcli", "Cliente", UiSelectRs("pcli", Cnn.Execute("sp_cg_Catalogo_Listar @Catalogo='clientes'"), "", "Todos los clientes", "0", ""), False, False)
	html = html & UiCampo("pmanos", "En manos de", "<select class='ui-select' id='pmanos' name='pmanos'><option value='T'>Todos</option><option value='FUNCIONAL'>Funcional</option><option value='TECNICO'>T&eacute;cnico</option><option value='CLIENTE'>Cliente</option></select>", False, False)
	html = html & UiCampo("pdias", "Sin avance hace m&aacute;s de (d&iacute;as)", UiInputTexto("pdias", False, "type='number' min='1' step='1' inputmode='numeric' placeholder='Cualquier antig&uuml;edad'"), False, False)
	html = html & UiFormGridCierra() & UiCardCierra()
	html = html & UiFormAcciones(UiBotonEnviar("Generar reporte", puedeCons, "No tiene permiso para consultar."))
	html = html & "</form>"
End If

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
html = html & "$('#agfecd, #agfech').datepicker();"
' [UI] calendario tambien en las fechas del reporte de agenda
html = html & "    $( ""#fecd"" ).datepicker();"
html = html & "    $( ""#fech"" ).datepicker();"
html = html & "  } );"
html = html & "  </script>"

html = html & formx04 & "</body>"
html = html & "</html>"

response.write html

Function armaCombo(sql,cn,clave,valor,texto,comboid)

	Set nrs = cn.Execute(sql)

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

	armaCombo = xhtml & "</select>"

End Function

Function getResp(cn,xid)

	isql = "sp_cg_Responsables @Modo='nombre_id', @IdSeg=" & SqlStr(xid)
	Set irs = cn.Execute(isql)

	xResp = ""
	if not irs.eof then
		xResp = irs(0)
	else
		xResp = "TODOS"
	end if
	getResp = xResp

End Function


' [UI] Pestana de la barra de tipos de reporte (misma marca de activa que Consultas: is-active y aria-current)
Function RepTab(valor, texto, actual)
	Dim url
	url = "reportes.asp?rep=" & valor & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu
	If valor = actual Then
		RepTab = "<a class='ui-tabbar__link is-active' aria-current='page' href='" & url & "'>" & texto & "</a>"
	Else
		RepTab = "<a class='ui-tabbar__link' href='" & url & "'>" & texto & "</a>"
	End If
End Function
' [UI] Usuarios activos (login como valor, nombre y apellido como texto)
Function RepSelectUsuarios(id)
	Dim rsU, h
	h = "<select class='ui-select' id='" & id & "' name='" & id & "'><option value=''>Todos los usuarios</option>"
	Set rsU = Cnn.Execute("sp_cg_Responsables @Modo='listar_activos_sin', @ExcluirId=''")
	Do While Not rsU.EOF
		h = h & "<option value='" & Server.HTMLEncode(rsU(0) & "") & "'>" & Server.HTMLEncode(Trim(rsU(1) & " " & rsU(2))) & "</option>"
		rsU.MoveNext
	Loop
	rsU.Close
	RepSelectUsuarios = h & "</select>"
End Function
%>