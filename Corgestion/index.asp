<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/ui_componentes.inc" -->
<%
'[UI-ANT] ' ===== Cuerpo anterior completo: reemplazado por el bloque [UI] que sigue. Se borra entero buscando [UI-ANT] =====
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

'[UI-ANT] XXX = MAILREC()

'[UI-ANT] 'html = "<html>"
'[UI-ANT] 'html = html & "<head>"

'[UI-ANT] 'html = html & "<title>Cor Gestion</title>"
'[UI-ANT] 'html = html & "<meta http-equiv='Content-Type' content='text/html;charset=ISO-8859-1'>"
'[UI-ANT] 'html = html & "<meta name='viewport' content='width=device-width, initial-scale=1' />"
'[UI-ANT] 'html = html & "<link rel='stylesheet' href='assets/css/main.css' />"
'[UI-ANT] 'html = html & "<script src='assets/js/jquery.min.js'></script>"
'[UI-ANT] 'html = html & "<script src='assets/js/skel.min.js'></script>"
'[UI-ANT] 'html = html & "<script src='assets/js/util.js'></script>"
'[UI-ANT] 'html = html & "<script src='assets/js/main.js'></script>"
'[UI-ANT] 'html = html & "<script type='text/javascript'>"
'[UI-ANT] 'html = html & "$(document).ready(function(){"
'[UI-ANT] ''html = html & "$('#body').fadeIn(1500);"
'[UI-ANT] 'html = html & "});"
'[UI-ANT] 'html = html & "</script>"
'[UI-ANT] 'html = html & "</head>"

'[UI-ANT] html=headhtml()
'[UI-ANT] html = html & "<body id='body' name='body' style='height:100%;width:100%;' onload='mueveReloj();'>"

'[UI-ANT] 'SCRIPTS SLIDER
'[UI-ANT] 'html = html & "<script type='text/javascript' src='assets/js/jquery-1.9.1.min.js'></script>"
'[UI-ANT] 'html = html & "<script type='text/javascript' src='assets/js/jssor.core.js'></script>"
'[UI-ANT] 'html = html & "<script type='text/javascript' src='assets/js/jssor.utils.js'></script>"
'[UI-ANT] 'html = html & "<script type='text/javascript' src='assets/js/jssor.slider.mini.js'></script>"

'[UI-ANT] 'html = html & "<script>"

'[UI-ANT] 'html = html & "var _SlideshowTransitions = [ "
'[UI-ANT] 'html = html & " {$Duration: 700, $Opacity: 2, $Brother: { $Duration: 1000, $Opacity: 2} } "
'[UI-ANT] 'html = html & "]; "
'[UI-ANT] 'html = html & "</script>"

'[UI-ANT] 'html = html & "<script>"
'[UI-ANT] 'html = html & "jQuery(document).ready(function ($) { "
'[UI-ANT] 'html = html & "var options = { "
'[UI-ANT] 'html = html & "$FillMode: 3, "                                     '[Optional] The way to fill image in slide, 0 stretch, 1 contain (keep aspect ratio and put all inside slide), 2 cover (keep aspect ratio and cover whole slide), 4 actuall size, default value is 0
'[UI-ANT] 'html = html & "$DragOrientation: 1, "                              '[Optional] Orientation to drag slide, 0 no drag, 1 horizental, 2 vertical, 3 either, default value is 1 (Note that the $DragOrientation should be the same as $PlayOrientation when $DisplayPieces is greater than 1, or parking position is not 0)
'[UI-ANT] 'html = html & "$AutoPlay: true, "                                  '[Optional] Whether to auto play, to enable slideshow, this option must be set to true, default value is false
'[UI-ANT] 'html = html & "$AutoPlayInterval: 700, "                          '[Optional] Interval (in milliseconds) to go for next slide since the previous stopped if the slider is auto playing, default value is 3000
'[UI-ANT] 'html = html & "$SlideshowOptions: { "                              '[Optional] Options to specify and enable slideshow or not
'[UI-ANT] 'html = html & "$Class: $JssorSlideshowRunner$, "                   '[Required] Class to create instance of slideshow
'[UI-ANT] 'html = html & "$Transitions: _SlideshowTransitions, "              '[Required] An array of slideshow transitions to play slideshow
'[UI-ANT] 'html = html & "$TransitionsOrder: 1, "                             '[Optional] The way to choose transition to play slide, 1 Sequence, 0 Random
'[UI-ANT] 'html = html & "$ShowLink: true "                                   '[Optional] Whether to bring slide link on top of the slider when slideshow is running, default value is false
'[UI-ANT] 'html = html & "}"
'[UI-ANT] 'html = html & "};"

'[UI-ANT] 'html = html & "var jssor_slider1 = new $JssorSlider$('slider1_container', options);"
'[UI-ANT] 'html = html & "});"
'[UI-ANT] 'html = html & "</script>"

'[UI-ANT] html = html & "<div id='wrapper' >"

'[UI-ANT] '''''''''''''''''TEST JQUERY''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'[UI-ANT] 'html = html & "<form id='test' name='test' style='display:none;' method='post' action='test.asp'></form>"
'[UI-ANT] 'html = html & "<input type='button' value='JQuery' onclick='testeo();'>"
'[UI-ANT] ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'[UI-ANT] html = html & "<header id='header' style='background:#0e1949;padding:1rem 0;'>"

'[UI-ANT] html = html & "<div >"

'[UI-ANT] html = html & "<table style='width:100%;'><tr style='display:flex;padding: 0 1rem;'>"

'[UI-ANT] html = html & "<td style='width:75%;' valign=center><a href='index.asp' class='logo' >"
'[UI-ANT] html = html & "<img src='images/iso_cor.png' style='width:6rem;' >"
'[UI-ANT] html = html & "<font style='font-size:15pt;color:white;'><b>Cor Gestion</b> | Sistema Administrativo</font></a></td>"
'[UI-ANT] html = html & "<td style='width:25%;display:flex;align-items:center;flex-direction:row-reverse;' align='right'>"
'[UI-ANT] 'html = html & "<form name='form_reloj'>"
'[UI-ANT] 'html = html & "<input class='reloj' type='text' name='reloj' size='10' style='width:8rem;color:#000000;font-family:sans-serif;font-size:18pt;background:#FFFFFF;'>"
'[UI-ANT] 'html = html & "</form>"
'[UI-ANT] 'html = html & "<nav>"
'[UI-ANT] 'html = html & "<ul>"
'[UI-ANT] 'html = html & "<li><a href='#menu'></a></li>"
'[UI-ANT] 'html = html & "</ul>"
'[UI-ANT] 'html = html & "</nav>"

'[UI-ANT] ' Reemplazamos la lista vacía por un botón hamburguesa con icono SVG moderno
'[UI-ANT] html = html & "<nav>"
'[UI-ANT] html = html & "<a href='#menu' class='btn-hamburguesa' aria-label='Abrir menú'>"
'[UI-ANT] html = html & "<svg xmlns='http://www.w3.org/2000/svg' fill='none' viewBox='0 0 24 24' stroke='currentColor' stroke-width='2'>"
'[UI-ANT] html = html & "<path stroke-linecap='round' stroke-linejoin='round' d='M4 6h16M4 12h16M4 18h16' />"
'[UI-ANT] html = html & "</svg>"
'[UI-ANT] html = html & "</a>"
'[UI-ANT] html = html & "</nav>"
'[UI-ANT] html = html & "</td>"

'[UI-ANT] html = html & "</tr></table>"

'[UI-ANT] html = html & "</div>"

'[UI-ANT] 'html = html & "<nav>"
'[UI-ANT] 'html = html & "<ul>"
'[UI-ANT] 'html = html & "<li><a href='#menu'></a></li>"
'[UI-ANT] 'html = html & "</ul>"
'[UI-ANT] 'html = html & "</nav>"

'[UI-ANT] html = html & "</header>"

'[UI-ANT] html = html & "</div>"

'[UI-ANT] 'ARMA MENU DE PARAMETROS
'[UI-ANT] html = html & getMenuParam()

'[UI-ANT] html = html & "<div id='main' >"

'[UI-ANT] html = html & "<div class='titulo-seccion'>"
'[UI-ANT] 'html = html & "<table style='height:100%;width:100%;'><tr>"
'[UI-ANT] 'html = html & "<td style='height:100%;width:60%;'><font style='margin-left:20px;font-family:Serif;color:#000000;font-size:25pt;letter-spacing:4px;' >"
'[UI-ANT] html = html & "<h1>"
'[UI-ANT] 'html = html & "Sistema Administrativo</font>"
'[UI-ANT] html = html & "SECCIONES</h1>"
'[UI-ANT] 'html = html & "SISTEMA ADMINISTRATIVO</font></td>"
'[UI-ANT] 'html = html & "</tr></table>"
'[UI-ANT] html = html & "</div>"

'[UI-ANT] 'ARMA MENU PRINCIPAL
'[UI-ANT] html = html & getMenuPrinc()

'[UI-ANT] 'html = html & "<div id='slider1_container' style='position:relative;width:1300px;height:200px;background-color:#FFFFFF;margin-left:10px;'>"

'[UI-ANT] <!-- Loading Screen -->
'[UI-ANT] 'html = html & "<div u='loading' style='position: absolute; top: 0px; left: 0px;'>"
'[UI-ANT] 'html = html & "<div style='filter:alpha(opacity=70);opacity:0.7;position:absolute;display:block;background-color:#FFFFFF;top:0px;left:0px;width:1300px;height:200px;'>"
'[UI-ANT] 'html = html & "</div>"
'[UI-ANT] 'html = html & "<div style='position:absolute;display:block;background:url(../img/loading.gif) no-repeat center center;top:0px;left:0px;width:1300px;height:200px;'>"
'[UI-ANT] 'html = html & "</div>"
'[UI-ANT] 'html = html & "</div>"

'[UI-ANT] <!-- Slides Container -->
'[UI-ANT] 'html = html & "<div u='slides' style='cursor:move;position:absolute;left:0px;top:0px;width:1300px;height:200px;overflow: hidden;'>"

'[UI-ANT] 'NUEVO SLIDER
'[UI-ANT] html = html & "<div class='contenedor-marquee'>"
'[UI-ANT] html = html & "<div class='track-marquee'>"

'[UI-ANT] Set rc = Exec(Cnn, "sp_cg_Utilidades @Modo='listar_logos', @Id=37")

'[UI-ANT] 'response.write sql

'[UI-ANT] ' Armamos el HTML de los logos en una variable temporal
'[UI-ANT] logosHTML = ""
'[UI-ANT] do while not rc.eof
'[UI-ANT] 'html = html & "<div>"
'[UI-ANT] 'html = html & "<a u=image href='incidentes.asp?xmenu=9&xtmenu=INICIO&ord=&cli=" & rc(0) & "'>"
'[UI-ANT] 'html = html & "<img src='../Corgestion/images/" & rc(1) & "'  width='100px' height='100px' />"
'[UI-ANT] 'html = html & "</a>"
'[UI-ANT] 'html = html & "</div>"
	'[UI-ANT] logosHTML = logosHTML & "<a class='item-marquee' href='incidentes.asp?xmenu=9&xtmenu=INICIO&ord=&cli=" & rc(0) & "'>"
	'[UI-ANT] logosHTML = logosHTML & "<img src='../Corgestion/images/" & rc(1) & "' alt='Cliente' />"
	'[UI-ANT] logosHTML = logosHTML & "</a>"
	'[UI-ANT] rc.movenext
'[UI-ANT] loop

'[UI-ANT] ' Para que la animación CSS sea un ciclo perfecto e infinito sin cortes,
'[UI-ANT] ' concatenamos la lista de logos DOS VECES en el contenedor.
'[UI-ANT] html = html & logosHTML & logosHTML

'[UI-ANT] html = html & "</div>" ' Cierre de track-marquee
'[UI-ANT] html = html & "</div>" ' Cierre de contenedor-marquee

'[UI-ANT] ' Footer moderno activado (con clase limpia y año automático)
'[UI-ANT] anioActual = Year(Now())
'[UI-ANT] html = html & "<div class='footer-legal'>"
'[UI-ANT] html = html & "&copy; " & anioActual & " Cor Consulting S.R.L. Todos los derechos reservados."
'[UI-ANT] html = html & "</div>"

'[UI-ANT] 'html = html & "<div style='width:100%;heith:30px;'>"
'[UI-ANT] 'html = html & "&nbsp;"
'[UI-ANT] 'html = html & "</div>"

'[UI-ANT] 'html = html & "<div style='width:100%;heith:30px;' align=center>"
'[UI-ANT] 'html = html & "<font style='font-size:20pt;color:#000000;'>2018 © Cor Consulting S.R.L.</font>"
'[UI-ANT] 'html = html & "</div>"


'[UI-ANT] html = html & "</div>" ' Cierre de main


'[UI-ANT] html = html & "<script language='JavaScript'> "
'[UI-ANT] html = html & "function mueveReloj(){ "
'[UI-ANT] html = html & " momentoActual = new Date(); "
'[UI-ANT] html = html & " hora = momentoActual.getHours();"
'[UI-ANT] html = html & " minuto = momentoActual.getMinutes(); "
'[UI-ANT] html = html & " segundo = momentoActual.getSeconds(); "
'[UI-ANT] html = html & "	str_segundo = new String (segundo); "
'[UI-ANT] html = html & " if (str_segundo.length == 1){ "
'[UI-ANT] html = html & " segundo = '0' + segundo}; "
'[UI-ANT] html = html & " str_minuto = new String (minuto); "
'[UI-ANT] html = html & " if (str_minuto.length == 1){ "
'[UI-ANT] html = html & " minuto = '0' + minuto } ;"
'[UI-ANT] html = html & " str_hora = new String (hora); "
'[UI-ANT] html = html & " if (str_hora.length == 1) {"
'[UI-ANT] html = html & " hora = '0' + hora} ; "
'[UI-ANT] html = html & " horaImprimible = hora + ' : ' + minuto + ' : ' + segundo; "
'[UI-ANT] html = html & " document.form_reloj.reloj.value = horaImprimible; "
'[UI-ANT] html = html & " setTimeout('mueveReloj()',1000); "
'[UI-ANT] html = html & "} "

'[UI-ANT] html = html & "function testeo(){ "
'[UI-ANT] html = html & " var conf = confirm('NO ACEPTES ! ! !'); "
'[UI-ANT] html = html & " if (conf) { "
'[UI-ANT] html = html & " test.submit(); "
'[UI-ANT] html = html & "}"
'[UI-ANT] html = html & "} "
'[UI-ANT] html = html & "</script> "

'[UI-ANT] html = html & "</body>"
'[UI-ANT] html = html & "</html>"

'[UI-ANT] response.write html

' =========================================================================
' [UI] PORTADA con el sistema visual (IncFile/ui_componentes.inc)
' El cuerpo anterior completo queda arriba, comentado con '[UI-ANT].
'
' - Cabecera comun (cabecera()) en lugar de la propia. El menu lateral de parametros
'   (getMenuParam) se mantiene y se abre con el boton "Parametros" de la barra.
' - Saludo en lugar del titulo "SECCIONES".
' - Secciones del usuario (sp_cg_Menu 'principal': mismas que mostraba getMenuPrinc, segun permisos)
'   agrupadas en Gestion y Analisis, cada una con una linea que explica para que sirve.
' - Clientes en grilla fija (antes carrusel en movimiento): logo o nombre, con acceso a sus
'   incidentes (el mismo enlace de antes) y a sus proyectos.
' - Se quitaron el reloj (escribia en un formulario que ya no existe: daba error cada segundo)
'   y la funcion testeo() (sin uso). MAILREC() se mantiene.
' =========================================================================
Dim usuario, nombreUsuario, primerNombre, partesNombre, saludo, textoHoy, DIAS, MESES, accTop
Dim funcionUsr, rsFun, rsParam, itemsParam, itemsSeg, enlaceParam, menuParamHtml, rsMenu, gestion, analisis, clave, item, nombres, rsCli, clientesHtml, cantCli, nombreCli, logo, tieneInc, tienePry, urlInc, urlPry, enlacesCli

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

XXX = MAILREC()

' Usuario: la sesion puede venir vacia (el index original no la usaba: getMenuPrinc identifica al usuario solo).
' Se toma la sesion y, si falta, UserLogon(), como hace detalleCliente.asp
usuario = Trim(Session("usuario") & "")
If usuario = "" Then usuario = Trim(UserLogon() & "")

' Funcion (FUNC / TEC): de la sesion o, si falta, de la base segun el login
funcionUsr = Trim(Session("FUNC") & "")
If funcionUsr = "" And usuario <> "" Then
	Set rsFun = Exec(Cnn, "sp_cg_Responsables @Modo='id_funcion_id', @IdSeg=" & SqlStr(usuario))
	If Not rsFun.eof Then funcionUsr = Trim(rsFun(1) & "")
	rsFun.Close
End If

' --- Saludo y fecha ---
DIAS = Array("Domingo", "Lunes", "Martes", "Mi&eacute;rcoles", "Jueves", "Viernes", "S&aacute;bado")
MESES = Array("enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre")
textoHoy = DIAS(Weekday(Date) - 1) & " " & Day(Date) & " de " & MESES(Month(Date) - 1)

nombreUsuario = Trim(getName(Cnn, "", "cor_responsables", usuario) & "")
primerNombre = ""
If nombreUsuario <> "" Then
	partesNombre = Split(nombreUsuario, " ")
	primerNombre = UiCapitalizar(partesNombre(0))
End If
If Hour(Now) < 12 Then
	saludo = "Buenos d&iacute;as"
ElseIf Hour(Now) < 20 Then
	saludo = "Buenas tardes"
Else
	saludo = "Buenas noches"
End If
If primerNombre <> "" Then saludo = saludo & ", " & UiTexto(primerNombre)

html = headhtml()
html = html & "<body>"
html = html & "<div id='wrapper' role='banner'>" & cabecera() & "</div>"

' --- Menu de Parametros y Seguridad: desplegable propio ---
' [UI] Antes era un panel lateral (getMenuParam) que abria main.js desde la cabecera propia; sin esa
' cabecera no se abria. Ahora se arma con las mismas opciones (sp_cg_Menu 'param': las habilitadas para
' el usuario, incluidas las de Seguridad) y se abre con un boton propio, sin depender de main.js.
itemsParam = ""
itemsSeg = ""
If usuario <> "" Then
	Set rsParam = Exec(Cnn, "sp_cg_Menu @Modo='param', @IdUsr=" & SqlStr(usuario) & ", @FuncUsr=" & SqlStr(funcionUsr) & ", @EsAdmin=1")
	Do While Not rsParam.eof
		enlaceParam = "<li><a class='ui-menu-param__link' href='" & Server.HTMLEncode(rsParam("modulo") & "") & "'>" & UiTexto(UiCapitalizar(rsParam("descripcion"))) & "</a></li>"
		If UCase(rsParam("tipo") & "") = "SEG" Then
			itemsSeg = itemsSeg & enlaceParam
		Else
			itemsParam = itemsParam & enlaceParam
		End If
		rsParam.movenext
	Loop
	rsParam.Close
End If

If itemsParam <> "" Or itemsSeg <> "" Then
	menuParamHtml = "<div class='ui-menu-param'>" & _
		"<button type='button' class='ui-button ui-button--light ui-button--sm' id='btnParam' aria-expanded='false' aria-controls='panelParam'>Par&aacute;metros <span aria-hidden='true'>&#9662;</span></button>" & _
		"<nav class='ui-menu-param__panel' id='panelParam' aria-label='Par&aacute;metros' hidden>"
	If itemsParam <> "" Then menuParamHtml = menuParamHtml & "<p class='ui-menu-param__grupo'>Par&aacute;metros</p><ul>" & itemsParam & "</ul>"
	If itemsSeg <> "" Then menuParamHtml = menuParamHtml & "<p class='ui-menu-param__grupo'>Seguridad</p><ul>" & itemsSeg & "</ul>"
	menuParamHtml = menuParamHtml & "</nav></div>"
	accTop = menuParamHtml
Else
	' Respaldo: si no se pudo identificar al usuario, queda el menu lateral original
	html = html & getMenuParam()
	accTop = "<a class='ui-button ui-button--light ui-button--sm' href='#menu'>Par&aacute;metros</a>"
End If

html = html & "<main id='contenido'>"
' Mismo ancho que los listados a los que lleva (incidentes, proyectos): 1240px, sin salto al navegar
html = html & "<div class='ui-app ui-topbar'><div class='ui-topbar__inner ui-topbar__inner--wide'>" & _
	"<div class='ui-topbar__top'>" & _
		"<div class='ui-topbar__ident'><h1 class='ui-topbar__title'>" & saludo & "</h1><p class='ui-topbar__sub'>" & textoHoy & "</p></div>" & _
		"<div class='ui-topbar__actions'>" & accTop & "</div>" & _
	"</div></div></div>"

html = html & "<div class='ui-app ui-page'><div class='ui-container ui-container--wide'>"

' =============================================================================
' SECCIONES (segun los permisos del usuario)
' =============================================================================
Set rsMenu = Exec(Cnn, "sp_cg_Menu @Modo='principal', @IdUsr=" & SqlStr(usuario) & ", @FuncUsr=" & SqlStr(funcionUsr))
gestion = ""
analisis = ""
tieneInc = False
tienePry = False
Do While Not rsMenu.eof
	clave = IdxClave(rsMenu("modulo"), rsMenu("descripcion"))
	If clave = "incidentes" Then tieneInc = True
	If clave = "proyectos" Then tienePry = True
	'[UI-ANT] item = "<li><a class='ui-acceso' href='" & Server.HTMLEncode(rsMenu("modulo") & "") & "'>" & _
	' [UI] La Agenda abre la vista semanal (desde ahi, 'Vista en lista' lleva a la lista)
	item = "<li><a class='ui-acceso' href='" & Server.HTMLEncode(IdxDestino(rsMenu("modulo"), clave)) & "'>" & _
		"<span class='ui-acceso__icono' aria-hidden='true'>" & IdxIcono(clave) & "</span>" & _
		"<span class='ui-acceso__texto'><span class='ui-acceso__nombre'>" & IdxNombre(clave, rsMenu("descripcion")) & "</span>" & _
		"<span class='ui-acceso__desc'>" & IdxDescripcion(clave) & "</span></span></a></li>"
	If clave = "consultas" Or clave = "estadisticas" Or clave = "reportes" Then
		analisis = analisis & item
	Else
		gestion = gestion & item
	End If
	rsMenu.movenext
Loop
rsMenu.Close

If gestion = "" And analisis = "" Then
	' Red de seguridad: si no se pudo identificar al usuario o la consulta no devolvio secciones,
	' se muestra el menu original (getMenuPrinc), que resuelve el usuario por su cuenta. Nadie se queda sin acceso.
	html = html & UiCardAbre("Secciones", "") & "<div class='ui-menu-original'>" & getMenuPrinc() & "</div>" & UiCardCierra()
	' Con el menu original no se sabe que secciones tiene: los accesos por cliente se ofrecen como antes (incidentes)
	tieneInc = True
Else
	html = html & "<div class='ui-portal'>"
	If gestion <> "" Then html = html & "<section class='ui-portal__grupo' aria-labelledby='tituloGestion'><h2 class='ui-portal__titulo' id='tituloGestion'>Gesti&oacute;n</h2><ul class='ui-portal__lista'>" & gestion & "</ul></section>"
	If analisis <> "" Then html = html & "<section class='ui-portal__grupo' aria-labelledby='tituloAnalisis'><h2 class='ui-portal__titulo' id='tituloAnalisis'>An&aacute;lisis</h2><ul class='ui-portal__lista'>" & analisis & "</ul></section>"
	html = html & "</div>"
End If

' =============================================================================
' ACCESO RAPIDO POR CLIENTE (mismos logos y enlace a incidentes de antes, sin carrusel)
' =============================================================================
' Nombres de los clientes, para mostrarlos si el logo no carga
Set nombres = Server.CreateObject("Scripting.Dictionary")
Set rsCli = Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='clientes'")
Do While Not rsCli.eof
	nombres(CStr(rsCli(0))) = rsCli(1) & ""
	rsCli.movenext
Loop
rsCli.Close

' Los accesos por cliente respetan los permisos: solo se ofrecen las secciones habilitadas
Set rc = Exec(Cnn, "sp_cg_Utilidades @Modo='listar_logos', @Id=37")
clientesHtml = ""
cantCli = 0
Do While Not rc.eof And (tieneInc Or tienePry)
	nombreCli = ""
	If nombres.Exists(CStr(rc(0))) Then nombreCli = nombres(CStr(rc(0)))
	If nombreCli = "" Then nombreCli = "Cliente " & rc(0)
	If Trim(rc(1) & "") <> "" Then
		' Si la imagen no carga, se oculta y queda el nombre
		logo = "<img class='ui-cliente__logo' src='../Corgestion/images/" & Server.HTMLEncode(rc(1) & "") & "' alt='' onerror=""this.hidden = true; this.nextElementSibling.hidden = false;"">" & _
			"<span class='ui-cliente__nombre' hidden>" & UiTexto(nombreCli) & "</span>"
	Else
		logo = "<span class='ui-cliente__nombre'>" & UiTexto(nombreCli) & "</span>"
	End If
	urlInc = "incidentes.asp?xmenu=9&xtmenu=INICIO&ord=&cli=" & rc(0)
	urlPry = "proyectos.asp?xmenu=10&xtmenu=INICIO&ord=&cli=" & rc(0)
	enlacesCli = ""
	If tieneInc Then enlacesCli = enlacesCli & "<a href='" & urlInc & "' aria-label='Incidentes de " & UiTexto(nombreCli) & "'>Incidentes</a>"
	If tienePry Then enlacesCli = enlacesCli & "<a href='" & urlPry & "' aria-label='Proyectos de " & UiTexto(nombreCli) & "'>Proyectos</a>"
	' El logo lleva al primer acceso disponible; con teclado se saltea (repite el enlace de abajo)
	If tieneInc Then
		clientesHtml = clientesHtml & "<li class='ui-cliente'><a class='ui-cliente__marca' href='" & urlInc & "' tabindex='-1' aria-hidden='true'>" & logo & "</a>"
	Else
		clientesHtml = clientesHtml & "<li class='ui-cliente'><a class='ui-cliente__marca' href='" & urlPry & "' tabindex='-1' aria-hidden='true'>" & logo & "</a>"
	End If
	clientesHtml = clientesHtml & "<div class='ui-cliente__nombre-sr ui-sr-only'>" & UiTexto(nombreCli) & "</div><div class='ui-cliente__links'>" & enlacesCli & "</div></li>"
	cantCli = cantCli + 1
	rc.movenext
Loop
rc.Close

If cantCli > 0 Then
	html = html & UiCardAbre("Acceso r&aacute;pido por cliente", "")
	html = html & "<ul class='ui-clientes'>" & clientesHtml & "</ul>"
	html = html & UiCardCierra()
End If

html = html & "<p class='ui-legal'>&copy; " & Year(Now()) & " Cor Consulting S.R.L. Todos los derechos reservados.</p>"
html = html & "</div></div>"
html = html & "</main>"

' Desplegable de Parametros: abre y cierra con el boton, Escape o un clic afuera
html = html & "<script>"
html = html & "(function () {"
html = html & "  var btn = document.getElementById('btnParam'), panel = document.getElementById('panelParam');"
html = html & "  if (!btn || !panel) return;"
html = html & "  function abrir(si, foco) { panel.hidden = !si; btn.setAttribute('aria-expanded', String(si)); if (si && foco) { var a = panel.querySelector('a'); if (a) a.focus(); } }"
html = html & "  btn.addEventListener('click', function (e) { e.stopPropagation(); abrir(panel.hidden, e.detail === 0); });"
html = html & "  document.addEventListener('click', function (e) { if (!panel.hidden && !panel.contains(e.target)) abrir(false); });"
html = html & "  document.addEventListener('keydown', function (e) { if (e.key === 'Escape' && !panel.hidden) { abrir(false); btn.focus(); } });"
html = html & "})();"
html = html & "</script>"

html = html & "<script src='assets/js/jquery.min.js'></script>"
html = html & "<script src='assets/js/skel.min.js'></script>"
html = html & "<script src='assets/js/util.js'></script>"
html = html & "<script src='assets/js/main.js'></script>"

html = html & "</body>"
html = html & "</html>"

Cnn.Close
Set Cnn = Nothing
response.write html


' =============================================================================
' FUNCIONES DE ESTA PANTALLA
' =============================================================================

' Clave de la seccion a partir de su pagina (o de su nombre si la pagina no se reconoce)
' [UI] Destino de cada acceso. La Agenda entra por la vista semanal: agenda.asp pasa a agenda_semana.asp
' conservando los parametros que traiga el menu (xmenu, xtmenu...). El resto, tal cual viene del menu.
Function IdxDestino(modulo, clave)
	Dim m
	m = Trim(modulo & "")
	If clave = "agenda" And LCase(Left(m, 10)) = "agenda.asp" Then m = "agenda_semana.asp" & Mid(m, 11)
	IdxDestino = m
End Function

Function IdxClave(modulo, descripcion)
	Dim m
	m = LCase(modulo & "") & " " & LCase(descripcion & "")
	If InStr(m, "incidente") > 0 Then
		IdxClave = "incidentes"
	ElseIf InStr(m, "proyecto") > 0 Then
		IdxClave = "proyectos"
	ElseIf InStr(m, "apl") > 0 Then
		IdxClave = "apl"
	ElseIf InStr(m, "consulta") > 0 Then
		IdxClave = "consultas"
	ElseIf InStr(m, "licencia") > 0 Then
		IdxClave = "licencias"
	ElseIf InStr(m, "agenda") > 0 Then
		IdxClave = "agenda"
	ElseIf InStr(m, "estadistica") > 0 Then
		IdxClave = "estadisticas"
	ElseIf InStr(m, "reporte") > 0 Then
		IdxClave = "reportes"
	ElseIf InStr(m, "todo") > 0 Or InStr(m, "para hacer") > 0 Then
		IdxClave = "todo"
	Else
		IdxClave = "otra"
	End If
End Function

' Nombre visible (las conocidas, bien escritas; el resto, como viene de la base)
Function IdxNombre(clave, descripcion)
	Select Case clave
		Case "incidentes"
			IdxNombre = "Incidentes"
		Case "proyectos"
			IdxNombre = "Proyectos"
		Case "apl"
			IdxNombre = "APL"
		Case "consultas"
			IdxNombre = "Consultas"
		Case "licencias"
			IdxNombre = "Licencias"
		Case "agenda"
			IdxNombre = "Agenda"
		Case "estadisticas"
			IdxNombre = "Estad&iacute;sticas"
		Case "reportes"
			IdxNombre = "Reportes"
		Case "todo"
			IdxNombre = "Para hacer"
		Case Else
			IdxNombre = UiTexto(UiCapitalizar(descripcion))
	End Select
End Function

Function IdxDescripcion(clave)
	Select Case clave
		Case "incidentes"
			IdxDescripcion = "Pedidos y errores reportados por los clientes"
		Case "proyectos"
			IdxDescripcion = "Desarrollos con horas, etapas y entrega"
		Case "apl"
			IdxDescripcion = "Paquetes enviados a cada cliente"
		Case "consultas"
			IdxDescripcion = "Busc&aacute; por cliente, estado o fecha"
		Case "licencias"
			IdxDescripcion = "Ausencias y licencias del equipo"
		Case "agenda"
			IdxDescripcion = "Recordatorios y reuniones"
		Case "estadisticas"
			IdxDescripcion = "Accesos y actividad por caso y usuario"
		Case "reportes"
			IdxDescripcion = "Res&uacute;menes por cliente y per&iacute;odo"
		Case "todo"
			IdxDescripcion = "Incidentes por atender"
		Case Else
			IdxDescripcion = "Abrir secci&oacute;n"
	End Select
End Function

' Iconos de linea (SVG en la pagina: no dependen de archivos ni fuentes externas)
Function IdxIcono(clave)
	Dim d
	Select Case clave
		Case "incidentes"
			d = "<path d='M7 18v-6a5 5 0 0 1 10 0v6'/><path d='M5 21h14a1 1 0 0 0 1-1v-1a2 2 0 0 0-2-2H6a2 2 0 0 0-2 2v1a1 1 0 0 0 1 1Z'/><path d='M12 4V2M4.2 6.2 3 5M19.8 6.2 21 5'/>"
		Case "proyectos"
			d = "<rect x='2' y='7' width='20' height='14' rx='2'/><path d='M16 7V5a2 2 0 0 0-2-2h-4a2 2 0 0 0-2 2v2'/><path d='M2 13h20'/>"
		Case "apl"
			d = "<path d='M21 8 12 3 3 8v8l9 5 9-5Z'/><path d='M3 8l9 5 9-5M12 13v8'/>"
		Case "consultas"
			d = "<circle cx='11' cy='11' r='7'/><path d='m21 21-4.3-4.3'/>"
		Case "licencias"
			d = "<path d='M12 3 14.4 5l3.1-.2.5 3.1L20.6 10 19 12.7l.6 3.1-3 1-1.5 2.7-2.9-1.3-2.9 1.3-1.5-2.7-3-1 .6-3.1L3.4 10 5.9 7.9l.5-3.1L9.6 5Z'/><path d='m9 12 2 2 4-4'/>"
		Case "agenda"
			d = "<rect x='3' y='4' width='18' height='18' rx='2'/><path d='M16 2v4M8 2v4M3 10h18'/><path d='M8 14h.01M12 14h.01M16 14h.01M8 18h.01M12 18h.01'/>"
		Case "estadisticas"
			d = "<path d='M3 3v18h18'/><path d='M8 17V11M13 17V7M18 17v-4'/>"
		Case "reportes"
			d = "<rect x='5' y='3' width='14' height='18' rx='2'/><path d='M9 3h6v3H9zM9 11h6M9 15h4'/>"
		Case "todo"
			d = "<path d='M9 11l2 2 4-4'/><rect x='3' y='3' width='18' height='18' rx='2'/>"
		Case Else
			d = "<rect x='3' y='3' width='7' height='7' rx='1'/><rect x='14' y='3' width='7' height='7' rx='1'/><rect x='3' y='14' width='7' height='7' rx='1'/><rect x='14' y='14' width='7' height='7' rx='1'/>"
	End Select
	IdxIcono = "<svg viewBox='0 0 24 24' width='26' height='26' fill='none' stroke='currentColor' stroke-width='1.8' stroke-linecap='round' stroke-linejoin='round' focusable='false'>" & d & "</svg>"
End Function
%>