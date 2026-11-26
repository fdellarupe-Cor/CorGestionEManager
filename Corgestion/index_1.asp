<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<%
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

'html = "<html>"
'html = html & "<head>"

'html = html & "<title>Cor Gestion</title>"
'html = html & "<meta http-equiv='Content-Type' content='text/html;charset=ISO-8859-1'>"
'html = html & "<meta name='viewport' content='width=device-width, initial-scale=1' />"
'html = html & "<link rel='stylesheet' href='assets/css/main.css' />"
'html = html & "<script src='assets/js/jquery.min.js'></script>"
'html = html & "<script src='assets/js/skel.min.js'></script>"
'html = html & "<script src='assets/js/util.js'></script>"
'html = html & "<script src='assets/js/main.js'></script>"
'html = html & "<script type='text/javascript'>"
'html = html & "$(document).ready(function(){"
''html = html & "$('#body').fadeIn(1500);"
'html = html & "});"
'html = html & "</script>"
'html = html & "</head>"

html=headhtml()
html = html & "<body id='body' name='body' style='height:100%;width:100%;' onload='mueveReloj();'>"

'SCRIPTS SLIDER
'html = html & "<script type='text/javascript' src='assets/js/jquery-1.9.1.min.js'></script>"
'html = html & "<script type='text/javascript' src='assets/js/jssor.core.js'></script>"
'html = html & "<script type='text/javascript' src='assets/js/jssor.utils.js'></script>"
'html = html & "<script type='text/javascript' src='assets/js/jssor.slider.mini.js'></script>"

'html = html & "<script>"

'html = html & "var _SlideshowTransitions = [ "
'html = html & " {$Duration: 700, $Opacity: 2, $Brother: { $Duration: 1000, $Opacity: 2} } "
'html = html & "]; "
'html = html & "</script>"

'html = html & "<script>"
'html = html & "jQuery(document).ready(function ($) { "
'html = html & "var options = { "
'html = html & "$FillMode: 3, "                                     '[Optional] The way to fill image in slide, 0 stretch, 1 contain (keep aspect ratio and put all inside slide), 2 cover (keep aspect ratio and cover whole slide), 4 actuall size, default value is 0
'html = html & "$DragOrientation: 1, "                              '[Optional] Orientation to drag slide, 0 no drag, 1 horizental, 2 vertical, 3 either, default value is 1 (Note that the $DragOrientation should be the same as $PlayOrientation when $DisplayPieces is greater than 1, or parking position is not 0)
'html = html & "$AutoPlay: true, "                                  '[Optional] Whether to auto play, to enable slideshow, this option must be set to true, default value is false
'html = html & "$AutoPlayInterval: 700, "                          '[Optional] Interval (in milliseconds) to go for next slide since the previous stopped if the slider is auto playing, default value is 3000
'html = html & "$SlideshowOptions: { "                              '[Optional] Options to specify and enable slideshow or not
'html = html & "$Class: $JssorSlideshowRunner$, "                   '[Required] Class to create instance of slideshow
'html = html & "$Transitions: _SlideshowTransitions, "              '[Required] An array of slideshow transitions to play slideshow
'html = html & "$TransitionsOrder: 1, "                             '[Optional] The way to choose transition to play slide, 1 Sequence, 0 Random
'html = html & "$ShowLink: true "                                   '[Optional] Whether to bring slide link on top of the slider when slideshow is running, default value is false
'html = html & "}"
'html = html & "};"

'html = html & "var jssor_slider1 = new $JssorSlider$('slider1_container', options);"
'html = html & "});"
'html = html & "</script>"

html = html & "<div id='wrapper' >"

'''''''''''''''''TEST JQUERY''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'html = html & "<form id='test' name='test' style='display:none;' method='post' action='test.asp'></form>"
'html = html & "<input type='button' value='JQuery' onclick='testeo();'>"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

html = html & "<header id='header' style='background:#0e1949;padding:1rem 0;'>"

html = html & "<div >"

html = html & "<table style='width:100%;'><tr style='display:flex;padding: 0 1rem;'>"

html = html & "<td style='width:75%;' valign=center><a href='index.asp' class='logo' >"
html = html & "<img src='images/iso_cor.png' style='width:6rem;' >"
html = html & "<font style='font-size:15pt;color:white;'><b>Cor Gestion</b> | Sistema Administrativo</font></a></td>"
html = html & "<td style='width:25%;display:flex;align-items:center;flex-direction:row-reverse;' align='right'>"
'html = html & "<form name='form_reloj'>"
'html = html & "<input class='reloj' type='text' name='reloj' size='10' style='width:8rem;color:#000000;font-family:sans-serif;font-size:18pt;background:#FFFFFF;'>"
'html = html & "</form>"
'html = html & "<nav>"
'html = html & "<ul>"
'html = html & "<li><a href='#menu'></a></li>"
'html = html & "</ul>"
'html = html & "</nav>"

' Reemplazamos la lista vacía por un botón hamburguesa con icono SVG moderno
html = html & "<nav>"
html = html & "<a href='#menu' class='btn-hamburguesa' aria-label='Abrir menú'>"
html = html & "<svg xmlns='http://www.w3.org/2000/svg' fill='none' viewBox='0 0 24 24' stroke='currentColor' stroke-width='2'>"
html = html & "<path stroke-linecap='round' stroke-linejoin='round' d='M4 6h16M4 12h16M4 18h16' />"
html = html & "</svg>"
html = html & "</a>"
html = html & "</nav>"
html = html & "</td>"

html = html & "</tr></table>"

html = html & "</div>"

'html = html & "<nav>"
'html = html & "<ul>"
'html = html & "<li><a href='#menu'></a></li>"
'html = html & "</ul>"
'html = html & "</nav>"

html = html & "</header>"

html = html & "</div>"

'ARMA MENU DE PARAMETROS
html = html & getMenuParam()

html = html & "<div id='main' >"

html = html & "<div class='titulo-seccion'>"
'html = html & "<table style='height:100%;width:100%;'><tr>"
'html = html & "<td style='height:100%;width:60%;'><font style='margin-left:20px;font-family:Serif;color:#000000;font-size:25pt;letter-spacing:4px;' >"
html = html & "<h1>"
'html = html & "Sistema Administrativo</font>"
html = html & "SECCIONES</h1>"
'html = html & "SISTEMA ADMINISTRATIVO</font></td>"
'html = html & "</tr></table>"
html = html & "</div>"

'ARMA MENU PRINCIPAL
html = html & getMenuPrinc()

'html = html & "<div id='slider1_container' style='position:relative;width:1300px;height:200px;background-color:#FFFFFF;margin-left:10px;'>"

<!-- Loading Screen -->
'html = html & "<div u='loading' style='position: absolute; top: 0px; left: 0px;'>"
'html = html & "<div style='filter:alpha(opacity=70);opacity:0.7;position:absolute;display:block;background-color:#FFFFFF;top:0px;left:0px;width:1300px;height:200px;'>"
'html = html & "</div>"
'html = html & "<div style='position:absolute;display:block;background:url(../img/loading.gif) no-repeat center center;top:0px;left:0px;width:1300px;height:200px;'>"
'html = html & "</div>"
'html = html & "</div>"

<!-- Slides Container -->
'html = html & "<div u='slides' style='cursor:move;position:absolute;left:0px;top:0px;width:1300px;height:200px;overflow: hidden;'>"

'NUEVO SLIDER
html = html & "<div class='contenedor-marquee'>"
html = html & "<div class='track-marquee'>"

Set rc = Exec(Cnn, "sp_cg_Utilidades @Modo='listar_logos', @Id=37")

'response.write sql

' Armamos el HTML de los logos en una variable temporal
logosHTML = ""
do while not rc.eof
'html = html & "<div>"
'html = html & "<a u=image href='incidentes.asp?xmenu=9&xtmenu=INICIO&ord=&cli=" & rc(0) & "'>"
'html = html & "<img src='../Corgestion/images/" & rc(1) & "'  width='100px' height='100px' />"
'html = html & "</a>"
'html = html & "</div>"
	logosHTML = logosHTML & "<a class='item-marquee' href='incidentes.asp?xmenu=9&xtmenu=INICIO&ord=&cli=" & rc(0) & "'>"
	logosHTML = logosHTML & "<img src='../Corgestion/images/" & rc(1) & "' alt='Cliente' />"
	logosHTML = logosHTML & "</a>"
	rc.movenext
loop

' Para que la animación CSS sea un ciclo perfecto e infinito sin cortes, 
' concatenamos la lista de logos DOS VECES en el contenedor.
html = html & logosHTML & logosHTML

html = html & "</div>" ' Cierre de track-marquee
html = html & "</div>" ' Cierre de contenedor-marquee

' Footer moderno activado (con clase limpia y año automático)
anioActual = Year(Now())
html = html & "<div class='footer-legal'>"
html = html & "&copy; " & anioActual & " Cor Consulting S.R.L. Todos los derechos reservados."
html = html & "</div>"

'html = html & "<div style='width:100%;heith:30px;'>"
'html = html & "&nbsp;"
'html = html & "</div>"

'html = html & "<div style='width:100%;heith:30px;' align=center>"
'html = html & "<font style='font-size:20pt;color:#000000;'>2018 © Cor Consulting S.R.L.</font>"
'html = html & "</div>"


html = html & "</div>" ' Cierre de main


html = html & "<script language='JavaScript'> "
html = html & "function mueveReloj(){ "
html = html & " momentoActual = new Date(); "
html = html & " hora = momentoActual.getHours();"
html = html & " minuto = momentoActual.getMinutes(); "
html = html & " segundo = momentoActual.getSeconds(); "
html = html & "	str_segundo = new String (segundo); "
html = html & " if (str_segundo.length == 1){ "
html = html & " segundo = '0' + segundo}; "
html = html & " str_minuto = new String (minuto); "
html = html & " if (str_minuto.length == 1){ "
html = html & " minuto = '0' + minuto } ;"
html = html & " str_hora = new String (hora); "
html = html & " if (str_hora.length == 1) {"
html = html & " hora = '0' + hora} ; "
html = html & " horaImprimible = hora + ' : ' + minuto + ' : ' + segundo; "
html = html & " document.form_reloj.reloj.value = horaImprimible; "
html = html & " setTimeout('mueveReloj()',1000); "
html = html & "} "

html = html & "function testeo(){ "
html = html & " var conf = confirm('NO ACEPTES ! ! !'); "
html = html & " if (conf) { "
html = html & " test.submit(); "
html = html & "}"
html = html & "} "
html = html & "</script> "

html = html & "</body>"
html = html & "</html>"

response.write html
%>