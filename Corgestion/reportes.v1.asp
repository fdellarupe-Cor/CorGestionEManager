<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<%

html = "<html>"
html = html & "<head>"

html = html & "<title>Cor Gestion</title>"
html = html & "<meta charset='utf-8' />"
html = html & "<meta name='viewport' content='width=device-width, initial-scale=1' />"
html = html & "<link rel='stylesheet' href='assets/css/main.css' />"
html = html & "<link rel='stylesheet' href='assets/js/jquery-ui.css'>"

html = html & "</head>"
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

html = html & "<div>"

html = html & "<header id='header' style='background:#0e1949;padding:1rem 0;'>"

html = html & "<div >"

html = html & "<table style='width:100%;'><tr style='display:flex;padding: 0 1rem;'>"

html = html & "<td style='width:75%;' valign=center><a href='index.asp' class='logo' >"
html = html & "<img src='images/iso_cor.png' style='width:6rem;' >"
html = html & "<font style='font-size:15pt;color:white;'><b>Cor Gestion</b> | Sistema Administrativo</font></a></td>"
html = html & "<td style='width:25%;display:flex;align-items:center;flex-direction:row-reverse;' align='right'>"
html = html & "<nav>"
html = html & "<ul>"
'html = html & "<li><a href='#menu'></a></li>"
html = html & "</ul>"
html = html & "</nav>"
html = html & "</td>"

html = html & "</tr></table>"

html = html & "</div>"

html = html & "</header>"

'html = html & "<div style='height:9%;width:100%;border-top:solid 5px #FF3D3D;border-bottom:double 6px #001BFF;'>"

html = html & "<table class='titulo-seccion' >"
html = html & "<tr style='display:flex;align-items:center;justify-content:space-between;'>"

html = html & "<td><h2>REPORTES</h2></td>"

formx04 = "<form id='formx04' method='post' action='index.asp'></form>"
botonx04 = "<input class='bot' type='button' value='VOLVER' onclick=""document.getElementById('formx04').submit();"">"
html = html & "<td>" & botonx04 & "</td>"

html = html & "</tr>"

html = html & "</table>"

html = html & "</div><br>"

	' --- INICIO NORMALIZACION: FORMULARIO GENERAR REPORTE ---
	' Contenedor centrado y compacto
	'html = html & "<div style='height:75%;width:100%;'><br>"
	html = html & "<div style='max-width: 600px; margin: 2rem auto;'>"

	html = html & "<form id='xForm' method='post' action='reportesDet.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"

	'html = html & "<table class='tabl' style='height:50%;width:40%;' align=center cellspacing=0>"
	html = html & "<div class='bloque-datos'>"

	' --- ENCABEZADO ---
	'html = html & "<tr>"
	'html = html & "<td class='fil' style='font-size:12pt;' colspan=2 align=center valign=center><b>GENERAR REPORTE DE INCIDENTES</b></td>"
	'html = html & "</tr>"
	html = html & "<div class='bloque-header'>GENERAR REPORTE DE INCIDENTES</div>"
	
	html = html & "<div class='bloque-body'>"

	'ESPACIO
	'html = html & "<tr>"
	'html = html & "<td class='fil' colspan=2 >&nbsp;</td>"
	'html = html & "</tr>"
	
	' --- INICIO GRILLA ---
	html = html & "<div class='grid-datos-doble'>"

	'CLIENTE
	'html = html & "<tr>"
	'html = html & "<td class='fil' align=center valign=center>CLIENTE</td>"
	'html = html & "<td class='fil' valign=center>"
	'html = html & armaCombo("select id,descripcion from clientes",cnn,request.form("cli"),0,1,"cli")
	'html = html & "</td>"
	'html = html & "</tr>"
	html = html & "<div class='dato-label'>CLIENTE</div>"
	html = html & "<div class='dato-value'>"
	html = html & armaCombo("sp_cg_Catalogo_Listar @Catalogo='clientes'",cnn,request.form("cli"),0,1,"cli")
	html = html & "</select></div>" ' <--- PARCHE: Forzamos el cierre del combo

	'INCIDENTES (Mantenemos la secci�n comentada original)
	'html = html & "<tr>"
	'html = html & "<td class='fil' align=center valign=center>INCIDENTES</td>"
	'html = html & "<td class='fil' valign=center>"
	'html = html & "<select class='selec' id='inc' name='inc' style='width:20%;'>"
	'html = html & "<option value='0'>NO</option>"
	'html = html & "<option value='1'>SI</option>"
	'html = html & "</select>"
	'html = html & "</td>"
	'html = html & "</tr>"

	'PROYECTOS (Mantenemos la secci�n comentada original)
	'html = html & "<tr>"
	'html = html & "<td class='fil' align=center valign=center>PROYECTOS</td>"
	'html = html & "<td class='fil' valign=center>"
	'html = html & "<select class='selec' id='proy' name='proy' style='width:20%;'>"
	'html = html & "<option value='0'>NO</option>"
	'html = html & "<option value='1'>SI</option>"
	'html = html & "</select>"
	'html = html & "</td>"
	'html = html & "</tr>"

	'APLs (Mantenemos la secci�n comentada original)
	'html = html & "<tr>"
	'html = html & "<td class='fil' align=center valign=center>APLs</td>"
	'html = html & "<td class='fil' valign=center>"
	'html = html & "<select class='selec' id='apl' name='apl' style='width:20%;'>"
	'html = html & "<option value='0'>NO</option>"
	'html = html & "<option value='1'>SI</option>"
	'html = html & "</select>"
	'html = html & "</td>"
	'html = html & "</tr>"

	'FECHA DESDE
	'html = html & "<tr>"
	'html = html & "<td class='fil' align=center valign=center>FECHA DESDE</td>"
	'html = html & "<td class='fil' valign=center>"
	'html = html & "<input class='inpu' id='fecd' name='fecd' style='width:40%;' type='text' />"
	'html = html & "</td>"
	'html = html & "</tr>"
	html = html & "<div class='dato-label'>FECHA DESDE</div>"
	html = html & "<div><input type='text' id='fecd' name='fecd' class='input-moderno'></div>"

	'FECHA HASTA
	'html = html & "<tr>"
	'html = html & "<td class='fil' align=center valign=center>FECHA HASTA</td>"
	'html = html & "<td class='fil' valign=center>"
	'html = html & "<input class='inpu' id='fech' name='fech' style='width:40%;' type='text' />"
	'html = html & "</td>"
	'html = html & "</tr>"
	html = html & "<div class='dato-label'>FECHA HASTA</div>"
	html = html & "<div><input type='text' id='fech' name='fech' class='input-moderno'></div>"

	'ESPACIO
	'html = html & "<tr>"
	'html = html & "<td class='fil' colspan=2 >&nbsp;</td>"
	'html = html & "</tr>"

	' --- FIN GRILLA ---
	html = html & "</div>" 

	' --- BOTON CONSULTAR ---
	'html = html & "<tr>"
	'html = html & "<td style='height:10%;width:50%;' colspan=2 align=center>"
	html = html & "<hr style='border: 0; border-top: 1px solid #e2e8f0; margin: 1.5rem 0;'>"
	html = html & "<div style='text-align: center; margin-bottom: 0.5rem;'>"

	if getPermisos(xmenu,xtmenu,"CONS") = 1 then
		'html = html & "<input class='bot' type='button' value='Consultar' onclick=""document.getElementById('xForm').submit();"">"
		html = html & "<input type='button' class='btn-principal' value='Consultar' onclick=""document.getElementById('xForm').submit();"">"
	else
		'html = html & "<input class='botsp' type='button' value='Consultar' onclick=""alert('No tiene permiso para consultar.');"">"
		' Usamos la clase de deshabilitado para que mantenga la coherencia visual
		html = html & "<input type='button' class='btn-principal btn-accion-deshabilitado' value='Consultar' onclick=""alert('No tiene permiso para consultar.');"">"
	end if

	'html = html & "</td>"
	'html = html & "</tr>"
	html = html & "</div>"

	'html = html & "</table>"
	html = html & "</div>" ' Fin .bloque-body
	html = html & "</div>" ' Fin .bloque-datos

	html = html & "</form>"

	'html = html & "</div>"
	html = html & "</div>" ' Fin max-width















html = html & "<script type='text/javascript'>"
html = html & " function modificar(formu){ "
html = html & " document.getElementById(formu).submit();"
html = html & " alert('Descripcion actualizada');} "
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
html = html & "dayNames: ['Domingo', 'Lunes', 'Martes', 'Miercoles', 'Jueves', 'Viernes', 'Sabado'],"
html = html & "dayNamesShort: ['Dom','Lun','Mar','Mi�','Juv','Vie','S�b'],"
html = html & "dayNamesMin: ['Do','Lu','Ma','Mi','Ju','Vi','Sa'],"
html = html & "weekHeader: 'Sm',"
html = html & "dateFormat: 'dd/mm/yy',"
html = html & "firstDay: 1,"
html = html & "isRTL: false,"
html = html & "showMonthAfterYear: false,"
html = html & "yearSuffix: ''"
html = html & "};"
html = html & "$.datepicker.setDefaults($.datepicker.regional['es']);"

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

%>