<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<%

html = "<html>"
html = html & "<head>"

html = html & "<title>Cor Gestion</title>"
html = html & "<meta charset='utf-8' />"
html = html & "<meta name='viewport' content='width=device-width, initial-scale=1' />"
html = html & "<link rel='stylesheet' href='assets/css/main.css' />"

html = html & "</head>"
html = html & "<body>"

nFecini = 0

If nFecini < CLng(CDate("01-01-1900")) Then
	nFecini = CLng(CDate("01-01-1900"))
Else
	nFecini = 0
End If

sCli = request.form("cli")
if sCli & "" = "" then sCli = request.querystring("cli")

'sInc = request.form("inc")
'if sInc & "" = "" then sInc = request.querystring("inc")
sInc = 1

sProy = request.form("proy")
if sProy & "" = "" then sProy = request.querystring("proy")

sApl = request.form("apl")
if sApl & "" = "" then sApl = request.querystring("apl")

sFDesde = request.form("fecd")
if sFDesde & "" = "" then sFDesde = request.querystring("fecd")
if sFDesde & "" <> "" then sFDesde = CLng(CDate(sFDesde)) - nFecini

sFHasta = request.form("fech")
if sFHasta & "" = "" then sFHasta = request.querystring("fech")
if sFHasta & "" <> "" then sFHasta = CLng(CDate(sFHasta)) - nFecini

sOpc = request.querystring("opc")

If sOpc = "exp" Then

	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "content-disposition", " filename=reporte.xls"

End If

Set Cnn = Server.CreateObject("ADODB.Connection")
Cnn.open scon1

if sOpc <> "exp" then

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

html = html & "</div>"

'html = html & "<div style='height:4%;width:100%;border-top:solid 5px #FF3D3D;border-bottom:double 6px #001BFF;'>"
html = html & "<table class='titulo-seccion' >"
html = html & "<tr style='display:flex;align-items:center;justify-content:space-between;'>"

html = html & "<td><h2>REPORTE GENERADO</h2></td>"

'formx50 = "<form id='formx50' style='height:70%;' method='post' action='reportesDet.asp?" & request.form & "&opc=exp'></form>"
'botonx50 = "<input class='bot' type='button' value='GENERAR REPORTE' onclick=""document.getElementById('formx50').submit();"">"
'html = html & "<td>" & botonx50 & "</td>"

'formx10 = "<form id='formx10' style='height:70%;' method='post' action='reportes.asp'></form>"
'botonx10 = "<input class='bot' type='button' value='VOLVER' onclick=""document.getElementById('formx10').submit();"">"
'html = html & "<td>" & botonx10 & "</td>"

html = html & "</tr>"

html = html & "</table>"

html = html & "</div>"

end if

html = html & "<div>"

if sInc = 1 then

iSql = "sp_cg_Reportes_Detalle @Modo='incidentes'"

if sCli <> 0 then
	iSql = iSql & ", @Cli=" & sCli
else
	iSql = iSql & ", @Cli=0"
end if

if sFDesde & "" <> "" then
	iSql = iSql & ", @FDesde=" & sFDesde
end if

if sFHasta & "" <> "" then
	iSql = iSql & ", @FHasta=" & sFHasta
end if

Set rs = Cnn.Execute(iSql)

if not rs.eof then

	' --- INICIO NORMALIZACION: CABECERA DE REPORTE ---
	html = html & "<div class='inner' style='padding: 2rem;'>"
	
	' Contenedor de la cabecera (Filtros aplicados)
	html = html & "<div class='bloque-datos' style='margin-bottom: 0;border-bottom:none;border-radius:8px 8px 0px 0px;'>"

	'html = html & "<table xclass='tabl' style='width:900px;color:#000000;font-family:Serif;text-align:center;' cellspacing=0 align=center >"
	'html = html & "<tr>"
	'html = html & "<td colspan=7 style='font-size:17pt;' align=left><b><u>INCIDENTES Y CONSULTAS</u></b></td>"
	'html = html & "</tr>"
	html = html & "<div class='bloque-header'>INCIDENTES Y CONSULTAS</div>"

	html = html & "<div class='bloque-body grid-datos-doble'>"

	'html = html & "<tr>"
	'html = html & "<td colspan=7 >&nbsp;</td>"
	'html = html & "</tr>"

	' --- LÓGICA DE CLIENTE ---
	nCli = getCli(sCli)

	'html = html & "<tr>"
	'html = html & "<td colspan=7 style='font-size:12pt;' align=left><b>CLIENTE:</b>&nbsp;" & nCli & "</td>"
	'html = html & "</tr>"
	html = html & "<div class='dato-label'>CLIENTE</div>"
	html = html & "<div class='dato-value' style='font-weight: 600;'>" & nCli & "</div>"

	'html = html & "<tr>"
	'html = html & "<td colspan=7 >&nbsp;</td>"
	'html = html & "</tr>"

	' --- LÓGICA DE FECHAS ---
	if SFDesde & "" <> "" then
		sFecDesde = Fechador(SFDesde + nFecini)
	else
		sFecDesde = "TODO"
	end if

	if SFHasta & "" <> "" then
		sFecHasta = Fechador(SFHasta + nFecini)
	else
		sFecHasta = "TODO"
	end if

	'html = html & "<tr>"
	'html = html & "<td style='font-size:12pt;' align=left><b>FECHA DESDE:</b>&nbsp;" & sFecDesde & "</td>"
	'html = html & "</tr>"
	html = html & "<div class='dato-label'>FECHA DESDE</div>"
	html = html & "<div class='dato-value'>" & sFecDesde & "</div>"

	'html = html & "<tr>"
	'html = html & "<td style='font-size:12pt;' align=left><b>FECHA HASTA:</b>&nbsp;" & sFecHasta & "</td>"
	'html = html & "</tr>"
	html = html & "<div class='dato-label'>FECHA HASTA</div>"
	html = html & "<div class='dato-value'>" & sFecHasta & "</div>"

	'html = html & "<tr>"
	'html = html & "<td colspan=7 >&nbsp;</td>"
	'html = html & "</tr>"
	'html = html & "<tr>"
	'html = html & "</table>"
	
	html = html & "</div>" ' Fin .grid-datos-doble
	html = html & "</div>" ' Fin .bloque-datos (Cabecera)

	' --- INICIO NORMALIZACION: CUERPO DEL REPORTE (TABLA) ---
	' Abrimos la tarjeta contenedora para la tabla
	html = html & "<div class='bloque-datos table-responsive' style='border-radius:0px 0px 8px 8px;'>"
	
	'html = html & "<table class='tablRep' style='width:900px;color:#000000;font-family:Serif;' cellspacing=0 cellpadding=0 align=center >"
	html = html & "<table class='tabla-moderna'>"

	' --- NUEVO: EXTRAEMOS ENCABEZADOS FUERA DEL BUCLE ---
	html = html & "<thead>"
	html = html & "<tr>"
	html = html & "<th style='width:5%;'>ID</th>"
	html = html & "<th style='width:20%;'>CONTACTO</th>"
	html = html & "<th style='width:20%;'>RESPONSABLE</th>"
	html = html & "<th style='width:10%;'>FEC ING</th>"
	html = html & "<th style='width:10%;'>FEC RESOL</th>"
	html = html & "<th style='width:10%;'>TIPO</th>"
	html = html & "<th style='width:25%;'>CLIENTE</th>"
	html = html & "</tr>"
	html = html & "</thead>"
	html = html & "<tbody>"

	do while not rs.eof

		sBco = ""
		if rs(9) = 38 then
			sBco = "BCO ENTRE RIOS"
		elseif rs(9) = 39 then
			sBco = "BCO SAN JUAN"
		elseif rs(9) = 40 then
			sBco = "BCO SANTA CRUZ"
		elseif rs(9) = 41 then
			sBco = "BCO SANTA FE"
		elseif rs(9) = 4 then
			sBco = "GRUPO PETERSEN"
		else
			sBco = rs(6)
		end if

		' --- ENCABEZADOS LEGACY (COMENTADOS: AHORA ESTÁN EN EL THEAD) ---
		'html = html & "<tr>"
		'html = html & "<td style='witdh:1%;'><b><center>ID</center></b></td>"
		'html = html & "<td style='witdh:26%;'><b><center>CONTACTO</center></b></td>"
		'html = html & "<td style='witdh:20%;'><b><center>RESPONSABLE</center></b></td>"
		'html = html & "<td style='witdh:10%;'><b><center>FEC ING</center></b></td>"
		'html = html & "<td style='witdh:10%;'><b><center>FEC RESOL</center></b></td>"
		'html = html & "<td style='witdh:10%;'><b><center>TIPO</center></b></td>"
		'html = html & "<td style='witdh:23%;padding-right:5px;'><b><center>CLIENTE</center></b></td>"
		'html = html & "</tr>"

		' --- FILA PRINCIPAL DE DATOS ---
		'html = html & "<tr>"
		html = html & "<tr>"

		'html = html & "<td align=center>" & rs(0) & "</td>"
		html = html & "<td>" & rs(0) & "</td>"
		
		'html = html & "<td align=center>" & rs(1) & "</td>"
		html = html & "<td>" & rs(1) & "</td>"
		
		'html = html & "<td align=center>" & rs(2) & "</td>"
		html = html & "<td>" & rs(2) & "</td>"

		sFecIng = Fechador(rs(3) + nFecini)
		'html = html & "<td align=center>" & sFecIng & "</td>"
		html = html & "<td>" & sFecIng & "</td>"

		sFecResol = Fechador(rs(4) + nFecini)
		'html = html & "<td align=center>" & sFecResol & "</td>"
		html = html & "<td>" & sFecResol & "</td>"

		if Trim(rs(5) & "") = "0" then
			xtipo = "CONSULTA"
		else
			xtipo = "DESARROLLO"
		end if
		
		'html = html & "<td align=center>" & xtipo & "</td>"
		html = html & "<td>" & xtipo & "</td>"
		
		'html = html & "<td align=center style='padding-right:5px;'>" & sBco & "</td>"
		html = html & "<td>" & sBco & "</td>"

		'html = html & "</tr>"
		html = html & "</tr>"

		' --- SUB-FILA 1: DETALLE ---
		'html = html & "<tr>"
		'html = html & "<td style='padding-left:8px;padding-right:5px;' colspan=7><b>DETALLE:</b> " & rs(7) & "</td>"
		'html = html & "</tr>"
		'html = html & "<tr>"
		'html = html & "<td colspan=7>" & rs(7) & "</td>"
		'html = html & "</tr>"
		
		html = html & "<tr>"
		' Ajustamos paddings y le quitamos el borde inferior para que parezca conectada a la fila de arriba
		html = html & "<td colspan='7' style='padding-top: 0.5rem; padding-bottom: 0.25rem; border-bottom: none; color: #475569;'><strong style='color: #0f172a;'>DETALLE:</strong> " & rs(7) & "</td>"
		html = html & "</tr>"

		' --- SUB-FILA 2: AVANCE ---
		'html = html & "<tr>"
		'html = html & "<td style='padding-left:8px;padding-right:8px;border-bottom:dashed 1px #AAAAAA;' colspan=7><b>AVANCE:</b> "  & rs(8) & "</td>"
		'html = html & "</tr>"
		'html = html & "<tr>"
		'html = html & "<td style='border-bottom:solid 2px #000000;margin-bot:100px;' colspan=7>" & rs(8) & "</td>"
		'html = html & "</tr>"
		
		html = html & "<tr>"
		' El borde inferior lo hereda del CSS de .tabla-moderna, cerrando visualmente el registro
		html = html & "<td colspan='7' style='padding-top: 0.25rem; padding-bottom: 1rem; color: #475569;'><strong style='color: #0f172a;'>AVANCE:</strong> " & rs(8) & "</td>"
		html = html & "</tr>"

		rs.movenext
	loop

	'html = html & "</table><br>"
	html = html & "</tbody>"
	html = html & "</table>"
	
	html = html & "</div>" ' Fin .bloque-datos table-responsive
	html = html & "</div>" ' Fin .inner (abierto en el fragmento anterior)

end if

end if


if sProy = 1 then

iSql = "sp_cg_Reportes_Detalle @Modo='proyectos'"

if sCli <> 0 then
	iSql = iSql & ", @Cli=" & sCli
else
	iSql = iSql & ", @Cli=0"
end if

Set rs = Cnn.Execute(iSql)

if not rs.eof then

html = html & "<table style='width:900px;color:#000000;font-family:Serif;text-align:center;' cellspacing=0 cellpadding=0 align=center >"

html = html & "<tr>"

html = html & "<td colspan=7 style='font-size:17pt;' align=left><b><u>PROYECTOS</u></b></td>"

html = html & "</tr>"

html = html & "<tr>"

html = html & "<td colspan=7 >&nbsp;</td>"

html = html & "</tr>"

html = html & "<tr>"

html = html & "</table>"

html = html & "<table class='tabl' style='width:900px;color:#000000;font-family:Serif;' cellspacing=0 cellpadding=0 align=center >"

do while not rs.eof

sBco = ""
if rs(9) = 38 then
	sBco = "ENTRE RIOS"
elseif rs(9) = 39 then
	sBco = "SAN JUAN"
elseif rs(9) = 40 then
	sBco = "SANTA CRUZ"
elseif rs(9) = 41 then
	sBco = "SANTA FE"
elseif rs(9) = 4 then
	sBco = "PETERSEN"
else
	sBco = rs(6)
end if

html = html & "<tr>"

html = html & "<td style='witdh:1%;'><b>ID</b></td>"
html = html & "<td style='witdh:26%;'><b>CONTACTO</b></td>"
html = html & "<td style='witdh:23%;'><b>RESPONSABLE</b></td>"
html = html & "<td style='witdh:10%;'><b>FEC ING</b></td>"
html = html & "<td style='witdh:10%;'><b>FEC RESOL</b></td>"
html = html & "<td style='witdh:10%;'><b>TIPO</b></td>"
html = html & "<td style='witdh:20%;'><b>BANCO</b></td>"

html = html & "</tr>"

html = html & "<tr>"

html = html & "<td>" & rs(0) & "</td>"
html = html & "<td>" & rs(1) & "</td>"
html = html & "<td>" & rs(2) & "</td>"



html = html & "<td>" & rs(3) & "</td>"
html = html & "<td>" & rs(4) & "</td>"
if Trim(rs(5) & "") = "0" then
	xtipo = "CONSULTA"
else
	xtipo = "DESARROLLO"
end if
html = html & "<td>" & xtipo & "</td>"
html = html & "<td>" & rs(6) & "</td>"

html = html & "</tr>"

html = html & "<tr>"

html = html & "<td colspan=7><b>DETALLE:</b></td>"

html = html & "</tr>"

html = html & "<tr>"

html = html & "<td colspan=7>" & rs(7) & "</td>"

html = html & "</tr>"

html = html & "<tr>"

html = html & "<td colspan=7><b>AVANCE:</b></td>"

html = html & "</tr>"

html = html & "<tr>"

html = html & "<td style='border-bottom:solid 1px #aaaaaa;' colspan=7>" & rs(8) & "</td>"

html = html & "</tr>"

rs.movenext
loop

html = html & "</table><br>"

end if

end if

if sApl = 1 then

iSql = "sp_cg_Reportes_Detalle @Modo='apl'"

if sCli <> 0 then
	iSql = iSql & ", @Cli=" & sCli
else
	iSql = iSql & ", @Cli=0"
end if

Set rs = Cnn.Execute(iSql)

if not rs.eof then

html = html & "<table style='width:900px;color:#000000;font-family:Serif;text-align:center;' cellspacing=0 align=center border>"

html = html & "<tr>"

html = html & "<td colspan=4><b>APLs</b></td>"

html = html & "</tr>"

html = html & "<tr>"

html = html & "<td style='witdh:5%;'><b>ID</b></td>"
html = html & "<td style='witdh:31%;'><b>NOMBRE</b></td>"
html = html & "<td style='witdh:33%;'><b>RESPONSABLE</b></td>"
html = html & "<td style='witdh:31%;'><b>FECHA ENVIO</b></td>"

html = html & "</tr>"

do while not rs.eof

html = html & "<tr>"

html = html & "<td>" & rs(0) & "</td>"
html = html & "<td>" & rs(1) & "</td>"
html = html & "<td>" & rs(2) & "</td>"
html = html & "<td>" & rs(3) & "</td>"

html = html & "</tr>"

rs.movenext
loop

html = html & "</table>"

end if

end if

html = html & "<script type='text/javascript'>"
html = html & " function borrar(formu){ "
html = html & " var r = confirm('Desea eliminar proyecto?'); "
html = html & " if (r){ "
html = html & " document.getElementById(formu).submit(); "
html = html & " }else{ "
html = html & " return false;"
html = html & " }} "
html = html & "</script>"


html = html & "<script src='assets/js/jquery.min.js'></script>"
html = html & "<script src='assets/js/skel.min.js'></script>"
html = html & "<script src='assets/js/util.js'></script>"
html = html & "<script src='assets/js/main.js'></script>"

html = html & formx22 & formx23 & xhtml & formx10 & xformx & formx1 & formx2 & formx3 & formx29 & formx19 & formx50 & "</body>"
html = html & "</html>"

response.write html
%>