<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/ui_componentes.inc" -->
<%
'[UI-ANT] ' ===== Cuerpo anterior completo: reemplazado por el bloque [UI] que sigue. Se borra entero buscando [UI-ANT] =====

'[UI-ANT] html = "<html>"
'[UI-ANT] html = html & "<head>"

'[UI-ANT] html = html & "<title>Cor Gestion</title>"
'[UI-ANT] html = html & "<meta charset='utf-8' />"
'[UI-ANT] html = html & "<meta name='viewport' content='width=device-width, initial-scale=1' />"
'[UI-ANT] html = html & "<link rel='stylesheet' href='assets/css/main.css' />"

'[UI-ANT] html = html & "</head>"
'[UI-ANT] html = html & "<body>"

'[UI-ANT] nFecini = 0

'[UI-ANT] If nFecini < CLng(CDate("01-01-1900")) Then
	'[UI-ANT] nFecini = CLng(CDate("01-01-1900"))
'[UI-ANT] Else
	'[UI-ANT] nFecini = 0
'[UI-ANT] End If

'[UI-ANT] sCli = request.form("cli")
'[UI-ANT] if sCli & "" = "" then sCli = request.querystring("cli")

'[UI-ANT] 'sInc = request.form("inc")
'[UI-ANT] 'if sInc & "" = "" then sInc = request.querystring("inc")
'[UI-ANT] sInc = 1

'[UI-ANT] sProy = request.form("proy")
'[UI-ANT] if sProy & "" = "" then sProy = request.querystring("proy")

'[UI-ANT] sApl = request.form("apl")
'[UI-ANT] if sApl & "" = "" then sApl = request.querystring("apl")

'[UI-ANT] sFDesde = request.form("fecd")
'[UI-ANT] if sFDesde & "" = "" then sFDesde = request.querystring("fecd")
'[UI-ANT] if sFDesde & "" <> "" then sFDesde = CLng(CDate(sFDesde)) - nFecini

'[UI-ANT] sFHasta = request.form("fech")
'[UI-ANT] if sFHasta & "" = "" then sFHasta = request.querystring("fech")
'[UI-ANT] if sFHasta & "" <> "" then sFHasta = CLng(CDate(sFHasta)) - nFecini

'[UI-ANT] sOpc = request.querystring("opc")

'[UI-ANT] If sOpc = "exp" Then

	'[UI-ANT] Response.ContentType = "application/vnd.ms-excel"
	'[UI-ANT] Response.AddHeader "content-disposition", " filename=reporte.xls"

'[UI-ANT] End If

'[UI-ANT] Set Cnn = Server.CreateObject("ADODB.Connection")
'[UI-ANT] Cnn.open scon1

'[UI-ANT] if sOpc <> "exp" then

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

'[UI-ANT] html = html & "</div>"

'[UI-ANT] 'html = html & "<div style='height:4%;width:100%;border-top:solid 5px #FF3D3D;border-bottom:double 6px #001BFF;'>"
'[UI-ANT] html = html & "<table class='titulo-seccion' >"
'[UI-ANT] html = html & "<tr style='display:flex;align-items:center;justify-content:space-between;'>"

'[UI-ANT] html = html & "<td><h2>REPORTE GENERADO</h2></td>"

'[UI-ANT] 'formx50 = "<form id='formx50' style='height:70%;' method='post' action='reportesDet.asp?" & request.form & "&opc=exp'></form>"
'[UI-ANT] 'botonx50 = "<input class='bot' type='button' value='GENERAR REPORTE' onclick=""document.getElementById('formx50').submit();"">"
'[UI-ANT] 'html = html & "<td>" & botonx50 & "</td>"

'[UI-ANT] 'formx10 = "<form id='formx10' style='height:70%;' method='post' action='reportes.asp'></form>"
'[UI-ANT] 'botonx10 = "<input class='bot' type='button' value='VOLVER' onclick=""document.getElementById('formx10').submit();"">"
'[UI-ANT] 'html = html & "<td>" & botonx10 & "</td>"

'[UI-ANT] html = html & "</tr>"

'[UI-ANT] html = html & "</table>"

'[UI-ANT] html = html & "</div>"

'[UI-ANT] end if

'[UI-ANT] html = html & "<div>"

'[UI-ANT] if sInc = 1 then

'[UI-ANT] iSql = "sp_cg_Reportes_Detalle @Modo='incidentes'"

'[UI-ANT] if sCli <> 0 then
	'[UI-ANT] iSql = iSql & ", @Cli=" & sCli
'[UI-ANT] else
	'[UI-ANT] iSql = iSql & ", @Cli=0"
'[UI-ANT] end if

'[UI-ANT] if sFDesde & "" <> "" then
	'[UI-ANT] iSql = iSql & ", @FDesde=" & sFDesde
'[UI-ANT] end if

'[UI-ANT] if sFHasta & "" <> "" then
	'[UI-ANT] iSql = iSql & ", @FHasta=" & sFHasta
'[UI-ANT] end if

'[UI-ANT] Set rs = Cnn.Execute(iSql)

'[UI-ANT] if not rs.eof then

	'[UI-ANT] ' --- INICIO NORMALIZACION: CABECERA DE REPORTE ---
	'[UI-ANT] html = html & "<div class='inner' style='padding: 2rem;'>"
	
	'[UI-ANT] ' Contenedor de la cabecera (Filtros aplicados)
	'[UI-ANT] html = html & "<div class='bloque-datos' style='margin-bottom: 0;border-bottom:none;border-radius:8px 8px 0px 0px;'>"

	'[UI-ANT] 'html = html & "<table xclass='tabl' style='width:900px;color:#000000;font-family:Serif;text-align:center;' cellspacing=0 align=center >"
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td colspan=7 style='font-size:17pt;' align=left><b><u>INCIDENTES Y CONSULTAS</u></b></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='bloque-header'>INCIDENTES Y CONSULTAS</div>"

	'[UI-ANT] html = html & "<div class='bloque-body grid-datos-doble'>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td colspan=7 >&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"

	'[UI-ANT] ' --- LOGICA DE CLIENTE ---
	'[UI-ANT] nCli = getCli(sCli)

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td colspan=7 style='font-size:12pt;' align=left><b>CLIENTE:</b>&nbsp;" & nCli & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>CLIENTE</div>"
	'[UI-ANT] html = html & "<div class='dato-value' style='font-weight: 600;'>" & nCli & "</div>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td colspan=7 >&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"

	'[UI-ANT] ' --- LOGICA DE FECHAS ---
	'[UI-ANT] if SFDesde & "" <> "" then
		'[UI-ANT] sFecDesde = Fechador(SFDesde + nFecini)
	'[UI-ANT] else
		'[UI-ANT] sFecDesde = "TODO"
	'[UI-ANT] end if

	'[UI-ANT] if SFHasta & "" <> "" then
		'[UI-ANT] sFecHasta = Fechador(SFHasta + nFecini)
	'[UI-ANT] else
		'[UI-ANT] sFecHasta = "TODO"
	'[UI-ANT] end if

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td style='font-size:12pt;' align=left><b>FECHA DESDE:</b>&nbsp;" & sFecDesde & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>FECHA DESDE</div>"
	'[UI-ANT] html = html & "<div class='dato-value'>" & sFecDesde & "</div>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td style='font-size:12pt;' align=left><b>FECHA HASTA:</b>&nbsp;" & sFecHasta & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>FECHA HASTA</div>"
	'[UI-ANT] html = html & "<div class='dato-value'>" & sFecHasta & "</div>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td colspan=7 >&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "</table>"
	
	'[UI-ANT] html = html & "</div>" ' Fin .grid-datos-doble
	'[UI-ANT] html = html & "</div>" ' Fin .bloque-datos (Cabecera)

	'[UI-ANT] ' --- INICIO NORMALIZACION: CUERPO DEL REPORTE (TABLA) ---
	'[UI-ANT] ' Abrimos la tarjeta contenedora para la tabla
	'[UI-ANT] html = html & "<div class='bloque-datos table-responsive' style='border-radius:0px 0px 8px 8px;'>"
	
	'[UI-ANT] 'html = html & "<table class='tablRep' style='width:900px;color:#000000;font-family:Serif;' cellspacing=0 cellpadding=0 align=center >"
	'[UI-ANT] html = html & "<table class='tabla-moderna'>"

	'[UI-ANT] ' --- NUEVO: EXTRAEMOS ENCABEZADOS FUERA DEL BUCLE ---
	'[UI-ANT] html = html & "<thead>"
	'[UI-ANT] html = html & "<tr>"
	'[UI-ANT] html = html & "<th style='width:5%;'>ID</th>"
	'[UI-ANT] html = html & "<th style='width:20%;'>CONTACTO</th>"
	'[UI-ANT] html = html & "<th style='width:20%;'>RESPONSABLE</th>"
	'[UI-ANT] html = html & "<th style='width:10%;'>FEC ING</th>"
	'[UI-ANT] html = html & "<th style='width:10%;'>FEC RESOL</th>"
	'[UI-ANT] html = html & "<th style='width:10%;'>TIPO</th>"
	'[UI-ANT] html = html & "<th style='width:25%;'>CLIENTE</th>"
	'[UI-ANT] html = html & "</tr>"
	'[UI-ANT] html = html & "</thead>"
	'[UI-ANT] html = html & "<tbody>"

	'[UI-ANT] do while not rs.eof

		'[UI-ANT] sBco = ""
		'[UI-ANT] if rs(9) = 38 then
			'[UI-ANT] sBco = "BCO ENTRE RIOS"
		'[UI-ANT] elseif rs(9) = 39 then
			'[UI-ANT] sBco = "BCO SAN JUAN"
		'[UI-ANT] elseif rs(9) = 40 then
			'[UI-ANT] sBco = "BCO SANTA CRUZ"
		'[UI-ANT] elseif rs(9) = 41 then
			'[UI-ANT] sBco = "BCO SANTA FE"
		'[UI-ANT] elseif rs(9) = 4 then
			'[UI-ANT] sBco = "GRUPO PETERSEN"
		'[UI-ANT] else
			'[UI-ANT] sBco = rs(6)
		'[UI-ANT] end if

		'[UI-ANT] ' --- ENCABEZADOS LEGACY (COMENTADOS: AHORA ESTAN EN EL THEAD) ---
		'[UI-ANT] 'html = html & "<tr>"
		'[UI-ANT] 'html = html & "<td style='witdh:1%;'><b><center>ID</center></b></td>"
		'[UI-ANT] 'html = html & "<td style='witdh:26%;'><b><center>CONTACTO</center></b></td>"
		'[UI-ANT] 'html = html & "<td style='witdh:20%;'><b><center>RESPONSABLE</center></b></td>"
		'[UI-ANT] 'html = html & "<td style='witdh:10%;'><b><center>FEC ING</center></b></td>"
		'[UI-ANT] 'html = html & "<td style='witdh:10%;'><b><center>FEC RESOL</center></b></td>"
		'[UI-ANT] 'html = html & "<td style='witdh:10%;'><b><center>TIPO</center></b></td>"
		'[UI-ANT] 'html = html & "<td style='witdh:23%;padding-right:5px;'><b><center>CLIENTE</center></b></td>"
		'[UI-ANT] 'html = html & "</tr>"

		'[UI-ANT] ' --- FILA PRINCIPAL DE DATOS ---
		'[UI-ANT] 'html = html & "<tr>"
		'[UI-ANT] html = html & "<tr>"

		'[UI-ANT] 'html = html & "<td align=center>" & rs(0) & "</td>"
		'[UI-ANT] html = html & "<td>" & rs(0) & "</td>"
		
		'[UI-ANT] 'html = html & "<td align=center>" & rs(1) & "</td>"
		'[UI-ANT] html = html & "<td>" & rs(1) & "</td>"
		
		'[UI-ANT] 'html = html & "<td align=center>" & rs(2) & "</td>"
		'[UI-ANT] html = html & "<td>" & rs(2) & "</td>"

		'[UI-ANT] sFecIng = Fechador(rs(3) + nFecini)
		'[UI-ANT] 'html = html & "<td align=center>" & sFecIng & "</td>"
		'[UI-ANT] html = html & "<td>" & sFecIng & "</td>"

		'[UI-ANT] sFecResol = Fechador(rs(4) + nFecini)
		'[UI-ANT] 'html = html & "<td align=center>" & sFecResol & "</td>"
		'[UI-ANT] html = html & "<td>" & sFecResol & "</td>"

		'[UI-ANT] if Trim(rs(5) & "") = "0" then
			'[UI-ANT] xtipo = "CONSULTA"
		'[UI-ANT] else
			'[UI-ANT] xtipo = "DESARROLLO"
		'[UI-ANT] end if
		
		'[UI-ANT] 'html = html & "<td align=center>" & xtipo & "</td>"
		'[UI-ANT] html = html & "<td>" & xtipo & "</td>"
		
		'[UI-ANT] 'html = html & "<td align=center style='padding-right:5px;'>" & sBco & "</td>"
		'[UI-ANT] html = html & "<td>" & sBco & "</td>"

		'[UI-ANT] 'html = html & "</tr>"
		'[UI-ANT] html = html & "</tr>"

		'[UI-ANT] ' --- SUB-FILA 1: DETALLE ---
		'[UI-ANT] 'html = html & "<tr>"
		'[UI-ANT] 'html = html & "<td style='padding-left:8px;padding-right:5px;' colspan=7><b>DETALLE:</b> " & rs(7) & "</td>"
		'[UI-ANT] 'html = html & "</tr>"
		'[UI-ANT] 'html = html & "<tr>"
		'[UI-ANT] 'html = html & "<td colspan=7>" & rs(7) & "</td>"
		'[UI-ANT] 'html = html & "</tr>"
		
		'[UI-ANT] html = html & "<tr>"
		'[UI-ANT] ' Ajustamos paddings y le quitamos el borde inferior para que parezca conectada a la fila de arriba
		'[UI-ANT] html = html & "<td colspan='7' style='padding-top: 0.5rem; padding-bottom: 0.25rem; border-bottom: none; color: #475569;'><strong style='color: #0f172a;'>DETALLE:</strong> " & rs(7) & "</td>"
		'[UI-ANT] html = html & "</tr>"

		'[UI-ANT] ' --- SUB-FILA 2: AVANCE ---
		'[UI-ANT] 'html = html & "<tr>"
		'[UI-ANT] 'html = html & "<td style='padding-left:8px;padding-right:8px;border-bottom:dashed 1px #AAAAAA;' colspan=7><b>AVANCE:</b> "  & rs(8) & "</td>"
		'[UI-ANT] 'html = html & "</tr>"
		'[UI-ANT] 'html = html & "<tr>"
		'[UI-ANT] 'html = html & "<td style='border-bottom:solid 2px #000000;margin-bot:100px;' colspan=7>" & rs(8) & "</td>"
		'[UI-ANT] 'html = html & "</tr>"
		
		'[UI-ANT] html = html & "<tr>"
		'[UI-ANT] ' El borde inferior lo hereda del CSS de .tabla-moderna, cerrando visualmente el registro
		'[UI-ANT] html = html & "<td colspan='7' style='padding-top: 0.25rem; padding-bottom: 1rem; color: #475569;'><strong style='color: #0f172a;'>AVANCE:</strong> " & rs(8) & "</td>"
		'[UI-ANT] html = html & "</tr>"

		'[UI-ANT] rs.movenext
	'[UI-ANT] loop

	'[UI-ANT] 'html = html & "</table><br>"
	'[UI-ANT] html = html & "</tbody>"
	'[UI-ANT] html = html & "</table>"
	
	'[UI-ANT] html = html & "</div>" ' Fin .bloque-datos table-responsive
	'[UI-ANT] html = html & "</div>" ' Fin .inner (abierto en el fragmento anterior)

'[UI-ANT] end if

'[UI-ANT] end if


'[UI-ANT] if sProy = 1 then

'[UI-ANT] iSql = "sp_cg_Reportes_Detalle @Modo='proyectos'"

'[UI-ANT] if sCli <> 0 then
	'[UI-ANT] iSql = iSql & ", @Cli=" & sCli
'[UI-ANT] else
	'[UI-ANT] iSql = iSql & ", @Cli=0"
'[UI-ANT] end if

'[UI-ANT] Set rs = Cnn.Execute(iSql)

'[UI-ANT] if not rs.eof then

'[UI-ANT] html = html & "<table style='width:900px;color:#000000;font-family:Serif;text-align:center;' cellspacing=0 cellpadding=0 align=center >"

'[UI-ANT] html = html & "<tr>"

'[UI-ANT] html = html & "<td colspan=7 style='font-size:17pt;' align=left><b><u>PROYECTOS</u></b></td>"

'[UI-ANT] html = html & "</tr>"

'[UI-ANT] html = html & "<tr>"

'[UI-ANT] html = html & "<td colspan=7 >&nbsp;</td>"

'[UI-ANT] html = html & "</tr>"

'[UI-ANT] html = html & "<tr>"

'[UI-ANT] html = html & "</table>"

'[UI-ANT] html = html & "<table class='tabl' style='width:900px;color:#000000;font-family:Serif;' cellspacing=0 cellpadding=0 align=center >"

'[UI-ANT] do while not rs.eof

'[UI-ANT] sBco = ""
'[UI-ANT] if rs(9) = 38 then
	'[UI-ANT] sBco = "ENTRE RIOS"
'[UI-ANT] elseif rs(9) = 39 then
	'[UI-ANT] sBco = "SAN JUAN"
'[UI-ANT] elseif rs(9) = 40 then
	'[UI-ANT] sBco = "SANTA CRUZ"
'[UI-ANT] elseif rs(9) = 41 then
	'[UI-ANT] sBco = "SANTA FE"
'[UI-ANT] elseif rs(9) = 4 then
	'[UI-ANT] sBco = "PETERSEN"
'[UI-ANT] else
	'[UI-ANT] sBco = rs(6)
'[UI-ANT] end if

'[UI-ANT] html = html & "<tr>"

'[UI-ANT] html = html & "<td style='witdh:1%;'><b>ID</b></td>"
'[UI-ANT] html = html & "<td style='witdh:26%;'><b>CONTACTO</b></td>"
'[UI-ANT] html = html & "<td style='witdh:23%;'><b>RESPONSABLE</b></td>"
'[UI-ANT] html = html & "<td style='witdh:10%;'><b>FEC ING</b></td>"
'[UI-ANT] html = html & "<td style='witdh:10%;'><b>FEC RESOL</b></td>"
'[UI-ANT] html = html & "<td style='witdh:10%;'><b>TIPO</b></td>"
'[UI-ANT] html = html & "<td style='witdh:20%;'><b>BANCO</b></td>"

'[UI-ANT] html = html & "</tr>"

'[UI-ANT] html = html & "<tr>"

'[UI-ANT] html = html & "<td>" & rs(0) & "</td>"
'[UI-ANT] html = html & "<td>" & rs(1) & "</td>"
'[UI-ANT] html = html & "<td>" & rs(2) & "</td>"



'[UI-ANT] html = html & "<td>" & rs(3) & "</td>"
'[UI-ANT] html = html & "<td>" & rs(4) & "</td>"
'[UI-ANT] if Trim(rs(5) & "") = "0" then
	'[UI-ANT] xtipo = "CONSULTA"
'[UI-ANT] else
	'[UI-ANT] xtipo = "DESARROLLO"
'[UI-ANT] end if
'[UI-ANT] html = html & "<td>" & xtipo & "</td>"
'[UI-ANT] html = html & "<td>" & rs(6) & "</td>"

'[UI-ANT] html = html & "</tr>"

'[UI-ANT] html = html & "<tr>"

'[UI-ANT] html = html & "<td colspan=7><b>DETALLE:</b></td>"

'[UI-ANT] html = html & "</tr>"

'[UI-ANT] html = html & "<tr>"

'[UI-ANT] html = html & "<td colspan=7>" & rs(7) & "</td>"

'[UI-ANT] html = html & "</tr>"

'[UI-ANT] html = html & "<tr>"

'[UI-ANT] html = html & "<td colspan=7><b>AVANCE:</b></td>"

'[UI-ANT] html = html & "</tr>"

'[UI-ANT] html = html & "<tr>"

'[UI-ANT] html = html & "<td style='border-bottom:solid 1px #aaaaaa;' colspan=7>" & rs(8) & "</td>"

'[UI-ANT] html = html & "</tr>"

'[UI-ANT] rs.movenext
'[UI-ANT] loop

'[UI-ANT] html = html & "</table><br>"

'[UI-ANT] end if

'[UI-ANT] end if

'[UI-ANT] if sApl = 1 then

'[UI-ANT] iSql = "sp_cg_Reportes_Detalle @Modo='apl'"

'[UI-ANT] if sCli <> 0 then
	'[UI-ANT] iSql = iSql & ", @Cli=" & sCli
'[UI-ANT] else
	'[UI-ANT] iSql = iSql & ", @Cli=0"
'[UI-ANT] end if

'[UI-ANT] Set rs = Cnn.Execute(iSql)

'[UI-ANT] if not rs.eof then

'[UI-ANT] html = html & "<table style='width:900px;color:#000000;font-family:Serif;text-align:center;' cellspacing=0 align=center border>"

'[UI-ANT] html = html & "<tr>"

'[UI-ANT] html = html & "<td colspan=4><b>APLs</b></td>"

'[UI-ANT] html = html & "</tr>"

'[UI-ANT] html = html & "<tr>"

'[UI-ANT] html = html & "<td style='witdh:5%;'><b>ID</b></td>"
'[UI-ANT] html = html & "<td style='witdh:31%;'><b>NOMBRE</b></td>"
'[UI-ANT] html = html & "<td style='witdh:33%;'><b>RESPONSABLE</b></td>"
'[UI-ANT] html = html & "<td style='witdh:31%;'><b>FECHA ENVIO</b></td>"

'[UI-ANT] html = html & "</tr>"

'[UI-ANT] do while not rs.eof

'[UI-ANT] html = html & "<tr>"

'[UI-ANT] html = html & "<td>" & rs(0) & "</td>"
'[UI-ANT] html = html & "<td>" & rs(1) & "</td>"
'[UI-ANT] html = html & "<td>" & rs(2) & "</td>"
'[UI-ANT] html = html & "<td>" & rs(3) & "</td>"

'[UI-ANT] html = html & "</tr>"

'[UI-ANT] rs.movenext
'[UI-ANT] loop

'[UI-ANT] html = html & "</table>"

'[UI-ANT] end if

'[UI-ANT] end if

'[UI-ANT] html = html & "<script type='text/javascript'>"
'[UI-ANT] html = html & " function borrar(formu){ "
'[UI-ANT] html = html & " var r = confirm('Desea eliminar proyecto?'); "
'[UI-ANT] html = html & " if (r){ "
'[UI-ANT] html = html & " document.getElementById(formu).submit(); "
'[UI-ANT] html = html & " }else{ "
'[UI-ANT] html = html & " return false;"
'[UI-ANT] html = html & " }} "
'[UI-ANT] html = html & "</script>"


'[UI-ANT] html = html & "<script src='assets/js/jquery.min.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/skel.min.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/util.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/main.js'></script>"

'[UI-ANT] html = html & formx22 & formx23 & xhtml & formx10 & xformx & formx1 & formx2 & formx3 & formx29 & formx19 & formx50 & "</body>"
'[UI-ANT] html = html & "</html>"

'[UI-ANT] response.write html

' =========================================================================
' [UI] PANTALLA con el sistema visual (IncFile/ui_componentes.inc)
' Reporte generado (incidentes y consultas; proyectos y APL si se piden).
' El cuerpo anterior completo queda arriba, comentado con '[UI-ANT].
' Consultas: las mismas de antes. Cambios: cliente y fechas validados, Volver (no habia),
' Imprimir, aviso sin resultados y textos del detalle y avance seguros (con sus saltos de linea).
' =========================================================================
Dim accTop, nCli, sFecDesde, sFecHasta, hayDatos, cantInc, cantCons, cantDes, items, xtipo, claseTipo, sBco

html = headhtml()
html = html & "<body>"

nFecini = 0
If nFecini < CLng(CDate("01-01-1900")) Then
	nFecini = CLng(CDate("01-01-1900"))
Else
	nFecini = 0
End If

sCli = request.form("cli")
if sCli & "" = "" then sCli = request.querystring("cli")

sInc = 1

sProy = request.form("proy")
if sProy & "" = "" then sProy = request.querystring("proy")

sApl = request.form("apl")
if sApl & "" = "" then sApl = request.querystring("apl")

' [UI] Cliente validado como numero y fechas verificadas: antes un texto iba a la consulta y una fecha mal escrita rompia la pagina
If IsNumeric(sCli & "") And Trim(sCli & "") <> "" Then sCli = CLng(sCli) Else sCli = 0
If sProy & "" <> "1" Then sProy = 0
If sApl & "" <> "1" Then sApl = 0

sFDesde = request.form("fecd")
if sFDesde & "" = "" then sFDesde = request.querystring("fecd")
If IsDate(sFDesde & "") Then sFDesde = CLng(CDate(sFDesde)) - nFecini Else sFDesde = ""

sFHasta = request.form("fech")
if sFHasta & "" = "" then sFHasta = request.querystring("fech")
If IsDate(sFHasta & "") Then sFHasta = CLng(CDate(sFHasta)) - nFecini Else sFHasta = ""

sOpc = request.querystring("opc")

If sOpc = "exp" Then
	Response.ContentType = "application/vnd.ms-excel"
	Response.AddHeader "content-disposition", " filename=reporte.xls"
End If

Set Cnn = Server.CreateObject("ADODB.Connection")
Cnn.open scon1

' Rango del reporte, para el encabezado
If sFDesde & "" <> "" Then sFecDesde = UiFecha(sFDesde) Else sFecDesde = "Sin l&iacute;mite"
If sFHasta & "" <> "" Then sFecHasta = UiFecha(sFHasta) Else sFecHasta = "Sin l&iacute;mite"
If sCli <> 0 Then nCli = UiTexto(getCli(sCli)) Else nCli = "Todos los clientes"

If sOpc <> "exp" Then
	html = html & "<div id='wrapper' role='banner'>" & cabecera() & "</div>"
	' [UI] Antes no habia forma de volver (el boton estaba comentado)
	formx10 = "<form id='formx10' class='form-oculto' method='post' action='reportes.asp'></form>"
	accTop = UiBoton("Imprimir", "light", "window.print();", True, "") & UiBoton("Volver", "secondary", "document.getElementById('formx10').submit();", True, "")
	html = html & UiPantallaAbre("Reporte generado", accTop, "")
Else
	html = html & "<main>"
End If

' --- Parametros del reporte ---
html = html & "<section class='ui-card' aria-labelledby='tituloParam'><h2 class='ui-card__title' id='tituloParam'>Par&aacute;metros del reporte</h2><div class='ui-fields'>"
html = html & UiCampoLectura("Cliente", nCli, False)
html = html & UiCampoLectura("Desde", sFecDesde, False)
html = html & UiCampoLectura("Hasta", sFecHasta, False)
html = html & "</div></section>"

hayDatos = False

' =============================================================================
' INCIDENTES Y CONSULTAS
' =============================================================================
if sInc = 1 then

	iSql = "sp_cg_Reportes_Detalle @Modo='incidentes', @Cli=" & sCli
	if sFDesde & "" <> "" then iSql = iSql & ", @FDesde=" & sFDesde
	if sFHasta & "" <> "" then iSql = iSql & ", @FHasta=" & sFHasta

	Set rs = Cnn.Execute(iSql)

	items = ""
	cantInc = 0
	cantCons = 0
	cantDes = 0
	do while not rs.eof
		' Bancos del Grupo Petersen: se muestra el banco (el mismo criterio de antes)
		sBco = RepBanco(rs(9), rs(6), True)
		if Trim(rs(5) & "") = "0" then
			xtipo = "Consulta"
			claseTipo = "blue"
			cantCons = cantCons + 1
		else
			xtipo = "Desarrollo"
			claseTipo = "purple"
			cantDes = cantDes + 1
		end if
		items = items & "<article class='ui-reporte__item'>" & _
			"<header class='ui-reporte__head'><span class='ui-tag-id'>INC-" & rs(0) & "</span>" & UiBadge(xtipo, claseTipo) & "<span class='ui-reporte__cliente'>" & sBco & "</span></header>" & _
			"<dl class='ui-reporte__datos'>" & _
				"<div><dt>Contacto</dt><dd>" & UiDato(rs(1)) & "</dd></div>" & _
				"<div><dt>Responsable</dt><dd>" & UiDato(rs(2)) & "</dd></div>" & _
				"<div><dt>Ingreso</dt><dd>" & UiFecha(rs(3)) & "</dd></div>" & _
				"<div><dt>Resoluci&oacute;n</dt><dd>" & UiFecha(rs(4)) & "</dd></div>" & _
			"</dl>" & _
			RepTexto("Detalle", rs(7)) & RepTexto("Avance", rs(8)) & _
			"</article>"
		cantInc = cantInc + 1
		rs.movenext
	loop

	If cantInc > 0 Then
		hayDatos = True
		html = html & "<section class='ui-card' aria-labelledby='tituloInc'><h2 class='ui-card__title' id='tituloInc'>Incidentes y consultas</h2>"
		html = html & "<div class='ui-reporte'>" & items & "</div>"
		html = html & UiPie(RepCant(cantInc, "caso", "casos") & " &middot; " & cantDes & " de desarrollo &middot; " & RepCant(cantCons, "consulta", "consultas"), "")
		html = html & "</section>"
	End If

end if

' =============================================================================
' PROYECTOS (solo si se pide con proy=1)
' =============================================================================
if sProy = 1 then

	iSql = "sp_cg_Reportes_Detalle @Modo='proyectos', @Cli=" & sCli
	Set rs = Cnn.Execute(iSql)

	items = ""
	cantInc = 0
	do while not rs.eof
		sBco = RepBanco(rs(9), rs(6), False)
		if Trim(rs(5) & "") = "0" then xtipo = "Consulta" else xtipo = "Desarrollo"
		items = items & "<article class='ui-reporte__item'>" & _
			"<header class='ui-reporte__head'><span class='ui-tag-id ui-tag-id--pry'>PRY-" & rs(0) & "</span>" & UiBadge(xtipo, "gray") & "<span class='ui-reporte__cliente'>" & sBco & "</span></header>" & _
			"<dl class='ui-reporte__datos'>" & _
				"<div><dt>Contacto</dt><dd>" & UiDato(rs(1)) & "</dd></div>" & _
				"<div><dt>Responsable</dt><dd>" & UiDato(rs(2)) & "</dd></div>" & _
				"<div><dt>Ingreso</dt><dd>" & UiFecha(rs(3)) & "</dd></div>" & _
				"<div><dt>Resoluci&oacute;n</dt><dd>" & UiFecha(rs(4)) & "</dd></div>" & _
			"</dl>" & _
			RepTexto("Detalle", rs(7)) & RepTexto("Avance", rs(8)) & _
			"</article>"
		cantInc = cantInc + 1
		rs.movenext
	loop

	If cantInc > 0 Then
		hayDatos = True
		html = html & "<section class='ui-card' aria-labelledby='tituloPry'><h2 class='ui-card__title' id='tituloPry'>Proyectos</h2>"
		html = html & "<div class='ui-reporte'>" & items & "</div>" & UiPie(RepCant(cantInc, "proyecto", "proyectos"), "") & "</section>"
	End If

end if

' =============================================================================
' APL (solo si se pide con apl=1)
' =============================================================================
if sApl = 1 then

	iSql = "sp_cg_Reportes_Detalle @Modo='apl', @Cli=" & sCli
	Set rs = Cnn.Execute(iSql)

	items = ""
	cantInc = 0
	do while not rs.eof
		items = items & "<tr><td>" & rs(0) & "</td><td><span class='ui-table__titulo'>" & UiTexto(rs(1)) & "</span></td><td>" & UiDato(rs(2)) & "</td><td>" & UiFecha(rs(3)) & "</td></tr>"
		cantInc = cantInc + 1
		rs.movenext
	loop

	If cantInc > 0 Then
		hayDatos = True
		html = html & UiCardAbre("APL", "") & UiTablaAbre("APL del reporte")
		html = html & "<thead><tr>" & UiTh("ID", "") & UiTh("Nombre", "") & UiTh("Responsable", "") & UiTh("Fecha de env&iacute;o", "") & "</tr></thead><tbody>" & items
		html = html & UiTablaCierra() & UiPie(cantInc & " APL", "") & UiCardCierra()
	End If

end if

' [UI] Antes, sin resultados la pagina quedaba en blanco
If Not hayDatos Then
	html = html & UiCardAbre("Sin resultados", "") & "<div class='ui-empty'>No hay incidentes para el cliente y el per&iacute;odo elegidos.</div>" & UiCardCierra()
End If

If sOpc <> "exp" Then
	html = html & UiPantallaCierra()
Else
	html = html & "</main>"
End If

html = html & "<script src='assets/js/jquery.min.js'></script>"
html = html & "<script src='assets/js/skel.min.js'></script>"
html = html & "<script src='assets/js/util.js'></script>"
html = html & "<script src='assets/js/main.js'></script>"

html = html & formx10 & "</body>"
html = html & "</html>"

response.write html


' =============================================================================
' FUNCIONES DE ESTA PANTALLA
' =============================================================================

' Nombre del banco: los del Grupo Petersen por su id (mismo criterio de antes), el resto como viene
Function RepBanco(idBanco, nombre, completo)
	Dim prefijo
	If completo Then prefijo = "BCO " Else prefijo = ""
	Select Case CStr(idBanco & "")
		Case "38"
			RepBanco = prefijo & "ENTRE RIOS"
		Case "39"
			RepBanco = prefijo & "SAN JUAN"
		Case "40"
			RepBanco = prefijo & "SANTA CRUZ"
		Case "41"
			RepBanco = prefijo & "SANTA FE"
		Case "4"
			If completo Then RepBanco = "GRUPO PETERSEN" Else RepBanco = "PETERSEN"
		Case Else
			RepBanco = UiTexto(nombre)
	End Select
End Function

' Cantidad con su palabra en singular o plural
Function RepCant(n, singular, plural)
	If n = 1 Then RepCant = "1 " & singular Else RepCant = n & " " & plural
End Function

' Bloque de texto largo (detalle, avance): escapado y con sus saltos de linea
Function RepTexto(titulo, valor)
	Dim t
	t = Trim(valor & "")
	If t = "" Then
		RepTexto = "<div class='ui-reporte__texto'><h3>" & titulo & "</h3><p class='ui-status-empty'>Sin " & LCase(titulo) & " cargado.</p></div>"
	Else
		t = Server.HTMLEncode(t)
		' Los <br> que pueda traer la base se respetan como saltos de linea; cualquier otra etiqueta queda como texto
		t = Replace(Replace(Replace(t, "&lt;br&gt;", "<br>", 1, -1, 1), "&lt;br/&gt;", "<br>", 1, -1, 1), "&lt;br /&gt;", "<br>", 1, -1, 1)
		t = Replace(Replace(t, vbCrLf, "<br>"), vbLf, "<br>")
		RepTexto = "<div class='ui-reporte__texto'><h3>" & titulo & "</h3><p>" & Replace(t, vbCr, "<br>") & "</p></div>"
	End If
End Function
%>