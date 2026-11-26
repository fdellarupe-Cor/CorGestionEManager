<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/ui_componentes.inc" -->
<%
'[UI-ANT] ' ===== Cuerpo anterior completo: reemplazado por el bloque [UI] que sigue. Se borra entero buscando [UI-ANT] =====

'[UI-ANT] html = headhtml()
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

'[UI-ANT] tipo = request.querystring("tipo")
'[UI-ANT] sel = request.querystring("sel")
'[UI-ANT] xmenu = request.querystring("xmenu")
'[UI-ANT] xtmenu = request.querystring("xtmenu")

'[UI-ANT] nFecini = 0

'[UI-ANT] If nFecini < CLng(CDate("01-01-1900")) Then
	'[UI-ANT] nFecini = CLng(CDate("01-01-1900"))
'[UI-ANT] Else
	'[UI-ANT] nFecini = 0
'[UI-ANT] End If

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

'[UI-ANT] html = html & "<td><h2>ESTADISTICAS DETALLE</h2></td>"

'[UI-ANT] formx04 = "<form id='formx04' method='post' action='estadisticas.asp?sel=" & sel & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
'[UI-ANT] botonx04 = "<input class='bot' type='button' value='VOLVER' onclick=""document.getElementById('formx04').submit();"">"
'[UI-ANT] html = html & "<td align=center>" & botonx04 & "</td>"

'[UI-ANT] html = html & "</tr>"

'[UI-ANT] html = html & "</table>"

'[UI-ANT] html = html & "</div>"

'[UI-ANT] 'CUERPO DETALLE ESTADIST

'[UI-ANT] if tipo = "I" then

'[UI-ANT] XXX = SYSLOGON(cnn,"ESTADIST_INC")

'[UI-ANT] 'html = html & "<div style='height:75%;width:100%;'><br>"
'[UI-ANT] html = html & "<div class='inner' style='padding: 2rem;'>"

'[UI-ANT] cmd = "sp_cg_Estadisticas @Modo='detalle_I'"

'[UI-ANT] 'WHERE ID
'[UI-ANT] id = request.form("id")
'[UI-ANT] if id & "" <> "" then
	'[UI-ANT] cmd = cmd & ", @Id=" & id
'[UI-ANT] end if

'[UI-ANT] 'WHERE CLIENTE
'[UI-ANT] cli = request.form("cli")
'[UI-ANT] if cli <> 0 then
	'[UI-ANT] cmd = cmd & ", @Cli=" & cli
'[UI-ANT] end if

'[UI-ANT] 'WHERE ULTIMO USUARIO QUE ACCEDIO
'[UI-ANT] uau = request.form("uau")
'[UI-ANT] if Trim(uau & "") <> "0" And Trim(uau & "") <> "" then
	'[UI-ANT] cmd = cmd & ", @Uau=" & SqlStr(uau)
'[UI-ANT] end if

'[UI-ANT] 'WHERE ULT TIPO ACCESO
'[UI-ANT] uta = request.form("uta")
'[UI-ANT] if cstr(uta) & "" <> "" then
	'[UI-ANT] cmd = cmd & ", @Uta='" & uta & "'"
'[UI-ANT] end if

'[UI-ANT] 'WHERE FECHA CREACION DESDE
'[UI-ANT] fcd = request.form("fcd")
'[UI-ANT] if fcd & "" <> "" then
	'[UI-ANT] fcd = clng(cdate(fcd)) - nFecini
	'[UI-ANT] cmd = cmd & ", @Fcd=" & fcd
'[UI-ANT] end if

'[UI-ANT] 'WHERE FECHA CREACION HASTA
'[UI-ANT] fch = request.form("fch")
'[UI-ANT] if fch & "" <> "" then
	'[UI-ANT] fch = clng(cdate(fch)) - nFecini
	'[UI-ANT] cmd = cmd & ", @Fch=" & fch
'[UI-ANT] end if

'[UI-ANT] 'WHERE FECHA ULTIMO ACCESO DESDE
'[UI-ANT] fad = request.form("fad")
'[UI-ANT] if fad & "" <> "" then
	'[UI-ANT] cmd = cmd & ", @Fad='" & fad & "'"
'[UI-ANT] end if

'[UI-ANT] 'WHERE FECHA ULTIMO ACCESO HASTA
'[UI-ANT] fah = request.form("fah")
'[UI-ANT] if fah & "" <> "" then
	'[UI-ANT] cmd = cmd & ", @Fah='" & fah & "'"
'[UI-ANT] end if

'[UI-ANT] Set rs = Exec(cnn, cmd)

'[UI-ANT] if not rs.eof then

'[UI-ANT] do while not rs.eof

	'[UI-ANT] ' --- INICIO TARJETA DE INCIDENTE ---
	'[UI-ANT] 'html = html & "<table class='tabl' xstyle='width:80%;font-size:10pt;color:#000000;font-family:Serif;' align=center cellspacing=0>"
	'[UI-ANT] html = html & "<div class='bloque-datos' style='margin-bottom: 2rem; max-width: 1000px; margin-left: auto; margin-right: auto;'>"

	'[UI-ANT] ' Encabezado con enlace
	'[UI-ANT] if getPermisos(9,"INICIO","CONS") = 1 then
			'[UI-ANT] ref = "<a class='col-link' href='detalles.asp?opc=veri&id=" & rs(0) & "'>INCIDENTE " & rs(0) & "</a>"
	'[UI-ANT] else
			'[UI-ANT] ref = "<a class='col-link' onclick=""alert('Sin permiso para consultar.');"">INCIDENTE " & rs(0) & "</a>"
	'[UI-ANT] end if

	'[UI-ANT] 'html = html & "<tr><td class='fil' style='font-size:12pt;' colspan=3 align=center>" & ref & "</td></tr>"
	'[UI-ANT] html = html & "<div class='bloque-header' style='text-align: center; color: #ef4444;'>" & ref & "</div>"
			
	'[UI-ANT] html = html & "<div class='bloque-body'>"
				
		'[UI-ANT] ' Cliente (Fila destacada)
		'[UI-ANT] 'html = html & "<tr><td class='filc' colspan=3>CLIENTE: " & rs(1) & "</td></tr>"
		'[UI-ANT] html = html & "<div style='margin-bottom: 1rem; border-bottom: 1px solid #e2e8f0; padding-bottom: 0.5rem;'>"
		'[UI-ANT] html = html & "<span class='dato-label'>CLIENTE:</span> <span class='dato-value'>" & rs(1) & "</span>"
		'[UI-ANT] html = html & "</div>"

		'[UI-ANT] ' Grilla de 3 columnas para los datos estad?sticos
		'[UI-ANT] html = html & "<div style='display: grid; grid-template-columns: repeat(3, 1fr); gap: 1rem;'>"
					
			'[UI-ANT] ' Columna 1
			'[UI-ANT] html = html & "<div>"
			'[UI-ANT] html = html & "<div class='dato-label'>CANTIDAD ACCESOS</div><div class='dato-value'>" & rs(2) & "</div>"
			'[UI-ANT] html = html & "<div class='dato-label' style='margin-top:0.8rem;'>GESTOR</div><div class='dato-value'>" & rs(5) & "</div>"
			'[UI-ANT] html = html & "<div class='dato-label' style='margin-top:0.8rem;'>FECHA CREACION</div><div class='dato-value'>" & Fechador(rs(8) + nFecini) & "</div>"
			'[UI-ANT] html = html & "</div>"

			'[UI-ANT] ' Columna 2
			'[UI-ANT] html = html & "<div>"
			'[UI-ANT] html = html & "<div class='dato-label'>ULTIMO ACCESO</div><div class='dato-value'>" & rs(3) & "</div>"
			'[UI-ANT] html = html & "<div class='dato-label' style='margin-top:0.8rem;'>TECNICO</div><div class='dato-value'>" & rs(6) & "</div>"
					
			'[UI-ANT] sFecFunc = "S/D"
			'[UI-ANT] if rs(9) <> 0 then sFecFunc = Fechador(rs(9) + nFecini)
			'[UI-ANT] html = html & "<div class='dato-label' style='margin-top:0.8rem;'>FECHA OK FUNCIONAL</div><div class='dato-value'>" & sFecFunc & "</div>"
			'[UI-ANT] html = html & "</div>"

			'[UI-ANT] ' Columna 3
			'[UI-ANT] html = html & "<div>"
			'[UI-ANT] html = html & "<div class='dato-label'>ULTIMO TIPO ACC</div><div class='dato-value'>" & rs(4) & "</div>"
			'[UI-ANT] html = html & "<div class='dato-label' style='margin-top:0.8rem;'>ULTIMO ACCESO (USR)</div><div class='dato-value'>" & rs(7) & "</div>"
					
			'[UI-ANT] sFecDesa = "S/D"
			'[UI-ANT] if rs(10) <> 0 then sFecDesa = Fechador(rs(10) + nFecini)
			'[UI-ANT] html = html & "<div class='dato-label' style='margin-top:0.8rem;'>FECHA OK DESARROLLO</div><div class='dato-value'>" & sFecDesa & "</div>"
			'[UI-ANT] html = html & "</div>"

		'[UI-ANT] html = html & "</div>" ' Fin grilla
	'[UI-ANT] html = html & "</div>" ' Fin bloque-body
	'[UI-ANT] html = html & "</div>" ' Fin tarjeta
	'[UI-ANT] 'html = html & "</table><br>"

'[UI-ANT] rs.movenext
'[UI-ANT] loop



'[UI-ANT] else

	'[UI-ANT] ' --- ESTADO SIN DATOS NORMALIZADO ---
	'[UI-ANT] 'html = html & "<table style='width:50%;font-size:10pt;color:#000000;font-family:Serif;' align=center cellspacing=0>"
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='brd1' colspan=3 align=center>SIN DATOS</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] 'html = html & "</table>"

	'[UI-ANT] html = html & "<div class='bloque-datos' style='max-width: 500px; margin: 2rem auto; border: 1px dashed #cbd5e1; background: #f8fafc;'>"
	'[UI-ANT] html = html & "    <div class='bloque-body' style='text-align: center; padding: 3rem;'>"
	'[UI-ANT] html = html & "        <div style='font-size: 3rem; margin-bottom: 1rem;'>??</div>"
	'[UI-ANT] html = html & "        <div style='color: #64748b; font-weight: 600; text-transform: uppercase; letter-spacing: 0.05em;'>No se encontraron resultados</div>"
	'[UI-ANT] html = html & "        <div style='color: #94a3b8; font-size: 0.9rem; margin-top: 0.5rem;'>Prob? ajustando los filtros de b?squeda e intent? de nuevo.</div>"
	'[UI-ANT] html = html & "    </div>"
	'[UI-ANT] html = html & "</div>"

'[UI-ANT] end if

'[UI-ANT] html = html & "</div>"

'[UI-ANT] elseif tipo = "P" then

'[UI-ANT] XXX = SYSLOGON(cnn,"ESTADIST_PROY")

'[UI-ANT] 'html = html & "<div style='height:75%;width:100%;'><br>"
'[UI-ANT] html = html & "<div class='inner' style='padding: 2rem;'>"

'[UI-ANT] cmd = "sp_cg_Estadisticas @Modo='detalle_P'"

'[UI-ANT] 'WHERE ID
'[UI-ANT] id = request.form("id")
'[UI-ANT] if id & "" <> "" then
	'[UI-ANT] cmd = cmd & ", @Id=" & id
'[UI-ANT] end if

'[UI-ANT] 'WHERE CLIENTE
'[UI-ANT] cli = request.form("cli")
'[UI-ANT] if cli <> 0 then
	'[UI-ANT] cmd = cmd & ", @Cli=" & cli
'[UI-ANT] end if

'[UI-ANT] 'WHERE ULTIMO USUARIO QUE ACCEDIO
'[UI-ANT] uau = request.form("uau")
'[UI-ANT] if Trim(uau & "") <> "0" And Trim(uau & "") <> "" then
	'[UI-ANT] cmd = cmd & ", @Uau=" & SqlStr(uau)
'[UI-ANT] end if

'[UI-ANT] 'WHERE ULT TIPO ACCESO
'[UI-ANT] uta = request.form("uta")
'[UI-ANT] if cstr(uta) & "" <> "" then
	'[UI-ANT] cmd = cmd & ", @Uta='" & uta & "'"
'[UI-ANT] end if

'[UI-ANT] 'WHERE FECHA CREACION DESDE
'[UI-ANT] fcd = request.form("fcd")
'[UI-ANT] if fcd & "" <> "" then
	'[UI-ANT] fcd = clng(cdate(fcd)) - nFecini
	'[UI-ANT] cmd = cmd & ", @Fcd=" & fcd
'[UI-ANT] end if

'[UI-ANT] 'WHERE FECHA CREACION HASTA
'[UI-ANT] fch = request.form("fch")
'[UI-ANT] if fch & "" <> "" then
	'[UI-ANT] fch = clng(cdate(fch)) - nFecini
	'[UI-ANT] cmd = cmd & ", @Fch=" & fch
'[UI-ANT] end if

'[UI-ANT] 'WHERE FECHA ULTIMO ACCESO DESDE
'[UI-ANT] fad = request.form("fad")
'[UI-ANT] if fad & "" <> "" then
	'[UI-ANT] cmd = cmd & ", @Fad='" & fad & "'"
'[UI-ANT] end if

'[UI-ANT] 'WHERE FECHA ULTIMO ACCESO HASTA
'[UI-ANT] fah = request.form("fah")
'[UI-ANT] if fah & "" <> "" then
	'[UI-ANT] cmd = cmd & ", @Fah='" & fah & "'"
'[UI-ANT] end if

'[UI-ANT] Set rs = Exec(cnn, cmd)

'[UI-ANT] if not rs.eof then



'[UI-ANT] ' --- INICIO NORMALIZACION: DETALLE ESTADISTICAS (PROYECTOS) ---
'[UI-ANT] do while not rs.eof

	'[UI-ANT] 'html = html & "<table class='tabl' align=center cellspacing=0>"
	'[UI-ANT] html = html & "<div class='bloque-datos' style='margin-bottom: 2rem; max-width: 1000px; margin-left: auto; margin-right: auto;'>"

	'[UI-ANT] if getPermisos(10,"INICIO","CONS") = 1 then
		'[UI-ANT] 'html = html & "<td class='fil' style='font-size:12pt;' colspan=3 align=center>" & ref & "</td>"
		'[UI-ANT] ref = "<a class='col-link' href='detalles.asp?opc=verp&id=" & rs(0) & "'>PROYECTO " & rs(0) & "</a>"
	'[UI-ANT] else
		'[UI-ANT] ref = "<a class='col-link' onclick=""alert('Sin permiso para consultar.');"">PROYECTO " & rs(0) & "</a>"
	'[UI-ANT] end if

	'[UI-ANT] ' Encabezado de la tarjeta
	'[UI-ANT] html = html & "<div class='bloque-header' style='text-align: center;'>" & ref & "</div>"
		
	'[UI-ANT] html = html & "<div class='bloque-body'>"
			
		'[UI-ANT] ' Cliente (Fila destacada)
		'[UI-ANT] 'html = html & "<tr><td class='filc' colspan=3>CLIENTE: " & rs(1) & "</td></tr>"
		'[UI-ANT] html = html & "<div style='margin-bottom: 1rem; border-bottom: 1px solid #e2e8f0; padding-bottom: 0.5rem;'>"
		'[UI-ANT] html = html & "<span class='dato-label'>CLIENTE:</span> <span class='dato-value'>" & rs(1) & "</span>"
		'[UI-ANT] html = html & "</div>"

		'[UI-ANT] ' Grilla de 3 columnas para los datos estad?sticos del Proyecto
		'[UI-ANT] html = html & "<div style='display: grid; grid-template-columns: repeat(3, 1fr); gap: 1rem;'>"
				
			'[UI-ANT] ' Columna 1
			'[UI-ANT] 'html = html & "<td class='fil'>CANTIDAD ACCESOS: " & rs(2) & "</td>"
			'[UI-ANT] 'html = html & "<td class='fil'>GESTOR: " & rs(5) & "</td>"
			'[UI-ANT] 'html = html & "<td class='fil'>FECHA CREACION: " & rs(8) & "</td>"
			'[UI-ANT] html = html & "<div>"
			'[UI-ANT] html = html & "<div class='dato-label'>CANTIDAD ACCESOS</div><div class='dato-value'>" & rs(2) & "</div>"
			'[UI-ANT] html = html & "<div class='dato-label' style='margin-top:0.8rem;'>GESTOR</div><div class='dato-value'>" & rs(5) & "</div>"
			'[UI-ANT] html = html & "<div class='dato-label' style='margin-top:0.8rem;'>FECHA CREACION</div><div class='dato-value'>" & Fechador(rs(8) + nFecini) & "</div>"
			'[UI-ANT] html = html & "</div>"

			'[UI-ANT] ' Columna 2
			'[UI-ANT] 'html = html & "<td class='filb'>ULTIMO ACCESO: " & rs(3) & "</td>"
			'[UI-ANT] 'html = html & "<td class='filb'>TECNICO: " & rs(6) & "</td>"
			'[UI-ANT] 'html = html & "<td class='filb'>FECHA OK FUNCIONAL: " & rs(9) & "</td>"
			'[UI-ANT] html = html & "<div>"
			'[UI-ANT] html = html & "<div class='dato-label'>ULTIMO ACCESO</div><div class='dato-value'>" & rs(3) & "</div>"
			'[UI-ANT] html = html & "<div class='dato-label' style='margin-top:0.8rem;'>TECNICO</div><div class='dato-value'>" & rs(6) & "</div>"
			'[UI-ANT] sFecFunc = "S/D"
			'[UI-ANT] If IsNumeric(rs(9)) Then
				'[UI-ANT] If CLng(rs(9)) > 0 Then sFecFunc = Fechador(CLng(rs(9)) + nFecini)
			'[UI-ANT] End If
			'[UI-ANT] html = html & "<div class='dato-label' style='margin-top:0.8rem;'>FECHA OK FUNCIONAL</div><div class='dato-value'>" & sFecFunc & "</div>"
			'[UI-ANT] html = html & "</div>"

			'[UI-ANT] ' Columna 3
			'[UI-ANT] 'html = html & "<td class='filb'>ULTIMO TIPO ACC: " & rs(4) & "</td>"
			'[UI-ANT] 'html = html & "<td class='filb'>ULTIMO ACCESO: " & rs(7) & "</td>"
			'[UI-ANT] 'html = html & "<td class='filb'>FECHA OK DESARROLLO: " & rs(10) & "</td>"
			'[UI-ANT] html = html & "<div>"
			'[UI-ANT] html = html & "<div class='dato-label'>ULTIMO TIPO ACC</div><div class='dato-value'>" & rs(4) & "</div>"
			'[UI-ANT] html = html & "<div class='dato-label' style='margin-top:0.8rem;'>ULTIMO ACCESO (USR)</div><div class='dato-value'>" & rs(7) & "</div>"
			'[UI-ANT] sFecDesa = "S/D"
			'[UI-ANT] If IsNumeric(rs(10)) Then
				'[UI-ANT] If CLng(rs(10)) > 0 Then sFecDesa = Fechador(CLng(rs(10)) + nFecini)
			'[UI-ANT] End If
			'[UI-ANT] html = html & "<div class='dato-label' style='margin-top:0.8rem;'>FECHA OK DESARROLLO</div><div class='dato-value'>" & sFecDesa & "</div>"
			'[UI-ANT] html = html & "</div>"

		'[UI-ANT] html = html & "</div>" ' Fin grilla
	'[UI-ANT] html = html & "</div>" ' Fin bloque-body
	'[UI-ANT] html = html & "</div>" ' Fin bloque-datos
	'[UI-ANT] 'html = html & "</table><br>"

'[UI-ANT] rs.movenext
'[UI-ANT] loop



'[UI-ANT] else

'[UI-ANT] ' --- ESTADO SIN DATOS NORMALIZADO (PROYECTOS) ---
	'[UI-ANT] 'html = html & "<table style='width:50%;font-size:10pt;color:#000000;font-family:Serif;' align=center cellspacing=0>"
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='brd1' colspan=3 align=center>SIN DATOS</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] 'html = html & "</table>"

	'[UI-ANT] html = html & "<div class='bloque-datos' style='max-width: 500px; margin: 2rem auto; border: 1px dashed #cbd5e1; background: #f8fafc;'>"
	'[UI-ANT] html = html & "    <div class='bloque-body' style='text-align: center; padding: 3rem;'>"
	'[UI-ANT] html = html & "        <div style='font-size: 3rem; margin-bottom: 1rem;'>??</div>"
	'[UI-ANT] html = html & "<div style='color: #64748b; font-weight: 600; text-transform: uppercase; letter-spacing: 0.05em;'>Sin proyectos encontrados</div>"
	'[UI-ANT] html = html & "<div style='color: #94a3b8; font-size: 0.9rem; margin-top: 0.5rem;'>No hay estad?sticas de proyectos que coincidan con los filtros aplicados.</div>"
	'[UI-ANT] html = html & "    </div>"
	'[UI-ANT] html = html & "</div>"

'[UI-ANT] end if

'[UI-ANT] html = html & "</div>"

'[UI-ANT] elseif tipo = "A" then

'[UI-ANT] XXX = SYSLOGON(cnn,"ESTADIST_APL")

'[UI-ANT] 'html = html & "<div style='height:75%;width:100%;'><br>"
'[UI-ANT] html = html & "<div class='inner' style='padding: 2rem;'>"

'[UI-ANT] cmd = "sp_cg_Estadisticas @Modo='detalle_A'"

'[UI-ANT] 'WHERE ID
'[UI-ANT] id = request.form("id")
'[UI-ANT] if id & "" <> "" then
	'[UI-ANT] cmd = cmd & ", @Id=" & id
'[UI-ANT] end if

'[UI-ANT] 'WHERE CLIENTE
'[UI-ANT] cli = request.form("cli")
'[UI-ANT] if cli <> 0 then
	'[UI-ANT] cmd = cmd & ", @Cli=" & cli
'[UI-ANT] end if

'[UI-ANT] 'WHERE ULTIMO USUARIO QUE ACCEDIO
'[UI-ANT] uau = request.form("uau")
'[UI-ANT] if Trim(uau & "") <> "0" And Trim(uau & "") <> "" then
	'[UI-ANT] cmd = cmd & ", @Uau=" & SqlStr(uau)
'[UI-ANT] end if

'[UI-ANT] 'WHERE ULT TIPO ACCESO
'[UI-ANT] uta = request.form("uta")
'[UI-ANT] if cstr(uta) & "" <> "" then
	'[UI-ANT] cmd = cmd & ", @Uta='" & uta & "'"
'[UI-ANT] end if

'[UI-ANT] 'WHERE FECHA CREACION DESDE  (la base guarda DIAS desde 1900-01-01)
'[UI-ANT] fcd = request.form("fcd")
'[UI-ANT] if fcd & "" <> "" then
	'[UI-ANT] fcd = clng(cdate(fcd)) - nFecini
	'[UI-ANT] cmd = cmd & ", @Fcd=" & fcd
'[UI-ANT] end if

'[UI-ANT] 'WHERE FECHA CREACION HASTA
'[UI-ANT] fch = request.form("fch")
'[UI-ANT] if fch & "" <> "" then
	'[UI-ANT] fch = clng(cdate(fch)) - nFecini
	'[UI-ANT] cmd = cmd & ", @Fch=" & fch
'[UI-ANT] end if

'[UI-ANT] 'WHERE FECHA ULTIMO ACCESO DESDE
'[UI-ANT] fad = request.form("fad")
'[UI-ANT] if fad & "" <> "" then
	'[UI-ANT] cmd = cmd & ", @Fad='" & fad & "'"
'[UI-ANT] end if

'[UI-ANT] 'WHERE FECHA ULTIMO ACCESO HASTA
'[UI-ANT] fah = request.form("fah")
'[UI-ANT] if fah & "" <> "" then
	'[UI-ANT] cmd = cmd & ", @Fah='" & fah & "'"
'[UI-ANT] end if

'[UI-ANT] Set rs = Exec(cnn, cmd)

'[UI-ANT] if not rs.eof then

'[UI-ANT] ' --- INICIO NORMALIZACION: DETALLE ESTADISTICAS (APLs) ---
	'[UI-ANT] do while not rs.eof

		'[UI-ANT] 'html = html & "<table class='tabl' align=center cellspacing=0>"
		'[UI-ANT] html = html & "<div class='bloque-datos' style='margin-bottom: 2rem; max-width: 1000px; margin-left: auto; margin-right: auto;'>"

		'[UI-ANT] if getPermisos(11,"INICIO","CONS") = 1 then
			'[UI-ANT] 'ref = "<a class='red' href='detalles2.asp?opc=ver&id=" & rs(0) & "'><u>APL " & rs(0) & "</u></a>"
			'[UI-ANT] ref = "<a class='col-link' href='detalles2.asp?opc=ver&id=" & rs(0) & "'>APL " & rs(0) & "</a>"
		'[UI-ANT] else
			'[UI-ANT] ref = "<a class='col-link' onclick=""alert('Sin permiso para consultar.');"">APL " & rs(0) & "</a>"
		'[UI-ANT] end if

		'[UI-ANT] ' Encabezado de la tarjeta (Limpio, sin subrayado)
		'[UI-ANT] html = html & "<div class='bloque-header' style='text-align: center;'>" & ref & "</div>"
		
		'[UI-ANT] html = html & "<div class='bloque-body'>"
			
			'[UI-ANT] ' Grid de datos principales (Cliente y Fecha Creaci?n)
			'[UI-ANT] html = html & "<div class='grid-datos-doble' style='margin-bottom: 1rem;'>"
				'[UI-ANT] html = html & "<div><span class='dato-label'>CLIENTE:</span> <span class='dato-value'>" & rs(1) & "</span></div>"
				'[UI-ANT] html = html & "<div><span class='dato-label'>FECHA CREACION:</span> <span class='dato-value'>" & Fechador(rs(7) + nFecini) & "</span></div>"
			'[UI-ANT] html = html & "</div>"

			'[UI-ANT] ' Separador sutil
			'[UI-ANT] html = html & "<hr style='border:0; border-top: 1px solid #f1f5f9; margin: 1rem 0;'>"

			'[UI-ANT] ' Grid de estad?sticas de acceso (3 columnas)
			'[UI-ANT] html = html & "<div style='display: grid; grid-template-columns: repeat(3, 1fr); gap: 1rem;'>"
				
				'[UI-ANT] ' Columna 1: Accesos
				'[UI-ANT] html = html & "<div>"
				'[UI-ANT] html = html & "<div class='dato-label'>CANTIDAD ACCESOS</div><div class='dato-value'>" & rs(2) & "</div>"
				'[UI-ANT] html = html & "</div>"

				'[UI-ANT] ' Columna 2: ?ltimo Acceso (Fecha/Hora)
				'[UI-ANT] html = html & "<div>"
				'[UI-ANT] html = html & "<div class='dato-label'>ULTIMO ACCESO</div><div class='dato-value'>" & rs(3) & "</div>"
				'[UI-ANT] html = html & "</div>"

				'[UI-ANT] ' Columna 3: Tipo
				'[UI-ANT] html = html & "<div>"
				'[UI-ANT] html = html & "<div class='dato-label'>ULTIMO TIPO ACC</div><div class='dato-value'>" & rs(4) & "</div>"
				'[UI-ANT] html = html & "</div>"

			'[UI-ANT] html = html & "</div>"

			'[UI-ANT] ' Fila inferior: Responsables
			'[UI-ANT] html = html & "<div class='grid-datos-doble' style='margin-top: 1rem; padding-top: 1rem; border-top: 1px solid #f1f5f9;'>"
				'[UI-ANT] html = html & "<div><span class='dato-label'>RESPONSABLE ENVIO:</span> <span class='dato-value'>" & rs(5) & "</span></div>"
				'[UI-ANT] html = html & "<div><span class='dato-label'>ULTIMO ACCESO (USR):</span> <span class='dato-value'>" & rs(6) & "</span></div>"
			'[UI-ANT] html = html & "</div>"

		'[UI-ANT] html = html & "</div>" ' Fin bloque-body
		'[UI-ANT] html = html & "</div>" ' Fin bloque-datos
		'[UI-ANT] 'html = html & "</table><br>"

'[UI-ANT] rs.movenext
'[UI-ANT] loop



'[UI-ANT] else

	'[UI-ANT] ' --- ESTADO SIN DATOS NORMALIZADO (APLs) ---
	'[UI-ANT] html = html & "<div class='bloque-datos' style='max-width: 500px; margin: 2rem auto; border: 1px dashed #cbd5e1; background: #f8fafc;'>"
	'[UI-ANT] html = html & "    <div class='bloque-body' style='text-align: center; padding: 3rem;'>"
	'[UI-ANT] html = html & "        <div style='font-size: 3rem; margin-bottom: 1rem;'>??</div>"
	'[UI-ANT] html = html & "<div style='color: #64748b; font-weight: 600; text-transform: uppercase; letter-spacing: 0.05em;'>Sin APLs registradas</div>"
	'[UI-ANT] html = html & "<div style='color: #94a3b8; font-size: 0.9rem; margin-top: 0.5rem;'>No se encontraron movimientos para los criterios seleccionados.</div>"
	'[UI-ANT] html = html & "    </div>"
	'[UI-ANT] html = html & "</div>"

'[UI-ANT] end if

'[UI-ANT] html = html & "</div>"

'[UI-ANT] elseif tipo = "U" then

'[UI-ANT] XXX = SYSLOGON(cnn,"ESTADIST_USR")

'[UI-ANT] 'html = html & "<div style='height:75%;width:100%;'><br><br>"
'[UI-ANT] html = html & "<div class='inner' style='padding: 2rem;'>"

'[UI-ANT] id = request.form("uau")

'[UI-ANT] cmd = "sp_cg_Estadisticas @Modo='detalle_U'"

'[UI-ANT] if Trim(id & "") <> "0" And Trim(id & "") <> "" then
	'[UI-ANT] cmd = cmd & ", @Uau=" & SqlStr(id)
'[UI-ANT] end if

'[UI-ANT] Set rs = Exec(cnn, cmd)

'[UI-ANT] if not rs.eof then

'[UI-ANT] ' --- INICIO NORMALIZACION: DETALLE ESTADISTICAS (USUARIOS) ---
'[UI-ANT] do while not rs.eof

	'[UI-ANT] 'html = html & "<table class='tabl' style='width:55%;' align=center cellspacing=0>"
	'[UI-ANT] html = html & "<div class='bloque-datos' style='margin-bottom: 1.5rem; max-width: 800px; margin-left: auto; margin-right: auto;'>"

	'[UI-ANT] ' --- ENCABEZADO: NOMBRE DEL USUARIO ---
	'[UI-ANT] 'html = html & "<td class='filc' style='font-size:12pt;' colspan=3><b>" & rs(1) & "</b></td>"
	'[UI-ANT] html = html & "<div class='bloque-header' style='text-align: center; font-weight: 600;'>" & rs(1) & "</div>"
		
	'[UI-ANT] html = html & "<div class='bloque-body'>"
			
		'[UI-ANT] 'Grilla de 3 columnas para m?tricas de acceso
		'[UI-ANT] 'html = html & "<td class='fil' style='width:17%;'>CANT ACC: " & rs(3) & "</td>"
		'[UI-ANT] 'html = html & "<td class='filb' style='width:17%;'>ULT ACC: " & rs(4) & " </td>"
		'[UI-ANT] 'html = html & "<td class='filb' style='width:21%;'>ULT TIPO ACC: " & rs(5) & "</td>"
			
		'[UI-ANT] html = html & "<div style='display: grid; grid-template-columns: repeat(3, 1fr); gap: 1rem; text-align: center;'>"
				
			'[UI-ANT] ' Cantidad de Accesos
			'[UI-ANT] html = html & "<div>"
			'[UI-ANT] html = html & "<div class='dato-label'>CANT. ACCESOS</div><div class='dato-value'>" & rs(3) & "</div>"
			'[UI-ANT] html = html & "</div>"

			'[UI-ANT] ' ?ltimo Acceso (Fecha)
			'[UI-ANT] html = html & "<div>"
			'[UI-ANT] html = html & "<div class='dato-label'>ULTIMO ACCESO</div><div class='dato-value'>" & rs(4) & "</div>"
			'[UI-ANT] html = html & "</div>"

			'[UI-ANT] ' Tipo de Acceso
			'[UI-ANT] html = html & "<div>"
			'[UI-ANT] html = html & "<div class='dato-label'>TIPO ACCESO</div><div class='dato-value'>" & rs(5) & "</div>"
			'[UI-ANT] html = html & "</div>"

		'[UI-ANT] html = html & "</div>" ' Fin grilla
	'[UI-ANT] html = html & "</div>" ' Fin bloque-body
	'[UI-ANT] html = html & "</div>" ' Fin bloque-datos
	'[UI-ANT] 'html = html & "</table><br>"

'[UI-ANT] rs.movenext
'[UI-ANT] loop

'[UI-ANT] else

	'[UI-ANT] ' --- ESTADO SIN DATOS NORMALIZADO (USUARIOS) ---
	'[UI-ANT] html = html & "<div class='bloque-datos' style='max-width: 500px; margin: 2rem auto; border: 1px dashed #cbd5e1; background: #f8fafc;'>"
	'[UI-ANT] html = html & "    <div class='bloque-body' style='text-align: center; padding: 3rem;'>"
	'[UI-ANT] html = html & "        <div style='font-size: 3rem; margin-bottom: 1rem;'>??</div>"
	'[UI-ANT] html = html & "<div style='color: #64748b; font-weight: 600; text-transform: uppercase; letter-spacing: 0.05em;'>Sin actividad de usuarios</div>"
	'[UI-ANT] html = html & "<div style='color: #94a3b8; font-size: 0.9rem; margin-top: 0.5rem;'>No se registraron accesos para el usuario seleccionado en este per?odo.</div>"
	'[UI-ANT] html = html & "    </div>"
	'[UI-ANT] html = html & "</div>"

'[UI-ANT] end if

'[UI-ANT] html = html & "</div>"

'[UI-ANT] end if

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

'[UI-ANT] response.write html


' =========================================================================
' [UI] PANTALLA con el sistema visual (IncFile/ui_componentes.inc)
' Resultado de estadisticas: incidentes (I), proyectos (P), APL (A) o usuarios (U).
' El cuerpo anterior completo queda arriba, comentado con '[UI-ANT].
' Procedimientos, registro de acceso y permisos de los enlaces: los mismos de antes.
' Cambios: filtros validados (antes se pegaban tal cual en la consulta), un solo bloque
' de filtros para I/P/A (eran tres copias) y resultados en tabla con el total de accesos.
' =========================================================================
Dim accTop, titulo, cmd, filas, cantRes, totAcc, puedeVer, caso, urlCaso, etiqueta, cols, encab

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

tipo = UCase(request.querystring("tipo") & "")
sel = request.querystring("sel")
xmenu = request.querystring("xmenu")
xtmenu = request.querystring("xtmenu")
If tipo <> "I" And tipo <> "P" And tipo <> "A" And tipo <> "U" Then tipo = "I"

nFecini = 0
If nFecini < CLng(CDate("01-01-1900")) Then
	nFecini = CLng(CDate("01-01-1900"))
Else
	nFecini = 0
End If

html = html & "<div id='wrapper' role='banner'>" & cabecera() & "</div>"

formx04 = "<form id='formx04' class='form-oculto' method='post' action='estadisticas.asp?sel=" & Server.URLEncode(sel & "") & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
accTop = UiBoton("Volver", "secondary", "document.getElementById('formx04').submit();", True, "")

' =============================================================================
' CONSULTA
' =============================================================================
If tipo = "U" Then
	XXX = SYSLOGON(cnn,"ESTADIST_USR")
	titulo = "Estad&iacute;sticas de usuarios"
	cmd = "sp_cg_Estadisticas @Modo='detalle_U'"
	id = request.form("uau")
	if Trim(id & "") <> "0" And Trim(id & "") <> "" then cmd = cmd & ", @Uau=" & SqlStr(id)
Else
	If tipo = "I" Then
		XXX = SYSLOGON(cnn,"ESTADIST_INC")
		titulo = "Estad&iacute;sticas de incidentes"
	ElseIf tipo = "P" Then
		XXX = SYSLOGON(cnn,"ESTADIST_PROY")
		titulo = "Estad&iacute;sticas de proyectos"
	Else
		XXX = SYSLOGON(cnn,"ESTADIST_APL")
		titulo = "Estad&iacute;sticas de APL"
	End If
	cmd = "sp_cg_Estadisticas @Modo='detalle_" & tipo & "'"

	' [UI] Filtros validados antes de ir a la consulta
	id = request.form("id")
	if IsNumeric(id & "") And Trim(id & "") <> "" then cmd = cmd & ", @Id=" & CLng(id)

	cli = request.form("cli")
	if IsNumeric(cli & "") And Trim(cli & "") <> "" then
		if CLng(cli) <> 0 then cmd = cmd & ", @Cli=" & CLng(cli)
	end if

	uau = request.form("uau")
	if Trim(uau & "") <> "0" And Trim(uau & "") <> "" then cmd = cmd & ", @Uau=" & SqlStr(uau)

	' Tipo de acceso: solo los valores del formulario
	uta = UCase(Trim(request.form("uta") & ""))
	if uta = "CONSULTA" Or uta = "ACTUALIZA" then cmd = cmd & ", @Uta='" & uta & "'"

	fcd = request.form("fcd")
	if IsDate(fcd & "") then cmd = cmd & ", @Fcd=" & (clng(cdate(fcd)) - nFecini)

	fch = request.form("fch")
	if IsDate(fch & "") then cmd = cmd & ", @Fch=" & (clng(cdate(fch)) - nFecini)

	' Ultimo acceso: va como texto, igual que antes, pero escapado y solo si es una fecha
	fad = request.form("fad")
	if IsDate(fad & "") then cmd = cmd & ", @Fad=" & SqlStr(fad)

	fah = request.form("fah")
	if IsDate(fah & "") then cmd = cmd & ", @Fah=" & SqlStr(fah)
End If

Set rs = Exec(cnn, cmd)

' =============================================================================
' RESULTADOS
' =============================================================================
filas = ""
cantRes = 0
totAcc = 0

' Permiso para abrir el caso (los mismos menus de antes: 9 incidentes, 10 proyectos, 11 APL)
If tipo = "I" Then
	puedeVer = (getPermisosAtrm("CG_INCIDENTES", "CONS") = 1)
ElseIf tipo = "P" Then
	puedeVer = (getPermisosAtrm("CG_PROYECTOS", "CONS") = 1)
ElseIf tipo = "A" Then
	puedeVer = (getPermisosAtrm("CG_APL", "CONS") = 1)
End If

do while not rs.eof
	If IsNumeric(rs(2) & "") And tipo <> "U" Then totAcc = totAcc + CDbl(rs(2))
	If IsNumeric(rs(3) & "") And tipo = "U" Then totAcc = totAcc + CDbl(rs(3))

	If tipo = "U" Then
		filas = filas & "<tr><td><span class='ui-table__titulo'>" & UiTexto(rs(1)) & "</span></td>" & _
			"<td class='ui-table__cell--center'>" & UiDato(rs(3)) & "</td>" & _
			"<td>" & UiDato(rs(4)) & "</td><td>" & UiDato(rs(5)) & "</td></tr>"
	Else
		If tipo = "I" Then
			etiqueta = "<span class='ui-tag-id'>INC-" & rs(0) & "</span>"
			urlCaso = "detalles.asp?opc=veri&id=" & rs(0)
		ElseIf tipo = "P" Then
			etiqueta = "<span class='ui-tag-id ui-tag-id--pry'>PRY-" & rs(0) & "</span>"
			urlCaso = "detalles.asp?opc=verp&id=" & rs(0)
		Else
			etiqueta = "<span class='ui-tag-id'>APL-" & rs(0) & "</span>"
			urlCaso = "detalles2.asp?opc=ver&id=" & rs(0)
		End If
		If puedeVer Then
			caso = "<a class='ui-link' href='" & urlCaso & "'>" & etiqueta & "</a>"
		Else
			caso = etiqueta & UiSub("Sin permiso para abrirlo")
		End If

		If tipo = "A" Then
			' APL: 0 id, 1 cliente, 2 accesos, 3 ultimo acceso, 4 tipo, 5 responsable envio, 6 usuario, 7 creacion
			filas = filas & "<tr><td>" & caso & "</td><td>" & UiTexto(rs(1)) & "</td>" & _
				"<td class='ui-table__cell--center'><b>" & UiDato(rs(2)) & "</b></td>" & _
				"<td>" & UiDato(rs(3)) & UiSub(UiTexto(rs(4)) & " &middot; " & UiTexto(rs(6))) & "</td>" & _
				"<td>" & UiDato(rs(5)) & "</td>" & _
				"<td><span class='ui-table__nowrap'>" & UiFecha(rs(7)) & "</span></td></tr>"
		Else
			' Incidentes y proyectos: 0 id, 1 cliente, 2 accesos, 3 ultimo acceso, 4 tipo, 5 gestor, 6 tecnico,
			' 7 usuario, 8 creacion, 9 OK funcional, 10 OK desarrollo
			filas = filas & "<tr><td>" & caso & "</td><td>" & UiTexto(rs(1)) & "</td>" & _
				"<td class='ui-table__cell--center'><b>" & UiDato(rs(2)) & "</b></td>" & _
				"<td>" & UiDato(rs(3)) & UiSub(UiTexto(rs(4)) & " &middot; " & UiTexto(rs(7))) & "</td>" & _
				"<td>" & UiDato(rs(5)) & UiSub(UiTexto(rs(6))) & "</td>" & _
				"<td><span class='ui-table__nowrap'>" & UiFecha(rs(8)) & "</span></td>" & _
				"<td><span class='ui-table__nowrap'>" & EstFechaOk(rs(9)) & "</span></td>" & _
				"<td><span class='ui-table__nowrap'>" & EstFechaOk(rs(10)) & "</span></td></tr>"
		End If
	End If
	cantRes = cantRes + 1
	rs.movenext
loop

' --- Encabezado de la tabla segun el tipo ---
If tipo = "U" Then
	encab = UiTh("Usuario", "") & UiTh("Accesos", "ui-table__cell--center") & UiTh("&Uacute;ltimo acceso", "") & UiTh("Tipo de acceso", "")
	cols = 4
ElseIf tipo = "A" Then
	encab = UiTh("APL", "") & UiTh("Cliente", "") & UiTh("Accesos", "ui-table__cell--center") & UiTh("&Uacute;ltimo acceso", "") & UiTh("Responsable del env&iacute;o", "") & UiTh("Creaci&oacute;n", "")
	cols = 6
Else
	encab = UiTh("Caso", "") & UiTh("Cliente", "") & UiTh("Accesos", "ui-table__cell--center") & UiTh("&Uacute;ltimo acceso", "") & UiTh("Gestor / t&eacute;cnico", "") & _
		UiTh("Creaci&oacute;n", "") & UiTh("OK funcional", "") & UiTh("OK desarrollo", "")
	cols = 8
End If

html = html & UiPantallaAbre(titulo, accTop, "wide")
html = html & UiCardAbre("Resultados", "")
html = html & UiTablaAbre(titulo)
html = html & "<thead><tr>" & encab & "</tr></thead><tbody>"
If cantRes = 0 Then
	html = html & UiFilaVacia(cols, "No hay resultados para los filtros elegidos. Prob&aacute; ampliando las fechas o quitando filtros.")
Else
	html = html & filas
End If
html = html & UiTablaCierra()
If cantRes > 0 Then
	If cantRes = 1 Then
		html = html & UiPie("1 resultado &middot; " & totAcc & " accesos en total", "")
	Else
		html = html & UiPie(cantRes & " resultados &middot; " & totAcc & " accesos en total", "")
	End If
End If
html = html & UiCardCierra()
html = html & UiPantallaCierra()

html = html & "<script src='assets/js/jquery.min.js'></script>"
html = html & "<script src='assets/js/skel.min.js'></script>"
html = html & "<script src='assets/js/util.js'></script>"
html = html & "<script src='assets/js/main.js'></script>"

html = html & formx04 & "</body>"
html = html & "</html>"

response.write html


' =============================================================================
' FUNCIONES DE ESTA PANTALLA
' =============================================================================

' Fecha de OK: vacia o en cero se muestra como pendiente
Function EstFechaOk(v)
	If IsNumeric(v & "") Then
		If CDbl(v) > 0 Then
			EstFechaOk = UiFecha(v)
			Exit Function
		End If
	End If
	EstFechaOk = "<span class='ui-status-empty'>Pendiente</span>"
End Function
%>