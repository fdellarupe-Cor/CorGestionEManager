<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
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

tipo = request.querystring("tipo")
sel = request.querystring("sel")
xmenu = request.querystring("xmenu")
xtmenu = request.querystring("xtmenu")

nFecini = 0

If nFecini < CLng(CDate("01-01-1900")) Then
	nFecini = CLng(CDate("01-01-1900"))
Else
	nFecini = 0
End If

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

html = html & "<td><h2>ESTADISTICAS DETALLE</h2></td>"

formx04 = "<form id='formx04' method='post' action='estadisticas.asp?sel=" & sel & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
botonx04 = "<input class='bot' type='button' value='VOLVER' onclick=""document.getElementById('formx04').submit();"">"
html = html & "<td align=center>" & botonx04 & "</td>"

html = html & "</tr>"

html = html & "</table>"

html = html & "</div>"

'CUERPO DETALLE ESTADIST

if tipo = "I" then

XXX = SYSLOGON(cnn,"ESTADIST_INC")

'html = html & "<div style='height:75%;width:100%;'><br>"
html = html & "<div class='inner' style='padding: 2rem;'>"

cmd = "sp_cg_Estadisticas @Modo='detalle_I'"

'WHERE ID
id = request.form("id")
if id & "" <> "" then
	cmd = cmd & ", @Id=" & id
end if

'WHERE CLIENTE
cli = request.form("cli")
if cli <> 0 then
	cmd = cmd & ", @Cli=" & cli
end if

'WHERE ULTIMO USUARIO QUE ACCEDIO
uau = request.form("uau")
if Trim(uau & "") <> "0" And Trim(uau & "") <> "" then
	cmd = cmd & ", @Uau=" & SqlStr(uau)
end if

'WHERE ULT TIPO ACCESO
uta = request.form("uta")
if cstr(uta) & "" <> "" then
	cmd = cmd & ", @Uta='" & uta & "'"
end if

'WHERE FECHA CREACION DESDE
fcd = request.form("fcd")
if fcd & "" <> "" then
	fcd = clng(cdate(fcd)) - nFecini
	cmd = cmd & ", @Fcd=" & fcd
end if

'WHERE FECHA CREACION HASTA
fch = request.form("fch")
if fch & "" <> "" then
	fch = clng(cdate(fch)) - nFecini
	cmd = cmd & ", @Fch=" & fch
end if

'WHERE FECHA ULTIMO ACCESO DESDE
fad = request.form("fad")
if fad & "" <> "" then
	cmd = cmd & ", @Fad='" & fad & "'"
end if

'WHERE FECHA ULTIMO ACCESO HASTA
fah = request.form("fah")
if fah & "" <> "" then
	cmd = cmd & ", @Fah='" & fah & "'"
end if

Set rs = Exec(cnn, cmd)

if not rs.eof then

do while not rs.eof

	' --- INICIO TARJETA DE INCIDENTE ---
	'html = html & "<table class='tabl' xstyle='width:80%;font-size:10pt;color:#000000;font-family:Serif;' align=center cellspacing=0>"
	html = html & "<div class='bloque-datos' style='margin-bottom: 2rem; max-width: 1000px; margin-left: auto; margin-right: auto;'>"

	' Encabezado con enlace
	if getPermisos(9,"INICIO","CONS") = 1 then
			ref = "<a class='col-link' href='detalles.asp?opc=veri&id=" & rs(0) & "'>INCIDENTE " & rs(0) & "</a>"
	else
			ref = "<a class='col-link' onclick=""alert('Sin permiso para consultar.');"">INCIDENTE " & rs(0) & "</a>"
	end if

	'html = html & "<tr><td class='fil' style='font-size:12pt;' colspan=3 align=center>" & ref & "</td></tr>"
	html = html & "<div class='bloque-header' style='text-align: center; color: #ef4444;'>" & ref & "</div>"
			
	html = html & "<div class='bloque-body'>"
				
		' Cliente (Fila destacada)
		'html = html & "<tr><td class='filc' colspan=3>CLIENTE: " & rs(1) & "</td></tr>"
		html = html & "<div style='margin-bottom: 1rem; border-bottom: 1px solid #e2e8f0; padding-bottom: 0.5rem;'>"
		html = html & "<span class='dato-label'>CLIENTE:</span> <span class='dato-value'>" & rs(1) & "</span>"
		html = html & "</div>"

		' Grilla de 3 columnas para los datos estad�sticos
		html = html & "<div style='display: grid; grid-template-columns: repeat(3, 1fr); gap: 1rem;'>"
					
			' Columna 1
			html = html & "<div>"
			html = html & "<div class='dato-label'>CANTIDAD ACCESOS</div><div class='dato-value'>" & rs(2) & "</div>"
			html = html & "<div class='dato-label' style='margin-top:0.8rem;'>GESTOR</div><div class='dato-value'>" & rs(5) & "</div>"
			html = html & "<div class='dato-label' style='margin-top:0.8rem;'>FECHA CREACION</div><div class='dato-value'>" & Fechador(rs(8) + nFecini) & "</div>"
			html = html & "</div>"

			' Columna 2
			html = html & "<div>"
			html = html & "<div class='dato-label'>ULTIMO ACCESO</div><div class='dato-value'>" & rs(3) & "</div>"
			html = html & "<div class='dato-label' style='margin-top:0.8rem;'>TECNICO</div><div class='dato-value'>" & rs(6) & "</div>"
					
			sFecFunc = "S/D"
			if rs(9) <> 0 then sFecFunc = Fechador(rs(9) + nFecini)
			html = html & "<div class='dato-label' style='margin-top:0.8rem;'>FECHA OK FUNCIONAL</div><div class='dato-value'>" & sFecFunc & "</div>"
			html = html & "</div>"

			' Columna 3
			html = html & "<div>"
			html = html & "<div class='dato-label'>ULTIMO TIPO ACC</div><div class='dato-value'>" & rs(4) & "</div>"
			html = html & "<div class='dato-label' style='margin-top:0.8rem;'>ULTIMO ACCESO (USR)</div><div class='dato-value'>" & rs(7) & "</div>"
					
			sFecDesa = "S/D"
			if rs(10) <> 0 then sFecDesa = Fechador(rs(10) + nFecini)
			html = html & "<div class='dato-label' style='margin-top:0.8rem;'>FECHA OK DESARROLLO</div><div class='dato-value'>" & sFecDesa & "</div>"
			html = html & "</div>"

		html = html & "</div>" ' Fin grilla
	html = html & "</div>" ' Fin bloque-body
	html = html & "</div>" ' Fin tarjeta
	'html = html & "</table><br>"

rs.movenext
loop



else

	' --- ESTADO SIN DATOS NORMALIZADO ---
	'html = html & "<table style='width:50%;font-size:10pt;color:#000000;font-family:Serif;' align=center cellspacing=0>"
	'html = html & "<tr>"
	'html = html & "<td class='brd1' colspan=3 align=center>SIN DATOS</td>"
	'html = html & "</tr>"
	'html = html & "</table>"

	html = html & "<div class='bloque-datos' style='max-width: 500px; margin: 2rem auto; border: 1px dashed #cbd5e1; background: #f8fafc;'>"
	html = html & "    <div class='bloque-body' style='text-align: center; padding: 3rem;'>"
	html = html & "        <div style='font-size: 3rem; margin-bottom: 1rem;'>??</div>"
	html = html & "        <div style='color: #64748b; font-weight: 600; text-transform: uppercase; letter-spacing: 0.05em;'>No se encontraron resultados</div>"
	html = html & "        <div style='color: #94a3b8; font-size: 0.9rem; margin-top: 0.5rem;'>Prob� ajustando los filtros de b�squeda e intent� de nuevo.</div>"
	html = html & "    </div>"
	html = html & "</div>"

end if

html = html & "</div>"

elseif tipo = "P" then

XXX = SYSLOGON(cnn,"ESTADIST_PROY")

'html = html & "<div style='height:75%;width:100%;'><br>"
html = html & "<div class='inner' style='padding: 2rem;'>"

cmd = "sp_cg_Estadisticas @Modo='detalle_P'"

'WHERE ID
id = request.form("id")
if id & "" <> "" then
	cmd = cmd & ", @Id=" & id
end if

'WHERE CLIENTE
cli = request.form("cli")
if cli <> 0 then
	cmd = cmd & ", @Cli=" & cli
end if

'WHERE ULTIMO USUARIO QUE ACCEDIO
uau = request.form("uau")
if Trim(uau & "") <> "0" And Trim(uau & "") <> "" then
	cmd = cmd & ", @Uau=" & SqlStr(uau)
end if

'WHERE ULT TIPO ACCESO
uta = request.form("uta")
if cstr(uta) & "" <> "" then
	cmd = cmd & ", @Uta='" & uta & "'"
end if

'WHERE FECHA CREACION DESDE
fcd = request.form("fcd")
if fcd & "" <> "" then
	fcd = clng(cdate(fcd)) - nFecini
	cmd = cmd & ", @Fcd=" & fcd
end if

'WHERE FECHA CREACION HASTA
fch = request.form("fch")
if fch & "" <> "" then
	fch = clng(cdate(fch)) - nFecini
	cmd = cmd & ", @Fch=" & fch
end if

'WHERE FECHA ULTIMO ACCESO DESDE
fad = request.form("fad")
if fad & "" <> "" then
	cmd = cmd & ", @Fad='" & fad & "'"
end if

'WHERE FECHA ULTIMO ACCESO HASTA
fah = request.form("fah")
if fah & "" <> "" then
	cmd = cmd & ", @Fah='" & fah & "'"
end if

Set rs = Exec(cnn, cmd)

if not rs.eof then



' --- INICIO NORMALIZACION: DETALLE ESTADISTICAS (PROYECTOS) ---
do while not rs.eof

	'html = html & "<table class='tabl' align=center cellspacing=0>"
	html = html & "<div class='bloque-datos' style='margin-bottom: 2rem; max-width: 1000px; margin-left: auto; margin-right: auto;'>"

	if getPermisos(10,"INICIO","CONS") = 1 then
		'html = html & "<td class='fil' style='font-size:12pt;' colspan=3 align=center>" & ref & "</td>"
		ref = "<a class='col-link' href='detalles.asp?opc=verp&id=" & rs(0) & "'>PROYECTO " & rs(0) & "</a>"
	else
		ref = "<a class='col-link' onclick=""alert('Sin permiso para consultar.');"">PROYECTO " & rs(0) & "</a>"
	end if

	' Encabezado de la tarjeta
	html = html & "<div class='bloque-header' style='text-align: center;'>" & ref & "</div>"
		
	html = html & "<div class='bloque-body'>"
			
		' Cliente (Fila destacada)
		'html = html & "<tr><td class='filc' colspan=3>CLIENTE: " & rs(1) & "</td></tr>"
		html = html & "<div style='margin-bottom: 1rem; border-bottom: 1px solid #e2e8f0; padding-bottom: 0.5rem;'>"
		html = html & "<span class='dato-label'>CLIENTE:</span> <span class='dato-value'>" & rs(1) & "</span>"
		html = html & "</div>"

		' Grilla de 3 columnas para los datos estad�sticos del Proyecto
		html = html & "<div style='display: grid; grid-template-columns: repeat(3, 1fr); gap: 1rem;'>"
				
			' Columna 1
			'html = html & "<td class='fil'>CANTIDAD ACCESOS: " & rs(2) & "</td>"
			'html = html & "<td class='fil'>GESTOR: " & rs(5) & "</td>"
			'html = html & "<td class='fil'>FECHA CREACION: " & rs(8) & "</td>"
			html = html & "<div>"
			html = html & "<div class='dato-label'>CANTIDAD ACCESOS</div><div class='dato-value'>" & rs(2) & "</div>"
			html = html & "<div class='dato-label' style='margin-top:0.8rem;'>GESTOR</div><div class='dato-value'>" & rs(5) & "</div>"
			html = html & "<div class='dato-label' style='margin-top:0.8rem;'>FECHA CREACION</div><div class='dato-value'>" & Fechador(rs(8) + nFecini) & "</div>"
			html = html & "</div>"

			' Columna 2
			'html = html & "<td class='filb'>ULTIMO ACCESO: " & rs(3) & "</td>"
			'html = html & "<td class='filb'>TECNICO: " & rs(6) & "</td>"
			'html = html & "<td class='filb'>FECHA OK FUNCIONAL: " & rs(9) & "</td>"
			html = html & "<div>"
			html = html & "<div class='dato-label'>ULTIMO ACCESO</div><div class='dato-value'>" & rs(3) & "</div>"
			html = html & "<div class='dato-label' style='margin-top:0.8rem;'>TECNICO</div><div class='dato-value'>" & rs(6) & "</div>"
			sFecFunc = "S/D"
			If IsNumeric(rs(9)) Then
				If CLng(rs(9)) > 0 Then sFecFunc = Fechador(CLng(rs(9)) + nFecini)
			End If
			html = html & "<div class='dato-label' style='margin-top:0.8rem;'>FECHA OK FUNCIONAL</div><div class='dato-value'>" & sFecFunc & "</div>"
			html = html & "</div>"

			' Columna 3
			'html = html & "<td class='filb'>ULTIMO TIPO ACC: " & rs(4) & "</td>"
			'html = html & "<td class='filb'>ULTIMO ACCESO: " & rs(7) & "</td>"
			'html = html & "<td class='filb'>FECHA OK DESARROLLO: " & rs(10) & "</td>"
			html = html & "<div>"
			html = html & "<div class='dato-label'>ULTIMO TIPO ACC</div><div class='dato-value'>" & rs(4) & "</div>"
			html = html & "<div class='dato-label' style='margin-top:0.8rem;'>ULTIMO ACCESO (USR)</div><div class='dato-value'>" & rs(7) & "</div>"
			sFecDesa = "S/D"
			If IsNumeric(rs(10)) Then
				If CLng(rs(10)) > 0 Then sFecDesa = Fechador(CLng(rs(10)) + nFecini)
			End If
			html = html & "<div class='dato-label' style='margin-top:0.8rem;'>FECHA OK DESARROLLO</div><div class='dato-value'>" & sFecDesa & "</div>"
			html = html & "</div>"

		html = html & "</div>" ' Fin grilla
	html = html & "</div>" ' Fin bloque-body
	html = html & "</div>" ' Fin bloque-datos
	'html = html & "</table><br>"

rs.movenext
loop



else

' --- ESTADO SIN DATOS NORMALIZADO (PROYECTOS) ---
	'html = html & "<table style='width:50%;font-size:10pt;color:#000000;font-family:Serif;' align=center cellspacing=0>"
	'html = html & "<tr>"
	'html = html & "<td class='brd1' colspan=3 align=center>SIN DATOS</td>"
	'html = html & "</tr>"
	'html = html & "</table>"

	html = html & "<div class='bloque-datos' style='max-width: 500px; margin: 2rem auto; border: 1px dashed #cbd5e1; background: #f8fafc;'>"
	html = html & "    <div class='bloque-body' style='text-align: center; padding: 3rem;'>"
	html = html & "        <div style='font-size: 3rem; margin-bottom: 1rem;'>??</div>"
	html = html & "<div style='color: #64748b; font-weight: 600; text-transform: uppercase; letter-spacing: 0.05em;'>Sin proyectos encontrados</div>"
	html = html & "<div style='color: #94a3b8; font-size: 0.9rem; margin-top: 0.5rem;'>No hay estad�sticas de proyectos que coincidan con los filtros aplicados.</div>"
	html = html & "    </div>"
	html = html & "</div>"

end if

html = html & "</div>"

elseif tipo = "A" then

XXX = SYSLOGON(cnn,"ESTADIST_APL")

'html = html & "<div style='height:75%;width:100%;'><br>"
html = html & "<div class='inner' style='padding: 2rem;'>"

cmd = "sp_cg_Estadisticas @Modo='detalle_A'"

'WHERE ID
id = request.form("id")
if id & "" <> "" then
	cmd = cmd & ", @Id=" & id
end if

'WHERE CLIENTE
cli = request.form("cli")
if cli <> 0 then
	cmd = cmd & ", @Cli=" & cli
end if

'WHERE ULTIMO USUARIO QUE ACCEDIO
uau = request.form("uau")
if Trim(uau & "") <> "0" And Trim(uau & "") <> "" then
	cmd = cmd & ", @Uau=" & SqlStr(uau)
end if

'WHERE ULT TIPO ACCESO
uta = request.form("uta")
if cstr(uta) & "" <> "" then
	cmd = cmd & ", @Uta='" & uta & "'"
end if

'WHERE FECHA CREACION DESDE  (la base guarda DIAS desde 1900-01-01)
fcd = request.form("fcd")
if fcd & "" <> "" then
	fcd = clng(cdate(fcd)) - nFecini
	cmd = cmd & ", @Fcd=" & fcd
end if

'WHERE FECHA CREACION HASTA
fch = request.form("fch")
if fch & "" <> "" then
	fch = clng(cdate(fch)) - nFecini
	cmd = cmd & ", @Fch=" & fch
end if

'WHERE FECHA ULTIMO ACCESO DESDE
fad = request.form("fad")
if fad & "" <> "" then
	cmd = cmd & ", @Fad='" & fad & "'"
end if

'WHERE FECHA ULTIMO ACCESO HASTA
fah = request.form("fah")
if fah & "" <> "" then
	cmd = cmd & ", @Fah='" & fah & "'"
end if

Set rs = Exec(cnn, cmd)

if not rs.eof then

' --- INICIO NORMALIZACION: DETALLE ESTADISTICAS (APLs) ---
	do while not rs.eof

		'html = html & "<table class='tabl' align=center cellspacing=0>"
		html = html & "<div class='bloque-datos' style='margin-bottom: 2rem; max-width: 1000px; margin-left: auto; margin-right: auto;'>"

		if getPermisos(11,"INICIO","CONS") = 1 then
			'ref = "<a class='red' href='detalles2.asp?opc=ver&id=" & rs(0) & "'><u>APL " & rs(0) & "</u></a>"
			ref = "<a class='col-link' href='detalles2.asp?opc=ver&id=" & rs(0) & "'>APL " & rs(0) & "</a>"
		else
			ref = "<a class='col-link' onclick=""alert('Sin permiso para consultar.');"">APL " & rs(0) & "</a>"
		end if

		' Encabezado de la tarjeta (Limpio, sin subrayado)
		html = html & "<div class='bloque-header' style='text-align: center;'>" & ref & "</div>"
		
		html = html & "<div class='bloque-body'>"
			
			' Grid de datos principales (Cliente y Fecha Creaci�n)
			html = html & "<div class='grid-datos-doble' style='margin-bottom: 1rem;'>"
				html = html & "<div><span class='dato-label'>CLIENTE:</span> <span class='dato-value'>" & rs(1) & "</span></div>"
				html = html & "<div><span class='dato-label'>FECHA CREACION:</span> <span class='dato-value'>" & Fechador(rs(7) + nFecini) & "</span></div>"
			html = html & "</div>"

			' Separador sutil
			html = html & "<hr style='border:0; border-top: 1px solid #f1f5f9; margin: 1rem 0;'>"

			' Grid de estad�sticas de acceso (3 columnas)
			html = html & "<div style='display: grid; grid-template-columns: repeat(3, 1fr); gap: 1rem;'>"
				
				' Columna 1: Accesos
				html = html & "<div>"
				html = html & "<div class='dato-label'>CANTIDAD ACCESOS</div><div class='dato-value'>" & rs(2) & "</div>"
				html = html & "</div>"

				' Columna 2: �ltimo Acceso (Fecha/Hora)
				html = html & "<div>"
				html = html & "<div class='dato-label'>ULTIMO ACCESO</div><div class='dato-value'>" & rs(3) & "</div>"
				html = html & "</div>"

				' Columna 3: Tipo
				html = html & "<div>"
				html = html & "<div class='dato-label'>ULTIMO TIPO ACC</div><div class='dato-value'>" & rs(4) & "</div>"
				html = html & "</div>"

			html = html & "</div>"

			' Fila inferior: Responsables
			html = html & "<div class='grid-datos-doble' style='margin-top: 1rem; padding-top: 1rem; border-top: 1px solid #f1f5f9;'>"
				html = html & "<div><span class='dato-label'>RESPONSABLE ENVIO:</span> <span class='dato-value'>" & rs(5) & "</span></div>"
				html = html & "<div><span class='dato-label'>ULTIMO ACCESO (USR):</span> <span class='dato-value'>" & rs(6) & "</span></div>"
			html = html & "</div>"

		html = html & "</div>" ' Fin bloque-body
		html = html & "</div>" ' Fin bloque-datos
		'html = html & "</table><br>"

rs.movenext
loop



else

	' --- ESTADO SIN DATOS NORMALIZADO (APLs) ---
	html = html & "<div class='bloque-datos' style='max-width: 500px; margin: 2rem auto; border: 1px dashed #cbd5e1; background: #f8fafc;'>"
	html = html & "    <div class='bloque-body' style='text-align: center; padding: 3rem;'>"
	html = html & "        <div style='font-size: 3rem; margin-bottom: 1rem;'>??</div>"
	html = html & "<div style='color: #64748b; font-weight: 600; text-transform: uppercase; letter-spacing: 0.05em;'>Sin APLs registradas</div>"
	html = html & "<div style='color: #94a3b8; font-size: 0.9rem; margin-top: 0.5rem;'>No se encontraron movimientos para los criterios seleccionados.</div>"
	html = html & "    </div>"
	html = html & "</div>"

end if

html = html & "</div>"

elseif tipo = "U" then

XXX = SYSLOGON(cnn,"ESTADIST_USR")

'html = html & "<div style='height:75%;width:100%;'><br><br>"
html = html & "<div class='inner' style='padding: 2rem;'>"

id = request.form("uau")

cmd = "sp_cg_Estadisticas @Modo='detalle_U'"

if Trim(id & "") <> "0" And Trim(id & "") <> "" then
	cmd = cmd & ", @Uau=" & SqlStr(id)
end if

Set rs = Exec(cnn, cmd)

if not rs.eof then

' --- INICIO NORMALIZACION: DETALLE ESTADISTICAS (USUARIOS) ---
do while not rs.eof

	'html = html & "<table class='tabl' style='width:55%;' align=center cellspacing=0>"
	html = html & "<div class='bloque-datos' style='margin-bottom: 1.5rem; max-width: 800px; margin-left: auto; margin-right: auto;'>"

	' --- ENCABEZADO: NOMBRE DEL USUARIO ---
	'html = html & "<td class='filc' style='font-size:12pt;' colspan=3><b>" & rs(1) & "</b></td>"
	html = html & "<div class='bloque-header' style='text-align: center; font-weight: 600;'>" & rs(1) & "</div>"
		
	html = html & "<div class='bloque-body'>"
			
		'Grilla de 3 columnas para m�tricas de acceso
		'html = html & "<td class='fil' style='width:17%;'>CANT ACC: " & rs(3) & "</td>"
		'html = html & "<td class='filb' style='width:17%;'>ULT ACC: " & rs(4) & " </td>"
		'html = html & "<td class='filb' style='width:21%;'>ULT TIPO ACC: " & rs(5) & "</td>"
			
		html = html & "<div style='display: grid; grid-template-columns: repeat(3, 1fr); gap: 1rem; text-align: center;'>"
				
			' Cantidad de Accesos
			html = html & "<div>"
			html = html & "<div class='dato-label'>CANT. ACCESOS</div><div class='dato-value'>" & rs(3) & "</div>"
			html = html & "</div>"

			' �ltimo Acceso (Fecha)
			html = html & "<div>"
			html = html & "<div class='dato-label'>ULTIMO ACCESO</div><div class='dato-value'>" & rs(4) & "</div>"
			html = html & "</div>"

			' Tipo de Acceso
			html = html & "<div>"
			html = html & "<div class='dato-label'>TIPO ACCESO</div><div class='dato-value'>" & rs(5) & "</div>"
			html = html & "</div>"

		html = html & "</div>" ' Fin grilla
	html = html & "</div>" ' Fin bloque-body
	html = html & "</div>" ' Fin bloque-datos
	'html = html & "</table><br>"

rs.movenext
loop

else

	' --- ESTADO SIN DATOS NORMALIZADO (USUARIOS) ---
	html = html & "<div class='bloque-datos' style='max-width: 500px; margin: 2rem auto; border: 1px dashed #cbd5e1; background: #f8fafc;'>"
	html = html & "    <div class='bloque-body' style='text-align: center; padding: 3rem;'>"
	html = html & "        <div style='font-size: 3rem; margin-bottom: 1rem;'>??</div>"
	html = html & "<div style='color: #64748b; font-weight: 600; text-transform: uppercase; letter-spacing: 0.05em;'>Sin actividad de usuarios</div>"
	html = html & "<div style='color: #94a3b8; font-size: 0.9rem; margin-top: 0.5rem;'>No se registraron accesos para el usuario seleccionado en este per�odo.</div>"
	html = html & "    </div>"
	html = html & "</div>"

end if

html = html & "</div>"

end if

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

html = html & "    $( ""#fec"" ).datepicker();"

html = html & "  } );"
html = html & "  </script>"

html = html & formx04 & "</body>"
html = html & "</html>"

response.write html

%>