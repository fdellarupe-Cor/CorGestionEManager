<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<%


xhtml = headhtml()

xhtml = xhtml & "<body>"

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

xmenu = request.querystring("xmenu")
xtmenu = request.querystring("xtmenu")

XXX = SYSLOGON(cnn,"MENU_LICENCIAS")

xhtml = xhtml & cabecera()

xhtml = xhtml & "<div class='titulo-seccion ui-page-header' style='padding: 1rem 2rem;'>"
xhtml = xhtml & "<div class='ui-toolbar' style='display:flex;align-items:center;justify-content:space-between;'>"

xhtml = xhtml & "<div><h2>LICENCIAS</h2></div>"

xhtml = xhtml & "<div class='ui-toolbar__actions' style='display:flex; gap: 1rem;'>"

xform = "<form id='xform' class='form-oculto' method='post' action='agregar2.asp?tipo=LIC&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
if getPermisos(xmenu,xtmenu,"ALTA") = 1 then
	xboton = "<input class='bot ui-button ui-button--primary' type='button' value='CARGAR NUEVA LICENCIA' onclick=""document.getElementById('xform').submit();"">"
else
	xboton = "<input class='botsp ui-button' type='button' value='CARGAR NUEVA LICENCIA' onclick=""alert('Sin permiso para agregar.');"">"
end if
xhtml = xhtml & xboton

xform2 = "<form id='xform2' class='form-oculto' method='post' action='index.asp'></form>"
xboton2 = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('xform2').submit();"">"
xhtml = xhtml & xboton2

xhtml = xhtml & "</div>" ' Fin ui-toolbar__actions

xhtml = xhtml & "</div>" ' Fin ui-toolbar
xhtml = xhtml & "</div>" ' Fin titulo-seccion

' --- INICIO NORMALIZACION: TABLA DE LICENCIAS (ENCABEZADOS) ---
	
'xhtml = xhtml & "<div><br>"
xhtml = xhtml & "<div class='inner ui-main' style='padding: 2rem;'>"
xhtml = xhtml & "<div class='bloque-datos table-responsive ui-table-wrap'>"

'xhtml = xhtml & "<table style='width:60%;color:#000000;font-family:Serif;text-align:center;' cellspacing=0 align=center>"
xhtml = xhtml & "<table class='tabla-moderna ui-table'>"
	
'Agregamos THEAD por buenas prácticas y compatibilidad de exportación
xhtml = xhtml & "<thead>"
	
'xhtml = xhtml & "<tr>"
xhtml = xhtml & "<tr>"

'xhtml = xhtml & "<td class='encabe' style='width:2%;'>*</td>"
xhtml = xhtml & "<th style='width: 50px; text-align: center;'>&nbsp;</th>" ' Columna del botón eliminar

'xhtml = xhtml & "<td class='encabe' style='width:15%;'>RESPONSABLE</td>"
xhtml = xhtml & "<th>RESPONSABLE</th>"

'xhtml = xhtml & "<td class='encabe' style='width:14%;'>FECHA DESDE</td>"
xhtml = xhtml & "<th style='text-align: center;'>FECHA DESDE</th>"

'xhtml = xhtml & "<td class='encabe' style='width:14%;'>FECHA HASTA</td>"
xhtml = xhtml & "<th style='text-align: center;'>FECHA HASTA</th>"

'xhtml = xhtml & "<td class='encabe' style='width:15%;'>MOTIVO</td>"
xhtml = xhtml & "<th>MOTIVO</th>"

'xhtml = xhtml & "</tr>"
xhtml = xhtml & "</tr>"

	sql = "sp_cg_Licencias_Listar"
	Set rs = Exec(Cnn, sql)

	'rs.Move (CantRegMover)

	'PosReg = 0

	sColor = "#FFFFFF"
	xform3 = ""

	Do While (Not rs.eof) ' And (PosReg < CantRegMostrar)


		xhtml = xhtml & "<tr>"

		xform3 = xform3 & "<form id='formx" & rs(0) & "' style='height:1px;' method='post' action='eliminar.asp?opc=lic&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=licencias.asp'></form>"
		if getPermisos(xmenu,xtmenu,"ELIM") = 1 then
			'xboton3 = "<input class='bot' type='button' value='X' onclick=""borrar('formx" & rs(0) & "');"">"
			xboton3 = "<input class='btn-accion btn-accion-peligro ui-button ui-button--icon ui-button--danger' type='button' value='X' onclick=""borrar('formx" & rs(0) & "');"">"
		else
			'xboton3 = "<input class='botsp' type='button' value='X' onclick=""alert('Sin permiso para eliminar.');"">"
			xboton3 = "<input class='btn-accion btn-accion-deshabilitado ui-button ui-button--icon' type='button' value='X' onclick=""alert('Sin permiso para eliminar.');"">"
		end if

		fIni = ""
		if rs(2) <> 0 then
			fIni = cdate(rs(2)) + CLng(CDate("01-01-1900"))
		end if

		fFin = ""
		if rs(3) <> 0 then
			fFin = cdate(rs(3)) + CLng(CDate("01-01-1900"))
		end if

		' --- RENDERIZADO DE FILA (LICENCIAS) ---
		'xhtml = xhtml & "<td class='fil' >" & xboton3 & "</td>"
		xhtml = xhtml & "<td style='text-align: center;'>" & xboton3 & "</td>"

		'xhtml = xhtml & "<td class='fil' style='background-color:" & sColor & ";' >" & rs(1) & "</td>"
		xhtml = xhtml & "<td>" & Server.HTMLEncode(rs(1) & "") & "</td>"

		'xhtml = xhtml & "<td class='fil' style='background-color:" & sColor & ";'>" & fIni & "</td>"
		xhtml = xhtml & "<td style='text-align: center;'>" & fIni & "</td>"

		'xhtml = xhtml & "<td class='fil' style='background-color:" & sColor & ";'>" & fFin & "</td>"
		xhtml = xhtml & "<td style='text-align: center;'>" & fFin & "</td>"

		'xhtml = xhtml & "<td class='fil' style='background-color:" & sColor & ";'>" & rs(4) & "</td>"
		xhtml = xhtml & "<td>" & Server.HTMLEncode(rs(4) & "") & "</td>"

		'xhtml = xhtml & "</tr>"
		xhtml = xhtml & "</tr>"

		'if sColor = "#FFFFFF" then
			'sColor = "#ECECEC"
		'elseif sColor = "#ECECEC" then
			'sColor = "#FFFFFF"
		'end if

		'PosReg = PosReg + 1

		rs.movenext
	Loop

	xhtml = xhtml & "</table>"

xhtml = xhtml & "</div>"









xhtml = xhtml & "</div>"


xhtml = xhtml & "<script type='text/javascript'>"
xhtml = xhtml & " function borrar(formu){ "
xhtml = xhtml & " var r = confirm('Desea eliminar licencia?'); "
xhtml = xhtml & " if (r){ "
xhtml = xhtml & " document.getElementById(formu).submit(); "
xhtml = xhtml & " }else{ "
xhtml = xhtml & " return false;"
xhtml = xhtml & " }} "
xhtml = xhtml & "</script>"



xhtml = xhtml & "<script src='assets/js/jquery.min.js'></script>"
xhtml = xhtml & "<script src='assets/js/skel.min.js'></script>"
xhtml = xhtml & "<script src='assets/js/util.js'></script>"
xhtml = xhtml & "<script src='assets/js/main.js'></script>"

xhtml = xhtml & xform & xform2 & xform3 & "</body></html>"

response.write xhtml
%>