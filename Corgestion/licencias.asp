<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/ui_componentes.inc" -->
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

'[UI-ANT] ' ===== Vista anterior completa: reemplazada por el bloque [UI] que sigue. Se borra entera buscando [UI-ANT] =====
'[UI-ANT] xhtml = xhtml & cabecera()

'[UI-ANT] xhtml = xhtml & "<div class='titulo-seccion ui-page-header' style='padding: 1rem 2rem;'>"
'[UI-ANT] xhtml = xhtml & "<div class='ui-toolbar' style='display:flex;align-items:center;justify-content:space-between;'>"

'[UI-ANT] xhtml = xhtml & "<div><h2>LICENCIAS</h2></div>"

'[UI-ANT] xhtml = xhtml & "<div class='ui-toolbar__actions' style='display:flex; gap: 1rem;'>"

'[UI-ANT] xform = "<form id='xform' class='form-oculto' method='post' action='agregar2.asp?tipo=LIC&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
'[UI-ANT] if getPermisos(xmenu,xtmenu,"ALTA") = 1 then
	'[UI-ANT] xboton = "<input class='bot ui-button ui-button--primary' type='button' value='CARGAR NUEVA LICENCIA' onclick=""document.getElementById('xform').submit();"">"
'[UI-ANT] else
	'[UI-ANT] xboton = "<input class='botsp ui-button' type='button' value='CARGAR NUEVA LICENCIA' onclick=""alert('Sin permiso para agregar.');"">"
'[UI-ANT] end if
'[UI-ANT] xhtml = xhtml & xboton

'[UI-ANT] xform2 = "<form id='xform2' class='form-oculto' method='post' action='index.asp'></form>"
'[UI-ANT] xboton2 = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('xform2').submit();"">"
'[UI-ANT] xhtml = xhtml & xboton2

'[UI-ANT] xhtml = xhtml & "</div>" ' Fin ui-toolbar__actions

'[UI-ANT] xhtml = xhtml & "</div>" ' Fin ui-toolbar
'[UI-ANT] xhtml = xhtml & "</div>" ' Fin titulo-seccion

'[UI-ANT] ' --- INICIO NORMALIZACION: TABLA DE LICENCIAS (ENCABEZADOS) ---
	
'[UI-ANT] 'xhtml = xhtml & "<div><br>"
'[UI-ANT] xhtml = xhtml & "<div class='inner ui-main' style='padding: 2rem;'>"
'[UI-ANT] xhtml = xhtml & "<div class='bloque-datos table-responsive ui-table-wrap'>"

'[UI-ANT] 'xhtml = xhtml & "<table style='width:60%;color:#000000;font-family:Serif;text-align:center;' cellspacing=0 align=center>"
'[UI-ANT] xhtml = xhtml & "<table class='tabla-moderna ui-table'>"
	
'[UI-ANT] 'Agregamos THEAD por buenas practicas y compatibilidad de exportacion
'[UI-ANT] xhtml = xhtml & "<thead>"
	
'[UI-ANT] 'xhtml = xhtml & "<tr>"
'[UI-ANT] xhtml = xhtml & "<tr>"

'[UI-ANT] 'xhtml = xhtml & "<td class='encabe' style='width:2%;'>*</td>"
'[UI-ANT] xhtml = xhtml & "<th style='width: 50px; text-align: center;'>&nbsp;</th>" ' Columna del boton eliminar

'[UI-ANT] 'xhtml = xhtml & "<td class='encabe' style='width:15%;'>RESPONSABLE</td>"
'[UI-ANT] xhtml = xhtml & "<th>RESPONSABLE</th>"

'[UI-ANT] 'xhtml = xhtml & "<td class='encabe' style='width:14%;'>FECHA DESDE</td>"
'[UI-ANT] xhtml = xhtml & "<th style='text-align: center;'>FECHA DESDE</th>"

'[UI-ANT] 'xhtml = xhtml & "<td class='encabe' style='width:14%;'>FECHA HASTA</td>"
'[UI-ANT] xhtml = xhtml & "<th style='text-align: center;'>FECHA HASTA</th>"

'[UI-ANT] 'xhtml = xhtml & "<td class='encabe' style='width:15%;'>MOTIVO</td>"
'[UI-ANT] xhtml = xhtml & "<th>MOTIVO</th>"

'[UI-ANT] 'xhtml = xhtml & "</tr>"
'[UI-ANT] xhtml = xhtml & "</tr>"

	'[UI-ANT] sql = "sp_cg_Licencias_Listar"
	'[UI-ANT] Set rs = Exec(Cnn, sql)

	'[UI-ANT] 'rs.Move (CantRegMover)

	'[UI-ANT] 'PosReg = 0

	'[UI-ANT] sColor = "#FFFFFF"
	'[UI-ANT] xform3 = ""

	'[UI-ANT] Do While (Not rs.eof) ' And (PosReg < CantRegMostrar)


		'[UI-ANT] xhtml = xhtml & "<tr>"

		'[UI-ANT] xform3 = xform3 & "<form id='formx" & rs(0) & "' style='height:1px;' method='post' action='eliminar.asp?opc=lic&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=licencias.asp'></form>"
		'[UI-ANT] if getPermisos(xmenu,xtmenu,"ELIM") = 1 then
			'[UI-ANT] 'xboton3 = "<input class='bot' type='button' value='X' onclick=""borrar('formx" & rs(0) & "');"">"
			'[UI-ANT] xboton3 = "<input class='btn-accion btn-accion-peligro ui-button ui-button--icon ui-button--danger' type='button' value='X' onclick=""borrar('formx" & rs(0) & "');"">"
		'[UI-ANT] else
			'[UI-ANT] 'xboton3 = "<input class='botsp' type='button' value='X' onclick=""alert('Sin permiso para eliminar.');"">"
			'[UI-ANT] xboton3 = "<input class='btn-accion btn-accion-deshabilitado ui-button ui-button--icon' type='button' value='X' onclick=""alert('Sin permiso para eliminar.');"">"
		'[UI-ANT] end if

		'[UI-ANT] fIni = ""
		'[UI-ANT] if rs(2) <> 0 then
			'[UI-ANT] fIni = cdate(rs(2)) + CLng(CDate("01-01-1900"))
		'[UI-ANT] end if

		'[UI-ANT] fFin = ""
		'[UI-ANT] if rs(3) <> 0 then
			'[UI-ANT] fFin = cdate(rs(3)) + CLng(CDate("01-01-1900"))
		'[UI-ANT] end if

		'[UI-ANT] ' --- RENDERIZADO DE FILA (LICENCIAS) ---
		'[UI-ANT] 'xhtml = xhtml & "<td class='fil' >" & xboton3 & "</td>"
		'[UI-ANT] xhtml = xhtml & "<td style='text-align: center;'>" & xboton3 & "</td>"

		'[UI-ANT] 'xhtml = xhtml & "<td class='fil' style='background-color:" & sColor & ";' >" & rs(1) & "</td>"
		'[UI-ANT] xhtml = xhtml & "<td>" & Server.HTMLEncode(rs(1) & "") & "</td>"

		'[UI-ANT] 'xhtml = xhtml & "<td class='fil' style='background-color:" & sColor & ";'>" & fIni & "</td>"
		'[UI-ANT] xhtml = xhtml & "<td style='text-align: center;'>" & fIni & "</td>"

		'[UI-ANT] 'xhtml = xhtml & "<td class='fil' style='background-color:" & sColor & ";'>" & fFin & "</td>"
		'[UI-ANT] xhtml = xhtml & "<td style='text-align: center;'>" & fFin & "</td>"

		'[UI-ANT] 'xhtml = xhtml & "<td class='fil' style='background-color:" & sColor & ";'>" & rs(4) & "</td>"
		'[UI-ANT] xhtml = xhtml & "<td>" & Server.HTMLEncode(rs(4) & "") & "</td>"

		'[UI-ANT] 'xhtml = xhtml & "</tr>"
		'[UI-ANT] xhtml = xhtml & "</tr>"

		'[UI-ANT] 'if sColor = "#FFFFFF" then
			'[UI-ANT] 'sColor = "#ECECEC"
		'[UI-ANT] 'elseif sColor = "#ECECEC" then
			'[UI-ANT] 'sColor = "#FFFFFF"
		'[UI-ANT] 'end if

		'[UI-ANT] 'PosReg = PosReg + 1

		'[UI-ANT] rs.movenext
	'[UI-ANT] Loop

	'[UI-ANT] xhtml = xhtml & "</table>"

'[UI-ANT] xhtml = xhtml & "</div>"









'[UI-ANT] xhtml = xhtml & "</div>"


'[UI-ANT] xhtml = xhtml & "<script type='text/javascript'>"
'[UI-ANT] xhtml = xhtml & " function borrar(formu){ "
'[UI-ANT] xhtml = xhtml & " var r = confirm('Desea eliminar licencia?'); "
'[UI-ANT] xhtml = xhtml & " if (r){ "
'[UI-ANT] xhtml = xhtml & " document.getElementById(formu).submit(); "
'[UI-ANT] xhtml = xhtml & " }else{ "
'[UI-ANT] xhtml = xhtml & " return false;"
'[UI-ANT] xhtml = xhtml & " }} "
'[UI-ANT] xhtml = xhtml & "</script>"

' =========================================================================
' [UI] VISTA con el sistema visual (IncFile/ui_componentes.inc)
' El bloque anterior completo queda arriba, comentado con '[UI-ANT].
' Consulta, formularios y permisos: los mismos de antes.
' =========================================================================
Dim accTop, puedeEliminar, hoySerial, cantLic, desdeLic, hastaLic, estadoLic

xhtml = xhtml & cabecera()

' --- Barra superior: Cargar y Volver ---
xform = "<form id='xform' class='form-oculto' method='post' action='agregar2.asp?tipo=LIC&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
accTop = UiBoton("Cargar licencia", "light", "document.getElementById('xform').submit();", getPermisosAtrm("CG_LICENCIAS","ALTA") = 1, "Sin permiso para agregar.")
xform2 = "<form id='xform2' class='form-oculto' method='post' action='index.asp'></form>"
accTop = accTop & UiBoton("Volver", "secondary", "document.getElementById('xform2').submit();", True, "")

xhtml = xhtml & UiPantallaAbre("Licencias", accTop, "")
xhtml = xhtml & UiCardAbre("Listado de licencias", "")

xhtml = xhtml & UiTablaAbre("Licencias registradas")
xhtml = xhtml & "<thead><tr>" & _
	UiTh("<span class='ui-sr-only'>Acciones</span>", "ui-table__col--action") & _
	UiTh("Responsable", "") & _
	UiTh("Desde", "") & _
	UiTh("Hasta", "") & _
	UiTh("Estado", "") & _
	UiTh("Motivo", "") & _
	"</tr></thead><tbody>"

sql = "sp_cg_Licencias_Listar"
Set rs = Exec(Cnn, sql)

' Hoy en el formato de la base (dias desde 01-01-1900), para mostrar el estado de cada licencia
hoySerial = CLng(Date) - CLng(CDate("01-01-1900"))
puedeEliminar = (getPermisosAtrm("CG_LICENCIAS","ELIM") = 1)
xform3 = ""
cantLic = 0

Do While (Not rs.eof)

	xform3 = xform3 & "<form id='formx" & rs(0) & "' class='form-oculto' method='post' action='eliminar.asp?opc=lic&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=licencias.asp'></form>"

	' Estado segun las fechas (solo visual): Proxima, En curso o Finalizada
	desdeLic = CLng("0" & rs(2))
	hastaLic = CLng("0" & rs(3))
	If desdeLic > 0 And hoySerial < desdeLic Then
		estadoLic = UiBadge("Pr&oacute;xima", "blue")
	ElseIf hastaLic > 0 And hoySerial > hastaLic Then
		estadoLic = UiBadge("Finalizada", "gray")
	ElseIf desdeLic > 0 Then
		estadoLic = UiBadge("En curso", "amber")
	Else
		estadoLic = "<span class='ui-status-empty'>&mdash;</span>"
	End If

	xhtml = xhtml & "<tr>" & _
		"<td>" & UiBotonIcono("&times;", "Eliminar licencia " & rs(0), "danger", "borrar('formx" & rs(0) & "');", puedeEliminar, "Sin permiso para eliminar") & "</td>" & _
		"<td><strong>" & UiTexto(rs(1)) & "</strong></td>" & _
		"<td><span class='ui-table__nowrap'>" & UiFecha(rs(2)) & "</span></td>" & _
		"<td><span class='ui-table__nowrap'>" & UiFecha(rs(3)) & "</span></td>" & _
		"<td>" & estadoLic & "</td>" & _
		"<td>" & UiDato(rs(4)) & "</td>" & _
		"</tr>"

	cantLic = cantLic + 1
	rs.movenext
Loop

If cantLic = 0 Then xhtml = xhtml & UiFilaVacia(6, "No hay licencias registradas.")
xhtml = xhtml & UiTablaCierra()
' Sin licencias, la fila vacia ya lo dice: el pie solo aparece cuando hay registros
If cantLic = 1 Then
	xhtml = xhtml & UiPie("1 licencia", "")
ElseIf cantLic > 1 Then
	xhtml = xhtml & UiPie(cantLic & " licencias", "")
End If

xhtml = xhtml & UiCardCierra()
xhtml = xhtml & UiPantallaCierra()

xhtml = xhtml & "<script type='text/javascript'>"
xhtml = xhtml & " function borrar(formu){ "
xhtml = xhtml & " var r = confirm('\u00bfDesea eliminar la licencia?'); "
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