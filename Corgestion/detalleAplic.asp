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
'[UI-ANT] idapli = request.querystring("id")
'[UI-ANT] idcli = request.querystring("idcli")
'[UI-ANT] opc= request.querystring("opc")
'[UI-ANT] xmenu = request.querystring("xmenu")
'[UI-ANT] xtmenu = request.querystring("xtmenu")

'[UI-ANT] If opc = "modif" Then

	'[UI-ANT] ExecNonQuery Cnn, "sp_cg_ClienAplic @Modo='actualizar_tiene', @IdCliente=" & idcli & ", @IdAplic=" & idapli & ", @Tiene='" & request.form("tiene") & "'"

'[UI-ANT] End If

'[UI-ANT] Set rs = Exec(Cnn, "sp_cg_Utilidades @Modo='consultar_campo', @Tabla='aplicativos', @Id=" & idapli)

'[UI-ANT] tit = ""
'[UI-ANT] If Not rs.eof Then tit = rs(0)

'[UI-ANT] html = html & cabecera()

'[UI-ANT] html = html & "<div class='titulo-seccion ui-page-header' style='padding: 0.5rem 2rem;background-color:#e8edff;'>"
'[UI-ANT] html = html & "<div class='ui-toolbar' style='display:flex;align-items:center;justify-content:space-between;'>"

'[UI-ANT] html = html & "<div><h2>" & Server.HTMLEncode(tit) & "</h2></div>"

'[UI-ANT] formx5 = "<form id='formx5' class='form-oculto' method='post' action='aplicativos.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
'[UI-ANT] botonx5 = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('formx5').submit();"">"
'[UI-ANT] html = html & "<div class='ui-toolbar__actions'>" & botonx5 & "</div>"

'[UI-ANT] html = html & "</div>" ' Fin ui-toolbar
'[UI-ANT] html = html & "</div>" ' Fin titulo-seccion

'[UI-ANT] ' --- INICIO NORMALIZACION: DETALLES DE APLICACION (TIENE / NO TIENE) ---

'[UI-ANT] 'html = html & "<div style='height:100%;width:100%;float:left'>"
'[UI-ANT] html = html & "<div class='inner ui-main' style='padding: 2rem;'>"
'[UI-ANT] ' Centramos la tabla y le damos un ancho maximo para que no se estire excesivamente
'[UI-ANT] html = html & "<div class='bloque-datos table-responsive ui-table-wrap' style='max-width: 600px; margin: 0 auto;'>"

'[UI-ANT] 'html = html & "<table style='width:30%;color:#000000;font-family:Serif;text-align:left;' cellspacing=0 align=center>"
'[UI-ANT] html = html & "<table class='tabla-moderna ui-table'>"

'[UI-ANT] html = html & "<thead>"
'[UI-ANT] 'html = html & "<tr>"
'[UI-ANT] html = html & "<tr>"
	
'[UI-ANT] 'html = html & "<td class='encabe' colspan=3><center>DETALLES</center></td>"
'[UI-ANT] html = html & "<th colspan='2' style='text-align: center; text-transform: uppercase;'>DETALLES</th>"
	
'[UI-ANT] 'html = html & "</tr>"
'[UI-ANT] html = html & "</tr>"
'[UI-ANT] html = html & "</thead>"

'[UI-ANT] Set rx = Exec(Cnn, "sp_cg_ClienAplic @Modo='listar_x_aplic', @IdAplic=" & idapli)

'[UI-ANT] html = html & "<tbody>"

'[UI-ANT] Do While Not rx.eof

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] html = html & "<tr>"

	'[UI-ANT] 'html = html & "<td class='regis'>" & rx(1) & "</td>"
	'[UI-ANT] html = html & "<td style='vertical-align: middle; color: #475569; font-weight: 500;'>" & Server.HTMLEncode(rx(1) & "") & "</td>"

	'[UI-ANT] ' OJO ACA: Correccion de estructura HTML. El form entra DENTRO del td.
	'[UI-ANT] 'html = html & "<form id='sForm" & rx(0) & "' method='post' action='detalleAplic.asp?opc=modif&id=" & idapli & "&idcli=" & rx(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "' >"
	'[UI-ANT] 'html = html & "<td class='regis'>"
	'[UI-ANT] html = html & "<td style='vertical-align: middle; text-align: right;'>"
	'[UI-ANT] html = html & "<form id='sForm" & rx(0) & "' style='margin: 0; padding: 0;' method='post' action='detalleAplic.asp?opc=modif&id=" & idapli & "&idcli=" & rx(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"

	'[UI-ANT] ' Envolvemos en dato-value para activar el parche CSS de los combos. Le pongo max-width para que no ocupe toda la celda.
	'[UI-ANT] 'html = html & "<select id='tiene' name='tiene' style='height:20px;width:70%;border:solid 1px FF0000;color:#000000;font-size:10pt;' onchange=""actualizar('sForm" & rx(0) & "');"">"
	'[UI-ANT] html = html & "<div class='dato-value' style='display: inline-block; width: 140px; text-align: left;'>"
	'[UI-ANT] html = html & "<select class='ui-select' name='tiene' onchange=""actualizar('sForm" & rx(0) & "');"">"

	'[UI-ANT] If rx(5) Then

		'[UI-ANT] html = html & "<option value='1' selected>TIENE</option>"
		'[UI-ANT] html = html & "<option value='0'>NO TIENE</option>"

	'[UI-ANT] Else

		'[UI-ANT] html = html & "<option value='1'>TIENE</option>"
		'[UI-ANT] html = html & "<option value='0' selected>NO TIENE</option>"

	'[UI-ANT] End If

	'[UI-ANT] 'html = html & "</select></td></form></tr>"
	'[UI-ANT] html = html & "</select></div></form></td></tr>"

	'[UI-ANT] rx.movenext
'[UI-ANT] Loop

'[UI-ANT] 'html = html & "</table>"
'[UI-ANT] html = html & "</tbody>"
'[UI-ANT] html = html & "</table>"
	
'[UI-ANT] html = html & "</div>" ' Cierre de .bloque-datos
	
'[UI-ANT] 'html = html & "</div>"
'[UI-ANT] html = html & "</div>" ' Cierre de .inner

'[UI-ANT] html = html & "<script type='text/javascript'>"
'[UI-ANT] html = html & " function actualizar(formu){ "
'[UI-ANT] html = html & " document.getElementById(formu).submit();"
'[UI-ANT] html = html & " alert('Cambios realizados exitosamente');} "
'[UI-ANT] html = html & "</script>"

'[UI-ANT] html = html & "<script src='assets/js/jquery.min.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/skel.min.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/util.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/main.js'></script>"

'[UI-ANT] html = html & formx5 & "</body>"
'[UI-ANT] html = html & "</html>"

'[UI-ANT] response.write html

' =========================================================================
' [UI] PANTALLA con el sistema visual (IncFile/ui_componentes.inc)
' Clientes que tienen un aplicativo (se llega desde Aplicativos > Ver / modificar).
' El cuerpo anterior completo queda arriba, comentado con '[UI-ANT].
' Consultas: las mismas de antes. Cambios: ids y valor validados, permiso de
' modificar chequeado en el servidor y confirmacion despues de guardar.
' =========================================================================
Dim puedeModif, aviso, tiene, guardado, nombreGuardado, accTop, filas, cantCli, cantTienen, selTiene, nombreCli

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
idapli = request.querystring("id")
idcli = request.querystring("idcli")
opc= request.querystring("opc")
xmenu = request.querystring("xmenu")
xtmenu = request.querystring("xtmenu")

' [UI] Los ids van a la consulta SQL: se validan como numeros (antes se pegaban tal cual)
If IsNumeric(idapli & "") Then idapli = CLng(idapli) Else idapli = 0
If IsNumeric(idcli & "") Then idcli = CLng(idcli) Else idcli = 0
puedeModif = (getPermisosAtrm("CG_APLICATIVOS", "MODIF") = 1)
' [LOG] Consulta del detalle de un aplicativo (CG_LOG)
If idapli > 0 Then CgLog "APLICATIVO", idapli, "CONSULTA"

aviso = ""
guardado = False
If opc = "modif" Then
	tiene = request.form("tiene") & ""
	If Not puedeModif Then
		' [UI] El permiso se chequea tambien aca: antes solo dependia del enlace de Aplicativos
		aviso = "<div class='ui-alert ui-alert--danger' role='alert'>No ten&eacute;s permiso para modificar.</div>"
	ElseIf (tiene = "1" Or tiene = "0") And idapli > 0 And idcli > 0 Then
		ExecNonQuery Cnn, "sp_cg_ClienAplic @Modo='actualizar_tiene', @IdCliente=" & idcli & ", @IdAplic=" & idapli & ", @Tiene='" & tiene & "'"
		CgLog "APLICATIVO", idapli, "ACTUALIZA"
		guardado = True
	End If
End If

Set rs = Exec(Cnn, "sp_cg_Utilidades @Modo='consultar_campo', @Tabla='aplicativos', @Id=" & idapli)
tit = ""
If Not rs.eof Then tit = rs(0)

html = html & "<div id='wrapper' role='banner'>" & cabecera() & "</div>"

formx5 = "<form id='formx5' class='form-oculto' method='post' action='aplicativos.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
accTop = UiBoton("Volver", "secondary", "document.getElementById('formx5').submit();", True, "")

' --- Filas: cada cliente con su "lo tiene / no lo tiene" (se guarda al cambiar, como antes) ---
Set rx = Exec(Cnn, "sp_cg_ClienAplic @Modo='listar_x_aplic', @IdAplic=" & idapli)
filas = ""
cantCli = 0
cantTienen = 0
nombreGuardado = ""
Do While Not rx.eof
	nombreCli = UiTexto(rx(1))
	If rx(5) Then
		selTiene = "<option value='1' selected>S&iacute;</option><option value='0'>No</option>"
		cantTienen = cantTienen + 1
	Else
		selTiene = "<option value='1'>S&iacute;</option><option value='0' selected>No</option>"
	End If
	If guardado And CStr(rx(0)) = CStr(idcli) Then nombreGuardado = nombreCli

	filas = filas & "<tr><td><span class='ui-table__titulo'>" & nombreCli & "</span></td><td>"
	filas = filas & "<form id='sForm" & rx(0) & "' class='form-linea' method='post' action='detalleAplic.asp?opc=modif&id=" & idapli & "&idcli=" & rx(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
	If puedeModif Then
		filas = filas & "<select class='ui-select ui-select--compacta' name='tiene' aria-label='&iquest;" & nombreCli & " tiene el aplicativo?' onchange=""document.getElementById('sForm" & rx(0) & "').submit();"">" & selTiene & "</select>"
	Else
		filas = filas & "<select class='ui-select ui-select--compacta' name='tiene' aria-label='&iquest;" & nombreCli & " tiene el aplicativo?' disabled title='Sin permiso para modificar'>" & selTiene & "</select>"
	End If
	filas = filas & "</form></td></tr>"
	cantCli = cantCli + 1
	rx.movenext
Loop

' Confirmacion despues de guardar (antes el aviso salia antes de enviar, aunque fallara)
If guardado Then
	aviso = "<div class='ui-alert ui-alert--success' role='status'>Cambio guardado"
	If nombreGuardado <> "" Then aviso = aviso & ": " & nombreGuardado
	aviso = aviso & ".</div>"
End If

If tit = "" Then
	html = html & UiPantallaParametrosAbre("Aplicativo", accTop, "narrow")
Else
	html = html & UiPantallaParametrosAbre(UiTexto(tit), accTop, "narrow")
End If
html = html & aviso
html = html & UiCardAbre("Clientes que tienen el aplicativo", "")
html = html & UiTablaAbre("Clientes y si tienen el aplicativo")
html = html & "<thead><tr>" & UiTh("Cliente", "") & UiTh("&iquest;Lo tiene?", "ui-table__col--select") & "</tr></thead><tbody>"
If cantCli = 0 Then
	html = html & UiFilaVacia(2, "No hay clientes para mostrar.")
Else
	html = html & filas
End If
html = html & UiTablaCierra()
If cantCli > 0 Then
	If cantTienen = 1 Then
		html = html & UiPie("1 de " & cantCli & " clientes lo tiene", "")
	Else
		html = html & UiPie(cantTienen & " de " & cantCli & " clientes lo tienen", "")
	End If
End If
html = html & UiCardCierra()
html = html & UiPantallaCierra()

html = html & "<script src='assets/js/jquery.min.js'></script>"
html = html & "<script src='assets/js/skel.min.js'></script>"
html = html & "<script src='assets/js/util.js'></script>"
html = html & "<script src='assets/js/main.js'></script>"

html = html & formx5 & "</body>"
html = html & "</html>"

response.write html
%>