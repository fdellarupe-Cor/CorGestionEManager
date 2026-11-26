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
idapli = request.querystring("id")
idcli = request.querystring("idcli")
opc= request.querystring("opc")
xmenu = request.querystring("xmenu")
xtmenu = request.querystring("xtmenu")

If opc = "modif" Then

	ExecNonQuery Cnn, "sp_cg_ClienAplic @Modo='actualizar_tiene', @IdCliente=" & idcli & ", @IdAplic=" & idapli & ", @Tiene='" & request.form("tiene") & "'"

End If

Set rs = Exec(Cnn, "sp_cg_Utilidades @Modo='consultar_campo', @Tabla='aplicativos', @Id=" & idapli)

tit = ""
If Not rs.eof Then tit = rs(0)

html = html & cabecera()

html = html & "<div class='titulo-seccion ui-page-header' style='padding: 0.5rem 2rem;background-color:#e8edff;'>"
html = html & "<div class='ui-toolbar' style='display:flex;align-items:center;justify-content:space-between;'>"

html = html & "<div><h2>" & Server.HTMLEncode(tit) & "</h2></div>"

formx5 = "<form id='formx5' class='form-oculto' method='post' action='aplicativos.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
botonx5 = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('formx5').submit();"">"
html = html & "<div class='ui-toolbar__actions'>" & botonx5 & "</div>"

html = html & "</div>" ' Fin ui-toolbar
html = html & "</div>" ' Fin titulo-seccion

' --- INICIO NORMALIZACION: DETALLES DE APLICACIÓN (TIENE / NO TIENE) ---

'html = html & "<div style='height:100%;width:100%;float:left'>"
html = html & "<div class='inner ui-main' style='padding: 2rem;'>"
' Centramos la tabla y le damos un ancho máximo para que no se estire excesivamente
html = html & "<div class='bloque-datos table-responsive ui-table-wrap' style='max-width: 600px; margin: 0 auto;'>"

'html = html & "<table style='width:30%;color:#000000;font-family:Serif;text-align:left;' cellspacing=0 align=center>"
html = html & "<table class='tabla-moderna ui-table'>"

html = html & "<thead>"
'html = html & "<tr>"
html = html & "<tr>"
	
'html = html & "<td class='encabe' colspan=3><center>DETALLES</center></td>"
html = html & "<th colspan='2' style='text-align: center; text-transform: uppercase;'>DETALLES</th>"
	
'html = html & "</tr>"
html = html & "</tr>"
html = html & "</thead>"

Set rx = Exec(Cnn, "sp_cg_ClienAplic @Modo='listar_x_aplic', @IdAplic=" & idapli)

html = html & "<tbody>"

Do While Not rx.eof

	'html = html & "<tr>"
	html = html & "<tr>"

	'html = html & "<td class='regis'>" & rx(1) & "</td>"
	html = html & "<td style='vertical-align: middle; color: #475569; font-weight: 500;'>" & Server.HTMLEncode(rx(1) & "") & "</td>"

	' OJO ACÁ: Corrección de estructura HTML. El form entra DENTRO del td.
	'html = html & "<form id='sForm" & rx(0) & "' method='post' action='detalleAplic.asp?opc=modif&id=" & idapli & "&idcli=" & rx(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "' >"
	'html = html & "<td class='regis'>"
	html = html & "<td style='vertical-align: middle; text-align: right;'>"
	html = html & "<form id='sForm" & rx(0) & "' style='margin: 0; padding: 0;' method='post' action='detalleAplic.asp?opc=modif&id=" & idapli & "&idcli=" & rx(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"

	' Envolvemos en dato-value para activar el parche CSS de los combos. Le pongo max-width para que no ocupe toda la celda.
	'html = html & "<select id='tiene' name='tiene' style='height:20px;width:70%;border:solid 1px FF0000;color:#000000;font-size:10pt;' onchange=""actualizar('sForm" & rx(0) & "');"">"
	html = html & "<div class='dato-value' style='display: inline-block; width: 140px; text-align: left;'>"
	html = html & "<select class='ui-select' name='tiene' onchange=""actualizar('sForm" & rx(0) & "');"">"

	If rx(5) Then

		html = html & "<option value='1' selected>TIENE</option>"
		html = html & "<option value='0'>NO TIENE</option>"

	Else

		html = html & "<option value='1'>TIENE</option>"
		html = html & "<option value='0' selected>NO TIENE</option>"

	End If

	'html = html & "</select></td></form></tr>"
	html = html & "</select></div></form></td></tr>"

	rx.movenext
Loop

'html = html & "</table>"
html = html & "</tbody>"
html = html & "</table>"
	
html = html & "</div>" ' Cierre de .bloque-datos
	
'html = html & "</div>"
html = html & "</div>" ' Cierre de .inner

html = html & "<script type='text/javascript'>"
html = html & " function actualizar(formu){ "
html = html & " document.getElementById(formu).submit();"
html = html & " alert('Cambios realizados exitosamente');} "
html = html & "</script>"

html = html & "<script src='assets/js/jquery.min.js'></script>"
html = html & "<script src='assets/js/skel.min.js'></script>"
html = html & "<script src='assets/js/util.js'></script>"
html = html & "<script src='assets/js/main.js'></script>"

html = html & formx5 & "</body>"
html = html & "</html>"

response.write html
%>