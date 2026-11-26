<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<%
html = headhtml()
html = html & "<body>"

id = request.querystring("id")
apli = request.querystring("apli")
opc = request.querystring("opc")
sel = request.form("sel")
xmenu = request.querystring("xmenu")
xtmenu = request.querystring("xtmenu")

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

On Error Resume Next
XXX = SYSLOGON(cnn, "PARAM_CLIEN")
If Err.Number <> 0 Then
	Response.Write "ERROR SYSLOGON: " & Err.Number & ": " & Err.Description
	Response.End
End If
On Error Goto 0

If apli & "" = "" Then
	If sel & "" <> "" Then
		apli = sel
	Else
		On Error Resume Next
		Set rsPri = Exec(Cnn, "sp_cg_ClienAplic @Modo='primer_id', @IdCliente=" & id)
		If Err.Number <> 0 Then
			Response.Write "ERROR primer_id: " & Err.Number & ": " & Err.Description
			Response.End
		End If
		On Error Goto 0

		If Not rsPri.eof Then
			apli = rsPri(0)
		Else
			apli = 0
		End If
		rsPri.Close
	End If
End If

If opc = "UPD" Then

	ambiente = Trim(request.form("AMBIENTE") & "")
	base_em  = Trim(request.form("BASE_EM") & "")
	base_rqt = Trim(request.form("BASE_RQT") & "")
	ruta_sv  = Trim(request.form("RUTA_SV") & "")
	tipo     = Trim(request.querystring("tipo") & "")

	If ambiente = "" Or base_em = "" Or base_rqt = "" Or ruta_sv = "" Then
		html = html & "<script>alert('Todos los campos son obligatorios.'); window.history.back();</script>"
	'ElseIf tipo = "" Then
	'	html = html & "<script>alert('Error: tipo de ambiente no valido.'); window.history.back();</script>"
	Else

		On Error Resume Next
		Set chk = Exec(Cnn, "sp_cg_ClienAplic @Modo='obtener', @IdCliente=" & id & ", @IdAplic=" & apli)
		If Err.Number <> 0 Then
			Response.Write "ERROR obtener: " & Err.Number & ": " & Err.Description
			Response.End
		End If
		On Error Goto 0

		encontro = False
		Do While Not chk.eof
			If UCase(Trim(chk("TIPO_AMB") & "")) = UCase(tipo) Then
				encontro = True
				Exit Do
			End If
			chk.movenext
		Loop

		If encontro Then

			sSQL = "sp_cg_ClienAplic @Modo='actualizar'"
			sSQL = sSQL & ", @IdCliente=" & id
			sSQL = sSQL & ", @IdAplic=" & apli
			sSQL = sSQL & ", @BaseEm='" & Replace(base_em, "'", "''") & "'"
			sSQL = sSQL & ", @BaseRqt='" & Replace(base_rqt, "'", "''") & "'"
			sSQL = sSQL & ", @RutaSv='" & Replace(ruta_sv, "'", "''") & "'"
			sSQL = sSQL & ", @Ambiente='" & Replace(ambiente, "'", "''") & "'"
			sSQL = sSQL & ", @TipoAmb='" & Replace(tipo, "'", "''") & "'"

			On Error Resume Next
			ExecNonQuery Cnn, sSQL
			If Err.Number <> 0 Then
				Response.Write "ERROR actualizar: " & Err.Number & ": " & Err.Description
				Response.End
			End If
			On Error Goto 0

			html = html & "<script>alert('Actualizado exitosamente.'); window.location.href='detalleCliente.asp?id=" & id & "&apli=" & apli & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "';</script>"
		Else

			sSQL = "sp_cg_ClienAplic @Modo='insertar'"
			sSQL = sSQL & ", @IdCliente=" & id
			sSQL = sSQL & ", @IdAplic=" & apli
			sSQL = sSQL & ", @Tiene='1'"
			sSQL = sSQL & ", @BaseEm='" & Replace(base_em, "'", "''") & "'"
			sSQL = sSQL & ", @BaseRqt='" & Replace(base_rqt, "'", "''") & "'"
			sSQL = sSQL & ", @RutaSv='" & Replace(ruta_sv, "'", "''") & "'"
			sSQL = sSQL & ", @Ambiente='" & Replace(ambiente, "'", "''") & "'"
			sSQL = sSQL & ", @TipoAmb='" & Replace(tipo, "'", "''") & "'"			

			On Error Resume Next
			ExecNonQuery Cnn, sSQL
			If Err.Number <> 0 Then
				Response.Write "ERROR insertar: " & Err.Number & ": " & Err.Description
				Response.End
			End If
			On Error Goto 0

			If ambiente <> "" Or tipo <> "" Then
				sSQL2 = "sp_cg_ClienAplic @Modo='actualizar'"
				sSQL2 = sSQL2 & ", @IdCliente=" & id
				sSQL2 = sSQL2 & ", @IdAplic=" & apli
				If ambiente <> "" Then sSQL2 = sSQL2 & ", @Ambiente='" & Replace(ambiente, "'", "''") & "'"
				If tipo <> "" Then sSQL2 = sSQL2 & ", @TipoAmb='" & Replace(tipo, "'", "''") & "'"

				On Error Resume Next
				ExecNonQuery Cnn, sSQL2
				If Err.Number <> 0 Then
					Response.Write "ERROR insertar ambiente: " & Err.Number & ": " & Err.Description
					Response.End
				End If
				On Error Goto 0
			End If

			html = html & "<script>alert('Registrado exitosamente.'); window.location.href='detalleCliente.asp?id=" & id & "&apli=" & apli & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "';</script>"
		End If

		chk.Close
	End If

ElseIf opc = "DEL" Then

	delId = Trim(request.querystring("delId") & "")
	If delId = "" Then delId = Trim(request.form("delId") & "")
	If delId <> "" And IsNumeric(delId) Then
		usuarioDel = session("usuario") & ""
		If usuarioDel = "" Then usuarioDel = UserLogon() & ""
		sSQL = "sp_cg_Registro_Eliminar @Opt='clxapli', @Id=" & CLng(delId) & ", @Usuario='" & Replace(usuarioDel, "'", "''") & "'"
		On Error Resume Next
		Cnn.Execute(sSQL)
		If Err.Number <> 0 Then
			Response.Write "ERROR eliminar ambiente: " & Err.Number & ": " & Err.Description
			Response.End
		End If
		On Error Goto 0
		html = html & "<script>alert('Ambiente eliminado.'); window.location.href='detalleCliente.asp?id=" & id & "&apli=" & apli & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "';</script>"
	Else
		html = html & "<script>alert('ID de ambiente no valido.'); window.history.back();</script>"
	End If

End If

On Error Resume Next
Set rs = Exec(Cnn, "sp_cg_Utilidades @Modo='consultar_campo', @Tabla='clientes', @Id=" & id)
If Err.Number <> 0 Then
	Response.Write "ERROR consultar_campo: " & Err.Number & ": " & Err.Description
	Response.End
End If
On Error Goto 0

tit = ""
If Not rs.eof Then tit = rs(0)
rs.Close

html = html & cabecera()

html = html & "<div class='titulo-seccion ui-page-header' style='padding: 0.5rem 2rem;background-color:#e8edff;'>"
html = html & "<div class='ui-toolbar' style='display:flex;align-items:center;justify-content:space-between;'>"

html = html & "<div><h2>" & Server.HTMLEncode(tit) & "</h2></div>"

html = html & "<div style='display:flex;align-items:center;gap:0.5rem;'>"
html = html & "<strong>Aplicativo:</strong>"

html = html & "<form id='sForm' method='post' action='detalleCliente.asp?id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"

html = html & "<select class='selec ui-select' id='sel' name='sel' style='height:50%;width:70%;' onchange=""document.getElementById('sForm').submit();"">"

On Error Resume Next
Set rx = Exec(Cnn, "sp_cg_ClienAplic @Modo='listar_x_cliente', @IdCliente=" & id)
If Err.Number <> 0 Then
	Response.Write "ERROR listar_x_cliente: " & Err.Number & ": " & Err.Description
	Response.End
End If
On Error Goto 0

Do While Not rx.eof
	If CStr(apli & "") = CStr(rx(0) & "") Then
		html = html & "<option value='" & rx(0) & "' selected>" & Server.HTMLEncode(rx(1) & "") & "</option>"
	Else
		html = html & "<option value='" & rx(0) & "'>" & Server.HTMLEncode(rx(1) & "") & "</option>"
	End If
	rx.movenext
Loop

html = html & "</select>"
html = html & "</form>"
html = html & "</div>" ' Fin selector Aplicativo

formx5 = "<form id='formx5' class='form-oculto' method='post' action='clientes.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
botonx5 = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('formx5').submit();"">"
html = html & "<div class='ui-toolbar__actions'>" & botonx5 & "</div>"

html = html & "</div>" ' Fin ui-toolbar
html = html & "</div>" ' Fin titulo-seccion

html = html & "<div class='inner ui-main' style='padding: 2rem; display: flex; flex-direction: column; gap: 2rem;'>"

On Error Resume Next
Set rz = Exec(Cnn, "sp_cg_ClienAplic @Modo='obtener', @IdCliente=" & id & ", @IdAplic=" & apli)
If Err.Number <> 0 Then
	Response.Write "ERROR obtener datos: " & Err.Number & ": " & Err.Description
	Response.End
End If
On Error Goto 0

numReg = 0

If Not rz.eof Then

	Do While Not rz.eof

		numReg = numReg + 1
		vId         = rz("ID") & ""
		vAmbiente   = rz("AMBIENTE") & ""
		vBaseEm     = rz("BASE_EM") & ""
		vBaseRqt    = rz("BASE_RQT") & ""
		vRutaSv     = rz("RUTA_SV") & ""
		vTipo       = rz("TIPO_AMB") & ""

		If vTipo = "T" Then
			stit = "TESTING"
		ElseIf vTipo = "P" Then
			stit = "PRODUCCION"
		Else
			stit = "AMBIENTE UNICO"
		End If

		frmId = "frmCliente_" & numReg

		html = html & "<form id='" & frmId & "' method='post' action='detalleCliente.asp?opc=UPD&id=" & id & "&apli=" & apli & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&tipo=" & Server.HTMLEncode(vTipo) & "'>"

		html = html & "<div class='bloque-datos ui-panel' style='max-width: 900px; margin: 0 auto;'>"
		html = html & "<div class='bloque-header ui-panel__header'>" & stit & "</div>"
		html = html & "<div class='bloque-body ui-panel__body'>"
		html = html & "<div class='grid-datos-doble ui-field-grid--dual'>"

		html = html & "<div class='dato-label'>AMBIENTE (Url)</div>"
		html = html & "<div>" & HtmlInputCampo("AMBIENTE", vAmbiente, False) & "</div>"

		html = html & "<div class='dato-label'>BASE DATOS EM</div>"
		html = html & "<div>" & HtmlInputCampo("BASE_EM", vBaseEm, False) & "</div>"

		html = html & "<div class='dato-label'>BASE DATOS RQT</div>"
		html = html & "<div>" & HtmlInputCampo("BASE_RQT", vBaseRqt, False) & "</div>"

		html = html & "<div class='dato-label'>RUTA SV-40 (Webintranet)</div>"
		html = html & "<div>" & HtmlInputCampo("RUTA_SV", vRutaSv, False) & "</div>"

		html = html & "</div>"
		html = html & "</div>"

		html = html & "<div style='display: flex; justify-content: flex-end; gap: 1rem; padding: 1rem 0;'>"

		If getPermisos(xmenu, xtmenu, "MODIF") = 1 Then
			html = html & "<input class='btn-editar ui-button' type='button' value='&#9998; Editar' title='Editar ambiente' style='background-color:#6b7280; color:white; padding:0.5rem 1rem; border:none; border-radius:4px; font-size:0.95rem; cursor:pointer;' onclick=""toggleEditarFila('" & frmId & "', this);"">"
		End If
		If getPermisos(xmenu, xtmenu, "ELIM") = 1 Then
			html = html & "<input class='btn-eliminar ui-button ui-button--danger' type='button' value='&#128465; Eliminar' title='Eliminar ambiente' style='background-color:#dc2626; color:white; padding:0.5rem 1rem; border:none; border-radius:4px; font-size:0.95rem; cursor:pointer;' onclick=""if(confirm('¿Eliminar este ambiente? Esta acción no se puede deshacer.')){window.location.href='detalleCliente.asp?opc=DEL&delId=" & vId & "&id=" & id & "&apli=" & apli & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "';}"">"
		End If

		html = html & "</div>"

		html = html & "</div>"
		html = html & "</form>"

		rz.movenext
	Loop

	' --- Boton + Form para AGREGAR nuevo ambiente (visible siempre que hay datos) ---
	If getPermisos(xmenu, xtmenu, "MODIF") = 1 Then
		html = html & "<div style='max-width: 900px; margin: 0 auto; display:flex; justify-content:flex-end; padding:0.5rem 0;'>"
		html = html & "<input class='bot ui-button ui-button--primary' type='button' value='+ AGREGAR AMBIENTE' style='background-color:#10b981; color:white; padding:0.75rem 1.5rem; border:none; border-radius:4px; font-size:1rem; cursor:pointer;' onclick=""document.getElementById('frmAgregar').style.display=(document.getElementById('frmAgregar').style.display=='none'?'block':'none');"">"
		html = html & "</div>"

		html = html & "<form id='frmAgregar' method='post' action='#' style='display:none;'>"
		html = html & "<div class='bloque-datos ui-panel' style='max-width: 900px; margin: 0 auto; border:1px dashed #10b981;'>"
		html = html & "<div class='bloque-header ui-panel__header'>NUEVO AMBIENTE</div>"
		html = html & "<div class='bloque-body ui-panel__body'>"
		html = html & "<div class='grid-datos-doble ui-field-grid--dual'>"
		html = html & "<div class='dato-label'>TIPO AMBIENTE</div>"
		html = html & "<div><select class='input-moderno ui-select' id='agregar_tipo' style='width:90%;'><option value=''>AMBIENTE UNICO</option><option value='T'>TESTING</option><option value='P'>PRODUCCION</option></select></div>"
		html = html & "<div class='dato-label'>AMBIENTE (Url)</div><div>" & HtmlInputCampo("AMBIENTE", "", True) & "</div>"
		html = html & "<div class='dato-label'>BASE DATOS EM</div><div>" & HtmlInputCampo("BASE_EM", "", True) & "</div>"
		html = html & "<div class='dato-label'>BASE DATOS RQT</div><div>" & HtmlInputCampo("BASE_RQT", "", True) & "</div>"
		html = html & "<div class='dato-label'>RUTA SV-40 (Webintranet)</div><div>" & HtmlInputCampo("RUTA_SV", "", True) & "</div>"
		html = html & "</div></div>"
		html = html & "<div style='display:flex; justify-content:flex-end; gap:1rem; padding:1rem 0;'>"
		html = html & "<input class='btn-accion ui-button ui-button--primary' type='button' value='Agregar' style='background-color:#10b981; color:white; padding:0.75rem 1.5rem; border:none; border-radius:4px; cursor:pointer;' onclick=""var t=document.getElementById('agregar_tipo').value; document.getElementById('frmAgregar').action='detalleCliente.asp?opc=UPD&id=" & id & "&apli=" & apli & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'+'&tipo='+t; document.getElementById('frmAgregar').submit();"">"
		html = html & "<input class='bot ui-button' type='button' value='Cancelar' style='background-color:#6b7280; color:white; padding:0.75rem 1.5rem; border:none; border-radius:4px; cursor:pointer;' onclick=""document.getElementById('frmAgregar').style.display='none';"">"
		html = html & "</div></div></form>"
	End If

Else

	html = html & "<form id='frmNuevo' method='post' action='#'>"

	html = html & "<div class='bloque-datos ui-panel' style='max-width: 900px; margin: 0 auto;'>"
	html = html & "<div class='bloque-header ui-panel__header'>NUEVO AMBIENTE</div>"
	html = html & "<div class='bloque-body ui-panel__body'>"
	html = html & "<div class='grid-datos-doble ui-field-grid--dual'>"

	html = html & "<div class='dato-label'>TIPO AMBIENTE</div>"
	html = html & "<div>"
	html = html & "<select class='input-moderno ui-select' id='nuevo_tipo' name='nuevo_tipo' style='width:90%;'>"
	html = html & "<option value=''>AMBIENTE UNICO</option>"
	html = html & "<option value='T'>TESTING</option>"
	html = html & "<option value='P'>PRODUCCION</option>"
	html = html & "</select>"
	html = html & "</div>"

	html = html & "<div class='dato-label'>AMBIENTE (Url)</div>"
	html = html & "<div>" & HtmlInputCampo("AMBIENTE", "", True) & "</div>"

	html = html & "<div class='dato-label'>BASE DATOS EM</div>"
	html = html & "<div>" & HtmlInputCampo("BASE_EM", "", True) & "</div>"

	html = html & "<div class='dato-label'>BASE DATOS RQT</div>"
	html = html & "<div>" & HtmlInputCampo("BASE_RQT", "", True) & "</div>"

	html = html & "<div class='dato-label'>RUTA SV-40 (Webintranet)</div>"
	html = html & "<div>" & HtmlInputCampo("RUTA_SV", "", True) & "</div>"

	html = html & "</div>"
	html = html & "</div>"

	html = html & "<div style='display: flex; justify-content: flex-end; gap: 1rem; padding: 1rem 0;'>"

	If getPermisos(xmenu, xtmenu, "MODIF") = 1 Then
		html = html & "<input class='btn-accion ui-button ui-button--primary' type='button' value='Agregar' onclick=""var t=document.getElementById('nuevo_tipo').value; document.getElementById('frmNuevo').action='detalleCliente.asp?opc=UPD&id=" & id & "&apli=" & apli & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'+'&tipo='+t; document.getElementById('frmNuevo').submit();"">"
	End If

	html = html & "</div>"

	html = html & "</div>"
	html = html & "</form>"

End If

rz.Close

html = html & "<div style='max-width: 900px; margin: 0 auto; width: 100%; margin-top: 1rem; padding-top: 1rem; border-top: 1px dashed #cbd5e1; display: flex; justify-content: flex-end; gap: 1rem;'>"

html = html & "<input class='bot ui-button' type='button' value='CERRAR' style='background-color: #6b7280; color: white; padding: 0.75rem 1.5rem; border: none; border-radius: 4px; font-size: 1rem; cursor: pointer;' onclick=""document.getElementById('formx5').submit();"">"

html = html & "</div>"

html = html & "</div>"

html = html & ScriptEdicionCatalogo()

formx5 = "<form id='formx5' class='form-oculto' method='post' action='clientes.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"

html = html & "<script src='assets/js/jquery.min.js'></script>"
html = html & "<script src='assets/js/skel.min.js'></script>"
html = html & "<script src='assets/js/util.js'></script>"
html = html & "<script src='assets/js/main.js'></script>"

html = html & formx5 & "</body>"
html = html & "</html>"

response.write html
%>