<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/ui_componentes.inc" -->
<%
'[UI-ANT] ' ===== Cuerpo anterior completo: reemplazado por el bloque [UI] que sigue. Se borra entero buscando [UI-ANT] =====
'[UI-ANT] html = headhtml()
'[UI-ANT] html = html & "<body>"

'[UI-ANT] id = request.querystring("id")
'[UI-ANT] apli = request.querystring("apli")
'[UI-ANT] opc = request.querystring("opc")
'[UI-ANT] sel = request.form("sel")
'[UI-ANT] xmenu = request.querystring("xmenu")
'[UI-ANT] xtmenu = request.querystring("xtmenu")

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

'[UI-ANT] On Error Resume Next
'[UI-ANT] XXX = SYSLOGON(cnn, "PARAM_CLIEN")
'[UI-ANT] If Err.Number <> 0 Then
	'[UI-ANT] Response.Write "ERROR SYSLOGON: " & Err.Number & ": " & Err.Description
	'[UI-ANT] Response.End
'[UI-ANT] End If
'[UI-ANT] On Error Goto 0

'[UI-ANT] If apli & "" = "" Then
	'[UI-ANT] If sel & "" <> "" Then
		'[UI-ANT] apli = sel
	'[UI-ANT] Else
		'[UI-ANT] On Error Resume Next
		'[UI-ANT] Set rsPri = Exec(Cnn, "sp_cg_ClienAplic @Modo='primer_id', @IdCliente=" & id)
		'[UI-ANT] If Err.Number <> 0 Then
			'[UI-ANT] Response.Write "ERROR primer_id: " & Err.Number & ": " & Err.Description
			'[UI-ANT] Response.End
		'[UI-ANT] End If
		'[UI-ANT] On Error Goto 0

		'[UI-ANT] If Not rsPri.eof Then
			'[UI-ANT] apli = rsPri(0)
		'[UI-ANT] Else
			'[UI-ANT] apli = 0
		'[UI-ANT] End If
		'[UI-ANT] rsPri.Close
	'[UI-ANT] End If
'[UI-ANT] End If

'[UI-ANT] If opc = "UPD" Then

	'[UI-ANT] ambiente = Trim(request.form("AMBIENTE") & "")
	'[UI-ANT] base_em  = Trim(request.form("BASE_EM") & "")
	'[UI-ANT] base_rqt = Trim(request.form("BASE_RQT") & "")
	'[UI-ANT] ruta_sv  = Trim(request.form("RUTA_SV") & "")
	'[UI-ANT] tipo     = Trim(request.querystring("tipo") & "")

	'[UI-ANT] If ambiente = "" Or base_em = "" Or base_rqt = "" Or ruta_sv = "" Then
		'[UI-ANT] html = html & "<script>alert('Todos los campos son obligatorios.'); window.history.back();</script>"
	'[UI-ANT] 'ElseIf tipo = "" Then
	'[UI-ANT] '	html = html & "<script>alert('Error: tipo de ambiente no valido.'); window.history.back();</script>"
	'[UI-ANT] Else

		'[UI-ANT] On Error Resume Next
		'[UI-ANT] Set chk = Exec(Cnn, "sp_cg_ClienAplic @Modo='obtener', @IdCliente=" & id & ", @IdAplic=" & apli)
		'[UI-ANT] If Err.Number <> 0 Then
			'[UI-ANT] Response.Write "ERROR obtener: " & Err.Number & ": " & Err.Description
			'[UI-ANT] Response.End
		'[UI-ANT] End If
		'[UI-ANT] On Error Goto 0

		'[UI-ANT] encontro = False
		'[UI-ANT] Do While Not chk.eof
			'[UI-ANT] If UCase(Trim(chk("TIPO_AMB") & "")) = UCase(tipo) Then
				'[UI-ANT] encontro = True
				'[UI-ANT] Exit Do
			'[UI-ANT] End If
			'[UI-ANT] chk.movenext
		'[UI-ANT] Loop

		'[UI-ANT] If encontro Then

			'[UI-ANT] sSQL = "sp_cg_ClienAplic @Modo='actualizar'"
			'[UI-ANT] sSQL = sSQL & ", @IdCliente=" & id
			'[UI-ANT] sSQL = sSQL & ", @IdAplic=" & apli
			'[UI-ANT] sSQL = sSQL & ", @BaseEm='" & Replace(base_em, "'", "''") & "'"
			'[UI-ANT] sSQL = sSQL & ", @BaseRqt='" & Replace(base_rqt, "'", "''") & "'"
			'[UI-ANT] sSQL = sSQL & ", @RutaSv='" & Replace(ruta_sv, "'", "''") & "'"
			'[UI-ANT] sSQL = sSQL & ", @Ambiente='" & Replace(ambiente, "'", "''") & "'"
			'[UI-ANT] sSQL = sSQL & ", @TipoAmb='" & Replace(tipo, "'", "''") & "'"

			'[UI-ANT] On Error Resume Next
			'[UI-ANT] ExecNonQuery Cnn, sSQL
			'[UI-ANT] If Err.Number <> 0 Then
				'[UI-ANT] Response.Write "ERROR actualizar: " & Err.Number & ": " & Err.Description
				'[UI-ANT] Response.End
			'[UI-ANT] End If
			'[UI-ANT] On Error Goto 0

			'[UI-ANT] html = html & "<script>alert('Actualizado exitosamente.'); window.location.href='detalleCliente.asp?id=" & id & "&apli=" & apli & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "';</script>"
		'[UI-ANT] Else

			'[UI-ANT] sSQL = "sp_cg_ClienAplic @Modo='insertar'"
			'[UI-ANT] sSQL = sSQL & ", @IdCliente=" & id
			'[UI-ANT] sSQL = sSQL & ", @IdAplic=" & apli
			'[UI-ANT] sSQL = sSQL & ", @Tiene='1'"
			'[UI-ANT] sSQL = sSQL & ", @BaseEm='" & Replace(base_em, "'", "''") & "'"
			'[UI-ANT] sSQL = sSQL & ", @BaseRqt='" & Replace(base_rqt, "'", "''") & "'"
			'[UI-ANT] sSQL = sSQL & ", @RutaSv='" & Replace(ruta_sv, "'", "''") & "'"
			'[UI-ANT] sSQL = sSQL & ", @Ambiente='" & Replace(ambiente, "'", "''") & "'"
			'[UI-ANT] sSQL = sSQL & ", @TipoAmb='" & Replace(tipo, "'", "''") & "'"

			'[UI-ANT] On Error Resume Next
			'[UI-ANT] ExecNonQuery Cnn, sSQL
			'[UI-ANT] If Err.Number <> 0 Then
				'[UI-ANT] Response.Write "ERROR insertar: " & Err.Number & ": " & Err.Description
				'[UI-ANT] Response.End
			'[UI-ANT] End If
			'[UI-ANT] On Error Goto 0

			'[UI-ANT] If ambiente <> "" Or tipo <> "" Then
				'[UI-ANT] sSQL2 = "sp_cg_ClienAplic @Modo='actualizar'"
				'[UI-ANT] sSQL2 = sSQL2 & ", @IdCliente=" & id
				'[UI-ANT] sSQL2 = sSQL2 & ", @IdAplic=" & apli
				'[UI-ANT] If ambiente <> "" Then sSQL2 = sSQL2 & ", @Ambiente='" & Replace(ambiente, "'", "''") & "'"
				'[UI-ANT] If tipo <> "" Then sSQL2 = sSQL2 & ", @TipoAmb='" & Replace(tipo, "'", "''") & "'"

				'[UI-ANT] On Error Resume Next
				'[UI-ANT] ExecNonQuery Cnn, sSQL2
				'[UI-ANT] If Err.Number <> 0 Then
					'[UI-ANT] Response.Write "ERROR insertar ambiente: " & Err.Number & ": " & Err.Description
					'[UI-ANT] Response.End
				'[UI-ANT] End If
				'[UI-ANT] On Error Goto 0
			'[UI-ANT] End If

			'[UI-ANT] html = html & "<script>alert('Registrado exitosamente.'); window.location.href='detalleCliente.asp?id=" & id & "&apli=" & apli & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "';</script>"
		'[UI-ANT] End If

		'[UI-ANT] chk.Close
	'[UI-ANT] End If

'[UI-ANT] ElseIf opc = "DEL" Then

	'[UI-ANT] delId = Trim(request.querystring("delId") & "")
	'[UI-ANT] If delId = "" Then delId = Trim(request.form("delId") & "")
	'[UI-ANT] If delId <> "" And IsNumeric(delId) Then
		'[UI-ANT] usuarioDel = session("usuario") & ""
		'[UI-ANT] If usuarioDel = "" Then usuarioDel = UserLogon() & ""
		'[UI-ANT] sSQL = "sp_cg_Registro_Eliminar @Opt='clxapli', @Id=" & CLng(delId) & ", @Usuario='" & Replace(usuarioDel, "'", "''") & "'"
		'[UI-ANT] On Error Resume Next
		'[UI-ANT] Cnn.Execute(sSQL)
		'[UI-ANT] If Err.Number <> 0 Then
			'[UI-ANT] Response.Write "ERROR eliminar ambiente: " & Err.Number & ": " & Err.Description
			'[UI-ANT] Response.End
		'[UI-ANT] End If
		'[UI-ANT] On Error Goto 0
		'[UI-ANT] html = html & "<script>alert('Ambiente eliminado.'); window.location.href='detalleCliente.asp?id=" & id & "&apli=" & apli & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "';</script>"
	'[UI-ANT] Else
		'[UI-ANT] html = html & "<script>alert('ID de ambiente no valido.'); window.history.back();</script>"
	'[UI-ANT] End If

'[UI-ANT] End If

'[UI-ANT] On Error Resume Next
'[UI-ANT] Set rs = Exec(Cnn, "sp_cg_Utilidades @Modo='consultar_campo', @Tabla='clientes', @Id=" & id)
'[UI-ANT] If Err.Number <> 0 Then
	'[UI-ANT] Response.Write "ERROR consultar_campo: " & Err.Number & ": " & Err.Description
	'[UI-ANT] Response.End
'[UI-ANT] End If
'[UI-ANT] On Error Goto 0

'[UI-ANT] tit = ""
'[UI-ANT] If Not rs.eof Then tit = rs(0)
'[UI-ANT] rs.Close

'[UI-ANT] html = html & cabecera()

'[UI-ANT] html = html & "<div class='titulo-seccion ui-page-header' style='padding: 0.5rem 2rem;background-color:#e8edff;'>"
'[UI-ANT] html = html & "<div class='ui-toolbar' style='display:flex;align-items:center;justify-content:space-between;'>"

'[UI-ANT] html = html & "<div><h2>" & Server.HTMLEncode(tit) & "</h2></div>"

'[UI-ANT] html = html & "<div style='display:flex;align-items:center;gap:0.5rem;'>"
'[UI-ANT] html = html & "<strong>Aplicativo:</strong>"

'[UI-ANT] html = html & "<form id='sForm' method='post' action='detalleCliente.asp?id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"

'[UI-ANT] html = html & "<select class='selec ui-select' id='sel' name='sel' style='height:50%;width:70%;' onchange=""document.getElementById('sForm').submit();"">"

'[UI-ANT] On Error Resume Next
'[UI-ANT] Set rx = Exec(Cnn, "sp_cg_ClienAplic @Modo='listar_x_cliente', @IdCliente=" & id)
'[UI-ANT] If Err.Number <> 0 Then
	'[UI-ANT] Response.Write "ERROR listar_x_cliente: " & Err.Number & ": " & Err.Description
	'[UI-ANT] Response.End
'[UI-ANT] End If
'[UI-ANT] On Error Goto 0

'[UI-ANT] Do While Not rx.eof
	'[UI-ANT] If CStr(apli & "") = CStr(rx(0) & "") Then
		'[UI-ANT] html = html & "<option value='" & rx(0) & "' selected>" & Server.HTMLEncode(rx(1) & "") & "</option>"
	'[UI-ANT] Else
		'[UI-ANT] html = html & "<option value='" & rx(0) & "'>" & Server.HTMLEncode(rx(1) & "") & "</option>"
	'[UI-ANT] End If
	'[UI-ANT] rx.movenext
'[UI-ANT] Loop

'[UI-ANT] html = html & "</select>"
'[UI-ANT] html = html & "</form>"
'[UI-ANT] html = html & "</div>" ' Fin selector Aplicativo

'[UI-ANT] formx5 = "<form id='formx5' class='form-oculto' method='post' action='clientes.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
'[UI-ANT] botonx5 = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('formx5').submit();"">"
'[UI-ANT] html = html & "<div class='ui-toolbar__actions'>" & botonx5 & "</div>"

'[UI-ANT] html = html & "</div>" ' Fin ui-toolbar
'[UI-ANT] html = html & "</div>" ' Fin titulo-seccion

'[UI-ANT] html = html & "<div class='inner ui-main' style='padding: 2rem; display: flex; flex-direction: column; gap: 2rem;'>"

'[UI-ANT] On Error Resume Next
'[UI-ANT] Set rz = Exec(Cnn, "sp_cg_ClienAplic @Modo='obtener', @IdCliente=" & id & ", @IdAplic=" & apli)
'[UI-ANT] If Err.Number <> 0 Then
	'[UI-ANT] Response.Write "ERROR obtener datos: " & Err.Number & ": " & Err.Description
	'[UI-ANT] Response.End
'[UI-ANT] End If
'[UI-ANT] On Error Goto 0

'[UI-ANT] numReg = 0

'[UI-ANT] If Not rz.eof Then

	'[UI-ANT] Do While Not rz.eof

		'[UI-ANT] numReg = numReg + 1
		'[UI-ANT] vId         = rz("ID") & ""
		'[UI-ANT] vAmbiente   = rz("AMBIENTE") & ""
		'[UI-ANT] vBaseEm     = rz("BASE_EM") & ""
		'[UI-ANT] vBaseRqt    = rz("BASE_RQT") & ""
		'[UI-ANT] vRutaSv     = rz("RUTA_SV") & ""
		'[UI-ANT] vTipo       = rz("TIPO_AMB") & ""

		'[UI-ANT] If vTipo = "T" Then
			'[UI-ANT] stit = "TESTING"
		'[UI-ANT] ElseIf vTipo = "P" Then
			'[UI-ANT] stit = "PRODUCCION"
		'[UI-ANT] Else
			'[UI-ANT] stit = "AMBIENTE UNICO"
		'[UI-ANT] End If

		'[UI-ANT] frmId = "frmCliente_" & numReg

		'[UI-ANT] html = html & "<form id='" & frmId & "' method='post' action='detalleCliente.asp?opc=UPD&id=" & id & "&apli=" & apli & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&tipo=" & Server.HTMLEncode(vTipo) & "'>"

		'[UI-ANT] html = html & "<div class='bloque-datos ui-panel' style='max-width: 900px; margin: 0 auto;'>"
		'[UI-ANT] html = html & "<div class='bloque-header ui-panel__header'>" & stit & "</div>"
		'[UI-ANT] html = html & "<div class='bloque-body ui-panel__body'>"
		'[UI-ANT] html = html & "<div class='grid-datos-doble ui-field-grid--dual'>"

		'[UI-ANT] html = html & "<div class='dato-label'>AMBIENTE (Url)</div>"
		'[UI-ANT] html = html & "<div>" & HtmlInputCampo("AMBIENTE", vAmbiente, False) & "</div>"

		'[UI-ANT] html = html & "<div class='dato-label'>BASE DATOS EM</div>"
		'[UI-ANT] html = html & "<div>" & HtmlInputCampo("BASE_EM", vBaseEm, False) & "</div>"

		'[UI-ANT] html = html & "<div class='dato-label'>BASE DATOS RQT</div>"
		'[UI-ANT] html = html & "<div>" & HtmlInputCampo("BASE_RQT", vBaseRqt, False) & "</div>"

		'[UI-ANT] html = html & "<div class='dato-label'>RUTA SV-40 (Webintranet)</div>"
		'[UI-ANT] html = html & "<div>" & HtmlInputCampo("RUTA_SV", vRutaSv, False) & "</div>"

		'[UI-ANT] html = html & "</div>"
		'[UI-ANT] html = html & "</div>"

		'[UI-ANT] html = html & "<div style='display: flex; justify-content: flex-end; gap: 1rem; padding: 1rem 0;'>"

		'[UI-ANT] If getPermisos(xmenu, xtmenu, "MODIF") = 1 Then
			'[UI-ANT] html = html & "<input class='btn-editar ui-button' type='button' value='&#9998; Editar' title='Editar ambiente' style='background-color:#6b7280; color:white; padding:0.5rem 1rem; border:none; border-radius:4px; font-size:0.95rem; cursor:pointer;' onclick=""toggleEditarFila('" & frmId & "', this);"">"
		'[UI-ANT] End If
		'[UI-ANT] If getPermisos(xmenu, xtmenu, "ELIM") = 1 Then
			'[UI-ANT] html = html & "<input class='btn-eliminar ui-button ui-button--danger' type='button' value='&#128465; Eliminar' title='Eliminar ambiente' style='background-color:#dc2626; color:white; padding:0.5rem 1rem; border:none; border-radius:4px; font-size:0.95rem; cursor:pointer;' onclick=""if(confirm('?Eliminar este ambiente? Esta accion no se puede deshacer.')){window.location.href='detalleCliente.asp?opc=DEL&delId=" & vId & "&id=" & id & "&apli=" & apli & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "';}"">"
		'[UI-ANT] End If

		'[UI-ANT] html = html & "</div>"

		'[UI-ANT] html = html & "</div>"
		'[UI-ANT] html = html & "</form>"

		'[UI-ANT] rz.movenext
	'[UI-ANT] Loop

	'[UI-ANT] ' --- Boton + Form para AGREGAR nuevo ambiente (visible siempre que hay datos) ---
	'[UI-ANT] If getPermisos(xmenu, xtmenu, "MODIF") = 1 Then
		'[UI-ANT] html = html & "<div style='max-width: 900px; margin: 0 auto; display:flex; justify-content:flex-end; padding:0.5rem 0;'>"
		'[UI-ANT] html = html & "<input class='bot ui-button ui-button--primary' type='button' value='+ AGREGAR AMBIENTE' style='background-color:#10b981; color:white; padding:0.75rem 1.5rem; border:none; border-radius:4px; font-size:1rem; cursor:pointer;' onclick=""document.getElementById('frmAgregar').style.display=(document.getElementById('frmAgregar').style.display=='none'?'block':'none');"">"
		'[UI-ANT] html = html & "</div>"

		'[UI-ANT] html = html & "<form id='frmAgregar' method='post' action='#' style='display:none;'>"
		'[UI-ANT] html = html & "<div class='bloque-datos ui-panel' style='max-width: 900px; margin: 0 auto; border:1px dashed #10b981;'>"
		'[UI-ANT] html = html & "<div class='bloque-header ui-panel__header'>NUEVO AMBIENTE</div>"
		'[UI-ANT] html = html & "<div class='bloque-body ui-panel__body'>"
		'[UI-ANT] html = html & "<div class='grid-datos-doble ui-field-grid--dual'>"
		'[UI-ANT] html = html & "<div class='dato-label'>TIPO AMBIENTE</div>"
		'[UI-ANT] html = html & "<div><select class='input-moderno ui-select' id='agregar_tipo' style='width:90%;'><option value=''>AMBIENTE UNICO</option><option value='T'>TESTING</option><option value='P'>PRODUCCION</option></select></div>"
		'[UI-ANT] html = html & "<div class='dato-label'>AMBIENTE (Url)</div><div>" & HtmlInputCampo("AMBIENTE", "", True) & "</div>"
		'[UI-ANT] html = html & "<div class='dato-label'>BASE DATOS EM</div><div>" & HtmlInputCampo("BASE_EM", "", True) & "</div>"
		'[UI-ANT] html = html & "<div class='dato-label'>BASE DATOS RQT</div><div>" & HtmlInputCampo("BASE_RQT", "", True) & "</div>"
		'[UI-ANT] html = html & "<div class='dato-label'>RUTA SV-40 (Webintranet)</div><div>" & HtmlInputCampo("RUTA_SV", "", True) & "</div>"
		'[UI-ANT] html = html & "</div></div>"
		'[UI-ANT] html = html & "<div style='display:flex; justify-content:flex-end; gap:1rem; padding:1rem 0;'>"
		'[UI-ANT] html = html & "<input class='btn-accion ui-button ui-button--primary' type='button' value='Agregar' style='background-color:#10b981; color:white; padding:0.75rem 1.5rem; border:none; border-radius:4px; cursor:pointer;' onclick=""var t=document.getElementById('agregar_tipo').value; document.getElementById('frmAgregar').action='detalleCliente.asp?opc=UPD&id=" & id & "&apli=" & apli & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'+'&tipo='+t; document.getElementById('frmAgregar').submit();"">"
		'[UI-ANT] html = html & "<input class='bot ui-button' type='button' value='Cancelar' style='background-color:#6b7280; color:white; padding:0.75rem 1.5rem; border:none; border-radius:4px; cursor:pointer;' onclick=""document.getElementById('frmAgregar').style.display='none';"">"
		'[UI-ANT] html = html & "</div></div></form>"
	'[UI-ANT] End If

'[UI-ANT] Else

	'[UI-ANT] html = html & "<form id='frmNuevo' method='post' action='#'>"

	'[UI-ANT] html = html & "<div class='bloque-datos ui-panel' style='max-width: 900px; margin: 0 auto;'>"
	'[UI-ANT] html = html & "<div class='bloque-header ui-panel__header'>NUEVO AMBIENTE</div>"
	'[UI-ANT] html = html & "<div class='bloque-body ui-panel__body'>"
	'[UI-ANT] html = html & "<div class='grid-datos-doble ui-field-grid--dual'>"

	'[UI-ANT] html = html & "<div class='dato-label'>TIPO AMBIENTE</div>"
	'[UI-ANT] html = html & "<div>"
	'[UI-ANT] html = html & "<select class='input-moderno ui-select' id='nuevo_tipo' name='nuevo_tipo' style='width:90%;'>"
	'[UI-ANT] html = html & "<option value=''>AMBIENTE UNICO</option>"
	'[UI-ANT] html = html & "<option value='T'>TESTING</option>"
	'[UI-ANT] html = html & "<option value='P'>PRODUCCION</option>"
	'[UI-ANT] html = html & "</select>"
	'[UI-ANT] html = html & "</div>"

	'[UI-ANT] html = html & "<div class='dato-label'>AMBIENTE (Url)</div>"
	'[UI-ANT] html = html & "<div>" & HtmlInputCampo("AMBIENTE", "", True) & "</div>"

	'[UI-ANT] html = html & "<div class='dato-label'>BASE DATOS EM</div>"
	'[UI-ANT] html = html & "<div>" & HtmlInputCampo("BASE_EM", "", True) & "</div>"

	'[UI-ANT] html = html & "<div class='dato-label'>BASE DATOS RQT</div>"
	'[UI-ANT] html = html & "<div>" & HtmlInputCampo("BASE_RQT", "", True) & "</div>"

	'[UI-ANT] html = html & "<div class='dato-label'>RUTA SV-40 (Webintranet)</div>"
	'[UI-ANT] html = html & "<div>" & HtmlInputCampo("RUTA_SV", "", True) & "</div>"

	'[UI-ANT] html = html & "</div>"
	'[UI-ANT] html = html & "</div>"

	'[UI-ANT] html = html & "<div style='display: flex; justify-content: flex-end; gap: 1rem; padding: 1rem 0;'>"

	'[UI-ANT] If getPermisos(xmenu, xtmenu, "MODIF") = 1 Then
		'[UI-ANT] html = html & "<input class='btn-accion ui-button ui-button--primary' type='button' value='Agregar' onclick=""var t=document.getElementById('nuevo_tipo').value; document.getElementById('frmNuevo').action='detalleCliente.asp?opc=UPD&id=" & id & "&apli=" & apli & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'+'&tipo='+t; document.getElementById('frmNuevo').submit();"">"
	'[UI-ANT] End If

	'[UI-ANT] html = html & "</div>"

	'[UI-ANT] html = html & "</div>"
	'[UI-ANT] html = html & "</form>"

'[UI-ANT] End If

'[UI-ANT] rz.Close

'[UI-ANT] html = html & "<div style='max-width: 900px; margin: 0 auto; width: 100%; margin-top: 1rem; padding-top: 1rem; border-top: 1px dashed #cbd5e1; display: flex; justify-content: flex-end; gap: 1rem;'>"

'[UI-ANT] html = html & "<input class='bot ui-button' type='button' value='CERRAR' style='background-color: #6b7280; color: white; padding: 0.75rem 1.5rem; border: none; border-radius: 4px; font-size: 1rem; cursor: pointer;' onclick=""document.getElementById('formx5').submit();"">"

'[UI-ANT] html = html & "</div>"

'[UI-ANT] html = html & "</div>"

'[UI-ANT] html = html & ScriptEdicionCatalogo()

'[UI-ANT] formx5 = "<form id='formx5' class='form-oculto' method='post' action='clientes.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"

'[UI-ANT] html = html & "<script src='assets/js/jquery.min.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/skel.min.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/util.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/main.js'></script>"

'[UI-ANT] html = html & formx5 & "</body>"
'[UI-ANT] html = html & "</html>"

'[UI-ANT] response.write html

' =========================================================================
' [UI] PANTALLA con el sistema visual (IncFile/ui_componentes.inc)
' Informacion local del cliente por aplicativo: ambientes (testing, produccion o unico) con su URL,
' bases y ruta. Se llega desde Clientes > Info local.
' El cuerpo anterior completo queda arriba, comentado con '[UI-ANT].
' Consultas y procedimientos: los mismos de antes. Cambios: ids y tipo validados, permisos
' chequeados en el servidor y sin alta de ambientes cuando el cliente no tiene aplicativos.
' =========================================================================
Dim puedeModif, puedeElim, accTop, filtrosHtml, selHtml, cantApl, sinPermiso, tituloAmb, frmId, accAmb

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

' [UI] Los ids van a la consulta SQL: se validan como numeros (antes se pegaban tal cual)
If IsNumeric(id & "") Then id = CLng(id) Else id = 0
If Not IsNumeric(sel & "") Then sel = ""
If Not IsNumeric(apli & "") Then apli = ""
puedeModif = (getPermisosAtrm("CG_CLIENTES", "MODIF") = 1)
' [LOG] Consulta del detalle de un cliente (CG_LOG)
If IsNumeric(id & "") Then CgLog "CLIENTE", id, "CONSULTA"
puedeElim = (getPermisosAtrm("CG_CLIENTES", "ELIM") = 1)
sinPermiso = "<script>alert('No ten\u00e9s permiso para esta acci\u00f3n.'); window.history.back();</script>"

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
apli = CLng("0" & apli)

If opc = "UPD" Then

	ambiente = Trim(request.form("AMBIENTE") & "")
	base_em  = Trim(request.form("BASE_EM") & "")
	base_rqt = Trim(request.form("BASE_RQT") & "")
	ruta_sv  = Trim(request.form("RUTA_SV") & "")
	tipo     = UCase(Trim(request.querystring("tipo") & ""))
	' [UI] El tipo solo puede ser testing (T), produccion (P) o unico (vacio)
	If tipo <> "T" And tipo <> "P" Then tipo = ""

	If Not puedeModif Then
		html = html & sinPermiso
	ElseIf apli = 0 Then
		html = html & "<script>alert('El cliente no tiene aplicativos asignados.'); window.history.back();</script>"
	ElseIf ambiente = "" Or base_em = "" Or base_rqt = "" Or ruta_sv = "" Then
		html = html & "<script>alert('Todos los campos son obligatorios.'); window.history.back();</script>"
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
			sSQL = sSQL & ", @TipoAmb='" & tipo & "'"

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
			sSQL = sSQL & ", @TipoAmb='" & tipo & "'"

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
				If tipo <> "" Then sSQL2 = sSQL2 & ", @TipoAmb='" & tipo & "'"

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
	If Not puedeElim Then
		' [UI] El permiso se chequea tambien aca: antes solo se ocultaba el boton
		html = html & sinPermiso
	ElseIf delId <> "" And IsNumeric(delId) Then
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
		html = html & "<script>alert('ID de ambiente no v\u00e1lido.'); window.history.back();</script>"
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

html = html & "<div id='wrapper' role='banner'>" & cabecera() & "</div>"

' --- Barra superior: aplicativo (se cambia solo al elegirlo, como antes) y Volver ---
On Error Resume Next
Set rx = Exec(Cnn, "sp_cg_ClienAplic @Modo='listar_x_cliente', @IdCliente=" & id)
If Err.Number <> 0 Then
	Response.Write "ERROR listar_x_cliente: " & Err.Number & ": " & Err.Description
	Response.End
End If
On Error Goto 0

cantApl = 0
selHtml = "<select class='ui-select' id='sel' name='sel' onchange=""document.getElementById('sForm').submit();"">"
Do While Not rx.eof
	If CStr(apli & "") = CStr(rx(0) & "") Then
		selHtml = selHtml & "<option value='" & rx(0) & "' selected>" & Server.HTMLEncode(rx(1) & "") & "</option>"
	Else
		selHtml = selHtml & "<option value='" & rx(0) & "'>" & Server.HTMLEncode(rx(1) & "") & "</option>"
	End If
	cantApl = cantApl + 1
	rx.movenext
Loop
selHtml = selHtml & "</select>"

formx5 = "<form id='formx5' class='form-oculto' method='post' action='clientes.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
accTop = ""
If cantApl > 0 Then
	accTop = "<form id='sForm' class='ui-card__filtros' method='post' action='detalleCliente.asp?id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>" & _
		UiFiltroCompacto("sel", "Aplicativo", selHtml) & "</form>"
End If
accTop = accTop & UiBoton("Volver", "secondary", "document.getElementById('formx5').submit();", True, "")

If tit = "" Then
	html = html & UiPantallaParametrosAbre("Informaci&oacute;n local", accTop, "")
Else
	html = html & UiPantallaParametrosAbre(UiTexto(tit), accTop, "")
End If

If cantApl = 0 Then
	' [UI] Sin aplicativos no hay donde guardar un ambiente (antes ofrecia el alta igual, con aplicativo 0)
	html = html & UiCardAbre("Informaci&oacute;n local", "") & "<div class='ui-empty'>El cliente no tiene aplicativos asignados. Se asignan desde Aplicativos &gt; Ver / modificar.</div>" & UiCardCierra()
Else

	On Error Resume Next
	Set rz = Exec(Cnn, "sp_cg_ClienAplic @Modo='obtener', @IdCliente=" & id & ", @IdAplic=" & apli)
	If Err.Number <> 0 Then
		Response.Write "ERROR obtener datos: " & Err.Number & ": " & Err.Description
		Response.End
	End If
	On Error Goto 0

	numReg = 0

	If Not rz.eof Then

		' --- Una tarjeta por ambiente: los campos se habilitan con Editar (toggleEditarFila, de funciones.inc) ---
		Do While Not rz.eof
			numReg = numReg + 1
			vId         = rz("ID") & ""
			vAmbiente   = rz("AMBIENTE") & ""
			vBaseEm     = rz("BASE_EM") & ""
			vBaseRqt    = rz("BASE_RQT") & ""
			vRutaSv     = rz("RUTA_SV") & ""
			vTipo       = rz("TIPO_AMB") & ""

			If vTipo = "T" Then
				tituloAmb = "Testing"
			ElseIf vTipo = "P" Then
				tituloAmb = "Producci&oacute;n"
			Else
				tituloAmb = "Ambiente &uacute;nico"
			End If

			frmId = "frmCliente_" & numReg
			accAmb = ""
			If puedeModif Then accAmb = accAmb & "<input class='ui-button ui-button--light ui-button--sm' type='button' value='Editar' aria-label='Editar ambiente " & DetTextoPlano(tituloAmb) & "' onclick=""toggleEditarFila('" & frmId & "', this);"">"
			If puedeElim Then accAmb = accAmb & UiBoton("Eliminar", "danger", "if (confirm('\u00bfEliminar este ambiente? No se puede deshacer.')) { window.location.href='detalleCliente.asp?opc=DEL&delId=" & vId & "&id=" & id & "&apli=" & apli & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'; }", True, "")

			html = html & "<form id='" & frmId & "' method='post' action='detalleCliente.asp?opc=UPD&id=" & id & "&apli=" & apli & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&tipo=" & Server.HTMLEncode(vTipo) & "'>"
			If accAmb <> "" Then accAmb = "<div class='ui-card__header-acciones'>" & accAmb & "</div>"
			html = html & UiCardAbre(tituloAmb, accAmb) & "<div class='ui-form-grid'>"
			html = html & DetCampo("Ambiente (URL)", HtmlInputCampo("AMBIENTE", vAmbiente, False), True)
			html = html & DetCampo("Base de datos EM", HtmlInputCampo("BASE_EM", vBaseEm, False), False)
			html = html & DetCampo("Base de datos RQT", HtmlInputCampo("BASE_RQT", vBaseRqt, False), False)
			html = html & DetCampo("Ruta SV-40 (Webintranet)", HtmlInputCampo("RUTA_SV", vRutaSv, False), True)
			html = html & "</div>" & UiCardCierra() & "</form>"

			rz.movenext
		Loop

		' --- Agregar ambiente: tarjeta oculta que se muestra con el boton ---
		If puedeModif Then
			html = html & "<div class='ui-acciones-pie'>" & UiBoton("+ Agregar ambiente", "light", "var f = document.getElementById('frmAgregar'); f.hidden = !f.hidden; if (!f.hidden) f.querySelector('select').focus();", True, "") & "</div>"
			html = html & HtmlNuevoAmbiente("frmAgregar", True)
		End If

	Else
		' --- Sin ambientes: alta directa ---
		If puedeModif Then
			html = html & HtmlNuevoAmbiente("frmNuevo", False)
		Else
			html = html & UiCardAbre("Ambientes", "") & "<div class='ui-empty'>El aplicativo no tiene ambientes cargados para este cliente.</div>" & UiCardCierra()
		End If
	End If

	rz.Close
End If

html = html & UiPantallaCierra()

html = html & ScriptEdicionCatalogo()

html = html & "<script src='assets/js/jquery.min.js'></script>"
html = html & "<script src='assets/js/skel.min.js'></script>"
html = html & "<script src='assets/js/util.js'></script>"
html = html & "<script src='assets/js/main.js'></script>"

html = html & formx5 & "</body>"
html = html & "</html>"

response.write html


' =============================================================================
' FUNCIONES DE ESTA PANTALLA
' =============================================================================

' Campo con su etiqueta. El label envuelve al control: queda asociado aunque HtmlInputCampo no le ponga id
Function DetCampo(etiqueta, controlHtml, ancho)
	Dim clase
	clase = "ui-field"
	If ancho Then clase = clase & " ui-form-grid__ancho"
	DetCampo = "<label class='" & clase & "'><span class='ui-field__label'>" & etiqueta & "</span>" & controlHtml & "</label>"
End Function

' Tarjeta de alta de ambiente (tipo + 4 campos). oculta: empieza escondida (se abre con Agregar ambiente)
Function HtmlNuevoAmbiente(formId, oculta)
	Dim h, urlBase
	urlBase = "detalleCliente.asp?opc=UPD&id=" & id & "&apli=" & apli & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&tipo="
	h = "<form id='" & formId & "' method='post' action='" & urlBase & "'"
	If oculta Then h = h & " hidden"
	h = h & " onsubmit=""this.action = '" & urlBase & "' + this.querySelector('select').value;"">"
	h = h & UiCardAbre("Nuevo ambiente", "") & "<div class='ui-form-grid'>"
	h = h & DetCampo("Tipo de ambiente", "<select class='ui-select'><option value=''>Ambiente &uacute;nico</option><option value='T'>Testing</option><option value='P'>Producci&oacute;n</option></select>", False)
	h = h & "<div></div>"
	h = h & DetCampo("Ambiente (URL) <span class='ui-field__required' aria-hidden='true'> *</span>", HtmlInputCampo("AMBIENTE", "", True), True)
	h = h & DetCampo("Base de datos EM <span class='ui-field__required' aria-hidden='true'> *</span>", HtmlInputCampo("BASE_EM", "", True), False)
	h = h & DetCampo("Base de datos RQT <span class='ui-field__required' aria-hidden='true'> *</span>", HtmlInputCampo("BASE_RQT", "", True), False)
	h = h & DetCampo("Ruta SV-40 (Webintranet) <span class='ui-field__required' aria-hidden='true'> *</span>", HtmlInputCampo("RUTA_SV", "", True), True)
	h = h & "</div><div class='ui-card__acciones'>"
	If oculta Then h = h & UiBoton("Cancelar", "secondary", "document.getElementById('" & formId & "').hidden = true;", True, "")
	h = h & UiBotonEnviar("Agregar ambiente", True, "") & "</div>" & UiCardCierra() & "</form>"
	HtmlNuevoAmbiente = h
End Function

' Titulo del ambiente sin entidades HTML (para aria-label)
Function DetTextoPlano(v)
	DetTextoPlano = Replace(Replace(v, "&oacute;", "o"), "&uacute;", "u")
End Function
%>