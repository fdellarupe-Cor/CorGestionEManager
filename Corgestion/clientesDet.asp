<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/datosActualizar.inc" -->
<!-- #INCLUDE FILE="IncFile/ui_componentes.inc" -->
<%
'[UI-ANT] ' ===== Cuerpo anterior completo: reemplazado por el bloque [UI] que sigue. Se borra entero buscando [UI-ANT] =====
'[UI-ANT] html = headhtml()
'[UI-ANT] html = html & "<body>"

'[UI-ANT] id      = request.querystring("id")
'[UI-ANT] modo    = request.querystring("modo")
'[UI-ANT] opc     = request.querystring("opc")
'[UI-ANT] xmenu   = request.querystring("xmenu")
'[UI-ANT] xtmenu  = request.querystring("xtmenu")

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
'[UI-ANT] XXX = SYSLOGON(cnn,"PARAM_CLIEN")
'[UI-ANT] If Err.Number <> 0 Then
	'[UI-ANT] Response.Write "ERROR SYSLOGON: " & Err.Number & ": " & Err.Description
	'[UI-ANT] Response.End
'[UI-ANT] End If
'[UI-ANT] On Error Goto 0

'[UI-ANT] ' --- PROCESAMIENTO DE GUARDADO ---
'[UI-ANT] If opc = "UPD_CLI" Then

	'[UI-ANT] direccion   = Trim(request.form("direccion") & "")
	'[UI-ANT] ruta_ficha  = Trim(request.form("ruta_ficha") & "")

	'[UI-ANT] If direccion = "" Then
		'[UI-ANT] html = html & "<script>alert('La direccion es obligatoria.'); window.history.back();</script>"
	'[UI-ANT] Else
		'[UI-ANT] ' sSQL = "sp_cg_Catalogo_Actualizar @Tabla='clientes', @Campo='direccion', @Valor='" & Replace(direccion,"'","''") & "', @Id=" & id
		'[UI-ANT] ' On Error Resume Next
		'[UI-ANT] ' Cnn.Execute(sSQL)
		'[UI-ANT] ' If Err.Number <> 0 Then
			'[UI-ANT] ' Response.Write "ERROR actualizar direccion: " & Err.Number & ": " & Err.Description
			'[UI-ANT] ' Response.End
		'[UI-ANT] ' End If
		'[UI-ANT] ' On Error Goto 0

		'[UI-ANT] sSQL = "sp_cg_Catalogo_Actualizar @Tabla='clientes_ficha', @RutaFicha='" & Replace(ruta_ficha,"'","''") & "',@Direccion='" & Replace(direccion,"'","''") &"', @Id=" & id
		'[UI-ANT] response.write sSQL
		'[UI-ANT] On Error Resume Next
		'[UI-ANT] ExecNonQuery Cnn, sSQL
		'[UI-ANT] If Err.Number <> 0 Then
			'[UI-ANT] Response.Write "ERROR actualizar ruta_ficha: " & Err.Number & ": " & Err.Description
			'[UI-ANT] Response.End
		'[UI-ANT] End If
		'[UI-ANT] On Error Goto 0

		'[UI-ANT] html = html & "<script>alert('Datos del cliente actualizados.'); window.location.href='clientesDet.asp?id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "';</script>"
	'[UI-ANT] End If

'[UI-ANT] ElseIf opc = "UPD_ALL_DET" Then

	'[UI-ANT] totalEntornos = Trim(request.form("total_entornos") & "")
	'[UI-ANT] If totalEntornos = "" Then totalEntornos = "-1"

	'[UI-ANT] EliminarClientesDet Cnn, id

	'[UI-ANT] i = 0
	'[UI-ANT] Do While i <= CInt(totalEntornos)
		'[UI-ANT] srv   = Trim(request.form("srv_" & i) & "")
		'[UI-ANT] amb   = Trim(request.form("amb_" & i) & "")
		'[UI-ANT] dbrqt = Trim(request.form("dbrqt_" & i) & "")
		'[UI-ANT] lic   = Trim(request.form("lic_" & i) & "")

		'[UI-ANT] If srv <> "" And amb <> "" Then
			'[UI-ANT] srv   = Replace(srv, "'", "''")
			'[UI-ANT] amb   = Replace(amb, "'", "''")
			'[UI-ANT] dbrqt = Replace(dbrqt, "'", "''")
			'[UI-ANT] lic   = Replace(lic, "'", "''")
			'[UI-ANT] InsertarClientesDet Cnn, id, srv, amb, dbrqt, lic
		'[UI-ANT] End If
		'[UI-ANT] i = i + 1
	'[UI-ANT] Loop

	'[UI-ANT] html = html & "<script>alert('Entornos actualizados.'); window.location.href='clientesDet.asp?id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "';</script>"

'[UI-ANT] ElseIf opc = "INS_DET" Then

	'[UI-ANT] servidor  = Trim(request.form("servidor") & "")
	'[UI-ANT] ambiente  = Trim(request.form("ambiente") & "")
	'[UI-ANT] db_rqt    = Trim(request.form("db_rqt") & "")
	'[UI-ANT] licencia  = Trim(request.form("licencia") & "")

	'[UI-ANT] If servidor = "" Or ambiente = "" Then
		'[UI-ANT] html = html & "<script>alert('Servidor y Ambiente son obligatorios.'); window.history.back();</script>"
	'[UI-ANT] Else
		'[UI-ANT] servidor  = Replace(servidor,"'","''")
		'[UI-ANT] ambiente  = Replace(ambiente,"'","''")
		'[UI-ANT] db_rqt    = Replace(db_rqt,"'","''")
		'[UI-ANT] licencia  = Replace(licencia,"'","''")
		'[UI-ANT] InsertarClientesDet Cnn, id, servidor, ambiente, db_rqt, licencia
		'[UI-ANT] html = html & "<script>alert('Entorno agregado.'); window.location.href='clientesDet.asp?id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "';</script>"
	'[UI-ANT] End If

'[UI-ANT] End If

'[UI-ANT] ' --- LECTURA DE DATOS ---
'[UI-ANT] If id & "" = "" Then
	'[UI-ANT] Set rs = Exec(Cnn, "sp_cg_Utilidades @Modo='consultar_cliente_det', @Id=0")
	'[UI-ANT] If Not rs.eof Then id = rs(0)
	'[UI-ANT] rs.close
'[UI-ANT] End If

'[UI-ANT] Set rs = Exec(Cnn, "sp_cg_Utilidades @Modo='consultar_cliente_det', @Id=" & id)

'[UI-ANT] xform  = "<form id='xform' class='form-oculto' method='post' action='clientes.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
'[UI-ANT] xboton = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('xform').submit();""></form>"

'[UI-ANT] If Not rs.eof Then

	'[UI-ANT] vNombre     = rs(0) & ""
	'[UI-ANT] vRutaFicha  = rs(1) & ""
	'[UI-ANT] vDireccion  = rs(2) & ""
	'[UI-ANT] vServidor   = rs(3) & ""
	'[UI-ANT] vAmbiente   = rs(4) & ""
	'[UI-ANT] vLicencia   = rs(6) & ""
	'[UI-ANT] vDbRqt      = rs(8) & ""

	'[UI-ANT] tiene_cli   = (Trim(vDireccion) <> "" Or Trim(vRutaFicha) <> "")
	'[UI-ANT] tiene_det   = (Trim(vServidor) <> "")

	'[UI-ANT] html = html & cabecera()

	'[UI-ANT] html = html & "<div class='titulo-seccion ui-page-header' style='padding: 0.5rem 2rem;background-color:#e8edff;'>"
	'[UI-ANT] html = html & "<div class='ui-toolbar' style='display:flex;align-items:center;justify-content:space-between;'>"
	'[UI-ANT] html = html & "<div><h2>CLIENTES</h2></div>"
	'[UI-ANT] html = html & "<div class='ui-toolbar__actions'>" & xboton & "</div>"
	'[UI-ANT] html = html & "</div>" ' Fin ui-toolbar
	'[UI-ANT] html = html & "</div>" ' Fin titulo-seccion

	'[UI-ANT] html = html & "<div class='inner ui-main' style='padding: 2rem;'>"
	'[UI-ANT] html = html & "<div class='bloque-datos ui-panel' style='max-width: 800px; margin: 0 auto;'>"
	'[UI-ANT] html = html & "<div class='bloque-header ui-panel__header' style='text-transform: uppercase;'>" & Server.HTMLEncode(vNombre) & "</div>"
	'[UI-ANT] html = html & "<div class='bloque-body ui-panel__body'>"

	'[UI-ANT] ' ============================================================
	'[UI-ANT] ' ZONA 1: DATOS GENERALES (CLIENTES.direccion / ruta_ficha)
	'[UI-ANT] ' - readonly si existen, editable + boton GUARDAR si faltan
	'[UI-ANT] ' ============================================================
	'[UI-ANT] If Trim(vDireccion) <> "" And Trim(vRutaFicha) <> "" Then
		'[UI-ANT] ' Con datos -> readonly
		'[UI-ANT] html = html & "<div class='grid-datos-doble ui-field-grid--dual' style='margin-bottom: 2rem; padding-bottom: 1.5rem; border-bottom: 1px dashed #cbd5e1;'>"
		'[UI-ANT] html = html & "<div class='dato-label'>DIRECCION</div>"
		'[UI-ANT] html = html & "<div class='dato-value' style='color: #334155;'>" & Server.HTMLEncode(vDireccion) & "</div>"
		'[UI-ANT] lnk = "<a href='" & Server.HTMLEncode(vRutaFicha) & "' style='color: blue; text-decoration: none; font-weight: 500;' target='_blank'>" & Server.HTMLEncode(vRutaFicha) & "</a>"
		'[UI-ANT] html = html & "<div class='dato-label'>RUTA FICHA CLIENTE</div>"
		'[UI-ANT] html = html & "<div class='dato-value'>" & lnk & "</div>"
		'[UI-ANT] html = html & "</div>"
	'[UI-ANT] Else
		'[UI-ANT] ' Sin datos -> habilitar campos + boton guardar
		'[UI-ANT] html = html & "<form id='frmCli' method='post' action='clientesDet.asp?opc=UPD_CLI&id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
		'[UI-ANT] html = html & "<div class='grid-datos-doble ui-field-grid--dual' style='margin-bottom: 1rem; padding-bottom: 1.5rem; border-bottom: 1px dashed #cbd5e1;'>"
		'[UI-ANT] html = html & "<div class='dato-label'>DIRECCION</div>"
		'[UI-ANT] html = html & "<div>" & HtmlInputCampo("direccion", vDireccion, True) & "</div>"
		'[UI-ANT] html = html & "<div class='dato-label'>RUTA FICHA CLIENTE</div>"
		'[UI-ANT] html = html & "<div>" & HtmlInputCampo("ruta_ficha", vRutaFicha, True) & "</div>"
		'[UI-ANT] html = html & "</div>"
		'[UI-ANT] html = html & "<div style='display:flex; justify-content:flex-end; gap:1rem; padding:0 0 1.5rem 0; margin-bottom:1.5rem; border-bottom:1px dashed #cbd5e1;'>"
		'[UI-ANT] html = html & "<input class='bot ui-button ui-button--primary' type='submit' value='GUARDAR DATOS CLIENTE' style='background-color:#2563eb; color:white; padding:0.75rem 1.5rem; border:none; border-radius:4px; font-size:1rem; cursor:pointer;'>"
		'[UI-ANT] html = html & "</div>"
		'[UI-ANT] html = html & "</form>"
	'[UI-ANT] End If

	'[UI-ANT] ' ============================================================
	'[UI-ANT] ' ZONA 2: GRILLA ENTORNOS (CLIENTES_DET)
	'[UI-ANT] ' - readonly si existen, editable + boton guardar si no existen
	'[UI-ANT] ' - boton Agregar nuevos entornos siempre visible cuando hay datos
	'[UI-ANT] ' - modo=edit: form unico indexado, JS agregar fila, GUARDAR borra+inserta
	'[UI-ANT] ' ============================================================
	'[UI-ANT] If Not tiene_det Then
		'[UI-ANT] ' Sin entornos -> habilitar campos + boton guardar
		'[UI-ANT] html = html & "<div class='dato-label' style='margin-bottom: 1rem; font-size: 1rem;'>ENTORNOS - AGREGAR PRIMER ENTORNO</div>"
		'[UI-ANT] html = html & "<form id='frmDet' method='post' action='clientesDet.asp?opc=INS_DET&id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
		'[UI-ANT] html = html & "<div class='grid-datos-doble ui-field-grid--dual'>"
		'[UI-ANT] html = html & "<div class='dato-label'>SERVIDOR</div>"
		'[UI-ANT] html = html & "<div>" & HtmlInputCampo("servidor", "", True) & "</div>"
		'[UI-ANT] html = html & "<div class='dato-label'>AMBIENTE</div>"
		'[UI-ANT] html = html & "<div>" & HtmlInputCampo("ambiente", "", True) & "</div>"
		'[UI-ANT] html = html & "<div class='dato-label'>DB RQT</div>"
		'[UI-ANT] html = html & "<div>" & HtmlInputCampo("db_rqt", "", True) & "</div>"
		'[UI-ANT] html = html & "<div class='dato-label'>LICENCIA</div>"
		'[UI-ANT] html = html & "<div>" & HtmlInputCampo("licencia", "", True) & "</div>"
		'[UI-ANT] html = html & "</div>"
		'[UI-ANT] html = html & "<div style='display:flex; justify-content:flex-end; gap:1rem; padding:1rem 0;'>"
		'[UI-ANT] html = html & "<input class='bot ui-button ui-button--primary' type='submit' value='GUARDAR ENTORNO' style='background-color:#10b981; color:white; padding:0.75rem 1.5rem; border:none; border-radius:4px; font-size:1rem; cursor:pointer;'>"
		'[UI-ANT] html = html & "</div>"
		'[UI-ANT] html = html & "</form>"

	'[UI-ANT] ElseIf modo = "edit" Then
		'[UI-ANT] ' MODO EDICION: form unico con todos los entornos indexados
		'[UI-ANT] html = html & "<form id='frmAllDet' method='post' action='clientesDet.asp?opc=UPD_ALL_DET&id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
		'[UI-ANT] html = html & "<div class='dato-label' style='margin-bottom: 1rem; font-size: 1rem;'>EDITAR ENTORNOS</div>"

		'[UI-ANT] numReg = 0
		'[UI-ANT] Do While Not rs.eof
			'[UI-ANT] vDetSrv   = rs(3) & ""
			'[UI-ANT] vDetAmb   = rs(4) & ""
			'[UI-ANT] vDetLic   = rs(6) & ""
			'[UI-ANT] vDetDbRqt = rs(8) & ""
			'[UI-ANT] html = html & "<div class='fila-entorno' style='padding:1rem; margin-bottom:0.75rem; background:#f8fafc; border:1px solid #e2e8f0; border-radius:6px;'>"
			'[UI-ANT] html = html & "<div class='grid-datos-doble ui-field-grid--dual'>"
			'[UI-ANT] html = html & "<div class='dato-label'>SERVIDOR</div>"
			'[UI-ANT] html = html & "<div><input class='input-moderno ui-input' type='text' name='srv_" & numReg & "' value='" & Server.HTMLEncode(vDetSrv) & "' style='width:90%;'></div>"
			'[UI-ANT] html = html & "<div class='dato-label'>AMBIENTE</div>"
			'[UI-ANT] html = html & "<div><input class='input-moderno ui-input' type='text' name='amb_" & numReg & "' value='" & Server.HTMLEncode(vDetAmb) & "' style='width:90%;'></div>"
			'[UI-ANT] html = html & "<div class='dato-label'>DB RQT</div>"
			'[UI-ANT] html = html & "<div><input class='input-moderno ui-input' type='text' name='dbrqt_" & numReg & "' value='" & Server.HTMLEncode(vDetDbRqt) & "' style='width:90%;'></div>"
			'[UI-ANT] html = html & "<div class='dato-label'>LICENCIA</div>"
			'[UI-ANT] html = html & "<div><input class='input-moderno ui-input' type='text' name='lic_" & numReg & "' value='" & Server.HTMLEncode(vDetLic) & "' style='width:90%;'></div>"
			'[UI-ANT] html = html & "</div>"
			'[UI-ANT] html = html & "</div>"
			'[UI-ANT] numReg = numReg + 1
			'[UI-ANT] rs.movenext
		'[UI-ANT] Loop

		'[UI-ANT] html = html & "<input type='hidden' name='total_entornos' id='total_entornos' value='" & (numReg - 1) & "'>"
		'[UI-ANT] html = html & "<div style='display:flex; justify-content:space-between; padding:1rem 0;'>"
		'[UI-ANT] html = html & "<input class='bot ui-button ui-button--primary' type='button' value='+ AGREGAR ENTORNO' onclick='agregarFila()' style='background-color:#10b981; color:white; padding:0.75rem 1.5rem; border:none; border-radius:4px; font-size:1rem; cursor:pointer;'>"
		'[UI-ANT] html = html & "</div>"
		'[UI-ANT] html = html & "<div id='contenedor-nuevos'></div>"
		'[UI-ANT] html = html & "<div style='display:flex; justify-content:flex-end; gap:1rem; padding:1rem 0;'>"
		'[UI-ANT] html = html & "<input class='bot ui-button ui-button--primary' type='submit' value='GUARDAR' style='background-color:#2563eb; color:white; padding:0.75rem 1.5rem; border:none; border-radius:4px; font-size:1rem; cursor:pointer;'>"
		'[UI-ANT] html = html & "<a class='ui-button' href='clientesDet.asp?id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "' style='background-color:#6b7280; color:white; padding:0.75rem 1.5rem; border:none; border-radius:4px; font-size:1rem; text-decoration:none;'>CANCELAR</a>"
		'[UI-ANT] html = html & "</div>"
		'[UI-ANT] html = html & "</form>"

		'[UI-ANT] html = html & "<script>"
		'[UI-ANT] html = html & "var contadorEntornos = " & numReg & ";"
		'[UI-ANT] html = html & "function agregarFila() {"
		'[UI-ANT] html = html & "  var c = document.getElementById('contenedor-nuevos');"
		'[UI-ANT] html = html & "  var div = document.createElement('div');"
		'[UI-ANT] html = html & "  div.className = 'fila-entorno';"
		'[UI-ANT] html = html & "  div.style.cssText = 'padding:1rem; margin-bottom:0.75rem; background:#fff; border:1px dashed #10b981; border-radius:6px;';"
		'[UI-ANT] html = html & "  div.innerHTML = "
		'[UI-ANT] html = html & "    '<div class=""grid-datos-doble ui-field-grid--dual"">' +"
		'[UI-ANT] html = html & "    '<div class=""dato-label"">SERVIDOR</div>' +"
		'[UI-ANT] html = html & "    '<div><input class=""input-moderno ui-input"" type=""text"" name=""srv_' + contadorEntornos + '"" style=""width:90%;""></div>' +"
		'[UI-ANT] html = html & "    '<div class=""dato-label"">AMBIENTE</div>' +"
		'[UI-ANT] html = html & "    '<div><input class=""input-moderno ui-input"" type=""text"" name=""amb_' + contadorEntornos + '"" style=""width:90%;""></div>' +"
		'[UI-ANT] html = html & "    '<div class=""dato-label"">DB RQT</div>' +"
		'[UI-ANT] html = html & "    '<div><input class=""input-moderno ui-input"" type=""text"" name=""dbrqt_' + contadorEntornos + '"" style=""width:90%;""></div>' +"
		'[UI-ANT] html = html & "    '<div class=""dato-label"">LICENCIA</div>' +"
		'[UI-ANT] html = html & "    '<div><input class=""input-moderno ui-input"" type=""text"" name=""lic_' + contadorEntornos + '"" style=""width:90%;""></div>' +"
		'[UI-ANT] html = html & "    '</div>';"
		'[UI-ANT] html = html & "  c.appendChild(div);"
		'[UI-ANT] html = html & "  document.getElementById('total_entornos').value = contadorEntornos;"
		'[UI-ANT] html = html & "  contadorEntornos++;"
		'[UI-ANT] html = html & "}"
		'[UI-ANT] html = html & "</script>"

	'[UI-ANT] ElseIf modo = "add" Then
		'[UI-ANT] html = html & "<div class='dato-label' style='margin-bottom: 1rem; font-size: 1rem;'>AGREGAR NUEVO ENTORNO</div>"
		'[UI-ANT] html = html & "<form id='frmDet' method='post' action='clientesDet.asp?opc=INS_DET&id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
		'[UI-ANT] html = html & "<div class='grid-datos-doble ui-field-grid--dual'>"
		'[UI-ANT] html = html & "<div class='dato-label'>SERVIDOR</div><div>" & HtmlInputCampo("servidor", "", True) & "</div>"
		'[UI-ANT] html = html & "<div class='dato-label'>AMBIENTE</div><div>" & HtmlInputCampo("ambiente", "", True) & "</div>"
		'[UI-ANT] html = html & "<div class='dato-label'>DB RQT</div><div>" & HtmlInputCampo("db_rqt", "", True) & "</div>"
		'[UI-ANT] html = html & "<div class='dato-label'>LICENCIA</div><div>" & HtmlInputCampo("licencia", "", True) & "</div>"
		'[UI-ANT] html = html & "</div>"
		'[UI-ANT] html = html & "<div style='display:flex; justify-content:flex-end; gap:1rem; padding:1rem 0;'>"
		'[UI-ANT] html = html & "<input class='bot ui-button ui-button--primary' type='submit' value='GUARDAR ENTORNO' style='background-color:#10b981; color:white; padding:0.75rem 1.5rem; border:none; border-radius:4px; font-size:1rem; cursor:pointer;'>"
		'[UI-ANT] html = html & "<a class='ui-button' href='clientesDet.asp?id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "' style='background-color:#6b7280; color:white; padding:0.75rem 1.5rem; border:none; border-radius:4px; font-size:1rem; text-decoration:none;'>CANCELAR</a>"
		'[UI-ANT] html = html & "</div>"
		'[UI-ANT] html = html & "</form>"

	'[UI-ANT] Else
		'[UI-ANT] ' MODO LECTURA: grilla readonly + botones EDITAR / AGREGAR
		'[UI-ANT] html = html & "<div class='dato-label' style='margin-bottom: 1rem; font-size: 1rem;'>ENTORNOS ASIGNADOS</div>"
		'[UI-ANT] html = html & "<div style='display:flex; flex-direction:column; gap:0.75rem;'>"
		'[UI-ANT] Do While Not rs.eof
			'[UI-ANT] vDetSrv   = rs(3) & ""
			'[UI-ANT] vDetAmb   = rs(4) & ""
			'[UI-ANT] vDetLic   = rs(6) & ""
			'[UI-ANT] vDetDbRqt = rs(8) & ""
			'[UI-ANT] xtipo = ""
			'[UI-ANT] If vDetAmb = "T" Then
				'[UI-ANT] xtipo = "Test"
			'[UI-ANT] ElseIf vDetAmb = "P" Then
				'[UI-ANT] xtipo = "Prod"
			'[UI-ANT] Else
				'[UI-ANT] xtipo = vDetAmb
			'[UI-ANT] End If
			'[UI-ANT] html = html & "<div style='display:flex; flex-wrap:wrap; justify-content:space-between; align-items:center; padding:1rem; background-color:#f8fafc; border:1px solid #e2e8f0; border-radius:6px; color:#475569; font-size:0.9rem;'>"
			'[UI-ANT] html = html & "<div style='flex:1 1 25%; min-width:150px;'><strong style='color:#1e293b;'>Srv:</strong> " & Server.HTMLEncode(vDetSrv) & " <span style='background-color:#e0e7ff; color:#4338ca; padding:0.1rem 0.4rem; border-radius:4px; font-size:0.8rem; margin-left:0.3rem; font-weight:bold;'>" & Server.HTMLEncode(xtipo) & "</span></div>"
			'[UI-ANT] html = html & "<div style='flex:1 1 25%; min-width:150px;'><strong style='color:#1e293b;'>DB RQT:</strong> " & Server.HTMLEncode(vDetDbRqt) & "</div>"
			'[UI-ANT] html = html & "<div style='flex:1 1 25%; min-width:150px;'><strong style='color:#1e293b;'>Lic:</strong> " & Server.HTMLEncode(vDetLic) & "</div>"
			'[UI-ANT] html = html & "</div>"
			'[UI-ANT] rs.movenext
		'[UI-ANT] Loop
		'[UI-ANT] html = html & "</div>"
		'[UI-ANT] html = html & "<div style='display:flex; justify-content:flex-end; gap:1rem; padding:1rem 0;'>"
		'[UI-ANT] If getPermisos(xmenu,xtmenu,"MODIF") = 1 Then
			'[UI-ANT] html = html & "<a class='ui-button ui-button--primary' href='clientesDet.asp?id=" & id & "&modo=edit&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "' style='background-color:#2563eb; color:white; padding:0.75rem 1.5rem; border:none; border-radius:4px; font-size:1rem; text-decoration:none;'>EDITAR</a>"
		'[UI-ANT] End If
		'[UI-ANT] html = html & "<a class='ui-button ui-button--primary' href='clientesDet.asp?id=" & id & "&modo=add&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "' style='background-color:#10b981; color:white; padding:0.75rem 1.5rem; border:none; border-radius:4px; font-size:1rem; text-decoration:none;'>AGREGAR NUEVO</a>"
		'[UI-ANT] html = html & "</div>"
	'[UI-ANT] End If

	'[UI-ANT] html = html & "</div>"
	'[UI-ANT] html = html & "</div>"
	'[UI-ANT] html = html & "</div>"

	'[UI-ANT] rs.close
'[UI-ANT] Else

	'[UI-ANT] html = html & cabecera()
	'[UI-ANT] html = html & "<div class='inner ui-main' style='padding: 2rem;'>"
	'[UI-ANT] html = html & "<div class='bloque-datos ui-panel' style='max-width: 800px; margin: 0 auto;'>"
	'[UI-ANT] html = html & "<div class='bloque-header ui-panel__header' style='text-transform: uppercase;'>AGREGAR CLIENTE</div>"
	'[UI-ANT] html = html & "<div class='bloque-body ui-panel__body'>"
	'[UI-ANT] html = html & "<div class='grid-datos-doble ui-field-grid--dual' style='margin-bottom: 2rem; padding-bottom: 1.5rem; border-bottom: 1px dashed #cbd5e1;'>"
	'[UI-ANT] html = html & "<div class='dato-label'>NOMBRE</div><div><input class='input-moderno ui-input' id='nombre' name='nombre' type='text' required></div>"
	'[UI-ANT] html = html & "<div class='dato-label'>DIRECCION</div><div><input class='input-moderno ui-input' id='direccion' name='direccion' type='text'></div>"
	'[UI-ANT] html = html & "<div class='dato-label'>RUTA FICHA CLIENTE</div><div><input class='input-moderno ui-input' id='ruta_ficha' name='ruta_ficha' type='text'></div>"
	'[UI-ANT] html = html & "</div>"
	'[UI-ANT] html = html & "<div style='text-align: center; margin: 2rem 0;'>"
	'[UI-ANT] html = html & "<input class='bot ui-button ui-button--primary' type='submit' value='GUARDAR CLIENTE' style='background-color:#2563eb; color:white; padding:0.75rem 1.5rem; border:none; border-radius:4px; font-size:1rem; cursor:pointer;'>"
	'[UI-ANT] html = html & "</div>"
	'[UI-ANT] html = html & "<div style='text-align: center; margin-top: 1rem;'>" & xboton & "</div>"
	'[UI-ANT] html = html & "</div>" ' Fin .bloque-body
	'[UI-ANT] html = html & "</div>" ' Fin .bloque-datos
	'[UI-ANT] html = html & "</div>" ' Fin .inner
'[UI-ANT] End If

'[UI-ANT] html = html & "<script src='assets/js/jquery.min.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/skel.min.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/util.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/main.js'></script>"

'[UI-ANT] html = html & xform & "</body>"
'[UI-ANT] html = html & "</html>"

'[UI-ANT] response.write html

' =========================================================================
' [UI] PANTALLA con el sistema visual (IncFile/ui_componentes.inc)
' Ficha del cliente: datos (direccion y ficha) y entornos. Se llega desde Clientes > Ver ficha.
' El cuerpo anterior completo queda arriba, comentado con '[UI-ANT].
' Consultas y procedimientos: los mismos de antes. Cambios: id validado, sin el SQL que se
' mostraba en pantalla, permisos chequeados en el servidor y datos del cliente siempre editables.
' =========================================================================
Dim puedeModif, puedeAlta, accTop, editarDatos, tablaEnt, cantEnt, fila, i, sinPermiso

html = headhtml()
html = html & "<body>"

id      = request.querystring("id")
modo    = request.querystring("modo")
opc     = request.querystring("opc")
xmenu   = request.querystring("xmenu")
xtmenu  = request.querystring("xtmenu")

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
XXX = SYSLOGON(cnn,"PARAM_CLIEN")
If Err.Number <> 0 Then
	Response.Write "ERROR SYSLOGON: " & Err.Number & ": " & Err.Description
	Response.End
End If
On Error Goto 0

' [UI] El id va a la consulta SQL: se valida como numero. Sin un id valido no se opera sobre nada
' (antes, sin id tomaba rs(0) de la consulta como id, pero esa columna es el NOMBRE: nunca funciono)
If IsNumeric(id & "") And Trim(id & "") <> "" Then id = CLng(id) Else id = 0
If id = 0 Then opc = ""
puedeModif = (getPermisosAtrm("CG_CLIENTES", "MODIF") = 1)
' [LOG] Consulta del detalle de un cliente (CG_LOG)
If id > 0 Then CgLog "CLIENTE", id, "CONSULTA"
puedeAlta = (getPermisosAtrm("CG_CLIENTES", "ALTA") = 1)
sinPermiso = "<script>alert('No ten\u00e9s permiso para esta acci\u00f3n.'); window.history.back();</script>"

If opc = "UPD_CLI" Then

	direccion   = Trim(request.form("direccion") & "")
	ruta_ficha  = Trim(request.form("ruta_ficha") & "")

	If Not puedeModif Then
		html = html & sinPermiso
	ElseIf direccion = "" Then
		html = html & "<script>alert('La direcci\u00f3n es obligatoria.'); window.history.back();</script>"
	Else
		sSQL = "sp_cg_Catalogo_Actualizar @Tabla='clientes_ficha', @RutaFicha='" & Replace(ruta_ficha,"'","''") & "',@Direccion='" & Replace(direccion,"'","''") &"', @Id=" & id
		' [UI] Se quito el response.write sSQL que mostraba la consulta en pantalla
		On Error Resume Next
		ExecNonQuery Cnn, sSQL
		If Err.Number <> 0 Then
			Response.Write "ERROR actualizar ruta_ficha: " & Err.Number & ": " & Err.Description
			Response.End
		End If
		On Error Goto 0
		CgLog "CLIENTE", id, "ACTUALIZA"
		html = html & "<script>alert('Datos del cliente actualizados.'); window.location.href='clientesDet.asp?id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "';</script>"
	End If

ElseIf opc = "UPD_ALL_DET" Then

	If Not puedeModif Then
		html = html & sinPermiso
	Else
		totalEntornos = Trim(request.form("total_entornos") & "")
		If Not IsNumeric(totalEntornos) Then totalEntornos = "-1"

		EliminarClientesDet Cnn, id

		i = 0
		Do While i <= CInt(totalEntornos)
			srv   = Trim(request.form("srv_" & i) & "")
			amb   = Trim(request.form("amb_" & i) & "")
			dbrqt = Trim(request.form("dbrqt_" & i) & "")
			lic   = Trim(request.form("lic_" & i) & "")

			If srv <> "" And amb <> "" Then
				srv   = Replace(srv, "'", "''")
				amb   = Replace(amb, "'", "''")
				dbrqt = Replace(dbrqt, "'", "''")
				lic   = Replace(lic, "'", "''")
				InsertarClientesDet Cnn, id, srv, amb, dbrqt, lic
			End If
			i = i + 1
		Loop

		CgLog "CLIENTE", id, "ACTUALIZA"
		html = html & "<script>alert('Entornos actualizados.'); window.location.href='clientesDet.asp?id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "';</script>"
	End If

ElseIf opc = "INS_DET" Then

	servidor  = Trim(request.form("servidor") & "")
	ambiente  = Trim(request.form("ambiente") & "")
	db_rqt    = Trim(request.form("db_rqt") & "")
	licencia  = Trim(request.form("licencia") & "")

	If Not puedeAlta Then
		html = html & sinPermiso
	ElseIf servidor = "" Or ambiente = "" Then
		html = html & "<script>alert('Servidor y ambiente son obligatorios.'); window.history.back();</script>"
	Else
		servidor  = Replace(servidor,"'","''")
		ambiente  = Replace(ambiente,"'","''")
		db_rqt    = Replace(db_rqt,"'","''")
		licencia  = Replace(licencia,"'","''")
		InsertarClientesDet Cnn, id, servidor, ambiente, db_rqt, licencia
		CgLog "CLIENTE", id, "ACTUALIZA"
		html = html & "<script>alert('Entorno agregado.'); window.location.href='clientesDet.asp?id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "';</script>"
	End If

End If

Set rs = Exec(Cnn, "sp_cg_Utilidades @Modo='consultar_cliente_det', @Id=" & id)

xform  = "<form id='xform' class='form-oculto' method='post' action='clientes.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
accTop = UiBoton("Volver", "secondary", "document.getElementById('xform').submit();", True, "")

html = html & "<div id='wrapper' role='banner'>" & cabecera() & "</div>"

If Not rs.eof Then

	vNombre     = rs(0) & ""
	vRutaFicha  = rs(1) & ""
	vDireccion  = rs(2) & ""
	vServidor   = rs(3) & ""
	tiene_det   = (Trim(vServidor) <> "")

	html = html & UiPantallaParametrosAbre(UiTexto(vNombre), accTop, "")

	' ===================== Datos del cliente =====================
	' Se editan si falta alguno de los dos (como antes) o si se pide con Editar (antes no se podian cambiar)
	editarDatos = (Trim(vDireccion) = "" Or Trim(vRutaFicha) = "" Or modo = "datos")
	If editarDatos And puedeModif Then
		html = html & "<form id='frmCli' method='post' action='clientesDet.asp?opc=UPD_CLI&id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
		html = html & UiCardAbre("Datos del cliente", "") & UiFormGridAbre()
		html = html & UiCampo("direccion", "Direcci&oacute;n", UiInputTexto("direccion", True, "value='" & Server.HTMLEncode(vDireccion) & "'"), True, True)
		html = html & UiCampo("ruta_ficha", "Ruta de la ficha del cliente", UiInputTexto("ruta_ficha", False, "value='" & Server.HTMLEncode(vRutaFicha) & "' placeholder='http://...'"), False, True)
		html = html & UiFormGridCierra()
		html = html & "<div class='ui-card__acciones'>"
		If Trim(vDireccion) <> "" And Trim(vRutaFicha) <> "" Then
			html = html & "<a class='ui-button ui-button--secondary' href='clientesDet.asp?id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>Cancelar</a>"
		End If
		html = html & UiBotonEnviar("Guardar datos", True, "") & "</div>"
		html = html & UiCardCierra() & "</form>"
	Else
		html = html & UiCardAbre("Datos del cliente", UiBotonEnlaceSi("Editar", "clientesDet.asp?id=" & id & "&modo=datos&xmenu=" & xmenu & "&xtmenu=" & xtmenu, puedeModif, "Sin permiso para modificar."))
		html = html & "<div class='ui-fields'>"
		html = html & UiCampoLectura("Direcci&oacute;n", UiDato(vDireccion), False)
		If Trim(vRutaFicha) = "" Then
			html = html & UiCampoLectura("Ficha del cliente", "<span class='ui-status-empty'>&mdash;</span>", False)
		Else
			html = html & UiCampoLectura("Ficha del cliente", "<a class='ui-link' href='" & Server.HTMLEncode(vRutaFicha) & "' target='_blank'>Abrir ficha</a>", False)
		End If
		html = html & "</div>" & UiCardCierra()
	End If

	' ===================== Entornos =====================
	If Not tiene_det Or modo = "add" Then
		' Alta de entorno (el primero, o uno nuevo con Agregar)
		If puedeAlta Then
			html = html & "<form id='frmDet' method='post' action='clientesDet.asp?opc=INS_DET&id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
			If tiene_det Then
				html = html & UiCardAbre("Nuevo entorno", "")
			Else
				html = html & UiCardAbre("Entornos: agregar el primero", "")
			End If
			html = html & UiFormGridAbre()
			html = html & UiCampo("servidor", "Servidor", UiInputTexto("servidor", True, "autocomplete='off'"), True, False)
			html = html & UiCampo("ambiente", "Ambiente", UiInputTexto("ambiente", True, "autocomplete='off' placeholder='T (test) o P (producci&oacute;n)'"), True, False)
			html = html & UiCampo("db_rqt", "Base RQT", UiInputTexto("db_rqt", False, "autocomplete='off'"), False, False)
			html = html & UiCampo("licencia", "Licencia", UiInputTexto("licencia", False, "autocomplete='off'"), False, False)
			html = html & UiFormGridCierra()
			html = html & "<div class='ui-card__acciones'>"
			If tiene_det Then html = html & "<a class='ui-button ui-button--secondary' href='clientesDet.asp?id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>Cancelar</a>"
			html = html & UiBotonEnviar("Guardar entorno", True, "") & "</div>"
			html = html & UiCardCierra() & "</form>"
		Else
			html = html & UiCardAbre("Entornos", "") & "<div class='ui-empty'>El cliente no tiene entornos cargados.</div>" & UiCardCierra()
		End If

	ElseIf modo = "edit" And puedeModif Then
		' Edicion de todos los entornos: se guardan juntos (mismo procedimiento de antes)
		html = html & "<form id='frmAllDet' method='post' action='clientesDet.asp?opc=UPD_ALL_DET&id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
		html = html & UiCardAbre("Editar entornos", "")
		html = html & UiTablaAbre("Entornos del cliente, editables")
		html = html & "<thead><tr>" & UiTh("Servidor", "") & UiTh("Ambiente", "") & UiTh("Base RQT", "") & UiTh("Licencia", "") & "</tr></thead><tbody id='filasEntornos'>"
		cantEnt = 0
		Do While Not rs.eof
			html = html & "<tr>" & _
				"<td><input class='ui-input' type='text' name='srv_" & cantEnt & "' value='" & Server.HTMLEncode(rs(3) & "") & "' aria-label='Servidor del entorno " & (cantEnt + 1) & "'></td>" & _
				"<td><input class='ui-input' type='text' name='amb_" & cantEnt & "' value='" & Server.HTMLEncode(rs(4) & "") & "' aria-label='Ambiente del entorno " & (cantEnt + 1) & "'></td>" & _
				"<td><input class='ui-input' type='text' name='dbrqt_" & cantEnt & "' value='" & Server.HTMLEncode(rs(8) & "") & "' aria-label='Base RQT del entorno " & (cantEnt + 1) & "'></td>" & _
				"<td><input class='ui-input' type='text' name='lic_" & cantEnt & "' value='" & Server.HTMLEncode(rs(6) & "") & "' aria-label='Licencia del entorno " & (cantEnt + 1) & "'></td>" & _
				"</tr>"
			cantEnt = cantEnt + 1
			rs.movenext
		Loop
		html = html & UiTablaCierra()
		html = html & "<input type='hidden' name='total_entornos' id='total_entornos' value='" & (cantEnt - 1) & "'>"
		html = html & "<p class='ui-table__nota'>Para quitar un entorno, dej&aacute; vac&iacute;os el servidor y el ambiente.</p>"
		html = html & "<div class='ui-card__acciones'>" & UiBoton("+ Agregar fila", "secondary", "agregarFila();", True, "") & _
			"<a class='ui-button ui-button--secondary' href='clientesDet.asp?id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>Cancelar</a>" & _
			UiBotonEnviar("Guardar entornos", True, "") & "</div>"
		html = html & UiCardCierra() & "</form>"

		' Fila nueva con los mismos nombres de campo (srv_N, amb_N, dbrqt_N, lic_N)
		html = html & "<script>"
		html = html & "var contadorEntornos = " & cantEnt & ";"
		html = html & "function agregarFila() {"
		html = html & "  var n = contadorEntornos, num = n + 1, tr = document.createElement('tr');"
		html = html & "  var campo = function (pref, etiqueta) { return '<td><input class=""ui-input"" type=""text"" name=""' + pref + n + '"" aria-label=""' + etiqueta + ' del entorno ' + num + '""></td>'; };"
		html = html & "  tr.innerHTML = campo('srv_', 'Servidor') + campo('amb_', 'Ambiente') + campo('dbrqt_', 'Base RQT') + campo('lic_', 'Licencia');"
		html = html & "  document.getElementById('filasEntornos').appendChild(tr);"
		html = html & "  document.getElementById('total_entornos').value = n;"
		html = html & "  contadorEntornos++;"
		html = html & "  tr.querySelector('input').focus();"
		html = html & "}"
		html = html & "</script>"

	Else
		' Entornos asignados (solo lectura)
		html = html & UiCardAbre("Entornos", "<div class='ui-card__header-acciones'>" & _
			UiBotonEnlaceSi("Editar", "clientesDet.asp?id=" & id & "&modo=edit&xmenu=" & xmenu & "&xtmenu=" & xtmenu, puedeModif, "Sin permiso para modificar.") & _
			UiBotonEnlaceSi("Agregar", "clientesDet.asp?id=" & id & "&modo=add&xmenu=" & xmenu & "&xtmenu=" & xtmenu, puedeAlta, "Sin permiso para agregar.") & "</div>")
		html = html & UiTablaAbre("Entornos del cliente")
		html = html & "<thead><tr>" & UiTh("Servidor", "") & UiTh("Ambiente", "") & UiTh("Base RQT", "") & UiTh("Licencia", "") & "</tr></thead><tbody>"
		cantEnt = 0
		Do While Not rs.eof
			vDetAmb = rs(4) & ""
			If vDetAmb = "T" Then
				fila = UiBadge("Test", "purple")
			ElseIf vDetAmb = "P" Then
				fila = UiBadge("Producci&oacute;n", "green")
			Else
				fila = UiDato(vDetAmb)
			End If
			html = html & "<tr><td><span class='ui-table__titulo'>" & UiTexto(rs(3)) & "</span></td><td>" & fila & "</td><td>" & UiDato(rs(8)) & "</td><td>" & UiDato(rs(6)) & "</td></tr>"
			cantEnt = cantEnt + 1
			rs.movenext
		Loop
		html = html & UiTablaCierra()
		If cantEnt = 1 Then
			html = html & UiPie("1 entorno", "")
		Else
			html = html & UiPie(cantEnt & " entornos", "")
		End If
		html = html & UiCardCierra()
	End If

	html = html & UiPantallaCierra()
	rs.close
Else
	' [UI] Cliente inexistente: antes mostraba un formulario de alta sin <form> (no guardaba nada)
	html = html & UiPantallaParametrosAbre("Ficha del cliente", accTop, "narrow")
	html = html & UiCardAbre("Cliente no encontrado", "") & "<div class='ui-empty'>No se encontr&oacute; el cliente. Los clientes nuevos se cargan desde Clientes &gt; Agregar cliente.</div>" & UiCardCierra()
	html = html & UiPantallaCierra()
End If

html = html & "<script src='assets/js/jquery.min.js'></script>"
html = html & "<script src='assets/js/skel.min.js'></script>"
html = html & "<script src='assets/js/util.js'></script>"
html = html & "<script src='assets/js/main.js'></script>"

html = html & xform & "</body>"
html = html & "</html>"

response.write html
%>