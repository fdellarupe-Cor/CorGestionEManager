<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/datosActualizar.inc" -->

<%
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

' --- PROCESAMIENTO DE GUARDADO ---
If opc = "UPD_CLI" Then

	direccion   = Trim(request.form("direccion") & "")
	ruta_ficha  = Trim(request.form("ruta_ficha") & "")

	If direccion = "" Then
		html = html & "<script>alert('La direccion es obligatoria.'); window.history.back();</script>"
	Else
		' sSQL = "sp_cg_Catalogo_Actualizar @Tabla='clientes', @Campo='direccion', @Valor='" & Replace(direccion,"'","''") & "', @Id=" & id
		' On Error Resume Next
		' Cnn.Execute(sSQL)
		' If Err.Number <> 0 Then
			' Response.Write "ERROR actualizar direccion: " & Err.Number & ": " & Err.Description
			' Response.End
		' End If
		' On Error Goto 0

		sSQL = "sp_cg_Catalogo_Actualizar @Tabla='clientes_ficha', @RutaFicha='" & Replace(ruta_ficha,"'","''") & "',@Direccion='" & Replace(direccion,"'","''") &"', @Id=" & id
		response.write sSQL
		On Error Resume Next
		ExecNonQuery Cnn, sSQL
		If Err.Number <> 0 Then
			Response.Write "ERROR actualizar ruta_ficha: " & Err.Number & ": " & Err.Description
			Response.End
		End If
		On Error Goto 0

		html = html & "<script>alert('Datos del cliente actualizados.'); window.location.href='clientesDet.asp?id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "';</script>"
	End If

ElseIf opc = "UPD_ALL_DET" Then

	totalEntornos = Trim(request.form("total_entornos") & "")
	If totalEntornos = "" Then totalEntornos = "-1"

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

	html = html & "<script>alert('Entornos actualizados.'); window.location.href='clientesDet.asp?id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "';</script>"

ElseIf opc = "INS_DET" Then

	servidor  = Trim(request.form("servidor") & "")
	ambiente  = Trim(request.form("ambiente") & "")
	db_rqt    = Trim(request.form("db_rqt") & "")
	licencia  = Trim(request.form("licencia") & "")

	If servidor = "" Or ambiente = "" Then
		html = html & "<script>alert('Servidor y Ambiente son obligatorios.'); window.history.back();</script>"
	Else
		servidor  = Replace(servidor,"'","''")
		ambiente  = Replace(ambiente,"'","''")
		db_rqt    = Replace(db_rqt,"'","''")
		licencia  = Replace(licencia,"'","''")
		InsertarClientesDet Cnn, id, servidor, ambiente, db_rqt, licencia
		html = html & "<script>alert('Entorno agregado.'); window.location.href='clientesDet.asp?id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "';</script>"
	End If

End If

' --- LECTURA DE DATOS ---
If id & "" = "" Then
	Set rs = Exec(Cnn, "sp_cg_Utilidades @Modo='consultar_cliente_det', @Id=0")
	If Not rs.eof Then id = rs(0)
	rs.close
End If

Set rs = Exec(Cnn, "sp_cg_Utilidades @Modo='consultar_cliente_det', @Id=" & id)

xform  = "<form id='xform' class='form-oculto' method='post' action='clientes.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
xboton = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('xform').submit();""></form>"

If Not rs.eof Then

	vNombre     = rs(0) & ""
	vRutaFicha  = rs(1) & ""
	vDireccion  = rs(2) & ""
	vServidor   = rs(3) & ""
	vAmbiente   = rs(4) & ""
	vLicencia   = rs(6) & ""
	vDbRqt      = rs(8) & ""

	tiene_cli   = (Trim(vDireccion) <> "" Or Trim(vRutaFicha) <> "")
	tiene_det   = (Trim(vServidor) <> "")

	html = html & cabecera()

	html = html & "<div class='titulo-seccion ui-page-header' style='padding: 0.5rem 2rem;background-color:#e8edff;'>"
	html = html & "<div class='ui-toolbar' style='display:flex;align-items:center;justify-content:space-between;'>"
	html = html & "<div><h2>CLIENTES</h2></div>"
	html = html & "<div class='ui-toolbar__actions'>" & xboton & "</div>"
	html = html & "</div>" ' Fin ui-toolbar
	html = html & "</div>" ' Fin titulo-seccion

	html = html & "<div class='inner ui-main' style='padding: 2rem;'>"
	html = html & "<div class='bloque-datos ui-panel' style='max-width: 800px; margin: 0 auto;'>"
	html = html & "<div class='bloque-header ui-panel__header' style='text-transform: uppercase;'>" & Server.HTMLEncode(vNombre) & "</div>"
	html = html & "<div class='bloque-body ui-panel__body'>"

	' ============================================================
	' ZONA 1: DATOS GENERALES (CLIENTES.direccion / ruta_ficha)
	' - readonly si existen, editable + boton GUARDAR si faltan
	' ============================================================
	If Trim(vDireccion) <> "" And Trim(vRutaFicha) <> "" Then
		' Con datos -> readonly
		html = html & "<div class='grid-datos-doble ui-field-grid--dual' style='margin-bottom: 2rem; padding-bottom: 1.5rem; border-bottom: 1px dashed #cbd5e1;'>"
		html = html & "<div class='dato-label'>DIRECCION</div>"
		html = html & "<div class='dato-value' style='color: #334155;'>" & Server.HTMLEncode(vDireccion) & "</div>"
		lnk = "<a href='" & Server.HTMLEncode(vRutaFicha) & "' style='color: blue; text-decoration: none; font-weight: 500;' target='_blank'>" & Server.HTMLEncode(vRutaFicha) & "</a>"
		html = html & "<div class='dato-label'>RUTA FICHA CLIENTE</div>"
		html = html & "<div class='dato-value'>" & lnk & "</div>"
		html = html & "</div>"
	Else
		' Sin datos -> habilitar campos + boton guardar
		html = html & "<form id='frmCli' method='post' action='clientesDet.asp?opc=UPD_CLI&id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
		html = html & "<div class='grid-datos-doble ui-field-grid--dual' style='margin-bottom: 1rem; padding-bottom: 1.5rem; border-bottom: 1px dashed #cbd5e1;'>"
		html = html & "<div class='dato-label'>DIRECCION</div>"
		html = html & "<div>" & HtmlInputCampo("direccion", vDireccion, True) & "</div>"
		html = html & "<div class='dato-label'>RUTA FICHA CLIENTE</div>"
		html = html & "<div>" & HtmlInputCampo("ruta_ficha", vRutaFicha, True) & "</div>"
		html = html & "</div>"
		html = html & "<div style='display:flex; justify-content:flex-end; gap:1rem; padding:0 0 1.5rem 0; margin-bottom:1.5rem; border-bottom:1px dashed #cbd5e1;'>"
		html = html & "<input class='bot ui-button ui-button--primary' type='submit' value='GUARDAR DATOS CLIENTE' style='background-color:#2563eb; color:white; padding:0.75rem 1.5rem; border:none; border-radius:4px; font-size:1rem; cursor:pointer;'>"
		html = html & "</div>"
		html = html & "</form>"
	End If

	' ============================================================
	' ZONA 2: GRILLA ENTORNOS (CLIENTES_DET)
	' - readonly si existen, editable + boton guardar si no existen
	' - boton Agregar nuevos entornos siempre visible cuando hay datos
	' - modo=edit: form unico indexado, JS agregar fila, GUARDAR borra+inserta
	' ============================================================
	If Not tiene_det Then
		' Sin entornos -> habilitar campos + boton guardar
		html = html & "<div class='dato-label' style='margin-bottom: 1rem; font-size: 1rem;'>ENTORNOS - AGREGAR PRIMER ENTORNO</div>"
		html = html & "<form id='frmDet' method='post' action='clientesDet.asp?opc=INS_DET&id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
		html = html & "<div class='grid-datos-doble ui-field-grid--dual'>"
		html = html & "<div class='dato-label'>SERVIDOR</div>"
		html = html & "<div>" & HtmlInputCampo("servidor", "", True) & "</div>"
		html = html & "<div class='dato-label'>AMBIENTE</div>"
		html = html & "<div>" & HtmlInputCampo("ambiente", "", True) & "</div>"
		html = html & "<div class='dato-label'>DB RQT</div>"
		html = html & "<div>" & HtmlInputCampo("db_rqt", "", True) & "</div>"
		html = html & "<div class='dato-label'>LICENCIA</div>"
		html = html & "<div>" & HtmlInputCampo("licencia", "", True) & "</div>"
		html = html & "</div>"
		html = html & "<div style='display:flex; justify-content:flex-end; gap:1rem; padding:1rem 0;'>"
		html = html & "<input class='bot ui-button ui-button--primary' type='submit' value='GUARDAR ENTORNO' style='background-color:#10b981; color:white; padding:0.75rem 1.5rem; border:none; border-radius:4px; font-size:1rem; cursor:pointer;'>"
		html = html & "</div>"
		html = html & "</form>"

	ElseIf modo = "edit" Then
		' MODO EDICION: form unico con todos los entornos indexados
		html = html & "<form id='frmAllDet' method='post' action='clientesDet.asp?opc=UPD_ALL_DET&id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
		html = html & "<div class='dato-label' style='margin-bottom: 1rem; font-size: 1rem;'>EDITAR ENTORNOS</div>"

		numReg = 0
		Do While Not rs.eof
			vDetSrv   = rs(3) & ""
			vDetAmb   = rs(4) & ""
			vDetLic   = rs(6) & ""
			vDetDbRqt = rs(8) & ""
			html = html & "<div class='fila-entorno' style='padding:1rem; margin-bottom:0.75rem; background:#f8fafc; border:1px solid #e2e8f0; border-radius:6px;'>"
			html = html & "<div class='grid-datos-doble ui-field-grid--dual'>"
			html = html & "<div class='dato-label'>SERVIDOR</div>"
			html = html & "<div><input class='input-moderno ui-input' type='text' name='srv_" & numReg & "' value='" & Server.HTMLEncode(vDetSrv) & "' style='width:90%;'></div>"
			html = html & "<div class='dato-label'>AMBIENTE</div>"
			html = html & "<div><input class='input-moderno ui-input' type='text' name='amb_" & numReg & "' value='" & Server.HTMLEncode(vDetAmb) & "' style='width:90%;'></div>"
			html = html & "<div class='dato-label'>DB RQT</div>"
			html = html & "<div><input class='input-moderno ui-input' type='text' name='dbrqt_" & numReg & "' value='" & Server.HTMLEncode(vDetDbRqt) & "' style='width:90%;'></div>"
			html = html & "<div class='dato-label'>LICENCIA</div>"
			html = html & "<div><input class='input-moderno ui-input' type='text' name='lic_" & numReg & "' value='" & Server.HTMLEncode(vDetLic) & "' style='width:90%;'></div>"
			html = html & "</div>"
			html = html & "</div>"
			numReg = numReg + 1
			rs.movenext
		Loop

		html = html & "<input type='hidden' name='total_entornos' id='total_entornos' value='" & (numReg - 1) & "'>"
		html = html & "<div style='display:flex; justify-content:space-between; padding:1rem 0;'>"
		html = html & "<input class='bot ui-button ui-button--primary' type='button' value='+ AGREGAR ENTORNO' onclick='agregarFila()' style='background-color:#10b981; color:white; padding:0.75rem 1.5rem; border:none; border-radius:4px; font-size:1rem; cursor:pointer;'>"
		html = html & "</div>"
		html = html & "<div id='contenedor-nuevos'></div>"
		html = html & "<div style='display:flex; justify-content:flex-end; gap:1rem; padding:1rem 0;'>"
		html = html & "<input class='bot ui-button ui-button--primary' type='submit' value='GUARDAR' style='background-color:#2563eb; color:white; padding:0.75rem 1.5rem; border:none; border-radius:4px; font-size:1rem; cursor:pointer;'>"
		html = html & "<a class='ui-button' href='clientesDet.asp?id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "' style='background-color:#6b7280; color:white; padding:0.75rem 1.5rem; border:none; border-radius:4px; font-size:1rem; text-decoration:none;'>CANCELAR</a>"
		html = html & "</div>"
		html = html & "</form>"

		html = html & "<script>"
		html = html & "var contadorEntornos = " & numReg & ";"
		html = html & "function agregarFila() {"
		html = html & "  var c = document.getElementById('contenedor-nuevos');"
		html = html & "  var div = document.createElement('div');"
		html = html & "  div.className = 'fila-entorno';"
		html = html & "  div.style.cssText = 'padding:1rem; margin-bottom:0.75rem; background:#fff; border:1px dashed #10b981; border-radius:6px;';"
		html = html & "  div.innerHTML = "
		html = html & "    '<div class=""grid-datos-doble ui-field-grid--dual"">' +"
		html = html & "    '<div class=""dato-label"">SERVIDOR</div>' +"
		html = html & "    '<div><input class=""input-moderno ui-input"" type=""text"" name=""srv_' + contadorEntornos + '"" style=""width:90%;""></div>' +"
		html = html & "    '<div class=""dato-label"">AMBIENTE</div>' +"
		html = html & "    '<div><input class=""input-moderno ui-input"" type=""text"" name=""amb_' + contadorEntornos + '"" style=""width:90%;""></div>' +"
		html = html & "    '<div class=""dato-label"">DB RQT</div>' +"
		html = html & "    '<div><input class=""input-moderno ui-input"" type=""text"" name=""dbrqt_' + contadorEntornos + '"" style=""width:90%;""></div>' +"
		html = html & "    '<div class=""dato-label"">LICENCIA</div>' +"
		html = html & "    '<div><input class=""input-moderno ui-input"" type=""text"" name=""lic_' + contadorEntornos + '"" style=""width:90%;""></div>' +"
		html = html & "    '</div>';"
		html = html & "  c.appendChild(div);"
		html = html & "  document.getElementById('total_entornos').value = contadorEntornos;"
		html = html & "  contadorEntornos++;"
		html = html & "}"
		html = html & "</script>"

	ElseIf modo = "add" Then
		html = html & "<div class='dato-label' style='margin-bottom: 1rem; font-size: 1rem;'>AGREGAR NUEVO ENTORNO</div>"
		html = html & "<form id='frmDet' method='post' action='clientesDet.asp?opc=INS_DET&id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
		html = html & "<div class='grid-datos-doble ui-field-grid--dual'>"
		html = html & "<div class='dato-label'>SERVIDOR</div><div>" & HtmlInputCampo("servidor", "", True) & "</div>"
		html = html & "<div class='dato-label'>AMBIENTE</div><div>" & HtmlInputCampo("ambiente", "", True) & "</div>"
		html = html & "<div class='dato-label'>DB RQT</div><div>" & HtmlInputCampo("db_rqt", "", True) & "</div>"
		html = html & "<div class='dato-label'>LICENCIA</div><div>" & HtmlInputCampo("licencia", "", True) & "</div>"
		html = html & "</div>"
		html = html & "<div style='display:flex; justify-content:flex-end; gap:1rem; padding:1rem 0;'>"
		html = html & "<input class='bot ui-button ui-button--primary' type='submit' value='GUARDAR ENTORNO' style='background-color:#10b981; color:white; padding:0.75rem 1.5rem; border:none; border-radius:4px; font-size:1rem; cursor:pointer;'>"
		html = html & "<a class='ui-button' href='clientesDet.asp?id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "' style='background-color:#6b7280; color:white; padding:0.75rem 1.5rem; border:none; border-radius:4px; font-size:1rem; text-decoration:none;'>CANCELAR</a>"
		html = html & "</div>"
		html = html & "</form>"

	Else
		' MODO LECTURA: grilla readonly + botones EDITAR / AGREGAR
		html = html & "<div class='dato-label' style='margin-bottom: 1rem; font-size: 1rem;'>ENTORNOS ASIGNADOS</div>"
		html = html & "<div style='display:flex; flex-direction:column; gap:0.75rem;'>"
		Do While Not rs.eof
			vDetSrv   = rs(3) & ""
			vDetAmb   = rs(4) & ""
			vDetLic   = rs(6) & ""
			vDetDbRqt = rs(8) & ""
			xtipo = ""
			If vDetAmb = "T" Then 
				xtipo = "Test" 
			ElseIf vDetAmb = "P" Then 
				xtipo = "Prod" 
			Else 
				xtipo = vDetAmb
			End If
			html = html & "<div style='display:flex; flex-wrap:wrap; justify-content:space-between; align-items:center; padding:1rem; background-color:#f8fafc; border:1px solid #e2e8f0; border-radius:6px; color:#475569; font-size:0.9rem;'>"
			html = html & "<div style='flex:1 1 25%; min-width:150px;'><strong style='color:#1e293b;'>Srv:</strong> " & Server.HTMLEncode(vDetSrv) & " <span style='background-color:#e0e7ff; color:#4338ca; padding:0.1rem 0.4rem; border-radius:4px; font-size:0.8rem; margin-left:0.3rem; font-weight:bold;'>" & Server.HTMLEncode(xtipo) & "</span></div>"
			html = html & "<div style='flex:1 1 25%; min-width:150px;'><strong style='color:#1e293b;'>DB RQT:</strong> " & Server.HTMLEncode(vDetDbRqt) & "</div>"
			html = html & "<div style='flex:1 1 25%; min-width:150px;'><strong style='color:#1e293b;'>Lic:</strong> " & Server.HTMLEncode(vDetLic) & "</div>"
			html = html & "</div>"
			rs.movenext
		Loop
		html = html & "</div>"
		html = html & "<div style='display:flex; justify-content:flex-end; gap:1rem; padding:1rem 0;'>"
		If getPermisos(xmenu,xtmenu,"MODIF") = 1 Then
			html = html & "<a class='ui-button ui-button--primary' href='clientesDet.asp?id=" & id & "&modo=edit&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "' style='background-color:#2563eb; color:white; padding:0.75rem 1.5rem; border:none; border-radius:4px; font-size:1rem; text-decoration:none;'>EDITAR</a>"
		End If
		html = html & "<a class='ui-button ui-button--primary' href='clientesDet.asp?id=" & id & "&modo=add&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "' style='background-color:#10b981; color:white; padding:0.75rem 1.5rem; border:none; border-radius:4px; font-size:1rem; text-decoration:none;'>AGREGAR NUEVO</a>"
		html = html & "</div>"
	End If

	html = html & "</div>"
	html = html & "</div>"
	html = html & "</div>"

	rs.close
Else

	html = html & cabecera()
	html = html & "<div class='inner ui-main' style='padding: 2rem;'>"
	html = html & "<div class='bloque-datos ui-panel' style='max-width: 800px; margin: 0 auto;'>"
	html = html & "<div class='bloque-header ui-panel__header' style='text-transform: uppercase;'>AGREGAR CLIENTE</div>"
	html = html & "<div class='bloque-body ui-panel__body'>"
	html = html & "<div class='grid-datos-doble ui-field-grid--dual' style='margin-bottom: 2rem; padding-bottom: 1.5rem; border-bottom: 1px dashed #cbd5e1;'>"
	html = html & "<div class='dato-label'>NOMBRE</div><div><input class='input-moderno ui-input' id='nombre' name='nombre' type='text' required></div>"
	html = html & "<div class='dato-label'>DIRECCION</div><div><input class='input-moderno ui-input' id='direccion' name='direccion' type='text'></div>"
	html = html & "<div class='dato-label'>RUTA FICHA CLIENTE</div><div><input class='input-moderno ui-input' id='ruta_ficha' name='ruta_ficha' type='text'></div>"
	html = html & "</div>"
	html = html & "<div style='text-align: center; margin: 2rem 0;'>"
	html = html & "<input class='bot ui-button ui-button--primary' type='submit' value='GUARDAR CLIENTE' style='background-color:#2563eb; color:white; padding:0.75rem 1.5rem; border:none; border-radius:4px; font-size:1rem; cursor:pointer;'>"
	html = html & "</div>"
	html = html & "<div style='text-align: center; margin-top: 1rem;'>" & xboton & "</div>"
	html = html & "</div>" ' Fin .bloque-body
	html = html & "</div>" ' Fin .bloque-datos
	html = html & "</div>" ' Fin .inner
End If

html = html & "<script src='assets/js/jquery.min.js'></script>"
html = html & "<script src='assets/js/skel.min.js'></script>"
html = html & "<script src='assets/js/util.js'></script>"
html = html & "<script src='assets/js/main.js'></script>"

html = html & xform & "</body>"
html = html & "</html>"

response.write html
%>