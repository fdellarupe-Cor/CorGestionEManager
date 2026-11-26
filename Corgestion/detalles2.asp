<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/datosDetalles.inc" -->
<!-- #INCLUDE FILE="IncFile/ui_componentes.inc" -->
<%
'[UI-ANT] ' ===== Cuerpo anterior completo: reemplazado por el bloque [UI] que sigue. Se borra entero buscando [UI-ANT] =====

	'[UI-ANT] html = headhtml()

	'[UI-ANT] html = html & "<body>"

	'[UI-ANT] Set Cnn = Server.CreateObject("ADODB.Connection")
	'Cnn.open "[credenciales eliminadas]"
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

	'[UI-ANT] id = request.querystring("id")
	'[UI-ANT] opc = request.querystring("opc")
	'[UI-ANT] tipo = request.querystring("tipo")
	'[UI-ANT] vuel = request.querystring("vuel")
	'[UI-ANT] sol = request.querystring("sol")
	'[UI-ANT] flg = True
	'[UI-ANT] vl = request.querystring("vl")
	'[UI-ANT] incid = request.querystring("incid")
	'[UI-ANT] idproy = request.querystring("idproy")
	'[UI-ANT] cons = request.querystring("cons")
	'[UI-ANT] sel = request.querystring("sel")
	'[UI-ANT] cli = request.querystring("cli")
	'[UI-ANT] aplic = request.querystring("aplic")
	'[UI-ANT] xmenu = request.querystring("xmenu")
	'[UI-ANT] xtmenu = request.querystring("xtmenu")
	'[UI-ANT] idAplic = request.form("aplic")
	'[UI-ANT] cantregmover = request.querystring("cantregmover")

		'[UI-ANT] if idAplic & "" = "" then idAplic = request.querystring("aplic")

		'[UI-ANT] if idAplic & "" = "" then idAplic = 0

		'[UI-ANT] idtip = request.form("tip")


		'[UI-ANT] if idtip & "" = "" then idtip = request.querystring("tip")

		'[UI-ANT] if idtip & "" = "" then idtip = 0


	' response.write "idaplic: " & idaplic

		'[UI-ANT] nFecini = 0

		'[UI-ANT] If nFecini < CLng(CDate("01-01-1900")) Then
			'[UI-ANT] nFecini = CLng(CDate("01-01-1900"))
		'[UI-ANT] Else
			'[UI-ANT] nFecini = 0
		'[UI-ANT] End If

		'[UI-ANT] If request.querystring("id") & "" = "" Then id = request.form("id")

		'[UI-ANT] if id & "" = "" then id = 0

		'[UI-ANT] If tipo = "cons" Then

			'[UI-ANT] If Not ExisteApl(Cnn, id) Then flg = False

		'[UI-ANT] End If



	'[UI-ANT] html = html & "<div>"
	'[UI-ANT] html = html & "<header id='header' class='app-header' style='background:#0e1949;padding:1rem 0;'>"

	'[UI-ANT] html = html & "<div >"

	'[UI-ANT] html = html & "<div class='app-header-row' style='display:flex;padding: 0 1rem;align-items:center;'>"

	'[UI-ANT] html = html & "<div style='width:75%;' class='app-header-brand'><a href='index.asp' class='logo' >"
	'[UI-ANT] html = html & "<img src='images/iso_cor.png' style='width:6rem;' >"
	'[UI-ANT] html = html & "<span class='logo-text' style='font-size:15pt;font-family:sans-serif;color:white;'><b>Cor Gestion</b> | Sistema Administrativo</span></a></div>"
	'[UI-ANT] html = html & "<div style='width:25%;display:flex;align-items:center;flex-direction:row-reverse;'>"
	'[UI-ANT] html = html & "<nav>"
	'[UI-ANT] html = html & "<ul>"
	'html = html & "<li><a href='#menu'></a></li>"
	'[UI-ANT] html = html & "</ul>"
	'[UI-ANT] html = html & "</nav>"
	'[UI-ANT] html = html & "</div>"

	'[UI-ANT] html = html & "</div>"

	'[UI-ANT] html = html & "</div>"

	'[UI-ANT] html = html & "</header>"


	'[UI-ANT] If opc = "ver" And flg Then

		'[UI-ANT] xxx = ObjAcc(id,"A",cnn,"D")
		'[UI-ANT] XXX = SYSLOGON(cnn,"CONSULTA_APL_" & id)

		'[UI-ANT] html = html & "<div class='titulo-seccion page-toolbar' style='padding:0 2rem;background-color:#e8edff;' >"

		'[UI-ANT] apl = getName(cnn,"nombre","apl",id)

		'[UI-ANT] html = html & "<div><h2>DETALLE - '" & apl & "' (ID:" & id & ")</h2></div>"

		'[UI-ANT] form9 = "<form id='form9' style='height:70%;' method='post' action='detalles2.asp?cli=" & cli & "&id=" & id & "&opc=act&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&Aplic=" & idAplic & "&tip=" & idtip & "'></form>"

		'[UI-ANT] if getPermisos(xmenu,xtmenu,"MODIF") = 1 then
			'[UI-ANT] boton9 = "<input class='bot ui-button' type='button' value='ACTUALIZAR APL' onclick=""document.getElementById('form9').submit();"">"
			'[UI-ANT] else
				'[UI-ANT] boton9 = "<input class='botsp ui-button' type='button' value='ACTUALIZAR APL' onclick=""alert('Sin permiso para modificar.');"">"
		'[UI-ANT] end if

		'[UI-ANT] html = html & "<div class='page-toolbar-action'>" & boton9 & "</div>"

		'[UI-ANT] If tipo <> "cons" Then

			'[UI-ANT] cli = request.querystring("cli")

			'[UI-ANT] form9x = "<form id='form9x' method='post' action='apl.asp?cli=" & cli & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&Aplic=" & idAplic & "&tip=" & idtip &  "'></form>"

			'[UI-ANT] boton8 = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('form9x').submit();""></form>"

			'[UI-ANT] html = html & "<div class='page-toolbar-action'>" & boton8 & "</div>"

			'[UI-ANT] Else

				'[UI-ANT] cli = request.querystring("cli")
					'[UI-ANT] if cli & "" = "" then cli = 0
				'[UI-ANT] IdAplic = request.querystring("aplic")
					'[UI-ANT] if IdAplic & "" = "" then IdAplic = 0
				'[UI-ANT] sel = request.querystring("sel")
					'[UI-ANT] if sel & "" = "" then sel = 0

				'[UI-ANT] asp = request.querystring("busc")

				'[UI-ANT] form8 = "<form id='form8' style='height:70%;' method='post' action='apl.asp?sel=" & sel & "&opc=consi&cli=" & cli & "&idaplic=" & idaplic & "&busc=" & asp & "'>"

				'[UI-ANT] boton8 = "<input class='bot ui-button' style='height:70%;font-size:13pt;' type='button' value='VOLVER' onclick=""document.getElementById('form8').submit();""></form>"

				'[UI-ANT] html = html & "<div class='page-toolbar-action'>" & form8 & boton8 & "</div>"

		'[UI-ANT] End If

		'[UI-ANT] html = html & "</div>"
		'[UI-ANT] html = html & "</div>"


		'[UI-ANT] Set rs = ObtenerDetalleApl(Cnn, id)

		'[UI-ANT] If Not rs.eof Then

			'[UI-ANT] html = html & "<div style='width:100%;'>"
			'[UI-ANT] html = html & "<div id='tabs' style='border-style: none;'>"
			
			'[UI-ANT] html = html & "<ul class='nav-tabs ui-tabbar'>"
			'[UI-ANT] html = html & "<li><a class='nav-tab active ui-tabbar__link is-active' style='display: block;' href='#tabs-1'>DATOS</a></li>"
			'[UI-ANT] html = html & "<li><a class='nav-tab ui-tabbar__link' style='display: block;' href='#tabs-2'>ARREGLOS Y MODULOS</a></li>"
			'[UI-ANT] html = html & "<li><a class='nav-tab ui-tabbar__link' style='display: block;' href='#tabs-3'>INCIDENTES Y PROYECTOS RELACIONADOS</a></li>"
			'[UI-ANT] html = html & "</ul>"

		'DATOS
			'[UI-ANT] html = html & "<div id='tabs-1'>"

			'[UI-ANT] html = html & "<div class='bloque-datos ui-panel'>"
			'[UI-ANT] html = html & "<div class='bloque-body ui-panel__body grid-datos-doble ui-field-grid--dual'>"

			'[UI-ANT] html = html & "<div class='dato-label'>ID</div>"
			'[UI-ANT] html = html & "<div class='dato-value' style='font-weight: 600;'>" & rs(0) & "</div>"

			'[UI-ANT] html = html & "<div class='dato-label'>NOMBRE</div>"
			'[UI-ANT] html = html & "<div class='dato-value'>" & rs(1) & "</div>"

			'[UI-ANT] html = html & "<div class='dato-label'>CLIENTE</div>"
			'[UI-ANT] html = html & "<div class='dato-value'>" & rs(2) & "</div>"

			'[UI-ANT] html = html & "<div class='dato-label'>APLICATIVO</div>"
			'[UI-ANT] html = html & "<div class='dato-value'>" & rs(3) & "</div>"

			'[UI-ANT] html = html & "<div class='dato-label'>RESPONSABLE ENVIO</div>"
			'[UI-ANT] html = html & "<div class='dato-value'>" & rs(4) & "</div>"

			'[UI-ANT] sFec = Fechador(rs("fecha_creacion") + nFecini)
			'[UI-ANT] html = html & "<div class='dato-label'>FECHA INGRESO</div>"
			'[UI-ANT] html = html & "<div class='dato-value'>" & sFec & "</div>"

			'[UI-ANT] sFE = ""
			'[UI-ANT] if rs(5) <> 0 then
				'[UI-ANT] sFE = Fechador(rs(5) + nFecini)
			'[UI-ANT] end if

			'[UI-ANT] html = html & "<div class='dato-label'>FECHA ENVIO</div>"
			'[UI-ANT] html = html & "<div class='dato-value'>" & sFE & "</div>"

			'[UI-ANT] sFIT = ""
			'[UI-ANT] if rs(6) <> 0 then
				'[UI-ANT] sFIT = Fechador(rs(6) + nFecini)
			'[UI-ANT] end if

			'[UI-ANT] html = html & "<div class='dato-label'>FECHA INSTALACION TEST</div>"
			'[UI-ANT] html = html & "<div class='dato-value'>" & sFIT & "</div>"

			'[UI-ANT] sFIP = ""
			'[UI-ANT] if rs(8) <> 0 then
				'[UI-ANT] sFIP = Fechador(rs(8) + nFecini)
			'[UI-ANT] end if

			'[UI-ANT] html = html & "<div class='dato-label'>FECHA INSTALACION PROD</div>"
			'[UI-ANT] html = html & "<div class='dato-value'>" & sFIP & "</div>"

			'[UI-ANT] html = html & "</div>" ' Fin .bloque-body
			'[UI-ANT] html = html & "</div>" ' Fin .bloque-datos
			'[UI-ANT] html = html & "</div>" ' Fin #tabs-1

		'DETALLES
			'[UI-ANT] html = html & "<div id='tabs-2'>"

			'[UI-ANT] html = html & "<div class='bloque-datos ui-panel'>"
			'[UI-ANT] html = html & "<div class='bloque-body ui-panel__body'>"

			'[UI-ANT] html = html & "<div class='dato-label' style='margin-bottom: 0.5rem;'>DETALLE DEL ARREGLO/MODULO</div>"
			'[UI-ANT] html = html & "<div class='dato-value' style='width: 100%;'>" & Replace(rs(7) & "",Chr(13),"<br>") & "</div>"

			'[UI-ANT] html = html & "</div>" ' Fin .bloque-body
			'[UI-ANT] html = html & "</div>" ' Fin .bloque-datos
			'[UI-ANT] html = html & "</div>" ' Fin #tabs-2

		'INCIDENTES O PROYECTOS RELACIONADOS
			'[UI-ANT] html = html & "<div id='tabs-3'>"

			'[UI-ANT] html = html & "<div class='bloque-datos ui-panel'>"
			'[UI-ANT] html = html & "<div class='bloque-body ui-panel__body grid-datos-doble ui-field-grid--dual'>"
			' --- SECCION: INCIDENTES ---
			'[UI-ANT] html = html & "<div>"

			'[UI-ANT] Set rz = ObtenerRelacionesApl(Cnn, id, "I")

			'[UI-ANT] html = html & "<div class='dato-label'>INCIDENTES RELACIONADOS</div>"
			'[UI-ANT] html = html & "<div style='display: flex; flex-direction: column; gap: 0.8rem; margin-top: 1rem;'>"

			'[UI-ANT] if not rz.eof then
				'[UI-ANT] do while not rz.eof
					'[UI-ANT] html = html & "<div>"
					'[UI-ANT] html = html & "<a href='detalles.asp?opc=veri&id=" & rz(2) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "' style='text-decoration:none;'>"
					'[UI-ANT] html = html & "<span style='cursor:pointer; text-decoration:underline; color:#FF0000; font-weight:600;'>" & rz(2) & " - " & rz(6) & "</span>"
					'[UI-ANT] html = html & "</a>"
					'[UI-ANT] html = html & "</div>"
					'[UI-ANT] rz.movenext
				'[UI-ANT] loop
			'[UI-ANT] else
				'[UI-ANT] html = html & "<div class='dato-value'>-----</div>"
			'[UI-ANT] end if

			'[UI-ANT] html = html & "</div>" ' Fin flex col Incidentes
			'[UI-ANT] html = html & "</div>" ' Fin columna 1 de la grid

			' --- SECCION: PROYECTOS ---
			'[UI-ANT] html = html & "<div>"

			'[UI-ANT] Set rz = ObtenerRelacionesApl(Cnn, id, "P")

			'[UI-ANT] html = html & "<div class='dato-label'>PROYECTOS RELACIONADOS</div>"
			'[UI-ANT] html = html & "<div style='display: flex; flex-direction: column; gap: 0.8rem; margin-top: 1rem;'>"

			'[UI-ANT] if not rz.eof then
				'[UI-ANT] do while not rz.eof
					'[UI-ANT] html = html & "<div>"
					'[UI-ANT] html = html & "<a href='detalles.asp?opc=verp&id=" & rz(2) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "' style='text-decoration:none;'>"
					'[UI-ANT] html = html & "<span style='cursor:pointer; text-decoration:underline; color:#FF0000; font-weight:600;'>" & rz(2) & " - " & rz(6) & "</span>"
					'[UI-ANT] html = html & "</a>"
					'[UI-ANT] html = html & "</div>"
					'[UI-ANT] rz.movenext
				'[UI-ANT] loop
			'[UI-ANT] else
				'[UI-ANT] html = html & "<div class='dato-value'>-----</div>"
			'[UI-ANT] end if

			'[UI-ANT] html = html & "</div>" ' Fin flex col Proyectos
			'[UI-ANT] html = html & "</div>" ' Fin columna 2 de la grid
		'[UI-ANT] end if

		'[UI-ANT] html = html & "</div>" ' Fin .bloque-body
		'[UI-ANT] html = html & "</div>" ' Fin .bloque-datos
		'[UI-ANT] html = html & "</div>" ' Fin #tabs-3

		'[UI-ANT] html = html & "</div>" ' Cierra el <div id='tabs'>
		'[UI-ANT] html = html & "</div>" ' Cierra el <div style='width:100%;'> principal

	'End If

	'[UI-ANT] ElseIf opc = "act" And flg Then

		'[UI-ANT] xmenu = request.querystring("xmenu")
		'[UI-ANT] xtmenu = request.querystring("xtmenu")

		'[UI-ANT] zzz = ObjAcc(id,"A",cnn,"A")
		'[UI-ANT] XXX = SYSLOGON(cnn,"ACTUALIZA_APL_" & id)

		'[UI-ANT] html = html & "<div class='titulo-seccion page-toolbar' style='padding:0 2rem;background-color:#e8edff;' >"
		'[UI-ANT] html = html & "<div><h2>ACTUALIZAR - APL NRO '" & id & "'</h2></div>"

		'[UI-ANT] cliente = request.querystring("cli")
		'[UI-ANT] cantregmover = request.querystring("cantregmover")

		'[UI-ANT] form10 = "<form id='form10' style='height:70%;' method='post' action='apl.asp?cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&cantregmover=" & cantregmover & "'></form>"

		'[UI-ANT] boton10 = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('form10').submit();""></form>"

		'[UI-ANT] html = html & "<div class='page-toolbar-action'>" & boton10 & "</div>"
		'[UI-ANT] html = html & "</div>"
		'[UI-ANT] html = html & "</div>"

		'[UI-ANT] Set rs = ObtenerAplEdicion(Cnn, id)

		'[UI-ANT] If Not rs.eof Then

			' CORRECCION 1: Se remueve el style='height:1px;' para permitir que el form tenga dimension completa
			'[UI-ANT] html = html & "<form id='sForm' method='post' action='actualizar.asp?tabla=apl&id=" & id & "&opt=actapl&cli=" & cli & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=detalles2.asp?opc=act&cantregmover=" & cantregmover & "'>"

			'[UI-ANT] html = html & "<div class='inner ui-main' style='padding: 2rem;'>"

			' --- COLUMNA IZQUIERDA ---
			'[UI-ANT] html = html & "<div style='width: 50%; float: left; padding-right: 1.5rem;'>"

			'[UI-ANT] html = html & "<div class='bloque-datos ui-panel'>"
			'[UI-ANT] html = html & "<div class='bloque-body ui-panel__body grid-datos-doble ui-field-grid--dual'>"

			' ID
			'[UI-ANT] html = html & "<div class='dato-label'>ID</div>"
			'[UI-ANT] html = html & "<div id='nid' class='dato-value' style='font-weight: 600; padding-left: 0.5rem;'>" & rs(0) & "</div>"

			' NOMBRE
			'[UI-ANT] html = html & "<div class='dato-label'>NOMBRE</div>"
			'[UI-ANT] html = html & "<div>"
			'[UI-ANT] html = html & "<input id='nombre' name='nombre' type='text' class='input-moderno ui-input' value='" & rs(1) & "'>"
			'[UI-ANT] html = html & "</div>"

			' SECUENCIA
			'[UI-ANT] html = html & "<div class='dato-label'>SECUENCIA</div>"
			'[UI-ANT] html = html & "<div>"
			'[UI-ANT] html = html & "<input id='sec' name='sec' type='text' class='input-moderno ui-input' autocomplete='off' value='" & rs(11) & "'>"
			'[UI-ANT] html = html & "</div>"

			' TIPO
			'[UI-ANT] Set rsTipo = GetListaCombo(Cnn, "TIPO_APL")

			'[UI-ANT] html = html & "<div class='dato-label'>TIPO</div>"
			'[UI-ANT] html = html & "<div class='dato-value'>"
			'[UI-ANT] html = html & "<select class='ui-select' id='tipo' name='tipo'>"
			'[UI-ANT] html = html & "<option value=" & rs(12) & " selected>" & rs(13) & "</option>"

			'[UI-ANT] Do While Not rsTipo.EOF
				'[UI-ANT] html = html & "<option value=" & rsTipo(0) & ">" & rsTipo(1) & "</option>"
				'[UI-ANT] rsTipo.movenext
			'[UI-ANT] Loop

			'[UI-ANT] html = html & "</select></div>"

			' CLIENTE
			'[UI-ANT] html = html & "<div class='dato-label'>CLIENTE</div>"
			'[UI-ANT] html = html & "<div class='dato-value'>"
			'[UI-ANT] html = html & ArmarComboHTML(GetListaCombo(Cnn, "CLIENTES"), rs(2), "id_cliente")
			'[UI-ANT] html = html & "</div>"

			' APLICACION
			'[UI-ANT] html = html & "<div class='dato-label'>APLICACION</div>"
			'[UI-ANT] html = html & "<div class='dato-value'>"
			'[UI-ANT] html = html & ArmarComboHTML(GetListaCombo(Cnn, "APLICATIVOS"), rs(3), "id_aplic")
			'[UI-ANT] html = html & "</div>"

			' RESPONSABLE ENVIO
			'[UI-ANT] html = html & "<div class='dato-label'>RESPONSABLE ENVIO</div>"
			'[UI-ANT] html = html & "<div class='dato-value'>"
			'[UI-ANT] html = html & ArmarComboHTML(GetListaCombo(Cnn, "RESPONSABLES_FUNC"), rs(4), "respon_envio")
			'[UI-ANT] html = html & "</div>"

			' FECHA ENVIO
			'[UI-ANT] sFE = ""
			'[UI-ANT] if rs(5) <> 0 then
				'[UI-ANT] sFE = Fechador(rs(5) + nFecini)
			'[UI-ANT] end if

			'[UI-ANT] html = html & "<div class='dato-label'>FECHA ENVIO</div>"
			'[UI-ANT] html = html & "<div>"
			'[UI-ANT] html = html & "<input id='fecha_envio' name='fecha_envio' type='text' class='input-moderno ui-input' value='" & sFE & "'>"
			'[UI-ANT] html = html & "</div>"

			' FECHA TEST
			'[UI-ANT] sFIT = ""
			'[UI-ANT] if rs(6) <> 0 then
				'[UI-ANT] sFIT = Fechador(rs(6) + nFecini)
			'[UI-ANT] end if

			'[UI-ANT] if rs(5) & "" <> "" then
				'[UI-ANT] html = html & "<div class='dato-label'>FECHA INSTALACION TEST</div>"
				'[UI-ANT] html = html & "<div>"
				'[UI-ANT] html = html & "<input id='fecha_instalacion_test' name='fecha_instalacion_test' type='text' class='input-moderno ui-input' value='" & sFIT & "'>"
				'[UI-ANT] html = html & "</div>"
			'[UI-ANT] else
				'[UI-ANT] html = html & "<div class='dato-label'>FECHA INSTALACION TEST</div>"
				'[UI-ANT] html = html & "<div class='dato-value'>" & sFIT & "</div>"
			'[UI-ANT] end if

			' FECHA PROD
			'[UI-ANT] sFIP = ""
			'[UI-ANT] if rs(8) <> 0 then
				'[UI-ANT] sFIP = Fechador(rs(8) + nFecini)
			'[UI-ANT] end if

			'[UI-ANT] if rs(6) & "" <> "" then
				'[UI-ANT] html = html & "<div class='dato-label'>FECHA INSTALACION PROD</div>"
				'[UI-ANT] html = html & "<div>"
				'[UI-ANT] html = html & "<input id='fecha_instalacion_prod' name='fecha_instalacion_prod' type='text' class='input-moderno ui-input' value='" & sFIP & "'>"
				'[UI-ANT] html = html & "</div>"
			'[UI-ANT] else
				'[UI-ANT] html = html & "<div class='dato-label'>FECHA INSTALACION PROD</div>"
				'[UI-ANT] html = html & "<div class='dato-value'>" & sFIP & "</div>"
			'[UI-ANT] end if

			' BOTON AGREGAR
			'[UI-ANT] html = html & "<div style='grid-column: 1 / -1; text-align: center; margin-top: 1.5rem;'>"
			'[UI-ANT] html = html & "<span id='sAgre' class='ui-button ui-button--quiet' style='color:#FF0000; cursor:pointer; font-weight:600; text-transform:uppercase;'>AGREGAR INC O PROY</span>"
			'[UI-ANT] html = html & "</div>"

			'[UI-ANT] html = html & "</div>" ' Fin .bloque-body
			'[UI-ANT] html = html & "</div>" ' Fin .bloque-datos

			'[UI-ANT] html = html & "</div>" ' Fin columna izquierda

			' --- COLUMNA DERECHA ---
			'[UI-ANT] html = html & "<div style='width: 50%; float: right; padding-left: 1.5rem;'>"

			'[UI-ANT] html = html & "<div class='bloque-datos ui-panel'>"

			'[UI-ANT] html = html & "<div class='bloque-header ui-panel__header'>ARREGLOS QUE IMPLICA</div>"

			'[UI-ANT] html = html & "<div class='bloque-body ui-panel__body'>"

			'[UI-ANT] html = html & "<textarea id='motivo_apl' name='motivo_apl' class='input-moderno ui-input' style='height: 400px; resize: vertical;'>" & rs(7) & "</textarea>"

			'[UI-ANT] html = html & "<div class='ui-edit-actions' style='text-align: center; margin-top: 1rem;'>"
			'[UI-ANT] html = html & "<input class='btn-principal ui-button ui-button--primary' type='button' value='ACTUALIZAR' onclick=""actualizar('sForm');"">"
			'[UI-ANT] html = html & "</div>"
			'[UI-ANT] html = html & "</div>" ' Fin .bloque-body
			'[UI-ANT] html = html & "</div>" ' Fin .bloque-datos

			'[UI-ANT] html = html & "</div>" ' Fin columna derecha

			'[UI-ANT] html = html & "<div style='clear: both;'></div>"
			'[UI-ANT] html = html & "</div>" ' Fin .inner

			' CORRECCION 2: Cierre de formulario antes de insertar el contenedor del Modal Dialog
			'[UI-ANT] html = html & "</form>"

			' El modal se ubica independiente del formulario
			'[UI-ANT] html = html & "<div id='dialog' title=''></div>"

		'[UI-ANT] End If

	'[UI-ANT] End If

' CORRECCION 3: Importacion de script JS antes de la inicializacion jQuery
'[UI-ANT] html = html & "<script src='assets/js/relaciones.js'></script>"

'[UI-ANT] html = html & "<script type='text/javascript'>"
'[UI-ANT] html = html & " function actualizar(formu){ "
'[UI-ANT] html = html & " var r = confirm('Desea actualizar apl?'); "
'[UI-ANT] html = html & " if (r){ "
'[UI-ANT] html = html & " document.getElementById(formu).submit(); "
'[UI-ANT] html = html & " }else{ "
'[UI-ANT] html = html & " return false;"
'[UI-ANT] html = html & " }} "
'[UI-ANT] html = html & "</script>"

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
'[UI-ANT] html = html & "dayNamesShort: ['Dom','Lun','Mar','Mi','Juv','Vie','Sab'],"
'[UI-ANT] html = html & "dayNamesMin: ['Do','Lu','Ma','Mi','Ju','Vi','Sa'],"
'[UI-ANT] html = html & "weekHeader: 'Sm',"
'[UI-ANT] html = html & "dateFormat: 'dd/mm/yy',"
'[UI-ANT] html = html & "firstDay: 1,"
'[UI-ANT] html = html & "isRTL: false,"
'[UI-ANT] html = html & "showMonthAfterYear: false,"
'[UI-ANT] html = html & "yearSuffix: ''"
'[UI-ANT] html = html & "};"
'[UI-ANT] html = html & "$.datepicker.setDefaults($.datepicker.regional['es']);"

'[UI-ANT] html = html & "    $( ""#fecha_envio"" ).datepicker();"
'[UI-ANT] html = html & "    $( ""#fecha_instalacion_test"" ).datepicker();"
'[UI-ANT] html = html & "    $( ""#fecha_instalacion_prod"" ).datepicker();"
'[UI-ANT] html = html & "  } );"

'[UI-ANT] html = html & "$( '#tabs' ).tabs();"

'[UI-ANT] html = html & "$( '#dialog' ).dialog({"
'[UI-ANT] html = html & "title: 'Agregar relacion',"
'[UI-ANT] html = html & "autoOpen: false,"
'[UI-ANT] html = html & "resizable: false,"
'[UI-ANT] html = html & "buttons: {"
'[UI-ANT] html = html & "'Aceptar': function () {"
'[UI-ANT] html = html & "capturarYProcesar();"
'[UI-ANT] html = html & "$(this).dialog('close');"
'[UI-ANT] html = html & "},"
'[UI-ANT] html = html & "'Cerrar': function () {"
'[UI-ANT] html = html & "$(this).dialog('close');"
'[UI-ANT] html = html & "}"
'[UI-ANT] html = html & "},"
'[UI-ANT] html = html & "show: {"
'[UI-ANT] html = html & "effect: 'blind',"
'[UI-ANT] html = html & "duration: 100"
'[UI-ANT] html = html & "      },"
'[UI-ANT] html = html & "hide: {"
'[UI-ANT] html = html & "effect: 'blind',"
'[UI-ANT] html = html & "duration: 100"
'[UI-ANT] html = html & "}"
'[UI-ANT] html = html & "});"

' CORRECCION 4: Captura delegada del evento click en $(document)
'[UI-ANT] html = html & "$(document).on('click', '#sAgre', function() {"
'[UI-ANT] html = html & "    $.get('armaCargaParam.asp?param=IOP', function( data ) {"
'[UI-ANT] html = html & "        $( '#dialog' ).html(data);"
'[UI-ANT] html = html & "        if (typeof inicializarRelaciones === 'function') {"
'[UI-ANT] html = html & "            inicializarRelaciones();"
'[UI-ANT] html = html & "        }"
'[UI-ANT] html = html & "    });"
'[UI-ANT] html = html & "    $( '#dialog' ).dialog( 'open' );"
'[UI-ANT] html = html & "});"

'[UI-ANT] html = html & "</script>"

'[UI-ANT] html = html & form9x & form9 & form10 & "</body>"
'[UI-ANT] html = html & "</html>"


'[UI-ANT] response.write html

' =========================================================================
' [UI] PANTALLA con el sistema visual (IncFile/ui_componentes.inc)
' Detalle de APL (opc=ver) y actualizacion (opc=act).
' El cuerpo anterior completo queda arriba, comentado con '[UI-ANT].
' - ver: detalle de trabajo del APL (ciclo, arreglos y relacionados) con Actualizar segun permiso.
'   La vista de consulta de APL es otra pantalla: vistaconsulta.asp?tipo=A (desde Consultas).
' - act: mismo formulario y destino (actualizar.asp) y mismo dialogo de relaciones (relaciones.js).
' Cambios: cabecera comun, datos escapados, fechas legibles, id validado, permiso de modificar
' chequeado tambien en el servidor y Volver de la consulta con todos sus filtros.
' =========================================================================
Dim accTop, metas, esConsulta, apl, urlRel, relInc, relPry, cantRel, rz, tituloRel, puedeModif, fechaEnv, fechaTest, fechaProd

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

id = request.querystring("id")
opc = request.querystring("opc")
tipo = request.querystring("tipo")
cli = request.querystring("cli")
xmenu = request.querystring("xmenu")
xtmenu = request.querystring("xtmenu")
cantregmover = request.querystring("cantregmover")
idAplic = request.form("aplic")
If idAplic & "" = "" Then idAplic = request.querystring("aplic")
If idAplic & "" = "" Then idAplic = 0
idtip = request.form("tip")
If idtip & "" = "" Then idtip = request.querystring("tip")
If idtip & "" = "" Then idtip = 0
If request.querystring("id") & "" = "" Then id = request.form("id")

' [UI] El id va a las consultas: se valida como numero
If IsNumeric(id & "") And Trim(id & "") <> "" Then id = CLng(id) Else id = 0
esConsulta = (tipo = "cons")
puedeModif = (getPermisosAtrm("CG_APL", "MODIF") = 1)

flg = (id > 0)
' [LOG] Consulta de un APL (CG_LOG)
If flg And (opc = "ver" Or esConsulta) Then CgLog "APL", id, "CONSULTA"
If esConsulta And flg Then
	If Not ExisteApl(Cnn, id) Then flg = False
End If

html = html & "<div id='wrapper' role='banner'>" & cabecera() & "</div>"

' --- Volver: a la consulta (con todos sus filtros) o al listado de APL ---
If esConsulta Then
	' [UI] Antes mandaba idaplic (apl.asp lee aplic) y perdia la fecha: volvia sin el filtro de aplicacion
	form8 = "<form id='form8' class='form-oculto' method='post' action='apl.asp?opc=consi&sel=2&cli=" & Server.URLEncode(cli & "") & "&aplic=" & Server.URLEncode(idAplic & "") & _
		"&busc=" & Server.URLEncode(request.querystring("busc") & "") & "&fec=" & Server.URLEncode(request.querystring("fec") & "") & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
Else
	form8 = "<form id='form8' class='form-oculto' method='post' action='apl.asp?cli=" & Server.URLEncode(cli & "") & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&aplic=" & Server.URLEncode(idAplic & "") & "&tip=" & Server.URLEncode(idtip & "") & "&cantregmover=" & Server.URLEncode(cantregmover & "") & "'></form>"
End If

' [CONSULTA] Si se llego desde vistaconsulta.asp (vc=1), Volver regresa a la vista de consulta del APL con los mismos filtros
If request.querystring("vc") = "1" Then
	form8 = "<form id='form8' class='form-oculto' method='post' action='vistaconsulta.asp?tipo=A&id=" & id & "&cli=" & Server.URLEncode(cli & "") & "&aplic=" & Server.URLEncode(idAplic & "") & _
		"&busc=" & Server.URLEncode(request.querystring("busc") & "") & "&fec=" & Server.URLEncode(request.querystring("fec") & "") & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
End If

' =============================================================================
' VER (y vista de consulta)
' =============================================================================
If opc = "ver" And flg Then

	xxx = ObjAcc(id, "A", cnn, "D")
	XXX = SYSLOGON(cnn, "CONSULTA_APL_" & id)

	Set rs = ObtenerDetalleApl(Cnn, id)

	' Actualizar (mismo destino de antes), segun el permiso de modificar
	form9 = "<form id='form9' class='form-oculto' method='post' action='detalles2.asp?cli=" & Server.URLEncode(cli & "") & "&id=" & id & "&opc=act&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&aplic=" & Server.URLEncode(idAplic & "") & "&tip=" & Server.URLEncode(idtip & "") & "&cantregmover=" & Server.URLEncode(cantregmover & "") & "'></form>"
	accTop = UiBoton("Actualizar APL", "light", "document.getElementById('form9').submit();", puedeModif, "Sin permiso para modificar.")
	accTop = accTop & UiBoton("Volver", "secondary", "document.getElementById('form8').submit();", True, "")

	If rs.eof Then
		html = html & UiPantallaAbre("APL", accTop, "")
		html = html & UiCardAbre("APL", "") & "<div class='ui-empty'>No se encontr&oacute; el APL " & id & ".</div>" & UiCardCierra()
		html = html & UiPantallaCierra()
	Else
		' --- Barra superior: identificacion y datos clave ---
		metas = AplMeta("Cliente", UiDato(rs(2))) & AplMeta("Aplicativo", UiDato(rs(3))) & AplMeta("Responsable del env&iacute;o", UiDato(rs(4))) & AplMeta("Secuencia", UiDato(rs(12)))
		html = html & "<main id='contenido'>"
		html = html & "<div class='ui-app ui-topbar'><div class='ui-topbar__inner'>" & _
			"<div class='ui-topbar__top'>" & _
				"<div class='ui-topbar__ident'><span class='ui-topbar__id'>APL-" & rs(0) & "</span><h1 class='ui-topbar__title'>" & UiTexto(rs(1)) & "</h1></div>" & _
				"<div class='ui-topbar__actions'>" & accTop & "</div>" & _
			"</div>" & _
			"<div class='ui-topbar__bottom'><div class='ui-topbar__meta'>" & metas & "</div><div></div></div>" & _
			"</div></div>"
		html = html & "<div class='ui-app ui-page'><div class='ui-container'>"

		' --- Ciclo del APL + Arreglos ---
		html = html & "<div class='ui-grid-2'>"
		html = html & "<section class='ui-card' aria-labelledby='tituloCiclo'><h2 class='ui-card__title' id='tituloCiclo'>Ciclo del APL</h2><div class='ui-lifecycle'>"
		html = html & AplHito("Creaci&oacute;n", rs(9))
		html = html & AplHito("Env&iacute;o al cliente", rs(5))
		html = html & AplHito("Instalaci&oacute;n en test", rs(6))
		html = html & AplHito("Instalaci&oacute;n en producci&oacute;n", rs(8))
		html = html & "</div></section>"

		html = html & "<section class='ui-card' aria-labelledby='tituloArreglos'><h2 class='ui-card__title' id='tituloArreglos'>Arreglos y m&oacute;dulos</h2>"
		If Trim(rs(7) & "") = "" Then
			html = html & "<div class='ui-empty'>Sin detalle de arreglos cargado.</div>"
		Else
			html = html & "<div class='ui-prose'>" & AplTexto(rs(7)) & "</div>"
		End If
		html = html & "</section>"
		html = html & "</div>"

		'[UI-ANT] ' --- Incidentes y proyectos relacionados ---
		'[UI-ANT] relInc = AplRelacionados(id, "I")
		'[UI-ANT] relPry = AplRelacionados(id, "P")
		'[UI-ANT] html = html & "<section class='ui-card' aria-labelledby='tituloRel'><h2 class='ui-card__title' id='tituloRel'>Incidentes y proyectos relacionados</h2>"
		'[UI-ANT] html = html & "<div class='ui-grid-2 ui-grid-2--plano'>"
		'[UI-ANT] html = html & "<div><h3 class='ui-apl-rel__titulo'>Incidentes</h3>" & relInc & "</div>"
		'[UI-ANT] html = html & "<div><h3 class='ui-apl-rel__titulo'>Proyectos</h3>" & relPry & "</div>"
		'[UI-ANT] html = html & "</div></section>"
		' --- Incidentes y proyectos relacionados, con el pop-up para relacionar uno (tipo y numero) ---
		relInc = AplRelacionados(id, "I")
		relPry = AplRelacionados(id, "P")
		If puedeModif Then
			html = html & UiCardAbre("Incidentes y proyectos relacionados", "<button type='button' class='ui-button ui-button--light ui-button--sm' id='btnRel' aria-haspopup='dialog' onclick=""aplAbrirRel()"">+ Relacionar incidente o proyecto</button>")
		Else
			html = html & UiCardAbre("Incidentes y proyectos relacionados", UiBoton("+ Relacionar incidente o proyecto", "light", "", False, "Sin permiso para modificar."))
		End If
		html = html & AplAvisoRelacion()
		html = html & "<div class='ui-grid-2 ui-grid-2--plano'>"
		html = html & "<div><h3 class='ui-apl-rel__titulo'>Incidentes</h3>" & relInc & "</div>"
		html = html & "<div><h3 class='ui-apl-rel__titulo'>Proyectos</h3>" & relPry & "</div>"
		html = html & "</div>"
		If puedeModif Then html = html & AplDialogoRelacion(id)
		html = html & UiCardCierra()

		html = html & "</div></div></main>"
	End If

' =============================================================================
' ACTUALIZAR
' =============================================================================
ElseIf opc = "act" And flg Then

	accTop = UiBoton("Volver", "secondary", "document.getElementById('form10').submit();", True, "")
	form10 = "<form id='form10' class='form-oculto' method='post' action='apl.asp?cli=" & Server.URLEncode(cli & "") & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&cantregmover=" & Server.URLEncode(cantregmover & "") & "'></form>"
	html = html & UiPantallaAbre("Actualizar APL-" & id, accTop, "")

	If Not puedeModif Then
		' [UI] El permiso se chequea tambien aca: antes se podia entrar directo por la URL
		html = html & UiCardAbre("Actualizar APL", "") & "<div class='ui-empty'>No ten&eacute;s permiso para modificar APL.</div>" & UiCardCierra()
	Else
		zzz = ObjAcc(id, "A", cnn, "A")
		XXX = SYSLOGON(cnn, "ACTUALIZA_APL_" & id)

		Set rs = ObtenerAplEdicion(Cnn, id)
		If rs.eof Then
			html = html & UiCardAbre("Actualizar APL", "") & "<div class='ui-empty'>No se encontr&oacute; el APL " & id & ".</div>" & UiCardCierra()
		Else
			fechaEnv = AplFechaInput(rs(5))
			fechaTest = AplFechaInput(rs(6))
			fechaProd = AplFechaInput(rs(8))

			' Mismo destino y campos de antes
			html = html & "<form id='sForm' method='post' action='actualizar.asp?tabla=apl&id=" & id & "&opt=actapl&cli=" & Server.URLEncode(cli & "") & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=detalles2.asp?opc=act&cantregmover=" & Server.URLEncode(cantregmover & "") & "'>"
			' ID del APL: lo lee relaciones.js (capturarYProcesar) para armar la relacion
			html = html & "<span id='nid' hidden>" & Server.HTMLEncode(id & "") & "</span>"
			html = html & UiCardAbre("Datos del APL", "") & UiFormGridAbre()
			html = html & UiCampo("nombre", "Nombre", UiInputTexto("nombre", False, "value='" & Server.HTMLEncode(rs(1) & "") & "'"), False, False)
			html = html & UiCampo("sec", "Secuencia", UiInputTexto("sec", False, "autocomplete='off' value='" & Server.HTMLEncode(rs(11) & "") & "'"), False, False)
			html = html & UiCampo("tipo", "Tipo", UiSelectRs("tipo", GetListaCombo(Cnn, "TIPO_APL"), rs(12), "", "", ""), False, False)
			html = html & UiCampo("id_cliente", "Cliente", UiSelectRs("id_cliente", GetListaCombo(Cnn, "CLIENTES"), rs(2), "", "", ""), False, False)
			html = html & UiCampo("id_aplic", "Aplicaci&oacute;n", UiSelectRs("id_aplic", GetListaCombo(Cnn, "APLICATIVOS"), rs(3), "", "", ""), False, False)
			html = html & UiCampo("respon_envio", "Responsable del env&iacute;o", UiSelectRs("respon_envio", GetListaCombo(Cnn, "RESPONSABLES_FUNC"), rs(4), "", "", ""), False, False)
			html = html & UiFormGridCierra()

			html = html & UiGrupoAbre("Fechas")
			html = html & UiCampo("fecha_envio", "Env&iacute;o al cliente", UiInputTexto("fecha_envio", False, "autocomplete='off' placeholder='dd/mm/aaaa' value='" & fechaEnv & "'"), False, False)
			' Mismo criterio de antes para habilitar las fechas de instalacion
			If rs(5) & "" <> "" Then
				html = html & UiCampo("fecha_instalacion_test", "Instalaci&oacute;n en test", UiInputTexto("fecha_instalacion_test", False, "autocomplete='off' placeholder='dd/mm/aaaa' value='" & fechaTest & "'"), False, False)
			Else
				html = html & UiCampoLectura("Instalaci&oacute;n en test", UiDato(fechaTest), False)
			End If
			If rs(6) & "" <> "" Then
				html = html & UiCampo("fecha_instalacion_prod", "Instalaci&oacute;n en producci&oacute;n", UiInputTexto("fecha_instalacion_prod", False, "autocomplete='off' placeholder='dd/mm/aaaa' value='" & fechaProd & "'"), False, False)
			Else
				html = html & UiCampoLectura("Instalaci&oacute;n en producci&oacute;n", UiDato(fechaProd), False)
			End If
			html = html & UiGrupoCierra()
			html = html & UiCardCierra()

			html = html & UiCardAbre("Arreglos que implica", "")
			html = html & "<label class='ui-sr-only' for='motivo_apl'>Arreglos que implica</label>"
			html = html & "<textarea id='motivo_apl' name='motivo_apl' class='ui-input ui-textarea ui-apl-arreglos'>" & Server.HTMLEncode(rs(7) & "") & "</textarea>"
			html = html & UiCardCierra()

			'[UI-ANT] ' Relacionar con un incidente o proyecto: mismo dialogo de antes (armaCargaParam.asp + relaciones.js)
			'[UI-ANT] html = html & UiFormAcciones("<button type='button' id='sAgre' class='ui-button ui-button--secondary ui-button--lg'>+ Relacionar incidente o proyecto</button>" & _
			'[UI-ANT] UiBotonPrincipal("Actualizar APL", "actualizar('sForm');"))
			' [UI] Los incidentes y proyectos se relacionan desde la vista del APL (pop-up), como en incidentes y proyectos
			html = html & UiFormAcciones(UiBotonPrincipal("Actualizar APL", "actualizar('sForm');"))
			html = html & "</form>"
			html = html & "<div id='dialog' title=''></div>"
		End If
	End If
	html = html & UiPantallaCierra()

Else
	' [UI] Antes la pagina quedaba en blanco si el APL no existia o faltaba el id
	accTop = UiBoton("Volver", "secondary", "document.getElementById('form8').submit();", True, "")
	html = html & UiPantallaAbre("APL", accTop, "narrow")
	html = html & UiCardAbre("APL", "") & "<div class='ui-empty'>No se encontr&oacute; el APL solicitado.</div>" & UiCardCierra()
	html = html & UiPantallaCierra()
End If

html = html & "<script src='assets/js/relaciones.js'></script>"
html = html & "<script type='text/javascript'>"
html = html & " function actualizar(formu){ if (confirm('\u00bfActualizar el APL?')) { document.getElementById(formu).submit(); } }"
html = html & "</script>"

If opc = "act" And flg And puedeModif Then
	html = html & "<script>"
	html = html & "  $( function() {"
	html = html & "$.datepicker.regional['es'] = {"
	html = html & "closeText: 'Cerrar',"
	html = html & "prevText: '< Ant',"
	html = html & "nextText: 'Sig >',"
	html = html & "currentText: 'Hoy',"
	html = html & "monthNames: ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'],"
	html = html & "monthNamesShort: ['Ene','Feb','Mar','Abr', 'May','Jun','Jul','Ago','Sep', 'Oct','Nov','Dic'],"
	html = html & "dayNames: ['Domingo', 'Lunes', 'Martes', 'Mi\u00e9rcoles', 'Jueves', 'Viernes', 'S\u00e1bado'],"
	html = html & "dayNamesShort: ['Dom','Lun','Mar','Mi\u00e9','Jue','Vie','S\u00e1b'],"
	html = html & "dayNamesMin: ['Do','Lu','Ma','Mi','Ju','Vi','Sa'],"
	html = html & "weekHeader: 'Sm',"
	html = html & "dateFormat: 'dd/mm/yy',"
	html = html & "firstDay: 1,"
	html = html & "isRTL: false,"
	html = html & "showMonthAfterYear: false,"
	html = html & "yearSuffix: ''"
	html = html & "};"
	html = html & "$.datepicker.setDefaults($.datepicker.regional['es']);"
	html = html & "    $( ""#fecha_envio"" ).datepicker();"
	html = html & "    $( ""#fecha_instalacion_test"" ).datepicker();"
	html = html & "    $( ""#fecha_instalacion_prod"" ).datepicker();"
	html = html & "  } );"
	html = html & "$( '#dialog' ).dialog({"
	html = html & "title: 'Relacionar incidente o proyecto',"
	html = html & "autoOpen: false,"
	html = html & "resizable: false,"
	html = html & "buttons: {"
	html = html & "'Aceptar': function () {"
	html = html & "capturarYProcesar();"
	html = html & "$(this).dialog('close');"
	html = html & "},"
	html = html & "'Cerrar': function () {"
	html = html & "$(this).dialog('close');"
	html = html & "}"
	html = html & "}"
	html = html & "});"
	html = html & "$(document).on('click', '#sAgre', function() {"
	html = html & "    $.get('armaCargaParam.asp?param=IOP', function( data ) {"
	html = html & "        $( '#dialog' ).html(data);"
	html = html & "        if (typeof inicializarRelaciones === 'function') {"
	html = html & "            inicializarRelaciones();"
	html = html & "        }"
	html = html & "    });"
	html = html & "    $( '#dialog' ).dialog( 'open' );"
	html = html & "});"
	html = html & "</script>"
End If

html = html & form8 & form9 & form10 & "</body>"
html = html & "</html>"

response.write html


' =============================================================================
' FUNCIONES DE ESTA PANTALLA
' =============================================================================

Function AplMeta(etiqueta, valorHtml)
	AplMeta = "<div><span class='ui-topbar__label'>" & etiqueta & "</span><span class='ui-topbar__value'>" & valorHtml & "</span></div>"
End Function

' Hito del ciclo: cumplido si tiene fecha (dias desde 1900, mayor que cero)
Function AplHito(etiqueta, valorFecha)
	If IsNumeric(valorFecha & "") Then
		If CDbl(valorFecha) > 0 Then
			AplHito = "<div class='ui-lifecycle__item ui-lifecycle__item--done'><span class='ui-lifecycle__dot'></span><div><div class='ui-lifecycle__label'>" & etiqueta & "</div><div class='ui-lifecycle__value'>" & UiFecha(valorFecha) & "</div></div></div>"
			Exit Function
		End If
	End If
	AplHito = "<div class='ui-lifecycle__item ui-lifecycle__item--pend'><span class='ui-lifecycle__dot'></span><div><div class='ui-lifecycle__label'>" & etiqueta & "</div><div class='ui-lifecycle__value'>Pendiente</div></div></div>"
End Function

' Fecha para un campo editable: dd/mm/aaaa o vacia
Function AplFechaInput(v)
	AplFechaInput = ""
	If IsNumeric(v & "") Then
		If CDbl(v) > 0 Then AplFechaInput = UiFecha(v)
	End If
End Function

' Texto largo de la base: escapado y con sus saltos de linea
Function AplTexto(v)
	Dim t
	t = Server.HTMLEncode(Trim(v & ""))
	t = Replace(Replace(Replace(t, "&lt;br&gt;", "<br>", 1, -1, 1), "&lt;br/&gt;", "<br>", 1, -1, 1), "&lt;br /&gt;", "<br>", 1, -1, 1)
	AplTexto = Replace(Replace(Replace(t, vbCrLf, "<br>"), vbLf, "<br>"), vbCr, "<br>")
End Function

' Lista de incidentes (I) o proyectos (P) relacionados: abren su detalle (los mismos enlaces de antes)
Function AplRelacionados(nId, t)
	Dim r, h, n, url, etiqueta
	Set r = ObtenerRelacionesApl(Cnn, nId, t)
	h = ""
	n = 0
	Do While Not r.eof
		If t = "I" Then
			etiqueta = "<span class='ui-tag-id'>INC-" & r(2) & "</span>"
			url = "detalles.asp?opc=veri&id=" & r(2)
		Else
			etiqueta = "<span class='ui-tag-id ui-tag-id--pry'>PRY-" & r(2) & "</span>"
			url = "detalles.asp?opc=verp&id=" & r(2)
		End If
		url = url & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu
		h = h & "<li><a class='ui-relac__link' href='" & url & "'>" & etiqueta & " " & UiTexto(r(6)) & "</a></li>"
		n = n + 1
		r.movenext
	Loop
	If n = 0 Then
		If t = "I" Then AplRelacionados = "<p class='ui-status-empty'>Sin incidentes relacionados.</p>" Else AplRelacionados = "<p class='ui-status-empty'>Sin proyectos relacionados.</p>"
	Else
		AplRelacionados = "<ul class='ui-apl-rel'>" & h & "</ul>"
	End If
End Function

' [UI] Aviso despues de relacionar un incidente o proyecto (lo devuelve actualizar_v2.asp, opt=relcaso)
Function AplAvisoRelacion()
	Dim r, caso
	r = Request.QueryString("rel") & ""
	If Request.QueryString("reltipo") = "P" Then caso = "proyecto" Else caso = "incidente"
	caso = caso & " " & Server.HTMLEncode(Request.QueryString("relid") & "")
	AplAvisoRelacion = ""
	If r = "ok" Then AplAvisoRelacion = "<div class='ui-alert ui-alert--success' role='status'>Se relacion&oacute; el " & caso & ".</div>"
	If r = "ya" Then AplAvisoRelacion = "<div class='ui-alert ui-alert--warning' role='status'>El " & caso & " ya estaba relacionado con este APL.</div>"
	If r = "noexiste" Then AplAvisoRelacion = "<div class='ui-alert ui-alert--danger' role='alert'>No existe el " & caso & ".</div>"
	If r = "error" Then AplAvisoRelacion = "<div class='ui-alert ui-alert--danger' role='alert'>No se pudo relacionar: eleg&iacute; el tipo, complet&aacute; el n&uacute;mero y verific&aacute; tu permiso para modificar.</div>"
End Function

' [UI] Pop-up para relacionar un incidente o proyecto: solo el tipo y el numero. Envia a actualizar_v2.asp (opt=relcaso)
Function AplDialogoRelacion(nId)
	Dim h
	h = "<dialog id='dlgRel' class='ui-dialogo ui-dialogo--chico' aria-labelledby='dlgRelTitulo'>"
	h = h & "<form method='post' action='actualizar_v2.asp?tabla=apl&opt=relcaso&id=" & nId & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&cantregmover=0&cli=" & Server.URLEncode(cli & "") & "&aplic=" & Server.URLEncode(idAplic & "") & "&tip=" & Server.URLEncode(idtip & "") & "'>"
	h = h & "<div class='ui-dialogo__head'><h2 class='ui-dialogo__titulo' id='dlgRelTitulo'>Relacionar incidente o proyecto</h2><button type='button' class='ui-dialogo__cerrar' aria-label='Cerrar' onclick=""document.getElementById('dlgRel').close()"">&times;</button></div>"
	h = h & "<div class='ui-dialogo__body'>"
	h = h & "<div class='ui-field'><label class='ui-field__label' for='tipo_relac'>Tipo</label><select class='ui-select' id='tipo_relac' name='tipo_relac'><option value='I'>Incidente</option><option value='P'>Proyecto</option></select></div>"
	h = h & "<div class='ui-field'><label class='ui-field__label' for='id_caso'>N&uacute;mero <span class='ui-field__required' aria-hidden='true'> *</span></label>"
	h = h & "<input class='ui-input' id='id_caso' name='id_caso' type='number' min='1' step='1' inputmode='numeric' required aria-required='true' aria-describedby='id_caso_ayuda'>"
	h = h & "<p class='ui-dialogo__ayuda' id='id_caso_ayuda'>El incidente o proyecto se relaciona con este APL.</p></div>"
	h = h & "</div>"
	h = h & "<div class='ui-dialogo__acciones'><button type='button' class='ui-button ui-button--secondary' onclick=""document.getElementById('dlgRel').close()"">Cancelar</button><button type='submit' class='ui-button ui-button--primary'>Relacionar</button></div>"
	h = h & "</form></dialog>"
	h = h & "<script>function aplAbrirRel(){var d=document.getElementById('dlgRel');d.showModal();document.getElementById('tipo_relac').focus();}"
	h = h & "document.getElementById('dlgRel').addEventListener('close',function(){var b=document.getElementById('btnRel');if(b)b.focus();});"
	h = h & "document.getElementById('dlgRel').addEventListener('click',function(e){if(e.target===this)this.close();});</script>"
	AplDialogoRelacion = h
End Function
%>