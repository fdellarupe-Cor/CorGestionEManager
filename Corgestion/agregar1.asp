<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/ui_componentes.inc" -->
<%
' [CG] Control de usuario E-Manager (igual que CorEManager/Acciones.asp)
cgOk = ofv.CheckUsuario()
If cgOk <> "Y" Then
	cgUv = ofv.uv
	If cgUv = "NO" Then cgOk = ofv.ValidarUsuario
	Response.Write cgOk
	Response.End
End If



html = headhtml()
html = html & "<body>"

opc = request.querystring("opc")
tipo = request.querystring("tipo")
xmenu = request.querystring("xmenu")
xtmenu = request.querystring("xtmenu")
sBack = request.querystring("back")


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

If opc <> "ins" Then

	Select Case tipo

		Case "AP"

		tabla = "APLICATIVOS"
		tit = "APLICATIVO"
		volver = "aplicativos.asp?xmenu=" & xmenu & "_xtmenu=" & xtmenu

		Case "BR"

		tabla = "BROWSERS"
		tit = "BROWSER"
		volver = "browsers.asp?xmenu=" & xmenu & "_xtmenu=" & xtmenu

		Case "CL"

		tabla = "CLIENTES"
		tit = "CLIENTE"
		volver = "clientes.asp?xmenu=" & xmenu & "_xtmenu=" & xtmenu

		Case "ES"

		tabla = "ESTADOS"
		tit = "ESTADO"
		volver = "estados.asp?xmenu=" & xmenu & "_xtmenu=" & xtmenu

		Case "PR"

		tabla = "PRIORIDADES"
		tit = "PRIORIDAD"
		volver = "prioridades.asp?xmenu=" & xmenu & "_xtmenu=" & xtmenu

		Case "VP"

		tabla = "VIA_PEDIDO"
		tit = "VIA DE PEDIDO"
		volver = "viapedido.asp?xmenu=" & xmenu & "_xtmenu=" & xtmenu

		Case "CO"

		tabla = "CONTACTOS"
		tit = "CONTACTO"
		volver = "contactos.asp?xmenu=" & xmenu & "_xtmenu=" & xtmenu

		Case "CR"

		tabla = "COR_RESPONSABLES"
		tit = "RESPONSABLE"
		volver = "corresp.asp?xmenu=" & xmenu & "_xtmenu=" & xtmenu

		Case "TA"

		tabla = "TIPO_AVANCE"
		tit = "TIPO AVANCE"
		volver = "tipoavance.asp?xmenu=" & xmenu & "_xtmenu=" & xtmenu

	End Select

End If

If opc = "ins" Then

	tabla = request.querystring("tabla")

	If tabla = "CLIENTES" Then

		Dim cmdInsCli
		Set cmdInsCli = Server.CreateObject("ADODB.Command")
		With cmdInsCli
			.ActiveConnection = Cnn
			.CommandText = "sp_cg_Registro_Insertar"
			.CommandType = 4
			.Parameters.Append .CreateParameter("@Entidad", 200, 1, 20, "auxiliar")
			.Parameters.Append .CreateParameter("@Param", 200, 1, 30, "CLIENTE")
			.Parameters.Append .CreateParameter("@Dato1", 200, 1, 255, CStr(request.form("descri") & ""))
			.Execute
		End With
		Set cmdInsCli = Nothing

		Set rm = Exec(Cnn, "sp_cg_Utilidades @Modo='ultimo_id', @Tabla='clientes'")

		idx = ""
		If Not rm.eof Then idx = rm(0)

		Set rn = Exec(Cnn, "sp_cg_Utilidades @Modo='listar_id', @Tabla='aplicativos'")

		Do While Not rn.eof

			Dim cmdCA1
			Set cmdCA1 = Server.CreateObject("ADODB.Command")
			With cmdCA1
				.ActiveConnection = Cnn
				.CommandText = "sp_cg_ClienAplic"
				.CommandType = 4
				.Parameters.Append .CreateParameter("@Modo", 200, 1, 30, "insertar")
				.Parameters.Append .CreateParameter("@IdCliente", 3, 1, , CLng(idx))
				.Parameters.Append .CreateParameter("@IdAplic", 3, 1, , CLng(rn(0)))
				.Parameters.Append .CreateParameter("@Tiene", 200, 1, 10, CStr(request.form(rn(0) & "") & ""))
				.Execute
			End With
			Set cmdCA1 = Nothing

			rn.movenext
		Loop

	ElseIf tabla = "APLICATIVOS" Then

		Set rh = Exec(Cnn, "sp_cg_Utilidades @Modo='max_orden'")

		orden = ""
		If Not rh.eof Then orden = rh(0) + 10

		Dim cmdInsApl
		Set cmdInsApl = Server.CreateObject("ADODB.Command")
		With cmdInsApl
			.ActiveConnection = Cnn
			.CommandText = "sp_cg_Registro_Insertar"
			.CommandType = 4
			.Parameters.Append .CreateParameter("@Entidad", 200, 1, 20, "auxiliar")
			.Parameters.Append .CreateParameter("@Param", 200, 1, 30, "APLICACION")
			.Parameters.Append .CreateParameter("@Dato1", 200, 1, 255, CStr(request.form("descri") & ""))
			.Parameters.Append .CreateParameter("@Dato2", 200, 1, 255, CStr(orden))
			.Execute
		End With
		Set cmdInsApl = Nothing

		Set ra = Exec(Cnn, "sp_cg_Utilidades @Modo='ultimo_id', @Tabla='aplicativos'")

		id = ""
		If Not ra.eof Then id = ra(0)

		Set rl = Exec(Cnn, "sp_cg_Utilidades @Modo='listar_id', @Tabla='clientes'")

		Do While Not rl.eof

			Dim cmdCA2
			Set cmdCA2 = Server.CreateObject("ADODB.Command")
			With cmdCA2
				.ActiveConnection = Cnn
				.CommandText = "sp_cg_ClienAplic"
				.CommandType = 4
				.Parameters.Append .CreateParameter("@Modo", 200, 1, 30, "insertar")
				.Parameters.Append .CreateParameter("@IdCliente", 3, 1, , CLng(rl(0)))
				.Parameters.Append .CreateParameter("@IdAplic", 3, 1, , CLng(id))
				.Parameters.Append .CreateParameter("@Tiene", 200, 1, 10, CStr(request.form(rl(0) & "") & ""))
				.Execute
			End With
			Set cmdCA2 = Nothing

			rl.movenext
		Loop

	ElseIf tabla = "CONTACTOS" Then

		Dim cmdCont
		Set cmdCont = Server.CreateObject("ADODB.Command")
		With cmdCont
			.ActiveConnection = Cnn
			.CommandText = "sp_cg_Contactos"
			.CommandType = 4
			.Parameters.Append .CreateParameter("@Modo", 200, 1, 15, "insertar")
			.Parameters.Append .CreateParameter("@Nombre", 200, 1, 100, CStr(request.form("nom") & ""))
			.Parameters.Append .CreateParameter("@Apellido", 200, 1, 100, CStr(request.form("ape") & ""))
			.Parameters.Append .CreateParameter("@Telefono", 200, 1, 50, CStr(request.form("tel") & ""))
			.Parameters.Append .CreateParameter("@Mail", 200, 1, 100, CStr(request.form("mail") & ""))
			.Parameters.Append .CreateParameter("@IdCliente", 3, 1, , CLng(request.form("cli")))
			.Parameters.Append .CreateParameter("@Tipo", 3, 1, , CLng(request.form("tipo")))
			.Execute
		End With
		Set cmdCont = Nothing

	ElseIf tabla = "BROWSERS" or tabla = "VIA_PEDIDO" or tabla = "PRIORIDADES" or tabla = "ESTADOS" or tabla = "TIPO_AVANCE" Then

		If tabla = "ESTADOS" Then tParam = "ESTADO"
		If tabla = "BROWSERS" Then tParam = "BROWSER"
		If tabla = "VIA_PEDIDO" Then tParam = "VIA PEDIDO"
		If tabla = "PRIORIDADES" Then tParam = "PRIORIDAD"
		If tabla = "TIPO_AVANCE" Then tParam = "TIPO_AVANCE"

		Dim cmdInsAux
		Set cmdInsAux = Server.CreateObject("ADODB.Command")
		With cmdInsAux
			.ActiveConnection = Cnn
			.CommandText = "sp_cg_Registro_Insertar"
			.CommandType = 4
			.Parameters.Append .CreateParameter("@Entidad", 200, 1, 20, "auxiliar")
			.Parameters.Append .CreateParameter("@Param", 200, 1, 30, tParam)
			.Parameters.Append .CreateParameter("@Dato1", 200, 1, 255, CStr(request.form("descri") & ""))
			.Execute
		End With
		Set cmdInsAux = Nothing

	Else
		'corresp
		Dim cmdResp
		Set cmdResp = Server.CreateObject("ADODB.Command")
		With cmdResp
			.ActiveConnection = Cnn
			.CommandText = "sp_cg_Responsables"
			.CommandType = 4
			.Parameters.Append .CreateParameter("@Modo", 200, 1, 25, "insertar")
			.Parameters.Append .CreateParameter("@Nombre", 200, 1, 100, CStr(request.form("nom") & ""))
			.Parameters.Append .CreateParameter("@Apellido", 200, 1, 100, CStr(request.form("ape") & ""))
			.Parameters.Append .CreateParameter("@Funcion", 200, 1, 50, CStr(request.form("func") & ""))
			.Parameters.Append .CreateParameter("@Activo", 200, 1, 10, "S")
			.Parameters.Append .CreateParameter("@IdSeg", 200, 1, 100, CStr(request.form("LOGUIN") & ""))
			.Execute
		End With
		Set cmdResp = Nothing

		set rs = Cnn.Execute("sp_cg_Responsables @Modo='ultimo_id_funcion'")

		if not rs.eof then

			'iSql = "select id,tipo from CG_MENUES"
			IF rs(1)="TEC" THEN
			set rx = Cnn.Execute("sp_cg_MenuUsuario @Modo='plantilla', @FuncUsr='TEC'")
			ELSE
			set rx = Cnn.Execute("sp_cg_MenuUsuario @Modo='plantilla', @FuncUsr='FUNC'")
			END IF

			do while not rx.eof

			ExecNonQuery Cnn, "sp_cg_MenuUsuario @Modo='insertar', @IdMenu=" & rx(0) & ", @TipoMenu='" & rx(1) & "', @IdUsr=" & SqlStr(rs(0)) & ", @FuncUsr='" & rs(1) & "', @Dato1=" & rx(2) & ", @Dato2=" & rx(3) & ", @Dato3=" & rx(4) & ", @Dato4=" & rx(5) & ", @Dato5=" & rx(6)

			rx.movenext
			loop

		end if


	End If

	' [LOG] Alta de maestro (CG_LOG): el ultimo ID de la tabla recien cargada
	Select Case UCase(tabla & "")
		Case "CLIENTES": CgLogUltimo "CLIENTE", "CG_CLIENTES", "ALTA"
		Case "APLICATIVOS": CgLogUltimo "APLICATIVO", "CG_APLICATIVOS", "ALTA"
		Case "CONTACTOS": CgLogUltimo "CONTACTO", "CG_CONTACTOS", "ALTA"
		Case "BROWSERS": CgLogUltimo "BROWSER", "CG_BROWSERS", "ALTA"
		Case "VIA_PEDIDO": CgLogUltimo "VIA_PEDIDO", "CG_VIA_PEDIDO", "ALTA"
		Case "PRIORIDADES": CgLogUltimo "PRIORIDAD", "CG_PRIORIDADES", "ALTA"
		Case "ESTADOS": CgLogUltimo "ESTADO", "CG_ESTADOS", "ALTA"
		Case "TIPO_AVANCE": CgLogUltimo "TIPO_AVANCE", "CG_TIPO_AVANCE", "ALTA"
		Case "COR_RESPONSABLES": CgLogUltimo "RESPONSABLE", "CG_COR_RESPONSABLES", "ALTA"
	End Select

End If

'[UI-ANT] ' ===== Vista anterior completa: reemplazada por el bloque [UI] que sigue. Se borra entera buscando [UI-ANT] =====
'[UI-ANT] html = html & "<div>"

'[UI-ANT] html = html & cabecera()

'[UI-ANT] html = html & "<div class='titulo-seccion page-toolbar'>"

'[UI-ANT] html = html & "<div><h2>AGREGAR " & tit & "</h2></div>"
'[UI-ANT] formx5 = "<form id='formx5' style='height:100%;' method='post' action='" & replace(volver,"_","&") & "'></form>"
'[UI-ANT] botonx5 = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('formx5').submit();"">"
'[UI-ANT] html = html & "<div class='page-toolbar-action'>" & botonx5 & "</div>"

'[UI-ANT] html = html & "</div>"

'[UI-ANT] html = html & "</div>"

'[UI-ANT] If tipo <> "CO" And tipo <> "CR" Then

	'[UI-ANT] 'html = html & "<div style='height:70%;width:100%;'>"
	'[UI-ANT] html = html & "<div class='form-shell'>"

	'[UI-ANT] 'html = html & "<form id='agre' method='post' action='agregar1.asp?opc=ins&tabla=" & tabla & "&volver=" & volver & "&back=" & sBack & "'>"
	'[UI-ANT] html = html & "<form id='agre' method='post' action='agregar1.asp?opc=ins&tabla=" & tabla & "&volver=" & volver & "&back=" & sBack & "'>"

	'[UI-ANT] 'html = html & "<table class='tabl' style='width:40%;' align=center cellspacing=0 cellpadding=3>"
	'[UI-ANT] html = html & "<div class='bloque-datos ui-panel'>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' style='font-size:12pt;' colspan=2 align=center><u><b>COMPLETE LOS SIGUIENTES DATOS</b></u></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='bloque-header'>COMPLETE LOS SIGUIENTES DATOS</div>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td colspan=2>&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"

	'[UI-ANT] html = html & "<div class='bloque-body ui-panel__body'>"

	'[UI-ANT] ' --- INICIO GRILLA 2 COLUMNAS ---
	'[UI-ANT] html = html & "<div class='grid-datos-doble ui-field-grid--dual'>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil'>DESCRIPCION:</td>"
	'[UI-ANT] 'html = html & "<td class='fil'><input class='inpu' style='width:90%;' id='descri' name='descri' type='text'></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>DESCRIPCION</div>"
	'[UI-ANT] html = html & "<div><input id='descri' name='descri' type='text' class='input-moderno ui-input'></div>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td colspan=2>&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"

	'[UI-ANT] If tipo <> "AP" And tipo <> "BR" And tipo <> "ES" And tipo <> "PR" And tipo <> "VP" Then

		'[UI-ANT] Set rf = Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='aplicativos'")

		'[UI-ANT] Do While Not rf.eof

			'[UI-ANT] 'html = html & "<tr>"
			'[UI-ANT] 'html = html & "<td class='fil'>TIENE " & rf(1) & "</td>"
			'[UI-ANT] 'html = html & "<td class='fil'><select class='selec' style='width:30%;' id='" & rf(0) & "' name='" & rf(0) & "'>"
			'[UI-ANT] 'html = html & "<option value='1'>SI</option>"
			'[UI-ANT] 'html = html & "<option value='0'>NO</option>"
			'[UI-ANT] 'html = html & "</select></td>"
			'[UI-ANT] 'html = html & "</tr>"
			'[UI-ANT] html = html & "<div class='dato-label'>TIENE " & Server.HTMLEncode(UCase(rf(1) & "")) & "</div>"
			'[UI-ANT] html = html & "<div class='dato-value'><select class='ui-select' id='" & rf(0) & "' name='" & rf(0) & "'>"
			'[UI-ANT] html = html & "<option value='1'>SI</option>"
			'[UI-ANT] html = html & "<option value='0'>NO</option>"
			'[UI-ANT] html = html & "</select></div>"

			'[UI-ANT] rf.movenext
		'[UI-ANT] Loop

		'[UI-ANT] 'html = html & "<tr>"
		'[UI-ANT] 'html = html & "<td colspan=2>&nbsp;</td>"
		'[UI-ANT] 'html = html & "</tr>"

	'[UI-ANT] End If

	'[UI-ANT] If tipo = "AP" Then

		'[UI-ANT] Set rf = Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='clientes'")

		'[UI-ANT] Do While Not rf.eof

			'[UI-ANT] 'html = html & "<tr>"
			'[UI-ANT] 'html = html & "<td class='fil'>TIENE " & rf(1) & "</td>"
			'[UI-ANT] 'html = html & "<td class='fil'><select class='selec' style='width:30%;' id='" & rf(0) & "' name='" & rf(0) & "'>"
			'[UI-ANT] 'html = html & "<option value='1'>SI</option>"
			'[UI-ANT] 'html = html & "<option value='0'>NO</option>"
			'[UI-ANT] 'html = html & "</select></td>"
			'[UI-ANT] 'html = html & "</tr>"
			'[UI-ANT] html = html & "<div class='dato-label'>TIENE " & Server.HTMLEncode(UCase(rf(1) & "")) & "</div>"
			'[UI-ANT] html = html & "<div class='dato-value'><select class='ui-select' id='" & rf(0) & "' name='" & rf(0) & "'>"
			'[UI-ANT] html = html & "<option value='1'>SI</option>"
			'[UI-ANT] html = html & "<option value='0'>NO</option>"
			'[UI-ANT] html = html & "</select></div>"

			'[UI-ANT] rf.movenext
		'[UI-ANT] Loop

		'[UI-ANT] 'html = html & "<tr>"
		'[UI-ANT] 'html = html & "<td colspan=2>&nbsp;</td>"
		'[UI-ANT] 'html = html & "</tr>"

	'[UI-ANT] End If

'[UI-ANT] ' --- FIN GRILLA 2 COLUMNAS ---
'[UI-ANT] html = html & "</div>"

'[UI-ANT] 'html = html & "<tr>"
'[UI-ANT] 'html = html & "<td colspan=2 align=center><input class='bot' type='button' value='Agregar' onclick='agregar();' ></td>"
'[UI-ANT] 'html = html & "</tr>"
'[UI-ANT] html = html & "<div style='text-align: center; margin-bottom: 1rem; margin-top: 1rem;'>"
'[UI-ANT] html = html & "<input type='button' class='btn-principal ui-button ui-button--primary' value='Agregar' onclick='agregar();'>"
'[UI-ANT] html = html & "</div>"

'[UI-ANT] 'html = html & "</table>"
'[UI-ANT] html = html & "</div>" ' Fin .bloque-body
'[UI-ANT] html = html & "</div>" ' Fin .bloque-datos

'[UI-ANT] 'html = html & "</form>"
'[UI-ANT] html = html & "</form>"

'[UI-ANT] 'html = html & "</div>"
'[UI-ANT] html = html & "</div>" ' Fin max-width

'[UI-ANT] ElseIf tipo = "CO" Then

	'[UI-ANT] 'html = html & "<div style='height:57%;width:100%;'><br><br>"
	'[UI-ANT] html = html & "<div class='form-shell'>"

	'[UI-ANT] 'html = html & "<form id='agre' method='post' action='agregar1.asp?opc=ins&tabla=" & tabla & "&volver=" & volver & "&back=" & sBack & "'>"
	'[UI-ANT] html = html & "<form id='agre' method='post' action='agregar1.asp?opc=ins&tabla=" & tabla & "&volver=" & volver & "&back=" & sBack & "'>"

	'[UI-ANT] 'html = html & "<table class='tabl' style='width:40%;' align=center cellspacing=0 cellpadding=3>"
	'[UI-ANT] html = html & "<div class='bloque-datos ui-panel'>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' style='font-size:12pt;' colspan=2 align=center><b><u>COMPLETE LOS SIGUIENTES DATOS</u></b></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='bloque-header'>COMPLETE LOS SIGUIENTES DATOS</div>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td colspan=2>&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"

	'[UI-ANT] html = html & "<div class='bloque-body ui-panel__body'>"

	'[UI-ANT] ' --- INICIO GRILLA 2 COLUMNAS ---
	'[UI-ANT] html = html & "<div class='grid-datos-doble ui-field-grid--dual'>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil'>NOMBRE:</td>"
	'[UI-ANT] 'html = html & "<td class='fil'><input class='inpu' style='width:80%;' id='nom' name='nom' type='text'></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>NOMBRE</div>"
	'[UI-ANT] html = html & "<div><input id='nom' name='nom' type='text' class='input-moderno ui-input'></div>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil'>APELLIDO:</td>"
	'[UI-ANT] 'html = html & "<td class='fil'><input class='inpu' style='width:80%;' id='ape' name='ape' type='text'></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>APELLIDO</div>"
	'[UI-ANT] html = html & "<div><input id='ape' name='ape' type='text' class='input-moderno ui-input'></div>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil'>TELEFONO:</td>"
	'[UI-ANT] 'html = html & "<td class='fil'><input class='inpu' style='width:80%;' id='tel' name='tel' type='text'></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>TELEFONO</div>"
	'[UI-ANT] html = html & "<div><input id='tel' name='tel' type='text' class='input-moderno ui-input'></div>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil'>MAIL:</td>"
	'[UI-ANT] 'html = html & "<td class='fil'><input class='inpu' style='width:80%;' id='mail' name='mail' type='text'></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>MAIL</div>"
	'[UI-ANT] html = html & "<div><input id='mail' name='mail' type='text' class='input-moderno ui-input'></div>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil'>CLIENTE:</td>"
	'[UI-ANT] 'html = html & "<td class='fil'><select class='selec' style='width:80%;' id='cli' name='cli'>"
	'[UI-ANT] html = html & "<div class='dato-label'>CLIENTE</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><select class='ui-select' id='cli' name='cli'>"

		'[UI-ANT] Set Cnn = Server.CreateObject("ADODB.Connection")
		'[UI-ANT] Cnn.open "[credenciales eliminadas]"

		'[UI-ANT] Set rs = Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='clientes'")

		'[UI-ANT] Do While Not rs.eof
			'[UI-ANT] html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
			'[UI-ANT] rs.movenext
		'[UI-ANT] Loop

	'[UI-ANT] 'html = html & "</select></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "</select></div>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil'>TIPO:</td>"
	'[UI-ANT] 'html = html & "<td class='fil'><select class='selec' style='width:80%;' id='tipo' name='tipo'>"
	'[UI-ANT] 'html = html & "<option value='1'>generico</option>"
	'[UI-ANT] 'html = html & "<option value='10'>normal</option>"
	'[UI-ANT] 'html = html & "</select></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>TIPO</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><select class='ui-select' id='tipo' name='tipo'>"
	'[UI-ANT] html = html & "<option value='1'>generico</option>"
	'[UI-ANT] html = html & "<option value='10'>normal</option>"
	'[UI-ANT] html = html & "</select></div>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td colspan=2>&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"

	'[UI-ANT] ' --- FIN GRILLA 2 COLUMNAS ---
	'[UI-ANT] html = html & "</div>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td colspan=2 align=center><input class='bot' type='button' value='Agregar' onclick='agregar1();' ></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div style='text-align: center; margin-bottom: 1rem; margin-top: 1rem;'>"
	'[UI-ANT] html = html & "<input type='button' class='btn-principal ui-button ui-button--primary' value='Agregar' onclick='agregar1();'>"
	'[UI-ANT] html = html & "</div>"

	'[UI-ANT] 'html = html & "</table>"
	'[UI-ANT] html = html & "</div>" ' Fin .bloque-body
	'[UI-ANT] html = html & "</div>" ' Fin .bloque-datos

	'[UI-ANT] 'html = html & "</form>"
	'[UI-ANT] html = html & "</form>"

	'[UI-ANT] 'html = html & "</div>"
	'[UI-ANT] html = html & "</div>" ' Fin max-width

'[UI-ANT] Else

	'[UI-ANT] 'html = html & "<div style='height:57%;width:100%;'><br><br>"
	'[UI-ANT] html = html & "<div class='form-shell'>"

	'[UI-ANT] 'html = html & "<form id='agre' method='post' action='agregar1.asp?opc=ins&tabla=" & tabla & "&volver=" & volver & "&back=" & sBack & "'>"
	'[UI-ANT] html = html & "<form id='agre' method='post' action='agregar1.asp?opc=ins&tabla=" & tabla & "&volver=" & volver & "&back=" & sBack & "'>"

	'[UI-ANT] 'html = html & "<table class='tabl' style='width:40%;' align=center cellspacing=0 cellpadding=3>"
	'[UI-ANT] html = html & "<div class='bloque-datos ui-panel'>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' style='font-size:12pt;' colspan=2 align=center><b><u>COMPLETE LOS SIGUIENTES DATOS</u></b></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='bloque-header'>COMPLETE LOS SIGUIENTES DATOS</div>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td colspan=2>&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"

	'[UI-ANT] html = html & "<div class='bloque-body ui-panel__body'>"

	'[UI-ANT] ' --- INICIO GRILLA 2 COLUMNAS ---
	'[UI-ANT] html = html & "<div class='grid-datos-doble ui-field-grid--dual'>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil'>NOMBRE:</td>"
	'[UI-ANT] 'html = html & "<td class='fil'><input class='inpu' style='width:80%;' id='nom' name='nom' type='text'></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>NOMBRE</div>"
	'[UI-ANT] html = html & "<div><input id='nom' name='nom' type='text' class='input-moderno ui-input'></div>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil'>APELLIDO:</td>"
	'[UI-ANT] 'html = html & "<td class='fil'><input class='inpu' style='width:80%;' id='ape' name='ape' type='text'></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>APELLIDO</div>"
	'[UI-ANT] html = html & "<div><input id='ape' name='ape' type='text' class='input-moderno ui-input'></div>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil'>FUNCION:</td>"
	'[UI-ANT] 'html = html & "<td class='fil'><select class='selec' style='width:80%;' id='func' name='func'>"
	'[UI-ANT] 'html = html & "<option value='TEC'>TECNICO</option>"
	'[UI-ANT] 'html = html & "<option value='FUNC'>FUNCIONAL</option>"
	'[UI-ANT] 'html = html & "</select></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>FUNCION</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><select class='ui-select' id='func' name='func'>"
	'[UI-ANT] html = html & "<option value='TEC'>TECNICO</option>"
	'[UI-ANT] html = html & "<option value='FUNC'>FUNCIONAL</option>"
	'[UI-ANT] html = html & "</select></div>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil'>LOGUIN:</td>"
	'[UI-ANT] 'html = html & "<td class='fil'><input class='inpu' style='width:80%;' id='LOGUIN' name='LOGUIN' type='text'></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>LOGUIN</div>"
	'[UI-ANT] html = html & "<div><input id='LOGUIN' name='LOGUIN' type='text' class='input-moderno ui-input'></div>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td colspan=2>&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"

	'[UI-ANT] ' --- FIN GRILLA 2 COLUMNAS ---
	'[UI-ANT] html = html & "</div>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' colspan=2 align=center><input class='bot' type='button' value='Agregar' onclick='agregar1();' ></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div style='text-align: center; margin-bottom: 1rem; margin-top: 1rem;'>"
	'[UI-ANT] html = html & "<input type='button' class='btn-principal ui-button ui-button--primary' value='Agregar' onclick='agregar1();'>"
	'[UI-ANT] html = html & "</div>"

	'[UI-ANT] 'html = html & "</table>"
	'[UI-ANT] html = html & "</div>" ' Fin .bloque-body
	'[UI-ANT] html = html & "</div>" ' Fin .bloque-datos

	'[UI-ANT] 'html = html & "</form>"
	'[UI-ANT] html = html & "</form>"

	'[UI-ANT] 'html = html & "</div>"
	'[UI-ANT] html = html & "</div>" ' Fin max-width

'[UI-ANT] End If


'[UI-ANT] html = html & "<script src='assets/js/jquery.min.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/skel.min.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/util.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/main.js'></script>"

'[UI-ANT] html = html & "<script type='text/javascript'>"
'[UI-ANT] html = html & " function agregar(){ "
'[UI-ANT] html = html & " var des = document.getElementById('descri').value;"
'[UI-ANT] html = html & " if (des == ''){ "
'[UI-ANT] html = html & " alert('Descripcion no debe ser nula'); return false;} "
'[UI-ANT] html = html & " else{ "
'[UI-ANT] html = html & " document.getElementById('agre').submit();"
'[UI-ANT] html = html & " }} "
'[UI-ANT] html = html & "</script>"

'[UI-ANT] html = html & "<script type='text/javascript'>"
'[UI-ANT] html = html & " function agregar1(){ "
'[UI-ANT] html = html & " var nom = document.getElementById('nom').value;"
'[UI-ANT] html = html & " var ape = document.getElementById('ape').value;"
'[UI-ANT] html = html & " if (nom == '' || ape == ''){ "
'[UI-ANT] html = html & " alert('Nombre y apellido no deben ser nulo'); return false; "
'[UI-ANT] html = html & " } "
'[UI-ANT] html = html & " document.getElementById('agre').submit();"
'[UI-ANT] html = html & " } "
'[UI-ANT] html = html & "</script>"

' =========================================================================
' [UI] VISTA con el sistema visual (IncFile/ui_componentes.inc)
' El bloque anterior completo queda arriba, comentado con '[UI-ANT].
' Ids y nombres de los campos, formulario y validaciones: los mismos de antes.
' =========================================================================
Dim tituloPantalla, tituloDatos, accTop, anchoPantalla, botonAgregar
Select Case tipo
	Case "AP"
		tituloPantalla = "Nuevo aplicativo"
		tituloDatos = "Datos del aplicativo"
	Case "BR"
		tituloPantalla = "Nuevo browser"
		tituloDatos = "Datos del browser"
	Case "CL"
		tituloPantalla = "Nuevo cliente"
		tituloDatos = "Datos del cliente"
	Case "ES"
		tituloPantalla = "Nuevo estado"
		tituloDatos = "Datos del estado"
	Case "PR"
		tituloPantalla = "Nueva prioridad"
		tituloDatos = "Datos de la prioridad"
	Case "VP"
		tituloPantalla = "Nueva v&iacute;a de pedido"
		tituloDatos = "Datos de la v&iacute;a de pedido"
	Case "TA"
		tituloPantalla = "Nuevo tipo de avance"
		tituloDatos = "Datos del tipo de avance"
	Case "CO"
		tituloPantalla = "Nuevo contacto"
		tituloDatos = "Datos del contacto"
	Case "CR"
		tituloPantalla = "Nuevo responsable"
		tituloDatos = "Datos del responsable"
	Case Else
		tituloPantalla = "Agregar " & UiCapitalizar(tit)
		tituloDatos = "Datos"
End Select

' Catalogos de un solo dato: pantalla angosta. Cliente, aplicativo y contacto: ancho normal
anchoPantalla = "narrow"
If tipo = "AP" Or tipo = "CL" Or tipo = "CO" Then anchoPantalla = ""

html = html & "<div>"
html = html & cabecera()
html = html & "</div>"

formx5 = "<form id='formx5' class='form-oculto' method='post' action='" & replace(volver,"_","&") & "'></form>"
accTop = UiBoton("Volver", "secondary", "document.getElementById('formx5').submit();", True, "")
html = html & UiPantallaAbre(tituloPantalla, accTop, anchoPantalla)

html = html & "<form id='agre' method='post' action='agregar1.asp?opc=ins&tabla=" & tabla & "&volver=" & volver & "&back=" & sBack & "'>"
html = html & UiCardAbre(tituloDatos, "") & UiNotaObligatorios() & UiFormGridAbre()

If tipo <> "CO" And tipo <> "CR" Then

	html = html & UiCampo("descri", "Descripci&oacute;n", UiInputTexto("descri", True, "autocomplete='off'"), True, True)
	html = html & UiFormGridCierra() & UiCardCierra()

	' Cliente: que aplicativos tiene (Si / No por cada uno)
	If tipo <> "AP" And tipo <> "BR" And tipo <> "ES" And tipo <> "PR" And tipo <> "VP" And tipo <> "TA" Then
		Set rf = Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='aplicativos'")
		html = html & UiCardAbre("Aplicativos del cliente", "") & "<div class='ui-form-grid ui-form-grid--4'>"
		Do While Not rf.eof
			html = html & UiCampo(rf(0), Server.HTMLEncode(rf(1) & ""), UiSelectOpciones(rf(0), "1|S&iacute;||0|No", "1", ""), False, False)
			rf.movenext
		Loop
		html = html & "</div>" & UiCardCierra()
	End If

	' Aplicativo: que clientes lo tienen (Si / No por cada uno)
	If tipo = "AP" Then
		Set rf = Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='clientes'")
		html = html & UiCardAbre("Clientes que lo tienen", "") & "<div class='ui-form-grid ui-form-grid--4'>"
		Do While Not rf.eof
			html = html & UiCampo(rf(0), Server.HTMLEncode(rf(1) & ""), UiSelectOpciones(rf(0), "1|S&iacute;||0|No", "1", ""), False, False)
			rf.movenext
		Loop
		html = html & "</div>" & UiCardCierra()
	End If

	botonAgregar = UiBotonPrincipal("Agregar", "agregar();")

ElseIf tipo = "CO" Then

	html = html & UiCampo("nom", "Nombre", UiInputTexto("nom", True, "autocomplete='off'"), True, False)
	html = html & UiCampo("ape", "Apellido", UiInputTexto("ape", True, "autocomplete='off'"), True, False)
	html = html & UiCampo("tel", "Tel&eacute;fono", UiInputTexto("tel", False, "autocomplete='off'"), False, False)
	html = html & UiCampo("mail", "Mail", UiInputTexto("mail", False, "autocomplete='off'"), False, False)
	' Usa la conexion abierta al inicio de la pagina (antes se reabria con usuario y contrasena escritos en el codigo)
	Set rs = Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='clientes'")
	html = html & UiCampo("cli", "Cliente", UiSelectRs("cli", rs, "", "", "", ""), False, False)
	html = html & UiCampo("tipo", "Tipo", UiSelectOpciones("tipo", "1|Gen&eacute;rico||10|Normal", "1", ""), False, False)
	html = html & UiFormGridCierra() & UiCardCierra()

	botonAgregar = UiBotonPrincipal("Agregar", "agregar1();")

Else
	' Responsable (tipo CR)
	html = html & UiCampo("nom", "Nombre", UiInputTexto("nom", True, "autocomplete='off'"), True, False)
	html = html & UiCampo("ape", "Apellido", UiInputTexto("ape", True, "autocomplete='off'"), True, False)
	html = html & UiCampo("func", "Funci&oacute;n", UiSelectOpciones("func", "TEC|T&eacute;cnico||FUNC|Funcional", "TEC", ""), False, False)
	html = html & UiCampo("LOGUIN", "Usuario (login)", UiInputTexto("LOGUIN", False, "autocomplete='off'"), False, False)
	html = html & UiFormGridCierra() & UiCardCierra()

	botonAgregar = UiBotonPrincipal("Agregar", "agregar1();")

End If

html = html & UiFormAcciones(botonAgregar)
html = html & "</form>"
html = html & UiPantallaCierra()

html = html & "<script src='assets/js/jquery.min.js'></script>"
html = html & "<script src='assets/js/skel.min.js'></script>"
html = html & "<script src='assets/js/util.js'></script>"
html = html & "<script src='assets/js/main.js'></script>"

html = html & "<script type='text/javascript'>"
html = html & " function agregar(){ "
html = html & " var des = document.getElementById('descri');"
html = html & " if (des.value == ''){ "
html = html & " alert('Complet\u00e1 la descripci\u00f3n.'); des.focus(); return false;} "
html = html & " else{ "
html = html & " document.getElementById('agre').submit();"
html = html & " }} "
html = html & "</script>"

html = html & "<script type='text/javascript'>"
html = html & " function agregar1(){ "
html = html & " var nom = document.getElementById('nom');"
html = html & " var ape = document.getElementById('ape');"
html = html & " if (nom.value == '' || ape.value == ''){ "
html = html & " alert('Complet\u00e1 el nombre y el apellido.'); (nom.value == '' ? nom : ape).focus(); return false; "
html = html & " } "
html = html & " document.getElementById('agre').submit();"
html = html & " } "
html = html & "</script>"
'response.write volver & "--"
If opc = "ins" Then

	vuelta = request.querystring("volver")
	vuelta = replace(vuelta,"_","&")
	if sBack = "si" then vuelta = "carga1.asp?opc=noact"

	html = html & "<FORM name='init' method='post'  ACTION='" & vuelta & "' >"
	html = html & "<INPUT TYPE='SUBMIT' VALUE='INICIAR' style='display:none;' >"
	html = html & "</FORM>"

	html = html & "<script type='text/javascript'>"
	html = html & " alert('Agregado exitosamente.');"
	html = html & " document.all.init.submit();"
	html = html & "</script>"

End If

html = html & formx5 & "</body>"
html = html & "</html>"

response.write html
%>