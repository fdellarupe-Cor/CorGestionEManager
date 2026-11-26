<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<%
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->


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

End If

html = html & "<div>"

html = html & cabecera()

html = html & "<div class='titulo-seccion page-toolbar'>"

html = html & "<div><h2>AGREGAR " & tit & "</h2></div>"
formx5 = "<form id='formx5' style='height:100%;' method='post' action='" & replace(volver,"_","&") & "'></form>"
botonx5 = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('formx5').submit();"">"
html = html & "<div class='page-toolbar-action'>" & botonx5 & "</div>"

html = html & "</div>"

html = html & "</div>"

If tipo <> "CO" And tipo <> "CR" Then

	'html = html & "<div style='height:70%;width:100%;'>"
	html = html & "<div class='form-shell'>"

	'html = html & "<form id='agre' method='post' action='agregar1.asp?opc=ins&tabla=" & tabla & "&volver=" & volver & "&back=" & sBack & "'>"
	html = html & "<form id='agre' method='post' action='agregar1.asp?opc=ins&tabla=" & tabla & "&volver=" & volver & "&back=" & sBack & "'>"

	'html = html & "<table class='tabl' style='width:40%;' align=center cellspacing=0 cellpadding=3>"
	html = html & "<div class='bloque-datos ui-panel'>"

	'html = html & "<tr>"
	'html = html & "<td class='fil' style='font-size:12pt;' colspan=2 align=center><u><b>COMPLETE LOS SIGUIENTES DATOS</b></u></td>"
	'html = html & "</tr>"
	html = html & "<div class='bloque-header'>COMPLETE LOS SIGUIENTES DATOS</div>"

	'html = html & "<tr>"
	'html = html & "<td colspan=2>&nbsp;</td>"
	'html = html & "</tr>"

	html = html & "<div class='bloque-body ui-panel__body'>"

	' --- INICIO GRILLA 2 COLUMNAS ---
	html = html & "<div class='grid-datos-doble ui-field-grid--dual'>"

	'html = html & "<tr>"
	'html = html & "<td class='fil'>DESCRIPCION:</td>"
	'html = html & "<td class='fil'><input class='inpu' style='width:90%;' id='descri' name='descri' type='text'></td>"
	'html = html & "</tr>"
	html = html & "<div class='dato-label'>DESCRIPCION</div>"
	html = html & "<div><input id='descri' name='descri' type='text' class='input-moderno ui-input'></div>"

	'html = html & "<tr>"
	'html = html & "<td colspan=2>&nbsp;</td>"
	'html = html & "</tr>"

	If tipo <> "AP" And tipo <> "BR" And tipo <> "ES" And tipo <> "PR" And tipo <> "VP" And tipo <> "TA"  Then

		Set rf = Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='aplicativos'")

		Do While Not rf.eof

			'html = html & "<tr>"
			'html = html & "<td class='fil'>TIENE " & rf(1) & "</td>"
			'html = html & "<td class='fil'><select class='selec' style='width:30%;' id='" & rf(0) & "' name='" & rf(0) & "'>"
			'html = html & "<option value='1'>SI</option>"
			'html = html & "<option value='0'>NO</option>"
			'html = html & "</select></td>"
			'html = html & "</tr>"
			html = html & "<div class='dato-label'>TIENE " & Server.HTMLEncode(UCase(rf(1) & "")) & "</div>"
			html = html & "<div class='dato-value'><select class='ui-select' id='" & rf(0) & "' name='" & rf(0) & "'>"
			html = html & "<option value='1'>SI</option>"
			html = html & "<option value='0'>NO</option>"
			html = html & "</select></div>"

			rf.movenext
		Loop

		'html = html & "<tr>"
		'html = html & "<td colspan=2>&nbsp;</td>"
		'html = html & "</tr>"

	End If

	If tipo = "AP" Then

		Set rf = Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='clientes'")

		Do While Not rf.eof

			'html = html & "<tr>"
			'html = html & "<td class='fil'>TIENE " & rf(1) & "</td>"
			'html = html & "<td class='fil'><select class='selec' style='width:30%;' id='" & rf(0) & "' name='" & rf(0) & "'>"
			'html = html & "<option value='1'>SI</option>"
			'html = html & "<option value='0'>NO</option>"
			'html = html & "</select></td>"
			'html = html & "</tr>"
			html = html & "<div class='dato-label'>TIENE " & Server.HTMLEncode(UCase(rf(1) & "")) & "</div>"
			html = html & "<div class='dato-value'><select class='ui-select' id='" & rf(0) & "' name='" & rf(0) & "'>"
			html = html & "<option value='1'>SI</option>"
			html = html & "<option value='0'>NO</option>"
			html = html & "</select></div>"

			rf.movenext
		Loop

		'html = html & "<tr>"
		'html = html & "<td colspan=2>&nbsp;</td>"
		'html = html & "</tr>"

	End If

' --- FIN GRILLA 2 COLUMNAS ---
html = html & "</div>"

'html = html & "<tr>"
'html = html & "<td colspan=2 align=center><input class='bot' type='button' value='Agregar' onclick='agregar();' ></td>"
'html = html & "</tr>"
html = html & "<div style='text-align: center; margin-bottom: 1rem; margin-top: 1rem;'>"
html = html & "<input type='button' class='btn-principal ui-button ui-button--primary' value='Agregar' onclick='agregar();'>"
html = html & "</div>"

'html = html & "</table>"
html = html & "</div>" ' Fin .bloque-body
html = html & "</div>" ' Fin .bloque-datos

'html = html & "</form>"
html = html & "</form>"

'html = html & "</div>"
html = html & "</div>" ' Fin max-width

ElseIf tipo = "CO" Then

	'html = html & "<div style='height:57%;width:100%;'><br><br>"
	html = html & "<div class='form-shell'>"

	'html = html & "<form id='agre' method='post' action='agregar1.asp?opc=ins&tabla=" & tabla & "&volver=" & volver & "&back=" & sBack & "'>"
	html = html & "<form id='agre' method='post' action='agregar1.asp?opc=ins&tabla=" & tabla & "&volver=" & volver & "&back=" & sBack & "'>"

	'html = html & "<table class='tabl' style='width:40%;' align=center cellspacing=0 cellpadding=3>"
	html = html & "<div class='bloque-datos ui-panel'>"

	'html = html & "<tr>"
	'html = html & "<td class='fil' style='font-size:12pt;' colspan=2 align=center><b><u>COMPLETE LOS SIGUIENTES DATOS</u></b></td>"
	'html = html & "</tr>"
	html = html & "<div class='bloque-header'>COMPLETE LOS SIGUIENTES DATOS</div>"

	'html = html & "<tr>"
	'html = html & "<td colspan=2>&nbsp;</td>"
	'html = html & "</tr>"

	html = html & "<div class='bloque-body ui-panel__body'>"

	' --- INICIO GRILLA 2 COLUMNAS ---
	html = html & "<div class='grid-datos-doble ui-field-grid--dual'>"

	'html = html & "<tr>"
	'html = html & "<td class='fil'>NOMBRE:</td>"
	'html = html & "<td class='fil'><input class='inpu' style='width:80%;' id='nom' name='nom' type='text'></td>"
	'html = html & "</tr>"
	html = html & "<div class='dato-label'>NOMBRE</div>"
	html = html & "<div><input id='nom' name='nom' type='text' class='input-moderno ui-input'></div>"

	'html = html & "<tr>"
	'html = html & "<td class='fil'>APELLIDO:</td>"
	'html = html & "<td class='fil'><input class='inpu' style='width:80%;' id='ape' name='ape' type='text'></td>"
	'html = html & "</tr>"
	html = html & "<div class='dato-label'>APELLIDO</div>"
	html = html & "<div><input id='ape' name='ape' type='text' class='input-moderno ui-input'></div>"

	'html = html & "<tr>"
	'html = html & "<td class='fil'>TELEFONO:</td>"
	'html = html & "<td class='fil'><input class='inpu' style='width:80%;' id='tel' name='tel' type='text'></td>"
	'html = html & "</tr>"
	html = html & "<div class='dato-label'>TELEFONO</div>"
	html = html & "<div><input id='tel' name='tel' type='text' class='input-moderno ui-input'></div>"

	'html = html & "<tr>"
	'html = html & "<td class='fil'>MAIL:</td>"
	'html = html & "<td class='fil'><input class='inpu' style='width:80%;' id='mail' name='mail' type='text'></td>"
	'html = html & "</tr>"
	html = html & "<div class='dato-label'>MAIL</div>"
	html = html & "<div><input id='mail' name='mail' type='text' class='input-moderno ui-input'></div>"

	'html = html & "<tr>"
	'html = html & "<td class='fil'>CLIENTE:</td>"
	'html = html & "<td class='fil'><select class='selec' style='width:80%;' id='cli' name='cli'>"
	html = html & "<div class='dato-label'>CLIENTE</div>"
	html = html & "<div class='dato-value'><select class='ui-select' id='cli' name='cli'>"

		Set Cnn = Server.CreateObject("ADODB.Connection")
		Cnn.open scon1

		Set rs = Exec(Cnn, "sp_cg_Catalogo_Listar @Catalogo='clientes'")

		Do While Not rs.eof
			html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
			rs.movenext
		Loop

	'html = html & "</select></td>"
	'html = html & "</tr>"
	html = html & "</select></div>"

	'html = html & "<tr>"
	'html = html & "<td class='fil'>TIPO:</td>"
	'html = html & "<td class='fil'><select class='selec' style='width:80%;' id='tipo' name='tipo'>"
	'html = html & "<option value='1'>generico</option>"
	'html = html & "<option value='10'>normal</option>"
	'html = html & "</select></td>"
	'html = html & "</tr>"
	html = html & "<div class='dato-label'>TIPO</div>"
	html = html & "<div class='dato-value'><select class='ui-select' id='tipo' name='tipo'>"
	html = html & "<option value='1'>generico</option>"
	html = html & "<option value='10'>normal</option>"
	html = html & "</select></div>"

	'html = html & "<tr>"
	'html = html & "<td colspan=2>&nbsp;</td>"
	'html = html & "</tr>"

	' --- FIN GRILLA 2 COLUMNAS ---
	html = html & "</div>"

	'html = html & "<tr>"
	'html = html & "<td colspan=2 align=center><input class='bot' type='button' value='Agregar' onclick='agregar1();' ></td>"
	'html = html & "</tr>"
	html = html & "<div style='text-align: center; margin-bottom: 1rem; margin-top: 1rem;'>"
	html = html & "<input type='button' class='btn-principal ui-button ui-button--primary' value='Agregar' onclick='agregar1();'>"
	html = html & "</div>"

	'html = html & "</table>"
	html = html & "</div>" ' Fin .bloque-body
	html = html & "</div>" ' Fin .bloque-datos

	'html = html & "</form>"
	html = html & "</form>"

	'html = html & "</div>"
	html = html & "</div>" ' Fin max-width

Else

	'html = html & "<div style='height:57%;width:100%;'><br><br>"
	html = html & "<div class='form-shell'>"

	'html = html & "<form id='agre' method='post' action='agregar1.asp?opc=ins&tabla=" & tabla & "&volver=" & volver & "&back=" & sBack & "'>"
	html = html & "<form id='agre' method='post' action='agregar1.asp?opc=ins&tabla=" & tabla & "&volver=" & volver & "&back=" & sBack & "'>"

	'html = html & "<table class='tabl' style='width:40%;' align=center cellspacing=0 cellpadding=3>"
	html = html & "<div class='bloque-datos ui-panel'>"

	'html = html & "<tr>"
	'html = html & "<td class='fil' style='font-size:12pt;' colspan=2 align=center><b><u>COMPLETE LOS SIGUIENTES DATOS</u></b></td>"
	'html = html & "</tr>"
	html = html & "<div class='bloque-header'>COMPLETE LOS SIGUIENTES DATOS</div>"

	'html = html & "<tr>"
	'html = html & "<td colspan=2>&nbsp;</td>"
	'html = html & "</tr>"

	html = html & "<div class='bloque-body ui-panel__body'>"

	' --- INICIO GRILLA 2 COLUMNAS ---
	html = html & "<div class='grid-datos-doble ui-field-grid--dual'>"

	'html = html & "<tr>"
	'html = html & "<td class='fil'>NOMBRE:</td>"
	'html = html & "<td class='fil'><input class='inpu' style='width:80%;' id='nom' name='nom' type='text'></td>"
	'html = html & "</tr>"
	html = html & "<div class='dato-label'>NOMBRE</div>"
	html = html & "<div><input id='nom' name='nom' type='text' class='input-moderno ui-input'></div>"

	'html = html & "<tr>"
	'html = html & "<td class='fil'>APELLIDO:</td>"
	'html = html & "<td class='fil'><input class='inpu' style='width:80%;' id='ape' name='ape' type='text'></td>"
	'html = html & "</tr>"
	html = html & "<div class='dato-label'>APELLIDO</div>"
	html = html & "<div><input id='ape' name='ape' type='text' class='input-moderno ui-input'></div>"

	'html = html & "<tr>"
	'html = html & "<td class='fil'>FUNCION:</td>"
	'html = html & "<td class='fil'><select class='selec' style='width:80%;' id='func' name='func'>"
	'html = html & "<option value='TEC'>TECNICO</option>"
	'html = html & "<option value='FUNC'>FUNCIONAL</option>"
	'html = html & "</select></td>"
	'html = html & "</tr>"
	html = html & "<div class='dato-label'>FUNCION</div>"
	html = html & "<div class='dato-value'><select class='ui-select' id='func' name='func'>"
	html = html & "<option value='TEC'>TECNICO</option>"
	html = html & "<option value='FUNC'>FUNCIONAL</option>"
	html = html & "</select></div>"

	'html = html & "<tr>"
	'html = html & "<td class='fil'>LOGUIN:</td>"
	'html = html & "<td class='fil'><input class='inpu' style='width:80%;' id='LOGUIN' name='LOGUIN' type='text'></td>"
	'html = html & "</tr>"
	html = html & "<div class='dato-label'>LOGUIN</div>"
	html = html & "<div><input id='LOGUIN' name='LOGUIN' type='text' class='input-moderno ui-input'></div>"

	'html = html & "<tr>"
	'html = html & "<td colspan=2>&nbsp;</td>"
	'html = html & "</tr>"

	' --- FIN GRILLA 2 COLUMNAS ---
	html = html & "</div>"

	'html = html & "<tr>"
	'html = html & "<td class='fil' colspan=2 align=center><input class='bot' type='button' value='Agregar' onclick='agregar1();' ></td>"
	'html = html & "</tr>"
	html = html & "<div style='text-align: center; margin-bottom: 1rem; margin-top: 1rem;'>"
	html = html & "<input type='button' class='btn-principal ui-button ui-button--primary' value='Agregar' onclick='agregar1();'>"
	html = html & "</div>"

	'html = html & "</table>"
	html = html & "</div>" ' Fin .bloque-body
	html = html & "</div>" ' Fin .bloque-datos

	'html = html & "</form>"
	html = html & "</form>"

	'html = html & "</div>"
	html = html & "</div>" ' Fin max-width

End If


html = html & "<script src='assets/js/jquery.min.js'></script>"
html = html & "<script src='assets/js/skel.min.js'></script>"
html = html & "<script src='assets/js/util.js'></script>"
html = html & "<script src='assets/js/main.js'></script>"

html = html & "<script type='text/javascript'>"
html = html & " function agregar(){ "
html = html & " var des = document.getElementById('descri').value;"
html = html & " if (des == ''){ "
html = html & " alert('Descripcion no debe ser nula'); return false;} "
html = html & " else{ "
html = html & " document.getElementById('agre').submit();"
html = html & " }} "
html = html & "</script>"

html = html & "<script type='text/javascript'>"
html = html & " function agregar1(){ "
html = html & " var nom = document.getElementById('nom').value;"
html = html & " var ape = document.getElementById('ape').value;"
html = html & " if (nom == '' || ape == ''){ "
html = html & " alert('Nombre y apellido no deben ser nulo'); return false; "
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