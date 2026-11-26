<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/datosActualizar.inc" -->
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
html = html & "<body style='background:#2F81C6;'>"

opt = request.querystring("opt")
id = request.querystring("id")
tabla = request.querystring("tabla")
cli = request.querystring("cli")
relainc = request.querystring("relainc")
xmenu = request.querystring("xmenu")
xtmenu = request.querystring("xtmenu")
cantregmover = cdbl(request.querystring("cantregmover"))

Set Cnn = Server.CreateObject("ADODB.Connection")
'Cnn.open "Provider=SQLOLEDB;Data Source=server-2025;Initial Catalog=dbcorgestionv5; user id=sa;password=bb4515$;"
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

nFecini = 0

If nFecini < CLng(CDate("01-01-1900")) Then
	nFecini = CLng(CDate("01-01-1900"))
Else
	nFecini = 0
End If

sFec = CLng(Date) - nFecini

Dim avanceFecha, avanceHora, avanceResponsable
avanceFecha = 0
avanceHora = 0
If IsDate(request.form("avance_fecha")) Then
	avanceFecha = CLng(CDate(request.form("avance_fecha"))) - nFecini
End If
If IsDate(request.form("avance_hora")) Then
	avanceHora = (DatePart("h", CDate(request.form("avance_hora"))) * 3600) + (DatePart("n", CDate(request.form("avance_hora"))) * 60)
End If
avanceResponsable = UCase(Trim(request.form("responsable_avance") & ""))

' [UI] Avance desde la vista del incidente (pop-up de detalles.asp). Registra SOLO el avance con
' RegistrarAvance, igual que la edicion (mismos calculos de fecha, hora y responsable de arriba).
' No actualiza los campos del incidente ni sincroniza estados. Exige el permiso de modificar.
If opt = "avainc" Then
	If IsNumeric(id & "") And Trim(id & "") <> "" Then
		If Trim(request.form("avance") & "") <> "" And getPermisosAtrm("CG_INCIDENTES", "MODIF") = 1 Then
			RegistrarAvanceAnt Cnn, "incidente", id, request.form("avance"), avanceFecha, avanceHora, avanceResponsable, request.form("tipo_avance"), request.form("avance_origen")
			CgLog "INCIDENTE", id, "ACTUALIZA"
			Response.Redirect "detalles.asp?opc=veri&id=" & CLng(id) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&av=ok"
		Else
			Response.Redirect "detalles.asp?opc=veri&id=" & CLng(id) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&av=error"
		End If
	Else
		Response.Redirect "index.asp"
	End If
End If

' [UI] Avance desde la vista del proyecto (pop-up de detalles.asp): igual que avainc, para proyectos.
If opt = "avaproy" Then
	If IsNumeric(id & "") And Trim(id & "") <> "" Then
		If Trim(request.form("avance") & "") <> "" And getPermisosAtrm("CG_PROYECTOS", "MODIF") = 1 Then
			RegistrarAvanceAnt Cnn, "proyecto", id, request.form("avance"), avanceFecha, avanceHora, avanceResponsable, request.form("tipo_avance"), request.form("avance_origen")
			CgLog "PROYECTO", id, "ACTUALIZA"
			Response.Redirect "detalles.asp?opc=verp&id=" & CLng(id) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&av=ok"
		Else
			Response.Redirect "detalles.asp?opc=verp&id=" & CLng(id) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&av=error"
		End If
	Else
		Response.Redirect "index.asp"
	End If
End If

' [ESTADO] Cambiar el estado de un avance desde el pop-up de detalle del avance (detalles.asp). Exige el permiso de modificar
' de la pantalla (incidentes o proyectos). Solo cambia CG_AVANCES.IDESTADO del avance indicado (parametro "avance").
If opt = "estadoav" Then
	Dim esIncAv, formAv, opcAv, okAv
	esIncAv = (LCase(Trim(tabla & "")) = "incidentes")
	If esIncAv Then
		formAv = "CG_INCIDENTES"
		opcAv = "veri"
	Else
		formAv = "CG_PROYECTOS"
		opcAv = "verp"
	End If
	okAv = False
	If IsNumeric(id & "") And Trim(id & "") <> "" Then
		If IsNumeric(request.querystring("avance") & "") And Trim(request.querystring("avance") & "") <> "" And IsNumeric(request.form("estado") & "") And Trim(request.form("estado") & "") <> "" And getPermisosAtrm(formAv, "MODIF") = 1 Then
			okAv = CambiarEstadoAvance(Cnn, tabla, id, request.querystring("avance"), request.form("estado"))
		End If
		If okAv Then
			CgLog IIf(esIncAv, "INCIDENTE", "PROYECTO"), id, "ACTUALIZA"
			Response.Redirect "detalles.asp?opc=" & opcAv & "&id=" & CLng(id) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&av=estado"
		Else
			Response.Redirect "detalles.asp?opc=" & opcAv & "&id=" & CLng(id) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&av=estadoerror"
		End If
	Else
		Response.Redirect "index.asp"
	End If
End If

' [UI] Relacionar un APL desde la vista del incidente o del proyecto (pop-up de detalles.asp: solo el ID del APL).
' Usa sp_cg_Registro_Insertar (auxiliar / APL), que valida que el APL y el caso existan. Si el APL ya
' estaba relacionado no lo vuelve a insertar. Exige el permiso de modificar.
If opt = "relapl" Then
	Dim tipoRelApl, vistaRelApl, idAplRel, rsRelApl, yaRelApl, destinoRelApl
	If tabla = "proyectos" Then
		tipoRelApl = "P"
		vistaRelApl = "verp"
	Else
		tipoRelApl = "I"
		vistaRelApl = "veri"
	End If
	If IsNumeric(id & "") And Trim(id & "") <> "" Then
		destinoRelApl = "detalles.asp?opc=" & vistaRelApl & "&id=" & CLng(id) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu
		idAplRel = Trim(request.form("id_apl") & "")
		If IsNumeric(idAplRel) And idAplRel <> "" And getPermisosAtrm(IIf(tipoRelApl = "P", "CG_PROYECTOS", "CG_INCIDENTES"), "MODIF") = 1 Then
			yaRelApl = False
			Set rsRelApl = Cnn.Execute("sp_cg_RelacionAPL @Modo='relac', @IdRelac=" & CLng(id) & ", @TipoRelac='" & tipoRelApl & "'")
			Do While Not rsRelApl.EOF
				If CStr(rsRelApl("id_apl")) = CStr(CLng(idAplRel)) Then yaRelApl = True
				rsRelApl.MoveNext
			Loop
			rsRelApl.Close
			If yaRelApl Then Response.Redirect destinoRelApl & "&apl=ya"
			On Error Resume Next
			Cnn.Execute "sp_cg_Registro_Insertar @Entidad='auxiliar', @Param='APL', @Dato1='" & CLng(idAplRel) & "', @Dato2='" & CLng(id) & "', @Dato3='" & tipoRelApl & "'"
			If Err.Number <> 0 Then
				Err.Clear
				On Error Goto 0
				Response.Redirect destinoRelApl & "&apl=noexiste&idapl=" & CLng(idAplRel)
			End If
			On Error Goto 0
			CgLog IIf(tipoRelApl = "P", "PROYECTO", "INCIDENTE"), id, "ACTUALIZA"
			CgLog "APL", idAplRel, "ACTUALIZA"
			Response.Redirect destinoRelApl & "&apl=ok&idapl=" & CLng(idAplRel)
		Else
			Response.Redirect destinoRelApl & "&apl=error"
		End If
	Else
		Response.Redirect "index.asp"
	End If
End If

' [UI] Relacionar un incidente o proyecto desde la vista del APL (pop-up de detalles2.asp: tipo y numero).
' Usa la misma rama validada que relapl (sp_cg_Registro_Insertar auxiliar / APL: valida que el APL y el
' caso existan; escribe la misma fila que la rama IOP, que no valida). No duplica. Exige MODIF.
If opt = "relcaso" Then
	Dim tipoRelCaso, idCasoRel, rsRelCaso, yaRelCaso, destinoRelCaso
	tipoRelCaso = UCase(Trim(request.form("tipo_relac") & ""))
	idCasoRel = Trim(request.form("id_caso") & "")
	If IsNumeric(id & "") And Trim(id & "") <> "" Then
		destinoRelCaso = "detalles2.asp?opc=ver&id=" & CLng(id) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & _
			"&cli=" & Server.URLEncode(request.querystring("cli") & "") & "&aplic=" & Server.URLEncode(request.querystring("aplic") & "") & _
			"&tip=" & Server.URLEncode(request.querystring("tip") & "")
		If (tipoRelCaso = "I" Or tipoRelCaso = "P") And IsNumeric(idCasoRel) And idCasoRel <> "" And getPermisosAtrm(IIf(tipoRelCaso = "P", "CG_PROYECTOS", "CG_INCIDENTES"), "MODIF") = 1 Then
			destinoRelCaso = destinoRelCaso & "&reltipo=" & tipoRelCaso & "&relid=" & CLng(idCasoRel)
			yaRelCaso = False
			Set rsRelCaso = Cnn.Execute("sp_cg_RelacionAPL @Modo='relac', @IdRelac=" & CLng(idCasoRel) & ", @TipoRelac='" & tipoRelCaso & "'")
			Do While Not rsRelCaso.EOF
				If CStr(rsRelCaso("id_apl")) = CStr(CLng(id)) Then yaRelCaso = True
				rsRelCaso.MoveNext
			Loop
			rsRelCaso.Close
			If yaRelCaso Then Response.Redirect destinoRelCaso & "&rel=ya"
			On Error Resume Next
			Cnn.Execute "sp_cg_Registro_Insertar @Entidad='auxiliar', @Param='APL', @Dato1='" & CLng(id) & "', @Dato2='" & CLng(idCasoRel) & "', @Dato3='" & tipoRelCaso & "'"
			If Err.Number <> 0 Then
				Err.Clear
				On Error Goto 0
				Response.Redirect destinoRelCaso & "&rel=noexiste"
			End If
			On Error Goto 0
			CgLog "APL", id, "ACTUALIZA"
			CgLog IIf(tipoRelCaso = "P", "PROYECTO", "INCIDENTE"), idCasoRel, "ACTUALIZA"
			Response.Redirect destinoRelCaso & "&rel=ok"
		Else
			Response.Redirect destinoRelCaso & "&rel=error"
		End If
	Else
		Response.Redirect "index.asp"
	End If
End If


''update  actincidentes'''


If opt = "actinc" Then

	' "ok_cliente" se muestra formateado (dd/mm/aaaa) en detalles.asp para que
	' se pueda leer; igual que "fecha_entrega" en proyectos, hay que convertirlo
	' de vuelta a dias-desde-1900 antes de guardar (la columna es numerica).
	Dim sOkCliInc
	sOkCliInc = ""
	If request.form("ok_cliente") & "" <> "" And request.form("ok_cliente") & "" <> "0" Then
		sOkCliInc = clng(cdate(request.form("ok_cliente"))) - nFecini
	End If

	' Cualquier campo posteado se actualiza en "tabla" (misma logica que el
	' original: recorre TODO request.form). La validacion de que cada NAME
	' sea una columna real ahora la hace el SP (ver sp_cg_Registro_ActualizarCampos).
	camposJson = ArmarJsonDesdeForm(request.form, "ok_cliente", CStr(sOkCliInc), "", "")
	ActualizarCamposDinamico Cnn, tabla, id, camposJson
	RegistrarAvance Cnn, "incidente", id, request.form("avance"), avanceFecha, avanceHora, avanceResponsable, request.form("tipo_avance")

	Dim estadoDesaVal
	estadoDesaVal = request.form("estado_desa")
	If Not IsNumeric(estadoDesaVal) Then estadoDesaVal = 0

	SincronizarEstadosIncidente Cnn, id, estadoDesaVal, sFec, relainc

ElseIf opt = "actproy" Then

	' Caso especial preservado del original: "fecha_entrega" no se manda
	' como vino del form, se convierte primero a dias-desde-1900.
	Dim sFecEn
	sFecEn = ""
	If request.form("fecha_entrega") & "" <> "" And request.form("fecha_entrega") & "" <> "0" Then
		sFecEn = clng(cdate(request.form("fecha_entrega"))) - nFecini
	End If

	' "ok_cliente" tambien se muestra formateado (dd/mm/aaaa) en detalles.asp;
	' misma conversion de vuelta a dias-desde-1900 antes de guardar.
	Dim sOkCliProy
	sOkCliProy = ""
	If request.form("ok_cliente") & "" <> "" And request.form("ok_cliente") & "" <> "0" Then
		sOkCliProy = clng(cdate(request.form("ok_cliente"))) - nFecini
	End If

	camposJson = ArmarJsonDesdeForm(request.form, "fecha_entrega", CStr(sFecEn), "ok_cliente", CStr(sOkCliProy))
	ActualizarCamposDinamico Cnn, tabla, id, camposJson
	RegistrarAvance Cnn, "proyecto", id, request.form("avance"), avanceFecha, avanceHora, avanceResponsable, request.form("tipo_avance")
	response.write "Jsoncampos:" & camposJson
	Dim estadoDesaValP
	estadoDesaValP = request.form("estado_desa")
	If Not IsNumeric(estadoDesaValP) Then estadoDesaValP = 0

	SincronizarEstadosProyecto Cnn, id, estadoDesaValP, sFec
	
	response.write "Sincronico proyecto"

ElseIf opt = "actapl" Then

	sNom = request.form("nombre")
	sCli = request.form("id_cliente")
	sApl = request.form("id_aplic")
	srelainc = request.form("id")
	sResp = request.form("respon_envio")
	sFecEnv = request.form("fecha_envio")
	cantregmover = cdbl(request.querystring("cantregmover"))


	if sFecEnv & "" <> "0" and sFecEnv & "" <> "" then
		sFecEnv = clng(CDATE(sFecEnv)) - CLng(CDate("01-01-1900"))
	else
		sFecEnv = 0
	end if
	sFecTest = request.form("fecha_instalacion_test")
	if sFecTest & "" <> "0" and sFecTest & "" <> "" then
		sFecTest = clng(CDATE(sFecTest)) - CLng(CDate("01-01-1900"))
	else
		sFecTest = 0
	end if
	sFecProd = request.form("fecha_instalacion_prod")
	if sFecProd & "" <> "0" and sFecProd & "" <> "" then
		sFecProd = clng(CDATE(sFecProd)) - CLng(CDate("01-01-1900"))
	else
		sFecProd = 0
	end if
	sMotivo = request.form("motivo_apl")
	sIncidente = request.form("incidente")
	if sIncidente = "" then sIncidente = 0
	sProyecto = request.form("proyecto")
	if sProyecto = "" then sProyecto = 0

	secuencia = request.form("sec")
	tipo = request.form("tipo")

	ActualizarApl Cnn, id, sNom, sCli, sApl, sResp, sFecEnv, sFecTest, sFecProd, sMotivo, sIncidente, sProyecto, secuencia, tipo

End If


' [LOG] Actualizacion de incidente / proyecto / APL (con o sin avance)
If opt = "actinc" Then CgLog "INCIDENTE", id, "ACTUALIZA"
If opt = "actproy" Then CgLog "PROYECTO", id, "ACTUALIZA"
If opt = "actapl" Then CgLog "APL", id, "ACTUALIZA"

'vuelta
vuelta = request.querystring("vuelta")
' html = html & "<script>alert('" & vuelta & "')</script>"
'cantregmover = cdbl(request.querystring("cantregmover"))

if opt = "actinc" then
	vuelta = vuelta & "&id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu
elseif opt = "actproy" then
	vuelta = vuelta & "&id=" & id & "&sol=" & request.querystring("sol") & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu
 elseif opt = "actapl" then
	vuelta = vuelta & "&id=" & id & "&cli=" & cli & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&cantregmover=" & cantregmover
end if

html = html & "<script>alert('" & vuelta & "')</script>"

html = html & "<FORM name='init' method='post'  ACTION='" & vuelta & "' >"
html = html & "<INPUT TYPE='SUBMIT' VALUE='INICIAR' style='display:none;' >"
html = html & "</FORM>"

html = html & "<script type='text/javascript'>"
html = html & " alert('Actualizado exitosamente');"
html = html & " document.all.init.submit();"
html = html & "</script>"



html = html & "</body>"
html = html & "</html>"

response.write html
%>