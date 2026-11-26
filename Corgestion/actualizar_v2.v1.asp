<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/datosActualizar.inc" -->
<%


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
	RegistrarAvance Cnn, "incidente", id, request.form("avance"), avanceFecha, avanceHora, avanceResponsable

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
	RegistrarAvance Cnn, "proyecto", id, request.form("avance"), avanceFecha, avanceHora, avanceResponsable
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