<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/feriados.inc" -->
<%
' ============================================================================
' feriados_test.asp - Diagnostico de la sincronizacion de feriados (CG_FERIADOS)
'
'   feriados_test.asp                -> solo lectura: prueba la API y muestra el estado
'   feriados_test.asp?anio=2026      -> otro anio
'   feriados_test.asp?sync=1         -> ademas ejecuta la sincronizacion (escribe en CG_FERIADOS)
'   feriados_test.asp?reset=1        -> borra las marcas de "ya consultado" del anio (memoria de la aplicacion)
' Es una pagina de ayuda: se puede quitar del servidor cuando ya no haga falta.
' ============================================================================
' [CG] Control de usuario E-Manager (igual que CorEManager/Acciones.asp)
cgOk = ofv.CheckUsuario()
If cgOk <> "Y" Then
	cgUv = ofv.uv
	If cgUv = "NO" Then cgOk = ofv.ValidarUsuario
	Response.Write cgOk
	Response.End
End If

Dim Cnn, anio, doSync, doReset, txt, xml, nF, n, ini, fin, rs, filas, fila, t0, cambio
Set Cnn = Server.CreateObject("ADODB.Connection")
On Error Resume Next
Cnn.Open scon1
If Err.Number <> 0 Then
	Response.Write "<div style='color:red;font-family:monospace'>ERROR DE CONEXION A BASE DE DATOS (" & Err.Number & "): " & Server.HTMLEncode(Err.Description) & "</div>"
	Response.End
End If
On Error Goto 0

anio = Trim(Request.QueryString("anio") & "")
If anio = "" Or Not IsNumeric(anio) Then anio = Year(Date()) Else anio = CLng(anio)
doSync = (Request.QueryString("sync") & "" = "1")
doReset = (Request.QueryString("reset") & "" = "1")

If doReset Then
	Application.Lock
	Application.Contents.Remove("FER_TS_" & anio)
	Application.Contents.Remove("FER_OK_" & anio)
	Application.Contents.Remove("FER_ERR")
	Application.UnLock
End If

Response.Write "<html><head><title>Diagnostico de feriados</title><meta charset='windows-1252' />"
Response.Write "<style>body{font-family:Segoe UI,Arial,sans-serif;margin:24px;color:#0f172a}table{border-collapse:collapse}td,th{border:1px solid #cbd5e1;padding:6px 10px;text-align:left}th{background:#e8edff}.ok{color:#15803d;font-weight:600}.mal{color:#b91c1c;font-weight:600}code{background:#f1f5f9;padding:1px 4px}</style></head><body>"
Response.Write "<h2>Diagnostico de feriados - anio " & anio & "</h2>"
Response.Write "<p>Servidor: " & Server.HTMLEncode(Request.ServerVariables("SERVER_NAME") & "") & " | Usuario del proceso: " & Server.HTMLEncode(Request.ServerVariables("LOGON_USER") & "") & "</p>"

Response.Write "<h3>1. Descarga desde la API</h3>"
txt = FeriadosApiTexto(anio)
Response.Write "<table><tr><th>URL</th><td>https://api.argentinadatos.com/v1/feriados/" & anio & "</td></tr>"
Response.Write "<tr><th>Estado HTTP</th><td>" & Server.HTMLEncode(gFerHttpStatus & "") & "</td></tr>"
Response.Write "<tr><th>Tiempo</th><td>" & gFerHttpMs & " ms</td></tr>"
Response.Write "<tr><th>Bytes recibidos</th><td>" & Len(txt) & "</td></tr>"
If Len(txt) >= 10 Then
	Response.Write "<tr><th>Resultado</th><td class='ok'>La API respondio</td></tr></table>"
Else
	Response.Write "<tr><th>Resultado</th><td class='mal'>No se pudo descargar</td></tr></table>"
	Response.Write "<p>Motivo registrado: <code>" & Server.HTMLEncode(Application("FER_ERR") & "") & "</code></p>"
	Response.Write "<p>Revisar: salida a Internet del servidor, proxy de WinHTTP (<code>netsh winhttp show proxy</code>) y TLS 1.2.</p>"
End If

Response.Write "<h3>2. Lectura del JSON</h3>"
xml = ""
nF = 0
If Len(txt) >= 10 Then
	xml = FeriadosJsonAXml(txt)
	n = InStr(xml, "<f ")
	Do While n > 0
		nF = nF + 1
		n = InStr(n + 1, xml, "<f ")
	Loop
End If
If nF > 0 Then
	Response.Write "<p class='ok'>Feriados leidos: " & nF & "</p>"
Else
	Response.Write "<p class='mal'>No se leyo ningun feriado del JSON.</p>"
End If

Response.Write "<h3>3. Tabla CG_FERIADOS del anio</h3>"
ini = DateDiff("d", DateSerial(1900, 1, 1), DateSerial(anio, 1, 1))
fin = DateDiff("d", DateSerial(1900, 1, 1), DateSerial(anio, 12, 31))
On Error Resume Next
Set rs = Cnn.Execute("SELECT FECHA, TIPO, NOMBRE FROM dbo.CG_FERIADOS WHERE FECHA BETWEEN " & ini & " AND " & fin & " ORDER BY FECHA")
If Err.Number <> 0 Then
	Response.Write "<p class='mal'>No se pudo leer la tabla (" & Err.Number & "): " & Server.HTMLEncode(Err.Description) & "</p>"
	Err.Clear
Else
	filas = 0
	fila = "<table><tr><th>Fecha</th><th>Tipo</th><th>Nombre</th></tr>"
	Do While Not rs.EOF
		filas = filas + 1
		fila = fila & "<tr><td>" & FormatDateTime(DateSerial(1900, 1, 1) + CLng(rs("FECHA")), 2) & "</td><td>" & Server.HTMLEncode(rs("TIPO") & "") & "</td><td>" & Server.HTMLEncode(rs("NOMBRE") & "") & "</td></tr>"
		rs.MoveNext
	Loop
	rs.Close
	Response.Write "<p>Filas del anio en la tabla: <b>" & filas & "</b></p>"
	If filas > 0 Then Response.Write fila & "</table>"
End If
On Error Goto 0

Response.Write "<h3>4. Guardado (sp_cg_Feriados)</h3>"
If doSync Then
	If nF > 0 Then
		cambio = FeriadosSincronizarAnio(Cnn, anio, txt)
		If gFerSyncOk Then
			If cambio Then
				Response.Write "<p class='ok'>Sincronizado: la tabla se actualizo con los datos de la API.</p>"
			Else
				Response.Write "<p class='ok'>El procedimiento funciono: la tabla ya estaba igual que la API (sin cambios).</p>"
			End If
		Else
			Response.Write "<p class='mal'>El guardado fallo.</p>"
		End If
	Else
		Response.Write "<p class='mal'>No se ejecuta el guardado: no hay feriados que guardar.</p>"
	End If
Else
	Response.Write "<p>No ejecutado (solo lectura). Para probarlo: <a href='feriados_test.asp?anio=" & anio & "&sync=1'>feriados_test.asp?anio=" & anio & "&amp;sync=1</a> (escribe en CG_FERIADOS).</p>"
End If

Response.Write "<h3>5. Estado interno de la aplicacion</h3><table>"
Response.Write "<tr><th>Ultimo error (FER_ERR)</th><td>" & Server.HTMLEncode(Application("FER_ERR") & "") & "</td></tr>"
Response.Write "<tr><th>Ultima consulta del anio (FER_TS_" & anio & ")</th><td>" & Server.HTMLEncode(Application("FER_TS_" & anio) & "") & "</td></tr>"
Response.Write "<tr><th>Ultima consulta correcta (FER_OK_" & anio & ")</th><td>" & Server.HTMLEncode(Application("FER_OK_" & anio) & "") & "</td></tr></table>"
Response.Write "<p><a href='feriados_test.asp?anio=" & anio & "&reset=1'>Borrar las marcas del anio</a> (hace que agenda_semana.asp vuelva a consultar la API en la proxima carga).</p>"
Response.Write "</body></html>"
Cnn.Close
Set Cnn = Nothing
%>
