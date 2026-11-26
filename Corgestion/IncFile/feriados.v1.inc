<%
'==============================================================================
' IncFile/feriados.inc  -  Feriados (CG_FERIADOS) para agenda_semana.asp
' Requiere una conexion abierta (cn). SP y tabla: DB/migracion_cg/16_feriados.sql
'
' FeriadosAsegurar(cn, desde, hasta) devuelve un Dictionary dia -> nombre del feriado
' (dia = dias desde 1900-01-01, como en el resto del proyecto) para la semana pedida.
' Si la semana NO tiene feriados en la tabla, se confirma contra la API
' https://api.argentinadatos.com/v1/feriados/<anio> y, si la tabla no esta al dia, se actualiza
' (el anio completo). Para no llamar a la API en cada vista, cada anio se consulta como maximo
' una vez por dia (15 minutos si la API fallo). Si la API o la tabla no estan disponibles,
' la agenda se muestra igual, sin feriados.
'==============================================================================

' Feriados entre dos dias -> Dictionary dia -> nombre (varios en el mismo dia se unen con " / ")
Function FeriadosSemana(cn, desde, hasta)
	Dim d, cmd, rs, k, nom
	Set d = Server.CreateObject("Scripting.Dictionary")
	On Error Resume Next
	Set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = cn
	cmd.CommandType = 4
	cmd.CommandText = "sp_cg_Feriados"
	cmd.Parameters.Append cmd.CreateParameter("@Modo", 200, 1, 15, "semana")
	cmd.Parameters.Append cmd.CreateParameter("@Desde", 3, 1, 4, CLng(desde))
	cmd.Parameters.Append cmd.CreateParameter("@Hasta", 3, 1, 4, CLng(hasta))
	Set rs = cmd.Execute
	If Err.Number = 0 Then
		Do While Not rs.EOF
			k = CStr(rs("FECHA"))
			nom = Trim(rs("NOMBRE") & "")
			If nom = "" Then nom = "Feriado"
			If d.Exists(k) Then
				If InStr(d(k), nom) = 0 Then d(k) = d(k) & " / " & nom
			Else
				d.Add k, nom
			End If
			rs.MoveNext
		Loop
		rs.Close
	End If
	Err.Clear
	On Error Goto 0
	Set cmd = Nothing
	Set FeriadosSemana = d
End Function

' Descarga el JSON del anio desde la API. Devuelve "" si falla.
Function FeriadosApiTexto(anio)
	Dim http, url
	FeriadosApiTexto = ""
	url = "https://api.argentinadatos.com/v1/feriados/" & CLng(anio)
	On Error Resume Next
	Set http = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
	If Err.Number <> 0 Then
		Err.Clear
		Set http = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	End If
	If Err.Number = 0 Then
		http.setTimeouts 4000, 4000, 4000, 8000   ' resolver, conectar, enviar, recibir (ms)
		http.Open "GET", url, False
		http.setRequestHeader "Accept", "application/json"
		http.Send
		If Err.Number = 0 Then
			If http.Status = 200 Then FeriadosApiTexto = http.responseText
		End If
	End If
	Err.Clear
	On Error Goto 0
	Set http = Nothing
End Function

' Texto de un valor JSON: quita escapes (\" \\ \/ \n \t \uXXXX)
Function FeriadosJsonTexto(s)
	Dim r, p, hx
	r = s
	p = InStr(r, "\u")
	Do While p > 0
		hx = Mid(r, p + 2, 4)
		If Len(hx) = 4 And IsNumeric("&H" & hx) Then
			r = Left(r, p - 1) & ChrW(CLng("&H" & hx)) & Mid(r, p + 6)
			p = InStr(p + 1, r, "\u")
		Else
			p = InStr(p + 2, r, "\u")
		End If
	Loop
	r = Replace(r, "\""", """")
	r = Replace(r, "\/", "/")
	r = Replace(r, "\n", " ")
	r = Replace(r, "\t", " ")
	r = Replace(r, "\\", "\")
	FeriadosJsonTexto = r
End Function

Function FeriadosXmlAttr(s)
	Dim r
	r = Replace(s & "", "&", "&amp;")
	r = Replace(r, "<", "&lt;")
	r = Replace(r, ">", "&gt;")
	r = Replace(r, """", "&quot;")
	FeriadosXmlAttr = r
End Function

' Valor de una clave en un objeto JSON plano ({"fecha":"...","tipo":"...","nombre":"..."})
Function FeriadosJsonValor(obj, clave)
	Dim re, m
	FeriadosJsonValor = ""
	Set re = New RegExp
	re.Pattern = """" & clave & """\s*:\s*""((?:[^""\\]|\\.)*)"""
	Set m = re.Execute(obj)
	If m.Count > 0 Then FeriadosJsonValor = FeriadosJsonTexto(m(0).SubMatches(0))
End Function

' JSON de la API -> <r><f d="aaaa-mm-dd" t="tipo" n="nombre"/>...</r>  ("" si no hay feriados)
Function FeriadosJsonAXml(json)
	Dim re, ms, mo, x, f, n
	Set re = New RegExp
	re.Global = True
	re.Pattern = "\{[^{}]*\}"
	Set ms = re.Execute(json)
	x = ""
	n = 0
	For Each mo In ms
		f = FeriadosJsonValor(mo.Value, "fecha")
		If f <> "" Then
			x = x & "<f d=""" & FeriadosXmlAttr(f) & """ t=""" & FeriadosXmlAttr(FeriadosJsonValor(mo.Value, "tipo")) & """ n=""" & FeriadosXmlAttr(FeriadosJsonValor(mo.Value, "nombre")) & """/>"
			n = n + 1
		End If
	Next
	If n = 0 Then FeriadosJsonAXml = "" Else FeriadosJsonAXml = "<r>" & x & "</r>"
End Function

' Deja CG_FERIADOS igual a la API para el anio (json = respuesta ya descargada). Devuelve True si hubo cambios.
Function FeriadosSincronizarAnio(cn, anio, json)
	Dim cmd, rs, xml
	FeriadosSincronizarAnio = False
	If Len(json) < 10 Then Exit Function
	xml = FeriadosJsonAXml(json)
	If xml = "" Then Exit Function
	On Error Resume Next
	Set cmd = Server.CreateObject("ADODB.Command")
	cmd.ActiveConnection = cn
	cmd.CommandType = 4
	cmd.CommandText = "sp_cg_Feriados"
	cmd.Parameters.Append cmd.CreateParameter("@Modo", 200, 1, 15, "sincronizar")
	cmd.Parameters.Append cmd.CreateParameter("@Desde", 3, 1, 4, Null)
	cmd.Parameters.Append cmd.CreateParameter("@Hasta", 3, 1, 4, Null)
	cmd.Parameters.Append cmd.CreateParameter("@Anio", 3, 1, 4, CLng(anio))
	cmd.Parameters.Append cmd.CreateParameter("@Datos", 203, 1, Len(xml) + 1, xml)   ' adLongVarWChar -> parametro XML
	Set rs = cmd.Execute
	If Err.Number = 0 Then
		If Not rs.EOF Then FeriadosSincronizarAnio = (CLng(rs("CAMBIOS")) = 1)
		rs.Close
	End If
	Err.Clear
	On Error Goto 0
	Set cmd = Nothing
End Function

' True si hay que confirmar el anio contra la API (no se consulto en las ultimas 24 h / 15 min si fallo)
Function FeriadosDebeConsultar(anio)
	Dim ts, esperaMin
	ts = Application("FER_TS_" & anio)
	FeriadosDebeConsultar = True
	If IsDate(ts) Then
		If Application("FER_OK_" & anio) = True Then esperaMin = 1440 Else esperaMin = 15
		If DateDiff("n", CDate(ts), Now()) < esperaMin Then FeriadosDebeConsultar = False
	End If
End Function

Sub FeriadosMarcarConsulta(anio, ok)
	Application.Lock
	Application("FER_TS_" & anio) = Now()
	Application("FER_OK_" & anio) = ok
	Application.UnLock
End Sub

' Punto de entrada: feriados de la semana desde-hasta (dias desde 1900), confirmando con la API si no hay ninguno
Function FeriadosAsegurar(cn, desde, hasta)
	Dim d, a1, a2, a, ok, cambio, txt
	Set d = FeriadosSemana(cn, desde, hasta)
	If d.Count = 0 Then
		a1 = Year(DateSerial(1900, 1, 1) + CLng(desde))
		a2 = Year(DateSerial(1900, 1, 1) + CLng(hasta))
		cambio = False
		For a = a1 To a2
			If FeriadosDebeConsultar(a) Then
				txt = FeriadosApiTexto(a)
				ok = (Len(txt) >= 10)
				If ok Then
					If FeriadosSincronizarAnio(cn, a, txt) Then cambio = True
				End If
				FeriadosMarcarConsulta a, ok
			End If
		Next
		If cambio Then Set d = FeriadosSemana(cn, desde, hasta)
	End If
	Set FeriadosAsegurar = d
End Function
%>
