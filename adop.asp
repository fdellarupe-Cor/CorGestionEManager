<%
On Error Resume Next

Sub CheckDriver(progID, friendlyName)
    Dim objConn
    Set objConn = Server.CreateObject("ADODB.Connection")
    ' Intentamos instanciar el driver sin abrir conexión real
    objConn.Provider = progID

    If Err.Number = 0 Then
        Response.Write "<p style='color:green;'><strong>[INSTALADO]</strong> " & friendlyName & " (" & progID & ")</p>"
    Else
        Response.Write "<p style='color:red;'><strong>[NO DISPONIBLE]</strong> " & friendlyName & " (" & progID & ")</p>"
        Err.Clear
    End If
    Set objConn = Nothing
End Sub

Response.Write "<h2>Verificación de Drivers en el Servidor</h2>"

' Verificamos los tres sospechosos habituales
CheckDriver "SQLNCLI11", "SQL Server Native Client 11.0 (SQL 2012+)"
CheckDriver "SQLNCLI10", "SQL Server Native Client 10.0 (SQL 2008)"
CheckDriver "SQLOLEDB",  "Microsoft OLE DB Provider for SQL (Antiguo)"

If Err.Number <> 0 Then Err.Clear
%>