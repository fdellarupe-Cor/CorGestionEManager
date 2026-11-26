<%
' Función para ejecutar el SP parametrizado y devolver el recordset de APLs
Function GetAPLs(Cnn, cons, cliente, idAplic, idtip, asp, fec, ord)
    Dim cmd
    Set cmd = Server.CreateObject("ADODB.Command")
    cmd.ActiveConnection = Cnn
    cmd.CommandText = "sp_cg_ConsultarEntidad"
    cmd.CommandType = 4 ' adCmdStoredProc
    
    cmd.Parameters.Append cmd.CreateParameter("@Entidad", 200, 1, 15, "apl")
    cmd.Parameters.Append cmd.CreateParameter("@Modo", 200, 1, 15, "lista")
    cmd.Parameters.Append cmd.CreateParameter("@Cons", 200, 1, 20, cons)
    cmd.Parameters.Append cmd.CreateParameter("@IdCliente", 3, 1, 4, cliente)
    cmd.Parameters.Append cmd.CreateParameter("@IdAplicativo", 3, 1, 4, idAplic)
    cmd.Parameters.Append cmd.CreateParameter("@IdTip", 3, 1, 4, idtip)
    cmd.Parameters.Append cmd.CreateParameter("@Asp", 200, 1, 100, asp)
    cmd.Parameters.Append cmd.CreateParameter("@Fec", 3, 1, 4, fec)
    cmd.Parameters.Append cmd.CreateParameter("@Orden", 200, 1, 50, ord)
    
    On Error Resume Next
    Set GetAPLs = cmd.Execute()
    If Err.Number <> 0 Then
        Response.Write "<div style='color:red;font-family:monospace;padding:10px;border:1px solid red;margin:10px;'>"
        Response.Write "<b>ERROR DE BASE DE DATOS</b><br>"
        Response.Write "Contexto: GetAPLs (cons=" & Server.HTMLEncode(cons) & ")<br>"
        Response.Write "Codigo: " & Err.Number & "<br>"
        Response.Write "Detalle: " & Server.HTMLEncode(Err.Description) & "</div>"
        Response.End
    End If
    On Error Goto 0
    Set cmd = Nothing
End Function

' GetListaComboExtended (catalogos para combos) vive ahora en IncFile/funciones.inc,
' compartida por todas las pantallas -- no duplicar aca.
%>