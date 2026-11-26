<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<%
'Option Explicit

' --- Funciones de Sanitización y Conversión ---
Function ConvertirFechaASerial(fecha)
    If IsDate(fecha) Then
        ConvertirFechaASerial = CLng(CDate(fecha)) - CLng(CDate("01/01/1900"))
    Else
        ConvertirFechaASerial = 0
    End If
End Function

Function SanitizeInt(val)
    If IsNumeric(val) And Trim(val) <> "" Then
        SanitizeInt = CLng(val)
    Else
        SanitizeInt = 0
    End If
End Function

' --- Captura de Variables QueryString ---
Dim opt, id, tabla, cli, xmenu, xtmenu, cantregmover, vuelta, sol
opt          = Trim(Request.QueryString("opt"))
id           = SanitizeInt(Request.QueryString("id"))
tabla        = Trim(Request.QueryString("tabla"))
cli          = Trim(Request.QueryString("cli"))
xmenu        = Trim(Request.QueryString("xmenu"))
xtmenu       = Trim(Request.QueryString("xtmenu"))
sol          = Trim(Request.QueryString("sol"))
cantregmover = SanitizeInt(Request.QueryString("cantregmover"))
vuelta       = Trim(Request.QueryString("vuelta"))

' --- Armado de la Redirección (Vuelta) ---
Select Case opt
    Case "actinc"
        vuelta = vuelta & "&id=" & id & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu
    Case "actproy"
        vuelta = vuelta & "&id=" & id & "&sol=" & sol & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu
    Case "actapl"
        vuelta = vuelta & "&id=" & id & "&cli=" & cli & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&cantregmover=" & cantregmover
End Select

' --- Conexión e Invocación del Stored Procedure ---
Dim Cnn, Cmd
Set Cnn = Server.CreateObject("ADODB.Connection")
'Cnn.Open "Provider=SQLOLEDB;Data Source=server-2025;Initial Catalog=dbcorgestionv5;User ID=sa;password=bb4515$;"
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

Set Cmd = Server.CreateObject("ADODB.Command")
Cmd.ActiveConnection = Cnn
Cmd.CommandType = 4 ' adCmdStoredProc
Cmd.CommandText = "sp_cg_ActualizarGestion"

' Parámetros Principales
Cmd.Parameters.Append Cmd.CreateParameter("@Opt", 200, 1, 20, opt)
Cmd.Parameters.Append Cmd.CreateParameter("@Id", 3, 1, , id)
Cmd.Parameters.Append Cmd.CreateParameter("@Tabla", 200, 1, 50, tabla)

Dim estadoDesa
estadoDesa = SanitizeInt(Request.Form("estado_desa"))
Cmd.Parameters.Append Cmd.CreateParameter("@EstadoDesa", 3, 1, , estadoDesa)

' Parámetros Condicionales según Operación
If opt = "actapl" Then
    Cmd.Parameters.Append Cmd.CreateParameter("@Nombre", 200, 1, 250, Request.Form("nombre"))
    Cmd.Parameters.Append Cmd.CreateParameter("@IdCliente", 3, 1, , SanitizeInt(Request.Form("id_cliente")))
    Cmd.Parameters.Append Cmd.CreateParameter("@IdAplic", 3, 1, , SanitizeInt(Request.Form("id_aplic")))
    Cmd.Parameters.Append Cmd.CreateParameter("@ResponEnvio", 200, 1, 100, Request.Form("respon_envio"))
    Cmd.Parameters.Append Cmd.CreateParameter("@FechaEnvio", 3, 1, , ConvertirFechaASerial(Request.Form("fecha_envio")))
    Cmd.Parameters.Append Cmd.CreateParameter("@FechaTest", 3, 1, , ConvertirFechaASerial(Request.Form("fecha_instalacion_test")))
    Cmd.Parameters.Append Cmd.CreateParameter("@FechaProd", 3, 1, , ConvertirFechaASerial(Request.Form("fecha_instalacion_prod")))
    Cmd.Parameters.Append Cmd.CreateParameter("@MotivoApl", 200, 1, -1, Request.Form("motivo_apl"))
    Cmd.Parameters.Append Cmd.CreateParameter("@Incidente", 3, 1, , SanitizeInt(Request.Form("incidente")))
    Cmd.Parameters.Append Cmd.CreateParameter("@Proyecto", 3, 1, , SanitizeInt(Request.Form("proyecto")))
    Cmd.Parameters.Append Cmd.CreateParameter("@Sec", 3, 1, , SanitizeInt(Request.Form("sec")))
    Cmd.Parameters.Append Cmd.CreateParameter("@Tipo", 3, 1, , SanitizeInt(Request.Form("tipo")))
ElseIf opt = "actproy" Then
    Dim fEntrega
    fEntrega = ConvertirFechaASerial(Request.Form("fecha_entrega"))
    Cmd.Parameters.Append Cmd.CreateParameter("@FechaEntrega", 3, 1, , fEntrega)
End If

' Ejecutar el SP
Cmd.Execute

' Liberar Conexión
Set Cmd = Nothing
Cnn.Close
Set Cnn = Nothing
%>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Cor Gestion</title>
    <link rel="stylesheet" href="assets/css/main.css" />
</head>
<body style="background:#2F81C6;">

    <form name="init" id="initForm" method="post" action="<%= Server.HTMLEncode(vuelta) %>">
        <input type="submit" value="INICIAR" style="display:none;" />
    </form>

    <script type="text/javascript">
        alert('<%= Server.HTMLEncode(vuelta) %>');
        alert('Actualizado exitosamente');
        document.getElementById('initForm').submit();
    </script>

</body>
</html>