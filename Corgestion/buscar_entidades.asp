<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<%
' --- CONFIGURACIÓN ---
Dim qTermino, qTabla, qVolverA, qCampo
qTermino = Request.QueryString("t") ' Término de búsqueda
qTabla = Request.QueryString("tabla") ' Tabla objetivo (AGENDA, INCIDENTE, PROYECTO)
qVolverA = Request.QueryString("v") ' URL de retorno (opcional)
qCampo = Request.QueryString("campo") & "" ' Campo del formulario padre que recibirá el ID

// Conexión - usar tu función existente
Dim Cnn, rs

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


' --- LÓGICA DE BÚSQUEDA ---
Dim  cmd
Set cmd = Server.CreateObject("ADODB.Command")
cmd.ActiveConnection = Cnn
cmd.CommandType = 4 ' adCmdStoredProc
cmd.CommandText = "sp_cg_BuscarEntidades"

' Parámetros del SP
cmd.Parameters.Append cmd.CreateParameter("@TerminoBusqueda", 202, 1, 100, qTermino)
cmd.Parameters.Append cmd.CreateParameter("@TablaObjetivo", 202, 1, 50, qTabla)

Set rs = cmd.Execute()

' --- PÁGINA HTML ---
Response.Write "<!DOCTYPE html>"
Response.Write "<html>"
Response.Write "<head>"
Response.Write "<meta charset=""utf-8"">"
Response.Write "<title>Búsqueda de Entidades</title>"
Response.Write "<style>"
Response.Write "  body { font-family: Arial, sans-serif; margin: 20px; }"
Response.Write "  .contenido { max-width: 800px; margin: 0 auto; }"
Response.Write "  table { width: 100%; border-collapse: collapse; margin-top: 20px; }"
Response.Write "  th, td { border: 1px solid #ddd; padding: 8px; text-align: left; }"
Response.Write "  th { background-color: #f2f2f2; }"
Response.Write "  .btn-seleccionar { padding: 5px 10px; background: #28a745; color: white; border: none; cursor: pointer; }"
Response.Write "</style>"
Response.Write "</head>"
Response.Write "<body>"
Response.Write "<div class=""contenido"">"

' --- Cabecera ---
Response.Write "<h2>?? Búsqueda de Entidades</h2>"
Response.Write "<p>Término: <strong>" & Server.HTMLEncode(qTermino) & "</strong></p>"

' --- Formulario de búsqueda ---
Response.Write "<form method=""get"" action=""./buscar_entidades.asp"">"
Response.Write "<input type=""hidden"" name=""campo"" value=""" & Server.HTMLEncode(qCampo) & """/>"
Response.Write "<input type=""hidden"" name=""tabla"" value=""" & Server.HTMLEncode(qTabla) & """/>"
Response.Write "<div style=""margin-bottom: 20px;"">"
Response.Write "<label>Buscar en: <strong>" & Server.HTMLEncode(qTabla) & "</strong></label><br/>"
Response.Write "<label>Término:</label> "
Response.Write "<input type=""text"" name=""t"" value=""" & Server.HTMLEncode(qTermino) & """ style=""width: 300px;"" onfocus=""this.select()""/> "
Response.Write "<button type=""submit"" style=""margin-left: 10px;"">Buscar</button>"
Response.Write "</div>"
Response.Write "</form>"

' --- Mostrar resultados ---
Response.Write "<h3>Resultados</h3>"
If rs.Fields.Count = 0 Then
    Response.Write "<p>Ingrese un término para iniciar la búsqueda.</p>"
ElseIf rs.EOF And rs.BOF Then
    Response.Write "<p>No se encontraron resultados.</p>"
Else
    Response.Write "<table>"
    Response.Write "<thead>"
    Response.Write "<tr>"
    Response.Write "<th>#</th>"
    Response.Write "<th>Entidad</th>"
    Response.Write "<th>Tipo</th>"
    Response.Write "<th>Acción</th>"
    Response.Write "</tr>"
    Response.Write "</thead>"
    Response.Write "<tbody>"
    
    Dim i
    i = 0
    Do While Not rs.EOF
        i = i + 1
        Dim idRegistro, nombreRegistro, tipoRegistro
        idRegistro = rs.Fields("ID").Value
        nombreRegistro = rs.Fields("Nombre").Value
        tipoRegistro = rs.Fields("Tipo").Value
        
        Response.Write "<tr>"
        Response.Write "<td>" & i & "</td>"
        Response.Write "<td title=""" & Server.HTMLEncode(nombreRegistro) & """>" & Left(nombreRegistro, 80) & "</td>"
        Response.Write "<td>" & tipoRegistro & "</td>"
        Response.Write "<td>"
        ' Botón para seleccionar y devolver el ID al campo indicado en la ventana padre.
        Response.Write "<button class=""btn-seleccionar"" data-id=""" & Server.HTMLEncode(idRegistro & "") & """ data-nombre=""" & Server.HTMLEncode(nombreRegistro & "") & """ onclick=""seleccionarEntidad(this); return false;"">Seleccionar</button>"
        Response.Write "</td>"
        Response.Write "</tr>"
        
        rs.MoveNext
    Loop
    
    Response.Write "</tbody>"
    Response.Write "</table>"
End If

' --- Cerrar conexión ---
If Not rs Is Nothing Then
    If rs.State <> 0 Then rs.Close
End If
Cnn.Close
Set rs = Nothing
Set Cnn = Nothing

Response.Write "</div>"
Response.Write "<script>function seleccionarEntidad(boton){var op=window.opener;if(!op||typeof op.agendaExtSeleccionarEntidad!=='function'){alert('No se encontró la pantalla de agenda.');return;}op.agendaExtSeleccionarEntidad('" & Replace(qCampo, "'", "\\'") & "',boton.getAttribute('data-id'),boton.getAttribute('data-nombre'));window.close();}</script>"
Response.Write "</body>"
Response.write "</html>"
%>