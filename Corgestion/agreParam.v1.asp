<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<%
'Option Explicit
On Error Resume Next

Dim Cnn, sParam, sDato1, sDato2, sDato3, sAccion, rs

sParam  = Trim(Request.QueryString("param"))
sDato1  = Trim(Request.QueryString("dato1"))
sDato2  = Trim(Request.QueryString("dato2"))
sDato3  = Trim(Request.QueryString("dato3"))
sAccion = Trim(Request.QueryString("accion"))

Set Cnn = Server.CreateObject("ADODB.Connection")
On Error Resume Next
Cnn.Open scon1
If Err.Number <> 0 Then
	Response.Write "<div style='color:red;font-family:monospace;padding:10px;border:1px solid red;margin:10px;'>"
	Response.Write "<b>ERROR DE CONEXION A BASE DE DATOS</b><br>"
	Response.Write "Codigo: " & Err.Number & "<br>"
	Response.Write "Detalle: " & Server.HTMLEncode(Err.Description) & "</div>"
	Response.End
End If
On Error Goto 0

If sParam = "APL" And sAccion = "preview" Then
    ' PASO 1: Consulta de datos y respuesta para el Popup
    Set rs = ConsultarRelacionAPL(Cnn, sDato1, sDato2, sDato3)
    
    If Err.Number <> 0 Then
        Response.Write "ERROR: " & Err.Description
    ElseIf Not rs.EOF Then
        ' Devolvemos una cadena formateada con los datos a confirmar
        Response.Write "OK|" & rs("Cliente") & "|" & rs("Aplicativo") & "|" & rs("Detalle")
    Else
        Response.Write "ERROR: No se encontraron los datos especificados."
    End If
    
    If Not rs Is Nothing Then rs.Close : Set rs = Nothing

ElseIf sAccion = "guardar" Or sParam <> "APL" Then
    ' PASO 2: Inserción directa
    Call InsertarDatoAuxiliar(Cnn, sParam, sDato1, sDato2, sDato3)
    
    If Err.Number <> 0 Then
        Response.Write "ERROR: " & Err.Description
    Else
        Response.Write "SUCCESS"
    End If
End If

Cnn.Close
Set Cnn = Nothing
%>
