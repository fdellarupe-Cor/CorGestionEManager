<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<%
'Option Explicit

Dim Cnn, sdato, sRef, origenSP, rs

' Inicialización de la conexión
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

' Captura y limpieza de parámetros
sdato = Trim(Request.QueryString("param"))
sRef = Trim(Request.QueryString("ref"))
%>


<% If sdato = "CLIENTE" Or sdato = "APLICACION" Or sdato = "VIA PEDIDO" Or sdato = "PRIORIDAD" Or sdato = "BROWSER" Then %>

    <table>
        <tr>
            <td>Descripcion</td>
            <td><input class='inp' type='text' id='valor1' name='valor1'></td>
        </tr>
    </table>

<% ElseIf sdato = "RESPON GESTION" Or sdato = "RESPON TECNICO" Or sdato = "CONTACTO" Then %>

    <table>
        <tr>
            <td>Nombre</td>
            <td><input class='inp' type='text' id='valor1' name='valor1'></td>
        </tr>
        <tr>
            <td>Apellido</td>
            <td><input class='inp' type='text' id='valor2' name='valor2'></td>
        </tr>
    </table>

<% ElseIf sdato = "IOP" Then %>

    <table align="center">
        <tr>
            <td style='width:50%;'>ID</td>
            <td style='width:50%;'><input class='inp' style='width:60px;' type='text' id='valor1' name='valor1'></td>
        </tr>
        <tr>
            <td style='width:50%;'>TIPO</td>
            <td style='width:50%;'>
                <select id='valor2' name='valor2' style='width:60px;height:60%;border:solid 1px black;'>
                    <option value='I'>INC</option>
                    <option value='P'>PROY</option>
                </select>
            </td>
        </tr>
    </table>

<% ElseIf sdato = "APL" Then %>

    <table align="center">
        <tr>
            <td style='width:50%;'>ID</td>
            <td style='width:50%;'><input class='inp' style='width:60px;' type='text' id='valor1' name='valor1'></td>
			<!-- En el formulario desplegado -->
			<input type="hidden" id="hdnIdAPL" name="hdnIdAPL" value="<%=idAPL%>" />
        </tr>
    </table>

<% ElseIf sdato = "EDITAR" Then 
    
    ' 1. Mapeamos el 'sRef' al origen esperado por tu Stored Procedure
    origenSP = ""
    Select Case sRef
        Case "CLIENTE":        origenSP = "id_cliente"
        Case "APLICACION":     origenSP = "id_aplic"
        Case "RESPON GESTION": origenSP = "respon_gestion"
        Case "VIA PEDIDO":     origenSP = "via_pedido"
        Case "PRIORIDAD":      origenSP = "prioridad"
        Case "RESPON TECNICO": origenSP = "respon_tec"
        Case "BROWSER":        origenSP = "browser"
        Case "CONTACTO":       origenSP = "contacto"
    End Select

    ' 2. Si es un origen válido, llamamos a la base de datos y renderizamos los inputs
    If origenSP <> "" Then
        Set rs = GetDatosCombo(Cnn, origenSP, "")
        
        Do While Not rs.EOF
%>
            <input type='text' id='<%=rs("id")%>' name='<%=rs("id")%>' value='<%=rs("descripcion")%>' />
<%
            rs.MoveNext
        Loop
        
        ' Limpieza de objetos del recordset
        rs.Close
        Set rs = Nothing
    End If

End If 

' Limpieza de la conexión
Cnn.Close
Set Cnn = Nothing
%>

<!-- Inclusión del script JavaScript -->
<script src="assets/js/relaciones.js"></script>