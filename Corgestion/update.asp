<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<%
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

sPadre = request.querystring("padre")
sDato = request.querystring("dato")
sTipo = request.querystring("tipo")

if sTipo = "INC" then

	select case sPadre

		case "licli"
			sql = "id_cliente"

		case "liaplic"
			sql = "id_aplic"

		case "lirg"
			sql = "respon_gestion"

		case "lirp"
			sql = "req_presup"

		case "livp"
			sql = "via_pedido"

		case "librow"
			sql = "browser"

		case "liprio"
			sql = "prioridad"

		case "lirt"
			sql = "respon_tec"

		case "licon"
			sql = "contacto"

		case "iTit"
			sql = "titulo"

		case "iDeta"
			sql = "detalles"

	End Select

ExecNonQuery Cnn, "sp_cg_Registro_ActualizarCampos @tabla='INCIDENTES', @id=" & session("ID") & ", @camposJson='[{""col"":""" & sql & """,""val"":""" & sDato & """}]'"

Elseif sTipo = "ELIM" then

	ExecNonQuery Cnn, "sp_cg_Registro_Eliminar @Opt='incf', @Id=" & session("ID")

end if

response.write request.querystring

%>