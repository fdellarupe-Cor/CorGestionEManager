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

sParam = request.querystring("param")

if sParam = "licli" then

	sql = "sp_cg_Catalogo_Listar @Catalogo='clientes'"

elseif sParam = "liaplic" then

	sql = "sp_cg_Catalogo_Listar @Catalogo='aplicativos'"

elseif sParam = "lirg" then

	sql = "sp_cg_Responsables @Modo='combo_funcion', @Funcion='FUNC'"

elseif sParam = "livp" then

	sql = "sp_cg_Catalogo_Listar @Catalogo='via_pedido'"

elseif sParam = "librow" then

	sql = "sp_cg_Catalogo_Listar @Catalogo='browsers'"

elseif sParam = "liprio" then

	sql = "sp_cg_Catalogo_Listar @Catalogo='prioridades'"

elseif sParam = "lirt" then

	sql = "sp_cg_Responsables @Modo='combo_funcion', @Funcion='TEC'"

elseif sParam = "licon" then

	sql = "sp_cg_Combo_Listar @Origen='contacto'"

end if


if sParam <> "lirp" then

	Set rc = Exec(Cnn, sql)
	do while not rc.eof

		shtml = shtml & "<li class='losli' data-padre='" & sParam & "' data-id='" & rc(0) & "' >" & rc(1) & "</li>"

	rc.movenext
	loop

else

	shtml = shtml & "<li id='1' class='losli' data-padre='lirp' data-id='NO'>NO</li>"
	shtml = shtml & "<li id='2' class='losli' data-padre='lirp' data-id='SI'>SI</li>"

end if

response.write shtml
%>