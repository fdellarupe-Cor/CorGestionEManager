<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<%

html = headhtml()
html = html & "<body>"
'response.write request.querystring & "<br>" & request.form
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

cons = request.querystring("cons")
opc = request.querystring("opc")
xmenu = request.querystring("xmenu")
xtmenu = request.querystring("xtmenu")
CantRegMostrar = 10
CantRegMover = CDbl(Request.querystring("CantRegMover"))
if CantRegMover < 0 then CantRegMover = 0

nFecini = 0

If nFecini < CLng(CDate("01-01-1900")) Then
	nFecini = CLng(CDate("01-01-1900"))
Else
	nFecini = 0
End If

cliente = request.form("cli")
if cliente & "" = "" then cliente = request.querystring("cli")
if cliente & "" = "" then cliente = 0

ord = request.querystring("ord")

par = request.querystring("par")
if par & "" = "" then par = "asc"
'if par = "asc" then
'	par = "desc"
'else
'	par = "asc"
'end if

Response.Write "Fecha y hora actual es: " & now & "<br><br>"


	Dim StrNombre

	'Lista de los valores de la coleccion Request.Querystring del formulario anterior en apl2.asp

	For Each StrNombre in Request.Querystring
	Response.write StrNombre & ": " & Request.Querystring(StrNombre) & "<br>"
	Response.write "<br>"

	Next

html = "</body>"
html = html & "</html>"



response.write html
%>