<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<%

html = "<html>"
html = html & "<head>"

html = html & "<title>Cor Gestion</title>"
html = html & "<meta charset='utf-8' />"
html = html & "<meta name='viewport' content='width=device-width, initial-scale=1' />"
html = html & "<link rel='stylesheet' href='assets/css/main.css' />"
html = html & "<style style='text/css'>"
html = html & ".tablX{"
html = html & "width:100%; "
html = html & "border-collapse:collapse; "
html = html & "}"
html = html & ".tablX td{ "
html = html & "padding:7px; "
html = html & "}"
html = html & ".tablX tr{"
html = html & "background: #FFFFFF;"
html = html & "}"
html = html & ".tablX tr:hover {"
html = html & "background-color: #D4D4D4;"
html = html & "}"
html = html & "</style>"




html = html & "</head>"
html = html & "<body>"
'response.write request.querystring & "<br>" & request.form
Set Cnn = Server.CreateObject("ADODB.Connection")
Cnn.open "Provider=SQLOLEDB;Data Source=server-2025;Initial Catalog=dbcorgestionv5; user id=sa;password=bb4515$;"

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


	StrElNombre = ""
	StrCancion1 = ""
	StrCancion2 = ""
	StrCancion3 = ""

	StrElNombre = Request.form("Nombre")
	StrCancion1 = Request.form("Cancion1")
	StrCancion2 = Request.form("Cancion2")
	StrCancion3 = Request.form("Cancion3")

	Response.write "Hola " & StrElNombre & "! <br><br>"
	Response.write "Tus 3 canciones preferidas son: <br><br>"
	Response.write "<table border='1' align='center'> "
	Response.write "<tr><td>&nbsp;01&nbsp;</td><td>&nbsp;" & StrCancion1 & "&nbsp;</td></tr>"
	Response.write "<tr><td>&nbsp;02&nbsp;</td><td>&nbsp;" & StrCancion2 & "&nbsp;</td></tr>"
	Response.write "<tr><td>&nbsp;03&nbsp;</td><td>&nbsp;" & StrCancion3 & "&nbsp;</td></tr>"
	Response.write "</table>"

	response.write "<br><br><table border='1'><tr><td>" & "hola" & "</td></tr></table>"



html = "</body>"
html = html & "</html>"



response.write html
%>