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
Cnn.open scon1

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


nombre = request.form("Nombre")
edad = request.form("Edad")

response.write nombre & edad


html = "</body>"
html = html & "</html>"



response.write html
%>