<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<HTML>
<BODY>
	<DIV style="border:solid 1px black">Hola</div>

</BODY>

</HTML>

<%

SHTML = "<SHTML>"
SHTML = SHTML & "<head>"

SHTML = SHTML & "<title>Cor Gestion</title>"
SHTML = SHTML & "<meta charset='utf-8' />"
SHTML = SHTML & "<meta name='viewport' content='width=device-width, initial-scale=1' />"
SHTML = SHTML & "<link rel='stylesheet' href='assets/css/main.css' />"
SHTML = SHTML & "<style style='text/css'>"
SHTML = SHTML & ".tablX{"
SHTML = SHTML & "width:100%; "
SHTML = SHTML & "border-collapse:collapse; "
SHTML = SHTML & "}"
SHTML = SHTML & ".tablX td{ "
SHTML = SHTML & "padding:7px; "
SHTML = SHTML & "}"
SHTML = SHTML & ".tablX tr{"
SHTML = SHTML & "background: #FFFFFF;"
SHTML = SHTML & "}"
SHTML = SHTML & ".tablX tr:hover {"
SHTML = SHTML & "background-color: #D4D4D4;"
SHTML = SHTML & "}"
SHTML = SHTML & "</style>"




SHTML = SHTML & "</head>"



SHTML = SHTML & "<body>"
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


Response.Write "Se esta ejecutando la pagina <br>" & Request.ServerVariables ("SCRIPT_NAME") & "<br>"

id = request.form("ElId")

response.write id

dim strElRodri
strElRodri = "Rodri"

dim iSinDefinir
iSinDefinir = 1/1

Response.write strElRodri& "<br>"
Response.write iSinDefinir& "<br>"

Dim elDiv
elDiv = "<div style='border:solid black 1px;width:150px;height:50px;border-radius:2.5px;text-align:center;'>Cliente</div><br>"
Response.write elDiv

	set rssql = Server.CreateObject("ADODB.Recordset")
	set rssql = cnn.Execute("sp_cg_Utilidades @Modo='listar_clientes'")

'response.write rssql(1)

	Dim TdClase1
	TdClase1 = "class = 'TdDetClase1'"
	Dim TdDetClase1
'	TdDetClase1 = "style='border:solid 1px red;text-align:right;width:40px;font-family:tahoma;font-size:8pt;"
	TdDetClase1 = "style='height:20px;border:solid 1px white;border-right-color:grey;border-bottom-color:grey;text-align:right;width:40px;font-family:tahoma;font-size:8pt;box-shadow:1px 1px grey;border-radius:2.5px;"
	Dim TdDetClase2
	'TdDetClase2 = "style='border:solid 1px red;text-align:left;width:200px;font-family:tahoma;font-size:8pt;"
	TdDetClase2 = "style='height:20px;border:solid 1px white;border-right-color:#a2cadf;border-bottom-color:#a2cadf;text-align:left;width:200px;font-family:tahoma;font-size:8pt;box-shadow:1px 1px grey;background-color:a8e8f8;color:black;border-radius:2.5px;"
	'#a8e8f8
	TdDetClase3 = "style='height:20px;border:solid 1px white;border-right-color:#a2cadf;border-bottom-color:#a2cadf;text-align:center;width:50px;font-family:tahoma;font-size:8pt;box-shadow:1px 1px grey;background-color:#a8e8f8;border-radius:2.5px;"




		Do while Not rssql.eof

			IDCli = rssql(0) & "&nbsp;&nbsp;"
			DescriCli = "&nbsp;&nbsp;" & rssql(1)
			DirecCli =  "&nbsp;&nbsp;" & rssql(3)

		'response.write descricli
		Dim TablaClientes
		TablaClientes = "<table style='margin: 0 auto'><tr><td " & TdDetClase1 & "'>" & IDCli  & "</td><td " & TdDetClase2 & "'>" & DescriCli & "</td><td " & TdDetClase2 & "'>" & DirecCli & "</td><td " & TdDetClase3 & "'><input type='submit' value='APL'</input></td></tr></table>"

		response.write tablaClientes
		rssql.movenext
		loop




'Dim TablaClientes
'TablaClientes = "<table><tr><td " & TdDetClase1 & "'>" & DescriCli & "</td></tr></table>"
'response.write TablaClientes

SHTML = SHTML & "<DIV style='border:solid 1px black'>Holaque tal</div>"
SHTML = "</body>"
SHTML = SHTML & "</SHTML>"



response.write SHTML
%>