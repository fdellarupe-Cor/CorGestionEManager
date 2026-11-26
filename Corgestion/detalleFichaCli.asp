<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

html = headhtml()
html = html & "<body>"

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

id = request.querystring("id")
nom = request.querystring("nom")

html = html & "<div>"

html = html & "<header id='header' style='height:11%'>"

html = html & "<div>"

html = html & "<table style='width:100%;'><tr>"

html = html & "<td style='width:75%;' valign=center><a href='index.asp' class='logo' >"
html = html & "<img src='images/logocorP4.gif' width='10%' >"
html = html & "<font style='margin-left:100px;font-size:30pt;align:center;valign:center;font-family:Serif;color:#000000;'>Cor Consulting</font>"
html = html & "</a></td>"

html = html & "</tr></table>"

html = html & "</div>"

html = html & "</header>"

html = html & "<div style='height:10%;width:100%;border-top:solid 5px #FF3D3D;border-bottom:double 6px #001BFF;'>"

html = html & "<table style='height:100%;width:100%;'>"

html = html & "<tr>"

html = html & "<td style='width:80%;' valign=center><font style='font-size:25pt;font-family:Serif;color:#000000;letter-spacing:0.2em;'><b>" & nom & "</b></font></td>"

formx5 = "<form id='formx5' style='height:1px;' method='post' action='clientes.asp'></form>"
botonx5 = "<input class='bot' style='height:35px;font-size:13pt;' type='button' value='VOLVER' onclick=""document.getElementById('formx5').submit();"">"
html = html & "<td style='width:20%;' valign=center align=center>" & botonx5 & "</td>"

html = html & "</tr>"

html = html & "</table>"

html = html & "</div><br>"


html = html & "<div style='height:70%;width:100%;'>"

html = html & "<table style='width:50%;' align=center>"

html = html & "<tr>"
html = html & "<td class='encabe' colspan=2 align=center>Detalle ficha cliente</td>"
html = html & "</tr>"

html = html & "<tr>"
html = html & "<td class='regis'>Col1</td><td class='regis'>Dato1</td>"
html = html & "</tr>"

html = html & "<tr>"
html = html & "<td class='regis'>Col2</td><td class='regis'>Dato2</td>"
html = html & "</tr>"

html = html & "<tr>"
html = html & "<td class='regis'>Col3</td><td class='regis'>Dato3</td>"
html = html & "</tr>"

html = html & "</table>"

html = html & "</div>"

html = html & "<script type='text/javascript'>"
html = html & " function actu(){ "
html = html & " document.getElementById('upd').submit();"
html = html & " alert('Cliente actualizado');} "
html = html & "</script>"

html = html & "<script src='assets/js/jquery.min.js'></script>"
html = html & "<script src='assets/js/skel.min.js'></script>"
html = html & "<script src='assets/js/util.js'></script>"
html = html & "<script src='assets/js/main.js'></script>"

html = html & formx5 & "</body>"
html = html & "</html>"

response.write html
%>