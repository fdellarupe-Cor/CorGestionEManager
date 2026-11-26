<%
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->

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
sel = request.querystring("sel")
response.write "FORM: " & request.form & "<br>" & "STRING: " & request.querystring

		html = html & "<div>"

		html = html & "<header id='header' style='height:20%'>"

		html = html & "<div>"

		html = html & "<table style='width:100%;'><tr>"

		html = html & "<td style='width:75%;' valign=center><a href='index.asp' class='logo' >"
		html = html & "<img src='images/logocorP4.gif' width='10%' >"
		html = html & "<font style='font-size:40pt;align:center;valign:center;font-family:Serif;color:#000000;'>Cor Consulting</font>"
		html = html & "</a></td>"

		html = html & "</tr></table>"

		html = html & "</div>"

		html = html & "</header>"

		html = html & "</div>"

		html = html & "<div style='height:8%;width:100%;border-top:solid 5px #FF3D3D;border-bottom:double 6px #001BFF;'>"

		html = html & "<table style='height:100%;width:100%;'>"

		html = html & "<tr>"

		html = html & "<td><font style='margin-left:5%;font-size:20pt;font-family:Serif;color:#000000;letter-spacing:0.2em;'><b>TITULO</b></font></td>"

		html = html & "</tr>"

		html = html & "</table>"

		html = html & "</div>"

		html = html & "<div>"

		html = html & "ACA VA TABLA"

		html = html & "</div>"

		'sql = "select a.id,b.descripcion,c.descripcion,f.descripcion,g.descripcion,a.fecha_ingreso,h.nombre,i.nombre "
		'sql = sql & "from incidentes a "
		'sql = sql & "inner join clientes b on (a.id_cliente=b.id) "
		'sql = sql & "inner join prioridades c on (a.prioridad=c.id) "
		'sql = sql & "inner join CG_ESTADOS f on (a.estado_desa=f.id) "
		'sql = sql & "inner join CG_ESTADOS g on (a.estado_func=g.id) "
		'sql = sql & "inner join CG_COR_RESPONSABLES h on (a.respon_gestion = h.id) "
		'sql = sql & "inner join CG_COR_RESPONSABLES i on (a.respon_tec = i.id) "
		'sql = sql & "where a.id_cliente=" & cliente
		'Set rs = Cnn.Execute(sql)



html = html & "</body></html>"

response.write html
%>