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

shtml = "<html>"

shtml = shtml & "<head>"




shtml = shtml & "<meta charset='utf-8'>"
shtml = shtml & "<meta name='viewport' content='width=device-width, initial-scale=1'>"
shtml = shtml & "<title>jQuery test Mauro</title>"
shtml = shtml & "<link rel='stylesheet' href='assets/css/main2.css' />"
shtml = shtml & "<link rel='stylesheet' href='assets/js/jquery-ui.css'>"
shtml = shtml & "<script type='text/javascript' src='assets/js/jquery.min.js'></script>"
shtml = shtml & "<script type='text/javascript' src='assets/js/jquery-ui.js'></script>"
shtml = shtml & "<script type='text/javascript' src='assets/js/jquery.ui.touch-punch.min.js'></script>"






shtml = shtml & "<style>"

shtml = shtml & "@media all and (min-width: 1000px){"
shtml = shtml & "td{"
shtml = shtml & "font-size: 2vw;"
shtml = shtml & "}"
shtml = shtml & "}"

shtml = shtml & "</style>"







shtml = shtml & "<script>"

shtml = shtml & "$( function() {"

shtml = shtml & "var ww = $(window).width();"
shtml = shtml & "var wh = $(window).height();"

shtml = shtml & "$('body').css({"
shtml = shtml & "'width': ww,"
shtml = shtml & "'height': wh,"
shtml = shtml & "'margin': '0',"
shtml = shtml & "'background-image': 'linear-gradient(#BCEAFF, #FFFFFF)',"
shtml = shtml & "'font-family': 'Comic Sans MS'"
shtml = shtml & "});"

shtml = shtml & "$('#tabla1').css({"
shtml = shtml & "'width': $('body').width(),"
shtml = shtml & "'height': ($('body').height() / 3) * 2,"
shtml = shtml & "'border-spacing': '0',"
shtml = shtml & "'padding': '10'"
shtml = shtml & "});"

shtml = shtml & "$('#tabla2').css({"
shtml = shtml & "'width': $('body').width(),"
shtml = shtml & "'height': $('body').height() / 3,"
shtml = shtml & "'border-spacing': '0',"
shtml = shtml & "'padding': '10'"
shtml = shtml & "});"

shtml = shtml & "$('#tabla3').css({"
shtml = shtml & "'width': '100%',"
shtml = shtml & "'height': '100%',"
shtml = shtml & "'border-spacing': '0',"
shtml = shtml & "'padding': '10'"
shtml = shtml & "});"

shtml = shtml & "$('#tabla4').css({"
shtml = shtml & "'width': '100%',"
shtml = shtml & "'height': '100%',"
shtml = shtml & "'border-spacing': '0',"
shtml = shtml & "'padding': '10'"
shtml = shtml & "});"

shtml = shtml & "$('#t1c1, #t1c2, #t1c3').css({"
shtml = shtml & "'width': '22%'"
shtml = shtml & "});"

shtml = shtml & "$('#t1c4').css({"
shtml = shtml & "'width': '34%'"
shtml = shtml & "});"

shtml = shtml & "$('#t2c1, #t2c2, #t2c3').css({"
shtml = shtml & "'width': $('#tabla2').width() / 3,"
shtml = shtml & "'height': $('#tabla2').height()"
shtml = shtml & "});"

shtml = shtml & "$('#t3c1').css({"
shtml = shtml & "'width': '100%',"
shtml = shtml & "'height': $('#tabla3').height() / 3"
shtml = shtml & "});"

shtml = shtml & "$('#t3c2, #t3c3').css({"
shtml = shtml & "'width': '50%',"
shtml = shtml & "'height': $('#tabla3').height() / 3"
shtml = shtml & "});"

shtml = shtml & "$('#t4c1, #t4c2').css({"
shtml = shtml & "'width': '50%',"
shtml = shtml & "'height': $('#tabla4').height()"
shtml = shtml & "});"

shtml = shtml & "$('td').css({"
shtml = shtml & "'text-align': 'center'"
shtml = shtml & "});"

shtml = shtml & "$( '.getLis' ).click(function(e) {"
shtml = shtml & "alert($(this).text());"
shtml = shtml & "});"


shtml = shtml & "});"

shtml = shtml & "</script>"

shtml = shtml & "</head>"






shtml = shtml & "<body>"

'TABLA 1
shtml = shtml & "<table id='tabla1' border>"

shtml = shtml & "<tr>"
shtml = shtml & "<td id='t1c1'>TITULO</td>"
shtml = shtml & "<td id='t1c2' class='getLis'>APLIC</td>"
shtml = shtml & "<td id='t1c3' class='getLis'>CONTACT</td>"
shtml = shtml & "<td id='t1c4' rowspan=4>LISTA ITEMS</td>"
shtml = shtml & "</tr>"

shtml = shtml & "<tr>"
shtml = shtml & "<td>SOLIC PRO</td>"
shtml = shtml & "<td class='getLis'>R GEST</td>"
shtml = shtml & "<td class='getLis'>PRIORI</td>"
shtml = shtml & "</tr>"

shtml = shtml & "<tr>"
shtml = shtml & "<td>F ENTRE</td>"
shtml = shtml & "<td class='getLis'>R TEC</td>"
shtml = shtml & "<td class='getLis'>BROW</td>"
shtml = shtml & "</tr>"

shtml = shtml & "<tr>"
shtml = shtml & "<td class='getLis'>CLIENTE</td>"
shtml = shtml & "<td class='getLis'>VIA PED</td>"
shtml = shtml & "<td class='getLis'>REQ PRESUP</td>"
shtml = shtml & "</tr>"

shtml = shtml & "</table>"

'TABLA 2
shtml = shtml & "<table id='tabla2' border>"

shtml = shtml & "<tr>"

shtml = shtml & "<td id='t2c1'>DETA</td>"

shtml = shtml & "<td id='t2c2'>"
shtml = shtml & "<table id='tabla3' border>"
shtml = shtml & "<tr>"
shtml = shtml & "<td id='t3c1' colspan=2>HORAS ESTIMADAS</td>"
shtml = shtml & "</tr>"
shtml = shtml & "<tr>"
shtml = shtml & "<td id='t3c2'>R Y D</td>"
shtml = shtml & "<td id='t3c3'>DIS</td>"
shtml = shtml & "</tr>"
shtml = shtml & "<tr>"
shtml = shtml & "<td>DESA</td>"
shtml = shtml & "<td>TEST</td>"
shtml = shtml & "</tr>"
shtml = shtml & "</table>"
shtml = shtml & "</td>"

shtml = shtml & "<td id='t2c3'>"
shtml = shtml & "<table id='tabla4' border>"
shtml = shtml & "<tr>"
shtml = shtml & "<td id='t4c1'>TILDE</td>"
shtml = shtml & "<td id='t4c2'>CRUZ</td>"
shtml = shtml & "</tr>"
shtml = shtml & "</table>"
shtml = shtml & "</td>"

shtml = shtml & "</tr>"

shtml = shtml & "</table>"

shtml = shtml & "</body>"

shtml = shtml & "</html>"

response.write shtml
%>
