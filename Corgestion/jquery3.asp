<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<%
Set Cnn = Server.CreateObject("ADODB.Connection")
Cnn.open "Provider=SQLOLEDB;Data Source=server-2025;Initial Catalog=dbcorgestionv5; user id=sa;password=bb4515$;"

shtml = "<html>"

shtml = shtml & "<head>"

shtml = shtml & "<meta charset='utf-8'>"
shtml = shtml & "<meta name='viewport' content='width=device-width, initial-scale=1'>"
shtml = shtml & "<title>jQuery test Mauro</title>"
shtml = shtml & "<link rel='stylesheet' href='assets/css/main2.css' />"
shtml = shtml & "<link rel='stylesheet' href='//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css'>"
shtml = shtml & "<script type='text/javascript' src='assets/js/jquery.min.js'></script>"
shtml = shtml & "<script type='text/javascript' src='assets/js/jquery-ui.js'></script>"

shtml = shtml & "<script>"

shtml = shtml & "$( function() {"

shtml = shtml & "$( '#li1' ).click(function() {"
shtml = shtml & "alert('hola');"
shtml = shtml & "});"

shtml = shtml & "});"

shtml = shtml & "</script>"

shtml = shtml & "</head>"






shtml = shtml & "<body style='width:100%;height:100%;'> "



shtml = shtml & "</body>"

shtml = shtml & "</html>"

response.write shtml
%>
