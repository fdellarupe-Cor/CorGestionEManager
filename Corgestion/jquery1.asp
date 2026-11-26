<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<%
Set Cnn = Server.CreateObject("ADODB.Connection")
Cnn.open "Provider=SQLOLEDB;Data Source=server-2025;Initial Catalog=dbcorgestionv5; user id=sa;password=bb4515$;"
shtml = "<html>"
shtml = shtml & "<head>"

shtml = shtml & "<meta charset='utf-8'>"
shtml = shtml & "<meta name='viewport' content='width=device-width, initial-scale=1'>"

shtml = shtml & "<title>jQuery test Mauro</title>"

shtml = shtml & "<link rel='stylesheet' href='//code.jquery.com/ui/1.12.1/themes/base/jquery-ui.css'>"
shtml = shtml & "<link rel='stylesheet' href='/resources/demos/style.css'>"

shtml = shtml & "<style>"

shtml = shtml & "ul#sortable { "
shtml = shtml & "float:left; "
shtml = shtml & "width:100%; "
shtml = shtml & "margin:0px; "
shtml = shtml & "padding:0px; "
shtml = shtml & "} "

shtml = shtml & "ul#sortable li { "
shtml = shtml & "float:left; "
shtml = shtml & "width:200px; "
shtml = shtml & "margin-right:10px; "
shtml = shtml & "margin-top:10px; "
shtml = shtml & "background-color:#CCCCCC; "
shtml = shtml & "border-radius:4px; "
shtml = shtml & "padding:5px; "
shtml = shtml & "box-shadow: 5px 5px 2px #888; "
shtml = shtml & "cursor:pointer; "
shtml = shtml & "list-style:none; "
shtml = shtml & "} "

shtml = shtml & "</style>"

shtml = shtml & "<script type='text/javascript' src='assets/js/jquery.min.js'></script>"
shtml = shtml & "<script type='text/javascript' src='assets/js/jquery-ui.js'></script>"

shtml = shtml & "<script>"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CAPTAR VALUE DEL DATA-
'ui.item.attr('data-id')    ejemplo: <li data-id='10'> este seria el value del data-id (en este caso id)

'VER VALOR DEL ITEM
'ui.item.text()     ejemplo: <li>este seria el valor del text</li>

'BORRAR ITEM DEL DOM
'ui.item.remove();     ejemplo: iria dentro de la funcion jquery del elemento

'TOMAR VALOR DATA- DEL ELEMENTO DE DONDE VIENE EL ITEM ENVIADO
'ui.sender.attr('data-tab')

'FUNCION $.get
'$.get( 'deletear.asp?id=' + ui.item.attr('data-id'), function( data ) {       el get navega al asp, procesa, y al volver ejecuta >>> function(data)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

shtml = shtml & "$( function() {"

shtml = shtml & "$('#destino').sortable({"
shtml = shtml & "connectWith: '#contenedor',"
shtml = shtml & "update : function(event, ui) {"
'shtml = shtml & "alert('Nombre: ' + ui.item.text() + '\nEdad: ' + ui.item.attr('data-age') + '\nSexo: ' + ui.item.attr('data-sex'));"
shtml = shtml & "$.get( 'php.asp?id=' + ui.item.attr('data-id'), function( data ) {"
shtml = shtml & "$('#contenedor').hide();"
shtml = shtml & "$('#destino').hide();"
shtml = shtml & "})"

shtml = shtml & "}"
shtml = shtml & "});"

shtml = shtml & "$('#contenedor').sortable({"
shtml = shtml & "connectWith: '#destino',"
shtml = shtml & "update : function(event, ui) {"
shtml = shtml & "alert('La lista de opciones se ha actualizado');"
shtml = shtml & "}});"

shtml = shtml & "});"

shtml = shtml & "</script>"

shtml = shtml & "</head>"

shtml = shtml & "<body>"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'NOMBRES

shtml = shtml & "<div id='contenedor' style='width:100px;height:10px;border:solid 1px green;float:left;'>"
	shtml = shtml & "<div data-age='37' data-sex='H' style='border:solid 1px red;'>LUIS</div>"
	shtml = shtml & "<div data-age='21' data-sex='H' style='border:solid 1px green;'>OSCAR</div>"
	shtml = shtml & "<div data-age='57' data-sex='M' style='border:solid 1px orange;'>MARIA</div>"
	shtml = shtml & "<div data-age='34' data-sex='H' style='border:solid 1px yellow;'>CARLOS</div>"
	shtml = shtml & "<div data-age='45' data-sex='M' style='border:solid 1px brown;'>BARBARA</div>"
shtml = shtml & "</div>"

'''''''''''''''''''''''''''''

'DIV
shtml = shtml & "<div id='destino' style='width:500px;height:500px;border:solid;float:right;'>"

shtml = shtml & "</div>"

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

shtml = shtml & "</body>"
shtml = shtml & "</html>"

response.write shtml
%>
