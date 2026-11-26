$(".your-draggable")
	.enableSelection()
	.removeData("draggable")
	.unbind(".draggable")
	.removeClass("ui-draggable ui-draggable-dragging ui-draggable-disabled");

<%

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

shtml = shtml & "$( '#tit' ).mousedown(function() {"
shtml = shtml & "  $( '#mover' ).draggable();"
shtml = shtml & "});"

shtml = shtml & "$( '#tit' ).mouseup(function() {"
shtml = shtml & "  $( '#mover' ).draggable('destroy');"
shtml = shtml & "});"

shtml = shtml & "$('#licli').sortable({"
shtml = shtml & "});"
shtml = shtml & "});"

shtml = shtml & "</script>"

shtml = shtml & "</head>"




shtml = shtml & "<body style='width:100%;height:100%;'> "
shtml = shtml & "<div id=mover style='background-color: coral;width:150px;height:500px' ><div id=tit>xxxxxxxxxxxxx</div><ul  id='licli'>"


for i = 0 to 10
shtml = shtml & "<li class='item' data-padre='licli'>" & i & "</li>"
next

shtml = shtml & "</ul></div>"


shtml = shtml & "</body>"
shtml = shtml & "</html>"

response.write shtml

%>