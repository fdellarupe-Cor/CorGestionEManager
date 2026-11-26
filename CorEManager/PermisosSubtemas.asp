<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

response.expires=0
Id = request.querystring("ID")
volver = request.querystring("Volver")

sHTML = "<html><head><title>COR EManager - Permisos de Subtema</title></head>" & chr(13)

sHTML = sHTML & "<frameset rows=*,2 >"
sHTML = sHTML & "<FRAME SRC=" & chr(34) & "subtemaskr.asp?vlt=" & volver & "&id=" & Id
sHTML = sHTML & chr(34)
sHTML = sHTML & " NAME=Barra SCROLLING=no>"
sHTML = sHTML & "<FRAME NAME=Principal "
sHTML = sHTML & "src=" & chr(34) & "subtemaskrdet.asp?id=" & Id
sHTML = sHTML & chr(34)
sHTML = sHTML & " SCROLLING=auto></frameset>"
sHTML = sHTML & "<body></body></html>"

response.write sHTML

%>