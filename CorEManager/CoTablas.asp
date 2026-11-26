<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

response.expires=0
  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")

sHTML = "<html>" & chr(13)
sHTML = sHTML & "<frameset rows=40,*  >"
sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 SRC=" & chr(34) & "CoTablas2.asp?"
sHTML = sHTML & "Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & chr(34)
sHTML = sHTML & " NAME=arch1 SCROLLING=auto>"
sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME=arch "
sHTML = sHTML & "src=" & chr(34) & "nulo.asp"
sHTML = sHTML & chr(34)
sHTML = sHTML & " SCROLLING=auto></frameset>"
sHTML = sHTML & "<body  ><head><title>COR EManager - Edicion de Esquemas Normativos</title></head></body></html>"

response.write sHTML

%>