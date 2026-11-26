<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

response.expires=0
  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")
  sID = request.querystring("ID")

sHTML = "<html>" & chr(13)
sHTML = sHTML & "<frameset rows=30,* FRAMESPACING=0  FRAMEBORDER=0 BORDER=0 >"
sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 SRC=" & chr(34) & "EdEstrHead.asp?"
sHTML = sHTML & "Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&ID=" & sID & chr(34)
sHTML = sHTML & " NAME=est1 SCROLLING=no>"
sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME=est "
sHTML = sHTML & "SRC=" & chr(34) & "EdEstrdet.asp?"
sHTML = sHTML & "Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&ID=" & sID & chr(34)
sHTML = sHTML & " SCROLLING=auto></frameset>"
sHTML = sHTML & "<body  bgcolor=#eeeeee ><head><title>COR EManager - Edicion de Esquemas Normativos</title></head></body></html>"

response.write sHTML

%>