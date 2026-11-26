<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

response.expires = 0

sTipo = request.querystring("Tipo")
sNombre = request.querystring("Nombre")
sNombre = ofv.convertircar(sNombre,"_"," ")
sMenu = request.querystring("Menu")




%>

<html>
<TITLE>CorEManager 2.8 - Novedades</title>
<frameset rows="*,40" >
<FRAME  MARGINWIDTH=0 MARGINHEIGHT=0 SRC="novedmenu.asp" NAME="tope" SCROLLING=auto>
<FRAME  MARGINWIDTH=0 MARGINHEIGHT=0 NAME="data" sRC="nulo2.asp" SCROLLING=auto>
</frameset>
</body></html>