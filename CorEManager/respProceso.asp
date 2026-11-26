<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

response.expires=0

sTipo = request.querystring("Tipobjeto")
nCodigo = request.querystring("Codobjeto")






response.write "<br><br><br>" & ncodigo & "-" & stipo & "<br>"



%>