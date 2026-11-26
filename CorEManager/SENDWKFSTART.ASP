<%
 idpedido=request.querystring("idpedido")
 sTipo = request.querystring("Tipobjeto")


 session("NCliente")=IDPEDIDO
 session("TCliente")=sTipo

 RESPONSE.REDIRECT "EDWKFSTART.ASP?MODAL=S&ACT=S&OPCION=AUTO&acc=PEDIDOOK"

%>