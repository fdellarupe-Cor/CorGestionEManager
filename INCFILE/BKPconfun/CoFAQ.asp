<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTFAQ.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")

  set cn = ofv.conectar(ofv.strconn2)

     shtml = ofv.FormHeader(session("FormN")) 



     shtml = shtml & GENFMTFAQ(nCodigo,STIPO)


     shtml = shtml & "</body></html>"


end if

response.write sHTML

%>