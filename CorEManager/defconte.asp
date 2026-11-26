<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%
ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
  response.write shtml 
else
  session("Nomodificar") = "N"
  response.redirect "edestrprtymst.asp?codobjeto=0&tipobjeto=SYSTEM&tipo=TITULO"
end if

%>