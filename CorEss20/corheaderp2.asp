<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
   ' cc=ConUsrset()
    op=request.querystring("op")
    empre=request.querystring("empre")
    shtml = "<html>" & CorStyler2()
    'shtml = shtml & " <script src='../v48/line-awesome/jquery.min.js'></script>"
    'shtml = shtml & " <script src='../v48/line-awesome/jqheader.js'></script>"
    sHTML = sHTML & "<body class=menu>" & corpaneltop()
    sHTML = sHTML & "<div id='XHDM' >cargame </div></body></html>"
 END IF

 RESPONSE.WRITE SHTML
%>

