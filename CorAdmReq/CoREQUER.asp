<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires = 0



    sHTML = "<html><head><title>COR ESolution Suite 2.8.2 </title></head>"
    sHTML = sHTML & "<frameset cols=25%,* >"
    sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME=Index SCROLLING=no "
    sHTML = sHTML & "src=REQUERIndex.asp style='border-style:solid;border-width:1;border-color:#6699cc;' ><FRAME MARGINWIDTH=0 MARGINHEIGHT=0 "
    sHTML = sHTML & "NAME=Cont SCROLLING=yes src=EDREQUER.asp style='border-style:solid;border-width:1;border-color:#6699cc;' ></frameset><noframes><body>"
    sHTML = sHTML & "<p>Esta página usa marcos, pero su explorador no los admite.</p>"
    sHTML = sHTML & "</body></noframes></html>"

end if

Response.Write sHTML

%>