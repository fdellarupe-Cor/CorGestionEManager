<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires = 0

  ID = request.querystring("ID")



  sHTML = "<html><head><title>COR ESolution Suite V4.8 </title></head>"
  sHTML = sHTML & "<frameset rows=120,* >"
  sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME=tope SCROLLING=no "
  sHTML = sHTML & "src=ImagenCont1.asp?ID=" & ID & "><FRAME MARGINWIDTH=0 MARGINHEIGHT=0 "
  sHTML = sHTML & "NAME=dato SCROLLING=yes src=ImagenCont2.asp?ID=" & ID & "></frameset><noframes><body>"
  sHTML = sHTML & "<p>Esta página usa marcos, pero su explorador no los admite.</p>"
  sHTML = sHTML & "</body></noframes></html>"

end if

Response.Write sHTML

%>