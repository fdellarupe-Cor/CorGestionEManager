<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires = 0

  Marco = request.querystring("Tipo")
  dir = "ftp://" & Ucase(Request.ServerVariables("SERVER_NAME")) & "/CorEss" & Marco


    sHTML = "<html><head><title>COR ESolution Suite V4.8</title></head>"
    sHTML = sHTML & "<frameset rows=10%,* >"
    sHTML = sHTML & "<FRAME MARGINWIDTH=0 MARGINHEIGHT=0 NAME=tope SCROLLING=no "
    sHTML = sHTML & "src=ImgHead.asp?tipo=" & Marco & "><FRAME MARGINWIDTH=0 MARGINHEIGHT=0 "
    sHTML = sHTML & "NAME=dato SCROLLING=yes src=" & dir & " ></frameset><noframes><body>"
    sHTML = sHTML & "<p>Esta página usa marcos, pero su explorador no los admite.</p>"
    sHTML = sHTML & "</body></noframes></html>"

end if

Response.Write sHTML

%>