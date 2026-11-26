<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
  Response.Write sHTML
else
  response.expires = 0

  solapa = request.querystring("solapa")

  isol = instr(solapa,".")
  if isol > 0 then
     isol = len(solapa) - isol
     usol = right(solapa,isol)
     if Ucase(usol) = "TIFF" or Ucase(usol) = "JPG" or Ucase(usol) = "JPEG" or Ucase(usol) = "GIF" or Ucase(usol) = "TIFF" or Ucase(usol) = "BMP" then
        vasolo = "no"
     else
        vasolo = "si"
     end if
  else
     vasolo = "si"
  end if

  if vasolo = "si" then
    ''' response.write "legajos/" & session("NCliente") & "/" & solapa
     response.redirect "legajos/" & session("NCliente") & "/" & solapa
  else
    
    sHTML = "<html><head><title>COR ESolution Suite 2.8.2 </title></head><table bgcolor=#f7eedd width=100% height=100% ><tr>"
    sHTML = sHTML & "<td align=center valign=middle style='filter:shadow(color=#888888,direction=125);' >"
    sHTML = sHTML & "<img valign=middle src='legajos/" & session("NCliente") & "/" & solapa & "' border=1 >"
    sHTML = sHTML & "</td></tr></table></body></html>"
    Response.Write sHTML
  end if

end if



%>