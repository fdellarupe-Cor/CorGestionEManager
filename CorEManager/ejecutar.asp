<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->

<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
  response.write sHTML
else
   sHTML = "<html>"
     SPEDIDO=request.querystring("pedido")
     Scodform=request.querystring("codform")
     seg=REQUEST.querystring("seg")
     sDOCMODAL=request.querystring("DOCMODAL")
     sTipo=request.querystring("tipo")
     wxs = "<script language=JavaScript>" & Chr(13)
     wxs = wxs & " var Fbtini = 0;" & Chr(13)
     wxs = wxs & " function IniWin() { " & Chr(13)
     wxs = wxs & "  if (document.all.SSOLACT) { " & chr(13)
     wxs = wxs & " document.SSOLACT.submit(); " & Chr(13)
     wxs = wxs & " } else { " & chr(13)
     wxs = wxs & "           if (Fbtini > 10) {} else { Fbtini++; " & chr(13)
     wxs = wxs & " setTimeout (""IniWin()"", 100); } " & chr(13)
     wxs = wxs & "  } " & chr(13)
     wxs = wxs & " } " & Chr(13)
     wxs = wxs & " </script>" & Chr(13)

    sHTML = sHTML & wxs

  sHTML = sHTML & "<body onload='IniWin();' >"
  		if sTipo = "SN" then
  			xverasp="crearnormativa.asp?pedido=" & SPEDIDO & "&codform=" & SCodform & "&seg=" & seg & "&DOCMODAL=" & sDOCMODAL
  		else
  			xverasp="crearpedido.asp?pedido=" & SPEDIDO & "&codform=" & SCodform & "&seg=" & seg & "&DOCMODAL=" & sDOCMODAL
  		end if
 sHTML1 = sHTML & "<form method=post  name=SSOLACT id=SSOLACT  action=" &chr(34) & xverasp & chr(34) & " target="

               sHTML1 = sHTML1 & "Principal"

            sHTML1 = sHTML1 & " style='display:none;' >" & chr(13)
            sHTML1 = sHTML1 & "<input type=hidden size=10 name=boton value=menu></input>" & chr(13)
            sHTML1 = sHTML1 & "</form>" & chr(13)

response.write sHTML1 & "</body>"
end if
%>