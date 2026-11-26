<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
  response.write sHTML
else

  nCodigo = request.querystring("Codobjeto")
  sTipo = request.querystring("Tipobjeto")

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "SELECT * from csprocmaster "
  Strsql = Strsql & "Where codobjeto = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    nID = rs(0)
    ' if session("APLMDCLI") <> "BC" then
   '  '   response.redirect "Cursograma.asp?Codproced=" & nID & "&Tipobjeto=" & sTipo
   '  else
      '  response.redirect "CursogramaView.asp?Codproced=" & nID & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo

    aspc= "CursogramaView.asp?Codproced=" & nID & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo

        sInput = "<html><head></head><body><form id=curview action='"& aspc &"' method=post target=_parent ></form>"
		sinput=   sinput & "</body></html><script>document.getElementById('curview').submit();</script>"
    response.write sinput
     '  server.transfer( "CursogramaView.asp")'?Codproced=" & nID & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo)
   '  end if
  else

    sInput = "<br><br><br><br><p align=center><font color=#336699 face=arial-narrow size=3><b>"
    sInput = sInput  & Ucase("Sin Información para graficar") & "</b></font></p>"
   response.write sInput
  end if
  call ofv.cerrarconsulta(rs)
end if



%>