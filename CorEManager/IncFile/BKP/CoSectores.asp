<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
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

  sHTML = ofv.FormHeader(session("FormN"))
  sMenu = Session("menu")

  sHTML = sHTML & "<br><br><br><br><center>"

  sImagen4si = "<FONT FACE=corflaticon COLOR=#f7aa22 SIZE=3>G</FONT>"

  'Encabezado sector
  sHTML = sHTML & "<table width=70% cellspacing=0 cellpadding=3   >"
  sHTML = sHTML & "<tr height=30 ><td width=100% "
  sHTML = sHTML & "class=ttt align=center colspan=2 >SECTORES AFECTADOS</td></tr>"

  set cn = ofv.conectar(ofv.strconn2)
  Strsql = "SELECT ECO.descri, OBJSECTOR.*  FROM OBJSECTOR, ECO "
  Strsql = Strsql & "where OBJSECTOR.codsector = ECO.eco and OBJSECTOR.codobjeto=" & nCodigo
  Strsql = Strsql & " Order By Calificador Desc"

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  do until rs.eof
    sHTML = sHTML & "<tr height=35 >"

    sHTML = sHTML & "<td  class=uts  >"
    sHTML = sHTML & sImagen4si & "&nbsp;&nbsp;"
    sHTML = sHTML & "</td>"

      'Texto sectores origen

    sHTML = sHTML & "<td  class=ut width=98% >"
    SArma = "" & rs(0) & ""
    sHTML = sHTML & SArma

    sHTML = sHTML & "</td></tr>"

    rs.movenext
  loop

  sHTML = sHTML & "<tr ><td  class=vt >&nbsp;</td></tr>"
  sHTML = sHTML & "</table>"

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</center></body></html>"
end if

response.write sHTML

%>