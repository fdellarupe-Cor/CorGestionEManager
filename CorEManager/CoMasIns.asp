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

  sHTML = ofv.FormBody
  sMenu = Session("menu")

  sHTML = sHTML & "<br><br><br><br><center>"

  'Encabezado sector
  sHTML = sHTML & "<table width=70% cellspacing=1 cellpadding=3 border ><tr><td width=100% "
  sHTML = sHTML & "class=tt align=center colspan=3 style={border-style:outset;border-width:2;} >Origen de Documentacion Adicionales</td></tr>"

  sHTML = sHTML & "<tr><td width=40% "
  sHTML = sHTML & "class=tt align=center style={border-style:outset;border-width:2;} >Origen</td>"
  sHTML = sHTML & "<td width=40% "
  sHTML = sHTML & "class=tt align=center style={border-style:outset;border-width:2;} >Instrumento</td>"
  sHTML = sHTML & "<td width=20% "
  sHTML = sHTML & "class=tt align=center style={border-style:outset;border-width:2;} >Identificacion</td></tr>"


  set cn = ofv.conectar(ofv.strconn2)
  Strsql = "Select * From Objprtyadv Where codobjeto=" & nCodigo & " Order By secuen"

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  do until rs.eof
    sHTML = sHTML & "<tr class=cc>"

    sHTML = sHTML & "<td  class=cc>"
      'Texto  origen
      sOrigen = ""
      if cstr(rs(3)) <> "" then
        Strsql = "Select descri From Origen Where objid = " & rs(3)
        set rsOrigen = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsOrigen.eof then sOrigen = rsOrigen(0)
        call ofv.cerrarconsulta(rsOrigen)
      end if

    SArma = "<B>" & sOrigen & "</B>"
    sHTML = sHTML & SArma

    sHTML = sHTML & "</td>"

    sHTML = sHTML & "<td  class=cc>"
      'Texto  origen
      sInstrumento = ""
      if cstr(rs(4)) <> "" then
        Strsql = "Select descri From Instrumen Where objid = " & rs(4)
        set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsInstrumento.eof then sInstrumento = rsInstrumento(0)
        call ofv.cerrarconsulta(rsInstrumento)
      end if
    SArma = "<B>" & sInstrumento & "</B>"
    sHTML = sHTML & SArma

    sHTML = sHTML & "</td>"

    sHTML = sHTML & "<td  class=cc>"
      'Texto  identificacion

    SArma = "<B>" & rs(5) & "</B>"
    sHTML = sHTML & SArma

    sHTML = sHTML & "</td>"

    sHTML = sHTML & "</tr>"

    rs.movenext
  loop
  sHTML = sHTML & "</table>"

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</center></body></html>"
end if

response.write sHTML

%>