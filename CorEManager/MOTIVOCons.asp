<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrarx = 10
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

 ' sHTML = ofv.FormHeader(session("FormN"))
  response.write "<table  marginwidth=0  cellspacing=0 cellpadding=0>"
  response.write "<tr>"
  response.write "<td width=100% >"
  response.write ofv.Repbody(session("FormN")) & ofv.RepHeader2(ofv.convertircar(session("FormN"),"_"," "),"600")
  response.write "<table id=margen1 width=600 marginwidth=0  cellspacing=1 cellpadding=2>"
  response.write "<tr>"
  response.write "<td width=1% >&nbsp;</td><td width=10% >&nbsp;</td>"
  response.write "<td width=40% >&nbsp;</td><td width=49% >&nbsp;</td>"
  response.write "</tr></table>"

  'Encabezados
  'Formulario = ofv.Encabezados("Segmento","Descripcion")

  Formulario=chr(13)&" <TR ><TD  class=tt width=30% >Descripcion</TD></tr>" & chr(13)


  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From FORMADMMOTIVO Order By DESCRIPCION "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  'if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF)

      Celda =""




    'descrip'
    sInput = rs(2)
    Celda = Celda & ofv.gencelda("","ut","","90%","","","",sInput,"si")

    '
    Fila = ofv.genrow("","","","","","","",Celda,"si")


    Formulario = Formulario & fila

    posi = posi + 1
    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>