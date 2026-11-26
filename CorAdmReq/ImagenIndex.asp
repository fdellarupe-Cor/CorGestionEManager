<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  solapa = request.querystring("solapa")



  set cn = ofv.conectar(ofv.strconn5)


  sHTML = ofv.MenuHeader(session("CMFMENU"),"")

  StrSql = "Select * From ItemV where id = '" & solapa & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then


    sInput = ""
    if rs(7) <> "D" then
       sInput = sinput & rs(3)
    else
       sInput = sinput & rs(5)
    end if


  end if



    sLink = "<b><big>" & sinput & "</big></b>"
    Celda = ofv.GenCelda("","cc colspan=2","","100%","","","",sLink,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")

    sLink = "<hr width=100% size=1 color=#993333 >"
    Celda = ofv.GenCelda("","tt colspan=2","","100%","","","",sLink,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")


  StrSql = "Select * From ItemData where item = '" & solapa & "' "
  strsql = strsql & " order by orden,id"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  i = 0
  do until rs.eof

    sImagen = ofv.GenImage("","","","","imagenes/botrep.jpg","0","")
    sHref = "ImagenCont.asp?solapa=" & rs(6)
    sLink = ofv.GenLink("","tt",sHref,"Cont","","",sImagen,"si")
    Celda = ofv.GenCelda("","tt","","1%","","","",sLink,"si")
    sLink = ofv.GenLink("","cc",sHref,"Cont","","",rs(4),"si")
    Celda = Celda & ofv.GenCelda("","cc","","99%","","","",sLink,"si")
    Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")



    rs.movenext
  loop

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & ofv.GenTabla("","","","100%","","","","","","","",Fila,"si")
  sHTML = sHTML & "</center></body></html>"
end if

Response.Write sHTML

%>