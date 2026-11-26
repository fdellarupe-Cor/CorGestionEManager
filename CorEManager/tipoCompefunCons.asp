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
  response.write ofv.Repbody(request.querystring("mnn")) & ofv.RepHeader2(ofv.convertircar(request.querystring("mnn"),"_"," "),"600")
  response.write "<table id=margen1 width=600 marginwidth=0  cellspacing=1 cellpadding=2>"
  response.write "<tr>"
  response.write "<td width=1% >&nbsp;</td><td width=10% >&nbsp;</td>"
  response.write "<td width=40% >&nbsp;</td><td width=49% >&nbsp;</td>"
  response.write "</tr></table>"

  'Encabezados


  Formulario=chr(13)&" <TR ><TD  class=tt width=2% >Id</TD><TD  class=tt width=50% >Descripcion</TD><TD  class=tt width=30% >tipo</TD></tr>" & chr(13)


  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From Tipocompetencia Order By descripcion"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  'if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) 'and (posi < CantRegMostrarx)

      Celda =""


    'id
    sInput = rs(0)
    Celda = Celda & ofv.gencelda("","vt","right","50","","","",sInput,"si")

    'Descripcion
    sInput = rs(1) 'ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"100","dt","")
    Celda = Celda & ofv.gencelda("","ut","","","","","",sInput,"si")
   
    select case rs(2)&""
    case "0" 
     xtipo="Libre"
    case "1"
    xtipo="Parametrizado"
    case else
    xtipo="&nbsp;"
    end select  'ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"100","dt","")
       Celda = Celda & ofv.gencelda("","ut","","","","","",xtipo,"si")



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