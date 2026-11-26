<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->

<%


ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  CantRegMostrarx = 8
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))





       response.write "<table  marginwidth=0  cellspacing=0 cellpadding=0>"
       response.write "<tr>"
       response.write "<td width=100% >"
    response.write ofv.Repbody(request.querystring("mnn")) & ofv.RepHeader2(ofv.convertircar(request.querystring("mnn"),"_"," "),"600")
       response.write "<table id=margen1 width=600 marginwidth=0  cellspacing=1 cellpadding=2>"
       response.write "<tr>"
       response.write "<td width=10% >&nbsp;</td>"
       response.write "<td width=90% >&nbsp;</td>"
       response.write "</tr></table>"

  'Encabezados

  Formulario=chr(13)&" <TR ><TD  class=tt width=10% >ESTADO</TD><TD  class=tt width=40% >DESCRIPCION</TD><TD  class=tt width=10% >ENVIA BCRA</TD><TD  class=tt width=40% >BOTONES</TD></tr>" & chr(13)

  if request.querystring("seg") <> "" then
    Session("Form") = ofv.convertircar(request.querystring("seg"),"_"," ")
  end if

  set cn = ofv.conectar(ofv.strconn2)



 



  WKSETTABLA = "ORESTADOS"


  Strsql = "Select * From " & WKSETTABLA  & " WHERE TIPO <> '0' ORDER BY DESCRIPCION"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  posI = 0
  Do While (Not rs.EOF)
    sForm = "form" & posi


      Celda = ""

       sInput = "<B>" & RS(1) & "</B>"
      Celda = Celda & ofv.gencelda("","cc","","","","","",sInput,"si")

       sInput = "<B>" & RS(2) & "</B>"
      Celda = Celda & ofv.gencelda("","cc","","","","","",sInput,"si")


       sInput = "<B>" & RS(14) & "</B>"
      Celda = Celda & ofv.gencelda("","cc","","","","","",sInput,"si")

       sInput = "&nbsp;"
      Celda = Celda & ofv.gencelda("","cc","","","","","",sInput,"si")

    Formulario = Formulario & ofv.genrow("","","","","","","",Celda,"si")


'''''''''''''''''''''''

  FOR POSI=1 TO 10


    IF rs(3 + POSI) > 0 THEN 


       Strsql = "Select * From ORESTADOS WHERE OBJID = 0" & rs(3 + POSI)
       set rsxx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rsxx.eof then

          WKTEVENTO = RSxx(2)

       end if



      WKFCOLOR = "#99CCDD"
      Celda = ""

    sInput = "&nbsp;"
    Celda = Celda & ofv.gencelda("","uTi colspan=2 ","","","","","",sInput,"si")


    'BOTON
  '  sInput = "BOTON&nbsp;" & POSI
    Celda = Celda & ofv.gencelda("","uTi   ","","","","","",sInput,"si")

    'descripcion


    sInput = "<B>" & WKTEVENTO & "</B>"
    Celda = Celda & ofv.gencelda("","UTI style='border-style:solid;border-width:1;border-color:#6699cc;' ","","","","","",sInput,"si")

    Formulario = Formulario & ofv.genrow("","AAA  ","","","","","",Celda,"si")


   end if


  NEXT


       sInput = "<hr>"
      Celda = ofv.gencelda("","ut colspan=4 ","","","","","",sInput,"si")

    Formulario = Formulario & ofv.genrow("","","","","","","",Celda,"si")

'''''''''''''''''''''''







    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si")



  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML





%>
