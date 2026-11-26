<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/reportes.inc" -->
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

   response.write corstyler()


 ' sHTML = ofv.FormHeader(session("FormN"))
'  response.write "<table  marginwidth=0  cellspacing=0 cellpadding=0>"
'  response.write "<tr>"
'  response.write "<td width=100% >"
'  response.write ofv.Repbody(request.querystring("mnn")) & ofv.RepHeader2(ofv.convertircar(request.querystring("mnn"),"_"," "),"600")
'  response.write "<table id=margen1 width=600 marginwidth=0  cellspacing=1 cellpadding=2>"
'  response.write "<tr>"
'  response.write "<td width=1% >&nbsp;</td><td width=10% >&nbsp;</td>"
'  response.write "<td width=40% >&nbsp;</td><td width=49% >&nbsp;</td>"
'  response.write "</tr></table>"

  'Encabezados


 ' Formulario=chr(13)&" <TR ><TD  class=tt width=2% >Id</TD><TD  class=tt width=30% >Descripcion</TD></tr>" & chr(13)


  wwwheader = chr(13) & " <TR ><TD  class=tt width=2% >Id</TD><TD  class=tt width=30% >Descripcion</TD></tr>" & chr(13)


    set oPrtG = new ClsReportes ' server.createobject("CorGeneratorEXLR3.Reportes")
    Set oPrtG.Ofv = Ofv

   oPrtG.RCliLogo = session("clilogo")
   oPrtG.RCliNomb = ofv.convertircar(session("FormN"),"_"," ")
   Formulario = oPrtG.RepHeader2(cstr(""), cstr("600"))



  set cn = ofv.conectar(ofv.strconn2)

  WKSETTABLA = "TIPOPRODUCTO"

  IF WKAPLIC = "OR" THEN WKSETTABLA = "ORTIPOPRODUCTO"



  Strsql = "Select * From " & WKSETTABLA & " Order By descri"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  'if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) 'and (posi < CantRegMostrarx)


  if (POSI=0) OR (POSI > 34) then

     IF   posi = 0 THEN

        FormularioD = ""
        FormularioD = FormularioD & wwwheader

        posi = 2

     ELSE

          Formulario  = Formulario & ofv.GenTabla("","","","80%","","0","","0","1","","",FormularioD,"si")
          FormularioD = ""

     END IF 

     IF (POSI > 34) then


        Formulario = Formulario & "<DIV >&nbsp;</DIV> "
        Formulario = Formulario & oPrtG.RepHeader2(cstr(""), cstr("600"))
        POSI=2

        FormularioD = FormularioD & wwwheader


     END IF






  end if





      Celda =""


    'id
    sInput = rs(0)
    Celda = Celda & ofv.gencelda("","vt","right","50","","","",sInput,"si")

    'Descripcion
    sInput = rs(1) 'ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"100","dt","")
    Celda = Celda & ofv.gencelda("","ut","","","","","",sInput,"si")


    Fila = ofv.genrow("","","","","","","",Celda,"si")


    FormularioD = FormularioD & fila 

    posi = posi + 1
    rs.Movenext
  Loop
 ''' sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")

    IF POSI > 2 THEN Formulario = Formulario & ofv.GenTabla("","","","80%","","0","","0","1","","",FormularioD,"si")

  sHTML = sHTML & Formulario


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>