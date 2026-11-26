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






 wwwheader = chr(13) & " <TR ><TD  class=tt width=5% >ID</TD><TD  class=tt width=5% >CODIGO&nbsp;INTERNO</TD><TD  class=tt width=90% >Descripcion</TD></tr>" & chr(13)


    set oPrtG = new ClsReportes ' server.createobject("CorGeneratorEXLR3.Reportes")
    Set oPrtG.Ofv = Ofv

   oPrtG.RCliLogo = session("clilogo")
   oPrtG.RCliNomb = ofv.convertircar(session("FormN"),"_"," ")
   Formulario = oPrtG.RepHeader2(cstr(""), cstr("600"))




   set cn = ofv.conectar(ofv.strconn2)

   if session("APLMDLOR") <> "BE"  THEN
       Sordenar="	Order By DESCRIPCION "
     else
       Sordenar = " Order By CAST(CODINTER AS INT) "
    end if

  Strsql = "Select * From sucursal " & Sordenar
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  posI = 0
  Do While (Not rs.EOF)

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


    Celda = ""

   'ID
    sInput = rs(0) & "&nbsp;"
    Celda = Celda & ofv.gencelda("","vt align=right ","","","","","",sInput,"si")

   'Descripcion
    sInput = rs(1) & "&nbsp;"
    Celda = Celda & ofv.gencelda("","vt align=right ","","","","","",sInput,"si")

    sInput = rs(2)
    Celda = Celda & ofv.gencelda("","vt  ","","","","","",sInput,"si")
    Fila = ofv.genrow("","","","","","","",Celda,"si")
    FormularioD = FormularioD & fila
    posi = posi + 1
    rs.Movenext
  Loop
 ''' sHTML = sHTML & ofv.GenTabla("","","","70%  ","","0","","-1","1","","",Formulario,"si")

  IF POSI > 2 THEN Formulario = Formulario & ofv.GenTabla("","","","80%","","0","","0","1","","",FormularioD,"si")

  sHTML = sHTML & Formulario


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>