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



  sImagenCampo = "ARCHIVO"
  CantRegMostrarx = 100

   response.write corstyler()


  'Encabezados

  wwwheader = chr(13) & " <TR ><TD  class=tt width=2% >COD.</TD><TD  class=tt width=46% >DESCRIPCION</TD><TD  class=tt width=30% >FORMULA</TD><TD  class=tt width=2% >FECHA</TD></tr>" & chr(13)


    set oPrtG = new ClsReportes ' server.createobject("CorGeneratorEXLR3.Reportes")
    Set oPrtG.Ofv = Ofv

   oPrtG.RCliLogo = session("clilogo")
   oPrtG.RCliNomb = ofv.convertircar(session("FormN"),"_"," ")
   Formulario = oPrtG.RepHeader2(cstr(""), cstr("600"))



 ''' Formulario= wwwheader


  if request.querystring("seg") <> "" then
    Session("Form") = ofv.convertircar(request.querystring("seg"),"_"," ")
  end if
          xTrColor = "#FFFFFF"

 eveant = 0

  set cn = ofv.conectar(ofv.strconn2)



  Strsql = "select * FROM orformulas"
 ' response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  posi = 0
  Do While (Not rs.EOF)


  if (POSI=0) OR (POSI > 30) then

     IF   posi = 0 THEN

        FormularioD = ""
        FormularioD = FormularioD & wwwheader

        posi = 2

     ELSE

          Formulario  = Formulario & ofv.GenTabla("","","","80%","","0","","0","1","","",FormularioD,"si")
          FormularioD = ""

     END IF 

     IF (POSI > 30) then


        Formulario = Formulario & "<DIV >&nbsp;</DIV> "
        Formulario = Formulario & oPrtG.RepHeader2(cstr(""), cstr("600"))
        POSI=2

        FormularioD = FormularioD & wwwheader

     END IF






  end if



      Celda =""

      sInput = rs(0) & "&nbsp;"
      Celda  = Celda & ofv.gencelda("","vt","RIGHT","","","","",sInput,"si")

      sInput = rs(2)
      Celda  = Celda & ofv.gencelda("","utI","","","","","",sInput,"si")

      Strsql = "Select descri From ORTIPOFORMULA where tipo = 0" & rs(1)
      set rsxx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if rsxx.eof then
         sInput = "S/D"
      else
         sInput = rsxx(0)
      end if 

      Celda = Celda & ofv.gencelda("","utI","","","","","",sInput,"si")

    sInput = GTFechadorSH(RS("FECHA"))
    Celda = Celda & ofv.gencelda("","utI","","","","","",sInput,"si")


    Fila = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","25","","",Celda,"si")

       if xTrColor = "#FFFFFF" then
          xTrColor = "#F0F0F0"
       ELSE
          xTrColor = "#FFFFFF"
       END IF



    FormularioD = FormularioD & fila

    posi = posi + 1
    rs.Movenext
  Loop
''  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","1","","",Formulario,"si")

    IF POSI > 2 THEN Formulario = Formulario & ofv.GenTabla("","","","80%","","0","","0","1","","",FormularioD,"si")

  sHTML = sHTML & Formulario

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>