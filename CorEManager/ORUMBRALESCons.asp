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

  wwwheader = chr(13) & " <TR ><TD  class=tt width=2% >VALOR</TD><TD  class=tt width=46% >DESCRIPCION</TD><TD  class=tt width=30% >COLOR</TD><TD  class=tt width=2% >&nbsp;</TD></tr>" & chr(13)


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

  WKSETTABLA = "ORUMBRALES"

  Strsql = "Select * From " & WKSETTABLA & " WHERE IDINDICADOR=0  Order By ORDEN"
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

      sInput = rs(1) & "&nbsp;"
      Celda = Celda & ofv.gencelda("","vt","RIGHT","","","","",sInput,"si")

      sInput=rs(3)
      Celda = Celda & ofv.gencelda("","utI","","","","","",sInput,"si")


   select case (RS(7) & "")

         case "#00DD00"
              sInput = "VERDE CLARO"

         case "#F7EEDD"
              sInput = "BEIGE"


         case "#F7DD77"
              sInput = "AMARILLO"


         case "#FF6666"
              sInput = "ROJO CLARO"


         case "#993333"
              sInput = "ROJO OSCURO"


         case "#BBBBBB"
              sInput = "GRIS CLARO"


         case "#555555"
              sInput = "GRIS INTERMEDIO"


         case "#333333"
              sInput = "GRIS OSCURO"


         case "#000000"
              sInput = "NEGRO"


         case "#FF8C00"
              sInput = "NARANJA"


         case "#888888"
              sInput = "GRIS"

         case ELSE
              sInput = "S/D"


    end select


      Celda = Celda & ofv.gencelda("","utI","","","","","",sInput,"si")

    sInput = "&nbsp;"
    Celda = Celda & ofv.gencelda("","UT STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#DDDDDD;BACKGROUND-COLOR:" & RS(7) & ";' ","CENTER","","","","",sInput,"si")


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