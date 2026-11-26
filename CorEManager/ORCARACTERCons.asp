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
  
  set oPrtG = new ClsReportes ' server.createobject("CorGeneratorEXLR3.Reportes")
    Set oPrtG.Ofv = Ofv

   oPrtG.RCliLogo = session("clilogo")
   oPrtG.RCliNomb = ofv.convertircar(session("FormN"),"_"," ")
    response.write  oPrtG.RepHeader2(cstr(""), cstr("600"))
	
	
  'CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  'vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)


    '   response.write "<table  marginwidth=0  cellspacing=0 cellpadding=0>"
    '   response.write "<tr>"
    '   response.write "<td width=100% >"
    '   response.write ofv.Repbody(request.querystring("mnn")) & ofv.RepHeader2(ofv.convertircar(request.querystring("mnn"),"_"," "),"600")
    '   response.write "<table id=margen1 width=600 marginwidth=0  cellspacing=1 cellpadding=2>"
    '   response.write "<tr>"
    '   response.write "<td width=1% >&nbsp;</td><td width=10% >&nbsp;</td>"
    '   response.write "<td width=40% >&nbsp;</td><td width=49% >&nbsp;</td>"
    '   response.write "</tr></table>"

   response.write corstyler()

  'Encabezados

  Formulario = chr(13) & "<TR >"
  Formulario = Formulario & "<TD  class=tt width=2%  >CODBCRA</TD>"
  Formulario = Formulario & "<TD  class=tt width=2%  >CODIGO</TD>"
  Formulario = Formulario & "<TD  class=tt width=65% >Descripcion</TD>"
  Formulario = Formulario & "<TD  class=tt width=5%  >Acumula</TD>"
  Formulario = Formulario & "<TD  class=tt width=5%  >I.Contable</TD>"
  Formulario = Formulario & "<TD  class=tt width=5%  >Patron</TD>"
  Formulario = Formulario & "<TD  class=tt width=10%  >ID</TD>"
  Formulario = Formulario & "</tr>" & chr(13)

    




  if request.querystring("seg") <> "" then
    Session("Form") = ofv.convertircar(request.querystring("seg"),"_"," ")
  end if


  set cn = ofv.conectar(ofv.strconn2)

  WKSETTABLA = "ORTIPOCARACTER"

  Sordenar="	Order By CODBCRA "
  Strsql = "Select * From " & WKSETTABLA & Sordenar
  'response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  'if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
   ' sForm = "form" & posi


      Celda ="" ' ofv.GenCelda("","","","","","","","&nbsp;","si")


    sInput = rs("CODBCRA")
    Celda = Celda & ofv.gencelda("","vt ALIGN=RIGHT","","","","","",sInput,"si")
    'Tipo

    sInput = rs(0)
    Celda = Celda & ofv.gencelda("","vt ALIGN=RIGHT","","","","","",sInput,"si")

    'Descripcion
    sInput = rs(1)
    Celda = Celda & ofv.gencelda("","ut","","","","","",sInput,"si")

    'ACUMULA

   select case (RS(2) & "")

         case "OM"
              sInput = "OTROS MONTOS"

         case "REC"
              sInput = "RECUPEROS"

         case "DES"
              sInput = "DESAFECTACIONES"

         case "PER"
              sInput = "PERDIDAS"

         case else
              sInput = "NO CONTABLES"


    end select

    Celda = Celda & ofv.gencelda("","ut","","","","","",sInput,"si")

    'IMPCONT

   select case (RS(4) & "")

         case "S"
              sInput = "SI"

         case else
              sInput = "NO"


    end select
    Celda = Celda & ofv.gencelda("","ut","","","","","",sInput,"si")

    'PATRON

   select case (RS(5) & "")

         case "DEF"
              sInput = "DEFAULT"
               STRSQL="SELECT OBJID FROM ORTRACER WHERE CODTRACER='DOBLIGA'"
               set rsG = ofv.crearconsultaEx(StrSql,cn,1,parametros)
               XID=RSG(0)
               CALL OFV.CERRARCONSULTA(RSG)


         case else
              sInput = "ESPECIAL"
              STRSQL="SELECT OBJID FROM ORTRACER WHERE CODTRACER='DOBLIGA"&RS(0)&"'"
			  set rsG = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			  XID=RSG(0)
              CALL OFV.CERRARCONSULTA(RSG)


    end select

    Celda = Celda & ofv.gencelda("","ut","","","","","",sInput,"si")

     Celda = Celda & ofv.gencelda("","ut ALIGN=RIGHT","","","","","",xID,"si")





    Fila = ofv.genrow("","","","","","","",Celda,"si") & chr(13)




    Formulario = Formulario & fila

    posi = posi + 1
    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","1","","",Formulario,"si")




  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>