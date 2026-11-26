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



   '    response.write "<table  marginwidth=0  cellspacing=0 cellpadding=0>"
   '    response.write "<tr>"
   '    response.write "<td width=100% >"
   '    response.write ofv.Repbody(request.querystring("mnn")) & ofv.RepHeader2(ofv.convertircar(request.querystring("mnn"),"_"," "),"600")
   '    response.write "<table id=margen1 width=600 marginwidth=0  cellspacing=1 cellpadding=2>"
   '    response.write "<tr>"
   '    response.write "<td width=10% >&nbsp;</td>"
   '    response.write "<td width=50% >&nbsp;</td>"
   '    response.write "<td width=40% >&nbsp;</td>"
   '    response.write "</tr></table>"

  'Encabezados

  wwwheader = chr(13) & " <TR ><TD  class=tt width=2% >CODBCRA</TD><TD  class=tt width=46% >CATEGORIAS (NIVEL 2)</TD><TD  class=tt width=46% >EVENTOS (NIVEL 3)</TD></tr>" & chr(13)


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

  WKSETTABLA = "ORIMPRIESGO"

 ' Strsql = "Select * From " & WKSETTABLA & " Order By agrupa"

  strsql="  Select a.*,b.descri as t1,c.descri as t2 ,B.codbcra AS NIVEL1, c.codbcra as nivel2 From  ORIMPRIESGO a "
  strsql = strsql & " inner  join  ORTIPOEVENTO b on (a.agrupa=b.objid)"
  strsql =strsql & "  inner  join  ORTIPOCATEGO c on (a.catego=c.objid)"
  strsql = strsql & " Order By b.codbcra,c.codbcra ,a.descri"
  'response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  posi = 2
  Do While (Not rs.EOF)


  if (eveant <> rs(3)) OR (POSI > 30) then

     Formulario  = Formulario & ofv.GenTabla("","","","80%","","0","","0","1","","",FormularioD,"si")
     FormularioD = ""

     IF (POSI > 30) then


        Formulario = Formulario & "<DIV >&nbsp;</DIV> "
        Formulario = Formulario & oPrtG.RepHeader2(cstr(""), cstr("600"))
        POSI=2

     END IF



    FormularioD = FormularioD & wwwheader

    Celda =""

    sInput = "SIN DEFINIR"
    Strsql = "Select DESCRI From ORTIPOEVENTO WHERE OBJID = 0" & RS(3)
    set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    IF NOT RS1.EOF THEN
       sInput= rs1(0)
       sInput=RS("NIVEL1")& "&nbsp;" &rs("t1")
    END IF

    Celda = Celda & ofv.gencelda("","vti colspan=3 STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#990000;' ","","","","","","NIVEL 1:&nbsp;" & sInput,"si")


    Fila = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","25","","",Celda,"si")
    FormularioD = FormularioD & fila

    posi = posi + 3


     eveant = rs(3)

  end if



      Celda =""
     'CODBCRA

      sInput = rs("nivel2")
      Celda = Celda & ofv.gencelda("","vt","","10%","","","",sInput,"si")

      sInput=rs("t2")
      Celda = Celda & ofv.gencelda("","ut","","40%","","","",sInput,"si")

      sInput=rs(2)
      Celda = Celda & ofv.gencelda("","ut","","50%","","","",sInput,"si")




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