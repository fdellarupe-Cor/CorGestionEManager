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
   '     response.write ofv.Repbody(ofv.convertircar(session("FormN"),"_"," ")) & ofv.RepHeader2(ofv.convertircar(session("FormN"),"_"," "),"600")
   '    response.write "<table id=margen1 width=600 marginwidth=0  cellspacing=1 cellpadding=2>"
   '    response.write "<tr>"
   '    response.write "<td width=10% >&nbsp;</td>"
   '    response.write "<td width=90% >&nbsp;</td>"
   '    response.write "</tr></table>"

  'Encabezados

 ''''' Formulario=""'chr(13)&" <TR ><TD  class=tt width=2% >CODBCRA</TD><TD  class=tt width=96% >Descripcion</TD></tr>" & chr(13)


  wwwheader = chr(13) & " <TR ><TD  class=tt width=2% >CODBCRA</TD><TD  class=tt width=96% >Descripcion</TD></tr>" & chr(13)


    set oPrtG = new ClsReportes ' server.createobject("CorGeneratorEXLR3.Reportes")
    Set oPrtG.Ofv = Ofv

   oPrtG.RCliLogo = session("clilogo")
   oPrtG.RCliNomb = ofv.convertircar(session("FormN"),"_"," ")
   Formulario = oPrtG.RepHeader2(cstr(""), cstr("600"))



  if request.querystring("seg") <> "" then
    Session("Form") = ofv.convertircar(request.querystring("seg"),"_"," ")
  end if
          xTrColor = "#FFFFFF"

 eveant = 0
 eveant2 = 0

  set cn = ofv.conectar(ofv.strconn2)

  WKSETTABLA = "ORLINNEGOCIOS"

 ' Strsql = "Select * From " & WKSETTABLA & "  Order By codbcra"
 '
 'response.write strsql
 strsql ="Select a.*,b.descripcion as t1 ,B.codbcra AS NIVEL1 , c.codbcra AS NIVEL2 ,c.descripcion as t2 From ORLINNEGOCIOS a "
 strsql=strsql & " inner join ORLINNEGO1 b on (a.nivel1=b.OBJID) "
 strsql=strsql & " inner join ORLINNEGO2 c on (a.nivel2=c.OBJID) "
 strsql=strsql & " Order by b.codbcra,c.codbcra ,A.DESCRI"
' Strsql=strsql & " Order by A.codbcra,B.CODBCRA "
' response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  posi = 0

  Do While (Not rs.EOF) and (posi < CantRegMostrarx)


  if (eveant <> rs(1)) OR (POSI > 34) then

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





    Celda =""

    sInput = "SIN DEFINIR"
    Strsql = "Select DESCRIPCION From ORLINNEGO1 WHERE OBJID = 0" & RS(1)
    set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    IF NOT RS1.EOF THEN
       sInput=rs1(0)

    END IF
    sinput=RS("NIVEL1")& "&nbsp;" & rs("t1")
    Celda = Celda & ofv.gencelda("","vti colspan=3 STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#990000;' ","","","","","","NIVEL 1 :&nbsp;&nbsp;" & sInput,"si")


    FilaH = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","25","","",Celda,"si")



    FormularioD = FormularioD & filaH

    posi = posi + 2


    if (eveant = rs(1)) AND (eveant2 = rs(2)) then

        Celda =""

        sinput="<U>" & RS("NIVEL2")& "&nbsp;" & rs("t2") & "</U>"
        Celda = Celda & ofv.gencelda("","vti colspan=3  ","","","","","","&nbsp;&nbsp;&nbsp;&nbsp;<U>NIVEL 2 :</U>&nbsp;&nbsp;" & sInput,"si")

        Fila = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","25","","",Celda,"si")
        FormularioD = FormularioD & fila

        posi = posi + 2


    END IF



     eveant = rs(1)
  end if

  if (eveant2 <> rs(2))   then

     Formulario  = Formulario & ofv.GenTabla("","","","80%","","0","","0","1","","",FormularioD,"si")
     FormularioD = ""


    Celda =""



    sinput="<U>" & RS("NIVEL2")& "&nbsp;" & rs("t2") & "</U>"

    Celda = Celda & ofv.gencelda("","vti colspan=3  ","","","","","","&nbsp;&nbsp;&nbsp;&nbsp;<U>NIVEL 2 :</U>&nbsp;&nbsp;" & sInput,"si")

    Fila = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","25","","",Celda,"si")
    FormularioD = FormularioD & fila

    posi = posi + 2


     eveant2 = rs(2)
  end if



      Celda =""
    'CODBCRA
    sInput = rs("CODBCRA")
    Celda = Celda & ofv.gencelda("","vt ALIGN=RIGHT","","10%","","","",sInput,"si")


    'Descripcion
    sInput=rs(3)
    Celda = Celda & ofv.gencelda("","ut","","90%","","","",sInput,"si")


    Fila = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","25","","",Celda,"si")

       if xTrColor = "#FFFFFF" then
          xTrColor = "#FFFFFF" '"#F0F0F0"
       ELSE
          xTrColor = "#FFFFFF"
       END IF



    FormularioD = FormularioD & fila

    posi = posi + 1
    rs.Movenext
  Loop
  '''''''''''''sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","1","","",Formulario,"si")

    IF POSI > 2 THEN Formulario = Formulario & ofv.GenTabla("","","","80%","","0","","0","1","","",FormularioD,"si")

  sHTML = sHTML & Formulario


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>