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
  '     response.write "<table id=margen1 width=600 marginwidth=0  cellspacing=1 cellpadding=2>"
   '    response.write "<tr>"
   '    response.write "<td width=10% >&nbsp;</td>"
   '    response.write "<td width=90% >&nbsp;</td>"
   '    response.write "</tr></table>"

  'Encabezados

 ''' Formulario=chr(13)&" <TR ><TD  class=tt width=2% >Tipo</TD><TD  class=tt width=2% >Descripcion</TD></tr>" & chr(13)



  wwwheader = chr(13) & " <TR ><TD  class=tt width=2% >CODIGO</TD><TD  class=tt width=2% >Descripcion</TD></tr>" & chr(13)


    set oPrtG = new ClsReportes ' server.createobject("CorGeneratorEXLR3.Reportes")
    Set oPrtG.Ofv = Ofv

   oPrtG.RCliLogo = session("clilogo")
   oPrtG.RCliNomb = ofv.convertircar(session("FormN"),"_"," ")
   Formulario = oPrtG.RepHeader2(cstr(""), cstr("600"))



  if request.querystring("seg") <> "" then
    Session("Form") = ofv.convertircar(request.querystring("seg"),"_"," ")
  end if


  set cn = ofv.conectar(ofv.strconn2)

  WKSETTABLA = "ORAMENAZAS"

  Strsql = "Select * From " & WKSETTABLA & " Order By DESCRI"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)


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

    'Tipo
    sInput = rs(1)
    Celda = Celda & ofv.gencelda("","vt","","10%","","","",sInput,"si")

    'Descripcion

    sInput = rs(2)
    Celda = Celda & ofv.gencelda("","ut","","90%","","","",sInput,"si")


    Fila = ofv.genrow("","","","","","","",Celda,"si") & chr(13)

    FormularioD = FormularioD & fila

    posi = posi + 1
    rs.Movenext
  Loop
 ''''' sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","1","","",Formulario,"si")

  IF POSI > 2 THEN Formulario = Formulario & ofv.GenTabla("","","","80%","","0","","0","1","","",FormularioD,"si")

  sHTML = sHTML & Formulario


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>