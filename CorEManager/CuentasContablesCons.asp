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



    'Encabezados
  'sHTML = ofv.FormHeader(session("FormN"))
  'Formulario=chr(13)&" <TR ><TD  class=tt width=10% >Numero</TD><TD  class=tt width=10% >Numero</TD><TD  class=tt width=30% >Descripcion</TD></tr>" & chr(13)


    set oPrtG = new ClsReportes ' server.createobject("CorGeneratorEXLR3.Reportes")


    Set oPrtG.Ofv = Ofv


   sHTML = corstyler()

 '  response.write ofv.Repbody(request.querystring("mnn"))
   xsHTML=   "<table  marginwidth=0  cellspacing=0 cellpadding=0>"
   xsHTML=  xsHTML & "<tr>"
   xsHTML=  xsHTML & "<td width=100% >"
 '  xsHTML=  xsHTML &  ofv.RepHeader2(ofv.convertircar(request.querystring("mnn"),"_"," "),"600")
   xsHTML=  xsHTML & "<table id=margen1 width=600 marginwidth=0  cellspacing=1 cellpadding=2>"
   xsHTML=  xsHTML & "<tr>"
   xsHTML=  xsHTML & "<td width=1% >&nbsp;</td><td width=10% >&nbsp;</td>"
   xsHTML=  xsHTML & "<td width=40% >&nbsp;</td><td width=49% >&nbsp;</td>"
   xsHTML=  xsHTML & "</tr></table>"

   oPrtG.RCliLogo = session("clilogo")
   oPrtG.RCliNomb = ofv.convertircar(session("FormN"),"_"," ")
   xsHTML=  oPrtG.RepHeader2(cstr(""), cstr("600"))

   xfila=chr(13)&" <TR ><TD  class=tt width=10% >ID</TD><TD  class=tt width=20% >Numero</TD><TD  class=tt width=70% >Descripcion</TD></tr>" & chr(13)

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From ObjCuenta Order By numero"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  ifila=0
  XDIV=""

  Do While (Not rs.EOF)
    Celda = ""


    if ifila=40 then
		       sHTML=shtml &  XDIV & xshtml  & ofv.GenTabla("","AAA " ,"","600 ","","0","","0","-1","","",xfila & Fila,"si")
		       XDIV= "<DIV >&nbsp;</DIV> "
		      ' sHTML=shtml & xsHTML&  "<DIV style='page-break-before :always'></DIV> "
		      ifila=0
		      fila=""
		      Celda = ""
	 end if

    'ID
    sInput = rs(0)
    Celda = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")

    'Numero
    sInput = rs(2)
    Celda = Celda & ofv.GenCelda("","vt align=right","","","","","",sInput,"si")
    'Descripcion
    sInput ="&nbsp;&nbsp;&nbsp;" & rs(1)
    Celda = Celda & ofv.GenCelda("","ut","","","","","",sInput,"si")
    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

    Formulario = Formulario & fila

    ifila = ifila + 1
    rs.Movenext
  Loop

  'sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")
   if fila<>"" then
       sHTML=shtml & XDIV & xshtml &  ofv.GenTabla("","AAA  " ,"","600 ","","0","","0","-1","","",xfila & Fila,"si") '& "<DIV style='page-break-before:always'></DIV> "
       ifila=0
        fila=""
        Celda = ""
  end if


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>