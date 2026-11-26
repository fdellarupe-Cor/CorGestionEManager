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


sEMPRE = request.querystring("EMPRE")

    APLCHECK()
    set cn = ofv.conectar(ofv.strconn2)
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
xfila=""
   	sql = "select cliente from clienteplan where id=" & sEMPRE
   	set rz = ofv.crearconsultaEx(sql,cn,1,parametros)
       if not rz.eof then
   	Celda = ofv.GenCelda("","tt colspan=3","center","5%","","","",rz(0),"")
   	xfila = ofv.GenRow("","","","","","","",Celda,"si")
   	   end if
   	 call ofv.cerrarconsulta(rz)

  'Encabezados
  Celda = ofv.GenCelda("","tt","","5%","","","","Codigo","si")
  Celda = Celda & ofv.GenCelda("","tt","","45%","","","","Descripcion del Sector","si")
  if session("APLMDCLI")="BC" THEN
  Celda = Celda & ofv.GenCelda("","tt","","45%","","","","DEPENDENCIA&nbsp;CONTABLE","si")
  END IF
  xfila = xfila & ofv.GenRow("","","","","","","",Celda,"si")




   if session("APLMDLOR") <> "BE"  THEN
       Sordenar = " Order By DESCRI "
     else
       Sordenar = " Order By CODINTER "
    end if

if session("UN") <> 1 or Session("EMPMUN") = "N"  then
   Strsql = "Select * From ECO where ecoUN=" & sEMPRE & " order by descri"
else
   Strsql = "Select a.*,b.cliente From ECO a inner join clienteplan b on (a.ecoUN=b.ID) order by b.cliente"
end if
  'response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


  ifila=0
  XDIV=""
  Do While (Not rs.EOF)
    Celda = ""



    if ifila=40 then
	       sHTML=shtml &  XDIV & xshtml  & ofv.GenTabla("","AAA border" ,"","600 ","","0","","0","-1","","",xfila & Fila,"si")
	       XDIV= "<DIV >&nbsp;</DIV> "
	      ' sHTML=shtml & xsHTML&  "<DIV style='page-break-before :always'></DIV> "
	      ifila=0
	      fila=""
	      Celda = ""
    end if

    sInput = rs("eco")
    Celda = Celda & ofv.gencelda("","vt ALIGN=center","","","5%","","",sInput,"si")
   'Descripcion
    sInput =  REPLACE(rs(1) & "", " " ,"&nbsp;") &"&nbsp;"
    if session("UN") <> 1 or Session("EMPMUN") = "N" then
    	Celda = Celda & ofv.gencelda("","vt","","45%","","","",sInput,"si")
    else
    	Celda = Celda & ofv.gencelda("","vt","","45%","","","",rs(11) & "&nbsp;-&nbsp;" & sInput,"si")
    end if
    'Fila = ofv.genrow("","","","","","","",Celda,"si")
    if session("APLMDCLI")="BC" THEN
    sInput =  REPLACE(rs("ECOMAP")&"", " " ,"&nbsp;") &"&nbsp;"
	Celda = Celda & ofv.gencelda("","vt","","45%","","","",sInput,"si")
	END IF
    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")
   ' Formulario = Formulario & fila

    ifila=ifila +1
    rs.Movenext
  Loop

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
  if fila<>"" then
     sHTML=shtml & XDIV & xshtml &  ofv.GenTabla("","AAA BORDER","","600","","0","","0","-1","","",xfila & Fila,"si") '& "<DIV style='page-break-before:always'></DIV> "
     ifila=0
      fila=""
      Celda = ""
  end if
  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>