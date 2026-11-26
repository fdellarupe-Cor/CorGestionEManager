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
  'CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  'vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

		set oPrtG = new ClsReportes ' server.createobject("CorGeneratorEXLR3.Reportes")
		Set oPrtG.Ofv = Ofv

		oPrtG.RCliLogo = session("clilogo")
		oPrtG.RCliNomb = ofv.convertircar(session("FormN"),"_"," ")
		response.write  oPrtG.RepHeader2(cstr(""), cstr("600"))
		response.write corstyler() 

  'Encabezados

  Formulario=chr(13)&" <TR ><TD  class=tt width=2% >Codigo</TD><TD  class=tt width=2% >Descripcion</TD></tr>" & chr(13)

  if request.querystring("seg") <> "" then
    Session("Form") = ofv.convertircar(request.querystring("seg"),"_"," ")
  end if


  set cn = ofv.conectar(ofv.strconn2)

  WKSETTABLA = "ORCOSTOPOR"

  Strsql = "Select * From " & WKSETTABLA & " Order By DESCRI"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  'if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
   ' sForm = "form" & posi


      Celda ="" ' ofv.GenCelda("","","","","","","","&nbsp;","si")


    'Tipo
    sInput = rs(0)
    Celda = Celda & ofv.gencelda("","vti","","10%","","","",sInput,"si")

    'Descripcion
    'sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"90","dt","")
    sInput=rs(1)
    Celda = Celda & ofv.gencelda("","uti","","90%","","","",sInput,"si")



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