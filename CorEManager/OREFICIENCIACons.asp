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
		response.write corstyler() 

  'Encabezados

  Formulario=chr(13)&" <TR ><TD  class=tt width=2% >Valor</TD><TD  class=tt width=2% >Descripcion</TD></tr>" & chr(13)

  if request.querystring("seg") <> "" then
    Session("Form") = ofv.convertircar(request.querystring("seg"),"_"," ")
  end if


  set cn = ofv.conectar(ofv.strconn2)

  WKSETTABLA = "ORNIVEFICIENCIA"

  Strsql = "Select * From " & WKSETTABLA & " Order By VALOR"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)

      Celda ="" 

    'Tipo
    sInput = rs(1)
    Celda = Celda & ofv.gencelda("","vt","","10%","","","",sInput,"si")

    'Descripcion

    sInput=rs(4)
    Celda = Celda & ofv.gencelda("","ut","","90%","","","",sInput,"si")



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