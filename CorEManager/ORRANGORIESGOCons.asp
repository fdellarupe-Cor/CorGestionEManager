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

  'Encabezados

    Formulario=chr(13)&" <TR >"
    Formulario = Formulario & "<TD  class=tt width=2% >valor</TD>"
    Formulario = Formulario & "<TD  class=tt width=2% >Descripcion</TD>"
    Formulario = Formulario & "<TD  class=tt width=2% >color</TD>"
    Formulario = Formulario & "</tr>" & chr(13)


  if request.querystring("seg") <> "" then
    Session("Form") = ofv.convertircar(request.querystring("seg"),"_"," ")
  end if


  set cn = ofv.conectar(ofv.strconn2)

  WKSETTABLA = "ORRANGORIESGO"

  Strsql = "Select * From " & WKSETTABLA & " Order By VALOR DESC"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)

      Celda ="" 

    'Tipo
    sInput = rs(1)
    Celda = Celda & ofv.gencelda("","vt","","5%","","","",sInput,"si")

    'Descripcion

    sInput=rs(2)
    Celda = Celda & ofv.gencelda("","ut","","30%","","","",sInput,"si")


    'color

    sInput="&nbsp;"
    Celda = Celda & ofv.gencelda("","ut bgcolor='" & rs(3) & "' ","","5%","","","",sInput,"si")



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