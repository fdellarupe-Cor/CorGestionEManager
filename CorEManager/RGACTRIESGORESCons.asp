<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
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



       response.write "<table  marginwidth=0  cellspacing=0 cellpadding=0>"
       response.write "<tr>"
       response.write "<td width=100% >"
    response.write ofv.Repbody(request.querystring("mnn")) & ofv.RepHeader2(ofv.convertircar(request.querystring("mnn"),"_"," "),"600")
       response.write "<table id=margen1 width=600 marginwidth=0  cellspacing=1 cellpadding=2>"
       response.write "<tr>"
       response.write "<td  >&nbsp;</td>"
       response.write "<td  >&nbsp;</td>"
       response.write "<td  >&nbsp;</td>"
       response.write "<td  >&nbsp;</td>"
       response.write "<td  >&nbsp;</td>"
       response.write "</tr></table>"

  'Encabezados

  'Encabezados

    Formulario=chr(13)&" <TR >"
    Formulario = Formulario & "<TD  class=tt width=2% >valor</TD>"
    Formulario = Formulario & "<TD  class=tt width=2% >Descripcion</TD>"
    Formulario = Formulario & "<TD  class=tt width=2% >DESDE</TD>"
    Formulario = Formulario & "<TD  class=tt width=2% >HASTA</TD>"
    Formulario = Formulario & "<TD  class=tt width=2% >REQ.PLAN</TD>"
    Formulario = Formulario & "</tr>" & chr(13)


  if request.querystring("seg") <> "" then
    Session("Form") = ofv.convertircar(request.querystring("seg"),"_"," ")
  end if


  set cn = ofv.conectar(ofv.strconn2)

  WKSETTABLA = "ORACTRIESGO"

  Strsql = "Select * From " & WKSETTABLA & " Order By VALORDESDE"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)

      Celda ="" 

    'Tipo
    sInput = rs(0)
    Celda = Celda & ofv.gencelda("","vt","","5%","","","",sInput,"si")

    'Descripcion

    sInput="SIN DEFINIR"

    Strsql = "Select DESCRI From ORRANGORIESGO WHERE OBJID =0" & RS(1)
    set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    IF NOT RS1.EOF THEN sInput=rs1(0)
    Celda = Celda & ofv.gencelda("","ut","","30%","","","",sInput,"si")


    'DESDE

    sInput=rs(2)
    Celda = Celda & ofv.gencelda("","ut","","5%","","","",sInput,"si")

    'HASTA

    sInput=rs(3)
    Celda = Celda & ofv.gencelda("","ut","","5%","","","",sInput,"si")

    'REQPLAN

    sInput=rs(4)
    Celda = Celda & ofv.gencelda("","ut","","5%","","","",sInput,"si")





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