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
       response.write "</tr></table>"

  'Encabezados

  Formulario=chr(13)&" <TR ><TD  class=tt width=2% >USUARIO</TD><TD  class=tt width=2% >CODIGO</TD><TD  class=tt width=2% >GRUPO</TD></tr>" & chr(13)

  if request.querystring("seg") <> "" then
    Session("Form") = ofv.convertircar(request.querystring("seg"),"_"," ")
  end if


  set cn = ofv.conectar(ofv.strconn2)

  WKSETTABLA = "ORREFERENTES"

  Strsql = "Select * From " & WKSETTABLA & " Order By USUARIO"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)

      Celda ="" 

    'Tipo
    sInput = rs(2)
    Celda = Celda & ofv.gencelda("","vt","","40%","","","",sInput,"si")

    'Descripcion

    sInput=rs(1)
    Celda = Celda & ofv.gencelda("","ut","","15%","","","",sInput,"si")

    'GRUPO

      sGrupo = "-------"

      if rs(3) <> "" then
        Strsql = "Select descripcion From Grupos Where ID = '" & rs(3) & "'"
        set rsGrupo = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsGrupo.eof then sGrupo = rsGrupo(0)
        call ofv.cerrarconsulta(rsGrupo)
      end if

    sInput=sGrupo
    Celda = Celda & ofv.gencelda("","ut","","40%","","","",sInput,"si")



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