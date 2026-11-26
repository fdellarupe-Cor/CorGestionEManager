<%@LCID = 11274%>
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0



  sopcion = request.querystring("opcion")

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)


'response.write  Session("EMPMUN")
  if Session("EMPMUN") <> "S" then

  IF session("APLMDCLI")="BC" THEN
       response.redirect "OrganigramaPuestos.asp?UN=1"
   ELSE
      response.redirect "OrganigramaUn.asp?UN=1"
   END IF

   END IF




  'end if





    sImagen4 = "<B><FONT FACE=CorFlaticon COLOR=#aa0000 SIZE=2>&#0232;</FONT></B>"





''  sHTML = ofv.FormHeader(session("FormN"))

  sHTML = ofv.FormHeader(session("FormN")) & chr(13)

  'Encabezados
  Formulario = ofv.Encabezados("Site/Sucursal","Consultar")

  if Session("EMPMUN")="S" then
    Strsql = "Select * From Clienteplan order by ID "
  else
    Strsql = "Select * From Clienteplan where id=1 "
  end if

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  posi = 0
  Do While (Not rs.EOF)

    sForm = "form" & posi

    'Eliminar

      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")


    'Descripcion
    sInput = rs(0) & "." & rs(1)
    Celda = Celda & ofv.gencelda("","ttt","","50%","","","",sInput,"si")


    sImagenP = "&#0232;"
    sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
    sinput = "<a class='ft1n' " & sFuncAnc & " href='OrganigramaUn.asp?UN=" & rs(0)

     IF session("APLMDCLI")="GP" OR session("APLMDCLI")="BP" OR session("APLMDCLI")="FINANSUR" OR session("APLMDCLI")="BPN" OR session("APLMDCLI")="BNX" OR session("APLMDCLI")="RIVA" THEN
	     sinput = "<a class='ft1n' " & sFuncAnc & " href='OrganigramaUn.asp?UN=" & rs(0)
	    ELSE
	    'response.redirect "OrganigramaPuestos.asp"
	     sinput = "<a class='ft1n' " & sFuncAnc & " href='OrganigramaPuestos.asp?UN=" & rs(0)
     END IF

    sinput = sinput  & "' title='Grafico de " & rs(1) & "' >" & sImagenP & "</a>"
    Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")


    Fila = ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""

    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext
  Loop


    sHTML = sHTML & ofv.GenTabla("","","","100% ","","0","","0","0","","",Formulario,"si")






  'Botonera
  Agregar = ""

  Buscar = ""



  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>