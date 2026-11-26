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



  'if Session("EMPMUN") <> "S" then

     response.redirect "SectoresUn.asp?UN=" & session("UN")

  'end if





    sImagen4 = "<B><FONT FACE=CorFlaticon COLOR=#aa0000 SIZE=2>&#0232;</FONT></B>"





  sHTML = ofv.FormHeader(session("FormN"))



  'Encabezados
  Formulario = ofv.Encabezados("Id","Sites/Sucursales","Sectores")


  Strsql = "Select * From Clienteplan order by ID"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  posi = 0
  Do While (Not rs.EOF)

    sForm = "form" & posi

    'Eliminar

      if bSuprimir then
	  	 		sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
	  	 		sHref = sHref & "&Tabla=Clienteplan&volver=" & vuelta & "&Eliminar=NSNC"
	  	 	    Celda = ofv.BotonEliminar(rs(0),sHref)
	  	 	else
	  	 		Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'Descripcion
    sInput = rs(0)
    Celda = Celda & ofv.gencelda("","cc","","5","","","",sInput,"si")

    'Descripcion
    sInput = rs(1)
    Celda = Celda & ofv.gencelda("","ttt","","100%","","","",sInput,"si")



    sImagenP = "&#0232;"
    sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
    sinput = "<a class='ft1n' " & sFuncAnc & " href='SectoresUn.asp?UN=" & rs(0)
    sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx
    sinput = sinput  & "' title='Areas/sectores de " & rs(1) & "' >" & sImagenP & "</a>"
    Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")



    Fila = ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""

    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext
  Loop


  ftab = ofv.GenTabla("","","","100% ","","0","","0","0","","",Formulario,"si")

    Celda = ofv.GenCelda("","","center","","","","",ftab,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa ","","90%","","0","","0","0","","",filas,"si")




  'Botonera
  Agregar = ""

  Buscar = ""


  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>