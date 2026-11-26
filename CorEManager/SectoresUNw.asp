<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
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

  sUN = request.querystring("UN")

  IF sUN <> "" THEN

     SESSION("SUN") = SUN

  ELSE

    SUN = SESSION("SUN")

  END IF

  sHTML = ofv.FormHeader(session("FormN"))

  if SESSION("APLCHK") <> "S" then xxxx = APLCHECK()



  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From CLIENTEPLAN  WHERE ID = " & SUN
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     SNOMEMPRE = RS(1)

  END IF

    sImagen4 = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=2>4</FONT></B>"
    sImagenP = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=2>&#0232;</FONT></B>"

  Celda = ofv.gencelda("","ttt","","","","","","<BIG>" & sImagen4 & " " & SNOMEMPRE & "</BIG>","si")
  Fila = ofv.genrow("","","","","","","",Celda,"si")
  sHTML = sHTML & ofv.GenTabla("","","","100% ","","0","","0","0","","",Fila,"si") & "<BR>"

  'Encabezados

  IF  session("APLMDLOR") <> "BE"  THEN
   Formulario = ofv.Encabezados("Id","Descripcion del Sector","Propiedades","ESTADO")
  ELSE
   Formulario = ofv.Encabezados("Id","CODIGO","Descripcion del Sector","Propiedades")
  END IF

  if session("APLMDLOR") <> "BE"  THEN
       Sordenar="	Order By DESCRI "
      else
       Sordenar = " Order By CODINTER  "
  end if

  Strsql = "Select * From ECO where ecoun = " & SUN & Sordenar

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   if Not rs.EOF then   rs.move(CantRegMoverx)
  posI = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=Eco&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(1),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'id
    sInput = rs(0)
    Celda = Celda & ofv.GenCelda("","cc","center","5","","","",sInput,"si")

    if session("APLMDLOR") = "BE"  THEN
	  SINPUT="<input name='"& rs(2).name   &"'  value='" & trim(rs(2)&"") & "' type=text MAXLENGTH=2 class=dtn onblur="&chr(34)&"validarnum(this,'"&sForm&"','"&trim(rs(2)&"")&"',2);"&chr(34) & ">"
	  Celda = Celda & ofv.GenCelda("","cc","center","5","","","",sInput,"si")
	end if




    'Descripcion
    if rs("estado") = 3 then

    sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"100","dtbs","")
    else
    sinput=rs(1)
    end if
    Celda = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")

    'Propiedades


  if (session("APLMDLOR") & "") = "BE" then

    sinput = "&nbsp;---"

  else

    sImagenP = "&#0232;"
    sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
    sinput = "<a class='ft1n' " & sFuncAnc & " href='SectoresProp.asp?wperfil=" & rs(0) & "&CantRegMover=" & CantRegMoverx
    sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx
    sinput = sinput  & "' title='Propiedades de " & rs(1) & "' >" & sImagenP & "</a>"
      if rs("estado") = 3 then
      Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")
      else
      Celda = Celda & ofv.gencelda("","ft","center","","15%","","","&nbsp;","si")
      end if
         SCOMSEP = chr(9)
	     scomK = "3" & SCOMSEP & "4"
		 scomD = "VIGENTE" & SCOMSEP & "NO&nbsp;VIGENTE"
	     sInput = ofv.gencomboFX(sForm,"ESTADO",RS("ESTADO"),scomK,scomD,"")

     'Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")

  end if

     Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")

     Fila = ofv.genrow("","","","","","","",Celda,"si")



    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=Eco&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")


    posi = posi + 1
    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","-1","1","","",Formulario,"si")


  retornar = ""
  if Session("EMPMUN") = "S" then
     sInput = ofv.generarinput("","Volver","submit","Volver","10","bt","")
     sInput = ofv.gencelda("","","center","","","","",sInput,"si")
     sAccion = "Sectores.asp"
     retornar = ofv.genform("Volver","",sAccion,"","post","",sInput,"si")

     sAccion = "Sectores.asp"
     retornar = ofv.BotonVolver(sAccion)



  end if


  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "AgregarSector.asp?Regisvuelta=" & CantRegMoverx & "&Opcion=si&tabla=Eco"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select descri From Eco where ecoun = " & SUN & Sordenar    ' orden pagina
  str2 = "Select descri From Eco where ecoun = " & SUN & Sordenar    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

     Listar  = ""

  if (session("APLMDLOR") & "") = "BE" then

  else
     'Listar Organigrama

     xxxx    = CorGenBotonXp(16, "NO", "NO", "NO", "NO", "Listar el organigrama", "NO", "Organigrama")
     sAccion = "LisOrganigrama.asp"
     Listar  = ofv.GenBoton(16, "esp1Bot", "BT", "S", "", "S", "A", "N", "esp1Frm", sAccion & "")

  end if

  'consulta

  sAccion = "SectoresCons.asp?" & request.querystring
  Celda = ofv.BotonConsultar(sAccion)


  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","",Agregar,Buscar,Listar,celda & retornar)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
  shtml=SHTML & 	ValidarInput()
  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>