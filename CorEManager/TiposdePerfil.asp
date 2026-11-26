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
   wwwMNN = request.querystring("MNN") & ""
 '  response.write wwwMNN
 APLCHECK()

  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados

  IF session("APLMDCLI")="BC" THEN
     Formulario = ofv.Encabezados("Id","Descripcion","Codigo","Funciones Generales","ESTADO")
  ELSE
     Formulario = ofv.Encabezados("Id","Descripcion","Codigo","ESTADO")
  END IF


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  'call ofv.cerrarconn(cn)

  set cn = ofv.conectar(ofv.strconn2)


  Strsql = "Select * From TipoPerfil Order By descri"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if Not rs.EOF then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=Tipoperfil&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(1),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'Perfil
    sInput = rs(0)
    Celda = Celda & ofv.gencelda("","cc","","50","","","",sInput,"si")

    'Descripcion
    if rs("estado")=3 then
    sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"95","dt","")
    else
    sInput =rs(1)
    end if

    Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

     if rs("estado")=3 then
    sInput = ofv.GenerarInput(rs(3).name,rs(3),"textsb",sForm,"15","dt","")
    else
    sInput =rs(3)
    end if
    Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

    IF session("APLMDCLI")="BC" AND RS("ESTADO")=3 THEN

    'FUNCIONES
    'sImagenP = "&#0232;"
    'sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
    sinput = "<a href='TiposdePerfilFUN.asp?PERFIL=" & rs(0) & "&CantRegMover=" & CantRegMoverx
    sinput = sinput  & "' title='FUNCIONES DEL PUESTO' ><i class='las la-arrow-right'></i></a>"
    Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")
    ELSE
    	IF session("APLMDCLI")="BC" THEN
         Celda = Celda & ofv.gencelda("","ft","center","","15%","","","&nbsp;","si")
      END IF
    END IF


    SCOMSEP = chr(9)
		     scomK = "3" & SCOMSEP & "4"
			 scomD = "VIGENTE" & SCOMSEP & "NO&nbsp;VIGENTE"
	     sInput = ofv.gencomboFX(sForm,"ESTADO",RS("ESTADO"),scomK,scomD,"")

     Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")
    Fila = ofv.genrow("","","","","","","",Celda,"si")


    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=TipoPerfil&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext


  Loop
  'sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si") '& cn.ConnectionTimeout
  
    ftab = ofv.GenTabla("","","","100%","","0","","0","0","","",Formulario,"si")

    Celda = ofv.GenCelda("","","center","","","","",ftab,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa ","","","","0","","0","0","","",filas,"si")
  
  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "AgregarTipoPerfil.asp?Regisvuelta=" & CantRegMoverx & "&Opcion=si&tabla=TipoPerfil"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select descri From Tipoperfil Order By descri "    ' orden pagina
  str2 = "Select descri From Tipoperfil Order By descri "    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  'consulta
  sAccion = "TiposDePerfilCons.asp?" & request.querystring
''Celda   = ofv.BotonConsultar(sAccion)
  Reportes=BtnReportes(16,"","xform",sAccion)

 'consulta x sector
''xxxx    = CorGenBotonXp(16, "NO", "NO", "NO", "NO", "Perfiles por sector", "NO", "Sectores")
  xxxx    = CorGenBotonXp(16, "NO", "NO", "NO", "NO", "Perfiles por sector", "NO", "Reporte por Sectores")

  sAccion = "TiposDePerfilxsector.asp?" & request.querystring
  Celda2  = ofv.GenBoton(16, "esp1Bot", "BT", "S", "", "S", "A", "N", "esp1Frm", sAccion & "")

''sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda & celda2)
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","",Agregar,Buscar,Reportes,celda2)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = shtml & "</body></html>" '& SESSION("STRSQLX")
end if

Response.Write sHTML

%>