<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrarx = 8
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
  SOPCION = request.querystring("OPCION")

  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados
  Formulario = ofv.Encabezados("Id","Descripcion","tipo","detalle")


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)


''''''''''''''''''''''''''''''''''''''''''''''''
'' INCORPORA competencias - OPCION = SI '''
''''''''''''''''''''''''''''''''''''''''''''''''

 IF SOPCION = "SI" THEN





        strsql = "Insert into tipocompetencia "
        strsql = strsql & " (descripcion,tipocompe)"
        strsql = strsql & " Values ("
        strsql = strsql & "'','0'"
  ''''''strsql = strsql & "'NUEVA COMPETENCIA','0'"
        strsql = strsql & ")"

        call ofv.crearconsultaEx(StrSql,cn,1,volver)




 END IF




  Strsql = "Select * From Tipocompetencia Order By descripcion"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=Tipocompetencia&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(1),sHref)

    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'id
    sInput = rs(0)
    Celda = Celda & ofv.gencelda("","cc","right","50","","","",sInput,"si")


    'Descripcion
    sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"100","dt","")
    Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

    'tipo
    scomK = "0" & chr(9) & "1"
    scomD = "Libre" & chr(9) & "Parametrizado"


    if bModificar then
        sInput = ofv.gencomboFX(sForm,rs(2).name,rs(2),scomK,scomD," ")
    else
        sInput = ofv.GenerarInput("",rs(2),"readonly","","10","dt","")
    end if
    Celda = Celda & ofv.gencelda("","ut style='border-radius:20px;border-style:solid;border-color:#888888;border-width:1;' ","","","","","",sInput,"si")


    'FUNCIONES

   if rs(2) = 1 then

    'sImagenP = "&#0232;"
    'sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
    sinput = "<a href='tipocompefundet.asp?compe=" & rs(0)
    sinput = sinput  & "' title='Parametrizacion de competencias' ><i class='las la-arrow-right'></i></a>"
    Celda = Celda & ofv.gencelda("","ft","center","","5%","","",sInput,"si")

   else

    Celda = Celda & ofv.gencelda("","ut","center","","","","","----","si")

   end if


    Fila = ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=Tipocompetencia&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext
  Loop

  ftab = ofv.GenTabla("","","","100% ","","0","","0","0","","",Formulario,"si")

    Celda = ofv.GenCelda("","","center","","","","",ftab,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa ","","","","0","","0","0","","",filas,"si")

  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "AgregarTIPOCOMPEFUN.asp?Regisvuelta=" & CantRegMoverx & "&Opcion=si&tabla=Tipocompetencia"
    Agregar = ofv.BotonAgregar(sAccion)

  end if

  str1 = "Select descripcion From Tipocompetencia Order By descripcion "    ' orden pagina
  str2 = "Select descripcion From Tipocompetencia Order By descripcion "    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))


  'consulta

  'sAccion = "TipodscptoCons.asp?" & request.querystring
  sAccion = "TIPOCOMPEFUNCONS.asp?" & request.querystring

''Celda = ofv.BotonConsultar(sAccion)
  Reportes=BtnReportes(16,"","xform",sAccion)

''sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","",Agregar,Buscar,"no",Reportes)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>