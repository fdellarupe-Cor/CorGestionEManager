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

  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados
  Formulario = ofv.Encabezados("Id","Descripcion")


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  WKSETTABLA = "TIPOPRODUCTO"

  IF WKAPLIC = "OR" THEN WKSETTABLA = "ORTIPOPRODUCTO"



  sOPCION       = request.querystring("OPCION")
 ' response.write sOPCION & "22222222222222"
  IF SOPCION = "SI" THEN

        strsql = "Insert into " & WKSETTABLA & " "
        strsql = strsql & " (DESCRI)"
        strsql = strsql & "  Values "
        strsql = strsql &  "('NUEVO PRODUCTO')"


        call ofv.crearconsultaEx(StrSql,cn,1,volver)

  END IF





  Strsql = "Select * From " & WKSETTABLA & " Order By descri"
 ' response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=" & WKSETTABLA & "&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(0),sHref)

    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'id
    sInput = ofv.GenerarInput(rs(0).name,rs(0),"text",sForm,"100","dt","no")
    Celda = Celda & ofv.gencelda("","cc","right","50","","","",sInput,"si")


    'Descripcion
    sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"100","dt","no")
    Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

    Fila = ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=" & WKSETTABLA & "&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla(""," ","","80%","","0","","0","1","","",Formulario,"si")

  'Botonera
  Agregar = ""
  if bAgregar then

    sAccion = "AgregarTipoPRODUCTO.asp?Regisvuelta=" & CantRegMoverx & "&Opcion=si&tabla=" & WKSETTABLA
    'sAccion = "TipoPRODUCTO.asp?Opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)

  end if



  str1 = "Select descri From " & WKSETTABLA & " Order By descri "    ' orden pagina
  str2 = "Select descri From " & WKSETTABLA & " Order By descri "    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))


  'consulta



  sinput = "<a  href=TipoPRODUCTOCons.asp?"&request.querystring & " target=_new ><b>consulta</b></a>"
  Celda = ofv.gencelda("","tt style='border-style:outset;border-width:2;' ","center","","15%","","",sInput,"si")

  sAccion = "TipoPRODUCTOCons.asp?" & request.querystring
  Reportes=BtnReportes(16,"","xform",sAccion)
  'Celda = ofv.BotonConsultar(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,Reportes)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>