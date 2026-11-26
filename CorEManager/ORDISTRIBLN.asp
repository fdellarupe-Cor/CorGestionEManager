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
  Formulario = ofv.Encabezados("LINEA","PORCENTAJE")

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  sOPCION       = request.querystring("OPCION")
  IF SOPCION = "SI" THEN

        strsql = "Insert into ORDISTRIBLN "
        strsql = strsql & " (LINNEGO,PORCENTAJE)"
        strsql = strsql & "  Values "
        strsql = strsql &  "('0','0')"

  '''response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

  END IF

  WKSETTABLA = "ORDISTRIBLN"

  Strsql = "Select * From " & WKSETTABLA & " Order By LINNEGO"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posI = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=" & WKSETTABLA & "&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar("LINEA A DISTRIBUIR",sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'LINEA DE NEGOCIOS
    strsql = "Select * From ORLINNEGO1 Order By Descripcion"
    sInput = ofv.GenerarCombo(SFORM,rs(1).name,rs(1),strsql,cn,0,1,"")
    Celda = Celda & ofv.gencelda("","UT","","80%","","","",sInput,"si")


    'PORCENTAJE
    sInput = ofv.GenerarInput(rs(2).name,REPLACE(CSTR(rs(2) & ""),",","."),"text",sForm,"10","dtN","")
    Celda = Celda & ofv.gencelda("","UT","RIGHT","20%","","","",sInput & "%","si")



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
  sHTML = sHTML & ofv.GenTabla("","","","60% ","","0","","0","0","","",Formulario,"si")

  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "ORDISTRIBLN.asp?Opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if


  Buscar = ""

  'consulta
  sAccion = "ORDISTRIBLNCons.asp?" & request.querystring
  Celda = ofv.BotonConsultar(sAccion)
  
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>