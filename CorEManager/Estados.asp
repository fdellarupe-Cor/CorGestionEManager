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

  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados
  Formulario = ofv.Encabezados("Estado","Descripcion")


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From estados Order By objid"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

'eliminar
   Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")

    'estado
    sInput = rs(0)
    Celda = Celda & ofv.gencelda("","cc","right","10%","","","",sInput,"si")

    'Descripcion
    sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"200","dt","")
    Celda = Celda & ofv.gencelda("","ut","","80%","","","",sInput,"si")

  '  Celda = Celda & ofv.gencelda("","","","","","","","","si")
    Fila = ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=Estados&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext
  Loop



  ftab = ofv.GenTabla("","","","50% ","","0","","0","0","","",Formulario,"si")

    Celda = ofv.GenCelda("","","center","","","","",ftab,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa  ","","","","0","","0","0","","",filas,"si")


 ' sHTML = sHTML & ofv.GenTabla("","","","50%","","0","","0","0","","",Formulario,"si")

 ' consulta
  sAccion = "EstadosCons.asp?" & request.querystring
  Celda = ofv.BotonConsultar(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no","no","no",celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>