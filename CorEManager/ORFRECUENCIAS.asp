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
  Formulario = ofv.Encabezados("Valor","Descripcion","Activo")

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  'bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  sOPCION       = request.querystring("OPCION")


  WKSETTABLA = "ORFRECUENCIA"

  Strsql = "Select * From " & WKSETTABLA & " Order By VALOR ASC"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posI = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

	XCOL = " BGCOLOR='#F7EEDD' "
    IF (RS(3) & "") = "N" THEN XCOL = " BGCOLOR='#DDDDDD' "

    'Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=" & WKSETTABLA & "&volver=" & vuelta & "&Eliminar=NSNC&Nombre=" & rs(3)
      Celda = ofv.BotonEliminar(rs(0),sHref)
    else
      Celda = ofv.GenCelda("","UT" & XCOL,"","","","","","&nbsp;","si")
    end if

    'VALOR FINAL
	sInput = rs(2)
	Celda = Celda & ofv.gencelda("","UT ALIGN=LEFT" & XCOL,"","","","","",sInput,"si")

    'Descripcion
	sInput = rs(1)
	Celda = Celda & ofv.gencelda("","UT ALIGN=LEFT" & XCOL,"","","","","",sInput,"si")



    scomK = "S" & chr(9) & "N"
    scomD = "SI" & chr(9) & "NO"
    sInput = ofv.gencomboFX(sForm,rs(3).name,rs(3),scomK,scomD," ")
    Celda = Celda & ofv.gencelda("","UT" & XCOL,"","20%","","","",sInput,"si")




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
  sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si")

  'Botonera
  Agregar = ""


''  consulta
  sAccion = "ORFRECUENCIASCons.asp?" & request.querystring
  Celda = ofv.BotonConsultar(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no","",Celda,volver)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>