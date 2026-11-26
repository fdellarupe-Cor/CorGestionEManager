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
  Formulario = ofv.Encabezados("Descripcion")

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)


  sOPCION       = request.querystring("OPCION")
  IF SOPCION = "SI" THEN

        strsql = "Insert into ORTIPOCONTROLES"
        strsql = strsql & " (DESCRI)"
        strsql = strsql & "  Values "
        strsql = strsql &  "('')"
        xxx=SqlExLog("INSERT",strsql,CN,"ORTIPOCONTROLES","","","PAR")
 ' response.write strsql & "<br>"
   '     call ofv.crearconsultaEx(StrSql,cn,1,volver)

  END IF


  WKSETTABLA = "ORTIPOCONTROLES"

  Strsql = "Select * From " & WKSETTABLA  & "   Order By DESCRI"
  'RESPONSE.WRITE STRSQL
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posI = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir  then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=" & WKSETTABLA & "&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(0),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'Tipo
    sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"100","dt","")
    Celda = Celda & ofv.gencelda("","cc","","50 ALIGN=LEFT","","","",sInput,"si")







    'color
    scomK = "#00DD00" & chr(9) & "#F7EEDD" & chr(9) & "#F7DD77" & chr(9) & "#FF6666" & chr(9) & "#993333" & chr(9) & "#BBBBBB" & chr(9) & "#888888"
    scomK = scomK & chr(9) & "#555555" & chr(9)  & "#333333" & chr(9)  & "#000000"

    scomD = "VERDE CLARO" & chr(9) & "BEIGE" & chr(9) & "AMARILLO" & chr(9)
    scomD = scomD & "ROJO CLARO" & chr(9) & "ROJO OSCURO" & chr(9) & "GRIS CLARO" & chr(9)
    scomD = scomD & "GRIS" & chr(9) & "GRIS INTERMEDIO" & chr(9) & "GRIS OSCURO" & chr(9)
    scomD = scomD & "NEGRO"
   ' sInput = ofv.gencomboFX(sForm,rs(7).name,rs(7),scomK,scomD," ")
   ' Celda = Celda & ofv.gencelda("","UT","","20%","","","",sInput,"si")

   ' sInput = "&nbsp;"
   ' Celda = Celda & ofv.gencelda("","UT STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#DDDDDD;BACKGROUND-COLOR:" & RS(7) & ";' ","CENTER","5%","","","",sInput,"si")



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
  if bAgregar then
    sAccion = "ORTIPOCONTROL.asp?Opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select descripcion From " & WKSETTABLA & " Order By DESCRI "    ' orden pagina
  str2 = "Select descripcion From " & WKSETTABLA & " Order By DESCRI "    ' orden alfabetico
  Buscar = ""'ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  'consulta
  sAccion = "ORTIPOCONTROLCons.asp?" & request.querystring
  Celda   = ofv.BotonConsultar(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>