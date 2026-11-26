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
  Formulario = ofv.Encabezados("DESCRIPCION")

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  sOPCION       = request.querystring("OPCION")
  IF SOPCION = "SI" THEN

        strsql = "Insert into RGACTVULNERABILIDAD "
        strsql = strsql & " (descripcion)"
        strsql = strsql & "  Values "
        strsql = strsql &  "('DESCRIPCION')"

        xxx=SqlExLog("INSERT",strsql,CN,"RGACTCAPAS","","","PAR")

  '''response.write strsql & "<br>"
     ''   call ofv.crearconsultaEx(StrSql,cn,1,volver)

  END IF


  WKSETTABLA = "RGACTVULNERABILIDAD"

  Strsql = "Select * From " & WKSETTABLA & " Order BY DESCRIPCION"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posI = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=" & WKSETTABLA & "&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(1),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if


    'descripcion

       nCol = 80
       if  not  isnull(rs(1)) then
           wdescri = rs(1)
       end if

       if  len(wdescri) = 0 then wdescri = " "

       t3b = "<font class=at><textarea class=dt style='font-size:7pt;' "
       t3b = t3b & "name=" & rs(1).name & " cols=" & nCol & " rows=2 "
       if  bModificar then
           t3b = t3b & " onblur=" & sForm & ".submit() "
           t3b = t3b & " onkeypress=" & chr(34)
           t3b = t3b & " if (this.value.length > 450) {alert('El texto es demasiado"
           t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
       else
           t3b = t3b & " readonly >"
       end if
       xdatos =  wdescri
       t3b    = t3b & xdatos & "</textarea></font>"
       SINPUT = t3b

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
  sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si")

  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "RGACTVULNERABILIDAD.asp?OPCION=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select DESCRIPCION From " & WKSETTABLA & " Order By DESCRIPCION "    ' orden pagina
  str2 = "Select DESCRIPCION From " & WKSETTABLA & " Order By DESCRIPCION "    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  'consulta
  sAccion = "RGACTVULNERABILIDADCons.asp?" & request.querystring
  Celda = ofv.BotonConsultar(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>