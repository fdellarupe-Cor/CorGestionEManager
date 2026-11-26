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
  Formulario = ofv.Encabezados("Descripcion","VALOR","OBSERVACIONES","CALIF." )

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  sOPCION       = request.querystring("OPCION")
  IF SOPCION = "SI" THEN

        strsql = "Insert into RGACTEFICIENCIASV "
        strsql = strsql & " (VALORD,DESCRIPCION,VALORH,OBSERVACIONES,CALIF)"
        strsql = strsql & "  Values "
        strsql = strsql &  "('0','NUEVO NIVEL EFICIENCIA','0',' ',' ')"
        xxx=SqlExLog("INSERT",strsql,CN,"RGACTEFICIENCIASV","","","PAR")
  '''response.write strsql & "<br>"
       ' call ofv.crearconsultaEx(StrSql,cn,1,volver)

  END IF

  WKSETTABLA = "RGACTEFICIENCIASV"

  Strsql = "Select * From " & WKSETTABLA & " Order By VALORD"
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


    'Descripcion
    sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"70","dt","")
    Celda = Celda & ofv.gencelda("","UT","","","","","",sInput,"si")

    'VALOR FINAL
    sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"10","dtN","")
    Celda = Celda & ofv.gencelda("","cc","","","","","",sInput,"si")

    'OBSERVACION

       nCol = 80
       if  not  isnull(rs(4)) then
           wdescri = rs(4)
       end if

       if  len(wdescri) = 0 then wdescri = " "

       t3b = "<font class=at><textarea class=dt style='font-size:7pt;' "
       t3b = t3b & "name=" & rs(4).name & " cols=" & nCol & " rows=2 "
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

    'CALIF
    sInput = ofv.GenerarInput(rs(5).name,rs(5),"text",sForm,"70","dt","")
    Celda = Celda & ofv.gencelda("","UT","","","","","",sInput,"si")


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
    sAccion = "RGACTEFICIENCIASV.asp?Opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select descriPCION From " & WKSETTABLA & " Order By VALORD "    ' orden pagina
  str2 = "Select descriPCION From " & WKSETTABLA & " Order By descriPCION "    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  'consulta
  sAccion = "RGACTEFICIENCIASVCons.asp?" & request.querystring
  Celda = ofv.BotonConsultar(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>