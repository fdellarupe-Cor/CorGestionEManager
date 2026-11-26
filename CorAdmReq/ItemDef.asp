<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))

  'Encabezados
  formulario =  ofv.Encabezados("Descripcion","activa","Orden")

  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn5)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

 sopcion = request.querystring("opcion")

 if sopcion = "si" then

    strsql = "insert into itemdef (descripcion,orden,activa) values ('INFO','0','1')"
    call ofv.crearconsultaEx(StrSql,cn,1,parametros)

 end if


  StrSql = "Select * From ItemDef order by orden,id"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)

    if rs(2) then
       sc = "tt"
    else
       sc = "cc"
    end if

    sInput1 = ofv.GenerarInput(rs(0).name,rs(0),"hidden","niveles" & rs(0),"10","dt","")
    celdas = ofv.GenCelda("",sc,"","","","","","&nbsp;","si")

    if bModificar then
      sInput = ofv.GenerarInput(rs(1).name,rs(1),"text","niveles" & rs(0),"90","dt","")
    else
      sInput = ofv.GenerarInput("",rs(1),"readonly","","90","dt","")
    end if
    celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")

    if bModificar then
      sInput = replace(ofv.combosino("niveles" & rs(0),rs(2),rs(2).name),"-1","1")
    else
      sInput = ofv.GenerarInput("",rs(2),"readonly","","10","dt","")
    end if
    celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")

    if bModificar then
      sInput = ofv.GenerarInput(rs(3).name,rs(3),"text","niveles" & rs(0),"5","dtn","")
    else
      sInput = ofv.GenerarInput("",rs(3),"readonly","","5","dtn","")
    end if
    celdas = celdas & ofv.GenCelda("",sc,"","","","","",sInput,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")

    if bModificar then
      sAccion = "asp/Actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*") & "&volver=" & vuelta & "&Tabla=itemdef"
      formulario = formulario & ofv.GenForm("niveles" & rs(0),"",sAccion,"","post","",filas,"si")
    else
      formulario = formulario & ofv.GenForm("","","","","post","",filas,"si")
    end if

    posi = posi + 1
    rs.Movenext
  loop

  sHTML = sHTML  & ofv.GenTabla("","","","98%","","0","","0","0","","",formulario,"si")


  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "Itemdef.asp?Opcion=si"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select descripcion From ItemDef order by orden,id "    ' orden pagina
  str2 = "Select descripcion From ItemDef order by descripcion "    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))
  

  
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML

%>