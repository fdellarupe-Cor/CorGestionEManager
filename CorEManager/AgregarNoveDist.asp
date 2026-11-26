<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
  TA = request.querystring("TA")
  CantRegMoverx = request.querystring("Cantregmover")
  ID = request.querystring("ID")
  volver = request.querystring("volver")
  sxopc1 = ""


  if TA = "U" then

  set cn = ofv.conectar(ofv.strconn0)

  strsql = "Select * From Usuarios  order by descripcion"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

  do until rs.eof

     sxopc1 = sxopc1 & "<option value=" & rs(0) & ">Usuario:" & rs(1) & "</option> "

  rs.movenext
  loop

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select * From novegenparDist "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)


  sHTML = ofv.FormHeader("Agregar Usuario")

  f1a = "asp/Agregar.asp?Tabla=novegenpardist&volver="
  f1a = f1a & volver
  sHTML = sHTML & ofv.GenForm("AgreArchb","",f1a,"","post","","","no")

  sHTML = sHTML & ofv.GenTabla("panel","","","80%","",0,0,0,0,"","","","no")

  'Encabezados Descripcion y Secuencia
  h1 = ofv.GenCelda("","tt","","2%","","","","Usuario","si")
  h1 = h1 & ofv.GenCelda("","","","2%","","","","&nbsp;","si")

  sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")


  t1 = ofv.GenCelda("","","","","","","","<select class=at name=" & rs(2).name & " >" & sxopc1 & "</select>","si")
  t1 = t1 & ofv.GenCelda("","","","2%","","","","&nbsp;" & ofv.GenerarInput(rs(3).name,"0","hidden","","","",""),"si")
  t1 = t1 & ofv.GenCelda("","","","2%","","","","&nbsp;" & ofv.GenerarInput(rs(1).name,ID,"hidden","","","",""),"si")
  t1 = t1 & ofv.GenCelda("","","","2%","","","","&nbsp;" & ofv.GenerarInput(rs(4).name,"0","hidden","","","",""),"si")


  sHTML = sHTML & ofv.GenRow("","","","","","","",t1 ,"si")
  sHTML = sHTML & ofv.GenFTabla()


  elseIF  TA = "E" THEN

 ' TA = "E" - (Agrega unA EMPRESA COMPLETA)

  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select * From CLIENTEPLAN order by CLIENTE"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

  do until rs.eof


        sxopc1 = sxopc1 & "<option value=" & rs(0) & ">" & rs(1) & "</option> "


  rs.movenext
  loop

  call ofv.cerrarconsulta(rs)


  strsql = "Select * From novegenparDist "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)


  sHTML = ofv.FormHeader("Agregar EMPRESA")

  f1a = "asp/Agregar.asp?Tabla=novegenpardist&volver="
  f1a = f1a & volver


  sHTML = sHTML & ofv.GenForm("AgreArchb","",f1a,"","post","","","no")

  sHTML = sHTML & ofv.GenTabla("panel","","","80%","",0,0,0,0,"","","","no")

  'Encabezados Descripcion y Secuencia
  h1 = ofv.GenCelda("","tt","","2%","","","","EMPRESA","si")
  h1 = h1 & ofv.GenCelda("","","","2%","","","","&nbsp;","si")

  sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")


  t1 = ofv.GenCelda("","","","","","","","<select class=at name=" & rs(3).name & " >" & sxopc1 & "</select>","si")
  t1 = t1 & ofv.GenCelda("","","","2%","","","","&nbsp;" & ofv.GenerarInput(rs(2).name,"TODOS","hidden","","","",""),"si")
  t1 = t1 & ofv.GenCelda("","","","2%","","","","&nbsp;" & ofv.GenerarInput(rs(1).name,ID,"hidden","","","",""),"si")
  t1 = t1 & ofv.GenCelda("","","","2%","","","","&nbsp;" & ofv.GenerarInput(rs(4).name,"0","hidden","","","",""),"si")

  sHTML = sHTML & ofv.GenRow("","","","","","","",t1 ,"si")
  sHTML = sHTML & ofv.GenFTabla()


  else

 ' TA = "S" - (Agrega un sector completo)

  set cn = ofv.conectar(ofv.strconn2)

  strsql = "Select * From Eco  order by descri"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)

  do until rs.eof


        sxopc1 = sxopc1 & "<option value=" & rs(0) & ">" & rs(1) & "</option> "


  rs.movenext
  loop

  call ofv.cerrarconsulta(rs)


  strsql = "Select * From novegenparDist "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)


  sHTML = ofv.FormHeader("Agregar Sector")

  f1a = "asp/Agregar.asp?Tabla=novegenpardist&volver="
  f1a = f1a & volver


  sHTML = sHTML & ofv.GenForm("AgreArchb","",f1a,"","post","","","no")

  sHTML = sHTML & ofv.GenTabla("panel","","","80%","",0,0,0,0,"","","","no")

  'Encabezados Descripcion y Secuencia
  h1 = ofv.GenCelda("","tt","","2%","","","","Usuario","si")
  h1 = h1 & ofv.GenCelda("","","","2%","","","","&nbsp;","si")

  sHTML = sHTML & ofv.GenRow("","","","","","","",h1,"si")


  t1 = ofv.GenCelda("","","","","","","","<select class=at name=" & rs(3).name & " >" & sxopc1 & "</select>","si")
  t1 = t1 & ofv.GenCelda("","","","2%","","","","&nbsp;" & ofv.GenerarInput(rs(2).name,"SECT-","hidden","","","",""),"si")
  t1 = t1 & ofv.GenCelda("","","","2%","","","","&nbsp;" & ofv.GenerarInput(rs(1).name,ID,"hidden","","","",""),"si")
  t1 = t1 & ofv.GenCelda("","","","2%","","","","&nbsp;" & ofv.GenerarInput(rs(4).name,"0","hidden","","","",""),"si")

  sHTML = sHTML & ofv.GenRow("","","","","","","",t1 ,"si")
  sHTML = sHTML & ofv.GenFTabla()

  end if


  'Botones Aceptar y Cancelar
  sAccion = "AdNovedadesDist.asp?ID=" & ID & "&CantRegMover=" & Cantregmoverx
  sHTML = sHTML & ofv.AgregarBoton("AgreArchb",sAccion,"_self")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

response.write sHTML

%>