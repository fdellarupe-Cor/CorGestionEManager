<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0



  Id = request.querystring("Id")
  if id = "" then id = session("subt")

  vuelta = ofv.vueltaasp2("../")
  vuelta2 = ofv.vueltaasp2("")


  sHTML = ofv.menuheader("","")



  set cn = ofv.conectar(ofv.strconn2)

  dopc = "N"

  Strsql = "SELECT * FROM prodserv where codprodserv = " & Id
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if not rs.eof then
      dsubt = rs(1)
      if rs("segaceso") <> "0" or rs("segactual") <> "0" then dopc = "S"
  end if
  call ofv.cerrarconsulta(rs)

  Celda = ofv.gencelda("","tt colspan=3","center","2%","","","","Grupos y Usuarios habilitados para el subtema: " & dsubt,"si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")

  if dopc = "N" then
     Celda = ofv.gencelda("","cc colspan=3","center","2%","","","","Todos los Grupos y Usuarios habilitados","si")
     Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

     Celda = ofv.gencelda("","","","2%","","","","&nbsp;","si")
     Celda = Celda & ofv.gencelda("","","","2%","","","","&nbsp;","si")
     Celda = Celda & ofv.gencelda("","","","2%","","","","&nbsp;","si")
     Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
    
  else


     Celda = ofv.gencelda("","tt","","2%","","","","&nbsp;","si")
     Celda = Celda & ofv.gencelda("","tt","","2%","","","","Usuario/Grupo","si")
     Celda = Celda & ofv.gencelda("","tt","","2%","","","","Permiso","si")

     Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

  
  Strsql = "SELECT * FROM subtemaskr  "
  Strsql = Strsql & " Where codprodserv = " & id


  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then
     Celda = ofv.gencelda("","cc colspan=3","center","2%","","","","Ningun Grupo ni Usuario habilitado","si")
     Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

  else

  do until rs.eof

    sForm = "Sdk" & rs(0)


      sHref = "asp/EliminarRegistro.asp?ID=" & rs(0) & "&Tabla=subtemaskr&volver="
      sHref = sHref & vuelta
      sHref = sHref & "&Eliminar=NSNC"

      Celda = ofv.BotonEliminar("Permiso",sHref)


    sUbic = instr(rs(4),"USR-")
    if not sUbic > 0 then
       sTabla = "Grupo"
    else
       sTabla = "Usuario"
    end if  

    sBusca = mid(rs(4),5,len(rs(4))-4)

    sGrpUsr = "S/D"
    set cx = ofv.conectar(ofv.strconn0)

    Strsql = "SELECT * FROM " & sTabla & "s"
    Strsql = Strsql & " Where ID = '" & sBusca & "' "

    set rx = ofv.crearconsultaEx(StrSql,cx,1,parametros)
    if not rx.eof then sGrpUsr = sTabla & ": " & rx(0) & " - " & rx(1)
    call ofv.cerrarconsulta(rx)
    call ofv.cerrarconn(cx)

    sinput = ofv.gencelda("","cc ","","75%","","","",sGrpUsr,"si")

    Celda = Celda & sinput

    if rs(3) = "A" then 
       sinput = "<font color=#225522>Acceso</font>"
    else
       sinput = "<font color=#993333>Acceso y Actualizacion</font>"
    end if  

    sinput = ofv.gencelda("","cc ","","20%","","","",sinput,"si")

    Celda = Celda & sinput

    Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

   rs.movenext
   loop

  end if

  end if

  celda = ofv.gencelda("","","","1%","","","","&nbsp;","si")
  AgregarA = "AgregarPermisoSubtema.asp"
  Agregar = ofv.GenerarInput("bac1","Agregar","submit","","10","bt","")
  Agregar = ofv.GenCelda("","","","50%","","","",Agregar,"si")
  Agregar = ofv.GenForm("Agregar","",AgregarA,"","post","",Agregar,"si") 


  Retornar = ofv.GenCelda("","","center","50%","","","","&nbsp;","si")

    Fila = fila & ofv.GenRow("","","","","","","",celda & Agregar & Retornar,"si")



  sHTML = sHTML & "<center>"
  sHTML = sHTML & ofv.GenTabla("","","","90% border=0 ","","0","0","1","0","#bbbbbb","",fila,"si")


  call ofv.cerrarconsulta(rs)

  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</center></body></html>"
end if

response.write sHTML

%>