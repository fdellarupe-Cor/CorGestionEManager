<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  vuelta = ""
  sHTML = ofv.FormBody

  Id     = request.querystring("Id")
  Volver = request.querystring("vlt")
  if  Volver = "" then   Volver = request.querystring("Volver")

  if  Id <> "" then
      session("Subt") = Id
      session("vlt")  = volver
  else
      Id              = Session("subt")
      volver          = session("vlt")
  end if


    shtml = ofv.FormHeader("Permisos de Subtema")


  vuelta = ofv.vueltaasp2("../")

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "SELECT * FROM prodserv where codprodserv = " & Id
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if not rs.eof then

     sForm = "perst"

    sHTML = sHTML & "<center>"

 '''''   shtml = replace(shtml,"<body ","<body onload=document.activ.submit(); ")

    Celda = ofv.gencelda("","ttr","center","2%","","","","Acceso","si")
    Celda = celda & ofv.gencelda("","ttr","center","2%","","","","Actualizacion","si")
    Fila  = ofv.GenRow("","","","","","","",Celda,"si")

    sxsel = ""
    if rs("segaceso") <> "0" then sxsel = " selected "


    scomK = "0" & chr(9) & "1"
    scomD = "Publico" & chr(9) & "Restringido"


    sInput = ofv.gencomboFX(sForm,rs("segaceso").name,rs("segaceso"),scomK,scomD," ")

    Celda = ofv.gencelda("","","center","50%","","","",sInput,"si")

    sxsel = ""
    if rs("segaceso") <> "0" or rs("segactual") <> "0" then sxsel = " selected "


    scomK = "0" & chr(9) & "1"
    scomD = "Publico" & chr(9) & "Restringido"


    sInput = ofv.gencomboFX(sForm,rs("segactual").name,rs("segactual"),scomK,scomD," ")

    Celda = Celda & ofv.gencelda("","","center","50%","","","",sInput,"si")


    Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")



    sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
    sAccion = sAccion & "&tabla=prodserv&volver=" & vuelta
  '    response.write vuelta

    Fila = ofv.GenForm("perst","",sAccion,"","post","",Fila,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa","","30%","","0","0","1","0","","",fila,"si")

    sAccion = "subtemaskrdet.asp?id=" & Id
    sHTML = sHTML & ofv.GenForm("activ","",sAccion,"Principal","post","","&nbsp;","si")


  dopc = "N"

  Strsql = "SELECT * FROM prodserv where codprodserv = " & Id
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if not rs.eof then
      dsubt = rs(1)
      if rs("segaceso") <> "0" or rs("segactual") <> "0" then dopc = "S"
  end if
  call ofv.cerrarconsulta(rs)

  Celda = ofv.gencelda("","ttt colspan=3","center","2%","","","","Grupos y Usuarios habilitados para el subtema: " & dsubt,"si")
  Fila = ofv.GenRow("","","","","25","","",Celda,"si")

  if dopc = "N" then
     Celda = ofv.gencelda("","vt colspan=3","center","2%","","","","<big>Todos los Grupos y Usuarios habilitados</big>","si")
     Fila = fila & ofv.GenRow("","aaa bgcolor=#f9f9f9 ","","","30","","",Celda,"si")

  else

     Celda = ofv.gencelda("","ttr","","1%","","","","&nbsp;","si")
     Celda = Celda & ofv.gencelda("","ttr","","75%","","","","Usuario/Grupo","si")
     Celda = Celda & ofv.gencelda("","ttr","","25%","","","","Permiso","si")

     Fila = fila & ofv.GenRow("","","","","25","","",Celda,"si")


     Strsql = "SELECT * FROM subtemaskr  "
     Strsql = Strsql & " Where codprodserv = " & id


  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then
     Celda = ofv.gencelda("","vt colspan=3","center","2%","","","","<big>No hay Grupos ni Usuarios habilitados</big>","si")
     Fila = fila & ofv.GenRow("","aaa bgcolor=#f9f9f9 ","","","30","","",Celda,"si")

  else

  do until rs.eof

    sForm = "Sdk" & rs(0)


      sHref = "asp/EliminarRegistro.asp?ID=" & rs(0) & "&Tabla=subtemaskr&volver="
      sHref = sHref & vuelta
      sHref = sHref & "&Eliminar=NSNC"

      Celda = ofv.BotonEliminar("Permiso",sHref)


    sUbic = instr(rs(4),"USR-")
    sBusca = mid(rs(4),5,len(rs(4))-4)
    sGrpUsr = "S/D"

    if not sUbic > 0 then
       sTabla = "Grupo"

       set cx = ofv.conectar(ofv.strconn0)
       Strsql = "SELECT * FROM " & sTabla & "s"
       Strsql = Strsql & " Where ID = '" & sBusca & "' "

       set rx = ofv.crearconsultaEx(StrSql,cx,1,parametros)
       if not rx.eof then sGrpUsr = sTabla & ": " & rx(0) & " - " & rx(1)
       call ofv.cerrarconsulta(rx)
       call ofv.cerrarconn(cx)


    else
        sTabla = "Usuario"

        xxxx    = getInfoUser(sBusca)
        sGrpUsr = session("SGU_nombre")
        sGrpUsr = sTabla & ": " & sBusca & " - " & sGrpUsr
    end if




    sinput = ofv.gencelda("","ttt ","","","","","",sGrpUsr,"si")

    Celda = Celda & sinput

    if rs(3) = "A" then
       sinput = "<font color=#225522>Acceso</font>"
    else
       sinput = "<font color=#993333>Acceso y Actualizacion</font>"
    end if

    sinput = ofv.gencelda("","vt ","","","","","",sinput,"si")

    Celda = Celda & sinput

    Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

   rs.movenext
   loop

  end if

  end if



  sHTML = sHTML & ofv.GenTabla("","aaa ","","70%","","0","0","1","0","","",fila,"si")



 end if

  sHTML = sHTML & "</center>"

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


'  Agregar = ""
  Buscar  = ""

  'Botonera
  sAccion = "AgregarPermisoSubtema.asp"
  Agregar = ofv.BotonAgregar(sAccion)

  sAccion = replace(volver,"_","&")
  Celda   = ofv.BotonVolver(sAccion)


  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)


  sHTML = sHTML & "</center></body></html>"
end if

response.write sHTML

%>