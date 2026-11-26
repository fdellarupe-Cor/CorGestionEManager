<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  ofv.ObtenerAtributos Session("Form2"),""
  bModificar = ofv.AModificar
  if session("Nomodificar") = "S" then bModificar = False

  nCodigo = request.querystring("Codobjeto")
  sTipo   = request.querystring("Tipobjeto")
  sVolver = request.querystring("volver")
  sCod    = request.querystring("cod")
  scircuito    = request.querystring("CIRCUITO")
  vuelta  = ofv.vueltaasp2("../")
  vuelta2 = ofv.vueltaasp2("")




  sHTML   = ofv.FormHeader(session("FormN"))


  set cn = ofv.conectar(ofv.strconn2)
  session("sCod")=sCod
  select case sCod

         case "1"
              sinput = "Eliminar&nbsp;-&nbsp;Copiar&nbsp;-&nbsp;Reemplazar&nbsp;-&nbsp;"
              sinput = sinput & "Restringir&nbsp;-&nbsp;"
              sinput = sinput & "Autorizar:"
              sPerm = sinput
         case "2"
              sPerm = "Modificar"
         case "3"
              sPerm = "Acceder por Consulta"
         case "4"
              sPerm = "Cambiar de Estado"
  end select

   sPerm="Permisos"

  Celda = ofv.gencelda("","ttt colspan=3","center","2%","","","","Autorizados&nbsp;para&nbsp;" & sPerm,"si")
  Fila = fila & ofv.GenRow("","","","","20","","",Celda,"si")

  Celda = ofv.gencelda("","ttr","","2%","","","","&nbsp;","si")
  Celda = Celda & ofv.gencelda("","ttr colspan=2","","2%","","","","Usuario/Grupo","si")
  Fila = fila & ofv.GenRow("","","","","25","","",Celda,"si")







       if session("sCod")<> 1 then
        usr= " and not usracceso is  null"
       else
        usr = " and not usrresp  is null"
       end if
       Strsql = "SELECT * FROM orusrlst  "
       Strsql = Strsql & " Where indicador=" & nCodigo & usr
       response.write strsql
        WKSPTABLA = "orusrlst"



  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  do until rs.eof

    sForm = "Sdk" & rs(0)

    if bModificar then
      sHref = "asp/EliminarEN.asp?ID=" & rs(0) & "&Tabla=" & WKSPTABLA & "&volver="
      sHref = sHref  & vuelta & "?Cod=" & sCod & "&CIRCUITO=" & SCIRCUITO
      sHref = sHref & "&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&Eliminar=NSNC"
    end if
    if sHref <> "" then
      Celda = ofv.BotonEliminar("Permiso",sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if


    if session("scod") <>1 then
    sUbic  = instr(rs(3),"USR-")

    sBusca = mid(rs(3),5,len(rs(3))-4)
    else
    sUbic  = instr(rs(2),"USR-")

    sBusca = mid(rs(2),5,len(rs(2))-4)
    end if

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

        xxxx     = getInfoUser(sBusca)
        sGrpUsr = session("SGU_nombre")

        sGrpUsr = sTabla & ":&nbsp;" & sBusca & " - " & sGrpUsr



    end if





    sinput = ofv.gencelda("","ttt colspan=2","","60%","","","",sGrpUsr,"si")

    Celda = Celda & sinput

    Fila  = fila & ofv.GenRow("","","","","20","","",Celda,"si")

   rs.movenext
   loop

  sHTML = sHTML & "<br><center>"
  sHTML = sHTML & ofv.GenTabla("","","","90% style='border-style:solid;border-width:1;border-color:#999999;' ","","0","0","-2","1","","",fila,"si")
  sHTML = sHTML & "</center>"


  sAccion = "OrAgregarPermiso.asp?Codobjeto=" & nCodigo & "&CIRCUITO=" & SCIRCUITO
  sAccion = sAccion & "&Tipobjeto=" & sTipo & "&Cod=" & sCod & "&volver=" & vuelta
  Agregar = ofv.BotonAgregar(sAccion)

  sAccion  = "orcargainfoindi.asp?id=" & nCodigo
  Retornar = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Retornar,"")

  call ofv.cerrarconsulta(rs)

  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</center></body></html>"
end if

response.write sHTML


%>