<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  ofv.ObtenerAtributos Session("Form"),""
  bModificar = ofv.AModificar


  CantRegMostrarx = 8
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
  SOPCION = REQUEST.QUERYSTRING("OPCION")

  session("modalidad") = "PRGESTADO"

  SNIVEL1  = REQUEST.QUERYSTRING("NIVEL1")


  IF SNIVEL1 <> "" THEN
     SESSION("NIVEL1") = SNIVEL1
  ELSE
     SNIVEL1 = SESSION("NIVEL1")
  END IF

'  vuelta  = ofv.vueltaasp2("../")
'  vuelta2 = ofv.vueltaasp2("")

  set cn = ofv.conectar(ofv.strconn2)


  IF SOPCION = "SI" THEN

     WKUSRGRP = REQUEST.FORM("USRGRP")

        strsql = "Insert into ORESTADOSPER "
        strsql = strsql & " (ESTADO,PERMISO,USRGRP)"
        strsql = strsql & "  Values "
        strsql = strsql &  "('" & SNIVEL1 & "','1','" & WKUSRGRP & "')"
   '''   response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

  END IF




  sHTML   = ofv.FormHeader(session("FormN"))


  WKTEVENTO  = "NO DEFINIDO"
  WKSETTABLA = "ORESTADOS"

  Strsql = "Select * From ORESTADOS WHERE OBJID = 0" & SNIVEL1
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     WKTEVENTO = RS(2)


    Celda = ""
    Celda = Celda & ofv.gencelda("","VTI STYLE='COLOR:#000099;' ","","","","","","ESTADO:&nbsp;" & WKTEVENTO,"si")

    Fila = ofv.genrow("","","","","25","","",Celda,"si")
    sHTML = sHTML & ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ","","100% ","","0","","-1","1","","",Fila,"si") & "<BR>"


  END IF


  Celda = ""
  Fila  = ""


   Celda = ofv.gencelda("","ttt colspan=3","center","2%","","","","Autorizados&nbsp;para&nbsp;" & sPerm,"si")
  Fila = fila & ofv.GenRow("","","","","20","","",Celda,"si")

  Celda = ofv.gencelda("","ttr","","2%","","","","&nbsp;","si")
  Celda = Celda & ofv.gencelda("","ttr colspan=2","","2%","","","","Usuario/Grupo","si")
  Fila = fila & ofv.GenRow("","","","","25","","",Celda,"si")





       Strsql = "SELECT OBJID, * FROM ORESTADOSPER  "
       Strsql = Strsql & " Where ESTADO = " & SNIVEL1

      WKSPTABLA = "ORESTADOSPER"

   response.write strsql

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  IF RS.EOF THEN

        strsql = "Insert into ORESTADOSPER "
        strsql = strsql & " (ESTADO,PERMISO,USRGRP)"
        strsql = strsql & "  Values "
        strsql = strsql &  "('" & SNIVEL1 & "','0','TODOS')"
   '''   response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

       Strsql = "SELECT OBJID, * FROM ORESTADOSPER  "
       Strsql = Strsql & " Where ESTADO = " & SNIVEL1
       set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


  END IF




  do until rs.eof

    sForm = "Sdk" & rs(0)


    if bModificar and (rs(4) & "") <> "TODOS" then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=" & WKSPTABLA & "&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar("PERMISO",sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if


    SDATO = rs(4) & ""
    sUbic  = instr(rs(4),"USR-")
    sBusca = mid(rs(4),5,len(rs(4))-4)
    sGrpUsr = "S/D"

    if not sUbic > 0 then


     IF SDATO = "TODOS" THEN
        sGrpUsr = "TODOS LOS USUARIOS"

     ELSE
       sTabla = "Grupo"

       set cx = ofv.conectar(ofv.strconn0)
       Strsql = "SELECT * FROM " & sTabla & "s"
       Strsql = Strsql & " Where ID = '" & sBusca & "' "
       set rx = ofv.crearconsultaEx(StrSql,cx,1,parametros)
       if not rx.eof then sGrpUsr = sTabla & ": " & rx(0) & " - " & rx(1)
       call ofv.cerrarconsulta(rx)
       call ofv.cerrarconn(cx)

     END IF


    else
       sTabla = "Usuario"
       ' response.write sbusca
        xxxx     = getInfoUser(sBusca)
        sGrpUsr = session("SGU_nombre")

        sGrpUsr = sTabla & ":&nbsp;" & sBusca & " - " & sGrpUsr



    end if


    sinput = ofv.gencelda("","VTI","","60%","","","",sGrpUsr,"si")

    Celda = Celda & sinput

    'ENVIA
    scomK = "0" & chr(9) & "1" & chr(9) & "2" & chr(9) & "3"
    scomD = "NO" & chr(9) & "ACTUALIZAR" & chr(9) & "CAMBIO DE ESTADO" & chr(9) & "SOLO CONSULTAR"
    sInput = ofv.gencomboFX(sForm,rs(3).name,rs(3),scomK,scomD," ")
    Celda = Celda & ofv.gencelda("","UT","","20%","","","",sInput,"si")

    FilaX = ofv.genrow("","","","","20","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=" & WKSPTABLA & "&volver=" & vuelta
    end if
    Fila  = fila & ofv.genform(sForm,"",sAccion,"","post","",FilaX,"si")



   rs.movenext
   loop

  sHTML = sHTML & "<br><center>"
  sHTML = sHTML & ofv.GenTabla("","","","90% style='border-style:solid;border-width:1;border-color:#999999;' ","","0","0","-2","1","","",fila,"si")
  sHTML = sHTML & "</center>"


  sAccion = "AgregarPermiso.asp?Codobjeto=0&CIRCUITO=" & SNIVEL1
  sAccion = sAccion & "&Tipobjeto=0&Cod=1&volver=" & vuelta
  Agregar = ofv.BotonAgregar(sAccion)

  sAccion  = "ORESTADOEVENTO.asp"
  Retornar = ofv.BotonVolver(sAccion)



  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Retornar,"")

  call ofv.cerrarconsulta(rs)

  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</center></body></html>"
end if

response.write sHTML

%>