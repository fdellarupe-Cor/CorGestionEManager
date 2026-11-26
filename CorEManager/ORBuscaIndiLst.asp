<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  CantRegMostrarx = 6
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta        = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)





    if sMODAL = "ENCA"  then
       WSPSESION = "ORWHERE2"
    ELSE
       WSPSESION = "ORWHERE"
    END IF



  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)


  IF SOPCION = "VS" THEN


    SORCODIGO   = request.FORM("ORCODIGO")
    sORCARACTER = request.FORM("ORCARACTER")
    SORTEXTO    = request.FORM("ORTEXTO")
    sORRIESGO   = request.FORM("ORRIESGO")
    SORCUENTA   = request.FORM("ORCUENTA")
    SORCLIENTE  = request.FORM("ORCLIENTE")
    SORMONTOD   = request.FORM("ORMONTOD")
    SORMONTOH   = request.FORM("ORMONTOH")
    SORPROCESO  = request.FORM("ORPROCESO")
    SORLOTE     = request.FORM("ORLOTE")





       SESSION(WSPSESION) =  ""

    IF SORCODIGO <> "TODOS" THEN

       SORCODIGOW   = " (A.OBJID = 0" & SORCODIGO & ") "

       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORCODIGOW
       SCONECTOR = " AND "

    END IF

    IF sORCARACTER <> "TODOS" THEN
       sORCARACTERW = " (A.TRCCARACTER = 0" & sORCARACTER & ") "

       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & sORCARACTERW
       SCONECTOR = " AND "

    END IF


    IF SORTEXTO <> "TODOS" THEN
       SORTEXTOW    = " (A.TRCOBSERV LIKE '%" & SORTEXTO & "%') "

       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORTEXTOW
       SCONECTOR = " AND "

    END IF


    IF sORRIESGO <> "TODOS" THEN
       sORRIESGOW   = " (A.TRCEVENTORIESGO1 = 0" & sORRIESGO & ") "

       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & sORRIESGOW
       SCONECTOR = " AND "

    END IF









  ''response.write strsql & "<br>"



   Celda = ""

   wkproccod = "&nbsp;"

   Celda = Celda & ofv.GenCelda("","uti  ","","90%","","","","BASE DE EVENTOS DE RIESGO " & WKACCION,"si")
   Celda = Celda & ofv.GenCelda("","ut   ","","2%","","","",wkproccod,"si")
   Fila  = ofv.GenRow("","","","","","","",Celda,"si")

  sHTML = sHTML & ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ","","100% ","30","0","","-1","1","","",Fila,"si")

   Celda = ""
   Celda = Celda & ofv.GenCelda("","uti  ","","10%","","","","SELECCION: ","si")
   Celda = Celda & ofv.GenCelda("","ut   ","","90%","","","",WSPWHERE,"si")
   Fila  = ofv.GenRow("","","","","","","",Celda,"si")

 ' sHTML = sHTML & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;BORDER-COLOR:#993333;' valign='middle'  ","","100% ","30","0","","-1","1","","",Fila,"si")
  sHTML = sHTML & "<br>"











  'Encabezados

   Formulario = ""
   FILAX = ""
   Celda = ""


' response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then


  ELSE
  posI = 0

  xTrColor = "#f0f0f0"
  '''response.write strsql & "<br>"

  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)


    sForm = "form" & posi


    Filax  = ""

    Celda = ""

    'CODIGO
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;' ","","5%","","","","CODIGO","si")


    'INCIDENTE
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;' ","","30%","","","","FECHA","si")

    'CARACTER
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","30%","","","","CARACTER","si")

    'CUENTA
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","30%","","","","CUENTA","si")

  ''  Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")



    Celda = ""


    'CODIGO
    SINPUT = RS(0)







    'F.INCIDENTE
    if clng(rs(3)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSH(rs(3))
    end if
    SINPUT = dFechan
    Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","FECHA:&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","25%","","","",sInput,"si")



    'CARACTER

    sInputx = "-------"

    strsql = "Select descri From ORTIPOCARACTER where OBJID = 0" & rs(24)
    set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rsniv.eof then sInputx = rsniv(0)
    call ofv.cerrarconsulta(rsniv)

    sInput = sInputx
    Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","CARACTER:&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")



    'CUENTA CONTABLE
    if (rs(29) & "") = "" then
      SINPUT = "-------"
    else
      SINPUT = RS(29)
    end if
    Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","CUENTA:&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")




    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")




    Celda = ""
    'CLIENTE
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;' ","","","","","","CLIENTE","si")

    'MONTO
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","MONTO","si")

    'ESTADO
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","ESTADO","si")

  '  Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")



    Celda = ""
    'CLIENTE
    if (rs(28) & "") = "" then
      SINPUT = "-------"
    else
      SINPUT = RS(28)
    end if
    Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","CLIENTE:&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    'MONTO
    if clng("0" & rs(16)) = 0 then
      SINPUT = "-------"
    else
      SINPUT = RS(16)
    end if
    Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","MONTO NETO:&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    'ESTADO
    wkestado = rs(17) & ""

    select case wkestado

           case "CAN"

                 wkestadoD = "<font color='#888888' >CANCELADO</font>"

           case "REG"

                 wkestadoD = "<font color='#993333' >REGISTRADO</font>"


           case "ENR"

                 wkestadoD = "<font color='#009900' >ENRIQUECIDO</font>"

           case "ENRF"

                 wkestadoD = "<font color='#009900' >PENDIENTE</font>"


           case "FIN"

                 wkestadoD = "<font color='#000000' ><B>FINALIZADO</B></font>"

           CASE ELSE


                 wkestadoD = "<font color='#000099' >EN REVISION</font>"


    end select

    sInput = wkestadoD

    Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","ESTADO:&nbsp;","si")
    Celda = Celda & ofv.gencelda("","uTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")




    Celda = ""
    'DESCRIPCION
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;' COLSPAN=3 ","","","","","","DESCRIPCION","si")

  '  Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")



    Celda = ""
    'OBSERVACIONES


    SINPUT = RS(11) & "&nbsp;"
    Celda = Celda & ofv.gencelda("","UTi style='background-color:#dddddd;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","DESCRIPCION:&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UTi COLSPAN=5 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")



    Formulario = Formulario & Filax


    sInput = "&nbsp;"
    Celda = ofv.gencelda("","UT colspan=7 style='border-TOP-style:SOLID;border-TOP-width:2;border-TOP-color:#000000;' ","","","","","",sInput,"si")

    Filax  = ofv.genrow("","AAA bgcolor='#fff9f9' ","","","25","","",Celda,"si")


 ''''

      Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"_BLANK","post","",FilaX,"si")





    posi = posi + 1
    rs.Movenext
  Loop
  END IF
  sHTML = sHTML & ofv.GenTabla("","","","98% ","","0","","0","0","","",Formulario,"si")

  'Botonera
  Agregar = ""
  if bAgregar AND sMODAL = "REG" then
    sAccion = "ORBASERIESGOSADD.asp?MODAL=AAG"
    Agregar = ofv.BotonAgregar(sAccion)
  end if
  Buscar  = ""
  Celda   = ""

  sAccion = "orbaseriesgos.asp?modal=" & smodal & "&CODTRACER=" & SCODTRACER
  Celda = ofv.BotonVolver(sAccion)

  WSPREEMP = "modal=" & smodal  & "&CODTRACER=" & SCODTRACER & "&"

  sHTML = sHTML &  replace(ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","",Agregar,Buscar,celda),"cantregmostrar=",WSPREEMP & "cantregmostrar=")

  call ofv.cerrarconsulta(rs)







  else



  end if


  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML



%>