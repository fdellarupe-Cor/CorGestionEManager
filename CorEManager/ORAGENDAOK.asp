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
  vuelta        = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
  svision       = request.querystring("VISION")
  svisionMAIN   = request.querystring("VISIONMAIN")
  sACTIVRIESGO  = request.querystring("ACTIVRIESGO")
  sFECHA        = request.querystring("FECHA")

  if svision = "" then svision = "AGN"

  sHTML = ofv.FormHeader(session("FormN"))

  WKVISION = ""
 if svision="PENALL" THEN

    WKVISION="PENALL"

 ELSEIF svision="AGNALL" THEN
      WKVISION="AGNALL"

 ELSEIF  svision = "PEN" then

     WKVISION = "PEN"

  ELSEIF  svision = "AGN" then

     WKVISION = "AGN"

  ELSEIF  svision = "IAG" then


    if svisionMAIN = "PEN" then

       WKVISION = "AGN"

    ELSE

       WKVISION = "PEN"

    END IF

  ELSE

    if svisionMAIN = "PEN" then

       WKVISION = "PEN"

    ELSE

       WKVISION = "AGN"

    END IF

  END IF




response.write WKVISION
SELECT CASE WKVISION

  if WKVISION = "PEN" then

       WKVISION = "AGN"
       WKVISDES = "AGENDA DE SEGUIMIENTO"
       sImage = "AGENDA"

  ELSE

       WKVISION = "PEN"
       WKVISDES = "ACTIVIDADES CON PLAN DE MITIGACION REQUERIDO Y NO AGENDADO."
       sImage = "PENDIENTES"


  end if

       wkproccodL = "<a class=nt href='ORAGENDA.ASP?VISION=" & WKVISION & "' "
       wkproccodL = wkproccodL & " TITLE='" & WKVISDES & "' ><U>"
       wkproccodL = wkproccodL & sImage
       wkproccodL = wkproccodL & "</U></a>"

    Celda = ""



   Celda = Celda & ofv.GenCelda("","uti  ","","90%","","","","SEGUIMIENTO PLAN DE MITIGACION","si")
 '  IF svision="PENALL" THEN
 '  wkproccodL = "<a class=nt href='ORAGENDA.ASP?VISION=" & svision & "' "
  '        wkproccodL = wkproccodL & " TITLE='" & WKVISDES & "' ><U>"
 '         wkproccodL = wkproccodL & sImage
 '      wkproccodL = wkproccodL & "</U></a>"
 '  END IF


   Celda = Celda & ofv.GenCelda("","ut  style='border-style:OUTSET;border-width:2;' ","CENTER","8%","","","",wkproccodL,"si")

   Celda = Celda & ofv.GenCelda("","ut   ","","2%","","","","&nbsp;","si")
   Fila  = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML & ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ","","100% ","30","0","","-1","1","","",Fila,"si")


 ' OFV.ObtenerAtributosNodo Session("Form"), "", "USUARIO", "TODOS"
  'bUSUARIO = OFV.AConsultar

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)


'''''''''''''''''''''''''''''''''''''''
'' POR AAG - INCORPORO ACTIVIDADES EN LA AGENDA

  if svision = "AAG" then

        strsql = "Insert into ORAGENDA "
        strsql = strsql & " (DESCRI,FECCUMPLI,FECREVISION,ESTADO, RESPONSABLE,"
        strsql = strsql & " ACTIVRIESGO, ACTIVIDAD, PERIODO, CODAGENDA, FECREAL) "
        strsql = strsql & " Values ("
        strsql = strsql &  "'" & REQUEST.FORM("DESCRI") & "','" & REQUEST.FORM("FECCUMPLI") & "','" & REQUEST.FORM("FECREVISION") & "',"
        strsql = strsql &  "'" & REQUEST.FORM("ESTADO") & "','" & REQUEST.FORM("RESPONSABLE") & "',"
        strsql = strsql & "'" & REQUEST.FORM("ACTIVRIESGO") & "','" & REQUEST.FORM("ACTIVIDAD") & "',"
        strsql = strsql & "'" & REQUEST.FORM("PERIODO") & "','" & REQUEST.FORM("CODAGENDA") & "',"
        strsql = strsql & "'" & REQUEST.FORM("FECREAL") & "'"
        strsql = strsql & ")"
  '''response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)


        strsql = "UPDATE ORACTIVRIESGO SET INCPLAN = 'SI' WHERE CODRIESGO = '" & REQUEST.FORM("ACTIVRIESGO") & "' "
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

    svision = svisionMAIN

  end if






     if session("ORPERIODO") = "" then

        Strsql = "Select * From ORPERIODOS  where ESTADO IN ('OK','CL')  "
        Strsql = Strsql & " ORDER BY FINAL,INICIO, ESTADO DESC"
        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if NOT rs.eof then
           session("ORPERIODO")     = RS(0)
           session("ORPERIODODESC") = RS(1)
        END IF

     END IF

'RESPONSE.WRITE svision
  if svision = "PEN" OR svision="PENALL" then




  'Encabezados
  Formulario = ""

    Celda = ""

    Celda = Celda & ofv.gencelda("","VTI style='background-color:#F7CC88;' ","","40%","","","","PROCESO","si")
    Celda = Celda & ofv.gencelda("","VTI style='background-color:#F7DD88;' ","","25%","","","","ACTIVIDAD","si")
    Celda = Celda & ofv.gencelda("","VTI style='background-color:#F7EE88;' COLSPAN ","","45%","","","","RIESGO","si")
   Celda = Celda & ofv.gencelda("","VTI style='background-color:#F0F0F0;' ","","10% ALIGN=LEFT","","","","PERIODO","si")
    Formulario = ofv.genrow("","","","","25","","",Celda,"si")


  WKSETTABLA = "ORAGENDA"

  xTrColor = "#FFFFFF"


  IF SVISION="PEN" THEN
   Strsql = ""

    strsql= "select         d.secuen, "
    strsql =strsql & "     ISNULL(e.descri,D.DESCRIACC), g.descri, F.CODRIESGO, "
    strsql =strsql & "a.*  ,b1.fecvigen,H.PERIODO  "
    strsql =strsql & "from csrelpathdescri a    "
    strsql =strsql & "inner join objetos b on (a.cobj = b.codobjeto) "
 	strsql =strsql & "inner join objlnk c on (a.cobj = c.codobjeto) "
    strsql =strsql & " inner join objaccion d on (c.codobjetol = d.codobjeto and c.iden4 = d.codproced) "
    strsql =strsql & "LEFT join tipoaccalt e on (d.tipoaccion = e.tipoaccionalt)      "
    strsql =strsql & "left join oractivriesgo f on (d.codaccion = f.actividad )     "
    strsql =strsql & " left join oRPERIODOS H on (H.OBJID = F.PERIODO) "
    strsql =strsql & "INNER join ORRSKCIRCUITO b1 on (f.tipobjeto=b1.tipobjeto and f.periodo=b1.periodo and  f.codobjeto = b1.codobjeto )     "
    strsql =strsql & "inner join OBJWFLDET b2 on (b1.objid = b2.codobjeto and b2.tipobjeto='RSK' AND B2.comen1='REFERENTE' and B2.usr='" & session("usuario") & "' )   "
    strsql =strsql & "left join orimpriesgo g on "

    strsql =strsql & "(f.itemriesgo = g.objid) where d.critico = 'SI' and  F.requiereplan = 'SI' and   (F.incplan IS NULL OR F.incplan <> 'SI' )  order by b1.fecvigen ASC "


  ELSE
   strsql= "select         d.secuen, "
   strsql =strsql & "     ISNULL(e.descri,D.DESCRIACC), g.descri, F.CODRIESGO, "
   strsql =strsql & "a.*  ,H.PERIODO "
   strsql =strsql & "from csrelpathdescri a    "
   strsql =strsql & "inner join objetos b on (a.cobj = b.codobjeto) "
	strsql =strsql & "inner join objlnk c on (a.cobj = c.codobjeto) "
   strsql =strsql & " inner join objaccion d on (c.codobjetol = d.codobjeto and c.iden4 = d.codproced) "
   strsql =strsql & "LEFT join tipoaccalt e on (d.tipoaccion = e.tipoaccionalt)      "
   strsql =strsql & "left join oractivriesgo f on (d.codaccion = f.actividad )     "
   strsql =strsql & " left join oRPERIODOS H on (H.OBJID = F.PERIODO) "
   strsql =strsql & "INNER join ORRSKCIRCUITO b1 on (f.tipobjeto=b1.tipobjeto and f.periodo=b1.periodo and  f.codobjeto = b1.codobjeto )     "
   strsql =strsql & "inner join OBJWFLDET b2 on (b1.objid = b2.codobjeto and b2.tipobjeto='RSK'  )   "
   strsql =strsql & "left join orimpriesgo g on "

   strsql =strsql & "(f.itemriesgo = g.objid) where d.critico = 'SI' and  F.requiereplan = 'SI' and   (F.incplan IS NULL OR F.incplan <> 'SI' )  order by b1.fecvigen ASC "

  END IF

 'response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then


  ELSE
  posI = 0
  Do While (Not rs.EOF)
    sForm = "form" & posi


      Celda = ""

      WKIPROCD = ""

       FOR WKI=4 TO 11

           WKIPROC = RS(WKI) & ""

           IF WKIPROC <> "" THEN
              WKIPROCD = WKIPROCD & RS(WKI) & " / "
           END IF


       NEXT

    IF LEN(WKIPROCD) > 4 THEN
       WKIPROCD = LEFT(WKIPROCD,LEN(WKIPROCD)-3)
       WKIPROCDA = "<A HREF='javascript:void(0);' onclick='alert(this.innerHTML);' "
       WKIPROCDA = WKIPROCDA & " STYLE='COLOR:#888888;FONT-SIZE:7PT;FONT-FAMILY:""TIMES NEW ROMAN"";' >" & WKIPROCD & "</A>"
    END IF

    'PROCESO
    sInput = WKIPROCDA
    Celda = Celda & ofv.gencelda("","vT","","","","","",sInput,"si")

    'ACTIVIDAD
    sInput = rs(0) & "." & RS(1)
    Celda = Celda & ofv.gencelda("","vT style='color:#993333;' ","","","","","",sInput,"si")

    'RIESGO
    sInput = rs(2) & ""
    Celda = Celda & ofv.gencelda("","UT","","","","","",sInput,"si")
    'FECHA

     if rs("fecvigen")="0" or rs("fecvigen"&"") ="" then
      sInput =rs("PERIODO")& ""
     else
      sInput = GTFechadorSH(rs("fecvigen"))
     end if
    Celda = Celda & ofv.gencelda("","UT","","","","","",sInput,"si")

    'AGENDAR
   IF SVISION="PENALL" THEN
       SINPUT="&nbsp;"
    ELSE
      sInput = "<INPUT TYPE='SUBMIT' CLASS=BT VALUE='AGENDAR' > "
    END IF
    Celda = Celda & ofv.gencelda("","UT","","","","","",sInput,"si")

    Filax  = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

       if xTrColor = "#FFFFFF" then
          xTrColor = "#F0F0F0"
       ELSE
          xTrColor = "#FFFFFF"
       END IF


    sAccion = ""
    if bModificar then
      sAccion = "ORAGENDA.ASP?VISION=IAG&VISIONMAIN=" & SVISION & "&ACTIVRIESGO=" & RS(3)
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Filax,"si")

    posi = posi + 1
    rs.Movenext
  Loop
  END IF
  sHTML = sHTML & ofv.GenTabla("","","","98% ","","0","","0","0","","",Formulario,"si")

  'Botonera
  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)




  elseif svision = "AGN" OR SVISION="AGNALL" then

    xTrColor = "#FFFFFF"

       wkstylo1 = " style='background-color:#F7CC88;color:#ffffff;' "

  sImageva    = "<font face='corflaticon' size=2 color='#990000'>&#0110;</font>"
  sImagevm    = "<font face='corflaticon' size=2 color='#f7aa22'>&#0110;</font>"
  sImagevb    = "<font face='corflaticon' size=2 color='#009900'>&#0110;</font>"
  sImagevn    = "<font face='corflaticon' size=2 color='#dddddd'>&#0110;</font>"

  sImageresp    = "<font face='corflaticon' size=2 color='#990000' style='border-style:OUTSET;border-width:1;border-color:#F7CC88;cursor:pointer;' >&#0128;</font>"

  sImagePROC    = "<font face='corflaticon' size=2 color='#888888' style='border-style:OUTSET;border-width:1;border-color:#F7CC88;cursor:pointer;' >2</font>"

  sImagePROY    = "<font face='corflaticon' size=2 color='#6699cc' style='border-style:OUTSET;border-width:1;border-color:#F7CC88;cursor:pointer;' >&#0204;</font>"

  sImageFNL     = "<font face='corflaticon' size=3 color='#009900' style='border-style:OUTSET;border-width:1;background-color:#F7eedd;' >&#097;</font>"

  'Encabezados
  Formulario = ""

    Celda = ""

    Celda = Celda & ofv.gencelda("","VTI style='background-color:#F7CC88;' COLSPAN=6 ","","75%","","","","DESCRIPCION","si")
    Celda = Celda & ofv.gencelda("","VTI style='background-color:#F7DD88;' COLSPAN=2 ","","15%","","","","FECHA","si")
    Celda = Celda & ofv.gencelda("","VTI style='background-color:#F7EE88;'  ","","10%","","","","ESTADO","si")
    Formulario = ofv.genrow("","","","","25","","",Celda,"si")


  WKSETTABLA = "ORAGENDA"
 IF SVISION="AGN" THEN
  Strsql = "Select * From " & WKSETTABLA & " WHERE ESTADO > 0 AND ESTADO < 5 and responsable='" & session("usuario") & "' Order By FECCUMPLI DESC, ESTADO "
 ELSE
  Strsql = "Select * From " & WKSETTABLA & " WHERE ESTADO > 0 AND ESTADO < 5  Order By FECCUMPLI DESC, ESTADO "
 END IF
 ' Strsql = "Select * From " & WKSETTABLA & " WHERE ESTADO > 0 AND ESTADO < 5  Order By FECCUMPLI DESC, ESTADO "

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posI = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi



      Celda = ""

    sAccion = ""
    if bModificar then
      sAccion = "ORAGENDAEDIT.ASP?VISION=" & SVISION & "&ACTIVRIESGO=" & RS(0) & "&CantRegMover=" & CantRegMoverx
    end if



    'Descripcion
    sInput = sImageFNL  & ofv.genform("FNEDIT" & SFORM,"AAA STYLE='DISPLAY:NONE;' ",sAccion,"","post","","","si")
    Celda = Celda & ofv.gencelda("","UT  title='EDITAR plan.' onclick='document.FNEDIT" & SFORM & ".submit();' style='cursor:pointer;'  ","CENTER","1%","","","",sInput,"si")





    'Descripcion
    sInput = rs(1) & ""
    Celda = Celda & ofv.gencelda("","UTI ","","30%","","","",sInput,"si")



    sfunc  = " onclick=""document.all.XTDX5.innerHTML = document.all.AGR" & posi & ".innerHTML;document.all.DX5.style.display='block';"" "
    sInput = sImageresp
    Celda = Celda & ofv.gencelda("","UT TITLE='Responsable del Plan.' " & sfunc & " ","CENTER","2%","","","",sInput,"si")

    sfunc  = " onclick=""document.all.XTDX5.innerHTML = document.all.AGP" & posi & ".innerHTML;document.all.DX5.style.display='block';"" "
    sInput = sImagePROC
    Celda = Celda & ofv.gencelda("","UT TITLE='Proceso-Actividad.' " & sfunc & " ","CENTER","2%","","","",sInput,"si")

    sInput = sImagePROY
    sfunc  = " onclick='document.AGPRY" & posi & ".submit();' "
  '''  Celda = Celda & ofv.gencelda("","UT TITLE='Proyecto asociado.' " & sfunc & " ","CENTER","2%","","","",sInput,"si")

    sInput = "&nbsp;"
    Celda = Celda & ofv.gencelda("","UT ","CENTER","2%","","","",sInput,"si")

    sInput = "&nbsp;"
    Celda = Celda & ofv.gencelda("","UT ","CENTER","2%","","","",sInput,"si")

    'FECHA CUMPLIMIENTO



    if clng(rs(2)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSh(rs(2))
    end if

    sInput = dFechan

    sInput = sinput

    Celda = Celda & ofv.gencelda("","VTI " ,"","10%","","","",sInput,"si")

    sInput = sImagevN

    DHOY = CLNG(DATE()) - clng(cdate("01-01-1900"))
    DFEC = clng(rs(2))

    if DHOY > DFEC then
       sInput = sImageva
    ELSEIF  DHOY = DFEC THEN
       sInput = sImagevM
    ELSE
       sInput = sImagevB
    END IF

    Celda = Celda & ofv.gencelda("","UT","CENTER","5%","","","",sInput,"si")


    'estado

    DFEC = rs(5) & ""

    if DFEC = "0" then
       sInput = "CANCELADO"
    ELSEIF DFEC = "1" then
       sInput = "PLANIFICADO"
    ELSEIF  DFEC = "2" then
       sInput = "EN CURSO"
    ELSEIF  DFEC = "5" then
       sInput = "FINALIZADO"
    ELSE
       sInput = "REPLANIFICADO"
    END IF


    Celda = Celda & ofv.gencelda("","UTI","","5%","","","",sInput,"si")

'''''''''''''''''''''
'' ventana de trabajo - proyectos
'''''''''''''''''''''''''''''''''

    sAccion = ""
    if bModificar then
      sAccion = "ORAGENDAPROY.asp?opcion=CP&AGENDA=" & rs(10) & ""
    end if

    Celdaf = Celdaf & ofv.genform("AGPRY" & POSI,"",sAccion,"_blank","post","","","si")


'''''''''''''''''''''
'' ventana de trabajo - procesos
'''''''''''''''''''''''''''''''''

  Strsql = "     select d.secuen,ISNULL(e.descri,D.DESCRIACC), g.descri,  a.* "
  Strsql = Strsql & "  from csrelpathdescri a"
  Strsql = Strsql & "     inner join objetos b on (a.cobj = b.codobjeto)"
  Strsql = Strsql & "     inner join objlnk c on (a.cobj = c.codobjeto)"
  Strsql = Strsql & "     inner join objaccion d on (c.codobjetol = d.codobjeto and c.iden4 = d.codproced)"
  Strsql = Strsql & "     LEFT join tipoaccalt e on (d.tipoaccion = e.tipoaccionalt)"
  Strsql = Strsql & "     left join oractivriesgo f on (d.codaccion = f.actividad )"
  Strsql = Strsql & "     left join orimpriesgo g on (f.itemriesgo = g.objid)"
  Strsql = Strsql & " where F.CODRIESGO = '" & rs(7) & "' "


'''  response.write strsql & "<br>"
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs1.eof then


      WKIPROCD = ""

       FOR WKI=3 TO 10

           WKIPROC = RS1(WKI) & ""

           IF WKIPROC <> "" THEN
              WKIPROCD = WKIPROCD & RS1(WKI) & " / "
           END IF


       NEXT

       IF LEN(WKIPROCD) > 4 THEN
          WKIPROCDA = LEFT(WKIPROCD,LEN(WKIPROCD)-3)
       END IF

       Formulariop = ""

       Celdap = ""
       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","PROCESO&nbsp;","si")
       sInput = WKIPROCDA
       Celdap = Celdap & ofv.GenCelda("","UtI ","","","","","",sInput,"si")
       Formulariop = Formulariop & ofv.GenRow("","","","","25","","",Celdap,"si")

       Celdap = ""
       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","ACTIVIDAD&nbsp;","si")
       sInput = RS1(0) & "." & RS1(1)
       Celdap = Celdap & ofv.GenCelda("","UtI ","","","","","",sInput,"si")
       Formulariop = Formulariop & ofv.GenRow("","","","","25","","",Celdap,"si")

       Celdap = ""
       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","RIESGO&nbsp;","si")
       sInput = RS1(2)
       Celdap = Celdap & ofv.GenCelda("","UtI ","","","","","",sInput,"si")
       Formulariop = Formulariop & ofv.GenRow("","","","","25","","",Celdap,"si")


       Filan = ofv.GenTabla("","aaa RULES='ROWS' STYLE='border-style:solid;border-width:3;border-color:#F7CC88;' ","","95%","","0","","-1","1","","",Formulariop,"si")




       sInput = Filan

  else

     sInput = "error al buscar la actividad"

  end if



    Celda = Celda & ofv.gencelda("AGP" & POSI,"aaa style='display:none;' ","","","","","",sInput,"si")


    Strsql = "     select A.descripcion, B.DESCrI FROM USUARIOS A  "
    Strsql = Strsql & "     inner join ECO b on (a.SECTOR = b.ECO) "
    Strsql = Strsql & " WHERE A.ID = '" & RS(6) & "' "


'''  response.write strsql & "<br>"
    set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs1.eof then


       Formulariop = ""

       Celdap = ""
       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","RESPONSABLE&nbsp;","si")
       sInput = RS1(0)
       Celdap = Celdap & ofv.GenCelda("","UtI ","","","","","",sInput,"si")
       Formulariop = Formulariop & ofv.GenRow("","","","","25","","",Celdap,"si")

       Celdap = ""
       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","AREA&nbsp;","si")
       sInput = RS1(1)
       Celdap = Celdap & ofv.GenCelda("","UtI ","","","","","",sInput,"si")
       Formulariop = Formulariop & ofv.GenRow("","","","","25","","",Celdap,"si")

       Filan = ofv.GenTabla("","aaa RULES='ROWS' STYLE='border-style:solid;border-width:3;border-color:#F7CC88;' ","","95%","","0","","-1","1","","",Formulariop,"si")




       sInput = Filan

    else

     sInput = "error al buscar EL RESPONSABLE"

    end if



    Celda = Celda & ofv.gencelda("AGR" & POSI,"aaa style='display:none;' ","","","","","",sInput,"si")



    Formulario = Formulario & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","25","","",Celda,"si")

       if xTrColor = "#FFFFFF" then
          xTrColor = "#F0F0F0"
       ELSE
          xTrColor = "#FFFFFF"
       END IF

    posi = posi + 1
    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","","","98% ","","0","","0","0","","",Formulario,"si") &  Celdaf



''''''''''''''''''''''''''''''''''''''''''''''''
'' ventana de trabajo
'''''''''''''''''''''''''''''''''''''''''''''''

       FilaN = ""

          sImage = "<font face='corflaticon' size=1 color='#FFFFFF' style='border-style:OUTSET;border-width:1;' >&#0114;</font>"

          wkproccod2 = "<a class=nt onclick=""document.all.DX5.style.display='none';"" href='javascript:void(0);' "
          wkproccod2 = wkproccod2 & " title='Cerrar Ventana.' >"
          wkproccod2 = wkproccod2 & sImage
          wkproccod2 = wkproccod2 & "</a>"

       Formulario = ""

       sInput = wkproccod2
       Celda = ""
       Celda = Celda & ofv.GenCelda("","UTI" & wkstylo1,"right","","","","",sInput,"si")
       FilaN = ofv.GenRow("","","","","15","","",Celda,"si")
       Formulario = Formulario & FilaN

       sInput = "&nbsp;"
       Celda = ""
       Celda = Celda & ofv.GenCelda("XTDX5","UTI","","","","","",sInput,"si")
       FilaN = ofv.GenRow("","","","","","","",Celda,"si")
       Formulario = Formulario & FilaN

       Filan = ofv.GenTabla("","aaa STYLE='border-style:solid;border-width:3;border-color:#F7CC88;' ","","95%","","0","","-1","1","","",Formulario,"si")

       Filan = "<CENTER><BR>" & Filan & "</CENTER>"

       Celda = ""
       Celda = Celda & "<div id='DX5' style='z-index:1;margin-top:1;margin-left:1;position:absolute;left:400px;top:140px;display:none;"
       Celda = Celda & "width:450;height:250;overflow:auto;border-style:SOLID;border-width:1;Border-color:#888888;"
       Celda = Celda & "background-color:#F9F9F9;"
       Celda = Celda & "' >" & Chr(13) & Filan & "</DIV>"

       sHTML = sHTML & Celda



  'Botonera
  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)


  sHTML = sHTML & GTCCALENDAR("M")


  elseif svision = "IAG" then




       wkstylo1 = " style='background-color:#F7CC88;color:#ffffff;' "
       FilaN = ""

  Strsql = "     select d.secuen,ISNULL(e.descri,D.DESCRIACC), g.descri, F.CODRIESGO, d.codaccion, f.objid,f.periodo, a.* "
  Strsql = Strsql & "  from csrelpathdescri a"
  Strsql = Strsql & "     inner join objetos b on (a.cobj = b.codobjeto)"
  Strsql = Strsql & "     inner join objlnk c on (a.cobj = c.codobjeto)"
  Strsql = Strsql & "     inner join objaccion d on (c.codobjetol = d.codobjeto and c.iden4 = d.codproced)"
  Strsql = Strsql & "     LEFT join tipoaccalt e on (d.tipoaccion = e.tipoaccionalt)"
  Strsql = Strsql & "     left join oractivriesgo f on (d.codaccion = f.actividad )"
  Strsql = Strsql & "     left join orimpriesgo g on (f.itemriesgo = g.objid)"
  Strsql = Strsql & " where F.CODRIESGO = '" & sACTIVRIESGO & "' "


'''  response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     sForm = "FORMIAG"

        Strsql = "select * from oragenda "
        set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)


       IF SFECHA  = "" THEN
          SFECHA  = CLNG(DATE()) - clng(cdate("01-01-1900"))
       end if

       sdescri = rs(0) & "." & rs(1) & " - " & rs(2)


        nCol = 70



        t3b = "<font class=at><textarea class=dt "
        t3b = t3b & "name=" & rs1(1).name & " cols=" & nCol & " rows=2 "
        t3b = t3b & " onkeypress=" & chr(34)
        t3b = t3b & " if (this.value.length > 195) {alert('El texto es demasiado"
        t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
        t3b    = t3b & sdescri & "</textarea></font>"


   sinput =  t3b   ' descri
   sinput = sinput & ofv.GenerarInput(rs1(2).name,SFECHA,"HIDDEN","","1","dt","") ' feccumpli
   sinput = sinput & ofv.GenerarInput(rs1(3).name,"0","HIDDEN","","1","dt","")     ' fecrevision
   sinput = sinput & ofv.GenerarInput(rs1(4).name,"NO","HIDDEN","","1","dt","")     ' proyrel
   sinput = sinput & ofv.GenerarInput(rs1(5).name,"1","HIDDEN","","1","dt","")     ' estado
   sinput = sinput & ofv.GenerarInput(rs1(6).name,session("usuario"),"HIDDEN","","1","dt","")     ' responsable
   sinput = sinput & ofv.GenerarInput(rs1(7).name,sACTIVRIESGO,"HIDDEN","","1","dt","")     ' activriesgo
   sinput = sinput & ofv.GenerarInput(rs1(8).name,rs(4),"HIDDEN","","1","dt","")     ' actividad
   sinput = sinput & ofv.GenerarInput(rs1(9).name,rs(6),"HIDDEN","","1","dt","")     ' periodo
   sinput = sinput & ofv.GenerarInput(rs1(10).name,gtkey(rs(5)),"HIDDEN","","1","dt","")     ' codagenda
   sinput = sinput & ofv.GenerarInput(rs1(11).name,"0","HIDDEN","","1","dt","")     ' fecreal
   sinputd = sinput

      WKIPROCD = ""

       FOR WKI=7 TO 14

           WKIPROC = RS(WKI) & ""

           IF WKIPROC <> "" THEN
              WKIPROCD = WKIPROCD & RS(WKI) & " / "
           END IF


       NEXT

    IF LEN(WKIPROCD) > 4 THEN
       WKIPROCD = LEFT(WKIPROCD,LEN(WKIPROCD)-3)
       WKIPROCDA = "<A HREF='javascript:void(0);' onclick='alert(this.innerHTML);' "
       WKIPROCDA = WKIPROCDA & " STYLE='COLOR:#888888;FONT-SIZE:7PT;FONT-FAMILY:""TIMES NEW ROMAN"";' >" & WKIPROCD & "</A>"
    END IF


       Celda = ""
       Celda = Celda & ofv.GenCelda("","VTI COLSPAN=3 " & wkstylo1,"","","","","","PLANIFICAR MITIGACION EN AGENDA","si")
       FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")


       Celda = ""
       Celda = Celda & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","PROCESO&nbsp;","si")
       sInput = WKIPROCDA
       Celda = Celda & ofv.GenCelda("","UtI colspan=2 ","","","","","",sInput,"si")
       FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")


       Celda = ""
       Celda = Celda & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","ACTIVIDAD&nbsp;","si")
       sInput = rs(0) & "." & rs(1)
       Celda = Celda & ofv.GenCelda("","UtI colspan=2 ","","","","","",sInput,"si")
       FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")

       Celda = ""
       Celda = Celda & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","RIESGO&nbsp;","si")
       sInput = rs(2)
       Celda = Celda & ofv.GenCelda("","vtI colspan=2 ","","","","","",sInput,"si")
       FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")

       Celda = ""
       Celda = Celda & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","CUMPLIMIENTO&nbsp;","si")
       sInput = GTFECHADOR(SFECHA)

       wkproccod2 = sInput

       Celda = Celda & ofv.GenCelda("CELLFEC","vtI colspan=2 style='cursor:pointer;color:#993333;text-decoration:underline;' title='CAMBIAR FECHA DE CUMPLIMIENTO.' "  & GTCSETCALR(sForm,rs1(2).name,"N","1","CELLFEC","",0,2,0,0,0) ,"","","","","",wkproccod2,"si")
       FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")


       Celda = ""
       Celda = Celda & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","DESCRIPCION&nbsp;","si")
       sInput = sinputd
       Celda = Celda & ofv.GenCelda("","UtI colspan=2 ","","","","","",sInput,"si")
       FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")




       Celda = ""
       Celda = Celda & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","ESTADO&nbsp;","si")
       sInput = "PLANIFICADO"
       Celda = Celda & ofv.GenCelda("","UtI colspan=2 ","","","","","",sInput,"si")
       FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")

       Celda = ""
       Celda = Celda & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","RESPONSABLE&nbsp;","si")
       sInput = SESSION("SEGUSUARIO")
       Celda = Celda & ofv.GenCelda("","UtI colspan=2 ","","","","","",sInput,"si")
       FilaN = filaN & ofv.GenRow("","","","","20","","",Celda,"si")

       sInput = "<INPUT TYPE='SUBMIT' CLASS=BT VALUE='AGENDAR' > "
       Celda = ""
       Celda = Celda & ofv.GenCelda("","UTI COLSPAN=3 " & wkstylo1,"RIGHT","","","","",sInput,"si")
       FilaN = filaN & ofv.GenRow("","","","","25","","",Celda,"si")

'''''''''''''''''

       sAccion = "ORAGENDA.asp?VISION=AAG&VISIONMAIN=" & SVISIONMAIN
       Formulario = ofv.genform(sForm,"",sAccion,"","post","",Filan,"si")

      else

       Celda = ""
       Celda = Celda & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","OBSERVACIONES&nbsp;","si")
       sInput = "NO EXISTE LA ACTIVIDAD-RIESGO PARA PLANIFICAR"
       Celda = Celda & ofv.GenCelda("","vt colspan=2 ","","","","","",sInput,"si")
       FilaN = filaN & ofv.GenRow("","","","","20","","",Celda,"si")

      end if


       Filan = ofv.GenTabla("","aaa RULES='ROWS' STYLE='border-style:solid;border-width:3;border-color:#F7CC88;' ","","80%","","0","","-1","1","","",Formulario,"si")

       sHTML = sHTML & "<CENTER><BR><BR>" & Filan & "</CENTER>"


      'Botonera
       Agregar = ""
       Buscar  = ""
       Celda   = ""

      sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)

      sHTML = sHTML & GTCCALENDAR("M")

      call ofv.cerrarconsulta(rs)

  else



  end if


  call ofv.cerrarconn(cn)


    sHTML = sHTML & "<script type='text/javascript'>"
    sHTML = sHTML & "function disp_confirm(disp_msg,disp_frm)"
    sHTML = sHTML & "  {"
    sHTML = sHTML & "  var r=confirm(disp_msg);"
    sHTML = sHTML & "  if (r==true)"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "  var fr= document.getElementById(disp_frm);"
  '''  sHTML = sHTML & "    alert(fr.name);"
    sHTML = sHTML & "    fr.submit();"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  else"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "    alert('Se cancelo la accion!');"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  }"

''    sHTML = sHTML & "  }"
    sHTML = sHTML & "</script>"


  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML



Function FechaHAB(wkfecha)


      swd = datepart("w",cdate(wkfecha + clng(cdate("01-01-1900"))))

      IF SWD > 6 THEN
         wkfecha = wkfecha + 2
      END IF

      IF swd = 1 THEN
         wkfecha = wkfecha + 1
      END IF


      FechaHAB  = wkfecha

End Function




%>