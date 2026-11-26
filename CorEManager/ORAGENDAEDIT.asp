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

  svision       = cstr(request.querystring("vision")&"")
  sOPCION       = request.querystring("OPCION")
  sACTIVRIESGO  = request.querystring("ACTIVRIESGO")

    ofv.ObtenerAtributos Session("Form"),""
    bAgregar = ofv.AAgregar
    bModificar = ofv.AModificar
    bSuprimir = ofv.AEliminar
    bConsultar = ofv.AConsultar

    set cn = ofv.conectar(ofv.strconn2)

'  response.write svision
  if svision="AGNALL" or svision="PORMATIMPACTO" THEN
   SREAD=" READONLY"
  ELSE
   SREAD=""
  END IF

'  IF svision="PORMATIMPACTO" THEN
'     ScodRiesgo= request.querystring("ScodRiesgo")
 '    SActividad= request.querystring("SActividad")
 '    SPERIODO=request.querystring("SPERIODO")
 '    strsql= "select * from ORAGENDA  where  actividad='" & SActividad  & "' and activriesgo='" & ScodRiesgo &"' AND PERIODO=0" & SPERIODO
   '  response.write strsql
 '    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  '    sACTIVRIESGO=0
  '   if not rs.eof  then sACTIVRIESGO=rs(0)
  '   call ofv.cerrarconsulta(rs)
  'END IF


  SREAD=cstr(SREAD)
  vuelta        = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_ACTIVRIESGO=" & sACTIVRIESGO & "_VISION=" & SVISION

  sHTML = ofv.FormHeader(session("FormN"))

  Celda = ""


   Celda = Celda & ofv.GenCelda("","uti  ","","90%","","","","SEGUIMIENTO PLAN DE MITIGACION","si")
   Fila  = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML & ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ","","100% ","30","0","","-1","1","","",Fila,"si")





  DIM TABRANRIESGOS(10,5)


  Strsql = "Select * From ORRANGORIESGO ORDER BY VALOR DESC"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     POSI = 0
     DO WHILE NOT RS.EOF
        POSI = POSI + 1

          TABRANRIESGOS(POSI,0) = RS(0) 'CODIGO
          TABRANRIESGOS(POSI,1) = RS(1) 'VALOR
          TABRANRIESGOS(POSI,2) = RS(2) 'DESCRIPCION
          TABRANRIESGOS(POSI,3) = RS(3) 'COLOR

          TTABRANRIESGO = POSI

        RS.MOVENEXT
     LOOP



  END IF

    DHOY = CLNG(DATE()) - clng(cdate("01-01-1900"))
    dFechaHOY = GTFechadorSh(DHOY)

'''''''''''''''''''''''''''''''''''''''
'' POR UAG - FINALIZO ENTRADAS EN LA AGENDA

  if sOPCION = "UAG" then


        strsql = "UPDATE ORAGENDA SET ESTADO = '3' WHERE OBJID = '0" & sACTIVRIESGO & "' "
''  response.write strsql & "<br>" & chr(13)
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

      '  xxxr=GENLOGEXTDATA("UPDATE","objaccion","objid",ncodigo,"RO")
  end if


'''''''''''''''''''''''''''''''''''''''
'' POR CAN - CANCELO PLAN

  if sOPCION = "CAN" then


        strsql = "UPDATE ORAGENDA SET ESTADO = '0', FECFINAL = '" & DHOY & "' WHERE OBJID = '0" & sACTIVRIESGO & "' "
''  response.write strsql & "<br>" & chr(13)
        call ofv.crearconsultaEx(StrSql,cn,1,volver)


  end if

'''''''''''''''''''''''''''''''''''''''
'' POR INI - INICIO PLAN

  if sOPCION = "INI" then


        strsql = "UPDATE ORAGENDA SET ESTADO = '2', FECREAL = '" & DHOY & "' WHERE OBJID = '0" & sACTIVRIESGO & "' "
''  response.write strsql & "<br>" & chr(13)
        call ofv.crearconsultaEx(StrSql,cn,1,volver)


  end if

'''''''''''''''''''''''''''''''''''''''
'' POR FIN - FINALIZO PLAN

  if sOPCION = "FIN" then


        strsql = "UPDATE ORAGENDA SET ESTADO = '5', FECFINAL = '" & DHOY & "' WHERE OBJID = '0" & sACTIVRIESGO & "' "
''  response.write strsql & "<br>" & chr(13)
        call ofv.crearconsultaEx(StrSql,cn,1,volver)


  end if


'''''''''''''''''''''''''''''''''''''''
'' POR REP - REPLANIFICO PLAN

  if sOPCION = "REP" then


        strsql = "UPDATE ORAGENDA SET ESTADO = '4', FECREPLAN = FECCUMPLI, OBSERVACIONES = OBSERVACIONES + ' - REPLANIFICADO EL " & dFechaHOY & "' WHERE OBJID = '0" & sACTIVRIESGO & "' "
''  response.write strsql & "<br>" & chr(13)
        call ofv.crearconsultaEx(StrSql,cn,1,volver)


  end if


'''''''''''''''''''''''''''''''''''''''




    xTrColor = "#FFFFFF"

    wkstylo1 = " style='background-color:#F7CC88;color:#ffffff;' "
    wkstylo2 = " style='background-color:#336699;color:#ffffff;' "

    wkstyloD = " style='border-style:SOLID;border-width:1;border-color:#888888;' "

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




  WKSETTABLA = "ORAGENDA"

  Strsql = "Select * From " & WKSETTABLA & " WHERE objid = 0" & sACTIVRIESGO
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then


     IF CLNG("0" & RS(14)) > 0 AND CLNG("0" & RS(15)) > 0 THEN


        strsql = "UPDATE ORAGENDA SET COSTOTOTAL = '" & (CLNG("0" & RS(14)) * CLNG("0" & RS(15)) )& "' WHERE objid = 0" & sACTIVRIESGO
''  response.write strsql & "<br>" & chr(13)
        call ofv.crearconsultaEx(StrSql,cn,1,volver)



        Strsql = "Select * From " & WKSETTABLA & " WHERE objid = 0" & sACTIVRIESGO
        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     END IF

  posI = 0

    sForm = "form" & posi
      Celda = ""
    sAccion = ""
    if bModificar then
      sAccion = "ORAGENDAEX.ASP?VISION=UAG&VISIONMAIN=" & SVISION & "&ACTIVRIESGO=" & RS(0)
    end if
   'PLAN

       Celda = ""
       Celda = Celda & ofv.GenCelda("","UTI " & wkstylo1,"","15%","","","","PLAN&nbsp;","si")

    sInput = rs(0) & ""


    Celda = Celda & ofv.gencelda("","UTI  " & wkstyloD,"","35%","","","",sInput,"si")
    'Descripcion
    sInput = rs(1) & ""


       Celda = Celda & ofv.GenCelda("","UTI " & wkstylo1,"","15%","","","","DESCRIPCION&nbsp;","si")


      nCol = 60

    ''  if  len(wdescri) = 0 then wdescri = "Sin Descripcion"

          wdescri = ""
      if  not  isnull(rs(1)) then
          wdescri = rs(1)
      end if

      t3b = "<font class=at><textarea class=dt "
      t3b = t3b & "name=" & rs(1).name & " cols=" & nCol & " rows=3 "
      if  bModificar AND SREAD="" then
          t3b = t3b & " onblur=" & sForm & ".submit() "
          t3b = t3b & " onkeypress=" & chr(34)
          t3b = t3b & " if (this.value.length > 150) {alert('El texto es demasiado"
          t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
      else
          t3b = t3b & " readonly >"
      end if
      xdatos =  wdescri
      t3b    = t3b & xdatos & "</textarea></font>"
      SINPUT = t3b


       Celda = Celda & ofv.GenCelda("","UtI " & wkstyloD,"","35%","","","",sInput,"si")
       Formulariop = Formulariop & ofv.GenRow("","","","","25","","",Celda,"si")


   'estado

       Celda = ""
       Celda = Celda & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","ESTADO&nbsp;","si")

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

       sInput = "<div STYLE='background-color:#888888;height:100%;width:150px;' ><center><FONT color='#ffffff' style='position:relative;top:+3;' >&nbsp;" & SINPUT & "&nbsp;</FONT></center><div>"


    Celda = Celda & ofv.gencelda("","UTI " & wkstyloD,"","","","","",sInput,"si")




    'FECHA CUMPLIMIENTO


       Celda = Celda & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","FECHA CUMPLIMIENTO&nbsp;","si")


    sInput2 = sImagevN

    DHOY = CLNG(DATE()) - clng(cdate("01-01-1900"))
    DFEC2 = clng(rs(2))

    if DHOY > DFEC2 then
       sInput2 = sImageva
    ELSEIF  DHOY = DFEC2 THEN
       sInput2 = sImagevM
    ELSE
       sInput2 = sImagevB
    END IF


 '     sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
 '     sAccion = sAccion & "&tabla=ORAGENDA&volver=" & vuelta

      sInputD = ofv.GenerarInput(rs(2).name,rs(2).value,"HIDDEN",sForm,"1","dt","")

    if clng(rs(2)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSh(rs(2))
    end if

    sInput = dFechan & "&nbsp;" & sInput2 & sInputD

    sInput = "<u>" & sinput & "</u>"
 '   SFORMCU = ofv.genform("FCU" & sForm,"",sAccion,"","post","",sInputD,"si")

    IF SREAD="" THEN
		if (rs("estado").value = "5" or rs("estado").value = "0") then
			Celda = Celda & ofv.gencelda("","VTI ","","","","","",sInput,"si")
		else
			Celda = Celda & ofv.gencelda("","VTI STYLE='cursor:pointer;' " & GTCSETCALR(sForm,rs(2).name,GTCSMT,GTCFCH,GTCFCHDAT,"",0,2,0,0,0)  & wkstyloD,"","","","","",sInput,"si")
		end if
    ELSE
     Celda = Celda & ofv.gencelda("","VTI ","","","","","",sInput,"si")
    END IF

       Formulariop = Formulariop & ofv.GenRow("","","","","25","","",Celda,"si")


    if clng("0" & rs(12)) = 0 then
      dFechan = "-------"
    else


      'FECHA ORIGINAL

         Celda = ""
         Celda = Celda & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","FECHA ORIGINAL&nbsp;","si")

         dFechan = GTFechadorSh(rs(12))
         sInput = dFechan

         Celda = Celda & ofv.gencelda("","VTI COLSPAN=3 "  & wkstyloD ,"","","","","",sInput,"si")

         Formulariop = Formulariop & ofv.GenRow("","","","","25","","",Celda,"si")

     end if


    if clng("0" & rs(11)) = 0 then
      dFechan = "-------"
    else


      'FECHA INICIO

         Celda = ""
         Celda = Celda & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","FECHA INICIO&nbsp;","si")

         dFechan = GTFechadorSh(rs(11))
         sInput = dFechan

         Celda = Celda & ofv.gencelda("","VTI COLSPAN=3 "  & wkstyloD ,"","","","","",sInput,"si")

         Formulariop = Formulariop & ofv.GenRow("","","","","25","","",Celda,"si")

     end if


    if clng("0" & rs(17)) = 0 then
      dFechan = "-------"
    else


      'FECHA FINALIZACION

         Celda = ""
         Celda = Celda & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","FECHA FINALIZACION&nbsp;","si")

         dFechan = GTFechadorSh(rs(17))
         sInput = dFechan

         Celda = Celda & ofv.gencelda("","VTI COLSPAN=3 "  & wkstyloD,"","","","","",sInput,"si")

         Formulariop = Formulariop & ofv.GenRow("","","","","25","","",Celda,"si")

     end if


'''''''''''''''''''''
'' ventana de trabajo - procesos
'''''''''''''''''''''''''''''''''

       wspvalor = 0
       WSPTABLA = "&nbsp;"

  Strsql = "     select d.secuen,e.descri,g.descri, f.probabilidad,f.impacto,"
  Strsql = Strsql & "  f.riesgobruto,f.riesgoresidual,f.riesgoresidualNV,F.CODRIESGO, F.PERIODO,"
  Strsql = Strsql & "   a.* "
  Strsql = Strsql & "  from csrelpathdescri a"
  Strsql = Strsql & "     inner join objetos b on (a.cobj = b.codobjeto)"
  Strsql = Strsql & "     inner join objlnk c on (a.cobj = c.codobjeto)"
  Strsql = Strsql & "     inner join objaccion d on (c.codobjetol = d.codobjeto and c.iden4 = d.codproced)"
  Strsql = Strsql & "     left join tipoaccalt e on (d.tipoaccion = e.tipoaccionalt)"
  Strsql = Strsql & "     left join oractivriesgo f on (d.codaccion = f.actividad )"
  Strsql = Strsql & "     left join orimpriesgo g on (f.itemriesgo = g.objid)"
  Strsql = Strsql & " where F.CODRIESGO = '" & rs(7) & "' "


' response.write strsql & "<br>"
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs1.eof then


      WKIPROCD = ""

       FOR WKI=10 TO 17

           WKIPROC = RS1(WKI) & ""

           IF WKIPROC <> "" THEN
              WKIPROCD = WKIPROCD & RS1(WKI) & " / "
           END IF


       NEXT

       IF LEN(WKIPROCD) > 4 THEN
          WKIPROCDA = LEFT(WKIPROCD,LEN(WKIPROCD)-3)
       END IF



       Celdap = ""
       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","PROCESO&nbsp;","si")
       sInput = WKIPROCDA
       Celdap = Celdap & ofv.GenCelda("","UtI "  & wkstyloD,"","","","","",sInput,"si")

       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","ACTIVIDAD&nbsp;","si")
       sInput = RS1(0) & "." & RS1(1)
       Celdap = Celdap & ofv.GenCelda("","UtI " & wkstyloD,"","","","","",sInput,"si")
       Formulariop = Formulariop & ofv.GenRow("","","","","25","","",Celdap,"si")

       Celdap = ""
       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","RIESGO&nbsp;","si")
       sInput = RS1(2)
       Celdap = Celdap & ofv.GenCelda("","UtI " & wkstyloD,"","","","","",sInput,"si")

       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","PERIODO&nbsp;","si")

       Speriododsc = RS1(9)

        Strsql = "Select * From ORPERIODOS  where OBJID = 0" & RS1(9)
        set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if NOT rsX1.eof then
           Speriododsc              = RSX1(1)

        end if

       sInput = Speriododsc
       Celdap = Celdap & ofv.GenCelda("","UtI " & wkstyloD,"","","","","",sInput,"si")
       Formulariop = Formulariop & ofv.GenRow("","","","","25","","",Celdap,"si")


             wspporc = 100

             Strsql  = "Select VALOR,porcentaje,obs1,DESCRI From ORNIVPROBOCURRENCIA where OBJID = 0" & RS1(3)
             set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if NOT rsX1.eof then
                wspDATPX      = "0" & RSX1(0)
                wspDATP       = CINT(wspDATPX)
                wspporc       = CINT("0" & RSX1(1))
                wspDATprodsc  = RSX1(3) & "&nbsp-&nbsp;" & RSX1(1) & "&nbsp;%&nbsp-&nbsp;" & RSX1(2)
             END IF

             wspmonto = 0

             Strsql  = "Select VALOR,impdesde,imphasta,obs1,DESCRI From ORNIVIMPACTO where OBJID = 0" & RS1(4)
             set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if NOT rsX1.eof then
                wspDATIX      = "0" & RSX1(0)
                wspDATI       = CINT(wspDATIX)
                wspmonto      = clng((Clng("0" & RSX1(1)) + Clng("0" & RSX1(2))) / 2)
                wspDATimpdsc  = RSX1(4) & "&nbsp-&nbsp;$&nbsp;" & wspmonto & "&nbsp;-&nbsp;" & RSX1(3)
             END IF

             wspmontobruto = clng(wspmonto * wspporc / 100)
             wspDATA       = wspDATP * wspDATI
             wspDATB       = CDBL("0" & RS1(6))
             wspDATBB      = CDBL("0" & RS1(7))
             wspDATC       = CINT(wspDATA / wspDATB)
             wspDATCN      = CINT(wspDATA / wspDATBB)

             Strsql  = "Select descri From ORNIVRIESGOBRUTO where NIVPROB =  0" & RS1(3) & " AND NIVIMPA = 0" & RS1(4) & " "
 '''  response.write strsql & "<br>" & chr(13)
             set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if NOT rsX1.eof then

                FOR II=1 TO TTABRANRIESGO

                    IF RSX1(0)  =  TABRANRIESGOS(II,0) THEN
                       wspDATAdsc    = TABRANRIESGOS(II,2)   'DESCRIPCION
                       wspDATAdsccl    = TABRANRIESGOS(II,3) 'COLOR
                       EXIT FOR
                    END IF

                NEXT


             END IF

             wspDATAdsc    = wspDATAdsc & "&nbsp-&nbsp;$&nbsp;" & wspmontobruto

             Strsql  = "Select descri,REQPLAN From ORNIVriesgo where "
             Strsql  = Strsql & " NOT (0" & replace(cstr(wspDATB & ""),",",".") & " > VALORHASTA) AND "
             Strsql  = Strsql & " NOT (0" & replace(cstr(wspDATB & ""),",",".") & " < VALORDESDE) "
 ''  response.write strsql & "<br>" & chr(13)
             set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if NOT rsX1.eof then

                FOR II=1 TO TTABRANRIESGO

                    IF RSX1(0)  =  TABRANRIESGOS(II,0) THEN
 ''  response.write RS1(0) & "<br>" & chr(13)
                       wspDATArrdsc    = TABRANRIESGOS(II,2)   'DESCRIPCION
                       wspDATArrdsccl    = TABRANRIESGOS(II,3) 'COLOR
                       EXIT FOR
                    END IF

                NEXT


             END IF

             Strsql  = "Select descri,REQPLAN From ORNIVriesgo where "
             Strsql  = Strsql & " NOT (0" & replace(cstr(wspDATBB & ""),",",".") & " > VALORHASTA) AND "
             Strsql  = Strsql & " NOT (0" & replace(cstr(wspDATBB & ""),",",".") & " < VALORDESDE) "
 ''  response.write strsql & "<br>" & chr(13)
             set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if NOT rsX1.eof then

                FOR II=1 TO TTABRANRIESGO

                    IF RSX1(0)  =  TABRANRIESGOS(II,0) THEN
 ''  response.write RS1(0) & "<br>" & chr(13)
                       wspDATArNdsc    = TABRANRIESGOS(II,2)   'DESCRIPCION
                       wspDATArNdsccl    = TABRANRIESGOS(II,3) 'COLOR
                       EXIT FOR
                    END IF

                NEXT


             END IF

             wspvalor      = wspmontobruto / wspDATC
             wspvalorN     = wspmontobruto / wspDATCN

             wspvalor      = ROUND(wspvalor,2)
             wspvalorN     = ROUND(wspvalorN,2)

             wspDATArrdsc    = wspDATArrdsc & "&nbsp-&nbsp;$&nbsp;" & wspvalor
             wspDATArNdsc    = wspDATArNdsc & "&nbsp-&nbsp;$&nbsp;" & wspvalorN




          Strsql = "Select DESCRI From ORNIVEFICIENCIA WHERE VALOR =0" & wspDATC
 '''  response.write strsql & "<br>" & chr(13)
             set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if NOT rsX1.eof then

                wspDATMdsc    = RSX1(0)   'DESCRIPCION

             END IF


          Strsql = "Select DESCRI From ORNIVEFICIENCIA WHERE VALOR =0" & wspDATCN
 '''  response.write strsql & "<br>" & chr(13)
             set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if NOT rsX1.eof then

                wspDATMNdsc    = RSX1(0)   'DESCRIPCION

             END IF


             Strsql = "Select A.* From oractivrcontroles A "
             strsql = strsql & "  where  "
             strsql = strsql & "  A.CODRIESGO =  '" & RS1(8) & "' "
             strsql = strsql & "  AND A.TIPO = '1' "

			'	RESPONSE.WRITE(CHR(13) & CHR(13))
			'	RESPONSE.WRITE(STRSQL)
			'	RESPONSE.WRITE(CHR(13) & CHR(13))
             set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if not rsX1.eof then

                FILACONT = ""

                   Celda  = ""
                   wspDATCDSC1    = "CONTROL"
                   Celda  = Celda & ofv.GenCelda("","UT BGCOLOR='#CCCCCC' " & wkstyloD,"","60%","","","",wspDATCDSC1,"si")

                   wspDATCDSC2    = "TIPO"
                   Celda  = Celda & ofv.GenCelda("","UT BGCOLOR='#CCCCCC' " & wkstyloD,"","20%","","","",wspDATCDSC2,"si")

                   wspDATCDSC3    = "CALIFICACION"
                   Celda  = Celda & ofv.GenCelda("","UT BGCOLOR='#CCCCCC' " & wkstyloD,"","20%","","","",wspDATCDSC3,"si")

                   FILACONT   = FILACONT & ofv.GenRow("","","","","20","","",Celda,"si")


                do until rsX1.eof
                   Celda  = ""

                   Strsql = "Select DESCRI From ORCONTROLRIESGO WHERE OBJID = 0" & RSX1(7)

                   set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                   if NOT rs2.eof then
                      wspDATCDSC1    = RS2(0)
                   ELSE
                      wspDATCDSC1    = "SIN DEFINIR"
                   END IF
                   Celda  = Celda & ofv.GenCelda("","UT " & wkstyloD,"","","","","",wspDATCDSC1,"si")

                   'if RSX1(11) = "1" then
                   '   wspDATCDSC2    = "AUTOMATICO"
                   'ELSEIf RSX1(11) = "2" then
                   '   wspDATCDSC2    = "SEMIAUTOMATICO"
                   'ELSEIf RSX1(11) = "3" then
                   '   wspDATCDSC2    = "MANUAL"
                   'ELSE
                   '   wspDATCDSC2    = "SIN DEFINIR"
                   'END IF


					STRSQL = "Select DESCRI From ORTIPOCONTROLES WHERE OBJID = 0" & RSX1("modocont")
					'RESPONSE.WRITE(CHR(13) & CHR(13))
					'RESPONSE.WRITE(STRSQL)
					'RESPONSE.WRITE(CHR(13) & CHR(13))
					set RS2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
					IF NOT RS2.EOF THEN
						wspDATCDSC2 = RS2("DESCRI")
					ELSE
						wspDATCDSC2 = "SIN DEFINIR"
					END IF





                   Celda  = Celda & ofv.GenCelda("","UT " & wkstyloD,"","","","","",wspDATCDSC2,"si")


                   Strsql = "Select DESCRI From ORNIVEFICIENCIA WHERE OBJID = 0" & RSX1(8)
                   set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                   if NOT rs2.eof then
                      wspDATCDSC3    = RS2(0)
                   ELSE
                      wspDATCDSC3    = "----"
                   END IF
                   Celda  = Celda & ofv.GenCelda("","UT " & wkstyloD,"","","","","",wspDATCDSC3,"si")


                   FILACONT   = FILACONT & ofv.GenRow("","","","","20","","",Celda,"si")

                   IF (RSX1(12) & "") <> "" THEN

                      Celda      = ofv.GenCelda("","UT COLSPAN=3 " & wkstyloD,"","","","","",RSX1(12) & "&nbsp;","si")
                      FILACONT   = FILACONT & ofv.GenRow("","","","","20","","",Celda,"si")

                   END IF

                      Celda      = ofv.GenCelda("","UT COLSPAN=3 BGCOLOR='#DDDDDD' ","","","","","","&nbsp;","si")
                      FILACONT   = FILACONT & ofv.GenRow("","","","","5","","",Celda,"si")

                   rsX1.movenext
                loop

                WSPTABLA = ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;border-color:#888888;' BGCOLOR='#F0F0F0' ","","80%","","0","","0","0","","",FILACONT,"si")

             end if

       sInput = Filan

  else

     sInput = "error al buscar la actividad"

  end if



    Strsql = "     select A.descripcion, B.DESCrI FROM USUARIOS A  "
    Strsql = Strsql & "     inner join ECO b on (a.SECTOR = b.ECO) "
    Strsql = Strsql & " WHERE A.ID = '" & RS(6) & "' "


'''  response.write strsql & "<br>"
    set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if not rs1.eof then




       Celdap = ""
       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","RESPONSABLE&nbsp;","si")
       sInput = RS1(0)
       Celdap = Celdap & ofv.GenCelda("","UtI " & wkstyloD,"","","","","",sInput,"si")

       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","AREA&nbsp;","si")
       sInput = RS1(1)
       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstyloD,"","","","","",sInput,"si")
       Formulariop = Formulariop & ofv.GenRow("","","","","25","","",Celdap,"si")


       sInput = Filan

    else

     sInput = "error al buscar EL RESPONSABLE"

    end if



       Celdap = ""
       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","TOTAL HS.&nbsp;","si")
       IF SREAD="" THEN
        sInput = ofv.GenerarInput(rs(14).name,rs(14),"text",sForm,"15","dtN","") & "&nbsp;Hs."
       ELSE
        sInput=rs(14) & "&nbsp;Hs."
       END IF

       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstyloD,"","","","","",sInput,"si")

       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","COSTO $ x HS.&nbsp;","si")
       IF SREAD="" THEN
        sInput = ofv.GenerarInput(rs(15).name,rs(15),"text",sForm,"15","dtN","") & "&nbsp;$"
       ELSE
         sInput = rs(15) & "&nbsp;$"
       END IF
       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstyloD,"","","","","",sInput,"si")
       Formulariop = Formulariop & ofv.GenRow("","","","","25","","",Celdap,"si")



       Celdap = ""
       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstylo2,"","","","","","COSTO TOTAL&nbsp;","si")
       IF SREAD="" THEN
        sInput = ofv.GenerarInput(rs(16).name,rs(16),"text",sForm,"15","dtN","") & "&nbsp;$"
       ELSE
         sInput = rs(16) & "&nbsp;$"
       END IF

       Celdap = Celdap & ofv.GenCelda("","UtI style='border-style:SOLID;border-width:1;border-color:#000099;' ","","","","","",sInput,"si")

       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstylo2,"","","","","","VALOR RIESGO&nbsp;","si")
       IF SREAD="" THEN
        sInput = ofv.GenerarInput("ormnres",wspvalor,"readonly",sForm,"15","dtN","") & "&nbsp;$"
       ELSE
        sInput = wspvalor& "&nbsp;$"
       END IF

       Celdap = Celdap & ofv.GenCelda("","UtI style='border-style:SOLID;border-width:1;border-color:#990000;' ","","","","","",sInput,"si")
       Formulariop = Formulariop & ofv.GenRow("","","","","25","","",Celdap,"si")

       Celdap = ""
       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","PROBABILIDAD&nbsp;","si")
       sInput = wspDATprodsc
       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstyloD,"","","","","",sInput,"si")

       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","IMPACTO&nbsp;","si")
       sInput = wspDATimpdsc
       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstyloD,"","","","","",sInput,"si")
       Formulariop = Formulariop & ofv.GenRow("","","","","25","","",Celdap,"si")



       Celdap = ""
       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","RIESGO INHERENTE&nbsp;","si")
       sInput = wspDATAdsc
       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstyloD,"","","","","",sInput,"si")

       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","CONTROLES&nbsp;","si")
       sInput = wspDATMdsc
       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstyloD,"","","","","",sInput,"si")
       Formulariop = Formulariop & ofv.GenRow("","","","","25","","",Celdap,"si")





       Celdap = ""
       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","RIESGO RESIDUAL&nbsp;","si")
       sInput = wspDATArrdsc
       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstyloD,"","","","","",sInput,"si")

       Celdap = Celdap & ofv.GenCelda("","UTI ","","","","",""," ","si")
       sInput =  ""
       Celdap = Celdap & ofv.GenCelda("","UTI ","","","","","",sInput,"si")
       Formulariop = Formulariop & ofv.GenRow("","","","","25","","",Celdap,"si")


       Celdap = ""
       Celdap = Celdap & ofv.GenCelda("","UTI COLSPAN=4 " & wkstylo2,"","","","","","PROPUESTA&nbsp;","si")
       Formulariop = Formulariop & ofv.GenRow("","","","","25","","",Celdap,"si")



       Celdap = ""
       Celdap = Celdap & ofv.GenCelda("","UTI COLSPAN=4 " & wkstylo2 & " style='border-style:SOLID;border-width:4;border-color:#336699;' ","CENTER","","","","",WSPTABLA,"si")
       Formulariop = Formulariop & ofv.GenRow("","","","","25","","",Celdap,"si")


       Celdap = ""
       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstylo2,"","","","","","RIESGO NUEVO&nbsp;","si")
       sInput = wspDATArNdsc
       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstyloD,"","","","","",sInput,"si")
       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstylo2,"","","","","","CONTROLES NUEVOS&nbsp;","si")
       sInput = wspDATMNdsc
       Celdap = Celdap & ofv.GenCelda("","UTI " & wkstyloD,"","","","","",sInput,"si")
       Formulariop = Formulariop & ofv.GenRow("","","","","25","","",Celdap,"si")




    Celda = ""
    'OBSERVACIONES
       Celda = Celda & ofv.GenCelda("","UTI " & wkstylo1,"","","","","","OBSERVACIONES&nbsp;","si")
      nCol = 80

    ''  if  len(wdescri) = 0 then wdescri = "Sin Descripcion"
          wdescri = ""
      if  not  isnull(rs(13)) then
          wdescri = rs(13)
      end if

      t3b = "<font class=at><textarea class=dt "
      t3b = t3b & "name=" & rs(13).name & " cols=" & nCol & " rows=3 "
      if  bModificar AND SREAD= "" and (rs("estado").value <> "5" or rs("estado").value <> "0") then
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
    Celda = Celda & ofv.gencelda("","UTI COLSPAN=2 " & wkstyloD,"","","","","",SINPUT,"si")


       sAccionDET = "ORAGENDAEDITDET.ASP?VISION=" & SVISION & "&ACTIVRIESGO=" & RS(0) & "&CantRegMover=" & CantRegMoverx & "&CODAGENDA=" & RS(10)
       FILA1 = "<INPUT TYPE='BUTTON' VALUE='DETALLE' CLASS='BT' title='Detalle del plan.' STYLE='cursor:pointer;' onclick=""disp_confirm('Confirma acceder al detalle de este plan de mitigacion???','FAGDET');""  > "
       FormuDET = ofv.GenForm("FAGDET","",sAccionDET,"","post","","","si")


     Celda = Celda & ofv.gencelda("","UTI  " & wkstyloD,"","","","","",FILA1,"si")
     Formulariop = Formulariop & ofv.GenRow("","","","","25","","",Celda,"si")

	'response.Write("Busqueda " & rs("estado").value)

	if (rs("estado").value = "5" or rs("estado").value = "0") then
		sAccion = ""
	else
	    sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
    	sAccion = sAccion & "&tabla=ORAGENDA&volver=" & vuelta
	end if

    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Formulariop,"si") & FormuDET

       if xTrColor = "#FFFFFF" then
          xTrColor = "#F0F0F0"
       ELSE
          xTrColor = "#FFFFFF"
       END IF

    posi = posi + 1

       Formu1 = ""
       Formu2 = ""

    if DFEC = "1" OR DFEC = "2" OR DFEC = "4" then



    if DFEC = "1" then
       sAccion = "ORAGENDAEDIT.ASP?VISION=" & SVISION & "&ACTIVRIESGO=" & RS(0) & "&CantRegMover=" & CantRegMoverx & "&OPCION=CAN"
       FILA1 = "<INPUT TYPE='BUTTON' VALUE='CANCELAR' CLASS='BT' title='Cancelar plan.' STYLE='cursor:pointer;' onclick=""disp_confirm('Confirma la CANCELACION de este plan de mitigacion???','FCAN');""  > "
       Formu1 = ofv.GenForm("FCAN","",sAccion,"","post","","","si")

       sAccion = "ORAGENDAEDIT.ASP?VISION=" & SVISION & "&ACTIVRIESGO=" & RS(0) & "&CantRegMover=" & CantRegMoverx & "&OPCION=INI"
       FILA2 = "<INPUT TYPE='BUTTON' VALUE='INICIAR' CLASS='BT' title='Iniciar plan.' STYLE='cursor:pointer;' onclick=""disp_confirm('Confirma la INICIACION de este plan de mitigacion???','FINI');""  > "
       Formu2 = ofv.GenForm("FINI","",sAccion,"","post","","","si")

       Celda = ""
       Celda = Celda & ofv.gencelda("","UT","","","","","",FILA1,"si")
       Celda = Celda & ofv.gencelda("","UT","","","","","",FILA2,"si")
       FormulACC = ofv.GenRow("","","","","25","","",Celda,"si")

    ELSEIF  DFEC = "2" then
       sAccion = "ORAGENDAEDIT.ASP?VISION=" & SVISION & "&ACTIVRIESGO=" & RS(0) & "&CantRegMover=" & CantRegMoverx & "&OPCION=REP"
       FILA1 = "<INPUT TYPE='BUTTON' VALUE='REPLANIFICAR' CLASS='BT' title='Replanificar plan.' STYLE='cursor:pointer;' onclick=""disp_confirm('Confirma la REPLANIFICACION de este plan de mitigacion???','FCAN');""  > "
       Formu1 = ofv.GenForm("FCAN","",sAccion,"","post","","","si")

       sAccion = "ORAGENDAEDIT.ASP?VISION=" & SVISION & "&ACTIVRIESGO=" & RS(0) & "&CantRegMover=" & CantRegMoverx & "&OPCION=FIN"
       FILA2 = "<INPUT TYPE='BUTTON' VALUE='FINALIZAR' CLASS='BT' title='Finalizar plan.' STYLE='cursor:pointer;' onclick=""disp_confirm('Confirma la finalizacion de este plan de mitigacion???','FINI');""  > "
       Formu2 = ofv.GenForm("FINI","",sAccion,"","post","","","si")

       Celda = ""
       Celda = Celda & ofv.gencelda("","UT","","","","","",FILA1,"si")
       Celda = Celda & ofv.gencelda("","UT","","","","","",FILA2,"si")
       FormulACC = ofv.GenRow("","","","","25","","",Celda,"si")

    ELSE

       sAccion = "ORAGENDAEDIT.ASP?VISION=" & SVISION & "&ACTIVRIESGO=" & RS(0) & "&CantRegMover=" & CantRegMoverx & "&OPCION=REP"
       FILA1 = "<INPUT TYPE='BUTTON' VALUE='REPLANIFICAR' CLASS='BT' title='Replanificar plan.' STYLE='cursor:pointer;' onclick=""disp_confirm('Confirma la REPLANIFICACION de este plan de mitigacion???','FCAN');""  > "
       Formu1 = ofv.GenForm("FCAN","",sAccion,"","post","","","si")


       sAccion = "ORAGENDAEDIT.ASP?VISION=" & SVISION & "&ACTIVRIESGO=" & RS(0) & "&CantRegMover=" & CantRegMoverx & "&OPCION=FIN"
       FILA2 = "<INPUT TYPE='BUTTON' VALUE='FINALIZAR' CLASS='BT' title='Finalizar plan.' STYLE='cursor:pointer;' onclick=""disp_confirm('Confirma la finalizacion de este plan de mitigacion???','FINI');""  > "
       Formu2 = ofv.GenForm("FINI","",sAccion,"","post","","","si")

       Celda = ""
       Celda = Celda & ofv.gencelda("","UT","","","","","",FILA1,"si")
       Celda = Celda & ofv.gencelda("","UT","","","","","",FILA2,"si")
       FormulACC = ofv.GenRow("","","","","25","","",Celda,"si")
    END IF

     IF SREAD="" THEN
      TABLAACC = ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;border-color:#555555;'  ","","","","0","","0","0","","",FormulACC,"si")
     ELSE
      TABLAACC = ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;border-color:#555555;'  ","","","","0","","0","0","","","","si")
     END IF
      Celda =ofv.gencelda("","UTI COLSPAN=4 BGCOLOR='#BBBBBB' ","RIGHT","","","","",TABLAACC & "&nbsp;","si")
      Formulario = Formulario & ofv.GenRow("","","","","25","","",Celda,"si")


    END IF


 end if
  sHTML = sHTML & "<CENTER><BR>"
  sHTML = sHTML & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;border-color:#555555;'  ","","90% ","","0","","0","0","","",Formulario,"si")
  sHTML = sHTML & "</CENTER>" & Formu1 & Formu2





  'Botonera
  Agregar = ""
  Buscar  = ""

  sAccion = "ORAGENDAEX.ASP?CantRegMover=" & CantRegMoverx & "&vision=" & svision
  Celda = ofv.BotonVolver(sAccion)


  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)


  sHTML = sHTML & GTCCALENDAR("M")





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