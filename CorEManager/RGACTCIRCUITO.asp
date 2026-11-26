<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else

 response.expires = 0

 sACTIVIDAD = request.querystring("ACTIVIDAD")
 scodmaster = request.querystring("codmaster") & ""
 stipmaster = request.querystring("tipmaster") & ""
 sOPCION    = request.querystring("OPCION")
 scodriesgo    = request.querystring("codriesgo")

 session("ModMain") = "ED"

''''''''''''''''''''''''''''

  set cn = ofv.conectar(ofv.strconn2)


  FUNCION = ""

  sHTML = ofv.MENUheader("",funcion)
  sHTML = sHTML & ofv.CorFormHeader



  Fila    = ""
  Celda   = ""
  sHref   = ""

''''''''''''''''''''''''''''''''''''''''''''
''' ESPECIFICACION DE PERIODOS PARA OR  ''''
''''''''''''''''''''''''''''''''''''''''''''

  wktabperiodo = ""
  Celda        = ""


   Celda = Celda & ofv.GenCelda("","uti style='background-color:#888888;color:#ffffff;' ","","10%","","","","PERIODO:","si")
   Celda = Celda & ofv.GenCelda("","uti ","","65%","","","",session("ORPERIODODESC"),"si")
   Celda = Celda & ofv.GenCelda("","ut  ","","20%","","","","&nbsp;","si")
   Fila  = ofv.GenRow("","","","","","","",Celda,"si")
   wktabperiodo = ofv.GenTabla("","","","100% ","100%","0","","-1","1","","",Fila,"si")


''''''''''''''''''''''''''''''''''''''''''''
 sImage = "<font face='corflaticon' size=2 Xcolor='#FFFFFF' style='border-style:OUTSET;border-width:2;' >&#0233;</font>"

 wkproccod = "<a class=nt  href='" & Session("ASPSTR") & "' ><b>"
 wkproccod = wkproccod & sImage
 wkproccod = wkproccod & "</b></a>"

 AGREGAR = ""
 wkproccodN = ""


      Strsql = "Select A.* From OBJETOS A"
      Strsql = Strsql & " where  "
      Strsql = Strsql & "  A.CODOBJETO = '" & scodmaster & "' and A.TIPOBJETO = '" & stipmaster & "'  "


 '''  response.write strsql & "<br>" & chr(13)
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT rs.eof then


       WKPROCDESC = RS(3) & "&nbsp;"
       WKPROCVINC = CLNG("0" & RS(15))
       wspTIPO    = (rs(17) & "")


    END IF









 '' strsql = "select a.* from OBJACCION a "
 '' strsql = strsql & " WHERE a.CODACCION = '" &  SACTIVIDAD & "' "


  strsql = "SELECT A.* FROM RGACTIVOPRTY A "
  strsql = strsql & " WHERE A.CODOBJETO =  0" & SACTIVIDAD & "  "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     sCODPROCED = RS(0)
     WKTIPOACC  = "M" & RS(0)

     sTIPOBJETO = RS(2)

     sACTVALUACION = RS(6)
     sACTFAMILY    = RS(7)

   IF wspTIPO = "XCR" OR wspTIPO = "VN" THEN

      WKSPTITLE = "Agregar riesgo asociado."

   ELSEIF wspTIPO = "CR"  THEN

      WKSPTITLE = "Agregar riesgo asociado."
   else

      WKSPTITLE = "Agregar amenaza asociada."

   end if





      WKPROCACCDESC = "ACTIVO INFORMATICO"


      sAccion = "RGACTCIRCUITO.ASP?ACTIVIDAD=" & SACTIVIDAD & "&codmaster=" & Scodmaster & "&TIPmaster=" & STIPmaster
      sAccion = sAccion & "&Opcion=SI"


    IF SESSION("ORPERIODOST") = "OK" AND (SESSION("ORRSKST") <> 3 AND SESSION("ORRSKST") <> 4  AND SESSION("ORRSKST") <> 5  AND SESSION("ORRSKST") <> 2 ) AND (sACTVALUACION > 0 AND sACTFAMILY > 0) THEN

       sImage = "<font face='corflaticon' size=2 color='#990000' style='border-style:OUTSET;border-width:2;' >5</font>"

       wkproccodN = "<a class=nt href='" & sAccion & "' id='NewElem' "
       wkproccodN = wkproccodN & " TITLE='" & WKSPTITLE & "'   >"
       wkproccodN = wkproccodN & sImage
       wkproccodN = wkproccodN & "</a>"

    END IF



  else

      WKPROCACCDESC = "no se encuentra " &  SACTIVIDAD


  END IF


           select case wspTIPO




                  case "VN"

                        sInput = "EVENTO VINCULADO EMANAGER"
                        sInput2 = "ACTIVIDAD"
                        wspTIPOCLC = "#008800"

                  case "CR"

                        sInput = "EVENTO MANUAL"
                        sInput2 = "ACTIVIDAD"
                        wspTIPOCLC = "#99CCDD"

                  case "AI"

                        sInput = "ACTIVO INFORMATICO"
                        sInput2 = "ACTIVO"
                        wspTIPOCLC = "#0000CC"

                  case else

                        sInput = "SIN DETERMINAR"
                        sInput2 = "----"
                        wspTIPOCLC = "#000000"

           end select


    sImagePRC = "<font face='corflaticon' size=2 color='" & wspTIPOCLC & "'  >&#060;</font>"

    WKDIVSECUEN = GENPROCSECUENEX()

    WKPROCDESC = WKDIVSECUEN & sInput & ":&nbsp;" & WKPROCDESC


    Celda = ""

   Celda = Celda & ofv.GenCelda("","Ut style='background-color:#f0f0f0;color:#990000;' ","CENTER","1%","","","",sImagePRC,"si")
   Celda = Celda & ofv.GenCelda("","ut style='background-color:#f0f0f0;color:#888888;' ","","85%","","","", WKPROCDESC,"si")
   Celda = Celda & ofv.GenCelda("","ut style='Xbackground-color:#993333;Xcolor:#ffffff;' ","CENTER","1%","","","",wkproccodN,"si")
   Celda = Celda & ofv.GenCelda("","ut style='Xbackground-color:#993333;Xcolor:#ffffff;' ","CENTER","1%","","","",wkproccod,"si")
   Fila  = ofv.GenRow("","","","","","","",Celda,"si")
   TablaH = ofv.GenTabla("","aaa style='background-color:#f0eedd;border-style:solid;border-width:1;border-color:#000000;' valign='middle' align='center' ","","100% ","100%","0","","-1","1","","",Fila,"si")




  ' Celda = ""
  ' Celda = Celda & ofv.GenCelda("","uti style='background-color:#993333;color:#ffffff;' ","","","","","",sInput2 & ":&nbsp;" & WKPROCACCDESC,"si")
  ' Fila  = ofv.GenRow("","","","","30","","",Celda,"si")
  ' TablaH2 = ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;background-color:#aaaaaa;' valign='middle' align='center' ","","100% ","","0","","-1","1","","",Fila,"si")

    TablaH2 = "" 


''''''''''''''''''''''''''''''''''''''''''''''''
'' INCORPORA RIESGOS ASOCIADOS - OPCION = SI '''
''''''''''''''''''''''''''''''''''''''''''''''''

 IF SOPCION = "SI" THEN

        strsql = "Insert into RGACTANALISIS "
        strsql = strsql & " (PERIODO,ACTIVIDAD,EVENTO,codobjeto, tipobjeto,"
        strsql = strsql & " SECUEN, CODRIESGO, ITEMRIESGO, PROBABILIDAD,IMPACTO, "
        strsql = strsql & " IMPACTOCONF,IMPACTOINTE,IMPACTODISP,RIESGOBRUTO, RIESGORESIDUAL, "
        strsql = strsql & " RIESGORESIDUALNV, VULNERABILIDAD,REQUIEREPLAN"
        strsql = strsql & " ) "
        strsql = strsql & " Values ("
        strsql = strsql &  "'" & session("ORPERIODO") & "','" & sACTIVIDAD & "','" & sCODPROCED & "','" & sACTIVIDAD & "','" & sTIPOBJETO & "',"
        strsql = strsql & "'10','" & GTkey(SCODMASTER) & "','0','0','0',"
        strsql = strsql & " '0','0','0','0','0',"
        strsql = strsql & " '0',' ','NO'"
        strsql = strsql & ")"

        call ofv.crearconsultaEx(StrSql,cn,1,volver)


 END IF


 IF SOPCION = "DEL" THEN

        strsql = "Delete From RGACTANALISIS Where CODRIESGO = '" & SCODRIESGO & "' "

        call ofv.crearconsultaEx(StrSql,cn,1,volver)

        strsql = "Delete From ORACTIVRCONTROLES Where CODRIESGO = '" & SCODRIESGO & "' "

        call ofv.crearconsultaEx(StrSql,cn,1,volver)




 END IF

''''''''''''''''''''''''''''''''''''''''''''''''''

   Fila = ""

   wspcolor1  = "#990000"
   wspcolor2  = "#009900"
   wspcolor3  = "#D7CF77"
   wspcolor4  = "#EEEEEE"
   sImageDAP  = "<font face='corflaticon' size=3 color='#FFFFFF'>3</font>"

  DIM TABRANRIESGOS(10,5)


  Strsql = "Select * From rgactRANGORIESGO ORDER BY VALOR DESC"
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







   Strsql = "Select * From RGACTANALISIS where PERIODO = 0" & session("ORPERIODO") & " AND ACTIVIDAD = '" & SACTIVIDAD & "' "
   Strsql = Strsql & " ORDER BY SECUEN "

' response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then


     SNOHAYDSC = "NO EXISTEN AMENAZAS ASOCIADAS AL ACTIVO"

     IF (sACTVALUACION > 0 AND sACTFAMILY > 0) THEN


     ELSE
        SNOHAYDSC = SNOHAYDSC & "<BR><HR>NO SE PUEDE ASOCIAR AMENAZAS AL ACTIVO<BR>SI NO ESTA DETERMINADA SU CLASE<BR>Y/O SI NO TIENE VALUACION DEFINIDA."
     END IF  


     Celda = ""
     Celda = Celda & ofv.GenCelda("","vtG ","CENTER","","","","",SNOHAYDSC,"si")

     Fila = ofv.GenRow("","","","","","","",Celda,"si")
     Tabla = ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-COLOR:#993333;' BGCOLOR='#F9F9F9' ","","300 ","","0","","8","8","","",Fila,"si")

     Tabla = "<CENTER><BR><BR><BR><BR>" & Tabla & "</CENTER>"

  else


       wkcoddelete = "N"
       wktitdelete = " "
       sImageELS    = "<font face='corflaticon' size=3 color='#bb0000'>&#0253;</font>"
       sImageELN    = "<font face='corflaticon' size=3 color='#dddddd'>&#0253;</font>"

    xTrColor = "#FFFFFF"

    CeldaK  = ""
    Celda  = ""

    Celda  = Celda & ofv.GenCelda("","uti colspan=3  style='color:#000000;border-style:solid;border-width:1;border-COLOR:#888888;' ","","","","","","AMENAZAS ASOCIADAS","si")

    Celda  = Celda & ofv.GenCelda("","Ut  style='color:#000000;border-style:solid;border-width:1;border-COLOR:#888888;' ","CENTER","5%","","","","INHERENTE&nbsp;&nbsp;&nbsp;","si")
    Celda  = Celda & ofv.GenCelda("","Ut  style='color:#000000;border-style:solid;border-width:1;border-COLOR:#888888;' ","CENTER","5%","","","","RESIDUAL","si")
    Celda  = Celda & ofv.GenCelda("","Ut  style='color:#000000;border-style:solid;border-width:1;border-COLOR:#888888;' ","CENTER","5%","","","","PLAN&nbsp;&nbsp;&nbsp;&nbsp;","si")
    Celda  = Celda & ofv.GenCelda("","Ut  style='color:#000000;border-style:solid;border-width:1;border-COLOR:#888888;' ","CENTER","5%","","","","RESIDUAL NUEVO","si")

  ''  Celda  = Celda & ofv.GenCelda("","Ut  style='color:#000000;border-style:solid;border-width:1;border-COLOR:#888888;' ","CENTER","5%","","","","INDICADOR","si")
    FilaN  = filaN & ofv.GenRow("","aaa bgcolor='#bbccdf' ","","","25","","",Celda,"si")

    do until rs.eof


       wspcolor  = "#336699"
       sImageP   = "<font face='corflaticon' size=2 color='" & wspcolor & "'>&#0159;</font>"

       wkproccod = "RGACTCIRCUITODET.ASP?ACTIVIDAD=" & SACTIVIDAD
       wkproccod = wkproccod & "&RIESGO=" & RS(7) & "&codmaster=" & scodmaster & "&tipmaster=" & sTIPmaster


       WKCODRIESGO = "NO"

       sRIESGODESC = "<FONT COLOR='#880000' style='border-style:solid;border-width:1;border-COLOR:#993333;' TITLE='ESTABLECER AMENAZA ASOCIADA.'  > ESTABLECER AMENAZA ASOCIADA </FONT>"

      'CODIGO DE AMENAZA
       Strsql = "Select * From ORAMENAZAS Where OBJID = 0" & rs(8)
       set rsAcciones = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if  not rsAcciones.eof then
           sCODNEMO    = rsAcciones(1)
           sDESCRI     = rsAcciones(2)
           sRIESGODESC = "<U TITLE='ACCEDER / MODIFICAR AMENAZA ASOCIADA.' >" & sCODNEMO & "&nbsp;-&nbsp;" & sDESCRI & "</U>"
       end if
       call ofv.cerrarconsulta(rsAcciones)

       wktitdelete = "Confirma la eliminacion de la amenaza asociada???."







    IF RS(8) > 0 THEN WKCODRIESGO = "SI"


       sHref     = wkproccod
       sLink     = ofv.GenLink("","nt",sHref,"","","",sRIESGODESC,"si")

       Celdak = ""
       Celda  = ""


       wkcoddelete = "S"

       if wkcoddelete = "S" AND SESSION("ORPERIODOST") = "OK" AND (SESSION("ORRSKST") <> 3 AND SESSION("ORRSKST") <> 4   AND SESSION("ORRSKST") <> 5  AND SESSION("ORRSKST") <> 2 ) then
          sImageEL = sImageELS
          sformobj = "FRMRGA" & rs(0)
       '''   sAccion    = "respproceso.asp?codobjeto=" & rs(7) & "&tipobjeto=" & "RGA" & ""

          sAccion = "RGACTCIRCUITO.ASP?ACTIVIDAD=" & SACTIVIDAD & "&codmaster=" & Scodmaster & "&TIPmaster=" & STIPmaster
          sAccion = sAccion & "&Opcion=DEL&CODRIESGO=" & RS(7)

          formobj    = ofv.genform(sformobj,"AAA STYLE='display:none;' ",sAccion,"","post","","","si")
          sfunobj    = " onclick="" disp_confirm('" & wktitdelete & "','" & sformobj & "'); "" "


       else
          wktitdelete = "No se puede eliminar"
          sImageEL = sImageELN
          formobj    = ""
          sfunobj    = ""
       end if




       Celda     = ""
       Celda = Celda & ofv.GenCelda("","ut style='cursor:pointer;' title='" & wktitdelete & "' " & sfunobj & " ","","2%","","","",sImageEL,"si") & formobj


       Celda  = Celda & ofv.GenCelda("","ut","CENTER","1%","","","",sImageP,"si")
       Celda  = Celda & ofv.GenCelda("","UT","","65%","","","",sLink,"si")


       IF WKCODRIESGO = "NO" THEN

          wspcolor0A  = wspcolor4
          wspcolor0B  = wspcolor4
          wspcolor0P  = wspcolor4
          wspcolor0Bb  = wspcolor4

          wspDATA     = "---"
          wspDATB     = "---"
          wspDATP     = "---"
          wspDATBb     = "---"

             wspmontobruto    = "0"
             wspmontoresidual = "0"
             wspmontoresidaj  = "0"
             wspDATAdsc       = "-----"
             wspDATArrdsc     = "-----"
             wspDATArradsc    = "-----"

             wspmontobrn      = ""
             wspmontoren      = ""
             wspmontoran      = ""

       ELSE




             wspDATA     = rs(14)
             wspDATB     = rs(15)
             wspDATP     = rs(17)
             wspDATBB     = rs(20)

             wspDATP      = 0
             wspDATI      = 0

             wspmontobruto    = "0"
             wspmontoresidual = "0"
             wspmontoresidaj  = "0"

             wspmontobrn      = ""
             wspmontoren      = ""
             wspmontoran      = ""

             wspDATAdsc       = "-----"
             wspDATArrdsc     = "-----"
             wspDATArradsc    = "-----"

             WKTCONFIDENP = CINT("0" & RS(11))
             WKTINTEGRIDP = CINT("0" & RS(12))
             WKTDISPONIBP = CINT("0" & RS(13))


             WKTIMPACTO = WKTCONFIDENP


             IF WKTIMPACTO < WKTINTEGRIDP THEN WKTIMPACTO = WKTINTEGRIDP
             IF WKTIMPACTO < WKTDISPONIBP THEN WKTIMPACTO = WKTDISPONIBP

             wspporc = 100

             Strsql  = "Select VALOR,porcentaje From ORACTPROBOCURRENCIA where OBJID = 0" & RS(9)
             set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if NOT rs1.eof then
                wspDATPX    = "0" & RS1(0)
                wspDATP     = CINT(wspDATPX)
                wspporc     = CINT("0" & RS1(1))
             END IF




             wspmonto = 0


             wspDEGRADA = WKTIMPACTO




             Strsql  = "Select VALOR,impdesde,imphasta,obs1,DESCRI From ORACTIMPACTO where OBJID = 0" & RS(10)
             set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if  NOT rs1.eof then
                  wspDATIX    = "0" & RS1(0)
                  wspDATI     = CINT(wspDATIX)
                  wspDEGRADA  = Clng("0" & RS1(1))
             END IF


             wspmonto    = clng("0" & sACTVALUACION)





            if wspDATP > 0 and wspDATI > 0 then





         '    Strsql  = "Select descri From ORACTRIESGOBRUTO where NIVPROB =  0" & RS(9) & " AND NIVIMPA = 0" & RS(10) & " "
 '''  response.write strsql & "<br>" & chr(13)
          '   set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          '   if NOT rs1.eof then

           '     FOR II=1 TO TTABRANRIESGO

           '         IF RS1(0)  =  TABRANRIESGOS(II,0) THEN
           '            wspDATAdsc    = TABRANRIESGOS(II,2)   'DESCRIPCION
           '            wspDATAdsccl    = TABRANRIESGOS(II,3) 'COLOR
           '            EXIT FOR
           '         END IF

           '     NEXT


           '  END IF




             WSPNEWRB =  clng(wspporc * wspDEGRADA  / 100)


             SVTOPE = 0
             SVCODI = 0
             SVDSC  = 0
             WSPNEWRBdsc    = "NO"


             Strsql = "Select * From RGACTRANGORIESGO ORDER BY PORCENTAJE"
        '''  response.write strsql & "<br>" & chr(13)
             set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             DO UNTIL rs1.eof 

                 IF cint(WSPNEWRB) = 0  THEN
                    SVTOPE = RS1(4)
                    SVCODI = RS1(0)
                    SVDSC  = RS1(2)
                    SVCOL  = RS1(3)
                    WSPNEWRBdsc    = SVDSC   'DESCRIPCION
                    wspDATAdsc     = SVDSC   'DESCRIPCION
                    wspDATAdsccl   = SVCOL   'COLOR
                    EXIT DO
                 END IF

                 IF (NOT WSPNEWRB < SVTOPE) AND (WSPNEWRB < CDBL("0" & RS1(4))) THEN
                    SVTOPE = RS1(4)
                    SVCODI = RS1(0)
                    SVDSC  = RS1(2)
                    SVCOL  = RS1(3)
                    WSPNEWRBdsc    = SVDSC   'DESCRIPCION
                    wspDATAdsc     = SVDSC   'DESCRIPCION
                    wspDATAdsccl   = SVCOL   'COLOR
                    EXIT DO
                 END IF

                    SVTOPE = RS1(4)
                    SVCODI = RS1(0)
                    SVDSC  = RS1(2)
                    SVCOL  = RS1(3)

             RS1.MOVENEXT
             LOOP 

             IF WSPNEWRBdsc = "NO" THEN
                IF SVCODI = 0 THEN
                   WSPNEWRBdsc    = "NO ESPECIFICADO"
                ELSE
                    WSPNEWRBdsc    = SVDSC   'DESCRIPCION
                    wspDATAdsc     = SVDSC   'DESCRIPCION
                    wspDATAdsccl   = SVCOL   'COLOR
                END IF

             END IF 




        '     Strsql  = "Select descri From ORACTRIESGO where "
        '     Strsql  = Strsql & " NOT (0" & replace(cstr(wspDATB & ""),",",".") & " > VALORHASTA) AND "
        '     Strsql  = Strsql & " NOT (0" & replace(cstr(wspDATB & ""),",",".") & " < VALORDESDE) "

           ''  Strsql  = "Select descri From ORNIVriesgo where valorhasta > 0" & replace(cstr(wspDATB & ""),",",".")
  '' response.write strsql & "<br>" & chr(13)
        '     set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        '     if NOT rs1.eof then

        '        FOR II=1 TO TTABRANRIESGO

        '            IF RS1(0)  =  TABRANRIESGOS(II,0) THEN
        '               wspDATArrdsc    = TABRANRIESGOS(II,2)   'DESCRIPCION
        '               wspDATArrdsccl    = TABRANRIESGOS(II,3) 'COLOR
        '               EXIT FOR
        '            END IF

        '        NEXT

              ''  wspDATArrdsc    = RS1(0)
        '     END IF


        '     Strsql  = "Select descri From ORACTRIESGO where "
        '     Strsql  = Strsql & " NOT (0" & replace(cstr(wspDATBB & ""),",",".") & " > VALORHASTA) AND "
        '     Strsql  = Strsql & " NOT (0" & replace(cstr(wspDATBB & ""),",",".") & " < VALORDESDE) "

           ''  Strsql  = "Select descri From ORNIVriesgo where valorhasta > 0" & replace(cstr(wspDATBB & ""),",",".")
 ''  response.write strsql & "<br>" & chr(13)
        '     set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        '     if NOT rs1.eof then

        '        FOR II=1 TO TTABRANRIESGO

        '            IF RS1(0)  =  TABRANRIESGOS(II,0) THEN
        '               wspDATArrAdsc    = TABRANRIESGOS(II,2)   'DESCRIPCION
        '               wspDATArrAdsccl    = TABRANRIESGOS(II,3) 'COLOR
        '               EXIT FOR
        '            END IF

        '        NEXT


           ''     wspDATArradsc    = RS1(0)
        '     END IF


             Strsql  = "Select * From ORACTriesgo where "
             Strsql  = Strsql & " NOT (0" & replace(cstr(wspDATB & ""),",",".") & " > VALORHASTA) AND "
             Strsql  = Strsql & " NOT (0" & replace(cstr(wspDATB & ""),",",".") & " < VALORDESDE) "
  ' response.write wspDATMBCN & "-----"

             set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if NOT rs1.eof then

                       wspDATArrdsc      = RS1(1) 'DESCRIPCION
                       wspDATArrdsccl    = RS1(5) 'COLOR

             END IF


             Strsql  = "Select * From ORACTriesgo where "
             Strsql  = Strsql & " NOT (0" & replace(cstr(wspDATBB & ""),",",".") & " > VALORHASTA) AND "
             Strsql  = Strsql & " NOT (0" & replace(cstr(wspDATBB & ""),",",".") & " < VALORDESDE) "
 ' response.write strsql & "<br>" & chr(13)
             set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

             if NOT rs1.eof then
                       wspDATArrAdsc     = RS1(1)   'DESCRIPCION
                       wspDATArrAdsccl   = RS1(5) 'COLOR

             END IF




             wspDATA     = cdbl(wspDATA)
             wspDATB     = cdbl(wspDATB)
             wspDATBB    = cdbl(wspDATBB)


               wspDATM    = 0
               wspDATMB   = 0
               wspDATMBCN =0

               wspDATmX   = 0
               wspDATmt   = 0

               wspDATMCB  = 0
               wspDATMBe  = 0
               wspDATMBCN = 0
               
               wspDATMC   = 0
               wspDATMe   = 0

             Strsql = "Select A.* From oractivrcontroles A "
             strsql = strsql & "  where a.periodo = 0" & session("ORPERIODO") & " "
             strsql = strsql & "  AND a.actividad = '" & SACTIVIDAD & "' AND A.CODRIESGO =  '" & RS(7) & "' "
             strsql = strsql & "  ORDER BY A.TIPO "

          ' response.write strsql & "<br>"
             set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
             if not rs1.eof then
                do until rs1.eof



                   Strsql  = "Select VALORD From RGACTEFICIENCIASV where OBJID = 0" & RS1(8)
                   set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                   if NOT rs2.eof then
                      wspDATmX    = "0" & RS2(0)
                      wspDATmt     = CINT(wspDATmX)

                    IF RS1("TIPO") = 1 or (RS1("TIPO") = 0 and (rs1("SECUEN") & "") <> "0" ) THEN
                      wspDATMCB    = wspDATMCB + 1
                      wspDATMBe     = wspDATMBe + wspDATmt
                      IF RS1("TIPO") = 1 then wspDATMBCN     =  wspDATMBCN + 1
                    END IF

                    IF RS1("TIPO") = 0 THEN
                      wspDATMC    = wspDATMC + 1
                      wspDATMe     = wspDATMe + wspDATmt
                    END IF



                   END IF



                   rs1.movenext
                loop

                if wspDATMC > 0 and wspDATMe > 0 then
                    wspDATM = cdbl(round(wspDATMe / wspDATMC,2))
                else
                    wspDATM = cdbl(0)
                end if

                if (wspDATMCB > 0 and wspDATMBe > 0)  then
                    wspDATMB = cdbl(round((wspDATMBe) / (wspDATMCB),2))
                else
                    wspDATMB = cdbl(0)
                end if

             end if



             
             wspDATM = (100 - wspDATM)
             wspDATMB = (100 - wspDATMB)            



             wspmontobruto    = clng(clng(wspmonto * wspporc / 100) * wspDEGRADA / 100)

             wspmontobrn      = "<hr>$&nbsp;" & wspmontobruto


             wspmontoresidual = clng(wspmontobruto * wspDATM / 100)
             wspmontoresidaj  = clng(wspmontobruto * wspDATMB / 100)



             wspmontoren      = "<hr>$&nbsp;" & wspmontoresidual
             wspmontoran      = "<hr>$&nbsp;" & wspmontoresidaj




            end if






          if rs(17) = "SI"  then

             wspcolor0P  = wspcolor2

          else

             wspcolor0P  = wspcolor4

          end if




       END IF



       wsstyloF   = " style='border-style:solid;border-width:1;border-COLOR:#888888;COLOR:#000000;"
       wsstyloF   = wsstyloF & "background-color:#f0f0f0;' "

       Celdak  = Celdak & ofv.GenCelda("","Ut rowspan=2 " & wsstyloF,"CENTER","5%","","","",wspDATAdsc & "" & wspmontobrn,"si")

       wsstyloF   = " style='border-style:solid;border-width:1;border-COLOR:#000077;COLOR:#888888;"
       wsstyloF   = wsstyloF & "background-color:#f0f0f0;' "

       Celdak  = Celdak & ofv.GenCelda("","Ut  rowspan=2  " & wsstyloF,"CENTER","5%","","","",wspDATArrdsc & "" & wspmontoren,"si")

       wsstyloF   = " style='border-style:solid;border-width:1;border-COLOR:#888888;COLOR:#888888;"
       wsstyloF   = wsstyloF & "background-color:" & wspcolor0P & ";' "

       Celdak  = Celdak & ofv.GenCelda("","Ut  rowspan=2  " & wsstyloF,"CENTER","5%","","","",sImageDAP,"si")


       wsstyloF   = " style='border-style:solid;border-width:1;border-COLOR:#993333;COLOR:#888888;"
       wsstyloF   = wsstyloF & "background-color:#f0f0f0;' "

       Celdak  = Celdak & ofv.GenCelda("","Ut  rowspan=2  " & wsstyloF,"CENTER","5%","","","",wspDATArradsc & "" & wspmontoran,"si")


        wsstyloF   = " style='border-style:solid;border-width:1;border-COLOR:#888888;COLOR:#888888;"

    IF WSPSIGUE = "SI" THEN

       if wspDATP <> "---" then
          if  wspDATP > 0 and wspDATI  then
              strsql= "SELECT  * FROM ORRELAINDIRIESGO WHERE  CODRIESGO='" & rs(7) &"'"
	      SET RSREL = ofv.crearconsultaEx(StrSql,cn,1,volver)

              IF not RSREL.EOF THEN
	         wsstyloF   = wsstyloF & "background-color:" & wspcolor0P & ";cursor:pointer;' "
	      else
	         wspcolor0P ="#993333"
	         wsstyloF   = wsstyloF & "background-color:" & wspcolor0P & ";cursor:pointer;' "
	      end if
              call ofv.cerrarconsulta(RSREL)
              url ="'ORASIGNAINDICADOR.asp?opcion=CONSULTA&riesgo=" &rs(7) & "' ,'mywindow','status=0,toolbar=0,width=900,height=400'"
              Celdak  = Celdak & ofv.GenCelda("","Ut  onclick="&chr(34)&"window.open("&url&");"&chr(34)&"  rowspan=2  " & wsstyloF,"CENTER","5%","","","",sImageDAP,"si")
          else
              wspcolor0P = "#993333"
	      wsstyloF   = wsstyloF & "background-color:" & wspcolor0P & ";cursor:pointer;' "
              celdak     = celdak & ofv.GenCelda("","Ut  rowspan=2  " & wsstyloF,"CENTER","5%","","","",sImageDAP,"si")
          end if
       else
          url ="'ORASIGNAINDICADOR.asp?opcion=CONSULTA&riesgo=" &rs(7) & "' ,'mywindow','status=0,toolbar=0,width=900,height=400'"
          wspcolor0P ="#993333"
	  wsstyloF   = wsstyloF & "background-color:" & wspcolor0P & ";cursor:pointer;' "
          celdak     = celdak  & ofv.GenCelda("","Ut  onclick="&chr(34)&"window.open("&url&");"&chr(34)&"    rowspan=2  " & wsstyloF,"CENTER","5%","","","",sImageDAP,"si")
       end if

      END IF


       Celda = Celda & Celdak
       FilaN = filaN & ofv.GenRow("","AAA bgcolor=" & xTrColor & " ","","","20","","",Celda,"si")

       Celda  = ""
       Celda  = Celda & ofv.GenCelda("","uti colspan=2   ","","","","","","&nbsp;","si")
       Celda  = Celda & ofv.GenCelda("","uti   style='border-style:dotted;border-width:1;border-COLOR:#bbbbbb;' ","","","","","",rs(16) & "&nbsp;","si")
       FilaN = filaN & ofv.GenRow("","AAA bgcolor=" & xTrColor & " ","","","20","","",Celda,"si")


       if xTrColor = "#FFFFFF" then
          xTrColor = "#F0F0F0"
       ELSE
          xTrColor = "#FFFFFF"
       END IF

      rs.movenext
    loop

    Tabla = "<BR>" & ofv.GenTabla("","","","99%","","0","","1","2","","",Filan,"si")

  end if

    sHTML = sHTML & "<script type='text/javascript'>"
    sHTML = sHTML & "function disp_confirm(disp_msg,disp_frm)"
    sHTML = sHTML & "  {"
    sHTML = sHTML & "  var r=confirm(disp_msg);"
    sHTML = sHTML & "  if (r==true)"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "  var fr= document.getElementById(disp_frm);"
  '  sHTML = sHTML & "    alert(fr.name);"
    sHTML = sHTML & "    fr.submit();"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  else"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "    alert('Se cancelo la accion!');"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  }"
    sHTML = sHTML & "</script>"


  sHTML = sHTML & "<SCRIPT language=" & Chr(34) & "VBScript" & Chr(34) & ">" & Chr(13)

  sHTML = sHTML & " function Abre() " & Chr(13)
  sHTML = sHTML & " if window.event.button=2  then  " & Chr(13)
'  sHTML = sHTML & "    if window.event.srcElement.tagName = ""A"" or window.event.srcElement.tagName = ""IMG""  or window.event.srcElement.tagName = ""FONT"" then  " & Chr(13)
'  sHTML = sHTML & "       window.alert(""dato "" & window.event.srcElement.ID & "" tag "" & window.event.srcElement.tagName ) " & Chr(13)
  sHTML = sHTML & "       Editar(window.event.srcElement.ID) " & Chr(13)
 ' sHTML = sHTML & "    end if " & Chr(13)
  sHTML = sHTML & " end if " & Chr(13)
  sHTML = sHTML & " end function " & Chr(13) & Chr(13)

  sHTML = sHTML & " function Editar(Elemento) " & Chr(13)
  sHTML = sHTML & " if elemento = " & Chr(34) & Chr(34) & " Then " & Chr(13)
  sHTML = sHTML & "    exit function " & Chr(13)
  sHTML = sHTML & " end if " & Chr(13)
  sHTML = sHTML & " dato = " & Chr(34) & "edicion2.asp?"



  sHTML = sHTML & "ID=" & scodmaster & "-" & STIPMASTER & "X-X" & Chr(34) & Chr(13)


  sHTML = sHTML & " window.showModalDialog "
  sHTML = sHTML & "dato,window," & Chr(34) & "dialogWidth:600px;dialogHeight:360px;help:no;resizable:yes;dialogTop:250px;dialogLeft:300px;status:no" & Chr(34) & Chr(13)



  sHTML = sHTML & " end function " & Chr(13)

  sHTML = sHTML & "</SCRIPT> " & Chr(13)

'''  sHTML = sHTML & FormuP



  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

'''''''''''''''''''''''''''

  Celda = ""

  Celda = Celda & "<div id='DX01' style='margin-top:1;margin-left:1;position:absolute;left:5px;top:110px;display:block;"
  Celda = Celda & "width:95%;height:30;overflow:auto;border-style:solid;border-width:1;border-color:#888888;"
  Celda = Celda & "background-color:#f9f9f9;"
  Celda = Celda & "' >"
  Celda = Celda & wktabperiodo
  Celda = Celda & "</div>"

  Celda = Celda & "<div id='DX0' style='margin-top:1;margin-left:1;position:absolute;left:5px;top:140px;display:block;"
  Celda = Celda & "width:95%;height:30;overflow:auto;"
  Celda = Celda & "background-color:#f9f9f9;"
  Celda = Celda & "' >"
  Celda = Celda & TablaH
  Celda = Celda & "</div>"

  Celda = Celda & "<div id='DX2' style='margin-top:1;margin-left:1;position:absolute;left:5px;top:170px;display:block;"
  Celda = Celda & "width:95%;height:70%;overflow:auto;border-style:solid;border-width:1;border-color:#888888;"
  Celda = Celda & "background-color:" & session("APLBGD") & ";"
  Celda = Celda & "' >"
  Celda = Celda & TablaH2
  Celda = Celda & TABLA
  Celda = Celda & "</div>"

  sHTML = sHTML & Celda


'''''''''''''''''''''''''''''''''''''''''''''



  sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no","","","","no",Agregar,"no","")

  sHTML = sHTML & "</body></html>"

 END IF

  response.write  sHTML

 FUNCTION GENPROCSECUENEX()

    GENPROCSECUENEX = ""

    Strsql = "Select * From CSRELPATHDESCRI  where COBJ = " & scodmaster
 '' response.write strsql & "<br>"
    set Wrs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT Wrs.eof then

       WKITOPE = 8

       FOR WKI=0 TO 7

           WKIPROC = WRS(WKI) & ""

           IF WKIPROC <> "" THEN
              WKITOPE = WKI
              EXIT FOR
           END IF


       NEXT

       WKICCOL = 7 - WKITOPE

       WKINCOL = 0

       FilaW = ""

           Celda = ""
           Celda2 = ""

       FOR WKI=WKITOPE TO 7




           Celda = Celda & Celda2 & WRS(WKI)

           Celda2 = "<b><font color='#00aa00' size=2 >&#0124;</font></b>"



       NEXT


       IF WKITOPE < 8 THEN

          GENPROCSECUENEX = Celda & Celda2

       END IF



    END IF


 END FUNCTION


%>

