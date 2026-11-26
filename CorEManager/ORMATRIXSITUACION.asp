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
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  SPROCESO = REQUEST.QUERYSTRING("PROCESO")
  SFILTRO  = REQUEST.QUERYSTRING("FILTRO")
  SFTIPO   = REQUEST.QUERYSTRING("TIPO")

  Speriodo = REQUEST.QUERYSTRING("periodo")


  if Speriodo = "" then Speriodo = REQUEST.FORM("PERIODO")



  if Speriodo = "" then Speriodo = session("ORPERIODO")


  if SFILTRO  = "" then
     SFILTRO  = "PRO"

  end if

  if SFTIPO  = "" then
     SFTIPO  = "TODOS"

  end if


  SFTIPOSRCH = SFTIPO

  IF SPROCESO = "000" THEN SFTIPOSRCH = "TODOS"

  DIM TABPROCESOSCRT(200,20) 


  sHTML = ofv.FormHeader(session("FormN"))


  ofv.ObtenerAtributos Session("Form"),""
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

     if Speriodo = "" then

        Strsql = "Select * From ORPERIODOS  where APLIC = '" & wkaplic & "' and ESTADO IN ('OK','CL')  " 
        Strsql = Strsql & " ORDER BY FINAL,INICIO, ESTADO DESC"

        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if NOT rs.eof then
           session("ORPERIODO")     = RS(0)
           session("ORPERIODODESC") = RS(1)
           Speriodo                 = session("ORPERIODO")
           Speriododsc              = session("ORPERIODODESC")

        end if

      else

        Strsql = "Select * From ORPERIODOS  where OBJID = 0" & Speriodo
        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if NOT rs.eof then
           Speriododsc              = RS(1)

        end if

      end if 

  Fila = ""
  xTrColor = "#FFFFFF"

  sImageva    = "<font face='corflaticon' size=2 color='#990000'>&#0110;</font>"
  sImagevm    = "<font face='corflaticon' size=2 color='#f7aa22'>&#0110;</font>"
  sImagevb    = "<font face='corflaticon' size=2 color='#009900'>&#0110;</font>"
  sImagevn    = "<font face='corflaticon' size=2 color='#dddddd'>&#0110;</font>"

  wkstylof = "aaa style='font-family:arial;font-size:6pt;color:#888888;' "






  Strsql = "     select * from ORNIVPROBOCURRENCIA order by objid "

 '' response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then 
     POSI = 0
     DO WHILE NOT RS.EOF
        POSI = POSI + 1


          TABPROCESOSCRT(0,POSI + 3) = RS(2) ' valor de probabilidad


        RS.MOVENEXT
     LOOP  



  END IF


 IF SFTIPOSRCH = "PROCESO" THEN


  Strsql = "     select B.CODOBJETO,B.DESCRI, "
  Strsql = Strsql & " a.* from csrelpathdescri a"
  Strsql = Strsql & "     inner join objetos b on (a.cobj = b.codobjeto)"
  Strsql = Strsql & " where "
  Strsql = Strsql & " B.APLIC = 'RG' AND "

  Strsql = Strsql & "  ("
'  Strsql = Strsql & " A.MC1 = " & SPROCESO & " OR "
'  Strsql = Strsql & " A.MC2 = " & SPROCESO & " OR "
'  Strsql = Strsql & " A.MC3 = " & SPROCESO & " OR "
'  Strsql = Strsql & " A.MC4 = " & SPROCESO & " OR "
'  Strsql = Strsql & " A.MC5 = " & SPROCESO & " OR "
'  Strsql = Strsql & " A.MC6 = " & SPROCESO & " OR "
'  Strsql = Strsql & " A.MC7 = " & SPROCESO & " OR "
  Strsql = Strsql & " A.COBJ = " & SPROCESO & "  "
  Strsql = Strsql & " ) "

  Strsql = Strsql & " ORDER BY B.PRIORIDAD DESC"

 '' response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then 
     POSI = 0
     TOPE = POSI
     DO WHILE NOT RS.EOF
        POSI = POSI + 1
        TOPE = POSI  

          TABPROCESOSCRT(POSI,0) = RS(0) 'COD.PROCESO
          TABPROCESOSCRT(POSI,1) = RS(1) 'NOM.PROCESO
          TABPROCESOSCRT(POSI,2) = 0     'CANT.ACTIVIDADES
          TABPROCESOSCRT(POSI,3) = 0     'FREE
          TABPROCESOSCRT(POSI,4) = 0     'MB
          TABPROCESOSCRT(POSI,5) = 0     'B
          TABPROCESOSCRT(POSI,6) = 0     'M
          TABPROCESOSCRT(POSI,7) = 0     'A
          TABPROCESOSCRT(POSI,8) = 0     'MA
          TABPROCESOSCRT(POSI,9) = 0     'FREE
          TABPROCESOSCRT(POSI,10) = 0    'FREE

        RS.MOVENEXT
     LOOP  



  END IF



 ELSE



  Strsql = "     select B.CODOBJETO,B.DESCRI, "
  Strsql = Strsql & " a.* from csrelpathdescri a"
  Strsql = Strsql & "     inner join objetos b on (a.cobj = b.codobjeto)"
  Strsql = Strsql & " where "
  Strsql = Strsql & " B.APLIC = 'RG' AND "
  Strsql = Strsql & " A.MC7 = 0 "
  Strsql = Strsql & " ORDER BY B.PRIORIDAD DESC"

 '' response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then 
     POSI = 0
     TOPE = POSI
     DO WHILE NOT RS.EOF
        POSI = POSI + 1
        TOPE = POSI  

          TABPROCESOSCRT(POSI,0) = RS(0) 'COD.PROCESO
          TABPROCESOSCRT(POSI,1) = RS(1) 'NOM.PROCESO
          TABPROCESOSCRT(POSI,2) = 0     'CANT.ACTIVIDADES
          TABPROCESOSCRT(POSI,3) = 0     'FREE
          TABPROCESOSCRT(POSI,4) = 0     'MB
          TABPROCESOSCRT(POSI,5) = 0     'B
          TABPROCESOSCRT(POSI,6) = 0     'M
          TABPROCESOSCRT(POSI,7) = 0     'A
          TABPROCESOSCRT(POSI,8) = 0     'MA
          TABPROCESOSCRT(POSI,9) = 0     'FREE
          TABPROCESOSCRT(POSI,10) = 0    'FREE

        RS.MOVENEXT
     LOOP  



  END IF


 END IF

  MAINPROC = ""
  MAINACTI = ""
  MAINRIES = ""


  Strsql = "     select D.CODACCION,F.CODRIESGO, f.probabilidad, f.impacto, f.riesgobruto, f.riesgoresidual, f.riesgoresidualNV, "
  Strsql = Strsql & " a.* from csrelpathdescri a"
  Strsql = Strsql & "     inner join objetos b on (a.cobj = b.codobjeto)"

  Strsql = Strsql & "     inner join orrskcircuito w on (a.cobj = w.codobjeto AND W.ESTADO <> 4)"
  Strsql = Strsql & "     inner join objlnk x on (a.cobj = x.codobjeto and (x.ltipo = 'LRGVN' or x.ltipo = 'LRGCR'))"

  Strsql = Strsql & "     inner join objproced c on ( c.codobjeto = x.codobjetol and c.codproced = x.iden4 )"


  Strsql = Strsql & "     inner join objaccion d on ( C.codproced = d.codproced)  "
  Strsql = Strsql & "     LEFT join tipoaccalt e on (d.tipoaccion = e.tipoaccionalt)"

  Strsql = Strsql & "     INNER join oractivriesgo f on (d.codaccion = f.actividad )"
  Strsql = Strsql & " where d.critico = 'SI' and w.periodo = 0" & Speriodo & "  and f.periodo = w.periodo "

 IF SFTIPOSRCH = "PROCESO" THEN

  Strsql = Strsql & " AND ("
  Strsql = Strsql & " A.MC1 = " & SPROCESO & " OR "
  Strsql = Strsql & " A.MC2 = " & SPROCESO & " OR "
  Strsql = Strsql & " A.MC3 = " & SPROCESO & " OR "
  Strsql = Strsql & " A.MC4 = " & SPROCESO & " OR "
  Strsql = Strsql & " A.MC5 = " & SPROCESO & " OR "
  Strsql = Strsql & " A.MC6 = " & SPROCESO & " OR "
  Strsql = Strsql & " A.MC7 = " & SPROCESO & " OR "
  Strsql = Strsql & " A.COBJ = " & SPROCESO & "  "
  Strsql = Strsql & " ) "

 END IF


  Strsql = Strsql & " ORDER BY A.TSEC "

''  response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then 

     CANTREG = 0

     CANTACT = 0

     CANTACT4 = 0
     CANTACT5 = 0
     CANTACT6 = 0
     CANTACT7 = 0
     CANTACT8 = 0

     TCANTACT0 = 0
     TCANTACT4 = 0
     TCANTACT5 = 0
     TCANTACT6 = 0
     TCANTACT7 = 0
     TCANTACT8 = 0


     CANTACT4W = 0
     CANTACT5W = 0
     CANTACT6W = 0
     CANTACT7W = 0
     CANTACT8W = 0

               WKWRB = 0
               WKWRR = 0
               WKWRN = 0
               WKWIM = 0
               WKWPO = 0
               WKWPR = 0

     DO WHILE NOT RS.EOF

        WKIPROCD = ""
        WKIPROCN = 0

        FOR WKI=7 TO 14 

            WKIPROC = RS(WKI) & ""

            IF WKIPROC <> "" THEN
               WKIPROCN = RS(WKI + 8)
               EXIT FOR
            END IF 


        NEXT


        IF MAINPROC <> WKIPROCN THEN                     ' CAMBIA PROCESO

          IF MAINPROC <> "" THEN


                 WKWPR = CLNG(WKWPR / CANTREG)


                 SELECT CASE WKWPR

                     CASE 1

                     CANTACT4W = CANTACT4W + 1

                     CASE 2
                     CANTACT5W = CANTACT5W + 1

                     CASE 3
                     CANTACT6W = CANTACT6W + 1

                     CASE 4
                     CANTACT7W = CANTACT7W + 1

                     CASE 5
                     CANTACT8W = CANTACT8W + 1


                 END SELECT

                 CANTACT   = CANTACT + 1


            TABPROCESOSCRT(POSI,2) = CANTACT       'CANT.ACTIVIDADES
            TABPROCESOSCRT(POSI,4) = CANTACT4W     'MB
            TABPROCESOSCRT(POSI,5) = CANTACT5W     'B
            TABPROCESOSCRT(POSI,6) = CANTACT6W     'M
            TABPROCESOSCRT(POSI,7) = CANTACT7W     'A
            TABPROCESOSCRT(POSI,8) = CANTACT8W     'MA

               MAINACTI = RS(0) & RS(1)

               CANTREG  = 0

               WKWRB = 0
               WKWRR = 0
               WKWRN = 0
               WKWIM = 0
               WKWPO = 0
               WKWPR = 0

          END IF


            MAINPROC = WKIPROCN

            FOR WKI=1 TO TOPE 

                IF TABPROCESOSCRT(WKI,0) = WKIPROCN THEN
                   POSI = WKI
                   EXIT FOR
                END IF 

            NEXT


            CANTACT4W = 0
            CANTACT5W = 0
            CANTACT6W = 0
            CANTACT7W = 0
            CANTACT8W = 0

            CANTACT   = 0

    Celda = ""
    Celda = Celda & ofv.gencelda("","UT colspan=5 ","","","","","","-----------------------------------------------------","si")
    Fila2 = FILA2 & ofv.genrow("","","","","25","","",Celda,"si")

      END IF  


         IF MAINACTI <> (RS(0) & RS(1)) THEN                      ' CAMBIA ACTIVIDAD

           IF MAINACTI <> "" THEN

                 WKWPR = CLNG(WKWPR / CANTREG)


                 SELECT CASE WKWPR

                     CASE 1

                     CANTACT4W = CANTACT4W + 1

                     CASE 2
                     CANTACT5W = CANTACT5W + 1

                     CASE 3
                     CANTACT6W = CANTACT6W + 1

                     CASE 4
                     CANTACT7W = CANTACT7W + 1

                     CASE 5
                     CANTACT8W = CANTACT8W + 1


                 END SELECT

                 CANTACT   = CANTACT + 1

             END IF


               MAINACTI = RS(0) & RS(1)

               CANTREG  = 0

               WKWRB = 0
               WKWRR = 0
               WKWRN = 0
               WKWIM = 0
               WKWPO = 0
               WKWPR = 0



          END IF   

  


          CANTREG = CANTREG + 1

          WKWRB = WKWRB + RS(4)
          WKWRR = WKWRR + RS(5)
          WKWRN = WKWRN + RS(5)

          WKWIM = WKWIM + RS(3)
          WKWPO = WKWPO + RS(2)


  if SFILTRO  = "PRO" then

          WKWPR = WKWPR + RS(2)

  elseif SFILTRO  = "IMP" then

          WKWPR = WKWPR + RS(3)

  elseif SFILTRO  = "RBR" then

          if rs(4) < 6  then
             WKWPRW  = 1
          elseif rs(4) < 11  then
             WKWPRW  = 2
          elseif rs(4) < 16  then
             WKWPRW  = 3
          elseif rs(4) < 21  then
             WKWPRW  = 4
          elseif rs(4) > 20  then
             WKWPRW  = 5
          end if  

          WKWPR = WKWPR + WKWPRW

  elseif SFILTRO  = "RRE" then

          if rs(5) < 6  then
             WKWPRW  = 1
          elseif rs(5) < 11  then
             WKWPRW  = 2
          elseif rs(5) < 16  then
             WKWPRW  = 3
          elseif rs(5) < 21  then
             WKWPRW  = 4
          elseif rs(5) > 20  then
             WKWPRW  = 5
          end if  

          WKWPR = WKWPR + WKWPRW

  elseif SFILTRO  = "RNV" then

          if rs(6) < 6  then
             WKWPRW  = 1
          elseif rs(6) < 11  then
             WKWPRW  = 2
          elseif rs(6) < 16  then
             WKWPRW  = 3
          elseif rs(6) < 21  then
             WKWPRW  = 4
          elseif rs(6) > 20  then
             WKWPRW  = 5
          end if  

          WKWPR = WKWPR + WKWPRW


  end if




    Celda = ""
    Celda = Celda & ofv.gencelda("","UT ","","","","","",CANTACT,"si")
    Celda = Celda & ofv.gencelda("","UT ","","","","","",CANTREG,"si")
    Celda = Celda & ofv.gencelda("","UT ","","","","","",WKIPROC,"si")
    Celda = Celda & ofv.gencelda("","UT ","","","","","",RS(4),"si")
    Celda = Celda & ofv.gencelda("","UT","","","","","",RS(2),"si")
    Fila2 = FILA2 & ofv.genrow("","","","","25","","",Celda,"si")


        RS.MOVENEXT
     LOOP  
  END IF


     IF CANTREG > 0 THEN 

                 WKWPR = CLNG(WKWPR / CANTREG)


                 SELECT CASE WKWPR

                     CASE 1

                     CANTACT4W = CANTACT4W + 1

                     CASE 2
                     CANTACT5W = CANTACT5W + 1

                     CASE 3
                     CANTACT6W = CANTACT6W + 1

                     CASE 4
                     CANTACT7W = CANTACT7W + 1

                     CASE 5
                     CANTACT8W = CANTACT8W + 1


                 END SELECT

            CANTACT   = CANTACT + 1

            TABPROCESOSCRT(POSI,2) = CANTACT       'CANT.ACTIVIDADES
            TABPROCESOSCRT(POSI,4) = CANTACT4W     'MB
            TABPROCESOSCRT(POSI,5) = CANTACT5W     'B
            TABPROCESOSCRT(POSI,6) = CANTACT6W     'M
            TABPROCESOSCRT(POSI,7) = CANTACT7W     'A
            TABPROCESOSCRT(POSI,8) = CANTACT8W     'MA


     END IF

'''bgcolor='#ddddff'
''#F7CC88

       wkstylo1 = " style='background-color:#ddddff;color:#888888;border-style:solid;border-width:1;border-color:" & BCOLORG & ";' " 


'    sHTML = sHTML & "<CENTER>" & ofv.GenTabla("","aaa BORDER=1 STYLE='border-style:solid;border-width:3;border-color:#F0F0F0;' ","","80% ","","0","","-1","0","","",Fila2,"si") & "</CENTER>" 

     BCOLORG = "#AAAAAA"
     BCOLORD = "#6699CC"

    BTIT1 = "&nbsp;PROBABILIDAD&nbsp;"
    BTIT2 = "&nbsp;IMPACTO&nbsp;"
    BTIT3 = "&nbsp;RIESGO&nbsp;BRUTO&nbsp;"
    BTIT4 = "&nbsp;RIESGO&nbsp;RESIDUAL&nbsp;"
    BTIT5 = "&nbsp;RESIDUAL&nbsp;NUEVO&nbsp;"


  if SFILTRO  = "PRO" then

          WKTIT = BTIT1

     B1 = "N"
     B2 = "S"
     B3 = "S"
     B4 = "S"
     B5 = "S"
     

  elseif SFILTRO  = "IMP" then

          WKTIT = BTIT2

     B1 = "S"
     B2 = "N"
     B3 = "S"
     B4 = "S"
     B5 = "S"

  elseif SFILTRO  = "RBR" then

          WKTIT = BTIT3

     B1 = "S"
     B2 = "S"
     B3 = "N"
     B4 = "S"
     B5 = "S"

  elseif SFILTRO  = "RRE" then

          WKTIT = BTIT4

     B1 = "S"
     B2 = "S"
     B3 = "S"
     B4 = "N"
     B5 = "S"

  elseif SFILTRO  = "RNV" then

          WKTIT = BTIT5

     B1 = "S"
     B2 = "S"
     B3 = "S"
     B4 = "S"
     B5 = "N"


  end if

 IF SFTIPO = "PROCESO" THEN

     sInput = Speriododsc


 ELSE

    sform = "ORMATPER"

     Strsql = "Select *, (99999 - FINAL) AS XFFINAL From ORPERIODOS  where APLIC = '" & wkaplic & "' and "
     Strsql = Strsql & " ESTADO IN ('OK','CL')  " 
     Strsql = Strsql & " ORDER BY XFFINAL,INICIO, ESTADO DESC"
     sInput = ofv.GenerarCombo(sForm,"PERIODO",SPERIODO,Strsql,cn,0,1,"")

 END IF



    Celda = ""
    Celda = Celda & ofv.gencelda("","VTI style='color:#ffffff;border-style:solid;border-width:1;border-color:" & BCOLORG & ";' ","","25%","","","","PERIODO DE EVALUACION:","si")
    Celda = Celda & ofv.gencelda("","VTI " & wkstylo1,"","80%","","","",sInput,"si")

    saccion = "ORMATRIXRIESGO.ASP?PERIODO=&FILTRO=" & SFILTRO

    Fila = FILA & ofv.genrow("","AAA BGCOLOR='#888888' ","","","25","","",Celda,"si")
    Fila =  ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")
    sHTML = sHTML & "<CENTER>" & ofv.GenTabla("","aaa  STYLE='border-style:solid;border-width:3;border-color:" & BCOLORG & ";' ","","80% ","","0","","1","1","","",Fila,"si") & "</CENTER>" 

    Fila = ""

  WKTITG = "MATRIZ DE RIESGO POR" & WKTIT


    BTIT1 = "<A CLASS=CC HREF='ORMATRIXSITUACION.ASP?PERIODO=" & SPERIODO & "&FILTRO=PRO&PROCESO=" & sproceso & "&TIPO=" & sftipo & "' >" & BTIT1 & "</A>"
    BTIT2 = "<A CLASS=CC HREF='ORMATRIXSITUACION.ASP?PERIODO=" & SPERIODO & "&FILTRO=IMP&PROCESO=" & sproceso & "&TIPO=" & sftipo & "' >" & BTIT2 & "</A>"
    BTIT3 = "<A CLASS=CC HREF='ORMATRIXSITUACION.ASP?PERIODO=" & SPERIODO & "&FILTRO=RBR&PROCESO=" & sproceso & "&TIPO=" & sftipo & "' >" & BTIT3 & "</A>"
    BTIT4 = "<A CLASS=CC HREF='ORMATRIXSITUACION.ASP?PERIODO=" & SPERIODO & "&FILTRO=RRE&PROCESO=" & sproceso & "&TIPO=" & sftipo & "' >" & BTIT4 & "</A>"
    BTIT5 = "<A CLASS=CC HREF='ORMATRIXSITUACION.ASP?PERIODO=" & SPERIODO & "&FILTRO=RNV&PROCESO=" & sproceso & "&TIPO=" & sftipo & "' >" & BTIT5 & "</A>"


    Celda = ""
    Celda = Celda & ofv.gencelda("","VTI " & wkstylo1,"","55%","","","",WKTITG,"si")

    IF B1 = "S" THEN    Celda = Celda & ofv.gencelda("","VT  style='border-style:solid;border-width:1;border-color:" & BCOLORG & ";'  ","CENTER","15%","","","",BTIT1,"si")
    IF B2 = "S" THEN    Celda = Celda & ofv.gencelda("","VT  style='border-style:solid;border-width:1;border-color:" & BCOLORG & ";'  ","CENTER","15%","","","",BTIT2,"si")
    IF B3 = "S" THEN    Celda = Celda & ofv.gencelda("","VT  style='border-style:solid;border-width:1;border-color:" & BCOLORG & ";'  ","CENTER","15%","","","",BTIT3,"si")
    IF B4 = "S" THEN    Celda = Celda & ofv.gencelda("","VT  style='border-style:solid;border-width:1;border-color:" & BCOLORG & ";'  ","CENTER","15%","","","",BTIT4,"si")
    IF B5 = "S" THEN    Celda = Celda & ofv.gencelda("","VT  style='border-style:solid;border-width:1;border-color:" & BCOLORG & ";'  ","CENTER","15%","","","",BTIT5,"si")


    Fila = FILA & ofv.genrow("","AAA BGCOLOR='#F7EEDD' ","","","25","","",Celda,"si")
    sHTML = sHTML & "<CENTER>" & ofv.GenTabla("","aaa  STYLE='border-style:solid;border-width:3;border-color:" & BCOLORG & ";' ","","80% ","","0","","1","1","","",Fila,"si") & "</CENTER>" 


    Fila = ""


    Celda = ""
    Celda = Celda & ofv.gencelda("","VTI ROWSPAN=2 COLSPAN=3 " & wkstylo1,"","","","","","PROCESO","si")
    Celda = Celda & ofv.gencelda("","VT COLSPAN=5  " & wkstylo1,"CENTER","","","","",WKTIT,"si")
    Fila = FILA & ofv.genrow("","","","","25","","",Celda,"si")

    Celda = ""
    Celda = Celda & ofv.gencelda("","VT BGCOLOR='#00dd00' style='color:#ffffff;border-style:solid;border-width:1;border-color:" & BCOLORG & ";'  ","CENTER","","","","",TABPROCESOSCRT(0,4),"si")
    Celda = Celda & ofv.gencelda("","VT BGCOLOR='#6699cc' style='color:#ffffff;border-style:solid;border-width:1;border-color:" & BCOLORG & ";'  ","CENTER","","","","",TABPROCESOSCRT(0,5),"si")
    Celda = Celda & ofv.gencelda("","VT BGCOLOR='#F7DD88' style='color:#ffffff;border-style:solid;border-width:1;border-color:" & BCOLORG & ";'  ","CENTER","","","","",TABPROCESOSCRT(0,6),"si")
    Celda = Celda & ofv.gencelda("","VT BGCOLOR='#f7aa22' style='color:#ffffff;border-style:solid;border-width:1;border-color:" & BCOLORG & ";'  ","CENTER","","","","",TABPROCESOSCRT(0,7),"si")
    Celda = Celda & ofv.gencelda("","VT BGCOLOR='#dd0000' style='color:#ffffff;border-style:solid;border-width:1;border-color:" & BCOLORG & ";'  ","CENTER","","","","",TABPROCESOSCRT(0,8),"si")
    Fila = FILA & ofv.genrow("","","","","15","","",Celda,"si")

    POSI = 0


    FOR WKI=1 TO TOPE

      if TABPROCESOSCRT(WKI,2) > 0  then 

         POSI = POSI + 1

        Celda = ""
        Celda = Celda & ofv.gencelda("","UTI " & wkstylo1,"RIGHT","2%","","","",POSI & ".","si")
        Celda = Celda & ofv.gencelda("","UTGI bgcolor='" & xTrColor & "'  STYLE='COLOR:#000099;border-style:solid;border-width:1;border-color:" & BCOLORG & ";' ","","35%","","","",TABPROCESOSCRT(WKI,1),"si")
        Celda = Celda & ofv.gencelda("","ut " & wkstylo1,"center","2%","","","",TABPROCESOSCRT(WKI,2),"si")

        sHref     = "ORMATRIXIMPACTO.asp?PERIODO=" & SPERIODO & "&PROCESO=" & TABPROCESOSCRT(WKI,0) & "&FILTRO=PROCESO&FILSEL=" & SFILTRO & "&SUBFILTRO="


        CANTACT4 = "<b>" & TABPROCESOSCRT(WKI,4) & "</b>" 
        sHref4     = sHref & "1"
        sImage    = CANTACT4
        CANTACT4  = ofv.GenLink("","nt",sHref4,"_blank","","",sImage,"si")

        CANTACT5 = "<b>" & TABPROCESOSCRT(WKI,5) & "</b>" 
        sHref5     = sHref & "2"
        sImage    = CANTACT5
        CANTACT5  = ofv.GenLink("","nt",sHref5,"_blank","","",sImage,"si")

        CANTACT6 = "<b>" & TABPROCESOSCRT(WKI,6) & "</b>" 
        sHref6     = sHref & "3"
        sImage    = CANTACT6
        CANTACT6  = ofv.GenLink("","nt",sHref6,"_blank","","",sImage,"si")

        CANTACT7 = "<b>" & TABPROCESOSCRT(WKI,7) & "</b>" 
        sHref7     = sHref & "4"
        sImage    = CANTACT7
        CANTACT7  = ofv.GenLink("","nt",sHref7,"_blank","","",sImage,"si")

        CANTACT8 = "<b>" & TABPROCESOSCRT(WKI,8) & "</b>" 
        sHref8     = sHref & "5"
        sImage    = CANTACT8
        CANTACT8  = ofv.GenLink("","nt",sHref8,"_blank","","",sImage,"si")



        if TABPROCESOSCRT(WKI,4) = 0 then  CANTACT4 = "-"
        if TABPROCESOSCRT(WKI,5) = 0 then  CANTACT5 = "-"
        if TABPROCESOSCRT(WKI,6) = 0 then  CANTACT6 = "-"
        if TABPROCESOSCRT(WKI,7) = 0 then  CANTACT7 = "-"
        if TABPROCESOSCRT(WKI,8) = 0 then  CANTACT8 = "-"




        TCANTACT0 = TCANTACT0 + TABPROCESOSCRT(WKI,2)

        TCANTACT4 = TCANTACT4 + TABPROCESOSCRT(WKI,4)
        TCANTACT5 = TCANTACT5 + TABPROCESOSCRT(WKI,5)
        TCANTACT6 = TCANTACT6 + TABPROCESOSCRT(WKI,6)
        TCANTACT7 = TCANTACT7 + TABPROCESOSCRT(WKI,7)
        TCANTACT8 = TCANTACT8 + TABPROCESOSCRT(WKI,8)



        Celda = Celda & ofv.gencelda("","UTg bgcolor='" & xTrColor & "' STYLE='border-style:solid;border-width:1;border-color:" & BCOLORD & ";' ","CENTER","10%","","","",CANTACT4,"si")
        Celda = Celda & ofv.gencelda("","UTg bgcolor='" & xTrColor & "' STYLE='border-style:solid;border-width:1;border-color:" & BCOLORD & ";' ","CENTER","10%","","","",CANTACT5,"si")
        Celda = Celda & ofv.gencelda("","UTg bgcolor='" & xTrColor & "' STYLE='border-style:solid;border-width:1;border-color:" & BCOLORD & ";' ","CENTER","10%","","","",CANTACT6,"si")
        Celda = Celda & ofv.gencelda("","UTg bgcolor='" & xTrColor & "' STYLE='border-style:solid;border-width:1;border-color:" & BCOLORD & ";' ","CENTER","10%","","","",CANTACT7,"si")
        Celda = Celda & ofv.gencelda("","UTg bgcolor='" & xTrColor & "' STYLE='border-style:solid;border-width:1;border-color:" & BCOLORD & ";' ","CENTER","10%","","","",CANTACT8,"si")
        Fila = FILA & ofv.genrow("","","","","50","","",Celda,"si")


 
        if xTrColor = "#FFFFFF" then
           xTrColor = "#F0F0F0" 
        ELSE
           xTrColor = "#FFFFFF" 
        END IF


      end if 




    NEXT

        if TCANTACT4 = 0 then  TCANTACT4 = "-"
        if TCANTACT5 = 0 then  TCANTACT5 = "-"
        if TCANTACT6 = 0 then  TCANTACT6 = "-"
        if TCANTACT7 = 0 then  TCANTACT7 = "-"
        if TCANTACT8 = 0 then  TCANTACT8 = "-"

        Celda = ""
        Celda = Celda & ofv.gencelda("","VTI COLSPAN=2 " & wkstylo1,"","","","","","TOTAL GENERAL","si")
        Celda = Celda & ofv.gencelda("","ut " & wkstylo1,"center","2%","","","",TCANTACT0,"si")

        xTrColor = "#F7EEDD" 

        Celda = Celda & ofv.gencelda("","VT bgcolor='" & xTrColor & "' STYLE='border-style:solid;border-width:1;border-color:" & BCOLORD & ";' ","CENTER","10%","","","",TCANTACT4,"si")
        Celda = Celda & ofv.gencelda("","VT bgcolor='" & xTrColor & "' STYLE='border-style:solid;border-width:1;border-color:" & BCOLORD & ";' ","CENTER","10%","","","",TCANTACT5,"si")
        Celda = Celda & ofv.gencelda("","VT bgcolor='" & xTrColor & "' STYLE='border-style:solid;border-width:1;border-color:" & BCOLORD & ";' ","CENTER","10%","","","",TCANTACT6,"si")
        Celda = Celda & ofv.gencelda("","VT bgcolor='" & xTrColor & "' STYLE='border-style:solid;border-width:1;border-color:" & BCOLORD & ";' ","CENTER","10%","","","",TCANTACT7,"si")
        Celda = Celda & ofv.gencelda("","VT bgcolor='" & xTrColor & "' STYLE='border-style:solid;border-width:1;border-color:" & BCOLORD & ";' ","CENTER","10%","","","",TCANTACT8,"si")
        Fila = FILA & ofv.genrow("","","","","50","","",Celda,"si")


    sHTML = sHTML & "<CENTER>" & ofv.GenTabla("","aaa  STYLE='border-style:solid;border-width:3;border-color:" & BCOLORG & ";' ","","80% ","","0","","1","1","","",Fila,"si") & "</CENTER>" 

  'Botonera
  Agregar = ""
  Buscar  = ""
  Celda   = ""
  
  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>