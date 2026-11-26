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

  Speriodo = REQUEST.QUERYSTRING("periodo")


  if Speriodo = "" then Speriodo = REQUEST.FORM("PERIODO")



  if Speriodo = "" then Speriodo = session("ORPERIODO")


  if SFILTRO  = "" then
     SFILTRO  = "PRO"

  end if

  DIM TABPROCESOSCRT(200,20)
  DIM TABPROCESOSCRTM(200,20)


  sHTML = ofv.FormHeader(session("FormN"))


  ofv.ObtenerAtributos Session("Form"),""
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

     if Speriodo = "" then

        Strsql = "Select * From ORPERIODOS  where APLIC = '" & wkaplic & "' and ESTADO IN ('OK','CL')  and RGVALEAUTOEV = 'S' "
        Strsql = Strsql & " ORDER BY FINAL,INICIO, ESTADO DESC"

        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if NOT rs.eof then
           session("ORPERIODO")     = RS(0)
           session("ORPERIODODESC") = RS(1)
           session("ORPERIODOMAP")  = RS(9)
           Speriodo                 = session("ORPERIODO")
           Speriododsc              = session("ORPERIODODESC")
           Speriodomap              = session("ORPERIODOMAP")

        end if

      else

        Strsql = "Select * From ORPERIODOS  where OBJID = 0" & Speriodo
        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if NOT rs.eof then
           Speriododsc              = RS(1)
           Speriodomap              = RS(9)
        end if

      end if

  Fila = ""
  xTrColor = "#FFFFFF"

  sImageva    = "<font face='corflaticon' size=2 color='#990000'>&#0110;</font>"
  sImagevm    = "<font face='corflaticon' size=2 color='#f7aa22'>&#0110;</font>"
  sImagevb    = "<font face='corflaticon' size=2 color='#009900'>&#0110;</font>"
  sImagevn    = "<font face='corflaticon' size=2 color='#dddddd'>&#0110;</font>"

  wkstylof = "aaa style='font-family:arial;font-size:6pt;color:#888888;' "

''''b.codvinobj = 'VN' AND




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


  DIM TABMAPA(10,5)
  DIM TABMAPAM(10,5)

  DIM TABPROBABILIDAD(10,5)
  DIM TABIMPACTO(10,5)
  DIM TABRIESGOS(50,6)
  DIM TABRIESGOSRES(50,6)
  DIM TABRANRIESGOS(10,5)

  TTABPROB      = 0
  TTABIMPA      = 0
  TTABRIESGO    = 0
  TTABRIESGORES = 0

  FOR II=1 TO 10

      TABMAPA(II,01) = 0
      TABMAPA(II,02) = 0
      TABMAPA(II,03) = 0
      TABMAPA(II,04) = 0
      TABMAPA(II,05) = 0

      TABMAPAM(II,01) = 0
      TABMAPAM(II,02) = 0
      TABMAPAM(II,03) = 0
      TABMAPAM(II,04) = 0
      TABMAPAM(II,05) = 0


  NEXT


          TABPROCESOSCRT(0,00) = SFILTRO 'COD.PROCESO
          TABPROCESOSCRT(0,01) = "PRO"   'NOM.PROCESO
          TABPROCESOSCRT(0,02) = 0       'CANT.ACTIVIDADES
          TABPROCESOSCRT(0,03) = 0       'FREE

          TABPROCESOSCRT(0,04) = "NO"     'B
          TABPROCESOSCRT(0,05) = "NO"     'B
          TABPROCESOSCRT(0,06) = "NO"     'M
          TABPROCESOSCRT(0,07) = "NO"     'A
          TABPROCESOSCRT(0,08) = "NO"     'MA
          TABPROCESOSCRT(0,09) = "NO"     'FREE
          TABPROCESOSCRT(0,10) = "NO"     'FREE
          TABPROCESOSCRT(0,11) = "NO"     'FREE
          TABPROCESOSCRT(0,12) = "NO"     'FREE
          TABPROCESOSCRT(0,13) = "NO"     'FREE

          TABPROCESOSCRT(200,04) = 0     'B
          TABPROCESOSCRT(200,05) = 0     'B
          TABPROCESOSCRT(200,06) = 0     'M
          TABPROCESOSCRT(200,07) = 0     'A
          TABPROCESOSCRT(200,08) = 0     'MA
          TABPROCESOSCRT(200,09) = 0     'FREE
          TABPROCESOSCRT(200,10) = 0     'FREE
          TABPROCESOSCRT(200,11) = 0     'FREE
          TABPROCESOSCRT(200,12) = 0     'FREE
          TABPROCESOSCRT(200,13) = 0     'FREE

          TABPROCESOSCRTM(200,04) = 0     'B
          TABPROCESOSCRTM(200,05) = 0     'B
          TABPROCESOSCRTM(200,06) = 0     'M
          TABPROCESOSCRTM(200,07) = 0     'A
          TABPROCESOSCRTM(200,08) = 0     'MA
          TABPROCESOSCRTM(200,09) = 0     'FREE
          TABPROCESOSCRTM(200,10) = 0     'FREE
          TABPROCESOSCRTM(200,11) = 0     'FREE
          TABPROCESOSCRTM(200,12) = 0     'FREE
          TABPROCESOSCRTM(200,13) = 0     'FREE



  if SFILTRO  = "PRO" or SFILTRO  = "RRE" OR SFILTRO  = "RNV" OR SFILTRO  = "RBR" then

  Strsql = "Select * From ORNIVPROBOCURRENCIA ORDER BY VALOR"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     POSI = 0
     DO WHILE NOT RS.EOF
        POSI = POSI + 1

          TABPROBABILIDAD(POSI,0) = RS(0) 'CODIGO
          TABPROBABILIDAD(POSI,1) = RS(1) 'VALOR
          TABPROBABILIDAD(POSI,2) = RS(2) 'DESCRIPCION
          TABPROBABILIDAD(POSI,3) = RS(3) 'PORCENTAJE
          TABPROBABILIDAD(POSI,5) = RS(6) 'COLOR

          TABPROCESOSCRT(0,POSI + 3) = RS(2)     'DESCRIPCION

          TTABPROB = POSI

        RS.MOVENEXT
     LOOP



  END IF

  end if


  if SFILTRO  = "IMP" or SFILTRO  = "RRE" OR SFILTRO  = "RNV" OR SFILTRO  = "RBR" then

  Strsql = "Select * From ORNIVIMPACTO ORDER BY VALOR"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     POSI = 0
     DO WHILE NOT RS.EOF
        POSI = POSI + 1


          TABIMPACTO(POSI,0) = RS(0) 'CODIGO
          TABIMPACTO(POSI,1) = RS(1) 'VALOR
          TABIMPACTO(POSI,2) = RS(2) 'DESCRIPCION
          TABIMPACTO(POSI,3) = RS(3) 'MONTODESDE
          TABIMPACTO(POSI,4) = RS(4) 'MONTOHASTA
          TABIMPACTO(POSI,5) = RS(6) 'COLOR

          TABPROCESOSCRT(0,POSI + 3) = RS(2)     'DESCRIPCION

          TTABIMPA = POSI


        RS.MOVENEXT
     LOOP



  END IF

  end if

  if SFILTRO  = "RBR" OR SFILTRO  = "RNV" OR SFILTRO  = "RRE" then

  Strsql = "Select * From ORRANGORIESGO ORDER BY VALOR"

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     POSI = 0
     DO WHILE NOT RS.EOF
        POSI = POSI + 1

          TABRANRIESGOS(POSI,0) = RS(0) 'CODIGO
          TABRANRIESGOS(POSI,1) = RS(1) 'VALOR
          TABRANRIESGOS(POSI,2) = RS(2)  'DESCRIPCION
          TABRANRIESGOS(POSI,3) = RS(3) 'COLOR
          TABRANRIESGOS(POSI,4) = 0     'CANTIDAD


          TTABRANRIESGO = POSI


        RS.MOVENEXT

     LOOP



  END IF


  if SFILTRO  = "RBR" then

  Strsql = "Select * From ORNIVRIESGOBRUTO ORDER BY VALOR DESC"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     POSI = 0
     DO WHILE NOT RS.EOF
        POSI = POSI + 1

          TABRIESGOS(POSI,0) = RS(0) 'CODIGO
          TABRIESGOS(POSI,1) = RS(1) 'VALOR

          TABRIESGOS(POSI,3) = RS(3) 'PROBABILIDAD
          TABRIESGOS(POSI,4) = RS(4) 'IMPACTO

          FOR II=1 TO TTABRANRIESGO

              IF RS(2) =  TABRANRIESGOS(II,0) THEN
                 TABRIESGOS(POSI,2)    = TABRANRIESGOS(II,2)  'DESCRIPCION
                 TABRIESGOS(POSI,5)    = TABRANRIESGOS(II,3) 'COLOR
                 TABRIESGOS(POSI,6)    = II                  'RANGO
                 TABRANRIESGOS(ii,4)   = 1                   'CANTIDAD
                 EXIT FOR
              END IF

          NEXT

          TTABRIESGO = POSI

        RS.MOVENEXT
     LOOP



  END IF



  elseif SFILTRO  = "RNV" OR SFILTRO  = "RRE" then

  Strsql = "Select * From ORNIVRIESGO ORDER BY VALORDESDE "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     POSI = 0
     DO WHILE NOT RS.EOF
        POSI = POSI + 1

          TABRIESGOSRES(POSI,0) = RS(0) 'CODIGO

          TABRIESGOSRES(POSI,2) = CLNG(CDBL(RS(2)) * 100) 'VALOR DESDE
          TABRIESGOSRES(POSI,3) = CLNG(CDBL(RS(3)) * 100) 'VALOR HASTA

          FOR II=1 TO TTABRANRIESGO

              IF RS(1) =  TABRANRIESGOS(II,0) THEN
                 TABRIESGOSRES(POSI,1)    = TABRANRIESGOS(II,2) 'DESCRIPCION
                 TABRIESGOSRES(POSI,5)    = TABRANRIESGOS(II,3) 'COLOR
                 TABRIESGOSRES(POSI,6)    = II                  'RANGO
                 TABRANRIESGOS(ii,4)      = 1                   'CANTIDAD
                 EXIT FOR
              END IF

          NEXT

          TTABRIESGORES = POSI

        RS.MOVENEXT
     LOOP



  END IF


  end if

  ' RESPONSE.WRITE TTABRANRIESGO
  FOR II=1 TO TTABRANRIESGO
      TABPROCESOSCRT(0,II + 3) = "NO"
      IF TABRANRIESGOS(ii,4) = 1 THEN
          TABPROCESOSCRT(0,II + 3) = TABRANRIESGOS(ii,2)     'DESCRIPCION
      END IF

  NEXT



 END IF


  Strsql = "     select B.CODOBJETO,B.DESCRI, "
  Strsql = Strsql & " a.* from csrelpathdescri a"
  Strsql = Strsql & "     inner join objetos b on (a.cobj = b.codobjeto)"
  Strsql = Strsql & " where "
  Strsql = Strsql & " B.APLIC = 'RG' AND "
  Strsql = Strsql & " A.MC7 = 0" & Speriodomap
  Strsql = Strsql & " ORDER BY B.PRIORIDAD "

 '' response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     POSI = 0
     TOPE = POSI
     DO WHILE NOT RS.EOF
        POSI = POSI + 1
        TOPE = POSI

          TABPROCESOSCRT(POSI,00) = RS(0) 'COD.PROCESO
          TABPROCESOSCRT(POSI,01) = RS(1) 'NOM.PROCESO
          TABPROCESOSCRT(POSI,02) = 0     'CANT.ACTIVIDADES
          TABPROCESOSCRT(POSI,03) = 0     'FREE

          TABPROCESOSCRT(POSI,04) = 0     'MB
          TABPROCESOSCRT(POSI,05) = 0     'B
          TABPROCESOSCRT(POSI,06) = 0     'M
          TABPROCESOSCRT(POSI,07) = 0     'A
          TABPROCESOSCRT(POSI,08) = 0     'MA
          TABPROCESOSCRT(POSI,09) = 0     'FREE
          TABPROCESOSCRT(POSI,10) = 0     'FREE
          TABPROCESOSCRT(POSI,11) = 0     'FREE
          TABPROCESOSCRT(POSI,12) = 0     'FREE
          TABPROCESOSCRT(POSI,13) = 0     'FREE

          TABPROCESOSCRTM(POSI,04) = 0     'MB
          TABPROCESOSCRTM(POSI,05) = 0     'B
          TABPROCESOSCRTM(POSI,06) = 0     'M
          TABPROCESOSCRTM(POSI,07) = 0     'A
          TABPROCESOSCRTM(POSI,08) = 0     'MA
          TABPROCESOSCRTM(POSI,09) = 0     'FREE
          TABPROCESOSCRTM(POSI,10) = 0     'FREE
          TABPROCESOSCRTM(POSI,11) = 0     'FREE
          TABPROCESOSCRTM(POSI,12) = 0     'FREE
          TABPROCESOSCRTM(POSI,13) = 0     'FREE


        RS.MOVENEXT
     LOOP



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
  Strsql = Strsql & " where d.critico = 'SI' and w.periodo = 0" & Speriodo & "  and f.periodo = w.periodo  and f.probabilidad > 0 and f.impacto > 0 "
  Strsql = Strsql & " ORDER BY A.TSEC "

  'response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then

     CANTREG = 0
     CANTREG2 = 0
     CANTREG3 = 0

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

            FOR II=1 TO 10

                TABMAPA(II,01) = 0
                TABMAPA(II,02) = 0
                TABMAPA(II,03) = 0
                TABMAPA(II,04) = 0
                TABMAPA(II,05) = 0

                TABMAPAM(II,01) = 0
                TABMAPAM(II,02) = 0
                TABMAPAM(II,03) = 0
                TABMAPAM(II,04) = 0
                TABMAPAM(II,05) = 0


            NEXT

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


                 TABPROCESOSCRT(POSI,2) = CANTACT       'CANT.ACTIVIDADES

                 FOR II=1 TO 10
                     TABPROCESOSCRT(POSI,II + 3) =  TABMAPA(II,01)
                     TABPROCESOSCRTM(POSI,II + 3) =  TABMAPAM(II,01)
                 NEXT



            END IF


            MAINPROC = WKIPROCN

            FOR WKI=1 TO TOPE

                IF TABPROCESOSCRT(WKI,0) = WKIPROCN THEN
                   POSI = WKI
                   EXIT FOR
                END IF

            NEXT

            FOR II=1 TO 10

                TABMAPA(II,01) = 0
                TABMAPA(II,02) = 0
                TABMAPA(II,03) = 0
                TABMAPA(II,04) = 0
                TABMAPA(II,05) = 0

                TABMAPAM(II,01) = 0
                TABMAPAM(II,02) = 0
                TABMAPAM(II,03) = 0
                TABMAPAM(II,04) = 0
                TABMAPAM(II,05) = 0



            NEXT

            CANTACT   = 0

            Celda = ""
            Celda = Celda & ofv.gencelda("","UT colspan=5 ","","","","","","-----------------------------------------------------","si")
            Fila2 = FILA2 & ofv.genrow("","","","","25","","",Celda,"si")

        END IF


          CANTACT   = CANTACT + 1
          wspmontoF      = 0

          if SFILTRO  = "PRO" then '2

             FOR II=1 TO TTABPROB
                 if TABPROBABILIDAD(ii,0) = RS(2) then
                    TABMAPA(II,01)        = TABMAPA(II,01) + 1
                    TABMAPAM(II,01)       = TABMAPAM(II,01) + wspmontoF
                    exit for
                 end if
             NEXT


          elseif SFILTRO  = "IMP" then '3

              FOR II=1 TO TTABIMPA
                  if TABIMPACTO(ii,0) = RS(3) then
                     TABMAPA(II,01)   = TABMAPA(II,01) + 1
                     TABMAPAM(II,01)  = TABMAPAM(II,01) + wspmontoF
                     exit for
                  end if
              NEXT


          elseif SFILTRO  = "RBR" then '4

              IITP=0
              IITI=0

              FOR IITP=1 TO TTABPROB
                  if TABPROBABILIDAD(IITP,0) = RS(2) then
                     exit for
                  end if
              NEXT


              FOR IITI=1 TO TTABIMPA
                  if TABIMPACTO(IITI,0) = RS(3) then
                     exit for
                  end if
              NEXT

              CANTREG2 = CANTREG2 + 1



              FOR II=1 TO TTABRIESGO

                  IF TABRIESGOS(II,3) =  RS(2) AND TABRIESGOS(II,4) =  RS(3) THEN
                     TABMAPA(TABRIESGOS(II,6),01)  = TABMAPA(TABRIESGOS(II,6),01) + 1

                     wspmontoF      = 0

                     IF IITP > 0 AND IITI > 0  THEN

                        wspDATA       = CINT("0" & TABPROBABILIDAD(IITP,1)) * CINT("0" & TABIMPACTO(IITI,1))

                        wspmonto      = clng((Clng("0" & TABIMPACTO(IITI,3)) + Clng("0" & TABIMPACTO(IITI,4))) / 2)    'tabla impacto     3 Y 4
                        wspporc       = CINT("0" & TABPROBABILIDAD(IITP,3))                                            'tabla probabilidad  3
                        wspmontobruto = clng(wspmonto * wspporc / 100)
                        wspmontoF     = wspmontobruto
                        TABMAPAM(TABRIESGOS(II,6),01) = TABMAPAM(TABRIESGOS(II,6),01) + wspmontoF

                     CANTREG3 = CANTREG3 + 1

                     END IF



                     EXIT FOR
                  END IF

              NEXT


          elseif SFILTRO  = "RRE" then '5

              IITP=0
              IITI=0

              FOR IITP=1 TO TTABPROB
                  if TABPROBABILIDAD(IITP,0) = RS(2) then
                     exit for
                  end if
              NEXT


              FOR IITI=1 TO TTABIMPA
                  if TABIMPACTO(IITI,0) = RS(3) then
                     exit for
                  end if
              NEXT



              FOR II=1 TO TTABRIESGORES


           '''       RESPONSE.WRITE " TBD:" & TABRIESGOSRES(II,2) & "-TBH:" & TABRIESGOSRES(II,3) & "-RG:" & CLNG(CDBL(RS(5)) * 100) & "-<BR>"
              ''''    IF NOT (CDBL(TABRIESGOSRES(II,2)) <  CDBL(REPLACE("" & RS(5),",","."))) AND NOT ( CDBL(TABRIESGOSRES(II,3)) >  CDBL(REPLACE("" & RS(5),",","."))) THEN

                  IF CLNG(CDBL(RS(5)) * 100) < TABRIESGOSRES(II,2)   OR CLNG(CDBL(RS(5)) * 100) >  TABRIESGOSRES(II,3)  THEN

                  ELSE

                '''  RESPONSE.WRITE "--OK-- TBD:" & TABRIESGOSRES(II,2) & "-TBH:" & TABRIESGOSRES(II,3) & "-RG:" & CLNG(CDBL(RS(5)) * 100) & "-<BR>"
                     TABMAPA(TABRIESGOSRES(II,6),01) = TABMAPA(TABRIESGOSRES(II,6),01) + 1

                     wspmontoF      = 0

                     IF IITP > 0 AND IITI > 0 and CDBL(RS(5)) > 0 THEN

                        wspDATA       = CINT("0" & TABPROBABILIDAD(IITP,1)) * CINT("0" & TABIMPACTO(IITI,1))

                        wspmonto      = clng((Clng("0" & TABIMPACTO(IITI,3)) + Clng("0" & TABIMPACTO(IITI,4))) / 2)    'tabla impacto     3 Y 4
                        wspporc       = CINT("0" & TABPROBABILIDAD(IITP,3))                                            'tabla probabilidad  3
                        wspmontobruto = clng(wspmonto * wspporc / 100)
                        wspDATM       = cdbl( round( (wspDATA * 100) / round( CDBL(RS(5)) * 100 ,0)  ,2) )
                      ''''  response.write  " rre:" & CDBL(RS(5)) & "-data:" &  wspDATA & "-contr:" & wspDATM & "<br>"


                        wspmontoF     = clng(wspmontobruto / wspDATM)
                        TABMAPAM(TABRIESGOSRES(II,6),01) = TABMAPAM(TABRIESGOSRES(II,6),01) + wspmontoF

                     END IF


                     EXIT FOR
                  END IF

              NEXT


          elseif SFILTRO  = "RNV" then '6

              IITP=0
              IITI=0

              FOR IITP=1 TO TTABPROB
                  if TABPROBABILIDAD(IITP,0) = RS(2) then
                     exit for
                  end if
              NEXT


              FOR IITI=1 TO TTABIMPA
                  if TABIMPACTO(IITI,0) = RS(3) then
                     exit for
                  end if
              NEXT


              FOR II=1 TO TTABRIESGORES

                  IF CLNG(CDBL(RS(6)) * 100) < TABRIESGOSRES(II,2)   OR CLNG(CDBL(RS(6)) * 100) >  TABRIESGOSRES(II,3)  THEN

                  ELSE
                     TABMAPA(TABRIESGOSRES(II,6),01) = TABMAPA(TABRIESGOSRES(II,6),01) + 1

                     wspmontoF      = 0

                     IF IITP > 0 AND IITI > 0 and CDBL(RS(6)) > 0 THEN

                        wspDATA       = CINT("0" & TABPROBABILIDAD(IITP,1)) * CINT("0" & TABIMPACTO(IITI,1))

                        wspmonto      = clng((Clng("0" & TABIMPACTO(IITI,3)) + Clng("0" & TABIMPACTO(IITI,4))) / 2)    'tabla impacto     3 Y 4
                        wspporc       = CINT("0" & TABPROBABILIDAD(IITP,3))                                            'tabla probabilidad  3
                        wspmontobruto = clng(wspmonto * wspporc / 100)
                        wspDATM       = cdbl( round( (wspDATA * 100) / round( CDBL(RS(6)) * 100 ,0)  ,2) )
                      ''''  response.write  " rre:" & CDBL(RS(5)) & "-data:" &  wspDATA & "-contr:" & wspDATM & "<br>"


                        wspmontoF     = clng(wspmontobruto / wspDATM)
                        TABMAPAM(TABRIESGOSRES(II,6),01) = TABMAPAM(TABRIESGOSRES(II,6),01) + wspmontoF

                     END IF

                     EXIT FOR
                  END IF

              NEXT

          end if




          Celda = ""
          Celda = Celda & ofv.gencelda("","UT ","","","","","",CANTACT,"si")
          Celda = Celda & ofv.gencelda("","UT ","","","","","",CANTREG,"si")
          Celda = Celda & ofv.gencelda("","UT ","","","","","",WKIPROC,"si")
          Celda = Celda & ofv.gencelda("","UT ","","","","","",RS(4),"si")
          Celda = Celda & ofv.gencelda("","UT","","","","","",RS(2),"si")
          Fila2 = FILA2 & ofv.genrow("","","","","25","","",Celda,"si")


          CANTREG = CANTREG + 1


          RS.MOVENEXT
        LOOP

        TABPROCESOSCRT(POSI,2) = CANTACT         'CANT.ACTIVIDADES

        FOR II=1 TO 10
            TABPROCESOSCRT(POSI,II + 3)  =  TABMAPA(II,01)
            TABPROCESOSCRTM(POSI,II + 3) =  TABMAPAM(II,01)
        NEXT


     END IF







'''bgcolor='#ddddff'
''#F7CC88

       wkstylo1 = " style='background-color:#ddddff;color:#888888;border-style:solid;border-width:1;border-color:" & BCOLORG & ";' "


'    sHTML = sHTML & "<CENTER>" & ofv.GenTabla("","aaa BORDER=1 STYLE='border-style:solid;border-width:3;border-color:#F0F0F0;' ","","80% ","","0","","-1","0","","",Fila2,"si") & "</CENTER>"

     BCOLORG = "#AAAAAA"
     BCOLORD = "#6699CC"

    BTIT1 = "&nbsp;PROBABILIDAD&nbsp;"
    BTIT2 = "&nbsp;IMPACTO&nbsp;"
    BTIT3 = "&nbsp;RIESGO&nbsp;INHERENTE&nbsp;"
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

    sform = "ORMATPER"

     Strsql = "Select *, (99999 - FINAL) AS XFFINAL From ORPERIODOS  where APLIC = '" & wkaplic & "' and "
     Strsql = Strsql & " ESTADO IN ('OK','CL')  "
     Strsql = Strsql & " ORDER BY XFFINAL,INICIO, ESTADO DESC"
     sInput = ofv.GenerarCombo(sForm,"PERIODO",SPERIODO,Strsql,cn,0,1,"")



    Celda = ""
    Celda = Celda & ofv.gencelda("","VTI style='color:#ffffff;border-style:solid;border-width:1;border-color:" & BCOLORG & ";' ","","25%","","","","PERIODO DE EVALUACION:","si")
    Celda = Celda & ofv.gencelda("","VTI " & wkstylo1,"","80%","","","",sInput,"si")
  '''''  Celda = Celda & ofv.GenCelda("","ut  ","RIGHT","20%","","","",sInput,"si")

    saccion = "ORMATRIXRIESGO.ASP?PERIODO=&FILTRO=" & SFILTRO

    Fila = FILA & ofv.genrow("","AAA BGCOLOR='#888888' ","","","25","","",Celda,"si")
    Fila =  ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")
    sHTML = sHTML & "<CENTER>" & ofv.GenTabla("","aaa  STYLE='border-style:solid;border-width:3;border-color:" & BCOLORG & ";' ","","80% ","","0","","1","1","","",Fila,"si") & "</CENTER>"

    Fila = ""

  WKTITG = "MATRIZ DE RIESGO POR" & WKTIT


    BTIT1 = "<A CLASS=CC HREF='ORMATRIXRIESGO.ASP?PERIODO=" & SPERIODO & "&FILTRO=PRO' >" & BTIT1 & "</A>"
    BTIT2 = "<A CLASS=CC HREF='ORMATRIXRIESGO.ASP?PERIODO=" & SPERIODO & "&FILTRO=IMP' >" & BTIT2 & "</A>"
    BTIT3 = "<A CLASS=CC HREF='ORMATRIXRIESGO.ASP?PERIODO=" & SPERIODO & "&FILTRO=RBR' >" & BTIT3 & "</A>"
    BTIT4 = "<A CLASS=CC HREF='ORMATRIXRIESGO.ASP?PERIODO=" & SPERIODO & "&FILTRO=RRE' >" & BTIT4 & "</A>"
    BTIT5 = "<A CLASS=CC HREF='ORMATRIXRIESGO.ASP?PERIODO=" & SPERIODO & "&FILTRO=RNV' >" & BTIT5 & "</A>"


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

    TOPECELDA = 0

              WSPCOLOR = "#DDDDDD"

    FOR II=1 TO 10


        IF TABPROCESOSCRT(0,II + 3) <> "NO" THEN
           TOPECELDA = TOPECELDA + 1

           if SFILTRO  = "PRO" then

              WSPCOLOR = TABPROBABILIDAD(II,5)

           elseif SFILTRO  = "IMP" then

              WSPCOLOR = TABIMPACTO(II,5)

           elseif SFILTRO  = "RBR" OR SFILTRO  = "RRE" OR SFILTRO  = "RNV" then

              WSPCOLOR = TABRANRIESGOS(II,3)

           end if


           Celda = Celda & ofv.gencelda("","VT style='background-color:" & WSPCOLOR & ";color:#000000;border-style:solid;border-width:1;border-color:" & BCOLORG & ";'  ","CENTER","","","","",TABPROCESOSCRT(0,II + 3),"si")
        END IF
    NEXT


    FilaWK = ofv.genrow("","","","","15","","",Celda,"si")




    Celda = ""
    Celda = Celda & ofv.gencelda("","VTI ROWSPAN=2 COLSPAN=3 " & wkstylo1,"","","","","","PROCESO","si")
    Celda = Celda & ofv.gencelda("","VT COLSPAN=" & TOPECELDA & "  " & wkstylo1,"CENTER","","","","",WKTIT,"si")
    Fila = FILA & ofv.genrow("","","","","25","","",Celda,"si")

    Fila = FILA & FilaWK

    POSI = 0

    WFTOTALPESOS = 0


    FOR WKI=1 TO TOPE

      if TABPROCESOSCRT(WKI,2) > 0  then

         POSI = POSI + 1

        Celda = ""
        Celda = Celda & ofv.gencelda("","UTI " & wkstylo1,"RIGHT","2%","","","",POSI & ".","si")
        Celda = Celda & ofv.gencelda("","UTGI bgcolor='" & xTrColor & "'  STYLE='COLOR:#000099;border-style:solid;border-width:1;border-color:" & BCOLORG & ";' ","","35%","","","",TABPROCESOSCRT(WKI,1),"si")
        Celda = Celda & ofv.gencelda("","ut " & wkstylo1,"center","2%","","","",TABPROCESOSCRT(WKI,2),"si")

        sHref     = "ORMATRIXIMPACTO.asp?PERIODO=" & SPERIODO & "&PROCESO=" & TABPROCESOSCRT(WKI,0) & "&FILTRO=PROCESO&FILSEL=" & SFILTRO & "&SUBFILTRO="


        FOR II=1 TO 10

            IF TABPROCESOSCRT(0,II + 3) <> "NO" THEN

               TABPROCESOSCRT(200,II+3) = TABPROCESOSCRT(200,II+3) + TABPROCESOSCRT(WKI,II + 3)
               TABPROCESOSCRTM(200,II+3) = TABPROCESOSCRTM(200,II+3) + TABPROCESOSCRTM(WKI,II + 3)

               if SFILTRO  = "RBR" OR SFILTRO  = "RRE" OR SFILTRO  = "RNV" then

                  CANTACTW = "<b><FONT COLOR='#0000EE' >" & TABPROCESOSCRT(WKI,II + 3) & "</FONT><BR><HR SIZE=1 COLOR='#BBBBBB' >$&nbsp;" & Formatnumber(TABPROCESOSCRTM(WKI,II + 3) ,2,0,0,-1) & "</b>"
                  WFTOTALPESOS = WFTOTALPESOS + TABPROCESOSCRTM(WKI,II + 3)
               ELSE

                  CANTACTW = "<b>" & TABPROCESOSCRT(WKI,II + 3) & "</b>"

               END IF


               sHrefW     = sHref & II
               sImage    = CANTACTW
               CANTACTW  = ofv.GenLink("","nt STYLE='COLOR:#000000;' ",sHrefW,"_blank","","",sImage,"si")

               if TABPROCESOSCRT(WKI,II + 3) = 0 then  CANTACTW = "-"

               Celda = Celda & ofv.gencelda("","UTg bgcolor='" & xTrColor & "' STYLE='COLOR:#000000;border-style:solid;border-width:1;border-color:" & BCOLORD & ";' ","CENTER","10%","","","",CANTACTW,"si")
            END IF
        NEXT




        TCANTACT0 = TCANTACT0 + TABPROCESOSCRT(WKI,2)

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

        WFHEADER = "<font color='#555555' >TOTAL GENERAL"

        if SFILTRO  = "RBR" OR SFILTRO  = "RRE" OR SFILTRO  = "RNV" then

           WFHEADER = WFHEADER & ":&nbsp;$&nbsp;" & Formatnumber(WFTOTALPESOS ,2,0,0,-1)

'''''''''''''''''''''''''''' & "&nbsp;--&nbsp;" & CANTREG & "&nbsp;--&nbsp;" & CANTREG2 & "&nbsp;--&nbsp;" & CANTREG3

        END IF

        WFHEADER = WFHEADER & "</font>&nbsp;&nbsp;&nbsp;&nbsp;"

        Celda = Celda & ofv.gencelda("","VTI COLSPAN=2 " & wkstylo1,"RIGHT","","","","",WFHEADER,"si")
        Celda = Celda & ofv.gencelda("","ut " & wkstylo1,"center","2%","","","",TCANTACT0 ,"si")

        xTrColor = "#F7EEDD"

        FOR II=1 TO 10

            IF TABPROCESOSCRT(0,II + 3) <> "NO" THEN



               if SFILTRO  = "RBR" OR SFILTRO  = "RRE" OR SFILTRO  = "RNV" then

                  WSPIMPOF = TABPROCESOSCRTM(200,II+3)

                  WSPPORC  = WSPIMPOF / WFTOTALPESOS * 100
                  WSPPORC  = Formatnumber(WSPPORC ,2,-1,0,-1) & "&nbsp;%"

                  CANTACTW = ""

                  CANTACTW = CANTACTW & "<FONT COLOR='#555555' >" & TABPROCESOSCRT(200,II+3) & "</FONT>"
                  CANTACTW = CANTACTW & "<b>"
                  CANTACTW = CANTACTW & "<HR SIZE=1 COLOR='#993333' >$&nbsp;" & Formatnumber(TABPROCESOSCRTM(200,II+3) ,2,0,0,-1)
                  CANTACTW = CANTACTW & "</b>"
                  IF WSPIMPOF > 0 THEN
                     CANTACTW = CANTACTW & "<HR SIZE=1 COLOR='#993333' >" & WSPPORC
                  END IF



               ELSE

                  CANTACTW = "<b>" & TABPROCESOSCRT(200,II+3) & "</b>"

               END IF




               if TABPROCESOSCRT(200,II+3) = 0 then  CANTACTW = "-"

               Celda = Celda & ofv.gencelda("","UT bgcolor='" & xTrColor & "' STYLE='COLOR:#000000;border-style:solid;border-width:1;border-color:#555555;' ","CENTER","10%","","","",CANTACTW,"si")
            END IF
        NEXT




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