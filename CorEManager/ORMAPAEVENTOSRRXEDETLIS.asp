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

  sOPCION        = request.querystring("OPCION")
  stevento       = request.querystring("TEVENTO")
  SORPERIODO     = request.querystring("ORPERIODO")

  WSPSESION = "ORWHERE"


  sHTML = ofv.MENUHeader("","")


  ofv.ObtenerAtributos Session("Form"),""
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)





    WSPWHERE = SESSION(WSPSESION) & " AND (G.AGRUPA = 0" & STEVENTO & ")  "





  if SFILTRO  = "" then
     SFILTRO  = "RRE"

  end if

  DIM TABPROCESOSCRT(200,20)
  DIM TABPROCESOSCRTM(200,20)



  Strsql = "Select * From ORPERIODOS  where OBJID = 0" & SORPERIODO
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     Speriododsc              = RS(1)
     Speriodo                 = RS(0)

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


  DIM TABRIESGOSTE(110,110)

  TTABRIESGOSTE1 = 0
  TTABRIESGOSTE2 = 0

  ITABRIESGOSTE1 = 0
  ITABRIESGOSTE2 = 0

  TABRIESGOSTE(10,0) = 0             'CODIGO
  TABRIESGOSTE(10,1) = "SIN ASIGNAR" 'DESCRI
  TABRIESGOSTE(10,2) = 0             'CANTIDAD
  TABRIESGOSTE(10,4) = 0             'total

  TABRIESGOSTE(0,10) = 0             'CODIGO
  TABRIESGOSTE(1,10) = "SIN ASIGNAR" 'DESCRI
  TABRIESGOSTE(2,10) = 0             'CANTIDAD
  TABRIESGOSTE(4,10) = 0             'total

  Strsql = "Select * From ORTIPOEVENTO ORDER BY SECUEN, DESCRI"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     POSI = 10
     DO WHILE NOT RS.EOF
        POSI = POSI + 1

          TABRIESGOSTE(POSI,0) = RS(0) 'CODIGO
          TABRIESGOSTE(POSI,1) = RS(1) 'DESCRI
          TABRIESGOSTE(POSI,2) = 0     'CANTIDAD
          TABRIESGOSTE(POSI,4) = 0     'total

          TTABRIESGOSTE1 = POSI

        RS.MOVENEXT
     LOOP
  END IF


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
  Strsql = Strsql & " A.MC7 = 0 "
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




''' encabezados

    Fila = ""
    Fila1 = ""

     WSPTEVENTO = "SIN ASIGNAR"


  Strsql = "Select DESCRI From ORTIPOEVENTO WHERE OBJID = 0" & STEVENTO
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     WSPTEVENTO = RS(0) & ""
  END IF


    Celda = ""
    Celda = Celda & ofv.gencelda("","VTIg COLSPAN=2  style='border-style:solid;border-width:1;border-color:#888888;'  ","","","","","",session("FormN") & " - PERIODO: " & SPERIODODSC,"si")
    Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#F7EEDD' ","","","25","","",Celda,"si")

    Celda = ""
    Celda = Celda & ofv.gencelda("","VTI COLSPAN=2  style='border-style:solid;border-width:1;border-color:#888888;COLOR:#000000;'  ","","","","","","TIPO DE EVENTO: " & WSPTEVENTO,"si")
    Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda,"si")

    Celda = ""
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'   ","center","50%","","","","RIESGO","si")
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ","CENTER","50%","","","","&nbsp;MONTO&nbsp;","si")
    Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#555555' ","","","25","","",Celda,"si")


  MAINPROC = ""
  MAINACTI = ""
  MAINRIES = ""


  Strsql = "     select D.CODACCION,F.CODRIESGO, f.probabilidad, f.impacto, f.riesgobruto, f.riesgoresidual, f.riesgoresidualNV, "
  Strsql = Strsql & " G.DESCRI, G.AGRUPA, F.ITEMRIESGO,a.* from csrelpathdescri a"
  Strsql = Strsql & "     inner join objetos       b on (a.cobj = b.codobjeto)"

  Strsql = Strsql & "     inner join orrskcircuito w on (a.cobj = w.codobjeto AND W.ESTADO <> 4)"
  Strsql = Strsql & "     inner join objlnk        x on (a.cobj = x.codobjeto and (x.ltipo = 'LRGVN' or x.ltipo = 'LRGCR'))"

  Strsql = Strsql & "     inner join objproced     c on ( c.codobjeto = x.codobjetol and c.codproced = x.iden4 )"


  Strsql = Strsql & "     inner join objaccion     d on ( C.codproced = d.codproced)  "
  Strsql = Strsql & "     LEFT  join tipoaccalt    e on (d.tipoaccion = e.tipoaccionalt)"

  Strsql = Strsql & "     INNER join oractivriesgo f on (d.codaccion = f.actividad )"
  Strsql = Strsql & "     LEFT  join ORIMPRIESGO   G on (F.ITEMRIESGO = G.OBJID)"
  Strsql = Strsql & " where d.critico = 'SI' and w.periodo = 0" & Speriodo & " AND (G.AGRUPA = 0" & STEVENTO & ") and f.periodo = w.periodo  and f.probabilidad > 0 and f.impacto > 0 "
  Strsql = Strsql & " ORDER BY G.AGRUPA, G.DESCRI "

''  response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then


  RIESGOANT = "NO"
  RIESGODSC = "NO"
  TOTALXRSK = 0
  TOTALGRAL = 0

     DO WHILE NOT RS.EOF



         WSPRIESGO = CSTR(RS(9) & "")

         IF WSPRIESGO <> RIESGOANT THEN

            IF   RIESGOANT = "NO" THEN

            ELSE

               CELDA1 = ""
               Celda1 = Celda1 & ofv.gencelda("","UTI BGCOLOR='#E0E0FF' style='color:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","","20%","","","",RIESGODSC,"si")


               WSPIMPOF = TOTALXRSK
 
               IF WSPIMPOF = 0 THEN
                  WSPIMPOF = "&nbsp;-"
               ELSE
                  WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)

                  WSPIMPOF = "<A CLASS=NT HREF='ORMAPAEVENTOSRRXEDETACTLIS.asp?ORPERIODO=" & Speriodo & "&TEVENTO=" & STEVENTO  & "&TRIESGO=" & RIESGOANT & "' TITLE='Lista de actividades asociadas a " & RIESGODSC & ".' TARGET='_blank' STYLE='COLOR:#000000;' ><U>" & WSPIMPOF & "</U></A>"

               END IF 

               WSPMANDA = WSPIMPOF & "&nbsp;"


               Celda1 = Celda1 & ofv.gencelda("","UT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","",WSPMANDA,"si")
               Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda1,"si")

            END IF 

            TOTALXRSK = 0
            RIESGOANT = WSPRIESGO
            RIESGODSC = RS(7) & ""

         END IF






        if SFILTRO  = "RBR" then '4

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

                     wspmontoF      = 0

                     IF IITP > 0 AND IITI > 0  THEN

                        wspDATA       = CINT("0" & TABPROBABILIDAD(IITP,1)) * CINT("0" & TABIMPACTO(IITI,1))

                        wspmonto      = clng((Clng("0" & TABIMPACTO(IITI,3)) + Clng("0" & TABIMPACTO(IITI,4))) / 2)    'tabla impacto     3 Y 4
                        wspporc       = CINT("0" & TABPROBABILIDAD(IITP,3))                                            'tabla probabilidad  3
                        wspmontobruto = clng(wspmonto * wspporc / 100)
                        wspmontoF     = wspmontobruto


                        ''''BUSCAR TIPO EVENTO Y ACUMULAR


                           FOR IIX=11 TO TTABRIESGOSTE1

                               IF CSTR(TABRIESGOSTE(IIX,0)) = CSTR(RS(8) & "") THEN
                                  TABRIESGOSTE(IIX,2) = TABRIESGOSTE(IIX,2) + 1
                                  TABRIESGOSTE(IIX,4) = TABRIESGOSTE(IIX,4) + wspmontoF
                                  EXIT FOR
                               END IF


                           NEXT

                           IF IIX > TTABRIESGOSTE1 THEN 
                              IIX = 10
                              TABRIESGOSTE(IIX,2) = TABRIESGOSTE(IIX,2) + 1
                              TABRIESGOSTE(IIX,4) = TABRIESGOSTE(IIX,4) + wspmontoF
                           END IF

                        ''''

                           TOTALXRSK = TOTALXRSK + wspmontoF
                           TOTALGRAL = TOTALGRAL + wspmontoF


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



                  IF CLNG(CDBL(RS(5)) * 100) < TABRIESGOSRES(II,2)   OR CLNG(CDBL(RS(5)) * 100) >  TABRIESGOSRES(II,3)  THEN

                  ELSE



                     wspmontoF      = 0

                     IF IITP > 0 AND IITI > 0 and CDBL(RS(5)) > 0 THEN

                        wspDATA       = CINT("0" & TABPROBABILIDAD(IITP,1)) * CINT("0" & TABIMPACTO(IITI,1))

                        wspmonto      = clng((Clng("0" & TABIMPACTO(IITI,3)) + Clng("0" & TABIMPACTO(IITI,4))) / 2)    'tabla impacto     3 Y 4
                        wspporc       = CINT("0" & TABPROBABILIDAD(IITP,3))                                            'tabla probabilidad  3
                        wspmontobruto = clng(wspmonto * wspporc / 100)
                        wspDATM       = cdbl( round( (wspDATA * 100) / round( CDBL(RS(5)) * 100 ,0)  ,2) )



                        wspmontoF     = clng(wspmontobruto / wspDATM)




                        ''''BUSCAR TIPO EVENTO Y ACUMULAR


                           FOR IIX=11 TO TTABRIESGOSTE1

                               IF CSTR(TABRIESGOSTE(IIX,0)) = CSTR(RS(8) & "") THEN
                                  TABRIESGOSTE(IIX,2) = TABRIESGOSTE(IIX,2) + 1
                                  TABRIESGOSTE(IIX,4) = TABRIESGOSTE(IIX,4) + wspmontoF
                                  EXIT FOR
                               END IF


                           NEXT

                           IF IIX > TTABRIESGOSTE1 THEN 
                              IIX = 10
                              TABRIESGOSTE(IIX,2) = TABRIESGOSTE(IIX,2) + 1
                              TABRIESGOSTE(IIX,4) = TABRIESGOSTE(IIX,4) + wspmontoF
                           END IF

                        ''''

                           TOTALXRSK = TOTALXRSK + wspmontoF
                           TOTALGRAL = TOTALGRAL + wspmontoF

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


                     wspmontoF      = 0

                     IF IITP > 0 AND IITI > 0 and CDBL(RS(6)) > 0 THEN

                        wspDATA       = CINT("0" & TABPROBABILIDAD(IITP,1)) * CINT("0" & TABIMPACTO(IITI,1))

                        wspmonto      = clng((Clng("0" & TABIMPACTO(IITI,3)) + Clng("0" & TABIMPACTO(IITI,4))) / 2)    'tabla impacto     3 Y 4
                        wspporc       = CINT("0" & TABPROBABILIDAD(IITP,3))                                            'tabla probabilidad  3
                        wspmontobruto = clng(wspmonto * wspporc / 100)
                        wspDATM       = cdbl( round( (wspDATA * 100) / round( CDBL(RS(6)) * 100 ,0)  ,2) )


                        wspmontoF     = clng(wspmontobruto / wspDATM)


                        ''''BUSCAR TIPO EVENTO Y ACUMULAR


                           FOR IIX=11 TO TTABRIESGOSTE1

                               IF CSTR(TABRIESGOSTE(IIX,0)) = CSTR(RS(8) & "") THEN
                                  TABRIESGOSTE(IIX,2) = TABRIESGOSTE(IIX,2) + 1
                                  TABRIESGOSTE(IIX,4) = TABRIESGOSTE(IIX,4) + wspmontoF
                                  EXIT FOR
                               END IF


                           NEXT

                           IF IIX > TTABRIESGOSTE1 THEN 
                              IIX = 10
                              TABRIESGOSTE(IIX,2) = TABRIESGOSTE(IIX,2) + 1
                              TABRIESGOSTE(IIX,4) = TABRIESGOSTE(IIX,4) + wspmontoF
                           END IF


                        ''''

                           TOTALXRSK = TOTALXRSK + wspmontoF
                           TOTALGRAL = TOTALGRAL + wspmontoF

                     END IF

                     EXIT FOR
                  END IF

              NEXT

          end if







          CANTREG = CANTREG + 1


          RS.MOVENEXT
        LOOP


            IF  TOTALXRSK > 0 THEN


               CELDA1 = ""
               Celda1 = Celda1 & ofv.gencelda("","UTI BGCOLOR='#E0E0FF' style='color:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","","20%","","","",RIESGODSC,"si")


               WSPIMPOF = TOTALXRSK
 
               IF WSPIMPOF = 0 THEN
                  WSPIMPOF = "&nbsp;-"
               ELSE
                  WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)
                  WSPIMPOF = "<A CLASS=NT HREF='ORMAPAEVENTOSRRXEDETACTLIS.asp?ORPERIODO=" & Speriodo & "&TEVENTO=" & STEVENTO  & "&TRIESGO=" & RIESGOANT & "' TITLE='Lista de actividades asociadas a " & RIESGODSC & ".' TARGET='_blank' STYLE='COLOR:#000000;' ><U>" & WSPIMPOF & "</U></A>"
               END IF 

               WSPMANDA = WSPIMPOF & "&nbsp;"


               Celda1 = Celda1 & ofv.gencelda("","UT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","",WSPMANDA,"si")
               Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda1,"si")

            END IF 

               CELDA1 = ""
               Celda1 = Celda1 & ofv.gencelda("","VT  style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","RIGHT","20%","","","","TOTAL:&nbsp;","si")


               WSPIMPOF = TOTALGRAL
 
               IF WSPIMPOF = 0 THEN
                  WSPIMPOF = "&nbsp;-"
               ELSE
                  WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)
               END IF 

               WSPMANDA = WSPIMPOF & "&nbsp;"


               Celda1 = Celda1 & ofv.gencelda("","VT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","right","","","","",WSPMANDA,"si")
               Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#EEF7DD' ","","","25","","",Celda1,"si")


     END IF







'''bgcolor='#ddddff'
''#F7CC88

       wkstylo1 = " style='background-color:#ddddff;color:#888888;border-style:solid;border-width:1;border-color:" & BCOLORG & ";' "



     BCOLORG = "#AAAAAA"
     BCOLORD = "#6699CC"

    sHTML = sHTML & "<BR>" 
    sHTML = sHTML & "<CENTER>" & ofv.GenTabla("","aaa  STYLE='border-style:solid;border-width:3;border-color:#DDDDDD;' ","","60% ","","0","","1","1","","",Fila1,"si") & "</CENTER>"



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