<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  Server.ScriptTimeout=360
  CantRegMostrarx = 6
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta        = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
  sMODAL       = request.querystring("MODAL")
  sMODALMAIN   = request.querystring("MODALMAIN")
  sCODTRACER  = request.querystring("CODTRACER")
  sFECHA        = request.querystring("FECHA")
  sOPCION        = request.querystring("OPCION")

  if sMODAL = "" then sMODAL = "CSN"

  sHTML = ofv.MENUHeader("","")

  WKVISION = ""
  WSPSESION = "ORWHERE"

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

    SOREVALUADOS       = request.FORM("OREVALUADOS")
    SORRIESGOSIDEN     = request.FORM("ORRIESGOSIDEN")
    SORSPAIDEN         = request.FORM("ORSPAIDEN")

    SORFECINID   = request.FORM("ORFECINID")
    SORFECINIH   = request.FORM("ORFECINIH")

    SORPERIODO  = request.FORM("ORPERIODO")


    SORCODIGOW   = ""
    sORCARACTERW = ""
    SORTEXTOW    = ""
    sORRIESGOW   = ""
    SORCUENTAW   = ""
    SORCLIENTEW  = ""
    SORMONTODW   = ""
    SORMONTOHW   = ""
    SORPROCESOW  = ""
    SORLOTEW     = ""
    SOREVALUADOSW       = ""
    SORRIESGOSIDENW     = ""
    SORSPAIDENW         = ""

    SORFECINIDW   = ""
    SORFECINIHW   = ""
    SORPERIODOW   = ""

    SCONECTOR = ""
    SWHERE    = ""

    if sMODAL = "REG"  then
       SCONECTOR = " AND "
    ELSE
       SWHERE    = " WHERE "
    END IF

    SESSION(WSPSESION) =  ""

    IF SORPERIODO <> "TODOS" THEN
       SORPERIODOW   = " ( A.PERIODO > '0" & SORPERIODO & "') "

   '    SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORPERIODOW
   '    SCONECTOR = " AND "


        Strsql = "Select * From ORPERIODOS  where OBJID = 0" & SORPERIODO
        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if  NOT rs.eof then
            Speriododsc              = RS(1)
            Speriodo                 = RS(0)

            if RS(2) > 0 then SORFECINID  = Cdate(RS(2) + CLNG(CDATE("01-01-1900")) )
            if RS(3) > 0 then SORFECINIH  = Cdate(RS(3) + CLNG(CDATE("01-01-1900")) )

        end if



    END IF

    IF SORPROCESO <> "TODOS" THEN
       SORPROCESOW   = " (A.TRCPROCESO = 0" & sORPROCESO & ") "

       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORPROCESOW
       SCONECTOR = " AND "

    END IF

       WWSORFECINIDDSC = ""
       WWSORFECINIHDSC = ""

    IF SORFECINID = "" or SORFECINID = "0" THEN SORFECINID = "TODOS"
    IF SORFECINIH = "" or SORFECINIH = "0" THEN SORFECINIH = "TODOS"


    IF SORFECINID <> "TODOS" THEN

       if IsDate(SORFECINID) then
          WWSORFECINID = CLNG(CDATE(SORFECINID) - CLNG(CDATE("01-01-1900")))
          WWSORFECINIDDSC = " DESDE " & CDATE(SORFECINID)
       else
          WWSORFECINID = CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
          WWSORFECINIDDSC = " DESDE " & DATE()
       end if



       SORFECINIDW   = " (NOT A.FECORIG < " & WWSORFECINID & ") "

       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORFECINIDW
       SCONECTOR = " AND "

    END IF

    IF SORFECINIH <> "TODOS" THEN

       if IsDate(SORFECINIH) then
          WWSORFECINIH = CLNG(CDATE(SORFECINIH) - CLNG(CDATE("01-01-1900")))
          WWSORFECINIHDSC = " HASTA " & CDATE(SORFECINIH)
       else
          WWSORFECINIH = CLNG(DATE()) - CLNG(CDATE("01-01-1900"))
          WWSORFECINIHDSC = " HASTA " & DATE()
       end if



       SORFECINIHW   = " (NOT A.FECORIG > " & WWSORFECINIH & ") "

       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORFECINIHW
       SCONECTOR = " AND "

    END IF

    SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & "  (A.trccodtracerp IS NULL OR  A.trccodtracerp = '') "
    SESSION(WSPSESION) =  SESSION(WSPSESION) & " AND A.TRCESTADO <> 'CAN' AND A.TRCESTADO <> 'DOBLIGA'  AND A.TRCESTADO <> 'NO'  AND A.TRCESTADO <> 'REG'  "
    SESSION(WSPSESION) =  SESSION(WSPSESION) & " AND A.TRCPRODSER > 0 AND A.TRCIDENTACT <> 'NO' AND A.TRCEVENTO > 0 "



    IF SESSION(WSPSESION) <> "" THEN
       SESSION(WSPSESION) =  SWHERE & SESSION(WSPSESION)
    END IF

    WSPWHERE = SESSION(WSPSESION)


  ELSE

    WSPWHERE = SESSION(WSPSESION)


  END IF





  if sMODAL = "REG" or sMODAL = "ENR"  or sMODAL = "ENCA"  or sMODAL = "CSN"  then

  DIM TABPROBABILIDAD(10,5)
  DIM TABIMPACTO(10,5)
  DIM TABRIESGOSRES(50,6)
  DIM TABUMBRALES(10,4)

  TTABUMBRAL    = 0
  TTABPROB      = 0
  TTABIMPA      = 0
  TTABRIESGORES = 0

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

          TTABPROB = POSI

        RS.MOVENEXT
     LOOP



  END IF


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

          TTABIMPA = POSI


        RS.MOVENEXT
     LOOP



  END IF


  Strsql = "Select * From ORNIVRIESGO ORDER BY VALORDESDE "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     POSI = 0
     DO WHILE NOT RS.EOF
        POSI = POSI + 1

          TABRIESGOSRES(POSI,0) = RS(0) 'CODIGO

          TABRIESGOSRES(POSI,2) = CLNG(CDBL(RS(2)) * 100) 'VALOR DESDE
          TABRIESGOSRES(POSI,3) = CLNG(CDBL(RS(3)) * 100) 'VALOR HASTA

          TTABRIESGORES = POSI

        RS.MOVENEXT
     LOOP



  END IF



  TTABUMBRAL    = 4

'' 1ER UMB

   TABUMBRALES(1,1) = CDBL(0)
   TABUMBRALES(1,2) = CDBL(9.9)
   TABUMBRALES(1,3) = "#EEEEEE"

'' 2DO UMB

   TABUMBRALES(2,1) = CDBL(10)
   TABUMBRALES(2,2) = CDBL(29.9)
   TABUMBRALES(2,3) = "#00FF00"

'' 3ER UMB

   TABUMBRALES(3,1) = CDBL(30)
   TABUMBRALES(3,2) = CDBL(69.9)
   TABUMBRALES(3,3) = "#F7EEDD"

'' 4TO UMB

   TABUMBRALES(4,1) = CDBL(70)
   TABUMBRALES(4,2) = CDBL(99999)
   TABUMBRALES(4,3) = "#FF6666"




     Strsql = "     select A.*, M1.DESCRI AS SUBPROCESO, M2.DESCRI AS RIESGODSC, M3.ACTDES, COALESCE(B.PROBABILIDAD,'0'),  COALESCE(B.IMPACTO,'0'), COALESCE(B.RIESGORESIDUAL,'0') "
     Strsql = Strsql & " FROM ORTRACER A "
     Strsql = Strsql & " INNER JOIN OBJETOS         M1 ON (A.TRCPRODSER = M1.CODOBJETO)  "
     Strsql = Strsql & " INNER JOIN ORIMPRIESGO     M2 ON (A.TRCEVENTO  = M2.OBJID)  "
     Strsql = Strsql & " INNER JOIN ("
     Strsql = Strsql & "     select COALESCE(Me.descri,MD.DESCRIACC) as actdes, MD.CODACCION "
     Strsql = Strsql & "     from  objaccion Md "
     Strsql = Strsql & "     left join tipoaccalt Me on (Md.tipoaccion = Me.tipoaccionalt) "
     Strsql = Strsql & " )    "
     Strsql = Strsql & "                            M3  ON (A.TRCIDENTACT  = M3.codaccion)  "

     Strsql = Strsql & " LEFT JOIN ORACTIVRIESGO B ON (A.trcidentact = B.ACTIVIDAD AND A.TRCEVENTO = B.ITEMRIESGO AND B.PERIODO = " & SORPERIODO & " )  "
     Strsql = Strsql & WSPWHERE & "  ORDER BY A.TRCPRODSER , A.TRCIDENTACT, A.TRCEVENTO  "
     WKACCION = " - ENRIQUECIMIENTO DE EVENTOS DE PERDIDA"

     if sMODAL = "CSN"  then      WKACCION = " - PRINCIPALES PERDIDAS"

 ''' response.write strsql & "<br>"

   IF WWSORFECINIDDSC <> "" OR WWSORFECINIHDSC <> "" THEN

      Celda = ""
      Celda = Celda & ofv.GenCelda("","ut   ","","100%","","","",WWSORFECINIDDSC & WWSORFECINIHDSC,"si")
      Fila  = ofv.GenRow("","","","","","","",Celda,"si")

      sHTML = sHTML & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;BORDER-COLOR:#993333;' valign='middle'  ","","100% ","20","0","","-1","1","","",Fila,"si")
      sHTML = sHTML & "<br>"

   END IF

  sImageFNL     = "<font face='corflaticon' size=3 color='#336699'  >&#097;</font>"

  'Encabezados

   Formulario = ""
   FILAX = ""
   Celda = ""


    Fila = ""
    Fila1 = ""

    Celda = ""
    Celda = Celda & ofv.gencelda("","VTIg COLSPAN=7 BGCOLOR='#F7EEDD' style='border-style:solid;border-width:1;border-color:#888888;'  ","","","","","",session("FormN") & " - PERIODO: " & SPERIODODSC,"si")
    Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#555555' ","","","25","","",Celda,"si")

    Celda = ""
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ","center","20%","","","","SUBPROCESO","si")
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ","center","20%","","","","ACTIVIDAD","si")
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ","center","25%","","","","RIESGO","si")
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ","center","5%","","","","CODIGO","si")
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ","CENTER","10%","","","","PERDIDA","si")
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ","CENTER","10%","","","","R.RESIDUAL","si")
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ","CENTER","10%","","","","DESVIO","si")
    Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#555555' ","","","25","","",Celda,"si")




  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
 '  response.write strsql & "<br>" & rs(45).name
  if rs.eof then

  ELSE
    posI = 0

    xTrColor = "#f0f0f0"

     CORTEANT = "NO"

   posi = 0
   Do While (Not rs.EOF)

     wkmontonetox =  rs(16) & ""

     if wkmontonetox =  "" OR wkmontonetox =  "0" then

     else

        DATOACT = RS(7) & "-" & RS(9) & "-" & RS(10)

        IF DATOACT <> CORTEANT THEN

           IF CORTEANT = "NO" THEN

           ELSE

        CELDA1 = ""
        Celda1 = Celda1 & ofv.gencelda("","UT  COLSPAN=3 style='border-TOP-style:dotted;border-TOP-width:1;border-TOP-color:#888888;'  ","RIGHT","","","","","&nbsp;","si")

        Celda1 = Celda1 & ofv.gencelda("","VT  BGCOLOR='#DDDDDD' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","","","","","","TOTAL:&nbsp;","si")

            WSPIMPOF = wkmontonetoT
            WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,-1,0,-1)


        Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='#CCCCCC' style='COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")


            WSPIMPOF = wspmontoF
            WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,-1,0,-1)


        Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='#DDDDDD' style='COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")



            IF wkmontonetoT > 0 AND wspmontoF > 0 THEN
               IF  wkmontonetoT < wspmontoF THEN
                 '''  WSPPORC  = (wspmontoF / wkmontonetoT * 100) * -1
                   WSPPORC  = (100 - (wkmontonetoT / wspmontoF * 100)) * -1
               ELSE
             '''      WSPPORC  = wkmontonetoT / wspmontoF * 100
                   WSPPORC  = 100 - (wspmontoF / wkmontonetoT * 100)
               END IF
            ELSEIF wkmontonetoT = 0 THEN
               WSPPORC  = -100
            ELSE
               WSPPORC  = 100
            END IF

            WSPPORC2 = WSPPORC

            IF WSPPORC2 < 0 THEN
               WSPPORC2 = WSPPORC2 * -1
            END IF


            WSPCOLOR =  "#DDDDDD"

            FOR II=1 TO TTABUMBRAL

                  IF CDBL(WSPPORC2) < TABUMBRALES(ii,1)   OR CDBL(WSPPORC2) >  TABUMBRALES(ii,2)  THEN

                  ELSE
                     WSPCOLOR =  TABUMBRALES(ii,3)
                     EXIT FOR
                  END IF

            NEXT


            WSPPORC  = Formatnumber(WSPPORC ,2,-1,0,-1) & "%"

            WSPIMPOF = WSPPORC


        Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='" & WSPCOLOR & "' style='COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")


        Fila1 = FILA1 & ofv.genrow("","AAA  ","","","25","","",Celda1,"si")



           END IF

           CORTEANT = DATOACT

           CORTEANTS = RS("SUBPROCESO")
           CORTEANTA = RS("ACTDES")
           CORTEANTR = RS("RIESGODSC")

        wkmontonetoT =  0
        wkmontoRECUT =  0
        wkmontoBRUTT =  0
        wkmontoOMONT =  0
        wkmontoPERDT =  0
        wkmontoDESAT =  0

        CELDA1 = ""
        Celda1 = Celda1 & ofv.gencelda("","UT  COLSPAN=7   ","RIGHT","","","","","&nbsp;","si")
        Fila1 = FILA1 & ofv.genrow("","AAA  ","","","","","",Celda1,"si")


        CELDA1 = ""

        Celda1 = Celda1 & ofv.gencelda("","UT  BGCOLOR='#F7EEDD' style='color:#000000;border-style:solid;border-width:1;border-color:#888888;'  ","","","","","",CORTEANTS,"si")
        Celda1 = Celda1 & ofv.gencelda("","UT  BGCOLOR='#F7EEDD' style='color:#000000;border-style:solid;border-width:1;border-color:#888888;'  ","","","","","",CORTEANTA,"si")
        Celda1 = Celda1 & ofv.gencelda("","UT  BGCOLOR='#F7EEDD' style='color:#000000;border-style:solid;border-width:1;border-color:#888888;'  ","","","","","",CORTEANTR,"si")


        Celda1 = Celda1 & ofv.gencelda("","UT  COLSPAN=4   ","","","","","","&nbsp;","si")

        Fila1 = FILA1 & ofv.genrow("","AAA  ","","","25","","",Celda1,"si")


        'wkPROBAE   =  RS(43)
        'wkIMPAE    =  RS(44)
        'wkRREAE    =  RS(45)
         wkPROBAE   =  RS(52)
		 wkIMPAE    =  RS(53)
	     wkRREAE    =  RS(54)

        wspmontoF  = 0
        WSPDESVIO  = 0

'''''''''''''''''''

              IITP=0
              IITI=0

              FOR IITP=1 TO TTABPROB
                  if TABPROBABILIDAD(IITP,0) = wkPROBAE then
                     exit for
                  end if
              NEXT


              FOR IITI=1 TO TTABIMPA
                  if TABIMPACTO(IITI,0) = wkIMPAE then
                     exit for
                  end if
              NEXT



              FOR II=1 TO TTABRIESGORES

             '     response.write rs(51).name

                  IF cdbl(CDBL(wkRREAE) * 100) < TABRIESGOSRES(II,2)   OR cdbl(CDBL(wkRREAE) * 100) >  TABRIESGOSRES(II,3)  THEN

                  ELSE

                     wspmontoF      = 0

                     IF IITP > 0 AND IITI > 0 and CDBL(wkRREAE) > 0 THEN

                        wspDATA       = cdbl("0" & TABPROBABILIDAD(IITP,1)) * cdbl("0" & TABIMPACTO(IITI,1))

                        wspmonto      = cdbl((cdbl("0" & TABIMPACTO(IITI,3)) + cdbl("0" & TABIMPACTO(IITI,4))) / 2)    'tabla impacto     3 Y 4
                        wspporc       = CINT("0" & TABPROBABILIDAD(IITP,3))                                            'tabla probabilidad  3
                        wspmontobruto = cdbl(wspmonto * wspporc / 100)
                        wspDATM       = cdbl( round( (wspDATA * 100) / round( CDBL(wkRREAE) * 100 ,0)  ,4) )

                        wspmontoF     = cdbl(wspmontobruto / wspDATM)

                     END IF


                     EXIT FOR

                  END IF

              NEXT

'''''''''''''''''''''





        END IF




        wkmontoneto =  cdbl(rs(16))


        wkmontoRECU =  cdbl(rs(15))
        wkmontoBRUT =  cdbl(rs(12))
        wkmontoOMON =  cdbl(rs(14))
        wkmontoPERD =  cdbl(rs(13))
        wkmontoDESA =  cdbl(rs(36))



        wkmontonetoT =  wkmontonetoT + wkmontoneto
        wkmontoRECUT =  wkmontoRECUT + wkmontoRECU
        wkmontoBRUTT =  wkmontoBRUTT + wkmontoBRUT
        wkmontoOMONT =  wkmontoOMONT + wkmontoOMON
        wkmontoPERDT =  wkmontoPERDT + wkmontoPERD
        wkmontoDESAT =  wkmontoDESAT + wkmontoDESA




        CELDA1 = ""

        Celda1 = Celda1 & ofv.gencelda("","UT  COLSPAN=3   ","RIGHT","","","","","&nbsp;" & sImageFNL & "&nbsp;","si")


        Celda1 = Celda1 & ofv.gencelda("","UT  BGCOLOR='#EEEEEE' style='border-style:solid;border-width:1;border-color:#BBBBBB;'  ","RIGHT","","","","",RS(0) & "&nbsp;","si")




            WSPIMPOF = wkmontoneto


            IF WSPIMPOF = 0 THEN
               WSPIMPOF = "&nbsp;-"
            ELSE

               WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,0,0,-1)

            END IF




        Celda1 = Celda1 & ofv.gencelda("","VT style='border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")

        Celda1 = Celda1 & ofv.gencelda("","UT  COLSPAN=2   ","RIGHT","","","","","&nbsp;","si")


        Fila1 = FILA1 & ofv.genrow("","AAA  ","","","25","","",Celda1,"si")




     end if





    posi = posi + 1
    rs.Movenext
  Loop

        CELDA1 = ""
        Celda1 = Celda1 & ofv.gencelda("","UT  COLSPAN=3 style='border-TOP-style:dotted;border-TOP-width:1;border-TOP-color:#888888;'  ","RIGHT","","","","","&nbsp;","si")

        Celda1 = Celda1 & ofv.gencelda("","VT  BGCOLOR='#DDDDDD' style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","","","","","","TOTAL:&nbsp;","si")

            WSPIMPOF = wkmontonetoT
            WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,-1,0,-1)


        Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='#CCCCCC' style='COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")


            WSPIMPOF = wspmontoF
            WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,2,-1,0,-1)


        Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='#DDDDDD' style='COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")


            IF wkmontonetoT > 0 AND wspmontoF > 0 THEN
               IF  wkmontonetoT < wspmontoF THEN
                 '''  WSPPORC  = (wspmontoF / wkmontonetoT * 100) * -1
                   WSPPORC  = (100 - (wkmontonetoT / wspmontoF * 100)) * -1
               ELSE
             '''      WSPPORC  = wkmontonetoT / wspmontoF * 100
                   WSPPORC  = 100 - (wspmontoF / wkmontonetoT * 100)
               END IF
            ELSEIF wkmontonetoT = 0 THEN
               WSPPORC  = -100
            ELSE
               WSPPORC  = 100
            END IF


            WSPPORC2 = WSPPORC

            IF WSPPORC2 < 0 THEN
               WSPPORC2 = WSPPORC2 * -1
            END IF


            WSPCOLOR =  "#DDDDDD"

            FOR II=1 TO TTABUMBRAL

                  IF CDBL(WSPPORC2) < TABUMBRALES(ii,1)   OR CDBL(WSPPORC2) >  TABUMBRALES(ii,2)  THEN

                  ELSE
                     WSPCOLOR =  TABUMBRALES(ii,3)
                     EXIT FOR
                  END IF

            NEXT


            WSPPORC  = Formatnumber(WSPPORC ,2,-1,0,-1) & "%"

            WSPIMPOF = WSPPORC


        Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='" & WSPCOLOR & "' style='COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","right","","","","",WSPIMPOF & "&nbsp;","si")


        Fila1 = FILA1 & ofv.genrow("","AAA  ","","","25","","",Celda1,"si")






  END IF


        CELDA1 = ""
        Celda1 = Celda1 & ofv.gencelda("","UT  COLSPAN=7   ","RIGHT","","","","","&nbsp;","si")
        Fila1 = FILA1 & ofv.genrow("","AAA  ","","","","","",Celda1,"si")



  else



  end if







    sHTML = sHTML & "<BR>"

    sHTML = sHTML & "<CENTER>" & ofv.GenTabla("","aaa  STYLE='border-style:solid;border-width:3;border-color:#DDDDDD;' ","","95% ","","0","","1","1","","",Fila1,"si") & "</CENTER>"


'''bgcolor='#ddddff'
''#F7CC88

       wkstylo1 = " style='background-color:#ddddff;color:#888888;border-style:solid;border-width:1;border-color:" & BCOLORG & ";' "



  'Botonera
  Agregar = ""
  Buscar  = ""
  Celda   = ""

  sAccion = "ORMAPAEVENTOS.asp?modal=" & smodal & "&CODTRACER=" & SCODTRACER
  Celda = "" 'ofv.BotonVolver(sAccion)


  sHTML = sHTML &  ofv.botoneraESSX("0","",cantregmoverx,cantregmostrarx,"no","no",posi,"","","",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)


  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if



Response.Write sHTML



%>