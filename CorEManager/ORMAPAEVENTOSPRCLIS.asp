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


  WSPSESION = "ORWHERE"


  sHTML = ofv.MENUHeader("","")


  ofv.ObtenerAtributos Session("Form"),""
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
    SORPERIODO  = request.FORM("ORPERIODO")

  ''   RESPONSE.WRITE " PER:" & SORPERIODO & "<BR>"

    SOREVALUADOS       = request.FORM("OREVALUADOS")
    SORRIESGOSIDEN     = request.FORM("ORRIESGOSIDEN")
    SORSPAIDEN         = request.FORM("ORSPAIDEN")

    SORFECINID   = request.FORM("ORFECINID")
    SORFECINIH   = request.FORM("ORFECINIH")


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

       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORPERIODOW
       SCONECTOR = " AND "

    END IF


    IF SORPROCESO <> "TODOS" THEN
       SORPROCESOW   = " (A.TRCPROCESO = 0" & sORPROCESO & ") "

       SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & SORPROCESOW
       SCONECTOR = " AND "

    END IF

    SESSION(WSPSESION) =  SESSION(WSPSESION) & SCONECTOR & "  (A.trccodtracerp IS NULL OR  A.trccodtracerp = '') AND TRCESTADO <> 'CAN' AND TRCESTADO <> 'DOBLIGA' and A.FECORIG > 0 "




    IF SESSION(WSPSESION) <> "" THEN
       SESSION(WSPSESION) =  SWHERE & SESSION(WSPSESION)
    END IF

    WSPWHERE = SESSION(WSPSESION)


  ELSE

    WSPWHERE = SESSION(WSPSESION)


  END IF


  if SFILTRO  = "" then
     SFILTRO  = "RRE"

  end if

  DIM TABPROCESOSCRT(200,20)
  DIM TABPROCESOSCRTM(200,20)

  Strsql = "Select OBJID,PERIODO,RGAMBAUTOEV From ORPERIODOS  where OBJID = 0" & SORPERIODO
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     Speriododsc              = RS(1)
     Speriodo                 = RS(0)
     WKORPERIODOMAP           = RS(2)
  ELSE
     WKORPERIODOMAP           = 0
  END IF


  Fila = ""
  xTrColor = "#FFFFFF"

  sImageva    = "<font face='corflaticon' size=2 color='#990000'>&#0110;</font>"
  sImagevm    = "<font face='corflaticon' size=2 color='#f7aa22'>&#0110;</font>"
  sImagevb    = "<font face='corflaticon' size=2 color='#009900'>&#0110;</font>"
  sImagevn    = "<font face='corflaticon' size=2 color='#dddddd'>&#0110;</font>"

  wkstylof = "aaa style='font-family:arial;font-size:6pt;color:#888888;' "

''''b.codvinobj = 'VN' AND



  DIM TABMAPA(10,5)
  DIM TABMAPAM(10,5)




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









  Strsql = "     select B.CODOBJETO,B.DESCRI, "
  Strsql = Strsql & " a.* from csrelpathdescri a"
  Strsql = Strsql & "     inner join objetos b on (a.cobj = b.codobjeto)"
  Strsql = Strsql & " where "
  Strsql = Strsql & " B.APLIC = 'RG' AND "
  Strsql = Strsql & " A.MC7 = 0" & WKORPERIODOMAP

 IF SORPROCESO <> "TODOS" THEN

  Strsql = Strsql & " and ("
  Strsql = Strsql & " A.MC1 = " & SORPROCESO & " OR "
  Strsql = Strsql & " A.MC2 = " & SORPROCESO & " OR "
  Strsql = Strsql & " A.MC3 = " & SORPROCESO & " OR "
  Strsql = Strsql & " A.MC4 = " & SORPROCESO & " OR "
  Strsql = Strsql & " A.MC5 = " & SORPROCESO & " OR "
  Strsql = Strsql & " A.MC6 = " & SORPROCESO & " OR "
  Strsql = Strsql & " A.MC7 = " & SORPROCESO & " OR "
  Strsql = Strsql & " A.COBJ = " & SORPROCESO & "  "
  Strsql = Strsql & " ) "

 END IF

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

        RS.MOVENEXT
     LOOP



  END IF

  MAINPROC = ""
  MAINACTI = ""
  MAINRIES = ""


  Strsql = "     select COALESCE(W.ESTADO,'0'), "
  Strsql = Strsql & " a.* from csrelpathdescri a"
  Strsql = Strsql & " inner join objetos       b on (a.cobj = b.codobjeto)"
  Strsql = Strsql & " inner join objlnk        x on (a.cobj = x.codobjeto and (x.ltipo = 'LRGVN' or x.ltipo = 'LRGCR'))"
  Strsql = Strsql & " inner join objproced     c on (c.codobjeto = x.codobjetol and c.codproced = x.iden4 )"
  Strsql = Strsql & " inner join objPRTY       M on (a.cobj = M.codobjeto AND M.ESTADO <> 4)"
  Strsql = Strsql & " LEFT  join orrskcircuito w on (a.cobj = w.codobjeto AND w.periodo = 0" & Speriodo & " )"

 IF SORPROCESO <> "TODOS" THEN

  Strsql = Strsql & " WHERE ("
  Strsql = Strsql & " A.MC1 = " & SORPROCESO & " OR "
  Strsql = Strsql & " A.MC2 = " & SORPROCESO & " OR "
  Strsql = Strsql & " A.MC3 = " & SORPROCESO & " OR "
  Strsql = Strsql & " A.MC4 = " & SORPROCESO & " OR "
  Strsql = Strsql & " A.MC5 = " & SORPROCESO & " OR "
  Strsql = Strsql & " A.MC6 = " & SORPROCESO & " OR "
  Strsql = Strsql & " A.MC7 = " & SORPROCESO & " OR "
  Strsql = Strsql & " A.COBJ = " & SORPROCESO & "  "
  Strsql = Strsql & " ) "

 END IF

  Strsql = Strsql & " ORDER BY A.TSEC "

''  response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then

     CANTACT = 0

     DO WHILE NOT RS.EOF

        WKIPROCD = ""
        WKIPROCN = 0

        FOR WKI=1 TO 8

            WKIPROC = RS(WKI) & ""

            IF WKIPROC <> "" THEN
               WKIPROCN = RS(WKI + 8)
               EXIT FOR
            END IF


        NEXT


        IF MAINPROC <> WKIPROCN THEN                     ' CAMBIA PROCESO

           IF MAINPROC <> "" THEN


                 TABPROCESOSCRT(POSI,2) = CANTACT       'CANT.ACTIVIDADES

                 FOR II=0 TO 10
                     TABPROCESOSCRT(POSI,II + 4) =  TABMAPA(II,01)
                 NEXT



            END IF


            MAINPROC = WKIPROCN

            FOR WKI=1 TO TOPE

                IF TABPROCESOSCRT(WKI,0) = WKIPROCN THEN
                   POSI = WKI
                   EXIT FOR
                END IF

            NEXT

            FOR II=0 TO 10

                TABMAPA(II,01) = 0

            NEXT

            CANTACT   = 0



        END IF

          II = RS(0) 


          CANTACT   = CANTACT + 1
          TABMAPA(II,01) = TABMAPA(II,01) + 1


          RS.MOVENEXT
        LOOP

        TABPROCESOSCRT(POSI,2) = CANTACT       'CANT.ACTIVIDADES
        FOR II=0 TO 10
            TABPROCESOSCRT(POSI,II + 4) =  TABMAPA(II,01)
        NEXT

     END IF







'''bgcolor='#ddddff'
''#F7CC88

       wkstylo1 = " style='background-color:#ddddff;color:#888888;border-style:solid;border-width:1;border-color:" & BCOLORG & ";' "


'    sHTML = sHTML & "<CENTER>" & ofv.GenTabla("","aaa BORDER=1 STYLE='border-style:solid;border-width:3;border-color:#F0F0F0;' ","","80% ","","0","","-1","0","","",Fila2,"si") & "</CENTER>"

     BCOLORG = "#AAAAAA"
     BCOLORD = "#6699CC"



 


    Fila = ""





'''''''''''''''
'' REPORTE DE RIESGO RESIDUAL POR TIPO DE EVENTO
'''''''''''''''''''''''''''''''''''''''''''''''''''''

''' encabezados

    Fila = ""
    Fila1 = ""


    Celda = ""
    Celda = Celda & ofv.gencelda("","VTIg COLSPAN=3  style='border-style:solid;border-width:1;border-color:#888888;'  ","","","","","",session("FormN") & " - PERIODO: " & SPERIODODSC,"si")
    Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#F7EEDD' ","","","25","","",Celda,"si")


    Celda = ""
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'   ","center","50%","","","","PROCESO","si")
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;'  ","CENTER","50%","","","","&nbsp;ESTADO&nbsp;","si")
 ''   Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#555555' ","","","25","","",Celda,"si")


            WSPIMPOFTT = 0
            WSPCANTFTT = 0

    Celda = ""
    WSPMANDA = "&nbsp;"
    Celda = Celda & ofv.gencelda("","UT ","center","70%","","","",WSPMANDA,"si")
    Celda = Celda & ofv.gencelda("","UT ","CENTER","15%","","","",WSPMANDA,"si")
    Celda = Celda & ofv.gencelda("","UT ","CENTER","15%","","","",WSPMANDA,"si")
    Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda,"si")

    FOR IPROB=1 TO TOPE


        WSPPROCESO = "PROCESO:&nbsp;" & TABPROCESOSCRT(IPROB,1)

        CELDA1 = ""
        Celda1 = Celda1 & ofv.gencelda("","VTI BGCOLOR='#EEF7DD' COLSPAN=3 style='color:#000000;border-style:solid;border-width:1;border-color:#000000;'  ","","","","","",WSPPROCESO,"si")
        Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda1,"si")


        CELDA1 = ""
        WSPITEM = "&nbsp;&nbsp;&nbsp;PENDIENTES:&nbsp;"
        Celda1 = Celda1 & ofv.gencelda("","UTI style='color:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","","","","","",WSPITEM,"si")
        WSPIMPOF = TABPROCESOSCRT(IPROB,4)

        Celda1 = Celda1 & ofv.gencelda("","UT COLSPAN=2 style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","",WSPIMPOF,"si")
        Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda1,"si")



        CELDA1 = ""
        WSPITEM = "&nbsp;&nbsp;&nbsp;EN CURSO:&nbsp;"
        Celda1 = Celda1 & ofv.gencelda("","UTI style='color:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","","","","","",WSPITEM,"si")
        WSPIMPOF = TABPROCESOSCRT(IPROB,5) + TABPROCESOSCRT(IPROB,6) + TABPROCESOSCRT(IPROB,9) + TABPROCESOSCRT(IPROB,10)

        Celda1 = Celda1 & ofv.gencelda("","UT COLSPAN=2 style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","",WSPIMPOF,"si")
        Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda1,"si")



        CELDA1 = ""
        WSPITEM = "- INICIADOS:&nbsp;&nbsp;"
        Celda1 = Celda1 & ofv.gencelda("","UT style='color:#336699;border-style:solid;border-width:1;border-color:#DDDDDD;'  ","RIGHT","","","","",WSPITEM,"si")
        WSPIMPOF = TABPROCESOSCRT(IPROB,5)

        Celda1 = Celda1 & ofv.gencelda("","UT style='COLOR:#336699;border-style:solid;border-width:1;border-color:#DDDDDD;'  ","right","","","","",WSPIMPOF,"si")
        WSPMANDA = "&nbsp;"
        Celda1 = Celda1 & ofv.gencelda("","UT style='COLOR:#336699;border-style:solid;border-width:1;border-color:#DDDDDD;' ","CENTER","15%","","","",WSPMANDA,"si")
        Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda1,"si")

        CELDA1 = ""
        WSPITEM = "- VERIFICADOS:&nbsp;&nbsp;"
        Celda1 = Celda1 & ofv.gencelda("","UT style='color:#336699;border-style:solid;border-width:1;border-color:#DDDDDD;'  ","RIGHT","","","","",WSPITEM,"si")
        WSPIMPOF = TABPROCESOSCRT(IPROB,9)

        Celda1 = Celda1 & ofv.gencelda("","UT style='COLOR:#336699;border-style:solid;border-width:1;border-color:#DDDDDD;'  ","right","","","","",WSPIMPOF,"si")
        WSPMANDA = "&nbsp;"
        Celda1 = Celda1 & ofv.gencelda("","UT style='COLOR:#336699;border-style:solid;border-width:1;border-color:#DDDDDD;' ","CENTER","15%","","","",WSPMANDA,"si")
        Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda1,"si")

        CELDA1 = ""
        WSPITEM = "- AUTORIZADOS:&nbsp;&nbsp;"
        Celda1 = Celda1 & ofv.gencelda("","UT style='color:#336699;border-style:solid;border-width:1;border-color:#DDDDDD;'  ","RIGHT","","","","",WSPITEM,"si")
        WSPIMPOF = TABPROCESOSCRT(IPROB,6)

        Celda1 = Celda1 & ofv.gencelda("","UT style='COLOR:#336699;border-style:solid;border-width:1;border-color:#DDDDDD;'  ","right","","","","",WSPIMPOF,"si")
        WSPMANDA = "&nbsp;"
        Celda1 = Celda1 & ofv.gencelda("","UT style='COLOR:#336699;border-style:solid;border-width:1;border-color:#DDDDDD;' ","CENTER","15%","","","",WSPMANDA,"si")
        Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda1,"si")

        CELDA1 = ""
        WSPITEM = "- OBSERVADOS:&nbsp;&nbsp;"
        Celda1 = Celda1 & ofv.gencelda("","UT style='color:#336699;border-style:solid;border-width:1;border-color:#DDDDDD;'  ","RIGHT","","","","",WSPITEM,"si")
        WSPIMPOF = TABPROCESOSCRT(IPROB,10)

        Celda1 = Celda1 & ofv.gencelda("","UT style='COLOR:#336699;border-style:solid;border-width:1;border-color:#DDDDDD;'  ","right","","","","",WSPIMPOF,"si")
        WSPMANDA = "&nbsp;"
        Celda1 = Celda1 & ofv.gencelda("","UT style='COLOR:#336699;border-style:solid;border-width:1;border-color:#DDDDDD;' ","CENTER","15%","","","",WSPMANDA,"si")
        Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda1,"si")





        CELDA1 = ""
        WSPITEM = "&nbsp;&nbsp;&nbsp;FINALIZADOS:&nbsp;"
        Celda1 = Celda1 & ofv.gencelda("","UTI style='color:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","","","","","",WSPITEM,"si")

        WSPIMPOF = TABPROCESOSCRT(IPROB,7)

        Celda1 = Celda1 & ofv.gencelda("","UT COLSPAN=2 style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","",WSPIMPOF,"si")
        Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda1,"si")


        CELDA1 = ""
        WSPITEM = "&nbsp;&nbsp;&nbsp;CANCELADOS:&nbsp;"
        Celda1 = Celda1 & ofv.gencelda("","UTI style='color:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","","","","","",WSPITEM,"si")

        WSPIMPOF = TABPROCESOSCRT(IPROB,8)

        Celda1 = Celda1 & ofv.gencelda("","UT COLSPAN=2 style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","",WSPIMPOF,"si")
        Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda1,"si")


        CELDA1 = ""
        WSPITEM = "&nbsp;TOTAL:&nbsp;"
        Celda1 = Celda1 & ofv.gencelda("","VT style='color:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","",WSPITEM,"si")

        WSPIMPOF = TABPROCESOSCRT(IPROB,7) + TABPROCESOSCRT(IPROB,5) + TABPROCESOSCRT(IPROB,6) + TABPROCESOSCRT(IPROB,9) + TABPROCESOSCRT(IPROB,10) + TABPROCESOSCRT(IPROB,8) + TABPROCESOSCRT(IPROB,4)

        Celda1 = Celda1 & ofv.gencelda("","VT COLSPAN=2 style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;'  ","right","","","","",WSPIMPOF,"si")
        Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#DDDDDD' ","","","25","","",Celda1,"si")


        Celda1 = ""
        WSPMANDA = "&nbsp;"
        Celda1 = Celda1 & ofv.gencelda("","UT COLSPAN=3 ","right","","","","",WSPMANDA,"si")
        Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","","","",Celda1,"si")


    NEXT


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