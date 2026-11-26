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

  DIM TABMAPA(10,10,1)

  DIM TABPROBABILIDAD(10,5)
  DIM TABIMPACTO(10,5)
  DIM TABRIESGOS(50,5)
  DIM TABRANRIESGOS(10,5)

  DIM TABCOLOR(10)


  TABCOLOR(1)   = "#00DD00"
  TABCOLOR(2)   = "#F7EEDD"
  TABCOLOR(3)   = "#F7DD77"
  TABCOLOR(4)   = "#FF6666"
  TABCOLOR(5)   = "#993333"
  TABCOLOR(6)   = "#BBBBBB"
  TABCOLOR(7)   = "#888888"
  TABCOLOR(8)   = "#555555"
  TABCOLOR(9)   = "#333333"
  TABCOLOR(10)  = "#000000"


  TTABPROB = 0
  TTABIMPA = 0
  TTABRIESGO = 0

  sHTML = ofv.FormHeader(session("FormN"))


  ofv.ObtenerAtributos Session("Form"),""
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)


     if Speriodo = "" then

        Strsql = "Select * From ORPERIODOS  where APLIC = '" & wkaplic & "' and ESTADO IN ('OK','CL') and RGVALEAUTOEV = 'S' "
        Strsql = Strsql & " ORDER BY FINAL,INICIO, ESTADO DESC"

        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if NOT rs.eof then
           session("ORPERIODO")     = RS(0)
           session("ORPERIODODESC") = RS(1)
           session("ORPERIODOMAP")  = RS(8)
           Speriodo                 = session("ORPERIODO")
           Speriododsc              = session("ORPERIODODESC")
           Speriodomap              = session("ORPERIODOMAP")

        end if

      else

        Strsql = "Select * From ORPERIODOS  where OBJID = 0" & Speriodo
        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if NOT rs.eof then
           Speriododsc              = RS(1)
           Speriodomap              = RS(8)
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
                 TABRIESGOS(POSI,2) = TABRANRIESGOS(II,2) 'DESCRIPCION
                 TABRIESGOS(POSI,5) = TABRANRIESGOS(II,3) 'COLOR
                 EXIT FOR
              END IF

          NEXT

          TTABRIESGO = POSI

        RS.MOVENEXT
     LOOP



  END IF


  Strsql = "Select * From ORNIVPROBOCURRENCIA ORDER BY VALOR desc"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then
     POSI = 0
     DO WHILE NOT RS.EOF
        POSI = POSI + 1

          TABPROBABILIDAD(POSI,0) = RS(0) 'CODIGO
          TABPROBABILIDAD(POSI,1) = RS(1) 'VALOR
          TABPROBABILIDAD(POSI,2) = RS(2) 'DESCRIPCION
          TABPROBABILIDAD(POSI,3) = RS(3) 'PORCENTAJE

          TABMAPA(POSI,0,0)         = 0
          TABMAPA(POSI,0,1)         = 0

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

          TABMAPA(0,POSI,0)         = 0
          TABMAPA(0,POSI,1)         = 0

          FOR II=1 TO TTABPROB
              TABMAPA(II,POSI,0)         = 0
              TABMAPA(II,POSI,1)         = 0
          NEXT

          TTABIMPA = POSI


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
  Strsql = Strsql & " where d.critico = 'SI' and w.periodo = 0" & Speriodo & "  and f.periodo = w.periodo and f.probabilidad > 0 and f.impacto > 0 "
  Strsql = Strsql & " ORDER BY f.probabilidad, f.impacto,A.TSEC "

'response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then

     CANTREG = 0

     CANTACT = 0


     PROBANT = RS(2)
     IMPAANT = RS(3)

     IPROB = 0
     IIMPA = 0


     FOR II=1 TO TTABPROB
         if TABPROBABILIDAD(ii,0) = RS(2) then
            IPROB = ii
            exit for
         end if
     NEXT

     FOR II=1 TO TTABIMPA
         if TABIMPACTO(ii,0) = RS(3) then
            IIMPA = ii
            exit for
         end if
     NEXT


     PROBCANTA = 0
     IMPACANTA = 0

     PROBCANTP = 0
     IMPACANTP = 0

     PROBCANTI = 0
     IMPACANTI = 0




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


        if PROBANT <> RS(2) or IMPAANT <> RS(3) then

           TABMAPA(IPROB,IIMPA,0)         = IMPACANTI
           TABMAPA(IPROB,IIMPA,1)         = IMPACANTA

           TABMAPA(0,IIMPA,0)             = TABMAPA(0,IIMPA,0) + IMPACANTI
           TABMAPA(0,IIMPA,1)             = TABMAPA(0,IIMPA,1) + IMPACANTA

           IMPACANTA = 0
           IMPACANTI = 0
           IMPACANTP = 0

           IMPAANT = RS(3)



           FOR II=1 TO TTABIMPA
               if TABIMPACTO(ii,0) = RS(3) then
                  IIMPA = ii
                  exit for
               end if
           NEXT

           if  PROBANT <> RS(2) then

               TABMAPA(IPROB,0,0)         = PROBCANTI
               TABMAPA(IPROB,0,1)         = PROBCANTA


               PROBCANTA = 0
               PROBCANTP = 0
               PROBCANTI = 0

               PROBANT = RS(2)

               FOR II=1 TO TTABPROB
                   if TABPROBABILIDAD(ii,0) = RS(2) then
                      IPROB = ii
                      exit for
                   end if
               NEXT

           end if



        end if



        PROBCANTA     = PROBCANTA + 1
        IMPACANTA     = IMPACANTA + 1

        wspporc       = CINT("0" & TABPROBABILIDAD(IPROB,3))
        wspmonto      = clng((Clng("0" & TABIMPACTO(IIMPA,3)) + Clng("0" & TABIMPACTO(IIMPA,4))) / 2)

        wspmontobruto = clng(wspmonto * wspporc / 100)

        PROBCANTI     = PROBCANTI + wspmontobruto
        IMPACANTI     = IMPACANTI + wspmontobruto





        RS.MOVENEXT
     LOOP




           TABMAPA(IPROB,IIMPA,0)     = IMPACANTI
           TABMAPA(IPROB,IIMPA,1)     = IMPACANTA

           TABMAPA(0,IIMPA,0)         = TABMAPA(0,IIMPA,0) + IMPACANTI
           TABMAPA(0,IIMPA,1)         = TABMAPA(0,IIMPA,1) + IMPACANTA


           IMPACANTA                  = 0
           IMPACANTI                  = 0
           IMPACANTP                  = 0

           TABMAPA(IPROB,0,0)         = PROBCANTI
           TABMAPA(IPROB,0,1)         = PROBCANTA

           PROBCANTA                  = 0
           PROBCANTP                  = 0
           PROBCANTI                  = 0


  END IF


''' encabezados

    Fila = ""
    Fila1 = ""
    Fila2 = ""

    sform = "ORMATPER"

     Strsql = "Select *, (99999 - FINAL) AS XFFINAL From ORPERIODOS  where APLIC = '" & wkaplic & "' and "
     Strsql = Strsql & " ESTADO IN ('OK','CL') and RGVALEAUTOEV = 'S' "
     Strsql = Strsql & " ORDER BY XFFINAL,INICIO, ESTADO DESC"
     sInput = ofv.GenerarCombo(sForm,"PERIODO",SPERIODO,Strsql,cn,0,1,"")

    Celda = ""
    Celda = Celda & ofv.gencelda("","VTIg BGCOLOR='#F7EEDD' style='border-style:solid;border-width:1;border-color:#888888;'  ","","50%","","","","MAPA DE RIESGOS","si")
    Celda = Celda & ofv.gencelda("","VTI style='color:#888888;border-style:solid;border-width:1;border-color:#888888;' ","","25%","","","","PERIODO DE EVALUACION:","si")
    Celda = Celda & ofv.gencelda("","VTI style='color:#888888;border-style:solid;border-width:1;border-color:#888888;' ","","30%","","","",sInput,"si")
    FilaX = ofv.genrow("","AAA BGCOLOR='#ddddff' ","","","25","","",Celda,"si")
    saccion = "ORMAPARIESGOS.ASP?PERIODO=&FILTRO=" & SFILTRO
    FilaX =  ofv.genform(sForm,"",sAccion,"","post","",FilaX,"si")

    Celda = ""
    Celda = Celda & ofv.gencelda("","VTg BGCOLOR='#ddddff' style='color:#000000;border-style:solid;border-width:1;border-color:#888888;' colspan=3 rowspan=2 ","center","","","","","CUANTITATIVO","si")
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;' COLSPAN=" & TTABIMPA & " ","CENTER","","","","","IMPACTO","si")
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;' VALIGN='BOTTOM' ROWSPAN=4 ","CENTER","10%","","","","&nbsp;TOTAL&nbsp;","si")
    Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#555555' ","","","25","","",Celda,"si")

    Celda = ""
    Celda = Celda & ofv.gencelda("","VTg BGCOLOR='#ddddff' style='color:#000000;border-style:solid;border-width:1;border-color:#888888;' colspan=3 rowspan=2 ","center","","","","","CUALITATIVO","si")
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;' COLSPAN=" & TTABIMPA & " ","CENTER","","","","","IMPACTO","si")
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;' VALIGN='BOTTOM' ROWSPAN=4 ","CENTER","10%","","","","&nbsp;TOTAL&nbsp;","si")
    Fila2 = FILA2 & ofv.genrow("","AAA BGCOLOR='#555555' ","","","25","","",Celda,"si")


    Celda = ""
    FOR II=1 TO TTABIMPA
        Celda = Celda & ofv.gencelda("","VTI BGCOLOR='#888888' style='color:#ffffff;border-style:solid;border-width:1;border-color:#000000;'  ","","","","","",TABIMPACTO(II,2),"si")
    NEXT
    Fila = FILA & ofv.genrow("","AAA BGCOLOR='#f0f0f0' ","","","25","","",Celda,"si")


    Fila1 = FILA1 & FILA
    Fila2 = FILA2 & FILA



    Celda = ""
    Celda = Celda & ofv.gencelda("","VT style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;' colspan=2 rowspan=2 ","center","","","","","PROBABILIDAD","si")
    Celda = Celda & ofv.gencelda("","Vt style='color:#FFFFFF;border-style:solid;border-width:1;border-color:#888888;' rowspan=2 ","center","","","","","CANT","si")

    CELDA1 = CELDA
    CELDA2 = CELDA



    FOR II=1 TO TTABIMPA
        Celda1 = Celda1 & ofv.gencelda("","ut BGCOLOR='#f0f0f0' style='color:#000000;border-style:solid;border-width:1;border-color:#990000;'  ","center","10%","","","",TABIMPACTO(II,3),"si")
        Celda2 = Celda2 & ofv.gencelda("","ut BGCOLOR='#f0f0f0' style='color:#000000;border-style:solid;border-width:1;border-color:#990000;'  ","center","10%","","","",TABIMPACTO(II,1),"si")
    NEXT



    Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#555555' ","","","25","","",Celda1,"si")

    Fila2 = FILA2 & ofv.genrow("","AAA BGCOLOR='#555555' ","","","25","","",Celda2,"si")

    CELDA1 = ""
    CELDA2 = ""

    FOR II=1 TO TTABIMPA
        Celda1 = Celda1 & ofv.gencelda("","ut BGCOLOR='#f0f0f0' style='color:#000000;border-style:solid;border-width:1;border-color:#990000;'  ","center","","","","",TABIMPACTO(II,4),"si")
        Celda2 = Celda2 & ofv.gencelda("","ut   ","center","","","","","&nbsp;","si")
    NEXT
    Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#ddddff' ","","","25","","",Celda1,"si")
    Fila2 = FILA2 & ofv.genrow("","AAA BGCOLOR='#ddddff' ","","","5","","",Celda2,"si")

            WSPIMPOFTT = 0
            WSPCANTFTT = 0

    FOR IPROB=1 TO TTABPROB

        WSPRB = TABPROBABILIDAD(IPROB,1) * TABIMPACTO(IIMPA,1)

        CELDA1 = ""
        CELDA2 = ""


        Celda1 = Celda1 & ofv.gencelda("","VTI BGCOLOR='#888888' style='color:#ffffff;border-style:solid;border-width:1;border-color:#000000;'  ","","20%","","","",TABPROBABILIDAD(IPROB,2),"si")
        Celda1 = Celda1 & ofv.gencelda("","ut style='color:#000000;border-style:solid;border-width:1;border-color:#000080;'  ","right","5%","","","",TABPROBABILIDAD(IPROB,3) & "%&nbsp;","si")
        Celda1 = Celda1 & ofv.gencelda("","ut style='color:#000000;border-style:solid;border-width:1;border-color:#000080;'  ","right","5%","","","",TABMAPA(IPROB,0,1) & "&nbsp;","si")


        Celda2 = Celda2 & ofv.gencelda("","VTI BGCOLOR='#888888' style='color:#ffffff;border-style:solid;border-width:1;border-color:#000000;'  ","","20%","","","",TABPROBABILIDAD(IPROB,2),"si")
        Celda2 = Celda2 & ofv.gencelda("","ut style='color:#000000;border-style:solid;border-width:1;border-color:#000080;'  ","right","5%","","","",TABPROBABILIDAD(IPROB,1) & "&nbsp;","si")
        Celda2 = Celda2 & ofv.gencelda("","ut style='color:#000000;border-style:solid;border-width:1;border-color:#000080;'  ","right","5%","","","",TABMAPA(IPROB,0,1) & "&nbsp;","si")

        FOR IIMPA=1 TO TTABIMPA

            wspcolor = "#eeeeff"


            FOR II=1 TO TTABRIESGO

              IF TABRIESGOS(II,3) =  TABPROBABILIDAD(IPROB,0) AND TABRIESGOS(II,4) =  TABIMPACTO(IIMPA,0) THEN
               ''  TABRIESGOS(POSI,2) = TABRANRIESGOS(II,2) 'DESCRIPCION
                 wspcolor = TABRIESGOS(II,5) 'COLOR
                 EXIT FOR
              END IF

            NEXT



            WSPIMPOF = TABMAPA(IPROB,IIMPA,0)
            WSPCANTF = TABMAPA(IPROB,IIMPA,1)

            IF WSPIMPOF = 0 THEN
               WSPIMPOF = "&nbsp;-"
            ELSE
             ''  WSPIMPOF = "$&nbsp;" & WSPIMPOF


               WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,0,0,0,-1)

            END IF

            IF WSPCANTF = 0 THEN
               WSPCANTF = "&nbsp;-"
            ELSE
               WSPCANTF = WSPCANTF
            END IF
            'aca
            Celda1 = Celda1 & ofv.gencelda("","VT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;' BGCOLOR='" & wspcolor & "' ","right","","","","",WSPIMPOF & "&nbsp;","si")
            Celda2 = Celda2 & ofv.gencelda("","VT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;' BGCOLOR='" & wspcolor & "' ","right","","","","",WSPCANTF & "&nbsp;","si")
        NEXT

            wspcolor = "#eeeeff"

            WSPIMPOF = TABMAPA(IPROB,0,0)
            WSPCANTF = TABMAPA(IPROB,0,1)

            WSPIMPOFTT = WSPIMPOFTT + WSPIMPOF
            WSPCANTFTT = WSPCANTFTT + WSPCANTF

            IF WSPIMPOF = 0 THEN
               WSPIMPOF = "&nbsp;-"
            ELSE
           ''    WSPIMPOF = "$&nbsp;" & WSPIMPOF

               WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,0,0,0,-1)

            END IF

            IF WSPCANTF = 0 THEN
               WSPCANTF = "&nbsp;-"
            ELSE
               WSPCANTF = WSPCANTF
            END IF


        Celda1 = Celda1 & ofv.gencelda("","VT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;' BGCOLOR='" & wspcolor & "' ","right","","","","",WSPIMPOF & "&nbsp;","si")
        Celda2 = Celda2 & ofv.gencelda("","VT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;' BGCOLOR='" & wspcolor & "' ","right","","","","",WSPCANTF & "&nbsp;","si")

        Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda1,"si")
        Fila2 = FILA2 & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda2,"si")

    NEXT


        CELDA1 = ""
        CELDA2 = ""


        Celda1 = Celda1 & ofv.gencelda("","VTI BGCOLOR='#888888' style='color:#ffffff;border-style:solid;border-width:1;border-color:#000000;' colspan=3  ","RIGHT","","","","","TOTAL&nbsp;","si")
        Celda2 = Celda2 & ofv.gencelda("","VTI BGCOLOR='#888888' style='color:#ffffff;border-style:solid;border-width:1;border-color:#000000;' colspan=3  ","RIGHT","","","","","TOTAL&nbsp;","si")



        FOR IIMPA=1 TO TTABIMPA

            wspcolor = "#eeeeff"





            WSPIMPOF = TABMAPA(0,IIMPA,0)
            WSPCANTF = TABMAPA(0,IIMPA,1)




            IF WSPIMPOF = 0 THEN
               WSPIMPOF = "&nbsp;-"
            ELSE
           ''    WSPIMPOF = "$&nbsp;" & WSPIMPOF

               WSPIMPOF = "$&nbsp;" & Formatnumber(WSPIMPOF ,0,0,0,-1)

            END IF

            IF WSPCANTF = 0 THEN
               WSPCANTF = "&nbsp;-"
            ELSE
               WSPCANTF = WSPCANTF
            END IF

            Celda1 = Celda1 & ofv.gencelda("","VT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;' BGCOLOR='" & wspcolor & "' ","right","","","","",WSPIMPOF & "&nbsp;","si")
            Celda2 = Celda2 & ofv.gencelda("","VT style='COLOR:#000000;border-style:solid;border-width:1;border-color:#BBBBBB;' BGCOLOR='" & wspcolor & "' ","right","","","","",WSPCANTF & "&nbsp;","si")
        NEXT


            IF WSPIMPOFTT = 0 THEN
               WSPIMPOFTT = "&nbsp;-"
            ELSE
            '   WSPIMPOFTT = "$&nbsp;" & WSPIMPOFTT

               WSPIMPOFTT = "$&nbsp;" & Formatnumber(WSPIMPOFTT ,0,0,0,-1)

            END IF

            IF WSPCANTFTT = 0 THEN
               WSPCANTFTT = "&nbsp;-"
            ELSE
               WSPCANTFTT = WSPCANTFTT
            END IF


            Celda1 = Celda1 & ofv.gencelda("","VT BGCOLOR='#F7EEDD' style='COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;' BGCOLOR='" & wspcolor & "' ","right","","","","",WSPIMPOFTT & "&nbsp;","si")
            'aca
            Celda2 = Celda2 & ofv.gencelda("","VT BGCOLOR='#F7EEDD' style='COLOR:#000000;border-style:solid;border-width:1;border-color:#000000;' BGCOLOR='" & wspcolor & "' ","right","","","","",WSPCANTFTT & "&nbsp;","si")

        Fila1 = FILA1 & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda1,"si")
        Fila2 = FILA2 & ofv.genrow("","AAA BGCOLOR='#F0F0F0' ","","","25","","",Celda2,"si")

    sHTML = sHTML & "<CENTER>" & ofv.GenTabla("","aaa  STYLE='border-style:solid;border-width:3;border-color:#DDDDDD;' ","","90% ","","0","","1","1","","",FilaX,"si") & "</CENTER><BR>"

    sHTML = sHTML & "<CENTER>" & ofv.GenTabla("","aaa  STYLE='border-style:solid;border-width:3;border-color:#DDDDDD;' ","","90% ","","0","","1","1","","",Fila2,"si") & "</CENTER><BR>"

    sHTML = sHTML & "<CENTER>" & ofv.GenTabla("","aaa  STYLE='border-style:solid;border-width:3;border-color:#DDDDDD;' ","","90% ","","0","","1","1","","",Fila1,"si") & "</CENTER>"


'''bgcolor='#ddddff'
''#F7CC88

       wkstylo1 = " style='background-color:#ddddff;color:#888888;border-style:solid;border-width:1;border-color:" & BCOLORG & ";' "







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