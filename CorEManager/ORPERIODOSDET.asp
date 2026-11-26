<%@LCID = 11274%>
<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorCopyGen.inc" -->
<%

Server.ScriptTimeout=60*5
ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  nCodigo    = ""
  sTipo      = ""

  nCodESQS   = ""
  sTipESQS   = ""

  nCodESQSO  = ""
  sTipESQSO  = ""



  sUsuario   = ""
  sTipobjeto = ""

  CCGIOCOD   = ""
  CCGIOTIP   = ""

  CCGINCOD   = ""
  CCGINTIP   = ""

  CCGICESQ   = ""
  CCGITESQ   = ""

  CCGOCESQ   = ""
  CCGOTESQ   = ""


  CCGINCODx   = ""
  CCGINTIPx   = ""

  sUsuario   = Session("Usuario")


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  sopcion  = request.querystring("opcion") & ""
  speriodo = request.querystring("periodo") & ""
  CantRegMostrarx = 10
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_PERIODO=" & speriodo

  sHTML = ""

  'Encabezados


  Formulario = ""


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  WKSETTABLA = "ORPERIODOS"


  if sopcion = "" then sopcion = "MOD"



  if sopcion = "SI" then


     response.buffer = true


     sinifin = request.querystring("accion") & ""
     sfechai = clng(date()) - clng(cdate("01-01-1900"))


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  ORPERIODOBEV  = request.FORM("PERIODOBEV") & ""
  ORPERIODOBAI  = request.FORM("PERIODOBAI") & ""

  TTMAP = 0
  ITMAP = 0
  CTMAP = 0
  LTMAP = 0

  DIM TABPROCESOSMAP (500,10)




         Strsql2 = "Select OBJID,PERIODO,RGVALEAUTOEV,RGVALEACTINF,RGAMBAUTOEV,RGAMBACTINF,RGINSTANCIA From ORPERIODOS  where APLIC = '" & wkaplic & "' and OBJID = 0" & speriodo
         set rsX1 = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
         if NOT rsX1.eof then
            WKORPERIODODSC     = RSX1(1)
            WKORPERIODOMAPVEV  = RSX1(2)
            WKORPERIODOMAPVAI  = RSX1(3)
            WKORPERIODOMAPNEV  = RSX1(4)
            WKORPERIODOMAPNAI  = RSX1(5)
            WKORPERIODOINSTAN  = RSX1(6)
         ELSE
            WKORPERIODODSC  = "NO ENCONTRADO"
            WKORPERIODOMAPVEV  = "N"
            WKORPERIODOMAPVAI  = "N"
            WKORPERIODOMAPNEV  = 0
            WKORPERIODOMAPNAI  = 0
            WKORPERIODOINSTAN  = "DEFAULT"
         END IF




''''''''''' generapa nuevo arbol Procesos - (autoevaluacion)

  IF WKORPERIODOMAPVEV  = "S" THEN

     WKORPERIODOEVSIGUE  = "N"
     WKFCODAMBIENTEEV      = WKORPERIODOMAPNEV

    IF  WKORPERIODOINSTAN  <> "DEFAULT" THEN

        WKFNOMAMBIENTE = "MAPA PROCESOS - " & WKORPERIODODSC
        WKFCODAMBIENTEEV = 0
        WKFVIGEN       = (CLNG(DATE()) - CLNG(CDATE("01-01-1900")))


         Strsql2 = "Select MAX(CODAMBIENTE) From RGAMBIENTES"
         set rsX1 = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
         if NOT rsX1.eof then
            WKFCODAMBIENTEEV = rsX1(0) + 1
         END IF

        strsql = "Insert into RGAMBIENTES "
        strsql = strsql & " (CODAMBIENTE,TIPOAMBIENTE,DESCRI,OWNER,FECCRE,"
        strsql = strsql & " FECVIGEN,FECBAJA,ESTADO,APLIC,segct,"
        strsql = strsql & " segmo,segco,segce,RGMODO"
        strsql = strsql & " )"
        strsql = strsql & "  Values "
        strsql = strsql &  "('" & WKFCODAMBIENTEEV & "','PROCESO','" & WKFNOMAMBIENTE & "','" & sUsuario & "','" & WKFVIGEN & "',"
        strsql = strsql &  "'" & WKFVIGEN & "','0','3','RG','0',"
        strsql = strsql &  "'0','0','0','RGPRC'"
        strsql = strsql &  ")"
        xxx=SqlExLog("INSERT",strsql,CN,"RGAMBIENTES","","","PAR")
 '' response.write strsql & "<br>"
       '' call ofv.crearconsultaEx(StrSql,cn,1,volver)

         Strsql2 = "Select OBJID From RGAMBIENTES WHERE CODAMBIENTE = 0" & WKFCODAMBIENTEEV
         set rsX1 = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
         if NOT rsX1.eof then
            WKORPERIODOEVSIGUE  = "S"
         END IF

    END IF

    IF  WKORPERIODOEVSIGUE  = "XS" THEN


        WKORPERIODOMAPNEV = WKFCODAMBIENTEEV

         Strsql2 = "Select OBJID,PERIODO,RGAMBAUTOEV From ORPERIODOS  where APLIC = '" & wkaplic & "' and OBJID = 0" & ORPERIODOBEV
         set rsX1 = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
         if NOT rsX1.eof then
            WKORPERIODODSCBEV  = RSX1(1)
            WKORPERIODOMAPBEV  = RSX1(2)
         ELSE
            WKORPERIODODSCBEV  = "NO ENCONTRADO"
            WKORPERIODOMAPBEV  = 0
         END IF

         RESPONSE.WRITE " OK: AMBIENTEN:" & WKORPERIODOMAPNEV & " - AMBIENTEB:" & WKORPERIODOMAPBEV & "<BR>"


    END IF



    IF  WKORPERIODOEVSIGUE  = "S" THEN

        WKORPERIODOMAPNEV = WKFCODAMBIENTEEV



         Strsql2 = "Select OBJID,PERIODO,RGAMBAUTOEV From ORPERIODOS  where APLIC = '" & wkaplic & "' and OBJID = 0" & ORPERIODOBEV
         set rsX1 = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
         if NOT rsX1.eof then
            WKORPERIODODSCBEV  = RSX1(1)
            WKORPERIODOMAPBEV  = RSX1(2)
         ELSE
            WKORPERIODODSCBEV  = "NO ENCONTRADO"
            WKORPERIODOMAPBEV  = 0
         END IF


            TTMAP = 0
            ITMAP = 0
            CTMAP = 0
            LTMAP = 0


            TABPROCESOSMAP (ITMAP,1) = WKORPERIODOMAPNEV
            TABPROCESOSMAP (ITMAP,2) = "PROCESO"

            TABPROCESOSMAP (ITMAP,3) = WKORPERIODOMAPBEV
            TABPROCESOSMAP (ITMAP,4) = "PROCESO"



''''' MASTER LOOP DE PROCESOS ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

      scodmaster = WKORPERIODOMAPBEV
      stipmaster = "PROCESO"

      Strsql = "Select OBJETOS.* From relaobj, OBJETOS  where CODOJASOC = CODOBJETO "
      Strsql = Strsql & " AND codobjmaster = '" & scodmaster & "' and tipobjmaster = '" & stipmaster & "'  "
      Strsql = Strsql & " ORDER BY OBJETOS.CODOBJETO "
 '''     RESPONSE.WRITE STRSQL & "<br>" & "<br>"

      set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if NOT rsX1.eof then

         DO UNTIL RSX1.EOF

            WKFPROCESO = RSX1(0)


            Celda =  ofv.gencelda("","UTI style='background-color:#8888ff;color:#ffffff' COLSPAN=4","","1%","","","","PROCESO :&nbsp;" & RSX1(0) & "-" & RSX1(3),"si")
            xFila  = xFILA & ofv.genrow("","aaa bgcolor='#f9f9f9' ","","","","","",Celda,"si")

            Celda =  ofv.gencelda(""," COLSPAN=4","","","","","","&nbsp;","si")
            xFila  = xFILA & ofv.genrow(""," ","","","","","",Celda,"si")

            Celda = ofv.gencelda("","uti  " ,"","","","","","<b>Nº</b>","si")
            Celda = Celda & ofv.gencelda("","UTI " ,"","20%","","","","<b>SUBPROCESO<b/>","si")
            Celda = Celda & ofv.gencelda("","UTI " ,"","","","","","<b>Nº</b>","si")
            Celda = Celda & ofv.gencelda("","UTI " ,"","","","","","<b>ACTIVIDADES</b>","si")

       '     xFila  = xFILA & ofv.genrow("","aaa bgcolor='#F7EEDD' ","","","","","",Celda,"si")




            strsql =  "select   B.CODOBJETO,B.DESCRI,B.CODVINOBJ, B.TIPOBJETO, A.*    "
            strsql =  strsql &"     from csrelpathdescri a  "
            strsql =  strsql &"    inner join objetos b on (a.cobj = b.codobjeto)  "
            Strsql =  Strsql & "   inner join objprty n on (b.codobjeto=n.codobjeto and estado=3)"
            strsql =  strsql &"          where     ( "
            strsql =  strsql &"          A.MC1 =" & WKFPROCESO & " OR  A.MC2 = " & WKFPROCESO & " OR  A.MC3 = " & WKFPROCESO
            strsql =  strsql &"          OR  A.MC4 = " & WKFPROCESO & " OR  A.MC5 = " & WKFPROCESO & " OR  A.MC6 = " & WKFPROCESO & " OR  "
            strsql =  strsql &"          A.MC7 = " & WKFPROCESO & " OR  A.COBJ =" & WKFPROCESO & "   ) "
            strsql =  strsql &"          AND (  B.APLIC = 'RG' ) "
            strsql =  strsql &"           ORDER BY  A.TSEC"
    '''  RESPONSE.WRITE STRSQL & "<br>" & "<br>"
            set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
            if NOT rsX2.eof then

               DO UNTIL RSX2.EOF

                  WKFTSEC = "" & RSX2("TSEC")

                  WKFTSECN = LEN(WKFTSEC)

                  WKFTSECN = CINT(WKFTSECN / 7) - 1

                  WKFTSECD = "" & RSX2(1)

                  FOR III=2 TO  WKFTSECN

                      WKFTSECD = "--" & WKFTSECD

                  NEXT


                  WKFTSECTAB = SPLIT(WKFTSEC,".")


                  Celda = ofv.gencelda("","uti  " ,"","5%","","","",RSX2(0),"si")
                  Celda = Celda & ofv.gencelda("","UTI " ,"","70%","","","",WKFTSECD,"si")
                  Celda = Celda & ofv.gencelda("","UTI " ,"","15%","","","",RSX2("TSEC"),"si")
                  Celda = Celda & ofv.gencelda("","UTI " ,"","5%","","","",RSX2(2),"si")

                  xFila  = xFILA & ofv.genrow("","aaa bgcolor='#EEEEEE' ","","","","","",Celda,"si")

                  Iraiz = RSX2(0)
                  ICODVN = RSX2(2) & ""


                  if (ICODVN = "SP" OR  ICODVN = "SPI" OR ICODVN = "CR" OR ICODVN = "VN" OR ICODVN = "AI" ) then


                     ITMAPms = WKFTSECN - 1


                      CCGIOCOD   = RSX2(0)
                      CCGIOTIP   = RSX2(3)

                      CCGOCESQ   = TABPROCESOSMAP (ITMAPms,1)
                      CCGOTESQ   = TABPROCESOSMAP (ITMAPms,2)

                      CCGINCOD   = ""
                      CCGINTIP   = ""

                      CCGICESQ   = TABPROCESOSMAP (ITMAPms,3)
                      CCGITESQ   = TABPROCESOSMAP (ITMAPms,4)





                     CALL RGGENERAPROCESO(ICODVN,CCGIOCOD,CCGIOTIP,CCGICESQ,CCGITESQ,CCGOCESQ,CCGOTESQ)

                   '   CCGINCODx   = RSX2(0) + 800000
                    '  CCGINTIPx   = "N-" & RSX2(3)


                     ITMAP = WKFTSECN

                     TABPROCESOSMAP (ITMAP,1) = CCGINCODx
                     TABPROCESOSMAP (ITMAP,2) = CCGINTIPx

                     TABPROCESOSMAP (ITMAP,3) = RSX2(0)
                     TABPROCESOSMAP (ITMAP,4) = RSX2(3)



                  Celda = ofv.gencelda("","uti  " ,"","5%","","","",TABPROCESOSMAP (ITMAP,1) & "/" & TABPROCESOSMAP (ITMAP,2) & "-----------------" & TABPROCESOSMAP (ITMAP,3) & "/" & TABPROCESOSMAP (ITMAP,4),"si")
                  Celda = Celda & ofv.gencelda("","UTI " ,"","70%","","","","dep." & ITMAPms & "." ,"si")
                  Celda = Celda & ofv.gencelda("","UTI " ,"","15%","","","",TABPROCESOSMAP (ITMAPms,1) & "/" & TABPROCESOSMAP (ITMAPms,2) & "-----------------" & TABPROCESOSMAP (ITMAPms,3) & "/" & TABPROCESOSMAP (ITMAPms,4),"si")
                  Celda = Celda & ofv.gencelda("","UTI " ,"","5%","","","","ubi." & ITMAP & ".","si")

                  xFila  = xFILA & ofv.genrow("","aaa bgcolor='#EEEEEE' STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:2;BORDER-COLOR:#990000;' ","","","","","",Celda,"si")





                  END IF


                  RSX2.MOVENEXT
                LOOP

             END IF



            RSX1.MOVENEXT
         LOOP

      END IF




 ''  sHTML = shtml & xshtml &  ofv.GenTabla("","AAA  border=1 " & wkstylo1 ,"","100% ","","0","","0","-1","","",xfila,"si")
   call ofv.cerrarconsulta(RSX2)

  END IF

  END IF


''''''''''' generapa nuevo INVENTARIO DE ACTIVOS - (RG 4609)

  IF WKORPERIODOMAPVAI  = "S" THEN

     WKORPERIODOAISIGUE  = "N"
     WKFCODAMBIENTEAI      = WKORPERIODOMAPNAI

    IF  WKORPERIODOINSTAN  <> "DEFAULT" THEN

        WKFNOMAMBIENTE = "MAPA ACTIVOS - " & WKORPERIODODSC
        WKFCODAMBIENTEAI = 0
        WKFVIGEN       = (CLNG(DATE()) - CLNG(CDATE("01-01-1900")))


         Strsql2 = "Select MAX(CODAMBIENTE) From RGAMBIENTES"
         set rsX1 = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
         if NOT rsX1.eof then
            WKFCODAMBIENTEAI = rsX1(0) + 1
         END IF

        strsql = "Insert into RGAMBIENTES "
        strsql = strsql & " (CODAMBIENTE,TIPOAMBIENTE,DESCRI,OWNER,FECCRE,"
        strsql = strsql & " FECVIGEN,FECBAJA,ESTADO,APLIC,segct,"
        strsql = strsql & " segmo,segco,segce,RGMODO"
        strsql = strsql & " )"
        strsql = strsql & "  Values "
        strsql = strsql &  "('" & WKFCODAMBIENTEAI & "','ACTIVOS','" & WKFNOMAMBIENTE & "','" & sUsuario & "','" & WKFVIGEN & "',"
        strsql = strsql &  "'" & WKFVIGEN & "','0','3','RG','0',"
        strsql = strsql &  "'0','0','0','RGACT'"
        strsql = strsql &  ")"
        xxx=SqlExLog("INSERT",strsql,CN,"RGAMBIENTES","","","PAR")
 '' response.write strsql & "<br>"
       '' call ofv.crearconsultaEx(StrSql,cn,1,volver)

         Strsql2 = "Select OBJID From RGAMBIENTES WHERE CODAMBIENTE = 0" & WKFCODAMBIENTEAI
         set rsX1 = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
         if NOT rsX1.eof then
            WKORPERIODOAISIGUE  = "S"
         END IF

    END IF

    IF  WKORPERIODOAISIGUE  = "XS" THEN


        WKORPERIODOMAPNAI = WKFCODAMBIENTEAI

         Strsql2 = "Select OBJID,PERIODO,RGAMBACTINF From ORPERIODOS  where APLIC = '" & wkaplic & "' and OBJID = 0" & ORPERIODOBAI
         set rsX1 = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
         if NOT rsX1.eof then
            WKORPERIODODSCBAI  = RSX1(1)
            WKORPERIODOMAPBAI  = RSX1(2)
         ELSE
            WKORPERIODODSCBAI  = "NO ENCONTRADO"
            WKORPERIODOMAPBAI  = 1
         END IF

         RESPONSE.WRITE " OK: AMBIENTEN:" & WKORPERIODOMAPNAI & " - AMBIENTEB:" & WKORPERIODOMAPBAI & "<BR>"


    END IF


    IF  WKORPERIODOAISIGUE  = "S" THEN

        WKORPERIODOMAPNAI = WKFCODAMBIENTEAI

         Strsql2 = "Select OBJID,PERIODO,RGAMBACTINF From ORPERIODOS  where APLIC = '" & wkaplic & "' and OBJID = 0" & ORPERIODOBAI
         set rsX1 = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
         if NOT rsX1.eof then
            WKORPERIODODSCBAI  = RSX1(1)
            WKORPERIODOMAPBAI  = RSX1(2)
         ELSE
            WKORPERIODODSCBAI  = "NO ENCONTRADO"
            WKORPERIODOMAPBAI  = 1
         END IF


            TTMAP = 0
            ITMAP = 0
            CTMAP = 0
            LTMAP = 0


            TABPROCESOSMAP (ITMAP,1) = WKORPERIODOMAPNAI
            TABPROCESOSMAP (ITMAP,2) = "ACTIVOS"

            TABPROCESOSMAP (ITMAP,3) = WKORPERIODOMAPBAI
            TABPROCESOSMAP (ITMAP,4) = "ACTIVOS"

    END IF


  END IF


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

     if sinifin = "OK" then

        Strsql = "update " & WKSETTABLA & " set ESTADO = 'OK', INICIO = '" & sfechai & "', RGAMBAUTOEV  = '" & WKORPERIODOMAPNEV & "' ,RGAMBACTINF  = '" & WKORPERIODOMAPNAI & "' "
        Strsql = Strsql & " where objid = " & speriodo
      response.write strsql & "<br>" & "<br>" & "<br>"

        call ofv.crearconsultaEx(StrSql,cn,1,parametros)


     end if

    RESPONSE.REDIRECT "ORPERIODOS.ASP"


  ELSE

  sHTML = ofv.FormHeader(session("FORMN"))



  Strsql = "Select * From " & WKSETTABLA & " WHERE OBJID = 0" & speriodo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  posi = 0
  IF Not rs.EOF THEN
     sForm = "form" & posi

    Fila = ""


    if sopcion = "SET" then




       sInput2 = "<INPUT TYPE=BUTTON ONCLICK='document." & sForm & "" & ".submit();' VALUE='CONFIRMAR INICIO'  STYLE='background-color:#990000;color:#ffffff;' >"


       Celda = ""
       Celda = ofv.gencelda("","VT COLSPAN=2 STYLE='COLOR:#990000;'  ","CENTER","","","","","CONFIRMA INICIAR EL PERIODO CON LA SIGUIENTE CONFIGURACION?","si")
       Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")



       Celda = ""
       Celda = ofv.gencelda("","UT COLSPAN=2 style='BORDER-TOP-COLOR:#888888;BORDER-TOP-WIDTH:1;BORDER-TOP-STYLE:SOLID;' ","CENTER","","","","","&nbsp;","si")
       Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

      sInput3 = ""

    ELSE

      sAccion = "ORPERIODOSDET.asp?CantRegMover=" & CantRegMoverx & "&PERIODO=" & rs(0)
      sAccion = sAccion & "&OPCION=SET"
      sInput2 = "<INPUT TYPE=BUTTON ONCLICK='document." & sForm & "OK" & ".submit();' VALUE='INICIAR'  STYLE='background-color:#009900;color:#ffffff;' >"

      sInput3 = ofv.genform(sForm & "OK","",sAccion,"","post","","","si")

    END IF






     Celda = ""
    'Descripcion
    Celda  = Celda & ofv.GenCelda("","UT style='background-color:#888888;color:#ffffff;' ","","20%","","","","NOMBRE","si")

    if sopcion = "SET" then
       sInput = rs(1) & ""
    ELSE
       sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"50","dt","")
    END IF

    Celda = Celda & ofv.gencelda("","VTI","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


     Celda = ""
    'ESTADO
    Celda  = Celda & ofv.GenCelda("","UT style='background-color:#888888;color:#ffffff;' ","","","","","","ESTADO","si")

     SELECT CASE (RS(4) & "")
            CASE "NO"
                 sInput = "<FONT COLOR='#000099' >NO INICIADO</FONT>"


            CASE "OK"
                 sInput = "<FONT COLOR='#009900' >INICIADO</FONT>"


            CASE "CL"
                 sInput = "<FONT COLOR='#990000' >FINALIZADO</FONT>"


     END SELECT

    Celda = Celda & ofv.gencelda("","VTI ","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


     Celda = ""
    'fecha
    Celda  = Celda & ofv.GenCelda("","UT style='background-color:#888888;color:#ffffff;' ","","","","","","INICIO","si")
    if clng(rs(2)) = 0 then
       sInput1 = "------------"
    else
       sInput1 = cdate(clng(rs(2)) + clng(cdate("01-01-1900")) )
    end if

    sInput = sInput1
    Celda = Celda & ofv.gencelda("","VTI ","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

     Celda = ""
    'fecha
    Celda  = Celda & ofv.GenCelda("","UT style='background-color:#888888;color:#ffffff;' ","","","","","","FINAL","si")
    if clng(rs(3)) = 0 then
       sInput1 = "------------"
    else
       sInput1 = cdate(clng(rs(3)) + clng(cdate("01-01-1900")) )
    end if

    sInput = sInput1
    Celda = Celda & ofv.gencelda("","VTI ","","","","","",sInput,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

     Celda = ""
    Celda  = Celda & ofv.GenCelda("","UT style='background-color:#888888;color:#ffffff;' ","","","","","","INSTANCIA","si")

    if sopcion = "SET" then
       IF (rs(6) & "") = "MAPESP" THEN
          sInput = "MAPEO POR PERIODO"
       ELSE
          sInput = "MAPA DE PROCESO PRINCIPAL"
       END IF
    ELSE
       scomK = "DEFAULT" & chr(9) & "MAPESP"
       scomD = "MAPA DE PROCESO PRINCIPAL" & chr(9) & "MAPEO POR PERIODO"
       sInput = ofv.gencomboFX(sForm,rs(6).name,rs(6),scomK,scomD," ")
    END IF



    Celda = Celda & ofv.gencelda("","VTI","","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")




    if sopcion = "SET" then

 '' ORPERIODOBEV  = request.FORM("PERIODOBEV") & ""
 '' ORPERIODOBAI  = request.FORM("PERIODOBAI") & ""


       Celda = ""
       Celda  = Celda & ofv.GenCelda("","UT style='background-color:#888888;color:#ffffff;' ","","","","","","BASE AUTOEV.","si")

      Strsql = "Select OBJID,PERIODO,RGAMBAUTOEV, (99999 - FINAL) AS XFFINAL,INICIO, ESTADO From ORPERIODOS  where APLIC = '" & wkaplic & "' and "
      Strsql = Strsql & " ESTADO IN ('OK','CL') AND RGVALEAUTOEV = 'S' "
      Strsql = Strsql & " ORDER BY XFFINAL DESC,INICIO DESC, OBJID DESC"

      set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if NOT rsX1.eof then
         XTPPER          = RSX1(0)
      END IF


      sInput = ofv.GenerarCombo(SFORM,"PERIODOBEV",XTPPER,Strsql,cn,0,1,"no")

    Celda = Celda & ofv.gencelda("","VTI","","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

       Celda = ""
       Celda  = Celda & ofv.GenCelda("","UT style='background-color:#888888;color:#ffffff;' ","","","","","","BASE ACT.INF.","si")

      Strsql = "Select OBJID,PERIODO,RGAMBAUTOEV, (99999 - FINAL) AS XFFINAL,INICIO, ESTADO From ORPERIODOS  where APLIC = '" & wkaplic & "' and "
      Strsql = Strsql & " ESTADO IN ('OK','CL') AND RGVALEACTINF = 'S' "
      Strsql = Strsql & " ORDER BY XFFINAL DESC,INICIO DESC, OBJID DESC"

      set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if NOT rsX1.eof then
         XTPPER          = RSX1(0)
      END IF


      sInput = ofv.GenerarCombo(SFORM,"PERIODOBAI",XTPPER,Strsql,cn,0,1,"no")

    Celda = Celda & ofv.gencelda("","VTI","","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


    END IF


     Celda = ""
    Celda  = Celda & ofv.GenCelda("","UT style='background-color:#888888;color:#ffffff;' ","","","","","","AUTO EVAL.","si")

    if sopcion = "SET" then
       IF (rs(7) & "") = "S" THEN
          sInput = "<FONT COLOR='#008800' >SI</FONT>"
       ELSE
          sInput = "<FONT COLOR='#880000' >NO</FONT>"
       END IF
    ELSE
       scomK = "S" & chr(9) & "N"
       scomD = "SI" & chr(9) & "NO"
       sInput = ofv.gencomboFX(sForm,rs(7).name,rs(7),scomK,scomD," ")
    END IF



    Celda = Celda & ofv.gencelda("","VTI","","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


     Celda = ""
    Celda  = Celda & ofv.GenCelda("","UT style='background-color:#888888;color:#ffffff;' ","","","","","","ACT. IT.","si")

    if sopcion = "SET" then
       IF (rs(8) & "") = "S" THEN
          sInput = "<FONT COLOR='#008800' >SI</FONT>"
       ELSE
          sInput = "<FONT COLOR='#880000' >NO</FONT>"
       END IF
    ELSE
       scomK = "S" & chr(9) & "N"
       scomD = "SI" & chr(9) & "NO"
       sInput = ofv.gencomboFX(sForm,rs(8).name,rs(8),scomK,scomD," ")
    END IF



    Celda = Celda & ofv.gencelda("","VTI","","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


     Celda = ""
    Celda  = Celda & ofv.GenCelda("","UT style='background-color:#888888;color:#ffffff;' ","","","","","","MAPA PROCESOS","si")
    IF RS(6) <> "DEFAULT" AND (RS(9) & "") = "0" THEN
        sInput = "------"
    ELSE
       sInput = rs(9)

      Strsql = "Select A.* "
      Strsql = Strsql & " From RGAMBIENTES A"
      Strsql = Strsql & " where A.CODAMBIENTE = '" & rs(9) & "' and A.TIPOAMBIENTE = 'PROCESO'  "
 '''  response.write strsql & "<br>" & chr(13)
      set rsxx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if NOT rsxx.eof then
         sInput = rs(9) & "&nbsp;-&nbsp;" & RSxx(3)
      END IF

    END IF

    Celda = Celda & ofv.gencelda("","VTI","","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

     Celda = ""
    Celda  = Celda & ofv.GenCelda("","UT style='background-color:#888888;color:#ffffff;' ","","","","","","MAPA ACTIVOS","si")
    IF RS(6) <> "DEFAULT" AND (RS(10) & "") = "1" THEN
        sInput = "------"
    ELSE

       sInput = rs(10)

      Strsql = "Select A.* "
      Strsql = Strsql & " From RGAMBIENTES A"
      Strsql = Strsql & " where A.CODAMBIENTE = '" & rs(10) & "' and A.TIPOAMBIENTE = 'ACTIVOS'  "
 '''  response.write strsql & "<br>" & chr(13)
      set rsxx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if NOT rsxx.eof then
         sInput = rs(10) & "&nbsp;-&nbsp;" & RSxx(3)
      END IF
    END IF

    Celda = Celda & ofv.gencelda("","VTI","","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


    Celda = ""
    Celda = ofv.gencelda("","UT COLSPAN=2 ","CENTER","","","","","&nbsp;","si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")



     Celda = ""
    Celda = ofv.gencelda("","UT COLSPAN=2 style='BORDER-TOP-COLOR:#888888;BORDER-TOP-WIDTH:1;BORDER-TOP-STYLE:SOLID;' ","CENTER","","","","",sInput2,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    Formulario = Fila


    if sopcion = "SET" then

      sAccion = ""
      if bModificar then
         sAccion = "ORPERIODOSDET.asp?CantRegMover=" & CantRegMoverx & "&PERIODO=" & rs(0)
         sAccion = sAccion & "&OPCION=SI&ACCION=OK"

         Formulario = ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")
      end if

    ELSE

      sAccion = ""
      if bModificar then
         sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
         sAccion = sAccion & "&tabla=" & WKSETTABLA & "&volver=" & vuelta
         Formulario = ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")
      end if

    END IF

  END IF
  sHTML = sHTML & "<CENTER>" & ofv.GenTabla("","AAA style='BORDER-COLOR:#888888;BORDER-WIDTH:1;BORDER-STYLE:SOLID;'  ","","70%","","0","","0","0","","",Formulario,"si") & sInput3 & "</CENTER>"

  call ofv.cerrarconsulta(rs)

 END IF

  'Botonera
  Agregar = ""
  Buscar  = ""

    if sopcion = "SET" then
       sAccion = "ORPERIODOSDET.asp?CantRegMover=" & CantRegMoverx & "&PERIODO=" & SPERIODO
       sAccion = sAccion & "&OPCION=MOD"
    ELSE
       sAccion = "ORPERIODOS.ASP"
    END IF


  VOLVER = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no",posi,"","","no",Agregar,VOLVER,"")


  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

function RGGENERAPROCESO(SOPCIONRG,RGCODENTRA,RGTIPENTRA,RGCODMASENTRA,RGTIPMASENTRA,RGCODMASSALE,RGTIPMASSALE)


  sTipobjeto = RGTIPENTRA


'''''''''''''

  CCGIOCOD = RGCODENTRA
  CCGIOTIP = RGTIPENTRA

  CCGINCOD = 0
  CCGINTIP = RGTIPENTRA

  CCGICESQ = RGCODMASSALE
  CCGITESQ = RGTIPMASSALE

  CCGOCESQ = RGCODMASENTRA
  CCGOTESQ = RGTIPMASENTRA




     Strsql = "Select * From OBJETOS  where CODOBJETO = " & CCGIOCOD
     Strsql = Strsql & " and tipobjeto = '" & CCGIOTIP & "'  "

 '' response.write strsql & "<br>"
     set RSOBJIN = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if NOT RSOBJIN.eof then

        CCGIEMPR = RSOBJIN(7)
        CCGINPRF = "OBJMAIN"
        CCGIDESC = RSOBJIN(3)
        CCGICDID = RSOBJIN(1)
        CCGICREL = RSOBJIN(15)

     END IF


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' VERIFICA EL PERFIL DEL DOCUMENTO ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

          bSector        = "NO"
          bEvento        = "NO"
          bProcedimiento = "NO"
          bControl       = "NO"
          bEsquema       = "NO"
          bTablas        = "NO"
          bDocum         = "NO"
          bEditor        = "NO"
          bDescri        = "NO"
          bMain          = "NO"

       strsql = "Select * From Objperfildoc Where codobjeto =" & CCGIOCOD
'     response.write "sql " & strsql
       Set rsPerfil = OFV.crearconsultaEx(strsql, cn, 1, volver)
       If rsPerfil.EOF Then
          perfilsp = "N"
'          response.write "no "

          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
          '''' VERIFICA EL PERFIL GENERAL DEL OBJETO ''''''
          ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


          strsql = "Select Objperfil.* From Objperfil, Tipobjeto where "
          strsql = strsql & "Objperfil.tipoperfil = Tipobjeto.perfil "
          strsql = strsql & "and Tipobjeto.tipobjeto = '" & CCGIOTIP & "'"
          Set rsPerfil = OFV.crearconsultaEx(strsql, cn, 1, volver)

       Else
'         response.write "si "
          perfilsp = "s"

       End If

       If NOT rsPerfil.EOF Then

          If rsPerfil("objsector") Then          bSector = "SI"
            'Sectores

          If rsPerfil("objeven") Then            bEvento = "SI"
            'Eventos

          If rsPerfil("objproced") Then          bProcedimiento = "SI"
            'Procedimientos

          If rsPerfil("objcontrol") Then         bControl = "SI"
            'Controles

          If rsPerfil("objesquema") Then         bEsquema = "SI"
            'Esquemas

          If rsPerfil("objtabla") Then           bTablas = "SI"
            'Tablas

          If rsPerfil("objhtm") Then             bDocum = "SI"
            'Documentos

          If rsPerfil("objobj1") Then            bEditor = "SI"
            'Editor estructurado

          If rsPerfil("objaccion") Then          bDescri = "SI"
            'Descripcion

          If rsPerfil("objmain") Then            bMain = "SI"
            'Consideraciones Generales

        END IF


        Call OFV.cerrarconsulta(rsPerfil)
     ''''   Call OFV.cerrarconn(cn)





'''''''''''''''''' invoca generacion del objeto (BLKOBJ)

        xxxx = CopyGenFn("BLKOBJ","NO")


        CCGINCODx   = CCGINCOD
        CCGINTIPx   = CCGINTIP


'''''''''''''''''' invoca generacion del relaciones (BLKREL)

        xxxx = CopyGenFn("BLKREL","NO")


'''''''''''''''''' invoca generacion propietario del docuemnto (BLKAUT)

    ''''    xxxx = CopyGenFn("BLKAUT","NO")


'''''''''''''''''' invoca generacion PERMISOS del docuemnto (BLKSDK)

        xxxx = CopyGenFn("BLKSDK","NO")


'''''''''''''''''' invoca generacion propiedades del docuemnto (BLKPRT)

        xxxx = CopyGenFn("BLKPRT","NO")

'''''''''''''''''' invoca generacion perfil del docuemnto (BLKPRF)

        xxxx = CopyGenFn("BLKPRF","NO")


'''''''''''''''''' invoca generacion agrupamiento y segmento (BLKAGS)

        xxxx = CopyGenFn("BLKAGS","NO")

'''''''''''''''''' invoca generacion origen de normativa (BLKORG)

        xxxx = CopyGenFn("BLKORG","NO")


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''' seccion de bloques segun perfil '''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' DESCRIPCION ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        If bDescri = "SI" Then

'''''''''''''''''' invoca generacion bloque descripcion (BLKDSC)

           xxxx = CopyGenFn("BLKDSC","NO")
        End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' SECTORES ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Sectores
        If bSector = "SI" Then
'''''''''''''''''' invoca generacion bloque sectores (BLKSEC)

           xxxx = CopyGenFn("BLKSEC","NO")
        End If



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' EVENTOS ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        If bEvento = "SI" Then
'''''''''''''''''' invoca generacion bloque eventos (BLKEVE)

           xxxx = CopyGenFn("BLKEVE","NO")
        End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' PROCEDIMIENTOS ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

      If bProcedimiento = "SI" Then

'''''''''''''''''' invoca generacion bloque procedimientos (BLKPCD)

           xxxx = CopyGenFn("BLKPCD","NO")


      End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' DOCUMENTOS ASOCIADOS ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        If bDocum = "SI" Then
'''''''''''''''''' invoca generacion bloque docuemntos asociados (BLKDOC)

           xxxx = CopyGenFn("BLKDOC","NO")
      End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' TABLAS ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      If bTablas = "SI" Then
'''''''''''''''''' invoca generacion bloque tablas (BLKTAB)

           xxxx = CopyGenFn("BLKTAB","NO")

      End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' EDITOR ESTRUCTURADO ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      If bEditor = "SI" Then
'''''''''''''''''' invoca generacion bloque contenido (BLKEST)

           xxxx = CopyGenFn("BLKEST","NO")

      End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' CONTROLES ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      If bControl = "SI" Then
'''''''''''''''''' invoca generacion bloque descripcion (BLKCNT)

           xxxx = CopyGenFn("BLKCNT","NO")
      End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' ESQUEMAS CONTABLES ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      If bEsquema = "SI" Then
'''''''''''''''''' invoca generacion bloque descripcion (BLKESQ)

           xxxx = CopyGenFn("BLKESQ","NO")
      End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' DOCUMENTOS REDIRECCIONADOS ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''' invoca generacion bloque descripcion (BLKFAQ)

           xxxx = CopyGenFn("BLKFAQ","NO")



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''' CONSIDERACIONES GENERALES ''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If bMain = "SI" Then

'''''''''''''''''' invoca generacion bloque descripcion (BLKCON)

           xxxx = CopyGenFn("BLKCON","NO")


    End If




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' FUNCIONES ESPECIALES PARA PROCESOS Y SUBPROCESOS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

  '''  Set cn = OFV.conectar(OFV.strconn2)

''  CCGIOCOD = nCodigo
''  CCGIOTIP = sTipo
''  CCGINCOD = 0
''  CCGINTIP = sTipobjeto
''  CCGICESQ = nCodESQS
''  CCGITESQ = sTipESQS
''  CCGOCESQ = nCodESQSO
''  CCGOTESQ = sTipESQSO


                         strsql = "update OBJETOS set  "
                         strsql = strsql & " OBJREM = " & CCGIOCOD & ", APLIC='RG' "
                         strsql = strsql & " where  codobjeto = " & CCGINCOD
         ''' response.write strsql & "<br>"
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)

                         strsql = "update OBJPRTY set "
                         strsql = strsql & " ESTADO = 3, VERSION = 'V', NROVER = '1' "
                         strsql = strsql & " where  codobjeto = " & CCGINCOD
         ''' response.write strsql & "<br>"
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)

                         strsql = "update OBJETOS set OBJRDO = " & CCGINCOD
                         strsql = strsql & " where  codobjeto = " & CCGIOCOD
        '''  response.write strsql & "<br>"
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)




       IF (SOPCIONRG & "") = "CR" or (SOPCIONRG & "") = "AI" THEN

          CCGICESQ = CCGINCOD
          CCGITESQ = CCGINTIP

          CCGOCESQ = CCGIOCOD
          CCGOTESQ = CCGIOTIP

          wlnkid   = 0

          Strsql = "SELECT OBJID,codobjetoL, tipobjetoL,IDEN4, LTIPO From OBJLNK where iden1 = '" & CCGOCESQ & "' and iden2 = '" & CCGOCESQ & "' "
          strsql = strsql & " AND (LTipo = 'LRGVN' OR LTipo = 'LRGCR' ) "
          set rsVINC = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if not rsVINC.eof then
             wlnkid = rsVINC(0)
             wlnkidX1 = rsVINC(1)
             wlnkidX2 = rsVINC(2)
             wlnkidX3 = rsVINC(3)
             wlnkTIPL = rsVINC(4)

             CCGIOCOD = wlnkidX1
             CCGIOTIP = wlnkidX2
             CCGINTIP = wlnkidX2

          end if

          if wlnkid   > 0 then


          '''''''''''''''''' invoca generacion del objeto (BLKOBJ)

             xxxx = CopyGenFn("BLKOBJ","NO")


          '''''''''''''''''' invoca generacion del relaciones (BLKREL)

             xxxx = CopyGenFn("BLKREL","NO")


          '''''''''''''''''' invoca generacion propietario del docuemnto (BLKAUT)

         ''''    xxxx = CopyGenFn("BLKAUT","NO")


          '''''''''''''''''' invoca generacion PERMISOS del docuemnto (BLKSDK)

             xxxx = CopyGenFn("BLKSDK","NO")


          '''''''''''''''''' invoca generacion propiedades del docuemnto (BLKPRT)

             xxxx = CopyGenFn("BLKPRT","NO")

          '''''''''''''''''' invoca generacion perfil del docuemnto (BLKPRF)

             xxxx = CopyGenFn("BLKPRF","NO")


          '''''''''''''''''' invoca generacion agrupamiento y segmento (BLKAGS)

             xxxx = CopyGenFn("BLKAGS","NO")

          '''''''''''''''''' invoca generacion origen de normativa (BLKORG)

             xxxx = CopyGenFn("BLKORG","NO")

          '''''''''''''''''' invoca generacion bloque sectores (BLKSEC)

             xxxx = CopyGenFn("BLKSEC","NO")

          '''''''''''''''''' invoca generacion bloque eventos (BLKEVE)

             xxxx = CopyGenFn("BLKEVE","NO")

          '''''''''''''''''' invoca generacion bloque procedimientos (BLKPCD)

             xxxx = CopyGenFn("BLKPCD","NO")





              strsql = "SELECT TIPOBJETO,CODPROCED FROM OBJPROCED    "
              strsql = strsql & " WHERE CODOBJETO = 0" & CCGINCOD & " AND RELANT = '" & wlnkidX3 & "' "
       '''   response.write strsql & "<br>"
              set RSOBJOLD = ofv.crearconsultaEx(StrSql,cn,1,parametros)
              if NOT RSOBJOLD.eof then

                 NDOCUTIPO  = RSOBJOLD(0)
                 NPROCEDCOD = RSOBJOLD(1)

              END IF


                   wlink   = CCGINCOD
                   wlinkT  = CCGINTIP
                   wlinkN1 = "5"
                   wlinkN2 = NPROCEDCOD
                   wlinkN3 = "0"

                   wlinkDSC = CCGICESQ & "-" & CCGITESQ & "- VINCULACION EXPLICITA CR"



                      strsql = "Insert into OBJLNK (codobjeto, tipobjeto, iden1, iden2, codobjetoL, tipobjetoL, Lurl, LurlDesc, LTipo "
                      strsql = strsql & ",IDEN3L,IDEN4,IDEN4L) Values (" & CCGICESQ & ",'" & CCGITESQ & "',"
                      strsql = strsql & "'" & CCGICESQ & "','" & CCGICESQ & "',"
                      strsql = strsql & wlink & ",'" & wlinkT & "','','','" & wlnkTIPL & "','" & wlinkN1 & "','" & wlinkN2 & "','" & wlinkN3 & "'"
                      strsql = strsql & ")"
                      call ofv.crearconsultaEx(StrSql,cn,1,volver)

                      Strsql = "SELECT * From OBJLNK where iden1 = '" & CCGICESQ & "' and iden2 = '" & CCGICESQ & "' AND LTipo = '" & wlnkTIPL & "' "
                      set rsVINC2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                      if not rsVINC2.eof then
                         wlnkid = rsVINC2(0)

                         strsql = "update OBJETOS set rellink = " & wlnkid & ", APLIC='RG' "
                         strsql = strsql & " where  codobjeto = " & CCGICESQ
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)

                         strsql = "update OBJETOS set  APLIC='RG', DESCRI = '" & wlinkDSC & "' "
                         strsql = strsql & " where  codobjeto = " & CCGINCOD
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)


                         strsql = "update OBJPRTY set "
                         strsql = strsql & " ESTADO = 3, VERSION = 'V', NROVER = '1' "
                         strsql = strsql & " where  codobjeto = " & CCGICESQ
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)

                         strsql = "update OBJPRTY set "
                         strsql = strsql & " ESTADO = 3, VERSION = 'V', NROVER = '1' "
                         strsql = strsql & " where  codobjeto = " & CCGINCOD
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)


                         strsql = "update OBJPRTY set "
                         strsql = strsql & " ESTADO = 4, FECBAJA = '" & (CLNG(DATE()) - CLNG(CDATE("01-01-1900"))) & "' "
                         strsql = strsql & " where  codobjeto = " & CCGOCESQ
                     '''    call ofv.crearconsultaEx(StrSql,cn,1,volver) '''''''' NO PONE NO VIGENTE

                         strsql = "update OBJPRTY set "
                         strsql = strsql & " ESTADO = 4, FECBAJA = '" & (CLNG(DATE()) - CLNG(CDATE("01-01-1900"))) & "' "
                         strsql = strsql & " where  codobjeto = " & CCGIOCOD
                     '''    call ofv.crearconsultaEx(StrSql,cn,1,volver) '''''''' NO PONE NO VIGENTE


                      end if
                      call ofv.cerrarconsulta(rsVINC2)

           else


              Response.write "**!! Fallo reemplazo nueva version del subproceso: " & Minute(Now) & ":" & Second(Now) & "<br>"



           end if




       else


        Strsql = "SELECT OBJID,codobjetoL, tipobjetoL,IDEN4, LTIPO From OBJLNK where iden1 = '" & CCGIOCOD & "' and iden2 = '" & CCGIOCOD & "' "
        strsql = strsql & " AND (LTipo = 'LRGVN' OR LTipo = 'LRGCR' ) "
        set rsVINC = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsVINC.eof then
           wlnkid     = rsVINC(0)
           wlnkidX1   = rsVINC(1)
           wlnkidX2   = rsVINC(2)
           wlnkidX3   = rsVINC(3)
           wlnkTIPL   = rsVINC(4)

           NDOCUCOD   = wlnkidX1
           NDOCUTIPO  = wlnkidX2
           NPROCEDCOD = wlnkidX3


           wlink   = NDOCUCOD
           wlinkT  = NDOCUTIPO
           wlinkN1 = "5"
           wlinkN2 = NPROCEDCOD
           wlinkN3 = "0"


                      strsql = "Insert into OBJLNK (codobjeto, tipobjeto, iden1, iden2, codobjetoL, tipobjetoL, Lurl, LurlDesc, LTipo "
                      strsql = strsql & ",IDEN3L,IDEN4,IDEN4L) Values (" & CCGINCOD & ",'" & CCGINTIP & "',"
                      strsql = strsql & "'" & CCGINCOD & "','" & CCGINCOD & "',"
                      strsql = strsql & wlink & ",'" & wlinkT & "','','','" & wlnkTIPL & "','" & wlinkN1 & "','" & wlinkN2 & "','" & wlinkN3 & "'"
                      strsql = strsql & ")"
                      call ofv.crearconsultaEx(StrSql,cn,1,volver)

                      Strsql = "SELECT * From OBJLNK where iden1 = '" & CCGINCOD & "' and iden2 = '" & CCGINCOD & "' AND LTipo = '" & wlnkTIPL & "' "
                      set rsVINC2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                      if not rsVINC2.eof then
                         wlnkid = rsVINC2(0)

                         strsql = "update OBJETOS set rellink = " & wlnkid & ", APLIC='RG', "
                         strsql = strsql & " OBJREM = " & CCGIOCOD & " "
                         strsql = strsql & " where  codobjeto = " & CCGINCOD
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)

                         strsql = "update OBJPRTY set "
                         strsql = strsql & " ESTADO = 3, VERSION = 'V', NROVER = '1' "
                         strsql = strsql & " where  codobjeto = " & CCGINCOD
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)

                         strsql = "update OBJETOS set OBJRDO = " & CCGINCOD
                         strsql = strsql & " where  codobjeto = " & CCGIOCOD
                         call ofv.crearconsultaEx(StrSql,cn,1,volver)

                         strsql = "update OBJPRTY set "
                         strsql = strsql & " ESTADO = 4, FECBAJA = '" & (CLNG(DATE()) - CLNG(CDATE("01-01-1900"))) & "' "
                         strsql = strsql & " where  codobjeto = " & CCGIOCOD
                     ''    call ofv.crearconsultaEx(StrSql,cn,1,volver) '''''''''''''''''''''''''''' NO PONE NO VIGENTE


                      end if
                      call ofv.cerrarconsulta(rsVINC2)





       end if
        call ofv.cerrarconsulta(rsVINC)



    end if









end function



%>