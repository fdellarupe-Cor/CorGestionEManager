<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  sHTML = ofv.FormHeader("")


'''' VARIABLES DE FORMATO

 SFECHAIN  = REQUEST.QUERYSTRING("FCIN")   ' FECHA ACTUAL ------- FECHA DEL DIA SI NO VIENE

 SFYRANTES = REQUEST.QUERYSTRING("FCYRAN") ' AÑOS ANTES DE LA FECHA ---------- 0 SOLO EL AÑO DE LA FECHA / 99 EL AÑO DE LA FECHA Y EL SIGUIENTE SI SE PASA
 SFYRDESPS = REQUEST.QUERYSTRING("FCYRDP") ' AÑOS DESPUES DE LA FECHA -------- 0 SOLO EL AÑO DE LA FECHA / 99 EL AÑO DE LA FECHA Y EL ANTERIOR SI SE PASAN LOS MESES

 SFMNANTES = REQUEST.QUERYSTRING("FCMNAN") ' MESES ANTES DE LA FECHA --------- 0 SOLO EL MES DE LA FECHA / 98 TODOS LOS MESES / 99  Y EL ANTERIOR
 SFMNDESPS = REQUEST.QUERYSTRING("FCMNDP") ' MESES DESPUES DE LA FECHA ------- 0 SOLO EL MES DE LA FECHA / 99  Y EL POSTERIOR

 SFTDIAS   = REQUEST.QUERYSTRING("FCTDIA") ' 0 - DIAS CORRIDOS / 1 - DIAS HABILES 

 SFFLAG    = REQUEST.QUERYSTRING("FCFLAG") ' 0 - INICIO / 1 - TOMA DEL COMBO


''''''''''''''''''''''''''''
'' VALIDACION DE VARIABLES ''
'''''''''''''''''''''''''''''

 IF SFFLAG <> "" THEN

    IF NOT (ISNUMERIC(SFFLAG)) THEN

       SFFLAG = 0

    ELSE

       SFFLAG = CLNG(SFFLAG)

       IF SFFLAG = 1 THEN

          SFFLAGY    = REQUEST.FORM("FCFLAGY") ' AÑO DEL COMBO
          SFFLAGM    = REQUEST.FORM("FCFLAGM") ' MES DEL COMBO


       END IF 


    END IF


 ELSE

    SFFLAG = 0

 END IF





 IF SFECHAIN <> "" THEN

    IF NOT (ISNUMERIC(SFECHAIN)) THEN

       SFECHAIN = clng(DATE()) - clng(cdate("01-01-1900"))

    ELSE

       SFECHAIN = CLNG(SFECHAIN)

       IF SFECHAIN = 0  THEN SFECHAIN = clng(DATE()) - clng(cdate("01-01-1900"))

    END IF


 ELSE

    SFECHAIN = clng(DATE()) - clng(cdate("01-01-1900"))

 END IF





 IF SFYRANTES <> "" THEN

     IF NOT (ISNUMERIC(SFYRANTES)) THEN

       SFYRANTES = 0

    ELSE

       SFYRANTES = CLNG(SFYRANTES)

       IF SFYRANTES = 0 OR SFYRANTES = 99 THEN

       ELSEIF SFYRANTES < 9  THEN

       ELSE

              SFYRANTES = 0

       END IF 

    END IF

 ELSE

    SFYRANTES = 0

 END IF 


 IF SFYRDESPS <> "" THEN

     IF NOT (ISNUMERIC(SFYRDESPS)) THEN

       SFYRDESPS = 0

    ELSE

       SFYRDESPS = CLNG(SFYRDESPS)

       IF SFYRDESPS = 0 OR SFYRDESPS = 99 THEN


       ELSEIF SFYRDESPS < 9  THEN


       ELSE

              SFYRDESPS = 0

       END IF 

    END IF

 ELSE

    SFYRDESPS = 0

 END IF 

 IF SFMNANTES <> "" THEN

     IF NOT (ISNUMERIC(SFMNANTES)) THEN

       SFMNANTES = 0

    ELSE

       SFMNANTES = CLNG(SFMNANTES)

    END IF

 ELSE

    SFMNANTES = 0

 END IF 


 IF SFMNDESPS <> "" THEN

     IF NOT (ISNUMERIC(SFMNDESPS)) THEN

       SFMNDESPS = 0

    ELSE

       SFMNDESPS = CLNG(SFMNDESPS)

    END IF

 ELSE

    SFMNDESPS = 0

 END IF 


 IF SFTDIAS <> "" THEN

     IF NOT (ISNUMERIC(SFTDIAS)) THEN

       SFTDIAS = 0

    ELSE

       SFTDIAS = CLNG(SFTDIAS)

    END IF

 ELSE

    SFTDIAS = 0

 END IF 


''''''''''''''''''''''''''''''''


 '' sHTML = ""



    Celda = ""




 dim gtctabmes(12,12)

 dim gtctabAnO(10,5)


  set cn = ofv.conectar(ofv.strconn2)


 gtctabmes(1,0)  = "ENERO"
 gtctabmes(2,0)  = "FEBRERO"
 gtctabmes(3,0)  = "MARZO"
 gtctabmes(4,0)  = "ABRIL"
 gtctabmes(5,0)  = "MAYO"
 gtctabmes(6,0)  = "JUNIO"
 gtctabmes(7,0)  = "JULIO"
 gtctabmes(8,0)  = "AGOSTO"
 gtctabmes(9,0)  = "SEPTIEMBRE"
 gtctabmes(10,0) = "OCTUBRE"
 gtctabmes(11,0) = "NOVIEMBRE"
 gtctabmes(12,0) = "DICIEMBRE"

 gtctabmes(1,1)  = 31
 gtctabmes(2,1)  = 28
 gtctabmes(3,1)  = 31
 gtctabmes(4,1)  = 30
 gtctabmes(5,1)  = 31
 gtctabmes(6,1)  = 30
 gtctabmes(7,1)  = 31
 gtctabmes(8,1)  = 31
 gtctabmes(9,1)  = 30
 gtctabmes(10,1) = 31
 gtctabmes(11,1) = 30
 gtctabmes(12,1) = 31


 gtctabAnO(1,0)  = 2008
 gtctabAnO(2,0)  = 2007
 gtctabAnO(3,0)  = 2006
 gtctabAnO(4,0)  = 2005
 gtctabAnO(5,0)  = 2004

 gtcty = 0
 gtPOINTy = 0

 gtcrm = clng(cdate("01-01-1900"))





 GTWFECHAIN = cdate(SFECHAIN + gtcrm)

 gtcsanio   = datepart("yyyy",GTWFECHAIN)
 gtcsmes    = datepart("m",GTWFECHAIN)
 gtcsdia    = datepart("d",GTWFECHAIN)

 GTWAN      = "NO"
 GTWDS      = "NO"

 GTMDS      = 0
 GTMHS      = 0

'''' CARGA AÑOS ANTES

       IF SFYRANTES = 0  THEN

       ELSEIF SFYRANTES = 99  THEN

           IF SFMNANTES = 0 THEN

           ELSEIF SFMNANTES = 99 THEN

               if  gtcsmes = 1 then

                   GTMDS      = 12
                   GTMHS      = 12

                   GTWAN      = "SI"

               end if 


           ELSEIF SFMNANTES = 98 THEN

               if  gtcsmes < 6 then

                   GTMDS      = 12 - (12 - gtcsmes) - 1
                   GTMHS      = 12

                   GTWAN      = "SI"

               end if 

           ELSE



           END IF

           IF GTWAN      <> "NO" THEN

              gtcty = gtcty + 1
              GTWX = gtcsanio - 1

              IF SFFLAG = 1 THEN
                 IF cstr(SFFLAGY) = cstr(GTWX) THEN gtPOINTy = gtcty
              END IF

               gtctabAnO(gtcty,0) = GTWX
               gtctabAnO(gtcty,1) = 0
               gtctabAnO(gtcty,3) = 0
               gtctabAnO(gtcty,2) = 0
               gtctabAnO(gtcty,4) = GTMDS
               gtctabAnO(gtcty,5) = GTMHS

          END IF



       ELSE

           GTWYDESDE = gtcsanio - SFYRANTES
           GTWYHASTA = gtcsanio - 1

           for GTWX=GTWYDESDE TO GTWYHASTA

               gtcty = gtcty + 1

               IF SFFLAG = 1 THEN
                  IF cstr(SFFLAGY) = cstr(GTWX) THEN gtPOINTy = gtcty
               END IF


               gtctabAnO(gtcty,0) = GTWX
               gtctabAnO(gtcty,1) = 0
               gtctabAnO(gtcty,3) = 0
               gtctabAnO(gtcty,2) = 0
               gtctabAnO(gtcty,4) = 1
               gtctabAnO(gtcty,5) = 12

           NEXT


       END IF  


'''''' CARGA EL AÑO DE LA FECHA

       IF (SFYRANTES > 0  and SFYRANTES < 9) or (SFYRDESPS > 0  and SFYRDESPS < 9) THEN

                   GTMDS      = 1
                   GTMHS      = 12


       else 


           IF SFMNANTES = 0 THEN

              GTMDS      = gtcsmes

           ELSEIF SFMNANTES = 99 THEN

               if  gtcsmes > 1 then

                   GTMDS      = gtcsmes - 1

               end if 


           ELSEIF SFMNANTES = 98 THEN


                   GTMDS      = 1
                   GTMHS      = 12


           ELSE

                   GTMDS      = gtcsmes - SFMNANTES

                   IF GTMDS < 1 THEN GTMDS = 1

           END IF


           IF SFMNDESPS = 0 THEN

                   GTMHS      = gtcsmes

           ELSEIF SFMNDESPS = 99 THEN

               if  gtcsmes < 12 then

                   GTMHS      = gtcsmes + 1

               end if 



           ELSE

                   GTMHS      = gtcsmes + SFMNDESPS

                   IF GTMHS < 12 THEN GTMHS = 12

           END IF

         end if




               gtcty = gtcty + 1

              GTWX = gtcsanio

              IF SFFLAG = 1 THEN
                 IF cstr(SFFLAGY) = cstr(GTWX) THEN gtPOINTy = gtcty

  '''  response.write "compara-" & (SFFLAGY = GTWX) & " SFFLAGY-" & SFFLAGY & " gtcsanio-" & gtcsanio & " x" & gtctabAnO(gtciy,0) & " sf" & SFFLAG & " sfy" & SFFLAGy & " sfm" & SFFLAGm & chr(13)
              else
                 gtPOINTy = gtcty
              END IF

               gtctabAnO(gtcty,0) = GTWX
               gtctabAnO(gtcty,1) = 0
               gtctabAnO(gtcty,3) = 0
               gtctabAnO(gtcty,2) = 0
               gtctabAnO(gtcty,4) = cint(GTMDS)
               gtctabAnO(gtcty,5) = cint(GTMHS)

   '' response.write "desde-" & gtctabAnO(gtciy,4) & " hasta-" & gtctabAnO(gtciy,5) & " gtcsanio-" & gtcsanio & " x" & gtctabAnO(gtciy,0) & " sf" & SFFLAG & " sfy" & SFFLAGy & " sfm" & SFFLAGm & chr(13)


'''' CARGA AÑOS DESPUES

       IF SFYRDESPS = 0  THEN

       ELSEIF SFYRDESPS = 99  THEN

           IF SFMNDESPS = 0 THEN

           ELSEIF SFMNDESPS = 99 THEN

               if  gtcsmes = 12 then

                   GTMDS      = 1
                   GTMHS      = 1

                   GTWDS      = "SI"

               end if 



           ELSE



           END IF

           IF GTWDS      <> "NO" THEN

               gtcty = gtcty + 1

              GTWX = gtcsanio + 1

              IF SFFLAG = 1 THEN
                 IF cstr(SFFLAGY) = cstr(GTWX) THEN gtPOINTy = gtcty
              END IF

               gtctabAnO(gtcty,0) = GTWX
               gtctabAnO(gtcty,1) = 0
               gtctabAnO(gtcty,3) = 0
               gtctabAnO(gtcty,2) = 0
               gtctabAnO(gtcty,4) = GTMDS
               gtctabAnO(gtcty,5) = GTMHS

          END IF



       ELSE

           GTWYDESDE = gtcsanio + 1
           GTWYHASTA = gtcsanio + SFYRDESPS

           for GTWX=GTWYDESDE TO GTWYHASTA

               gtcty = gtcty + 1

               IF SFFLAG = 1 THEN
                  IF cstr(SFFLAGY) = cstr(GTWX) THEN gtPOINTy = gtcty
               END IF


               gtctabAnO(gtcty,0) = GTWX
               gtctabAnO(gtcty,1) = 0
               gtctabAnO(gtcty,3) = 0
               gtctabAnO(gtcty,2) = 0
               gtctabAnO(gtcty,4) = 1
               gtctabAnO(gtcty,5) = 12

           NEXT


       END IF  


 gtciy=gtPOINTy

     SCOMSEP = ""
     scomK   = ""
     scomD   = ""

     wky  = gtctabAnO(gtciy,0)
     wkyd = wky - 1900
     wkyb = fix(wkyd / 4)

     wkybt = fix(wky / 4) * 4

     if wkybt = wky then 
        wkybt = 0
        wkyb  = wkyb - 1
     else
        wkybt = 9
     end if

     wkyi = (365 * wkyd) + wkyb 
  
     gtctabAnO(gtciy,1) = wkyi
     gtctabAnO(gtciy,3) = wkyb
     gtctabAnO(gtciy,2) = wkybt

     wkyac = wkyi

     gtcimi = 3

     for gtcim=1 to 12



         gtctabmes(gtcim,2) = "N"


         IF not (gtcim < (gtctabAnO(gtciy,4))) AND not (gtcim > (gtctabAnO(gtciy,5))) THEN

            gtctabmes(gtcim,2) = "S"

            scomD = scomD & SCOMSEP & gtctabmes(gtcim,0)
            scomK = scomK & SCOMSEP & gtcim
            SCOMSEP = chr(9)

         END IF



         gtctabmes(gtcim,gtcimi) = wkyac

         wkmd = gtctabmes(gtcim,1)

         if wkybt = 0 and gtcim = 2 then wkmd = wkmd + 1

         wkyac = wkyac + wkmd

     next


    sForm = "FRMCALENDAR"

    gtcim = gtcsmes

    IF SFFLAG = 1 THEN
       gtcim = SFFLAGM
    END IF

    sInputMN = ofv.gencomboFX(sForm,"FCFLAGM",gtcim,scomK,scomD,"")



     SCOMSEP = ""
     scomK   = ""
     scomD   = ""

     For gtciy=1 to gtcty
         scomD = scomD & SCOMSEP & gtctabAnO(gtciy,0)
         scomK = scomK & SCOMSEP & gtctabAnO(gtciy,0)
         SCOMSEP = chr(9)
     next



    sInputYR = ofv.gencomboFX(sForm,"FCFLAGY",gtctabAnO(gtPOINTy,0),scomK,scomD,"")


    CELDAD = ""
    CeldaD = CeldaD & ofv.GenCelda("","VtI style='background-color:#999999;color:#ffffff;' ","","15%","","","","AÑO","si")
    CeldaD = CeldaD & ofv.GenCelda("","uti","","35%","","","",sInputYR,"si")
    CeldaD = CeldaD & ofv.GenCelda("","VtI style='background-color:#999999;color:#ffffff;' ","","15%","","","","MES","si")
    CeldaD = CeldaD & ofv.GenCelda("","uti","","35%","","","",sInputMN,"si")

    FilaD     = ofv.genrow("","","","","","","",CeldaD,"si")


    sAccion = "ORCALENDAR2.ASP?FCIN=" & SFECHAIN
    sAccion = sAccion & "&FCYRAN=" & SFYRANTES
    sAccion = sAccion & "&FCYRDP=" & SFYRDESPS
    sAccion = sAccion & "&FCMNAN=" & SFMNANTES
    sAccion = sAccion & "&FCMNDP=" & SFMNDESPS
    sAccion = sAccion & "&FCTDIA=" & SFTDIAS
    sAccion = sAccion & "&FCFLAG=1"

    Formulario = ofv.genform(sForm,"",sAccion,"","post","",Filad,"si")

    sFORMDAT = ofv.GenTabla("","","","100%","","0","","0","0","","",Formulario,"si")


    xtrcell = " style='border-style:SOLID;border-width:1;border-color:#aaaaaa;' " 

    dim gtcafech(40,1)
    dim gtcafechs(42,1)


    Fila = ""

    gtciy = gtPOINTy



    Celda = ofv.GenCelda("","vt colspan=7 ","","","","","","mes&nbsp;" & gtctabAnO(gtciy,0) & "&nbsp;" & gtctabmes(gtcim,0) & "&nbsp;dias&nbsp;" & gtctabAnO(gtciy,1)  & "&nbsp;bis&nbsp;" & gtctabAnO(gtciy,2)   & "&nbsp;dbis&nbsp;" & gtctabAnO(gtciy,3) ,"si")

 '   Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

    Celda = ""
    Celda = Celda & ofv.GenCelda("","VtI style='background-color:#999999;color:#ffffff;' ","","","","","","FECHA ","si")
    Celda = Celda & ofv.GenCelda("","VTI colspan=6 ","","","","","",GTFechador(SFECHAIN),"si")

    Fila = fila & ofv.genrow("","","","","25","","",Celda,"si")

    Celda = ofv.GenCelda("","ut colspan=7 style='border-style:SOLID;border-width:1;border-color:#993333;' ","CENTER","","","","",sFORMDAT,"si")

    Fila = fila & ofv.genrow("","","","","25","","",Celda,"si")


    CELDA = ""
    Celda = Celda & ofv.GenCelda("","Vt style='background-color:#f7aa22;color:#ffffff;border-style:outset;border-width:2;'","CENTER","10%","","","","D","si")
    Celda = Celda & ofv.GenCelda("","Vt style='background-color:#f7aa22;color:#ffffff;border-style:outset;border-width:2;'","CENTER","10%","","","","L","si")
    Celda = Celda & ofv.GenCelda("","Vt style='background-color:#f7aa22;color:#ffffff;border-style:outset;border-width:2;'","CENTER","10%","","","","M","si")
    Celda = Celda & ofv.GenCelda("","Vt style='background-color:#f7aa22;color:#ffffff;border-style:outset;border-width:2;'","CENTER","10%","","","","M","si")
    Celda = Celda & ofv.GenCelda("","Vt style='background-color:#f7aa22;color:#ffffff;border-style:outset;border-width:2;'","CENTER","10%","","","","J","si")
    Celda = Celda & ofv.GenCelda("","Vt style='background-color:#f7aa22;color:#ffffff;border-style:outset;border-width:2;'","CENTER","10%","","","","V","si")
    Celda = Celda & ofv.GenCelda("","Vt style='background-color:#f7aa22;color:#ffffff;border-style:outset;border-width:2;'","CENTER","10%","","","","S","si")

    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")



    CELDA = ""

   ''''''''''' desde


    gtcsfecd = gtctabmes(gtcim,3)

 '''   response.write "pointy-" & gtPOINTy & " fec-" & gtcsfecd & " mes " & gtcim & " x" & gtctabAnO(gtciy,0) & " sf" & SFFLAG & " sfy" & SFFLAGy & " sfm" & SFFLAGm & chr(13)

    sInput1 = cdate(gtcsfecd + gtcrm)

    gtcsanio1 = datepart("yyyy",sInput1)
    gtcsmes1  = datepart("m",sInput1)
    gtcsdia1  = datepart("d",sInput1)

    gtcswd = datepart("w",dateserial(gtcsanio1,gtcsmes1,gtcsdia1))

  '''  response.write "f " & gtcsanio1 & "-" & gtcsmes1 & "-" & gtcsdia1 & " d " & swd


   ''''''''''' hasta

    wkmd = gtctabmes(gtcim,1)
    if gtctabAnO(gtciy,2) = 0 and gtcim = 2 then wkmd = wkmd + 1

     
    gtcsfech = gtcsfecd + wkmd - 1


    Celda = ofv.GenCelda("","vt colspan=7 ","","","","","","desde&nbsp;" & GTFechador(gtcsfecd) & "&nbsp;hasta&nbsp;" & GTFechador(gtcsfech) & "&nbsp;dias&nbsp;" & wkmd,"si")

  '  Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


    gtcsrest = gtcrm
    gtctopefech = gtcsfech - gtcsfecd

    gtcyy = 0



    gtcwswd = gtcswd

    for gtcii = 0 to gtctopefech
        
        gtcafech(gtcii+1,0) = gtcwswd
        gtcwswd = gtcwswd + 1
        if gtcwswd > 7 then gtcwswd = 1        

        gtcafech(gtcii+1,1) = gtcsfecd + gtcii


    next


    for gtcii = 1 to (cint(gtcafech(1,0)) - 1)
        gtcafechs(gtcii,0) = "NO"
        gtcafechs(gtcii,1) = 0
    next

    gtcyy = 1


    for gtcii = gtcii to (gtcii + gtctopefecH )
        sInput1      = cdate(clng(gtcafech(gtcyy,1))  + gtcrm)
        gtcsdia1        = datepart("d",sInput1)
        gtcafechs(gtcii,0) = gtcsdia1
        gtcafechs(gtcii,1) = gtcafech(gtcyy,1)
        gtcyy = gtcyy + 1
    next


    for gtcii = gtcii to 42
        gtcafechs(gtcii,0) = "NO"
        gtcafechs(gtcii,1) = 0
    next


    CELDA = ""
    gtcyy = 0

    xcolor = "background-color:#f7aa22;"

    xTrColor = "#f0f0f0"

    for gtcii = 1 to 43

        if gtcyy = 7 then
           gtcyy = 1
           Fila = fila & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","25","","",Celda,"si")

           if xTrColor = "#f0f0f0" then
              xTrColor = "#F9F9F9" 
           ELSE
              xTrColor = "#f0f0f0" 
           END IF
           CELDA = ""
        else
           gtcyy = gtcyy + 1 
        end if

        if gtcii > 42 then exit for  

        if gtcafechs(gtcii,0) = "1" then xcolor = ""

        if gtcyy = 1 or gtcyy = 7 then
           sclase = "tt3 style='border-style:outset;border-width:2;'" 
        else
           sclase = "TT style='border-style:outset;border-width:2;" & xcolor & "'" 
        end if 

        if gtcafechs(gtcii,0) = "NO" then
           sinput = "&nbsp;"
           sclase = "CC " 
           Celda = Celda & ofv.GenCelda("","uT " & xtrcell,"CENTER","10%","","","",sinput,"si")
        else
           sinput1 = gtcafechs(gtcii,0)

           sinput1 = cstr(sinput1)

           if len(sinput1) = 1 then sinput1 = "0" & sinput1 

           sinput1 = " " & sinput1 & " " 

           sAccion = " "

           sf = "alert('"  & GTFechador(gtcafechs(gtcii,1)) & "');"

           t1 = ofv.GenerarInput("bvo1",sInput1,"BUTTON","","10",sclase,sf)
           t1 = ofv.GenCelda("","uT " & xtrcell,"center","10%","","","",t1,"si")
           Celda = Celda & t1

        end if


    next

    Celda = ofv.GenCelda("","ut colspan=7 ","","","","","","&nbsp;","si")

    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")




 shtml = shtml & "<br>" & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;border-color:#993333;' ","","50%","","0","","0","0","","",FILA,"si")

  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","","","","")
end if

 response.write shtml



%>