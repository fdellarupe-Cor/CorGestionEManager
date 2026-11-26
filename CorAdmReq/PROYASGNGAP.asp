<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0



  sHTML = ofv.MenuHeader("","")





  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn6)
  set cn2 = ofv.conectar(ofv.strconn0)
  set cn3 = ofv.conectar(ofv.strconn2)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar



if nFecini < clng(cdate("01-01-1900")) then
   nFecini = clng(cdate("01-01-1900"))
else
   nFecini = 0
end if

 XSET = request.querystring("XSET")
 XOPT = request.querystring("XOPT")
 sopcion = request.querystring("opcion")
 sRESP = request.querystring("RESP")


 IF XOPT = "" THEN XOPT = "XRESP"
 IF XSET = "" THEN XSET = "XPT"




    DIM ATAREA(200,40)
    DIM AORD(200,2)
    DIM AGAP(200,2)

    DIM WTFCH(6)

   ITART = 0
   ITATI = 0
   TTAR  = 0

   IGAP  = 0
   TGAP  = 0

   IETPT = 0
   IETPI = 0 
   TETP  = 0

   CDESP = 0
   CHASP = 0
   CDESR = 0
   CHASR = 0
   CMAX  = 0  


     sopcion = "VP"




SELECT CASE XOPT


 CASE "XGLB"

       XDEST="GLOBAL"

  Strsql = "Select * from clienteplan where id = 1"
  set rs = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
  if not rs.eof then
     NOMEMP = rs(1)
     LOGEMP = rs(5)
  end if
  call ofv.cerrarconsulta(rs)

       XDESTN=NOMEMP

      StrSql = "Select * From PROYTSKASGN where ESTADO < 2 "
      StrSql = StrSql & " AND ((VPLAN = PVPLAN ) "

      StrSql = StrSql & " ) "
      StrSql = StrSql & " order by GERENCIA, SECTOR, USRRESP, FINICIO, CODIGO, orden, TSKORD "


  CASE "XGRN"

       XDEST="GERENCIA"

     StrSql = "Select * From ECO where ECO = " & SRESP
     set rsX2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rsX2.eof then
        sniv = rsX2(1)
     ELSE
        sniv = "NO DETERMINADO"
     END IF
      call ofv.cerrarconsulta(rsX2)

       XDESTN=sniv

      StrSql = "Select * From PROYTSKASGN where ESTADO < 2 AND GERENCIA = " & sRESP
      StrSql = StrSql & " AND ((VPLAN = PVPLAN ) "
      StrSql = StrSql & " ) "
      StrSql = StrSql & " order by GERENCIA, SECTOR, USRRESP, FINICIO, CODIGO, orden, TSKORD "

  CASE "XSCT"

       XDEST="SECTOR"

     StrSql = "Select * From ECO where ECO = " & SRESP
     set rsX2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rsX2.eof then
        sniv = rsX2(1)
     ELSE
        sniv = "NO DETERMINADO"
     END IF
      call ofv.cerrarconsulta(rsX2)

       XDESTN=sniv

      StrSql = "Select * From PROYTSKASGN where ESTADO < 2 AND SECTOR = " & sRESP
      StrSql = StrSql & " AND ((VPLAN = PVPLAN ) "
      StrSql = StrSql & " ) "
      StrSql = StrSql & " order by GERENCIA, SECTOR, USRRESP, FINICIO, CODIGO, orden, TSKORD "

  CASE "XRESP"

       XDEST="RESPONSABLE"

      strsql = "Select * From USUARIOS where ID = '" & SRESP & "' "
      set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
      if not rsniv.eof then
         sniv = rsniv(1)
     ELSE
        sniv = "NO DETERMINADO"
     END IF
      call ofv.cerrarconsulta(rsniv)
       XDESTN=sniv

      StrSql = "Select * From PROYTSKASGN where ESTADO < 2 AND USRRESP = '" & sRESP  & "' "
      StrSql = StrSql & " AND ((VPLAN = PVPLAN ) "
      StrSql = StrSql & " ) "
      StrSql = StrSql & " order by GERENCIA, SECTOR, USRRESP, FINICIO, CODIGO, orden, TSKORD "

END SELECT




IF SESSION("PRCF") <> "S" THEN
   XXXX = OWKF.GETPRYCNFG()
END IF





    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"



    XFSTART = CLNG(DATE()) - clng(cdate("01-01-1900")) 
    XFCURR  = XFSTART
    XFHOY   = XFSTART
    XFREAL  = XFSTART

'''XXX = OWKF.ADMFERIADOS("C",XFSTART)



xfdes = CLNG(SESSION("FECDES"))
xfhas = CLNG(SESSION("FECHAS"))
xcdias = xfhas - xfdes + 1

xcdiast = " CORRIDOS" 

IF SESSION("FECOPC") <> "2" THEN

xcdias = 365

IF Session("DHC") = 1 THEN 

   xcdiast = " HABILES " 
   xcdiasC = xcdias
END IF

   xcdiast = xcdiast & xcdias & " (CORRIDOS " &  xcdiasC & ") "

ELSE







IF Session("DHC") = 1 THEN 

   xcdiast = " HABILES " 
   xcdiasC = xcdias

   xcdias = CLNG(xcdias * 0.733)

   IIX = 0

   iiX = OWKF.GETDHAB(xfHAS, iiX)

   XFHAS = XFHAS + IIX

   YWXAGRW = GETFECHA(xfdes, xcdias, 2, 1)

   IF WTFCH(3) <> xfHAS THEN

      XPOND = 1 

      IF WTFCH(3) > xfHAS THEN XPOND = -1 

      DO UNTIL WTFCH(3) = xfHAS

         xcdias = CLNG(xcdias + (1 * XPOND))

         YWXAGRW = GETFECHA(xfdes, xcdias, 2, 1)

      LOOP

   END IF


    '        WTFCH(0) = ical
    '        WTFCH(1) = ii3
    '        WTFCH(2) = FEDES
    '        WTFCH(3) = IHAB

   xcdiast = xcdiast & xcdias & " (CORRIDOS " &  xcdiasC & ") "


END IF

END IF

IF SESSION("FECOPC") = "2" THEN

    XPERIOD = "DESDE <FONT COLOR=#993333>" & UCASE(Fechador(xfdes)) & "</FONT> HASTA <FONT COLOR=#993333>" & UCASE(Fechador(xfHAS)) & "</FONT>"



ELSE

   XPERIOD = "INVENTARIO COMPLETO"

END IF



  ''    XFCURR = XFSTART

         dFecha = CDATE(CLNG(XFSTART) + clng(cdate("01-01-1900")))

      celdas = ofv.GenCelda("","cc colspan=5","","","","","",sImagen & "&nbsp;<BIG><B>GAP DE ASIGNACION POR RESPONSABLE " & XDESTN & "&nbsp;&nbsp;&nbsp;&nbsp; AL " & ofv.fechador & "</B></BIG>","si")
      filas  = ofv.GenRow("","","","","10","","",celdas,"si")


      celdas = ofv.GenCelda("","AA COLSPAN=5","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



  celdas = ofv.GenCelda("","vt ","","10%","","","","PERIODO:","si")
  celdas = celdas & ofv.GenCelda("","vt COLSPAN=5  ","","","","","",XPERIOD,"si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","vt ","","","","","","DIAS:","si")
  celdas = celdas & ofv.GenCelda("","vt COLSPAN=5  ","","","","","",xcdiast,"si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","AA COLSPAN=5","CENTER","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      FilaS1 = ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")




      filas = ""

    sImagen0 = "<FONT FACE=WINGDINGS SIZE=2>F</FONT>"
    sImagen1 = "<FONT FACE=WINGDINGS SIZE=2>ñ</FONT>"
    sImagen1N = "<FONT FACE=WINGDINGS COLOR=#AA0000 SIZE=2>l</FONT>"
    sImagen2 = "<FONT FACE=VERDANA SIZE=2>+</FONT>"
    sImagen3 = "<FONT FACE=WINGDINGS SIZE=2>2</FONT>"
    sImagen4s = "<FONT FACE=WEBDINGS COLOR=#993333 SIZE=2>&#0128;</FONT>"
    sImagen4n = "<FONT FACE=WEBDINGS COLOR=#FFFFFF SIZE=2>x</FONT>"
    sImagen6S = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=2>&#0252;</FONT></B>"
    sImagen6N = "<B><FONT FACE=WINGDINGS COLOR=#AA0000 SIZE=2>&#0233;</FONT></B>"
    sImagen6A = "<B><FONT FACE=WINGDINGS COLOR=#F7AA22 SIZE=3>&#0190;</FONT></B>"
    sImagen6W = "<B><FONT FACE=WINGDINGS COLOR=#CCCCCC SIZE=2>s</FONT></B>"
    sImagen6T = "<B><FONT FACE=WEBDINGS COLOR=#336699 SIZE=3>q</FONT></B>"
    sImagenE0 = "<B><FONT FACE=WINGDINGS COLOR=#6699CC SIZE=4>6</FONT></B>"
    sImagenE1 = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=4>&#0232;</FONT></B>"
    sImagenE2 = "<B><FONT FACE=WINGDINGS COLOR=#000080 SIZE=4>&#0251;</FONT></B>"
    sImagenE4 = "<B><FONT FACE=WINGDINGS COLOR=#993333 SIZE=4>&#0232;</FONT></B>"



''''''''''''''''
  tskant = "no"

  usrant = "no"

  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not  rs1.eof then
     do while not rs1.eof



  XVA = "S"


  IF SESSION("SELOPC") = "1" THEN
      StrSql = "Select * From PROYTIPOTSK where ID = 0" & RS1(17)
      set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if not  rsX1.eof then
      
         IF RSX1(2) = 0 THEN
            XVA = "N"

         ELSE

           StrSql = "Select TIPO From PROYECTO where CODIGO = '" & RS1(5) & "' "
           set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if not  rsX2.eof then

              XTIPO = RSX2(0)

           ELSE
              XTIPO = 0

           END IF

           StrSql = "Select * From TIPOS where ID = 0" & XTIPO
           set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if not  rsX2.eof then

              IF RSX2(4) = 0 THEN
                 XVA = "N"

              END IF

           END IF

         END IF

      END IF


  END IF 
  


   IF XVA = "S" THEN

       xnompry = rs1(38)
       xusrrsp = rs1(39)
       if rs1(37) > 0 then
          XFREAL = rs1(37)
       else
          XFREAL = rs1(36)
       end if
       XVPLAN = RS1(25)
       XFFCH  = RS1(35)

       XUSR   = RS1(13)
       XSCT   = RS1(26)
       XGRN   = RS1(27)





     SXFSTARTD = OWKF.GETFENDV(rs1(20), rs1(5))




    XIIX = OWKF.TIX





    XFREAL = OWKF.ATABTSK(cint(XIIX),1)
    TSKFS  = OWKF.ATABTSK(cint(XIIX),2)
    TSKFE  = OWKF.ATABTSK(cint(XIIX),3)
    TSKFSR = OWKF.ATABTSK(cint(XIIX),4)
    TSKFER = OWKF.ATABTSK(cint(XIIX),5)
    TSKST  = OWKF.ATABTSK(cint(XIIX),6)
    FSPLAN = OWKF.ATABTSK(cint(XIIX),7)
    FEPLAN = OWKF.ATABTSK(cint(XIIX),8)
    FSREAL = OWKF.ATABTSK(cint(XIIX),9)
    FEREAL = OWKF.ATABTSK(cint(XIIX),10)
    XFFDU  = OWKF.ATABTSK(cint(XIIX),11)
    XFFTI  = OWKF.ATABTSK(cint(XIIX),12)
    PADV   = OWKF.ATABTSK(cint(XIIX),13)
    XAGRD  = OWKF.ATABTSK(cint(XIIX),14)
    XAGRR  = OWKF.ATABTSK(cint(XIIX),15)
    XMODA  = OWKF.ATABTSK(cint(XIIX),16)
    XTDUR  = OWKF.ATABTSK(cint(XIIX),17)
    XUDUR  = OWKF.ATABTSK(cint(XIIX),18)
    XTDUP  = OWKF.ATABTSK(cint(XIIX),19)
    XUDUP  = OWKF.ATABTSK(cint(XIIX),20)
    XAGRP  = OWKF.ATABTSK(cint(XIIX),21)





  IF SESSION("FECOPC") = "2" THEN
     XFECDES = FSREAL

     XFECHAS = FEREAL

      
     IF CLNG(SESSION("FECDES")) > CLNG(XFECHAS) OR CLNG(SESSION("FECHAS")) < CLNG(XFECDES) THEN
        XVA = "N"
     END IF

   


  END IF 







  IF XVA = "S" THEN

     if   tskant <> rs1(2) then
          tskant = rs1(2) 

       if usrant <> rs1(13)  then 
 


          if usrant = "no" then
             usrant = rs1(13)
          ELSE
'''''''''''''''
  for itart=1 to ttar

 ''   ATAREA (ITART,11) = FSREAL
 ''   ATAREA (ITART,12) = FEREAL


      FOR WTART=1 TO TTAR

        IF ATAREA (ITART,0) <> ATAREA (WTART,0) THEN

          IF  ATAREA (ITART,11) <= ATAREA (WTART,12) AND ATAREA (ITART,12) >= ATAREA (WTART,11)  THEN

              IF (ATAREA (ITART,11) = ATAREA (WTART,12) OR ATAREA (ITART,12) = ATAREA (WTART,11))  THEN

'''AND ATAREA (ITART,29) = ATAREA (WTART,29)


              ELSE
'''                  ATAREA (ITART,30) = ATAREA (ITART,30) + 1
'''                  ATAREA (WTART,30) = ATAREA (WTART,30) + 1 
                  ATAREA (ITART,30) =  1
                  ATAREA (WTART,30) =  1 
              END IF 

          END IF

        END IF 


      NEXT
      


  NEXT


  XFINI = 0
  XFFIN = 0


 
  XHAYG = "N"

   XTAR1 = 0
   XTAR2 = 0

   IGAP  = 0
   TGAP  = 0

  for WTARt=1 to ttar

      ITART = AORD(WTART ,0)

'''ATAREA (ITART,18)

      IF XFINI = 0 THEN 
         XFINI = ATAREA (ITART,18)
         XFFIN = ATAREA (ITART,19)
         XTAR1 = WTARt
      END IF

      IF ATAREA (ITART,18) > XFFIN THEN
         XTAR2 = WTARt
         XHAYG = "S"
         IGAP = IGAP + 1
         TGAP = IGAP
         AGAP(IGAP ,1) = XTAR1
         AGAP(IGAP ,2) = XTAR2
         XFINI = ATAREA (ITART,18)
         XFFIN = ATAREA (ITART,19)
         XTAR1 = WTARt

      ELSEIF NOT ATAREA (ITART,19) < XFFIN THEN

             IF XTAR1 <> WTART THEN
                XFINI = ATAREA (ITART,18)
                XFFIN = ATAREA (ITART,19)
                XTAR1 = WTARt
             END IF

      END IF
      


  NEXT


  XRESPO = "&nbsp;&nbsp;RESPONSABLE: <b>" & ATAREA (1,9) & "</B>"
  sInputn0 = ""

 ''   ATAREA (ITART,31) = XGRN
 ''   ATAREA (ITART,32) = XSCT

   if XOPT = "XGRN" or XOPT = "XGLB" then

     StrSql = "Select * From ECO where ECO = " & ATAREA (1,32)
     set rsX2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rsX2.eof then
        sniv = rsX2(1)
     ELSE
        sniv = "NO DETERMINADO"
     END IF
      call ofv.cerrarconsulta(rsX2) 

     sInputn0 = sInputn0 & "<br>&nbsp;&nbsp;SECTOR:&nbsp;&nbsp;&nbsp;&nbsp;<b> " & sniv & "</b>"

   end if

   if XOPT = "XGLB" then

     StrSql = "Select * From ECO where ECO = " & ATAREA (1,31)
     set rsX2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rsX2.eof then
        sniv = rsX2(1)
     ELSE
        sniv = "NO DETERMINADO"
     END IF
      call ofv.cerrarconsulta(rsX2) 

     sInputn0 = sInputn0 & "<br>&nbsp;&nbsp;GERENCIA:&nbsp;&nbsp;<b> " & sniv & "</b>"

   end if

  XRESPO =   XRESPO & sInputn0


      celdas = ofv.GenCelda("","AA COLSPAN=4 STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#F7AA22;' BGCOLOR=#F7EEDD ","","","","","",XRESPO,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    celdas =  ofv.GenCelda("","VT  STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#F7AA22;' ","","10%","","","","&nbsp;","si")
    celdas = celdas & ofv.GenCelda("","VT   STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#F7AA22;'  ","CENTER","40%","","","", "ULTIMA TAREA","si")
    celdas = celdas & ofv.GenCelda("","VT  STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#F7AA22;'  ","CENTER","40%","","","", "PRIMERA TAREA","si")
    celdas = celdas & ofv.GenCelda("","VT  STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#F7AA22;' ","","10%","","","","DIFERENCIA","si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


  IF XHAYG = "S" THEN

  FOR IGAP=1 TO TGAP

      celdas = ofv.GenCelda("","VT ROWSPAN=3 STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#336699;'  ","CENTER","10%","","","","GAP" & IGAP,"si")




      WTART = AGAP(IGAP ,1)
      WTART2 = AGAP(IGAP ,2)

      ITART  = AORD(WTART ,0)
      ITART2 = AORD(WTART2 ,0)






      WSOLAPA = "&nbsp;*"
      WSFONT = "<FONT COLOR=#AA0000 SIZE=1 ><B>"



   sInputnP = ATAREA (ITART,3) & ".&nbsp;" & ATAREA (ITART,2)
   sInputnT = ATAREA (ITART,29)

   sInputnP2 = ATAREA (ITART2,3) & ".&nbsp;" & ATAREA (ITART2,2)
   sInputnT2 = ATAREA (ITART2,29)


    celdas = celdas & ofv.GenCelda("","VT    ","","40%","","","", "TAREA: " & sInputnP,"si")
    celdas = celdas & ofv.GenCelda("","VT    ","","40%","","","", "TAREA: " & sInputnP2,"si")

''''''' CALCULO DE DIFERENCIAS

xfdes = ATAREA (ITART,19)
xfhas = ATAREA (ITART2,18)

xcdias = xfhas - xfdes + 1




IF Session("DHC") = 1 THEN 

   xcdiasC = xcdias

   xcdias = CLNG(xcdias * 0.733)

   IIX = 0

   iiX = OWKF.GETDHAB(xfHAS, iiX)

   XFHAS = XFHAS + IIX

   YWXAGRW = GETFECHA(xfdes, xcdias, 2, 1)

   IF WTFCH(3) <> xfhas THEN

      XPOND = 1 

      IF WTFCH(3) > xfhas THEN XPOND = -1 

      DO UNTIL WTFCH(3) = xfhas

         xcdias = CLNG(xcdias + (1 * XPOND))

         YWXAGRW = GETFECHA(xfdes, xcdias, 2, 1)

      LOOP

   END IF


    '        WTFCH(0) = ical
    '        WTFCH(1) = ii3
    '        WTFCH(2) = FEDES
    '        WTFCH(3) = IHAB

   xcdiast = xcdias * SESSION("DPP")


END IF





    celdas = celdas & ofv.GenCelda("","VT  ROWSPAN=3 STYLE='BORDER-STYLE:SOLID;BORDER-WIDTH:1;BORDER-COLOR:#336699;' ","","10%","","","","&nbsp;<FONT COLOR=#993333>" & xcdiast & "&nbsp;Hs.</FONT>","si")

'''''''''

      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")




    celdas = ofv.GenCelda("","VT    ","","","","","","PROY: " & sInputnT,"si")
    celdas = celdas & ofv.GenCelda("","VT    ","","","","","","PROY: " & sInputnT2,"si")


      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")




       dFecha = cdate(ATAREA (ITART,18) + clng(cdate("01-01-1900")))
       SINPUTn4 = dFecha


       dFecha = cdate(ATAREA (ITART,19) + clng(cdate("01-01-1900")))
       SINPUTn5 = dFecha


       dFecha = cdate(ATAREA (ITART2,18) + clng(cdate("01-01-1900")))
       SINPUTn42 = dFecha


       dFecha = cdate(ATAREA (ITART2,19) + clng(cdate("01-01-1900")))
       SINPUTn52 = dFecha


      celdas = ofv.GenCelda("","VT    ","","","","","","FIN: " & SINPUTn5,"si")
      celdas = celdas & ofv.GenCelda("","VT    ","","","","","","INICIO: " & SINPUTn42,"si")


      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



''      celdas = celdas & ofv.GenCelda("","ut  STYLE='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#f7aa22;'","","75%","","","",sInputn0,"si")


      celdas = ofv.GenCelda("","AA COLSPAN=4 STYLE='border-TOP-style:RIDGE;border-TOP-width:2;border-TOP-color:#f7aa22;'  ","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")




  next

  ELSE

  XRESPO =   "<BR>NO REGISTRA GAPS DE ASIGNACION"


      celdas = ofv.GenCelda("","VT COLSPAN=4   ","CENTER","","","","",XRESPO,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


  END IF


      celdas = ofv.GenCelda("","AA COLSPAN=4   ","CENTER","","","","","&nbsp;<BR>&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


'''''''''''''''




          END IF


          USRANT = RS1(13)
          ITART = 0
          ITATI = 0
          TTAR  = 0 



       END IF

    IF MFMAX < FEREAL THEN MFMAX = FEREAL



       mXTACUM  = FSPLAN - MFMIN + 1 + XAGRP
       mXTACUMR = FSPLAN - MFMIN + 1 + XAGRR

   ''    response.write "tsk: " & rs1(3) & " sp " & fsplan & " ep " & feplan & " sr " & fsreal & " er " & fereal & " ap " & mXTACUM  & " ar " & mXTACUMr  & "<br>" 

       XTACUM   = FSPLAN - MFMIN + 1
       XTACUMR  = FSreal - MFMIN + 1


       CDESP    = XTACUM
       XTACUM   = XTACUM + XAGRP
       CHASP    = XTACUM       

       CDESR    = XTACUMR
       XTACUMR  = XTACUMR + XAGRR
       CHASR    = XTACUMR



    if RS1(13) = "NO" then
       sniv = "SIN AFECTAR"
    ELSE
      strsql = "Select * From USUARIOS where ID = '" & rs1(13) & "' "
      set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
      if not rsniv.eof then sniv = rsniv(1)
      call ofv.cerrarconsulta(rsniv)
    END IF



    ITART             = ITART + 1
    ATAREA (ITART,0)  = RS1(2)
    ATAREA (ITART,1)  = IETPT
    ATAREA (ITART,2)  = RS1(8)
    ATAREA (ITART,3)  = rs1(0) & "." & rs1(1)
    ATAREA (ITART,4)  = CDESP
    ATAREA (ITART,5)  = CDESR
    ATAREA (ITART,6)  = CHASP
    ATAREA (ITART,7)  = CHASR 
    ATAREA (ITART,8)  = TSKst
    ATAREA (ITART,9)  = sniv
    ATAREA (ITART,10) = PADV
    ATAREA (ITART,11) = FSREAL
    ATAREA (ITART,12) = FEREAL
    ATAREA (ITART,13) = FSPLAN
    ATAREA (ITART,14) = FEPLAN
    ATAREA (ITART,15) = XAGRD
    ATAREA (ITART,16) = TSKFS
    ATAREA (ITART,17) = TSKFE
    ATAREA (ITART,18) = TSKFSR
    ATAREA (ITART,19) = TSKFER
    ATAREA (ITART,20) = XFFDU
    ATAREA (ITART,21) = XFFTI
    ATAREA (ITART,22) = XAGRR
    ATAREA (ITART,23) = XMODA
    ATAREA (ITART,24) = XTDUR
    ATAREA (ITART,25) = XUDUR
    ATAREA (ITART,26) = XTDUP
    ATAREA (ITART,27) = XUDUP
    ATAREA (ITART,28) = XAGRP
    ATAREA (ITART,29) = xnompry
    ATAREA (ITART,30) = 0
    ATAREA (ITART,31) = XGRN
    ATAREA (ITART,32) = XSCT
    ATAREA (ITART,33) = XUSR




    TTAR              = ITART 

    IF ITART = 1 THEN
       AORD(ITART ,0) = ITART 
       AORD(ITART ,1) = FSREAL
       AORD(ITART ,2) = FEREAL
       
    ELSE

       FOR WTART=1 TO TTAR - 1
           IF FSREAL < AORD(WTART ,1) THEN
              EXIT FOR
           ELSEIF FSREAL = AORD(WTART ,1) THEN
                  IF FEREAL < AORD(WTART ,2) THEN
                     EXIT FOR
                  END IF
           END IF 
       NEXT
         
       IF WTART < TTAR THEN
          II = TTAR-1  
          FOR IIB=WTART TO TTAR-1
              AORD(II + 1 ,0) = AORD(II ,0)
              AORD(II + 1 ,1) = AORD(II ,1)
              AORD(II + 1 ,2) = AORD(II ,2)
              II = II - 1
          NEXT
       END IF
       
       AORD(WTART ,0) = ITART 
       AORD(WTART ,1) = FSREAL
       AORD(WTART ,2) = FEREAL

    END IF





    end if
   END IF
  END IF

    rs1.movenext
    loop
    call ofv.cerrarconsulta(rs1)
  end if





      celdas = ofv.GenCelda("","","CENTER","","2%","","","&nbsp;","si")
      celdas = celdas & ofv.GenCelda("","","CENTER","","5%","","","&nbsp;","si")
      celdas = celdas & ofv.GenCelda("","","CENTER","","73%","","","&nbsp;","si")
      celdas = celdas & ofv.GenCelda("","","CENTER","","5%","","","&nbsp;","si")
      celdas = celdas & ofv.GenCelda("","","CENTER","","5%","","","&nbsp;","si")
      celdas = celdas & ofv.GenCelda("","","CENTER","","5%","","","&nbsp;","si")
      celdas = celdas & ofv.GenCelda("","","CENTER","","5%","","","&nbsp;","si")
      celdas = celdas & ofv.GenCelda("","","CENTER","","5%","","","&nbsp;","si")
      celdas = celdas & ofv.GenCelda("","","CENTER","","5%","","","&nbsp;","si")
      celdas = celdas & ofv.GenCelda("","","CENTER","","1%","","","&nbsp;","si")
      celdas = celdas & ofv.GenCelda("","","CENTER","","1%","","","&nbsp;","si")

    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")


    FilaS2 = ofv.GenTabla("","aa style='border-style:ridge;border-width:2;border-color:#f7aa22;' ","","100%","","0","","0","0","","",filas,"si")


    Celda = ofv.GenCelda("","AAA COLSPAN=2","center","","","","",FilaS1 & FilaS2,"si")
    FilaS = ofv.GenRow("","","","","","","",Celda,"si")

     cnbot = "<input type=submit class=bt value='GANTT' >"
     Cons = ofv.gencelda("","","","","1%","","",cnbot,"si")
     sAccion = "PROYASGNgraf.asp?RESP=" & SRESP
     Cons = ofv.genform("cons","",sAccion,"_blank","post","",Cons,"si")

     if xopt = "XRESP" THEN
        celdas = Cons
     else
        celdas = ""
     end if

      celdas = celdas & ofv.GenCelda("","","CENTER","","80%%","","","&nbsp;","si")

    FilaS = FilaS & ofv.GenRow("","","","","","","",CeldaS,"si")

    sHTML = sHTML & ofv.GenTabla("","aa  ","","90%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"





'''''''''''''''












 







  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML

Function Fechador(wkfecha)


    
      fec = cdate(wkfecha + clng(cdate("01-01-1900")))
      ndia = ofv.cdia(Weekday(fec))
      nMes = ofv.cmes(Month(fec))
    
      Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)


End Function

Function GETFECHA(FEDES, FDIAS, FTIPO, FCH)

 IFEC = FEDES


 Select Case FTIPO

           Case 4
                  ical = CInt(FDIAS / CLng(Session("DPP")) / 60)


           Case 1
               ical = CInt(FDIAS / CLng(Session("DPP")))

        Case 2
               ical = FDIAS

        Case 3
               ical = FDIAS * CLng(Session("SPP"))

        Case Else
               ical = FDIAS

 End Select


 Select Case FCH

           Case 0


             ICORR = IFEC + ical

             WTFCH(0) = ical
             WTFCH(1) = ical
             WTFCH(2) = FEDES
             WTFCH(3) = ICORR


        Case 1

             IHAB = 0

             IFECD = CDate(IFEC + CLng(CDate("01-01-1900")))
             papa = DatePart("W", IFECD, 2)



             ii = 0
             ii3 = 0


             ii3 = OWKF.GETDHAB(IFEC, ii3)




             Do While ii < ical
                ii3 = ii3 + 1
                ii3 = OWKF.GETDHAB(IFEC, ii3)

                ii = ii + 1

             Loop


            IHAB = IFEC + ii3
            WTFCH(0) = ical
            WTFCH(1) = ii3
            WTFCH(2) = FEDES
            WTFCH(3) = IHAB

 End Select


End Function

%>