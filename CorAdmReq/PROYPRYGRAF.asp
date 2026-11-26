<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


''  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))
  sHTML = ofv.MenuHeader("","")

  sHTML = sHTML  & "<script language=javascript> "
  sHTML = sHTML  & "function AbrirCons(Datos) { "
  sHTML = sHTML  & " myWin= open(Datos," & chr(34) & "Cons2" & chr(34) & "," & chr(34) & "menubar=no,toolbar=no,height=450,width=600" & chr(34) & " ); "
  sHTML = sHTML  & " } "
  sHTML = sHTML  & "</script>"

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

'  IF session("ModMain") <> "WF" THEN
'     bAgregar = FALSE
'     bModificar = FALSE
'     bSuprimir = FALSE
'  END IF

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

     sopcion = "VP"


IF SESSION("PRCF") <> "S" THEN
   XXXX = OWKF.GETPRYCNFG()
END IF


 Whtml = ""

 Whtml = Whtml & "<table border=0 cellspacing=0 cellpadding=0 bgcolor=#eeeeee width=100% height=100% style='border-style:ridge;border-width:3;' >"
 Whtml = Whtml & "<tr>"
 Whtml = Whtml & "<td align=center style='border-style:ridge,border-width:2; filter:shadow(color=#777777,direction=135);' >"
 Whtml = Whtml & "<IMG border=0 src='../coress20/imagenes/corlogo.gif' style='filter: chroma(color=#ffffff);' >"
 Whtml = Whtml & "<FONT color=#ffffff face=verdana size=2>"
 Whtml = Whtml & "<B>&nbsp;Cor&nbsp;Consulting</B>"
 Whtml = Whtml & "</FONT>"
 Whtml = Whtml & "</td></tr></table>"


    sInput = "<B><h3>&nbsp;COR&nbsp;<i><font color=#f7aa22>On</font></i>Time&nbsp;-&nbsp;Consulta&nbsp;Gantt</h3></B>"
    celdas = ofv.GenCelda("","vt  ","","80%","","","",sInput,"si")

    sInput = Whtml
    celdas = celdas & ofv.GenCelda("","","","200","","","",sInput,"si")

    filas = ofv.GenRow("","","","","","","",celdas,"si")

  sHTML = sHTML  & ofv.GenTabla("","AAaa  ","","98%","","0","","-2","-1","","",filas,"si")


    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"



    XFSTART = CLNG(DATE()) - clng(cdate("01-01-1900")) 
    XFCURR  = XFSTART
    XFHOY   = XFSTART
    XFREAL  = XFSTART
    XTACUM  = 1
    XTACUMR = 1
    XTFTSK  = 0

    XMFMIN = 0
    XMFMAX = 0

    MFMAX = 0

    DIM ATAREA(200,40)
    DIM AORD(200,2)

   ITART = 0
   ITATI = 0
   TTAR  = 0

   IETPT = 0
   IETPI = 0 
   TETP  = 0

   CDESP = 0
   CHASP = 0
   CDESR = 0
   CHASR = 0
   CMAX  = 0  

    sImageS = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=3>4</FONT></B>"
   


XXX = OWKF.ADMFERIADOS("C",XFSTART)

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

      StrSql = "Select DISTINCT CODIGO From PROYTSKASGN where  "
      StrSql = StrSql & "  ((VPLAN = PVPLAN and CODIGO <> '" & SESSION("NCLIENTE") & "') "
      StrSql = StrSql & " OR (VPLAN = '" & SESSION("VPLAN") & "' AND CODIGO = '" & SESSION("NCLIENTE") & "')) "
''      StrSql = StrSql & " order by GERENCIA, SECTOR, USRRESP, FINICIO, CODIGO, orden, TSKORD "


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

      StrSql = "Select DISTINCT CODIGO From PROYTSKASGN where  GERENCIA = " & sRESP
      StrSql = StrSql & " AND ((VPLAN = PVPLAN and CODIGO <> '" & SESSION("NCLIENTE") & "') "
      StrSql = StrSql & " OR (VPLAN = '" & SESSION("VPLAN") & "' AND CODIGO = '" & SESSION("NCLIENTE") & "')) "
''      StrSql = StrSql & " order by GERENCIA, SECTOR, USRRESP, FINICIO, CODIGO, orden, TSKORD "

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

      StrSql = "Select DISTINCT CODIGO From PROYTSKASGN where SECTOR = " & sRESP
      StrSql = StrSql & " AND ((VPLAN = PVPLAN and CODIGO <> '" & SESSION("NCLIENTE") & "') "
      StrSql = StrSql & " OR (VPLAN = '" & SESSION("VPLAN") & "' AND CODIGO = '" & SESSION("NCLIENTE") & "')) "
''      StrSql = StrSql & " order by GERENCIA, SECTOR, USRRESP, FINICIO, CODIGO, orden, TSKORD "

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

      StrSql = "Select DISTINCT CODIGO From PROYTSKASGN where USRRESP = '" & sRESP  & "' "
      StrSql = StrSql & " AND ((VPLAN = PVPLAN and CODIGO <> '" & SESSION("NCLIENTE") & "') "
      StrSql = StrSql & " OR (VPLAN = '" & SESSION("VPLAN") & "' AND CODIGO = '" & SESSION("NCLIENTE") & "')) "
''      StrSql = StrSql & " order by GERENCIA, SECTOR, USRRESP, FINICIO, CODIGO, orden, TSKORD "

END SELECT






  ''    XFCURR = XFSTART

         dFecha = CDATE(CLNG(XFSTART) + clng(cdate("01-01-1900")))

      celdas = ofv.GenCelda("","vt colspan=4","","","","","",sImagen & "&nbsp;<BIG><B>PROYECTOS DE " & XDESTN & "&nbsp;&nbsp;&nbsp;&nbsp; AL " & ofv.fechador & "</B></BIG>","si")
      filas  = ofv.GenRow("","","","","10","","",celdas,"si")


      celdas = ofv.GenCelda("","AA COLSPAN=4","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      FilaS1 = ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")




''  StrSql = "Select * From PROYTSKASGN where ESTADO < 2 AND USRRESP = '" & sRESP  & "'  AND ((VPLAN = PVPLAN and CODIGO <> '" & SESSION("NCLIENTE") & "') OR (VPLAN = '" & SESSION("VPLAN") & "' AND CODIGO = '" & SESSION("NCLIENTE") & "'))  order by GERENCIA, SECTOR, USRRESP, FINICIO, CODIGO, orden, TSKORD "
''  response.write strsql
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not  rs1.eof then
     do while not rs1.eof


  XVA = "S"


  IF SESSION("SELOPC") = "1" THEN
           StrSql = "Select TIPO From PROYECTO where CODIGO = '" & RS1(0) & "' "
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
  


   IF XVA = "S" THEN


    StrSql = "Select * From PROYecto where codigo = '" & rs1(0) & "'  "
 '' response.write strsql
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT rs2.eof then
       xnompry = rs2(2)
       xusrrsp = rs2(14)
       if rs2(4) > 0 then
          XFREAL = rs2(4)
       else
          XFREAL = rs2(3)
       end if
       XPSTS  = RS2(11)
       XVPLAN = RS2(17)
       XFFCH  = RS2(18)

    end if
    call ofv.cerrarconsulta(rs2)




 





wtsktt = 0
wtskft = 0
wtskct = 0
wtskci = 0
wtskd  = 0


       StrSql = "Select * From PROYETPTSK where codigo = '" & rs1(0) & "' and VPLAN = " & XVPLAN & " ORDER BY ORDEN,TSKORD,FECHASPLAN DESC "
 '' response.write strsql
       set rs5 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       IF NOT rs5.EOF THEN
          YFRSTTSK =  RS5(20)

          do until rs5.eof

             USRTMP = RS5(15) & ""
             USRTMPU = RS5(16)
             IF USRTMP = "" THEN
                USRTMP = 1
                USRTMPU = 1
             END IF
             USRTMP = CINT(USRTMP)
             IF USRTMP = 0 THEN
                USRTMP = 1 
                USRTMPU = 1
             END IF

             SELECT CASE USRTMPU
                    CASE 4
                         xagrd = USRTMP / 60

                    CASE 1
                         xagrd = USRTMP

                    CASE 2

                         xagrd = (USRTMP * CLNG(SESSION("DPP")))
 
                    CASE ELSE

                         xagrd = (USRTMP * CLNG(SESSION("DPP")) * CLNG(SESSION("SPP")) )

              END select  
              xagrc = 0       

              if rs5(4) > 0 then
                 PADV = 0 
                 StrSql = "Select * From PROYTAREASADV where codigo = '" & rs5(5) & "' AND TSK = '" & RS5(20) & "' ORDER BY id desc "
                 set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
                 if NOT rsX.eof then
                    PADV = RSX(5)
                    xagrc = CINT((XAGRD * PADV / 100)) 
                    xagrr = XAGRD - CINT((XAGRD * PADV / 100)) 
                 END IF
               end if
               xagrr = xagrd - xagrc             

               wtsktt = wtsktt + xagrd
               wtskct = wtskct + xagrc

              if rs5(4) = 1 then
                 SXFSTARTD = OWKF.GETFENDV(RS5(20), rs1(0))

                 XIIX = OWKF.TIX

                 YFSPLAN = OWKF.ATABTSK(cint(XIIX),7)
                 YFEPLAN = OWKF.ATABTSK(cint(XIIX),8)
                 YFSREAL = OWKF.ATABTSK(cint(XIIX),9)
                 YFEREAL = OWKF.ATABTSK(cint(XIIX),10)

                 IF PADV > 0 THEN

                    IF XFHOY < YFEPLAN THEN
                       wtskd  = 0
                    ELSEIF XFHOY = YFEPLAN AND PADV > 99  THEN
                       wtskd  = 0
                    ELSE
                       wtskd  = 1
                    END IF

                 else

                    IF XFHOY > YFSREAL THEN
                       wtskd  = 2
                    END IF

                    IF XFHOY < YFEPLAN THEN

                    ELSE
                       wtskd  = 1
                    END IF

                 END IF


              end if 

          rs5.movenext
          loop
       END IF  
       call ofv.cerrarconsulta(rs5)



               wtskci = round((wtskct / wtsktt * 100),2)




 '' BUSCO LOS DATOS DE LA PRIMER TAREA


     SXFSTARTD = OWKF.GETFENDV(YFRSTTSK, rs1(0))



    XIIX = OWKF.TIX





    YXFREAL = OWKF.ATABTSK(cint(XIIX),1)
    YTSKFS  = OWKF.ATABTSK(cint(XIIX),2)
    YTSKFE  = OWKF.ATABTSK(cint(XIIX),3)
    YTSKFSR = OWKF.ATABTSK(cint(XIIX),4)
    YTSKFER = OWKF.ATABTSK(cint(XIIX),5)
    YTSKST  = OWKF.ATABTSK(cint(XIIX),6)
    YFSPLAN = OWKF.ATABTSK(cint(XIIX),7)
    YFEPLAN = OWKF.ATABTSK(cint(XIIX),8)
    YFSREAL = OWKF.ATABTSK(cint(XIIX),9)
    YFEREAL = OWKF.ATABTSK(cint(XIIX),10)
    YXFFDU  = OWKF.ATABTSK(cint(XIIX),11)
    YXFFTI  = OWKF.ATABTSK(cint(XIIX),12)
    YPADV   = OWKF.ATABTSK(cint(XIIX),13)
    YXAGRD  = OWKF.ATABTSK(cint(XIIX),14)
    YXAGRR  = OWKF.ATABTSK(cint(XIIX),15)
    YXMODA  = OWKF.ATABTSK(cint(XIIX),16)
    YXTDUR  = OWKF.ATABTSK(cint(XIIX),17)
    YXUDUR  = OWKF.ATABTSK(cint(XIIX),18)
    YXTDUP  = OWKF.ATABTSK(cint(XIIX),19)
    YXUDUP  = OWKF.ATABTSK(cint(XIIX),20)
    YXAGRP  = OWKF.ATABTSK(cint(XIIX),21)



 '' BUSCO LOS DATOS DE LA ULTIMA TAREA

     SXFSTARTD = OWKF.GETFENDV("999999",rs1(0))



    XIIX = OWKF.TIX



    ZXFREAL = OWKF.ATABTSK(cint(XIIX),1)
    ZTSKFS  = OWKF.ATABTSK(cint(XIIX),2)
    ZTSKFE  = OWKF.ATABTSK(cint(XIIX),3)
    ZTSKFSR = OWKF.ATABTSK(cint(XIIX),4)
    ZTSKFER = OWKF.ATABTSK(cint(XIIX),5)
    ZTSKST  = OWKF.ATABTSK(cint(XIIX),6)
    ZFSPLAN = OWKF.ATABTSK(cint(XIIX),7)
    ZFEPLAN = OWKF.ATABTSK(cint(XIIX),8)
    ZFSREAL = OWKF.ATABTSK(cint(XIIX),9)
    ZFEREAL = OWKF.ATABTSK(cint(XIIX),10)
    ZXFFDU  = OWKF.ATABTSK(cint(XIIX),11)
    ZXFFTI  = OWKF.ATABTSK(cint(XIIX),12)
    ZPADV   = OWKF.ATABTSK(cint(XIIX),13)
    ZXAGRD  = OWKF.ATABTSK(cint(XIIX),14)
    ZXAGRR  = OWKF.ATABTSK(cint(XIIX),15)
    ZXMODA  = OWKF.ATABTSK(cint(XIIX),16)
    ZXTDUR  = OWKF.ATABTSK(cint(XIIX),17)
    ZXUDUR  = OWKF.ATABTSK(cint(XIIX),18)
    ZXTDUP  = OWKF.ATABTSK(cint(XIIX),19)
    ZXUDUP  = OWKF.ATABTSK(cint(XIIX),20)
    ZXAGRP  = OWKF.ATABTSK(cint(XIIX),21)




  IF SESSION("FECOPC") = "2" THEN
     XFECDES = YFSREAL

     XFECHAS = ZFEREAL

      
     IF CLNG(SESSION("FECDES")) > CLNG(XFECHAS) OR CLNG(SESSION("FECHAS")) < CLNG(XFECDES) THEN
        XVA = "N"
     END IF

   


  END IF 

  IF XVA = "S" THEN







    IF MFMAX < ZFEPLAN THEN MFMAX = ZFEPLAN
    IF MFMAX < ZFEREAL THEN MFMAX = ZFEREAL

    IF MFMAX < SXFSTARTD THEN MFMAX = SXFSTARTD


    IF MFMIN = 0 THEN MFMIN = YFSPLAN

    IF MFMIN > YFSPLAN THEN MFMIN = YFSPLAN
    IF MFMIN > YFSREAL THEN MFMIN = YFSREAL







    TSKFS  = YTSKFS
    TSKFE  = ZTSKFE
    TSKFSR = YTSKFSR
    TSKFER = ZTSKFER

    FSPLAN = YFSPLAN
    FEPLAN = ZFEPLAN
    FSREAL = YFSREAL
    FEREAL = ZFEREAL


    IF XPSTS  = 1 THEN
       TSKST  = 0
    ELSEIF XPSTS  = 13 THEN
       TSKST  = 2
    ELSE
       TSKST  = 1
    END IF


    XAGRR  = FEREAL - FSREAL
    XAGRP  = FEPLAN - FSPLAN

''''''''''''
'' busca el avance del proyecto





'''''''''''









    if xusrrsp = "NO" then
       sniv = "SIN AFECTAR"
    ELSE
      strsql = "Select * From USUARIOS where ID = '" & xusrrsp & "' "
      set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
      if not rsniv.eof then sniv = rsniv(1)
      call ofv.cerrarconsulta(rsniv)
    END IF



    ITART             = ITART + 1
    ATAREA (ITART,0)  = RS1(0)
    ATAREA (ITART,1)  = IETPT
    ATAREA (ITART,2)  = xnompry
    ATAREA (ITART,3)  = sImageS
    ATAREA (ITART,4)  = CDESP
    ATAREA (ITART,5)  = CDESR
    ATAREA (ITART,6)  = CHASP
    ATAREA (ITART,7)  = CHASR 
    ATAREA (ITART,8)  = TSKst
    ATAREA (ITART,9)  = sniv
    ATAREA (ITART,10) = wtskci
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
    ATAREA (ITART,34) = wtskd

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





  END IF

  END IF
    rs1.movenext
    loop
    call ofv.cerrarconsulta(rs1)
  end if

  for itart=1 to ttar

      FSREAL =    ATAREA (ITART,11)
      FEREAL =    ATAREA (ITART,12)
      FSPLAN =    ATAREA (ITART,13)
      FEPLAN =    ATAREA (ITART,14)
      XAGRR  =    ATAREA (ITART,22)
      XAGRP  =    ATAREA (ITART,28)




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

    ATAREA (ITART,4)  = CDESP
    ATAREA (ITART,5)  = CDESR
    ATAREA (ITART,6)  = CHASP
    ATAREA (ITART,7)  = CHASR 



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

''''''''''''''''''''''

''' estilos

SPL  = "font-family:TAHOMA;font-size:3pt;border-bottom-style:solid;border-bottom-width:8;border-bottom-color:#CCEEFF;"
SNI  = "font-family:TAHOMA;font-size:3pt;border-top-style:solid;border-top-width:8;border-top-color:#DDDDDD;"
SNIY = "font-family:TAHOMA;font-size:3pt;border-top-style:double;border-top-width:8;border-top-color:#bbbbbb;"
SIF  = "font-family:TAHOMA;font-size:3pt;border-top-style:solid;border-top-width:8;border-top-color:#6699CC;"
SIC  = "font-family:TAHOMA;font-size:3pt;border-top-style:double;border-top-width:8;border-top-color:#008000;"
SID  = "font-family:TAHOMA;font-size:3pt;border-top-style:double;border-top-width:8;border-top-color:#AA0000;"
SIX  = "font-family:TAHOMA;font-size:3pt;border-top-style:double;border-top-width:8;border-top-color:#f7aa22;"
''''''''''''''''''''''






 '' response.write "t.etps: " & tetp & " t.tareas: " & ttar & "<br>"

 ''      response.write "min: " & mfmin & " max: " & mfmax & " cmax " & cmax & " xxxx " & XXXX &  "<br>" 




    CMAX = MFMAX - MFMIN + 1

    TSEM = CMAX / 7

    IF (CINT(TSEM) * 7) <> CMAX THEN
       TSEM = CINT(TSEM) + 1
       CMAX = TSEM * 7
    ELSE
       TSEM = CINT(TSEM)
    END IF  


  sXX =        "<script language=javascript> " & chr(13)
  sXX = sXX  & "function GUISP(dia,opc) { " & chr(13)
  sXX = sXX  & " var ttar=" & ttar & "; " & chr(13)
  sXX = sXX  & " var iit=0; " & chr(13)
  sXX = sXX  & " for (iit=1;iit<=ttar;iit++) " & chr(13)
  sXX = sXX  & " {  " & chr(13)
  sXX = sXX  & "   if (opc=='1') { " & chr(13)
  sXX = sXX  & "      document.all('D'+ dia + 'C1T'+ iit).style.borderLeftStyle='solid'; document.all('D'+ dia + 'C2T'+ iit).style.borderLeftStyle='solid';"
  sXX = sXX  & "      document.all('D'+ dia + 'C1T'+ iit).style.borderLeftColor='#F7AA22'; document.all('D'+ dia + 'C2T'+ iit).style.borderLeftColor='#F7AA22';"
  sXX = sXX  & "      document.all('D'+ dia + 'C1T'+ iit).style.borderLeftWidth=1; document.all('D'+ dia + 'C2T'+ iit).style.borderLeftWidth=1;"
  sXX = sXX  & "      document.all('CC'+ dia).style.backgroundColor='#F7AA22'; "
  sXX = sXX  & "      document.all('D'+ dia + 'C1T'+ iit).style.backgroundColor='#f7eedd'; document.all('D'+ dia + 'C2T'+ iit).style.backgroundColor='#f7eedd';"
  sXX = sXX  & "   } " & chr(13)
  sXX = sXX  & "   else { " & chr(13)
  sXX = sXX  & "     document.all('D' + dia + 'C1T' + iit).style.borderLeftStyle='none'; document.all('D' + dia + 'C2T' + iit).style.borderLeftStyle='none'; "
  sXX = sXX  & "     document.all('CC'+ dia).style.backgroundColor='#6699cc'; "
  sXX = sXX  & "     document.all('D'+ dia + 'C1T'+ iit).style.backgroundColor=''; document.all('D'+ dia + 'C2T'+ iit).style.backgroundColor='';"
  sXX = sXX  & "   } " & chr(13)
  sXX = sXX  & " } " & chr(13)
  sXX = sXX  & " } "
  sXX = sXX  & "</script>" & chr(13)

  sHTML = sHTML  & sXX



    ETPA = 0

'''    Celda = ofv.GenCelda("","abc colspan=" & (3+CMAX),"center","100%","","","","&nbsp;","si")
    FilaS =  ""



    
    Celdas = ofv.GenCelda("","abc STYLE='background-color:#993333;font-family:verdana;font-size:7pt;color:#ffffff;' colspan=2 rowspan=2 ","center","","","","","<b>PROYECTOS</b>","si")

    fshow = MFMIN

    FOR ISEM = 1 TO TSEM
       celdas = celdas & ofv.GenCelda("","aaa STYLE='border-style:solid;border-width:1;border-color:#eeeeee;background-color:#993333;font-family:verdana;font-size:7pt;color:#ffffff;' colspan=7 " ,"CENTER","","","","","SEMANA " & ISEM,"si")

    NEXT 

    xspan = (ttar * 2) + TETP + 2

 ''   celdas = celdas & ofv.GenCelda("XX","aaa rowspan=" & xspan & " STYLE='font-family:verdana;font-size:6pt;color:#ffffff;border-style:solid;border-color:#993333;' " ,"CENTER","1","","","","&nbsp;","si")

    FilaS =  filaS & ofv.GenRow("","","","","","","",Celdas,"si")

       celdas = ""

    ICOL = 0

    dHOY = clng(date) - clng(cdate("01-01-1900"))

    FOR ISEM = 1 TO TSEM
       for iday = 1 to 7



       ICOL = ICOL + 1 
        xid = "CC" & ICOL

       dFecha = cdate(fshow + clng(cdate("01-01-1900")))
       DFDIA = DatePart("W",dFecha,2)
       ndia = ofv.cdia(Weekday(dFecha))
       nMes = ofv.cmes(Month(dFecha))

       IF DFDIA > 5 and XFFCH = 1 THEN
        ''  XBRD2 = "display:none;"
       else
          xbrd2 = ""
       END IF 
    
       xfshow = "  " & ndia & ", " & Day(dFecha) & " de " & nMes & " de " & Year(dFecha)

''alert(document.all." & xid & ".offsetLeft);document.all.XX.style.pixelLeft=document.all." & xid & ".offsetLeft;alert(document.all.XX.style.Left);


       xlink = "<A  style='color:#ffffff;' onmouseover='this.style.textDecoration=" & chr(34)
       xlink = xlink & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
       xlink = xlink & "none" & chr(34) & ";' "
       xlink = xlink & " href='javascript:void(0);' " 
       xlink = xlink & " title='" & xfshow & "' ><center>"
       if fshow = dhoy then
          xlink = xlink & "<b><big>" & iday & "</big></b>" 
          xbrd = "outset"
          shtml = shtml & " <script language=javascript for=window event=onload > " & chr(13)
          shtml = shtml & " var cnt = 0; uno(); " & chr(13)
          shtml = shtml & " function uno() {  " & chr(13)
          shtml = shtml & " if (document.all.CC" &  ICOL &  ") { GUISP(" & chr(34) & ICOL & chr(34) & "," & chr(34) & "1" & chr(34) & "); " & chr(13)
          shtml = shtml & " } else {  if (cnt > 10) {} else { cnt++; " & chr(13)
          shtml = shtml & " setTimeout (" & chr(34) & "uno()" & chr(34) & ", 350); }}" & chr(13)
          shtml = shtml & " } " & chr(13)
          shtml = shtml & " </script>" & chr(13)

       else
          xlink = xlink & iday
          xbrd = "solid"
       end if 
       xlink = xlink & "</center></A>"

       xlinf = " onmouseover='"
       xlinf = xlinf & " GUISP(" & chr(34) & ICOL & chr(34) & "," & chr(34) & "1" & chr(34) & ");' onmouseout='"
       xlinf = xlinf & " GUISP(" & chr(34) & ICOL & chr(34) & "," & chr(34) & "2" & chr(34) & ");' "

           celdas = celdas & ofv.GenCelda(xid,"aaa " & xlinf & " STYLE='border-style:" & xbrd & xbrd2 & ";border-width:1;border-color:#eeeeee;background-color:#6699cc;font-family:verdana;font-size:6pt;color:#ffffff;' " ,"CENTER","2","","","",xlink,"si")
       fshow = fshow + 1 
       next
    NEXT 


    FilaS =  filaS & ofv.GenRow("","","","","","","",Celdas,"si")



'''    FOR ITART=1 TO TTAR


  for WTARt=1 to ttar

      ITART = AORD(WTART ,0)


'    ITART            = ITART + 1
'    ATAREA (ITART,1) = IETPT
'    ATAREA (ITART,2) = RS1(5)
'    ATAREA (ITART,3) = RS1(4)
'    ATAREA (ITART,4) = CDESP
'    ATAREA (ITART,5) = CDESR
'    ATAREA (ITART,6) = CHASP
'    ATAREA (ITART,7) = CHASR
'    ATAREA (ITART,8) = rs1(18)
'    TTAR             = ITART 

   CDESP = ATAREA (ITART,4)
   CHASP = ATAREA (ITART,6)
   CDESR = ATAREA (ITART,5)
   CHASR = ATAREA (ITART,7)

   PADV = ATAREA (ITART,10)


'    IF ETPA <> ATAREA (ITART,1) THEN
'
'       ETPA = ATAREA (ITART,1)
'
'       sInput  = AETAPA (ETPA,2)
'       XETPORD = sInput
'       celdas = ofv.GenCelda("","aaa bgcolor=#f7aa22   ","","2%","","","",SINPUT,"si")
'
'       sInput = AETAPA (ETPA,1)
'       celdas = celdas & ofv.GenCelda("","vt STYLE='background-color:#ffffff;border-bottom-style:solid;border-bottom-width:2;border-bottom-color:#f7aa22;' colspan=" & (1 + CMAX) ,"","100%","","","",xestr & ": " & sInput,"si")
'       FilaS =  filaS & ofv.GenRow("","","","","","","",Celdas,"si")
'
'
'    END IF



    celdas = "" 

    xidB = "C2T" & ITART
    xidct = 1

    FOR ICOL=1 TO CDESR-1
        xid = "D" & xidct & xidB
        xidct = xidct + 1
        celdas = celdas & ofv.GenCelda(xid,"ut STYLE='font-family:TAHOMA;font-size:3pt;'","","1","","","","&nbsp;","si")
    NEXT

    XTOT = CHASR - CDESR + 1

    if ATAREA (ITART,8) = "0" then
       xcolor = "#dddddd"
       XMSG = "<B><FONT FACE=WINGDINGS COLOR=#6699CC SIZE=2>6</FONT></B>"
       FOR ICOL=1 TO XTOT
           xid = "D" & xidct & xidB
           xidct = xidct + 1
           celdas = celdas & ofv.GenCelda(xid,"AAA STYLE='" & SNI & "'  ","","1","","","","&nbsp;","si")
       NEXT

    elseif ATAREA (ITART,8) = "1" then

'''       if CHASR > CHASP THEN 

       if ATAREA (ITART,34) = 1 THEN 
          xcolor = "#993333"
          XMSG = "<B><FONT COLOR=#800000 SIZE=1 >(" & PADV & "%)</FONT></B>"
          SSX = SID
       ELSEif ATAREA (ITART,34) = 2 THEN 
          xcolor = "#f7aa22"
          XMSG = "<B><FONT COLOR=#f7aa22 SIZE=1 >(" & PADV & "%)</FONT></B>"
          SSX = SIX
       ELSE
          xcolor = "#00aa00"
          XMSG = "<B><FONT COLOR=#008000 SIZE=1 >(" & PADV & "%)</FONT></B>"
          SSX = SIC
       END IF

       XTOT1 = CINT(XTOT * PADV / 100)

       FOR ICOL=1 TO XTOT1
           xid = "D" & xidct & xidB
           xidct = xidct + 1
           celdas = celdas & ofv.GenCelda(xid,"AAA STYLE='" & SSX & "'  ","","1","","","","&nbsp;","si")
       NEXT

       xcolor = "#dddddd"

       XTOT2 = XTOT - XTOT1

       FOR ICOL=1 TO XTOT2
           xid = "D" & xidct & xidB
           xidct = xidct + 1
           celdas = celdas & ofv.GenCelda(xid,"AAA STYLE='" & SNIY & "'  ","","1","","","","&nbsp;","si")
       NEXT


    else
       xcolor = "#6699cc"

       XMSG = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=2>&#0252;</FONT></B>"

       FOR ICOL=1 TO XTOT
           xid = "D" & xidct & xidB
           xidct = xidct + 1
           celdas = celdas & ofv.GenCelda(xid,"AAA STYLE='" & SIF & "'  ","","1","","","","&nbsp;","si")
       NEXT
    end if



    XTOT = CMAX - CHASR

    XMSG = "&nbsp;" & XMSG

 ''   celdas = celdas & ofv.GenCelda("","ut colspan=" & XTOT & " STYLE='font-family:verdana;font-size:6pt;color:#000080;' ","","","","","",XMSG,"si")


    FOR ICOL=CHASR TO CMAX-1
        xid = "D" & xidct & xidB
        xidct = xidct + 1
        celdas = celdas & ofv.GenCelda(xid,"ut STYLE='font-family:wingdings;font-size:6pt;'","","1","","","","&nbsp;","si")
    NEXT

    celdas2 = celdas
    celdas = ""


' ''''''''' primera fila

    sImagenI = "<b><FONT FACE=WEBDINGS COLOR=#F7AA22 SIZE=4>i</FONT></b>"

    xlink = "<A  href='PROYGRAF.asp?CPRY=" & ATAREA (ITART,0) & "' onmouseover='this.style.textDecoration=" & chr(34)
    xlink = xlink & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
    xlink = xlink & "none" & chr(34) & ";' "
    xlink = xlink & " TARGET=_blank "
    xlink = xlink & " >" & sImagenI & "</A>"

    sInput = xlink


    sInputn0 = "<b>" & replace(ATAREA (ITART,2)," ","&nbsp;") & "</b><br>resp.: " & replace(ATAREA (ITART,9)," ","&nbsp;") & "&nbsp;" & sInput & "&nbsp;"  & XMSG


    celdas = ofv.GenCelda("","aaa ROWSPAN=2 bgcolor=#ffffff   ","","1%","","","",ATAREA (ITART,3),"si")
    celdas = celdas & ofv.GenCelda("","ut ROWSPAN=2  STYLE='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#f7aa22;border-right-style:solid;border-right-width:1;border-right-color:#f7aa22;'  ","","25%","","","",sInputn0,"si")

    
    xidB = "C1T" & ITART
    xidct = 1
    FOR ICOL=1 TO CDESP-1
        xid = "D" & xidct & xidB
        xidct = xidct + 1
        celdas = celdas & ofv.GenCelda(xid,"ut  STYLE='font-family:TAHOMA;font-size:3pt;'","","","","","","&nbsp;","si")
    NEXT

    FOR ICOL=CDESP TO CHASP
        xid = "D" & xidct & xidB
        xidct = xidct + 1
        celdas = celdas & ofv.GenCelda(xid,"AAA STYLE='" & SPL & "' ","","","","","","&nbsp;","si")
    NEXT


    FOR ICOL=CHASP TO CMAX-1
        xid = "D" & xidct & xidB
        xidct = xidct + 1
        celdas = celdas & ofv.GenCelda(xid,"ut  STYLE='font-family:TAHOMA;font-size:3pt;'","","","","","","&nbsp;","si")
    NEXT


    FilaS =  filaS & ofv.GenRow("","","","","","","",Celdas,"si")

' ''''''''''
    FilaS =  filaS & ofv.GenRow("","","","","","","",Celdas2,"si")



    NEXT 



    xwidth = 100


    FilaS2 = ofv.GenTabla("","aa  style='border-style:ridge;border-width:2;border-color:#f7aa22;' ","",xwidth & "%","","0","","-2","-1","","",filas,"si")



    sImagen1p = "<FONT FACE=WINGDINGS COLOR=#cceeff SIZE=4>n</FONT>"
    sImagen1n = "<FONT FACE=WINGDINGS COLOR=#dddddd SIZE=4>n</FONT>"
    sImagen1d = "<FONT FACE=WINGDINGS COLOR=#993333 SIZE=4>n</FONT>"
    sImagen1c = "<FONT FACE=WINGDINGS COLOR=#008000 SIZE=4>n</FONT>"
    sImagen1f = "<FONT FACE=WINGDINGS COLOR=#336699 SIZE=4>n</FONT>"
    sImagen1x = "<FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=4>n</FONT>"

    sImagen1pD = "<FONT FACE=TERMINAL COLOR=#cceeff SIZE=4>&#0220;</FONT>"
    sImagen1nD = "<FONT FACE=TERMINAL COLOR=#dddddd SIZE=4>&#0223;</FONT>"
    sImagen1dD = "<FONT FACE=TERMINAL COLOR=#993333 SIZE=4>&#0223;</FONT>"
    sImagen1cD = "<FONT FACE=TERMINAL COLOR=#008000 SIZE=4>&#0223;</FONT>"
    sImagen1fD = "<FONT FACE=TERMINAL COLOR=#336699 SIZE=4>&#0223;</FONT>"


    sImagen6S = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=2>&#0252;</FONT></B>"

    sImagen6T = "<B><FONT FACE=WEBDINGS COLOR=#336699 SIZE=3>q</FONT></B>"

    sImagenE0 = "<B><FONT FACE=WINGDINGS COLOR=#6699CC SIZE=2>6</FONT></B>"

    sImagenP = "<B><FONT FACE=verdana COLOR=#f7aa22 SIZE=1>(nn%)</FONT></B>"

''sImagenI


     celdas = ofv.GenCelda("","","","4%","","","","&nbsp;","si")

     sinput = sImagenI & "&nbsp;Gantt&nbsp;del&nbsp;Proyecto" 
     celdas = celdas & ofv.GenCelda("","ut","","16%","","","",sInput,"si")

     sinput = sImagen1p & "&nbsp;Planificado" 
     celdas = celdas & ofv.GenCelda("","ut","","10%","","","",sInput,"si")

     sinput = sImagen1n & "&nbsp;" & sImagenE0 & "&nbsp;No&nbsp;iniciado" 
     celdas = celdas & ofv.GenCelda("","ut","","11%","","","",sInput,"si")
     sinput = sImagen1c & "&nbsp;En&nbsp;curso&nbsp;dentro&nbsp;del&nbsp;plan" 
     celdas = celdas & ofv.GenCelda("","ut","","16%","","","",sInput,"si")
     sinput = sImagen1d & "&nbsp;En&nbsp;curso&nbsp;desviado" 
     celdas = celdas & ofv.GenCelda("","ut","","16%","","","",sInput,"si")
     sinput = sImagen1x & "&nbsp;En&nbsp;curso&nbsp;demorado"
     celdas = celdas & ofv.GenCelda("","ut","","16%","","","",sInput,"si")
     sinput = sImagen1f & "&nbsp;Finalizado" 
     celdas = celdas & ofv.GenCelda("","ut","","11%","","","",sInput,"si")



     fila = ofv.GenRow("","","","","","","",celdas,"si")

     Fila = ofv.GenTabla("","aaa bgcolor=#ffffff ","",xwidth & "%","","0","","4","0","","",fila,"si")







    Celda = ofv.GenCelda("","","center","","","","",FilaS1 & FilaS2 & Fila,"si")
    FilaS = ofv.GenRow("","","","","","","",Celda,"si")




    sHTML = sHTML & ofv.GenTabla("","aa style='filter:shadow(color=#cccccc,direction=125);' ","",xwidth & "%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"


''' Session("CMFMENU")


 
  IF session("ModMain2") = "CL" THEN
     sinput = "<a  class=tt href=Clientes.asp"
     sinput = sinput  & " ><b>&nbsp;Volver&nbsp;a&nbsp;Cliente&nbsp;</b></a>" 
     celdas = ofv.GenCelda("","tt","","10%","","","",sInput,"si")

  ELSE
     sinput = "&nbsp;"
     celdas =  ofv.GenCelda("","","","1","","","",sInput,"si")
  END IF


  celdas = celdas & ofv.GenCelda("","","","","","","","&nbsp;","si")
  filas = ofv.GenRow("","","","","","","",celdas,"si")
  sHTML = sHTML  & ofv.GenTabla("","aa ","","98%","","0","","0","0","","",filas,"si")




  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML



%>