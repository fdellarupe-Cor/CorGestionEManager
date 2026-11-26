<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()
 
if ok <> "Y" then   
  sHTML = ok
else     
  response.expires=0 
  
  
  xls=request.querystring("xls")
  
 if xls="S" then
   Response.ContentType = "application/vnd.ms-excel"
 end if
 
    
   
  sHTML = ofv.MenuHeader("",funcion)


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




    DIM ATAREA(200,40)
    DIM AORD(200,2)
    DIM APROY(200,2)
    DIM Cproy
    DIM CSector
    
    Set Cproy = CrearCol()
    Set CSector = CrearCol()

   IPROY = 0
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





 'response.write strsql




    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"



    XFSTART = CLNG(DATE()) - clng(cdate("01-01-1900")) 
    XFCURR  = XFSTART
    XFHOY   = XFSTART
    XFREAL  = XFSTART

XXX = OWKF.ADMFERIADOS("C",XFSTART)








  ''    XFCURR = XFSTART

         dFecha = CDATE(CLNG(XFSTART) + clng(cdate("01-01-1900")))

      celdas = ofv.GenCelda("","ttt colspan=4","","","","","",sImagen & "&nbsp;<B>REPORTE DE ASIGNACION DE " & XDESTN & "&nbsp;&nbsp;&nbsp;&nbsp; AL " & ofv.fechador & "</B>","si")
      filas  = ofv.GenRow("","","","","10","","",celdas,"si")


      celdas = ofv.GenCelda("","AA COLSPAN=4","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      sHTML = sHTML & ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")







''''''''''''''''
  tskant = "no"

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

           StrSql = "Select TIPO,originante,titulo From PROYECTO where CODIGO = '" & RS1(5) & "' "
           set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
           if not  rsX2.eof then

              XTIPO = RSX2(0)
              XORIGINANTE=RSX2(1)&""
              XTitulo=RSX2(2)
              
              strsql = "Select A.DESCRIPCION,B.DESCRI From USUARIOS A,ECO B where A.SECTOR=B.ECO AND A.ID = '" & XORIGINANTE & "' "
	            set rsorig = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
	            if not rsorig.eof then
	               XORIGINANTE = rsorig(0)
	               XORIGINANTESECTOR = rsorig(1)
	           ELSE
	               XORIGINANTE = "NO DETERMINADO"
	               XORIGINANTESECTOR = rsorig(1)
	           END IF
              call ofv.cerrarconsulta(rsorig)
      

           ELSE
              XTIPO = 0
              XORIGINANTE="&nbsp;"

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
    ATAREA (ITART,3)  = rs1(5)    'rs1(0) & "." & rs1(1)
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
    ATAREA (ITART,34) = ""
    ATAREA (ITART,35) = xoriginante 
    ATAREA (ITART,36) = XORIGINANTESECTOR
    ATAREA (ITART,37) = XTitulo
    
    
    call AddCol(CProy ,rs1(5) ,"","")
    call AddCol(CSector ,XORIGINANTESECTOR ,"","")
    
    
   
 
 
 
  'DEPENDENCIAS

    
      XPRYNOM="&nbsp;"
    
      StrSql = "Select * From PROYtareasLST where codigo = '" & rs1(5) & "' and TAREA = '" & rs1(20) & "' order by orden "
      set rs1Dep = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if rs1Dep.eof then
    
          SINPUT = "&nbsp;"
  
    
      ELSE
        do while not rs1Dep.eof
          StrSql = "Select * From PROYECTO WHERE CODIGO = '" & rs1Dep(4) & "' "
          set RS2Dep = ofv.crearconsultaEx(StrSql,cn,1,parametros)
            IF NOT RS2Dep.EOF THEN
             XPRYNOM = XPRYNOM & RS2Dep(2) & "<br>"
            END IF
          call ofv.cerrarconsulta(RS2Dep)
        rs1Dep.movenext
        loop
     
      end if
    
      call ofv.cerrarconsulta(rs1Dep)
    
      ATAREA (ITART,34)=XPRYNOM






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
 


      
      
      
      celdas = ofv.GenCelda("","ttr ","CENTER","","","","","PROYECTO","si")
      celdas = celdas & ofv.GenCelda("","ttr","CENTER","","","","","Nombre","si")
      celdas = celdas & ofv.GenCelda("","ttr ","CENTER","","","","","TAREA","si")
      celdas = celdas & ofv.GenCelda("","ttr","CENTER","","","","","INI.PLAN","si")
      celdas = celdas & ofv.GenCelda("","ttr","CENTER","","","","","INI.REAL","si")
      celdas = celdas & ofv.GenCelda("","ttr","CENTER","","","","","FIN.REAL","si")
      celdas = celdas & ofv.GenCelda("","ttr ","CENTER","","","","","ESTADO/AVANCE","si")
      celdas = celdas & ofv.GenCelda("","ttr ","CENTER","","","","","ORIGINANTE","si")
      celdas = celdas & ofv.GenCelda("","ttr ","CENTER","","","","","SECTOR","si")
      celdas = celdas & ofv.GenCelda("","ttr ","CENTER","","","","","DEPENDENCIAS","si")
      
      filas = ofv.GenRow("","","","","20","","",celdas,"si")

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

                            xTrColor = "#FFFFFF" 

  for WTARt=1 to ttar

      ITART = AORD(WTART ,0)

    IF ATAREA (ITART,24) > 0 THEN
       SINPUT = ATAREA (ITART,24)
       if ATAREA (ITART,25) = 4 then
          SINPUT = SINPUT & "&nbsp;m"
       elseif ATAREA (ITART,25) = 1 then
          SINPUT = SINPUT & "&nbsp;h"
       elseif ATAREA (ITART,25) = 3 then
          SINPUT = SINPUT & "&nbsp;s"
       else
          SINPUT = SINPUT & "&nbsp;d"
       end if 
       SINPUT = "<b>" & SINPUT & "</b>"
    ELSE
       SINPUT = "-"
    END IF 

       SINPUTn9 = "<b> (FT)</b>"

       SINPUT = SINPUT & SINPUTn9

    IF ATAREA (ITART,23) > 0 AND (ATAREA (ITART,26) > 0 AND (ATAREA (ITART,26) <> ATAREA (ITART,24) OR ATAREA (ITART,27) <> ATAREA (ITART,25) )) THEN
       SINPUT = ATAREA (ITART,26)
       XFFTI = ATAREA (ITART,25)
       IF ATAREA (ITART,27) > 0  THEN XFFTI = ATAREA (ITART,27)
       if XFFTI = 4 then
          SINPUT = SINPUT & "&nbsp;m"
       elseif XFFTI = 1 then
          SINPUT = SINPUT & "&nbsp;h"
       elseif XFFTI = 3 then
          SINPUT = SINPUT & "&nbsp;s"
       else
          SINPUT = SINPUT & "&nbsp;d"
       end if 

      SELECT CASE ATAREA (ITART,25)

             CASE 4

                  XVDUR = ATAREA (ITART,24)
             CASE 1

                  XVDUR = ATAREA (ITART,24)

             CASE 2
                  XVDUR = ATAREA (ITART,24) * CLNG(SESSION("DPP"))


             CASE 3 
                  XVDUR = ATAREA (ITART,24) * CLNG(SESSION("SPP")) * CLNG(SESSION("DPP"))

      END SELECT

      SELECT CASE XFFTI

             CASE 4

                  XWDUR = ATAREA (ITART,26)

             CASE 1

                  XWDUR = ATAREA (ITART,26)

             CASE 2
                  XWDUR = ATAREA (ITART,26) * CLNG(SESSION("DPP"))


             CASE 3 
                  XWDUR = ATAREA (ITART,26) * CLNG(SESSION("SPP")) * CLNG(SESSION("DPP"))

      END SELECT

      XWVDUR = CINT(XVDUR / XWDUR * 100)


       SINPUTn9 = " (PT:" & XWVDUR & "%)"
       SINPUT = "<font color=#aa0000><b>" & SINPUT & SINPUTn9 & "</b></font>"


    END IF 



    SINPUTn1 = SINPUT 




   IF ATAREA (ITART,30) > 0 and xopt = "XRESP" THEN

      WSOLAPA = "&nbsp;*"
      WSFONT = "<FONT COLOR=#AA0000 SIZE=1 ><B>"

   ELSE

      WSOLAPA = "&nbsp;"
      WSFONT = "<FONT COLOR=#004400 SIZE=1 ><B>"
   END IF

  

     celdas = ofv.GenCelda("","VT STYLE='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#888888;' ","","7%","","","",WSFONT & ATAREA (ITART,3)   & WSOLAPA & "</B></FONT>" & "&nbsp;","si")
     celdas = celdas & ofv.GenCelda("","ut   STYLE='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#888888;' ","","5%","","","",ATAREA (ITART,29) & "&nbsp;","si")

    sInputn0 = "<b>" &  ATAREA (ITART,2) & "</b>"



   if XOPT = "XGRN" or XOPT = "XGLB" then

     StrSql = "Select * From ECO where ECO = " & ATAREA (ITART,32)
     set rsX2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rsX2.eof then
        sniv = rsX2(1)
     ELSE
        sniv = "NO DETERMINADO"
     END IF
      call ofv.cerrarconsulta(rsX2) 

     sInputn0 = sInputn0 & "<br>&nbsp;&nbsp;Sec.:<b>" & sniv & "</b>"

   end if

   if XOPT = "XGLB" then

     StrSql = "Select * From ECO where ECO = " & ATAREA (ITART,31)
     set rsX2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rsX2.eof then
        sniv = rsX2(1)
     ELSE
        sniv = "NO DETERMINADO"
     END IF
      call ofv.cerrarconsulta(rsX2) 

     sInputn0 = sInputn0 & "<br>&nbsp;&nbsp;Ger.:<b>" & sniv & "</b>"

   end if



       dFecha = cdate(ATAREA (ITART,16) + clng(cdate("01-01-1900")))
       SINPUTn2 = dFecha


       dFecha = cdate(ATAREA (ITART,17) + clng(cdate("01-01-1900")))
       SINPUTn3 = dFecha




   if  ATAREA (ITART,8) = "0" then


       dFecha = cdate(ATAREA (ITART,18) + clng(cdate("01-01-1900")))
       SINPUTn4 = dFecha
       c4 = " style='color:#bbbbbb;' "


       dFecha = cdate(ATAREA (ITART,19) + clng(cdate("01-01-1900")))
       SINPUTn5 = dFecha
       c5 = " style='color:#bbbbbb;' "


          sinput1 = sImagenE0
          sinput3 = SIMAGEN6W
          sinputt2 = SIMAGEN6W



   elseif ATAREA (ITART,8) = "2" then

          dFecha = cdate(ATAREA (ITART,18) + clng(cdate("01-01-1900")))
          SINPUTn4 = dFecha
          c4 = " style='color:#336699;font-weight:bold;' "


          dFecha = cdate(ATAREA (ITART,19) + clng(cdate("01-01-1900")))
          SINPUTn5 = dFecha
          c5 = " style='color:#336699;font-weight:bold;' "


          sinput1 = sImagenE2
          sinput3 = "100%"
          sinputt2 = SIMAGEN6S

   elseif ATAREA (ITART,8) = "1" then


          dFecha = cdate(ATAREA (ITART,18) + clng(cdate("01-01-1900")))
          SINPUTn4 = dFecha
          c4 = " style='color:#336699;font-weight:bold;' "


          sinput1  = sImagenE1
          sinputt2 = SIMAGEN6N

             sinput3 = ATAREA (ITART,10) &  "%"

          IF ATAREA (ITART,10) > 0 THEN

             IF XFHOY < ATAREA (ITART,17) THEN
                sinputt2 = SIMAGEN6S
                c5 = " style='color:#008000;font-weight:bold;' "

             ELSEIF XFHOY = ATAREA (ITART,17) AND ATAREA (ITART,10) > 99  THEN
                sinputt2 = SIMAGEN6S
                c5 = " style='color:#008000;font-weight:bold;' "

             ELSE
                sinputt2 = SIMAGEN6N
                sinputn0 = vfondo & "&nbsp;" & sinputn0
                c5 = " style='color:#993333;font-weight:bold;' "

             END IF

          else

             IF XFHOY > ATAREA (ITART,18) THEN
                sinputt2 = SIMAGEN6A
                c5 = " style='color:#F7AA22;font-weight:bold;' "
             END IF

             IF XFHOY < ATAREA (ITART,17) THEN
             '   sinputt2 = SIMAGEN6S
             '   c5 = " style='color:#008000;font-weight:bold;' "

             ELSE
                sinputt2 = SIMAGEN6N
                c5 = " style='color:#993333;font-weight:bold;' "

             END IF




          END IF


          SINPUTn5 = cdate(ATAREA (ITART,19) + clng(cdate("01-01-1900")))


   end if



      
      celdas = celdas & ofv.GenCelda("","ut STYLE='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#888888;' ","","75%","","","",replace(sInputn0&""," " ," "),"si")
      celdas = celdas & ofv.GenCelda("","ut STYLE='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#888888;' ","","5%","","","",sInputn2 ,"si")
      celdas = celdas & ofv.GenCelda("","ut STYLE='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#888888;' " & c4,"CENTER","5%","","","",sInputn4,"si") ' INICIO  REAL
      celdas = celdas & ofv.GenCelda("","ut STYLE='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#888888;' " & c5,"CENTER","5%","","","",sInputn5 ,"si") ' FIN REAL
      celdas = celdas & ofv.GenCelda("","UT STYLE='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#888888;' ","CENTER","","5%","","",sinput3,"si")
      celdas = celdas & ofv.GenCelda("","UT STYLE='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#888888;' ","CENTER","","5%","","",ATAREA (ITART,35) & "&nbsp;","si")
      celdas = celdas & ofv.GenCelda("","UT STYLE='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#888888;' ","CENTER","","5%","","",ATAREA (ITART,36) & "&nbsp;","si")
      celdas = celdas & ofv.GenCelda("","UT STYLE='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#888888;' ","CENTER","","5%","","",ATAREA (ITART,34) & "&nbsp;","si")

      filas = filas & ofv.GenRow("","aa  BGCOLOR='" & xTrColor & "'  ","","","30","","",celdas,"si")
    


	 if xTrColor = "#FFFFFF" then
	    xTrColor = "#F9F9F9" 
	 ELSE
	    xTrColor = "#FFFFFF" 
	 END IF



  next



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
    
    datoX="TOTAL DE PROYECTOS ASIGNADOS = "
    total=0
    dato=""
    For bb=0 to GetCount(CSector) -1
    
      total=total + cint("0" & Getitem(CSector,GetKey(CSector)(bb),""))
    
      dato=dato & GetKey(CSector)(bb) & "=" & Getitem(CSector,GetKey(CSector)(bb),"") & "<br>"
    
    next

    datoX=datox & total 
   

    celda = ofv.GenCelda("","VT colspan=19","left","","","","",datoX,"si")
    FilaS =  filaS & ofv.GenRow("","","","","30","","",Celda,"si")

    
    datoX="<b>DETALLE POR GERENCIA/SECTOR </b><br>" & dato
    
    celda = ofv.GenCelda("","VT colspan=19","left","","","","",datoX,"si")
    FilaS =  filaS & ofv.GenRow("","","","","30","","",Celda,"si")


    sHTML = sHTML & ofv.GenTabla("","aa   style='border-style:solid;border-width:1;border-color:#888888;' ","","100%","","0","","-1","-1","","",filas,"si")
    tt=ofv.GenTabla("","aa  border ","","100%","","0","","-1","-1","","",filas,"si")
    if xopt = "XRESP" THEN

       cnbot = "<input type=submit class=bt value='GANTT' >"
       Cons = ofv.gencelda("","","","","1%","","",cnbot,"si")
       sAccion = "PROYASGNgraf.asp?RESP=" & SRESP
       Cons = ofv.genform("cons","",sAccion,"_blank","post","",Cons,"si")

       celdas = Cons
       celdas = celdas & ofv.GenCelda("","","CENTER","","80%%","","","&nbsp;","si")
       FilaS = ofv.GenRow("","","","","","","",CeldaS,"si")
       
       sHTML = sHTML & ofv.GenTabla("","aa  ","","90%","","0","","0","0","","",filas,"si")

     else
        celdas = ""
     end if


    sHTML = sHTML & "</center>"





'''''''''''''''


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

if xls="S" then
  response.write tt
else
  response.write sHTML
end if



%>