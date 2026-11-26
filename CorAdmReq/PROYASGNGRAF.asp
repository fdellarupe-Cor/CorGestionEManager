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

      StrSql = "Select * From PROYTSKASGN where ESTADO < 2 "
      StrSql = StrSql & " AND ((VPLAN = PVPLAN ) "
    ''  StrSql = StrSql & " AND ((VPLAN = PVPLAN and CODIGO <> '" & SESSION("NCLIENTE") & "') "
    ''  StrSql = StrSql & " OR (VPLAN = '" & SESSION("VPLAN") & "' AND CODIGO = '" & SESSION("NCLIENTE") & "')) "
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
    ''  StrSql = StrSql & " AND ((VPLAN = PVPLAN and CODIGO <> '" & SESSION("NCLIENTE") & "') "
    ''  StrSql = StrSql & " OR (VPLAN = '" & SESSION("VPLAN") & "' AND CODIGO = '" & SESSION("NCLIENTE") & "')) "
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
    ''  StrSql = StrSql & " AND ((VPLAN = PVPLAN and CODIGO <> '" & SESSION("NCLIENTE") & "') "
    ''  StrSql = StrSql & " OR (VPLAN = '" & SESSION("VPLAN") & "' AND CODIGO = '" & SESSION("NCLIENTE") & "')) "
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
    ''  StrSql = StrSql & " AND ((VPLAN = PVPLAN and CODIGO <> '" & SESSION("NCLIENTE") & "') "
    ''  StrSql = StrSql & " OR (VPLAN = '" & SESSION("VPLAN") & "' AND CODIGO = '" & SESSION("NCLIENTE") & "')) "
      StrSql = StrSql & " ) "
      StrSql = StrSql & " order by GERENCIA, SECTOR, USRRESP, FINICIO, CODIGO, orden, TSKORD "

END SELECT






  ''    XFCURR = XFSTART

         dFecha = CDATE(CLNG(XFSTART) + clng(cdate("01-01-1900")))

      celdas = ofv.GenCelda("","vt colspan=4","","","","","",sImagen & "&nbsp;<BIG><B>TAREAS DE " & XDESTN & "&nbsp;&nbsp;&nbsp;&nbsp; AL " & ofv.fechador & "</B></BIG>","si")
      filas  = ofv.GenRow("","","","","10","","",celdas,"si")


      celdas = ofv.GenCelda("","AA COLSPAN=4","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      FilaS1 = ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")




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
    XSTSSIGN  = OWKF.ATABTSK(cint(XIIX),22)

  XVA = "S"

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


    IF MFMAX < FEPLAN THEN MFMAX = FEPLAN
    IF MFMAX < FEREAL THEN MFMAX = FEREAL

    IF MFMIN = 0 THEN MFMIN = FSPLAN

    IF MFMIN > FSPLAN THEN MFMIN = FSPLAN
    IF MFMIN > FSREAL THEN MFMIN = FSREAL






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
    ATAREA (ITART,34) = rs1(5)
    ATAREA (ITART,35) = XSTSSIGN

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

 end if

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
SIT  = "font-family:TAHOMA;font-size:3pt;border-top-style:double;border-top-width:8;border-top-color:#F7AA22;"
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
  sXX = sXX  & "      document.all('D'+ dia + 'C1T'+ iit).style.borderLeftColor='#993333'; document.all('D'+ dia + 'C2T'+ iit).style.borderLeftColor='#993333';"
  sXX = sXX  & "      document.all('D'+ dia + 'C1T'+ iit).style.borderLeftWidth=1; document.all('D'+ dia + 'C2T'+ iit).style.borderLeftWidth=1;"
  sXX = sXX  & "      document.all('CC'+ dia).style.backgroundColor='#993333'; "
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



    
    Celdas = ofv.GenCelda("","abc STYLE='background-color:#f7aa22;font-family:verdana;font-size:7pt;color:#ffffff;' colspan=2 rowspan=2 ","center","","","","","<b>ETAPAS/TAREAS</b>","si")

    fshow = MFMIN

    FOR ISEM = 1 TO TSEM
       celdas = celdas & ofv.GenCelda("","aaa STYLE='border-style:solid;border-width:1;border-color:#eeeeee;background-color:#f7aa22;font-family:verdana;font-size:7pt;color:#ffffff;' colspan=7 " ,"CENTER","","","","","SEMANA " & ISEM,"si")

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
             IF ATAREA (ITART,35) = 0 THEN
                xcolor = "#00aa00"
                XMSG = "<B><FONT COLOR=#008000 SIZE=1 >(" & PADV & "%)</FONT></B>"
                SSX = SIC
     
             ELSEIF ATAREA (ITART,35) = 1  THEN
                xcolor = "#F7AA22"
                XMSG = "<B><FONT COLOR=#F7AA22 FACE=WEBDINGS SIZE=2 >4</FONT><FONT COLOR=#008000 SIZE=1 >(" & PADV & "%)</FONT></B>"
                SSX = SIT

             ELSEIF ATAREA (ITART,35) = 2  THEN
                xcolor = "#993333"
                XMSG = "<B><FONT COLOR=#800000 SIZE=1 >(" & PADV & "%)</FONT></B>"
                SSX = SID

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

    xlink = "<A  href='&nbsp;' onmouseover='this.style.textDecoration=" & chr(34)
    xlink = xlink & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
    xlink = xlink & "none" & chr(34) & ";' "
    xlink = xlink & " onclick=" & chr(34) & "AbrirCons('PROYGRAFTSK.asp?TSK=" & ATAREA (ITART,0) & "&PRY=" & ATAREA (ITART,24) & "&FSPLAN=" & ATAREA (ITART,13) & "&FEPLAN=" & ATAREA (ITART,14)  & "&FSREAL=" & ATAREA (ITART,11) & "&FEREAL=" & ATAREA (ITART,12)  & "&STSSIGN=" & ATAREA (ITART,35)  & "'); return false;" & chr(34)
    xlink = xlink & " >" & sImagenI & "</A>"

    sInput = xlink

    IF SRESP <> "" THEN

    sInputn0 = "<b>" & replace(ATAREA (ITART,2)," ","&nbsp;") & "</b><br>" & replace(ATAREA (ITART,29)," ","&nbsp;") & "<br>&nbsp;&nbsp;resp.:<b>" & replace(ATAREA (ITART,9)," ","&nbsp;") & "</b>"

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

     sInputn0 = sInputn0 & "<br>&nbsp;" & sInput & "&nbsp;"  & XMSG

    ELSE

    sInputn0 = "<b>" & replace(ATAREA (ITART,2)," ","&nbsp;") & "</b><br>&nbsp;&nbsp;resp.:<b>" & replace(ATAREA (ITART,9)," ","&nbsp;") & "</b>"

     sInputn0 = sInputn0 & "&nbsp;" & sInput & "&nbsp;"  & XMSG

    END IF





    celdas = ofv.GenCelda("","aaa ROWSPAN=2 bgcolor=#ffffff   ","","1%","","","",XETPORD & "." & ATAREA (ITART,3),"si")
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

     sinput = sImagenI & "&nbsp;Info&nbsp;de&nbsp;la&nbsp;tarea" 
     celdas = celdas & ofv.GenCelda("","ut","","16%","","","",sInput,"si")

     sinput = sImagen1p & "&nbsp;Planificado" 
     celdas = celdas & ofv.GenCelda("","ut","","10%","","","",sInput,"si")

     sinput = sImagen1n & "&nbsp;" & sImagenE0 & "&nbsp;No&nbsp;iniciada" 
     celdas = celdas & ofv.GenCelda("","ut","","11%","","","",sInput,"si")
     sinput = sImagen1c & "&nbsp;En&nbsp;curso&nbsp;dentro&nbsp;del&nbsp;plan" 
     celdas = celdas & ofv.GenCelda("","ut","","16%","","","",sInput,"si")
     sinput = sImagen1d & "&nbsp;En&nbsp;curso&nbsp;desviada" 
     celdas = celdas & ofv.GenCelda("","ut","","16%","","","",sInput,"si")
     sinput = sImagen1f & "&nbsp;" & sImagen6S  & "&nbsp;Finalizada" 
     celdas = celdas & ofv.GenCelda("","ut","","11%","","","",sInput,"si")

     sinput = SIMAGENP & "&nbsp;Porcentaje&nbsp;de&nbsp;avance" 
     celdas = celdas & ofv.GenCelda("","ut","","16%","","","",sInput,"si")




     fila = ofv.GenRow("","","","","","","",celdas,"si")

     Fila = ofv.GenTabla("","aaa bgcolor=#ffffff ","",xwidth & "%","","0","","4","0","","",fila,"si")







    Celda = ofv.GenCelda("","","center","","","","",FilaS1 & FilaS2 & Fila,"si")
    FilaS = ofv.GenRow("","","","","","","",Celda,"si")




    sHTML = sHTML & ofv.GenTabla("","aa  ","",xwidth & "%","","0","","0","0","","",filas,"si")
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