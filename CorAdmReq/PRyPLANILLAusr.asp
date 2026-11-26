<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0



  SESSION("TSKVLV") = "" 
  session("NCliente") = ""


  sHTML = ofv.formHeader("")

  sHTML = sHTML  & "<script language=javascript> "
''  sHTML = sHTML  & "        document.getElementById('MDPTOP1').style.display='none';   document.getElementById('tbhMAX').style.display='none';       hdemenu();     "

  sHTML = sHTML  & "function AbrirCons(Datos) { "
  sHTML = sHTML  & " myWin= open(Datos," & chr(34) & "Cons2" & chr(34) & "," & chr(34) & "status=yes,scrollbars=yes,resizable=yes,menubar=no,toolbar=no,height=450,width=600" & chr(34) & " ); "
  sHTML = sHTML  & " } "
  sHTML = sHTML  & "</script>"

  sHTML = sHTML  & GENINCIDENCIASHED()


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

 XESTAD  = request.querystring("XESTAD")
 XSET    = request.querystring("XSET")
 XOPT    = request.querystring("XOPT")
 sopcion = request.querystring("opcion")
 sRESP   = session("usuario")


 IF XOPT   = "" THEN XOPT   = "XRESP"
 IF XSET   = "" THEN XSET   = "XPT"
 IF XESTAD = "" THEN XESTAD = "1"


    DIM ATAREA(200,40)
    DIM AORD(200,2)
    DIM APROY(200,2)
    DIM Cproy
    DIM CSector

 dim gtctabmes(12,12)
 dim gtctaDIAS(31)

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

 FOR YYY=1 to 31
     gtctaDIAS(YYY) = 0
 next

    
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

      StrSql = "Select * From PROYTSKASGN where ESTADO > 0 AND USRRESP = '" & sRESP  & "' "
      StrSql = StrSql & " AND ((VPLAN = PVPLAN ) "
      StrSql = StrSql & " ) "

      IF XESTAD <> "2" then 

         StrSql = StrSql & " order by FINICIO, CODIGO, orden, TSKORD "

      ELSE

         StrSql = StrSql & " order by FECHAEREAL DESC, CODIGO, orden, TSKORD "

      END IF




 ''response.write strsql




    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"



    XFSTART = 0

    smes  = CINT("0" & request.querystring("MM"))
    sANIO = CINT("0" & request.querystring("YY"))

    IF SMES = 0 THEN

       smes  = CINT("0" & SESSION("MM"))
       sANIO = CINT("0" & SESSION("YY"))

    END IF

    IF SMES = 0 THEN

       smes  = CINT("0" & Month(DATE()))
       sANIO = CINT("0" & YEAR(DATE()))

    END IF

    sDIAH  = CINT("0" & DAY(DATE()))
    smesH  = CINT("0" & Month(DATE()))
    sANIOH = CINT("0" & YEAR(DATE()))

    SESSION("MM") = smes
    SESSION("YY") = sANIO

    WFMESB = smes
    WFYEAB = sANIO

    WFDIAH =  gtctabmes(WFMESB,1)

    IF WFMESB = 2 THEN
     wky  = WFYEAB
     wkyd = wky - 1900
     wkyb = fix(wkyd / 4)

     wkybt = fix(wky / 4) * 4

     if wkybt = wky then 
        WFDIAH =  WFDIAH + 1
     end if

    END IF

    WFTOPE = WFDIAH

    XFSTART = clng(cdate("01-" & WFMESB & "-" & WFYEAB))  - clng(cdate("01-01-1900")) 
    WKTFECDES = XFSTART
    WKTFECHAS = clng(cdate(WFDIAH & "-" & WFMESB & "-" & WFYEAB))  - clng(cdate("01-01-1900"))

    XFCURR  = XFSTART
    XFHOY   = XFSTART
    XFREAL  = XFSTART

    XXX = OWKF.ADMFERIADOS("C",XFSTART)

    MMS = "S"
    MMA = "S"

    MMSM = 0
    MMSA = 0
    MMAM = 0
    MMAA = 0

    IF SANIO = SANIOH THEN
       IF SMES = SMESH THEN 
          MMS = "N"
          WFTOPE = sDIAH
       END IF 

          MMSM = SMES + 1
          MMSA = SANIO

          IF MMSM > 12 THEN
             MMSM = 1
             MMSA = SANIO + 1
          END IF 


          MMAM = SMES - 1
          MMAA = SANIO

          IF MMAM < 1 THEN
             MMAM = 12
             MMAA = SANIO - 1
          END IF

          MMSMDSC =  gtctabmes(MMSM,0)
          MMAMDSC =  gtctabmes(MMAM,0)


    ELSE
          MMSM = SMES + 1
          MMSA = SANIO

          IF MMSM > 12 THEN
             MMSM = 1
             MMSA = SANIO + 1
          END IF 


          MMAM = SMES - 1
          MMAA = SANIO

          IF MMAM < 1 THEN
             MMAM = 12
             MMAA = SANIO - 1
          END IF

          MMSMDSC =  gtctabmes(MMSM,0)
          MMAMDSC =  gtctabmes(MMAM,0)

    END IF

    IF MMS = "S" THEN
       MMSSET = "<A CLASS=Nt title='Mes siguiente.' HREF='PRYPLANILLAusr.asp?YY=" & MMSA & "&MM=" & MMSM & "' >" & MMSMDSC & "&nbsp;" & MMSA & "&nbsp;-</A>"

    ELSE
       MMSSET = "&nbsp;"
    END IF

    IF MMA = "S" THEN
       MMASET = "<A CLASS=Nt title='Mes anterior.' HREF='PRYPLANILLAusr.asp?YY=" & MMAA & "&MM=" & MMAM & "' >-&nbsp;" & MMAMDSC & "&nbsp;" & MMAA & "</A>"
    ELSE
       MMASET = "&nbsp;"
    END IF

    MESDSC =  gtctabmes(SMES,0)

    celdas = ofv.GenCelda("","UT BGCOLOR='#F7EEDD' style='border-style:solid;border-width:1;border-color:#000000;' ","center","20%","","","",MMASET,"si")
    celdas = celdas & ofv.GenCelda("","VT BGCOLOR='#AAAAAA' style='border-style:solid;border-width:1;border-color:#000000;color:#ffffff;' ","center","60%","","","","---- MES ACTUAL: " & MESDSC & "&nbsp;" & SANIO & " ----","si")
    celdas = celdas & ofv.GenCelda("","UT BGCOLOR='#F7EEDD' style='border-style:solid;border-width:1;border-color:#000000;' ","center","20%","","","",MMSSET,"si")

    filas = ofv.GenRow("","","","","22","","",celdas,"si")
    sHTML = sHTML & "<BR>" & ofv.GenTabla("","","","95%","","0","","0","0","","",filas,"si")



    XESTADDSC = "OTRAS"

    IF XESTAD = "0" then XESTADDSC = "PLANIFICADAS"
    IF XESTAD = "1" then XESTADDSC = "EN CURSO"
    IF XESTAD = "2" then XESTADDSC = "FINALIZADAS"


    XESTADDSC = "<FONT COLOR='#993333' >" & XESTADDSC & "</FONT>"



  ''    XFCURR = XFSTART

         dFecha = CDATE(CLNG(XFSTART) + clng(cdate("01-01-1900")))

      WKUSRDSC =  sImagen & "&nbsp;TAREAS ASIGNADAS A " & XDESTN & "&nbsp;&nbsp;DEL " & GTFechadorsh(WKTFECDES) & "&nbsp;&nbsp; AL " & GTFechadorsh(WKTFECHAS) & ""

      celdas = ofv.GenCelda("","UTI colspan=35 style='border-style:solid;border-width:1;border-color:#000000;' ","","","","","",WKUSRDSC,"si")
      filas  = ofv.GenRow("","","","","10","","",celdas,"si")



''''''''''''''''
  tskant = "no"

    TOTTSKREG = 1

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





 ' IF SESSION("FECOPC") = "2" THEN
     XFECDES = FSREAL

     XFECHAS = FEREAL

    '''''''''''''''     IF CLNG(SESSION("FECDES")) > CLNG(XFECHAS) OR CLNG(SESSION("FECHAS")) < CLNG(XFECDES) THEN
  
     IF CLNG(WKTFECDES) > CLNG(XFECHAS) OR CLNG(WKTFECHAS) < CLNG(XFECDES) THEN
        XVA = "N"
     END IF

   


 ' END IF 







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
    ATAREA (ITART,34) = XVPLAN
    ATAREA (ITART,35) = xoriginante 
    ATAREA (ITART,36) = XORIGINANTESECTOR
    ATAREA (ITART,37) = rs1(20)
    
    
    call AddCol(CProy ,rs1(5) ,"","")
    call AddCol(CSector ,XORIGINANTESECTOR ,"","")
    
    

    TTAR              = ITART 

 IF XESTAD = "2" THEN 


    IF ITART = 1 THEN



          AORD(ITART ,0) = ITART 
          AORD(ITART ,1) = 99999 - FEREAL
          AORD(ITART ,2) = 99999 - FSREAL

       
    ELSE

       FOR WTART=1 TO TTAR - 1
           IF (99999 - FEREAL) < AORD(WTART ,1) THEN
              EXIT FOR
           ELSEIF (99999 - FEREAL) = AORD(WTART ,1) THEN
                  IF (99999 - FSREAL) < AORD(WTART ,2) THEN
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
       AORD(WTART ,1) = 99999 - FEREAL
       AORD(WTART ,2) = 99999 - FSREAL

    END IF

 ELSE

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
  END IF



    rs1.movenext

    TOTTSKREG = TOTTSKREG + 1

    IF XESTAD = "2" AND ITART = 20 THEN EXIT DO


    loop
    call ofv.cerrarconsulta(rs1)
  end if

  
  for itart=1 to ttar



      FOR WTART=1 TO TTAR

        IF ATAREA (ITART,0) <> ATAREA (WTART,0) THEN

          IF  ATAREA (ITART,11) <= ATAREA (WTART,12) AND ATAREA (ITART,12) >= ATAREA (WTART,11)  THEN

              IF (ATAREA (ITART,11) = ATAREA (WTART,12) OR ATAREA (ITART,12) = ATAREA (WTART,11))  THEN



              ELSE

                  ATAREA (ITART,30) =  1
                  ATAREA (WTART,30) =  1 
              END IF 

          END IF

        END IF 


      NEXT
      


  NEXT
 

      celdas = ""
      
      

      celdas = celdas & ofv.GenCelda("","ut BGCOLOR='#999999' STYLE='COLOR:#FFFFFF;' ","","","","","","PROYECTO/TAREA","si")
      celdas = celdas & ofv.GenCelda("","ut BGCOLOR='#999999' STYLE='COLOR:#FFFFFF;' ","CENTER","","","","","INICIO","si")
      celdas = celdas & ofv.GenCelda("","ut BGCOLOR='#999999' STYLE='COLOR:#FFFFFF;' ","CENTER","","","","","FINAL","si")
      celdas = celdas & ofv.GenCelda("","ut BGCOLOR='#999999' STYLE='COLOR:#FFFFFF;' ","CENTER","","","","","%","si")

      FOR YYY=1 TO 31

        IF YYY > WFTOPE THEN
           WKDATO = "-"
           sclase = " BGCOLOR='#CCCCCC' "
        ELSE
           WKDATO = YYY
          swd = datepart("w",dateserial(SANIO,smes,YYY))
          if swd = 1 or swd = 7 then
           sclase = " BGCOLOR='#993333' " 
          else
           sclase = " BGCOLOR='#F78822' " 
          end if 

        END IF  

          celdas = celdas & ofv.GenCelda("","ut " & sclase & " style='border-style:solid;border-width:1;border-color:#888888;COLOR:#FFFFFF' ","","2%","","","",WKDATO,"si")
      NEXT

      
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

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




    sImagenI = "<b><FONT FACE=WEBDINGS COLOR=#F7AA22 SIZE=3>i</FONT></b>"

    xlink = "<A  href='&nbsp;' onmouseover='this.style.textDecoration=" & chr(34)
    xlink = xlink & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
    xlink = xlink & "none" & chr(34) & ";' "
    xlink = xlink & " onclick=" & chr(34) & "AbrirCons('PROYGRAFTSK.asp?TSK=" & ATAREA (ITART,0) & "&FSPLAN=" & ATAREA (ITART,13) & "&FEPLAN=" & ATAREA (ITART,14)  & "&FSREAL=" & ATAREA (ITART,11) & "&FEREAL=" & ATAREA (ITART,12)  & "&STSSIGN=" & ATAREA (ITART,16) & "&PADV=" & PADV & "'); return false;" & chr(34)
    xlink = xlink & " >" & sImagenI & "</A>"

    sInputv1 = xlink

    if XESTAD = "1" THEN

   
          sinput   = "<a style='text-decoration:none;color:#336699;' onmouseover='this.style.textDecoration=" & chr(34)
          sinput   = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
          sinput   = sinput  & "none" & chr(34) & ";'  href='PROYtaskAdv.asp?OPCION=PCA&TSK=" & ATAREA (ITART,37) & "&PRY=" & ATAREA (ITART,3) & "&VPLAN=" &  ATAREA (ITART,34) & "' "
          sinput   = sinput  & "  TITLE='Actualizar Avance de la Tarea'  >" & sImagen6S & "</a>" 
          sInputv2 = sinput
    ELSE

          sInputv2 = "-"

    END IF

     sInputn0 = "<b>" &  ATAREA (ITART,2) & "</b>"

       celdas = ""

      celdas = celdas & ofv.GenCelda("","ut style='border-style:solid;border-width:1;border-color:#888888;' ","","20%","","","",replace(ATAREA (ITART,29)," " ,"&nbsp;") & "<BR>" & replace(sInputn0&""," " ,"&nbsp;"),"si")
      celdas = celdas & ofv.GenCelda("","ut style='border-style:solid;border-width:1;border-color:#888888;' ","CENTER","5%","","","",sInputn4,"si") ' INICIO  REAL
      celdas = celdas & ofv.GenCelda("","ut style='border-style:solid;border-width:1;border-color:#888888;' ","CENTER","5%","","","",sInputn5 ,"si") ' FIN REAL
      celdas = celdas & ofv.GenCelda("","UT style='border-style:solid;border-width:1;border-color:#888888;' ","CENTER","","3%","","",sinput3,"si")    

  ''  WFMESB = CSTR(Month(DATE()))
  ''  WFYEAB = CSTR(YEAR(DATE()))
  ''  WFDIAH =  gtctabmes(WFMESB,1)

      FOR YYY=1 TO 31

        IF YYY > WFTOPE THEN
           WKDATO = "-"
           sclase = " BGCOLOR='#DDDDDD' "
           celdas = celdas & ofv.GenCelda("","ut " & sclase & " style='border-style:solid;border-width:1;border-color:#888888;COLOR:#FFFFFF' ","","2%","","","",WKDATO,"si")
        ELSE
           WKDATO = YYY




          WKTFECBUS = clng(cdate(YYY & "-" & WFMESB & "-" & WFYEAB))  - clng(cdate("01-01-1900"))

          StrSql = "Select * From PROYTAREASADV where codigo = '" & ATAREA (ITART,3) & "' AND TSK = '" & ATAREA (ITART,37)  & "' AND FUPD = '" & WKTFECBUS & "' "
          set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if NOT rsX.eof then
             TTIPO = RSX(7)
             TADV = 0
             wktotal = 0
           '  DO UNTIL RSX.EOF 
                TADV = TADV + RSX(4)
            ' RSX.MOVENEXT
           '  LOOP
             sInputt = "&nbsp;H"


             select case TTIPO

                    case 1
                         sInputt = "&nbsp;H"
                         wktotal = TADV

                    case 2
                         sInputt = "&nbsp;D"
                         wktotal = (TADV * CLNG(SESSION("DPP")))
                    case 3
                         sInputt = "&nbsp;S"
                         wktotal = (TADV * CLNG(SESSION("DPP")) * CLNG(SESSION("SPP")) )
                    case 4
                         sInputt = "&nbsp;m"
                         wktotal = round(TADV / 60,2)
                    case else
                         sInputt = "&nbsp;H"
                         wktotal = TADV
             end select 

             gtctaDIAS(YYY) = gtctaDIAS(YYY) + wktotal

          sinput   = "<a style='text-decoration:none;color:#336699;' onmouseover='this.style.textDecoration=" & chr(34)
          sinput   = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
          sinput   = sinput  & "none" & chr(34) & ";'  href='PROYtaskAdvmod.asp?MODAL=MOD&FUPD=" & WKTFECBUS & "&ADV=" & RSX(0) & "&TSK=" & ATAREA (ITART,37) & "&PRY=" & ATAREA (ITART,3) & "&VPLAN=" &  ATAREA (ITART,34) & "' "
          sinput   = sinput  & "  TITLE='Modificar/eliminar el Avance de la Tarea'  >" & TADV & sInputt & "</a>" 
          sInputv2 = sinput


             celdas = celdas & ofv.GenCelda("","ut BGCOLOR='#EECCAA' style='border-style:solid;border-width:1;border-color:#000000;' ","","2%","","","",sInputv2,"si")

          ELSE  

          sinput   = "<a style='text-decoration:none;color:#336699;' onmouseover='this.style.textDecoration=" & chr(34)
          sinput   = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
          sinput   = sinput  & "none" & chr(34) & ";'  href='PROYtaskAdvmod.asp?MODAL=NEW&FUPD=" & WKTFECBUS & "&TSK=" & ATAREA (ITART,37) & "&PRY=" & ATAREA (ITART,3) & "&VPLAN=" &  ATAREA (ITART,34) & "' "
          sinput   = sinput  & "  TITLE='Generar nueva inputacion'  >-</a>" 
          sInputv2 = sinput

             celdas = celdas & ofv.GenCelda("","ut BGCOLOR='#EEEEEE' style='border-style:solid;border-width:1;border-color:#BBBBBB;' ","CENTER","2%","","","",sInputv2,"si")

          END IF

        END IF  





      NEXT




      filas = filas & ofv.GenRow("","aaa  BGCOLOR='" & xTrColor & "' ","","","35","","",celdas,"si")
    


	 if xTrColor = "#FFFFFF" then
	    xTrColor = "#F4F4F4" 
	 ELSE
	    xTrColor = "#FFFFFF" 
	 END IF



  next




       celdas = ""

       wktotal = 0
 

      FOR YYY=1 TO 31

        IF YYY > WFTOPE THEN
           WKDATO = "-"
           sclase = " BGCOLOR='#888888' "
          '' celdas = celdas & ofv.GenCelda("","ut " & sclase & " style='border-style:solid;border-width:1;border-color:#888888;COLOR:#FFFFFF' ","","2%","","","",WKDATO,"si")
        ELSE

         swd = datepart("w",dateserial(SANIO,smes,YYY))
        if swd = 1 or swd = 7 then
           sclase = " BGCOLOR='#CCCCCC' " 
        else
           sclase = " BGCOLOR='#F7EEDD' " 
        end if 

         IF  gtctaDIAS(YYY) > 0 THEN

             wktotal = wktotal +   gtctaDIAS(YYY)
             WKDATO = gtctaDIAS(YYY)
          ''   celdas = celdas & ofv.GenCelda("","ut style='border-style:solid;border-width:1;border-color:#990000;color:#ffffff;'  BGCOLOR='#888888' ","","2%","","","",gtctaDIAS(YYY),"si")

          ELSE  
             WKDATO = "-"
          ''   celdas = celdas & ofv.GenCelda("","ut style='border-style:solid;border-width:1;border-color:#990000;'  BGCOLOR='#e5e5e5' ","CENTER","2%","","","","-","si")

          END IF

         END IF

          celdas = celdas & ofv.GenCelda("","Vt " & sclase & " style='border-style:solid;border-width:1;border-color:#990000;COLOR:#000000;'   ","CENTER","2%","","","",WKDATO,"si")

      NEXT

      celdas = ofv.GenCelda("","Vti colspan=4 style='border-style:solid;border-width:1;border-color:#990000;COLOR:#000000;' BGCOLOR='#EEEEEE' ","","","","","","TOTAL&nbsp;:&nbsp;" & wktotal & "hs.","si") & celdas 


    FilaS =  filaS & ofv.GenRow("","","","","35","","",CeldaS,"si")
    

    sHTML = sHTML & ofv.GenTabla("","aa  style='border-style:solid;border-width:1;border-color:#888888;' ","","95%","","0","","1","2","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"


  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","","","","")


'''''''''''''''


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML



%>