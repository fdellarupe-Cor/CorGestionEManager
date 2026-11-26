<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok                     
else
  response.expires=0 
           

  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))
   xls=request.querystring("xls")
  if xls="S" then
     Response.ContentType = "application/vnd.ms-excel"
  end if       

  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))     
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn6)
  set cn2 = ofv.conectar(ofv.strconn0)
  set cn3 = ofv.conectar(ofv.strconn2)
  dim adatostarea(1000,4)       
  
  iadatostarea=0         
     
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

 XSET = request.querystring("XSET")             
 XOPT = request.querystring("XOPT")
 sopcion = request.querystring("opcion")
 sRESP = request.querystring("RESP")

IF SESSION("PRCF") <> "S" THEN
   XXXX = OWKF.GETPRYCNFG()
END IF

SESSION("NCLIENTE") = "NO"


 IF XOPT = "" THEN XOPT = "XRESP"
 IF XSET = "" THEN XSET = "XPT"

    YSTART = CLng(Date) - CLng(CDate("01-01-1900"))

  '  XFSTART = CLNG(DATE()) - clng(cdate("01-01-1900")) - 1

 '   RESPONSE.WRITE "FECHA: " & XFSTART 

     XFSTART = CLNG(DATE()) - clng(cdate("01-01-1900")) - 1

  IF SESSION("FECOPC") = "1" THEN


     XFECDESS = XFSTART

     XFECHASS = XFSTART

  END IF 

  IF SESSION("FECOPC") = "2" THEN

 
     XFECDESS = CLNG(SESSION("FECDES"))

     XFECHASS = CLNG(SESSION("FECHAS"))

  END IF 

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
 
      StrSql = "Select * From PROYTSKASGN where PESTADO <> 20 "
      StrSql = StrSql & " AND ((VPLAN = PVPLAN ) "
    ''  StrSql = StrSql & " AND ((VPLAN = PVPLAN and CODIGO <> '" & SESSION("NCLIENTE") & "') "
    ''  StrSql = StrSql & " OR (VPLAN = '" & SESSION("VPLAN") & "' AND CODIGO = '" & SESSION("NCLIENTE") & "')) "
      StrSql = StrSql & " ) "
      StrSql = StrSql & " order by GERENCIA, SECTOR, USRRESP, FINICIO, CODIGO, orden, TSKORD "

   ''   response.write strsql & "<br>"


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

      StrSql = "Select PROYTSKASGN.*,PROYTIPOTSK.valoracion From PROYTSKASGN,PROYTIPOTSK where PROYTSKTIPO=PROYTIPOTSK.id and PESTADO <> 20 AND GERENCIA = " & sRESP
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

      StrSql = "Select PROYTSKASGN.*,PROYTIPOTSK.valoracion From PROYTSKASGN,PROYTIPOTSK where PROYTSKTIPO=PROYTIPOTSK.id and PESTADO <> 20 AND SECTOR = " & sRESP
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

      StrSql = "Select PROYTSKASGN.*,PROYTIPOTSK.valoracion From PROYTSKASGN,PROYTIPOTSK where PROYTSKTIPO=PROYTIPOTSK.id and PESTADO <> 20 AND USRRESP = '" & sRESP  & "' "
      StrSql = StrSql & " AND ((VPLAN = PVPLAN ) "
    ''  StrSql = StrSql & " AND ((VPLAN = PVPLAN and CODIGO <> '" & SESSION("NCLIENTE") & "') "
    ''  StrSql = StrSql & " OR (VPLAN = '" & SESSION("VPLAN") & "' AND CODIGO = '" & SESSION("NCLIENTE") & "')) "
      StrSql = StrSql & " ) "
      StrSql = StrSql & " order by GERENCIA, SECTOR, USRRESP, FINICIO, CODIGO, orden, TSKORD "

END SELECT
     'response.write strsql

   sImagen1 = "<FONT FACE=WINGDINGS SIZE=2>ñ</FONT>"
    sImagenO = "<B><FONT FACE=WEBDINGS COLOR=#336699 SIZE=3>&#0143;</FONT></B>"
    sImagenI = "<B><FONT FACE=WINGDINGS COLOR=#993333 SIZE=2>&#0243;</FONT></B>"
    sImagenE = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0240;</FONT></B>"
    sImagenC = "<B><FONT FACE=WebDINGS COLOR=#008000 SIZE=4>&#0194;</FONT></B>"
    sImagenPI = "<B><FONT FACE=WebDINGS COLOR=#99bbcc SIZE=3>4</FONT></B>"
    sImagenPF = "<B><FONT FACE=WebDINGS COLOR=#0000aa SIZE=3><</FONT></B>"
    sImagenPC = "<B><FONT FACE=WebDINGS COLOR=#6699cc SIZE=3>.</FONT></B>"
    sImagenCP = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=4>v</FONT></B>"
    sImagenI1 = "<B><FONT FACE=WEBDINGS COLOR=#000000 SIZE=3>&#0128;</FONT></B>"
    sImagenI2 = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=4>N</FONT></B>"
    sImagenI3 = "<B><FONT FACE=WEBDINGS COLOR=#cc0000 SIZE=4>&#064;</FONT></B>"

    sImagenO = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=3>&#0128;</FONT></B>"
    sImagenR = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=3>&#0128;</FONT></B>"
    sImagenA = "<B><FONT FACE=WEBDINGS COLOR=#f7aa22 SIZE=3>&#0128;</FONT></B>"
    

 xtt=0

''  StrSql = "Select * From PROYETPTSK where codigo = '" & session("NCliente") & "' AND VPLAN = " & SESSION("VPLAN") & " ORDER BY USRRESP,ORDEN,TSKORD "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then


 

    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"

  celdas = ofv.GenCelda("","cc colspan=4  ","","","","","",sImagen & "&nbsp;<BIG><B>DETALLE DE HORAS INSUMIDAS EN EL DIA DE " & XDESTN & "</B></BIG>","si")
  filas  = ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","AA COLSPAN=4","CENTER","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  sInput1 = "DESDE :" & UCASE(Fechador(XFECDESS)) & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;HASTA :" & UCASE(Fechador(XFECHASS)) 

  celdas = ofv.GenCelda("","VT COLSPAN=4","","","","","","AL&nbsp;" & sInput1,"si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","AA COLSPAN=4","CENTER","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


  celdas = ofv.GenCelda("","AA COLSPAN=4 style='border-top-style:ridge;border-top-width:2;border-top-color:#f7aa22;' ","CENTER","","","","","&nbsp;","si")
  filas2 = filas2 & ofv.GenRow("","","","","10","","",celdas,"si")


  tskant = "no"
  USRANT  = "XX"
  usrtsk  = 0
  prytsk  = 0
  prytskC = 0
  prytskR = 0
  xcant   = 0

  DO UNTIL RS.EOF

  XVA = "S"

  if rs(4) > 0 then


  IF SESSION("SELOPC") = "1" THEN
      StrSql = "Select * From PROYTIPOTSK where ID = 0" & RS(17)
      set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if not  rsX1.eof then
      
         IF RSX1(2) = 0 THEN
            XVA = "N"

         ELSE

           StrSql = "Select TIPO From PROYECTO where CODIGO = '" & RS(5) & "' "
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


 '    SXFSTARTD = OWKF.GETFENDV(rs(20), rs(5))




'    XIIX = OWKF.TIX





'    XFREAL = OWKF.ATABTSK(cint(XIIX),1)
'    TSKFS  = OWKF.ATABTSK(cint(XIIX),2)
'    TSKFE  = OWKF.ATABTSK(cint(XIIX),3)
'    TSKFSR = OWKF.ATABTSK(cint(XIIX),4)
'    TSKFER = OWKF.ATABTSK(cint(XIIX),5)
'    TSKST  = OWKF.ATABTSK(cint(XIIX),6)
'    FSPLAN = OWKF.ATABTSK(cint(XIIX),7)
'    FEPLAN = OWKF.ATABTSK(cint(XIIX),8)
'    FSREAL = OWKF.ATABTSK(cint(XIIX),9)
'    FEREAL = OWKF.ATABTSK(cint(XIIX),10)
'    XFFDU  = OWKF.ATABTSK(cint(XIIX),11)
'    XFFTI  = OWKF.ATABTSK(cint(XIIX),12)
'    PADV   = OWKF.ATABTSK(cint(XIIX),13)
'    XAGRD  = OWKF.ATABTSK(cint(XIIX),14)
'    XAGRR  = OWKF.ATABTSK(cint(XIIX),15)
'    XMODA  = OWKF.ATABTSK(cint(XIIX),16)
'    XTDUR  = OWKF.ATABTSK(cint(XIIX),17)
'    XUDUR  = OWKF.ATABTSK(cint(XIIX),18)
'    XTDUP  = OWKF.ATABTSK(cint(XIIX),19)
'    XUDUP  = OWKF.ATABTSK(cint(XIIX),20)
'    XAGRP  = OWKF.ATABTSK(cint(XIIX),21)



    FSPLAN = rs(3)
    FEPLAN = rs(10)
    FSREAL = rs(9)
    FEREAL = rs(11)

  if rs(4) = 1 then
     StrSql = "Select * From PROYtareas where id = " & rs(2)
     Set RS7 = OFV.crearconsultaEx(StrSql, CN, 1, "")
     If Not RS7.EOF Then
        XIIX = OWKF.GETTSKCLC(RS7, YSTART, YSTART, rs(35))
        XIIX = OWKF.TIX
        FEREAL = OWKF.ATABTSK(cint(XIIX),10)
     end if
  end if



  '  IF FSREAL < (XFSTART + 1) AND FEREAL > (XFSTART - 1) THEN


  IF SESSION("FECOPC") = "2" THEN
     XFECDES = FSREAL

     XFECHAS = FEREAL

      
     IF CLNG(XFECDESS) > CLNG(XFECHAS) OR CLNG(XFECHASS) < CLNG(XFECDES) THEN
        XVA = "N"
     END IF

   


  END IF 

  IF SESSION("FECOPC") = "1" THEN

     XFECDES = FSREAL

     XFECHAS = FEREAL
      
     IF CLNG(XFECDESS) > CLNG(XFECHAS) OR CLNG(XFECHASS) < CLNG(XFECDES) THEN
        XVA = "N"
     END IF

   


  END IF 

  IF XVA = "S" THEN

     if   tskant <> rs(2) then
          tskant = rs(2) 
           xtt=xtt+1
     xcant = xcant + 1 
    
     IF USRANT <> RS(13) THEN
     
     if usrtsk > 0 then

      
      celdas = ofv.GenCelda("","vt COLSPAN=2  ","","","","","","TOTAL " & USRDESC,"si")
      SINPUT = usrtsk
   '   celdas = celdas & ofv.GenCelda("","Ut","RIGHT","","","","",sInput,"si")
      SINPUT = usrtskc
      celdas = celdas & ofv.GenCelda("","Vt COLSPAN=2 ","RIGHT","","","","",sInput,"si")
      SINPUT = usrtskr
 '     celdas = celdas & ofv.GenCelda("","Ut","RIGHT","","","","",sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      celdas = ofv.GenCelda("","vt COLSPAN=2 ","","","","","",USRDESC,"si")
      SINPUT = usrtsk
     ' celdas = celdas & ofv.GenCelda("","Ut","RIGHT","","","","",sInput,"si")
      SINPUT = usrtskc 
      celdas = celdas & ofv.GenCelda("","Ut COLSPAN=2 ","RIGHT","","","","",sInput,"si")
      SINPUT = usrtskr
     ' celdas = celdas & ofv.GenCelda("","Ut","RIGHT","","","","",sInput,"si")
      filas2 = filas2 & ofv.GenRow("","","","","10","","",celdas,"si")

     end if


    if RS(13) = "NO" then
      sInput = "SIN ASIGNAR"

    else
      strsql = "Select * From USUARIOS where ID = '" & rs(13) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
        sInput = sImagenR & "&nbsp;<B>" & sniv & "</B>"
    END IF 

        USRDESC = sInput
        USRANT  = RS(13)
        usrtsk  = 0
        usrtskC = 0
        usrtskR = 0

     ' celdas = ofv.GenCelda("","vt style='border-BOTTOM-style:solid;border-BOTTOM-width:1;border-BOTTOM-color:#EEEEEE;' ","","","","","","<font size=2>" & USRDESC & "</font>","si")

     ' celdas = CELDAS & ofv.GenCelda("","AA COLSPAN=3 style='border-BOTTOM-style:solid;border-BOTTOM-width:1;border-BOTTOM-color:#EEEEEE;' ","CENTER","","","","","&nbsp;","si")
      'filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

        FILAS1 = ""
     END IF


        USRTMP = RS(15) & ""
        USRTMPU = RS(16)
        IF USRTMP = "" THEN
           USRTMP = 1
           USRTMPU = 1
        END IF
        USRTMP = CINT(USRTMP)
        IF USRTMP = 0 THEN
           USRTMP = 1 
           USRTMPU = 1
        END IF

      XWVDUR = 100


   IF RS(23) > 0 THEN 



      IF RS(21) > 0 THEN
         XTDUR = RS(21)
      ELSE
         XTDUR = RS(15)
      END IF
   
      IF RS(22) > 0 THEN
         XUDUR = RS(22)
      ELSE
         XUDUR = RS(16)
      END IF

      SELECT CASE RS(16)

             CASE 4

                  XVDUR = 1

             CASE 1

                  XVDUR = RS(15)

             CASE 2
                  XVDUR = RS(15) * CLNG(SESSION("DPP"))


             CASE 3 
                  XVDUR = RS(15) * CLNG(SESSION("SPP")) * CLNG(SESSION("DPP"))

      END SELECT

      SELECT CASE XUDUR

             CASE 4

                  XWDUR = 1

             CASE 1

                  XWDUR = XTDUR

             CASE 2
                  XWDUR = XTDUR * CLNG(SESSION("DPP"))


             CASE 3 
                  XWDUR = XTDUR * CLNG(SESSION("SPP")) * CLNG(SESSION("DPP"))

      END SELECT

     '   RESPONSE.WRITE " A" & XVDUR & " B" & XWDUR &  " C" & RS(15)  & " D" & XTDUR &  " E" & RS(16)  & " F" & RS(23)& " <BR> " 


      XWVDUR = CINT(XVDUR / XWDUR * 100)


   END IF

        SELECT CASE USRTMPU
               CASE 4
                    xagrd = USRTMP / 60

               CASE 1

                   IF USRTMP > SESSION("DPP") THEN
                      xagrd = CLNG(SESSION("DPP"))
                   ELSE
                      xagrd = USRTMP
                   END IF 

               CASE 2

                    xagrd = CLNG(SESSION("DPP"))
 

               CASE ELSE

                    xagrd = CLNG(SESSION("DPP"))

        END select         

  '      RESPONSE.WRITE XWVDUR & " <BR> " 

                    xagrd = round((xagrd * XWVDUR / 100),2)

                    usrtsk =  usrtsk + xagrd
                    prytsk =  prytsk + xagrd


                xagrc = 0
                xagrr = 0

   if rs(4) > 0 then



          PADV = 0 


           
          StrSql = "Select * From PROYTAREASADV where "
          StrSql = StrSql & " ( " & XFECDESS & " <=  FUPD and  " & XFECHASS & " >=  FUPD )"  
          StrSql = StrSql & " AND codigo = '" & rs(5) & "' AND TSK = '" & RS(20) & "' ORDER BY id desc "
          set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if NOT rsX.eof then


             PADV  = RSX(5)
             
             
             xagrC = 0
             xagrCInc=0
             xagrCIncstr=""
          do while not rsx.eof

       SELECT CASE USRTMPU
               CASE 4
                    if rs("valoracion")=1 then
                      xagrC = xagrC + (round((RSX(4) / 60),2))
                    else
                      xagrCInc=xagrCInc + (round((RSX(4) / 60),2))
                    end if  

               CASE ELSE
                   
                    if rs("valoracion")=1 then
                       xagrC = xagrC + clng(RSX(4))
                     else
                       xagrCInc = xagrCInc + clng(RSX(4))
                      ' xagrCIncstr=xagrCIncstr & cstr(RSX(4)) & "<br>"
                    end if 

        END select         


          rsx.movenext
          loop 
   

           '     xagrc = CINT((XAGRD * PADV / 100)) 
           '     xagrr = XAGRD - CINT((XAGRD * PADV / 100)) 
           
   if rs("valoracion")=1 then        
    adatostarea(iadatostarea,0)=rs("fechasplan")
    adatostarea(iadatostarea,1)=rs(8)&"&nbsp;"
    adatostarea(iadatostarea,2)=xagrC
   else
    adatostarea(iadatostarea,0)=rs("fechasplan")
    adatostarea(iadatostarea,3)=rs(8)&"&nbsp;"
    adatostarea(iadatostarea,4)=xagrCInc
    'respons.write "entro"
   end if
   
   
  iadatostarea=iadatostarea+1 


          END IF

     end if
                    
                    
                    xagrr = xagrd - xagrc

                    usrtskc =  usrtskc + xagrc
                    prytskc =  prytskc + xagrc
                    usrtskr =  usrtskr + xagrr
                    prytskr =  prytskr + xagrr

 ' prytsk  = 0
 ' prytskC = 0
'  prytskR = 0

       end if
      END IF

     END IF

    end if

    RS.MOVENEXT
    LOOP

     if usrtsk > 0 then

  '    celdas = ofv.GenCelda("","vt COLSPAN=2 ","","","","","","TOTAL " & USRDESC,"si")
'
  '    SINPUT = usrtsk 
 ''     celdas = celdas & ofv.GenCelda("","Ut","RIGHT","","","","",sInput,"si")
  '    SINPUT = usrtskc
  '    celdas = celdas & ofv.GenCelda("","Vt COLSPAN=2 ","RIGHT","","","","",sInput,"si")
  '    SINPUT = usrtskr
  ' '   celdas = celdas & ofv.GenCelda("","Ut","RIGHT","","","","",sInput,"si")
  '    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
'
  ''    celdas = ofv.GenCelda("","AA COLSPAN=4 style='border-top-style:ridge;border-top-width:2;border-top-color:#f7aa22;' ","CENTER","","","","","&nbsp;","si")
  '    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
'
  '    celdas = ofv.GenCelda("","vt COLSPAN=2 ","","","","","",USRDESC,"si")
'
  '    SINPUT = usrtsk
 ''     celdas = celdas & ofv.GenCelda("","Ut","RIGHT","","","","",sInput,"si")
  '    SINPUT = usrtskc
  '    celdas = celdas & ofv.GenCelda("","Ut COLSPAN=2 ","RIGHT","","","","",sInput,"si")
  '    SINPUT = usrtskr
  ''    celdas = celdas & ofv.GenCelda("","Ut","RIGHT","","","","",sInput,"si")
  '    filas2 = filas2 & ofv.GenRow("","","","","10","","",celdas,"si")
'
  '    celdas = ofv.GenCelda("","AA COLSPAN=4 style='border-top-style:ridge;border-top-width:2;border-top-color:#f7aa22;' ","CENTER","","","","","&nbsp;","si")
  '    filas2 = filas2 & ofv.GenRow("","","","","10","","",celdas,"si")

     end if

   end if

  '    rowspan= " rowspan=4 "
  '    celdas = ofv.GenCelda("","vt COLSPAN=2  ","","","","","","TOTALES DE " & XDESTN & "&nbsp;&nbsp;&nbsp;&nbsp;TAREAS:&nbsp;" & XCANT,"si")
'
  '    SINPUT = PRYtsk
 ''     celdas = celdas & ofv.GenCelda("","vt","RIGHT","","","","",sInput,"si")
  '    SINPUT = PRYtskc
  '    celdas = celdas & ofv.GenCelda("","vt COLSPAN=2  ","RIGHT","","","","",sInput,"si")
  '    SINPUT = PRYtskr
  ''    celdas = celdas & ofv.GenCelda("","vt","RIGHT","","","","",sInput,"si")
  '    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
'
  '    celdas = ofv.GenCelda("","AA COLSPAN=4 style='border-top-style:ridge;border-top-width:2;border-top-color:#f7aa22;' ","CENTER","","","","","&nbsp;","si")
  '    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")
'
  '    rowspan= " rowspan=4 "
  '    celdas = ofv.GenCelda("","vt COLSPAN=2  ","","","","","","TOTALES DE " & XDESTN & "&nbsp;&nbsp;&nbsp;&nbsp;TAREAS:&nbsp;" & XCANT,"si")
'
  '    SINPUT = PRYtsk
 ''     celdas = celdas & ofv.GenCelda("","vt","RIGHT","","","","",sInput,"si")
  '    SINPUT = PRYtskc
  '    celdas = celdas & ofv.GenCelda("","vt COLSPAN=2  ","RIGHT","","","","",sInput,"si")
  '    SINPUT = PRYtskr
  ''    celdas = celdas & ofv.GenCelda("","vt","RIGHT","","","","",sInput,"si")
  '    filas2 = filas2 & ofv.GenRow("","","","","10","","",celdas,"si")
'
  '    celdas = ofv.GenCelda("","AA COLSPAN=4 style='border-top-style:ridge;border-top-width:2;border-top-color:#f7aa22;' ","CENTER","","","","","&nbsp;","si")
  '    filas2 = filas2 & ofv.GenRow("","","","","10","","",celdas,"si")
'

  '  FilaS = ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")

  '  FilaS2 = ofv.GenTabla("","","","100%","","0","","0","0","","",filas2,"si")
'

  '  Celda = ofv.GenCelda("","","center","","","","",FilaS & "<BR>" & FilaS2,"si")
  '  FilaS = ofv.GenRow("","","","","","","",Celda,"si")

  '   sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
  '   sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
  '   sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
  '   sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
  '   sinput = sinput  & "none" & chr(34) & ";' "
  '   sinput = sinput  & "  href=REQsectoresBUS.asp?VASP=ES  "
  '   sinput = sinput  & "  TITLE='Afectar&nbsp;sector'  ><b>" & sImagenI3 & "&nbsp;Afectar&nbsp;sector</b></a>" 

 ''    sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34) & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34) & "none" & chr(34) & ";' class=tt href=REQsectoresBUS.asp?VASP=ES"
 ''    sinput = sinput  & " ><b>&nbsp;afectar&nbsp;sector&nbsp;</b></a>" 
  '   celdas = ofv.GenCelda("","","","","","","","&nbsp;","si")
  '   celdas = celdas & ofv.GenCelda("","vt","","10%","","","",sInput,"si")

  '   fila = ofv.GenRow("","","","","","","",celdas,"si")

  '   Fila = ofv.GenTabla("","","","100% border","","0","","0","0","","",fila,"si")

  '  Celda = ofv.GenCelda("","","center","","","","",Fila,"si")
 '   FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")
     tt1=ofv.GenTabla("","aa  ","","98%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & tt1 'ofv.GenTabla("","aa  ","","98%","","0","","0","0","","",filas,"si")
    
    filas=""
    celdas=ofv.GenCelda("","vt  align=center rowspan=2  style='border-style:solid;border-width:1;border-color:#F7AA22;'","","","","","","FECHA" ,"si")
    celdas=celdas & ofv.GenCelda("","vt  style='border-style:solid;border-width:1;border-color:#F7AA22;' align=center colspan=2 ","","","","","","TAREAS" ,"si")
    celdas=celdas & ofv.GenCelda("","vt align=center colspan=2 style='border-style:solid;border-width:1;border-color:#F7AA22;'","","","","","","TAREAS" ,"si")
    celdas=celdas & ofv.GenCelda("","vt rowspan=2 align=center  style='border-style:solid;border-width:1;border-color:#F7AA22;'","","","","","","TOTAL DE HORAS" ,"si")
    filas=ofv.GenRow("","","","","","","",Celdas,"si")
    
   ' celdas=ofv.GenCelda("","vt  ","","","","","","&nbsp;" ,"si")
    celdas= ofv.GenCelda("","vt  align=center style='border-style:solid;border-width:1;border-color:#F7AA22;'","","","","","","PROYECTO" ,"si")
    celdas=celdas & ofv.GenCelda("","vt  align=center style='border-style:solid;border-width:1;border-color:#F7AA22;'","","","","","","HORAS" ,"si")
    
    celdas=celdas & ofv.GenCelda("","vt  align=center style='border-style:solid;border-width:1;border-color:#F7AA22;'","","","","","","DE INCIDENCIAS" ,"si")
    celdas=celdas & ofv.GenCelda("","vt  align=center style='border-style:solid;border-width:1;border-color:#F7AA22;'","","","","","","HORAS" ,"si")
    
   ' celdas=celdas & ofv.GenCelda("","vt  align=center  ","","","","","","&nbsp;","si")
    filas=filas & ofv.GenRow("","","","","","","",Celdas,"si")
    totalProy=0
    totalincidencia=0
    totalparcialPry=0
    
    
    fechaact=adatostarea(0,0)
    
    for i=0 to iadatostarea-1
     
     celdas="" 
     if adatostarea(i,0)=fechaact then
      totalparcialPry=totalparcialPry + clng("0" & adatostarea(i,2)) + clng("0" & adatostarea(i,4))
     
     
     else
      'totalparcialPry=totalparcialPry + clng("0" & adatostarea(i,2)) + clng("0" & adatostarea(i,4))
      celdas = ofv.GenCelda("","vt  style='border-style:solid;border-width:1;'  colspan=6 align=right  ","","","","","","<b><font color=#ff0000> " & totalparcialPry &"</font></b>","si")
      'celdas = ofv.GenCelda("","Ut  colspan=6 align=right  style='border-BOTTOM-style:solid;border-BOTTOM-width:1;border-BOTTOM-color:#EEEEEE;'","","","","","","&nbsp;","si")
      filas=filas& ofv.GenRow("","","","","","","",Celdas,"si")
      totalparcialPry=clng("0" & adatostarea(i,2)) + clng("0" & adatostarea(i,4))
      ultimacelda=ultimacelda + adatostarea(i,2) + adatostarea(i,4)
      fechaact=adatostarea(i,0)
     
      celdas=""
     end if
     
     celdas = celdas & ofv.GenCelda("","vt  style='border-style:solid;border-width:1;'","","","","","",fechadorx(adatostarea(i,0)) ,"si")
     celdas = celdas & ofv.GenCelda("","vt  style='border-style:solid;border-width:1;'","","","","","",adatostarea(i,1)&"&nbsp;","si")
     celdas = celdas & ofv.GenCelda("","vt style='border-style:solid;border-width:1;'","","","","","",adatostarea(i,2) &"&nbsp;","si")
     celdas = celdas & ofv.GenCelda("","vt  style='border-style:solid;border-width:1;'","","","","","",adatostarea(i,3) &"&nbsp;","si")
     celdas = celdas & ofv.GenCelda("","vt  style='border-style:solid;border-width:1;'","","","","","",adatostarea(i,4) &"&nbsp;","si")
     'celdas = celdas & ofv.GenCelda("","xxx  ","","","","","",adatostarea(i,2) + adatostarea(i,4),"si")
     
    
     totalProy=totalProy + adatostarea(i,2)
     totalincidencia=totalincidencia + adatostarea(i,4)
     filas=filas& ofv.GenRow("","","","","","","",Celdas,"si")
    next
    celdas =  ofv.GenCelda("","xxx  ","","","","","","&nbsp;","si")
    filas=filas& ofv.GenRow("","","","","","","",Celdas,"si")
    
    celdas= ofv.GenCelda("","vt  align=left colspan=3 ","","","","","","TOTAL POR PROYECTOS =" & totalProy,"si")
    celdas=celdas & ofv.GenCelda("","vt  align=left colspan=3 ","","","","","","TOTAL POR INCIDENCIAS=" & totalincidencia,"si")
    
    filas=filas& ofv.GenRow("","","","","","","",Celdas,"si")
    
    celdas= ofv.GenCelda("","vt  align=left colspan=6 ","","","","","","TOTAL DE HORAS =" & totalProy + totalincidencia,"si")
    filas=filas& ofv.GenRow("","xxx ","","","","","",Celdas,"si")
    tt=ofv.GenTabla("","aa   ","","98% ","","0","","0","0","","",filas,"si")
    
    sHTML = sHTML & tt 'ofv.GenTabla("","aa   ","","98%","","0","","0","0","","",filas,"si")
  
    sHTML = sHTML & "</center><br><br>"





 




  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
  call ofv.cerrarconn(cn2)
  call ofv.cerrarconn(cn3)


  sHTML = sHTML  & "</body></html>"
end if

if xls="S" then
  response.write tt1 & tt
else
  response.write sHTML
end if

'response.write sHTML

Function Fechador(wkfecha)


    
      fec = cdate(wkfecha + clng(cdate("01-01-1900")))
      ndia = ofv.cdia(Weekday(fec))
      nMes = ofv.cmes(Month(fec))
    
      Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)


End Function


Function Fechadorx(wkfecha)


      
      fec = cdate(wkfecha + clng(cdate("01-01-1900")))
      ndia = day(Weekday(fec))
      
      if ndia < 10 then ndia= "0" & ndia
      
      
      nMes = month(Month(fec))
    
      if nmes <10 then nmes="0" & nmes
      
      Fechadorx=  ndia & "/" & nMes & "/" & Year(fec)


End Function

%>