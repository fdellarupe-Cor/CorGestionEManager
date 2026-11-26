<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


' FUNCION = " onload=" & chr(34)
' funcion = funcion & " if (window.name=='info') "
' funcion = funcion & " { document.all.UNO.style.display='block';} else  { document.all.UNO.style.display='none';}  " & chr(34)




  sHTML = ofv.MenuHeader("",funcion)

''  shtml = replace(shtml,"<body ","<body " & funcion)



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

 XSET    = request.querystring("XSET")
 XOPT    = request.querystring("XOPT")
 sopcion = request.querystring("opcion")
 sRESP   = request.querystring("RESP")


 SMES    = request.querystring("MES")
 SANIO   = request.querystring("ANIO")

'''      sfech = clng(dateserial(sanioc,smesc,28))


 IF SMES  = "" THEN
    SMES  = datepart("m",DATE())
    SANIO = datepart("yyyy",DATE())

 END IF




IF SESSION("PRCF") <> "S" THEN
   XXXX = OWKF.GETPRYCNFG()
END IF


 IF XOPT = "" THEN XOPT = "XRESP"
 IF XSET = "" THEN XSET = "XPT"




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


     sopcion = "VP"


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

    sImagenSO = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=3>5</FONT></B>"
    sImagenAN = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=3>3</FONT></B>"
    sImagenSI = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=3>4</FONT></B>"


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


      StrSql = "Select DISTINCT GERENCIA, SECTOR, USRRESP From PROYTSKASGN where "
      StrSql = StrSql & " ((VPLAN = PVPLAN and CODIGO <> '" & SESSION("NCLIENTE") & "') "
      StrSql = StrSql & " OR (VPLAN = '" & SESSION("VPLAN") & "' AND CODIGO = '" & SESSION("NCLIENTE") & "')) "
      StrSql = StrSql & " order by GERENCIA, SECTOR, USRRESP "

      StrSql2 = StrSql 

      StrSql = "Select * From PROYTSKASGN where "
      StrSql = StrSql & " ((VPLAN = PVPLAN and CODIGO <> '" & SESSION("NCLIENTE") & "') "
      StrSql = StrSql & " OR (VPLAN = '" & SESSION("VPLAN") & "' AND CODIGO = '" & SESSION("NCLIENTE") & "')) "
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

      StrSql = "Select DISTINCT GERENCIA, SECTOR, USRRESP From PROYTSKASGN where  GERENCIA = " & sRESP
      StrSql = StrSql & " AND ((VPLAN = PVPLAN and CODIGO <> '" & SESSION("NCLIENTE") & "') "
      StrSql = StrSql & " OR (VPLAN = '" & SESSION("VPLAN") & "' AND CODIGO = '" & SESSION("NCLIENTE") & "')) "
      StrSql = StrSql & " order by GERENCIA, SECTOR, USRRESP "

      StrSql2 = StrSql 

      StrSql = "Select * From PROYTSKASGN where  GERENCIA = " & sRESP
      StrSql = StrSql & " AND ((VPLAN = PVPLAN and CODIGO <> '" & SESSION("NCLIENTE") & "') "
      StrSql = StrSql & " OR (VPLAN = '" & SESSION("VPLAN") & "' AND CODIGO = '" & SESSION("NCLIENTE") & "')) "
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

      StrSql = "Select DISTINCT GERENCIA, SECTOR, USRRESP From PROYTSKASGN where  SECTOR = " & sRESP
      StrSql = StrSql & " AND ((VPLAN = PVPLAN and CODIGO <> '" & SESSION("NCLIENTE") & "') "
      StrSql = StrSql & " OR (VPLAN = '" & SESSION("VPLAN") & "' AND CODIGO = '" & SESSION("NCLIENTE") & "')) "
      StrSql = StrSql & " order by GERENCIA, SECTOR, USRRESP "

      StrSql2 = StrSql 

      StrSql = "Select * From PROYTSKASGN where  SECTOR = " & sRESP
      StrSql = StrSql & " AND ((VPLAN = PVPLAN and CODIGO <> '" & SESSION("NCLIENTE") & "') "
      StrSql = StrSql & " OR (VPLAN = '" & SESSION("VPLAN") & "' AND CODIGO = '" & SESSION("NCLIENTE") & "')) "
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

      StrSql = "Select DISTINCT GERENCIA, SECTOR, USRRESP From PROYTSKASGN where USRRESP = '" & sRESP  & "' "
      StrSql = StrSql & " AND ((VPLAN = PVPLAN and CODIGO <> '" & SESSION("NCLIENTE") & "') "
      StrSql = StrSql & " OR (VPLAN = '" & SESSION("VPLAN") & "' AND CODIGO = '" & SESSION("NCLIENTE") & "')) "
      StrSql = StrSql & " order by GERENCIA, SECTOR, USRRESP "

      StrSql2 = StrSql 

      StrSql = "Select * From PROYTSKASGN where USRRESP = '" & sRESP  & "' "
      StrSql = StrSql & " AND ((VPLAN = PVPLAN and CODIGO <> '" & SESSION("NCLIENTE") & "') "
      StrSql = StrSql & " OR (VPLAN = '" & SESSION("VPLAN") & "' AND CODIGO = '" & SESSION("NCLIENTE") & "')) "
      StrSql = StrSql & " order by GERENCIA, SECTOR, USRRESP, FINICIO, CODIGO, orden, TSKORD "

END SELECT










    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"



    XFSTART = CLNG(DATE()) - clng(cdate("01-01-1900")) 
    XFCURR  = XFSTART
    XFHOY   = XFSTART
    XFREAL  = XFSTART

XXX = OWKF.ADMFERIADOS("C",0)








  ''    XFCURR = XFSTART

         dFecha = CDATE(CLNG(XFSTART) + clng(cdate("01-01-1900")))





    THXD = CINT(SESSION("DPP"))


    DIM AUSR(31,400,2,12)

    SFDES = clng(dateserial(sanio,smes,1))
    SFHAS = clng(dateserial(sanio,smes,28))
    
  '  RESPONSE.WRITE SFDES & "-" &   SFHAS & "-" &  cdate(clng(SFDES)) & " " &  cdate(clng(SFHAS)) & "<BR>"

    yy = 0

    for ii = 1 to 3
        sfecha = SFHAS + ii
        sfecha = cdate(clng(sfecha))
        smess = datepart("m",sfecha)
        if cint(smes) <> cint(smess) then
           exit for
        else
           yy = ii
        end if

    next  

    SFHAS = SFHAS + YY

    SFANT = SFDES - 1
    SFNEX = SFHAS + 1


    sfecha = cdate(clng(SFANT))
    SMESA  = datepart("m",sfecha)
    SANIOA = datepart("yyyy",sfecha)

    sfecha = cdate(clng(SFNEX))
    SMESS  = datepart("m",sfecha)
    SANIOS = datepart("yyyy",sfecha)


     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=PROYASGNDET.asp?XSET=" & XSET & "&XOPT=" & XOPT & "&RESP=" & SRESP & "&MES=" & SMESA & "&ANIO=" & SANIOA
     sinput = sinput  & "  TITLE='ANTERIOR'  ><b>" & sImagenAN & "</b></a>" 

     sinputA = sinput

     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=PROYASGNDET.asp?XSET=" & XSET & "&XOPT=" & XOPT & "&RESP=" & SRESP & "&MES=" & SMESS & "&ANIO=" & SANIOS
     sinput = sinput  & "  TITLE='SIGUIENTE'  ><b>" & sImagenSI & "</b></a>" 

     sinputS = sinput


    TDM   = 28 + YY


    FOR IDAY=1 TO TDM
        SFWKT = SFANT + IDAY

        AUSR(IDAY,0,0,0) = SFWKT
 
        WKTH = OWKF.GETDHAB(SFWKT - CLNG(CDATE("01-01-1900")),0)

        IF WKTH = 0 THEN
           AUSR(IDAY,0,1,0) = 1
        ELSE
           AUSR(IDAY,0,1,0) = 0
        END IF
 ''RESPONSE.WRITE SFWKT & "-" &   cdate(clng(SFWKT)) & "-" &  WKTH & "<BR>"


    NEXT 

' RESPONSE.WRITE SFDES & "-" &   SFHAS & "-" &  cdate(clng(SFDES)) & " " &  cdate(clng(SFHAS)) &   SFANT & "-" &  cdate(clng(SFANT))  & "<BR>"
 
  IUSR = 0
  TUSR = 0 

''  response.write strsql2
  set rs1 = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
  if not  rs1.eof then
     do while not rs1.eof

''' DISTINCT GERENCIA, SECTOR, USRRESP

        IUSR = IUSR + 1
        TUSR = IUSR

        AUSR(0,IUSR,0,0) = RS1(0)
        AUSR(0,IUSR,1,0) = RS1(1)
        AUSR(0,IUSR,2,0) = RS1(2)


        FOR IDAY=1 TO TDM
            AUSR(IDAY,IUSR,1,0) = ROUND(0,2) ' HS. PLANIFICADAS
            AUSR(IDAY,IUSR,2,0) = ROUND(0,2) ' HS. CONSUMIDAS
        NEXT 


     RS1.MOVENEXT
     LOOP

  END IF 


        IUSR = 1







''''''''''''''''

''  StrSql = "Select * From PROYTSKASGN where ESTADO < 2 AND USRRESP = '" & sRESP  & "'  AND ((VPLAN = PVPLAN and CODIGO <> '" & SESSION("NCLIENTE") & "') OR (VPLAN = '" & SESSION("VPLAN") & "' AND CODIGO = '" & SESSION("NCLIENTE") & "'))  order by GERENCIA, SECTOR, USRRESP, FINICIO, CODIGO, orden, TSKORD "
''  response.write strsql
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not  rs1.eof then
     do while not rs1.eof






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


       IF AUSR(0,IUSR,2,0) <> XUSR THEN
          FOR IUSR=IUSR TO TUSR
              IF AUSR(0,IUSR,2,0) = XUSR THEN
                 EXIT FOR 
              END IF
          NEXT  
       END IF 


''        IUSR = IUSR + 1
''        TUSR = IUSR
''        AUSR(0,IUSR,0,0) = RS1(0)
''        AUSR(0,IUSR,1,0) = RS1(1)
''        AUSR(0,IUSR,2,0) = RS1(2)





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

   

 ''   SFDES = clng(dateserial(sanio,smes,1))
 ''   SFHAS = clng(dateserial(sanio,smes,28))

    FSPLAN = FSPLAN + clng(cdate("01-01-1900")) 
    FEPLAN = FEPLAN + clng(cdate("01-01-1900")) 


    IF (FSPLAN <= SFHAS) AND (FEPLAN >= SFDES) THEN

       IF FSPLAN > SFDES THEN
          SDDES = FSPLAN - SFANT
       ELSE
          SDDES = 1
       END IF 

       IF FEPLAN < SFHAS THEN
          SDHAS = TDM - (SFHAS - FEPLAN)
       ELSE
          SDHAS = TDM
       END IF 

'' ROUND(0,2)

      XWVDUR = ROUND(100,2)

      SELECT CASE RS1(16)

             CASE 4

                  XVDUR = rs1(15) / 60

             CASE 1

                  XVDUR = rs1(15)

             CASE 2
                  XVDUR = rs1(15) * CLNG(SESSION("DPP"))


             CASE 3 
                  XVDUR = rs1(15) * CLNG(SESSION("SPP")) * CLNG(SESSION("DPP"))

      END SELECT

      XWDUR = XVDUR


   IF rs1(24) > 0 THEN 



      IF rs1(22) > 0 THEN
         XTDUR = rs1(22)
      ELSE
         XTDUR = rs1(15)
      END IF
   
      IF rs1(23) > 0 THEN
         XUDUR = rs1(23)
      ELSE
         XUDUR = rs1(16)
      END IF



      SELECT CASE XUDUR

             CASE 4

                  XWDUR = XTDUR / 60

             CASE 1

                  XWDUR = XTDUR

             CASE 2
                  XWDUR = XTDUR * CLNG(SESSION("DPP"))


             CASE 3 
                  XWDUR = XTDUR * CLNG(SESSION("SPP")) * CLNG(SESSION("DPP"))

      END SELECT

      XWVDUR = ROUND((XVDUR / XWDUR * 100),2)

    END IF


''      XWDUR = XVDUR
''      XWVDUR = ROUND((XVDUR / XWDUR * 100),2)


        IF XWDUR < THXD THEN
           AUSR(SDDES,IUSR,1,0) = ROUND( AUSR(SDDES,IUSR,1,0) + XWDUR,2)

        ELSE

            FOR IDAY=SDDES TO SDHAS
                IF XFFCH > 0 THEN
                   IF  AUSR(IDAY,0,1,0)    = 1 THEN
                       AUSR(IDAY,IUSR,1,0) = ROUND( AUSR(IDAY,IUSR,1,0) + (THXD * XWVDUR / 100),2)
                   END IF
                ELSE
                   AUSR(IDAY,IUSR,1,0)     = ROUND( AUSR(IDAY,IUSR,1,0) + (THXD * XWVDUR / 100),2)
                END IF
            NEXT

        END IF

      IF TSKST = "2" THEN

        IF XWDUR < THXD THEN
           AUSR(SDDES,IUSR,2,0) = ROUND( AUSR(SDDES,IUSR,2,0) + XWDUR,2)

        ELSE

            FOR IDAY=SDDES TO SDHAS
                IF XFFCH > 0 THEN
                   IF  AUSR(IDAY,0,1,0)    = 1 THEN
                       AUSR(IDAY,IUSR,2,0) = ROUND( AUSR(IDAY,IUSR,2,0) + (THXD * XWVDUR / 100),2)
                   END IF
                ELSE
                   AUSR(IDAY,IUSR,2,0)     = ROUND( AUSR(IDAY,IUSR,2,0) + (THXD * XWVDUR / 100),2)
                END IF
            NEXT

        END IF

      END IF 


    END IF

    rs1.movenext
    loop
    call ofv.cerrarconsulta(rs1)
  end if

       dFecha = dateserial(sanio,smes,1)
       nMes   = ofv.cmes(Month(dFecha))
    
       xfMES = UCASE(nMes) & "&nbsp;" & Year(dFecha)


      celdas = ofv.GenCelda("","cc colspan=4","","","","","",sImagen & "&nbsp;<BIG><B>DETALLE DE ASIGNACION DE " & XDESTN & "&nbsp;&nbsp;&nbsp;&nbsp; AL " & ofv.fechador & "</B></BIG>","si")
      filas  = ofv.GenRow("","","","","10","","",celdas,"si")


      celdas = ofv.GenCelda("","VT COLSPAN=4","CENTER","","","","",SINPUTA & "&nbsp;<b>" & xfMES & "</b>&nbsp;" & SINPUTS,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      FilaS1 = ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")


  
  sfecha = cdate(clng(SFDES))
  YFEBNT = DatePart("W",sfecha,2)
  AAGR   = YFEBNT - 1
  
  sfecha2 = cdate(clng(SFHAS))
  YFEBNT = DatePart("W",sfecha2,2)
  DAGR   = 7 - YFEBNT



  TDAY = TDM + AAGR + DAGR

  TSEM = TDAY / 7

    fshow = SFDES - AAGR

 '   RESPONSE.WRITE sfecha & "-" &   sfecha2 & "-" &  AAGR & "-" &  DAGR & "-" & fshow & "-" &  TDAY & "-" &  TSEM & "<BR>"

  FilaS =  ""
    
    Celdas = ofv.GenCelda("","abc STYLE='background-color:#f7aa22;font-family:verdana;font-size:7pt;color:#ffffff;'  rowspan=2 ","center","","","","","&nbsp;<b>RESPONSABLES</b>&nbsp;","si")



    FOR ISEM = 1 TO TSEM
       celdas = celdas & ofv.GenCelda("","aaa STYLE='border-style:solid;border-width:1;border-color:#eeeeee;background-color:#f7aa22;font-family:verdana;font-size:7pt;color:#ffffff;' colspan=7 " ,"CENTER","","","","","SEMANA " & ISEM,"si")

    NEXT 


    FilaS =  filaS & ofv.GenRow("","","","","","","",Celdas,"si")

       celdas = ""

    ICOL = 0
    IDIA = 0

    FOR ISEM = 1 TO TSEM
       for iday = 1 to 7



       ICOL = ICOL + 1 
        xid = "CC" & ICOL

       IF ICOL > AAGR AND ICOL < (CINT(TDM) + CINT(AAGR) + 1) THEN
          IDIA = IDIA + 1
       ELSE  
          IDIA = 0
       END IF

       dFecha = cdate(fshow)
       DFDIA = DatePart("W",dFecha,2)
       ndia = ofv.cdia(Weekday(dFecha))
       nMes = ofv.cmes(Month(dFecha))

       xbrd2 = ""

    
       xfshow = "  " & ndia & ", " & Day(dFecha) & " de " & nMes & " de " & Year(dFecha)

''alert(document.all." & xid & ".offsetLeft);document.all.XX.style.pixelLeft=document.all." & xid & ".offsetLeft;alert(document.all.XX.style.Left);


       xlink = "<A  style='color:#ffffff;' onmouseover='this.style.textDecoration=" & chr(34)
       xlink = xlink & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
       xlink = xlink & "none" & chr(34) & ";' "
       xlink = xlink & " href='javascript:void(0);' " 
       xlink = xlink & " title='" & xfshow & "' ><center>"
       xlink = xlink & IDIA
       xbrd = "solid"
       xlink = xlink & "</center></A>"

       xlinf = ""

       IF IDIA = 0 THEN
          celdas = celdas & ofv.GenCelda(xid,"aaa " & xlinf & " STYLE='border-style:" & xbrd & xbrd2 & ";border-width:1;border-color:#eeeeee;background-color:#CCCCCC;font-family:verdana;font-size:6pt;color:#000000;' " ,"CENTER","2","","","","&nbsp;","si")
       ELSE
          IF AUSR(IDIA,0,1,0) = 1 THEN
             celdas = celdas & ofv.GenCelda(xid,"aaa " & xlinf & " STYLE='border-style:" & xbrd & xbrd2 & ";border-width:1;border-color:#eeeeee;background-color:#6699cc;font-family:verdana;font-size:6pt;color:#ffffff;' " ,"CENTER","2","","","",xlink,"si")
          ELSE
             celdas = celdas & ofv.GenCelda(xid,"aaa " & xlinf & " STYLE='border-style:" & xbrd & xbrd2 & ";border-width:1;border-color:#eeeeee;background-color:#F7AA22;font-family:verdana;font-size:6pt;color:#ffffff;' " ,"CENTER","2","","","",xlink,"si")
          END IF 

       END IF 
 

       fshow = fshow + 1 
       next
    NEXT 


    FilaS =  filaS & ofv.GenRow("","","","","","","",Celdas,"si")









''''''''''''''''''''''''''''''''''''''''''''

  for IUSR=1 to TUSR


''        IUSR = IUSR + 1
''        TUSR = IUSR
''        AUSR(0,IUSR,0,0) = RS1(0)
''        AUSR(0,IUSR,1,0) = RS1(1)
''        AUSR(0,IUSR,2,0) = RS1(2)


       XUSR   = AUSR(0,IUSR,2,0)
       XSCT   = AUSR(0,IUSR,1,0)
       XGRN   = AUSR(0,IUSR,0,0)


    if XUSR = "NO" then
       sniv = "SIN AFECTAR"
    ELSE
      strsql = "Select * From USUARIOS where ID = '" & XUSR & "' "
      set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
      if not rsniv.eof then sniv = rsniv(1)
      call ofv.cerrarconsulta(rsniv)
    END IF

       XUSRN   = sniv

     StrSql = "Select * From ECO where ECO = " & XSCT
     set rsX2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rsX2.eof then
        sniv = rsX2(1)
     ELSE
        sniv = "NO DETERMINADO"
     END IF
      call ofv.cerrarconsulta(rsX2)

       XSCTN   = sniv

     StrSql = "Select * From ECO where ECO = " & XGRN
     set rsX2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
     if NOT rsX2.eof then
        sniv = rsX2(1)
     ELSE
        sniv = "NO DETERMINADO"
     END IF
      call ofv.cerrarconsulta(rsX2)

       XGRNN   = sniv


     RSPAN  = THXD + 2

     CLSPAN = TDAY 






     celdas = ofv.GenCelda("","VT bgcolor=#ffffff  ROWSPAN=" & RSPAN,"","30%","","","",XUSRN & "<BR>" & XSCTN & "<BR>" & XGRNN,"si")
     celdas = celdas & ofv.GenCelda("","aaa bgcolor=#ffffff STYLE='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#f7aa22;FONT-FAMILY:ARIAL;FONT-SIZE:4PT;' COLSPAN=" & CLSPAN,"right","","","","","&nbsp;","si")

     FilaS =  filaS & ofv.GenRow("","","","","","","",Celdas,"si")




     WARM = ""

     FOR IDAY=1 TO TDM

'''     AUSR(IDAY,IUSR,1,0)     = ROUND( AUSR(IDAY,IUSR,1,0) + (THXD * XWVDUR),2)

        TDS = THXD
        TPL = AUSR(IDAY,IUSR,1,0) 
        TCS = AUSR(IDAY,IUSR,2,0) 

        TUS = CINT(TPL) - CINT(TCS) 

        IF TUS < 1 THEN
           TUS = CINT(TCS)
           TPL = 0 
           TCS = CINT(TCS)
        ELSE
           TUS = CINT(TPL)
           TPL = CINT(TPL) - CINT(TCS) 
           TCS = CINT(TCS)
        END IF

        IF CINT(TUS) > TDS THEN
           TSP = 1
           TDS = 0
           TUS = THXD
           IF TPL > 0 THEN
 '             TPL = TPL - 1
               TPL = THXD
           ELSE
  '            TCS = TCS - 1
               TCS = THXD
           END IF  
        ELSE  
           TSP = 0
           TDS = TDS - CINT(TUS)
           TUS = CINT(TUS)
        END IF

        IHORA = 1
        IVOY  = 1

     '   RESPONSE.WRITE IDAY & "-" &   CINT(TPL) & " " &  CINT(TDS) & " " &  CINT(TCS) & "<BR>"

       ''' STILO = " STYLE='border-style:SOLID;border-width:1;border-color:#CCCCCC;' "
        STILO = " STYLE='border-style:NONE;FONT-FAMILY:ARIAL;FONT-SIZE:4PT;' "

        IF TSP = 0 THEN
            AUSR(IDAY,IUSR,0,IHORA)     = AUSR(IDAY,IUSR,0,IHORA) & ofv.GenCelda("","aaa  " & STILO,"","","","","","&nbsp;","si")
        ELSE
            AUSR(IDAY,IUSR,0,IHORA)     = AUSR(IDAY,IUSR,0,IHORA) & ofv.GenCelda("","aaa " & STILO,"","","","","",sImagenSO,"si")
        END IF

        FOR IVOY=1 TO TDS
            IHORA = IHORA + 1
            IF AUSR(IDAY,0,1,0) = 1 THEN
               AUSR(IDAY,IUSR,0,IHORA)     = AUSR(IDAY,IUSR,0,IHORA) & ofv.GenCelda("","aaa bgcolor=#FF7766 " & STILO,"","","","","","&nbsp;","si")
            ELSE
               AUSR(IDAY,IUSR,0,IHORA)     = AUSR(IDAY,IUSR,0,IHORA) & ofv.GenCelda("","aaa  " & STILO,"","","","","","&nbsp;","si")
            END IF 
        NEXT
        
        FOR IVOY=1 TO TPL
            IHORA = IHORA + 1
            AUSR(IDAY,IUSR,0,IHORA)     = AUSR(IDAY,IUSR,0,IHORA) & ofv.GenCelda("","aaa bgcolor=#99AACC " & STILO,"","","","","","&nbsp;","si")
        NEXT

        FOR IVOY=1 TO TCS
            IHORA = IHORA + 1
            AUSR(IDAY,IUSR,0,IHORA)     = AUSR(IDAY,IUSR,0,IHORA) & ofv.GenCelda("","aaa bgcolor=#44AA22 " & STILO,"","","","","","&nbsp;","si")
        NEXT

     NEXT 


     FOR IHORA=1 TO THXD + 1


        IF AAGR > 0 THEN
           CLSAGR = " COLSPAN=" & AAGR
           celdas = ofv.GenCelda("","AAA " & STILO & CLSAGR,"","","","","","&nbsp;","si")
        ELSE
           celdas = ""
        END IF 

     FOR IDAY=1 TO TDM

        celdas = celdas & AUSR(IDAY,IUSR,0,IHORA)

     NEXT


        IF DAGR > 0 THEN
           CLSAGR = " COLSPAN=" & DAGR
           celdas = celdas & ofv.GenCelda("","AAA " & STILO & CLSAGR,"","","","","","&nbsp;","si")
        END IF 


        FilaS =  filaS & ofv.GenRow("","","","","","","",Celdas,"si")


     NEXT






  NEXT

















    FilaS2 = ofv.GenTabla("","aa  style='border-style:ridge;border-width:2;border-color:#f7aa22;' ","","100%","","0","","0","0","","",filas,"si")


    Celda = ofv.GenCelda("","AAA COLSPAN=2","center","","","","",FilaS1 & FilaS2,"si")
    FilaS = ofv.GenRow("","","","","","","",Celda,"si")

     cnbot = "<input type=submit class=bt value='GANTT' >"
     Cons = ofv.gencelda("","","","","1%","","",cnbot,"si")
     sAccion = "PROYASGNgraf.asp?XSET=" & XSET & "&XOPT=" & XOPT & "&RESP=" & SRESP
     Cons = ofv.genform("cons","",sAccion,"_blank","post","",Cons,"si")
     celdas = Cons

      celdas = celdas & ofv.GenCelda("","","CENTER","","80%%","","","&nbsp;","si")

    FilaS = FilaS & ofv.GenRow("","","","","","","",CeldaS,"si")

    sHTML = sHTML & ofv.GenTabla("","aa style='filter:shadow(color=#cccccc,direction=125);' ","","90%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"





'''''''''''''''












 







  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML



%>