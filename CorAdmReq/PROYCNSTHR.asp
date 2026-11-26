<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader(replace(session("FormN"),"_"," "))



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

 XSET = request.querystring("XSET")
 XOPT = request.querystring("XOPT")
 sopcion = request.querystring("opcion")
 sRESP = request.querystring("RESP")

IF SESSION("PRCF") <> "S" THEN
   XXXX = OWKF.GETPRYCNFG()
END IF


 IF XOPT = "" THEN XOPT = "XRESP"
 IF XSET = "" THEN XSET = "XPT"


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

       XDESTN="GLOBALES"

      StrSql = "Select * From PROYTSKASGN where PESTADO <> 13 "
      StrSql = StrSql & " AND ((VPLAN = PVPLAN ) "
    ''  StrSql = StrSql & " AND ((VPLAN = PVPLAN and CODIGO <> '" & SESSION("NCLIENTE") & "') "
    ''  StrSql = StrSql & " OR (VPLAN = '" & SESSION("VPLAN") & "' AND CODIGO = '" & SESSION("NCLIENTE") & "')) "
      StrSql = StrSql & " ) "
      StrSql = StrSql & " order by USRRESP, FINICIO, CODIGO, orden, TSKORD "


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

      StrSql = "Select * From PROYTSKASGN where PESTADO <> 13 AND GERENCIA = " & sRESP
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

      StrSql = "Select * From PROYTSKASGN where PESTADO <> 13 AND SECTOR = " & sRESP
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

      StrSql = "Select * From PROYTSKASGN where PESTADO <> 13 AND USRRESP = '" & sRESP  & "' "
      StrSql = StrSql & " AND ((VPLAN = PVPLAN ) "
    ''  StrSql = StrSql & " AND ((VPLAN = PVPLAN and CODIGO <> '" & SESSION("NCLIENTE") & "') "
    ''  StrSql = StrSql & " OR (VPLAN = '" & SESSION("VPLAN") & "' AND CODIGO = '" & SESSION("NCLIENTE") & "')) "
      StrSql = StrSql & " ) "
      StrSql = StrSql & " order by GERENCIA, SECTOR, USRRESP, FINICIO, CODIGO, orden, TSKORD "

END SELECT


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




''  StrSql = "Select * From PROYETPTSK where codigo = '" & session("NCliente") & "' AND VPLAN = " & SESSION("VPLAN") & " ORDER BY USRRESP,ORDEN,TSKORD "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then


 

    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"

  celdas = ofv.GenCelda("","cc colspan=4","","","","","",sImagen & "&nbsp;<BIG><B>TOTAL HORAS POR RESPONSABLE - " & XDESTN & "</B></BIG>","si")
  filas  = ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","AA COLSPAN=4","CENTER","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","vti ROWSPAN=2 style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","","","","","","RESPONSABLE","si")
  celdas = celdas & ofv.GenCelda("","vt COLSPAN=3 style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","CENTER","","","","","HORAS","si")
  filas = filas & ofv.GenRow("","aaa bgcolor='#ddddff' ","","","10","","",celdas,"si")

  celdas = ""
  celdas = celdas & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","","PLANIFICADAS","si")
  celdas = celdas & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","","CONSUMIDAS","si")
  celdas = celdas & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","","RESTANTES","si")
  filas = filas & ofv.GenRow("","aaa bgcolor='#ddddff'","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","AA COLSPAN=4   ","CENTER","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  tskant = "no"
  USRANT  = "XX"
  usrtsk  = 0
  prytsk  = 0
  prytskC = 0
  prytskR = 0

    xTrColor = "#FFFFFF" 

  DO UNTIL RS.EOF

  XVA = "S"


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

  IF SESSION("FECOPC") = "2" THEN

     SXFSTARTD = OWKF.GETFENDV(rs(20), rs(5))

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





     XFECDES = FSREAL

     XFECHAS = FEREAL

      
     IF CLNG(SESSION("FECDES")) > CLNG(XFECHAS) OR CLNG(SESSION("FECHAS")) < CLNG(XFECDES) THEN
        XVA = "N"
     END IF

   


  END IF 

  IF XVA = "S" THEN

     if   tskant <> rs(2) then
          tskant = rs(2) 

     IF USRANT <> RS(13) THEN

     if usrtsk > 0 then

       wkstyle = ""

      celdas = ofv.GenCelda("","vt style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#dddddd;'  ","","","","","",USRDESC,"si")

      SINPUT = GETGLBHMS(usrtsk)
      celdas = celdas & ofv.GenCelda("","ut style='border-style:solid;border-width:1;border-color:#dddddd;' ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = GETGLBHMS(usrtskc)
      celdas = celdas & ofv.GenCelda("","ut style='border-style:solid;border-width:1;border-color:#dddddd;' ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = GETGLBHMS(usrtskr)
      if usrtsk < usrtskc then
         wkstyle = "color:#990000;font-weight:bold;"   
         SINPUT = "-" & SINPUT
      end if
      celdas = celdas & ofv.GenCelda("","ut style='border-style:solid;border-width:1;border-color:#dddddd;" & wkstyle & "' ","RIGHT","","","","","&nbsp;" & sInput,"si")
      filas = filas & ofv.GenRow("","aaa  BGCOLOR='" & xTrColor & "' ","","","15","","",celdas,"si")

 '     celdas = ofv.GenCelda("","AA COLSPAN=4 style='border-top-style:ridge;border-top-width:2;border-top-color:#f7aa22;' ","CENTER","","","","","&nbsp;","si")
 '     filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

	 if xTrColor = "#FFFFFF" then
	    xTrColor = "#F4F4F4" 
	 ELSE
	    xTrColor = "#FFFFFF" 
	 END IF

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

        FILAS1 = ""
     END IF


'''    xagrd = round((xagrd),2)



        USRTMP = RS(15) & ""
        USRTMPU = RS(16)
        IF USRTMP = "" THEN
           USRTMP = 0
           USRTMPU = 0
        END IF
        USRTMP = CINT(USRTMP)
        IF USRTMP = 0 THEN
           USRTMP = 0 
           USRTMPU = 0
        END IF


        SELECT CASE USRTMPU
               CASE 4
                    xagrd = USRTMP 
               CASE 1
                    xagrd = USRTMP * 60

               CASE 2

                    xagrd = (USRTMP * CLNG(SESSION("DPP"))) * 60
 

               CASE ELSE

                    xagrd = (USRTMP * CLNG(SESSION("DPP")) * CLNG(SESSION("SPP")) ) * 60

        END select         

        xagrd = xagrd * 60

                    usrtsk =  usrtsk + xagrd
                    prytsk =  prytsk + xagrd



                xagrc = 0
                xagrr = 0

   if rs(4) > 0 then



          PADV = 0 


          StrSql = "Select * From PROYTAREASADV where codigo = '" & rs(5) & "' AND TSK = '" & RS(20) & "' ORDER BY id desc "
          set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if NOT rsX.eof then

             PADV = RSX(5)
             xagrc = Clng((XAGRD * PADV / 100)) 
             xagrr = XAGRD - Clng((XAGRD * PADV / 100)) 

             WKPSCTMP = GETGLBHMSW(xagrc)

           ''  WKPSCTMPH = CINT(LEFT(WKPSCTMP,2))
           ''  WKPSCTMPM = CINT(MID(WKPSCTMP,4,2))

             WKPSCTMPS = CINT(RIGHT(WKPSCTMP,2))

             xagrc = xagrc - WKPSCTMPS


          END IF

     end if

                xagrr = xagrd - xagrc
            '''    xagrr = round((xagrd - xagrc),2)

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

    RS.MOVENEXT
    LOOP

     if usrtsk > 0 then

      wkstyle = "" 

      celdas = ofv.GenCelda("","vt style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#dddddd;' ","","","","","",USRDESC,"si")

      SINPUT = GETGLBHMS(usrtsk)
      celdas = celdas & ofv.GenCelda("","ut style='border-style:solid;border-width:1;border-color:#dddddd;' ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = GETGLBHMS(usrtskc)
      celdas = celdas & ofv.GenCelda("","ut style='border-style:solid;border-width:1;border-color:#dddddd;' ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = GETGLBHMS(usrtskr)
      if usrtsk < usrtskc then
         wkstyle = "color:#990000;font-weight:bold;"   
         SINPUT = "-" & SINPUT
      end if
      celdas = celdas & ofv.GenCelda("","ut style='border-style:solid;border-width:1;border-color:#dddddd;" & wkstyle & "' ","RIGHT","","","","","&nbsp;" & sInput,"si")

      filas = filas & ofv.GenRow("","aaa  BGCOLOR='" & xTrColor & "' ","","","15","","",celdas,"si")


     end if

   end if

      celdas = ofv.GenCelda("","AA COLSPAN=4 style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#f7aa22;' ","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      rowspan= " rowspan=4 "
      celdas = ofv.GenCelda("","vti","","","","","","TOTALES " & XDESTN,"si")

      wkstyle = "" 

      SINPUT = GETGLBHMS(PRYtsk)
      celdas = celdas & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = GETGLBHMS(PRYtskc)
      celdas = celdas & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = GETGLBHMS(PRYtskr)
      if PRYtsk < PRYtskc then 
         wkstyle = "color:#990000;font-weight:bold;"   
         SINPUT = "-" & SINPUT
      end if
      celdas = celdas & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#aaaaaa;" & wkstyle & "' ","RIGHT","","","","","&nbsp;" & sInput,"si")


      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","AA COLSPAN=4 style='border-top-style:solid;border-top-width:1;border-top-color:#f7aa22;' ","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")




    FilaS = ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")

    Celda = ofv.GenCelda("","","center","","","","",FilaS,"si")
    FilaS = ofv.GenRow("","","","","","","",Celda,"si")

     sinput = "<a class=cc style='text-decoration:none;background-color:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "#f7aa22" & chr(34) & ";' onmouseout='this.style.backgroundColor=" & chr(34)
     sinput = sinput  & "" & chr(34) & ";this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' "
     sinput = sinput  & "  href=REQsectoresBUS.asp?VASP=ES  "
     sinput = sinput  & "  TITLE='Afectar&nbsp;sector'  ><b>" & sImagenI3 & "&nbsp;Afectar&nbsp;sector</b></a>" 

 ''    sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34) & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34) & "none" & chr(34) & ";' class=tt href=REQsectoresBUS.asp?VASP=ES"
 ''    sinput = sinput  & " ><b>&nbsp;afectar&nbsp;sector&nbsp;</b></a>" 
     celdas = ofv.GenCelda("","","","","","","","&nbsp;","si")
     celdas = celdas & ofv.GenCelda("","vt","","10%","","","",sInput,"si")

     fila = ofv.GenRow("","","","","","","",celdas,"si")

     Fila = ofv.GenTabla("","","","100%","","0","","0","0","","",fila,"si")

    Celda = ofv.GenCelda("","","center","","","","",Fila,"si")
 '   FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aa ","","98%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"





 




  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
  call ofv.cerrarconn(cn2)
  call ofv.cerrarconn(cn3)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML



%>