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



      StrSql = "Select * From PROYTSKASGN where  "
      StrSql = StrSql & "  ((VPLAN = PVPLAN ) "
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




      StrSql = "Select * From PROYTSKASGN where  GERENCIA = " & sRESP
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



      StrSql = "Select * From PROYTSKASGN where SECTOR = " & sRESP
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



      StrSql = "Select * From PROYTSKASGN where  USRRESP = '" & sRESP  & "' "
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


    DIM ATIPO(400,400,3)
    DIM WTFCH(6)

  ITIPO = 0
  TTIPO = 0 

  StrSql2 = "Select * From PROYTIPOTSK ORDER BY ID "
  set rs = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
  DO UNTIL RS.EOF

     ITIPO          = RS(0)
     ATIPO(ITIPO,0,0) = RS(1)


  RS.MOVENEXT
  LOOP

  TTIPO = ITIPO

  IUSR = 1
  ISCT = 1
  IGRN = 1
  IGLB = 1

IF SESSION("PRCF") <> "S" THEN
   XXXX = OWKF.GETPRYCNFG()
END IF



'''    IF CLNG(SESSION("FECDES")) > CLNG(XFECHAS) OR CLNG(SESSION("FECHAS")) < CLNG(XFECDES) THEN

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


   xcdiast = xcdiast & xcdias & " (CORRIDOS " &  xcdiasC & ") "


END IF
END IF

IF SESSION("FECOPC") = "2" THEN

   XPERIOD = "DESDE <FONT COLOR=#993333>" & UCASE(Fechador(xfdes)) & "</FONT> HASTA <FONT COLOR=#993333>" & UCASE(Fechador(xfHAS)) & "</FONT>"



ELSE

   XPERIOD = "INVENTARIO COMPLETO"

END IF


''  StrSql = "Select * From PROYETPTSK where codigo = '" & session("NCliente") & "' AND VPLAN = " & SESSION("VPLAN") & " ORDER BY USRRESP,ORDEN,TSKORD "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then


 

    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"

  celdas = ofv.GenCelda("","cc colspan=6","","","","","",sImagen & "&nbsp;<BIG><B>DISTRIBUCION POR TIPO DE TAREA POR RESPONSABLE - " & XDESTN & "</B></BIG>","si")
  filas  = ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","AA COLSPAN=6","CENTER","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","vtI COLSPAN=6 ","","","","","","PERIODO:&nbsp;" & XPERIOD,"si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","vtI COLSPAN=6 ","","","","","","DIAS:&nbsp;" & xcdiast,"si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","AA COLSPAN=6","CENTER","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")




  celdas = ofv.GenCelda("","vtI ROWSPAN=2 style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","","","","","","RESPONSABLE","si")
  celdas = celdas & ofv.GenCelda("","vt COLSPAN=3 style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","CENTER","","","","","TIEMPO INSUMIDO","si")
  celdas = celdas & ofv.GenCelda("","vt COLSPAN=2 style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","CENTER","","","","","PORCENTAJE","si")
  filas = filas & ofv.GenRow("","aaa bgcolor='#ddddff'","","","10","","",celdas,"si")

  celdas = ""
  celdas = celdas & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","","TAREAS","si")
  celdas = celdas & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","","IMPUTABLES","si")
  celdas = celdas & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","","AUSENCIAS","si")
  celdas = celdas & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","","TAREAS","si")
  celdas = celdas & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","","TIEMPO","si")
  filas = filas & ofv.GenRow("","aaa bgcolor='#ddddff'","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","AA COLSPAN=6  ","CENTER","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  tskant = "no"
  USRANT  = "XX"
  usrtsk  = 0
  prytsk  = 0
  prytskC = 0
  prytskR = 0

  DO UNTIL RS.EOF

  XVA = "S"

         XTIPOT =  0

      StrSql = "Select * From PROYTIPOTSK where ID = 0" & RS(17)
      set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if not  rsX1.eof then
      
         XTIPOT =  RSX1(2) 

      END IF


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
  

  IF SESSION("SELOPC") = "2" THEN
      StrSql = "Select * From PROYTIPOTSK where ID = 0" & RS(17)
      set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      if not  rsX1.eof then
      
         IF RSX1(2) = 1 THEN
            XVA = "N"



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

 '    on error resume next


     if   tskant <> rs(2) then
          tskant = rs(2) 

     IF USRANT <> RS(13) THEN

     if usrtsk > 0 then

       xTrColor = "#FFFFFF" 


      celdas = ofv.GenCelda("","vtI  COLSPAN=6 style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","","","","","",USRDESC,"si")
      filas = filas & ofv.GenRow("","aaa bgcolor='#ddddff'","","","10","","",celdas,"si")


      FOR ITIPO = 1 TO TTIPO

      IF  ATIPO(ITIPO,IUSR,1) <> "" THEN

          wktotalinput = usrtskc + usrtskr

          if wktotalinput = 0 then wktotalinput = 1


      SINPUT = ATIPO(ITIPO,0,0)
      celdas = ofv.GenCelda("","vti style='border-BOTTOM-style:solid;border-BOTTOM-width:1;border-BOTTOM-color:#DDDDDD;' ","","","","","","&nbsp;&nbsp;&nbsp;" & sImagenPI & SINPUT,"si")
      SINPUT = ATIPO(ITIPO,IUSR,1)
      celdas = celdas & ofv.GenCelda("","ut style='border-BOTTOM-style:solid;border-BOTTOM-width:1;border-BOTTOM-color:#DDDDDD;' ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = GETGLBHMS(ATIPO(ITIPO,IUSR,2))
      celdas = celdas & ofv.GenCelda("","ut style='border-BOTTOM-style:solid;border-BOTTOM-width:1;border-BOTTOM-color:#DDDDDD;' ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = GETGLBHMS(ATIPO(ITIPO,IUSR,3))
      celdas = celdas & ofv.GenCelda("","ut style='border-BOTTOM-style:solid;border-BOTTOM-width:1;border-BOTTOM-color:#DDDDDD;' ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = ROUND((ATIPO(ITIPO,IUSR,1) / usrtsk * 100 ),2) & "%"
      celdas = celdas & ofv.GenCelda("","ut style='border-BOTTOM-style:solid;border-BOTTOM-width:1;border-BOTTOM-color:#DDDDDD;' ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = ROUND( cdbl( (ATIPO(ITIPO,IUSR,2) + ATIPO(ITIPO,IUSR,3)) / (wktotalinput) * 100 ),2) & "%"

      celdas = celdas & ofv.GenCelda("","ut style='border-BOTTOM-style:solid;border-BOTTOM-width:1;border-BOTTOM-color:#DDDDDD;color:#0000aa;' ","RIGHT","","","","","&nbsp;" & sInput,"si")
      filas = filas & ofv.GenRow("","aaa  BGCOLOR='" & xTrColor & "' ","","","10","","",celdas,"si")

	 if xTrColor = "#FFFFFF" then
	    xTrColor = "#F0F0F0" 
	 ELSE
	    xTrColor = "#FFFFFF" 
	 END IF

      END IF

      NEXT

      celdas = ofv.GenCelda("","AA COLSPAN=6  ","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      celdas = ofv.GenCelda("","vt    ALIGN='right' ","","","","","","TOTAL&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","si")
      SINPUT = usrtsk
      celdas = celdas & ofv.GenCelda("","CC style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = GETGLBHMS(usrtskc)
      celdas = celdas & ofv.GenCelda("","CC style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = GETGLBHMS(usrtskr)
      celdas = celdas & ofv.GenCelda("","CC style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = "&nbsp;"
      celdas = celdas & ofv.GenCelda("","UT ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = "&nbsp;"
      celdas = celdas & ofv.GenCelda("","UT ","RIGHT","","","","","&nbsp;" & sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

''xcdias

   ''   xcaus = clng(usrtskr / CLNG(SESSION("DPP")))
      xcaus = clng(usrtskr )

      celdas = ofv.GenCelda("","vt   ALIGN='right' ","","","","","","PROMEDIO","si")
      SINPUT = round((usrtsk / (xcdias - clng(usrtskr / CLNG(SESSION("DPP"))))),4)

      celdas = celdas & ofv.GenCelda("","CC style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","","&nbsp;" & sInput,"si")


      SINPUT = ConvHM((round((clng(usrtskc ) / (xcdias   - clng(usrtskr / CLNG(SESSION("DPP"))))),2) ))     
      SINPUT2 = " (" & (round((clng(usrtskc ) / ((xcdias * CLNG(SESSION("DPP"))) - xcaus)) * 100,2)) & " %)" 


      celdas = celdas & ofv.GenCelda("","CC style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","","&nbsp;" & sInput & "&nbsp;" & sInput2,"si")
      SINPUT = "&nbsp;"
      celdas = celdas & ofv.GenCelda("","CC style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = "&nbsp;"
      celdas = celdas & ofv.GenCelda("","UT ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = "&nbsp;"
      celdas = celdas & ofv.GenCelda("","UT ","RIGHT","","","","","&nbsp;" & sInput,"si")
   '   filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      celdas = ofv.GenCelda("","AA COLSPAN=6 style='border-top-style:solid;border-top-width:1;border-top-color:#aaaaaa;' ","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","AA COLSPAN=6  ","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


     end if

   
    IUSR = IUSR + 1


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




''''''''''''''''''''''''''''''''

        USRTMP = RS(15) & ""
        USRTMPU = RS(16)
        IF USRTMP = "" THEN
           USRTMP = 1
           USRTMPU = 4
        END IF
        USRTMP = CINT(USRTMP)
        IF USRTMP = 0 THEN
           USRTMP = 1 
           USRTMPU = 4
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


''''''''''''''''''''''''''''''''''''''''''''

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

'''                xagrr = round((xagrd - xagrc),2)




                    usrtsk  =  usrtsk  + 1
                    prytsk  =  prytsk  + 1

      IF  ATIPO(RS(17),IUSR,1) <> "" THEN

          ATIPO(RS(17),IUSR,1) = ATIPO(RS(17),IUSR,1) + 1

      ELSE
          ATIPO(RS(17),IUSR,1) = 1

      END IF

      IF  ATIPO(RS(17),IGLB,1) <> "" THEN

          ATIPO(RS(17),IGLB,1) = ATIPO(RS(17),IGLB,1) + 1

      ELSE
          ATIPO(RS(17),IGLB,1) = 1

      END IF

                IF XTIPOT < 2 THEN

                    usrtskc =  usrtskc + xagrd
                    prytskc =  prytskc + xagrd

      IF  ATIPO(RS(17),IUSR,2) <> "" THEN

          ATIPO(RS(17),IUSR,2) = ATIPO(RS(17),IUSR,2) + xagrD

      ELSE
          ATIPO(RS(17),IUSR,2) = xagrD

      END IF

      IF  ATIPO(RS(17),IGLB,2) <> "" THEN

          ATIPO(RS(17),IGLB,2) = ATIPO(RS(17),IGLB,2) + xagrD

      ELSE
          ATIPO(RS(17),IGLB,2) = xagrD

      END IF




      IF  ATIPO(RS(17),IUSR,3) <> "" THEN


      ELSE
          ATIPO(RS(17),IUSR,3) = 0

      END IF

      IF  ATIPO(RS(17),IGLB,3) <> "" THEN


      ELSE
          ATIPO(RS(17),IGLB,3) = 0

      END IF




                ELSE

                    usrtskr =  usrtskr + xagrd
                    prytskr =  prytskr + xagrd

      IF  ATIPO(RS(17),IUSR,3) <> "" THEN

          ATIPO(RS(17),IUSR,3) = ATIPO(RS(17),IUSR,3) + xagrD

      ELSE
          ATIPO(RS(17),IUSR,3) = xagrD

      END IF

      IF  ATIPO(RS(17),IGLB,3) <> "" THEN

          ATIPO(RS(17),IGLB,3) = ATIPO(RS(17),IGLB,3) + xagrD

      ELSE
          ATIPO(RS(17),IGLB,3) = xagrD

      END IF



      IF  ATIPO(RS(17),IUSR,2) <> "" THEN


      ELSE
          ATIPO(RS(17),IUSR,2) = 0

      END IF

      IF  ATIPO(RS(17),IGLB,2) <> "" THEN


      ELSE
          ATIPO(RS(17),IGLB,2) = 0

      END IF



                END IF









 ' prytsk  = 0
 ' prytskC = 0
'  prytskR = 0

    end if

   END IF
 END IF

    RS.MOVENEXT
    LOOP

     if usrtsk > 0 then

      celdas = ofv.GenCelda("","vtI COLSPAN=6 style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","","","","","",USRDESC,"si")
      filas = filas & ofv.GenRow("","aaa bgcolor='#ddddff'","","","10","","",celdas,"si")


      FOR ITIPO = 1 TO TTIPO

      IF  ATIPO(ITIPO,IUSR,1) <> "" THEN

          wktotalinput = usrtskc + usrtskr

          if wktotalinput = 0 then wktotalinput = 1

      SINPUT = ATIPO(ITIPO,0,0)
      celdas = ofv.GenCelda("","vti style='border-BOTTOM-style:solid;border-BOTTOM-width:1;border-BOTTOM-color:#DDDDDD;' ","","","","","","&nbsp;&nbsp;&nbsp;" & sImagenPI & SINPUT,"si")
      SINPUT = ATIPO(ITIPO,IUSR,1)
      celdas = celdas & ofv.GenCelda("","ut style='border-BOTTOM-style:solid;border-BOTTOM-width:1;border-BOTTOM-color:#DDDDDD;' ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = GETGLBHMS(ATIPO(ITIPO,IUSR,2))
      celdas = celdas & ofv.GenCelda("","ut style='border-BOTTOM-style:solid;border-BOTTOM-width:1;border-BOTTOM-color:#DDDDDD;' ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = GETGLBHMS(ATIPO(ITIPO,IUSR,3))
      celdas = celdas & ofv.GenCelda("","ut style='border-BOTTOM-style:solid;border-BOTTOM-width:1;border-BOTTOM-color:#DDDDDD;' ","RIGHT","","","","","&nbsp;" & sInput,"si")

      SINPUT = ROUND((ATIPO(ITIPO,IUSR,1) / usrtsk * 100 ),2) & "%"
      celdas = celdas & ofv.GenCelda("","ut style='border-BOTTOM-style:solid;border-BOTTOM-width:1;border-BOTTOM-color:#DDDDDD;' ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = ROUND( ( (ATIPO(ITIPO,IUSR,2) + ATIPO(ITIPO,IUSR,3)) / (wktotalinput) * 100 ),2) & "%"
      celdas = celdas & ofv.GenCelda("","ut style='border-BOTTOM-style:solid;border-BOTTOM-width:1;border-BOTTOM-color:#DDDDDD;color:#0000aa;' ","RIGHT","","","","","&nbsp;" & sInput,"si")
      filas = filas & ofv.GenRow("","aaa  BGCOLOR='" & xTrColor & "' ","","","10","","",celdas,"si")

	 if xTrColor = "#FFFFFF" then
	    xTrColor = "#F0F0F0" 
	 ELSE
	    xTrColor = "#FFFFFF" 
	 END IF

      END IF

      NEXT

      celdas = ofv.GenCelda("","AA COLSPAN=6  ","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      celdas = ofv.GenCelda("","vt    ALIGN='right' ","","","","","","TOTAL&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","si")
      SINPUT = usrtsk
      celdas = celdas & ofv.GenCelda("","CC style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = GETGLBHMS(usrtskc) 
      celdas = celdas & ofv.GenCelda("","CC style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = GETGLBHMS(usrtskr)
      celdas = celdas & ofv.GenCelda("","CC style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = "&nbsp;"
      celdas = celdas & ofv.GenCelda("","Ut ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = "&nbsp;"
      celdas = celdas & ofv.GenCelda("","Ut ","RIGHT","","","","","&nbsp;" & sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

''xcdias

   ''   xcaus = clng(usrtskr / CLNG(SESSION("DPP")))
      xcaus = clng(usrtskr )

      celdas = ofv.GenCelda("","vt   ALIGN='right' ","","","","","","PROMEDIO","si")
      SINPUT = round((usrtsk / (xcdias - clng(usrtskr / CLNG(SESSION("DPP"))))),4)
      celdas = celdas & ofv.GenCelda("","CC style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","","&nbsp;" & sInput,"si")
  ''    SINPUT = (round((clng(usrtskc ) / (xcdias   - clng(usrtskr / CLNG(SESSION("DPP"))))),2) ) & " Hs.  "   & ConvHM((round((clng(usrtskc ) / (xcdias   - clng(usrtskr / CLNG(SESSION("DPP"))))),2) ))   

      SINPUT = ConvHM((round((clng(usrtskc ) / (xcdias   - clng(usrtskr / CLNG(SESSION("DPP"))))),2) ))     
      SINPUT2 = " (" & (round((clng(usrtskc ) / ((xcdias * CLNG(SESSION("DPP"))) - xcaus)) * 100,2)) & " %)" 

      celdas = celdas & ofv.GenCelda("","CC style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","","&nbsp;" & sInput & "&nbsp;" & sInput2,"si")
     
      SINPUT = "&nbsp;"
      celdas = celdas & ofv.GenCelda("","CC style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = "&nbsp;"
      celdas = celdas & ofv.GenCelda("","Ut ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = "&nbsp;"
      celdas = celdas & ofv.GenCelda("","Ut ","RIGHT","","","","","&nbsp;" & sInput,"si")
  '    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



      celdas = ofv.GenCelda("","AA COLSPAN=6  ","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","AA COLSPAN=6  ","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      celdas = ofv.GenCelda("","AA COLSPAN=6 style='border-top-style:ridge;border-top-width:2;border-top-color:#aaaaaa;' ","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","AA COLSPAN=6  ","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","AA COLSPAN=6  ","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



     end if

   end if

'      rowspan= " rowspan=4 "

      celdas = ofv.GenCelda("","vtI BGCOLOR=#993333 STYLE='COLOR:#FFFFFF;'  COLSPAN=6 ","","","","","","TOTALES " & XDESTN,"si")
      filas = filas & ofv.GenRow("","","","","25","","",celdas,"si")

      xTrColor = "#FFFFFF"

      FOR ITIPO = 1 TO TTIPO

      IF  ATIPO(ITIPO,IGLB,1) <> "" THEN

          wktotalinput = PRYtskc + PRYtskr

          if wktotalinput = 0 then wktotalinput = 1

      SINPUT = ATIPO(ITIPO,0,0)
      celdas = ofv.GenCelda("","vtI style='border-BOTTOM-style:solid;border-BOTTOM-width:1;border-BOTTOM-color:#DDDDDD;' ","","","","","","&nbsp;&nbsp;&nbsp;" & sImagenPI & SINPUT,"si")
      SINPUT = ATIPO(ITIPO,IGLB,1)
      celdas = celdas & ofv.GenCelda("","Ut style='border-BOTTOM-style:solid;border-BOTTOM-width:1;border-BOTTOM-color:#DDDDDD;' ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = GETGLBHMS(ATIPO(ITIPO,IGLB,2))
      celdas = celdas & ofv.GenCelda("","Ut style='border-BOTTOM-style:solid;border-BOTTOM-width:1;border-BOTTOM-color:#DDDDDD;' ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = GETGLBHMS(ATIPO(ITIPO,IGLB,3))
      celdas = celdas & ofv.GenCelda("","Ut style='border-BOTTOM-style:solid;border-BOTTOM-width:1;border-BOTTOM-color:#DDDDDD;' ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = ROUND((ATIPO(ITIPO,IGLB,1) / PRYtsk * 100 ),2) & "%"
      celdas = celdas & ofv.GenCelda("","Ut style='border-BOTTOM-style:solid;border-BOTTOM-width:1;border-BOTTOM-color:#DDDDDD;' ","RIGHT","","","","","&nbsp;" & sInput,"si")

      SINPUT = ROUND( ( (ATIPO(ITIPO,IGLB,2) + ATIPO(ITIPO,IGLB,3)) / (wktotalinput) * 100 ),2) & "%"
      celdas = celdas & ofv.GenCelda("","Ut style='border-BOTTOM-style:solid;border-BOTTOM-width:1;border-BOTTOM-color:#DDDDDD;color:#0000aa;' ","RIGHT","","","","","&nbsp;" & sInput,"si")
      filas = filas & ofv.GenRow("","aaa  BGCOLOR='" & xTrColor & "' ","","","10","","",celdas,"si")

	 if xTrColor = "#FFFFFF" then
	    xTrColor = "#F0F0F0" 
	 ELSE
	    xTrColor = "#FFFFFF" 
	 END IF

      END IF

      NEXT

      celdas = ofv.GenCelda("","AA COLSPAN=6  ","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      celdas = ofv.GenCelda("","vt   ALIGN='RIGHT' ","","","","","","TOTAL&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","si")
      SINPUT = PRYtsk
      celdas = celdas & ofv.GenCelda("","CC style='border-style:solid;border-width:1;border-color:#AAAAAA;' ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = GETGLBHMS(PRYtskc)
      celdas = celdas & ofv.GenCelda("","CC style='border-style:solid;border-width:1;border-color:#AAAAAA;' ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = GETGLBHMS(PRYtskr)
      celdas = celdas & ofv.GenCelda("","CC style='border-style:solid;border-width:1;border-color:#AAAAAA;' ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = "&nbsp;"
      celdas = celdas & ofv.GenCelda("","Ut ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = "&nbsp;"
      celdas = celdas & ofv.GenCelda("","Ut ","RIGHT","","","","","&nbsp;" & sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

''xcdias

      xcaus = clng(PRYtskr)

      celdas = ofv.GenCelda("","vt    ALIGN='RIGHT' ","","","","","","PROMEDIO","si")
      SINPUT = round((PRYtsk / (xcdias - clng(PRYtskr / CLNG(SESSION("DPP"))))),4)
      celdas = celdas & ofv.GenCelda("","CC style='border-style:solid;border-width:1;border-color:#AAAAAA;' ","RIGHT","","","","","&nbsp;" & sInput,"si")

      SINPUT = ConvHM((round((clng(PRYtskc ) / (xcdias   - clng(xcaus / CLNG(SESSION("DPP"))))),2) ))     
   '''   SINPUT = (round((clng(PRYtskc ) / (xcdias   - clng(xcaus/ CLNG(SESSION("DPP"))))),2) ) & " Hs."       
      SINPUT2 = " (" & (round((clng(PRYtskc ) / ((xcdias * CLNG(SESSION("DPP"))) - xcaus)) * 100,2)) & " %)" 

      celdas = celdas & ofv.GenCelda("","CC style='border-style:solid;border-width:1;border-color:#AAAAAA;' ","RIGHT","","","","","&nbsp;" & sInput & "&nbsp;" & sInput2,"si")
      
      SINPUT = "&nbsp;"
      celdas = celdas & ofv.GenCelda("","CC style='border-style:solid;border-width:1;border-color:#AAAAAA;' ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = "&nbsp;"
      celdas = celdas & ofv.GenCelda("","Ut ","RIGHT","","","","","&nbsp;" & sInput,"si")
      SINPUT = "&nbsp;"
      celdas = celdas & ofv.GenCelda("","Ut ","RIGHT","","","","","&nbsp;" & sInput,"si")
   '   filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","AA COLSPAN=6 style='border-top-style:SOLID;border-top-width:1;border-top-color:#AAAAAA;' ","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")




    FilaS = ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")

    Celda = ofv.GenCelda("","","center","","","","",FilaS,"si")
    FilaS = ofv.GenRow("","","","","","","",Celda,"si")


     celdas = ofv.GenCelda("","","","","","","","&nbsp;","si")
     celdas = celdas & ofv.GenCelda("","vt","","10%","","","",sInput,"si")

     fila = ofv.GenRow("","","","","","","",celdas,"si")

     Fila = ofv.GenTabla("","","","100%","","0","","0","0","","",fila,"si")

    Celda = ofv.GenCelda("","","center","","","","",Fila,"si")
 '   FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aa","","98%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"





 




  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
  call ofv.cerrarconn(cn2)
  call ofv.cerrarconn(cn3)


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

function ConvHM(hdecimal)

xhdec = cstr(hdecimal) & "00"

xvdec = instr(xhdec,".")

if xvdec <> 0 then
   xhdecD = mid(xhdec,xvdec+1,2)

   xhdecD = xhdecD * 60
   xhdecD = round(clng(xhdecD)/100,0) & "m "
   xhdecH = clng(replace(left(xhdec,xvdec),".","")) & "h " 
else
   xhdecD = ""
   xhdecH = clng(xhdec) & "h " 
end if

ConvHM = ""

if xhdecH <> "0h " then ConvHM = xhdecH

ConvHM = ConvHM & xhdecD 

end function

%>