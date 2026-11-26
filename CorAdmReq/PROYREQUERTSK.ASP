<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.MenuHeader("","")

  sHTML = sHTML  & "<script language=javascript> "
  sHTML = sHTML  & "function AbrirCons(Datos) { "
  sHTML = sHTML  & " myWin= open(Datos," & chr(34) & "Cons" & chr(34) & "," & chr(34) & "menubar=no,toolbar=no,height=570,width=790" & chr(34) & " ); "
  sHTML = sHTML  & " } "
  sHTML = sHTML  & "</script>"

  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  set cn = ofv.conectar(ofv.strconn6)
  set cn2 = ofv.conectar(ofv.strconn0)

  ofv.ObtenerAtributos session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  sopcion = request.querystring("opcion")
  sopc = request.querystring("opc")

  IF SOPC <> "" THEN
     SESSION("ROPC") = SOPC
  ELSE
     SOPC = SESSION("ROPC")
  END IF






    StrSql = "Select CODIGO,TAREA From PROYTAREAS where PROYTSKOBJ = '" & session("NCliente") & "' "
    set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT rsX.eof then
       XPRY = RSX(0)
       XTSK = RSX(1)
    END IF
    call ofv.cerrarconsulta(rsX)

   sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"



    XFSTART = CLNG(DATE()) - clng(cdate("01-01-1900")) 
    XFCURR  = XFSTART
    XFHOY  = XFSTART
    XFREAL  = XFSTART






    StrSql = "Select * From PROYecto where codigo = '" &  XPRY & "'  "
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT rs2.eof then
       xnompry = rs2(2)
       xusrrsp = rs2(14)
       if rs2(4) > 0 then
          XFSTART = rs2(4)
          XFCURR  = XFSTART
          XFREAL  = XFSTART
       else
          XFSTART = rs2(3)
          XFCURR  = XFSTART
          XFREAL  = XFSTART
       end if
       XVPLAN = RS2(17)
       XFFCH  = RS2(18)
       IF XFFCH > 0 THEN XXX = OWKF.ADMFERIADOS("C",XFSTART)
 
    end if
     call ofv.cerrarconsulta(rs2)

       XVPLANV    = "S"



  ''    XFCURR = XFSTART

         dFecha = CDATE(CLNG(XFSTART) + clng(cdate("01-01-1900")))

      celdas = ofv.GenCelda("","cc colspan=4","","","","","",sImagen & "&nbsp;<BIG><B>ESTADO DEL PROYECTO AL :  " & ofv.fechador & "</B></BIG>","si")
      filas  = ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","VT colspan=3 valign=top ","","60%","","","","PROYECTO:&nbsp;" & xnompry,"si")
      celdas = CELDAS & ofv.GenCelda("","UT ","RIGHT","","","","","PLAN VIGENTE:&nbsp;" & SESSION("VPLAND")& "&nbsp;","si")


     filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","AA COLSPAN=4","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    StrSql = "Select * From PROYPRTY where codigo = '" & XPRY & "'  "
    set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT rsX.eof then

       if rsX(4) = 0 then sInput = "BLOQUE"
       if rsX(4) = 1 then sInput = "ETAPA"
       if rsX(4) = 2 then sInput = "FASE"
   
       xtipo = rsX(3)
       xestr = sInput


    END IF



    FilaS1 = ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")

      celdas = ofv.GenCelda("","vt colspan=3","CENTER","","","","","ETAPA/TAREA","si")
      celdas = celdas & ofv.GenCelda("","VT","CENTER","","","","","DUR.","si")
      celdas = celdas & ofv.GenCelda("","VT","CENTER","","","","","INI.PLAN","si")
      celdas = celdas & ofv.GenCelda("","VT","CENTER","","","","","FIN.PLAN","si")
      celdas = celdas & ofv.GenCelda("","VT","CENTER","","","","","INI.REAL","si")
      celdas = celdas & ofv.GenCelda("","VT","CENTER","","","","","FIN.REAL","si")
      celdas = celdas & ofv.GenCelda("","VT COLSPAN=3","CENTER","","","","","ESTADO/AVANCE","si")
      filas = ofv.GenRow("","","","","10","","",celdas,"si")

    sImagen0  = "<FONT FACE=WINGDINGS SIZE=2>F</FONT>"
    sImagen1  = "<FONT FACE=WINGDINGS SIZE=2>ñ</FONT>"
    sImagen1N = "<FONT FACE=WINGDINGS COLOR=#AA0000 SIZE=2>l</FONT>"
    sImagen4  = "<FONT FACE=WebDINGS COLOR=#AA0000 SIZE=3>4</FONT>"
    sImagen2  = "<FONT FACE=VERDANA SIZE=2>+</FONT>"
    sImagen3  = "<FONT FACE=WINGDINGS SIZE=2>2</FONT>"
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

    StrSql = "Select MAX(ORDEN) From PROYETAPA where codigo = '" & XPRY & "' AND ORDEN < 999 AND VPLAN = " & XVPLAN
    set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if rsX.eof then
       XORDEN = 10
    ELSE
       XORDEN = RSX(0)
       IF XORDEN <> "" THEN
          XORDEN = XORDEN + 10
       ELSE
          XORDEN = 10
       END IF
    END IF
     call ofv.cerrarconsulta(rsx)


  StrSql = "Select * From PROYetapa where codigo = '" & XPRY & "' AND VPLAN = " & XVPLAN & " order by orden "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     do while not rs.eof


        sinput1 = "<B>" & sImagen1N & "</B>"
        sinput2 = "<B>" & sImagen1N & "</B>"



   

    sform = "ETP" & RS(0)


    SINPUT = RS(3)
    IF SINPUT = 999 THEN SINPUT = XORDEN
    XETPORD = SINPUT
    celdas = ofv.GenCelda("","aaa bgcolor=#f7aa22   ","","2%","","","",SINPUT,"si")

    sInput = rs(4)

    celdas = celdas & ofv.GenCelda("","vt STYLE='background-color:#ffffff;border-bottom-style:outset;border-bottom-width:2;border-bottom-color:#f7aa22;border-top-style:outset;border-top-width:2;border-top-color:#f7aa22;' colspan=10","","90%","","","",xestr & ": " & sInput,"si")




    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



  StrSql = "Select * From PROYtareas where codigo = '" & XPRY & "' AND VPLAN = " & XVPLAN & " and etapa = '" & rs(2) & "'  order by orden "
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not  rs1.eof then
     do while not rs1.eof

    sform = "TSK" & RS1(0)


    sinput1 = sImagenE0

    STYLE = " STYLE='BACKGROUND-COLOR:#FFFFFF;' "

    IF XTSK = RS1(3) THEN
       STYLE = " STYLE='BACKGROUND-COLOR:#F7EEDD;COLOR:#993333;FONT-WEIGHT:BOLD;' "

    END IF

    celdas = ofv.GenCelda("","aaa colspan=2  " & STYLE,"","7%","","","",XETPORD & "." & rs1(4),"si")

    sInputn0 = rs1(5)

    vfondo = sImagen1N

    IF RS1(13) <> "NO" THEN
       SINPUT4 = SIMAGEN4S
       XCOLOR = "#DDDDDD"
    ELSE
       SINPUT4 = "<B>" & SIMAGEN4N & "</B>"
       XCOLOR = "#AA0000"
    END IF 

    XIIX = OWKF.GETTSKCLC(RS1,XFHOY,XFREAL ,XFFCH)




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







    IF XTDUR > 0 THEN
       SINPUT = XTDUR
       if XUDUR = 1 then
          SINPUT = SINPUT & "&nbsp;h"
       elseif XUDUR = 3 then
          SINPUT = SINPUT & "&nbsp;s"
       else
          SINPUT = SINPUT & "&nbsp;d"
       end if 
       SINPUT = "<b>" & SINPUT & "</b>"
    ELSE
       SINPUT = "-"
    END IF 

    IF XMODA > 0 AND (XTDUP > 0 AND (XTDUP <> XTDUR OR XUDUP <> XUDUR )) THEN
       SINPUT = XTDUP
       IF XUDUP > 0  THEN XFFTI = XUDUP
       if XFFTI = 1 then
          SINPUT = SINPUT & "&nbsp;h"
       elseif XFFTI = 3 then
          SINPUT = SINPUT & "&nbsp;s"
       else
          SINPUT = SINPUT & "&nbsp;d"
       end if 
       SINPUT = "<font color=#aa0000><b>" & SINPUT & "</b></font>"
    END IF 


       SINPUTN1 = SINPUT


       dFecha = cdate(TSKFS + clng(cdate("01-01-1900")))
       SINPUTn2 = dFecha


       dFecha = cdate(TSKFE + clng(cdate("01-01-1900")))
       SINPUTn3 = dFecha




   if  TSKSt = "0" then


       dFecha = cdate(TSKFSR + clng(cdate("01-01-1900")))
       SINPUTn4 = dFecha
       c4 = " style='color:#bbbbbb;' "


       dFecha = cdate(TSKFER + clng(cdate("01-01-1900")))
       SINPUTn5 = dFecha
       c5 = " style='color:#bbbbbb;' "


          sinput1 = sImagenE0
          sinput3 = SIMAGEN6W
          sinputt2 = SIMAGEN6W



   elseif TSKSt = "2" then

          dFecha = cdate(TSKFSR + clng(cdate("01-01-1900")))
          SINPUTn4 = dFecha
          c4 = " style='color:#336699;font-weight:bold;' "


          dFecha = cdate(TSKFER + clng(cdate("01-01-1900")))
          SINPUTn5 = dFecha
          c5 = " style='color:#336699;font-weight:bold;' "


          sinput1 = sImagenE2
          sinput3 = "100%"
          sinputt2 = SIMAGEN6S

   elseif TSKSt = "1" then


          dFecha = cdate(TSKFSR + clng(cdate("01-01-1900")))
          SINPUTn4 = dFecha
          c4 = " style='color:#336699;font-weight:bold;' "


          sinput1  = sImagenE1
          sinputt2 = SIMAGEN6N

             sinput3 = PADV &  "%"

          IF PADV > 0 THEN

             IF XFHOY < TSKFE THEN
                sinputt2 = SIMAGEN6S
                c5 = " style='color:#008000;font-weight:bold;' "

             ELSEIF XFHOY = TSKFE AND PADV > 99  THEN
                sinputt2 = SIMAGEN6S
                c5 = " style='color:#008000;font-weight:bold;' "

             ELSE
                sinputt2 = SIMAGEN6N
                sinputn0 = vfondo & "&nbsp;" & sinputn0
                c5 = " style='color:#993333;font-weight:bold;' "

             END IF

          else

             IF XFHOY > TSKFSR THEN
                sinputt2 = SIMAGEN6A
                c5 = " style='color:#F7AA22;font-weight:bold;' "
             END IF

             IF XFHOY < TSKFE THEN
             '   sinputt2 = SIMAGEN6S
             '   c5 = " style='color:#008000;font-weight:bold;' "

             ELSE
                sinputt2 = SIMAGEN6N
                c5 = " style='color:#993333;font-weight:bold;' "

             END IF




          END IF


          SINPUTn5 = cdate(TSKFER + clng(cdate("01-01-1900")))





   end if



      sinputX = "&nbsp;"

    STYLE = "BACKGROUND-COLOR:#FFFFFF;"

    IF XTSK = RS1(3) THEN
       STYLE = "BACKGROUND-COLOR:#F7EEDD;COLOR:#993333;FONT-WEIGHT:BOLD;"

    END IF

      celdas = celdas & ofv.GenCelda("","ut COLSPAN=9   STYLE='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#f7aa22;" & STYLE & "'","","75%","","","",sInputn0,"si")


    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","aaa COLSPAN=2 ","center","5%","","","",sinputX,"si")
      celdas = celdas & ofv.GenCelda("","aaa bgcolor=#ffffff ","right","5%","","","",sInputn1,"si")
      celdas = celdas & ofv.GenCelda("","aaa bgcolor=#ffffff ","","5%","","","",sInputn2,"si")
      celdas = celdas & ofv.GenCelda("","aaa bgcolor=#ffffff ","","5%","","","",sInputn3,"si")
      celdas = celdas & ofv.GenCelda("","aaa bgcolor=#ffffff " & c4,"CENTER","5%","","","",sInputn4,"si")
      celdas = celdas & ofv.GenCelda("","aaa bgcolor=#ffffff " & c5,"CENTER","5%","","","",sInputn5,"si")
      celdas = celdas & ofv.GenCelda("","UT ","CENTER","","1%","","",sinput1,"si")
      celdas = celdas & ofv.GenCelda("","ut ","CENTER","","1%","","",sinputt2,"si")
      celdas = celdas & ofv.GenCelda("","UT  bgcolor=#FFFFFF ","CENTER","","5%","","",sinput3,"si")



    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    rs1.movenext
    loop
    call ofv.cerrarconsulta(rs1)
  end if



    rs.movenext
    loop
 
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


    Celda = ofv.GenCelda("","","center","","","","",FilaS1 & FilaS2,"si")
    FilaS = ofv.GenRow("","","","","","","",Celda,"si")





     celdas = ""

 '    celdas = celdas & ofv.GenCelda("","","","30%","","","","&nbsp;","si")

     sinput = sImagenE0 & "&nbsp;No&nbsp;iniciada" 
     celdas = celdas & ofv.GenCelda("","ut","","10%","","","",sInput,"si")
     sinput = sImagenE1 & "&nbsp;En&nbsp;curso" 
     celdas = celdas & ofv.GenCelda("","ut","","10%","","","",sInput,"si")
     sinput = sImagenE2 & "&nbsp;Finalizada" 
     celdas = celdas & ofv.GenCelda("","ut","","10%","","","",sInput,"si")


     sinput = SIMAGEN6S & "&nbsp;Dentro&nbsp;del&nbsp;plan" 
     celdas = celdas & ofv.GenCelda("","ut","","10%","","","",sInput,"si")
     sinput = SIMAGEN6N & "&nbsp;Con&nbsp;desvio" 
     celdas = celdas & ofv.GenCelda("","ut","","10%","","","",sInput,"si")



     fila = ofv.GenRow("","","","","","","",celdas,"si")

     Fila = ofv.GenTabla("","aaa bgcolor=#ffffff ","","","","0","","4","0","","",fila,"si")









    Celda = ofv.GenCelda("","","center","","","","",Fila,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aa style='filter:shadow(color=#cccccc,direction=125);' ","","90%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"



 

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML  & "</body></html>"

 END IF

  RESPONSE.WRITE SHTML


%>