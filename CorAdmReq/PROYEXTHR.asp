<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
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



 sopcion = request.querystring("opcion")


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




  StrSql = "Select * From PROYETPTSK where codigo = '" & session("NCliente") & "' AND VPLAN = " & SESSION("VPLAN") & " ORDER BY USRRESP,ORDEN,TSKORD "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then


 

    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"

  celdas = ofv.GenCelda("","cc colspan=5","","","","","",sImagen & "&nbsp;<BIG><B>TOTAL HORAS POR RESPONSABLE</B></BIG>","si")
  filas  = ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","AA COLSPAN=5","CENTER","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","vti ROWSPAN=2 style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","","","","","","RESPONSABLE","si")
  celdas = celdas & ofv.GenCelda("","vt COLSPAN=4 style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","CENTER","","","","","DETALLE DE TIEMPOS","si")
  filas = filas & ofv.GenRow("","aaa bgcolor='#ddddff' ","","","10","","",celdas,"si")

  celdas = ""
  celdas = celdas & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","","DURACION&nbsp;","si")
  celdas = celdas & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","","PLANIFICADO&nbsp;","si")
  celdas = celdas & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","","CONSUMIDO&nbsp;","si")
  celdas = celdas & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","","RESTANTE&nbsp;","si")
  filas = filas & ofv.GenRow("","aaa bgcolor='#ddddff' ","","","10","","",celdas,"si")

  celdas = ofv.GenCelda("","AA COLSPAN=5  ","CENTER","","","","","&nbsp;","si")
  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  USRANT  = "XX"
  usrtsk  = 0
  prytsk  = 0
  prytskC = 0
  prytskR = 0
  prytskD = 0



  DO UNTIL RS.EOF

     IF USRANT <> RS(13) THEN




     if usrtsk > 0 then

      wkstyle = ""
      celdas = ofv.GenCelda("","vt  bgcolor='#ffffff' style='border-top-style:solid;border-top-width:1;border-top-color:#AAAAAA;'  ","","","","","","&nbsp;","si")

      SINPUT = GETGLBHMS(usrtskD)
      celdas = celdas & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","",sInput,"si")

      SINPUT = GETGLBHMS(usrtsk)
      celdas = celdas & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","",sInput,"si")
      SINPUT = GETGLBHMS(usrtskc)
      celdas = celdas & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","",sInput,"si")
      SINPUT = GETGLBHMS(usrtskr)
      if usrtsk < usrtskc then
         wkstyle = "color:#990000;font-weight:bold;"   
         SINPUT = "-" & SINPUT
      end if
      celdas = celdas & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#aaaaaa;" & wkstyle & "' ","RIGHT","","","","",sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","AA COLSPAN=5  ","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

     end if

    xTrColor = "#FFFFFF" 

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

      celdas = ofv.GenCelda("","vt  bgcolor='#ddddff' colspan=5 style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","","","","","",USRDESC,"si")
    '  celdas = celdas & ofv.GenCelda("","vt colspan=3  style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#f7aa22;' ","RIGHT","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")




        USRANT  = RS(13)
        usrtsk  = 0
        usrtskC = 0
        usrtskR = 0
        usrtskD = 0

        FILAS1 = ""
     END IF


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

IF SESSION("PRCF") <> "S" THEN
   XXXX = GETPRYCNFG()
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

   wkmodcreate = cstr("M" & rs(23))

   if wkmodcreate = "M2" then

          xagrdD = 0

          StrSql = "Select * From PROYTAREAS where codigo = '" & session("NCliente") & "' AND tarea = '" & RS(20) & "' and vplan=0 "
          set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if NOT rsX.eof then


               
             USRTMP     = RSx(15) & ""
             USRTMPU    = RSx(16)
             IF USRTMP  = "" THEN
                USRTMP  = 0
                USRTMPU = 0
             END IF
             USRTMP = CINT(USRTMP)
             IF USRTMP  = 0 THEN
                USRTMP  = 0 
                USRTMPU = 0
             END IF

             SELECT CASE USRTMPU
                    CASE 4
                         xagrdd = USRTMP 
                    CASE 1
                         xagrdd = USRTMP * 60
                    CASE 2
                         xagrdd = (USRTMP * CLNG(SESSION("DPP"))) * 60
                    CASE ELSE
                         xagrdd = (USRTMP * CLNG(SESSION("DPP")) * CLNG(SESSION("SPP")) ) * 60
             END select         

             xagrdd = xagrdd * 60



          END IF






   else

      xagrdD = 0


   end if




   if rs(4) > 0 then

          PADV = 0 

          StrSql = "Select * From PROYTAREASADV where codigo = '" & session("NCliente") & "' AND TSK = '" & RS(20) & "' ORDER BY id desc "
          set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if NOT rsX.eof then

             PADV = RSX(5)
             xagrc = Clng((XAGRD * PADV / 100)) 
             xagrr = XAGRD - Clng((XAGRD * PADV / 100)) 

          END IF

     end if
                xagrr = xagrd - xagrc

                    usrtskc =  usrtskc + xagrc
                    prytskc =  prytskc + xagrc
                    usrtskr =  usrtskr + xagrr
                    prytskr =  prytskr + xagrr
                    usrtskD =  usrtskD + xagrDD
                    prytskD =  prytskD + xagrDD

 ' prytsk  = 0
 ' prytskC = 0
'  prytskR = 0

      SINPUT = "&nbsp;&nbsp;" & RS(0) & "." & RS(1) & "&nbsp;" & RS(8)
      celdas = ofv.GenCelda("","uti  style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#f0f0f0;' ","","","","","",SINPUT,"si")

      wkstyle = ""

      SINPUT = GETGLBHMS(xagrdD)
      celdas = celdas & ofv.GenCelda("","ut","RIGHT","","","","",sInput,"si")
      SINPUT = GETGLBHMS(xagrd)
      celdas = celdas & ofv.GenCelda("","ut","RIGHT","","","","",sInput,"si")
      SINPUT = GETGLBHMS(xagrc)
      celdas = celdas & ofv.GenCelda("","ut","RIGHT","","","","",sInput,"si")
      SINPUT = GETGLBHMS(xagrr)
      if xagrd < xagrc then
         wkstyle = "color:#990000;font-weight:bold;"   
         SINPUT = "-" & SINPUT
      end if
      celdas = celdas & ofv.GenCelda("","ut style='" & wkstyle & "' ","RIGHT","","","","",sInput,"si")

      filas = filas & ofv.GenRow("","aaa  BGCOLOR='" & xTrColor & "' ","","","10","","",celdas,"si")

	 if xTrColor = "#FFFFFF" then
	    xTrColor = "#F4F4F4" 
	 ELSE
	    xTrColor = "#FFFFFF" 
	 END IF

    RS.MOVENEXT
    LOOP

     if usrtsk > 0 then

      wkstyle = ""

      celdas = ofv.GenCelda("","vt  bgcolor='#ffffff' style='border-top-style:solid;border-top-width:1;border-top-color:#AAAAAA;' ","","","","","","&nbsp;","si")

      SINPUT = GETGLBHMS(usrtskD)
      celdas = celdas & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","",sInput,"si")
      SINPUT = GETGLBHMS(usrtsk)
      celdas = celdas & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","",sInput,"si")
      SINPUT = GETGLBHMS(usrtskc)
      celdas = celdas & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","",sInput,"si")
      SINPUT = GETGLBHMS(usrtskr)
      if usrtsk < usrtskc then
         wkstyle = "color:#990000;font-weight:bold;"   
         SINPUT = "-" & SINPUT
      end if
      celdas = celdas & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#aaaaaa;" & wkstyle & "' ","RIGHT","","","","",sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","AA COLSPAN=5 style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#f7aa22;' ","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

     end if

   end if

      rowspan= " rowspan=4 "
      celdas = ofv.GenCelda("","vt","","","","","","TOTALES DEL PROYECTO","si")

      wkstyle = ""

      SINPUT = GETGLBHMS(PRYtskD)
      celdas = celdas & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","",sInput,"si")
      SINPUT = GETGLBHMS(PRYtsk)
      celdas = celdas & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","",sInput,"si")
      SINPUT = GETGLBHMS(PRYtskc)
      celdas = celdas & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","RIGHT","","","","",sInput,"si")
      SINPUT = GETGLBHMS(PRYtskr)
      if PRYtsk < PRYtskc then
         wkstyle = "color:#990000;font-weight:bold;"   
         SINPUT = "-" & SINPUT
      end if
      celdas = celdas & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#aaaaaa;" & wkstyle & "' ","RIGHT","","","","",sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","AA COLSPAN=5 style='border-top-style:solid;border-top-width:1;border-top-color:#f7aa22;' ","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","20","","",celdas,"si")




    FilaS = ofv.GenTabla("","AAA style='border-style:solid;border-width:1;border-color:#aaaaaa;' ","","100%","","0","","0","0","","",filas,"si")

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

    sHTML = sHTML & ofv.GenTabla("","aa  ","","98%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"





 




  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
  call ofv.cerrarconn(cn2)
  call ofv.cerrarconn(cn3)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML



%>