<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.formHeader(" ")


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


 sopcion = request.querystring("opcion")

 

     sopcion = "VP"





    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"



  XCONTROL = "N" 
  XCPT     = "N" 
  XCADV    = "N" 

  StrSql = "Select * From PROYRESP where codigo = '" & session("NCliente") & "' AND USRAFC = '" & SESSION("USUARIO") & "' AND TIPO = 'AN' "
  set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if rsx1.eof then

     StrSql = "Select * From PROYRESP where codigo = '" & session("NCliente") & "' AND USRAFC = '" & SESSION("USUARIO") & "' "
     set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  end if
  
  
  if NOT rsX1.eof then

     IF RSX1(11) THEN
        XCONTROL = "S" 
        XCPT     = "S" 
        XCADV    = "S" 
     ELSE
        IF RSX1(13) THEN XCPT      = "S"
        IF RSX1(12) THEN XCADV     = "S"
     END IF

  END IF

  IF XCONTROL = "N" THEN

     bAgregar   = 0
     bModificar = 0
     bSuprimir  = 0

  END IF




    XFSTART = CLNG(DATE()) - clng(cdate("01-01-1900")) 
    XFCURR  = XFSTART
    XFHOY   = XFSTART
    XFREAL  = XFSTART






    StrSql = "Select * From PROYecto where codigo = '" & session("NCliente") & "'  "
    
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT rs2.eof then
       xnompry = rs2(2)
       xusrrsp = rs2(14)
       if SESSION("VPLAN")="" then SESSION("VPLAN")=RS2(17)
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

    IF XVPLAN  <> session("VPLAN") THEN
       bAgregar   = FALSE
       bModificar = FALSE
       bSuprimir  = FALSE
       XVPLANV    = "N"
    END IF

  ''    XFCURR = XFSTART

         dFecha = CDATE(CLNG(XFSTART) + clng(cdate("01-01-1900")))

      celdas = ofv.GenCelda("","ttt colspan=4","","","","","",sImagen & "&nbsp;ESTADO DEL PROYECTO AL :  " & ofv.fechador & " ","si")
      filas  = ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","VT colspan=3 valign=top ","","60%","","","","PROYECTO:&nbsp;" & xnompry,"si")
      celdas = CELDAS & ofv.GenCelda("","UT ","RIGHT","","","","","PLAN VIGENTE:&nbsp;" & SESSION("VPLAND")& "&nbsp;","si")


     filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","AA COLSPAN=4","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    StrSql = "Select * From PROYPRTY where codigo = '" & session("NCliente") & "'  "
    set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT rsX.eof then

       if rsX(4) = 0 then sInput = "BLOQUE"
       if rsX(4) = 1 then sInput = "ETAPA"
       if rsX(4) = 2 then sInput = "FASE"
   
       xtipo = rsX(3)
       xestr = sInput


    END IF



    FilaS1 = ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")



      celdas = ofv.GenCelda("","ut colspan=3","CENTER","","","","","ETAPA/TAREA","si")
      celdas = celdas & ofv.GenCelda("","uT","CENTER","","","","","DUR.","si")
      celdas = celdas & ofv.GenCelda("","uT","CENTER","","","","","INI.PLAN","si")
      celdas = celdas & ofv.GenCelda("","uT","CENTER","","","","","FIN.PLAN","si")
      celdas = celdas & ofv.GenCelda("","uT","CENTER","","","","","INI.REAL","si")
      celdas = celdas & ofv.GenCelda("","uT","CENTER","","","","","FIN.REAL","si")
      celdas = celdas & ofv.GenCelda("","uT COLSPAN=3","CENTER","","","","","ESTADO/AVANCE","si")
      filas = ofv.GenRow("","aaa bgcolor='#f9f9f9' ","","","10","","",celdas,"si")

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

    StrSql = "Select MAX(ORDEN) From PROYETAPA where codigo = '" & session("NCliente") & "' AND ORDEN < 999 AND VPLAN = " & SESSION("VPLAN")
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



   xTrColor = "#FFFFFF"


  StrSql = "Select * From PROYetapa where codigo = '" & session("NCliente") & "' AND VPLAN = " & SESSION("VPLAN") & " order by orden "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     do while not rs.eof

     sinput = "<a style='text-decoration:none;color:#ffffff;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";'  href=PROYtask.asp?opcion=SBE&etp=" & rs(2) 
     sinput = sinput  & " TITLE='SUBIR ETAPA' ><b>" & sImagen1 & "</b></a>" 
     sinput1 = sinput

     IF RS(3) = 0 OR RS(3) = 999 THEN
        sinput1 = "<B>" & sImagen1N & "</B>"
     END IF  


     sinput = "<a style='text-decoration:none;color:#ffffff;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";'  href=PROYtask.asp?opcion=AT&etp=" & rs(2) 
     sinput = sinput  & "  TITLE='AGREGAR TAREA'  ><b>" & sImagen2 & "</b></a>" 
     sinput2 = sinput

'     sinput = "<a style='text-decoration:none;color:#ffffff;' onmouseover='this.style.textDecoration=" & chr(34)
'     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
'     sinput = sinput  & "none" & chr(34) & ";'  href=PROYtaskprop.asp?opcion=EP&etp=" & rs(2) 
'     sinput = sinput  & "  TITLE='PROPIEDADES DE LA ETAPA' ><b>" & sImagen3 & "</b></a>" 
'     sinput3 = sinput
   

    sform = "ETP" & RS(0)


    SINPUT = RS(3)
    IF SINPUT = 999 THEN SINPUT = XORDEN
    XETPORD = SINPUT
    celdas = ofv.GenCelda("","aaa bgcolor=#f7aa22   ","","2%","","","",SINPUT,"si")

         sInput = rs(4)

    celdas = celdas & ofv.GenCelda("","ut STYLE='background-color:#bbddff;border-style:solid;border-width:1;border-color:#f7aa22;' colspan=10","","90%","","","",xestr & ": " & sInput,"si")


    if bModificar then
      sAccion = "asp/Actualizar.asp?FNAME=" & rs(0) & "&volver=" & vuelta & "&Tabla=PROYETAPA"
      formulario0 = ofv.GenForm(sform,"",sAccion,"","post","",celdas,"si")
    else
      formulario0 = ofv.GenForm("","","","","post","",celdas,"si")
    end if

    filas = filas & ofv.GenRow("","","","","10","","",formulario0,"si")



  StrSql = "Select * From PROYtareas where codigo = '" & session("NCliente") & "' and etapa = '" & rs(2) & "' AND VPLAN = " & SESSION("VPLAN") & " order by orden "
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not  rs1.eof then
     do while not rs1.eof

    sform = "TSK" & RS1(0)


     sinput1 = sImagenE0

    TSKMODCREATE = "M" & RS1(33)




 '''   celdas = ofv.GenCelda("","","CENTER","2%","","","","&nbsp;","si")
    celdas = ofv.GenCelda("","ut colspan=2  ","","7%","","","",XETPORD & "." & rs1(4),"si")

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
    XSTSSIGN  = OWKF.ATABTSK(cint(XIIX),22)

    IF XTDUR > 0 THEN
       SINPUT = XTDUR
       if XUDUR = 4 then
          SINPUT = SINPUT & "&nbsp;m"
       elseif XUDUR = 1 then
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
       if XFFTI = 4 then
          SINPUT = SINPUT & "&nbsp;m"
       elseif XFFTI = 1 then
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

       IF XVPLANV    <> "N" THEN
        if rs1(6) = 0 and rs1(7) = 0 then
         if XCONTROL = "S" then


        '    wwdif =  TSKFE - TSKFS

        '    dFecha = cdate(clng(date()) )
        '    SINPUTn2 = dFecha

        '    dFecha = cdate(clng(date())  + wwdif )
        '    SINPUTn3 = dFecha

        '    SINPUTn4 = SINPUTn2
        '    SINPUTn5 = SINPUTn3







          sinput = "<a style='text-decoration:none;color:#336699;' onmouseover='this.style.textDecoration=" & chr(34)
          sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
          sinput = sinput  & "none" & chr(34) & ";'  href=PROYtaskAdv.asp?TSK=" & rs1(3) 
          sinput = sinput  & "&opcion=LT  TITLE='Inicio manual de la Tarea'  ><b>" & sImagenE4 & "</b></a>" 
          sinput3 = sinput

         end if
       end if
      END IF

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


       IF TSKMODCREATE = "M2" THEN

          sinput = "<a style='text-decoration:none;color:#336699;' onmouseover='this.style.textDecoration=" & chr(34)
          sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
          sinput = sinput  & "none" & chr(34) & ";'  href='PROYtaskprop.asp?opcion=M2&TSK=" & rs1(3) & "' "
          sinput = sinput  & " TITLE='Ajustar tiempo consumido en la tarea.'  ><b>" & sinput3 & "</b></a>" 
          sinput3 = sinput

       END IF


   elseif TSKSt = "1" then


          dFecha = cdate(TSKFSR + clng(cdate("01-01-1900")))
          SINPUTn4 = dFecha
          c4 = " style='color:#336699;font-weight:bold;' "


          sinput1  = sImagenE1
          sinputt2 = SIMAGEN6N

             sinput3 = PADV &  "%"


             IF XSTSSIGN = 0 THEN
                sinputt2 = SIMAGEN6S
                c5 = " style='color:#008000;font-weight:bold;' "

             ELSEIF XSTSSIGN = 1  THEN
                sinputt2 = SIMAGEN6A
                c5 = " style='color:#F7AA22;font-weight:bold;' "

             ELSEIF XSTSSIGN = 2  THEN
                sinputt2 = SIMAGEN6N
                sinputn0 = vfondo & "&nbsp;" & sinputn0
                c5 = " style='color:#993333;font-weight:bold;' "

             ELSE
                sinputt2 = SIMAGEN6S
                c5 = " style='color:#008000;font-weight:bold;' "

             END IF





          SINPUTn5 = cdate(TSKFER + clng(cdate("01-01-1900")))

       IF XVPLANV    <> "N" THEN

          IF XCONTROL = "S" OR (XCADV = "S" AND RS1(13) = SESSION("USUARIO")) THEN

          sinput = "<a style='text-decoration:none;color:#336699;' onmouseover='this.style.textDecoration=" & chr(34)
          sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
          sinput = sinput  & "none" & chr(34) & ";'  href=PROYtaskAdv.asp?TSK=" & rs1(3) 
          sinput = sinput  & "  TITLE='Actualizar Avance de la Tarea'  ><b>" & sinput3 & "</b></a>" 
          sinput3 = sinput

          END IF 

      END IF 



   end if




      celdas = celdas & ofv.GenCelda("","ut ","","75%","","","",sInputn0,"si")
      celdas = celdas & ofv.GenCelda("","ut ","right","5%","","","",sInputn1,"si")
      celdas = celdas & ofv.GenCelda("","ut ","","5%","","","",sInputn2,"si")
      celdas = celdas & ofv.GenCelda("","ut ","","5%","","","",sInputn3,"si")
      celdas = celdas & ofv.GenCelda("","ut " & c4,"CENTER","5%","","","",sInputn4,"si")
      celdas = celdas & ofv.GenCelda("","ut " & c5,"CENTER","5%","","","",sInputn5,"si")
      celdas = celdas & ofv.GenCelda("","UT ","CENTER","","1%","","",sinput1,"si")
      celdas = celdas & ofv.GenCelda("","ut ","CENTER","","1%","","",sinputt2,"si")
      celdas = celdas & ofv.GenCelda("","UT ","CENTER","","5%","","",sinput3,"si")

    if bModificar then
      sAccion = "asp/Actualizar.asp?FNAME=" & rs1(0) & "&volver=" & vuelta & "&Tabla=PROYTAREAS"
      formulario0 = ofv.GenForm(sform,"",sAccion,"","post","",celdas,"si")
    else
      formulario0 = ofv.GenForm("","","","","post","",celdas,"si")
    end if

    filas = filas & ofv.GenRow("","aaa  BGCOLOR='" & xTrColor & "' ","","","10","","",formulario0,"si")


                         if xTrColor = "#FFFFFF" then
                            xTrColor = "#F0F0F0" 
                         ELSE
                            xTrColor = "#FFFFFF" 
                         END IF


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


    FilaS2 = ofv.GenTabla("","aa style='border-style:solid;border-width:1;border-color:#888888;' ","","100%","","0","","-1","2","","",filas,"si")


    Celda = ofv.GenCelda("","","center","","","","",FilaS1 & FilaS2,"si")
    FilaS = ofv.GenRow("","","","","","","",Celda,"si")


     sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' class=tt href=PROYgraf.asp"
     sinput = sinput  & " target='_blank' ><b>gantt</b></a>" 
''     celdas = ofv.GenCelda("","tt","center","10%","","","",sInput,"si")

     sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' class=tt href=PROYPROP.asp"
     sinput = sinput  & " ><b>VOLVER</b></a>" 
 ''    celdas = celdas & ofv.GenCelda("","tt","center","10%","","","",sInput,"si")


     cnbot = "<input type=submit class=bt value='GANTT' >"
     Cons = ofv.gencelda("","","","","1%","","",cnbot,"si")
     sAccion = "PROYgraf.asp"
     Cons = ofv.genform("cons","",sAccion,"_blank","post","",Cons,"si")

     celdas = Cons

     celdas = celdas & ofv.gencelda("","","","","1%","","","&nbsp;","si")

     celdas = celdas & ofv.GenCelda("","","","30%","","","","&nbsp;","si")

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

     Fila = ofv.GenTabla("","","","","","0","","4","0","","",fila,"si")









    Celda = ofv.GenCelda("","","center","","","","",Fila,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","","","90%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"







  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML



%>