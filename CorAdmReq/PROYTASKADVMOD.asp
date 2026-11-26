<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


 sopcion = request.querystring("opcion") & ""
 sMODAL = request.querystring("MODAL") & ""


 IF (SMODAL = "MOD" OR SMODAL = "NEW") AND SOPCION = "" THEN


     SESSION("TSKVLV") = "PRYPLANILLAUSR.ASP" 



     XPROYDEF = REQUEST.QUERYSTRING("PRY")

     session("NCliente") = XPROYDEF

     XPROYVPL = REQUEST.QUERYSTRING("VPLAN")

     SESSION("VPLAN") = XPROYVPL

     XFUPD = REQUEST.QUERYSTRING("FUPD")

     SESSION("FUPD") = XFUPD

     XTADV = REQUEST.QUERYSTRING("ADV")

     SESSION("XTADV") = XTADV


 END IF

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

  sHTML = ofv.formHeader("")

  sHTML = sHTML  & "<script language=javascript> "



  sHTML = sHTML  & "function CFEC(F1,F2) { "
  sHTML = sHTML  & " document.all.FSTART.value = F1;"
  sHTML = sHTML  & " document.all.FSTARTX.innerHTML = F2;"
  sHTML = sHTML  & " } "
  sHTML = sHTML  & "</script>"



IF SESSION("PRCF") <> "S" THEN
   XXXX = OWKF.GETPRYCNFG()
END IF


  XTAREA = REQUEST.QUERYSTRING("TSK")

  IF XTAREA <> "" THEN
     SESSION("TAREA") = XTAREA
  ELSE
     XTAREA = SESSION("TAREA")
  END IF



 IF SOPCION = "AT" or SOPCION = "ATT" OR SOPCION = "MT"  OR SOPCION = "ET" THEN
    XUSR = session("usuario")
    XFSTART = CLNG(DATE()) - clng(cdate("01-01-1900"))

    xtarea = SESSION("TAREA")

    IF SOPCION = "ET"  THEN

    elseIF SOPCION = "AT"  OR SOPCION = "MT" THEN
       TADV   = REQUEST.FORM("TADV") 
       PADV   = REQUEST.FORM("PADV") 
       XFSTART = REQUEST.FORM("FSTART")
    else
       TADV   = REQUEST.querystring("TADV") 
       PADV   = REQUEST.querystring("PADV") 
       FSTART = REQUEST.querystring("FSTART") 
       IF FSTART <> "NO" THEN
          StrSql = "UPDATE PROYTAREAS SET FECHASREAL = '" & FSTART & "' WHERE codigo = '" & session("NCliente") & "' AND TAREA = '" & SESSION("TAREA") & "' AND VPLAN = " & SESSION("VPLAN")
          CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
       END IF
    end if

  IF SOPCION = "AT" or SOPCION = "ATT" OR SOPCION = "MT" THEN


    TTIPO = REQUEST.FORM("TTIPO") 

    IF TTIPO = "" THEN TTIPO = 1
    XFDCUENTAGLB = 1

    StrSql = "Select TSKCUENTA From PROYTAREAS where codigo = '" & session("NCliente") & "' AND TAREA = '" & SESSION("TAREA") & "'  AND VPLAN = " & SESSION("VPLAN")
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    IF NOT RS.EOF THEN     XFDCUENTAGLB = RS(0)
    call ofv.cerrarconsulta(rs)

    IF SMODAL = "NEW" THEN

'''''''' agregar afectacion a cuentas ''''''''''''''''''''''''

       XFDCUENTA = OWKF.GENTSKCUENTA(XFSTART,TADV, TTIPO, XFDCUENTAGLB)


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    END IF


    IF SOPCION = "AT" or SOPCION = "ATT"  THEN


       XFDHOY = CLNG(DATE()) - clng(cdate("01-01-1900"))

       STRSQL = "INSERT INTO PROYTAREASADV (CODIGO,TSK,FUPD,TADV,TPRC,USRUPD,TTIPO,FCARGA,CODINPUT) VALUES ("  
       STRSQL = STRSQL & "'" & session("NCliente") & "','" & xtarea & "'," & XFSTART & ","
       STRSQL = STRSQL & TADV & "," & PADV & ",'" & XUSR & "','" & TTIPO & "','" & XFDHOY & "','" & XFDCUENTA & "')" 
       CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)


    ELSE

       STRSQL = "UPDATE PROYTAREASADV SET TADV = '" & TADV & "' where codigo = '" & session("NCliente") & "' AND TSK = '" & SESSION("TAREA") & "'  AND ID = '" & SESSION("XTADV") & "'"
       CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

    END IF


  ELSE

    STRSQL = "DELETE FROM PROYTAREASADV where codigo = '" & session("NCliente") & "' AND TSK = '" & SESSION("TAREA") & "'  AND ID = '" & SESSION("XTADV") & "'"
    CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

  END IF



    response.redirect SESSION("TSKVLV")


 END IF 

 IF (SOPCION <> "VT" and SOPCION <> "VM") THEN sopcion = "VP"

  StrSql = "Select * From PROYTAREAS where codigo = '" & session("NCliente") & "' AND TAREA = '" & SESSION("TAREA") & "'  AND VPLAN = " & SESSION("VPLAN")
''  RESPONSE.WRITE STRSQL & "<BR>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)



    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"
    sImagenA = "<FONT FACE=WINGDINGS COLOR=#AA0000 SIZE=2>n</FONT>"
    sImagenN = "<FONT FACE=WINGDINGS COLOR=#88bbdd SIZE=2>n</FONT>"
    sImageS = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=3>4</FONT></B>"


  IF  sopcion = "VP" OR sopcion = "VT"   OR sopcion = "VM"     OR sopcion = "VD" THEN


      IF SMODAL = "NEW" THEN
         WMODALDSC = sImagen & "&nbsp;<B>GENERAR NUEVA IMPUTACION PARA LA TAREA</B>"
      ELSE
         WMODALDSC = sImagen & "&nbsp;<B>MODIFICAR/ELIMINAR LA IMPUTACION PARA LA TAREA</B>"
      END IF 

      celdas = ofv.GenCelda("","VTI colspan=3","","","","","",WMODALDSC,"si")
      filas  = ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","vt  ","","","","","","TAREA:","si")
      celdas = celdas & ofv.GenCelda("","VTI COLSPAN=2 ","","","","","",rs(4) & "." & RS(5),"si")
    ''''  celdas = celdas & ofv.GenCelda("","TTT","","","","","",RS(5),"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



      celdas = ofv.GenCelda("","vt  ","","","","","","DURACION:","si")
     ''' celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")

        if rs(16) = 1 then sInput = "&nbsp;HORA"
        if rs(16) = 2 then sInput = "&nbsp;DIA"
        if rs(16) = 3 then sInput = "&nbsp;SEMANA"
        if rs(16) = 4 then sInput = "&nbsp;MINUTO"

        IF RS(15) > 1 THEN sinput = sinput & "S" 
   
        sinput = rs(15) & sinput 

        USRTMP     = RS(15)
        USRTMPU    = RS(16)

        SELECT CASE USRTMPU
               CASE 1
                    xagrd = USRTMP

               CASE 4
                    xagrd = USRTMP

               CASE 2

                    xagrd = (USRTMP * CLNG(SESSION("DPP")))
 

               CASE ELSE

                    xagrd = (USRTMP * CLNG(SESSION("DPP")) * CLNG(SESSION("SPP")) )

        END select       


      THTSK = xagrd  

        if rs(16) > 1 AND rs(16) < 4 then sInput = sinput & "&nbsp;(" & THTSK & "&nbsp;HORAS)"




      celdas = celdas & ofv.GenCelda("","Ut COLSPAN=2","","","","","",sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



      XMSG = "NO"

      if sopcion = "VP" then

         XFSR = XFUPD

         sclase = "CC" 

         if clng(XFSR) = 0 then
            sInput1 = "------------"
         else
            sInput1 = cdate(clng(XFSR) + clng(cdate("01-01-1900")) )
            swd = datepart("w",sInput1)
            sInput1 = Fechador(XFSR)
         END IF

         sInput2 = ofv.GenerarInput("FSTART",XFSR,"hidden","Usuarios","10","dtn","")


         celdas = ofv.GenCelda("","UT COLSPAN=3 ","","","","",""," " & sInput2,"si")

         filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


'''''''''''''''''''''''''''


         Celda = ofv.GenCelda("","VT","","25%","","","","FECHA IMPUTACION:","si")

         Celda = Celda & ofv.gencelda("FSTARTX","VT colspan=2 ","","60%","","","",sInput1,"si")

         Celda = Celda & ofv.gencelda("","UT ","","1%","","",""," ","si")

         Fila = ofv.genrow("","","","","","","",Celda,"si")


    'Descripcion
         CELDA = ""

         sanio = datepart("yyyy",date())
         smes = datepart("m",date())

         SMESC = SMES

         if not isnull(XFSR) then

            if clng(XFSR) > 0 then

               smesc = cdate(clng(XFSR) + clng(cdate("01-01-1900")) )
               smesc = datepart("m",smesc)

            end if 


        end if

    IF CA <> "" THEN 
       smesc = request.form("smesc")

       if isnull(smesc) then  SMESC = ca
       if len(smesc) = 0 then  SMESC = ca
       ca =  SMESC
      
    else
       ca = smesc

    end if  

    if smesc = 0 then smesc = smes

       session("ENCA") = smesc

    if cint(smesc) = cint(smes) then
       sanioc = sanio
    elseif cint(smesc) < cint(smes) then
       sanioc = sanio + 1
    else
       sanioc = sanio
    end if

    
      Celda = ofv.GenCelda("","VT","","15%","","","","MES:","si")

    'codigo

     Strsql = "Select * From meses Order By mes"
     sInput = ofv.GenerarCombo("" & sForm & "A","smesc",smesc,Strsql,cn2,1,2,"sb0")
     Celda = Celda & ofv.gencelda("","VT","","45%","","","",sInput,"si")

     Celda = Celda & ofv.GenCelda("","VT","","15%","","","","AÑO:","si")
     
     sinput = sanioc 
     Celda = Celda & ofv.gencelda("","VT","","20%","","","",sInput,"si")



    Formulario = Fila

    filas2 = ofv.GenTabla("","AAA ","","100%","","0","","0","0","","",Formulario,"si")

    XCONTROL = "S" 



      celdas = ofv.GenCelda("","vt COLSPAN=3 style='border-style:SOLID;border-width:1;border-color:#888888;' ","CENTER","","","","",filas2,"si")

      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

''''''''''''''''''''''''''''

       IF SMODAL = "NEW" THEN


             TADV = 0
             PADV = 0

          StrSql = "Select * From PROYTAREASADV where codigo = '" & session("NCliente") & "' AND TSK = '" & SESSION("TAREA") & "' ORDER BY id DESC "
          set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if NOT rsX.eof then
             PADV = RSX(5)
             DO UNTIL RSX.EOF 
                TADV = TADV + RSX(4)
             RSX.MOVENEXT
             LOOP
          END IF

          SINPUTT2 = GETPDV(PADV) 

          if rs(16) = 4 then
             sInputX1 = "&nbsp;MINUTO"
          ELSE
             sInputX1 = "&nbsp;HORA"
          END IF

          IF TADV > 1 THEN sInputX = sInputX1 & "S"


          celdas = ofv.GenCelda("","vt  ","","","","","","TOTAL IMPUTADO:","si")
          celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")

          sInput2 = ofv.GenerarInput("TADVA",TADV,"hidden","Usuarios","10","dtn","") & sInputX
          celdas = celdas & ofv.GenCelda("","vt","","","","","",TADV & sInput2,"si")
          filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

 

          celdas = ofv.GenCelda("","vt  ","","","","","","IMPUTACION ACTUAL:","si")
          celdas = celdas & ofv.GenCelda("","","center","","","","",sImageS,"si")

          sInputX = sInputX1 & "/S"
   
          sInput = ofv.GenerarInput("TADV","1","agregar","Usuarios","10","dtn  STYLE='background-color:#f7eedd;color:#993333;font-weight:bold;' ","") & sInputX
          celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")
          filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


          if bModificar then
             sAccion = "PROYTASKADVMOD.asp?MODAL=" & SMODAL & "&OPCION=VT"
             formulario0 = ofv.GenForm("Usuarios","",sAccion,"","post","",filas,"si")
          else
             formulario0 = ofv.GenForm("","","","","post","",filas,"si")
          end if


       END IF


       IF SMODAL = "MOD" THEN


             TADV = 0
             PADV = 0

          StrSql = "Select * From PROYTAREASADV where codigo = '" & session("NCliente") & "' AND TSK = '" & SESSION("TAREA") & "'  AND ID = '" & SESSION("XTADV") & "'"
          set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if NOT rsX.eof then
             PADV = RSX(5)
             DO UNTIL RSX.EOF 
                TADV = TADV + RSX(4)
             RSX.MOVENEXT
             LOOP
          END IF

          SINPUTT2 = GETPDV(PADV) 

          if rs(16) = 4 then
             sInputX1 = "&nbsp;MINUTO"
          ELSE
             sInputX1 = "&nbsp;HORA"
          END IF

          IF TADV > 1 THEN sInputX = sInputX1 & "S"


          celdas = ofv.GenCelda("","vt  ","","","","","","VALOR ACTUAL:","si")
          celdas = celdas & ofv.GenerarInput("TADVA",TADV,"hidden","Usuarios","10","dtn","")
          celdas = celdas & ofv.GenerarInput("DELOPC","NO","hidden","Usuarios","10","dtn","")
          celdas = celdas & ofv.GenCelda("","","center","","","","",sImageS,"si")

          sInputX =  sInputX1 & "/S"
   
          sInput = TADV & sInputX
          celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")
          filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


          celdas = ofv.GenCelda("","vt  ","","","","","","NUEVO VALOR:","si")
          celdas = celdas & ofv.GenCelda("","","center","","","","",sImageS,"si")

          sInputX = sInputX1 & "/S"
   
          sInput = ofv.GenerarInput("TADV",TADV,"agregar","Usuarios","10","dtn  STYLE='background-color:#f7eedd;color:#993333;font-weight:bold;' ","") & sInputX
          celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")
          filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


          if bModificar then
             sAccion = "PROYTASKADVMOD.asp?MODAL=" & SMODAL & "&OPCION=VM"
             formulario0 = ofv.GenForm("Usuarios","",sAccion,"","post","",filas,"si")
          else
             formulario0 = ofv.GenForm("","","","","post","",filas,"si")
          end if




      END IF


          filas = ofv.GenRow("","","","","","","",formulario0,"si")
          Celda = ofv.GenCelda("","abc colspan=3","center","100%","","","","&nbsp;","si")
          FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")


          FilaS = ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;border-color:#888888;' ","","100%","","0","","0","0","","",filas,"si")


          Celda = ofv.GenCelda("","","center","","","","",FilaS,"si")
          FilaS = ofv.GenRow("","","","","","","",Celda,"si")

          celdas = ""

       IF SMODAL = "MOD" THEN

          sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
          sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
          sinput = sinput  & "none" & chr(34) & ";' class=tt href='javascript:void(0);' onclick=" & chr(34) & "document.Usuarios.DELOPC.value='SI';Usuarios.submit();return false;" & chr(34) & " "
          sinput = sinput  & " TITLE='ELIMINAR LA IMPUTACION' ><b>ELIMINAR</b></a>" 
          celdas = celdas & ofv.GenCelda("","ttR","","10%","","","",sInput,"si")

       END IF

          sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
          sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
          sinput = sinput  & "none" & chr(34) & ";' class=tt href='javascript:void(0);' onclick='Usuarios.submit();return false;' "
          sinput = sinput  & " TITLE='INGRESAR NUEVO VALOR PARA IMPUTACION' ><b>INGRESAR</b></a>" 
          celdas = celdas & ofv.GenCelda("","ttR","","10%","","","",sInput,"si")

          sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
          sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
          sinput = sinput  & "none" & chr(34) & ";' class=tt "
          sinput = sinput  & " href='" & SESSION("TSKVLV") & "'"
          sinput = sinput  & " TITLE='CANCELAR Y VOLVER' ><b>VOLVER</b></a>" 

          celdas = celdas & ofv.GenCelda("","ttR","","10%","","","",sInput,"si")

          celdas = celdas & ofv.GenCelda("","","","80%","","","","&nbsp;","si")

          fila = ofv.GenRow("","","","","25","","",celdas,"si")

          Fila = ofv.GenTabla("","","","","","0","","4","0","","",fila,"si")



   else

     XMAXHS = 8
     XTTIPO = 1

     if rs(16) = 4 then XTTIPO = 4

     StrSql = "Select * From PROYCONFIG "
     set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if not rsx1.eof then 

        XMAXHS = RSX1(10)

     end if

           TADVA   = REQUEST.FORM("TADVA")
           TADV    = REQUEST.FORM("TADV")
           FSTART  = REQUEST.FORM("FSTART")

           SDELOPC = "NO"

           IF SOPCION = "VM" THEN SDELOPC = REQUEST.FORM("DELOPC") & ""

           PADV  = 0


        XACT = "S"

      '''     response.write XMAXHS & "--" & tadv & "--" & rs(16)   & "<br>" 

     IF rs(16) < 4 AND CINT(XMAXHS) < CINT(TADV) THEN

        XACT = "N"
     END IF 

     IF XACT = "S" THEN

        '''   response.write tadva & " -- " & tadv & " -- " & FSTART   & "<br>" 

      if sopcion = "CT" then
         PADV  = 100
         IF TADV = 0 THEN
            TADV = 1
            XTTIPO = 4
         END IF

      else

           IF TADV > 0 AND THTSK > 0 THEN
              PADV = (cint(TADVA) + cint(TADV)) / THTSK * 100
              PADV = CINT(PADV)
              
           END IF  

      end if

        SINPUTT2 = GETPDV(PADV) 

      celdas = ofv.GenCelda("","vt  ","","","","","","FECHA IMPUTACION:","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")
   
      sInput1 = Fechador(FSTART)
      celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput1,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")




          celdas = ofv.GenCelda("","vt  ","","","","","","IMPUTACION:","si")
          celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")

          sInput1 = ofv.GenerarInput("TADV",TADV,"hidden","Usuarios","10","dtn","")
          sInput2 = ofv.GenerarInput("PADV",PADV,"hidden","Usuarios","10","dtn","")
          sInput3 = ofv.GenerarInput("TTIPO",XTTIPO,"hidden","Usuarios","10","dtn","")
          sInput4 = ofv.GenerarInput("FSTART",FSTART,"hidden","Usuarios","10","dtn","")

          if XTTIPO = 4 then
             sInputN = "&nbsp;MINUTO"
          ELSE
             sInputN = "&nbsp;HORA"
          END IF

          sInput = sInputN


          IF TADV > 1 THEN sInput = sInput & "S"

   
          sInput = "<B>" & TADV & sInput & "</B>" & sInput1 & sInput2 & sInput3 & sInput4
          celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")
          filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      if sopcion <> "CT" AND sopcion <> "VM"  AND sopcion <> "VD" then

          celdas = ofv.GenCelda("","vt  ","","","","","","TOTAL IMPUTADO:","si")
          celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")

          sInput = "&nbsp;HORA"

          IF (cint(TADVA) + cint(TADV)) > 1 THEN sInput = sInput & "S"

          sInput = "<B>" & (cint(TADVA) + cint(TADV)) & sInput & "</B>"
          celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")
          filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")




          celdas = ofv.GenCelda("","vt  ","","","","","","PORCENTAJE:","si")
          celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")
   
          sInput = "<big><B>" & PADV & "%</B></big>&nbsp;&nbsp;" & sinputt2
          celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")
          filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      END IF

          celdas = ofv.GenCelda("","UT COLSPAN=3  ","CENTER","","","","","&nbsp;","si")
          filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

          IF SDELOPC = "SI" THEN
             XMSG = "<BR><FONT SIZE=2 COLOR='#993333' ><B>*** CONFIRMA LA ELIMINACION DE ESTA IMPUTACION PARA LA TAREA ?? ***</B></FONT>"
          ELSE
             XMSG = "<BR><FONT SIZE=2 COLOR='#993333' ><B>*** CONFIRMA LA IMPUTACION PARA LA TAREA ?? ***</B></FONT>"
          END IF 

          celdas = ofv.GenCelda("","vt  COLSPAN=3 style='border-style:SOLID;border-width:1;border-color:#993333;' ","CENTER","","","","",XMSG,"si")
          filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

         ' END IF

    IF SMODAL = "NEW" THEN
       WOPCION = "AT"
    ELSE
       WOPCION = "MT"
       IF SDELOPC = "SI" THEN WOPCION = "ET"
    END IF



    if bModificar then
      sAccion = "PROYTASKADVMOD.asp?MODAL=" & SMODAL & "&OPCION=" & WOPCION
      formulario0 = ofv.GenForm("Usuarios","",sAccion,"","post","",filas,"si")
    else
      formulario0 = ofv.GenForm("","","","","post","",filas,"si")
    end if


   ELSE

      celdas = ofv.GenCelda("","vt  ","","","","","","FECHA IMPUTACION:","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")
   
      sInput1 = Fechador(FSTART)
      celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput1,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")




          celdas = ofv.GenCelda("","vt  ","","","","","","IMPUTADO:","si")
          celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")


             sInput = "&nbsp;HORA"


          IF TADV > 1 THEN sInput = sInput & "S"

   
          sInput = "<B>" & TADV & sInput & "</B>"
          celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")
          filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


             XMSG = "<BR><FONT COLOR=#993333 SIZE=2 ><B>*** IMPUTACION INCORRECTA - EXCEDE EL MAXIMO AUTORIZADO DE " & XMAXHS & "hs. !!! ***</B></FONT><BR>"

             celdas = ofv.GenCelda("","vt  COLSPAN=3 style='border-style:SOLID;border-width:1;border-color:#993333;' ","CENTER","","","","",XMSG,"si")
             filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    '  celdas = ofv.GenCelda("","AAA COLSPAN=3 style='border-top-style:ridge;border-top-width:2;border-top-color:#f7aa22;' ","CENTER","","","","","&nbsp;","si")
     ' filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      formulario0 = filas


   END IF 


    filas = ofv.GenRow("","","","","","","",formulario0,"si")
    Celda = ofv.GenCelda("","abc colspan=3","center","100%","","","","&nbsp;","si")
    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")


      FilaS = ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;border-color:#888888;' ","","100%","","0","","0","0","","",filas,"si")


    Celda = ofv.GenCelda("","","center","","","","",FilaS,"si")
    FilaS = ofv.GenRow("","","","","","","",Celda,"si")

   IF XACT = "S" THEN


     sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' class=tt href='javascript:void(0);' onclick=" & chr(34) & "this.style.display='none';Usuarios.submit();return false;" & chr(34) & " "
     sinput = sinput  & " ><b>confirmar</b></a>" 
     celdas = ofv.GenCelda("","ttR","","10%","","","",sInput,"si")

     sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' class=tt "



     sinput = sinput  & " href='" & SESSION("TSKVLV") & "'"


     sinput = sinput  & " ><b>VOLVER</b></a>" 
     celdas = celdas & ofv.GenCelda("","ttR","","10%","","","",sInput,"si")

   ELSE

     sinput = "&nbsp;"
     celdas = ofv.GenCelda("","VT","","10%","","","",sInput,"si")

     sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' class=tt href='PROYTASKADVMOD.asp?MODAL=" & SMODAL & "' "
     sinput = sinput  & " ><b>VOLVER</b></a>" 
     celdas = celdas & ofv.GenCelda("","ttR","","10%","","","",sInput,"si")

   END IF

     celdas = celdas & ofv.GenCelda("","","","80%","","","","&nbsp;","si")

     fila = ofv.GenRow("","","","","","","",celdas,"si")

     Fila = ofv.GenTabla("","","","","","0","","4","0","","",fila,"si")




      end if










    Celda = ofv.GenCelda("","","center","","","","",Fila,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","","","90%","","0","","0","0","","",filas,"si")
    sHTML = sHTML & "</center><br><br>"




    end if



  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML

FUNCTION GETPDV(XPDV)

 XPDI = 0

 XTA = CINT(XPDV / 10)

 FOR XPDI=1 TO XTA
     GETPDV = GETPDV & SIMAGENA
 NEXT


 FOR XPDI=XTA + 1  TO 10
     GETPDV = GETPDV & SIMAGENN
 NEXT

END FUNCTION


Function Fechador(wkfecha)


    
      fec = cdate(wkfecha + clng(cdate("01-01-1900")))
      ndia = ofv.cdia(Weekday(fec))
      nMes = ofv.cmes(Month(fec))
    
      Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)


End Function



%>