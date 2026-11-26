<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


 sopcion = request.querystring("opcion")


 IF SOPCION = "PCA" OR SOPCION = "PCU" THEN

     IF SOPCION = "PCA" THEN

        SESSION("TSKVLV") = "PRYASGNUSR.ASP" 

     ELSE

        SESSION("TSKVLV") = "PRYADMULTIMPUT.ASP" 

     END IF 

     XPROYDEF = REQUEST.QUERYSTRING("PRY")

     session("NCliente") = XPROYDEF

     XPROYVPL = REQUEST.QUERYSTRING("VPLAN")

     SESSION("VPLAN") = XPROYVPL


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




    IF  SESSION("TSKVLV") <> "" THEN

        sHTML = ofv.formHeader("")

    ELSE

        sHTML = ofv.MENUHeader("","")

    END IF


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

 IF SOPCION = "LT" THEN
    XUSR = session("usuario")
    XFSTART = CLNG(DATE()) - clng(cdate("01-01-1900"))



       XXXX = OWKF.TSKLNCHERSV(session("NCliente"),SESSION("TAREA"),"S")



    response.redirect "proyest.asp"

 END IF 


 IF SOPCION = "AT" or SOPCION = "ATT" THEN
    XUSR = session("usuario")
    XFSTART = CLNG(DATE()) - clng(cdate("01-01-1900"))

    xtarea = SESSION("TAREA")

 IF SOPCION = "AT"  THEN
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

    TTIPO = REQUEST.FORM("TTIPO") 

    IF TTIPO = "" THEN TTIPO = 1
    XFDCUENTAGLB = 1

    StrSql = "Select TSKCUENTA From PROYTAREAS where codigo = '" & session("NCliente") & "' AND TAREA = '" & SESSION("TAREA") & "'  AND VPLAN = " & SESSION("VPLAN")
    set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    IF NOT RS.EOF THEN     XFDCUENTAGLB = RS(0)
    call ofv.cerrarconsulta(rs)


'''''''' agregar afectacion a cuentas ''''''''''''''''''''''''

   XFDCUENTA = OWKF.GENTSKCUENTA(XFSTART,TADV, TTIPO, XFDCUENTAGLB)


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    XFDHOY = CLNG(DATE()) - clng(cdate("01-01-1900"))




    STRSQL = "INSERT INTO PROYTAREASADV (CODIGO,TSK,FUPD,TADV,TPRC,USRUPD,TTIPO,FCARGA,CODINPUT) VALUES ("  
    STRSQL = STRSQL & "'" & session("NCliente") & "','" & xtarea & "'," & XFSTART & ","
    STRSQL = STRSQL & TADV & "," & PADV & ",'" & XUSR & "','" & TTIPO & "','" & XFDHOY & "','" & XFDCUENTA & "')" 
    CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)




    IF PADV > 99 THEN
       XXXX = OWKF.TSKLNCHER(PADV,session("NCliente"),xtarea,"S")
    END IF

 IF SOPCION = "AT"  THEN


    IF  SESSION("TSKVLV") <> "" THEN

        response.redirect SESSION("TSKVLV")

    ELSE

       response.redirect "proyest.asp"

    END IF



 else
    response.redirect "PROYTASKPROP.asp"
 end if


 END IF 

 IF (SOPCION <> "VT" and SOPCION <> "CT") THEN sopcion = "VP"

  StrSql = "Select * From PROYTAREAS where codigo = '" & session("NCliente") & "' AND TAREA = '" & SESSION("TAREA") & "'  AND VPLAN = " & SESSION("VPLAN")
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)



    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"
    sImagenA = "<FONT FACE=WINGDINGS COLOR=#AA0000 SIZE=2>n</FONT>"
    sImagenN = "<FONT FACE=WINGDINGS COLOR=#88bbdd SIZE=2>n</FONT>"
    sImageS = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=3>4</FONT></B>"


  IF  sopcion = "VP" OR sopcion = "VT"  OR sopcion = "CT" THEN

      celdas = ofv.GenCelda("","TTT colspan=3","","","","","",sImagen & "&nbsp;<B>ACTUALIZAR AVANCE DE LA TAREA</B>","si")
      filas  = ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","AA COLSPAN=3","CENTER","","","","","&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","vt  ","","","","","","TAREA:","si")
      celdas = celdas & ofv.GenCelda("","TTT COLSPAN=2 ","","","","","",rs(4) & "." & RS(5),"si")
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


   ''   celdas = ofv.GenCelda("","AAA COLSPAN=3 style='border-top-style:ridge;border-top-width:2;border-top-color:#f7aa22;' ","CENTER","","","","","&nbsp;","si")
    ''  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")




      XMSG = "NO"

      if sopcion = "VP" then

         XFSR = CLNG(DATE()) - clng(cdate("01-01-1900"))

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
     sInput = ofv.GenerarCombo("document." & sForm & "A","smesc",smesc,Strsql,cn2,1,2,"sb0")
     Celda = Celda & ofv.gencelda("","VT","","45%","","","",sInput,"si")

     Celda = Celda & ofv.GenCelda("","VT","","15%","","","","AÑO:","si")
     
     sinput = sanioc 
     Celda = Celda & ofv.gencelda("","VT","","20%","","","",sInput,"si")



    Formulario = Fila

    filas2 = ofv.GenTabla("","AAA ","","100%","","0","","0","0","","",Formulario,"si")

  XCONTROL = "N" 

  StrSql = "Select * From PROYRESP where codigo = '" & session("NCliente") & "' AND USRAFC = '" & SESSION("USUARIO") & "' AND TIPO = 'AN' "
  set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  if rsx1.eof then

     StrSql = "Select * From PROYRESP where codigo = '" & session("NCliente") & "' AND USRAFC = '" & SESSION("USUARIO") & "' "
     set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  end if
  if NOT rsX1.eof then

     IF RSX1(11) THEN XCONTROL = "S" 

  END IF

  IF XCONTROL = "S" THEN

     Celda = ofv.GenCelda("","VT","","30%","","","","FORZAR FIN DE LA TAREA ","si")

'''PROYTASKADV.asp?OPCION=VT


     sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
'''     sinput = sinput  & "none" & chr(34) & ";' class=tt href='javascript:void(0);' onclick='Usuarios.submit();return false;' "
     sinput = sinput  & "none" & chr(34) & ";' class=tt href='PROYTASKADV.asp?OPCION=CT'  "
     sinput = sinput  & " TITLE='FORZAR FIN DE LA TAREA' ><b>" & sImageS & "</b></a>" 

     Celda = Celda & ofv.gencelda("","VT","","70%","","","",sInput,"si")

     filas2 =  filas2 & ofv.GenTabla("","AAA ","","100%","","0","","0","0","","",Celda,"si")


    dim afech(40,1)
    dim afechs(42,1)

    CELDA = ""

    sfecds = "01-" & cstr(smesc) & "-" & cstr(sanioc)
    sfecd = clng(dateserial(sanioc,smesc,1))

    sfecd = sfecd - 4

       sInput1 = cdate(clng(sfecd) )

    sanio1 = datepart("yyyy",sInput1)
    smes1 = datepart("m",sInput1)
    sdia1 = datepart("d",sInput1)

    swd = datepart("w",dateserial(sanio1,smes1,sdia1))

   ''' response.write "f " & sanio1 & "-" & smes1 & "-" & sdia1 & " d " & swd
     
    sfechs = "28-" & cstr(smesc) & "-" & cstr(sanioc)
    sfech = clng(dateserial(sanioc,smesc,28))


    yy = 0

    for ii = 1 to 3
        sfecha = sfech + ii
        sfecha = cdate(clng(sfecha))
        smess = datepart("m",sfecha)
        if cint(smesc) <> cint(smess) then
           exit for
        else
           yy = ii
        end if

    next  

    sfech = sfech + yy

  '  srest = 0
    srest = clng(cdate("01-01-1900"))
    topefech = sfech - sfecd 

    yy = 0



    wswd = swd

    for ii = 0 to topefech
        
        afech(ii+1,0) = wswd
        wswd = wswd + 1
        if wswd > 7 then wswd = 1        

        afech(ii+1,1) = sfecd + ii


    next


    for ii = 1 to (cint(afech(1,0)) - 1)
        afechs(ii,0) = "NO"
        afechs(ii,1) = 0
    next

    yy = 1


    for ii = ii to (ii + topefecH )
        sInput1      = cdate(clng(afech(yy,1)) )
        sdia1        = datepart("d",sInput1)
        afechs(ii,0) = sdia1
        afechs(ii,1) = afech(yy,1)
        yy = yy + 1
    next


    for ii = ii to 42
        afechs(ii,0) = "NO"
        afechs(ii,1) = 0
    next

    CELDA = ""
    Celda = Celda & ofv.GenCelda("","Vt style='border-style:outset;border-width:2;'","CENTER","10%","","","","D","si")
    Celda = Celda & ofv.GenCelda("","Vt style='border-style:outset;border-width:2;'","CENTER","10%","","","","L","si")
    Celda = Celda & ofv.GenCelda("","Vt style='border-style:outset;border-width:2;'","CENTER","10%","","","","M","si")
    Celda = Celda & ofv.GenCelda("","Vt style='border-style:outset;border-width:2;'","CENTER","10%","","","","M","si")
    Celda = Celda & ofv.GenCelda("","Vt style='border-style:outset;border-width:2;'","CENTER","10%","","","","J","si")
    Celda = Celda & ofv.GenCelda("","Vt style='border-style:outset;border-width:2;'","CENTER","10%","","","","V","si")
    Celda = Celda & ofv.GenCelda("","Vt style='border-style:outset;border-width:2;'","CENTER","10%","","","","S","si")

    Fila = ofv.genrow("","","","","","","",Celda,"si")

    CELDA = ""
    yy = 0

    xcolor = "background-color:#f7aa22;"

    for ii = 1 to 43

        if yy = 7 then
           yy = 1
           Fila = fila & ofv.genrow("","","","","","","",Celda,"si")
           CELDA = ""
        else
           yy = yy + 1 
        end if

        if ii > 42 then exit for  

        if afechs(ii,0) = "1" then xcolor = ""

        if yy = 1 or yy = 7 then
           sclase = "tt3 style='border-style:outset;border-width:2;'" 
        else
           sclase = "TT style='border-style:outset;border-width:2;" & xcolor & "'" 
        end if 

        if afechs(ii,0) = "NO" then
           sinput = "&nbsp;"
           sclase = "CC " 
           Celda = Celda & ofv.GenCelda("","VT","CENTER","10%","","","",sinput,"si")
        else
           sinput1 = afechs(ii,0)

           sinput1 = cstr(sinput1)

           if len(sinput1) = 1 then sinput1 = "0" & sinput1 

           sinput1 = " " & sinput1 & " " 

           sAccion = " "

           sf = "CFEC('" & (afechs(ii,1) - srest) & "','"  & Fechador((afechs(ii,1) - srest)) & "');"

           t1 = ofv.GenerarInput("bvo1",sInput1,"BUTTON","","10",sclase,sf)
     '''      t1 = t1 & ofv.GenerarInput(rs(5).name,(afechs(ii,1) - srest),"hidden","","2","dt","")
           t1 = ofv.GenCelda("","VT","center","10%","","","",t1,"si")
      '''     t1 = ofv.GenForm("feac" & ii,"",sAccion,"","post","",t1,"si") 
           Celda = Celda & t1

        end if


    next



   filas2 = "<br>" & filas2 & "<br>" & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;border-color:#888888;'","","50%","","0","","0","0","","",FILA,"si")

  END IF

  ' filas2 = filas2 & "<br>" 

      celdas = ofv.GenCelda("","vt COLSPAN=3 style='border-style:SOLID;border-width:1;border-color:#888888;' ","CENTER","","","","",filas2,"si")

      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

''''''''''''''''''''''''''''




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
             sInputX = "&nbsp;MINUTO"
          ELSE
             sInputX = "&nbsp;HORA"
          END IF

          IF TADV > 1 THEN sInputX = sInputX & "S"


          celdas = ofv.GenCelda("","vt  ","","","","","","TOTAL IMPUTADO:","si")
          celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")

          sInput2 = ofv.GenerarInput("TADVA",TADV,"hidden","Usuarios","10","dtn","") & sInputX
          celdas = celdas & ofv.GenCelda("","vt","","","","","",TADV & sInput2,"si")
          filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

          celdas = ofv.GenCelda("","vt  ","","","","","","PORCENTAJE:","si")
          celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")
   
          sInput = "<big>" & PADV & "%</big>&nbsp;&nbsp;" & sinputt2
          celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")
          filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    '  celdas = ofv.GenCelda("","AAA COLSPAN=3 style='border-top-style:ridge;border-top-width:2;border-top-color:#f7aa22;' ","CENTER","","","","","&nbsp;","si")
    '  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


          celdas = ofv.GenCelda("","vt  ","","","","","","AVANCE:","si")
          celdas = celdas & ofv.GenCelda("","","center","","","","",sImageS,"si")

          sInputX = "&nbsp;HORA/S"
   
          sInput = ofv.GenerarInput("TADV","1","agregar","Usuarios","10","dtn  STYLE='background-color:#f7eedd;color:#993333;font-weight:bold;' ","") & sInputX
          celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")
          filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


     ' celdas = ofv.GenCelda("","AAA COLSPAN=3 style='border-top-style:ridge;border-top-width:2;border-top-color:#f7aa22;' ","CENTER","","","","","&nbsp;","si")
     ' filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    if bModificar then
      sAccion = "PROYTASKADV.asp?OPCION=VT"
      formulario0 = ofv.GenForm("Usuarios","",sAccion,"","post","",filas,"si")
    else
      formulario0 = ofv.GenForm("","","","","post","",filas,"si")
    end if


    filas = ofv.GenRow("","","","","","","",formulario0,"si")
    Celda = ofv.GenCelda("","abc colspan=3","center","100%","","","","&nbsp;","si")
    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")


      FilaS = ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;border-color:#888888;' ","","100%","","0","","0","0","","",filas,"si")


    Celda = ofv.GenCelda("","","center","","","","",FilaS,"si")
    FilaS = ofv.GenRow("","","","","","","",Celda,"si")


     sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' class=tt href='javascript:void(0);' onclick='Usuarios.submit();return false;' "
     sinput = sinput  & " ><b>ingresar</b></a>" 
     celdas = ofv.GenCelda("","ttR","","10%","","","",sInput,"si")

     sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' class=tt "

  IF  SESSION("TSKVLV") <> "" THEN

     sinput = sinput  & " href='" & SESSION("TSKVLV") & "'"

  ELSE

     sinput = sinput  & " href='PROYest.asp'"

  END IF


     sinput = sinput  & " ><b>VOLVER</b></a>" 
     celdas = celdas & ofv.GenCelda("","ttR","","10%","","","",sInput,"si")

     celdas = celdas & ofv.GenCelda("","","","80%","","","","&nbsp;","si")

     fila = ofv.GenRow("","","","","","","",celdas,"si")

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

           IF TADV > 0 THEN
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




          celdas = ofv.GenCelda("","vt  ","","","","","","AVANCE:","si")
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

      if sopcion <> "CT" then

          celdas = ofv.GenCelda("","vt  ","","","","","","TOTAL IMPUTADO:","si")
          celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")

          sInput = "&nbsp;HORA"

          IF (cint(TADVA) + cint(TADV)) > 1 THEN sInput = sInput & "S"

          sInput = "<B>" & (cint(TADVA) + cint(TADV)) & sInput & "</B>"
          celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")
          filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

     END IF


          celdas = ofv.GenCelda("","vt  ","","","","","","PORCENTAJE:","si")
          celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")
   
          sInput = "<big><B>" & PADV & "%</B></big>&nbsp;&nbsp;" & sinputt2
          celdas = celdas & ofv.GenCelda("","vt","","","","","",sInput,"si")
          filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

          IF PADV > 99 THEN 

             celdas = ofv.GenCelda("","UT COLSPAN=3  ","CENTER","","","","","&nbsp;","si")
             filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

             IF sopcion = "CT" THEN
                XMSG = "<BR><FONT SIZE=2 COLOR='#993333' ><B>***         FINALIZACION FORZADA DE LA TAREA          ***<BR>*** AL CONFIRMAR LA ACTUALIZACION FINALIZARA LA TAREA ***</B></FONT>"
             ELSE
                XMSG = "<BR><FONT SIZE=2 COLOR='#993333' ><B>*** AL CONFIRMAR LA ACTUALIZACION FINALIZARA LA TAREA ***</B></FONT>"
             END IF

             celdas = ofv.GenCelda("","vt  COLSPAN=3 style='border-style:SOLID;border-width:1;border-color:#993333;' ","CENTER","","","","",XMSG,"si")
             filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

          END IF


    '  celdas = ofv.GenCelda("","AAA COLSPAN=3 style='border-top-style:ridge;border-top-width:2;border-top-color:#f7aa22;' ","CENTER","","","","","&nbsp;","si")
     ' filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


    if bModificar then
      sAccion = "PROYTASKADV.asp?OPCION=AT"
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

  IF  SESSION("TSKVLV") <> "" THEN

     sinput = sinput  & " href='" & SESSION("TSKVLV") & "'"

  ELSE

     sinput = sinput  & " href='PROYest.asp'"

  END IF

     sinput = sinput  & " ><b>VOLVER</b></a>" 
     celdas = celdas & ofv.GenCelda("","ttR","","10%","","","",sInput,"si")

   ELSE

     sinput = "&nbsp;"
     celdas = ofv.GenCelda("","VT","","10%","","","",sInput,"si")

     sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' class=tt href=PROYTASKADV.asp"
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