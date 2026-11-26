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


     session("ENC7") = ""
     session("ENCA") = ""


 sopcion = request.querystring("opcion")

 if sopcion = "SI" THEN

 END IF

 if sopcion = "CF" THEN
        STRSQL = "UPDATE PROYPRTY SET CONF = 1 WHERE CODIGO = '" & session("NCliente") & "'  "
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

        OWKF.GENINITPRY()

 END IF

 if sopcion = "AT" THEN
    XETP = REQUEST.QUERYSTRING("ETP")
    StrSql = "Select MAX(ORDEN) From PROYTAREAS where codigo = '" & session("NCliente")
    StrSql = StrSql & "' AND ETAPA = '" & XETP & "' AND VPLAN = " & SESSION("VPLAN")
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
    call ofv.cerrarconsulta(rsX)
        Xtsk = OWKF.GENPRYTSK(session("NCliente"),SESSION("VPLAN"),XORDEN,XETP,"NUEVA TAREA","0","0","1","1","0","1")



 END IF

 if sopcion = "ST" THEN
    XTSK    = REQUEST.QUERYSTRING("TSK")
    XTSKANT = REQUEST.QUERYSTRING("TSKANT")
    StrSql = "Select ORDEN From PROYTAREAS where codigo = '" & session("NCliente")
    StrSql = StrSql & "' AND TAREA = '" & XTSK & "' AND VPLAN = " & SESSION("VPLAN")
    set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
'       RESPONSE.WRITE STRSQL & "<BR>"
    if NOT rsX.eof then
       XORDEN = RSX(0)
       XNORDN = XORDEN - 10  
       STRSQL = "UPDATE PROYTAREAS SET ORDEN = '" & XNORDN & "' where codigo = '" & session("NCliente")
       StrSql = StrSql & "' AND TAREA = '" & XTSK & "' AND VPLAN = " & SESSION("VPLAN")
       CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
'       RESPONSE.WRITE STRSQL & "<BR>"
       STRSQL = "UPDATE PROYTAREAS SET ORDEN = '" & XORDEN & "' where codigo = '" & session("NCliente")
       StrSql = StrSql & "' AND TAREA = '" & XTSKANT & "' AND VPLAN = " & SESSION("VPLAN")
       CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
 '      RESPONSE.WRITE STRSQL & "<BR>"
    END IF
    call ofv.cerrarconsulta(rsX)




 END IF

 if sopcion = "SE" THEN
    XETP    = REQUEST.QUERYSTRING("ETP")
    XETPANT = REQUEST.QUERYSTRING("ETPANT")
    StrSql = "Select ORDEN From PROYETAPA where codigo = '" & session("NCliente")
    StrSql = StrSql & "' AND ETAPA = '" & XETP & "' AND VPLAN = " & SESSION("VPLAN")
    set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
 '      RESPONSE.WRITE STRSQL & "<BR>"
    if NOT rsX.eof then
       XORDEN = RSX(0)
       XNORDN = XORDEN - 10  
       STRSQL = "UPDATE PROYETAPA SET ORDEN = '" & XNORDN & "' WHERE codigo = '" & session("NCliente")
       StrSql = StrSql & "' AND ETAPA = '" & XETP & "' AND VPLAN = " & SESSION("VPLAN")
       CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
 '      RESPONSE.WRITE STRSQL & "<BR>"
       STRSQL = "UPDATE PROYETAPA SET ORDEN = '" & XORDEN & "' WHERE codigo = '" & session("NCliente")
       StrSql = StrSql & "' AND ETAPA = '" & XETPANT & "' AND VPLAN = " & SESSION("VPLAN")
       CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)
 '      RESPONSE.WRITE STRSQL & "<BR>"
    END IF
    call ofv.cerrarconsulta(rsX)




 END IF

 if sopcion = "AE" THEN
        STRSQL = "UPDATE PROYPRTY SET CONF = 1 WHERE CODIGO = '" & session("NCliente") & "'  "
        CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

        OWKF.GENINITPRY()

 END IF

 sopcion = "VT"


  StrSql = "Select * From PROYprty where codigo = '" & session("NCliente") & "'  "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then

     sopcion = "VP"

     OWKF.GENPROP()

     StrSql = "Select * From PROYprty where codigo = '" & session("NCliente") & "'  "
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  ELSE
     IF RS(5) = 0 THEN
        sopcion = "VP"
     END IF



  END IF




    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"



  IF  sopcion = "VP" THEN

      celdas = ofv.GenCelda("","ttt colspan=3","","","","","",sImagen & "&nbsp;ESTRUCTURA","si")
      filas  = ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","UT COLSPAN=3","RIGHT","","","","","PLAN VIGENTE:&nbsp;" & SESSION("VPLAND")& "&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      celdas = ofv.GenCelda("","TT  ","center","","","","","TIPO","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")
    if bModificar  then
        sInput = "<font class=at><select name=" & rs(3).name & " onchange=submit()>"
        sInput = sInput & "<option value=0 "
        if rs(3) = 0 then sInput = sInput & "selected"
        sInput = sInput & " >ESQUEMA UNICO<option value=1 "
        if rs(3) = 1 then sInput = sInput & "selected"
        sInput = sInput & " >POR BLOQUES/ETAPAS/FASES</select></font>"
    else
        if rs(3) = 0 then sInput = "ESQUEMA UNICO"
        if rs(3) = 1 then sInput = "POR BLOQUES/ETAPAS/FASES"
        sInput = ofv.GenerarInput("",sInput,"readonly","","40","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","CC","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")




    if RS(3) > 0 then
      celdas = ofv.GenCelda("","TT  ","center","","","","","ESTRUCTURA","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")
      if bModificar  then
        sInput = "<font class=at><select name=" & rs(4).name & " onchange=submit()>"
        sInput = sInput & "<option value=0 "
        if rs(4) = 0 then sInput = sInput & "selected"
        sInput = sInput & " >BLOQUES<option value=1 "
        if rs(4) = 1 then sInput = sInput & "selected"
        sInput = sInput & " >ETAPAS<option value=2 "
        if rs(4) = 2 then sInput = sInput & "selected"
        sInput = sInput & " >FASES</select></font>"
      else
        if rs(4) = 0 then sInput = "BLOQUES"
        if rs(4) = 1 then sInput = "ETAPAS"
        if rs(4) = 2 then sInput = "FASES"
        sInput = ofv.GenerarInput("",sInput,"readonly","","40","dt","")
      end if
      celdas = celdas & ofv.GenCelda("","CC","","","","","",sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","TT  ","center","","","","","CANTIDAD<BR>(sin contar INICIAL y FINAL)","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","","&nbsp;","si")
      if bModificar then
         sInput = ofv.GenerarInput(rs(2).name,rs(2),"text","Usuarios","10","dtn","")
      else
         sInput = ofv.GenerarInput("",rs(2),"readonly","","10","dtn","")
      end if
      celdas = celdas & ofv.GenCelda("","CC","","","","","",sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    end if






    if bModificar then
      sAccion = "asp/Actualizar.asp?FNAME=" & rs(0) & "&volver=" & vuelta & "&Tabla=PROYPRTY"
      formulario0 = ofv.GenForm("Usuarios","",sAccion,"","post","",filas,"si")
    else
      formulario0 = ofv.GenForm("","","","","post","",filas,"si")
    end if







    filas = ofv.GenRow("","","","","","","",formulario0,"si")
    Celda = ofv.GenCelda("","abc colspan=3","center","100%","","","","&nbsp;","si")
    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")


      FilaS = ofv.GenTabla("","","","100%","","0","","0","0","","",filas,"si")


    Celda = ofv.GenCelda("","","center","","","","",FilaS,"si")
    FilaS = ofv.GenRow("","","","","","","",Celda,"si")

     sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' class=tt href=PROYtask.asp?opcion=CF"
     sinput = sinput  & " ><b>&nbsp;configurar&nbsp;proyecto&nbsp;</b></a>" 
     celdas = ofv.GenCelda("","tt","","10%","","","",sInput,"si")

     sinput = "<a style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' class=tt href=PROYPROP.asp"
     sinput = sinput  & " ><b>VOLVER</b></a>" 
     celdas = celdas & ofv.GenCelda("","tt","","10%","","","",sInput,"si")
     celdas = celdas & ofv.GenCelda("","","","80%","","","","&nbsp;","si")

     fila = ofv.GenRow("","","","","","","",celdas,"si")

     Fila = ofv.GenTabla("","","","","","0","","4","0","","",fila,"si")

    else


    sImagen0 = "<FONT FACE=WINGDINGS SIZE=2>F</FONT>"
    sImagen1 = "<FONT FACE=WINGDINGS SIZE=2>ñ</FONT>"
    sImagen1N = "<FONT FACE=WINGDINGS COLOR=#FFFFFF SIZE=2>l</FONT>"
    sImagen1S = "<B><FONT FACE=WINGDINGS COLOR=#FFFFFF SIZE=2>&#0233;</FONT></B>"
    sImagen2 = "<FONT FACE=VERDANA SIZE=2>+</FONT>"
    sImagen3 = "<FONT FACE=WINGDINGS SIZE=2>2</FONT>"
    sImagen4s = "<B1><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=2>&#0128;</FONT></B1>"
    sImagen4n = "<B1><FONT FACE=WEBDINGS COLOR=#993333 SIZE=2>x</FONT></B1>"
    sImagen6S = "<B1><FONT FACE=WINGDINGS COLOR=#008000 SIZE=2>&#0252;</FONT></B1>"
    sImagen6SS = "<B1><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=2>&#0246;</FONT></B1>"
    sImagen6N = "<B1><FONT FACE=WINGDINGS COLOR=#AA0000 SIZE=3>3</FONT></B1>"
    sImagen6OT = "<B1><FONT FACE=WEBDINGS COLOR=#008000 SIZE=3>&#0254;</FONT></B1>"
    sImagen6T = "<B1><FONT FACE=WEBDINGS COLOR=#336699 SIZE=3>q</FONT></B1>"
    sImagenSTN = "<B><FONT FACE=WINGDINGS COLOR=#336699 SIZE=3>6</FONT></B>"
    sImagenSTC = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0184;</FONT></B>"
    sImagenE0 = "<B><FONT FACE=WINGDINGS COLOR=#6699CC SIZE=4>6</FONT></B>"
    sImagenE1 = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=4>&#0232;</FONT></B>"
    sImagenE2 = "<B><FONT FACE=WINGDINGS COLOR=#000080 SIZE=4>&#0251;</FONT></B>"



    XFSTS = 0 
    XFHOY = CLNG(DATE()) - clng(cdate("01-01-1900")) 

    StrSql = "Select * From PROYecto where codigo = '" & session("NCliente") & "'  "
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT rs2.eof then
       xnompry = rs2(2)
       if rs2(4) > 0 then
          XFSTART = rs2(4)
       else
          XFSTART = rs2(3)
       end if
       XFSTS  = rs2(11) 
       XFFCH  = RS2(18)
       XVPLAN = RS2(17)
       XUSRRESP = RS2(14)
       IF XFFCH > 0 THEN XXX = OWKF.ADMFERIADOS("C",XFSTART)

    end if

    IF XFSTS > 1 THEN

       IF XFSTS = 15 OR XFSTS = 16 THEN
          IF CLNG(SESSION("VPLAN")) = CLNG(XVPLAN) THEN
             bAgregar   = FALSE
             bModificar = FALSE
             bSuprimir  = FALSE
          END IF
       ELSE
    '      bAgregar   = FALSE
    '      bModificar = FALSE
    '      bSuprimir  = FALSE
       END IF
    END IF




      XFCURR = XFSTART

      celdas = ofv.GenCelda("","ttt colspan=4","","","","","",sImagen & "&nbsp;ADMINISTRADOR DE TAREAS","si")
      filas  = ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","UT COLSPAN=4","RIGHT","","","","","PLAN VIGENTE:&nbsp;" & SESSION("VPLAND") & "&nbsp;","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      if XFSTS > 1 then
         if XFSTS = 13 then
            XFDSCSTS = "<BR>&nbsp;" & sImagenSTC & "&nbsp;<SMALLER>FINALIZADO</SMALLER>"
         ELSE
            XFDSCSTS = "<BR>&nbsp;" & sImagenSTC & "&nbsp;<SMALLER>EN&nbsp;CURSO</SMALLER>"
         END IF
      ELSE
         XFDSCSTS = "<BR>&nbsp;" & sImagenSTN & "&nbsp;<SMALLER>NO&nbsp;INICIADO</SMALLER>"
      END IF

      celdas = ofv.GenCelda("","VT rowspan=2 valign=top ","","60%","","","","PROYECTO:&nbsp;" & xnompry & XFDSCSTS,"si")


        if rs(3) = 0 then sInput = "ESQUEMA UNICO"
        if rs(3) = 1 then sInput = "POR BLOQUES/ETAPAS/FASES"

     celdas = celdas & ofv.GenCelda("","vt","","40%","","","","TIPO:&nbsp;" & sInput,"si")
     filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")




    if rs(4) = 0 then sInput = "BLOQUE"
    if rs(4) = 1 then sInput = "ETAPA"
    if rs(4) = 2 then sInput = "FASE"
   
    xtipo = rs(3)
    xestr = sInput

    if RS(3) > 0 then

      sInput = (CINT(rs(2)) + 2 ) & "&nbsp;" & xestr & "S"
      celdas = ofv.GenCelda("","vt","","","","","","ESTRUCTURA:&nbsp;" & sInput,"si")

    else

      celdas = ofv.GenCelda("","","","","","","","&nbsp;","si")

    end if

      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    


    FilaS1 = ofv.GenTabla("","","","100%","","0","","-2","2","","",filas,"si")



    FL1 = "<DIV STYLE='border-style:ridge;border-width:3;' >"



      celdas = ofv.GenCelda("","ut colspan=4","CENTER","","","","","ETAPA/TAREA","si")
      celdas = celdas & ofv.GenCelda("","uT","CENTER","","","","","DUR.","si")
      celdas = celdas & ofv.GenCelda("","uT","CENTER","","","","","F.INICIO","si")
      celdas = celdas & ofv.GenCelda("","uT","CENTER","","","","","F.FIN","si")
      celdas = celdas & ofv.GenCelda("","uT COLSPAN=6","CENTER","","","","","PROPIEDADES","si")

      filas = ofv.GenRow("","aaa bgcolor='#e9e9e9' ","","","10","","",celdas,"si")

 xTrColor = "#FFFFFF" 

    StrSql = "Select MAX(ORDEN) From PROYETAPA where codigo = '" & session("NCliente") & "' AND ORDEN < 999  AND VPLAN = " & SESSION("VPLAN")
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

  XTANTST  = 1
  XTANTTSK = "NO"
  XTANTETP = "NO"

  StrSql = "Select * From PROYetapa where codigo = '" & session("NCliente") & "' AND VPLAN = " & SESSION("VPLAN") & " order by orden "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
     do while not rs.eof

 
     sinput1 = "<B>" & sImagen1N & "</B>"


      '    bAgregar   = FALSE
      '    bModificar = FALSE
      '    bSuprimir  = FALSE
    

  if bAgregar then
     sinput = "<a style='text-decoration:none;color:#ffffff;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";'  href=PROYtask.asp?opcion=AT&etp=" & rs(2) 
     sinput = sinput  & "  TITLE='AGREGAR TAREA'  ><b>" & sImagen2 & "</b></a>" 
     sinput2 = sinput
  ELSE
     sinput2 = sImagen1N
  END IF 

'''  if bModificar AND RS(3) > 20 AND RS(3) < 999 then
  if bModificar AND RS(3) > 99999 then
     sinput = "<a style='text-decoration:none;color:#ffffff;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";'  href=PROYtask.asp?opcion=SE&ETP=" & rs(2) & "&ETPANT=" & XTANTETP
     sinput = sinput  & "  TITLE='SUBIR ETAPA'  ><b>" & sImagen1S & "</b></a>" 
     sinpuTST = sinput
  ELSE
     sinpuTST = sImagen1N
  END IF 
   

    sform = "ETP" & RS(0)


    SINPUT = RS(3)
    IF SINPUT = 999 THEN SINPUT = XORDEN
    XETPORD = SINPUT
    XTANTETP = rs(2)
    celdas = ofv.GenCelda("","aaa COLSPAN=2 bgcolor=#f7aa22   ","","2%","","","",SINPUT,"si")
      if bModificar then
         sInput = ofv.GenerarInput(rs(4).name,rs(4),"text",sform,"80","aa STYLE='border-style:solid;border-width:1;border-color:#cccccc;font-family:verdana;font-size:10px;' ","")
      else
         sInput = ofv.GenerarInput("",rs(4),"readonly","","80","aa STYLE='border-style:solid;border-width:1;border-color:#cccccc;font-family:verdana;font-size:10px;' ","")
      end if
    celdas = celdas & ofv.GenCelda("","ut  STYLE='background-color:#bbddff;border-style:solid;border-width:1;border-color:#f7aa22;' colspan=9","","90%","","","",xestr & ": " & sInput,"si")


      celdas = celdas & ofv.GenCelda("","ata bgcolor=#336699 ","CENTER","","1%","","",sinputST,"si")
      celdas = celdas & ofv.GenCelda("","ata bgcolor=#008000 ","CENTER","","1%","","",sinput2,"si")

    if bModificar then
      sAccion = "asp/Actualizar.asp?FNAME=" & rs(0) & "&volver=" & vuelta & "&Tabla=PROYETAPA"
      formulario0 = ofv.GenForm(sform,"",sAccion,"","post","",celdas,"si")
    else
      formulario0 = ofv.GenForm("","","","","post","",celdas,"si")
    end if

    filas = filas & ofv.GenRow("","","","","10","","",formulario0,"si")



  StrSql = "Select * From PROYtareas where codigo = '" & session("NCliente") & "' and etapa = '" & rs(2) & "' AND VPLAN = " & SESSION("VPLAN") & " order by orden "
''  RESPONSE.WRITE STRSQL
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not  rs1.eof then
     do while not rs1.eof

    sform = "TSK" & RS1(0)



     IF XFSTS > 1 THEN
      if  rs1(18) = 0 then
          sinput1 = sImagenE0

      elseif rs1(18) = 2 then

          sinput1 = sImagenE2

      else

          sinput1 = sImagenE1

      END IF
     ELSE
        sinput1 = "<B>" & sImagen1N & "</B>"
     END IF  

     sinput = "<a style='text-decoration:none;color:#ffffff;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";'  href=PROYtaskprop.asp?opcion=TP&TSK=" & rs1(3) 
     sinput = sinput  & " TITLE='PROPIEDADES DE LA TAREA' ><b>" & sImagen3 & "</b></a>" 
     sinput3 = sinput


  if bModificar AND (((rs1(4) > 10 AND XETPORD > 10) OR (rs1(4) > 20 AND XETPORD = 10)) AND rs1(18) = 0 AND XTANTST = 0) then
     sinput = "<a style='text-decoration:none;color:#ffffff;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";'  href=PROYtask.asp?opcion=ST&TSK=" & rs1(3) & "&TSKANT=" & XTANTTSK
     sinput = sinput  & "  TITLE='SUBIR TAREA'  ><b>" & sImagen1S & "</b></a>" 
     sinpuTST = sinput
  ELSE
     sinpuTST = sImagen1N
  END IF 


    'Eliminar
    if bSuprimir AND (((XETPORD > 10  AND XTANTST = 0) OR (rs1(4) > 20 AND XETPORD = 10 AND XTANTST = 0 )  OR (rs1(4) > 10 AND XETPORD = 10 AND RS1(6) = 0 AND RS1(7) = 0 )  ) AND rs1(18) = 0 ) then
      sHref = "asp/EliminarRegistro.asp?ID=" & rs1(0)
      sHref = sHref & "&Tabla=PROYTAREAS&volver=" & vuelta & "&Eliminar=NSNC"
      CeldaS = ofv.BotonEliminar(" TAREA: " & XETPORD & "." & rs1(4) & " " & rs1(5),sHref)
    else
      CeldaS = ofv.GenCelda("","","","","","","",sImagen0,"si")
    end if

  XTANTST  = rs1(18)
  XTANTTSK = rs1(3)


    celdas = celdas &  ofv.GenCelda("","ut colspan=2    ","","7%","","","","&nbsp;" & XETPORD & "." & rs1(4),"si")
      if bModificar then
         sInput = ofv.GenerarInput(rs1(5).name,rs1(5),"text",sform,"70","aa STYLE='border-style:solid;border-width:1;border-color:#cccccc;font-family:verdana;font-size:10px;' ","")
      else
         sInput = ofv.GenerarInput("",rs1(5),"readonly","","70","aa STYLE='border-style:solid;border-width:1;border-color:#cccccc;font-family:verdana;font-size:10px;' ","")
      end if

    IF RS1(6) = 0 THEN
     sinputt1 = SIMAGEN6S

     IF RS1(7) = 0 THEN

        sinputt2 = SIMAGEN6SS
     ELSEIF RS1(7) = 1 THEN
        sinputt2 = SIMAGEN6N
     ELSE
        sinputt2 = SIMAGEN6OT

     END IF

    ELSE
     sinputt1 = SIMAGEN6T
     sinputt2 = SIMAGEN6SS
    END IF




    celdas = celdas & ofv.GenCelda("","ut","","75%","","","",sInput,"si")

       SINPUT4 = SIMAGEN4S
       XCOLOR = "#eeeeee"

    IF RS1(13) <> "NO" THEN


    ELSE

       STRSQL = "UPDATE PROYTAREAS SET USRRESP = '" & XUSRRESP & "' where codigo = '" & session("NCliente")
       StrSql = StrSql & "' AND TAREA = '" & RS1(3) & "' AND VPLAN = " & SESSION("VPLAN")
       CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)



    END IF 


    XIIX = OWKF.GETTSKCLC(RS1,XFHOY,XFCURR ,XFFCH)


   

    XFCURR = OWKF.ATABTSK(cint(XIIX),1)
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




' '''' DURACION DE LA TAREA ''''''''''''''''

    celdas = celdas & ofv.GenCelda("","ut ","right","5%","","","",sInput,"si")


' '''' FECHA DE INICIO DE LA TAREA ''''''''''''''''
    dFecha = cdate(TSKFS + clng(cdate("01-01-1900")))
    SINPUT = dFecha
    celdas = celdas & ofv.GenCelda("","ut ","","5%","","","",sInput,"si")

' '''' FECHA DE FINALIZACION DE LA TAREA ''''''''''''''''
    dFecha = cdate(TSKFE + clng(cdate("01-01-1900")))
    SINPUT = dFecha
    celdas = celdas & ofv.GenCelda("","ut ","","5%","","","",sInput,"si")












      celdas = celdas & ofv.GenCelda("","uts  ","CENTER","","1%","","",sinput4,"si")
      celdas = celdas & ofv.GenCelda("","uts  ","CENTER","","1%","","",sinputt1,"si")
      celdas = celdas & ofv.GenCelda("","uts  ","CENTER","","1%","","",sinputt2 ,"si")
      celdas = celdas & ofv.GenCelda("","uts  ","CENTER","","1%","","",sinput1,"si")
      celdas = celdas & ofv.GenCelda("","ut bgcolor=#6699CC ","CENTER","","1%","","",sinputST,"si")
      celdas = celdas & ofv.GenCelda("","ut bgcolor=#f7aa22 ","CENTER","","1%","","",sinput3,"si")


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
 
  end if



    rs.movenext
    loop
 
  end if

      celdas = ofv.GenCelda("","","CENTER","","1%","","","&nbsp;","si")
      celdas = celdas & ofv.GenCelda("","","CENTER","","2%","","","&nbsp;","si")
      celdas = celdas & ofv.GenCelda("","","CENTER","","5%","","","&nbsp;","si")
      celdas = celdas & ofv.GenCelda("","","CENTER","","73%","","","&nbsp;","si")
      celdas = celdas & ofv.GenCelda("","","CENTER","","5%","","","&nbsp;","si")
      celdas = celdas & ofv.GenCelda("","","CENTER","","5%","","","&nbsp;","si")
      celdas = celdas & ofv.GenCelda("","","CENTER","","5%","","","&nbsp;","si")
      celdas = celdas & ofv.GenCelda("","","CENTER","","1%","","","&nbsp;","si")
      celdas = celdas & ofv.GenCelda("","","CENTER","","1%","","","&nbsp;","si")
      celdas = celdas & ofv.GenCelda("","","CENTER","","1%","","","&nbsp;","si")
      celdas = celdas & ofv.GenCelda("","","CENTER","","1%","","","&nbsp;","si")
      celdas = celdas & ofv.GenCelda("","","CENTER","","1%","","","&nbsp;","si")
      celdas = celdas & ofv.GenCelda("","","CENTER","","1%","","","&nbsp;","si")

    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")


    FilaS2 = ofv.GenTabla("","aa style='border-style:solid;border-width:1;border-color:#888888;' ","","99%","","0","","-1","2","","",filas,"si")

    Celda = ofv.GenCelda("","","center","","","","",FilaS1 & FilaS2,"si")
    FilaS = ofv.GenRow("","","","","","","",Celda,"si")



  if XFSTS = 1 then

     cnbot = "<input type=submit class=bt value='MODIFICAR&nbsp;INICIO' >"
  else
     cnbot = "<input type=submit class=bt value='VER&nbsp;INICIO' >"
  end if
  Cons = ofv.gencelda("","","center","","1%","","",cnbot,"si")
  sAccion = "PROYtaskpry.asp"
  Cons = ofv.genform("cons","",sAccion,"","post","",Cons,"si")

     celdas = Cons
 

     celdas = celdas & ofv.gencelda("","","center","","1%","","","&nbsp;","si")

     sinput = sImagen4s & "&nbsp;Responsable&nbsp;asignado" 
     celdas = celdas & ofv.GenCelda("","ut","","20%","","","",sInput,"si")
     sinput = sImagen4n & "&nbsp;Responsable&nbsp;no&nbsp;asignado" 
     celdas = celdas & ofv.GenCelda("","ut","","20%","","","",sInput,"si")
     sinput = SIMAGEN6S & "&nbsp;Inicio&nbsp;independiente" 
     celdas = celdas & ofv.GenCelda("","ut","","20%","","","",sInput,"si")
     sinput = SIMAGEN6t & "&nbsp;Inicio&nbsp;en&nbsp;secuencia" 
     celdas = celdas & ofv.GenCelda("","ut","","20%","","","",sInput,"si")



     fila = ofv.GenRow("","","","","","","",celdas,"si")

     celdas = ofv.GenCelda("","aaa colspan=2","","","","","","&nbsp;","si")
     sinput = SIMAGEN6SS & "&nbsp;Dependencia&nbsp;simple" 
     celdas = celdas & ofv.GenCelda("","ut","","20%","","","",sInput,"si")
     sinput = SIMAGEN6N & "&nbsp;Dependencia&nbsp;por&nbsp;lista" 
     celdas = celdas & ofv.GenCelda("","ut","","20%","","","",sInput,"si")
     sinput = SIMAGEN6OT & "&nbsp;Dependencia&nbsp;otros&nbsp;Proyectos" 
     celdas = celdas & ofv.GenCelda("","ut","","20%","","","",sInput,"si")
     celdas = celdas & ofv.GenCelda("","aaa","","","","","","&nbsp;","si")





     fila = fila & ofv.GenRow("","","","","","","",celdas,"si")

     Fila = ofv.GenTabla("","aaa bgcolor=#ffffff ","","","","0","","4","0","","",fila,"si")



    END IF





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