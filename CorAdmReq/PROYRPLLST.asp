<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.formHeader(" ")

  sHTML = sHTML  & "<script language=javascript> "
  sHTML = sHTML  & "function AbrirCons(Datos) { "
  sHTML = sHTML  & " myWin= open(Datos," & chr(34) & "Cons2" & chr(34) & "," & chr(34) & "menubar=no,toolbar=no,height=450,width=600" & chr(34) & " ); "
  sHTML = sHTML  & " } "
  sHTML = sHTML  & "</script>"

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

 sopcion = request.querystring("opcion")

 if sopcion <> "" then
    session("OPC") = sopcion
 else
    sopcion = session("OPC")
 end if 

 sacc = request.querystring("accion")






''    sImagen0 = "<FONT FACE=WINGDINGS SIZE=2>F</FONT>"
    sImagen1 = "<FONT FACE=WINGDINGS SIZE=2>ñ</FONT>"
    sImagen1N = "<FONT FACE=WINGDINGS COLOR=#FFFFFF SIZE=2>l</FONT>"
    sImagen2 = "<FONT FACE=VERDANA SIZE=2>+</FONT>"
    sImagen0 = "<FONT FACE=WINGDINGS COLOR=#6699cc  SIZE=2>2</FONT>"
    sImagen3 = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=3>&#0166;</FONT></B>"
    sImagen4n = "<B><FONT FACE=WEBDINGS COLOR=#f7aa22 SIZE=3>N</FONT></B>"
    sImagen6S = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=2>&#0252;</FONT></B>"
    sImagen6SS = "<B><FONT FACE=WEBDINGS COLOR=#008000 SIZE=3>&#0232;</FONT></B>"
    sImagen6N = "<FONT COLOR=#993333 FACE=WEBDINGS SIZE=2>4</FONT>"
    sImagen6M = "<FONT COLOR=#993333 FACE=WEBDINGS SIZE=2>8</FONT>"
    sImagen6OT = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=3>r</FONT></B>"
    sImagen6T = "<B><FONT FACE=WEBDINGS COLOR=#336699 SIZE=2>r</FONT></B>"
    sImagenSTN = "<B><FONT FACE=WINGDINGS COLOR=#336699 SIZE=3>6</FONT></B>"
    sImagenSTC = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>&#0184;</FONT></B>"
    sImagenE0 = "<B><FONT FACE=WINGDINGS COLOR=#6699CC SIZE=4>6</FONT></B>"
    sImagenE1 = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=4>&#0232;</FONT></B>"
    sImagenE2 = "<B><FONT FACE=WINGDINGS COLOR=#000080 SIZE=4>&#0251;</FONT></B>"
    sImagenCC = "<B><FONT FACE=WEBDINGS COLOR=#000000 SIZE=4>2</FONT></B>"
 




    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"


    XFSTS = 0 

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
    end if

  call ofv.cerrarconsulta(rs2)

     StrSql = "Select * From PROYprty where codigo = '" & session("NCliente") & "'  "
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     IF NOT RS.EOF THEN

      XFCURR = XFSTART

      celdas = ofv.GenCelda("","ttt colspan=4","","","","","",sImagen & "&nbsp;PLANIFICACIONES DEL PROYECTO","si")
      filas  = ofv.GenRow("","","","","10","","",celdas,"si")

        Strsql = "Select descri From reqestados Where id = " & XFSTS
        set rsInstrumento = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rsInstrumento.eof then sInstrumento = rsInstrumento(0)
        call ofv.cerrarconsulta(rsInstrumento)


      sInput = sInstrumento

      if XFSTS > 1 then
         XFDSCSTS = "<BR>&nbsp;" & sImagenSTC & "&nbsp;<SMALLER>" & sInput & "</SMALLER>"
      ELSE
         XFDSCSTS = "<BR>&nbsp;" & sImagenSTN & "&nbsp;<SMALLER>" & sInput & "</SMALLER>"
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

  END IF

      filas = ""


      SINPUT = "PLAN VIGENTE"
      celdas = ofv.GenCelda("","vt COLSPAN=2  ","","","","","",SINPUT,"si")

      SINPUT = "CREACION"
      celdas = celdas & ofv.GenCelda("","VT COLSPAN=2   ","","","","","",SINPUT,"si")

      SINPUT = "RESPONSABLE"
      celdas = celdas & ofv.GenCelda("","VT   ","","","","","",SINPUT,"si")

      SINPUT = "ESTADO"
      celdas = celdas & ofv.GenCelda("","VT COLSPAN=4  ","","","","","",SINPUT,"si")
      filas = filas & ofv.GenRow("","aaa BGCOLOR='#e9e9e9' ","","","10","","",celdas,"si")

      SINPUT = "&nbsp;"
      celdas = ofv.GenCelda("","ut colspan=10 STYLE='border-TOP-style:solid;border-TOP-width:1;border-TOP-color:#f7aa22;' ","CENTER","","","","",SINPUT,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


  StrSql = "Select * From PROYPLAN where codigo = '" & session("NCliente") & "' and ESTADO = 1 "
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs1.eof then

      SINPUT = "--------------"
      celdas = ofv.GenCelda("","ut colspan=10 ","","","","","",SINPUT,"si")

      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  ELSE
      SINPUT = sImagen6N & " " & RS1(2) & "."
      celdas = ofv.GenCelda("","ut ","RIGHT","8%","","","",SINPUT,"si")

    xlink = "<A CLASS=CC href='&nbsp;' onmouseover='this.style.textDecoration=" & chr(34)
    xlink = xlink & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
    xlink = xlink & "none" & chr(34) & ";' "
    xlink = xlink & " onclick=" & chr(34) & "AbrirCons('PROYRPLVIEW.asp?PLAN=" & RS1(0) & "'); return false;" & chr(34)
    xlink = xlink & " >" & RS1(3) & "</A>"

    sInput = xlink

    '  SINPUT = RS1(3)
      celdas = celdas & ofv.GenCelda("","uT","","30%","","","",SINPUT,"si")

      SINPUT = cdate(RS1(4) + clng(cdate("01-01-1900")))
      celdas = celdas & ofv.GenCelda("","uT","","7%","","","",SINPUT,"si")

      SINPUT = RS1(5)
      celdas = celdas & ofv.GenCelda("","uT","","7%","","","",SINPUT,"si")


    if RS1(6) = "NO" then
      sInput = "SIN RESPONSABLE"
      celdas = celdas & ofv.GenCelda("","uT","","40%","","","",sInput,"si")
    else
      strsql = "Select * From USUARIOS where ID = '" & RS1(6) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = sImagenO & " <B>" & sniv & "</B>"
      celdas = celdas & ofv.GenCelda("","uT","","40%","","","",SINPUT,"si")
    END IF


     sinput = "<a style='text-decoration:none;color:#ffffff;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' TARGET=_blank href=PROYRPLDSP.asp?opcion=HT&VPLAN=" & RS1(2)
     sinput = sinput  & " TITLE='VER/MODIFICAR TAREAS DEL PLAN' ><b>" & sImagen3 & "</b></a>" 
     sinput31 = sinput

     sinput = "<a style='text-decoration:none;color:#ffffff;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' TARGET=_blank href=PROYRPLDSP.asp?opcion=HE&VPLAN=" & RS1(2)
     sinput = sinput  & " TITLE='VER ESTADO/GANTT DEL PLAN' ><b>" & sImagen4n & "</b></a>" 
     sinput32 = sinput


     IF XFSTS = 12 THEN

        sinput = "<a style='text-decoration:none;color:#ffffff;' onmouseover='this.style.textDecoration=" & chr(34)
        sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
        sinput = sinput  & "none" & chr(34) & ";'  href=PROYRPLGEN.asp?opcion=CP&VPLAN=" & RS1(2)
        sinput = sinput  & " TITLE='GENERAR COPIA DEL PLAN' ><b>" & sImagenCC & "</b></a>" 
     ELSE
        sinput = "&nbsp;"
     END IF


     sinput33 = sinput

     sinput34 = "&nbsp;"


      sinput = sinput32 
      celdas = celdas & ofv.GenCelda("","uT","CENTER","","","","",SINPUT,"si")
      sinput = sinput31 
      celdas = celdas & ofv.GenCelda("","uT","CENTER","","","","",SINPUT,"si")
      sinput = sinput33 
      celdas = celdas & ofv.GenCelda("","uT","CENTER","","","","",SINPUT,"si")
      sinput = sinput34 
      celdas = celdas & ofv.GenCelda("","uT","CENTER","","","","",SINPUT,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


  END IF 

  call ofv.cerrarconsulta(rs1)

      SINPUT = "&nbsp;"
      celdas = ofv.GenCelda("","vt colspan=10 STYLE='border-BOTTOM-style:solid;border-BOTTOM-width:1;border-BOTTOM-color:#f7aa22;' ","CENTER","","","","",SINPUT,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      SINPUT = "REPLANIFICACION PENDIENTE"
      celdas = ofv.GenCelda("","vt COLSPAN=2 ","","","","","",SINPUT,"si")

      SINPUT = "CREACION"
      celdas = celdas & ofv.GenCelda("","VT COLSPAN=2  ","","","","","",SINPUT,"si")

      SINPUT = "RESPONSABLE"
      celdas = celdas & ofv.GenCelda("","VT  ","","","","","",SINPUT,"si")

      SINPUT = "ESTADO"
      celdas = celdas & ofv.GenCelda("","VT COLSPAN=4  ","","","","","",SINPUT,"si")
      filas = filas & ofv.GenRow("","aaa BGCOLOR='#e9e9e9' ","","","10","","",celdas,"si")

      SINPUT = "&nbsp;"
      celdas = ofv.GenCelda("","vt colspan=10 STYLE='border-TOP-style:solid;border-TOP-width:1;border-TOP-color:#f7aa22;' ","CENTER","","","","",SINPUT,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  StrSql = "Select * From PROYPLAN where codigo = '" & session("NCliente") & "' and ESTADO = 0 "
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs1.eof then

      SINPUT = "--------------"
      celdas = ofv.GenCelda("","ut colspan=10 ","","","","","",SINPUT,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  ELSE
      SINPUT = sImagen6N & " " & RS1(2) & "."
      celdas = ofv.GenCelda("","vt ","RIGHT","","","","",SINPUT,"si")

    xlink = "<A CLASS=CC href='&nbsp;' onmouseover='this.style.textDecoration=" & chr(34)
    xlink = xlink & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
    xlink = xlink & "none" & chr(34) & ";' "
    xlink = xlink & " onclick=" & chr(34) & "AbrirCons('PROYRPLVIEW.asp?PLAN=" & RS1(0) & "'); return false;" & chr(34)
    xlink = xlink & " >" & RS1(3) & "</A>"

    sInput = xlink

    '  SINPUT = RS1(3)
      celdas = celdas & ofv.GenCelda("","uT","","","","","",SINPUT,"si")

      SINPUT = cdate(RS1(4) + clng(cdate("01-01-1900")))
      celdas = celdas & ofv.GenCelda("","uT","","","","","",SINPUT,"si")

      SINPUT = RS1(5)
      celdas = celdas & ofv.GenCelda("","uT","","","","","",SINPUT,"si")

    if RS1(6) = "NO" then
      sInput = "SIN RESPONSABLE"
      celdas = celdas & ofv.GenCelda("","uT","","","","","",sInput,"si")
    else
      strsql = "Select * From USUARIOS where ID = '" & RS1(6) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = sImagenO & " <B>" & sniv & "</B>"
      celdas = celdas & ofv.GenCelda("","uT","","","","","",SINPUT,"si")
    END IF






     sinput = "<a style='text-decoration:none;color:#ffffff;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' TARGET=_blank href=PROYRPLDSP.asp?opcion=HT&VPLAN=" & RS1(2)
     sinput = sinput  & " TITLE='VER/MODIFICAR TAREAS DEL PLAN' ><b>" & sImagen3 & "</b></a>" 
     sinput31 = sinput

     sinput = "<a style='text-decoration:none;color:#ffffff;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' TARGET=_blank href=PROYRPLDSP.asp?opcion=HE&VPLAN=" & RS1(2)
     sinput = sinput  & " TITLE='VER ESTADO/GANTT DEL PLAN' ><b>" & sImagen4n & "</b></a>" 
     sinput32 = sinput

     sinput = "<a style='text-decoration:none;color:#ffffff;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";'  href=PROYRPLDSP.asp?opcion=AP&VPLAN=" & RS1(2)
     sinput = sinput  & " TITLE='ACTIVAR PLAN' ><b>" & sImagen6SS & "</b></a>" 
     sinput33 = sinput

     sinput = "<a style='text-decoration:none;color:#ffffff;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";'  href=PROYRPLDSP.asp?opcion=CP&VPLAN=" & RS1(2)
     sinput = sinput  & " TITLE='CANCELAR PLAN' ><b>" & sImagen6OT & "</b></a>" 
     sinput34 = sinput


      sinput = sinput32 
      celdas = celdas & ofv.GenCelda("","uT","CENTER","","","","",SINPUT,"si")
      sinput = sinput31 
      celdas = celdas & ofv.GenCelda("","uT","CENTER","","","","",SINPUT,"si")
      sinput = sinput33 
      celdas = celdas & ofv.GenCelda("","uT","CENTER","","","","",SINPUT,"si")
      sinput = sinput34 
      celdas = celdas & ofv.GenCelda("","uT","CENTER","","","","",SINPUT,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


  END IF 

  call ofv.cerrarconsulta(rs1)

      SINPUT = "&nbsp;"
      celdas = ofv.GenCelda("","vt colspan=10 STYLE='border-BOTTOM-style:solid;border-BOTTOM-width:1;border-BOTTOM-color:#f7aa22;' ","CENTER","","","","",SINPUT,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      SINPUT = "HISTORIAL DE PLANIFICACIONES"
      celdas = ofv.GenCelda("","vt COLSPAN=2 ","","","","","",SINPUT,"si")

      SINPUT = "CREACION"
      celdas = celdas & ofv.GenCelda("","VT COLSPAN=2 ","","","","","",SINPUT,"si")

      SINPUT = "RESPONSABLE"
      celdas = celdas & ofv.GenCelda("","VT  ","","","","","",SINPUT,"si")

      SINPUT = "ESTADO"
      celdas = celdas & ofv.GenCelda("","VT COLSPAN=4 ","","","","","",SINPUT,"si")
      filas = filas & ofv.GenRow("","aaa BGCOLOR='#e9e9e9' ","","","10","","",celdas,"si")

      SINPUT = "&nbsp;"
      celdas = ofv.GenCelda("","ut colspan=10 STYLE='border-TOP-style:solid;border-TOP-width:1;border-TOP-color:#f7aa22;' ","CENTER","","","","",SINPUT,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  StrSql = "Select * From PROYPLAN where codigo = '" & session("NCliente") & "' and ESTADO > 1 ORDER BY ESTADO, VPLAN DESC "
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs1.eof then

      SINPUT = "--------------"
      celdas = ofv.GenCelda("","ut colspan=10 ","","","","","",SINPUT,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  ELSE
     do while not rs1.eof

      IF RS1(7) = 2 THEN
         SINPUT = sImagen6M & " " & RS1(2) & "."
      ELSE
         SINPUT = sImagen6T & " " & RS1(2) & "."
      END IF
      celdas = ofv.GenCelda("","ut ","RIGHT","","","","",SINPUT,"si")

     xlink = "<A CLASS=CC href='&nbsp;' onmouseover='this.style.textDecoration=" & chr(34)
    xlink = xlink & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
    xlink = xlink & "none" & chr(34) & ";' "
    xlink = xlink & " onclick=" & chr(34) & "AbrirCons('PROYRPLVIEW.asp?PLAN=" & RS1(0) & "'); return false;" & chr(34)
    xlink = xlink & " >" & RS1(3) & "</A>"

    sInput = xlink

    '  SINPUT = RS1(3)
      celdas = celdas & ofv.GenCelda("","uT","","","","","",SINPUT,"si")

      SINPUT = cdate(RS1(4) + clng(cdate("01-01-1900")))
      celdas = celdas & ofv.GenCelda("","uT","","","","","",SINPUT,"si")

      SINPUT = RS1(5)
      celdas = celdas & ofv.GenCelda("","uT","","","","","",SINPUT,"si")

    if RS1(6) = "NO" then
      sInput = "SIN RESPONSABLE"
      celdas = celdas & ofv.GenCelda("","uT","","","","","",sInput,"si")
    else
      strsql = "Select * From USUARIOS where ID = '" & RS1(6) & "' "
        set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
        if not rsniv.eof then sniv = rsniv(1)
        call ofv.cerrarconsulta(rsniv)
      sInput = sImagenO & " <B>" & sniv & "</B>"
      celdas = celdas & ofv.GenCelda("","uT","","","","","",SINPUT,"si")
    END IF


 
     sinput31 = "&nbsp;"

     sinput = "<a style='text-decoration:none;color:#ffffff;' onmouseover='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     sinput = sinput  & "none" & chr(34) & ";' TARGET=_blank  href=PROYRPLDSP.asp?opcion=HE&VPLAN=" & RS1(2)
     sinput = sinput  & " TITLE='VER ESTADO/GANTT DEL PLAN' ><b>" & sImagen4n & "</b></a>" 
     sinput32 = sinput

     sinput33 = "&nbsp;"

     sinput34 = "&nbsp;"



      sinput = sinput32 
      celdas = celdas & ofv.GenCelda("","uT","CENTER","","","","",SINPUT,"si")
      sinput = sinput31 
      celdas = celdas & ofv.GenCelda("","uT","CENTER","","","","",SINPUT,"si")
      sinput = sinput33 
      celdas = celdas & ofv.GenCelda("","uT","CENTER","","","","",SINPUT,"si")
      sinput = sinput34 
      celdas = celdas & ofv.GenCelda("","uT","CENTER","","","","",SINPUT,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")






    rs1.movenext
    loop
 
  end if

  call ofv.cerrarconsulta(rs1)



     IF XFSTS = 12 THEN

        sinput = "<a CLASS=TT style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
        sinput = sinput  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
        sinput = sinput  & "none" & chr(34) & ";' class=tt href=PROYRPLGEN.ASP "
        sinput = sinput  & " ><b>&nbsp;GENERAR&nbsp;NUEVA&nbsp;PLANIFICACION</b></a>" 
        celdas = ofv.GenCelda("","TT ","RIGHT","10%","","","",sInput,"si")
     ELSE
        celdas = ofv.GenCelda("","","","","","","","&nbsp;","si")
     END IF


     celdas = CELDAS & ofv.GenCelda("","VT ","RIGHT","70%","","","","&nbsp;","si")


     fila = ofv.GenRow("","","","","","","",celdas,"si")

    FilaS2 = ofv.GenTabla("","","","100%","","0","","2","1","","",fila,"si")


    sHTML = sHTML & FILAS1 
    sHTML = sHTML & ofv.GenTabla("","AAA STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","98%","","0","","-1","2","","",filas,"si")
    sHTML = sHTML & FILAS2 & "</center><br><br>"





 




  call ofv.cerrarconsulta(rs1)
  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML





%>