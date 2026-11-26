<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FormHeader("")

  sHTML = sHTML  & "<script language=javascript> "
''  sHTML = sHTML  & "        document.getElementById('MDPTOP1').style.display='none';  document.getElementById('tbhMAX').style.display='none';        hdemenu();     "
  sHTML = sHTML  & "function AbrirCons(Datos) { "
  sHTML = sHTML  & " myWin= open(Datos," & chr(34) & "Cons" & chr(34) & "," & chr(34) & "status=yes,resizable=yes,menubar=no,toolbar=no,height=570,width=790" & chr(34) & " ); "
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

 IF SOPCION = "GEN" THEN
    Xorden = request.querystring("orden")
    Xetapa = request.querystring("etapa")
    xorden = xorden + 10
    xXFSTART = CLng(Date) - CLng(CDate("01-01-1900"))

    Xtsk = OWKF.GENPRYTSKTMP(session("NCliente"),SESSION("VPLAN"),XORDEN,Xetapa,"NUEVA INCIDENCIA",xXFSTART,xXFSTART,"1","4","0","1",SESSION("PRYCTA"))
    SESSION("TAREA") = Xtsk
 END IF 

 IF SOPCION = "FIN" THEN

    StrSql = "Select * From PROYecto where codigo = '" & session("NCliente") & "'  "
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT rs2.eof then
       XFFCH  = RS2(18)
    end if
  call ofv.cerrarconsulta(rs2)

  StrSql = "Select * From proytareastmp where codigo = '" & session("NCliente") & "' AND TAREA = '" & SESSION("TAREA") & "' AND VPLAN = " & SESSION("VPLAN")
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  IF NOT RS.EOF THEN


    XFHOY  = CLNG(DATE()) - clng(cdate("01-01-1900")) 
    XFCURR = CLNG("0" & rs(9))
    XIIX = OWKF.GETTSKCLC(RS,XFHOY,XFCURR ,XFFCH)

    FSREAL = OWKF.ATABTSK(cint(XIIX),9)
    FEREAL = OWKF.ATABTSK(cint(XIIX),10)
    XTDUR  = OWKF.ATABTSK(cint(XIIX),17)
    XUDUR  = OWKF.ATABTSK(cint(XIIX),18)

    XUSR    = session("usuario")
    XFSTART = CLNG("0" & rs(9))
    xtarea  = SESSION("TAREA")
    PADV    = 100
    TADV    = XTDUR
    TTIPO   = XUDUR

    StrSql = "UPDATE PROYTAREASTMP SET ESTADO = 2, FECHASREAL = " & FSREAL & ", FECHAEREAL = " & FEREAL
    StrSql = StrSql & " WHERE  codigo = '" & session("NCliente") & "' AND TAREA = '" & SESSION("TAREA") & "' "
    Call ofv.crearconsultaEx(StrSql,cn,1,parametros)

  END IF


  StrSql = "Select * From proytareastmp where codigo = '" & session("NCliente") & "' AND TAREA = '" & SESSION("TAREA") & "' AND VPLAN = " & SESSION("VPLAN")
  set RSXP = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  IF NOT RSXP.EOF THEN




     StrSql = "Insert into PROYtareas ("
     strsql2 = ") values ("

    For ii = 1 To (RSXP.Fields.Count - 1)
      If ii > 1 Then
         StrSql = StrSql & ","
         strsql2 = strsql2 & ","
      End If

      StrSql = StrSql & RSXP(ii).Name

      s = "'"
      sInput = RSXP(ii) & ""

      If UCase(sInput) = "VERDADERO" Then sInput = "1"
      If UCase(sInput) = "FALSO" Then sInput = "0"
      If UCase(sInput) = "TRUE" Then sInput = "1"
      If UCase(sInput) = "FALSE" Then sInput = "0"
 
      If RSXP(ii).Type = 131 Then sInput = Replace(RSXP(ii), ",", ".")
   
      strsql2 = strsql2 & s & sInput & s

    Next
  
    STRSQLX = StrSql & strsql2 & ") "
    Call ofv.crearconsultaEx(StrSqlX,cn,1,parametros)

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
    STRSQL = STRSQL & "" & TADV & "," & PADV & ",'" & XUSR & "','" & TTIPO & "','" & XFDHOY & "','" & XFDCUENTA & "')" 
    CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)




  END IF
  call ofv.cerrarconsulta(RSXP)
  call ofv.cerrarconsulta(RS)


  StrSql = "DELETE FROM PROYTAREASTMP  WHERE TAREA = '" & SESSION("TAREA") & "' "
  Call ofv.crearconsultaEx(StrSql,cn,1,parametros)

  RESPONSE.REDIRECT "pryadmtskdet.asp"


 END IF




  XTAREA = REQUEST.QUERYSTRING("TSK")

  IF XTAREA <> "" THEN
     SESSION("TAREA") = XTAREA
  ELSE
     XTAREA = SESSION("TAREA")
  END IF

 IF SOPCION = "SI" THEN
    XUSR = request.querystring("USR")
    XSEC = request.querystring("SECT")
    StrSql = "UPDATE proytareastmp SET SOLICUSR = '" & XUSR & "',SOLICSEC = '" & XSEC & "' WHERE codigo = '" & session("NCliente") & "' AND TAREA = '" & SESSION("TAREA") & "' AND VPLAN = " & SESSION("VPLAN")
    CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

 END IF 

 IF SOPCION = "AT" THEN
    XUSR = request.querystring("USR")
    StrSql = "UPDATE proytareastmp SET USRRESP = '" & XUSR & "' WHERE codigo = '" & session("NCliente") & "' AND TAREA = '" & SESSION("TAREA") & "' AND VPLAN = " & SESSION("VPLAN")
    CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

 END IF 


    sImagen0 = "<FONT FACE=WINGDINGS SIZE=2>F</FONT>"
    sImagen1 = "<FONT FACE=WINGDINGS SIZE=2>ñ</FONT>"
    sImagen1N = "<FONT FACE=WINGDINGS COLOR=#FFFFFF SIZE=2>l</FONT>"
    sImagen2 = "<FONT FACE=VERDANA SIZE=2>+</FONT>"
    sImagen3 = "<FONT FACE=WINGDINGS SIZE=2>2</FONT>"
    sImagen4s = "<B><FONT FACE=WEBDINGS COLOR=#6699cc SIZE=2>&#0128;</FONT></B>"
    sImagen4O = "<B><FONT FACE=WEBDINGS COLOR=#AA0000 SIZE=2>&#0128;</FONT></B>"
    sImagen4n = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=2>x</FONT></B>"
    sImagen6S = "<B><FONT FACE=WINGDINGS COLOR=#008000 SIZE=2>&#0252;</FONT></B>"
    sImagen6SS = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=2>&#0246;</FONT></B>"
    sImagen6N = "<B><FONT FACE=WINGDINGS COLOR=#AA0000 SIZE=3>3</FONT></B>"
    sImagen6OT = "<B><FONT FACE=WEBDINGS COLOR=#008000 SIZE=3>&#0254;</FONT></B>"
    sImagen6T = "<B><FONT FACE=WEBDINGS COLOR=#336699 SIZE=3>q</FONT></B>"
    sImagenut = "<B><FONT FACE=WINGDINGS COLOR=#6699cc SIZE=3>&#0192;</FONT></B>"
    sImagendu = "<B><FONT FACE=WINGDINGS COLOR=#993333 SIZE=3>h</FONT></B>"
    sImagenfi = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=3>R</FONT></B>"
    sImagenI3 = "<B><FONT FACE=WEBDINGS COLOR=#cc0000 SIZE=4>&#064;</FONT></B>"
    sImagenE0 = "<B><FONT FACE=WINGDINGS COLOR=#6699CC SIZE=4>6</FONT></B>"
    sImagenE1 = "<B><FONT FACE=WINGDINGS COLOR=#f7aa22 SIZE=4>&#0232;</FONT></B>"
    sImagenE2 = "<B><FONT FACE=WINGDINGS COLOR=#000080 SIZE=4>&#0251;</FONT></B>"

    sImagenF = "<B><FONT FACE=WEBDINGS COLOR=#F7AA22 SIZE=3>4</FONT></B>"

  sopcion = "VP"

  StrSql = "Select * From proytareastmp where codigo = '" & session("NCliente") & "' AND TAREA = '" & SESSION("TAREA") & "' AND VPLAN = " & SESSION("VPLAN")
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"


    StrSql = "Select * From PROYecto where codigo = '" & session("NCliente") & "'  "
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT rs2.eof then
       XFSTS  = rs2(11) 
       XVPLAN = RS2(17)
       XFFCH  = RS2(18)
    end if
  call ofv.cerrarconsulta(rs2)

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



  XTSKSTS = RS(18)

  IF XTSKSTS = 2 THEN
     bModificar = 0
  END IF




      celdas = ofv.GenCelda("","TTT colspan=4","","","","","","PROPIEDADES DE LA TAREA","si")
      filas  = ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","UT  ","","","","","","TAREA:","si")
      celdas = celdas & ofv.GenCelda("","UT","RIGHT","","","","",rs(4) & ".","si")
      if bModificar then
         sInput = ofv.GenerarInput(rs(5).name,rs(5),"text","Usuarios","80","dt","")
      else
         sInput = rs(5)
      end if
      celdas = celdas & ofv.GenCelda("","UT  COLSPAN=2","","","","","",sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      celdas = ofv.GenCelda("","UT  ","","","","","","CUENTA:","si")
      celdas = celdas & ofv.GenCelda("","UT","RIGHT","","","","","&nbsp;","si")
    if bModificar then
       StrSql = "Select * From PROYCUENTAS where estado = 1 ORDER BY ORDEN"
      sInput = ofv.GenerarCombo("Usuarios",rs(32).name,rs(32),strsql,cn,0,1,"")
    else
     StrSql = "Select * From PROYCUENTAS where ID = 0" & RS(32)
     set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if NOT rsX2.eof then
        sInput  = rsX2(1)

     ELSE
        sInput  = "SIN DETERMINAR"

     END IF
    end if
      celdas = celdas & ofv.GenCelda("","UT  COLSPAN=2","","","","","",sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



      celdas = ofv.GenCelda("","UT  ","","","","","","CLASE:","si")
      celdas = celdas & ofv.GenCelda("","UT","RIGHT","","","","","&nbsp;","si")
    if bModificar then
       StrSql = "Select * From PROYTIPOTSK ORDER BY ID"
      sInput = ofv.GenerarCombo("Usuarios",rs(19).name,rs(19),strsql,cn,0,1,"")
    else
     StrSql = "Select * From PROYTIPOTSK where ID = " & RS(19)
     set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if NOT rsX2.eof then
        sInput  = rsX2(1)

     ELSE
        sInput  = "SIN DETERMINAR"

     END IF
    end if
      celdas = celdas & ofv.GenCelda("","UT  COLSPAN=2","","","","","",sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


     StrSql = "Select * From PROYSUBTIPOTSK where TIPO = 0" & RS(19)
     set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if NOT rsX2.eof then

      celdas = ofv.GenCelda("","UT  ","","","","","","SUBCLASE:","si")
      celdas = celdas & ofv.GenCelda("","UT","RIGHT","","","","","&nbsp;","si")
    if bModificar then
       StrSql = "Select * From PROYSUBTIPOTSK WHERE TIPO = 0" & RS(19) & " ORDER BY ID"
       sInput = ofv.GenerarCombo("Usuarios",rs(14).name,rs(14),strsql,cn,0,1,"")
    else
     XSUBTIPO = RS(14) & ""
     IF XSUBTIPO = "NO" THEN
        XSUBTIPO = "0"
     END IF
     StrSql = "Select * From PROYSUBTIPOTSK where ID = " & XSUBTIPO
     set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if NOT rsX2.eof then
        sInput  = rsX2(1)

     ELSE
        sInput  = "SIN DETERMINAR"

     END IF
    end if
    
      celdas = celdas & ofv.GenCelda("","UT  COLSPAN=2","","","","","",sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    END IF



      celdas = ofv.GenCelda("","TTT  COLSPAN=4   ","","","","","","SOLICITANTE","si")
      filas = filas & ofv.GenRow("","","","","30","","",celdas,"si")

  if isnull(rs(26)) then 


       sinput2 = "<a CLASS=CC style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "none" & chr(34) & ";' class=tt href=PROYTASKSOLBUS.asp?TSK=" & SESSION("TAREA")
       sinput2 = sinput2  & " TITLE='AFECTAR' ><b>" & sImagenI3 & "</b></a>" 

      celdas = ofv.GenCelda("","Ut  COLSPAN=4  ","","","","","",sImagen4O & "&nbsp;SIN&nbsp;INFORMAR" & "&nbsp;&nbsp;&nbsp;&nbsp;" & sInput2,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


  ELSE

       sinput2 = "<a  CLASS=CC style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "none" & chr(34) & ";' class=tt href=PROYTASKSOLBUS.asp?TSK=" & SESSION("TAREA")
       sinput2 = sinput2  & " TITLE='MODIFICAR'  ><b>" & sImagenI3 & "</b></a>" 
       SCOLOR =  "ata bgcolor=#008000 "


      sInput = "SIN DETERMINAR"

      strsql = "Select * From USUARIOS where ID = '" & rs(26) & "' "
      set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
      if not rsniv.eof then
         sniv   = rsniv(1)
         sInput = "" & sniv & ""
      END IF
      call ofv.cerrarconsulta(rsniv)

      celdas = ofv.GenCelda("","UT  COLSPAN=4  ","","","","","",sImagen4O & "&nbsp;" & sInput & "&nbsp;&nbsp;&nbsp;&nbsp;" & sInput2,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      StrSql = "Select * From ECO where ECO = 0" & RS(25)
      set rs2 = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
      if NOT rs2.eof then
         sInput1 = rs2(1)
      ELSE
        sInput1 = "SIN DETERMINAR"
      END IF

      celdas = ofv.GenCelda("","UT  COLSPAN=4  ","","","","","",sImagenE1 & "&nbsp;" & sInput1,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



  end if 


  SFORM = "Usuarios"



      celdas = ofv.GenCelda("","UT  ","","","","","","HORA INICIO:","si")
      celdas = celdas & ofv.GenCelda("","UT","RIGHT","","","",""," ","si")
      if bModificar then
         sInput = ofv.GenerarInput(rs(28).name,rs(28),"text","Usuarios","10","dt","")
      else
         sInput = rs(28)
      end if
      celdas = celdas & ofv.GenCelda("","UT  COLSPAN=2","","","","","",sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","UT  ","","","","","","HORA FINALIZACION:","si")
      celdas = celdas & ofv.GenCelda("","UT","RIGHT","","","",""," ","si")
      if bModificar then
         sInput = ofv.GenerarInput(rs(29).name,rs(29),"text","Usuarios","10","dt","")
      else
         sInput = rs(29)
      end if
      celdas = celdas & ofv.GenCelda("","UT  COLSPAN=2","","","","","",sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      celdas = ofv.GenCelda("","TTT  COLSPAN=4   ","","","","","","OBSERVACIONES","si")
      filas = filas & ofv.GenRow("","","","","30","","",celdas,"si")



      nCol = 80

  if len(wdescri) = 0 then wdescri = "Sin Observaciones" 

  if not  isnull(rs(27)) then 
     wdescri = rs(27)
  end if 




    t3b = "<font class=at><textarea class=dt  style='background-color:transparent;padding-top:2px;' "
    t3b = t3b & "name=" & rs(27).name & " cols=" & nCol & " rows=5 "
    if bModificar then
      t3b = t3b & " onblur=Usuarios.submit() "
      t3b = t3b & " onkeypress=" & chr(34)
      t3b = t3b & " if (this.value.length > 1950) {alert('El texto es demasiado"
      t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
    else
      t3b = t3b & " readonly >"
    end if
    xdatos =  wdescri
    t3b = t3b & xdatos & "</textarea></font>"
    Celda = t3b



      celdas = ofv.GenCelda("","UT  COLSPAN=4   ","CENTER","","","","",Celda,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      celdas = ofv.GenCelda("","TTT  COLSPAN=4   ","","","","","","ESTADO, ASIGNACION Y DURACION PLANIFICADA","si")
      filas = filas & ofv.GenRow("","","","","30","","",celdas,"si")


      celdas = ofv.GenCelda("","UT  ","","","","","","ESTADO:","si")

      if  rs(18) = 0 then
          sinput1 = sImagenE0
          sinput2 = "NO INICIADA"

      elseif rs(18) = 2 then

          sinput1 = sImagenE2
          sinput2 = "FINALIZADA"

      else

          sinput1 = sImagenE1
          sinput2 = "EN CURSO"

      END IF


      celdas = celdas & ofv.GenCelda("","UT","","","","","",sinput1,"si")

      celdas = celdas & ofv.GenCelda("","UT  COLSPAN=2","","","","","",sinput2,"si")
   ''   filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      dFecha = cdate(CLNG("0" & rs(9)) + clng(cdate("01-01-1900")))

      sinput = "<a class=nt  STYLE='CURSOR:HAND;' href='pryadmtskdetagrFEC.asp' "
      sinput = sinput  & " title='modificar fecha de inicio' "
      sinput = sinput  & " ><font color='#993333' ><u><B>" & dfecha & "</B>&nbsp;" & sImagenF & "</u></font></a>" 

      dfecha = sinput


      celdas = ofv.GenCelda("","UT  ","","","","","","FECHA INICIO:","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","",sImagendu,"si")
''      if bModificar then
''         sInput = ofv.GenerarInput(rs(15).name,rs(15),"text","Usuarios","10","dtN","")
''      else
''         sInput = rs(15)
''      end if
      celdas = celdas & ofv.GenCelda("","UT  COLSPAN=2","","","","","",dFecha,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    XFHOY  = CLNG(DATE()) - clng(cdate("01-01-1900")) 
    XFCURR = CLNG("0" & rs(9))
 ''''   XFFCH  = 1

    XIIX = OWKF.GETTSKCLC(RS,XFHOY,XFCURR ,XFFCH)


   

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


    dFecha = cdate(CLNG("0" & FEREAL) + clng(cdate("01-01-1900")))


      celdas = ofv.GenCelda("","UT  ","","","","","","FECHA FIN:","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","",sImagendu,"si")
''      if bModificar then
''         sInput = ofv.GenerarInput(rs(15).name,rs(15),"text","Usuarios","10","dtN","")
''      else
''         sInput = rs(15)
''      end if
      celdas = celdas & ofv.GenCelda("","UT  COLSPAN=2","","","","","",dFecha,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



      celdas = ofv.GenCelda("","UT  ","","","","","","TIEMPO:","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","",sImagendu,"si")
      if bModificar then
         sInput = ofv.GenerarInput(rs(15).name,rs(15),"text","Usuarios","10","dtN","")
      else
         sInput = rs(15)
      end if
      celdas = celdas & ofv.GenCelda("","UT  COLSPAN=2","","","","","",sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      celdas = ofv.GenCelda("","UT  ","","","","","","UNIDAD DE TIEMPO:","si")
      celdas = celdas & ofv.GenCelda("","UT","center","","","","",sImagenut,"si")
    if bModificar AND XTSKSTS = 0 then
        scomK = "4" & chr(9) & "1" & chr(9) & "2" & chr(9) & "3"
        scomD = "MINUTOS" & chr(9) & "HORAS" & chr(9) & "DIAS" & chr(9) & "SEMANAS"

        sInput = ofv.gencomboFX(sForm,rs(16).name,rs(16),scomK,scomD," ")
    else
        if rs(16) = 1 then sInput = "HORAS"
        if rs(16) = 2 then sInput = "DIAS"
        if rs(16) = 3 then sInput = "SEMANAS"
        if rs(16) = 4 then sInput = "MINUTOS"
      ''  sInput = ofv.GenerarInput("",sInput,"readonly","","40","dt","")
    end if
    celdas = celdas & ofv.GenCelda("","UT  COLSPAN=2","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


     '  sinput2 = "<a class=cc style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
     '  sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
     '  sinput2 = sinput2  & "none" & chr(34) & ";' href=PROYtaskPropFin.asp"
     '  sinput2 = sinput2  & " ><b>" & sImagenI3  & "&nbsp;Finalizar Tarea</b></a>" 








    if bModificar then
      sAccion = "asp/Actualizar.asp?FNAME=" & rs(0) & "&volver=" & vuelta & "&Tabla=proytareastmp"
      formulario0 = ofv.GenForm("Usuarios","",sAccion,"","post","",filas,"si")
    else
      formulario0 = ofv.GenForm("","","","","post","",filas,"si")
    end if


    filas = ofv.GenRow("","","","","","","",formulario0,"si")

    Celda = ofv.GenCelda("","abc colspan=4","center","100%","","","","&nbsp;","si")
    FilaS =  filaS & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","AAA STYLE='border-style:SOLID;border-width:1;border-color:#888888;'  ","","70%","","0","","0","0","","",filas,"si")


''''''' confirmacion

    sInput = ofv.GenerarInput("btn1","CONFIRMAR","SUBMIT","confirma","10","BT","")

    Celda = ofv.GenCelda("","","center","100%","","","",sinput,"si")

    filas = ofv.GenRow("","","","","","","",Celda,"si")

    if bModificar then
      sAccion = "pryadmtskdetagr.asp?opcion=FIN"
      formulario2 = ofv.GenForm("confirma","",sAccion,"","post","",filas,"si")
    else
      formulario2 = ofv.GenForm("","","","","post","",filas,"si")
    end if


    filas = formulario2

    sHTML = sHTML & ofv.GenTabla("","AAA STYLE='border-style:SOLID;border-width:1;border-color:#993333;'  ","","70%","","0","","0","0","","",filas,"si")


 
    sHTML = sHTML & "</center>"




  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  Agregar = ""
  Buscar  = ""

  sAccion = "pryadmtskdet.asp"
  Celda = ofv.BotonVolver(sAccion)


  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,Celda)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML






%>