<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFNro.inc" -->
<%

ok=ofv.CheckUsuario()

if ok <> "Y" then
  sHTML = ok
else
  response.expires=0


  sHTML = ofv.FORMHeader("")



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
 sopcionACT = request.querystring("opcionACT")

  XTAREA = REQUEST.QUERYSTRING("TSK")

  IF XTAREA <> "" THEN
     SESSION("TAREA") = XTAREA

     XADMINC = REQUEST.QUERYSTRING("ADMINC")

     IF XADMINC = "" THEN
        XADMINC = "N"
     END IF
     SESSION("ADMINC") = XADMINC

  ELSE
     XTAREA = SESSION("TAREA")
  END IF


xSTS = 0


 IF SOPCION = "SI" THEN
    XUSR = request.querystring("USR")
    XSEC = request.querystring("SECT")
    StrSql = "UPDATE PROYTAREAS SET SOLICUSR = '" & XUSR & "',SOLICSEC = '" & XSEC & "' WHERE codigo = '" & session("NCliente") & "' AND TAREA = '" & SESSION("TAREA") & "' AND VPLAN = " & SESSION("VPLAN")
    CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

 END IF 

 IF SOPCION = "AT" THEN
    XUSR = request.querystring("USR")
    StrSql = "UPDATE PROYTAREAS SET USRRESP = '" & XUSR & "' WHERE codigo = '" & session("NCliente") & "' AND TAREA = '" & SESSION("TAREA") & "' AND VPLAN = " & SESSION("VPLAN")
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

  sopcion2 = sopcion

  IF SOPCION2 = "M2" THEN   vuelta = vuelta & "_OPCION=M2"


  sopcion = "VP"


    XFSTART = CLNG(DATE()) - clng(cdate("01-01-1900")) 
    XFCURR  = XFSTART
    XFHOY   = XFSTART
    XFREAL  = XFSTART

    StrSql = "Select * From PROYecto where codigo = '" & session("NCliente") & "'  "
    set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT rs2.eof then
       XFSTS  = rs2(11) 
       XVPLAN = RS2(17)

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

       XFFCH  = RS2(18)
       IF XFFCH > 0 THEN XXX = OWKF.ADMFERIADOS("C",XFSTART)

    end if
  call ofv.cerrarconsulta(rs2)


  StrSql = "Select * From PROYTAREAS where codigo = '" & session("NCliente") & "' AND TAREA = '" & SESSION("TAREA") & "' AND VPLAN = " & SESSION("VPLAN")
 '''' RESPONSE.WRITE STRSQL & "<BR>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

  IF NOT RS.EOF THEN

     TSKMODCREATE = "M" & RS(33)

     IF  TSKMODCREATE = "M2" THEN

         TSKMODASIGNA = "M" & RS(24)

         IF TSKMODASIGNA <> "M1" THEN
            
            XUSR = request.querystring("USR")
            StrSql = "UPDATE PROYTAREAS SET MASIG = 1, CASIG = TDUR, TASIG = UDUR  "
            StrSql = StrSql & " WHERE codigo = '" & session("NCliente") & "' AND TAREA = '" & SESSION("TAREA") & "' AND VPLAN = " & SESSION("VPLAN")
            CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

            StrSql = "Select * From PROYTAREAS where codigo = '" & session("NCliente") & "' AND TAREA = '" & SESSION("TAREA") & "' AND VPLAN = " & SESSION("VPLAN")
            set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)



         END IF


         IF SOPCIONACT = "M2" THEN

            WKTADV  = RS(15)
            WKTTIPO = RS(16)
            WKUSRUPD = SESSION("USUARIO")
            WKFCARGA = CLNG(DATE()) - CLNG(CDATE("01-01-1900"))

            StrSql = "UPDATE PROYTAREASADV SET TADV = '" & WKTADV & "', TTIPO = '" & WKTTIPO & "', USRUPD = '" & WKUSRUPD & "', FCARGA = '" & WKFCARGA & "' "
            STRSQL = STRSQL & " WHERE codigo = '" & session("NCliente") & "' AND TSK = '" & SESSION("TAREA") & "' "
            CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

         END IF 


     END IF 


    XIIX = OWKF.GETTSKCLC(RS,XFHOY,XFREAL ,XFFCH)




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



  END IF



    sHTML = sHTML & "<center>"

    sImagen = "<FONT FACE=WINGDINGS SIZE=4>4</FONT>"

    TSKMODCREATE = "M" & RS(33)






    IF XFSTS > 1 THEN

       IF XFSTS = 15 OR XFSTS = 16 THEN
          IF CLNG(SESSION("VPLAN")) = CLNG(XVPLAN) THEN
             bAgregar   = FALSE
             bModificar = FALSE
             bSuprimir  = FALSE
          END IF
       ELSE

       END IF
    END IF

  bModiesp  = FALSE

  if bModificar then   bModiesp  = bModificar


  XTSKSTS = RS(18)

  IF XTSKSTS = 2 THEN
     bModificar = 0
  END IF

  IF XTSKSTS = 0 THEN
     if rs(6) = 0 AND  rs(6) = 0 THEN
        xSTS = 1
     END IF
  END IF




     SINPUTF1 = ""
     SINPUTF2 = ""

     SFORM = "Usuarios1"


      celdas = ofv.GenCelda("","ttr colspan=4","","","","","","PROPIEDADES DE LA TAREA","si")
      filas  = ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","UT  ","","","","","","TAREA:","si")
      celdas = celdas & ofv.GenCelda("","UT","RIGHT","","","","",rs(4) & ".","si")
      if bModificar then
         sInput = ofv.GenerarInput(rs(5).name,rs(5),"text",SFORM,"60","dt","")
      else
         sInput = rs(5)
      end if
      celdas = celdas & ofv.GenCelda("","UT  COLSPAN=2","","","","","",sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","UT  ","","","","","","CLASE:","si")
      celdas = celdas & ofv.GenCelda("","UT","RIGHT","","","","","&nbsp;","si")
    if bModificar OR TSKMODCREATE = "M2" then
       StrSql = "Select * From PROYTIPOTSK ORDER BY ID"
      sInput = ofv.GenerarCombo(SFORM,rs(19).name,rs(19),strsql,cn,0,1,"")
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
    if bModificar  OR TSKMODCREATE = "M2" then
       StrSql = "Select * From PROYSUBTIPOTSK WHERE TIPO = 0" & RS(19) & " ORDER BY ID"
       sInput = ofv.GenerarCombo(SFORM,rs(14).name,rs(14),strsql,cn,0,1,"")
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



      celdas = ofv.GenCelda("","ttr  COLSPAN=4   ","","","","","","SOLICITANTE","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

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


    if bModificar  OR TSKMODCREATE = "M2"  then
      sAccion = "asp/Actualizar.asp?FNAME=" & rs(0) & "&volver=" & vuelta & "&Tabla=PROYTAREAS"
      formulario0 = ofv.GenForm(sform,"",sAccion,"","post","",filas,"si")
    else
      formulario0 = ofv.GenForm("","","","","post","",filas,"si")
    end if

    SINPUTF1 = SINPUTF1 & ofv.GenTabla("","aa STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","98%","","0","","0","0","","",formulario0,"si")
    SINPUTF1 = SINPUTF1 & "<br>"

  SFORM = "Usuarios12"

      filas = "" 

      celdas = ofv.GenCelda("","ttr  COLSPAN=4   ","","","","","","CUENTA DE PRESUPUESTO","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      if bModiesp then
         StrSql = "Select * From PROYCUENTAS WHERE ESTADO = 1 ORDER BY ORDEN"
         sInput = ofv.GenerarCombo(SFORM,"CTANUEVA",rs(32),strsql,cn,0,1,"")
      else
         StrSql = "Select DESCRIPCION From PROYCUENTAS where ID = " & RS(32)
         set rsX2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
         if NOT rsX2.eof then
            sInput  = rsX2(0)

         ELSE
            sInput  = "SIN ASIGNAR"

         END IF
      end if
      celdas = ofv.GenCelda("","UT  COLSPAN=4","CENTER","","","","",sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")




    if bModiesp then
      sAccion = "PRYCUENTASTRF.asp?TSK=" & rs(0)
      formulario0 = ofv.GenForm(sform,"",sAccion,"","post","",filas,"si")
    else
      formulario0 = ofv.GenForm("","","","","post","",filas,"si")
    end if

    SINPUTF2 = SINPUTF2 & ofv.GenTabla("","aa STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","98%","","0","","0","0","","",formulario0,"si")
    SINPUTF2 = SINPUTF2 & "<br>"


  IF XTSKSTS > 0  OR (rs(6) = 0) THEN
 



  SFORM = "Usuarios2"

      filas = "" 


      celdas = ofv.GenCelda("","ttr  COLSPAN=4   ","","","","","","OBSERVACIONES","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      wkfecha = rs(9)
      if wkfecha = 0 then  wkfecha = rs(8)
      if wkfecha = 0 then  
        '''''''' wkfecha = CLNG(DATE()) - clng(cdate("01-01-1900"))
         wkfecha = cdate(TSKFS)
      END IF


          
          celdas = ofv.GenCelda("","UT  ","","","","","","FECHA INICIO:","si")
          celdas = celdas & ofv.GenCelda("","UT","RIGHT","","","",""," ","si")
          sInput = cdate(wkfecha + clng(cdate("01-01-1900")))
          celdas = celdas & ofv.GenCelda("","UT  COLSPAN=2","","","","","",sInput,"si")
          filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")








      celdas = ofv.GenCelda("","UT  ","","","","","","HORA INICIO:","si")
      celdas = celdas & ofv.GenCelda("","UT","RIGHT","","","",""," ","si")
      if bModificar then
         sInput = ofv.GenerarInput(rs(28).name,rs(28),"text",SFORM,"10","dt","")
      else
         sInput = rs(28)
      end if
      celdas = celdas & ofv.GenCelda("","UT  COLSPAN=2","","","","","",sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      wkfecha = rs(11)
      if wkfecha = 0 then  wkfecha = rs(10)

      if wkfecha = 0 then  
 ''''        wkfecha = CLNG(DATE()) - clng(cdate("01-01-1900"))
         wkfecha = cdate(TSKFE)
      END IF


          celdas = ofv.GenCelda("","UT  ","","","","","","FECHA FINALIZACION:","si")
          celdas = celdas & ofv.GenCelda("","UT","RIGHT","","","",""," ","si")
          sInput = cdate(wkfecha + clng(cdate("01-01-1900")))
          celdas = celdas & ofv.GenCelda("","UT  COLSPAN=2","","","","","",sInput,"si")
          filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      celdas = ofv.GenCelda("","UT  ","","","","","","HORA FINALIZACION:","si")
      celdas = celdas & ofv.GenCelda("","UT","RIGHT","","","",""," ","si")
      if bModificar then
         sInput = ofv.GenerarInput(rs(29).name,rs(29),"text",SFORM,"10","dt","")
      else
         sInput = rs(29)
      end if
      celdas = celdas & ofv.GenCelda("","UT  COLSPAN=2","","","","","",sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      nCol = 50

      if len(wdescri) = 0 then wdescri = "Sin Observaciones" 

      if not  isnull(rs(27)) then 
         wdescri = rs(27)
      end if 




      t3b = "<font class=at><textarea class=dt  style='background-color:transparent;padding-top:2px;' "
      t3b = t3b & "name=" & rs(27).name & " cols=" & nCol & " rows=6 "
      if bModificar then
        t3b = t3b & " onblur=" & SFORM & ".submit() "
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


    if bModificar then
      sAccion = "asp/Actualizar.asp?FNAME=" & rs(0) & "&volver=" & vuelta & "&Tabla=PROYTAREAS"
      formulario0 = ofv.GenForm(sform,"",sAccion,"","post","",filas,"si")
    else
      formulario0 = ofv.GenForm("","","","","post","",filas,"si")
    end if

    SINPUTF2 = SINPUTF2 & ofv.GenTabla("","aa STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","98%","","0","","0","0","","",formulario0,"si")
    SINPUTF2 = SINPUTF2 & "<br>"

  END IF


    IF xSTS = 1 THEN

    filas = "" 

  cniv = request.querystring("Cniv")
  cA   = request.querystring("CA")

  if cniv = "" then 
     cniv = session("ENC7")
     cA   = session("ENCA")
  else
     session("ENC7") = cniv
     session("ENCA") = cA
  end if



  XFSR = FSPLAN


  posi = 0

  SFORM = "Usuarios2xtn"

   ''''' sForm = "form" & posi
       sclase = "CC" 

     Celda = ofv.GenCelda("","VT","","15%","","","","FECHA INICIO:","si")



    if clng(XFSR) = 0 then
       sInput1   = "------------"
    else
       sInput1   = cdate(clng(XFSR) + clng(cdate("01-01-1900")) )
       swd       = datepart("w",sInput1)
       sInput1   = Fechador(XFSR)
    END IF


     Celda = Celda & ofv.gencelda("","VT colspan=3 ","","80%","","","",sInput1,"si")



    Fila = ofv.genrow("","","","","","","",Celda,"si")

  ''  filas = filas & ofv.GenTabla("","","","50%","","0","","0","0","","",Fila,"si")

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
 '      response.write "smesc " & smesc & " ca " & ca
       if isnull(smesc) then  SMESC = ca
       if len(smesc) = 0 then  SMESC = ca
       ca =  SMESC
      
    else
       ca = smesc

    end if  

    if smesc = 0 then smesc = smes

       response.write "smesc " & smesc & " ca " & ca

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
     sInput = ofv.GenerarCombo(sForm & "A","smesc",smesc,Strsql,cn2,1,2,"sb0")
     Celda = Celda & ofv.gencelda("","VT","","45%","","","",sInput,"si")

     Celda = Celda & ofv.GenCelda("","VT","","15%","","","","AÑO:","si")
     
     sinput = sanioc 
     Celda = Celda & ofv.gencelda("","VT","","20%","","","",sInput,"si")

    Fila = FILA & ofv.genrow("","","","","","","",Celda,"si")


    sAccion = "PROYTASKPROP.asp?" & REQUEST.QUERYSTRING & "&cniv=" & cniv & "&cnom=" &  cnom & "&cx=" & cx & "&ca=" & smesc
    Formulario = ofv.genform(sForm & "A","",sAccion,"","post","",Fila,"si")

    filas = filas & ofv.GenTabla("","AAA style='border-style:ridge;border-width:2;border-color:#f7aa22;'","","50%","","0","","0","0","","",Formulario,"si")


    dim afech(33,1)
    dim afechs(42,1)

    CELDA = ""

    sfecds = "01-" & cstr(smesc) & "-" & cstr(sanioc)

    'sfecd = clng(cdate(sfecds))
    sfecd = clng(dateserial(sanioc,smesc,1))

    'swd = datepart("w",date(sfecds))
     swd = datepart("w",dateserial(sanioc,smesc,1))
     
    sfechs = "28-" & cstr(smesc) & "-" & cstr(sanioc)

    'sfech = clng(date(sfechs))
     sfech = clng(dateserial(sanioc,smesc,28))
 '   response.write " fd " & sfecd & " fh " & sfech

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

 '  response.write " wd " & swd & " fd " & sfecd & " fh " & sfech
    srest = clng(cdate("01-01-1900"))
   ' response.write  sfech &" -"&sfecd 
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

 '   RESPONSE.WRITE "II " & II & " TOPE " & (ii + topefecH ) & " TOPEFECH " & topefecH
    for ii = ii to (ii + topefecH )
        afechs(ii,0) = yy
        afechs(ii,1) = afech(yy,1)
        yy = yy + 1
    next

'    RESPONSE.WRITE "II " & II

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

    for ii = 1 to 43

        if yy = 7 then
           yy = 1
           Fila = fila & ofv.genrow("","","","","","","",Celda,"si")
           CELDA = ""
        else
           yy = yy + 1 
        end if

        if ii > 42 then exit for  

        if yy = 1 or yy = 7 then
           sclase = "tt3 style='border-style:outset;border-width:2;'" 
        else
           sclase = "TT style='border-style:outset;border-width:2;'" 
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

           sAccion = "asp/Actualizar.asp?FNAME=" & rs(0) & "&volver=" & vuelta & "&Tabla=PROYTAREAS"

           t1 = ofv.GenerarInput("bvo1",sInput1,"submit","","10",sclase,"")
           t1 = t1 & ofv.GenerarInput(rs(8).name,(afechs(ii,1) - srest),"hidden","","2","dt","")
        '''''''   t1 = t1 & ofv.GenerarInput(rs(3).name,(afechs(ii,1) - srest),"hidden","","2","dt","")
           t1 = ofv.GenCelda("","VT","center","10%","","","",t1,"si")
           t1 = ofv.GenForm("feac" & ii,"",sAccion,"","post","",t1,"si") 
           Celda = Celda & t1

        end if


    next

'''


   filas = filas & ofv.GenTabla("","AAA style='border-style:ridge;border-width:2;border-color:#f7aa22;'","","50%","","0","","0","0","","",FILA,"si")

    SINPUTF2 = SINPUTF2& filas
    SINPUTF2 = SINPUTF2 & "<br>"

 
  END IF


'''''''''''''''''''''








  SFORM = "Usuarios3"

      filas = "" 



      celdas = ofv.GenCelda("","ttr  COLSPAN=4   ","","","","","","ESTADO, ASIGNACION Y DURACION PLANIFICADA","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


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
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


'''''''''''''''' SOLO PARA TAREAS GENERADAS DESDE UN CIRCUITO DE WOKFLOW (MODCREATE = 2) ---- AJUSTE DE IMPUTACION -----------------

 IF TSKMODCREATE = "M2" THEN


      celdas = ofv.GenCelda("","ttr  COLSPAN=4   ","","","","","","AJUSTE DE IMPUTACION PARA LA ACTIVIDAD","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

 END IF


      celdas = ofv.GenCelda("","UT  ","","","","","","TIEMPO:","si")
      celdas = celdas & ofv.GenCelda("","","center","","","","",sImagendu,"si")
      if bModificar OR TSKMODCREATE = "M2" then
         wfopc = "text"
      if TSKMODCREATE = "M2" then wfopc = "agregar"
         sInput = ofv.GenerarInput(rs(15).name,rs(15),wfopc,SFORM,"10","dtN","")
      else
         sInput = rs(15)
      end if
      celdas = celdas & ofv.GenCelda("","UT  COLSPAN=2","","","","","",sInput,"si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      celdas = ofv.GenCelda("","UT  ","","","","","","UNIDAD DE TIEMPO:","si")
      celdas = celdas & ofv.GenCelda("","UT","center","","","","",sImagenut,"si")
    if (bModificar AND XTSKSTS = 0) OR TSKMODCREATE = "M2" then

         wfopc = ""
      if TSKMODCREATE = "M2" then wfopc = "no"

        scomK = "4" & chr(9) & "1" & chr(9) & "2" & chr(9) & "3"
        scomD = "MINUTOS" & chr(9) & "HORAS" & chr(9) & "DIAS" & chr(9) & "SEMANAS"

        sInput = ofv.gencomboFX(sForm,rs(16).name,rs(16),scomK,scomD,wfopc)

    else
        if rs(16) = 1 then sInput = "HORAS"
        if rs(16) = 2 then sInput = "DIAS"
        if rs(16) = 3 then sInput = "SEMANAS"
        if rs(16) = 4 then sInput = "MINUTOS"
    end if
    celdas = celdas & ofv.GenCelda("","UT  COLSPAN=2","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


 IF TSKMODCREATE = "M2" THEN

      sInput = ofv.GenerarInput("botonm2","ACTUALIZAR","SUBMIT",SFORM,"10","BT","")
      celdas = ofv.GenCelda("","ut  COLSPAN=4 STYLE='border-BOTTOM-style:solid;border-BOTTOM-width:1;border-BOTTOM-color:#888888;'  ","RIGHT","","","","",sinput,"si")
      filas = filas & ofv.GenRow("","","","","20","","",celdas,"si")

 END IF


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   

      celdas = ofv.GenCelda("","UT  ","","","","","","MODO DE ASIGNACION:","si")
      celdas = celdas & ofv.GenCelda("","UT","center","","","","",sImagenut,"si")
    if bModificar AND XTSKSTS = 0 AND rs(16) < 4 then

        scomK = "0" & chr(9) & "1"
        scomD = "FULL-TIME" & chr(9) & "PART-TIME"


        sInput = ofv.gencomboFX(sForm,rs(24).name,rs(24),scomK,scomD," ")



    else
        if rs(24) = 0 then sInput = "FULL-TIME"
        if rs(24) = 1 then sInput = "PART-TIME"

    end if
    celdas = celdas & ofv.GenCelda("","UT  COLSPAN=2","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



   IF RS(24) > 0 THEN 

      celdas = ofv.GenCelda("","TTT  COLSPAN=4  ","","","","","","PERIODO DE ASIGNACION","si")
      filas = filas & ofv.GenRow("","","","","30","","",celdas,"si")

      IF RS(22) > 0 THEN
         XTDUR = RS(22)
      ELSE
         XTDUR = RS(15) * 2
      END IF
   
      IF RS(23) > 0 THEN
         XUDUR = RS(23)
      ELSE
         XUDUR = RS(16)
      END IF



      SELECT CASE RS(16)

             CASE 4

                  XVDUR = RS(15) / 60

             CASE 1

                  XVDUR = RS(15)

             CASE 2
                  XVDUR = RS(15) * 7


             CASE 3 
                  XVDUR = RS(15) * 49

      END SELECT


      SELECT CASE XUDUR

             CASE 4

                  XWDUR = XTDUR / 60

             CASE 1

                  XWDUR = XTDUR

             CASE 2
                  XWDUR = XTDUR * 7


             CASE 3 
                  XWDUR = XTDUR * 49

      END SELECT

      if  XWDUR = 0 then XWDUR = 1
      if  XVDUR = 0 then XVDUR = 1


      XWVDUR = CINT(XVDUR / XWDUR * 100)

''XWVDUR = XVDUR & "-" & XWDUR


      celdas = ofv.GenCelda("","UT  ","","","","","","TIEMPO:","si")
      celdas = celdas & ofv.GenCelda("","UT","center","","","","",sImagendu,"si")
      if bModificar then
         sInput = ofv.GenerarInput(rs(22).name,XTDUR,"text",SFORM,"10","dtN","")
      else
         sInput = XTDUR
      end if
      celdas = celdas & ofv.GenCelda("","UT  COLSPAN=2","","","","","",sInput & "&nbsp;(AL&nbsp;" & XWVDUR & "%)","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


      celdas = ofv.GenCelda("","UT  ","","","","","","UNIDAD DE TIEMPO:","si")
      celdas = celdas & ofv.GenCelda("","UT","center","","","","",sImagenut,"si")
    if bModificar AND XTSKSTS = 0 then


        scomK = "1" & chr(9) & "2" & chr(9) & "3"
        scomD = "HORAS" & chr(9) & "DIAS" & chr(9) & "SEMANAS"
        sInput = ofv.gencomboFX(sForm,rs(23).name,rs(23),scomK,scomD," ")


    else
        if XUDUR = 1 then sInput = "HORAS"
        if XUDUR = 2 then sInput = "DIAS"
        if XUDUR = 3 then sInput = "SEMANAS"
        if XUDUR = 4 then sInput = "MINUTOS"

    end if
    celdas = celdas & ofv.GenCelda("","UT  COLSPAN=2","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")


  END IF


      celdas = ofv.GenCelda("","UT  ","","","","","","MODO DE INICIO:","si")

    IF RS(6) = 0 THEN

     sinputt1 = SIMAGEN6S
    ELSE
     sinputt1 = SIMAGEN6T

    END IF
      celdas = celdas & ofv.GenCelda("","UT","center","","","","",sinputt1,"si")
    if bModificar AND XTSKSTS = 0 then

        scomK = "0" & chr(9) & "1"
        scomD = "INDEPENDIENTE" & chr(9) & "EN SECUENCIA"

        sInput = ofv.gencomboFX(sForm,rs(6).name,rs(6),scomK,scomD," ")



    else
        if rs(6) = 0 then sInput = "INDEPENDIENTE"
        if rs(6) = 1 then sInput = "EN SECUENCIA"

    end if
    celdas = celdas & ofv.GenCelda("","UT  COLSPAN=2","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



      celdas = ofv.GenCelda("","UT  ","","","","","","DEPENDENCIA:","si")

     IF RS(7) = 0 THEN

     sinputt2 = SIMAGEN6SS
    ELSEIF RS(7) = 1 THEN
     sinputt2 = SIMAGEN6N
    ELSE
     sinputt2 = SIMAGEN6OT

    END IF

      celdas = celdas & ofv.GenCelda("","UT","center","","","","",sinputt2,"si")
    if bModificar AND XTSKSTS = 0  then

        scomK = "0" & chr(9) & "1" & chr(9) & "2"
        scomD = "INICIO MANUAL" & chr(9) & "LISTA DE TAREAS" & chr(9) & "OTROS PROYECTOS"

        sInput = ofv.gencomboFX(sForm,rs(7).name,rs(7),scomK,scomD," ")

    else
        if rs(7) = 0 then sInput = "INICIO MANUAL"
        if rs(7) = 1 then sInput = "LISTA DE TAREAS"
        if rs(7) = 2 then sInput = "OTROS PROYECTOS"

    end if
    if rs(7) = 0 then
       celdas = celdas & ofv.GenCelda("","UT   ","","","","","",sInput,"si")

       sinput2 = "<a class=cc style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
       sinput2 = sinput2  & "none" & chr(34) & ";' href=PROYtaskPropFin.asp"
       sinput2 = sinput2  & " ><b>" & sImagenI3  & "&nbsp;Finalizar Tarea</b></a>" 

       SINPUT = "&nbsp;"
       if bModificar AND XTSKSTS < 2   then SINPUT = sinput2

      celdas = celdas & ofv.GenCelda("","UT","","50%","","","",sinput,"si")
    elseif rs(7) = 1 then
       celdas = celdas & ofv.GenCelda("","UT ","","","","","",sInput,"si")

      sinput2 = "<a class=cc style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
      sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
      sinput2 = sinput2  & "none" & chr(34) & ";' class=tt href=PROYtasklst.asp?OPCION=TP"
      sinput2 = sinput2  & " ><b>" & sImagenI3  & "&nbsp;Lista de Tareas</b></a>" 

      SINPUT = "&nbsp;"
      if bModificar AND XTSKSTS = 0    then SINPUT = sinput2
      celdas = celdas & ofv.GenCelda("","UT","","50%","","","",sinput,"si")

    else
       celdas = celdas & ofv.GenCelda("","UT ","","","","","",sInput,"si")

      sinput2 = "<a class=cc style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
      sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
      sinput2 = sinput2  & "none" & chr(34) & ";' class=tt href=PROYtasklst.asp?OPCION=OP"
      sinput2 = sinput2  & " ><b>" & sImagenI3  & "&nbsp;Lista de Tareas</b></a>" 

      SINPUT = "&nbsp;"
      if bModificar AND XTSKSTS = 0    then SINPUT = sinput2
      celdas = celdas & ofv.GenCelda("","UT","","50%","","","",sinput,"si")

    end if



      if rs(6) = 0 then  filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","UT  ","","","","","","FECHA DE INICIO:","si")
      celdas = celdas & ofv.GenCelda("","UT","center","","","","",sImagenfi,"si")
    if bModificar  then

        scomK = "0" & chr(9) & "1"
        scomD = "NORMAL" & chr(9) & "PREDETERMINADA"

        sInput = ofv.gencomboFX(sForm,rs(17).name,rs(17),scomK,scomD," ")

    else
        if rs(17) = 0 then sInput = "NORMAL"
        if rs(17) = 1 then sInput = "PREDETERMINADA"

    end if
    celdas = celdas & ofv.GenCelda("","UT colspan=2","","","","","",sInput,"si")
 '''   filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")




      celdas = ofv.GenCelda("","UT  ","","","","","","RESPONSABLE:","si")



      sinput2 = "<a class=cc style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
      sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
      sinput2 = sinput2  & "none" & chr(34) & ";' class=tt href=PROYRESP.asp?OPCION=AT"
      sinput2 = sinput2  & " ><b>" & sImagen4s & "</b></a>" 

      SINPUT = "<b>" & sImagen4s & "</b>"
      if bModificar AND XTSKSTS = 0  then SINPUT = sinput2
      celdas = celdas & ofv.GenCelda("","UT","center","","","","",sinput,"si")

    if RS(13) = "NO" then
       sInput = "<B><I>SIN AFECTAR</I></B>"
    ELSE
      strsql = "Select * From USUARIOS where ID = '" & rs(13) & "' "
      set rsniv = ofv.crearconsultaEx(StrSql,cn2,1,parametros)
      if not rsniv.eof then sniv = rsniv(1)
      call ofv.cerrarconsulta(rsniv)
      sInput = "" & sniv & ""

      sinput2 = "<a class=cc style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
      sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
      sinput2 = sinput2  & "none" & chr(34) & ";' TARGET=_blank href=PROYASGN.asp?RESP=" & rs(13)
      sinput2 = sinput2  & " >" & sInput & "</a>" 

      sInput = sinput2
    END IF
    celdas = celdas & ofv.GenCelda("","UT  colspan=2","","","","","",sInput,"si")
    filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    vueltaACT = vuelta

    IF TSKMODCREATE = "M2" THEN

       vueltaACT = vueltaACT & "_OPCIONACT=M2"

    END IF



    if bModificar OR TSKMODCREATE = "M2" then
      sAccion = "asp/Actualizar.asp?FNAME=" & rs(0) & "&volver=" & vueltaACT & "&Tabla=PROYTAREAS"
      formulario0 = ofv.GenForm(SFORM,"",sAccion,"","post","",filas,"si")
    else
      formulario0 = ofv.GenForm("","","","","post","",filas,"si")
    end if

    SINPUTF1 = SINPUTF1 & ofv.GenTabla("","aa STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","98%","","0","","0","0","","",formulario0,"si")



'''''''''''''''''''''''''' lista de tareas dependencia '''''''''''''''''''''''''''''''''''''''''''''''''''
  IF rs(7) > 0 THEN


      filas = ""

      celdas = ofv.GenCelda("","ttr COLSPAN=3  ","","","","","","LISTA&nbsp;DE&nbsp;TAREAS&nbsp;DE&nbsp;DEPENDENCIA","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

      celdas = ofv.GenCelda("","UT ","CENTER","30%","","","","TAREA","si")
      celdas = celdas & ofv.GenCelda("","UT","CENTER","30%","","","","PROYECTO","si")
      celdas = celdas & ofv.GenCelda("","UT","CENTER","30%","","","","ESTADO","si")
      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")



  StrSql = "Select * From PROYtareasLST where codigo = '" & session("NCliente") & "' and TAREA = '" & SESSION("TAREA") & "' order by orden "
  set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs1.eof then

      SINPUT = "NO HAY TAREAS DE DEPENDENCIA INFORMADAS"
      celdas = ofv.GenCelda("","vt colspan=3 ","CENTER","","","","",SINPUT,"si")

      filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

  ELSE
     do while not rs1.eof
        StrSql = "Select * From PROYECTO WHERE CODIGO = '" & RS1(4) & "' "
        set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        IF NOT RS2.EOF THEN
           XVPLAN  = RS2(17)
           XPRYNOM = RS2(2)
        END IF


        StrSql = "Select * From PROYTAREAS WHERE TAREA = '" & RS1(5) & "' AND CODIGO = '" & RS1(4) & "' AND VPLAN = " & session("vplan")
        set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        IF NOT RS2.EOF THEN

           XTSKNOM = RS2(4) & "." & RS2(5)
        END IF
        SINPUT = XTSKNOM
        IF RS1(1) <> RS1(4) THEN
           SINPUT = SIMAGEN6OT & "&nbsp;" & SINPUT
        ELSE
           SINPUT = SIMAGEN6N & "&nbsp;" & SINPUT
        END IF
    
        celdas = ofv.GenCelda("","UT ","CENTER","","","","",SINPUT,"si")

        SINPUT = XPRYNOM
        celdas = celdas & ofv.GenCelda("","UT","CENTER","","","","",SINPUT,"si")
        if rs1(6) = 0 then sInput = "ACTIVA"
        if rs1(6) = 1 then sInput = "FINALIZADA"
        if rs1(6) = 2 then sInput = "CANCELADA"
        celdas = celdas & ofv.GenCelda("","UT","CENTER","","","","",SINPUT,"si")
        filas = filas & ofv.GenRow("","","","","10","","",celdas,"si")

    rs1.movenext
    loop
 
  end if

    SINPUTF2 = SINPUTF2 & ofv.GenTabla("","aa STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","98%","","0","","0","0","","",filas,"si")

 END IF

'''''''''''''''''''' fin lista dependencias '''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    celdas = ""
    celdas = celdas & ofv.GenCelda("","UT valign=top ","center","50%","","","",SINPUTF1,"si")
    celdas = celdas & ofv.GenCelda("","UT valign=top ","center","50%","","","",SINPUTF2,"si")
    filas = ofv.GenRow("","","","","10","","",celdas,"si")

    sHTML = sHTML & ofv.GenTabla("","aa STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","98%","","0","","0","0","","",filas,"si")



'''''''''''''''''''''

















 
    sHTML = sHTML & "</center>"

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  Agregar = ""
  Buscar  = ""

  if   SESSION("ADMINC") = "S" then
       sAccion = "pryadmtskdet.asp"
  else
       if Sopcion2 = "M2" then
          sAccion = "PROYEST.asp"
       else  
          sAccion = "PROYTASK.asp"
       end if 
  end if
  Celda = ofv.BotonVolver(sAccion)


  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,Celda)


  sHTML = sHTML  & "</body></html>"
end if

response.write sHTML


Function Fechador(wkfecha)


    
      fec = cdate(wkfecha + clng(cdate("01-01-1900")))
      ndia = ofv.cdia(Weekday(fec))
      nMes = ofv.cmes(Month(fec))
    
      Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)


End Function



%>