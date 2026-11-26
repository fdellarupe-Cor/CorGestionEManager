<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/CorWKFMANAGER.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  CantRegMostrarx = 8
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  sMODAL        = request.querystring("MODAL")
  sCODTRACER    = request.querystring("CODTRACER")
  sOPCION       = request.querystring("OPCION")
  SPRETRACER    = request.querystring("PRETRACER")
  SMODAL2       =request.querystring("MODAL2")
  sHTML = ofv.FormHeader(session("FormN"))

  IF SMODAL="VTR" and SMODAL2="" THEN SMODAL2="VTR"

  vuelta        = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_CODTRACER=" & sCODTRACER & "_MODAL=" & sMODAL

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)


'
'''''''''''''''''''''''''''''''''''''''


   Celda = ""

   wkproccod = "&nbsp;"

   Celda = Celda & ofv.GenCelda("","uti  ","","90%","","","","BASE DE EVENTOS DE RIESGO - ENCADENAMIENTO DE EVENTOS ","si")
   Celda = Celda & ofv.GenCelda("","ut   ","","2%","","","",wkproccod,"si")
   Fila  = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML & ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ","","100% ","30","0","","-1","1","","",Fila,"si")



 if SOPCION = "SI" THEN

        sFECMARCA  = CLNG(DATE()) - clng(cdate("01-01-1900"))

        strsql = "UPDATE ORTRACER SET TRCCODTRACERP = '" & SPRETRACER & "' Where CODTRACER = '" & SCODTRACER & "' "
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

        wwwkmarca = "K"

        Strsql = "select TRCMARCAACTP FROM ORTRACER WHERE CODTRACER = '" & SPRETRACER & "' "
        set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if not rs2.eof then
            
           if (RS2(0) & "") = "R"   then wwwkmarca = "KR"
           if (RS2(0) & "") = "XR"  then wwwkmarca = "KR"
           if (RS2(0) & "") = "X"   then wwwkmarca = "KR"
           if (RS2(0) & "") = "KR"  then wwwkmarca = "KR"

        end if



        Strsql = "UPDATE ORTRACER SET TRCMARCAACTP = '" & wwwkmarca & "', TRCFECACTP = '" & sFECMARCA & "' WHERE CODTRACER = '" & SPRETRACER & "' "
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

        XXXX = ORRECALCEVENTO(sCODTRACER,"NO")



 END IF

 if SOPCION = "DESAF" THEN

        strsql = "UPDATE ORTRACER SET TRCCODTRACERP = '' Where CODTRACER = '" & SCODTRACER & "' "
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

        XXXX = ORRECALCEVENTO(SPRETRACER,"SI")



 END IF

    WKEVENRELAC = 0

    sForm = "formACT"
    Formulario = ""
    sInput = "&nbsp;"
    Celda = ofv.gencelda("","VTI colspan=5 ","","","","","","EVENTO DE PERDIDA ACTUAL","si")
    Formulario = ofv.genrow("","AAA bgcolor='#F7EEDD' ","","","25","","",Celda,"si")

    Celda = ""

    Celda = Celda & ofv.gencelda("","UTI STYLE='BACKGROUND-color:#DDDDDD;color:#888888;' ","","5%","","","","CODIGO","si")
    Celda = Celda & ofv.gencelda("","UTI STYLE='BACKGROUND-color:#DDDDDD;color:#888888;' ","","20%","","","","CUENTA","si")
    Celda = Celda & ofv.gencelda("","UTI STYLE='BACKGROUND-color:#DDDDDD;color:#888888;' ","","30%","","","","CARACTER","si")
    Celda = Celda & ofv.gencelda("","UTI STYLE='BACKGROUND-color:#DDDDDD;color:#888888;' ","","15%","","","","F.REG.CONTABLE","si")
    Celda = Celda & ofv.gencelda("","UTI STYLE='BACKGROUND-color:#DDDDDD;color:#888888;' ","","20%","","","","MONTO","si")

    Formulario = Formulario & ofv.genrow("","","","","25","","",Celda,"si")

          WKADMMOD = "N"

    Celda = ""
       Strsql = "select * FROM ORTRACER WHERE CODTRACER = '" & SCODTRACER & "' "

''  response.write strsql & "<br>"
       set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if rs2.eof then

          WKCODIGO = "-----"
          WKDESCRI = "-- SIN DEFINIR --"
          WKCARAC = "-- SIN DEFINIR --"
          WKFCONT = "-- SIN DEFINIR --"
          WKMONTO = "-- SIN DEFINIR --"

       ELSE

          IF RS2("TRCMARCASEND") = "S" OR  RS2("TRCMARCASEND") = "R" OR RS2("TRCMARCASEND") = "J"  OR  RS2("TRCMARCAACTP") = "RM" OR RS2("TRCMARCAACTP") = "KR"   OR RS2("TRCMARCAACTP") = "XR" THEN

          ELSE
             WKADMMOD = "S"
          END IF

          WKCODIGO = RS2(0)

          if (rs2(29) & "") = "" then
             dFechan = "-------"
          else
             dFechan = RS2(29)
          end if

          WKDESCRI = dFechan

          sInputx = "-------"

          strsql = "Select descri From ORTIPOCARACTER where OBJID = 0" & rs2(24)
          set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if not rsniv.eof then sInputx = rsniv(0)
          call ofv.cerrarconsulta(rsniv)

          WKCARAC = sInputx

          if clng(rs2(5)) = 0 then
             dFechan = "-------"
          else
             dFechan = GTFechadorSH(rs2(5))
          end if
          WKFCONT = dFechan

          if clng(rs2(12)) = 0 then
             dFechan = "-------"
          else
             dFechan = RS2(12)
          end if
          WKMONTO = dFechan

       END IF

       Celda = Celda & ofv.gencelda("","VTI onclick='document." & sForm & ".submit();' title='ver evento.' STYLE='cursor:pointer;color:#000000;background-color:#ddddff;border-style:outset;border-width:2;' ","","","","","",WKCODIGO,"si")
       Celda = Celda & ofv.gencelda("","uTI STYLE='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",WKDESCRI,"si")
       Celda = Celda & ofv.gencelda("","uTI STYLE='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",WKCARAC,"si")
       Celda = Celda & ofv.gencelda("","uTI STYLE='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",WKFCONT,"si")
       Celda = Celda & ofv.gencelda("","uTI STYLE='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",WKMONTO,"si")

       sAccion = "ORBASERIESGOSADD.ASP?MODAL=CSN&CODTRACER=" & RS2(1)

       FILAX = ofv.genrow("","","","","25","","",Celda,"si")
       Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"_BLANK","post","",FilaX,"si")

  sHTML = sHTML & "<BR>" & ofv.GenTabla("","AAA STYLE='border-style:solid;border-width:1;border-color:#993333;' ","","98% ","","0","","0","0","","",Formulario,"si")

  sHTML = sHTML & "<BR><BR>"



    Celda = ""


    Formulario = ""

  Strsql = "select * FROM ORTRACER WHERE TRCCODTRACERP = '" & SCODTRACER & "' ORDER BY OBJID DESC   "

''  response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then


  ELSE
    posI = 0

          xTrColor = "#FFFFFF"

    sInput = "&nbsp;"
    Celda = ofv.gencelda("","VTI colspan=5","","","","","","EVENTOS RELACIONADOS","si")
    Formulario = ofv.genrow("","AAA bgcolor='#ddddff' ","","","25","","",Celda,"si")


    Celda = ""
    'REGISTRO
    Celda = Celda & ofv.gencelda("","UTI STYLE='BACKGROUND-color:#DDDDDD;color:#888888;' ","","5%","","","","CODIGO","si")
    Celda = Celda & ofv.gencelda("","UTI STYLE='BACKGROUND-color:#DDDDDD;color:#888888;' ","","20%","","","","CUENTA","si")
    Celda = Celda & ofv.gencelda("","UTI STYLE='BACKGROUND-color:#DDDDDD;color:#888888;' ","","30%","","","","CARACTER","si")
    Celda = Celda & ofv.gencelda("","UTI STYLE='BACKGROUND-color:#DDDDDD;color:#888888;' ","","15%","","","","F.REG.CONTABLE","si")
    Celda = Celda & ofv.gencelda("","UTI STYLE='BACKGROUND-color:#DDDDDD;color:#888888;' ","","20%","","","","MONTO","si")
    Formulario = Formulario & ofv.genrow("","AAA bgcolor='#ddddff' ","","","25","","",Celda,"si")


  Do While (Not rs.EOF)
    sForm = "form" & posi
    Filax  = ""

    WKEVENRELAC = WKEVENRELAC + 1

    Celda = ""

          WKCODIGO = RS(0)

          if (rs(29) & "") = "" then
             dFechan = "-------"
          else
             dFechan = RS(29)
          end if

          WKDESCRI = dFechan

          sInputx = "-------"

          strsql = "Select descri From ORTIPOCARACTER where OBJID = 0" & rs(24)
          set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if not rsniv.eof then sInputx = rsniv(0)
          call ofv.cerrarconsulta(rsniv)

          WKCARAC = sInputx

          if clng(rs(5)) = 0 then
             dFechan = "-------"
          else
             dFechan = GTFechadorSH(rs(5))
          end if
          WKFCONT = dFechan

          if cDBL(rs(12)) = 0 then
             dFechan = "-------"
          else
             dFechan = RS(12)
          end if
          WKMONTO = dFechan

       Celda = Celda & ofv.gencelda("","VTI onclick='document." & sForm & ".submit();' title='ver evento.' STYLE='cursor:pointer;color:#000000;background-color:#ddddff;border-style:outset;border-width:2;' ","","","","","",WKCODIGO,"si")
       Celda = Celda & ofv.gencelda("","uTI STYLE='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",WKDESCRI,"si")
       Celda = Celda & ofv.gencelda("","uTI STYLE='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",WKCARAC,"si")
       Celda = Celda & ofv.gencelda("","uTI STYLE='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",WKFCONT,"si")
       Celda = Celda & ofv.gencelda("","uTI STYLE='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",WKMONTO,"si")

       sAccion = "ORBASERIESGOSADD.ASP?MODAL=CSN&CODTRACER=" & RS(1)

       FILAX = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","25","","",Celda,"si")
       Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"_BLANK","post","",FilaX,"si")

       if xTrColor = "#FFFFFF" then
          xTrColor = "#F0F0F0"
       ELSE
          xTrColor = "#FFFFFF"
       END IF

    posi = posi + 1
    rs.Movenext
  Loop
  END IF
  sHTML = sHTML & ofv.GenTabla("","AAA STYLE='border-style:solid;border-width:1;border-color:#888888;'","","98% ","","0","","0","0","","",Formulario,"si")





  'Encabezados
  Formulario = ""

    Celda = ""


    Formulario = ""

  Strsql = "select * FROM ORTRACER WHERE CODTRACER = '" & SCODTRACER & "' "

''  response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then

     Celda = ""
     Celda = Celda & ofv.GenCelda("","vtG ","CENTER","","","","","NO SE ENCUANTRA EL EVENTO SOLICITADO","si")
     Fila = ofv.GenRow("","","","","","","",Celda,"si")
     Tabla = ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-COLOR:#993333;' BGCOLOR='#F9F9F9' ","","300 ","","0","","8","8","","",Fila,"si")

  sHTML = sHTML & "<CENTER><BR><BR><BR><BR>" & Tabla & "</CENTER>"


  ELSE
    posI = 0



    sInput = "&nbsp;"
    Celda = ofv.gencelda("","VTI colspan=6 ","","","","","","EVENTO PRECEDENTE","si")
    Formulario = ofv.genrow("","AAA bgcolor='#ddddff' ","","","25","","",Celda,"si")

    Celda = ""

    Celda = Celda & ofv.gencelda("","UTI STYLE='BACKGROUND-color:#DDDDDD;color:#888888;' ","","5%","","","","CODIGO","si")
    Celda = Celda & ofv.gencelda("","UTI STYLE='BACKGROUND-color:#DDDDDD;color:#888888;' ","","20%","","","","CUENTA","si")
    Celda = Celda & ofv.gencelda("","UTI STYLE='BACKGROUND-color:#DDDDDD;color:#888888;' ","","30%","","","","CARACTER","si")
    Celda = Celda & ofv.gencelda("","UTI STYLE='BACKGROUND-color:#DDDDDD;color:#888888;' ","","15%","","","","F.REG.CONTABLE","si")
    Celda = Celda & ofv.gencelda("","UTI STYLE='BACKGROUND-color:#DDDDDD;color:#888888;' ","","20%","","","","MONTO","si")
    Celda = Celda & ofv.gencelda("","UTI STYLE='BACKGROUND-color:#DDDDDD;color:#888888;' ","","10%","","","","&nbsp;","si")

    Formulario = Formulario & ofv.genrow("","","","","25","","",Celda,"si")


    WKFORMHEADER = Formulario

    Formulario = ""





    sForm = "formP" & posi
    Filax  = ""
       Celda = ""

    IF (RS(22) & "") = "" THEN

       IF  WKEVENRELAC = 0 THEN

           Formulario = WKFORMHEADER

           BTIT1 = "----"

           Celda = Celda & ofv.gencelda("","UTG COLSPAN=5 ","CENTER","","","","","NO SE REGISTRA UN EVENTO PRECEDENTE","si")
           IF WKADMMOD = "S" THEN
              BTIT1 = "<A CLASS=CC HREF='ORBASERIESGOS.asp?CODTRACER=" & SCODTRACER & "&MODAL=ENCA&MODAL2="& smodal2 &"' >RELACIONAR</A>"
           END IF 
           Celda = Celda & ofv.gencelda("","VT STYLE='background-color:#F7EEDD;color:#ffffff;border-style:solid;border-width:1;border-color:#888888;'","CENTER","5%","","","",BTIT1,"si")
       ELSE

           Celda = ""

       END IF



    ELSE


       Formulario = WKFORMHEADER

       Strsql = "select * FROM ORTRACER WHERE CODTRACER = '" & RS(22) & "' "

''  response.write strsql & "<br>"
       set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if rs2.eof then

          WKCODIGO = "-----"
          WKDESCRI = "-- SIN DEFINIR --"
          WKCARAC = "-- SIN DEFINIR --"
          WKFCONT = "-- SIN DEFINIR --"
          WKMONTO = "-- SIN DEFINIR --"

       ELSE

          WKCODIGO = RS2(0)

          if (rs2(29) & "") = "" then
             dFechan = "-------"
          else
             dFechan = RS2(29)
          end if

          WKDESCRI = dFechan

          sInputx = "-------"

          strsql = "Select descri From ORTIPOCARACTER where OBJID = 0" & rs2(24)
          set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          if not rsniv.eof then sInputx = rsniv(0)
          call ofv.cerrarconsulta(rsniv)

          WKCARAC = sInputx

          if clng(rs2(5)) = 0 then
             dFechan = "-------"
          else
             dFechan = GTFechadorSH(rs2(5))
          end if
          WKFCONT = dFechan

          if clng(rs2(12)) = 0 then
             dFechan = "-------"
          else
             dFechan = RS2(12)
          end if
          WKMONTO = dFechan

       END IF

       Celda = Celda & ofv.gencelda("","VTI onclick='document." & sForm & ".submit();' title='ver evento.' STYLE='cursor:pointer;color:#000000;background-color:#ddddff;border-style:outset;border-width:2;' ","","","","","",WKCODIGO,"si")
       Celda = Celda & ofv.gencelda("","uTI STYLE='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",WKDESCRI,"si")
       Celda = Celda & ofv.gencelda("","uTI STYLE='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",WKCARAC,"si")
       Celda = Celda & ofv.gencelda("","uTI STYLE='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",WKFCONT,"si")
       Celda = Celda & ofv.gencelda("","uTI STYLE='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",WKMONTO,"si")

       BTIT1 = "----"

       IF WKADMMOD = "S" THEN
          BTIT1 = "<A CLASS=CC HREF='ORBASERIESGOSENCA.asp?CODTRACER=" & SCODTRACER & "&MODAL=ENR&OPCION=DESAF&PRETRACER=" & RS2(1) & "&MODAL2="& SMODAL2 &"' >DESAFECTAR</A>"
       END IF

       Celda = Celda & ofv.gencelda("","VT title='Relacionar otro evento.' STYLE='background-color:#F7EEDD;color:#ffffff;border-style:outset;border-width:2;'","CENTER","5%","","","",BTIT1,"si")

    END IF

    sAccion = "ORBASERIESGOSADD.ASP?MODAL=CSN&CODTRACER=" & RS2(1)

    FILAX = ofv.genrow("","","","","25","","",Celda,"si")
    Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"_BLANK","post","",FilaX,"si")


  END IF
  sHTML = sHTML & "<BR><BR>" & ofv.GenTabla("","AAA STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","98% ","","0","","0","0","","",Formulario,"si")


  sHTML = sHTML & "<BR><BR>"




  'Botonera
  Agregar = ""
  Buscar  = ""
  Celda   = ""
  IF SMODAL2 = "VTR" THEN  SMODAL="VTR"
  sAccion = "ORBASERIESGOSADD.asp?CODTRACER=" & SCODTRACER & "&MODAL=" & SMODAL & "&CantRegMover=" & CantRegMoverX & ""
  Celda = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)





  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML



%>