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

  sMODAL       = request.querystring("MODAL")
  sMODALMAIN   = request.querystring("MODALMAIN")
  sCODTRACER   = request.querystring("CODTRACER")
  sFECHA       = request.querystring("FECHA")
  sOPCION       = request.querystring("OPCION")
  sOPCSET       = request.querystring("OPCSET")
  sOPX       = request.querystring("OPX")

  IF SMODAL    = "" THEN SMODAL = "AAG"
 ' response.write request.querystring
  SMODALORIG = SMODAL

  IF SOPX="DEF" THEN SOPCSET="DEF"

  IF SMODAL = "DOBLIGA" THEN
     IF SOPCSET <> "" THEN
        SESSION("OPCSET") = SOPCSET
     ELSE

        SOPCSET = SESSION("OPCSET")

     END IF


  END IF



'response.write request.querystring
  sHTML = ofv.FormHeader(session("FormN"))


    Celda = ""
    APLCHECK()


   Celda = Celda & ofv.GenCelda("","uti  ","","90%","","","","BASE DE EVENTOS DE RIESGO","si")
   Celda = Celda & ofv.GenCelda("","ut   ","","2%","","","","&nbsp;","si")
   Fila  = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML & ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ","","100% ","30","0","","-1","1","","",Fila,"si")




  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)
  set cn3 = ofv.conectar(ofv.strconn2)




  IF SOPCION = "RECALC" THEN

     XXXX = ORRECALCEVENTO(sCODTRACER,"NO")


  END IF

  IF SOPCION = "RECALCSI" THEN

     XXXX = ORRECALCEVENTO(sCODTRACER,"SI")


  END IF


'''''''''''''''''''''''''''''''''''''''
'' INCORPORO LINEA DE NEGOCIOS

  IF SOPCION = "ADDLINEA" THEN

        strsql = "Insert into ORDISTRIBLN "
        strsql = strsql & " (CODTRACER,LINNEGO,PORCENTAJE"
        strsql = strsql & " ) Values ("
        strsql = strsql &  "'" & sCODTRACER & "','0','0'"
        strsql = strsql & ")"
  '''response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

  end if



'''''''''''''''''''''''''''''''''''''''
'' INCORPORO EVENTO EN BASE DE EVENTOS DE PERDIDA

  if sMODAL = "AAG" then

     sCODTRACER = GTkey("001")
     sFECREGIS  = CLNG(DATE()) - clng(cdate("01-01-1900"))
  '''   sFECORIG   = CLNG(DATE()) - clng(cdate("01-01-1900"))
     sFECORIG   = sFECREGIS

        strsql = "Insert into ORTRACER "
        strsql = strsql & " (CODTRACER,FECREGIS,FECORIG,FECIDENT, FECCONTA,"
        strsql = strsql & " TRCPROCESO, TRCPRODSER, TRCNEGOLINE, TRCIDENTACT, TRCEVENTO, "
        strsql = strsql & " TRCOBSERV, TRCIMPBRUTO, TRCIMPACRECEN, TRCIMPOTROS, TRCIMPRECUPERA, "
        strsql = strsql & " TRCIMPNETO, TRCIMPSEGURO,TRCESTADO, TRCUSRREG,"
        strsql = strsql & " fecfinal, trccaracter, trcproducto, trcriesgoasoc, trcsecsoporte,"
        strsql = strsql & " TRCEVENTORIESGO1, TRCEVENTORIESGO2, TRCLOTEING, TRCTIPOING, TRCINFOING,"
        strsql = strsql & " FECFINENR, TRCIMPDESAFECTA,TRCDISTRIB,TRCPORCDISTRIB,TRCCOSTOPOR,"
        strsql = strsql & " TRCMARCASEND, TRCMARCAACTP,TRCMARCAACTH,TRCFECACTP,TRCFECACTH,TRCGRUPOWNER"
        strsql = strsql & " ) Values ("
        strsql = strsql &  "'" & sCODTRACER & "','" & SFECREGIS & "','" & sFECORIG & "',"
        strsql = strsql &  "'"&sFECORIG&"','0',"
        strsql = strsql & "'0','0','0','0','0','XXX','0','0','0','0','0','0','REG',"
        strsql = strsql & "'" & SESSION("USUARIO") & "',"
        strsql = strsql & "'0','0','0','4','0',"
        strsql = strsql & "'0','0','0','1','MANUAL','0','0','N','0','0',"
        strsql = strsql & "'N','N','N','0','0','" & SESSION("SEGGRP") & "'"
        strsql = strsql & ")"
  ''' response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

        sMODAL = "VTR"
         SMODALORIG = SMODAL

  end if


'''''''''''''''''''''''''''''''''''''''
'' INCORPORO EVENTO EN BASE DE EVENTOS DE PERDIDA EL PATRO DE DATOS OBLIGATORIOS

  if sMODAL = "DOBLIGA" then

    scomK = "1" & chr(9) & "0"
    scomD = "SI" & chr(9) & "NO"

     sCODTRACER = "DOBLIGA"


    if SOPCSET = "DEF" THEN
       sCODTRACER = "DOBLIGA"
    ELSE
       sCODTRACER = "DOBLIGA" & SOPCSET
    END IF

     sFECREGIS  = CLNG(DATE()) - clng(cdate("01-01-1900"))

     ''   strsql = strsql & "DELETE FROM ORTRACER WHERE CODTRACER = '" & SCODTRACER & "' "
  '''response.write strsql & "<br>"
    '    call ofv.crearconsultaEx(StrSql,cn,1,volver)


  Strsql = "  select * FROM ORTRACER WHERE CODTRACER = '" & SCODTRACER & "' "
'  RESPONSE.WRITE STRSQL
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if rs.eof then


        strsql = "Insert into ORTRACER "
        strsql = strsql & " (CODTRACER,FECREGIS,FECORIG,FECIDENT, FECCONTA,"
        strsql = strsql & " TRCPROCESO, TRCPRODSER, TRCNEGOLINE, TRCIDENTACT, TRCEVENTO, "
        strsql = strsql & " TRCOBSERV, TRCIMPBRUTO, TRCIMPACRECEN, TRCIMPOTROS, TRCIMPRECUPERA, "
        strsql = strsql & " TRCIMPNETO, TRCIMPSEGURO,TRCESTADO, TRCUSRREG,"
        strsql = strsql & " fecfinal, trccaracter, trcproducto, trcriesgoasoc, trcsecsoporte,"
        strsql = strsql & " TRCEVENTORIESGO1, TRCEVENTORIESGO2, TRCLOTEING, TRCTIPOING, TRCINFOING,"
        strsql = strsql & " TRCCLIENTE, TRCCTACBLE, trcecoresp, TRCIMPDESAFECTA,TRCDISTRIB,TRCPORCDISTRIB,TRCCOSTOPOR,trcAsientoCont"



        strsql = strsql & " ) Values ("
        strsql = strsql &  "'" & sCODTRACER & "','" & SFECREGIS & "','0',"
        strsql = strsql &  "'0','0',"
        strsql = strsql & "'0','0','0','0','0','0','0','0','0','0','0','0','DOBLIGA',"
        strsql = strsql & "'0',"
        strsql = strsql & "'0','1','0','0','0',"
        strsql = strsql & "'0','0','0','0','0',"
        strsql = strsql & "'0','0','0','0','NO','0','0','0'"
        strsql = strsql & ")"
  '''response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

   END IF

  end if


'''''''''''''''''''''''''''''''''''''''

  vuelta        = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_MODAL=" & SMODAL & "_CODTRACER=" & SCODTRACER

  'Encabezados
  Formulario = ""


    Filax  = ""

    Formulario = ""

          WSPSEN1 = "<FONT COLOR='#6699CC' SIZE=3>&nbsp;&nbsp;&#0149;</FONT>"
          WSPSEN2 = "<FONT COLOR='#DD0000' SIZE=3>&nbsp;&nbsp;&#0149;</FONT>"



  DIM ATABESTADOS(10,3)

  FOR XII=0 TO 10

      ATABESTADOS(XII,0) = 0

  NEXT

  wkestadoPERM = 3

  Strsql = "  select * FROM ORTRACER WHERE CODTRACER = '" & SCODTRACER & "' "

 'response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if NOT rs.eof then

'''''''''''''''' ver estado ''''''''''''''''''''''''''''''

    wkestado = rs(17) & ""

    Strsql = "Select * From ORESTADOS WHERE CODIGO = '" & wkestado & "' "
    set rsXX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    IF NOT rsXX.eof THEN

       wkestadodescri = rsxx(2) & ""

       FOR XII=1 TO 10

           ATABESTADOS(XII,0) = rsXX(3 + XII)

           IF ATABESTADOS(XII,0) > 0 THEN

              Strsql = "Select * From ORESTADOS WHERE OBJID = 0" & ATABESTADOS(XII,0)
              set rsXX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
              IF NOT rsXX1.eof THEN
                 ATABESTADOS(XII,1) = rsXX1(1)  'CODIGO
                 ATABESTADOS(XII,2) = rsXX1(2)  'DESCRIPCION
              END IF


           END IF

       NEXT


'''''''''''''''''''''''''''''''''''''''''''''



        Strsql = "select B.* from  orestadosper b "
        Strsql = Strsql & " where B.ESTADO = " & rsxx(0) & " AND B.PERMISO > 0 "

        Strsql = Strsql & " and ( "
        Strsql = Strsql & " (b.usrgrp = 'TODOS') OR "
        Strsql = Strsql & " (B.USRGRP = 'USR-" & Session("Usuario") & "')  "

        if  Session("DELLOG") = "S" then
            XMFTGRP = split(session("SegGrpLst"),chr(9))
            xcount  = ubound(XMFTGRP) + 1
            for XMFI = 0 to ubound(XMFTGRP)
                Strsql = Strsql & " OR (B.USRGRP = 'GRP-" & XMFTGRP(XMFI) & "') "
            next
        else
            Strsql = Strsql & " OR (B.USRGRP = 'GRP-" & SESSION("SEGGRP") & "') "
        end if

             Strsql = Strsql & " ) ORDER BY B.PERMISO"

             set rsxX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

             IF not rsxX1.eof THEN

                wkestadoPERM = rsxX1("PERMISO")

             END IF



'''''''''''''''''''''''''''''''''''''''''''''

    END IF

    SMODALSET = SMODAL
   '  RESPONSE.WRITE wkestado & "22222222222222"

    IF wkestadoPERM = 3 OR wkestadoPERM = 2 THEN

       IF wkestado = "NO" OR wkestado="DOBLIGA" THEN

          SMODALSET = SMODAL

       ELSE

          SMODALSET = "CSN"

       END IF


    END IF

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''



    XORWKESRAIZ = "NO"
    XORWKESRAIZDSC = "&nbsp;"
    XORWKENCADSC = "&nbsp;"
    XORWKENCA = "NO"
    XORWKENCACOD = "&nbsp;"

    Strsql = "select * FROM ORTRACER WHERE TRCCODTRACERP = '" & SCODTRACER & "' "

'  response.write strsql & "<br>"
    set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
    if NOT rsX.eof then
       XORWKESRAIZ = "SI"
       XORWKESRAIZDSC = "&nbsp;|&nbsp;<FONT COLOR='#FF0000' >(EVENTO&nbsp;RAIZ)</FONT>"
    END IF


    sForm = "form" & RS(0)



    IF (RS(22) & "") <> "" THEN

       Strsql = "select * FROM ORTRACER WHERE CODTRACER = '" & RS(22) & "' "

'response.write strsql & "<br>" & session("APLMDLOR") & ""
       set rsX = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if NOT rsX.eof then
          XORWKENCACOD = RSX(0)
       END IF

        XORWKENCA = "SI"
        XORWKENCADSC = "&nbsp;|&nbsp;<FONT COLOR='#F7AA22' >(RELACIONADO&nbsp;A&nbsp;EVENTO&nbsp;" & XORWKENCACOD & ")</FONT>&nbsp;"

    END IF



   if SMODALSET = "VTR" then
    SINPUT = "REGISTRACION DE EVENTO - CODIGO:&nbsp;" & RS(0) & XORWKESRAIZDSC & XORWKENCADSC
    SINPUT2 = "&nbsp;"
    WSPCOLOR = "#000088"

   '' if  (session("APLMDLOR") & "") = "BE" then
    BTIT2 = "<A CLASS=TT HREF='ORBASERIESGOSENCA.asp?CODTRACER=" & SCODTRACER & "&MODAL=" & SMODAL & "&CantRegMover=" & CantRegMoverX & "' >ENCADENAMIENTO</A>"
    Celda = ""
	 '   Celda = Celda & ofv.gencelda("","VT STYLE='background-color:#000088;color:#888888;' ","CENTER","50%","","","",BTIT1,"si")

	 Celda = Celda & ofv.gencelda("","VT STYLE='background-color:#000088;color:#888888;' ","CENTER","50%","","","",BTIT2,"si")
	 Fila2  = ofv.genrow("","","","","20","","",Celda,"si")
	 SINPUT2 = ofv.GenTabla("","aaa BORDER=1 STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","150","","0","","1","1","","",Fila2,"si")
   '' end if
   elseif SMODALSET = "CSN" then
      SINPUT = "CONSULTA DE EVENTO - CODIGO:&nbsp;" & RS(0) & XORWKESRAIZDSC & XORWKENCADSC
      SINPUT2 = "&nbsp;"
      WSPCOLOR = "#993333"

      BTIT2 = "<A CLASS=TT HREF='ORBASERIESGOSENCA.asp?CODTRACER=" & SCODTRACER & "&MODAL=" & SMODAL & "&CantRegMover=" & CantRegMoverX & "' >ENCADENAMIENTO</A>"
      Celda = ""
      Celda = Celda & ofv.gencelda("","VT STYLE='background-color:#000088;color:#888888;' ","CENTER","50%","","","",BTIT2,"si")
      Fila2  = ofv.genrow("","","","","20","","",Celda,"si")
      SINPUT2 = ofv.GenTabla("","aaa BORDER=1 STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","150","","0","","1","1","","",Fila2,"si")

   elseif SMODALSET = "DOBLIGA" then
    SINPUT = "PATRON DE DATOS OBLIGATORIOS - "

    IF SOPCSET = "DEF" THEN
       SINPUT = SINPUT & "DEFAULT"
    ELSE
       sInputx = "---"
       strsql = "Select descri From ORTIPOCARACTER where OBJID = 0" & SOPCSET
      ' response.write strsql
       set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rsniv.eof then sInputx = rsniv(0)
       call ofv.cerrarconsulta(rsniv)

       SINPUT = SINPUT & "ESPECIAL " & sInputx
    END IF


    SINPUT2 = "&nbsp;"
    WSPCOLOR = "#000000"

   else
    SINPUT = "ENRIQUECIMIENTO DE EVENTO - CODIGO:&nbsp;" & RS(0) & XORWKESRAIZDSC & XORWKENCADSC
    SINPUT2 = "&nbsp;"
    WSPCOLOR = "#000088"

  '  BTIT1 = "<A CLASS=TT HREF='ORBASERIESGOSCONT.asp?CODTRACER=" & SCODTRACER & "&MODAL=" & SMODAL & "' >CONTABILIZACION</A>"
    BTIT2 = "<A CLASS=TT HREF='ORBASERIESGOSENCA.asp?CODTRACER=" & SCODTRACER & "&MODAL=" & SMODAL & "&CantRegMover=" & CantRegMoverX & "' >ENCADENAMIENTO</A>"

    Celda = ""
 '   Celda = Celda & ofv.gencelda("","VT STYLE='background-color:#000088;color:#888888;' ","CENTER","50%","","","",BTIT1,"si")

    Celda = Celda & ofv.gencelda("","VT STYLE='background-color:#000088;color:#888888;' ","CENTER","50%","","","",BTIT2,"si")
    Fila2  = ofv.genrow("","","","","20","","",Celda,"si")
    SINPUT2 = ofv.GenTabla("","aaa BORDER=1 STYLE='border-style:solid;border-width:1;border-color:#888888;' ","","150","","0","","1","1","","",Fila2,"si")

   end if


    Celda = ofv.gencelda("","VTI COLSPAN=5 STYLE='background-color:" & WSPCOLOR & ";color:#ffffff;'  ","","","","","",SINPUT,"si")
    Celda = Celda & ofv.gencelda("","VTI  STYLE='background-color:#DDDDDD;color:#ffffff;'  ","RIGHT","","","","",SINPUT2,"si")

    Filax  = Filax & ofv.genrow("","","","","30","","",Celda,"si")


      sInputD1 = ""
      sInputD2 = ""

    IF SMODALSET = "CSN" then

    ELSEif SMODALSET = "DOBLIGA" then

    ELSE

      sFECREGIS  = CLNG(DATE()) - clng(cdate("01-01-1900"))

      sInputD1 = ofv.GenerarInput(rs(44).name,"A","HIDDEN",sForm,"1","dt","")            ''' MARCA DE ACTUALIZACION PROPIA
      sInputD2 = ofv.GenerarInput(rs(46).name,sFECREGIS,"HIDDEN",sForm,"1","dt","")      ''' FECHA ULTIMA ACTUALIZACION PROPIA

    END IF


    Celda = ""
   ' if (session("APLMDLOR") & "") = "BE" then
  '  Celda =  ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","Empresa:&nbsp;","si")
  '  Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","15%","","","",rs("idempresa"),"si")
  '  Celda =  ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","Nro. LOTE:&nbsp;","si")
  '  Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","15%","","","",RS("NROLOTE"),"si")
  '  Celda =  ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","USUARIO:&nbsp;","si")
  '  Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","15%","","","",RS("USUARIO"),"si")
  '  end if



    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","F.CARGA:&nbsp;","si")

    if clng(rs(2)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSH(rs(2))
    end if
    SINPUT = dFechan & sInputD1 & sInputD2
    Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","15%","","","",sInput,"si")



    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","ESTADO:&nbsp;","si")

    'ESTADO
  ''''''  wkestado = rs(17) & ""

    select case wkestado

           case "CAN"

                ' sMODAL = "CSN"

                 if clng("0" & rs(35)) = 0 then
                    dFechan = ""
                 else
                    dFechan = GTFechadorSH(rs(35))
                 end if

                 wkestadoD = "<font color='#888888' ><B>" & wkestadodescri & "&nbsp;&nbsp;(" & dFechan & ")&nbsp;</B></font>"


           case "REG"

                 wkestadoD = "<font color='#993333' >" & wkestadodescri & "</font>"


           case "ENR"

                 wkestadoD = "<font color='#009900' >" & wkestadodescri & "</font>"

           case "ENRF"

                 wkestadoD = "<font color='#009900' >" & wkestadodescri & "</font>"


           case "FIN"

             '    sMODAL = "CSN"

                 if clng("0" & rs(35)) = 0 then
                    dFechan = ""
                 else
                    dFechan = GTFechadorSH(rs(35))
                 end if

                 wkestadoD = "<font color='#000000' ><B>" & wkestadodescri & "&nbsp;&nbsp;(" & dFechan & ")&nbsp;</B></font>"

           case "DOBLIGA"

                 wkestadoD = "<font color='#000000' >PATRON DE DATOS</font>"

           CASE ELSE


                 wkestadoD = "<font color='#000000' >" & wkestadodescri & "</font>"

                 IF wkestado = "NO" OR wkestado = "" THEN wkestadoD = "<font color='#000080' >EN REVISION</font>"


    end select
    'sMODAL="DOBLIGA"
    sInput = wkestadoD
    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","15%","","","",sInput,"si")


    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","F.INICIO:&nbsp;","si")

    if clng(rs(3)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSH(rs(3))
    end if
    SINPUT = dFechan

    IF SMODALSET = "CSN" then
      Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","15%","","","",sInput,"si")

    ELSEif SMODALSET = "DOBLIGA" then

       IF CSTR(RS(3) & "") = "0" THEN
          WSPSEN = WSPSEN1
       ELSE
          WSPSEN = WSPSEN2
       END IF

       sInput = ofv.gencomboFX(sForm,rs(3).name,rs(3),scomK,scomD," ") & WSPSEN
       sInput = REPLACE(sInput,"<TABLE ","<TABLE STYLE='DISPLAY:INLINE;'  ")
      Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","15%","","","",sInput,"si")

    ELSE


      sInputD = ofv.GenerarInput(rs(3).name,rs(3).value,"HIDDEN",sForm,"1","dt","")
      sInput = "<u>" & sinput & "</u>" & sInputD
      Celda = Celda & ofv.gencelda("","UTi title='Modificar fecha del incidente.'  " & GTCSETCALREX(sForm,rs(3).name,"S","N","",0,15,0,2,0,0,0,rs(4),0,0) & "  style='cursor:pointer;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","15%","","","",sInput,"si")
    '  Celda = Celda & ofv.gencelda("","UTi title='Modificar fecha del incidente.'  " & GTCSETCALREX(sForm,rs(3).name,"S","N","",rs(2),18,0,2,0,0,RS(2) +  clng(cdate("01-01-1900"))) & "  style='cursor:pointer;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","15%","","","",sInput,"si")
    END IF

    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")


    Celda = ""




    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","F.DESCUBRIMIENTO:&nbsp;","si")



    if clng(rs(4)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSH(rs(4))
    end if
    SINPUT = dFechan

    IF SMODALSET = "CSN" then
       Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    ELSEif SMODALSET = "DOBLIGA" then

       IF CSTR(RS(4) & "") = "0" THEN
          WSPSEN = WSPSEN1
       ELSE
          WSPSEN = WSPSEN2
       END IF

       sInput = ofv.gencomboFX(sForm,rs(4).name,rs(4),scomK,scomD," ") & WSPSEN
       sInput = REPLACE(sInput,"<TABLE ","<TABLE STYLE='DISPLAY:INLINE;'  ")
      Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    ELSE

       sInputD = ofv.GenerarInput(rs(4).name,rs(4).value,"HIDDEN",sForm,"1","dt","")
       sInput = "<u>" & sinput & "</u>" & sInputD
       SFECHAFIN = clng(DATE()) - clng(cdate("01-01-1900"))
       Celda = Celda & ofv.gencelda("","UTi title='Modificar fecha deL DESCUBRIMIENTO.' " & GTCSETCALREX(sForm,rs(4).name,"S","N","",0,1,0,2,0,0,rs(3),SFECHAFIN,0,0) & " style='cursor:pointer;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")
    '   Celda = Celda & ofv.gencelda("","UTi title='Modificar fecha deL DESCUBRIMIENTO.' " & GTCSETCALR(sForm,rs(4).name,"S","N","",rs(2),15,0,2,0,0) & " style='cursor:pointer;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")
     '   GTCSETCALR(sForm,rs(3).name,"S","N","",rs(2),15,0,2,0,0)
    END IF

    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","F.R.CONTABLE:&nbsp;","si")



    if clng(rs(5)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSH(rs(5))
    end if
    SINPUT = dFechan

    IF SMODALSET = "CSN" then
       Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    ELSEif SMODALSET = "DOBLIGA" then

       IF CSTR(RS(5) & "") = "0" THEN
          WSPSEN = WSPSEN1
       ELSE
          WSPSEN = WSPSEN2
       END IF

       sInput = ofv.gencomboFX(sForm,rs(5).name,rs(5),scomK,scomD," ") & WSPSEN
       sInput = REPLACE(sInput,"<TABLE ","<TABLE STYLE='DISPLAY:INLINE;'  ")
      Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    ELSE


       sInputD = ofv.GenerarInput(rs(5).name,rs(5).value,"HIDDEN",sForm,"1","dt","")
       sInput = "<u>" & sinput & "</u>" & sInputD
     '  Celda = Celda & ofv.gencelda("","UTi title='Modificar fecha de contabilizacion.' " & GTCSETCALR(sForm,rs(5).name,"S","N","",rs(2),15,0,2,0,0) & " style='cursor:pointer;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")
       SFECHAFIN = clng(DATE()) - clng(cdate("01-01-1900"))
       Celda = Celda & ofv.gencelda("","UTi title='Modificar fecha de contabilizacion.' " & GTCSETCALREX(sForm,rs(5).name,"S","N","",0,15,0,2,0,0,rs(4),SFECHAFIN,0,0) & " style='cursor:pointer;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")
      ' Celda = Celda & ofv.gencelda("","UTi title='Modificar fecha del incidente.'  " &     GTCSETCALR(sForm,rs(3).name,"S","N","",rs(2),15,0,2,0,0) & "  style='cursor:pointer;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","15%","","","",sInput,"si")
    END IF

    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","F.FINALIZACION:&nbsp;","si")

    if clng(rs(23)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSH(rs(23))
    end if
    SINPUT = dFechan

    IF SMODALSET = "CSN" then
       Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    ELSEif SMODALSET = "DOBLIGA" then

       IF CSTR(RS(23) & "") = "0" THEN
          WSPSEN = WSPSEN1
       ELSE
          WSPSEN = WSPSEN2
       END IF

       sInput = ofv.gencomboFX(sForm,rs(23).name,rs(23),scomK,scomD," ") & WSPSEN
       sInput = REPLACE(sInput,"<TABLE ","<TABLE STYLE='DISPLAY:INLINE;'  ")
      Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    ELSE

       sInputD = ofv.GenerarInput(rs(23).name,rs(23).value,"HIDDEN",sForm,"1","dt","")
       sInput = "<u>" & sinput & "</u>" & sInputD
       SFECHAFIN = clng(DATE()) - clng(cdate("01-01-1900"))
       Celda = Celda & ofv.gencelda("","UTi title='Modificar fecha de Finalizacion.' " & GTCSETCALREX(sForm,rs(23).name,"S","N","",0,0,1,0,0,0,rs(4),SFECHAFIN,0,0) & " style='cursor:pointer;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")
    END IF






    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")


'''''''''''''''''''''''''''




    Celda = ""


    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","PROCESO:&nbsp;","si")

    IF SMODALSET = "CSN" then

       sInputx = "-------"

     if (session("APLMDLOR") & "") = "BE" then

       Strsql = "Select * From TABPROCESOS01 "
       strsql = strsql & " where objid = 0" & rs(6)

     else

       Strsql = "Select * From  OBJETOS   "
       'Strsql = "Select descri From  OBJETOS   "
       Strsql = Strsql & " where CODOBJETO = 0" & rs(6)

     end if

       set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rsniv.eof then sInputx = rsniv(3)
       call ofv.cerrarconsulta(rsniv)

       sInput = sInputx

    ELSEif SMODALSET = "DOBLIGA" then

       IF CSTR(RS(6) & "") = "0" THEN
          WSPSEN = WSPSEN1
       ELSE
          WSPSEN = WSPSEN2
       END IF

       sInput = ofv.gencomboFX(sForm,rs(6).name,rs(6),scomK,scomD," ") & WSPSEN
       sInput = REPLACE(sInput,"<TABLE ","<TABLE STYLE='DISPLAY:INLINE;'  ")

    ELSE

     if (session("APLMDLOR") & "") = "BE" then

       Strsql = "Select * From TABPROCESOS01 "
       strsql = strsql & " where "
       strsql = strsql & " nivel = '1' "

       sInput = ofv.GenerarCombo(sForm,RS(6).NAME,RS(6).value,Strsql,cn,0,3,"")
     else

       Strsql = "select a.CODOBJETO,(a.descri + '--' + b.ambiente) From OBJETOS a  "
       Strsql = Strsql & " inner join  (Select distinct x.CODOBJETO as tomaobj,c.descri as ambiente From OBJETOS x "
       Strsql = Strsql & " inner join relaobj x2 on (x.CODOBJETO = x2.CODOJASOC and x2.tipobjmaster = 'PROCESO'  ) "
       Strsql = Strsql & " inner join  rgambientes c on (c.codambiente=x2.codobjmaster ) "
       Strsql = Strsql & "  ) b on (a.codobjeto = b.tomaobj ) "
       Strsql = Strsql & "    ORDER BY a.descri,a.codobjeto desc "

       sInput = ofv.GenerarCombo(sForm,RS(6).NAME,RS(6).value,Strsql,cn,0,1,"")

     end if

      ' response.write strsql


    END IF



    'Celda = Celda & ofv.gencelda("","UTi COLSPAN=5 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")
    Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    Celda = Celda & ofv.gencelda("","UT  style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","NRO ASIENTO CONTABLE:&nbsp;","si")
    'ASIENTO CONTABLE


    if SMODALSET = "DOBLIGA" then

       IF CSTR(RS(41) & "") = "0" THEN
          WSPSEN = WSPSEN1
       ELSE
          WSPSEN = WSPSEN2
       END IF

       sInput = ofv.gencomboFX(sForm,rs(41).name,rs(41),scomK,scomD," ") & WSPSEN
       sInput = REPLACE(sInput,"<TABLE ","<TABLE STYLE='DISPLAY:INLINE;'  ")

    ELSEIF SMODALSET = "CSN" then

        SINPUT =  rs(41) & "&nbsp;"


    ELSE

       SINPUT="<input type=text class=dt name=" & rs(41).name & " align=left value='" & rs(41) & "'  onchange=document." & sForm & ".submit(); >"

    END IF


    Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")
    'USUARIO RESPONSABLE

    XXX=getInfoUser(RS("trcusrreg")&"")
    SINPUT=SESSION("SGU_nombre")
    Celda = Celda & ofv.gencelda("","UT  style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","RESPONSABLE:&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")


   ''    SMODALw    = SMODAL


   ''    IF SMODAL    = "VTR" and (wkestado = "NO" or wkestado = "VTR") THEN  SMODALw    = "REG"
   ''    IF SMODAL    = "VTR" and wkestado = "REG" THEN  SMODALw    = "ENR"
   ''    IF SMODAL    = "VTR" and wkestado = "ENR" THEN  SMODALw    = "ENR"
   ''    IF SMODAL    = "VTR" and wkestado = "ENRF" THEN  SMODALw    = "ENRF"

   ''    IF SMODAL    = "ENR" and (wkestado = "NO" or wkestado = "VTR") THEN  SMODALw    = "REG"
   ''    IF SMODAL    = "ENR" and wkestado = "REG" THEN  SMODALw    = "ENR"
   ''    IF SMODAL    = "ENR" and wkestado = "ENRF" THEN  SMODALw    = "ENRF"



  IF SMODALSET    = "ENR" OR SMODALSET = "DOBLIGA" OR SMODALSET = "CSN"  OR SMODALSET = "VTR"  THEN


    IF SMODALSET    = "ENR" OR SMODALSET = "DOBLIGA"   OR SMODALSET = "VTR"   THEN

       SMODALw    = wkestado
       IF (SMODALw = "NO" or SMODALw = "VTR") THEN  SMODALw    = "REG"


       sInput = ofv.GenerarInput(rs(17).name,SMODALW,"hidden",sForm,"1","dt","")
       Celda = ofv.gencelda("","AAA COLSPAN=6 STYLE='display:none;' ","","","","","",SINPUT,"si")
       Filax  = Filax & ofv.genrow("","AAA STYLE='display:none;' ","","","25","","",Celda,"si")

   END IF


    WKPROCESO = RS(6) & ""
    WKPROCSUB = RS(7) & ""
    WSPACTIVIDAD = RS(9) & ""

    sInputev  = "<FONT COLOR='#000000' >SUBPROCESO NO IDENTIFICADO</FONT>"

    IF SMODALSET = "DOBLIGA" OR (WKPROCESO <> "" AND WKPROCESO <> "0"  AND WKPROCESO <> "NO") THEN

       IF WKPROCSUB <> "" AND WKPROCSUB <> "0" AND WKPROCSUB <> "NO" THEN


     if (session("APLMDLOR") & "") = "BE" then

          sInputev  = "<FONT COLOR='#bbbbbb' >-----</FONT>"

     else

          sInputev  = "<FONT COLOR='#990000' >SUBPROCESO NO EVALUADO</FONT>"

          Strsql = "Select DISTINCT A.CODOBJETO From ORACTIVRIESGO A"
          Strsql = Strsql & " INNER JOIN OBJETOS B ON A.CODOBJETO = B.CODOBJETO "
          Strsql = Strsql & " WHERE A.CODOBJETO = 0" & WKPROCSUB & " "
          set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
          IF NOT RSX1.EOF THEN

             sInputev  = "<FONT COLOR='#008800' >SUBPROCESO EVALUADO</FONT>"

          END IF


     end if



      END IF



    IF SMODALSET = "CSN" then

       sInputx = "-------"




     if (session("APLMDLOR") & "") = "BE" then

       Strsql = "Select descripcion From TABPROCESOS01 "
       strsql = strsql & " where objid = 0" & rs(7)

     else

       Strsql = "Select DESCRI From  OBJETOS   "
       Strsql = Strsql & " where CODOBJETO = 0" & rs(7)

     end if


       set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rsniv.eof then sInputx = rsniv(0)
       call ofv.cerrarconsulta(rsniv)

       sInput = sInputx

    ELSEif SMODALSET = "DOBLIGA" then

       IF CSTR(RS(7) & "") = "0" THEN
          WSPSEN = WSPSEN1
       ELSE
          WSPSEN = WSPSEN2
       END IF

       sInput = ofv.gencomboFX(sForm,rs(7).name,rs(7),scomK,scomD," ") & WSPSEN
       sInput = REPLACE(sInput,"<TABLE ","<TABLE STYLE='DISPLAY:INLINE;'  ")

    ELSE


     if (session("APLMDLOR") & "") = "BE" then

    '   Strsql = "Select * From TABPROCESOS01 "
    '   strsql = strsql & " where objid = 0" & rs(6)


       Strsql = "Select A.* From TABPROCESOS01 A"
       strsql = strsql & " inner join TABPROCESOS01 b on (a.codproceso01 = b.codproceso and b.objid = 0" & RS(6) & ") "
       strsql = strsql & " where "
       strsql = strsql & " a.nivel = '2' order by a.secuen, a.descripcion "

 '   response.write "<br>" & strsql & "<br>" & chr(13)

       sInput =ofv.GenerarCombo(sForm,RS(7).NAME,RS(7).value,Strsql,cn,0,3,"")

     else

       'Strsql = "select a.cobj,B.DESCRI, a.* "
       'Strsql = Strsql & " from csrelpathdescri a "

       Strsql ="SELECT   ax.cobj,B.DESCRI , O1.descri, O2.descri AS Expr1, O3.descri AS Expr2, O4.descri AS Expr3, O5.descri AS Expr4, O6.descri AS Expr5, O7.descri AS Expr6,"
	   Strsql = strsql & "  O8.descri AS Expr7, ax.mc1, ax.mc2, ax.mc3, ax.mc4, ax.mc5, ax.mc6, ax.mc7, ax.cobj, ax.TSEC "
	   Strsql =Strsql & " FROM         dbo.CsRelPath ax LEFT OUTER JOIN"
	   Strsql =Strsql & "                       dbo.OBJETOS O1 ON ax.mc1 = O1.codobjeto LEFT OUTER JOIN"
	   Strsql =Strsql & "                       dbo.OBJETOS O2 ON ax.mc2 = O2.codobjeto LEFT OUTER JOIN"
	   Strsql =Strsql & "                       dbo.OBJETOS O3 ON ax.mc3 = O3.codobjeto LEFT OUTER JOIN"
	   Strsql =Strsql & "                       dbo.OBJETOS O4 ON ax.mc4 = O4.codobjeto LEFT OUTER JOIN"
	   Strsql =Strsql & "                       dbo.OBJETOS O5 ON ax.mc5 = O5.codobjeto LEFT OUTER JOIN"
	   Strsql =Strsql & "                       dbo.OBJETOS O6 ON ax.mc6 = O6.codobjeto LEFT OUTER JOIN"
	   Strsql =Strsql & "                       dbo.OBJETOS O7 ON ax.mc7 = O7.codobjeto LEFT OUTER JOIN"
       Strsql =Strsql & "                dbo.OBJETOS O8 ON ax.cobj = O8.codobjeto "
      ' Strsql = "select ax.cobj,B.DESCRI, ax.* "
      ' Strsql = Strsql & " from csrelpathdescri ax "
       Strsql = Strsql & " inner join objetos b on (ax.cobj = b.codobjeto) "
       Strsql = Strsql & " WHERE (B.CODVINOBJ = 'VN' or B.CODVINOBJ = 'CR' or B.CODVINOBJ = 'AI') AND ("
       Strsql = Strsql & " Ax.MC1 = 0" & RS(6) & " OR "
       Strsql = Strsql & " Ax.MC2 = 0" & RS(6) & " OR "
       Strsql = Strsql & " Ax.MC3 = 0" & RS(6) & " OR "
       Strsql = Strsql & " Ax.MC4 = 0" & RS(6) & " OR "
       Strsql = Strsql & " Ax.MC5 = 0" & RS(6) & " OR "
       Strsql = Strsql & " Ax.MC6 = 0" & RS(6) & " OR "
       Strsql = Strsql & " Ax.MC7 = 0" & RS(6) & "  "
       Strsql = Strsql & "  )"

 '   response.write "<br>" & strsql & "<br>" & chr(13)

       sInput =ofv.GenerarCombo(sForm,RS(7).NAME,RS(7).value,Strsql,cn,0,1,"")



     end if





    END IF



    Celda = ""
    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","SUBPROCESO:&nbsp;","si")

    Celda = Celda & ofv.gencelda("","UTi COLSPAN=3 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    if SMODALSET = "DOBLIGA" then sInputev = "&nbsp;"

    Celda = Celda & ofv.gencelda("","vTi COLSPAN=2 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInputev,"si")

    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")




    IF SMODALSET = "DOBLIGA" OR (WKPROCSUB <> "" AND WKPROCSUB <> "0" AND WKPROCSUB <> "NO") THEN


    IF SMODALSET = "CSN" then

       sInputx = "-------"


     if (session("APLMDLOR") & "") = "BE" then


     else

       Strsql = "select ISNULL(e.descri,D.DESCRIACC) as actdes "
       Strsql = Strsql & " from  objaccion d "
       Strsql = Strsql & " left join tipoaccalt e on (d.tipoaccion = e.tipoaccionalt) "
       Strsql = Strsql & " WHERE d.codaccion = '" & RS(9) & "'"

       set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rsniv.eof then sInputx = rsniv(0)
       call ofv.cerrarconsulta(rsniv)

     end if

       sInput = sInputx



    ELSEif SMODALSET = "DOBLIGA" then

       IF CSTR(RS(9) & "") = "0" THEN
          WSPSEN = WSPSEN1
       ELSE
          WSPSEN = WSPSEN2
       END IF

       sInput = ofv.gencomboFX(sForm,rs(9).name,rs(9),scomK,scomD," ") & WSPSEN
       sInput = REPLACE(sInput,"<TABLE ","<TABLE STYLE='DISPLAY:INLINE;'  ")

    ELSE


     if (session("APLMDLOR") & "") = "BE" then

       sInput = "-----"

     else


      ' Strsql = "select d.codaccion, ISNULL(e.descri,D.DESCRIACC) as actidesc,  a.* "
      ' Strsql = Strsql & " from csrelpathdescri a "

           Strsql ="SELECT   d.codaccion, ISNULL(e.descri,D.DESCRIACC) as actidesc, O1.descri, O2.descri AS Expr1, O3.descri AS Expr2, O4.descri AS Expr3, O5.descri AS Expr4, O6.descri AS Expr5, O7.descri AS Expr6,"
	   	   Strsql = strsql & "  O8.descri AS Expr7, ax.mc1, ax.mc2, ax.mc3, ax.mc4, ax.mc5, ax.mc6, ax.mc7, ax.cobj, ax.TSEC "
	   	  ' strsql="select * "
	   	   Strsql =Strsql & " FROM         dbo.CsRelPath ax LEFT OUTER JOIN"
	   	   Strsql =Strsql & "              dbo.OBJETOS O1 ON ax.mc1 = O1.codobjeto LEFT OUTER JOIN"
	   	   Strsql =Strsql & "              dbo.OBJETOS O2 ON ax.mc2 = O2.codobjeto LEFT OUTER JOIN"
	   	   Strsql =Strsql & "              dbo.OBJETOS O3 ON ax.mc3 = O3.codobjeto LEFT OUTER JOIN"
	   	   Strsql =Strsql & "              dbo.OBJETOS O4 ON ax.mc4 = O4.codobjeto LEFT OUTER JOIN"
	   	   Strsql =Strsql & "              dbo.OBJETOS O5 ON ax.mc5 = O5.codobjeto LEFT OUTER JOIN"
	   	   Strsql =Strsql & "              dbo.OBJETOS O6 ON ax.mc6 = O6.codobjeto LEFT OUTER JOIN"
	   	   Strsql =Strsql & "              dbo.OBJETOS O7 ON ax.mc7 = O7.codobjeto LEFT OUTER JOIN"
           Strsql =Strsql & "              dbo.OBJETOS O8 ON ax.cobj = O8.codobjeto "

       Strsql = Strsql & " inner join objetos b on (ax.cobj = b.codobjeto) "
       Strsql = Strsql & " inner join objlnk c on (ax.cobj = c.codobjeto) "
       Strsql = Strsql & " inner join objaccion d on (c.codobjetol = d.codobjeto and c.iden4 = d.codproced) "
       Strsql = Strsql & " left  join tipoaccalt e on (d.tipoaccion = e.tipoaccionalt and c.ltipo = 'LRGVN' ) "
       Strsql = Strsql & " WHERE Ax.COBJ = 0" & RS(7) '& " ORDER BY D.SECUEN"
       ' set rsnivdddd = ofv.crearconsultaEx(StrSql,cn3,1,parametros)
       'response.write strsql & "<br>"

       sInput = ofv.GenerarCombo(sForm,RS(9).NAME,RS(9).value,Strsql,cn,0,1,"")

     end if


    END IF





    Celda = ""
    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","ACTIVIDAD:&nbsp;","si")

    Celda = Celda & ofv.gencelda("","UTi COLSPAN=5 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")

   END IF

   END IF




'''''''''''''''''''''''''''


    Celda = ""


    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","CARACTER:&nbsp;","si")

    IF SMODALSET = "CSN" then

       sInputx = "-------"

       strsql = "Select descri From ORTIPOCARACTER where OBJID = 0" & rs(24)
       set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rsniv.eof then sInputx = rsniv(0)
       call ofv.cerrarconsulta(rsniv)

       sInput = sInputx

    ELSEif SMODALSET = "DOBLIGA" then

       IF CSTR(RS(24) & "") = "0" THEN
          WSPSEN = WSPSEN1
          sInput = "NO"
       ELSE
          WSPSEN = WSPSEN2
          sInput = "SI"
       END IF

    ''''''   sInput = ofv.gencomboFX(sForm,rs(24).name,rs(24),scomK,scomD," ") & WSPSEN
       sInput = REPLACE(sInput,"<TABLE ","<TABLE STYLE='DISPLAY:INLINE;'  ")

    ELSE

       Strsql = "Select * From ORTIPOCARACTER ORDER BY DESCRI "
       sInput = ofv.GenerarCombo(sForm,RS(24).NAME,RS(24).value,Strsql,cn,0,1,"")

    END IF
    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","CUENTA&nbsp;CONTABLE:&nbsp;","si")

    if SMODALSET = "DOBLIGA" then

       IF CSTR(RS(29) & "") = "0" THEN
          WSPSEN = WSPSEN1
       ELSE
          WSPSEN = WSPSEN2
       END IF

       sInput = ofv.gencomboFX(sForm,rs(29).name,rs(29),scomK,scomD," ") & WSPSEN
       sInput = REPLACE(sInput,"<TABLE ","<TABLE STYLE='DISPLAY:INLINE;'  ")

    ELSEIF SMODALSET = "CSN" then

        SINPUT = rs(29) & "&nbsp;"

    ELSE

       'SINPUT = ofv.GenerarInput(rs(29).name,rs(29),"text",sForm,"30","dt","")
        SINPUT="<input type=text class=dt name="&rs(29).name &" align=left value='" & rs(29) & "'  onchange=document."&sForm&".submit(); >"
    END IF

    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","REF.&nbsp;CLIENTE:&nbsp;","si")

    if SMODALSET = "DOBLIGA" then

       IF CSTR(RS(28) & "") = "0" THEN
          WSPSEN = WSPSEN1
       ELSE
          WSPSEN = WSPSEN2
       END IF

       sInput = ofv.gencomboFX(sForm,rs(28).name,rs(28),scomK,scomD," ") & WSPSEN
       sInput = REPLACE(sInput,"<TABLE ","<TABLE STYLE='DISPLAY:INLINE;'  ")

    ELSEIF SMODALSET = "CSN" then

        SINPUT = rs(28) & "&nbsp;"


    ELSE

       'SINPUT = ofv.GenerarInput(rs(28).name,rs(28),"text",sForm,"30","dt","")
       SINPUT="<input type=text class=dt name="&rs(28).name &" align=left value='" & rs(28) & "' onchange=document."&sForm&".submit();  >"
    END IF

    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")


    Celda = ""


    Celda = Celda & ofv.gencelda("","UT style='background-color:#F7AA22;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","MN.BRUTO&nbsp;","si")

    if SMODALSET = "DOBLIGA" then

       IF CSTR(RS(12) & "") = "0" THEN
          WSPSEN = WSPSEN1
       ELSE
          WSPSEN = WSPSEN2
       END IF

       sInput = ofv.gencomboFX(sForm,rs(12).name,rs(12),scomK,scomD," ") & WSPSEN
       sInput = REPLACE(sInput,"<TABLE ","<TABLE STYLE='DISPLAY:INLINE;'  ")

    ELSEIF SMODALSET = "CSN" then

        SINPUT = REPLACE(RS(12) & "",",",".") & "&nbsp;$"


    ELSE



       SINPUT = ofv.GenerarInput(rs(12).name,REPLACE(RS(12) & "",",","."),"text",sForm,"20","dtN","") & "&nbsp;$"
    END IF

    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;background-color:#f7eedd;' ","right","","","","",sInput,"si")


    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","COBERTURA SEGURO:&nbsp;%" ,"si")

    if SMODALSET = "DOBLIGA" then

       IF CSTR(RS(21) & "") = "0" THEN
          WSPSEN = WSPSEN1
       ELSE
          WSPSEN = WSPSEN2
       END IF

       sInput = ofv.gencomboFX(sForm,rs(21).name,rs(21),scomK,scomD," ") & WSPSEN
       sInput = REPLACE(sInput,"<TABLE ","<TABLE STYLE='DISPLAY:INLINE;'  ")

    ELSEIF SMODALSET = "CSN" then

        SINPUT =  rs(21) & "&nbsp;"


    ELSE

       SINPUT = ofv.GenerarInput(rs(21).name,rs(21),"text",sForm,"20","dtN","")
    END IF

    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","PRODUCTO/SERVICIO:&nbsp;","si")

    IF SMODALSET = "CSN" then

       sInputx = "-------"

       strsql = "Select descri From ORTIPOPRODUCTO where OBJID = 0" & rs(25)
       set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rsniv.eof then sInputx = rsniv(0)
       call ofv.cerrarconsulta(rsniv)

       sInput = sInputx

    ELSEif SMODALSET = "DOBLIGA" then

       IF CSTR(RS(25) & "") = "0" THEN
          WSPSEN = WSPSEN1
       ELSE
          WSPSEN = WSPSEN2
       END IF

       sInput = ofv.gencomboFX(sForm,rs(25).name,rs(25),scomK,scomD," ") & WSPSEN
       sInput = REPLACE(sInput,"<TABLE ","<TABLE STYLE='DISPLAY:INLINE;'  ")

    ELSE

       Strsql = "Select * From ORTIPOPRODUCTO ORDER BY DESCRI "
       sInput = ofv.GenerarCombo(sForm,RS(25).NAME,RS(25).value,Strsql,cn,0,1,"")

    END IF



    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")


  IF     XORWKENCA = "NO" AND SMODALSET <> "DOBLIGA"  THEN



''''''''''''''''''''''''''''''''''''''
    Celda = ""

    Celda = Celda & ofv.gencelda("","UT style='background-color:#336699;border-style:SOLID;border-width:1;border-color:#BBBBBB;color:#FFFFFF;' ","","","","","","OTROS MN.&nbsp;(+)","si")

    if SMODALSET = "DOBLIGA" then


       sInput = "-------"

    ELSE

     ''  SINPUT = ofv.GenerarInput(rs(15).name,REPLACE(RS(15) & "",",","."),"READONLY",sForm,"20","dtN","")
       SINPUT = REPLACE(RS(14) & "",",",".") & "&nbsp;$"

    END IF
    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#336699;background-color:#ddddff;' ","RIGHT","","","","",sInput,"si")




    sInput = "&nbsp;"
    Celda = Celda & ofv.gencelda("","UTi colspan=4 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")




    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")





    Celda = ""

    Celda = Celda & ofv.gencelda("","UT style='background-color:#336699;border-style:SOLID;border-width:1;border-color:#BBBBBB;color:#FFFFFF;' ","","","","","","PERDIDAS&nbsp;(+)","si")

    if SMODALSET = "DOBLIGA" then

       sInput = "-------"

    ELSE

     ''  SINPUT = ofv.GenerarInput(rs(13).name,REPLACE(RS(13) & "",",","."),"READONLY",sForm,"20","dtN","")
       SINPUT = REPLACE(RS(13) & "",",",".") & "&nbsp;$"

    END IF

    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#336699;background-color:#ddddff;' ","RIGHT","","","","",sInput,"si")


     sInput = "&nbsp;"
    Celda = Celda & ofv.gencelda("","UTi colspan=4 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")



    Celda = ""

    Celda = Celda & ofv.gencelda("","UT style='background-color:#336699;border-style:SOLID;border-width:1;border-color:#BBBBBB;color:#FFFFFF;' ","","","","","","RECUPERADO&nbsp;(-)","si")

    if SMODALSET = "DOBLIGA" then

       sInput = "-------"

    ELSE

       SINPUT = REPLACE(RS(15) & "",",",".") & "&nbsp;$"
     ''  SINPUT = ofv.GenerarInput(rs(15).name,REPLACE(RS(15) & "",",","."),"READONLY",sForm,"20","dtN","")

    END IF

    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#336699;background-color:#ddddff;' ","RIGHT","","","","",sInput,"si")


    sInput = "&nbsp;"
    Celda = Celda & ofv.gencelda("","UTi colspan=4 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")



    Celda = ""

    Celda = Celda & ofv.gencelda("","UT style='background-color:#336699;border-style:SOLID;border-width:1;border-color:#BBBBBB;color:#FFFFFF;' ","","","","","","DESAFECTACION&nbsp;(-)","si")

    if SMODALSET = "DOBLIGA" then

       sInput = "-------"

    ELSE

     '''  SINPUT = ofv.GenerarInput(rs(36).name,REPLACE(RS(36) & "",",","."),"READONLY",sForm,"20","dtN","")
       SINPUT = REPLACE(RS(36) & "",",",".") & "&nbsp;$"

    END IF

    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#336699;background-color:#ddddff;' ","RIGHT","","","","",sInput,"si")


    sInput = "&nbsp;"
    Celda = Celda & ofv.gencelda("","UTi colspan=4 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")


    Celda = ""

    Celda = Celda & ofv.gencelda("","UT style='background-color:#336699;border-style:SOLID;border-width:1;border-color:#BBBBBB;color:#FFFFFF;' ","","","","","","NO CONTABLES&nbsp;(X)","si")

    if SMODALSET = "DOBLIGA" then

       sInput = "-------"

    ELSE

     '''  SINPUT = ofv.GenerarInput(rs(36).name,REPLACE(RS(36) & "",",","."),"READONLY",sForm,"20","dtN","")
       SINPUT = REPLACE(RS(42) & "",",",".") & "&nbsp;$"

    END IF

    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#336699;background-color:#ffdddd;' ","RIGHT","","","","",sInput,"si")


    sInput = "&nbsp;"
    Celda = Celda & ofv.gencelda("","UTi colspan=4 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")




    Celda = ""

    Celda = Celda & ofv.gencelda("","UT style='background-color:#993333;border-style:SOLID;border-width:1;border-color:#BBBBBB;color:#FFFFFF;' ","","","","","","MN.NETO&nbsp;","si")

    if SMODALSET = "DOBLIGA" then

       sInput = "-------"

    ELSE

     '  wkmontoneto = cdbl(cdbl(rs(12)) + cdbl(rs(14)) - cdbl(rs(15)))
       SINPUT = REPLACE(RS(16) & "",",",".") & "&nbsp;$"
    END IF

    Celda = Celda & ofv.gencelda("","vTi style='border-style:SOLID;border-width:1;border-color:#993333;background-color:#ffcccc;color:#000000;' ","RIGHT","","","","",sInput,"si")


     sInput = "&nbsp;"
    Celda = Celda & ofv.gencelda("","UTi colspan=4 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")





''''''''''''''''''''''''''''''''''''''


  END IF




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    Celda = ""

    if session("APLMDLOR") <> "BE"  THEN
     Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","AREA/SECTOR","si")
    ELSE
     Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","SUCURSAL:<BR><BR>AREA:&nbsp;","si")

    END IF
    IF SMODALSET = "CSN" then

       sInputx = "-------"
       if session("APLMDCLI")="BC" THEN
            strsql = "Select ECOMAP From ECO where ECO = 0" & rs(20)
       ELSE
        strsql = "Select descri From ECO where ECO = 0" & rs(20)
       END IF
       set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rsniv.eof then sInputx = rsniv(0)
       call ofv.cerrarconsulta(rsniv)

       sInput = sInputx

    ELSEif SMODALSET = "DOBLIGA" then

       IF CSTR(RS(20) & "") = "0" THEN
          WSPSEN = WSPSEN1
       ELSE
          WSPSEN = WSPSEN2
       END IF

       sInput = ofv.gencomboFX(sForm,rs(20).name,rs(20),scomK,scomD," ") & WSPSEN
       sInput = REPLACE(sInput,"<TABLE ","<TABLE STYLE='DISPLAY:INLINE;'  ")

    ELSE

       if session("APLMDLOR") <> "BE" THEN
       IF  session("APLMDCLI")="BC" THEN
         STRSQL="Select * From ECO WHERE (ECOMAP  IS NOT NULL)  and ( len(ECOMAP) > 0)   ORDER BY ECOMAP "
         sInput = ofv.GenerarCombo(sForm,RS(20).NAME,RS(20).value,Strsql,cn,0,9,"")
       ELSE
       Strsql = "Select * From ECO ORDER BY DESCRI "
       sInput = ofv.GenerarCombo(sForm,RS(20).NAME,RS(20).value,Strsql,cn,0,1,"")
       END IF

       ELSE
       Strsql = "Select ECO.ECO ,CODINTER + '-' + DESCRI From ECO ORDER BY CODINTER "
       sInput = ofv.GenerarCombo(sForm,RS(20).NAME,RS(20).value,Strsql,cn,0,1,"")
       Strsql = "Select OBJID ,CODINTER + '-' + DESCRIPCION From SUCURSAL ORDER BY CODINTER "
       sInput =ofv.GenerarCombo(sForm,RS(40).NAME,RS(40).value,Strsql,cn,0,1,"") & "<BR>" & SINPUT
       END IF
    END IF


    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",  sInput,"si")

    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","	RIESGO&nbsp;ASOCIADO&nbsp;","si")


    IF SMODALSET = "CSN" then

       sInputx = "-------"

       strsql = "Select descri From ORRIESGOASOC where OBJID = 0" & rs(26)
       set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rsniv.eof then sInputx = rsniv(0)
       call ofv.cerrarconsulta(rsniv)

       sInput = sInputx

    ELSEif SMODALSET = "DOBLIGA" then

       IF CSTR(RS(26) & "") = "0" THEN
          WSPSEN = WSPSEN1
       ELSE
          WSPSEN = WSPSEN2
       END IF

       sInput = ofv.gencomboFX(sForm,rs(26).name,rs(26),scomK,scomD," ") & WSPSEN
       sInput = REPLACE(sInput,"<TABLE ","<TABLE STYLE='DISPLAY:INLINE;'  ")

    ELSE

      Strsql = "Select * From ORRIESGOASOC ORDER BY DESCRI "
      sInput = ofv.GenerarCombo(sForm,RS(26).NAME,RS(26).value,Strsql,cn,0,1,"")

    END IF


    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","AREA SOPORTE:&nbsp;","si")

    IF SMODALSET = "CSN" then

       sInputx = "-------"

       strsql = "Select descri From ECO where ECO = 0" & rs(27)
       set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rsniv.eof then sInputx = rsniv(0)
       call ofv.cerrarconsulta(rsniv)

       sInput = sInputx

    ELSEif SMODALSET = "DOBLIGA" then

       IF CSTR(RS(27) & "") = "0" THEN
          WSPSEN = WSPSEN1
       ELSE
          WSPSEN = WSPSEN2
       END IF

       sInput = ofv.gencomboFX(sForm,rs(27).name,rs(27),scomK,scomD," ") & WSPSEN
       sInput = REPLACE(sInput,"<TABLE ","<TABLE STYLE='DISPLAY:INLINE;'  ")

    ELSE

       Strsql = "Select * From ECO ORDER BY DESCRI "
       sInput = ofv.GenerarCombo(sForm,RS(27).NAME,RS(27).value,Strsql,cn,0,1,"")

    END IF

    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")
    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")









'30
    Celda = ""
    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","TIPO EVENTO:&nbsp;","si")
    sInput = WKTEVENTO

    IF SMODALSET = "CSN" then

       sInputx = "-------"

       Strsql = "Select DESCRI From  ORTIPOEVENTO   "
       Strsql = Strsql & " where OBJID = 0" & rs(30)
       set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rsniv.eof then sInputx = rsniv(0)
       call ofv.cerrarconsulta(rsniv)

       sInput = sInputx

    ELSEif SMODALSET = "DOBLIGA" then

       IF CSTR(RS(30) & "") = "0" THEN
          WSPSEN = WSPSEN1
       ELSE
          WSPSEN = WSPSEN2
       END IF

       sInput = ofv.gencomboFX(sForm,rs(30).name,rs(30),scomK,scomD," ") & WSPSEN
       sInput = REPLACE(sInput,"<TABLE ","<TABLE STYLE='DISPLAY:INLINE;'  ")

    ELSE

       Strsql = "Select * From ORTIPOEVENTO ORDER BY DESCRI"
       sInput = ofv.GenerarCombo(sForm,RS(30).NAME,RS(30).value,Strsql,cn,0,1,"")

    END IF

    Celda = Celda & ofv.gencelda("","UTi COLSPAN=3 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","COSTO OPORTUNIDAD:&nbsp;","si")

    IF SMODALSET = "CSN" then

       sInputx = "-------"

       strsql = "Select descri From ORCOSTOPOR where OBJID = 0" & rs(39)
       set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rsniv.eof then sInputx = rsniv(0)
       call ofv.cerrarconsulta(rsniv)

       sInput = sInputx

    ELSEif SMODALSET = "DOBLIGA" then

       IF CSTR(RS(39) & "") = "0" THEN
          WSPSEN = WSPSEN1
       ELSE
          WSPSEN = WSPSEN2
       END IF

       sInput = ofv.gencomboFX(sForm,rs(39).name,rs(39),scomK,scomD," ") & WSPSEN
       sInput = REPLACE(sInput,"<TABLE ","<TABLE STYLE='DISPLAY:INLINE;'  ")

    ELSE

       Strsql = "Select * From ORCOSTOPOR ORDER BY DESCRI "
       sInput = ofv.GenerarCombo(sForm,RS(39).NAME,RS(39).value,Strsql,cn,0,1,"")

    END IF

    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")


'31
    Celda = ""
    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","CATEGORIA:&nbsp;","si")
    sInput = WKTCATEGORIA

    IF SMODALSET = "CSN" then

       sInputx = "-------"

       Strsql = "Select DESCRI From  ORTIPOCATEGO   "
       Strsql = Strsql & " where OBJID = 0" & rs(31)
       set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rsniv.eof then sInputx = rsniv(0)
       call ofv.cerrarconsulta(rsniv)

       sInput = sInputx

    ELSEif SMODALSET = "DOBLIGA" then

       IF CSTR(RS(31) & "") = "0" THEN
          WSPSEN = WSPSEN1
       ELSE
          WSPSEN = WSPSEN2
       END IF

       sInput = ofv.gencomboFX(sForm,rs(31).name,rs(31),scomK,scomD," ") & WSPSEN
       sInput = REPLACE(sInput,"<TABLE ","<TABLE STYLE='DISPLAY:INLINE;'  ")

    ELSE

     Strsql = "Select DISTINCT C.OBJID, C.DESCRI From ORIMPRIESGO A "
     Strsql = Strsql & " INNER JOIN ORTIPOEVENTO B ON A.AGRUPA = B.OBJID "
     Strsql = Strsql & " INNER JOIN ORTIPOCATEGO C ON A.CATEGO = C.OBJID "
     Strsql = Strsql & " WHERE B.OBJID = 0" & RS(30) & "  "

    ''   Strsql = "Select * From ORTIPOCATEGO ORDER BY DESCRI"
       sInput = ofv.GenerarCombo(sForm,RS(31).NAME,RS(31).value,Strsql,cn,0,1,"")

    END IF


    Celda = Celda & ofv.gencelda("","UTi COLSPAN=5 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")





    Celda = ""
    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","EVENTO&nbsp;DE&nbsp;RIESGO:&nbsp;","si")

    IF SMODALSET = "CSN" then

       sInputx = "RIESGO NO IDENTIFICADO"

       Strsql = "Select DESCRI From  ORIMPRIESGO   "
       Strsql = Strsql & " where OBJID = 0" & rs(10)
       set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rsniv.eof then sInputx = rsniv(0)
       call ofv.cerrarconsulta(rsniv)

       sInput = sInputx

    ELSEif SMODALSET = "DOBLIGA" then

       IF CSTR(RS(10) & "") = "0" THEN
          WSPSEN = WSPSEN1
       ELSE
          WSPSEN = WSPSEN2
       END IF

       sInput = ofv.gencomboFX(sForm,rs(10).name,rs(10),scomK,scomD," ") & WSPSEN
       sInput = REPLACE(sInput,"<TABLE ","<TABLE STYLE='DISPLAY:INLINE;'  ")


    ELSE


  ''   Strsql = "Select A.OBJID, A.DESCRI From ORIMPRIESGO A "
  ''   Strsql = Strsql & " INNER JOIN ORTIPOEVENTO B ON A.AGRUPA = B.OBJID "
   ''  Strsql = Strsql & " INNER JOIN ORTIPOCATEGO C ON A.CATEGO = C.OBJID "
   ''  Strsql = Strsql & " WHERE A.OBJID = 0" & RS(30) & " AND C.OBJID = 0" & RS(31) & " "

    '''   Strsql = "Select * From ORIMPRIESGO ORDER BY DESCRI"


    ' Strsql = "Select DISTINCT B.OBJID, B.DESCRI From ORACTIVRIESGO A"
    ' Strsql = Strsql & " INNER JOIN ORIMPRIESGO B ON A.ITEMRIESGO = B.OBJID "
    ' Strsql = Strsql & " WHERE A.ACTIVIDAD = '" & WSPACTIVIDAD & "' ORDER BY B.DESCRI "

     if session("APLMDLOR") <> "BE" THEN
	        Strsql = "Select DISTINCT B.OBJID, B.DESCRI From ORACTIVRIESGO A"
	        Strsql = Strsql & " INNER JOIN ORIMPRIESGO B ON A.ITEMRIESGO = B.OBJID "
	        Strsql = Strsql & " WHERE A.ACTIVIDAD = '" & WSPACTIVIDAD & "' ORDER BY B.DESCRI "
	  else
	        Strsql = "Select  B.OBJID, B.DESCRI From "
	 	    Strsql = Strsql & " ORIMPRIESGO B  ORDER BY B.DESCRI "

     end if

     set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

     SCOMSEP = chr(9)
     scomK   = "0"
     scomD   = "RIESGO NO IDENTIFICADO"

     do while not rsx1.eof
         scomK = scomK & SCOMSEP & rsx1(0)
         scomD = scomD & SCOMSEP & rsx1(1)
     rsx1.movenext
     loop

    sInput = ofv.gencomboFX(sForm,RS(10).NAME,RS(10).value,scomK,scomD,"")



      '''' sInput = ofv.GenerarCombo(sForm,RS(10).NAME,RS(10).value,Strsql,cn,0,1,"")

    END IF



    Celda = Celda & ofv.gencelda("","UTi COLSPAN=5 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")









  END IF


    Celda = ""
    'OBSERVACIONES
    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","DESCRIPCION:&nbsp;","si")

    sinput2 = "&nbsp;"

    IF SMODALSET = "CSN" then

       SINPUT = rs(11) & "&nbsp;"

    ELSEif SMODALSET = "DOBLIGA" then

       IF CSTR(RS(11) & "") = "0" THEN
          WSPSEN = WSPSEN1
       ELSE
          WSPSEN = WSPSEN2
       END IF

       sInput = ofv.gencomboFX(sForm,rs(11).name,rs(11),scomK,scomD," ") & WSPSEN
       sInput = REPLACE(sInput,"<TABLE ","<TABLE STYLE='DISPLAY:INLINE;'  ")


    ELSE




       nCol = 120

       if  len(wdescri) = 0 then wdescri = "Sin Descripcion"

       if  not  isnull(rs(11)) then
           wdescri = rs(11) & ""
           sinput2 = len(wdescri) & ""
           if len(wdescri) > 300 then sinput2 = sinput2 & "<br><b>Cantidad de caracteres excede de 300.!!<b>"
       end if

       t3b = "<font class=at><textarea class=dt "
       t3b = t3b & "name=" & rs(11).name & " cols=" & nCol & " rows=3 "
       if  bModificar then
           t3b = t3b & " onblur=" & sForm & ".submit() "
           t3b = t3b & " onkeypress=" & chr(34)
           t3b = t3b & " if (this.value.length > 450) {alert('El texto es demasiado"
           t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
       else
           t3b = t3b & " readonly >"
       end if
       xdatos =  wdescri
       t3b    = t3b & xdatos & "</textarea></font>"
       SINPUT = t3b

    END IF


    Celda = Celda & ofv.gencelda("","UTi COLSPAN=4 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")
    Celda = Celda & ofv.gencelda("","UTi           style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sinput2,"si")

    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")




  IF SMODALSET    = "ENR" OR SMODALSET = "DOBLIGA" OR SMODALSET = "CSN"   OR SMODALSET = "VTR"   THEN


    wkdistrib = rs(37) & ""

    IF SMODALSET   <>  "DOBLIGA"  THEN



    Celda = ""
    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","DISTRIBUCION:&nbsp;","si")

    IF SMODALSET = "CSN" OR  XORWKENCA = "SI" then

          sInput = "-----"

          if (rs(37) & "") = "S" then
             sInput = "SI"
          else
             sInput = "NO"
          end if


    ELSE

          scomK = "N" & chr(9) & "S"
          scomD = "NO" & chr(9) & "SI"
          sInput = ofv.gencomboFX(sForm,rs(37).name,rs(37),scomK,scomD," ")

    END IF


    Celda = Celda & ofv.gencelda("","UTi COLSPAN=3 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    sInput = "-------"

    FORMADDB = ""


  IF wkdistrib = "S" and (SMODALSET  = "ENR"   OR SMODALSET = "VTR"  ) THEN


    sAccion = ""
    if bModificar then
      sAccion = "ORBASERIESGOSADD.asp?MODAL=ENR&CODTRACER=" & SCODTRACER & "&CantRegMover=" & CantRegMoverX & "&OPCION=ADDLINEA" & "&SOPX=" & sOPX
    end if
    sInput = "<INPUT TYPE='BUTTON' onclick=""disp_confirm('Confirma agregar linea???','FORMADDB');"" VALUE='AGREGAR LINEA' CLASS='BT'> "
    FORMADDB = ofv.GenForm("FORMADDB","",sAccion,"","post","","","si")

  END IF



    Celda = Celda & ofv.gencelda("","UTi COLSPAN=2 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")

   END IF









    Celda = ""
    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","LINEA NEGOCIOS:&nbsp;","si")

    IF SMODALSET = "CSN" then

       sInputx = "-------"

      ' Strsql = "Select DESCRIPCION From  ORLINNEGO1   "
      ' Strsql = Strsql & " where OBJID = 0" & rs(8)

       Strsql = ""
       Strsql = Strsql & " SELECT A.DESCRIPCION + '/' + B.DESCRIPCION AS DESCRI "
       Strsql = Strsql & "  FROM ORLINNEGO1 A "
       Strsql = Strsql & "  INNER JOIN ORLINNEGO2 B ON (A.OBJID = B.NIVEL1) "
       Strsql = Strsql & " where B.OBJID = 0" & rs(8)


       set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rsniv.eof then sInputx = rsniv(0)
       call ofv.cerrarconsulta(rsniv)

       sInput = sInputx

    ELSEif SMODALSET = "DOBLIGA" then

       IF CSTR(RS(8) & "") = "0" THEN
          WSPSEN = WSPSEN1
       ELSE
          WSPSEN = WSPSEN2
       END IF

       sInput = ofv.gencomboFX(sForm,rs(8).name,rs(8),scomK,scomD," ") & WSPSEN
       sInput = REPLACE(sInput,"<TABLE ","<TABLE STYLE='DISPLAY:INLINE;'  ")


    ELSE

       Strsql = ""
       Strsql = Strsql & " SELECT B.OBJID, A.DESCRIPCION + '/' + B.DESCRIPCION AS DESCRI "
       Strsql = Strsql & "  FROM ORLINNEGO1 A "
       Strsql = Strsql & "  INNER JOIN ORLINNEGO2 B ON (A.OBJID = B.NIVEL1) Order By A.OBJID, B.DESCRIPCION"

   '    Strsql = "Select OBJID,DESCRIPCION From ORLINNEGO1 Order By DESCRIPCION"
       sInput = ofv.GenerarCombo(sForm,RS(8).NAME,RS(8).value,Strsql,cn,0,1,"")

    END IF


    Celda = Celda & ofv.gencelda("","UTi COLSPAN=3 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

  IF wkdistrib = "S" THEN


    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","PORCENTAJE:&nbsp;" ,"si")


    IF SMODALSET = "CSN" then

       SINPUT = ofv.GenerarInput(rs(38).name,REPLACE(RS(38) & "",",","."),"READONLY",sForm,"5","dtN","") & "&nbsp;%"

     ELSE

       if cstr(RS(38)&"") ="" then
        Xporc=0
       else
        Xporc=REPLACE( RS(38) & "",",",".")
       end if

       SINPUT = ofv.GenerarInput(rs(38).name,Xporc,"text",sForm,"5","dtN","") & "&nbsp;%"

    END IF


    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

  ELSE




    sInput = "-------"


    Celda = Celda & ofv.gencelda("","UTi COLSPAN=2 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

  END IF


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")



  end if

      Formulario = ""


    IF SMODALSET = "CSN" then

      Formulario = FilaX

      FilaX = ""

    ELSE

    IF SMODALSET <> "DOBLIGA" then

       IF  XORWKENCA = "NO" THEN
           vuelta = vuelta & "_OPCION=RECALCSI"
       ELSE
           vuelta = vuelta & "_OPCION=RECALC"
       END IF
    END IF




    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=ORTRACER&volver=" & vuelta
    end if
      Formulario = ofv.GenForm(sForm,"",sAccion,"","post","",FilaX,"si") & FORMADDB

      FilaX = ""

   end if



  IF (SMODALSET    = "ENR" OR  SMODALSET = "CSN") AND wkdistrib = "S" THEN

     POSIB = 0

     Strsql = "  select * FROM ORDISTRIBLN WHERE CODTRACER = '" & SCODTRACER & "' "

'''  response.write strsql & "<br>"
     set rs3 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     if NOT rs3.eof then

        DO WHILE NOT RS3.EOF

        sFormB = "formDLN" & RS3(0)





    Celda = ""
    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","LINEA NEGOCIOS:&nbsp;","si")

    IF SMODALSET = "CSN" then

       sInputx = "-------"

    '   Strsql = "Select DESCRIPCION From  ORLINNEGO1   "
    '   Strsql = Strsql & " where OBJID = 0" & rs3(1)

       Strsql = ""
       Strsql = Strsql & " SELECT  A.DESCRIPCION + '/' + B.DESCRIPCION AS DESCRI "
       Strsql = Strsql & "  FROM ORLINNEGO1 A "
       Strsql = Strsql & "  INNER JOIN ORLINNEGO2 B ON (A.OBJID = B.NIVEL1) "
       Strsql = Strsql & " where B.OBJID = 0" & rs3(1)

       set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rsniv.eof then sInputx = rsniv(0)
       call ofv.cerrarconsulta(rsniv)

       sInput = sInputx

     ELSE

       Strsql = ""
       Strsql = Strsql & " SELECT B.OBJID, A.DESCRIPCION + '/' + B.DESCRIPCION AS DESCRI "
       Strsql = Strsql & "  FROM ORLINNEGO1 A "
       Strsql = Strsql & "  INNER JOIN ORLINNEGO2 B ON (A.OBJID = B.NIVEL1) Order By A.OBJID, B.DESCRIPCION"

     '  Strsql = "Select OBJID,DESCRIPCION From ORLINNEGO1 Order By DESCRIPCION"
       sInput = ofv.GenerarCombo(sFormB,RS3(1).NAME,RS3(1).value,Strsql,cn,0,1,"")

    END IF


    Celda = Celda & ofv.gencelda("","UTi COLSPAN=3 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")



    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","PORCENTAJE:&nbsp;","si")


    IF SMODALSET = "CSN" then

       SINPUT = ofv.GenerarInput(rs3(2).name,REPLACE(RS3(2) & "",",","."),"READONLY",sFormB,"5","dtN","") & "&nbsp;%"

     ELSE

       SINPUT = ofv.GenerarInput(rs3(2).name,REPLACE(RS3(2) & "",",","."),"text",sFormB,"5","dtN","") & "&nbsp;%"

    END IF


    Celda = Celda & ofv.gencelda("","UTi style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Filax  = ofv.genrow("","","","","25","","",Celda,"si")


    IF SMODALSET = "CSN" then

      Formulario = Formulario & FilaX

    ELSE


      sAccion = ""
      if bModificar then
         sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs3(0)," ","*")
         sAccion = sAccion & "&tabla=ORDISTRIBLN&volver=" & vuelta
      end if
      Formulario = Formulario & ofv.GenForm(sFormB,"",sAccion,"","post","",FilaX,"si")


    end if


    RS3.MOVENEXT
    LOOP


  end if

 end if





   if SMODALSET = "DOBLIGA" then

   ELSE


     Celda = ""
     Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","TIPO DE INGRESO:&nbsp;","si")

     IF RS(33) = "1" then

       sInput = "MANUAL"
       Celda  = Celda & ofv.gencelda("","UTi COLSPAN=5 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

     ELSE

       sInput = "INTERFASE"
       Celda  = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

       Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","LOTE DE INGRESO:&nbsp;","si")

       sInput = RS(32)
       Celda  = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

       Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","INFO.INTERFASE:&nbsp;","si")

       sInput = RS(34)
       Celda  = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

     END IF
     Filax  = ofv.genrow("","","","","25","","",Celda,"si")

     Celda =  ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","","FECHA ULT.ACT:&nbsp;","si")

     IF CLng("0" & RS("TRCFECACTP")) > 0 THEN
      fi = CLng(CDate("01-01-1900"))
      dULTACT = CDate(CLng("0" & RS("TRCFECACTP")) + fi)
      SINPUT=FECHADORCOR(dULTACT,"")
     ELSE
      SINPUT="---------"
    END IF
     Celda = Celda &  ofv.gencelda("","UTi COLSPAN=5 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

     Filax  = FILAX & ofv.genrow("","","","","25","","",Celda,"si")

   END IF



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

      Formulario = Formulario & FilaX


'''''   RESPONSE.WRITE " SMODAL:" & SMODAL & " SMODALSET:" & SMODALSET & "<BR>"


  IF SMODAL = "CSN" then



  ELSE

     Celda = ""


     Formu0 = "&nbsp;"
     Formu1 = "&nbsp;"
     Formu2 = "&nbsp;"

     IF wkestadoPERM < 3 THEN


        FOR XII=1 TO 10

         if ATABESTADOS(XII,0) > 0 then

            sAccion = ""
            Formu1 = "&nbsp;"
            if ATABESTADOS(XII,1) = "XXXXXXXXXXXXXXXXXENRF" then
               if bModificar then
                  sAccion = "ORBASERIESGOSFIN.asp?MODAL=" & SMODAL & "&CODTRACER=" & SCODTRACER & "&CantRegMover=" & CantRegMoverX & "&OPCION=PENDIENTE" '& "&SOPX=" & sOPX
             '''     sAccion = "ORBASERIESGOSADD.asp?MODAL=" & SMODAL & "&CODTRACER=" & SCODTRACER & "&CantRegMover=" & CantRegMoverX & "&OPCION=PENDIENTE"
               end if
               FILA0 = "<INPUT  TYPE='BUTTON' onclick=""disp_confirm('Confirma pasar este evento de perdida al estado " & ATABESTADOS(XII,2) & "???','FACC" & ATABESTADOS(XII,1) & "');"" VALUE='" & ATABESTADOS(XII,2) & "' CLASS='BTr'> "
               Formu1 = ofv.GenForm("FACC" & ATABESTADOS(XII,1),"",sAccion,"","post","",Fila0,"si")

            ELSEif ATABESTADOS(XII,1) = "XXXXXXXXXXXXXXXXXXXXXXXXXFIN" then

               if bModificar then
                  sAccion = "ORBASERIESGOSFIN.asp?MODAL=" & SMODAL & "&CODTRACER=" & SCODTRACER & "&CantRegMover=" & CantRegMoverX & "&OPCION=FIN"
               end if
               FILA0 = "<INPUT  TYPE='BUTTON' onclick=""disp_confirm('Confirma pasar este evento de perdida al estado " & ATABESTADOS(XII,2) & "???','FACC" & ATABESTADOS(XII,1) & "');"" VALUE='" & ATABESTADOS(XII,2) & "' CLASS='BTr'> "
               Formu1 = ofv.GenForm("FACC" & ATABESTADOS(XII,1),"",sAccion,"","post","",Fila0,"si")

            ELSEif ATABESTADOS(XII,1) = "XXXXXXXXXXXXXXXXXXXXXXXXXCAN" then
               if  XORWKESRAIZ = "NO" then
                   if bModificar then
                      sAccion = "ORBASERIESGOSFIN.asp?MODAL=" & SMODAL & "&CODTRACER=" & SCODTRACER & "&CantRegMover=" & CantRegMoverX & "&OPCION=CANCEL"
                    '''  sAccion = "ORBASERIESGOSADD.asp?MODAL=" & SMODAL & "&CODTRACER=" & SCODTRACER & "&CantRegMover=" & CantRegMoverX & "&OPCION=CANCEL"
                   end if
                   FILA1 = "<INPUT TYPE='BUTTON' onclick=""disp_confirm('Confirma pasar este evento de perdida al estado " & ATABESTADOS(XII,2) & "???','FACC" & ATABESTADOS(XII,1) & "');"" VALUE='" & ATABESTADOS(XII,2) & "' CLASS='BTr' > "
                   Formu1 = ofv.GenForm("FACC" & ATABESTADOS(XII,1),"",sAccion,"","post","",Fila1,"si")

               END IF

            ELSE

                if bModificar then
                   sAccion = "ORBASERIESGOSFIN.asp?MODAL=" & SMODAL & "&CODTRACER=" & SCODTRACER & "&CantRegMover=" & CantRegMoverX & "&OPCION=SET&OPCSET=" & ATABESTADOS(XII,1)
            '''       sAccion = "ORBASERIESGOSADD.asp?MODAL=" & SMODAL & "&CODTRACER=" & SCODTRACER & "&CantRegMover=" & CantRegMoverX & "&OPCION=SET&OPCSET=" & ATABESTADOS(XII,1)
                end if
                FILA1 = "<INPUT TYPE='BUTTON' onclick=""disp_confirm('Confirma pasar este evento de perdida al estado " & ATABESTADOS(XII,2) & "???','FACC" & ATABESTADOS(XII,1) & "');"" VALUE='" & ATABESTADOS(XII,2) & "' CLASS='BTr' > "
                Formu1 = ofv.GenForm("FACC" & ATABESTADOS(XII,1),"",sAccion,"","post","",Fila1,"si")


            END IF

            Celda = Celda & ofv.gencelda("","UT  align=center  ","","5%","","","",FORMU1,"si")

         end if

       NEXT

     END IF

     Filab = ofv.genrow("","","","","","","",Celda,"si")
     tablab = ofv.GenTabla("","aaa  ","","100%","","0","","0","0","","",Filab,"si")


  END IF

  END IF




  sHTML = sHTML & "<CENTER>"
  sHTML = sHTML & "<BR>" & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:3;border-color:#888888;' ","","90% ","","0","","0","0","","",Formulario,"si")
  sHTML = sHTML & "<BR>" & tablab



  sHTML = sHTML & "</CENTER>"

    sHTML = sHTML & "<script type='text/javascript'>"
    sHTML = sHTML & "function disp_confirm(disp_msg,disp_frm)"
    sHTML = sHTML & "  {"
    sHTML = sHTML & "  var r=confirm(disp_msg);"
    sHTML = sHTML & "  if (r==true)"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "  var fr= document.getElementById(disp_frm);"
  '  sHTML = sHTML & "    alert(fr.name);"
    sHTML = sHTML & "    fr.submit();"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  else"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "    alert('Se cancelo la accion!');"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  }"
    sHTML = sHTML & "function send_confirm(send_msg,send_href,send_aid)"
    sHTML = sHTML & "  {"
    sHTML = sHTML & "  var r=confirm(send_msg);"
    sHTML = sHTML & "  if (r==true)"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "    alert('ok ');"
    sHTML = sHTML & "    var fr= document.getElementById(send_aid);"
    sHTML = sHTML & "    fr.href = send_href;"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  else"
    sHTML = sHTML & "    {"
    sHTML = sHTML & "    alert('Se cancelo la accion! ');"
  '  sHTML = sHTML & "    return void(0);"
    sHTML = sHTML & "    }"
    sHTML = sHTML & "  }"
    sHTML = sHTML & "</script>"


  'Botonera
  Agregar = ""
  Buscar  = ""
  Celda   = ""

   if (SMODALORIG = "ENR" or SMODALORIG = "VTR" )  then

      sAccion = "ORBASERIESGOSLIS.asp?MODAL=" & SMODALORIG & "&CantRegMover=" & CantRegMoverX & ""
      Celda = ofv.BotonVolver(sAccion)

   END IF


   if SMODALORIG = "DOBLIGA"  AND SOPX = "" and  SOPCSET <> "DEF" then

      sAccion = "ORCARACTER.asp?CantRegMover=" & CantRegMoverX & ""
      Celda = ofv.BotonVolver(sAccion)

   END IF


  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)


  call ofv.cerrarconn(cn)

  sHTML = sHTML & GTCCALENDAR("M")

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML








%>