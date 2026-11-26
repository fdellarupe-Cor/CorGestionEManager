<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
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
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
  sopcion  = request.querystring("opcion") & ""
  sLOTE  = request.querystring("LOTE") & ""

  IF SLOTE = "" THEN
     SLOTE = SESSION("ORLOTE")
  ELSE
     SESSION("ORLOTE") = SLOTE
  END IF

  sHTML = ofv.FormHeader(session("FormN"))
  APLCHECK()

  'Encabezados
  Formulario = ""
  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)


'''''''''''''''''''''''''''''''''''''''
'' INCORPORO LOTE NUEVO

  if sopcion = "SI" then

     sfechai = clng(date()) - clng(cdate("01-01-1900"))

        strsql = "Insert into ORLOTEDET "
        strsql = strsql & " ("
        strsql = strsql & "LOTE,FECCONT,FECVALOR,CUENTA,SUCURSAL,"
        strsql = strsql & "MONTO,TEXTO,CLIENTE,FECPROC,ESTADO,"
        strsql = strsql & "INFTIPO"
        strsql = strsql & " ) Values ("
        strsql = strsql &  "'" & SLOTE & "','0','0','','',"
        strsql = strsql &  "'0','DESCRIPCION DEL EVENTO','','0','0','1'"
        strsql = strsql & ")"
 ''' response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

        strsql = "UPDATE ORLOTEING SET ESTADO = '1' WHERE CODIGO = '" & SLOTE & "' AND ESTADO = '0' "
 ''' response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)


  end if



  Formulario = ofv.Encabezados("CODIGO","CREACION","PROCESO","REVERSION","ESTADO")

  WKSETTABLA = "ORLOTEING"

  Strsql = "Select * From " & WKSETTABLA & " WHERE CODIGO = '" & SLOTE & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

    'CODIGO
    sInput = "LOTE:&nbsp;" & RS(1)
    Celda = ofv.gencelda("","CC COLSPAN=2 ","","5%","","","",sInput,"si")

    'FECCREATE
    if clng(rs(2)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSH(rs(2))
    end if
    SINPUT = dFechan
    Celda = Celda & ofv.gencelda("","UTI","","15%","","","",sInput,"si")

    'FECPROCESO
    if clng(rs(3)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSH(rs(3))
    end if
    SINPUT = dFechan
    Celda = Celda & ofv.gencelda("","UTI","","15%","","","",sInput,"si")

    'FECREVER
    if clng(rs(4)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSH(rs(4))
    end if
    SINPUT = dFechan
    Celda = Celda & ofv.gencelda("","UTI","","15%","","","",sInput,"si")



   'ESTADO

    wkestado = rs(5) & ""
    select case wkestado

           case "0"

                 wkestadoD = "<font color='#000000' >NUEVO</font>"

           case "1"

                 wkestadoD = "<font color='#993333' >REGISTRADO</font>"


           case "2"

                 wkestadoD = "<font color='#009900' >PROCESADO</font>"


           case "3"

                 wkestadoD = "<font color='#000099' >REVERTIDO</font>"

           case "4"

                 wkestadoD = "<font color='#336699' >SIN DATOS</font>"


            case "5"

                 wkestadoD = "<font color='#990000' >CANCELADO</font>"

           CASE ELSE


                 wkestadoD = "<font color='#000000' >EN REVISION</font>"


    end select

    sInput = wkestadoD
    Celda = Celda & ofv.gencelda("","UTI","","10%","","","",sInput,"si")

    Fila = ofv.genrow("","","","","","","",Celda,"si")

    Formulario = Formulario & Fila

   END IF
  sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si") & "<BR>"



  Formulario = ""

  WKSETTABLA = "ORLOTEDET"

  Strsql = "Select * From " & WKSETTABLA & " WHERE LOTE = '" & SLOTE & "' Order By OBJID DESC"
 ' response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posI = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    Filax  = ""
    wkestado = rs(10) & ""
    wkTIPO   = rs(15) & ""
    sMODAL   = "UPD"
    sDATO    = "text"


    IF wkTIPO   = "2" OR wkestado <> "0" THEN

       sMODAL   = "CSN"
       sDATO    = "readonly"

    END IF

    'Eliminar
    if bSuprimir AND sMODAL   = "UPD" then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=" & WKSETTABLA & "&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(1),sHref)
    else
      Celda = ofv.GenCelda("","CC style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","2%","","","",RS(0),"si")
    end if

 ''   Celda = ""

    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","F.CONTABLE:&nbsp;","si")


    'FECCONT
    if clng(rs(2)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSH(rs(2))
    end if
    SINPUT = dFechan

    IF sMODAL = "CSN" then
       Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    ELSE

       sInputD = ofv.GenerarInput(rs(2).name,rs(2).value,"HIDDEN",sForm,"1","dt","")
       sInput = "<u>" & sinput & "</u>" & sInputD
       Celda = Celda & ofv.gencelda("","UTi title='Modificar fecha CONTABLE.' " & GTCSETCALR(sForm,rs(2).name,"S","N","","",1,0,0,0,0) & " style='cursor:pointer;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")
    END IF






    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","F.VALOR:&nbsp;","si")

    'FECVALOR
    if clng(rs(3)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSH(rs(3))
    end if
    SINPUT = dFechan

    IF sMODAL = "CSN" then
       Celda = Celda & ofv.gencelda("","UTi  style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

    ELSE

       sInputD = ofv.GenerarInput(rs(3).name,rs(3).value,"HIDDEN",sForm,"1","dt","")
       sInput = "<u>" & sinput & "</u>" & sInputD
       Celda = Celda & ofv.gencelda("","UTi title='Modificar fecha VALOR CONTABLE.' " & GTCSETCALR(sForm,rs(3).name,"S","N","","",1,0,0,0,0) & " style='cursor:pointer;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")
    END IF




    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","CUENTA:&nbsp;","si")
    SINPUT = ofv.GenerarInput(rs(4).name,rs(4),sDATO,sForm,"40","dt","")
    Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","15%","","","",sInput,"si")

    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","F.PROCESO:&nbsp;","si")

    'FECPROC
    if clng(rs(9)) = 0 then
      dFechan = "-------"
    else
      dFechan = GTFechadorSH(rs(9))
    end if
    SINPUT = dFechan
    Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","15%","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")



    Celda = ""

    Celda = Celda & ofv.gencelda("","UT COLSPAN=2 style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","SUCURSAL:&nbsp;","si")
    SINPUT = ofv.GenerarInput(rs(5).name,rs(5),sDATO,sForm,"10","dt","")
    Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","10%","","","",sInput,"si")

    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","MONTO:&nbsp;","si")
    SINPUT = ofv.GenerarInput(rs(6).name,rs(6),sDATO,sForm,"40","dt","")
    Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","10%","","","",sInput,"si")

    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","CLIENTE:&nbsp;","si")
    SINPUT = ofv.GenerarInput(rs(8).name,rs(8),sDATO,sForm,"40","dt","")
    Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","10%","","","",sInput,"si")

   'ESTADO
    Celda = Celda & ofv.gencelda("","UT style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","ESTADO:&nbsp;","si")

    select case wkestado

           case "0"

                 wkestadoD = "<font color='#000000' >NUEVO</font>"

           case "1"

                 wkestadoD = "<font color='#993333' >REGISTRADO</font>"


           case "2"

                 wkestadoD = "<font color='#009900' >PROCESADO</font>"


           case "3"

                 wkestadoD = "<font color='#000099' >REVERTIDO</font>"

           CASE ELSE


                 wkestadoD = "<font color='#000000' >EN REVISION</font>"


    end select

    sInput = wkestadoD
    Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","10%","","","",sInput,"si")

    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")


    Celda = ""
    Celda = Celda & ofv.gencelda("","UT COLSPAN=2  style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","DESCRIPCION:&nbsp;","si")

      nCol = 120

       if  len(wdescri) = 0 then wdescri = "Sin Descripcion"

       if  not  isnull(rs(7)) then
           wdescri = rs(7)
       end if

       t3b = "<font class=at><textarea class=dt "
       t3b = t3b & "name=" & rs(7).name & " cols=" & nCol & " rows=3 "
       if  bModificar AND sMODAL   = "UPD" then
           t3b = t3b & " onblur=" & sForm & ".submit() "
           t3b = t3b & " onkeypress=" & chr(34)
           t3b = t3b & " if (this.value.length > 450) {alert('El texto es demasiado"
           t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
       else
           t3b = t3b & " readonly >"
       end if
       xdatos =  wdescri
       t3b    = t3b & xdatos & "</textarea></font>"
       SINPUT = t3b & ofv.GenerarInput(rs(10).name,"1","HIDDEN",sForm,"1","dt","")

    Celda = Celda & ofv.gencelda("","UTi COLSPAN=7 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")

   IF WKTIPO = "2" THEN

    Celda = ""
    Celda = Celda & ofv.gencelda("","UT COLSPAN=2  style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","INFO.INTERFASE:&nbsp;","si")


       SINPUT = RS(11)

    Celda = Celda & ofv.gencelda("","UTi COLSPAN=7 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")


    Celda = ""
    Celda = Celda & ofv.gencelda("","UT COLSPAN=2  style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","EVENTO:&nbsp;","si")

       sInputx = "-------" & RS(12)

       Strsql = "Select DESCRI From  ORTIPOCATEGO   "
       Strsql = Strsql & " where OBJID = 0" & RS(12)
     '  response.write strsql
       set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rsniv.eof then sInputx = rsniv(0)
       call ofv.cerrarconsulta(rsniv)

       sInput = sInputx


     ''  SINPUT = RS(12)

    Celda = Celda & ofv.gencelda("","UTi COLSPAN=7 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")

    Celda = ""
    Celda = Celda & ofv.gencelda("","UT COLSPAN=2  style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","PRODUCTO:&nbsp;","si")

       sInputx = "-------" & rs(13)

       strsql = "Select descri From TIPOPRODUCTO where OBJID = 0" & rs(13)
       set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rsniv.eof then sInputx = rsniv(0)
       call ofv.cerrarconsulta(rsniv)

       sInput = sInputx


    ''   SINPUT = RS(13)

    Celda = Celda & ofv.gencelda("","UTi COLSPAN=7 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")

    Celda = ""
    Celda = Celda & ofv.gencelda("","UT COLSPAN=2  style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","CARACTER:&nbsp;","si")

       sInputx = "-------" & RS(14)

       strsql = "Select descri From ORTIPOCARACTER where OBJID = 0" & RS(14)
       set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
       if not rsniv.eof then sInputx = rsniv(0)
       call ofv.cerrarconsulta(rsniv)

       sInput = sInputx


   ''    SINPUT = RS(14)

    Celda = Celda & ofv.gencelda("","UTi COLSPAN=7 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")

   ' PROCESO
    Celda = ""
	Celda = Celda & ofv.gencelda("","UT COLSPAN=2  style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","PROCESO:&nbsp;","si")

	   sInputx = "-------" & RS(14)

	'   Strsql = "Select DESCRI From  OBJETOS   "
	'   Strsql = Strsql & " where codobjeto = 0" & RS("PROCESO")

     if (session("APLMDLOR") & "") = "BE" then

       Strsql = "Select descripcion From TABPROCESOS01 "
       strsql = strsql & " where objid = 0" & RS("PROCESO")

     else

       Strsql = "Select DESCRI From  OBJETOS   "
       Strsql = Strsql & " where CODOBJETO = 0" & RS("PROCESO")

     end if


	   set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	   if not rsniv.eof then sInputx = rsniv(0)
	   call ofv.cerrarconsulta(rsniv)

	   sInput = sInputx


   ''    SINPUT = RS(14)

	Celda = Celda & ofv.gencelda("","UTi COLSPAN=7 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")


' SUBPROCESO
    Celda = ""
	Celda = Celda & ofv.gencelda("","UT COLSPAN=2  style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","SUBPROCESO:&nbsp;","si")

	   sInputx = "-------" & RS(14)

	'   Strsql = "Select DESCRI From  OBJETOS   "
	'   Strsql = Strsql & " where codobjeto = 0" & RS("SUBPROCESO")

     if (session("APLMDLOR") & "") = "BE" then

       Strsql = "Select descripcion From TABPROCESOS01 "
       strsql = strsql & " where objid = 0" & RS("SUBPROCESO")

     else

       Strsql = "Select DESCRI From  OBJETOS   "
       Strsql = Strsql & " where CODOBJETO = 0" & RS("SUBPROCESO")

     end if


	   set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	   if not rsniv.eof then sInputx = rsniv(0)
	   call ofv.cerrarconsulta(rsniv)

	   sInput = sInputx


   ''    SINPUT = RS(14)

	Celda = Celda & ofv.gencelda("","UTi COLSPAN=7 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")


    ' AREA /SECTOR
	    Celda = ""
		Celda = Celda & ofv.gencelda("","UT COLSPAN=2  style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","AREA/SECTOR:&nbsp;","si")

	   sInputx = "-------" & RS(14)

	  ' Strsql = "Select DESCRI From  ECO   "
	  ' Strsql = Strsql & " where ECO = 0" & RS("AREASECTOR")
	   IF session("APLMDCLI")="BC" THEN
	      Strsql = "Select ECOmap From  ECO   "
	  	  Strsql = Strsql & " where ECOMAP  LIKE '" & RS("AREASECTOR") & " -%'"

	   ELSE

	       Strsql = "Select descri From  ECO   "
	      Strsql = Strsql & " where ECO = 0" & RS("AREASECTOR")

		 END IF


	   set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	   if not rsniv.eof then sInputx = rsniv(0)
	   call ofv.cerrarconsulta(rsniv)

	   sInput = sInputx


	   ''    SINPUT = RS(14)

		Celda = Celda & ofv.gencelda("","UTi COLSPAN=7 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")



    ' AREA SOPORTE
	    Celda = ""
		Celda = Celda & ofv.gencelda("","UT COLSPAN=2  style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","AREA SOPORTE:&nbsp;","si")

	   sInputx = "-------" & RS(14)

	   Strsql = "Select DESCRI From  ECO   "
	  Strsql = Strsql & " where ECO = 0" & RS("AREASOPORTE")

	   set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	   if not rsniv.eof then sInputx = rsniv(0)
	   call ofv.cerrarconsulta(rsniv)

	   sInput = sInputx


	   ''    SINPUT = RS(14)

		Celda = Celda & ofv.gencelda("","UTi COLSPAN=7 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


    Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")

   ' LINEA NEGOCIOS
	Celda = ""
	Celda = Celda & ofv.gencelda("","UT COLSPAN=2  style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","LINEA DE NEGOCIOS:&nbsp;","si")

     sInputx = "-------" & RS(14)

	Strsql = "Select DESCRIPCION From  ORLINNEGO2   "
  Strsql = Strsql & " where OBJID = 0" & RS("LINEANEGOCIOS")

   set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   if not rsniv.eof then sInputx = rsniv(0)
   call ofv.cerrarconsulta(rsniv)

   sInput = sInputx


   ''    SINPUT = RS(14)

    IF session("APLMDCLI")="BC" THEN

      Celda = Celda & ofv.gencelda("","UTi COLSPAN=4 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")

      Celda = Celda & ofv.gencelda("","UT COLSPAN=1  style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","ASIENTO CONTABLE:&nbsp;","si")
      IF RS("TRCASIENTOCONT")&"" = "" THEN
       TRCASIENTOCONT="----------"
      ELSE
       TRCASIENTOCONT=RS("TRCASIENTOCONT")
      END IF

      Celda = Celda & ofv.gencelda("","UTi COLSPAN=3 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",TRCASIENTOCONT,"si")

    ELSE
   	   Celda = Celda & ofv.gencelda("","UTi COLSPAN=7 style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")
    END  IF

   Filax  = Filax & ofv.genrow("","","","","25","","",Celda,"si")







   END IF




    Celda = ""
    Celda = Celda & ofv.gencelda("","UT COLSPAN=9  style='background-color:#DDDDDD;border-TOP-style:SOLID;border-TOP-width:2;border-TOP-color:#000000;' ","","","","","","&nbsp;","si")
    Filax  = Filax & ofv.genrow("","","","","5","","",Celda,"si")

    Fila   = Filax

     sAccion = ""
    if bModificar AND sMODAL   = "UPD" then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=" & WKSETTABLA & "&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si")

  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "ORINTERFASESDET.asp?Opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  Buscar = ""

  sAccion = "ORINTERFASES.asp"
  Celda = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & GTCCALENDAR("M")

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>