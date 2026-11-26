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
  CantRegMoverx   = cdbl(request.querystring("CantRegMover"))
  vuelta          = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
  sMODAL          = request.querystring("MODAL")
  sMODALMAIN      = request.querystring("MODALMAIN")
  sCODTRACER      = request.querystring("CODTRACER")
  sFECHA          = request.querystring("FECHA")
  sTIPOLIS        = request.querystring("TIPOLIS")

  IF STIPOLIS <> "" THEN
     SESSION("STIPOLIS") = STIPOLIS
  ELSE
     STIPOLIS = SESSION("STIPOLIS")
  END IF


  IF STIPOLIS = "" THEN
     STIPOLIS = "BE"
     WKACCION2 = "BASE DE EVENTOS DE RIESGO "
  ELSE
     WKACCION2 = "CONSULTAS Y REPORTES "
  END IF

  if sMODAL = "" then sMODAL = "CSN"

  sHTML = ofv.FormHeader(session("FormN"))

  WKVISION = ""



  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

    sForm = "formlis"


  if sMODAL = "REG"  then

     WKACCION = " - REGISTRACION DE EVENTOS DE PERDIDA - USUARIO: " & SESSION("SEGUSUARIO")

  ELSEif sMODAL = "ENCA"  then

     WKACCION = " - ENCADENAMIENTO DE EVENTOS DE PERDIDA"

  ELSEif sMODAL = "CSN"  then

     WKACCION = ""

  ELSE

     WKACCION = " - ENRIQUECIMIENTO DE EVENTOS DE PERDIDA"


  END IF



   Celda = ""

   wkproccod = "&nbsp;"

   Celda = Celda & ofv.GenCelda("","uti  ","","90%","","","",WKACCION2 & WKACCION,"si")
   Celda = Celda & ofv.GenCelda("","ut   ","","2%","","","",wkproccod,"si")
   Fila  = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML & ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ","","100% ","30","0","","-1","1","","",Fila,"si")


  if sMODAL = "ENCA"  then

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

  sHTML = sHTML & "<BR>"

 END IF




  'Encabezados
  Formulario = ""

    Celda = ""
    Filax  = ""

    Formulario = ""


  IF  STIPOLIS = "BE" THEN




    Celda = ""

     'CARACTER
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","20%","","","","CODIGO","si")

    SINPUT = ofv.GenerarInput("ORCODIGO","TODOS","agregar",sForm,"10","dtN","")
    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","80%","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

  END IF

  IF  STIPOLIS = "BE" OR STIPOLIS = "ORPLEC" THEN

    Celda = ""

     'CARACTER
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","CARACTER","si")

    Strsql = "Select * From ORTIPOCARACTER ORDER BY DESCRI "
    sInput = ofv.GenerarCombo(sForm,"ORCARACTER","TODOS",Strsql,cn,0,1,"no")
    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

  END IF

  IF  STIPOLIS = "BE" THEN

    Celda = ""

     'CARACTER
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","DESCRIPCION ","si")
    SINPUT = ofv.GenerarInput("ORTEXTO","TODOS","agregar",sForm,"70","dt","")

    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

  END IF

  IF  STIPOLIS = "BE" OR STIPOLIS = "ORPLEC" THEN

    Celda = ""

     'CARACTER
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","TIPO EVENTO","si")

    Strsql = "Select * From ORTIPOEVENTO ORDER BY DESCRI "
    sInput = ofv.GenerarCombo(sForm,"ORRIESGO","TODOS",Strsql,cn,0,1,"no")
    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

  END IF

  IF  STIPOLIS = "BE" THEN

    Celda = ""

     'CARACTER
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","CUENTA ","si")

    SINPUT = ofv.GenerarInput("ORCUENTA","TODOS","agregar",sForm,"70","dt","")
    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")


    Celda = ""

     'CARACTER
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","CLIENTE ","si")

    SINPUT = ofv.GenerarInput("ORCLIENTE","TODOS","agregar",sForm,"70","dt","")
    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

  END IF

  IF  STIPOLIS = "BE" OR STIPOLIS = "ORRRXE"  OR STIPOLIS = "ORPO"  OR STIPOLIS = "ORPP"  OR STIPOLIS = "ORPXP"   OR STIPOLIS = "ORPCC"   OR STIPOLIS = "ORPER"   OR STIPOLIS = "ORPLEC"   OR STIPOLIS = "ORRRXELA" THEN


    Celda = ""

     'CARACTER
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","MONTO ","si")

    SINPUT = ofv.GenerarInput("ORMONTOD","TODOS","agregar",sForm,"30","dt","")
    Celda = Celda & ofv.gencelda("","VTI STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","40%","","","",SINPUT,"si")

    SINPUT = ofv.GenerarInput("ORMONTOH","TODOS","agregar",sForm,"30","dt","")
    Celda = Celda & ofv.gencelda("","VTI STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","40%","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

  END IF

  IF  STIPOLIS = "BE" OR STIPOLIS = "ORPRC"  OR STIPOLIS = "ORCAUT" THEN




     if (session("APLMDLOR") & "") = "BE" then

    Celda = ""
     'PROCESO
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","PROCESO","si")

       Strsql = "Select * From TABPROCESOS01 "
       strsql = strsql & " where "
       strsql = strsql & " nivel = '1' "


       sInput = ofv.GenerarCombo(sForm,"ORPROCESO","TODOS",Strsql,cn,0,3,"no")

       sInput = sInput & ofv.GenerarInput("ORSPAIDEN","TODOS","hidden",sForm,"1","dt","")
       sInput = sInput & ofv.GenerarInput("OREVALUADOS","TODOS","hidden",sForm,"1","dt","")
       sInput = sInput & ofv.GenerarInput("ORRIESGOSIDEN","TODOS","hidden",sForm,"1","dt","")



    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")


     else



    Celda = ""
     'PROCESO
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","PROCESO","si")

       WSPDATO = "TODOS"

     SCOMSEP = chr(9)
     scomK   = "TODOS"
     scomD   = "TODOS"



       Strsql = "Select A.* From OBJETOS A "
       Strsql = Strsql & " INNER JOIN RELAOBJ B ON ( B.CODOJASOC = A.CODOBJETO) "
    '''''   Strsql = Strsql & " where CODOJASOC = CODOBJETO "
       Strsql = Strsql & " where B.codobjmaster = '000' and B.tipobjmaster = 'PROCESO'  "
       Strsql = Strsql & " ORDER BY A.DESCRI "
      ' RESPONSE.WRITE STRSQL
    set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

    IF NOT RSX1.EOF THEN

       IF  STIPOLIS = "ORCAUT" THEN
           WSPDATO = rsx1(0)

           SCOMSEP = ""
           scomK   = ""
           scomD   = ""

       END IF


    END IF 

     do while not rsx1.eof
         scomK   = scomK & SCOMSEP & rsx1(0)
         scomD   = scomD & SCOMSEP & rsx1(3)
         SCOMSEP = chr(9)
     rsx1.movenext
     loop

    sInput = ofv.gencomboFX(sForm,"ORPROCESO",WSPDATO,scomK,scomD,"no")

    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")


   end if

  END IF 

  IF  STIPOLIS = "BE" THEN


     if (session("APLMDLOR") & "") = "BE" then




     else

    Celda = ""
     'identificacion
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","IDENTIFICACION","si")


       scomK = "1" & chr(9) & "2" & chr(9) & "3" & chr(9) & "TODOS"
       scomD = "SUBPROCESO Y ACTIVIDAD IDENTIFICADOS" & chr(9) & "SUBPROCESO IDENTIFICADO Y ACTIVIDAD NO IDENTIFICADA" & chr(9) & "SUBPROCESO Y ACTIVIDAD NO IDENTIFICADOS" & chr(9) & "TODOS"
       sInput = ofv.gencomboFX(sForm,"ORSPAIDEN","TODOS",scomK,scomD,"no")


    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")


    Celda = ""
     'EVALUADOS
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","EVALUADOS","si")


       scomK = "1" & chr(9) & "2" & chr(9) & "TODOS"
       scomD = "SI" & chr(9) & "NO" & chr(9) & "TODOS"
       sInput = ofv.gencomboFX(sForm,"OREVALUADOS","TODOS",scomK,scomD,"no")


    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

    Celda = ""
     'RIESGOS
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","RIESGOS","si")


    scomK = "1" & chr(9) & "2" & chr(9) & "TODOS"
    scomD = "IDENTIFICADOS" & chr(9) & "NO IDENTIFICADOS" & chr(9) & "TODOS"
    sInput = ofv.gencomboFX(sForm,"ORRIESGOSIDEN","TODOS",scomK,scomD,"no")
    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")


   end if 

    Celda = ""

     'LOTE
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","LOTE","si")

    Strsql = "Select * From ORLOTEING WHERE ESTADO = '2' ORDER BY FECHAPROCESO DESC "
  ''''  sInput = ofv.GenerarCombo(sForm,"ORLOTE","TODOS",Strsql,cn,1,1,"no")

    set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

     SCOMSEP = chr(9)
     scomK   = "TODOS"
     scomD   = "TODOS"

     do while not rsx1.eof
         scomK = scomK & SCOMSEP & rsx1(1)
         scomD = scomD & SCOMSEP & rsx1(1) & " - " & GTFechadorSH(rsx1(3))
     rsx1.movenext
     loop

    sInput = ofv.gencomboFX(sForm,"ORLOTE","TODOS",scomK,scomD,"no")



    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")



  END IF

  IF  STIPOLIS = "BE" OR STIPOLIS = "ORRRXE"  OR STIPOLIS = "ORPO"  OR STIPOLIS = "ORPP"  OR STIPOLIS = "ORPXP"   OR STIPOLIS = "ORPCC"   OR STIPOLIS = "ORPER"   OR STIPOLIS = "ORPLEC"   OR STIPOLIS = "ORRRXELA" THEN

    Celda = ""

     'CARACTER
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","FECHA INICIO ","si")

    SINPUT = ofv.GenerarInput("ORFECINID","TODOS","agregar",sForm,"30","dt","")
    Celda = Celda & ofv.gencelda("","VTI STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","40%","","","",SINPUT,"si")

    SINPUT = ofv.GenerarInput("ORFECINIH","TODOS","agregar",sForm,"30","dt","")
    Celda = Celda & ofv.gencelda("","VTI STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","40%","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

  END IF

  IF  STIPOLIS = "ORRRXE" OR STIPOLIS = "ORPRC" OR STIPOLIS = "ORCAUT" THEN

     Speriodo                 = session("ORPERIODO")

     if Speriodo = "" then

        Strsql = "Select * From ORPERIODOS  where APLIC = '" & wkaplic & "' and ESTADO IN ('OK','CL')  "
        Strsql = Strsql & " ORDER BY FINAL,INICIO, ESTADO DESC"

        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if NOT rs.eof then
           session("ORPERIODO")     = RS(0)
           session("ORPERIODODESC") = RS(1)
           Speriodo                 = session("ORPERIODO")
           Speriododsc              = session("ORPERIODODESC")

        end if

      END IF

    Celda = ""

     'PERIODO
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","PERIODO","si")

     Strsql = "Select *, (99999 - FINAL) AS XFFINAL From ORPERIODOS  where APLIC = '" & wkaplic & "' and "
     Strsql = Strsql & " ESTADO IN ('OK','CL')  "
     Strsql = Strsql & " ORDER BY XFFINAL,INICIO, ESTADO DESC"

    set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

     SCOMSEP = ""
     scomK   = ""
     scomD   = ""

     do while not rsx1.eof
         scomK = scomK & SCOMSEP & rsx1(0)
         scomD = scomD & SCOMSEP & rsx1(1)
         SCOMSEP = chr(9)
     rsx1.movenext
     loop

    sInput = ofv.gencomboFX(sForm,"ORPERIODO",Speriodo,scomK,scomD,"no")



    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")



  END IF


  IF  STIPOLIS = "ORPXP" THEN

    Celda = ""

     'PERIODO
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","PRODUCTO","si")

     Strsql = "SELECT * FROM ORTIPOPRODUCTO  ORDER BY DESCRI"
    set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

     SCOMSEP = chr(9)
     scomK   = "TODOS"
     scomD   = "TODOS"

     do while not rsx1.eof
         scomK = scomK & SCOMSEP & rsx1(0)
         scomD = scomD & SCOMSEP & rsx1(1)
         SCOMSEP = chr(9)
     rsx1.movenext
     loop

    sInput = ofv.gencomboFX(sForm,"ORPRODUCTO","TODOS",scomK,scomD,"no")



    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")



  END IF

  IF  STIPOLIS = "ORPXP" THEN

    Celda = ""

     'PERIODO
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","TIPO EVENTO","si")

     Strsql = "SELECT * FROM ORTIPOEVENTO ORDER BY SECUEN, DESCRI"
    set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

     SCOMSEP = chr(9)
     scomK   = "TODOS"
     scomD   = "TODOS"

     do while not rsx1.eof
         scomK = scomK & SCOMSEP & rsx1(0)
         scomD = scomD & SCOMSEP & rsx1(1)
         SCOMSEP = chr(9)
     rsx1.movenext
     loop

    sInput = ofv.gencomboFX(sForm,"ORTEVENTO","TODOS",scomK,scomD,"no")



    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")



  END IF



    Celda = ""
   SINPUT = "<INPUT TYPE='SUBMIT' VALUE='SELECCIONAR' CLASS='BT'>&nbsp;&nbsp;"
    Celda = Celda & ofv.gencelda("","ut colspan=3 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","right","","","","",SINPUT,"si")

       Filax  = Filax & ofv.genrow("","AAA bgcolor='#DDDDFF' ","","","30","","",Celda,"si")


  IF  STIPOLIS = "BE" THEN

         sAccion = "orMAPAEVENTOSlis.asp?modal=CSN&CODTRACER=" & SCODTRACER & "&OPCION=VS"

  ELSE

     IF STIPOLIS = "ORPO" THEN         ' PERDIDAS OPERACIONALES
         sAccion = "ORMAPAEVENTOSPOlis.asp?modal=CSN&OPCION=VS"
     ELSEIF STIPOLIS = "ORPP" THEN         ' PRINCIPALES PERDIDAS
         sAccion = "ORMAPAEVENTOSPPlis.asp?modal=CSN&OPCION=VS"
     ELSEIF STIPOLIS = "ORPPRNK" THEN         ' PRINCIPALES PERDIDAS
         sAccion = "ORMAPAEVENTOSPPRNKlis.asp?modal=CSN&OPCION=VS"
     ELSEIF STIPOLIS = "ORPRC" THEN         ' ESTADO DE LAS EVALUACIONES
         sAccion = "ORMAPAEVENTOSPRClis.asp?modal=CSN&OPCION=VS"
     ELSEIF STIPOLIS = "ORPCC" THEN         ' PERDIDAS X RUBRO CONTABLE
         sAccion = "ORMAPAEVENTOSPCClis.asp?modal=CSN&OPCION=VS"
     ELSEIF STIPOLIS = "ORPXP" THEN         ' PERDIDAS X PRODUCTO
         sAccion = "ORMAPAEVENTOSPXPlis.asp?modal=CSN&OPCION=VS"
     ELSEIF STIPOLIS = "ORPER" THEN         ' LISTA DE EVENTOS RAIZ
         sAccion = "ORMAPAEVENTOSPERlis.asp?modal=CSN&OPCION=VS"
     ELSEIF STIPOLIS = "ORPLEC" THEN         ' LISTA DE EVENTOS POR CARACTER
         sAccion = "ORMAPAEVENTOSPLEClis.asp?modal=CSN&OPCION=VS"
     ELSEIF STIPOLIS = "ORRRXE" THEN         ' RIESGO RESIDUAL X TIPO DE EVENTO
         sAccion = "ORMAPAEVENTOSRRXElis.asp?modal=CSN&OPCION=VS"
     ELSEIF STIPOLIS = "ORCAUT" THEN         ' CALIDAD DE LAS AUTOEVALUACIONES
         sAccion = "ORMAPAEVENTOSCAUTlis.asp?modal=CSN&OPCION=VS"
     END IF

  END IF

         Formulario = ofv.GenForm(sForm,"",sAccion,""," post target=myNewWin onsubmit=window.open('','myNewWin',',scrollbars=1,resizable=1;status=0,toolbar=0,width=900,height=600'); ","",FilaX,"si")



  sHTML = sHTML & "<br><br><center>" & ofv.GenTabla("","aaa STYLE='border-style:SOLID;border-width:1;border-color:#888888;' ","","70% ","","0","","0","0","","",Formulario,"si") & "</center>"

  'Botonera
  Agregar = ""
  Buscar  = ""
  Celda   = ""

  if sMODAL = "ENCA"  then
     sAccion = "ORBASERIESGOSENCA.asp?CODTRACER=" & SCODTRACER & "&MODAL=ENR"
     Celda = ofv.BotonVolver(sAccion)
  END IF

  sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)





  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML



%>