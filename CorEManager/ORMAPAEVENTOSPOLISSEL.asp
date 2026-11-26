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
  SOPCION         = request.querystring("OPCION")

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


  IF SOPCION = "SETFRM" THEN


    SORCODIGO       = UCASE(request.FORM("ORCODIGO"))
    sORCARACTER     = UCASE(request.FORM("ORCARACTER"))
    SORTEXTO        = UCASE(request.FORM("ORTEXTO"))
    sORRIESGO       = UCASE(request.FORM("ORRIESGO"))
    SORCUENTA       = UCASE(request.FORM("ORCUENTA"))
    SORCLIENTE      = UCASE(request.FORM("ORCLIENTE"))
    SORMONTOD       = UCASE(request.FORM("ORMONTOD"))
    SORMONTOH       = UCASE(request.FORM("ORMONTOH"))
    SORPROCESO      = UCASE(request.FORM("ORPROCESO"))
    SORLOTE         = UCASE(request.FORM("ORLOTE"))
    SOREVALUADOS    = UCASE(request.FORM("OREVALUADOS"))
    SORRIESGOSIDEN  = UCASE(request.FORM("ORRIESGOSIDEN"))
    SORSPAIDEN      = UCASE(request.FORM("ORSPAIDEN"))
    SORFECINID      = UCASE(request.FORM("ORFECINID"))
    SORFECINIH      = UCASE(request.FORM("ORFECINIH"))
    SORPERIODO      = UCASE(request.FORM("ORPERIODO"))
    SORRANGOY       = UCASE(request.FORM("ORRANGOY"))
    SORMILES        = UCASE(request.FORM("ORMILES"))
    SORRANKING      = UCASE(request.FORM("ORRANKING"))
    SORPRODUCTO     = UCASE(request.FORM("ORPRODUCTO"))
    SORTEVENTO      = UCASE(request.FORM("ORTEVENTO"))
    SORXLS          = UCASE(request.FORM("ORXLS"))
	
	
	
	SORFECREGCONTIN	= UCASE(request.FORM("ORFECREGCONTIN"))
	SORFECREGCONTOUT= UCASE(request.FORM("ORFECREGCONTOUT"))
	SORFECCARGAIN	= UCASE(request.FORM("ORFECCARGAIN"))
	SORFECCARGAOUT	= UCASE(request.FORM("ORFECCARGAOUT"))
	SORFECDESCUIN	= UCASE(request.FORM("ORFECDESCUIN"))
	SORFECDESCUOUT	= UCASE(request.FORM("ORFECDESCUOUT"))
	SORFECFININ		= UCASE(request.FORM("ORFECFININ"))
	SORFECFINOUT 	= UCASE(request.FORM("ORFECFINOUT"))
	SORFECESTFININ	= UCASE(request.FORM("ORFECESTFININ"))
	SORFECESTFINOUT	= UCASE(request.FORM("ORFECESTFINOUT"))
	SORESTADOS		= UCASE(request.FORM("ORESTADOS"))
	SFECHASELECCION	= UCASE(request.FORM("FECHASELECCION"))

  else

    SORCODIGO       = "TODOS"
    sORCARACTER     = "TODOS"
    SORTEXTO        = "TODOS"
    sORRIESGO       = "TODOS"
    SORCUENTA       = "TODOS"
    SORCLIENTE      = "TODOS"
    SORMONTOD       = "TODOS"
    SORMONTOH       = "TODOS"
    SORPROCESO      = "TODOS"
    SORLOTE         = "TODOS"
    SOREVALUADOS    = "TODOS"
    SORRIESGOSIDEN  = "TODOS"
    SORSPAIDEN      = "TODOS"
    SORFECINID      = "TODOS"
    SORFECINIH      = "TODOS"
    SORPERIODO      = "TODOS"
    SORRANGOY       = "TODOS"
    SORMILES        = "TODOS"
    SORRANKING      = "TODOS"
    SORPRODUCTO     = "TODOS"
    SORTEVENTO      = "TODOS"
    SORXLS          = "NO"
	
	SORFECREGCONTIN	= "TODOS"
	SORFECREGCONTOUT= "TODOS"
	SORFECCARGAIN	= "TODOS"
	SORFECCARGAOUT	= "TODOS"
	SORFECDESCUIN	= "TODOS"
	SORFECDESCUOUT	= "TODOS"
	SORFECFININ		= "TODOS"
	SORFECFINOUT 	= "TODOS"
	SORFECESTFININ	= "TODOS"
	SORFECESTFINOUT	= "TODOS"
	SORESTADOS		= "TODOS"
	SFECHASELECCION	= 0

  end if





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


  IF  STIPOLIS = "BE" OR STIPOLIS = "ORPER" THEN




    Celda = ""

     'CARACTER
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","20%","","","","CODIGO","si")

    SINPUT = ofv.GenerarInput("ORCODIGO",SORCODIGO,"TEXT",sForm,"10","dtN","")
    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","80%","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

  END IF

  IF  STIPOLIS = "BE" OR STIPOLIS = "ORPLEC" THEN

    Celda = ""

     'CARACTER
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","CARACTER","si")

    Strsql = "Select * From ORTIPOCARACTER ORDER BY DESCRI "
 ''   sInput = ofv.GenerarCombo(sForm,"ORCARACTER","TODOS",Strsql,cn,0,1,"no")

     set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     SCOMSEP = chr(9)
     scomK   = "TODOS"
     scomD   = "TODOS"

     do while not rsx1.eof
         scomK = scomK & SCOMSEP & rsx1(0)
         scomD = scomD & SCOMSEP & rsx1(1)
     rsx1.movenext
     loop

    sInput = ofv.gencomboFX(sForm,"ORCARACTER",SORCARACTER,scomK,scomD,"")


    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

  END IF

  IF  STIPOLIS = "BE" THEN

    Celda = ""

     'CARACTER
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","DESCRIPCION ","si")
    SINPUT = ofv.GenerarInput("ORTEXTO",SORTEXTO,"TEXT",sForm,"70","dt","")
    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

  END IF

  IF  STIPOLIS = "BE" OR STIPOLIS = "ORPLEC" THEN

    Celda = ""

     'CARACTER
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","TIPO EVENTO","si")

    Strsql = "Select * From ORTIPOEVENTO ORDER BY DESCRI "
  '''  sInput = ofv.GenerarCombo(sForm,"ORRIESGO","TODOS",Strsql,cn,0,1,"no")

     set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     SCOMSEP = chr(9)
     scomK   = "TODOS"
     scomD   = "TODOS"

     do while not rsx1.eof
         scomK = scomK & SCOMSEP & rsx1(0)
         scomD = scomD & SCOMSEP & rsx1(1)
     rsx1.movenext
     loop

    sInput = ofv.gencomboFX(sForm,"ORRIESGO",SORRIESGO,scomK,scomD,"")

    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

  END IF

  IF  STIPOLIS = "BE" THEN

    Celda = ""

     'CARACTER
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","CUENTA ","si")


    SINPUT = ofv.GenerarInput("ORCUENTA",SORCUENTA,"TEXT",sForm,"70","dt","")
    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")


    Celda = ""

     'CARACTER
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","CLIENTE ","si")
    SINPUT = ofv.GenerarInput("ORCLIENTE",SORCLIENTE,"TEXT",sForm,"70","dt","")

    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

  END IF

  IF  STIPOLIS = "BE" OR STIPOLIS = "ORRRXE"  OR STIPOLIS = "ORPO"  OR STIPOLIS = "ORPP"  OR STIPOLIS = "ORPPRNK"  OR STIPOLIS = "ORPXP"   OR STIPOLIS = "ORPCC"   OR STIPOLIS = "ORPER"   OR STIPOLIS = "ORPLEC"   OR STIPOLIS = "ORRRXELA" THEN


    Celda = ""

     'CARACTER
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","MONTO ","si")

    SINPUT = ofv.GenerarInput("ORMONTOD",SORMONTOD,"TEXT",sForm,"30","dt","")
    Celda = Celda & ofv.gencelda("","VTI STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","40%","","","",SINPUT,"si")

    SINPUT = ofv.GenerarInput("ORMONTOH",SORMONTOH,"TEXT",sForm,"30","dt","")
    Celda = Celda & ofv.gencelda("","VTI STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","40%","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

  END IF

  IF  STIPOLIS = "ORPO" THEN

    Celda = ""

     'RANGO
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","RANGO ","si")

       scomK = "1000" & chr(9) & "2000" & chr(9) & "5000" & chr(9)  & "10000" & chr(9) & "20000" & chr(9) & "50000" & chr(9)  & "100000" & chr(9) & "TODOS"
       scomD = "1000" & chr(9) & "2000" & chr(9) & "5000" & chr(9)  & "10000" & chr(9) & "20000" & chr(9) & "50000" & chr(9)  & "100000" & chr(9) & "AUTOMATICO"
       sInput = ofv.gencomboFX(sForm,"ORRANGOY",SORRANGOY,scomK,scomD,"")

    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")



  END IF

  IF  STIPOLIS = "ORPO" OR STIPOLIS = "ORPP"  OR STIPOLIS = "ORPPRNK"   OR STIPOLIS = "ORRRXE" THEN

    Celda = ""

     'RANGO
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","PRECISION ","si")

       scomK = "1000" & chr(9) & "TODOS"
       scomD = "MILES" & chr(9) & "AUTOMATICO"
       sInput = ofv.gencomboFX(sForm,"ORMILES",SORMILES,scomK,scomD,"")

    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")



  END IF

  IF  STIPOLIS = "ORPPRNK"  THEN

    Celda = ""

     'cantidad
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","RANKING ","si")

       scomK = "10" & chr(9) & "15" & chr(9) & "TODOS"
       scomD = "10 PRIMEROS" & chr(9) & "15 PRIMEROS" & chr(9) & "5 PRIMEROS"
       sInput = ofv.gencomboFX(sForm,"ORRANKING",SORRANKING,scomK,scomD,"")

    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")



  END IF


  IF  STIPOLIS = "BE" OR STIPOLIS = "ORRRXE" OR STIPOLIS = "ORPRC" OR STIPOLIS = "ORCAUT" THEN



     IF SORPERIODO = "TODOS" THEN

        Strsql = "Select OBJID,PERIODO,RGAMBAUTOEV, (99999 - FINAL) AS XFFINAL From ORPERIODOS  where APLIC = '" & wkaplic & "' and ESTADO IN ('OK','CL')  "
        Strsql = Strsql & " ORDER BY XFFINAL DESC,INICIO, ESTADO DESC"

        set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
        if NOT rs.eof then
           session("ORPERIODO")     = RS(0)
           session("ORPERIODODESC") = RS(1)
           SORPERIODO               = session("ORPERIODO")
           Speriododsc              = session("ORPERIODODESC")
           WKORPERIODOMAP           = RS(2)
        ELSE
           WKORPERIODOMAP  = 0
        END IF


    ELSE

         Strsql2 = "Select OBJID,PERIODO,RGAMBAUTOEV From ORPERIODOS  where APLIC = '" & wkaplic & "' and OBJID = 0" & SORPERIODO
         set rsX1 = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
         if NOT rsX1.eof then
            WKORPERIODOMAP  = RSX1(2)
         ELSE
            WKORPERIODOMAP  = 0
         END IF


    END IF

    Celda = ""

     'PERIODO
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","PERIODO","si")

     Strsql = "Select *, (99999 - FINAL) AS XFFINAL From ORPERIODOS  where APLIC = '" & wkaplic & "' and "
     Strsql = Strsql & " ESTADO IN ('OK','CL')  "
     Strsql = Strsql & " ORDER BY XFFINAL DESC,INICIO, ESTADO DESC"

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

    sInput = ofv.gencomboFX(sForm,"ORPERIODO",SORPERIODO,scomK,scomD,"")



    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")



  END IF




  IF  STIPOLIS = "BE" OR STIPOLIS = "ORPRC"  OR STIPOLIS = "ORCAUT"  OR STIPOLIS = "ORRRXE" THEN




''''''''''''''''''''''
     if (session("APLMDLOR") & "") = "BE" then

    Celda = ""
     'PROCESO
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","PROCESO","si")

       Strsql = "Select * From TABPROCESOS01 "
       strsql = strsql & " where "
       strsql = strsql & " nivel = '1' "


       sInput = ofv.GenerarCombo(sForm,"ORPROCESO",SORPROCESO,Strsql,cn,0,3,"")

       sInput = sInput & ofv.GenerarInput("ORSPAIDEN",SORSPAIDEN,"hidden",sForm,"1","dt","")
       sInput = sInput & ofv.GenerarInput("OREVALUADOS",SOREVALUADOS,"hidden",sForm,"1","dt","")
       sInput = sInput & ofv.GenerarInput("ORRIESGOSIDEN",SORRIESGOSIDEN,"hidden",sForm,"1","dt","")



    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")


     else



    Celda = ""
     'PROCESO
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","PROCESO","si")


     SCOMSEP = chr(9)
     scomK   = "TODOS"
     scomD   = "TODOS"



       Strsql = "Select A.* From OBJETOS A "
       Strsql = Strsql & " INNER JOIN RELAOBJ B ON ( B.CODOJASOC = A.CODOBJETO) "
    '''''   Strsql = Strsql & " where CODOJASOC = CODOBJETO "
       Strsql = Strsql & " where B.codobjmaster = '0" & WKORPERIODOMAP & "' and B.tipobjmaster = 'PROCESO'  "
       Strsql = Strsql & " ORDER BY A.DESCRI "
    '''   RESPONSE.WRITE STRSQL
    set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

    IF NOT RSX1.EOF THEN

       IF  STIPOLIS = "ORCAUT" THEN
           IF SORPROCESO = "TODOS" THEN SORPROCESO = rsx1(0)

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

    sInput = ofv.gencomboFX(sForm,"ORPROCESO",SORPROCESO,scomK,scomD,"")

    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")


   end if

'''''''''''''''''''''''''''





  END IF

  IF  STIPOLIS = "BE" THEN



     if (session("APLMDLOR") & "") = "BE" then




     else



    Celda = ""
     'identificacion
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","IDENTIFICACION","si")


       scomK = "1" & chr(9) & "2" & chr(9) & "3" & chr(9) & "TODOS"
       scomD = "SUBPROCESO Y ACTIVIDAD IDENTIFICADOS" & chr(9) & "SUBPROCESO IDENTIFICADO Y ACTIVIDAD NO IDENTIFICADA" & chr(9) & "SUBPROCESO Y ACTIVIDAD NO IDENTIFICADOS" & chr(9) & "TODOS"
       sInput = ofv.gencomboFX(sForm,"ORSPAIDEN",SORSPAIDEN,scomK,scomD,"")


    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")


    Celda = ""
     'EVALUADOS
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","EVALUADOS","si")


       scomK = "1" & chr(9) & "2" & chr(9) & "TODOS"
       scomD = "SI" & chr(9) & "NO" & chr(9) & "TODOS"
       sInput = ofv.gencomboFX(sForm,"OREVALUADOS",SOREVALUADOS,scomK,scomD,"")


    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

    Celda = ""
     'RIESGOS
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","RIESGOS","si")


    scomK = "1" & chr(9) & "2" & chr(9) & "TODOS"
    scomD = "IDENTIFICADOS" & chr(9) & "NO IDENTIFICADOS" & chr(9) & "TODOS"
    sInput = ofv.gencomboFX(sForm,"ORRIESGOSIDEN",SORRIESGOSIDEN,scomK,scomD,"")
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

    sInput = ofv.gencomboFX(sForm,"ORLOTE",SORLOTE,scomK,scomD,"")



    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

  END IF


'***************************************************************
'* Creación de las fechas para poder agregar la selección de fechas por un combo
'***************************************************************

  IF  STIPOLIS = "BE" OR STIPOLIS = "ORRRXE"  OR STIPOLIS = "ORPO"  OR STIPOLIS = "ORPP"  OR STIPOLIS = "ORPPRNK"  OR STIPOLIS = "ORPXP"   OR STIPOLIS = "ORPCC"   OR STIPOLIS = "ORPER"   OR STIPOLIS = "ORPLEC"   OR STIPOLIS = "ORRRXELA" THEN


	Celda  = ofv.GenCelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","FEHCA DE BUSQUEDA","si")

    SCOMSEP = chr(9)
	scomK = ""
	scomD = ""
	scomK = scomK & SCOMSEP & "1"
    scomD = scomD & SCOMSEP & "FECHA INICIO"
	scomK = scomK & SCOMSEP & "2"
    scomD = scomD & SCOMSEP & "FECHA DE REGISTRACION CONTABLE"
	scomK = scomK & SCOMSEP & "3"
    scomD = scomD & SCOMSEP & "FECHA DE CARGA"
	scomK = scomK & SCOMSEP & "4"
    scomD = scomD & SCOMSEP & "FECHA DE DESCUBRIMIENTO"
	scomK = scomK & SCOMSEP & "5"
    scomD = scomD & SCOMSEP & "FECHA DE FINALIZACION"
	scomK = scomK & SCOMSEP & "6"
    scomD = scomD & SCOMSEP & "FECHA DE ESTADO FINALIZADO"

	sInput = ofv.gencomboFX(SFORM,"FECHASELECCION",SFECHASELECCION,scomK,scomD,"")
	
	Celda  = Celda & ofv.GenCelda("","VTI STYLE='border-style:SOLID;border-width:1;border-color:#888888;' COLSPAN='2'","","","","","",sInput,"si")
	Filax   = Filax & ofv.GenRow("","","","","30","","",Celda,"si")

      'sAccion = "ORRUNLOTELISTADOS.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
      'Formulario = Formulario & ofv.GenForm("LISTADO","",sAccion,"","post","",Fila,"si")

    Celda = ""

     'CARACTER
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","RANGO DE FECHA ","si")

    SINPUT = ofv.GenerarInput("ORFECINID",SORFECINID,"TEXT",sForm,"30","dt","")
    Celda = Celda & ofv.gencelda("","VTI STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","40%","","","",SINPUT,"si")

    SINPUT = ofv.GenerarInput("ORFECINIH",SORFECINIH,"TEXT",sForm,"30","dt","")
    Celda = Celda & ofv.gencelda("","VTI STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","40%","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

  END IF
  
  '*****************************************************************************
  '* Agregado para la funcionalidad de filtros 2010/05/03
  '*****************************************************************************

	if STIPOLIS = "ORPO" OR STIPOLIS = "ORPP" OR STIPOLIS = "ORPER" OR STIPOLIS = "ORPCC" OR STIPOLIS = "ORPXP" OR STIPOLIS = "ORPLEC" OR STIPOLIS = "ORPPRNK" then
	
	
		Celda = ""
	     'ESTADOS
	 	Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","ESTADOS","si")

        Strsql = "select distinct a.codigo, a.descripcion "
        Strsql = Strsql & " from       orestados a "
        Strsql = Strsql & " inner join orestadosper b on (a.objid = b.estado) "
        Strsql = Strsql & " where a.tipo > 0 and b.permiso > 0 "
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

        Strsql = Strsql & " ) "
		
		Strsql = Strsql & " UNION ALL SELECT CODIGO = 'NO', DESCRIPCION = 'RESGUARDO RECTIFICATIVA' "

        set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

        SCOMSEP = chr(9)
        scomK   = "TODOS"
        scomD   = "TODOS"

        do while not rsx1.eof
        	scomK = scomK & SCOMSEP & rsx1(0)
	        scomD = scomD & SCOMSEP & rsx1(1)
    	    rsx1.movenext
        loop

	    sInput = ofv.gencomboFX(sForm,"ORESTADOS",SORESTADOS,scomK,scomD,"")
	    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")
	    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
			
	end if

  '*****************************************************************************
  '* Fin de agregado de filtros 2010/05/03
  '*****************************************************************************




    Celda = ""
   SINPUT = "<INPUT TYPE='SUBMIT' VALUE='SELECCIONAR' CLASS='BT'>&nbsp;&nbsp;"
    Celda = Celda & ofv.gencelda("","ut colspan=3 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","right","","","","",SINPUT,"si")

    '   Filax  = Filax & ofv.genrow("","AAA bgcolor='#DDDDFF' ","","","30","","",Celda,"si")


  IF  STIPOLIS = "BE" THEN

         sAccion = "orMAPAEVENTOSlis.asp?modal=CSN&CODTRACER=" & SCODTRACER & "&OPCION=VS"

  ELSE

     IF STIPOLIS = "ORPO" THEN         ' PERDIDAS OPERACIONALES
         sAccion = "ORMAPAEVENTOSPOlis.asp?modal=CSN&OPCION=VS"
     ELSEIF STIPOLIS = "ORPP" THEN         ' PRINCIPALES PERDIDAS
         sAccion = "ORMAPAEVENTOSPPlis.asp?modal=CSN&OPCION=VS"
     ELSEIF STIPOLIS = "ORPPRNK" THEN         ' PRINCIPALES PERDIDAS O RANKING DE PERDIDAS
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

  
            sAccion2 = "ORMAPAEVENTOSPOlisSEL.ASP?MODAL=" & SMODAL & "&CantRegMover=" & CantRegMoverx & "&MODALMAIN=" & sMODALMAIN & "&TIPOLIS=" & sTIPOLIS
            sAccion2 = sAccion2 & "&CODTRACER=" & sCODTRACER & "&FECHA=" & sFECHA & "&OPCION=SETFRM"


            Celda = ""

            Celda = Celda & ofv.gencelda("","ut  STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","CENTER","","","","","&nbsp;","si")


            SINPUT = "<INPUT TYPE='button' VALUE='SELECCIONAR'"  &"CLASS='BT' "
            SINPUT = SINPUT & " onclick=""document." & sform & ".action= '" & sAccion & "';"
            SINPUT = SINPUT & " document." & sform & ".target= 'myNewWin';"
            SINPUT = SINPUT & " window.open('','myNewWin',',scrollbars=1,resizable=1;status=0,toolbar=0,width=900,height=600'); "
            SINPUT = SINPUT & " document." & sform & ".submit(); "
            SINPUT = SINPUT & " document." & sform & ".action= '" & sAccion2 & "';"
            SINPUT = SINPUT & " document." & sform & ".target= '_self';"
            SINPUT = SINPUT & " "" "
  ''                  SINPUT = SINPUT & " alert(document." & sform & ".action); "
   ''                 SINPUT = SINPUT & " alert(document." & sform & ".target);"" "
            SINPUT = SINPUT & " >&nbsp;&nbsp;"

            Celda = Celda & ofv.gencelda("","ut colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","right","","","","",SINPUT,"si")

            Filax  = Filax & ofv.genrow("","AAA bgcolor='#DDDDFF' ","","","30","","",Celda,"si")

            sAccion = sAccion2

            Formulario = ofv.GenForm(sForm,"",sAccion,"","post","",FilaX,"si")

  

 ''' Formulario = ofv.GenForm(sForm,"",sAccion,""," post target=myNewWin onsubmit=window.open('','myNewWin',',scrollbars=1,resizable=1;status=0,toolbar=0,width=900,height=600'); ","",FilaX,"si")



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