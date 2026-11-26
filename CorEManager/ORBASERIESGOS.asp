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
  SOPCION         = request.querystring("OPCION")
  sMODAL2          = request.querystring("MODAL2")
   'response.write request.querystring
  APLCHECK()


  if sMODAL = "" then sMODAL = "REG"

  sHTML = ofv.FormHeader(session("FormN"))

  WKVISION = ""



  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)
  sFECREGIS = CLNG(DATE()) - clng(cdate("01-01-1900")) - 1
  Strsql = "DELETE FROM ORTRACER WHERE TRCESTADO = 'NO' and TRCMARCAACTP <> 'RM' and fecregis < " & sFECREGIS

  ' response.write strsql & "<br>"
  set rs2 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

    sForm = "formlis"


  if sMODAL = "REG"  then

     WKACCION = " - REGISTRACION DE EVENTOS DE PERDIDA - USUARIO: " & SESSION("SEGUSUARIO")

  ELSEif sMODAL = "ENCA"  then

     WKACCION = " - ENCADENAMIENTO DE EVENTOS DE PERDIDA"

  ELSEif sMODAL = "CSN"  then

     WKACCION = " - CONSULTA DE EVENTOS DE PERDIDA"

  ELSE

     WKACCION = " - ENRIQUECIMIENTO DE EVENTOS DE PERDIDA"


  END IF


  IF SOPCION = "SETFRM" THEN


    SORCODIGO       	= UCASE(request.FORM("ORCODIGO"))
    sORCARACTER     	= UCASE(request.FORM("ORCARACTER"))
    SORTEXTO        	= UCASE(request.FORM("ORTEXTO"))
    sORRIESGO       	= UCASE(request.FORM("ORRIESGO"))
    SORCUENTA       	= UCASE(request.FORM("ORCUENTA"))
    SORCLIENTE      	= UCASE(request.FORM("ORCLIENTE"))
    SORMONTOD       	= UCASE(request.FORM("ORMONTOD"))
    SORMONTOH       	= UCASE(request.FORM("ORMONTOH"))
    SORPROCESO      	= UCASE(request.FORM("ORPROCESO"))
    SORLOTE         	= UCASE(request.FORM("ORLOTE"))
    SOREVALUADOS    	= UCASE(request.FORM("OREVALUADOS"))
    SORRIESGOSIDEN  	= UCASE(request.FORM("ORRIESGOSIDEN"))
    SORSPAIDEN      	= UCASE(request.FORM("ORSPAIDEN"))
    SORESTADOS	    	= UCASE(request.FORM("ORESTADOS"))
    SORLINNEG       	= UCASE(request.FORM("ORLINNEG"))
    SORTIPORG       	= UCASE(request.FORM("ORTIPORG"))
    SORFALTANTES    	= UCASE(request.FORM("ORFALTANTES"))
    SORFECINID      	= UCASE(request.FORM("ORFECINID"))
    SORFECINIH      	= UCASE(request.FORM("ORFECINIH"))
    SORPERIODO      	= UCASE(request.FORM("ORPERIODO"))
    SORFECREGCONTABLED	= UCASE(request.FORM("ORFECREGCONTABLED"))
    SORFECREGCONTABLEH	= UCASE(request.FORM("ORFECREGCONTABLEH"))
    SORFECCARGAD		= UCASE(request.FORM("ORFECCARGAD"))
    SORFECCARGAH		= UCASE(request.FORM("ORFECCARGAH"))
    SORFECDESCUD		= UCASE(request.FORM("ORFECDESCUD"))
    SORFECDESCUH		= UCASE(request.FORM("ORFECDESCUH"))
    SORFECFIND			= UCASE(request.FORM("ORFECFIND"))
    SORFECFINH			= UCASE(request.FORM("ORFECFINH"))
    SORFECESTADOFIND	= UCASE(request.FORM("ORFECESTADOFIND"))
    SORFECESTADOFINH	= UCASE(request.FORM("ORFECESTADOFINH"))
    SORAREA				= UCASE(request.FORM("ORAREA"))
	SORCODIGOBCRA      	= UCASE(request.FORM("ORCODIGOBCRA"))


  else

    SORCODIGO       	= "TODOS"
    sORCARACTER     	= "TODOS"
    SORTEXTO        	= "TODOS"
    sORRIESGO       	= "TODOS"
    SORCUENTA       	= "TODOS"
    SORCLIENTE      	= "TODOS"
    SORMONTOD       	= "TODOS"
    SORMONTOH       	= "TODOS"
    SORPROCESO      	= "TODOS"
    SORLOTE         	= "TODOS"
    SOREVALUADOS    	= "TODOS"
    SORRIESGOSIDEN  	= "TODOS"
    SORSPAIDEN      	= "TODOS"
    SORESTADOS	    	= "TODOS"
    SORLINNEG       	= "TODOS"
    SORTIPORG       	= "TODOS"
    SORFALTANTES    	= "TODOS"
    SORFECINID      	= "TODOS"
    SORFECINIH      	= "TODOS"
    SORPERIODO      	= "TODOS"
    SORFECREGCONTABLED	= "TODOS"
    SORFECREGCONTABLEH	= "TODOS"
    SORFECCARGAD		= "TODOS"
    SORFECCARGAH		= "TODOS"
    SORFECDESCUD		= "TODOS"
    SORFECDESCUH		= "TODOS"
    SORFECFIND			= "TODOS"
    SORFECFINH			= "TODOS"
    SORFECESTADOFIND	= "TODOS"
    SORFECESTADOFINH	= "TODOS"
    SORAREA				= "TODOS"
	SORCODIGOBCRA		= "TODOS"
  end if


'---------------------------








    Celda =  ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","FECHA DE ESTADO FINALIZADO ","si")

	SINPUT = ofv.GenerarInput("ORFECESTADOFIND",SORFECINID,"TEXT",sForm,"30","dt","")
	Celda = Celda & ofv.gencelda("","VTI STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","40%","","","",SINPUT,"si")

	SINPUT = ofv.GenerarInput("ORFECESTADOFINH",SORFECINIH,"TEXT",sForm,"30","dt","")
	Celda = Celda & ofv.gencelda("","VTI STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","40%","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
'-------------------------


   Celda = ""

   wkproccod = "&nbsp;"

   Celda = Celda & ofv.GenCelda("","uti  ","","90%","","","","BASE DE EVENTOS DE RIESGO " & WKACCION,"si")
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

    Celda = ""
    sForm = "formlis"
     'CARACTER
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","20%","","","","NRO EVENTO","si")

    SINPUT = ofv.GenerarInput("ORCODIGO",SORCODIGO,"TEXT",sForm,"10","dtN","")
    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","80%","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
	
	Celda = ""
	Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","20%","","","","NRO BCRA","si")

    SINPUT = ofv.GenerarInput("ORCODIGOBCRA",SORCODIGOBCRA,"TEXT",sForm,"10","dtN","")
    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","80%","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")




    Celda = ""

     'CARACTER
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","FECHA INICIO ","si")

    SINPUT = ofv.GenerarInput("ORFECINID",SORFECINID,"TEXT",sForm,"30","dt","")
    Celda = Celda & ofv.gencelda("","VTI STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","40%","","","",SINPUT,"si")

    SINPUT = ofv.GenerarInput("ORFECINIH",SORFECINIH,"TEXT",sForm,"30","dt","")
    Celda = Celda & ofv.gencelda("","VTI STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","40%","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")


    Celda =  ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","FECHA DE REG. CONTABLE ","si")

	SINPUT = ofv.GenerarInput("ORFECREGCONTABLED",SORFECREGCONTABLED,"TEXT",sForm,"30","dt","")
	Celda = Celda & ofv.gencelda("","VTI STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","40%","","","",SINPUT,"si")

	SINPUT = ofv.GenerarInput("ORFECREGCONTABLEH",SORFECREGCONTABLEH,"TEXT",sForm,"30","dt","")
	Celda = Celda & ofv.gencelda("","VTI STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","40%","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")


    Celda =  ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","FECHA DE CARGA ","si")

	SINPUT = ofv.GenerarInput("ORFECCARGAD",SORFECCARGAD,"TEXT",sForm,"30","dt","")
	Celda = Celda & ofv.gencelda("","VTI STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","40%","","","",SINPUT,"si")

	SINPUT = ofv.GenerarInput("ORFECCARGAH",SORFECCARGAH,"TEXT",sForm,"30","dt","")
	Celda = Celda & ofv.gencelda("","VTI STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","40%","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")


    Celda =  ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","FECHA DE DESCUBRIMIENTO ","si")

	SINPUT = ofv.GenerarInput("ORFECDESCUD",SORFECDESCUD,"TEXT",sForm,"30","dt","")
	Celda = Celda & ofv.gencelda("","VTI STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","40%","","","",SINPUT,"si")

	SINPUT = ofv.GenerarInput("ORFECDESCUH",SORFECDESCUH,"TEXT",sForm,"30","dt","")
	Celda = Celda & ofv.gencelda("","VTI STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","40%","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

    Celda =  ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","FECHA DE FINALIZACION ","si")

	SINPUT = ofv.GenerarInput("ORFECFIND",SORFECFIND,"TEXT",sForm,"30","dt","")
	Celda = Celda & ofv.gencelda("","VTI STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","40%","","","",SINPUT,"si")

	SINPUT = ofv.GenerarInput("ORFECFINH",SORFECFINH,"TEXT",sForm,"30","dt","")
	Celda = Celda & ofv.gencelda("","VTI STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","40%","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

    Celda =  ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","FECHA DE ESTADO FINALIZADO ","si")

	SINPUT = ofv.GenerarInput("ORFECESTADOFIND",SORFECESTADOFIND,"TEXT",sForm,"30","dt","")
	Celda = Celda & ofv.gencelda("","VTI STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","40%","","","",SINPUT,"si")

	SINPUT = ofv.GenerarInput("ORFECESTADOFINH",SORFECESTADOFINH,"TEXT",sForm,"30","dt","")
	Celda = Celda & ofv.gencelda("","VTI STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","40%","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")




    Celda = ""

     'CARACTER
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","CARACTER","si")

    Strsql = "Select * From ORTIPOCARACTER ORDER BY DESCRI "
  '''''  sInput = ofv.GenerarCombo(sForm,"ORCARACTER",SORCARACTER,Strsql,cn,0,1,"")

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


    Celda = ""

     'CARACTER
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","DESCRIPCION ","si")
    SINPUT = ofv.GenerarInput("ORTEXTO",SORTEXTO,"TEXT",sForm,"70","dt","")

    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")


    Celda = ""

     'CARACTER
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","TIPO DE EVENTO","si")

    Strsql = "Select * From ORTIPOEVENTO ORDER BY DESCRI "
  ''  sInput = ofv.GenerarCombo(sForm,"ORRIESGO",SORRIESGO,Strsql,cn,0,1,"")

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


    Celda = ""

     'CARACTER
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","MONTO ","si")

    SINPUT = ofv.GenerarInput("ORMONTOD",SORMONTOD,"TEXT",sForm,"30","dt","")
    Celda = Celda & ofv.gencelda("","VTI STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","40%","","","",SINPUT,"si")

    SINPUT = ofv.GenerarInput("ORMONTOH",SORMONTOH,"TEXT",sForm,"30","dt","")
    Celda = Celda & ofv.gencelda("","VTI STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","40%","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")


    Celda = ""
     'PERIODO
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","PERIODO","si")


     Strsql = "Select *, (99999 - FINAL) AS XFFINAL From ORPERIODOS  where APLIC = '" & wkaplic & "' and "
     Strsql = Strsql & " ESTADO IN ('OK','CL') and RGVALEAUTOEV = 'S' "
     Strsql = Strsql & " ORDER BY XFFINAL DESC,INICIO, ESTADO DESC"

     set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
     SCOMSEP = chr(9)
     scomK   = "TODOS"
     scomD   = "TODOS"

  '   IF NOT RSX1.EOF THEN WKFPERIODO = rsx1(0)

  '   IF SORPERIODO = "TODOS" THEN SORPERIODO = WKFPERIODO

     do while not rsx1.eof
         scomK = scomK & SCOMSEP & rsx1(0)
         scomD = scomD & SCOMSEP & rsx1(1)
     rsx1.movenext
     loop

    sInput = ofv.gencomboFX(sForm,"ORPERIODO",SORPERIODO,scomK,scomD,"")



    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")



    Celda = ""
     'PROCESO
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","PROCESO","si")

    IF SORPERIODO = "TODOS" THEN

       SORPROCESO = "TODOS"

       SINPUT = "TODOS" & ofv.GenerarInput("ORPROCESO",SORPROCESO,"HIDDEN",sForm,"1","dt","")

    ELSE


         Strsql2 = "Select OBJID,PERIODO,RGAMBAUTOEV From ORPERIODOS  where APLIC = '" & wkaplic & "' and OBJID = 0" & SORPERIODO
         set rsX1 = ofv.crearconsultaEx(StrSql2,cn,1,parametros)
         if NOT rsX1.eof then
            WKORPERIODOMAP  = RSX1(2)
         ELSE
            WKORPERIODOMAP  = 0
         END IF


        Strsql = "Select OBJETOS.* From relaobj, OBJETOS  where CODOJASOC = CODOBJETO "
        Strsql = Strsql & " AND codobjmaster = '0" & WKORPERIODOMAP & "' and tipobjmaster = 'PROCESO'  "
        Strsql = Strsql & " ORDER BY OBJETOS.INDOBJ "



      set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
      SCOMSEP = chr(9)
      scomK   = "TODOS"
      scomD   = "TODOS"

      do while not rsx1.eof
         scomK = scomK & SCOMSEP & rsx1(0)
         scomD = scomD & SCOMSEP & rsx1(3)
      rsx1.movenext
      loop

      sInput = ofv.gencomboFX(sForm,"ORPROCESO",SORPROCESO,scomK,scomD,"")


    END IF



    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")

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



    Celda = ""


     'LINEA DE NEGOCIOS
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","LINEA DE NEGOCIOS","si")

    Strsql = "Select * From ORLINNEGO1  ORDER BY descripcion "
    set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

     SCOMSEP = chr(9)
     scomK   = "TODOS" & SCOMSEP & "0"
     scomD   = "TODOS" & SCOMSEP & "SIN ASIGNAR"

     do while not rsx1.eof
         scomK = scomK & SCOMSEP & rsx1(0)
         scomD = scomD & SCOMSEP & rsx1(1)
     rsx1.movenext
     loop

    sInput = ofv.gencomboFX(sForm,"ORLINNEG",SORLINNEG,scomK,scomD,"")



    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")


    Celda = ""


     'RIESGOS NIVEL 1
    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","EVENTO DE RIESGO","si")

    Strsql = "Select * From ORTIPOEVENTO ORDER BY SECUEN, DESCRI "
    set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

     SCOMSEP = chr(9)
     scomK   = "TODOS" & SCOMSEP & "0"
     scomD   = "TODOS" & SCOMSEP & "SIN ASIGNAR"

     do while not rsx1.eof
         scomK = scomK & SCOMSEP & rsx1(0)
         scomD = scomD & SCOMSEP & rsx1(1)
     rsx1.movenext
     loop

    sInput = ofv.gencomboFX(sForm,"ORTIPORG",SORTIPORG,scomK,scomD,"")



    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

    Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")






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


            Celda = ""
	     'DATOS FALTANTES
	    Celda = Celda & ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","DATOS FALTANTES","si")


	   scomK = "TODOS" & CHR(9) & "FECIDENT_N" & CHR(9) & "TRCPROCESO_X" & CHR(9) & "TRCPRODSER_X" & CHR(9) & "TRCIDENTACT_X"
	   scomK = scomK & CHR(9) & "TRCCARACTER_N" & CHR(9) & "TRCIMPBRUTO_N" & CHR(9) & "TRCECORESP_N" & CHR(9) & "TRCEVENTORIESGO1_N"
	   scomK = scomK & CHR(9) & "TRCEVENTORIESGO2_N" & CHR(9) & "TRCEVENTO_X" & CHR(9) & "TRCOBSERV_X" & CHR(9) & "TRCNEGOLINE_X"
	   scomK = scomK & CHR(9) & "TRCESTADO_X" & CHR(9) & "FECCONTA_N" & CHR(9) & "TRCCTACBLE_X" & CHR(9) & "TRCIMPSEGURO_N"
	   scomK = scomK & CHR(9) & "TRCRIESGOASOC_X" & CHR(9) & "FECORIG_N" & CHR(9) & "FECFINAL_N" & CHR(9) & "TRCCLIENTE_X"
	   scomK = scomK & CHR(9) & "TRCPRODUCTO_N" & CHR(9) & "TRCSECSOPORTE_X" & CHR(9) & "TRCCOSTOPOR_N"
       scomK = scomK & CHR(9) & "TRCSASIENTOCONT_N"



	   scomD = "SIN SELECCIONAR" & CHR(9) & "F.DESCUBRIMIENTO" & CHR(9) & "PROCESO" & CHR(9) & "SUBPROCESO" & CHR(9) & "ACTIVIDAD"
	   scomD = scomD & CHR(9) & "CARACTER" & CHR(9) & "MN.BRUTO" & CHR(9) & "AREA/SUCURSAL" & CHR(9) & "TIPO EVENTO" & CHR(9) & "CATEGORIA"
	   scomD = scomD & CHR(9) & "RIESGO" & CHR(9) & "DESCRIPCION" & CHR(9) & "LINEA NEGOCIOS" & CHR(9) & "ESTADO" & CHR(9) & "F.R.CONTABLE"
	   scomD = scomD & CHR(9) & "CUENTA" & CHR(9) & "COBERTURA SEGURO" & CHR(9) & "RIESGO ASOCIADO" & CHR(9) & "F.INICIO"
	   scomD = scomD & CHR(9) & "F.FINALIZACION" & CHR(9) & "CLIENTE" & CHR(9) & "PRODUCTO/SERVICIO" & CHR(9) & "AREA SOPORTE"
	   scomD = scomD & CHR(9) & "COSTO OPORTUNIDAD" & CHR(9) & "ASIENTO CONTABLE"


	   sInput = ofv.gencomboFX(sForm,"ORFALTANTES",SORFALTANTES,scomK,scomD,"")


	    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

	    Filax  = Filax & ofv.genrow("","AAA bgcolor='#F7AA22' ","","","","","",Celda,"si")




'--- area sector

	    Celda = ofv.gencelda("","VTI STYLE='background-color:#888888;color:#ffffff;'","","","","","","DEPENDENCIA/AREA/CENTRO DE COSTOS","si")


 if session("APLMDLOR") <> "BE" THEN

   IF  session("APLMDCLI")="BC" THEN
	 STRSQL="Select ECO,ECOMAP From ECO WHERE (ECOMAP  IS NOT NULL)  and ( len(ECOMAP) > 0)   ORDER BY ECOMAP "

   ELSE
   Strsql = "Select * From ECO ORDER BY DESCRI "

   END IF

  ELSE

   Strsql = "Select OBJID ,CODINTER + '-' + DESCRIPCION From SUCURSAL ORDER BY CODINTER "

  END IF


 '  RESPONSE.WRITE STRSQL

    set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)

	SCOMSEP = chr(9)
	scomK   = "TODOS"
	scomD   = "TODOS"

	do while not rsx1.eof
	   scomK = scomK & SCOMSEP & rsx1(0)
	   scomD = scomD & SCOMSEP & rsx1(1)
	   rsx1.movenext
	loop
    CALL OFV.CERRARCONSULTA(rsx1)
	sInput = ofv.gencomboFX(sForm,"ORAREA",SORAREA,scomK,scomD,"")






	    Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")

	    Filax  = Filax & ofv.genrow("","AAA bgcolor='"&xTrColor&"' ","","","","","",Celda,"si")



            Celda = ""

            IF sMODAL<> "ENCA" THEN
               SINPUT = "<INPUT TYPE='button' STYLE='background-color:#ff0000;color:#ffffff;' VALUE='EXTRAER' CLASS='BT2' "
               SINPUT = SINPUT & " onclick=""document." & sform & ".action= 'orbaseriesgoslis.asp?modal=" & smodal & "&CODTRACER=" & SCODTRACER & "&OPCION=VS&ESTILO=XTR';"
               SINPUT = SINPUT & " document." & sform & ".target='_blank'; "
               SINPUT = SINPUT & " document." & sform & ".submit();"" "
                     ''  SINPUT = SINPUT & " alert(document." & sform & ".action); "
                      '' SINPUT = SINPUT & " alert(document." & sform & ".target);"" "
               SINPUT = SINPUT & " >&nbsp;&nbsp;"
               Celda = Celda & ofv.gencelda("","ut  STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","CENTER","","","","",SINPUT,"si")
            else
               Celda = Celda & ofv.gencelda("","ut  STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","CENTER","","","","","&nbsp;","si")
            end if



            SINPUT = "<INPUT TYPE='button' VALUE='SELECCIONAR'"  &"CLASS='BT' "
            SINPUT = SINPUT & " onclick=""document." & sform & ".action= 'orbaseriesgoslis.asp?modal=" & smodal & "&CODTRACER=" & SCODTRACER & "&OPCION=VS&ESTILO=NORMAL&MODAL2="&SMODAL2&"';"
            SINPUT = SINPUT & " document." & sform & ".submit();"" "
  ''                  SINPUT = SINPUT & " alert(document." & sform & ".action); "
   ''                 SINPUT = SINPUT & " alert(document." & sform & ".target);"" "
            SINPUT = SINPUT & " >&nbsp;&nbsp;"

            Celda = Celda & ofv.gencelda("","ut colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","right","","","","",SINPUT,"si")

            Filax  = Filax & ofv.genrow("","AAA bgcolor='#DDDDFF' ","","","30","","",Celda,"si")



            sAccion = "ORBASERIESGOS.ASP?MODAL=" & SMODAL & "&CantRegMover=" & CantRegMoverx & "&MODALMAIN=" & sMODALMAIN
            sAccion = sAccion & "&CODTRACER=" & sCODTRACER & "&FECHA=" & sFECHA & "&OPCION=SETFRM"

            Formulario = ofv.GenForm(sForm,"",sAccion,"","post","",FilaX,"si")


            sHTML = sHTML & "<br><br><center>" & ofv.GenTabla("","aaa STYLE='border-style:SOLID;border-width:1;border-color:#888888;' ","","70% ","","0","","0","0","","",Formulario,"si") & "</center>"

  'Botonera
  Agregar = ""
  Buscar  = ""
  Celda   = ""

  if sMODAL = "ENCA"   then

  if  SMODAL2 <> "VTR"  then
     sAccion = "ORBASERIESGOSENCA.asp?CODTRACER=" & SCODTRACER & "&MODAL=ENR"
  ELSE
     sAccion = "ORBASERIESGOSENCA.asp?CODTRACER=" & SCODTRACER & "&MODAL=VTR"

  END IF

     Celda = ofv.BotonVolver(sAccion)
  END IF
' Celda = ofv.BotonVolver(sAccion)

  WSPREEMP = "modal2=" & smodal2 & "&modal=" & smodal  & "&"

  Agregar = ""
  if bAgregar AND sMODAL = "REG" AND sESTILO <> "XTR" then
    sAccion = "ORBASERIESGOSADD.asp?MODAL=AAG"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  sHTML = sHTML & replace(ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda),"cantregmostrar=",WSPREEMP & "cantregmostrar=") ' ofv.botoneraESSX("0","","","","","","","","","",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)





  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML



%>