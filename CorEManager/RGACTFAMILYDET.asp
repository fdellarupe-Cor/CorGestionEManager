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

  STIPOEV = REQUEST.QUERYSTRING("TIPOEV")

  IF STIPOEV <> "" THEN
     SESSION("TIPOEV") = STIPOEV
  ELSE
     STIPOEV = SESSION("TIPOEV")
  END IF

  sopcion = REQUEST.QUERYSTRING("opcion")



  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_TIPOEV=" & STIPOEV

  sHTML = ofv.FormHeader(session("FormN"))



  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)


  IF SOPCION = "SI" THEN

        strsql = "Insert into RGACTFAMILYRSK "
        strsql = strsql & " (ACTFAMILYCOD,ACTRSKCOD,ACTCONFIDEN,ACTINTEGRID,ACTDISPONIB,ACTRSKDEGRADA,RGMODO)"
        strsql = strsql & "  Values "
        strsql = strsql &  "('" & STIPOEV & "','0','100','0','0','0','S')"

  '''response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

  END IF


  IF SOPCION = "MODFAMILY" THEN

        SINPUTD1 = CINT("0" & REQUEST.Form("WCONFI"))
        SINPUTD2 = CINT("0" & REQUEST.Form("WINTEG"))
        SINPUTD3 = CINT("0" & REQUEST.Form("WDISPO"))


        IF NOT ((SINPUTD1 + SINPUTD2 + SINPUTD3) > 100) THEN

            strsql = "UPDATE RGACTFAMILY SET "
            strsql = strsql & " ACTCONFIDEN = '" & SINPUTD1 & "' ,"
            strsql = strsql & " ACTINTEGRID = '" & SINPUTD2 & "' ,"
            strsql = strsql & " ACTDISPONIB = '" & SINPUTD3 & "' "

            strsql = strsql &  " WHERE OBJID = " & STIPOEV
  '''response.write strsql & "<br>"
            call ofv.crearconsultaEx(StrSql,cn,1,volver)

            WSPOBSERV = ""

        ELSE


            WSPOBSERV = "LA SUMA DE LAS PONDERACIONES DE TODAS LAS DIMENSIONES NO PUEDE SUPERAR 100."

        END IF


  END IF





      wscomK = "0" & chr(9) & "10" & chr(9) & "20" & chr(9) & "30"
      wscomK = wscomK & chr(9) & "40" & chr(9) & "50" & chr(9) & "60"
      wscomK = wscomK & chr(9) & "70" & chr(9) & "80" & chr(9) & "90" & chr(9) & "100"


      wscomD = "N/A" & chr(9) & "10%" & chr(9) & "20%" & chr(9) & "30%"
      wscomD = wscomD & chr(9) & "40%" & chr(9) & "50%" & chr(9) & "60%"
      wscomD = wscomD & chr(9) & "70%" & chr(9) & "80%" & chr(9) & "90%" & chr(9) & "100%"


        '''  sInput = ofv.gencomboFX(sForm,rs(9).name,rs(9),wscomK,wscomD," ")
    sForm = "RGFAMILY"

  'Encabezados

    Celda = ""
    Celda = Celda & ofv.gencelda("","UTI style='color:#FFFFFF;background-color:#888888;border-style:SOLID;border-width:1;border-color:#000000;' ","","1%","","","","&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UTI style='color:#FFFFFF;background-color:#888888;border-style:SOLID;border-width:1;border-color:#000000;' ","","80%","","","","FAMILIA","si")

    Celda = Celda & ofv.gencelda("","UTI style='color:#FFFFFF;background-color:#888888;border-style:SOLID;border-width:1;border-color:#000000;' ","","5%","","","","CONFIDENCIALIDAD&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UTI style='color:#FFFFFF;background-color:#888888;border-style:SOLID;border-width:1;border-color:#000000;' ","","5%","","","","INTEGRIDAD&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UTI style='color:#FFFFFF;background-color:#888888;border-style:SOLID;border-width:1;border-color:#000000;' ","","5%","","","","DISPONIBILIDAD&nbsp;","si")

    Celda = Celda & ofv.gencelda("","UTI style='color:#FFFFFF;background-color:#888888;border-style:SOLID;border-width:1;border-color:#000000;' ","","5%","","","","&nbsp;","si")


    Fila = ofv.genrow("","","","","25","","",Celda,"si")

  Formulario = Fila

  Strsql = "Select * From RGACTFAMILY WHERE OBJID = " & STIPOEV
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then 

     WKTDESCRI   = RS(1)

     WKTCONFIDENM = CINT("0" & RS(2))
     WKTINTEGRIDM = CINT("0" & RS(3))
     WKTDISPONIBM = CINT("0" & RS(4))


     SINPUTD1 = ofv.gencomboFX(sForm,"WCONFI",WKTCONFIDENM,wscomK,wscomD," ")
     SINPUTD2 = ofv.gencomboFX(sForm,"WINTEG",WKTINTEGRIDM,wscomK,wscomD," ")
     SINPUTD3 = ofv.gencomboFX(sForm,"WDISPO",WKTDISPONIBM,wscomK,wscomD," ")


    Celda = ""
    Celda = Celda & ofv.gencelda("","VTI COLSPAN=2 style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","80%","","","",WKTDESCRI,"si")

    Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","CENTER","5%","","","",SINPUTD1,"si")
    Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","CENTER","5%","","","",SINPUTD2,"si")
    Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","CENTER","5%","","","",SINPUTD3,"si")

    Celda = Celda & ofv.gencelda("","UTI style='background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","5%","","","","&nbsp;","si")


    Fila = ofv.genrow("","","","","25","","",Celda,"si")

    sAccion = ""
    if bModificar then
       sAccion = "RGACTFAMILYDET.asp?OPCION=MODFAMILY&TIPOEV=" & STIPOEV & "&CANTREGMOVER=" & CANTREGMOVERX
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    IF WSPOBSERV <> "" THEN


       Celda = ""
       Celda = Celda & ofv.gencelda("","VTI COLSPAN=6 style='border-style:SOLID;border-width:1;border-color:#993333;color:#993333;' ","CENTER","","","","",WSPOBSERV,"si")
       Formulario = Formulario & ofv.genrow("","","","","25","","",Celda,"si")

    END IF



 '   sHTML = sHTML & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;border-COLOR:#000099;' ","","100% ","","0","","-1","1","","",Formulario,"si")

  END IF


  'Encabezados
'  Formulario = ofv.Encabezados("AMENAZAS","CONFIDENCIALIDAD","INTEGRIDAD","DISPONIBILIDAD","&nbsp;")

    Celda = ""
    Celda = Celda & ofv.gencelda("","UTI COLSPAN=6 ","","","","","","<HR>","si")

    Fila = ofv.genrow("","","","","25","","",Celda,"si")

  Formulario = Formulario & Fila

    Celda = ""
    Celda = Celda & ofv.gencelda("","UTI COLSPAN=2 style='color:#000000;background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#000000;' ","","","","","","AMENAZAS","si")
    Celda = Celda & ofv.gencelda("","UTI COLSPAN=3 style='color:#000000;background-color:#BBBBff;border-style:SOLID;border-width:1;border-color:#000000;' ","CENTER","","","","","DEGRADACION&nbsp;/&nbsp;IMPACTO","si")
    Celda = Celda & ofv.gencelda("","UTI  style='color:#000000;background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#000000;' ","","","","","","&nbsp;","si")
    Fila = ofv.genrow("","","","","25","","",Celda,"si")

  Formulario = Formulario & Fila



  WKSETTABLA = "RGACTFAMILYRSK"

  Strsql = "Select * From " & WKSETTABLA & " WHERE ACTFAMILYCOD = " & STIPOEV & " Order By  OBJID DESC"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  posI = 0
  Do While (Not rs.EOF)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=" & WKSETTABLA & "&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar("AMENAZA",sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'AMENAZA
    Strsql = "Select * From ORAMENAZAS ORDER BY CODNEMO"
    sInput = ofv.GenerarCombo(sForm,RS(2).NAME,RS(2).value,Strsql,cn,0,2,"")
    Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


     WKTCONFIDEN = CINT("0" & RS(3))
     WKTINTEGRID = CINT("0" & RS(4))
     WKTDISPONIB = CINT("0" & RS(5))


 ''    WKTCONFIDENM = CINT("0" & RS(2))
 ''    WKTINTEGRIDM = CINT("0" & RS(3))
 ''    WKTDISPONIBM = CINT("0" & RS(4))


     SINPUTD1 = "<HR SIZE=1 WIDTH=60% style='border-style:DOTTED;border-color:#993333;' >"
     SINPUTD2 = "<HR SIZE=1 WIDTH=60% style='border-style:DOTTED;border-color:#993333;'>"
     SINPUTD3 = "<HR SIZE=1 WIDTH=60% style='border-style:DOTTED;border-color:#993333;'>"



     IF WKTCONFIDENM > 0 THEN SINPUTD1 = ofv.gencomboFX(sForm,RS(3).NAME,WKTCONFIDEN,wscomK,wscomD," ")
     IF WKTINTEGRIDM > 0 THEN SINPUTD2 = ofv.gencomboFX(sForm,RS(4).NAME,WKTINTEGRID,wscomK,wscomD," ")
     IF WKTDISPONIBM > 0 THEN SINPUTD3 = ofv.gencomboFX(sForm,RS(5).NAME,WKTDISPONIB,wscomK,wscomD," ")


    Celda = Celda & ofv.gencelda("","UT style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","CENTER","","","","",SINPUTD1,"si")
    Celda = Celda & ofv.gencelda("","UT style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","CENTER","","","","",SINPUTD2,"si")
    Celda = Celda & ofv.gencelda("","UT style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","CENTER","","","","",SINPUTD3,"si")




    BTIT1 = "<A CLASS=CC HREF='RGACTFAMILYDET2.ASP?TIPOCAT=" & RS(0) & "&CANTREGMOVER=" & CANTREGMOVERX & "' >SALVAGUARDAS</A>"
    Celda = Celda & ofv.gencelda("","VT bgcolor='#F0F0F0' style='border-style:solid;border-width:1;border-color:#F7AA22;'  ","CENTER","10%","","","",BTIT1,"si")


    Fila = ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=" & WKSETTABLA & "&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext
  Loop

  IF POSI = 0 THEN
     Celda = ""
     Celda = Celda & ofv.gencelda("","UTI COLSPAN=6 ","","","","","","NO HAY AMENAZAS PREESTABLECIDAS PARA ESTA FAMILIA DE ACTIVOS.","si")
     Fila = ofv.genrow("","","","","25","","",Celda,"si")
     Formulario = Formulario & Fila
  END IF

    Celda = ""
    Celda = Celda & ofv.gencelda("","UTI COLSPAN=6 ","","","","","","<HR>","si")

    Fila = ofv.genrow("","","","","25","","",Celda,"si")

  Formulario = Formulario & Fila

  sHTML = sHTML & "<CENTER>" & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;border-COLOR:#000099;'","","80% ","","0","","0","0","","",Formulario,"si") & "</CENTER>"

  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "RGACTFAMILYDET.asp?OPCION=SI&TIPOEV=" & STIPOEV & "&CANTREGMOVER=" & CANTREGMOVERX
    Agregar = ofv.BotonAgregar(sAccion)
  end if


  Buscar = ""
  Celda = ""

  sAccion = "RGACTFAMILY.asp?CANTREGMOVER=" & CANTREGMOVERX
  VOLVER = ofv.BotonVolver(sAccion)
  
  sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no",posi,"","",Agregar,Buscar,celda,VOLVER)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>