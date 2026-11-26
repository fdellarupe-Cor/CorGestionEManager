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
        strsql = strsql &  "('0','" & STIPOEV & "','0','0','0','0','S')"

  '''response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

  END IF

  wscomK = "0" & chr(9) & "1"
  wscomD = "NO" & chr(9) & "SI"


    sForm = "RGFAMILY"



    Fila = ""

 WKTEVENTO = "SIN DEFINIR"


  Strsql = "Select * From ORAMENAZAS WHERE OBJID = 0" & STIPOEV
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     WKTEVENTO = RS(2)
  END IF

    Celda = ""
    Celda = Celda & ofv.gencelda("","VTI STYLE='COLOR:#000099;' ","","","","","","AMENAZA:&nbsp;" & WKTEVENTO,"si")

    Fila = ofv.genrow("","","","","25","","",Celda,"si")
    sHTML = sHTML & ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ","","100% ","","0","","-1","1","","",Fila,"si")






  'Encabezados
'  Formulario = ofv.Encabezados("FAMILIA ACTIVOS","CONFIDENCIALIDAD","INTEGRIDAD","DISPONIBILIDAD")

    Celda = ""
    Celda = Celda & ofv.gencelda("","UTI COLSPAN=6 ","","","","","","<HR>","si")

    Fila = ofv.genrow("","","","","25","","",Celda,"si")

'  Formulario = Formulario & Fila

    Celda = ""
    Celda = Celda & ofv.gencelda("","UTI COLSPAN=2 style='color:#000000;background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#000000;' ","","","","","","FAMILIA ACTIVOS","si")
    Celda = Celda & ofv.gencelda("","UTI COLSPAN=3 style='color:#000000;background-color:#BBBBff;border-style:SOLID;border-width:1;border-color:#000000;' ","CENTER","","","","","DEGRADACION&nbsp;/&nbsp;IMPACTO","si")
    Celda = Celda & ofv.gencelda("","UTI  style='color:#000000;background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#000000;' ","","","","","","&nbsp;","si")
    Fila = ofv.genrow("","","","","25","","",Celda,"si")

'  Formulario = Formulario & Fila



    Celda = ""
    Celda = Celda & ofv.gencelda("","UTI COLSPAN=5 ","","","","","","<HR>","si")

    Fila = ofv.genrow("","","","","25","","",Celda,"si")

  Formulario = Formulario & Fila

    Celda = ""
    Celda = Celda & ofv.gencelda("","UTI COLSPAN=2 style='color:#000000;background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#000000;' ","","","","","","FAMILIA ACTIVOS","si")
    Celda = Celda & ofv.gencelda("","UTI COLSPAN=3 style='color:#000000;background-color:#BBBBff;border-style:SOLID;border-width:1;border-color:#000000;' ","CENTER","","","","","DEGRADACION&nbsp;/&nbsp;IMPACTO","si")
    Fila = ofv.genrow("","","","","25","","",Celda,"si")

  Formulario = Formulario & Fila

    Celda = ""
    Celda = Celda & ofv.gencelda("","UTI COLSPAN=2 style='color:#000000;background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#000000;' ","","","","","","&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UTI  style='color:#000000;background-color:#BBBBff;border-style:SOLID;border-width:1;border-color:#000000;' ","CENTER","","","","","CONFIDENCIALIDAD","si")
    Celda = Celda & ofv.gencelda("","UTI  style='color:#000000;background-color:#BBBBff;border-style:SOLID;border-width:1;border-color:#000000;' ","CENTER","","","","","INTEGRIDAD","si")
    Celda = Celda & ofv.gencelda("","UTI  style='color:#000000;background-color:#BBBBff;border-style:SOLID;border-width:1;border-color:#000000;' ","CENTER","","","","","DISPONIBILIDAD","si")
    Fila = ofv.genrow("","","","","25","","",Celda,"si")

  Formulario = Formulario & Fila



  WKSETTABLA = "RGACTFAMILYRSK"

  Strsql = "Select * From " & WKSETTABLA & " WHERE ACTRSKCOD = " & STIPOEV & " Order By  OBJID DESC"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  posI = 0
  Do While (Not rs.EOF)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=" & WKSETTABLA & "&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar("FAMILIA",sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'AMENAZA
    Strsql = "Select * From RGACTFAMILY ORDER BY DESCRI"
    sInput = ofv.GenerarCombo(sForm,RS(1).NAME,RS(1).value,Strsql,cn,0,1,"")
    Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


     WKTCONFIDEN = CINT("0" & RS(3))
     WKTINTEGRID = CINT("0" & RS(4))
     WKTDISPONIB = CINT("0" & RS(5))

     SINPUTD1 = ofv.gencomboFX(sForm,RS(3).NAME,WKTCONFIDEN,wscomK,wscomD," ")
     SINPUTD2 = ofv.gencomboFX(sForm,RS(4).NAME,WKTINTEGRID,wscomK,wscomD," ")
     SINPUTD3 = ofv.gencomboFX(sForm,RS(5).NAME,WKTDISPONIB,wscomK,wscomD," ")


    Celda = Celda & ofv.gencelda("","UT style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","CENTER","","","","",SINPUTD1,"si")
    Celda = Celda & ofv.gencelda("","UT style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","CENTER","","","","",SINPUTD2,"si")
    Celda = Celda & ofv.gencelda("","UT style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","CENTER","","","","",SINPUTD3,"si")




  '  BTIT1 = "<A CLASS=CC HREF='RGACTFAMILYDET2.ASP?TIPOCAT=" & RS(0) & "&CANTREGMOVER=" & CANTREGMOVERX & "' >SALVAGUARDAS</A>"
   ' Celda = Celda & ofv.gencelda("","VT bgcolor='#F0F0F0' style='border-style:solid;border-width:1;border-color:#F7AA22;'  ","CENTER","10%","","","",BTIT1,"si")


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
     Celda = Celda & ofv.gencelda("","UTI COLSPAN=6 ","","","","","","NO HAY FAMILIA DE ACTIVOS RELACIONADA.","si")
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
    sAccion = "ORAMENAZASDETREL.asp?OPCION=SI&TIPOEV=" & STIPOEV
    Agregar = ofv.BotonAgregar(sAccion)
  end if


  Buscar = ""
  Celda = ""

  sAccion = "ORAMENAZASDET.asp"
  VOLVER = ofv.BotonVolver(sAccion)
  
  sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no",posi,"","",Agregar,Buscar,celda,VOLVER)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>