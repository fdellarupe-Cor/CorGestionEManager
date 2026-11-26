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

  STIPOCAT = REQUEST.QUERYSTRING("TIPOCAT")

  IF STIPOCAT <> "" THEN
     SESSION("TIPOCAT") = STIPOCAT
  ELSE
     STIPOCAT = SESSION("TIPOCAT")
  END IF

  STIPOEV = SESSION("TIPOEV")

  sopcion = REQUEST.QUERYSTRING("opcion")



  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_TIPOCAT=" & STIPOCAT

  sHTML = ofv.FormHeader(session("FormN"))



  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)


  IF SOPCION = "SI" THEN



        strsql = "Insert into RGACTFAMILYRSKMITG "
        strsql = strsql & " (ACTFAMILYCOD,ACTFAMILYRSKCOD,ACTMITGCOD,ACTNIVEL,RGMODO)"
        strsql = strsql & "  Values "
        strsql = strsql &  "('" & STIPOEV & "','" & STIPOCAT & "','0','0','S')"

  '''response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

  END IF


  'Encabezados



  Strsql = "Select * From RGACTFAMILY WHERE OBJID = " & STIPOEV
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then 

     WKTDESCRI   = RS(1)


    Celda = ""
    Celda = Celda & ofv.gencelda("","UTI style='color:#FFFFFF;background-color:#888888;border-style:SOLID;border-width:1;border-color:#000000;' ","","20%","","","","FAMILIA","si")
    Celda = Celda & ofv.gencelda("","UTI style='color:#000000;background-color:#ddddff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","80%","","","",WKTDESCRI,"si")
    Fila = ofv.genrow("","","","","25","","",Celda,"si")

    Formulario = Fila



  Strsql = "Select B.DESCRI From RGACTFAMILYRSK A INNER JOIN ORAMENAZAS B ON ( A.ACTRSKCOD = B.OBJID) "
  Strsql = Strsql & " WHERE A.ACTFAMILYCOD = " & STIPOEV & " AND A.OBJID =  " & STIPOCAT
  '''response.write strsql & "<br>"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then 

     WKTDESCRI   = RS(0)

    Celda = ""
    Celda = Celda & ofv.gencelda("","UTI style='color:#FFFFFF;background-color:#888888;border-style:SOLID;border-width:1;border-color:#000000;' ","","","","","","AMENAZA","si")
    Celda = Celda & ofv.gencelda("","UTI style='color:#000000;background-color:#BBBBff;border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",WKTDESCRI,"si")
    Fila = ofv.genrow("","","","","25","","",Celda,"si")

    Formulario = Formulario & Fila

  END IF




    sHTML = sHTML & "<CENTER>" & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;border-COLOR:#000099;' ","","80% ","","0","","0","0","","",Formulario,"si") & "</CENTER>" & "<BR>"

  END IF


  'Encabezados
    Celda = ""
    Celda = Celda & ofv.gencelda("","UTI style='color:#FFFFFF;background-color:#888888;border-style:SOLID;border-width:1;border-color:#000000;' ","","1%","","","","&nbsp;","si")
    Celda = Celda & ofv.gencelda("","UTI style='color:#FFFFFF;background-color:#888888;border-style:SOLID;border-width:1;border-color:#000000;' ","","70%","","","","SALVAGUARDA","si")
    Celda = Celda & ofv.gencelda("","UTI style='color:#FFFFFF;background-color:#888888;border-style:SOLID;border-width:1;border-color:#000000;' ","","30%","","","","NIVEL","si")
    Fila = ofv.genrow("","","","","25","","",Celda,"si")

  Formulario = Fila


  WKSETTABLA = "RGACTFAMILYRSKMITG"

  Strsql = "Select * From " & WKSETTABLA & " WHERE ACTFAMILYRSKCOD = " & STIPOCAT & " Order By  OBJID DESC"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  posI = 0
  Do While (Not rs.EOF)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=" & WKSETTABLA & "&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar("SALVAGUARDA",sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'SALVAGUARDA
    Strsql = "Select * From ORSALVAGUARDAS ORDER BY DESCRI"
    sInput = ofv.GenerarCombo(sForm,RS(3).NAME,RS(3).value,Strsql,cn,0,1,"")
    Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",sInput,"si")


          Strsql = "Select * From ORNIVEFICIENCIA ORDER BY VALOR"
          sInput = ofv.GenerarCombo(sForm,RS(4).NAME,RS(4).value,Strsql,cn,0,4,"")


    Celda = Celda & ofv.gencelda("","UTI style='border-style:SOLID;border-width:1;border-color:#BBBBBB;' ","","","","","",SINPUT,"si")


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
     Celda = Celda & ofv.gencelda("","UTI COLSPAN=3 ","","","","","","NO HAY SALVAGUARDAS PREESTABLECIDAS PARA ESTA AMENAZA.","si")
     Fila = ofv.genrow("","","","","25","","",Celda,"si")
     Formulario = Formulario & Fila
  END IF

    Celda = ""
    Celda = Celda & ofv.gencelda("","UTI COLSPAN=3 ","","","","","","<HR>","si")

    Fila = ofv.genrow("","","","","25","","",Celda,"si")

  Formulario = Formulario & Fila

  sHTML = sHTML & "<CENTER>" & ofv.GenTabla("","AAA style='border-style:SOLID;border-width:1;border-COLOR:#000099;'","","80% ","","0","","0","0","","",Formulario,"si") & "</CENTER>"

  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "RGACTFAMILYDET2.asp?OPCION=SI&TIPOCAT=" & STIPOCAT & "&CANTREGMOVER=" & CANTREGMOVERX
    Agregar = ofv.BotonAgregar(sAccion)
  end if


  Buscar = ""
  Celda = ""

  sAccion = "RGACTFAMILYDET.asp?TIPOEV=" & STIPOEV & "&CANTREGMOVER=" & CANTREGMOVERX
  VOLVER = ofv.BotonVolver(sAccion)
  
  sHTML = sHTML & ofv.botoneraESSX("0","","","","no","no",posi,"","",Agregar,Buscar,celda,VOLVER)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>