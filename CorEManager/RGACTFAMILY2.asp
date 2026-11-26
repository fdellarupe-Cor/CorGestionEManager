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
  sOPCION       = request.querystring("OPCION")

  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados
  Formulario = ofv.Encabezados("codigo","Descripcion","&nbsp;")

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

   xTrColor = "#FFFFFF" 
   sImage = "<font face='corflaticon' size=2 Xcolor='#000088' style='border-style:SOLID;border-width:1;' >4</font>"

  set cn = ofv.conectar(ofv.strconn2)

'''''''''''''''''''''''''''''''''''''''
'' INCORPORO EVENTO EN BASE DE EVENTOS DE PERDIDA

  SNIVEL1  = REQUEST.QUERYSTRING("NIVEL1")


  IF SNIVEL1 <> "" THEN
     SESSION("NIVEL1") = SNIVEL1
  ELSE
     SNIVEL1 = SESSION("NIVEL1")
  END IF


  if sopcion = "SI" then



        strsql = "Insert into RGACTFAMILY "
        strsql = strsql & " (DESCRI,ACTCONFIDEN,ACTINTEGRID,ACTDISPONIB,RGMODO,VALORCAPA"
        strsql = strsql & " ) Values ("
        strsql = strsql &  "'NUEVA FAMILIA DE ACTIVOS','0','0','0','S','" & SNIVEL1 & "'"
        strsql = strsql & ")"

        xxx=SqlExLog("INSERT",strsql,CN,"RGACTFAMILY","","","PAR")


 ''' response.write strsql & "<br>"
       '' call ofv.crearconsultaEx(StrSql,cn,1,volver)


  end if


    Fila = ""

 WKTEVENTO = "SIN DEFINIR"


  Strsql = "Select * From RGACTCAPAS WHERE OBJID = 0" & SNIVEL1
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     WKTEVENTO = RS(2)
  END IF

    Celda = ""
    Celda = Celda & ofv.gencelda("","VTI STYLE='COLOR:#000099;' ","","","","","","CAPA:&nbsp;" & WKTEVENTO,"si")

    Fila = ofv.genrow("","","","","25","","",Celda,"si")
    sHTML = sHTML & ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ","","100% ","","0","","-1","1","","",Fila,"si")


  WKSETTABLA = "RGACTFAMILY"

  Strsql = "Select * From " & WKSETTABLA & " where valorCAPA = '" & SNIVEL1 & "' Order By DESCRI"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posI = 0
  posI = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=" & WKSETTABLA & "&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(1),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'Tipo
    sInput = rs(0)
    Celda = Celda & ofv.gencelda("","cc","","50","","","",sInput,"si")

    'Descripcion
    sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"90","dt","")
    Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

    BTIT1 = "<A CLASS=CC HREF='RGACTFAMILY3.ASP?FAMILIAID=" & RS(0) & "&NIVEL1=" & SNIVEL1 & "&CANTREGMOVER=" & CANTREGMOVERX & "' >CLASE ACTIVOS</A>"
    Celda = Celda & ofv.gencelda("","VT bgcolor='#F0F0F0' style='border-style:solid;border-width:1;border-color:#F7AA22;'  ","CENTER","10%","","","",BTIT1,"si")


    Fila = ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","25","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=" & WKSETTABLA & "&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")


       if xTrColor = "#FFFFFF" then
          xTrColor = "#F0F0F0" 
       ELSE
          xTrColor = "#FFFFFF" 
       END IF


    posi = posi + 1
    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si")


  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "RGACTFAMILY2.asp?opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select descri From " & WKSETTABLA & " where valorCAPA = '" & SNIVEL1 & "' Order By DESCRI "    ' orden pagina
  str2 = "Select descri From " & WKSETTABLA & " where valorCAPA = '" & SNIVEL1 & "' Order By descri "    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  'consulta
  sAccion = "RGACTFAMILYCons.asp?" & request.querystring
  Celda = ofv.BotonConsultar(sAccion)

  sAccion = "RGACTFAMILY.asp"
  VOLVER = ofv.BotonVolver(sAccion)
  
 
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,VOLVER)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>