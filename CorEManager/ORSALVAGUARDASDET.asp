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


  SNIVEL1  = REQUEST.QUERYSTRING("NIVEL1")


  IF SNIVEL1 <> "" THEN
     SESSION("NIVEL1") = SNIVEL1
  ELSE
     SNIVEL1 = SESSION("NIVEL1")
  END IF



  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados
'''''''  Formulario = ofv.Encabezados("CODIGO","Descripcion","FAMILIAS ACTIVOS")

  Formulario = ofv.Encabezados("Descripcion","FAMILIAS ACTIVOS")

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  sOPCION       = request.querystring("OPCION")
  IF SOPCION = "SI" THEN


        strsql = "Insert into ORSALVAGUARDAS "
        strsql = strsql & " (DESCRI,GRUPO)"
        strsql = strsql & "  Values "
        strsql = strsql &  "('NUEVA SALVAGUARDA','" & SNIVEL1 & "')"

     '''' response.write strsql & "<br>"
       ''' call ofv.crearconsultaEx(StrSql,cn,1,volver)
        xxx=SqlExLog("INSERT",strsql,CN,"ORSALVAGUARDAS","","","PAR")
       ' Function SqlExLog(XLTIPO,sql,CNEX1,TABLA,wkeyname,VALOR,TIPOPARM)

  END IF


 WKTEVENTO = "SIN DEFINIR"


  Strsql = "Select * From RGACTGRUPOSALVAGUARDAS WHERE OBJID = 0" & SNIVEL1
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     WKTEVENTO = RS(2)
  END IF

    Celda = ""
    Celda = Celda & ofv.gencelda("","VTI STYLE='COLOR:#000099;' ","","","","","","GRUPO:&nbsp;" & WKTEVENTO,"si")

    Fila = ofv.genrow("","","","","25","","",Celda,"si")
    sHTML = sHTML & ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ","","100% ","","0","","-1","1","","",Fila,"si")



  WKSETTABLA = "ORSALVAGUARDAS"

  Strsql = "Select * From " & WKSETTABLA & " WHERE GRUPO = '" & SNIVEL1 & "' Order By DESCRI"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
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

 '   'Tipo
 '   sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"10","dt","")
 '   Celda = Celda & ofv.gencelda("","cc","","50","","","",sInput,"si")

    'Descripcion
       nCol = 80
       if  not  isnull(rs(1)) then
           wdescri = rs(1)
       end if

       if  len(wdescri) = 0 then wdescri = " "

       t3b = "<font class=at><textarea class=dt style='font-size:7pt;' "
       t3b = t3b & "name=" & rs(1).name & " cols=" & nCol & " rows=2 "
       if  bModificar then
           t3b = t3b & " onblur=" & sForm & ".submit() "
           t3b = t3b & " onkeypress=" & chr(34)
           t3b = t3b & " if (this.value.length > 250) {alert('El texto es demasiado"
           t3b = t3b & " extenso'); return false;}" & chr(34) & " >"
       else
           t3b = t3b & " readonly >"
       end if
       xdatos =  wdescri
       t3b    = t3b & xdatos & "</textarea></font>"
       SINPUT = t3b

    Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")


    BTIT1 = "<A CLASS=CC HREF='ORSALVAGUARDASDETREL.ASP?TIPOEV=" & RS(0) & "' >RELACION</A>"
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
  sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si")

  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "ORSALVAGUARDASDET.asp?Opcion=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select descri From " & WKSETTABLA & " WHERE GRUPO = '" & SNIVEL1 & "' Order By DESCRI "    ' orden pagina
  str2 = "Select descri From " & WKSETTABLA & " WHERE GRUPO = '" & SNIVEL1 & "' Order By descri "    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  sAccion = "ORSALVAGUARDAS.asp"
  VOLVER = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,VOLVER)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>