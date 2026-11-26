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

        strsql = "Insert into ORTIPOCATEGO "
        strsql = strsql & " (TIPOEVENTO,DESCRI,secuen)"
        strsql = strsql & "  Values "
        strsql = strsql &  "('" & STIPOEV & "','NUEVA CATEGORIA DE EVENTO','0')"

  '''response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

  END IF

 WKTEVENTO = "SIN DEFINIR"

'  if session("APLMDLOR") <> "BE"  THEN
'           Sordenar="	Order By objid "
'      else
'           Sordenar = " Order By CAST(CODBCRA AS INT) "
'  end if
  Sordenar="	Order By CODBCRA "
  Strsql = "Select DESCRI From ORTIPOEVENTO WHERE OBJID = " & STIPOEV
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     WKTEVENTO = RS(0)
  END IF

    Celda = ""
    Celda = Celda & ofv.gencelda("","VTI STYLE='COLOR:#000099;' ","","","","","","TIPO EVENTO:&nbsp;" & WKTEVENTO,"si")

    Fila = ofv.genrow("","","","","25","","",Celda,"si")
    sHTML = sHTML & ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ","","100% ","","0","","-1","1","","",Fila,"si")

  'Encabezados
  Formulario = ofv.Encabezados("COD. BCRA","CODIGO","Descripcion","&nbsp;")


  WKSETTABLA = "ORTIPOCATEGO"

  Strsql = "Select * From " & WKSETTABLA & " WHERE TIPOEVENTO = " & STIPOEV & Sordenar
'  RESPONSE.WRITE STRSQL
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posI = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=" & WKSETTABLA & "&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(2),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'sec
    'sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sForm,"3","dtn","")
    'Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

     ' COD BCRA

	 'sInput = ofv.GenerarInput(rs(4).name,rs(4) ,"text",sForm,"15","dtn  maxlength=2","SI")

'	 if session("APLMDLOR") <> "BE"  THEN
'	   sInput = ofv.GenerarInput(rs(4).name,rs(4) ,"text",sForm,"15","dtn  maxlength=2","SI")
'	 ELSE
	   SINPUT="<input name='"& rs(4).name   &"'  value='" & trim(rs(4)&"") & "' type=text MAXLENGTH=2 class=dtn onblur="&chr(34)&"validarnum(this,'"&sForm&"','"&trim(rs(4)&"")&"',2);"&chr(34) & ">"
'	 end if

	 Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

    'Tipo
    sInput = RS(0)
    Celda = Celda & ofv.gencelda("","cc","","50","","","",sInput,"si")

    'Descripcion
    sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"90","dt","")
    Celda = Celda & ofv.gencelda("","UT","","45%","","","",sInput,"si")


    BTIT1 = "<A CLASS=CC HREF='ORIMPACTORIESGODET2.ASP?TIPOCAT=" & RS(0) & "' >RIESGOS</A>"
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
    sAccion = "ORIMPACTORIESGODET.asp?OPCION=SI&TIPOEV=" & STIPOEV & "&CANTREGMOVER=" & CANTREGMOVERX
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select descri From " & WKSETTABLA & " WHERE TIPOEVENTO = " & STIPOEV & sordenar '" Order By CODBCRA "    ' orden pagina
  str2 = "Select descri From " & WKSETTABLA & " WHERE TIPOEVENTO = " & STIPOEV & sordenar    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  'consulta
  sAccion = "ORIMPACTOCons.asp?" & request.querystring
  Celda = ofv.BotonConsultar(sAccion)

  sAccion = "ORIMPACTORIESGO.asp"
  VOLVER = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","",Agregar,Buscar,celda,VOLVER)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
  shtml=SHTML & 	ValidarInput()
  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>