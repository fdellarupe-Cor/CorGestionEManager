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


  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados
  Formulario = ofv.Encabezados("ID","COD. BCRA","DESCRIPCION","NIVEL2")

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  sOPCION       = request.querystring("OPCION")
  IF SOPCION = "SI" THEN

        strsql = "Insert into ORLINNEGO1 "
        strsql = strsql & " (descripcion)"
        strsql = strsql & "  Values "
        strsql = strsql &  "('NUEVA LINEA DE NEGOCIOS')"

  '''response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

  END IF


  WKSETTABLA = "ORLINNEGO1"

'  if session("APLMDLOR") <> "BE"  THEN
'       Sordenar="	Order By objid "
'  else
'       Sordenar = " Order By CAST(CODBCRA AS INT) "
'  end if

  Sordenar="	Order By CODBCRA "
  Strsql = "Select * From " & WKSETTABLA & Sordenar

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
     sInput = RS(0)
	 Celda = Celda & ofv.gencelda("","cc","","","","","",sInput,"si")


     ' COD BCRA

	 'sInput = ofv.GenerarInput(rs(2).name,rs(2) ,"text",sForm,"15","dtn  maxlength=2","SI")

'	  if session("APLMDLOR") <> "BE"  THEN
'	 	sInput = ofv.GenerarInput(rs(2).name,rs(2) ,"text",sForm,"15","dtn  maxlength=2","SI")
'	   ELSE
	 	SINPUT="<input name='"& rs(2).name   &"'  value='" & trim(rs(2)&"") & "' type=text MAXLENGTH=2 class=dtn onblur="&chr(34)&"validarnum(this,'"&sForm&"','"&trim(rs(2)&"")&"',2);"&chr(34) & ">"
'	  end if

	 Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

    'descripcion
    sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"80","dt","")
    Celda = Celda & ofv.gencelda("","cc","","","","","",sInput,"si")


    BTIT1 = "<A CLASS=CC HREF='ORLINNEGOCIO2.ASP?NIVEL1=" & RS(0) & "' >NIVEL 2</A>"
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
    sAccion = "ORLINNEGOCIO.asp?OPCION=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select DESCRIPCION From " & WKSETTABLA & Sordenar    ' orden pagina
  str2 = "Select DESCRIPCION From " & WKSETTABLA & Sordenar    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  'consulta
  sAccion = "ORLINNEGOCIOCons.asp?" & request.querystring
  Celda = ofv.BotonConsultar(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
  shtml=SHTML & 	ValidarInput()
  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>