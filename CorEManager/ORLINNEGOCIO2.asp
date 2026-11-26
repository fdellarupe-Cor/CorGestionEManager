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
  Formulario = ofv.Encabezados("ID","COD. BCRA","DESCRIPCION","NIVEL3")

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  sOPCION       = request.querystring("OPCION")
  IF SOPCION = "SI" THEN

        strsql = "Insert into ORLINNEGO2 "
        strsql = strsql & " (nivel1,descripcion)"
        strsql = strsql & "  Values "
        strsql = strsql &  "('" & SNIVEL1 & "','NUEVA LINEA DE NEGOCIOS NIVEL 2')"

  '''response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

  END IF


 WKTEVENTO = "SIN DEFINIR"

'  if session("APLMDLOR") <> "BE"  THEN
'         Sordenar="	Order By objid "
'    else
'         Sordenar = " Order By CAST(CODBCRA AS INT) "
'  end if
  Sordenar="	Order By CODBCRA "
  Strsql = "Select * From ORLINNEGO1 WHERE OBJID = 0" & SNIVEL1 & Sordenar
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     WKTEVENTO = RS(1)
  END IF

    Celda = ""
    Celda = Celda & ofv.gencelda("","VTI STYLE='COLOR:#000099;' ","","","","","","LINEA DE NEGOCIOS:&nbsp;" & WKTEVENTO,"si")

    Fila = ofv.genrow("","","","","25","","",Celda,"si")
    sHTML = sHTML & ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ","","100% ","","0","","-1","1","","",Fila,"si")



  WKSETTABLA = "ORLINNEGO2"

  Strsql = "Select * From " & WKSETTABLA & " WHERE NIVEL1 = 0" & SNIVEL1 & Sordenar '" Order BY descripcion"
  'RESPONSE.WRITE STRSQL
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
    sInput = rs(0)
    Celda = Celda & ofv.gencelda("","cc","","","","","",sInput,"si")



    ' COD BCRA

	  'sInput = ofv.GenerarInput(rs(3).name,rs(3) ,"text",sForm,"15","dtn  maxlength=2","SI")

'	  if session("APLMDLOR") <> "BE"  THEN
'	   sInput = ofv.GenerarInput(rs(3).name,rs(3) ,"text",sForm,"15","dtn  maxlength=2","SI")
'	   ELSE
	    SINPUT="<input name='"& rs(3).name   &"'  value='" & trim(rs(3)&"") & "' type=text MAXLENGTH=2 class=dtn onblur="&chr(34)&"validarnum(this,'"&sForm&"','"&trim(rs(3)&"")&"',2);"&chr(34) & ">"
'	  end if

	  Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")


    'descripcion
    sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"80","dt","")
    Celda = Celda & ofv.gencelda("","cc","","","","","",sInput,"si")


    BTIT1 = "<A CLASS=CC HREF='ORLINNEGOCIO3.ASP?NIVEL2=" & RS(0) & "' >NIVEL 3</A>"
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
    sAccion = "ORLINNEGOCIO2.asp?OPCION=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select DESCRIPCION From " & WKSETTABLA & " WHERE NIVEL1 = 0" & SNIVEL1 & Sordenar    ' orden pagina
  str2 = "Select DESCRIPCION From " & WKSETTABLA & " WHERE NIVEL1 = 0" & SNIVEL1 & Sordenar    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  'consulta
  sAccion = "ORLINNEGOCIOCons.asp?" & request.querystring
  Celda = ofv.BotonConsultar(sAccion)

  sAccion = "ORLINNEGOCIO.asp"
  VOLVER = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","",Agregar,Buscar,celda,VOLVER)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
  shtml=SHTML & 	ValidarInput()
  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>