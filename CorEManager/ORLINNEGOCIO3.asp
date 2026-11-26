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

  SNIVEL2  = REQUEST.QUERYSTRING("NIVEL2")


  IF SNIVEL2 <> "" THEN
     SESSION("NIVEL2") = SNIVEL2
  ELSE
     SNIVEL2 = SESSION("NIVEL2")
  END IF


  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados
  Formulario = ofv.Encabezados("ID","COD. BCRA","DESCRIPCION")

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  sOPCION       = request.querystring("OPCION")
  IF SOPCION = "SI" THEN

        strsql = "Insert into ORLINNEGOCIOS "
        strsql = strsql & " (nivel1,NIVEL2,descri)"
        strsql = strsql & "  Values "
        strsql = strsql &  "('" & SESSION("NIVEL1") & "','" & SNIVEL2 & "','NUEVA LINEA DE NEGOCIOS NIVEL 3')"

  '''response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

  END IF


 WKTEVENTO = "SIN DEFINIR"
 WKTEVENTO2 = ""


  Strsql = "Select * From ORLINNEGO1 WHERE OBJID = 0" & SESSION("NIVEL1")
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     WKTEVENTO = RS(1)
  END IF

  Strsql = "Select * From ORLINNEGO2 WHERE OBJID = 0" & SNIVEL2
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     WKTEVENTO2 = RS(2)
  END IF

    Celda = ""
    Celda = Celda & ofv.gencelda("","VTI STYLE='COLOR:#000099;' ","","","","","","LINEA DE NEGOCIOS:&nbsp;" & WKTEVENTO & "&nbsp;-&nbsp;" & WKTEVENTO2,"si")

    Fila = ofv.genrow("","","","","25","","",Celda,"si")
    sHTML = sHTML & ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ","","100% ","","0","","-1","1","","",Fila,"si")



  WKSETTABLA = "ORLINNEGOCIOS"

  Strsql = "Select * From " & WKSETTABLA & " WHERE NIVEL2 = 0" & SNIVEL2 & " ORDER BY CODBCRA" '" Order By NIVEL1,NIVEL2,DESCRI"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posI = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=" & WKSETTABLA & "&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(3),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

      sInput = RS(0)
	 Celda = Celda & ofv.gencelda("","cc","","","","","",sInput,"si")
     ' COD BCRA

	  'sInput = ofv.GenerarInput(rs(4).name,rs(4) ,"text",sForm,"15","dtn  maxlength=3","SI")
	 ' SINPUT="<input type=text class=dtN maxlength=3 name="&rs(4).name &" align=left value='" & trim(rs(4)&"") & "' onchange=document."&sForm&".submit();  >"
	 ' if session("APLMDLOR") <> "BE"  THEN
	 '   sInput = ofv.GenerarInput(rs(4).name,rs(4) ,"text",sForm,"15","dtn  maxlength=3","SI")
	 '  ELSE
	    SINPUT="<input name='"& rs(4).name   &"'  value='" & trim(rs(4)&"") & "' type=text MAXLENGTH=3 class=dtn onblur="&chr(34)&"validarnum2(this,'"&sForm&"','"&trim(rs(4)&"")&"',3);"&chr(34) & ">"
	 ' end if

	  Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")
    'Descripcion
    sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sForm,"80","dt","")
    Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")



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
    sAccion = "ORLINNEGOCIO3.asp?OPCION=SI"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select descri From " & WKSETTABLA & " WHERE NIVEL2 = 0" & SNIVEL2 & " Order By CODBCRA "    ' orden pagina
  str2 = "Select descri From " & WKSETTABLA & " WHERE NIVEL2 = 0" & SNIVEL2 & " Order By CODBCRA "    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  'consulta
  sAccion = "ORLINNEGOCIOCons.asp?" & request.querystring
  Celda = ofv.BotonConsultar(sAccion)

  sAccion = "ORLINNEGOCIO2.asp"
  VOLVER = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","",Agregar,Buscar,celda,VOLVER)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
  shtml=SHTML & 	ValidarInput()
  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>