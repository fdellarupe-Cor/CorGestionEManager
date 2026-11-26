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

        strsql = "Insert into ORIMPRIESGO "
        strsql = strsql & " (DESCRI,AGRUPA,CATEGO,CODNEMO,secuen)"
        strsql = strsql & "  Values "
        strsql = strsql &  "('NUEVO EVENTO DE RIESGO','" & STIPOEV & "','" & STIPOCAT & "','XCOD','0')"

 '' response.write strsql & "<br>"
        call ofv.crearconsultaEx(StrSql,cn,1,volver)

  END IF

 WKTEVENTO = "SIN DEFINIR"
 WKTCATEGO = "SIN DEFINIR"


  Strsql = "Select DESCRI From ORTIPOEVENTO WHERE OBJID = " & STIPOEV
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     WKTEVENTO = RS(0)
  END IF

  Strsql = "Select DESCRI From ORTIPOCATEGO WHERE OBJID = " & STIPOCAT
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     WKTCATEGO = RS(0)
  END IF

    Celda = ""
    Celda = Celda & ofv.gencelda("","VTI STYLE='COLOR:#000099;' ","","","","","","TIPO EVENTO:&nbsp;" & WKTEVENTO & "&nbsp;/&nbsp;CATEGORIA:&nbsp;" & WKTCATEGO,"si")

    Fila = ofv.genrow("","","","","25","","",Celda,"si")
    sHTML = sHTML & ofv.GenTabla("","TWIND style='border-style:OUTSET;border-width:1;' valign='middle' align='center' ","","100% ","","0","","-1","1","","",Fila,"si")

  'Encabezados
  Formulario = ofv.Encabezados("COD. BCRA","CODIGO","COD.NEMOT.","Descripcion")


  WKSETTABLA = "ORIMPRIESGO"

'   if session("APLMDLOR") <> "BE"  THEN
'           Sordenar="	Order By objid "
'      else
 '          Sordenar = " Order By CAST(CODBCRA AS INT) "
'    end if

  Sordenar="	Order By CODBCRA "
  Strsql = "Select * From " & WKSETTABLA & " WHERE AGRUPA = " & STIPOEV & " AND CATEGO = " & STIPOCAT & Sordenar
  'response.write strsql
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
   ' sInput = ofv.GenerarInput(rs(5).name,rs(5),"text",sForm,"3","dtn","")
   ' Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")
    'sInput = ofv.GenerarInput(rs(6).name,rs(6) ,"text",sForm,"15","dtn  maxlength=2","SI")
   ' SINPUT="<input type=text class=dtN maxlength=3  name="&rs(6).name &" align=left value='" & rs(6) & "'  onchange=document."&sForm&".submit(); >"

   	'if session("APLMDLOR") <> "BE"  THEN
	  'sInput = ofv.GenerarInput(rs(6).name,rs(6) ,"text",sForm,"15","dtnSB  maxlength=3","SI")
	 ' SINPUT="<input name='"& rs(6).name   &"'  value='" & trim(rs(6)&"") & "' type=text MAXLENGTH=3 class=dtn onchange=document."&sForm&".submit();>"
   ' ELSE
       SINPUT="<input name='"& rs(6).name   &"'  value='" & trim(rs(6)&"") & "' type=text MAXLENGTH=4 class=dtn onblur="&chr(34)&"validarnum2(this,'"&sForm&"','"&trim(rs(6)&"")&"',4);"&chr(34) & ">"
	'end if

    Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

    'COD
    sInput = rs(0)
    Celda = Celda & ofv.gencelda("","cc","","50","","","",sInput,"si")

    'COD.NEMOTECNICO
    sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"10","dt","")
    Celda = Celda & ofv.gencelda("","cc","","50","","","",sInput,"si")

    'Descripcion
    sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"90","dt","")
    Celda = Celda & ofv.gencelda("","UT","","45%","","","",sInput,"si")




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
    sAccion = "ORIMPACTORIESGODET2.asp?OPCION=SI&TIPOCAT=" & STIPOCAT & "&CANTREGMOVER=" & CANTREGMOVERX
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select descri From " & WKSETTABLA & " WHERE AGRUPA = " & STIPOEV  & " AND CATEGO = " & STIPOCAT & Sordenar   ' orden pagina
  str2 = "Select descri From " & WKSETTABLA & " WHERE AGRUPA = " & STIPOEV & " AND CATEGO = " & STIPOCAT & Sordenar ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  'consulta
  sAccion = "ORIMPACTOCons.asp?" & request.querystring
  Celda = ofv.BotonConsultar(sAccion)

  sAccion = "ORIMPACTORIESGODET.asp?TIPOEV=" & STIPOEV
  VOLVER = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","",Agregar,Buscar,celda,VOLVER)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
  shtml=SHTML & 	ValidarInput()
  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>