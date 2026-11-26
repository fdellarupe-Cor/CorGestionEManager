<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrarx = 10
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  sUN = request.querystring("UN")

  IF sUN <> "" THEN

     SESSION("SUN") = SUN

  ELSE

    SUN = SESSION("SUN")

  END IF

  sHTML = ofv.FormHeader(session("FormN"))

  if SESSION("APLCHK") <> "S" then xxxx = APLCHECK()


  'response.write request.querystring
  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  sCod = request.querystring("cod")
  sTipo = request.querystring("tipo")

  Celda = OFV.GenCelda("", "tt", "center", "", "", "", "", "<B>SOLICITUD NORMATIVA&nbsp;" & sCod & "</B>", "si")
  Fila = OFV.GenRow("", "", "", "", "20", "", "", Celda, "si")
  Celda = OFV.GenCelda("", "", "", "", "", "", "", "&nbsp;", "si")
  Fila = Fila & OFV.GenRow("", "", "", "", "20", "", "", Celda, "si")
  Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "<B>DESARROLLADORES</B>", "si")
  Fila = Fila & OFV.GenRow("", "", "", "", "20", "", "", Celda, "si")

  sql = "select usr from objwfldet where codobjeto=" & sCod & " and tipobjeto='" & sTipo & "'"
  Set rs = OFV.crearconsultaEx(sql,cn,1,parametros)

  if not rs.eof then

  do while not rs.eof

	xxxx = getInfoUser(rs(0))

	Celda = OFV.GenCelda("", "dtn style='border:solid 1px;'", "", "", "", "", "","-&nbsp;" & session("SGU_NOMBRE"), "si")
	Fila = Fila & OFV.GenRow("", "", "", "", "20", "", "", Celda, "si")

  rs.movenext
  loop

  else

	Celda = OFV.GenCelda("", "dtn style='border:solid 1px;'", "", "", "", "", "","S/D", "si")
	Fila = Fila & OFV.GenRow("", "", "", "", "20", "", "", Celda, "si")

  end if

  sHTML = sHTML & OFV.GenTabla("", "a ", "center", "25%", "", "0", "", "-1", "0", "", "", Fila, "si") & "<br>"

  Celda = OFV.GenCelda("", "tt colspan=2","center", "", "", "", "", "<B>ESQUEMAS INVOLUCRADOS</B>", "si")
  Fila = OFV.GenRow("", "", "", "", "20", "", "", Celda, "si")

  Celda = OFV.GenCelda("", "tt", "", "", "", "", "", "DESCRIPCION", "si")
  Celda = Celda & OFV.GenCelda("", "tt", "", "", "", "", "", "ESTADO", "si")
  Fila = Fila & OFV.GenRow("", "", "", "", "20", "", "", Celda, "si")

  '''
  ssql = "select b.descri,c.estado "
  ssql = ssql & " from OBJPRTY a "
  ssql = ssql & " inner join objetos b on (a.codobjeto=b.codobjeto and a.tipobjeto=b.tipobjeto) "
  ssql = ssql & " inner join estados c on (a.estado=c.objid) "
  ssql = ssql & " where a.nrosolic=" & sCod
  Set rx = OFV.crearconsultaEx(ssql,cn,1,parametros)

  if not rx.eof then

  do while not rx.eof

  	Celda = OFV.GenCelda("", "dtn style='border:solid 1px;'", "", "", "", "", "",rx(0), "si")
  	Celda = Celda & OFV.GenCelda("", "dtn style='border:solid 1px;'", "", "", "", "", "",rx(1), "si")
  	Fila = Fila & OFV.GenRow("", "", "", "", "20", "", "", Celda, "si")

  rx.movenext
  loop
  ''''
  else

	Celda = OFV.GenCelda("", "dtn style='border:solid 1px;'", "", "", "", "", "","S/D", "si")
 	Celda = Celda & OFV.GenCelda("", "dtn style='border:solid 1px;'", "", "", "", "", "","S/D", "si")
  	Fila = Fila & OFV.GenRow("", "", "", "", "20", "", "", Celda, "si")

  end if

  sHTML = sHTML & OFV.GenTabla("", "a ", "center", "25%", "", "0", "", "-1", "0", "", "", Fila, "si")

  sAccion = "listadoseg.asp?menu=CONSGRAL&" & request.querystring
  Celda   = ofv.BotonVolver(sAccion)

  sHTML = sHTML & OFV.botoneraESSX("c", rs, CantRegMover, CantRegMostrar, "no", "no", PosReg, Addqry, "", "no", "", "", Celda)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>