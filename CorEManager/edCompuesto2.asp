<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  ofv.ObtenerAtributos Session("Form"),""
  bModificar = ofv.AModificar

  nCodigo = request.querystring("Codcomp")
  sTipo = request.querystring("Tipocomp")
  sOpcAgregar = request.querystring("OpcAgregar")
  vuelta = ofv.vueltaasp2("../")

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From Complementos Where codcomp=" & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then sNombre = rs(1) & "  -  " & rs(3)
  sNombre = ofv.convertircar(sNombre," ","&nbsp;")
  call ofv.cerrarconsulta(rs)

  Strsql = "Select * From Tipocomp Where tipo='" & sTipo & "'"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
    sImagen = rs(2)
    sComplemento = rs(1)
  end if
  call ofv.cerrarconsulta(rs)


  sHTML = ofv.FormHeader(session("FormN"))
'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

'-------------------------------------------------


  'sHTML = ofv.FormHeader("Edicion de " & sComplemento & " - Compuesto")

  'Encabezado
  sImagen = ofv.GenImage("","","","","Iconos/" & sImagen,"0","")
  sHref = "EditarImagen2.asp?ID=" & nCodigo
  sLink = ofv.GenLink("","",sHref,"_blank","","",sImagen,"si")
  Celda = ofv.GenCelda("","","","","","","",sLink,"si")
  Celda = Celda & ofv.GenCelda("","dt","","100%","","","",sNombre,"si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
'  Tabla = ofv.GenTabla("","","","80%","","0","","0","0","","",Fila,"si")

  'Encabezado Componentes
  Celda = ofv.GenCelda("","ttr","","2% colspan=3","","","","Componentes","si")
  Formulario = ofv.GenRow("","","","","25","","",Celda,"si")

  Strsql = "Select * From CSCompuesto Where codcomp = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  XPOS = 0
  do until rs.eof
    sForm = "edCompuesto2" & XPOS

      sql = "Select * from cscompuestosel "
      sql = sql & "Where codcomp = " & rs(4)
      set rsx = ofv.crearconsultaEx(Sql,cn,1,parametros)
      if not rsx.eof then sCompuesto = rsx(1)
      call ofv.cerrarconsulta(rsx)

    'Eliminar
    if bModificar then
      sHref = "asp/EliminarCompuesto.asp?Tabla=Compcompuesto&volver=" & vuelta & "&Codcomp="
      sHref = sHref & nCodigo & "&Tipocomp=" & sTipo & "&Eliminar=NSNC&ID=" & rs(8)
      Celda = ofv.BotonEliminar(sCompuesto,sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if


      sInput = ofv.GenerarInput("",sCompuesto,"readonly","","100","dt","")


    Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

      sID = "EdCompuestoCMP.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&opc=NO&CID=" & rs(8)
      sHref = sID
      sImage = ofv.GenImage("","","","","Iconos/ICorSTextov.jpg","0","")
      sLink = ofv.GenLink("","aaa title='Cambiar componente' ",sHref,"","","",sImage,"si")

    'sImagenP = "&#0232;"
    'sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
    sinput = "<a href='EdCompuestoCMP.asp?Codobjeto=" & nCodigo
    sinput = sinput & "&Tipobjeto=" & sTipo & "&opc=NO&CID=" & rs(8)
    sinput = sinput & "' title='Cambiar Componente'  ><i class='las la-arrow-right' ></i></a>" 
    Celda = Celda & ofv.gencelda("","ft","","","15%","","",sInput,"si")


  ''  Celda = Celda & ofv.GenCelda("","","left","","","","",sLink,"si")
    Fila = ofv.GenRow("","","","","","","",Celda,"si")

    sAccion = ""

    Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",Fila,"si")

    XPOS = XPOS + 1
    rs.movenext
  loop
  sHTML = sHTML & Tabla & ofv.GenTabla("","aaa","","80%","","0","","0","0","","",Formulario,"si")

  if sOpcAgregar = "si" then

     response.redirect "EdCompuestoCMP.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&opc=SI"



  else
    'Agregar
    Agregar = ""
    if bModificar then


       sAccion = "edCompuesto2.asp?Codcomp=" & nCodigo & "&Tipocomp=" & sTipo & "&OpcAgregar=si"
       Agregar = ofv.BotonAgregar(sAccion)

    end if
    '--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------
  
    sHTML = sHTML & ofv.botoneraESSX("8",rs,CantRegAMover,CantRegAMostrar,"","","",nCodigo,sTipo,"no",Agregar,"no","")
  end if
  sHTML = sHTML & "</body></html>"

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
end if

response.write sHTML

%>