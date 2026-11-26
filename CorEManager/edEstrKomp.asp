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

  nCodigo          = request.querystring("Codcomp")
  sTipo            = request.querystring("Tipocomp")
  CantRegAMover    = cdbl(request.querystring("CantRegAMover"))
  CantRegAMostrar  = 10
  vuelta           = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover)

  set cn = ofv.conectar(ofv.strconn2)

  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

'-------------------------------------------------

  'Encabezados
  Formulario1 = ofv.Encabezados("Estructura","Detalle")

  Strsql = "Select * From Compestr Where codcomp = " & nCodigo
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  nID = ""
  if  not rs.eof then rs.move(CantRegAMover)
  if  not rs.eof then

  do  until ((rs.eof) or (posi = CantRegAMostrar))
      nID    = rs(0)
      sForm  = "edEstr" & nID
      Celda  = ""
    'Eliminar
    if bModificar then
      sHref  = "asp/EliminarC.asp?ID=" & rs(0) & "&Tabla=Compestr&volver=" & vuelta
      sHref  = sHref & "&Codcomp=" & nCodigo & "&Tipocomp=" & sTipo & "&Eliminar=NSNC"
      Celda  = Celda & ofv.BotonEliminar(rs(3),sHref)
    else
      Celda  = Celda & ofv.GenCelda("","","","1%","","","","&nbsp;","si")
    end if

    'Estructura
    if bModificar then
      sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sForm,"90","dt","")
    else
      sInput = ofv.GenerarInput("",rs(3),"readonly","","120","dt","")
    end if
    'Celda    = Celda & ofv.GenCelda("","ut","","","","","",sInput,"si")

    'sImagenP = "&#0232;"
    'sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
    sinput2   = "<a href='edEstrKompDet.asp?Codcomp=" & nCodigo & "&Tipocomp=" & sTipo
    sinput2   = sinput2 & "&Kstr=" & nid
    sinput2   = sinput2 & "' title='Detalle de " & rs(3) & "' ><i class='las la-arrow-right' ></i></a>"
    'Celda    = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")

    'Fila     = ofv.GenRow("","","","","25","","",Celda,"si")

    sAccion  = ""
    if  bModificar then
        sAccion = "asp/ActualizarC.asp?ID=" & nID & "&Tabla=Compestr&volver=" & vuelta
        sAccion = sAccion & "&Codcomp=" & nCodigo & "&Tipocomp=" & sTipo
    end if
    Formulario = ofv.GenForm(sForm,"aa style=margin:0;",sAccion,"","post","",sInput & sinput2,"si")

      Celda = Celda & ofv.GenCelda("","vti","","90%","","","",formulario ,"si")
      Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

    rs.movenext
   loop

  else
      Celda      = ofv.GenCelda("","ttt colspan=3 ","","","","","","&nbsp;sin definir","si")
      Fila       = ofv.GenRow("","","","","30","","",Celda,"si")
      xtabla= ofv.GenTabla("","","","80%","","0","","0","0","","",formulario1 & fila,"si")
      Formulario = Formulario & ofv.GenForm(sForm,"aa style=margin:0;","","","post","",xtabla,"si")
  end if

  tabla= ofv.GenTabla("","aaa","","70%","","0","","0","0","","",formulario1 & fila,"si")


  sHTML = sHTML & tabla'' ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")

  'Botonera
  Agregar = ""
  if bModificar then
      sAccion = "AgregarEstrKomp.asp?Tabla=Compestr&volver=" & vuelta & "&Codcomp="
      sAccion = sAccion & nCodigo & "&Tipocomp=" & sTipo & "&ID=" & nID
      Agregar = ofv.BotonAgregar(sAccion)
  end if








'--------------------------------------------------------------------------------
  sHTML = sHTML & "</span></div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

  sHTML = sHTML & ofv.botoneraESSX("8",rs,CantRegAMover,CantRegAMostrar,"","","",nCodigo,sTipo,"no",Agregar,"no","")

  sHTML = sHTML & "</body></html>"

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
end if

response.write sHTML

%>