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

  nCodigo    = request.querystring("Codcomp")
  sTipo      = request.querystring("Tipocomp")
  OpcAgregar = request.querystring("Opcion")


  session("comp31") = "TODOS"
  session("comp32") = "JPG"

  set cn = ofv.conectar(ofv.strconn2)

  sHTML = ofv.FormHeader(session("FormN"))

'----------------------------------------------

  shtml = shtml & CorPnlFnc("N","")

  sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

'-------------------------------------------------

  CantRegAMover   = cdbl(request.querystring("CantRegAMover"))
  CantRegAMostrar = 1
  vuelta          = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover)

  'Encabezado


  sImagenCampo = "Archivo"
  Strsql = "Select * From Comparc Where codcomp= " & nCodigo & " Order By secuencia"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if  not rs.eof then rs.move(CantRegAMover)
  if  not rs.eof then
      nID = rs(0)
      sImagenValor = rs(3)

''''''''''''''''''''''''''''''''''
      Celda = ""
      Fila  = ""

    'Eliminar
    if bModificar then
      sHref = "asp/EliminarC.asp?ID=" & rs(0) & "&Tabla=Comparc&volver=" & vuelta
      sHref = sHref & "&CantRegAMover=" & CantRegAMover & "&Codcomp=" & nCodigo
      sHref = sHref & "&Tipocomp=" & sTipo & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(3),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    sInput  = "<a class=ttt>Archivo&nbsp;" & (CantRegAMover + 1) & "</a>"
    Celda   = Celda & ofv.GenCelda("","uti","","80%","","","",sInput,"si")
    Fila    = Fila & ofv.GenRow("","","","","25","","",Celda,"si")


    Celda   = ""
    sInput  = "Archivo"
    Celda   = Celda & ofv.GenCelda("","tt","","20%","","","",sInput,"si")

    sInput  = ofv.GenerarInput("",sImagenValor,"readonly","","70","dt","")
    Celda   = Celda & ofv.GenCelda("","uti","","80%","","","",sInput,"si")
    Fila    = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

    Celda   = ""
    sInput  = "Descripcion"
    Celda   = Celda & ofv.GenCelda("","tt","","20%","","","",sInput,"si")

    if  bModificar then
        sInput = ofv.GenerarInput(rs(4).name,rs(4),"text","edImagen2","70","dt","")
    else
        sInput = ofv.GenerarInput("",rs(4),"readonly","","70","dt","")
    end if
    Celda   = Celda & ofv.GenCelda("","uti","","80%","","","",sInput,"si")
    Fila    = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

    Celda   = ""
    sInput  = "Cuerpo"
    Celda   = Celda & ofv.GenCelda("","tt","","20%","","","",sInput,"si")
    if  bModificar then
        sInput = ofv.GenerarInput(rs(5).name,rs(5),"text","edImagen2","10","dtn","SI")
    else
        sInput = ofv.GenerarInput("",rs(5),"readonly","","10","dt","")
    end if
    Celda   = Celda & ofv.GenCelda("","uti","","80%","","","",sInput,"si")
    Fila    = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

    Celda   = ""
    sInput  = "Seleccionar"
    Celda   = Celda & ofv.GenCelda("","tt","","20%","","","",sInput,"si")

    sHref = "EdimgarcCMP.asp?codobjeto=" & nCodigo & "&tipobjeto=" & stipo & "&opc=NO&cid=" & rs(0) & "&CantRegAMover=" & CantRegAMover & "&TIPOCOMP=" & STIPO
    'sImagenP = "&#0232;"
    'sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
    sinput   = "<a href='" & sHref
    sinput   = sinput & "' title='Seleccionar imagen' ><i class='las la-arrow-right'></i></a>"

    Celda   = Celda & ofv.GenCelda("","uti","","80%","","","",sInput,"si")
    Fila    = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

    Celda   = ""
    sInput  = "Visualizar"
    Celda   = Celda & ofv.GenCelda("","tt","","20%","","","",sInput,"si")

    sHref =  "ImagenCont5.asp?Archivo=" & sImagenValor
    'sImagenP = "&#0232;"
    'sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
    sinput   = "<a href='" & sHref
    sinput   = sinput & "' title='Visualizar archivo' target='_blank' ><i class='las la-arrow-right'></i></a>"

    Celda   = Celda & ofv.GenCelda("","uti","","80%","","","",sInput,"si")
    Fila    = Fila & ofv.GenRow("","","","","25","","",Celda,"si")

    sAccion = ""
    if  bModificar then
      sAccion = "asp/ActualizarC.asp?ID=" & nID & "&volver=" & vuelta
      sAccion = sAccion & "&Tabla=Comparc&Codcomp=" & nCodigo & "&Tipocomp=" & sTipo
    end if
    table=  ofv.GenTabla("","aaa","","95%","","0","0","0","0","","",Fila,"si")
    Formulario = ofv.GenForm("edImagen2","",sAccion,"","post","",table,"si")
    sHTML = sHTML & Formulario

    rs.movenext
  end if

    'Agregar
   sRetornar = "&Codcomp=" & nCodigo & "&Tipocomp=" & sTipo
   Agregar = ""
   if  bModificar then
    '   sInput  = ofv.GenerarInput("","Agregar","submit","","10","bt","")
    '   Celda   = ofv.GenCelda("","","","","","","",sInput,"si")
    '   sAccion = "EdimgarcCMP.asp?codobjeto=" & nCodigo & "&tipobjeto=" & stipo
    '   sAccion = sAccion & "&opc=SI&cid=0&CantRegAMover=" & CantRegAMover
    '   Agregar = ofv.GenForm("Agregar","",sAccion,"","post","",Celda,"si")

       sAccion = "EdimgarcCMP.asp?codobjeto=" & nCodigo & "&tipobjeto=" & stipo
       sAccion = sAccion & "&opc=SI&cid=0&CantRegAMover=" & CantRegAMover & "&TIPOCOMP=" & STIPO
       Agregar = ofv.BotonAgregar(sAccion)

   end if


'--------------------------------------------------------------------------------
  sHTML = sHTML & "</div>"

  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
  Fila = ofv.GenRow("","","","","","","",Celda,"si")
  sHTML = sHTML &  ofv.GenTabla("TCEX1","onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

  'Estructura del Indice

  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------

  sHTML = sHTML & ofv.botoneraESSX("3",rs,CantRegAMover,CantRegAMostrar,"no","no",posi,"","","no",Agregar,"no",sRetornar)


  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>