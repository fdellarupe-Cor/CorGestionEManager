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

  sHTML = ofv.FormHeader(session("FormN"))  & chr(13)




  'Encabezados
  Formulario = ofv.Encabezados("Tipo","Descripcion","Pantalla","Cursograma","Propiedades")

  formulario = replace(formulario,">Icono"," align=center >Icono")


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From Modo Order By descri"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir and (rs(0) <> "I" and rs(0) <> "F") then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=Modo&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(1),sHref)
      celda = replace(celda,"<IMG  width=15  height=15  src=" & chr(34) & "imagenes/botdel.jpg" & chr(34) & "  border=0 >","<font face=corflaticon size=2 color=#aa0000>4</font>")
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'Tipo
    sInput = rs(0)
    Celda = Celda & ofv.gencelda("","cc","","50","","","",sInput,"si")

    'Descripcion
    If BrowserName() = "IE" OR BrowserName() = "CHROME" then
      nAncho = "150"
    else
      nAncho = "150"
    end if
    sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,nAncho,"dt","")
    Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")


    'Icono small
    sImagenValor = rs(2)
    sImage = ofv.GenImage("Imagen","","","","Iconos\" & sImagenValor,"","")
    Celda = Celda & ofv.gencelda("","ft  ","center","","","","",sImage,"si")


    'Icono large
    sImagenValorc = rs(10)
    if sImagenValorc & "" <> "DEFAULT" THEN
    sImage = ofv.GenImage("Imagen","","","","IconosCursograma\" & sImagenValorc,"","")
    Celda = Celda & ofv.gencelda("","ft ","center","","","","",sImage,"si")
    ELSE
      Celda = Celda & ofv.gencelda("","ft ","center","","","","","&nbsp;","si")
    END IF


    'sImagenP = "&#0232;"
    'sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
    sinput = "<a href='TiposdeAccionProp.asp?wperfil=" & rs(0) & "&CantRegMover=" & CantRegMoverx
    sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx
    sinput = sinput  & "' title='Propiedades de " & rs(1) & "' ><i class='las la-arrow-right' ></i></a>"
    Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")


    Fila = ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=Modo&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext
  Loop


  ftab = ofv.GenTabla("","","","100% ","","0","","0","0","","",Formulario,"si")

    Celda = ofv.GenCelda("","","center","","","","",ftab,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa ","","","","0","","0","0","","",filas,"si")


  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "AgregarTiposdeAccion.asp?Regisvuelta=" & CantRegMoverx & "&Opcion=si&tabla=Modo"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select descri From Modo Order By descri "    ' orden pagina
  str2 = "Select descri From Modo Order By descri "    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  'consulta

  sAccion = "TiposdeAccionCons.asp?" & request.querystring
  Celda = ofv.BotonConsultar(sAccion)


  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>