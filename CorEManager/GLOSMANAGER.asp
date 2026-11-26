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

  sHTML = ofv.FormHeader(session("FormN")) & chr(13)




  'Encabezados
  Formulario = ofv.Encabezados("PALABRA/FRASE","DESCRIPCION")



  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From GLOSARIO Order By KEYWORD"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir  then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=GLOSARIO&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(2),sHref)
      celda = replace(celda,"<IMG  width=15  height=15  src=" & chr(34) & "imagenes/botdel.jpg" & chr(34) & "  border=0 >","<font face=corflaticon size=2 color=#aa0000>4</font>")
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'Tipo
    sInput = rs(2)
    Celda = Celda & ofv.gencelda("","ttt","","25%","","","",sInput,"si")


    sInput = "&nbsp;"
 '   Celda = Celda & ofv.gencelda("","","","50%","","","",sInput,"si")



    'sImagenP = "&#0232;"
    'sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
    sinput = "<a href='GLOSMANAGERDSC.asp?GLOS=" & rs(0) & "&VMover=" & CantRegMoverx
    sinput = sinput  & "&CantRegMostrar=" & CantRegMostrarx
    sinput = sinput  & "' title='Consultar Descripciones' ><i class='las la-arrow-right'></i></a>"
    Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")


    Fila = ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=GLOSARIO&volver=" & vuelta
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
    sAccion = "AgregarGLOSARIO.asp?Regisvuelta=" & CantRegMoverx & "&Opcion=si&tabla=GLOSARIO"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select SHWWORD From GLOSARIO Order By KEYWORD "    ' orden pagina
  str2 = "Select SHWWORD From GLOSARIO Order By KEYWORD "    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

 ' sAccion = "GLOSMASTERCons.asp?" & request.querystring
 ' Celda = ofv.BotonConsultar(sAccion)
   Celda = ""

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>