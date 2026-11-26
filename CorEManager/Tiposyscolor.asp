<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<SCRIPT LANGUAGE="JavaScript" SRC="CorEManager/IncFile/jscolor/jscolor.js"></SCRIPT>
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
  Formulario = ofv.Encabezados("Id","Descripcion","codigo","*")


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From syscolor Order By descri"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=syscolor&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(1),sHref)

    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'id
    sInput = rs(0)
    Celda = Celda & ofv.gencelda("","cc","right","20","","","",sInput,"si")


    'Descripcion
    sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"80","dt","")
    Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

    'codigo
    sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"20","dt","")
    Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

    'muestra
    sInput = "<input type=button size=10 name=boton value='  ' style='border-radius:4px;border-width:1;background-color:" & rs(2) & "' >"
    ''sInput = ofv.GenerarInput("","&nbsp;&nbsp;","button",sForm,"10","dt style='background-color:" & rs(2) & "' ","")
    Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")






 '   Celda = Celda & ofv.gencelda("","","","","","","","","si")
    Fila = ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=syscolor&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext
  Loop
  'sHTML = sHTML & ofv.GenTabla(""," ","","80%","","0","","0","1","","",Formulario,"si")
  ftab = ofv.GenTabla("","","","100% ","","0","","0","0","","",Formulario,"si")

    Celda = ofv.GenCelda("","","center","","","","",ftab,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa ","","","","0","","0","0","","",filas,"si")

  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "AgregarTiposyscolor.asp?Regisvuelta=" & CantRegMoverx & "&Opcion=si&tabla=syscolor"
    Agregar = ofv.BotonAgregar(sAccion)

  end if



  str1 = "Select descri From syscolor Order By descri "    ' orden pagina
  str2 = "Select descri From syscolor Order By descri "    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  
  'consulta



  sinput = "<a  href=TiposyscolorCons.asp?"&request.querystring & " target=_new ><b>consulta</b></a>"
  Celda = ofv.gencelda("","tt style='border-style:outset;border-width:2;' ","center","","15%","","",sInput,"si")

  sAccion = "TiposyscolorCons.asp?" & request.querystring
  Celda = ofv.BotonConsultar(sAccion)
  
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>