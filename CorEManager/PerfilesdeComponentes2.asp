<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrarx = 1
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados
  Formulario = ofv.Encabezados("Perfil","Descripcion")

'  if request.querystring("seg") <> "" then
'    Session("Form") = ofv.convertircar(request.querystring("seg"),"_"," ")
'  end if

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From ComPerfil Order By descri"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=Comperfil&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(1),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'Perfil
    sInput = rs(0)
    Celda = Celda & ofv.gencelda("","cc","","50","","","",sInput,"si")

    'Descripcion
    sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"120","dt","")
    Celda = Celda & ofv.gencelda("","vt","","","","","",sInput,"si")

    Celda = Celda & ofv.gencelda("","","","","","","","","si")
    Fila = ofv.genrow("","","","","","","",Celda,"si")

    Celda = ofv.gencelda("","abc colspan=2","","","","","","","si")
    Celda = Celda & ofv.gencelda("","ut colspan=2 ","","2%","","","","&nbsp;","si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    'Perfiles
    Celda = ofv.gencelda("","abc colspan=2","","","","","","","si")
    Celda = Celda & ofv.gencelda("","ttt","","2%","","","","Propiedades","si")
    sInput = ofv.combosino(sForm,rs(2),rs(2).name)
    Celda = Celda & ofv.gencelda("","ft","","2%","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    Celda = ofv.gencelda("","abc colspan=2","","","","","","","si")
    Celda = Celda & ofv.gencelda("","ttt","","2%","","","","Estructura","si")
    sInput = ofv.combosino(sForm,rs(3),rs(3).name)
    Celda = Celda & ofv.gencelda("","ft","","2%","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    Celda = ofv.gencelda("","abc colspan=2","","","","","","","si")
    Celda = Celda & ofv.gencelda("","ttt","","2%","","","","Detalle de la estructura","si")
    sInput = ofv.combosino(sForm,rs(4),rs(4).name)
    Celda = Celda & ofv.gencelda("","ft","","2%","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    Celda = ofv.gencelda("","abc colspan=2","","","","","","","si")
    Celda = Celda & ofv.gencelda("","ttt","","2%","","","","Imagen de Complemento","si")
    sInput = ofv.combosino(sForm,rs(5),rs(5).name)
    Celda = Celda & ofv.gencelda("","ft","","2%","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    Celda = ofv.gencelda("","abc colspan=2","","","","","","","si")
    Celda = Celda & ofv.gencelda("","ttt","","2%","","","","Relacion Compuesta","si")
    sInput = ofv.combosino(sForm,rs(6),rs(6).name)
    Celda = Celda & ofv.gencelda("","ft","","2%","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")


    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=Comperfil&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","","","80% ","","0","","0","0","","",Formulario,"si")

  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "AgregarPerfilesdeComponentes.asp?Regisvuelta=" & CantRegMoverx
    sAccion = sAccion & "&Opcion=si&tabla=Comperfil"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select descri From Comperfil Order By descri "    ' orden pagina
  str2 = "Select descri From Comperfil Order By descri "    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))
  
  'consulta
  sinput = "<a  href=PerfilesdeComponentesCons.asp?"&request.querystring & " target=_new ><b>consulta</b></a>"
  Celda = ofv.gencelda("","tt style='border-style:outset;border-width:2;' ","center","","15%","","",sInput,"si")
  
  sAccion = "PerfilesdeComponentesCons.asp?" & request.querystring
  Celda = ofv.BotonConsultar(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>