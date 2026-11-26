<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


  CantRegMoverx = cdbl(request.querystring("CantRegMover"))

  wperfil = request.querystring("wperfil")
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_wperfil=" & wperfil


  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados
 '' Formulario = ofv.Encabezados("Tipo","Descripcion","Archivo","Icono","Perfil")



  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From Tipocomp where tipo = '" & wperfil & "' "
''response.write Strsql & " strsql " & chr(13)
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
  posi = 0

    sForm = "form" & posi

    'Tipo
     sInput = rs(0) & " - " & rs(1)
     Celda = ofv.gencelda("","ttt colspan=2","","","","","",sInput,"si")

     Fila = ofv.genrow("","","","","","","",Celda,"si")

    'Archivo
     Celda = ofv.gencelda("","tt","","20%","","","","Archivo","si")

    'Archivo
     sImagenValor = rs(2)
     sImagenCampo = rs(2).name
     sInput = ofv.GenerarCombo3(sForm,ofv.Getfiles("Iconos"),sImagenCampo,sImagenValor)
     Celda = Celda & ofv.gencelda("","vti ","","80%","","","",sInput,"si")

     Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

     'Icono
     Celda = ofv.gencelda("","tt","","","","","","Icono","si")

    'Icono
     sImage = ofv.GenImage("Imagen","","","","Iconos\" & sImagenValor,"","")
     Celda = Celda & ofv.gencelda("","fti  ","","","","","",sImage,"si")

    'response.write sImage & " sImage " & chr(13)
    'response.write sImagenValor & " sImagenValor " & chr(13)
    'response.write sImagenCampo & " sImagenCampo " & chr(13)


     Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    'Perfil
     Celda = ofv.gencelda("","tt","","","","","","Perfil","si")


    'Perfil
     Sql1 = "Select Tipoperfil From Comperfil Order By descri"
     sInput = ofv.GenerarCombo(sForm,rs(4).name,rs(4).value,Sql1,cn,0,0,rs(4).name)
     Celda = Celda & ofv.gencelda("","vti ","","","","","",sInput,"si")

     Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=Tipocomp&volver=" & vuelta
    end if
    Formulario = ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

  end if

  sHTML = sHTML & ofv.GenTabla("","aaa","","80%","","0","","0","0","","",Formulario,"si")

  'Botonera
  Agregar = ""
  Buscar = ""

  sAccion = "TiposdeComponentes.asp?Cantregmover=" & Cantregmoverx
  Celda = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>