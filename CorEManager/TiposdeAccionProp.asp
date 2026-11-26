<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrarx = cdbl(request.querystring("CantRegMostrar"))
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))

  Wperfil = request.querystring("Wperfil")

  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_wperfil=" & Wperfil

  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados



  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From Modo where tipomodo = '" & Wperfil & "' "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then
  posi = 0

    sForm = "form" & posi


    'Tipo
    Celda = ofv.gencelda("","tt","","10%","","","","Tipo","si")

    sInput = rs(0) & " - " & rs(1)
    Celda = Celda & ofv.gencelda("","vtig style='padding-left:5px;'","","80%","","","",sInput,"si")

    Fila = ofv.genrow("","","","","","","",Celda,"si")


    'Tipo
    Celda = ofv.gencelda("","ttt colspan=2 ","","","","","","ICONO PANTALLA","si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    Celda = ofv.gencelda("","tt","","","","","","Archivo","si")
    'Pantalla
    sImagenValor = rs(2)
    sImagenCampo = rs(2).name

    Xfiles=ofv.Getfiles("Iconos")

    sInput = ofv.GenerarCombo3(sForm,XFILES,sImagenCampo,sImagenValor)
    Celda = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    Celda = ofv.gencelda("","tt","","","","","","Imagen","si")
    'Icono small
    sImage = ofv.GenImage("Imagen","","","","Iconos\" & sImagenValor,"","")
    Celda = Celda & ofv.gencelda("","fti","","","","","",sImage,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    'Descripcion
    Celda = ofv.gencelda("","ttt colspan=2 ","","","","","","ICONO CURSOGRAMA","si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    Celda = ofv.gencelda("","tt","","","","","","Archivo","si")
    'Cursograma
    sImagenValorc = rs(10)
    sImagenCampoc = rs(10).name
   ' Xfiles=ofv.Getfiles("IconosCursograma")
   ' XFILES(0)="DEFAULT"
    sfiles=split(getfilesinfolder("IconosCursograma",""),"|")
      scomK = "DEFAULT" & chr(9)
      scomD = "DEFAULT" & chr(9)
      for f=0 to ubound(sfiles)-1
       fileT=split(sfiles(f),"###")

       scomK =  scomK & FILET(0) & CHR(9)
       scomD =  scomD & FILET(0) & CHR(9)

     NEXT

    'sInput = ofv.GenerarCombo3(sForm,XFILES,sImagenCampoc,sImagenValorc)
     sInput = ofv.gencomboFX(sForm,rs(10).name,rs(10),scomK,scomD," ")
    Celda = Celda & ofv.gencelda("","vti","","","","","",sInput,"si")
    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")

    Celda = ofv.gencelda("","tt","","","","","","imagen","si")
    'Icono large
    IF sImagenValorc <> "DEFAULT" ThEN
    sImage = ofv.GenImage("Imagen","","","","IconosCursograma\" & sImagenValorc,"","")
    ELSE
    sImage="&nbsp;"
    END IF
    Celda = Celda & ofv.gencelda("","fti","","","","","",sImage,"si")

    Fila = Fila & ofv.genrow("","","","","","","",Celda,"si")



    'Tipos
    Celda = ofv.gencelda("","ttt colspan=2 ","","","","","","OPCIONES","si")
    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


    'Control
    Celda = ofv.gencelda("","tt","","","","","","Control","si")

    sInput = ofv.combosino(sForm,rs(4),rs(4).name)
    Celda = celda & ofv.GenCelda("","fti","","","","","",sInput,"si")

    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


    'Documento
    Celda = ofv.gencelda("","tt","","","","","","Documento","si")

    sInput = ofv.combosino(sForm,rs(5),rs(5).name)
    Celda = celda & ofv.GenCelda("","fti","","","","","",sInput,"si")

    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


    'Link
    Celda = ofv.gencelda("","tt","","","","","","Links","si")

    sInput = ofv.combosino(sForm,rs(6),rs(6).name)
    Celda = celda & ofv.GenCelda("","fti","","","","","",sInput,"si")

    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


    'Fin
    Celda = ofv.gencelda("","tt","","","","","","Fin/Continuador","si")

    sInput = ofv.combosino(sForm,rs(7),rs(7).name)
    Celda = celda & ofv.GenCelda("","fti","","","","","",sInput,"si")

    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")


    'componentes
    Celda = ofv.gencelda("","tt","","","","","","Componentes","si")

    sInput = ofv.combosino(sForm,rs(8),rs(8).name)
    Celda = celda & ofv.GenCelda("","fti","","","","","",sInput,"si")


    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")




    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=Modo&volver=" & vuelta
    end if
    Formulario = ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa","","80%","","0","","-1","2","","",Formulario,"si")

    posi = posi + 1

  end if


  Agregar = ""
  Buscar  = ""


  sAccion = "TiposdeAccion.asp?Cantregmover=" & Cantregmoverx
  Celda = ofv.BotonVolver(sAccion)


  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)





  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>