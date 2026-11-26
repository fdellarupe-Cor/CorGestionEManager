<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrarx = 0
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))

  Wperfil = request.querystring("Wperfil")

  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx) & "_wperfil=" & Wperfil


  set cn = ofv.conectar(ofv.strconn2)


  sHTML = ofv.FormHeader(session("FormN"))



  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar



  Strsql = "Select * From Tipoaccalt where tipoaccionalt = " & Wperfil

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if Not rs.EOF then
  posI = 0

    sForm = "form" & posi
  
    'Descripcion
    sInput = rs(1)
    Celda = ofv.GenCelda("","ttt colspan=2 ","","","","","",sInput,"si")

    Fila = ofv.genrow("","","","","","","",Celda,"si")

    'Modo

    sInput = "tipo de accion"
    Celda = ofv.GenCelda("","tt","","20%","","","",sInput,"si")

    Strsql = "Select * From Modo Order By descri"
    sInput = ofv.GenerarCombo(sForm,rs(2).name,rs(2).value,Strsql,cn,0,1,"")
    Celda = Celda & ofv.GenCelda("","vti","","80%","","","",sInput,"si")

    Strsql1 = "Select archivo, descri From Modo where "
    Strsql1 = Strsql1 & " tipomodo = '" & RS(2) & "' "

    set rs1 = ofv.crearconsultaEx(StrSql1,cn,1,parametros)
    if Not rs1.EOF then 
       XSMODOd = RS1(1)
       XSMODOI = ofv.GenImage("","","","","Iconos/" & rs1(0),"","")
    ELSE
       XSMODOd = "S/D"
       XSMODOI = "&nbsp;"
    END IF

    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

    'Icono

    sInput = "icono"
    Celda = ofv.GenCelda("","tt","","","","","",sInput,"si")

    sImagen = XSMODOI
    Celda = Celda & ofv.GenCelda("","fti","","","","","",sImagen,"si")


    Fila = fila & ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=Tipoaccalt&volver=" & vuelta
    end if
    Formulario = ofv.GenForm(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
 
  end if
  sHTML = sHTML & ofv.GenTabla("","aaa","","80%","","0","0","0","0","","",Formulario,"si")

  'Botonera
  Agregar = ""
  Buscar  = ""


  sAccion = "Acciones.asp?Cantregmover=" & Cantregmoverx
  Celda = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>