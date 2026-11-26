<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrarx = 10
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  WKRSKCAT = request.querystring("CATEGO") & ""
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)


  IF WKRSKCAT = "" THEN
     WKRSKCAT = SESSION("RSKCAT")
  ELSE
     SESSION("RSKCAT") = WKRSKCAT
  END IF


  sHTML = ofv.FormHeader(session("FormN"))




  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)


  Strsql = "Select * From RSKCATEGO WHERE ID = " & WKRSKCAT
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then 

     SESSION("RSKSIT") = RS(1)
     SNOMEMPRE2 = RS(2)

  END IF


  Strsql = "Select * From RSKSITUACION WHERE ID = " & SESSION("RSKSIT")
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then 

     SNOMEMPRE = RS(1)

  END IF

    sImagen4 = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=2>4</FONT></B>"
    sImagenP = "<B><FONT FACE=corflaticon COLOR=#aa0000 SIZE=2>&#0232;</FONT></B>"

  Celda = ofv.gencelda("","ttt","","","","","","<BIG>" & sImagen4 & " " & SNOMEMPRE & " / " & SNOMEMPRE2 & "</BIG>","si")
  Fila = ofv.genrow("","","","","","","",Celda,"si")
  sHTML = sHTML & ofv.GenTabla("","","","100% ","","0","","0","0","","",Fila,"si") & "<BR>"



  'Encabezados
  Formulario = ofv.Encabezados("Id","Descripcion")

  Strsql = "Select * From RSKEVENTO WHERE CATEGO = '" & WKRSKCAT & "' Order By ID"
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    sInput = rs(0)

    'Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=RSKEVENTO&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(1),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'Segmento
    sInput = rs(0)
    Celda = Celda & ofv.gencelda("","cc","right","50","","","",sInput,"si")

    'Descripcion
    sInput = ofv.GenerarInput(rs(3).name,rs(3),"text",sForm,"100","dt","")
    Celda = Celda & ofv.gencelda("","","","","","","",sInput,"si")

    Celda = Celda & ofv.gencelda("","","","","","","","","si")
    Fila = ofv.genrow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=RSKEVENTO&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")

  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "AgregarEVENTORSK.asp?Regisvuelta=" & CantRegMoverx & "&Opcion=si&tabla=RSKEVENTO"
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select descri From RSKEVENTO WHERE CATEGO = '" & WKRSKCAT & "' Order By id "    ' orden pagina
  str2 = "Select descri From RSKEVENTO WHERE CATEGO = '" & WKRSKCAT & "' Order By descri "    ' orden alfabetico
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))
  
  'consulta

  sAccion = "EVENTOCons.asp?" & request.querystring
  Celda = ofv.BotonConsultar(sAccion)

  
  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>