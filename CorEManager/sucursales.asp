<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0
  Svalida=request.querystring("Svalida")
  CantRegMostrarx = 10
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  sUN = request.querystring("UN")

  IF sUN <> "" THEN

     SESSION("SUN") = SUN

  ELSE

    SUN = SESSION("SUN")

  END IF

  sHTML = ofv.FormHeader(session("FormN"))

  if SESSION("APLCHK") <> "S" then xxxx = APLCHECK()



  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From CLIENTEPLAN  WHERE ID =0" & SUN
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then

     SNOMEMPRE = RS(1)

  END IF

    sImagen4 = "<B><FONT FACE=CorFlaticon COLOR=#aa0000 SIZE=2>4</FONT></B>"
    sImagenP = "<B><FONT FACE=CorFlaticon COLOR=#aa0000 SIZE=2>&#0232;</FONT></B>"

  Celda = ofv.gencelda("","ttt","","","","","","<BIG>" & sImagen4 & " " & SNOMEMPRE & "</BIG>","si")
  Fila = ofv.genrow("","","","","","","",Celda,"si")
  sHTML = sHTML & ofv.GenTabla("","","","100% ","","0","","0","0","","",Fila,"si") & "<BR>"

  'Encabezados


   Formulario = ofv.Encabezados("Id","codigo interno","Descripcion de la sucursal")


   Sordenar = " Order By CAST(CODINTER AS INT) "


   Strsql = "Select * From SUCURSAL " & Sordenar

  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
   if Not rs.EOF then   rs.move(CantRegMoverx)
  posI = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    'Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=sucursal&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(0),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'id
    sInput = rs(0)
    Celda = Celda & ofv.GenCelda("","cc","center","5","","","",sInput,"si")
    SINPUT="<input id='"& rs(1).name   &"' name='"& rs(1).name   &"'  value='" & trim(rs(1)&"") & "' type=text MAXLENGTH=3 class=dtn onblur="&chr(34)&"validarnum(this,'"&sForm&"','"&trim(rs(1)&"")&"',3);"&chr(34) & ">"
    Celda = Celda & ofv.GenCelda("","cc","center","5","","","",sInput,"si")


   'Descripcion
    sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"100","dtbs","")

    Celda = Celda & ofv.gencelda("","vti","","","","","",sInput ,"si")



    'Propiedades


    Fila = ofv.genrow("","","","","","","",Celda,"si")



    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=sucursal&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")


    posi = posi + 1
    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","-1","1","","",Formulario,"si")



  str1 = "Select descripcion From sucursal "  & Sordenar
  str2=str1
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

   'Botonera
    Agregar = ""
    if bAgregar then
      sAccion = "AgregarSUCURSAL.asp?Regisvuelta=" & CantRegMoverx & "&Opcion=si&tabla=sucursal"
      Agregar = ofv.BotonAgregar(sAccion)
  end if

  'consulta

  sAccion = "SucursalCons.asp?" & request.querystring
  Celda = ofv.BotonConsultar(sAccion)


  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","",Agregar,Buscar,Listar,celda & retornar)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)
  shtml=SHTML & 	ValidarInput()

 ' if Svalida="SI" THEN
' sHTML = sHTML & "<SCRIPT LANGUAJE=JAVASCRIPT> alert('Codigo Repetido');</script>"
'  END IF
  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>