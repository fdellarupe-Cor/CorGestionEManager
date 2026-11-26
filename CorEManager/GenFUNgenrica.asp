<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrarx = 12
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  sHTML = ofv.FormHeader(session("FormN"))

  'Encabezados
  Formulario = ofv.Encabezados("DESCRI")


  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar
  nivel=request.querystring("nivel")
  set cn = ofv.conectar(ofv.strconn2)

  Strsql = "Select * From ECOPERFILFUN where tipo='GEN' AND ecoperfil=0" & nivel & " Order BY ptodescri"
  RESPONSE.WRITE STRSQL
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if not rs.eof then rs.move(CantRegMoverx)
  posi = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)
    sForm = "form" & posi

    sInput = rs(0)

    'Eliminar
    if bSuprimir and sinput > "0" then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=ECOPERFILFUN&volver=" & vuelta & "&Eliminar=NSNC"
      Celda = ofv.BotonEliminar(rs(1),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

     'DESCRI
	    sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"100","dt","")
       Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")



       Fila = ofv.GenRow("","","","","","","",Celda,"si")

    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=ECOPERFILFUN&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",Fila,"si")

    posi = posi + 1
    rs.Movenext
  Loop
  sHTML = sHTML & ofv.GenTabla("","","","80%","","0","","0","0","","",Formulario,"si")

  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "AgregarfunGen.asp?Regisvuelta=" & CantRegMoverx & "&Opcion=si&tabla=ECOPERFILFUN&NIVEL=" & NIVEL
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  str1 = "Select * From ECOPERFILFUN where tipo='GEN' AND ecoperfil=0" & nivel & "Order BY ptodescri"
  str2 = "Select * From ECOPERFILFUN where tipo='GEN' AND ecoperfil=0" & nivel & "Order BY ptodescri"
  Buscar = ofv.Combobuscar(str1,str2,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

  'consulta
  sinput = "<a  href=NivelPUESTOCons.asp?"&request.querystring & " target=_new ><b>consulta</b></a>"
  Celda = ""'ofv.gencelda("","tt style='border-style:outset;border-width:2;' ","center","","15%","","",sInput,"si")

  sAccion = "NivelePUESTOCons.asp?" & request.querystring
  Celda = ""'ofv.BotonConsultar(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>