<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0

  CantRegMostrarx = 6
  CantRegMoverx = cdbl(request.querystring("CantRegMover"))
  Regisvuelta = request.querystring("Regisvuelta")
  Regisperfil = request.querystring("Regisperfil")
  NOM = request.querystring("NOM")
  NOMU = request.querystring("NOM")
  NOM = ofv.convertircar(NOM,"*","&nbsp;")
  ID = request.querystring("ID")
  vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

  if Regisvuelta = "" then
    Regisvuelta = session("Regisvuel")
    Regisperfil = session("Regisperfil")
  else
    session("Regisvuel") = cdbl(Regisvuelta)
    session("Regisperfil") = cdbl(Regisperfil)
  end if

  if NOMU = "" then
    NOM = session("elNom")
    NOM = ofv.convertircar(NOM,"*","&nbsp;")
    NOMU = session("elNom")
  else
    session("elNom") = NOMU
  end if

  if ID = "" then
    ID = session("elID")
  else
    session("elID") =cdbl(ID)
  end if

  sHTML = ofv.FormHeader("DEPENDENCIA")

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

     simagend3 = "<B><FONT FACE=CorFlaticon COLOR=#005500 SIZE=2>&#0232;</FONT></B>"


  sInput = "Sector&nbsp;Dependiente:&nbsp;" & NOM

  Celda = ofv.gencelda("","ttt colspan=3","","","","","",sInput,"si")

  Formulario = ofv.genrow("","","","","20","","",Celda,"si")

  Formulario = Formulario & ofv.Encabezados("Sector Dependencia","site/sucursal")

  set cn = ofv.conectar(ofv.strconn2)

  Sql1 = "Select * From Eco order by descri"

  Strsql = "Select Ecorel.*,Eco.descri From Ecorel, Eco where eco.estado=3 and "
  Strsql = Strsql & "Ecorel.ecomaster = Eco.eco and ecodepend=" & ID
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if Not rs.EOF then rs.move(CantRegMoverx)
  posI = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)

    sForm = "form" & posi

    'Boton Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=Ecorel&volver=" & vuelta & "&Eliminar=NSNC"
	  valorCadena = ""
	  valorCadena = "Sector Dependencia " + rs(3)
	  Celda = ofv.BotonEliminar(valorCadena,sHref)

    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'Dependencia
    sInput = ofv.GenerarCombo(sForm,rs(1).name,rs(1).value,sql1,cn,0,1,"")
    Celda = Celda & ofv.gencelda("","ut ","","60%","","","",sInput,"si")

    sAccion  = "SectorDEPENSITES.asp?ID=" & rs(0)

    'sImagenP = "&#0232;"
    'sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
    sinput   = "<a " & sFuncAnc & " href='" & sAccion
    sinput   = sinput  & "' title='ACTUALIZAR SITE/SUCURSAL' ><i class='las la-arrow-right'></i></a>"
    Celda    = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")

    Fila = ofv.genrow("","","","","","","",Celda,"si")


    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=Ecorel&volver=" & vuelta
    end if
    Formulario = Formulario & ofv.genform(sForm,"",sAccion,"","post","",Fila,"si")


    posi = posi + 1
    rs.Movenext
  Loop

  sHTML = sHTML & ofv.GenTabla("","aaa","","","","0","","-1","1","","",Formulario,"si")


  'Botonera
  Agregar = ""
  Buscar = ""

  if bAgregar then
     sAccion = "AgregarSectorDepend.asp?Regisvuelta=" & CantRegMoverx & "&tabla=Ecorel&NOM=" & NOMU & "&ID=" & ID
     Agregar = ofv.BotonAgregar(sAccion)

  end if

  sAccion = "SectoresProp.asp?CantRegMover=" & regisvuelta   & "&Wperfil=" & regisperfil
  retornar = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",agregar,retornar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>