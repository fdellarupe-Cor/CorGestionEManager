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

 '' session("MODORIGEN") = "TIPOPERFIL"
  session("MODORIGEN") = "SECTORPERFIL"

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
    session("elNom") = Nom
  end if

  if ID = "" then
    ID = session("elID")
  else
    session("elID") =cdbl(ID)
  end if

  sHTML = ofv.FormHeader("PERFILES DEL SECTOR")

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar


  sInput = "Sector:&nbsp;" & NOM

  Celda = ofv.gencelda("","ttt colspan=3","","","","","",sInput,"si")

  Formulario = ofv.genrow("","","","","","","",Celda,"si")

  Formulario = Formulario & ofv.Encabezados("perfil","dotacion","FUNCIONES")

  set cn = ofv.conectar(ofv.strconn2)

  Sql1 = "Select * From TIPOPERFIL  order by descri"

  Strsql = "Select Ecoperfil.*,Tipoperfil.descri,tipoperfil.estado From Ecoperfil, Tipoperfil where "
  Strsql = Strsql & "Ecoperfil.codperfil = Tipoperfil.tipoperfil and codsector=" & ID
  	'response.write strsql
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if Not rs.EOF then rs.move(CantRegMoverx)
  posI = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)

    sForm = "form" & posi

    'Boton Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=Ecoperfil&volver=" & vuelta & "&Eliminar=NSNC"
      'Celda = ofv.BotonEliminar("Dependencia",sHref)
	  Celda = ofv.BotonEliminar(rs(4),sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if

    'Perfil
    if rs("estado")= 3 then
    sInput = ofv.GenerarCombo(sForm,rs(2).name,rs(2).value,sql1,cn,0,1,"")
    else
    sInput = rs("descri")
    end if
    Celda = Celda & ofv.gencelda("","ut ","","60%","","","",sInput,"si")

    'Dotacion
    if rs("estado") = 3 then
    sInput = ofv.GenerarInput(rs(3).name,rs(3),"text","form" & posi,"30","dtn","SI")
    else
    sinput=rs(3)
    end if

    Celda = Celda & ofv.gencelda("","ut ","","40%","","","",sInput,"si")

    'FUNCIONES
    'sImagenP = "&#0232;"
    'sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" "
    sinput = "<a href='SECTORPERFILESFUN.asp?PERFIL=" & rs(0)
    sinput = sinput  & "' title='FUNCIONES DEL PUESTO' ><i class='las la-arrow-right'></i></a>"
     if rs("estado") <> 3 then
      sinput="&nbsp;"
     end if
    Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")


    Fila = ofv.genrow("","","","","","","",Celda,"si")


    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=Ecoperfil&volver=" & vuelta
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
     sAccion = "AgregarSectorPerfiles.asp?Regisvuelta=" & CantRegMoverx & "&tabla=Ecoperfil&NOM=" & NOMU & "&ID=" & ID
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