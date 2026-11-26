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
    session("elNom") = Nom
  end if

  if ID = "" then
    ID = session("elID")
  else
    session("elID") =cdbl(ID)
  end if

  sHTML = ofv.FormHeader("SITES/SUCURSALES")

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

  'Encabezados

  sInput = "Sites/sucursales de:&nbsp;" & NOM

  Celda = ofv.gencelda("","ttt colspan=2","","","","","",sInput,"si")

  Formulario = ofv.genrow("","","","","","","",Celda,"si")

  Formulario = Formulario & ofv.Encabezados("site/sucursal")

  set cn = ofv.conectar(ofv.strconn2)

  Sql1 = "Select * From SITES  where siteun = " & session("sun") & " order by descri"
'response.write sql1
  Strsql = "Select EcoSITE.*,SITES.descri From EcoSITE, SITES where "
  Strsql = Strsql & "EcoSITE.ECOSITE = SITES.OBJID and ECO=" & ID
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  if Not rs.EOF then rs.move(CantRegMoverx)
  posI = 0
  Do While (Not rs.EOF) and (posi < CantRegMostrarx)

    sForm = "form" & posi

    'Boton Eliminar
    if bSuprimir then
      sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
      sHref = sHref & "&Tabla=EcoSITE&volver=" & vuelta & "&Eliminar=NSNC"

      SqlSector = "Select * From SITES  where siteun = " & session("sun") & " and objid = 0" &  rs(2)
      'response.write SqlSector
	  set rsSector = ofv.crearconsultaEx(SqlSector,cn,1,parametros)
	  valorCadena = ""
	  if not rsSector.eof then
		valorCadena = "Sites/sucursales " + rsSector(1)
	  end if
	  Celda = ofv.BotonEliminar(valorCadena,sHref)
    else
      Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
    end if


    'Perfil
    sInput = ofv.GenerarCombo(sForm,rs(2).name,rs(2).value,sql1,cn,0,1,"")
    Celda = Celda & ofv.gencelda("","ut","","60%","","","",sInput,"si")


    Fila = ofv.genrow("","","","","","","",Celda,"si")


    sAccion = ""
    if bModificar then
      sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
      sAccion = sAccion & "&tabla=EcoSITE&volver=" & vuelta
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


     sAccion = "AgregarSectorSITES.asp?Regisvuelta=" & CantRegMoverx & "&tabla=EcoSITE&NOM=" & NOMU & "&ID=" & ID
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