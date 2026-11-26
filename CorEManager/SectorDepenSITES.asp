<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
  uv = ofv.uv
  if uv = "NO" then ok = ofv.ValidarUsuario
  sHTML = ok
else
  response.expires=0


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

  if ID2 = "" then
    ID2 = session("elID")
  else
    session("elID") =cdbl(ID2)
  end if

  if ID = "" then
    ID = session("elIDS")
  else
    session("elIDS") =cdbl(ID)
  end if

  sHTML = ofv.FormHeader("DEPENDENCIA")

  ofv.ObtenerAtributos Session("Form"),""
  bAgregar = ofv.AAgregar
  bModificar = ofv.AModificar
  bSuprimir = ofv.AEliminar
  bConsultar = ofv.AConsultar

     simagend3 = "<B><FONT FACE=CorFlaticon COLOR=#005500 SIZE=2>&#0232;</FONT></B>"

  set cn = ofv.conectar(ofv.strconn2)

  NOMDEP = "S/D"

  Strsql = "Select DESCRI From Ecorel, ECO where CODREL = " & session("elIDS") & " AND ECOMASTER = ECO "
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  IF NOT RS.EOF THEN NOMDEP = RS(0)


  sHTML = sHTML & "<table width=60% class=aaa style='border-radius: 10px 10px 0px 0px;border-bottom:none;' marginwidth=0 cellspacing=" & ofv.wspacing
  sHTML = sHTML & " cellpadding=" & 1 + ofv.wspacing & "  > "
  sHTML = sHTML & "<tr><td  class=tt WIDTH=20% >"
  sHTML = sHTML & "Sector&nbsp;Dependiente:</td>"
  sHTML = sHTML & "<td  class=ttt >"
  sHTML = sHTML & NOM & "</td></tr>"
  sHTML = sHTML & "<tr><td  class=tt WIDTH=20% >"
  sHTML = sHTML & "DEPENDENCIA:</td>"
  sHTML = sHTML & "<td  class=ttt >"
  sHTML = sHTML & NOMDEP & "</td></tr>"
  sHTML = sHTML & "</table>"

  sHTML = sHTML & "<table width=60% class=aaa style='border-radius: 0px 0px 10px 10px;' marginwidth=0 cellspacing=0"
  sHTML = sHTML & " cellpadding=1 >"
  sHTML = sHTML & "<tr height=20 ><td width=1% class=tt>&nbsp;</td><td class=tt>SITE/SUCURSAL</td>"
  sHTML = sHTML & "</tr>"



  Sql1 = "Select * From SITES where SITEun = " & session("sun") & " order by descri"

  Strsql = "Select * From EcorelSITE where CODREL = " & session("elIDS")
  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  posI = 0
  IF RS.EOF THEN
     STRSQL = "INSERT INTO ECORELSITE (CODREL,SITE) VALUES(" & session("elIDS") & ",1) "
     CALL ofv.crearconsultaEx(StrSql,cn,1,parametros)

     Strsql = "Select * From EcorelSITE where CODREL = " & session("elIDS")
     set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
  END IF
  Do While (Not rs.EOF)
    if bModificar then
      sHTML = sHTML & "<form name=form" & posi & " action=asp/actualizar.asp?fname="
      sHTML = sHTML & rs(0) & "&tabla=EcorelSITE&volver=" & vuelta & " method=post>"
    else
      sHTML = sHTML & "<form name=form" & posi & " method=post>"
    end if

    sHTML = sHTML & "<tr><td>"
    'Boton Eliminar
    if bSuprimir then
      sHelp = ""
     ''' sHelp = rs(3).value

      sHTML = sHTML & "<A HREF=asp/EliminarRegistro.asp?ID=" & rs(0) & "&Tabla=EcorelSITE&volver="

	  SqlDependencia = "Select * From SITES where SITEun = " & session("sun") & " and objid = 0" & rs(2)
	  'response.write (SqlDependencia)
	  set rsDependencia = ofv.crearconsultaEx(SqlDependencia,cn,1,parametros)

	  if not rsDependencia.eof then
		sHelp = "Dependencia " + rsDependencia(1)
	  end if

	  sHelp = ofv.convertircar(sHelp," ","&nbsp;")

      sHTML = sHTML & vuelta & "&Eliminar=NSNC&Nombre=" & sHelp & " Title=Eliminar&nbsp;"

      sHTML = sHTML & sHelp & "><img src=imagenes/botdel.jpg border=0 height=15 width=15></a></td>"
    else
      sHTML = sHTML & "&nbsp;</td>"
    end if

    'Dependencia
    sHTML = sHTML & "<td width=10% >"
    sID = "form" & posi
    sHTML = sHTML & ofv.GenerarCombo(sID,rs(2).name,rs(2).value,sql1,cn,0,1,"xx")
    sHTML = sHTML & "</td>"



    sHTML = sHTML & "</tr></form>"

    posi = posi + 1
    rs.Movenext
  Loop
  sHTML = sHTML & "</table><br><br>"

  'Botones para agregar y volver

  'Botonera
  Agregar = ""
  if bAgregar then
    sAccion = "AgregarDepenSITE.asp?tabla=EcorelSITE&ID=" & ID & "&NOM=" & NOMU
    Agregar = ofv.BotonAgregar(sAccion)
  end if

  sAccion = "SectorDEPENDENCIAS.asp"
  retornar = ofv.BotonVolver(sAccion)

  sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",agregar,retornar,"")

  call ofv.cerrarconsulta(rs)
  call ofv.cerrarconn(cn)

  sHTML = sHTML & "</body></html>"
end if

Response.Write sHTML

%>