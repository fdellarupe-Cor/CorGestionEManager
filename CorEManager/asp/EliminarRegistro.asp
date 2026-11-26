<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

response.expires=0

Tabla = request.querystring("Tabla")
nCampo = request.querystring("Campo")
If nCampo = "" Then nCampo = 0
nCampo = CDbl(nCampo)
sID = request.querystring("ID")
sID = ofv.convertircar(sID,"_"," ")
sEliminar = request.querystring("Eliminar")
sNombre = request.querystring("Nombre")
Volver = request.querystring("Volver")
Volver = ofv.convertircar(Volver,"_","&")
If sEliminar = "NSNC" Then session("volver2") = volver

On Error Resume Next

Set cn = ofv.conectar(ofv.strconn2)

sMsg = ""
'Obtengo tipo de Clave
strsql = "Select * From " & tabla
Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
clave = rs(nCampo).name
s = ofv.QueEsx(rs(nCampo).type)
If rs(nCampo).type = 200 Then s = "'"
sID2 = s & ofv.convertircar(sID,"_"," ") & s
Call ofv.cerrarconsulta(rs)

'Validaciones
'response.Write(Tabla)
Select Case UCase(Tabla)
	Case UCase("Tipoperfil")
	strsql = "select * from ECOPERFIL Where codperfil=" & sID2

	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)

	If sMsg <> "NO" Then
		strsql = "select * from OBJPROCED Where codperfil=" & sID2
		Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If Not rs.eof Then sMsg = "NO"
		Call ofv.cerrarconsulta(rs)

		If sMsg <> "NO" Then
			strsql = "select * from OBJACCION Where perfilcont=" & sID2
			Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
			If Not rs.eof Then sMsg = "NO"
			Call ofv.cerrarconsulta(rs)

			If sMsg <> "NO" Then
				strsql = "select * from OBJcontrol Where codperfil=" & sID2
				Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
				If Not rs.eof Then sMsg = "NO"
				Call ofv.cerrarconsulta(rs)
			End If
		End If
	End If
	Case UCase("Tipoaccalt")
	strsql = "Select * From OBJACCION Where tipoaccion=" & sID2
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)
	If sMsg <> "NO" Then
		strsql = "Select * From OBJACCIONcontrol Where accion=" & sID2
		Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If Not rs.eof Then sMsg = "NO"
		Call ofv.cerrarconsulta(rs)
	End If
	Case UCase("Sisgen")
	strsql = "Select * From COMPPRTY Where sistorigen=" & sID2
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)
	Case UCase("Tipogen")
	strsql = "Select * From COMPPRTY Where tipogen=" & sID2
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)
	Case UCase("Origen")
	strsql = "Select * From OBJPRTY Where origen="  & sID2
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)
	Case UCase("Instrumen")
	strsql = "Select * From OBJPRTY Where instrumen="  & sID2
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)

	Case UCase("TIPOPRODUCTO")
	strsql = "Select * From objetos Where tsegmento="  & sID2
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)

	Case UCase("TipoSegmento")
	strsql = "Select * From objtsegagr Where tsegmento="  & sID2
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)

	Case UCase("TipoAgrupa")
	strsql = "Select * From objtsegagr Where tagrupa="  & sID2
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)

	Case UCase("Objcuenta")
	'RESPONSE.Write(CHR(13) & CHR(13))
	'RESPONSE.Write(session("APLMDLOR") & CHR(13) & sID2)
	If session("APLMDLOR") <> "BE" Then
		strsql = "Select DISTINCT OBJID From OBJESQUEMA2 Where codcuenta=" & sID2 & Chr(13)
		STRSQL = STRSQL & " UNION ALL " & Chr(13)
		STRSQL = STRSQL & " SELECT DISTINCT OBJID FROM ORTRACER WHERE TRCCTACBLE IN (SELECT NUMERO FROM ObjCuenta WHERE OBJID = " & sID2 & ") " & Chr(13)
		Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If Not rs.eof Then sMsg = "NO"
		Call ofv.cerrarconsulta(rs)
	Else
		STRSQL = " SELECT DISTINCT OBJID FROM ORTRACER WHERE TRCCTACBLE IN (SELECT NUMERO FROM ObjCuenta WHERE OBJID = " & sID2 & ") " & Chr(13)
		Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If Not rs.eof Then sMsg = "NO"
		Call ofv.cerrarconsulta(rs)
	End If
	Case UCase("Niveljerar")
	strsql = "Select * From ECO Where econiv=" & sID2
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)
	Case UCase("Segmento")
	strsql = "Select * From ECO Where ecoseg=" & sID2
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)
	Case UCase("Jerarquia")
	strsql = "Select * From ECO Where ecojerar=" & sID2
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)
	Case UCase("Eco")

	If session("APLMDLOR") <> "BE"     Then
		strsql = "Select * From ECOPERFIL Where codsector=" & sID2
		Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If Not rs.eof Then sMsg = "NO"
		Call ofv.cerrarconsulta(rs)

		If sMsg <> "NO" Then
			strsql = "Select * From ECOREL Where ecodepend=" & sID2
			Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
			If Not rs.eof Then sMsg = "NO"
			Call ofv.cerrarconsulta(rs)
		End If

		If sMsg <> "NO" Then
			strsql = "Select * From ECOREL Where ecomaster=" & sID2
			Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
			If Not rs.eof Then sMsg = "NO"
			Call ofv.cerrarconsulta(rs)
		End If

		If sMsg <> "NO" Then
			strsql = "Select * From OBJSECTOR Where codsector=" & sID2
			Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
			If Not rs.eof Then sMsg = "NO"
			Call ofv.cerrarconsulta(rs)
		End If
	Else
		STRSQL = " SELECT DISTINCT EcoSITE.OBJID FROM EcoSITE, SITES WHERE EcoSITE.ECOSITE = SITES.OBJID AND ECO = " & sID & Chr(13)
		STRSQL = STRSQL & " UNION ALL " & Chr(13)
		STRSQL = STRSQL & " SELECT DISTINCT Ecorel.CODREL AS OBJID FROM Ecorel, Eco WHERE Ecorel.ecomaster = Eco.eco AND ecodepend = " & sID & Chr(13)
		STRSQL = STRSQL & " UNION ALL " & Chr(13)
		STRSQL = STRSQL & " SELECT DISTINCT Ecoperfil.OBJID FROM Ecoperfil, Tipoperfil WHERE Ecoperfil.codperfil = Tipoperfil.tipoperfil AND codsector = " & sID & Chr(13)
		STRSQL = STRSQL & " UNION ALL " & Chr(13)
		STRSQL = STRSQL & " SELECT DISTINCT OBJID FROM EcoDESCRIPCION WHERE ECO = " & sID & Chr(13)
		STRSQL = STRSQL & " UNION ALL " & Chr(13)
		STRSQL = STRSQL & " SELECT DISTINCT OBJID FROM ORTRACER WHERE TRCECORESP = " & sID & Chr(13)
		STRSQL = STRSQL & " UNION ALL " & Chr(13)
		STRSQL = STRSQL & " SELECT DISTINCT CODSECTOR AS OBJID FROM OBJSECTOR WHERE CODSECTOR = " & sID & Chr(13)

		Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If Not rs.eof Then sMsg = "NO"
		Call ofv.cerrarconsulta(rs)

	End If
	Case UCase("Objperfil")
	strsql = "Select * From TIPOBJETO Where perfil=" & sID2
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)

	'    if sMsg <> "NO" then
	'      strsql = "Select * From SUBTEMA Where perfil=" & sID2
	'      set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	'      If not rs.eof then sMsg = "NO"
	'      call ofv.cerrarconsulta(rs)
	'    end if
	Case UCase("Subtema")
	strsql = "Select * From PRODSERV Where marcaps=" & sID2
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)

	If sMsg <> "NO" Then
		strsql = "Select * From RELAOBJ Where tipobjmaster=" & sID2
		Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If Not rs.eof Then sMsg = "NO"
		Call ofv.cerrarconsulta(rs)
	End If
	Case UCase("Tipobjeto")
	strsql = "Select * From OBJETOS Where tipobjeto=" & sID2
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)
	Case UCase("Comperfil")
	strsql = "Select * From TIPOCOMP Where perfil=" & sID2
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)
	Case UCase("Tipocomp")
	strsql = "Select * From COMPLEMENTOS Where tipocomp=" & sID2
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)
	If sMsg <> "NO" Then
		strsql = "Select * From OBJACCION Where tipocomp=" & sID2
		Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If Not rs.eof Then sMsg = "NO"
		Call ofv.cerrarconsulta(rs)
	End If
	Case UCase("tabperfil")
	strsql = "Select * From objtabdef Where tabperfil=" & sID2
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)

	Case UCase("Modo")
	strsql = "Select * From OBJACCION Where modo=" & sID2
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)
	If sMsg <> "NO" Then
		strsql = "Select * From TIPOACCALT Where modo=" & sID2
		Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	End If

	Case UCase ("REGIMENFIRMAS")
	strsql = "Select objid From OBJLNK Where CODOBJETO=" & sID2 & " AND TIPOBJETO='REGIFX' AND IDEN2='REGFIRMAS'"
	strsql ="union "
	strsql = "Select objid From ECOPERFILFIR Where PTODSCITEM='" & sID2 &"'"
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)


	Case "TABPROCESOS01"
	strsql = "Select * From TABPROCESOS01 Where objid=" & sID2
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	idelim=rs(1)
	PROCNIVEL=rs("nivel") +1
	Call ofv.cerrarconsulta(rs)
	Strsql = "Select * From TABPROCESOS01 "
	strsql = strsql & " where "

	If procnivel > 1 Then

		If PROCNIVEL = 2 Then
			strsql = strsql & " codproceso01 = '" & idelim & "' "
		ElseIf PROCNIVEL = 3 Then
			strsql = strsql & " codproceso02 = '" & idelim & "' "
		ElseIf PROCNIVEL = 4 Then
			strsql = strsql & " codproceso03 = '" & idelim & "' "
		Else
			strsql = strsql & " codproceso04 = '" & idelim & "' "
		End If

		strsql = strsql & " and "

	End If

	strsql = strsql & " nivel = '" & procnivel & "' "
	xx=strsql
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)
	If session("APLMDLOR") <> "BE"     Then
		If sMsg <> "NO" Then
			strsql = "Select * From objeven Where codproceso='" & idelim & "'"
			Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
			If Not rs.eof Then sMsg = "NO"
			Call ofv.cerrarconsulta(rs)
		End If
	End If

	'************************************************************************************
	' Control de borrado de tablas de parametros
	'************************************************************************************

	'****************************************************
	'Comienzo de Eventos de riesgo
	'****************************************************
	Case UCase("ORTIPOEVENTO")
	STRSQL = "SELECT OBJID FROM ORTIPOCATEGO WHERE TIPOEVENTO = " & sID & Chr(13)
	STRSQL = STRSQL & " UNION ALL " & Chr(13)
	STRSQL = STRSQL & " SELECT OBJID FROM ORTRACER WHERE TRCEVENTORIESGO1 = " & sID & Chr(13)
	STRSQL = STRSQL & " UNION ALL " & Chr(13)
	STRSQL = STRSQL & " SELECT OBJID FROM ORACTIVRIESGO WHERE ITEMRIESGO IN " & Chr(13)
	STRSQL = STRSQL & " (SELECT OBJID FROM ORIMPRIESGO WHERE CATEGO IN " & Chr(13)
	STRSQL = STRSQL & " (select OBJID from dbo.ORTIPOCATEGO WHERE TIPOEVENTO IN " & Chr(13)
	STRSQL = STRSQL & " ( select OBJID from dbo.ORTIPOEVENTO WHERE OBJID = " & sID & Chr(13)
	STRSQL = STRSQL & " ))) " & Chr(13)
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)

	Case UCase("ORTIPOCATEGO")
	STRSQL = " SELECT OBJID FROM ORIMPRIESGO WHERE CATEGO = " & sID & Chr(13)
	STRSQL = STRSQL & " UNION ALL " & Chr(13)
	STRSQL = STRSQL & " SELECT OBJID FROM ORTRACER WHERE TRCEVENTORIESGO2 = " & sID & Chr(13)
	STRSQL = STRSQL & " UNION ALL " & Chr(13)
	STRSQL = STRSQL & " SELECT OBJID FROM ORACTIVRIESGO WHERE ITEMRIESGO IN " & Chr(13)
	STRSQL = STRSQL & " (SELECT OBJID FROM ORIMPRIESGO WHERE CATEGO IN " & Chr(13)
	STRSQL = STRSQL & " (select OBJID from dbo.ORTIPOCATEGO WHERE TIPOEVENTO = " & sID & Chr(13)
	STRSQL = STRSQL & " )) " & Chr(13)
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)

	Case UCase("ORIMPRIESGO")
	STRSQL =   " SELECT OBJID FROM ORACTIVRIESGO WHERE ITEMRIESGO IN " & Chr(13)
	STRSQL = STRSQL & " (SELECT OBJID FROM ORIMPRIESGO WHERE OBJID = " & sID & Chr(13)
	STRSQL = STRSQL & " ) " & Chr(13)
	STRSQL = STRSQL & " UNION ALL " & Chr(13)
	STRSQL = STRSQL & " SELECT OBJID FROM ORTRACER WHERE trcevento = " & sID & Chr(13)
	'RESPONSE.WRITE STRSQL
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de Eventos de riesgo
	'****************************************************

	'****************************************************
	'Comienzo de Eventos de riesgo
	'****************************************************
	Case UCase("ORTIPOCARACTER")
	STRSQL = " SELECT * FROM ORTRACER where trccaracter = " & sID & Chr(13)
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de Eventos de riesgo
	'****************************************************

	'****************************************************
	'Comienzo de Nivel Impacto
	'****************************************************
	Case UCase("ORNIVIMPACTO")
	STRSQL = " SELECT * FROM ORACTIVRIESGO WHERE IMPACTO = " & sID & Chr(13)
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de Nivel Impacto
	'****************************************************

	'****************************************************
	'Comienzo de Nivel Impacto
	'****************************************************
	Case UCase("ORNIVPROBOCURRENCIA")
	STRSQL = " SELECT * FROM ORACTIVRIESGO WHERE PROBABILIDAD = " & sID & Chr(13)
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de Nivel Impacto
	'****************************************************

	'****************************************************
	'Comienzo de EFICIENCIA
	'****************************************************
	Case UCase("ORNIVEFICIENCIA")
	STRSQL = " SELECT * FROM oractivrcontroles WHERE CALIFMITIGANTE = " & sID & Chr(13)
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de EFICIENCIA
	'****************************************************

	'****************************************************
	'Comienzo de AMENAZAS
	'****************************************************
	Case UCase("RGACTGRUPOAMENAZA")
	STRSQL = " SELECT OBJID FROM ORAMENAZAS WHERE FAMILIA = " & sID & Chr(13)
	STRSQL = STRSQL & " UNION ALL " & Chr(13)
	STRSQL = STRSQL & "SELECT OBJID FROM ORAMENAZAS WHERE FAMILIA IN (SELECT OBJID FROM RGACTGRUPOAMENAZA WHERE OBJID = " & sID & Chr(13)
	STRSQL = STRSQL & " )"

	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)

	Case UCase("ORAMENAZAS")
	STRSQL = "SELECT OBJID FROM RGACTGRUPOAMENAZA WHERE ACTRSKCOD = " & sID & Chr(13)
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)


	'CASE UCASE("RGACTGRUPOAMENAZA")
	'STRSQL = " SELECT * FROM ORACTIVRIESGO WHERE PROBABILIDAD = " & sID & CHR(13)
	'set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	'If not rs.eof then sMsg = "NO"
	'call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de AMENAZAS
	'****************************************************

	'****************************************************
	'Comienzo de LINEA DE NEGOCIOS
	'****************************************************
	Case UCase("ORLINNEGO1")
	STRSQL = " SELECT OBJID FROM ORLINNEGO2 WHERE NIVEL1 = " & sID & Chr(13)
	STRSQL = STRSQL & " UNION ALL " & Chr(13)
	STRSQL = STRSQL & " SELECT OBJID FROM ORTRACER WHERE TRCNEGOLINE IN " & Chr(13)
	STRSQL = STRSQL & " (SELECT OBJID FROM ORLINNEGO2 WHERE NIVEL1 IN " & Chr(13)
	STRSQL = STRSQL & " (SELECT OBJID FROM ORLINNEGO1 WHERE OBJID = " & sID & Chr(13)
	STRSQL = STRSQL & " ))" & Chr(13)
	STRSQL = STRSQL & " UNION ALL " & Chr(13)
	STRSQL = STRSQL & " SELECT CODOBJETO AS OBJID FROM OBJETOS WHERE LINNEG IN " & Chr(13)
	STRSQL = STRSQL & " (SELECT OBJID FROM ORLINNEGOCIOS WHERE NIVEL2 IN " & Chr(13)
	STRSQL = STRSQL & " (SELECT OBJID FROM ORLINNEGO2 WHERE NIVEL1 IN " & Chr(13)
	STRSQL = STRSQL & " (SELECT OBJID FROM ORLINNEGO1 WHERE OBJID = " & sID & Chr(13)
	STRSQL = STRSQL & " )))" & Chr(13)
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)

	Case UCase("ORLINNEGO2")
	STRSQL = " SELECT * FROM ORLINNEGOCIOS WHERE NIVEL2 = " & sID & Chr(13)
	STRSQL = STRSQL & " UNION ALL " & Chr(13)
	STRSQL = STRSQL & " SELECT OBJID FROM ORTRACER WHERE TRCNEGOLINE IN " & Chr(13)
	STRSQL = STRSQL & " (SELECT OBJID FROM ORLINNEGO2 WHERE OBJID = " & sID & Chr(13)
	STRSQL = STRSQL & " ) " & Chr(13)
	STRSQL = STRSQL & " UNION ALL " & Chr(13)
	STRSQL = STRSQL & " SELECT CODOBJETO AS OBJID FROM OBJETOS WHERE LINNEG IN " & Chr(13)
	STRSQL = STRSQL & " (SELECT OBJID FROM ORLINNEGOCIOS WHERE NIVEL2 IN " & Chr(13)
	STRSQL = STRSQL & " (SELECT OBJID FROM ORLINNEGO2 WHERE OBJID = " & sID & Chr(13)
	STRSQL = STRSQL & " )) " & Chr(13)

	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)

	Case UCase("ORLINNEGOCIOS")
	STRSQL = " SELECT CODOBJETO AS OBJID FROM OBJETOS WHERE LINNEG IN " & Chr(13)
	STRSQL = STRSQL & " (SELECT OBJID FROM ORLINNEGOCIOS WHERE OBJID = " & sID & Chr(13)
	STRSQL = STRSQL & " ) " & Chr(13)
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de LINEA DE NEGOCIOS
	'****************************************************

	'****************************************************
	'Comienzo de AREAS Y SECTORES
	'****************************************************
	Case UCase("ECO")

	STRSQL = " SELECT DISTINCT EcoSITE.OBJID FROM EcoSITE, SITES WHERE EcoSITE.ECOSITE = SITES.OBJID AND ECO = " & sID & Chr(13)
	STRSQL = STRSQL & " UNION ALL " & Chr(13)
	STRSQL = STRSQL & " SELECT DISTINCT Ecorel.CODREL AS OBJID FROM Ecorel, Eco WHERE Ecorel.ecomaster = Eco.eco AND ecodepend = " & sID & Chr(13)
	STRSQL = STRSQL & " UNION ALL " & Chr(13)
	STRSQL = STRSQL & " SELECT DISTINCT Ecoperfil.OBJID FROM Ecoperfil, Tipoperfil WHERE Ecoperfil.codperfil = Tipoperfil.tipoperfil AND codsector = " & sID & Chr(13)
	STRSQL = STRSQL & " UNION ALL " & Chr(13)
	STRSQL = STRSQL & " SELECT DISTINCT OBJID FROM EcoDESCRIPCION WHERE ECO = " & sID & Chr(13)
	STRSQL = STRSQL & " UNION ALL " & Chr(13)
	STRSQL = STRSQL & " SELECT DISTINCT OBJID FROM ORTRACER WHERE TRCECORESP = " & sID & Chr(13)
	STRSQL = STRSQL & " UNION ALL " & Chr(13)
	STRSQL = STRSQL & " SELECT DISTINCT CODSECTOR AS OBJID FROM OBJSECTOR WHERE CODSECTOR = " & sID & Chr(13)

	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de AREAS Y SECTORES
	'****************************************************

	'****************************************************
	'Comienzo de PRODUCTOS/SERVICIOS
	'****************************************************
	Case UCase("TIPOPRODUCTO")
	STRSQL = " SELECT DISTINCT OBJID FROM ORTRACER WHERE TRCPRODUCTO = " & sID & Chr(13)
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)

	Case UCase("ORTIPOPRODUCTO")
	STRSQL = " SELECT DISTINCT OBJID FROM ORTRACER WHERE TRCPRODUCTO = " & sID & Chr(13)
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de PRODUCTOS/SERVICIOS
	'****************************************************

	'****************************************************
	'Comienzo de SALVAGUARDAS
	'****************************************************
	Case UCase("RGACTGRUPOSALVAGUARDAS")
	STRSQL = " SELECT DISTINCT OBJID FROM ORSALVAGUARDAS WHERE GRUPO = " & sID & Chr(13)
	STRSQL = STRSQL & " UNION ALL " & Chr(13)
	STRSQL = STRSQL & " SELECT DISTINCT OBJID FROM oractivrcontroles WHERE ITEMCONTROL = " & sID & Chr(13)
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)

	Case UCase("ORSALVAGUARDAS")
	STRSQL = " SELECT DISTINCT OBJID FROM RGACTFAMILYSVRSK WHERE ACTRSKCOD = " & sID & Chr(13)
	STRSQL = STRSQL & " UNION ALL " & Chr(13)
	STRSQL = STRSQL & " SELECT DISTINCT OBJID FROM oractivrcontroles WHERE ITEMCONTROL = " & sID & Chr(13)
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)

	'CASE UCASE("RGACTFAMILYSVRSK")
	'STRSQL = "  " & sID & CHR(13)
	'set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	'If not rs.eof then sMsg = "NO"
	'call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de SALVAGUARDAS
	'****************************************************

	'****************************************************
	'Comienzo de ACTIVIDADES
	'****************************************************
	Case UCase("Tipoaccalt")
	STRSQL = " SELECT MODO.TIPOMODO AS MODO FROM Modo LEFT JOIN Tipoaccalt ON MODO = TIPOMODO WHERE TIPOACCIONALT = " & sID & Chr(13)
	STRSQL = STRSQL & " UNION ALL " & Chr(13)
	STRSQL = STRSQL & " SELECT DISTINCT MODO FROM OBJACCION WHERE MODO IN ( " & Chr(13)
	STRSQL = STRSQL & " SELECT MODO.TIPOMODO AS MODO FROM Modo LEFT JOIN Tipoaccalt ON MODO = TIPOMODO WHERE TIPOACCIONALT = " & sID & Chr(13)
	STRSQL = STRSQL & " ) " & Chr(13)
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de ACTIVIDADES
	'****************************************************

	'****************************************************
	'Comienzo de RIESGO REPUTACIONAL
	'****************************************************
	Case UCase("ORRREPUTACIONAL")
	STRSQL = " SELECT DISTINCT OBJID FROM ORACTIVRIESGO WHERE RIESGOREPCOD = " & sID & Chr(13)
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de RIESGO REPUTACIONAL
	'****************************************************

	'****************************************************
	'Comienzo de COSTOS DE OPORTUNIDAD
	'****************************************************
	Case UCase("ORCOSTOPOR")
	STRSQL = " SELECT DISTINCT OBJID FROM ORTRACER WHERE TRCCOSTOPOR = " & sID & Chr(13)
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de COSTOS DE OPORTUNIDAD
	'****************************************************

	'****************************************************
	'Comienzo de FORMULAS
	'****************************************************
	Case UCase("ORFORMULAS")
	STRSQL = " select DISTINCT OBJID FROM orindicadores WHERE FORMULA = " & sID & Chr(13)
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de FORMULAS
	'****************************************************

	'****************************************************
	'Comienzo de TIPOS DE CLASIFICACION
	'****************************************************
	Case UCase("ORCLASIFICAINDI")
	STRSQL = " select DISTINCT OBJID FROM orindicadores WHERE CLASIFICACION = " & sID & Chr(13)
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de TIPOS DE CLASIFICACION
	'****************************************************

	'****************************************************
	'Comienzo de TIPO CONTROLES
	'****************************************************
	Case UCase("ORTIPOCONTROLES")
	STRSQL = " SELECT DISTINCT OBJID FROM oractivrcontroles WHERE MODOCONT = " & sID & Chr(13)
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de TIPO CONTROLES
	'****************************************************

	'****************************************************
	'Comienzo de CONTROLES
	'****************************************************
	Case UCase("ORCONTROLRIESGO")
	STRSQL = " SELECT DISTINCT OBJID FROM oractivrcontroles WHERE ITEMCONTROL = " & sID & Chr(13)
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de CONTROLES
	'****************************************************

	'****************************************************
	'Comienzo de RIESGO RESIDUAL
	'****************************************************
	Case UCase("ORNIVriesgo")
	STRSQL = " SELECT DISTINCT A.OBJID FROM ORACTIVRIESGO A,  ORNIVriesgo NR  WHERE A.RIESGORESIDUAL BETWEEN NR.VALORDESDE AND NR.VALORHASTA AND  NR.OBJID = " & sID & Chr(13)
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de RIESGO RESIDUAL
	'****************************************************

	'****************************************************
	'Comienzo de RIESGO INHERENTE
	'****************************************************
	Case UCase("ORNIVRIESGOBRUTO")
	STRSQL = " SELECT DISTINCT A.OBJID FROM ORACTIVRIESGO A LEFT JOIN ORNIVRIESGOBRUTO RI ON RI.NIVPROB = A.PROBABILIDAD AND RI.NIVIMPA = A.IMPACTO WHERE RI.OBJID = " & sID & Chr(13)
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de RIESGO INHERENTE
	'****************************************************

	'****************************************************
	'Comienzo de RANGOS RIESGO
	'****************************************************
	Case UCase("orrangoriesgo")
	STRSQL = " SELECT DISTINCT OBJID FROM ORNIVRIESGOBRUTO WHERE DESCRI = " & sID & Chr(13)
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de RANGOS RIESGO
	'****************************************************

	'****************************************************
	'Comienzo de CFA SUCURSAL
	'****************************************************
	Case UCase("SUCURSAL")
	STRSQL = " SELECT DISTINCT O.OBJID FROM ORTRACER O INNER JOIN SUCURSAL S ON S.OBJID = O.TRCSUCURSAL WHERE S.OBJID = " & sID & Chr(13)
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de CFA SUCURSAL
	'****************************************************

	'****************************************************
	'Comienzo de CFA PROCESO
	'****************************************************
	Case UCase("TABPROCESOS01")
	STRSQL = " SELECT DISTINCT O.OBJID from ORTRACER O JOIN TABPROCESOS01 P ON P.OBJID = O.TRCPROCESO where P.OBJID = " & sID & Chr(13)
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"
	Call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de CFA PROCESO
	'****************************************************


	'****************************************************
	'Comienzo de control listapuestos
	'****************************************************
	Case UCase("LISTADEPUESTOS")
	STRSQL = " SELECT DISTINCT OBJID from puestorel where IDPST=" & sid2
	'response.write strsql
    stradd=""
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
	If Not rs.eof Then sMsg = "NO"

	 Call ofv.cerrarconsulta(rs)



	'************************************************************************************
	' Fin Control de borrado de tablas de parametros
	'************************************************************************************

End Select

If sMsg = "NO" Then sMsg = "No puede eliminar " & sNombre & " porque esta siendo utilizado "

If sEliminar <> "NO" Then
	If sMsg = "" Then
		If sEliminar = "NSNC" Then
			sMensaje = "Esta seguro que desea eliminar " & sNombre
			dato = "../Mensajes3.asp?msg=" & ofv.convertircar(sMensaje," ","_") & "&volver="
			dato = dato & "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(sID," ","_")
			dato = dato & "&Botones=2" & "&Tabla=" & Tabla & "&Nombre=" & sNombre
			dato = dato & "&Campo=" & nCampo & Chr(13)
			response.redirect dato
		ElseIf sEliminar = "SI" Then


			If UCase(Tabla)="LISTADEPUESTOS" Then
				strsql= "DELETE from RELAPUESTO   WHERE PUESTODEPEN=0" & SID2
				'RESPONSE.WRITE STRSQL
				Set rsbcra= ofv.crearconsultaEx(StrSql,cn,1,volver)
				strsql= "DELETE from RELAPUESTO   WHERE PUESTOMASTER=0" & SID2
				'RESPONSE.WRITE STRSQL
				Set rsbcra= ofv.crearconsultaEx(StrSql,cn,1,volver)
				strsql= "DELETE from Ecoperfil   WHERE codperfil='****PUESTO' and codsector=0" & SID2
				'RESPONSE.WRITE STRSQL
				Set rsbcra= ofv.crearconsultaEx(StrSql,cn,1,volver)
			End If
			If UCase(Tabla)="ORREGINFOBCRA" Then
				strsql= "select codlote from ORREGINFOBCRA  Where " & clave & "=" & sID2
				Set rsbcra= ofv.crearconsultaEx(StrSql,cn,1,volver)

				strsql = "Delete From ORREGINFOBCRADET  Where codlote='" & rsbcra(0)&"'"
				Call ofv.crearconsultaEx(StrSql,cn,1,volver)

				strsql = "Delete From ORRIERRLOTE  Where codlote='" & rsbcra(0)&"'"
				Call ofv.crearconsultaEx(StrSql,cn,1,volver)

			End If

			strsql = "Delete From " & Tabla & " Where " & clave & "=" & sID2

			Call ofv.crearconsultaEx(StrSql,cn,1,volver)
			xxxr=GENLOGEXTDATA("DELETE",TABLA,clave,sID2,"PAR")
		End If
	Else
		dato = "../Mensajes3.asp?msg=" & ofv.convertircar(sMsg," ","_") & "&volver="
		dato = dato & "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(sID," ","_")
		dato = dato & "&Botones=1" & "&Tabla=" & Tabla & "&Nombre=" & sNombre
		dato = dato & "&Campo=" & nCampo & Chr(13)
		response.redirect dato
	End If
End If

Volver = session("volver2")

'RESPONSE.WRITE STRSQL
Set errores = cn.errors
If errores.count = 0 Then
	response.redirect volver
Else
	Call ofv.ErrorDB(cn,volver)
End If

Call ofv.cerrarconn(cn)

%>