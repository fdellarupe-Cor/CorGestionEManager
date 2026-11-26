<!-- #INCLUDE FILE="../IncFile/CorConect.inc" -->
<%

response.expires=0

Tabla = request.querystring("Tabla")
nCampo = request.querystring("Campo")
if nCampo = "" then nCampo = 0
nCampo = CDbl(nCampo)
sID = request.querystring("ID")
sID = ofv.convertircar(sID,"_"," ")
sEliminar = request.querystring("Eliminar")
sNombre = request.querystring("Nombre")
Volver = request.querystring("Volver")
Volver = ofv.convertircar(Volver,"_","&")
if sEliminar = "NSNC" then session("volver2") = volver

on error resume next

set cn = ofv.conectar(ofv.strconn2)

sMsg = ""
'Obtengo tipo de Clave
strsql = "Select * From " & tabla
set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
clave = rs(nCampo).name
s = ofv.QueEsx(rs(nCampo).type)
if rs(nCampo).type = 200 then s = "'"
sID2 = s & ofv.convertircar(sID,"_"," ") & s
call ofv.cerrarconsulta(rs)

'Validaciones
select case ucase(Tabla)
  case ucase("Tipoperfil")
    strsql = "select * from ECOPERFIL Where codperfil=" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    If not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)

    if sMsg <> "NO" then
      strsql = "select * from OBJPROCED Where codperfil=" & sID2
      set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
      If not rs.eof then sMsg = "NO"
      call ofv.cerrarconsulta(rs)

      if sMsg <> "NO" then
        strsql = "select * from OBJACCION Where perfilcont=" & sID2
        set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
        If not rs.eof then sMsg = "NO"
        call ofv.cerrarconsulta(rs)

        if sMsg <> "NO" then
           strsql = "select * from OBJcontrol Where codperfil=" & sID2
           set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
           If not rs.eof then sMsg = "NO"
           call ofv.cerrarconsulta(rs)
        end if
      end if
    end if
  case ucase("Tipoaccalt")
    strsql = "Select * From OBJACCION Where tipoaccion=" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    if not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)
    if sMsg <> "NO" then
       strsql = "Select * From OBJACCIONcontrol Where accion=" & sID2
       set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
       if not rs.eof then sMsg = "NO"
       call ofv.cerrarconsulta(rs)
    end if
  case ucase("Sisgen")
    strsql = "Select * From COMPPRTY Where sistorigen=" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    if not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)
  case ucase("Tipogen")
    strsql = "Select * From COMPPRTY Where tipogen=" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    if not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)
  case ucase("Origen")
    strsql = "Select * From OBJPRTY Where origen="  & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    if not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)
  case ucase("Instrumen")
    strsql = "Select * From OBJPRTY Where instrumen="  & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    if not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)
  case ucase("Objcuenta")
  	'RESPONSE.Write(CHR(13) & CHR(13))
	'RESPONSE.Write(session("APLMDLOR") & CHR(13) & sID2)
    if session("APLMDLOR") <> "BE" then
      strsql = "Select DISTINCT OBJID From OBJESQUEMA2 Where codcuenta=" & sID2 & CHR(13)
	  STRSQL = STRSQL & " UNION ALL " & CHR(13)
	  STRSQL = STRSQL & " SELECT DISTINCT OBJID FROM ORTRACER WHERE TRCCTACBLE IN (SELECT NUMERO FROM ObjCuenta WHERE OBJID = " & sID2 & ") " & CHR(13)
      set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
      if not rs.eof then sMsg = "NO"
      call ofv.cerrarconsulta(rs)
    end if
  case ucase("Niveljerar")
    strsql = "Select * From ECO Where econiv=" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    if not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)
  case ucase("Segmento")
    strsql = "Select * From ECO Where ecoseg=" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    if not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)
  case ucase("Jerarquia")
    strsql = "Select * From ECO Where ecojerar=" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    if not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)
  case ucase("Eco")

    if session("APLMDLOR") <> "BE"     THEN
    strsql = "Select * From ECOPERFIL Where codsector=" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    If not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)

    if sMsg <> "NO" then
      strsql = "Select * From ECOREL Where ecodepend=" & sID2
      set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
      If not rs.eof then sMsg = "NO"
      call ofv.cerrarconsulta(rs)
    end if

    if sMsg <> "NO" then
      strsql = "Select * From ECOREL Where ecomaster=" & sID2
      set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
      If not rs.eof then sMsg = "NO"
      call ofv.cerrarconsulta(rs)
    end if

      if sMsg <> "NO" then
        strsql = "Select * From OBJSECTOR Where codsector=" & sID2
        set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
        If not rs.eof then sMsg = "NO"
        call ofv.cerrarconsulta(rs)
      end if
   end if
  case ucase("Objperfil")
    strsql = "Select * From TIPOBJETO Where perfil=" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    If not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)

'    if sMsg <> "NO" then
'      strsql = "Select * From SUBTEMA Where perfil=" & sID2
'      set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
'      If not rs.eof then sMsg = "NO"
'      call ofv.cerrarconsulta(rs)
'    end if
  case ucase("Subtema")
    strsql = "Select * From PRODSERV Where marcaps=" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    If not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)

    if sMsg <> "NO" then
      strsql = "Select * From RELAOBJ Where tipobjmaster=" & sID2
      set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
      If not rs.eof then sMsg = "NO"
      call ofv.cerrarconsulta(rs)
    end if
  case ucase("Tipobjeto")
    strsql = "Select * From OBJETOS Where tipobjeto=" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    If not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)
  case ucase("Comperfil")
    strsql = "Select * From TIPOCOMP Where perfil=" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    If not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)
  case ucase("Tipocomp")
    strsql = "Select * From COMPLEMENTOS Where tipocomp=" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    If not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)
    if sMsg <> "NO" then
      strsql = "Select * From OBJACCION Where tipocomp=" & sID2
      set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
      If not rs.eof then sMsg = "NO"
      call ofv.cerrarconsulta(rs)
    end if
  case ucase("tabperfil")
    strsql = "Select * From objtabdef Where tabperfil=" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    If not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)

  case ucase("Modo")
    strsql = "Select * From OBJACCION Where modo=" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    If not rs.eof then sMsg = "NO"
    call ofv.cerrarconsulta(rs)
    if sMsg <> "NO" then
      strsql = "Select * From TIPOACCALT Where modo=" & sID2
      set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    end if


 case "TABPROCESOS01"
    strsql = "Select * From TABPROCESOS01 Where objid=" & sID2
    set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
    idelim=rs(1)
    PROCNIVEL=rs("nivel") +1
    call ofv.cerrarconsulta(rs)
    Strsql = "Select * From TABPROCESOS01 "
	  strsql = strsql & " where "

	  if procnivel > 1 then

	    IF PROCNIVEL = 2 THEN
	       strsql = strsql & " codproceso01 = '" & idelim & "' "
	    ELSEIF PROCNIVEL = 3 THEN
	       strsql = strsql & " codproceso02 = '" & idelim & "' "
	    ELSEIF PROCNIVEL = 4 THEN
	       strsql = strsql & " codproceso03 = '" & idelim & "' "
	    ELSE
	       strsql = strsql & " codproceso04 = '" & idelim & "' "
	    END IF

	       strsql = strsql & " and "

	  end if

	  strsql = strsql & " nivel = '" & procnivel & "' "
     xx=strsql
	  set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

    If not rs.eof then sMsg = "NO"
       call ofv.cerrarconsulta(rs)
   if session("APLMDLOR") <> "BE"     THEN
    if sMsg <> "NO" then
     strsql = "Select * From objeven Where codproceso='" & idelim & "'"
     set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
     If not rs.eof then sMsg = "NO"
     call ofv.cerrarconsulta(rs)
     end if
   END IF
   
'************************************************************************************
' Control de borrado de tablas de parametros
'************************************************************************************

	'****************************************************
	'Comienzo de Eventos de riesgo
	'****************************************************
	CASE UCASE("ORTIPOEVENTO")
		STRSQL = "SELECT OBJID FROM ORTIPOCATEGO WHERE TIPOEVENTO = " & sID & CHR(13)
		STRSQL = STRSQL & " UNION ALL " & CHR(13)
		STRSQL = STRSQL & " SELECT OBJID FROM ORTRACER WHERE TRCEVENTORIESGO1 = " & sID & CHR(13)
		STRSQL = STRSQL & " UNION ALL " & CHR(13)
		STRSQL = STRSQL & " SELECT OBJID FROM ORACTIVRIESGO WHERE ITEMRIESGO IN " & CHR(13)
		STRSQL = STRSQL & " (SELECT OBJID FROM ORIMPRIESGO WHERE CATEGO IN " & CHR(13)
		STRSQL = STRSQL & " (select OBJID from dbo.ORTIPOCATEGO WHERE TIPOEVENTO IN " & CHR(13)
		STRSQL = STRSQL & " ( select OBJID from dbo.ORTIPOEVENTO WHERE OBJID = " & sID & CHR(13)
		STRSQL = STRSQL & " ))) " & CHR(13)
		set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If not rs.eof then sMsg = "NO"
		call ofv.cerrarconsulta(rs)

	CASE UCASE("ORTIPOCATEGO")
		STRSQL = " SELECT OBJID FROM ORIMPRIESGO WHERE CATEGO = " & sID & CHR(13)
		STRSQL = STRSQL & " UNION ALL " & CHR(13)
		STRSQL = STRSQL & " SELECT OBJID FROM ORTRACER WHERE TRCEVENTORIESGO2 = " & sID & CHR(13)
		STRSQL = STRSQL & " UNION ALL " & CHR(13)
		STRSQL = STRSQL & " SELECT OBJID FROM ORACTIVRIESGO WHERE ITEMRIESGO IN " & CHR(13)
		STRSQL = STRSQL & " (SELECT OBJID FROM ORIMPRIESGO WHERE CATEGO IN " & CHR(13)
		STRSQL = STRSQL & " (select OBJID from dbo.ORTIPOCATEGO WHERE TIPOEVENTO = " & sID & CHR(13)
		STRSQL = STRSQL & " )) " & CHR(13)
		set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If not rs.eof then sMsg = "NO"
		call ofv.cerrarconsulta(rs)
		
	CASE UCASE("ORIMPRIESGO")
		STRSQL = STRSQL & " SELECT OBJID FROM ORACTIVRIESGO WHERE ITEMRIESGO IN " & CHR(13)
		STRSQL = STRSQL & " (SELECT OBJID FROM ORIMPRIESGO WHERE OBJID = " & sID & CHR(13)
		STRSQL = STRSQL & " ) " & CHR(13)
		set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If not rs.eof then sMsg = "NO"
		call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de Eventos de riesgo
	'****************************************************
	
	'****************************************************
	'Comienzo de Eventos de riesgo
	'****************************************************
	CASE UCASE("ORTIPOCARACTER")
		STRSQL = " SELECT * FROM ORTRACER where trccaracter = " & sID & CHR(13)
		set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If not rs.eof then sMsg = "NO"
		call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de Eventos de riesgo
	'****************************************************
	
	'****************************************************
	'Comienzo de Nivel Impacto
	'****************************************************
	CASE UCASE("ORNIVIMPACTO")
		STRSQL = " SELECT * FROM ORACTIVRIESGO WHERE IMPACTO = " & sID & CHR(13)
		set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If not rs.eof then sMsg = "NO"
		call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de Nivel Impacto
	'****************************************************
	
	'****************************************************
	'Comienzo de Nivel Impacto
	'****************************************************
	CASE UCASE("ORNIVPROBOCURRENCIA")
		STRSQL = " SELECT * FROM ORACTIVRIESGO WHERE PROBABILIDAD = " & sID & CHR(13)
		set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If not rs.eof then sMsg = "NO"
		call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de Nivel Impacto
	'****************************************************
	
	'****************************************************
	'Comienzo de EFICIENCIA
	'****************************************************
	CASE UCASE("ORNIVEFICIENCIA")
		STRSQL = " SELECT * FROM oractivrcontroles WHERE CALIFMITIGANTE = " & sID & CHR(13)
		set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If not rs.eof then sMsg = "NO"
		call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de EFICIENCIA
	'****************************************************
	
	'****************************************************
	'Comienzo de AMENAZAS
	'****************************************************
	CASE UCASE("RGACTGRUPOAMENAZA")
		STRSQL = " SELECT OBJID FROM ORAMENAZAS WHERE FAMILIA = " & sID & CHR(13)
		STRSQL = STRSQL & " UNION ALL " & CHR(13)
		STRSQL = STRSQL & "SELECT OBJID FROM ORAMENAZAS WHERE FAMILIA IN (SELECT OBJID FROM RGACTGRUPOAMENAZA WHERE OBJID = " & sID & CHR(13)
		STRSQL = STRSQL & " )"
		
		set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If not rs.eof then sMsg = "NO"
		call ofv.cerrarconsulta(rs)
		
	CASE UCASE("ORAMENAZAS")
		STRSQL = "SELECT OBJID FROM RGACTGRUPOAMENAZA WHERE ACTRSKCOD = " & sID & CHR(13)
		set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If not rs.eof then sMsg = "NO"
		call ofv.cerrarconsulta(rs)
		
		
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
	CASE UCASE("ORLINNEGO1")
		STRSQL = " SELECT OBJID FROM ORLINNEGO2 WHERE NIVEL1 = " & sID & CHR(13)
		STRSQL = STRSQL & " UNION ALL " & CHR(13)
		STRSQL = STRSQL & " SELECT OBJID FROM ORTRACER WHERE TRCNEGOLINE IN " & CHR(13)
		STRSQL = STRSQL & " (SELECT OBJID FROM ORLINNEGO2 WHERE NIVEL1 IN " & CHR(13)
		STRSQL = STRSQL & " (SELECT OBJID FROM ORLINNEGO1 WHERE OBJID = " & sID & CHR(13)
		STRSQL = STRSQL & " ))" & CHR(13)
		STRSQL = STRSQL & " UNION ALL " & CHR(13)
		STRSQL = STRSQL & " SELECT CODOBJETO AS OBJID FROM OBJETOS WHERE LINNEG IN " & CHR(13)
		STRSQL = STRSQL & " (SELECT OBJID FROM ORLINNEGOCIOS WHERE NIVEL2 IN " & CHR(13)
		STRSQL = STRSQL & " (SELECT OBJID FROM ORLINNEGO2 WHERE NIVEL1 IN " & CHR(13)
		STRSQL = STRSQL & " (SELECT OBJID FROM ORLINNEGO1 WHERE OBJID = " & sID & CHR(13)
		STRSQL = STRSQL & " )))" & CHR(13)
		set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If not rs.eof then sMsg = "NO"
		call ofv.cerrarconsulta(rs)
		
	CASE UCASE("ORLINNEGO2")
		STRSQL = " SELECT * FROM ORLINNEGOCIOS WHERE NIVEL2 = " & sID & CHR(13)
		STRSQL = STRSQL & " UNION ALL " & CHR(13)
		STRSQL = STRSQL & " SELECT OBJID FROM ORTRACER WHERE TRCNEGOLINE IN " & CHR(13)
		STRSQL = STRSQL & " (SELECT OBJID FROM ORLINNEGO2 WHERE OBJID = " & sID & CHR(13)
		STRSQL = STRSQL & " ) " & CHR(13)
		STRSQL = STRSQL & " UNION ALL " & CHR(13)
		STRSQL = STRSQL & " SELECT CODOBJETO AS OBJID FROM OBJETOS WHERE LINNEG IN " & CHR(13)
		STRSQL = STRSQL & " (SELECT OBJID FROM ORLINNEGOCIOS WHERE NIVEL2 IN " & CHR(13)
		STRSQL = STRSQL & " (SELECT OBJID FROM ORLINNEGO2 WHERE OBJID = " & sID & CHR(13)
		STRSQL = STRSQL & " )) " & CHR(13)

		set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If not rs.eof then sMsg = "NO"
		call ofv.cerrarconsulta(rs)
		
	CASE UCASE("ORLINNEGOCIOS")
		STRSQL = " SELECT CODOBJETO AS OBJID FROM OBJETOS WHERE LINNEG IN " & CHR(13)
		STRSQL = STRSQL & " (SELECT OBJID FROM ORLINNEGOCIOS WHERE OBJID = " & sID & CHR(13)
		STRSQL = STRSQL & " ) " & CHR(13)
		set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If not rs.eof then sMsg = "NO"
		call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de LINEA DE NEGOCIOS
	'****************************************************
	
	'****************************************************
	'Comienzo de AREAS Y SECTORES
	'****************************************************
	CASE UCASE("ECO")
		STRSQL = " SELECT DISTINCT EcoSITE.OBJID FROM EcoSITE, SITES WHERE EcoSITE.ECOSITE = SITES.OBJID AND ECO = " & sID & CHR(13)
		STRSQL = STRSQL & " UNION ALL " & CHR(13)
		STRSQL = STRSQL & " SELECT DISTINCT Ecorel.CODREL AS OBJID FROM Ecorel, Eco WHERE Ecorel.ecomaster = Eco.eco AND ecodepend = " & sID & CHR(13)
		STRSQL = STRSQL & " UNION ALL " & CHR(13)
		STRSQL = STRSQL & " SELECT DISTINCT Ecoperfil.OBJID FROM Ecoperfil, Tipoperfil WHERE Ecoperfil.codperfil = Tipoperfil.tipoperfil AND codsector = " & sID & CHR(13)
		STRSQL = STRSQL & " UNION ALL " & CHR(13)
		STRSQL = STRSQL & " SELECT DISTINCT OBJID FROM EcoDESCRIPCION WHERE ECO = " & sID & CHR(13)
		STRSQL = STRSQL & " UNION ALL " & CHR(13)
		STRSQL = STRSQL & " SELECT DISTINCT OBJID FROM ORTRACER WHERE TRCECORESP = " & sID & CHR(13)
		STRSQL = STRSQL & " UNION ALL " & CHR(13)
		STRSQL = STRSQL & " SELECT DISTINCT CODSECTOR AS OBJID FROM OBJSECTOR WHERE CODSECTOR = " & sID & CHR(13)
		set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If not rs.eof then sMsg = "NO"
		call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de AREAS Y SECTORES
	'****************************************************
	
	'****************************************************
	'Comienzo de PRODUCTOS/SERVICIOS
	'****************************************************
	CASE UCASE("TIPOPRODUCTO")
		STRSQL = " SELECT DISTINCT OBJID FROM ORTRACER WHERE TRCPRODUCTO = " & sID & CHR(13)
		set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If not rs.eof then sMsg = "NO"
		call ofv.cerrarconsulta(rs)
		
	CASE UCASE("ORTIPOPRODUCTO")
		STRSQL = " SELECT DISTINCT OBJID FROM ORTRACER WHERE TRCPRODUCTO = " & sID & CHR(13)
		set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If not rs.eof then sMsg = "NO"
		call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de PRODUCTOS/SERVICIOS
	'****************************************************
	
	'****************************************************
	'Comienzo de SALVAGUARDAS
	'****************************************************
	CASE UCASE("RGACTGRUPOSALVAGUARDAS")
		STRSQL = " SELECT DISTINCT OBJID FROM ORSALVAGUARDAS WHERE GRUPO = " & sID & CHR(13)
		STRSQL = STRSQL & " UNION ALL " & CHR(13)
		STRSQL = STRSQL & " SELECT DISTINCT OBJID FROM oractivrcontroles WHERE ITEMCONTROL = " & sID & CHR(13)
		set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If not rs.eof then sMsg = "NO"
		call ofv.cerrarconsulta(rs)
		
	CASE UCASE("ORSALVAGUARDAS")
		STRSQL = " SELECT DISTINCT OBJID FROM RGACTFAMILYSVRSK WHERE ACTRSKCOD = " & sID & CHR(13)
		STRSQL = STRSQL & " UNION ALL " & CHR(13)
		STRSQL = STRSQL & " SELECT DISTINCT OBJID FROM oractivrcontroles WHERE ITEMCONTROL = " & sID & CHR(13)
		set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If not rs.eof then sMsg = "NO"
		call ofv.cerrarconsulta(rs)
		
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
	CASE UCASE("Tipoaccalt")
		STRSQL = " SELECT MODO.TIPOMODO AS MODO FROM Modo LEFT JOIN Tipoaccalt ON MODO = TIPOMODO WHERE TIPOACCIONALT = " & sID & CHR(13)
		STRSQL = STRSQL & " UNION ALL " & CHR(13)
		STRSQL = STRSQL & " SELECT DISTINCT MODO FROM OBJACCION WHERE MODO IN ( " & CHR(13)
		STRSQL = STRSQL & " SELECT MODO.TIPOMODO AS MODO FROM Modo LEFT JOIN Tipoaccalt ON MODO = TIPOMODO WHERE TIPOACCIONALT = " & sID & CHR(13)
		STRSQL = STRSQL & " ) " & CHR(13)
		set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If not rs.eof then sMsg = "NO"
		call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de ACTIVIDADES
	'****************************************************
	
	'****************************************************
	'Comienzo de RIESGO REPUTACIONAL
	'****************************************************
	CASE UCASE("ORRREPUTACIONAL")
		STRSQL = " SELECT DISTINCT OBJID FROM ORACTIVRIESGO WHERE RIESGOREPCOD = " & sID & CHR(13)
		set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If not rs.eof then sMsg = "NO"
		call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de RIESGO REPUTACIONAL
	'****************************************************
	
	'****************************************************
	'Comienzo de COSTOS DE OPORTUNIDAD
	'****************************************************
	CASE UCASE("ORCOSTOPOR")
		STRSQL = " SELECT DISTINCT OBJID FROM ORTRACER WHERE TRCCOSTOPOR = " & sID & CHR(13)
		set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If not rs.eof then sMsg = "NO"
		call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de COSTOS DE OPORTUNIDAD
	'****************************************************
	
	'****************************************************
	'Comienzo de FORMULAS
	'****************************************************
	CASE UCASE("ORFORMULAS")
		STRSQL = " select DISTINCT OBJID FROM orindicadores WHERE FORMULA = " & sID & CHR(13)
		set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If not rs.eof then sMsg = "NO"
		call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de FORMULAS
	'****************************************************
	
	'****************************************************
	'Comienzo de TIPOS DE CLASIFICACION
	'****************************************************
	CASE UCASE("ORCLASIFICAINDI")
		STRSQL = " select DISTINCT OBJID FROM orindicadores WHERE CLASIFICACION = " & sID & CHR(13)
		set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If not rs.eof then sMsg = "NO"
		call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de TIPOS DE CLASIFICACION
	'****************************************************
	
	'****************************************************
	'Comienzo de TIPO CONTROLES
	'****************************************************
	CASE UCASE("ORTIPOCONTROLES")
		STRSQL = " SELECT DISTINCT OBJID FROM oractivrcontroles WHERE MODOCONT = " & sID & CHR(13)
		set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If not rs.eof then sMsg = "NO"
		call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de TIPO CONTROLES
	'****************************************************
	
	'****************************************************
	'Comienzo de CONTROLES
	'****************************************************
	CASE UCASE("ORCONTROLRIESGO")
		STRSQL = " SELECT DISTINCT OBJID FROM oractivrcontroles WHERE ITEMCONTROL = " & sID & CHR(13)
		set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If not rs.eof then sMsg = "NO"
		call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de CONTROLES
	'****************************************************
	
	'****************************************************
	'Comienzo de RIESGO RESIDUAL
	'****************************************************
	CASE UCASE("ORNIVriesgo")
		STRSQL = " SELECT DISTINCT A.OBJID FROM ORACTIVRIESGO A,  ORNIVriesgo NR  WHERE A.RIESGORESIDUAL BETWEEN NR.VALORDESDE AND NR.VALORHASTA AND  NR.OBJID = " & sID & CHR(13)
		set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If not rs.eof then sMsg = "NO"
		call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de RIESGO RESIDUAL
	'****************************************************
	
	'****************************************************
	'Comienzo de RIESGO INHERENTE
	'****************************************************
	CASE UCASE("ORNIVRIESGOBRUTO")
		STRSQL = " SELECT DISTINCT A.OBJID FROM ORACTIVRIESGO A LEFT JOIN ORNIVRIESGOBRUTO RI ON RI.NIVPROB = A.PROBABILIDAD AND RI.NIVIMPA = A.IMPACTO WHERE RI.OBJID = " & sID & CHR(13)
		set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If not rs.eof then sMsg = "NO"
		call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de RIESGO INHERENTE
	'****************************************************
	
	'****************************************************
	'Comienzo de RANGOS RIESGO
	'****************************************************
	CASE UCASE("orrangoriesgo")
		STRSQL = " SELECT DISTINCT OBJID FROM ORNIVRIESGOBRUTO WHERE DESCRI = " & sID & CHR(13)
		set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If not rs.eof then sMsg = "NO"
		call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de RANGOS RIESGO
	'****************************************************
	
	'****************************************************
	'Comienzo de CFA SUCURSAL
	'****************************************************
	CASE UCASE("SUCURSAL")
		STRSQL = " SELECT DISTINCT O.OBJID FROM ORTRACER O INNER JOIN SUCURSAL S ON S.OBJID = O.TRCSUCURSAL WHERE S.OBJID = " & sID & CHR(13)
		set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If not rs.eof then sMsg = "NO"
		call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de CFA SUCURSAL
	'****************************************************
	
	'****************************************************
	'Comienzo de CFA PROCESO
	'****************************************************
	CASE UCASE("TABPROCESOS01")
		STRSQL = " SELECT DISTINCT O.OBJID from ORTRACER O JOIN TABPROCESOS01 P ON P.OBJID = O.TRCPROCESO where P.OBJID = " & sID & CHR(13)
		set rs = ofv.crearconsultaEx(StrSql,cn,1,volver)
		If not rs.eof then sMsg = "NO"
		call ofv.cerrarconsulta(rs)
	'****************************************************
	'Fin de CFA PROCESO
	'****************************************************

'************************************************************************************
' Fin Control de borrado de tablas de parametros
'************************************************************************************
   
end select

if sMsg = "NO" then sMsg = "No puede eliminar " & sNombre & " porque esta siendo utilizado"

if sEliminar <> "NO" then
  if sMsg = "" then
    if sEliminar = "NSNC" then
      sMensaje = "Esta seguro que desea eliminar " & sNombre
      dato = "../Mensajes3.asp?msg=" & ofv.convertircar(sMensaje," ","_") & "&volver="
      dato = dato & "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(sID," ","_")
      dato = dato & "&Botones=2" & "&Tabla=" & Tabla & "&Nombre=" & sNombre
      dato = dato & "&Campo=" & nCampo & chr(13)
      response.redirect dato
    elseif sEliminar = "SI" then

      if ucase(Tabla)="ORREGINFOBCRA" then
       strsql= "select codlote from ORREGINFOBCRA  Where " & clave & "=" & sID2
       set rsbcra= ofv.crearconsultaEx(StrSql,cn,1,volver)

       strsql = "Delete From ORREGINFOBCRADET  Where codlote='" & rsbcra(0)&"'"
       call ofv.crearconsultaEx(StrSql,cn,1,volver)

       strsql = "Delete From ORRIERRLOTE  Where codlote='" & rsbcra(0)&"'"
       call ofv.crearconsultaEx(StrSql,cn,1,volver)

      end if

      strsql = "Delete From " & Tabla & " Where " & clave & "=" & sID2

      call ofv.crearconsultaEx(StrSql,cn,1,volver)
      xxxr=GENLOGEXTDATA("DELETE",TABLA,clave,sID2,"PAR")
    end if
  else
    dato = "../Mensajes3.asp?msg=" & ofv.convertircar(sMsg," ","_") & "&volver="
    dato = dato & "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(sID," ","_")
    dato = dato & "&Botones=1" & "&Tabla=" & Tabla & "&Nombre=" & sNombre
    dato = dato & "&Campo=" & nCampo & chr(13)
    response.redirect dato
  end if
end if

Volver = session("volver2")

'RESPONSE.WRITE STRSQL
set errores = cn.errors
if errores.count = 0 then
  response.redirect volver
else
  call ofv.ErrorDB(cn,volver)
end if

call ofv.cerrarconn(cn)

%>