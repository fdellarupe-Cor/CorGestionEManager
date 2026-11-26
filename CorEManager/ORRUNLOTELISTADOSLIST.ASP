<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/reportes.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
	uv = ofv.uv
	if uv = "NO" then ok = ofv.ValidarUsuario
sHTML = ok
else
	response.expires=0

	sImagenCampo = "ARCHIVO"
	CantRegMostrarx = 100

	response.write corstyler()
	'RESPONSE.Write(CHR(13) & CHR(13) & CHR(13))
	'RESPONSE.Write(REQUEST.QueryString("TRIMESTRE"))
	'RESPONSE.Write(REQUEST.QueryString("LISTADO"))
	'RESPONSE.Write(REQUEST.QueryString("CANTREGISTROS"))
	'RESPONSE.Write(CHR(13) & CHR(13) & CHR(13))

	set oPrtG = new ClsReportes ' server.createobject("CorGeneratorEXLR3.Reportes")
	Set oPrtG.Ofv = Ofv

	oPrtG.RCliLogo = session("clilogo")
	oPrtG.RCliNomb = ofv.convertircar(session("FormN"),"_"," ")
	Formulario = oPrtG.RepHeader2(cstr(""), cstr("600"))

	if request.querystring("seg") <> "" then
		Session("Form") = ofv.convertircar(request.querystring("seg"),"_"," ")
	end if
		xTrColor = "#FFFFFF"

		set cn = ofv.conectar(ofv.strconn2)
		strsql = ""

		IF (CSTR(REQUEST.QueryString("LISTADO")) <= "3") THEN

			wwwheader = chr(13) & " <TR ><TD  class=tt width=40% >DESCRIPCION</TD><TD  class=tt width=25% >IMPORTE</TD><TD  class=tt width=25% >BCRA</TD><TD  class=tt width=10% >CANT</TD></tr>" & chr(13)

			IF (CSTR(REQUEST.QueryString("LISTADO")) = "1") THEN
				TITULO = "<TR><TD COLSPAN='3'><H5>INFORME POR LINEA DE NEGOCIO<H5></TD></TR>"

				strsql = "SELECT B.LINEANEGOCIO, ISNULL(L1.DESCRIPCION, '') AS NIVEL1, ISNULL(L2.DESCRIPCION, '') AS NIVEL2, SUM(B.MONTO) AS TOTAL, COUNT(B.LINEANEGOCIO) AS CANTIDAD " & chr(13)
				strsql = strsql & ",SUM( CASE CARACTER WHEN '06' THEN B.MONTO * -1 WHEN '07' THEN B.MONTO * -1 ELSE B.MONTO END ) AS TOTALBCRA " & chr(13)
				strsql = strsql & "FROM ORREGINFOBCRA A inner JOIN ORREGINFOBCRADET B ON (A.CODLOTE=B.CODLOTE) LEFT JOIN ORLINNEGO2 L2 ON L2.CODBCRA = B.LINEANEGOCIO " & chr(13)
				strsql = strsql & "LEFT JOIN dbo.ORLINNEGO1 L1 ON L2.NIVEL1 = L1.OBJID " & chr(13)
				strsql = strsql & " WHERE A.OBJID = " & REQUEST.QueryString("TRIMESTRE") & chr(13)
				strsql = strsql & " AND B.ESTADO = 'N' " & chr(13)
				strsql = strsql & " GROUP BY B.LINEANEGOCIO, L1.DESCRIPCION, L2.DESCRIPCION " & chr(13)
				strsql = strsql & " ORDER BY L1.DESCRIPCION, L2.DESCRIPCION "

				strsqlR = "SELECT B.LINEANEGOCIO, ISNULL(L1.DESCRIPCION, '') AS NIVEL1, ISNULL(L2.DESCRIPCION, '') AS NIVEL2, SUM(B.MONTO) AS TOTAL, COUNT(B.LINEANEGOCIO) AS CANTIDAD " & chr(13)
				strsqlR = strsqlR & ",SUM( CASE CARACTER WHEN '06' THEN B.MONTO * -1 WHEN '07' THEN B.MONTO * -1 ELSE B.MONTO END ) AS TOTALBCRA " & chr(13)
				strsqlR = strsqlR & "FROM ORREGINFOBCRA A inner JOIN ORREGINFOBCRADET B ON (A.CODLOTE=B.CODLOTE) LEFT JOIN ORLINNEGO2 L2 ON L2.CODBCRA = B.LINEANEGOCIO " & chr(13)
				strsqlR = strsqlR & "LEFT JOIN dbo.ORLINNEGO1 L1 ON L2.NIVEL1 = L1.OBJID " & chr(13)
				strsqlR = strsqlR & " WHERE A.OBJID = " & REQUEST.QueryString("TRIMESTRE") & chr(13)
				strsqlR = strsqlR & " AND B.ESTADO = 'R' " & chr(13)
				strsqlR = strsqlR & " GROUP BY B.LINEANEGOCIO, L1.DESCRIPCION, L2.DESCRIPCION " & chr(13)
				strsqlR = strsqlR & " ORDER BY L1.DESCRIPCION, L2.DESCRIPCION "

				strsqlJ = "SELECT B.LINEANEGOCIO, ISNULL(L1.DESCRIPCION, '') AS NIVEL1, ISNULL(L2.DESCRIPCION, '') AS NIVEL2, SUM(B.MONTO) AS TOTAL, COUNT(B.LINEANEGOCIO) AS CANTIDAD " & chr(13)
				strsqlJ = strsqlJ & ",SUM( CASE CARACTER WHEN '06' THEN B.MONTO * -1 WHEN '07' THEN B.MONTO * -1 ELSE B.MONTO END ) AS TOTALBCRA " & chr(13)
				strsqlJ = strsqlJ & "FROM ORREGINFOBCRA A inner JOIN ORREGINFOBCRADET B ON (A.CODLOTE=B.CODLOTE) LEFT JOIN ORLINNEGO2 L2 ON L2.CODBCRA = B.LINEANEGOCIO " & chr(13)
				strsqlJ = strsqlJ & "LEFT JOIN dbo.ORLINNEGO1 L1 ON L2.NIVEL1 = L1.OBJID " & chr(13)
				strsqlJ = strsqlJ & " WHERE A.OBJID = " & REQUEST.QueryString("TRIMESTRE") & chr(13)
				strsqlJ = strsqlJ & " AND B.ESTADO = 'J' " & chr(13)
				strsqlJ = strsqlJ & " GROUP BY B.LINEANEGOCIO, L1.DESCRIPCION, L2.DESCRIPCION " & chr(13)
				strsqlJ = strsqlJ & " ORDER BY L1.DESCRIPCION, L2.DESCRIPCION "

			END IF

			IF (REQUEST.QueryString("LISTADO") = "2") THEN
				TITULO = "<TR><TD COLSPAN='3'><H5>INFORME POR CLASIFICACION TIPO EVENTO<H5></TD></TR>"

				strsql = "SELECT B.TIPOEVENTO, ISNULL(TE1.DESCRI,'') AS NIVEL1, ISNULL(TE2.DESCRI, '') AS NIVEL2, SUM(B.MONTO) AS TOTAL, COUNT(B.LINEANEGOCIO) AS CANTIDAD " & CHR(13)
				strsql = strsql & ",SUM( CASE CARACTER WHEN '06' THEN B.MONTO * -1 WHEN '07' THEN B.MONTO * -1 ELSE B.MONTO END ) AS TOTALBCRA " & chr(13)
				strsql = strsql & " FROM ORREGINFOBCRA A inner JOIN ORREGINFOBCRADET B 	ON (A.CODLOTE=B.CODLOTE) " & CHR(13)
				strsql = strsql & " LEFT JOIN ORTIPOCATEGO TE2 ON TE2.CODBCRA = B.TIPOEVENTO " & CHR(13)
				strsql = strsql & " LEFT JOIN ORTIPOEVENTO TE1 ON TE1.OBJID = TE2.TIPOEVENTO " & CHR(13)
				strsql = strsql & " WHERE A.OBJID = " & REQUEST.QueryString("TRIMESTRE") & CHR(13)
				strsql = strsql & " AND B.ESTADO = 'N' " & chr(13)
				strsql = strsql & " GROUP BY B.TIPOEVENTO, TE1.DESCRI, TE2.DESCRI " & CHR(13)
				strsql = strsql & " ORDER BY TE1.DESCRI, TE2.DESCRI "

				strsqlR = "SELECT B.TIPOEVENTO, ISNULL(TE1.DESCRI,'') AS NIVEL1, ISNULL(TE2.DESCRI, '') AS NIVEL2, SUM(B.MONTO) AS TOTAL, COUNT(B.LINEANEGOCIO) AS CANTIDAD " & CHR(13)
				strsqlR = strsqlR & ",SUM( CASE CARACTER WHEN '06' THEN B.MONTO * -1 WHEN '07' THEN B.MONTO * -1 ELSE B.MONTO END ) AS TOTALBCRA " & chr(13)
				strsqlR = strsqlR & " FROM ORREGINFOBCRA A inner JOIN ORREGINFOBCRADET B 	ON (A.CODLOTE=B.CODLOTE) " & CHR(13)
				strsqlR = strsqlR & " LEFT JOIN ORTIPOCATEGO TE2 ON TE2.CODBCRA = B.TIPOEVENTO " & CHR(13)
				strsqlR = strsqlR & " LEFT JOIN ORTIPOEVENTO TE1 ON TE1.OBJID = TE2.TIPOEVENTO " & CHR(13)
				strsqlR = strsqlR & " WHERE A.OBJID = " & REQUEST.QueryString("TRIMESTRE") & CHR(13)
				strsqlR = strsqlR & " AND B.ESTADO = 'R' " & chr(13)
				strsqlR = strsqlR & " GROUP BY B.TIPOEVENTO, TE1.DESCRI, TE2.DESCRI " & CHR(13)
				strsqlR = strsqlR & " ORDER BY TE1.DESCRI, TE2.DESCRI "

				strsqlJ = "SELECT B.TIPOEVENTO, ISNULL(TE1.DESCRI,'') AS NIVEL1, ISNULL(TE2.DESCRI, '') AS NIVEL2, SUM(B.MONTO) AS TOTAL, COUNT(B.LINEANEGOCIO) AS CANTIDAD " & CHR(13)
				strsqlJ = strsqlJ & ",SUM( CASE CARACTER WHEN '06' THEN B.MONTO * -1 WHEN '07' THEN B.MONTO * -1 ELSE B.MONTO END ) AS TOTALBCRA " & chr(13)
				strsqlJ = strsqlJ & " FROM ORREGINFOBCRA A inner JOIN ORREGINFOBCRADET B 	ON (A.CODLOTE=B.CODLOTE) " & CHR(13)
				strsqlJ = strsqlJ & " LEFT JOIN ORTIPOCATEGO TE2 ON TE2.CODBCRA = B.TIPOEVENTO " & CHR(13)
				strsqlJ = strsqlJ & " LEFT JOIN ORTIPOEVENTO TE1 ON TE1.OBJID = TE2.TIPOEVENTO " & CHR(13)
				strsqlJ = strsqlJ & " WHERE A.OBJID = " & REQUEST.QueryString("TRIMESTRE") & CHR(13)
				strsqlJ = strsqlJ & " AND B.ESTADO = 'J' " & chr(13)
				strsqlJ = strsqlJ & " GROUP BY B.TIPOEVENTO, TE1.DESCRI, TE2.DESCRI " & CHR(13)
				strsqlJ = strsqlJ & " ORDER BY TE1.DESCRI, TE2.DESCRI "

			END IF

			IF (REQUEST.QueryString("LISTADO") = "3") THEN
				TITULO = "<TR><TD COLSPAN='3'><H5>INFORME POR PROCESO<H5></TD></TR>"

				IF (SESSION("APLMDLOR") & "") = "BE" THEN
					strsql = " SELECT B.PROCESO AS NIVEL1 ,ISNULL(TP.DESCRIPCION, '') AS NIVEL2 ,SUM(B.MONTO) AS TOTAL ,COUNT(B.PROCESO) AS CANTIDAD  " & CHR(13)
					strsql = strsql & ",SUM( CASE CARACTER WHEN '06' THEN B.MONTO * -1 WHEN '07' THEN B.MONTO * -1 ELSE B.MONTO END ) AS TOTALBCRA " & chr(13)
					strsql = strsql & " FROM ORREGINFOBCRA A inner JOIN ORREGINFOBCRADET B ON (A.CODLOTE=B.CODLOTE) LEFT JOIN ORTRACER OT " & CHR(13)
					strsql = strsql & " ON B.OBJIDORTRACER = OT.OBJID LEFT JOIN TABPROCESOS01 TP ON TP.OBJID = OT.TRCPROCESO AND TP.NIVEL = 2 " & CHR(13)
					'strsql = strsql & " AND B.ESTADO = 'N' " & chr(13)
					strsql = strsql & " GROUP BY B.PROCESO ,TP.DESCRIPCION "
				ELSE
					strsql = " SELECT B.PROCESO AS NIVEL1 ,ISNULL(SP.DESCRI, '') AS NIVEL2 ,SUM(B.MONTO) AS TOTAL ,COUNT(B.PRODUCTO) AS CANTIDAD " & CHR(13)
					strsql = strsql & ",SUM( CASE CARACTER WHEN '06' THEN B.MONTO * -1 WHEN '07' THEN B.MONTO * -1 ELSE B.MONTO END ) AS TOTALBCRA " & chr(13)
					strsql = strsql & " FROM ORREGINFOBCRA A inner JOIN ORREGINFOBCRADET B 	ON (A.CODLOTE=B.CODLOTE) " & CHR(13)
					strsql = strsql & " LEFT JOIN ORTRACER T ON T.OBJID = B.OBJIDORTRACER LEFT JOIN " & CHR(13)
					strsql = strsql & " ( " & CHR(13)
					strsql = strsql & " 	SELECT ax.cobj AS COBJ, ISNULL(B.descri,'') AS DESCRI " & CHR(13)
					strsql = strsql & " 	FROM dbo.CsRelPath ax LEFT OUTER JOIN " & CHR(13)
					strsql = strsql & " 	dbo.OBJETOS O1 ON ax.mc1 = O1.codobjeto LEFT OUTER JOIN " & CHR(13)
					strsql = strsql & " 	dbo.OBJETOS O2 ON ax.mc2 = O2.codobjeto LEFT OUTER JOIN " & CHR(13)
					strsql = strsql & " 	dbo.OBJETOS O3 ON ax.mc3 = O3.codobjeto LEFT OUTER JOIN " & CHR(13)
					strsql = strsql & " 	dbo.OBJETOS O4 ON ax.mc4 = O4.codobjeto LEFT OUTER JOIN " & CHR(13)
					strsql = strsql & " 	dbo.OBJETOS O5 ON ax.mc5 = O5.codobjeto LEFT OUTER JOIN " & CHR(13)
					strsql = strsql & " 	dbo.OBJETOS O6 ON ax.mc6 = O6.codobjeto LEFT OUTER JOIN " & CHR(13)
					strsql = strsql & " 	dbo.OBJETOS O7 ON ax.mc7 = O7.codobjeto LEFT OUTER JOIN " & CHR(13)
					strsql = strsql & " 	dbo.OBJETOS O8 ON ax.cobj = O8.codobjeto " & CHR(13)
					strsql = strsql & " 	inner join objetos b on (ax.cobj = b.codobjeto) " & CHR(13)
					strsql = strsql & " 	inner join objPRTY b1 on (ax.cobj = b1.codobjeto) " & CHR(13)
					strsql = strsql & " 	WHERE (B.CODVINOBJ = 'VN' or B.CODVINOBJ = 'CR' or B.CODVINOBJ = 'AI') " & CHR(13)
					strsql = strsql & " ) SP ON SP.COBJ = T.TRCPRODSER " & CHR(13)
					strsql = strsql & " WHERE A.OBJID = " & REQUEST.QueryString("TRIMESTRE") & chr(13)
					strsql = strsql & " AND B.ESTADO = 'N' " & chr(13)
					strsql = strsql & " GROUP BY B.PROCESO, SP.DESCRI " & CHR(13)
					strsql = strsql & " ORDER BY B.PROCESO " & CHR(13)

					strsqlR = " SELECT B.PROCESO AS NIVEL1 ,ISNULL(SP.DESCRI, '') AS NIVEL2 ,SUM(B.MONTO) AS TOTAL ,COUNT(B.PRODUCTO) AS CANTIDAD " & CHR(13)
					strsqlR = strsqlR & ",SUM( CASE CARACTER WHEN '06' THEN B.MONTO * -1 WHEN '07' THEN B.MONTO * -1 ELSE B.MONTO END ) AS TOTALBCRA " & chr(13)
					strsqlR = strsqlR & " FROM ORREGINFOBCRA A inner JOIN ORREGINFOBCRADET B 	ON (A.CODLOTE=B.CODLOTE) " & CHR(13)
					strsqlR = strsqlR & " LEFT JOIN ORTRACER T ON T.OBJID = B.OBJIDORTRACER LEFT JOIN " & CHR(13)
					strsqlR = strsqlR & " ( " & CHR(13)
					strsqlR = strsqlR & " 	SELECT ax.cobj AS COBJ, ISNULL(B.descri,'') AS DESCRI " & CHR(13)
					strsqlR = strsqlR & " 	FROM dbo.CsRelPath ax LEFT OUTER JOIN " & CHR(13)
					strsqlR = strsqlR & " 	dbo.OBJETOS O1 ON ax.mc1 = O1.codobjeto LEFT OUTER JOIN " & CHR(13)
					strsqlR = strsqlR & " 	dbo.OBJETOS O2 ON ax.mc2 = O2.codobjeto LEFT OUTER JOIN " & CHR(13)
					strsqlR = strsqlR & " 	dbo.OBJETOS O3 ON ax.mc3 = O3.codobjeto LEFT OUTER JOIN " & CHR(13)
					strsqlR = strsqlR & " 	dbo.OBJETOS O4 ON ax.mc4 = O4.codobjeto LEFT OUTER JOIN " & CHR(13)
					strsqlR = strsqlR & " 	dbo.OBJETOS O5 ON ax.mc5 = O5.codobjeto LEFT OUTER JOIN " & CHR(13)
					strsqlR = strsqlR & " 	dbo.OBJETOS O6 ON ax.mc6 = O6.codobjeto LEFT OUTER JOIN " & CHR(13)
					strsqlR = strsqlR & " 	dbo.OBJETOS O7 ON ax.mc7 = O7.codobjeto LEFT OUTER JOIN " & CHR(13)
					strsqlR = strsqlR & " 	dbo.OBJETOS O8 ON ax.cobj = O8.codobjeto " & CHR(13)
					strsqlR = strsqlR & " 	inner join objetos b on (ax.cobj = b.codobjeto) " & CHR(13)
					strsqlR = strsqlR & " 	inner join objPRTY b1 on (ax.cobj = b1.codobjeto) " & CHR(13)
					strsqlR = strsqlR & " 	WHERE (B.CODVINOBJ = 'VN' or B.CODVINOBJ = 'CR' or B.CODVINOBJ = 'AI') " & CHR(13)
					strsqlR = strsqlR & " ) SP ON SP.COBJ = T.TRCPRODSER " & CHR(13)
					strsqlR = strsqlR & " WHERE A.OBJID = " & REQUEST.QueryString("TRIMESTRE") & chr(13)
					strsqlR = strsqlR & " AND B.ESTADO = 'R' " & chr(13)
					strsqlR = strsqlR & " GROUP BY B.PROCESO, SP.DESCRI " & CHR(13)
					strsqlR = strsqlR & " ORDER BY B.PROCESO " & CHR(13)

					strsqlJ = " SELECT B.PROCESO AS NIVEL1 ,ISNULL(SP.DESCRI, '') AS NIVEL2 ,SUM(B.MONTO) AS TOTAL ,COUNT(B.PRODUCTO) AS CANTIDAD " & CHR(13)
					strsqlJ = strsqlJ & ",SUM( CASE CARACTER WHEN '06' THEN B.MONTO * -1 WHEN '07' THEN B.MONTO * -1 ELSE B.MONTO END ) AS TOTALBCRA " & chr(13)
					strsqlJ = strsqlJ & " FROM ORREGINFOBCRA A inner JOIN ORREGINFOBCRADET B 	ON (A.CODLOTE=B.CODLOTE) " & CHR(13)
					strsqlJ = strsqlJ & " LEFT JOIN ORTRACER T ON T.OBJID = B.OBJIDORTRACER LEFT JOIN " & CHR(13)
					strsqlJ = strsqlJ & " ( " & CHR(13)
					strsqlJ = strsqlJ & " 	SELECT ax.cobj AS COBJ, ISNULL(B.descri,'') AS DESCRI " & CHR(13)
					strsqlJ = strsqlJ & " 	FROM dbo.CsRelPath ax LEFT OUTER JOIN " & CHR(13)
					strsqlJ = strsqlJ & " 	dbo.OBJETOS O1 ON ax.mc1 = O1.codobjeto LEFT OUTER JOIN " & CHR(13)
					strsqlJ = strsqlJ & " 	dbo.OBJETOS O2 ON ax.mc2 = O2.codobjeto LEFT OUTER JOIN " & CHR(13)
					strsqlJ = strsqlJ & " 	dbo.OBJETOS O3 ON ax.mc3 = O3.codobjeto LEFT OUTER JOIN " & CHR(13)
					strsqlJ = strsqlJ & " 	dbo.OBJETOS O4 ON ax.mc4 = O4.codobjeto LEFT OUTER JOIN " & CHR(13)
					strsqlJ = strsqlJ & " 	dbo.OBJETOS O5 ON ax.mc5 = O5.codobjeto LEFT OUTER JOIN " & CHR(13)
					strsqlJ = strsqlJ & " 	dbo.OBJETOS O6 ON ax.mc6 = O6.codobjeto LEFT OUTER JOIN " & CHR(13)
					strsqlJ = strsqlJ & " 	dbo.OBJETOS O7 ON ax.mc7 = O7.codobjeto LEFT OUTER JOIN " & CHR(13)
					strsqlJ = strsqlJ & " 	dbo.OBJETOS O8 ON ax.cobj = O8.codobjeto " & CHR(13)
					strsqlJ = strsqlJ & " 	inner join objetos b on (ax.cobj = b.codobjeto) " & CHR(13)
					strsqlJ = strsqlJ & " 	inner join objPRTY b1 on (ax.cobj = b1.codobjeto) " & CHR(13)
					strsqlJ = strsqlJ & " 	WHERE (B.CODVINOBJ = 'VN' or B.CODVINOBJ = 'CR' or B.CODVINOBJ = 'AI') " & CHR(13)
					strsqlJ = strsqlJ & " ) SP ON SP.COBJ = T.TRCPRODSER " & CHR(13)
					strsqlJ = strsqlJ & " WHERE A.OBJID = " & REQUEST.QueryString("TRIMESTRE") & chr(13)
					strsqlJ = strsqlJ & " AND B.ESTADO = 'J' " & chr(13)
					strsqlJ = strsqlJ & " GROUP BY B.PROCESO, SP.DESCRI " & CHR(13)
					strsqlJ = strsqlJ & " ORDER BY B.PROCESO " & CHR(13)
				END IF

			END IF

			'RESPONSE.Write(CHR(13) & CHR(13) & CHR(13))
			'RESPONSE.Write strsqlJ
			'RESPONSE.Write(CHR(13) & CHR(13) & CHR(13))

			set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			set rsR = ofv.crearconsultaEx(StrSqlR,cn,1,parametros)
			set rsJ = ofv.crearconsultaEx(StrSqlJ,cn,1,parametros)

			FormularioD = ""
			FormularioD = FormularioD & TITULO


			posi = 0
			TTOTALPESOS = 0
			TTOTALPESOSBCRA = 0
			TTOTALCANTIDAD = 0


'************************************************************************************************************
'* Comienzo configuracion para los datos de RECTIFICATIVA
'************************************************************************************************************

		IF(NOT RSR.EOF) THEN
			TSUBTOTALPESOS = 0
			TSUBTOTALPESOSBCRA = 0
			TSUBTOTALCANTIDAD = 0
			TSUBTOTALPESOSPARTICULAR = 0
			TSUBTOTALPESOSPARTICULARBCRA = 0
			TSUBTOTALCANTIDADPARTICULAR = 0
			SUBTOTALESRECTIFICATIVAS = 1
			NIVEL1 = ""
			sinput="<hr width='100%'>"
			Celda = ofv.gencelda("","vti colspan=3 ","","","","","",sInput,"si")
			Fila = ofv.genrow("","","","","25","","",Celda,"si")
			FormularioD = FormularioD & fila
			Fila = ""
			Celda = ""
			TITULON = "<TR><TD COLSPAN='3'><H5>ESTADO RECTIFICATIVO<H5></TD></TR>"
			FormularioD = FormularioD & TITULON
			FormularioD = FormularioD & wwwheader
				Do While (Not rsr.EOF)
					Celda =""

					IF (NIVEL1 <> RSR("NIVEL1")) THEN
						IF (CDBL(TSUBTOTALPESOS) <> 0) THEN
							'******************************************************************************
							'AGREGO EL SUB TOTAL DEL NIVEL
							'******************************************************************************
							sinput=""
							Celda = ofv.gencelda("","","","","","","",sInput,"si")

							sinput="<hr width='90%' align='right'>"
							Celda = Celda & ofv.gencelda("","vti colspan=2 ","","","","","",sInput,"si")

							Fila = ofv.genrow("","","","","25","","",Celda,"si")
							FormularioD = FormularioD & fila
							Fila = ""
							Celda = ""

							sinput="<pre style='display:inline'>&#09;&#09;</pre> TOTAL: "
							Celda = ofv.gencelda("","zz style='text-align=left;' ","","","","","",sInput,"si")

							'sinput= "<div class='xxxxxxxx'  align='right'>" & FormatNumber(TSUBTOTALPESOS,2) & "</div>"
							sinput=  FormatNumber(TSUBTOTALPESOS,2)
							Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

							sinput=  FormatNumber(TSUBTOTALPESOSBCRA,2)
							Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

							'sinput= "<div class='xxxxxxxx'  align='right'>" & TSUBTOTALCANTIDAD & "</div>"
							sinput = TSUBTOTALCANTIDAD
							Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

							Fila = ofv.genrow("","","","","25","","",Celda,"si")
							FormularioD = FormularioD & fila
							Fila = ""
							Celda = ""

							sinput="<hr width='100%'>"
							Celda = ofv.gencelda("","vti colspan=3","","","","","",sInput,"si")

							Fila = ofv.genrow("","","","","25","","",Celda,"si")
							FormularioD = FormularioD & fila
							Fila = ""
							Celda = ""

							'******************************************************************************
							'FINALIZO LA CARGA DEL SUB TOTAL
							'******************************************************************************
							TSUBTOTALPESOS = 0
							TSUBTOTALPESOSBCRA = 0
							TSUBTOTALCANTIDAD = 0
						END IF
						sinput="<U>" & RSR("NIVEL1") & "</U>"
						Celda = ofv.gencelda("","colspan=3","","","","","",sInput,"si")
						Fila = ofv.genrow("","","","","25","","",Celda,"si")
						FormularioD = FormularioD & fila
						NIVEL1 = RSR("NIVEL1")
						TSUBTOTALPESOS = 0
						TSUBTOTALPESOSBCRA = 0
						TSUBTOTALCANTIDAD = 0
					END IF

					Celda =""

					sinput= "<pre style='display:inline'>&#09;</pre>" & RSR("NIVEL2")
					Celda = ofv.gencelda("","","","","","","",sInput,"si")

					TSUBTOTALPESOS = TSUBTOTALPESOS + CDBL(RSR("TOTAL"))
					TTOTALPESOS = TTOTALPESOS + CDBL(RSR("TOTAL"))
					TSUBTOTALPESOSPARTICULAR = TSUBTOTALPESOSPARTICULAR + CDBL(RSR("TOTAL"))
					'sinput="<div class='xxxxxxxx'  align='right'>" & FormatNumber(RSR("TOTAL"),2) & "</div>"
					sinput = FormatNumber(RSR("TOTAL"),2)
					Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

					TSUBTOTALPESOSBCRA = TSUBTOTALPESOSBCRA + CDBL(RSR("TOTALBCRA"))
					TTOTALPESOSBCRA = TTOTALPESOSBCRA + CDBL(RSR("TOTALBCRA"))
					TSUBTOTALPESOSPARTICULARBCRA = TSUBTOTALPESOSPARTICULARBCRA + CDBL(RSR("TOTALBCRA"))
					'sinput="<div class='xxxxxxxx'  align='right'>" & FormatNumber(RSR("TOTAL"),2) & "</div>"
					sinput = FormatNumber(RSR("TOTALBCRA"),2)
					Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

					TSUBTOTALCANTIDAD = TSUBTOTALCANTIDAD + CDBL(RSR("CANTIDAD"))
					TTOTALCANTIDAD = TTOTALCANTIDAD + CDBL(RSR("CANTIDAD"))
					TSUBTOTALCANTIDADPARTICULAR = TSUBTOTALCANTIDADPARTICULAR + CDBL(RSR("CANTIDAD"))
					'sinput="<div class='xxxxxxxx'  align='right'>" & RSR("CANTIDAD") & "</div>"
					sinput= RSR("CANTIDAD")
					Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")
					Fila = ofv.genrow("","","","","25","","",Celda,"si")

					FormularioD = FormularioD & fila

					RSR.Movenext
				Loop
				Fila = ""
				Celda = ""

				sinput=""
				Celda = ofv.gencelda("","","","","","","",sInput,"si")

				sinput="<hr width='90%' align='right'>"
				Celda = Celda & ofv.gencelda("","vti colspan=2 ","","","","","",sInput,"si")

				Fila = ofv.genrow("","","","","25","","",Celda,"si")
				FormularioD = FormularioD & fila
				Fila = ""
				Celda = ""

				sinput="<pre style='display:inline'>&#09;&#09;</pre> TOTAL: "
				Celda = ofv.gencelda("","","","","","","",sInput,"si")

				'sinput= "<div class='xxxxxxxx'  align='right'>" & FormatNumber(TSUBTOTALPESOS,2) & "</div>"
				sinput= FormatNumber(TSUBTOTALPESOS,2)
				Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

				'sinput= "<div class='xxxxxxxx'  align='right'>" & FormatNumber(TSUBTOTALPESOS,2) & "</div>"
				sinput= FormatNumber(TSUBTOTALPESOSBCRA,2)
				Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

				'sinput= "<div class='xxxxxxxx'  align='right'>" & TSUBTOTALCANTIDAD & "</div>"
				sinput= TSUBTOTALCANTIDAD
				Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")
				Fila = ofv.genrow("","","","","25","","",Celda,"si")
				FormularioD = FormularioD & fila
				Fila = ""
				Celda = ""

				sinput="<hr width='100%'>"
				Celda = ofv.gencelda("","vti colspan=3 ","","","","","",sInput,"si")
				Fila = ofv.genrow("","","","","25","","",Celda,"si")
				FormularioD = FormularioD & fila
				Fila = ""
				Celda = ""
		END IF
		IF ( SUBTOTALESRECTIFICATIVAS = 1) THEN
			sinput="<pre style='display:inline'>&#09;&#09;</pre> TOTAL RECTIFICATIVO: "
			Celda = ofv.gencelda("","","","","","","",sInput,"si")

			'sinput= "<div class='xxxxxxxx'  align='right'>" & FormatNumber(TSUBTOTALPESOSPARTICULAR,2) & "</div>"
			sinput=  FormatNumber(TSUBTOTALPESOSPARTICULAR,2)
			Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")
			'sinput= "<div class='xxxxxxxx'  align='right'>" & FormatNumber(TSUBTOTALPESOSPARTICULAR,2) & "</div>"
			sinput=  FormatNumber(TSUBTOTALPESOSPARTICULARBCRA,2)
			Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")
			'sinput= "<div class='xxxxxxxx'  align='right'>" & TSUBTOTALCANTIDADPARTICULAR & "</div>"
			sinput= TSUBTOTALCANTIDADPARTICULAR
			Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")
			Fila = ofv.genrow("","","","","25","","",Celda,"si")
			FormularioD = FormularioD & fila
			Fila = ""
			Celda = ""
		END IF

'************************************************************************************************************
'* Fin configuracion para los datos de RECTIFICATIVA
'************************************************************************************************************

'************************************************************************************************************
'* Comienzo configuracion para los datos de NORMAL
'************************************************************************************************************
			IF(NOT RS.EOF) THEN
			TSUBTOTALPESOS = 0
			TSUBTOTALPESOSBCRA = 0
			TSUBTOTALCANTIDAD = 0
			TSUBTOTALPESOSPARTICULAR = 0
			TSUBTOTALPESOSPARTICULARBCRA = 0
			TSUBTOTALCANTIDADPARTICULAR = 0
			SUBTOTALESNORMALES = 1
			NIVEL1 = ""

			'sinput="<hr width='100%'>"
			sinput = "<div></div>"
			Celda = ofv.gencelda("","vti colspan=3 ","","","","","",sInput,"si")
			Fila = ofv.genrow("","","","","25","","",Celda,"si")
			FormularioD = FormularioD & fila
			Fila = ""
			Celda = ""
			TITULON = "<TR><TD COLSPAN='3'><H5>ESTADO NORMAL<H5></TD></TR>"
			FormularioD = FormularioD & TITULON
			FormularioD = FormularioD & wwwheader
				Do While (Not rs.EOF)
					Celda =""

					IF (NIVEL1 <> RS("NIVEL1")) THEN
						IF (CDBL(TSUBTOTALPESOS) <> 0) THEN
							'******************************************************************************
							'AGREGO EL SUB TOTAL DEL NIVEL
							'******************************************************************************
							sinput=""
							Celda = ofv.gencelda("","","","","","","",sInput,"si")

							sinput="<hr width='90%' align='right'>"
							Celda = Celda & ofv.gencelda("","vti colspan=2 ","","","","","",sInput,"si")

							Fila = ofv.genrow("","","","","25","","",Celda,"si")
							FormularioD = FormularioD & fila
							Fila = ""
							Celda = ""

							sinput="<pre style='display:inline'>&#09;&#09;</pre> TOTAL: "
							Celda = Celda & ofv.gencelda("","zz style='text-align=left;'","","","","","",sInput,"si")

							sinput=  FormatNumber(TSUBTOTALPESOS,2)
							Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

							sinput=  FormatNumber(TSUBTOTALPESOSBCRA,2)
							Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

							sinput=   TSUBTOTALCANTIDAD
							Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

							Fila = ofv.genrow("","","","","25","","",Celda,"si")
							FormularioD = FormularioD & fila
							Fila = ""
							Celda = ""

							sinput="<hr width='100%'>"
							Celda = ofv.gencelda("","vti colspan=3","","","","","",sInput,"si")

							Fila = ofv.genrow("","","","","25","","",Celda,"si")
							FormularioD = FormularioD & fila
							Fila = ""
							Celda = ""
							TSUBTOTALPESOS = 0
							TSUBTOTALPESOSBCRA = 0
							TSUBTOTALCANTIDAD = 0
							'******************************************************************************
							'FINALIZO LA CARGA DEL SUB TOTAL
							'******************************************************************************


						END IF
						sinput="<U>" & RS("NIVEL1") & "</U>"
						Celda = ofv.gencelda("","colspan=3","","","","","",sInput,"si")
						Fila = ofv.genrow("","","","","25","","",Celda,"si")
						FormularioD = FormularioD & fila
						NIVEL1 = RS("NIVEL1")
						TSUBTOTALPESOS = 0
						TSUBTOTALPESOSBCRA = 0
						TSUBTOTALCANTIDAD = 0
					END IF

					Celda =""

					sinput= "<pre style='display:inline'>&#09;</pre>" & RS("NIVEL2")
					Celda = ofv.gencelda("","","","","","","",sInput,"si")

					TSUBTOTALPESOS = TSUBTOTALPESOS + CDBL(RS("TOTAL"))
					TTOTALPESOS = TTOTALPESOS + CDBL(RS("TOTAL"))
					TSUBTOTALPESOSPARTICULAR = TSUBTOTALPESOSPARTICULAR + CDBL(RS("TOTAL"))
					sinput= FormatNumber(RS("TOTAL"),2)
					Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

					TSUBTOTALPESOSBCRA = TSUBTOTALPESOSBCRA + CDBL(RS("TOTALBCRA"))
					TTOTALPESOSBCRA = TTOTALPESOSBCRA + CDBL(RS("TOTALBCRA"))
					TSUBTOTALPESOSPARTICULARBCRA = TSUBTOTALPESOSPARTICULARBCRA + CDBL(RS("TOTALBCRA"))
					sinput= FormatNumber(RS("TOTALBCRA"),2)
					Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

					TSUBTOTALCANTIDAD = TSUBTOTALCANTIDAD + CDBL(RS("CANTIDAD"))
					TTOTALCANTIDAD = TTOTALCANTIDAD + CDBL(RS("CANTIDAD"))
					TSUBTOTALCANTIDADPARTICULAR = TSUBTOTALCANTIDADPARTICULAR + CDBL(RS("CANTIDAD"))
					sinput= RS("CANTIDAD")
					Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")
					Fila = ofv.genrow("","","","","25","","",Celda,"si")

					FormularioD = FormularioD & fila

					rs.Movenext
				Loop
				Fila = ""
				Celda = ""

				sinput=""
				Celda = ofv.gencelda("","","","","","","",sInput,"si")

				sinput="<hr width='90%' align='right'>"
				Celda = Celda & ofv.gencelda("","vti colspan=2 ","","","","","",sInput,"si")

				Fila = ofv.genrow("","","","","25","","",Celda,"si")
				FormularioD = FormularioD & fila
				Fila = ""
				Celda = ""

				sinput="<pre style='display:inline'>&#09;&#09;</pre> TOTAL: "
				Celda = ofv.gencelda("","","","","","","",sInput,"si")

				sinput= FormatNumber(TSUBTOTALPESOS,2)
				Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

				sinput= FormatNumber(TSUBTOTALPESOSBCRA,2)
				Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

				sinput=  TSUBTOTALCANTIDAD
				Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")
				Fila = ofv.genrow("","","","","25","","",Celda,"si")
				FormularioD = FormularioD & fila
				Fila = ""
				Celda = ""

				sinput="<hr width='100%'>"
				Celda = ofv.gencelda("","vti colspan=3 ","","","","","",sInput,"si")
				Fila = ofv.genrow("","","","","25","","",Celda,"si")
				FormularioD = FormularioD & fila
				Fila = ""
				Celda = ""
		END IF
		IF ( SUBTOTALESNORMALES = 1) THEN
			sinput="<pre style='display:inline'>&#09;&#09;</pre> TOTAL NORMAL: "
			Celda = ofv.gencelda("","","","","","","",sInput,"si")

			sinput = FormatNumber(TSUBTOTALPESOSPARTICULAR,2)
			Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")
			sinput = FormatNumber(TSUBTOTALPESOSPARTICULARBCRA,2)
			Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")
			sinput = TSUBTOTALCANTIDADPARTICULAR
			Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")
			Fila = ofv.genrow("","","","","25","","",Celda,"si")
			FormularioD = FormularioD & fila
			Fila = ""
			Celda = ""
		END IF
'************************************************************************************************************
'* Final configuracion para los datos de NORMAL
'************************************************************************************************************

'************************************************************************************************************
'* Comienzo configuracion para los datos de AJUSTE
'************************************************************************************************************
		IF(NOT RSJ.EOF) THEN
			TSUBTOTALPESOS = 0
			TSUBTOTALPESOSBCRA = 0
			TSUBTOTALCANTIDAD = 0
			TSUBTOTALPESOSPARTICULAR = 0
			TSUBTOTALPESOSPARTICULARBCRA = 0
			TSUBTOTALCANTIDADPARTICULAR = 0
			SUBTOTALESAJUSTES = 1
			NIVEL1 = ""
			'sinput="<hr width='100%'>"
			sinput = "<div></div>"
			Celda = ofv.gencelda("","vti colspan=3 ","","","","","",sInput,"si")
			Fila = ofv.genrow("","","","","25","","",Celda,"si")
			FormularioD = FormularioD & fila
			Fila = ""
			Celda = ""
			TITULON = "<TR><TD COLSPAN='3'><H5>ESTADO AJUSTE<H5></TD></TR>"
			FormularioD = FormularioD & TITULON
			FormularioD = FormularioD & wwwheader
				Do While (Not RSJ.EOF)
					Celda =""

					IF (NIVEL1 <> RSJ("NIVEL1")) THEN

						IF (CDBL(TSUBTOTALPESOS) <> 0) THEN

							'******************************************************************************
							'AGREGO EL SUB TOTAL DEL NIVEL
							'******************************************************************************
							sinput=""
							Celda = ofv.gencelda("","","","","","","",sInput,"si")

							sinput="<hr width='90%' align='right'>"
							Celda = Celda & ofv.gencelda("","vti colspan=2 ","","","","","",sInput,"si")

							Fila = ofv.genrow("","","","","25","","",Celda,"si")
							FormularioD = FormularioD & fila
							Fila = ""
							Celda = ""

							sinput="<pre style='display:inline'>&#09;&#09;</pre> TOTAL: "
							Celda = ofv.gencelda(""," text-align=rigth ","","","","","",sInput,"si")

							sinput= FormatNumber(TSUBTOTALPESOS,2)
							Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

							sinput= FormatNumber(TSUBTOTALPESOSBCRA,2)
							'sinput= "hola"
							Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

							sinput=  TSUBTOTALCANTIDAD
							Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

							Fila = ofv.genrow("","","","","25","","",Celda,"si")
							FormularioD = FormularioD & fila
							Fila = ""
							Celda = ""

							sinput="<hr width='100%'>"
							Celda = ofv.gencelda("","vti colspan=3","","","","","",sInput,"si")

							Fila = ofv.genrow("","","","","25","","",Celda,"si")
							FormularioD = FormularioD & fila
							Fila = ""
							Celda = ""

							'******************************************************************************
							'FINALIZO LA CARGA DEL SUB TOTAL
							'******************************************************************************
							TSUBTOTALPESOS = 0
							TSUBTOTALPESOSBCRA = 0
							TSUBTOTALCANTIDAD = 0

						END IF
						sinput="<U>" & RSJ("NIVEL1") & "</U>"
						Celda = ofv.gencelda("","colspan=3","","","","","",sInput,"si")
						Fila = ofv.genrow("","","","","25","","",Celda,"si")
						FormularioD = FormularioD & fila
						NIVEL1 = RSJ("NIVEL1")
						TSUBTOTALPESOS = 0
						TSUBTOTALPESOSBCRA = 0
						TSUBTOTALCANTIDAD = 0
					END IF

					Celda =""

					sinput= "<pre style='display:inline'>&#09;</pre>" & RSJ("NIVEL2")
					Celda = ofv.gencelda("","","","","","","",sInput,"si")

					TSUBTOTALPESOS = TSUBTOTALPESOS + CDBL(RSJ("TOTAL"))
					TTOTALPESOS = TTOTALPESOS + (CDBL(RSJ("TOTAL")))
					TSUBTOTALPESOSPARTICULAR = TSUBTOTALPESOSPARTICULAR + CDBL(RSJ("TOTAL"))
					sinput= FormatNumber(RSJ("TOTAL"),2)
					Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

					TSUBTOTALPESOSBCRA = TSUBTOTALPESOSBCRA + CDBL(RSJ("TOTALBCRA"))
					TTOTALPESOSBCRA = TTOTALPESOSBCRA + CDBL(RSJ("TOTALBCRA"))
					TSUBTOTALPESOSPARTICULARBCRA = TSUBTOTALPESOSPARTICULARBCRA + CDBL(RSJ("TOTALBCRA"))
					sinput= FormatNumber(RSJ("TOTALBCRA"),2)
					Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

					TSUBTOTALCANTIDAD = TSUBTOTALCANTIDAD + CDBL(RSJ("CANTIDAD"))
					TTOTALCANTIDAD = TTOTALCANTIDAD + CDBL(RSJ("CANTIDAD"))
					TSUBTOTALCANTIDADPARTICULAR = TSUBTOTALCANTIDADPARTICULAR + CDBL(RSJ("CANTIDAD"))
					sinput= RSJ("CANTIDAD")
					Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")
					Fila = ofv.genrow("","","","","25","","",Celda,"si")

					FormularioD = FormularioD & fila

					RSJ.Movenext
				Loop
				Fila = ""
				Celda = ""

				sinput=""
				Celda = ofv.gencelda("","","","","","","",sInput,"si")

				sinput="<hr width='90%' align='right'>"
				Celda = Celda & ofv.gencelda("","vti colspan=2 ","","","","","",sInput,"si")

				Fila = ofv.genrow("","","","","25","","",Celda,"si")
				FormularioD = FormularioD & fila
				Fila = ""
				Celda = ""

				sinput="<pre style='display:inline'>&#09;&#09;</pre> TOTAL: "
				Celda = ofv.gencelda("","","","","","","",sInput,"si")

				sinput= FormatNumber(TSUBTOTALPESOS,2)
				Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

				sinput= FormatNumber(TSUBTOTALPESOSBCRA,2)
				'sinput= "hola"
				Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

				sinput= TSUBTOTALCANTIDAD
				Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")
				Fila = ofv.genrow("","","","","25","","",Celda,"si")
				FormularioD = FormularioD & fila
				Fila = ""
				Celda = ""

				sinput="<hr width='100%'>"
				Celda = ofv.gencelda("","vti colspan=3 ","","","","","",sInput,"si")
				Fila = ofv.genrow("","","","","25","","",Celda,"si")
				FormularioD = FormularioD & fila
				Fila = ""
				Celda = ""
		END IF
		IF ( SUBTOTALESAJUSTES = 1) THEN
			sinput="<pre style='display:inline'>&#09;&#09;</pre> TOTAL AJUSTE: "
			Celda = ofv.gencelda("","","","","","","",sInput,"si")

			sinput= FormatNumber(TSUBTOTALPESOSPARTICULAR,2)
			Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")
			'sinput= FormatNumber(TSUBTOTALPESOSPARTICULARBCRA,2)
			'sinput= "Hola"
			Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")
			sinput= TSUBTOTALCANTIDADPARTICULAR
			Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")
			Fila = ofv.genrow("","","","","25","","",Celda,"si")
			FormularioD = FormularioD & fila
			Fila = ""
			Celda = ""
		END IF
'************************************************************************************************************
'* Final configuracion para los datos de AJUSTE
'************************************************************************************************************
		IF (CDBL(TTOTALPESOS) > 0) THEN
			'sinput="<hr width='100%'>"
			sinput = "<div></div>"
			Celda = ofv.gencelda("","vti colspan=3 ","","","","","",sInput,"si")
			Fila = ofv.genrow("","","","","25","","",Celda,"si")
			'FormularioD = FormularioD & fila
			FormularioD = FormularioD & fila
			Fila = ""
			Celda = ""
			TITULON = "<TR><TD COLSPAN='3'><H5>TOTALES GENERALES<H5></TD></TR>"
			FormularioD = FormularioD & TITULON
			sinput="<pre style='display:inline'>&#09;&#09;</pre> TOTAL GENERAL: "
			Celda = ofv.gencelda("","","","","","","",sInput,"si")

			sinput= FormatNumber(TTOTALPESOS,2)
			Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

			sinput= FormatNumber(TTOTALPESOSBCRA,2)
			Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

			sinput=  TTOTALCANTIDAD
			Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")
			Fila = ofv.genrow("","","","","25","","",Celda,"si")
			FormularioD = FormularioD & fila
			Fila = ""
			Celda = ""
			Formulario = Formulario & ofv.GenTabla("","","","80%","","0","","0","1","","",FormularioD,"si")
		END IF
		END IF


		IF (CSTR(REQUEST.QueryString("LISTADO")) = "4" OR CSTR(REQUEST.QueryString("LISTADO")) = "5") THEN
			wwwheader = chr(13) & " <TR ><TD  class=tt width=60% >DESCRIPCION</TD><TD  class=tt width=25% >IMPORTE</TD><TD  class=tt width=25% >BCRA</TD><TD  class=tt width=10% >CANT</TD></tr>" & chr(13)

			IF (REQUEST.QueryString("LISTADO") = "4") THEN
				TITULO = "<TR><TD COLSPAN='3'><H5>INFORME POR PRODUCTO<H5></TD></TR>"

				strsql = " SELECT B.PRODUCTO AS NIVEL1, SUM(B.MONTO) AS TOTAL, COUNT(B.PRODUCTO) AS CANTIDAD " & CHR(13)
				strsql = strsql & ",SUM( CASE CARACTER WHEN '06' THEN B.MONTO * -1 WHEN '07' THEN B.MONTO * -1 ELSE B.MONTO END ) AS TOTALBCRA " & chr(13)
				strsql = strsql & " FROM ORREGINFOBCRA A inner JOIN ORREGINFOBCRADET B 	ON (A.CODLOTE=B.CODLOTE)  " & CHR(13)
				strsql = strsql & " WHERE A.OBJID = " & REQUEST.QueryString("TRIMESTRE") & CHR(13)
				strsql = strsql & " AND B.ESTADO = 'N' " & chr(13)
				strsql = strsql & " GROUP BY B.PRODUCTO " & CHR(13)
				strsql = strsql & " ORDER BY B.PRODUCTO " & CHR(13)

				strsqlR = " SELECT B.PRODUCTO AS NIVEL1, SUM(B.MONTO) AS TOTAL, COUNT(B.PRODUCTO) AS CANTIDAD " & CHR(13)
				strsqlR = strsqlR & ",SUM( CASE CARACTER WHEN '06' THEN B.MONTO * -1 WHEN '07' THEN B.MONTO * -1 ELSE B.MONTO END ) AS TOTALBCRA " & chr(13)
				strsqlR = strsqlR & " FROM ORREGINFOBCRA A inner JOIN ORREGINFOBCRADET B 	ON (A.CODLOTE=B.CODLOTE)  " & CHR(13)
				strsqlR = strsqlR & " WHERE A.OBJID = " & REQUEST.QueryString("TRIMESTRE") & CHR(13)
				strsqlR = strsqlR & " AND B.ESTADO = 'R' " & chr(13)
				strsqlR = strsqlR & " GROUP BY B.PRODUCTO " & CHR(13)
				strsqlR = strsqlR & " ORDER BY B.PRODUCTO " & CHR(13)

				strsqlJ = " SELECT B.PRODUCTO AS NIVEL1, SUM(B.MONTO) AS TOTAL, COUNT(B.PRODUCTO) AS CANTIDAD " & CHR(13)
				strsqlJ = strsqlJ & ",SUM( CASE CARACTER WHEN '06' THEN B.MONTO * -1 WHEN '07' THEN B.MONTO * -1 ELSE B.MONTO END ) AS TOTALBCRA " & chr(13)
				strsqlJ = strsqlJ & " FROM ORREGINFOBCRA A inner JOIN ORREGINFOBCRADET B 	ON (A.CODLOTE=B.CODLOTE)  " & CHR(13)
				strsqlJ = strsqlJ & " WHERE A.OBJID = " & REQUEST.QueryString("TRIMESTRE") & CHR(13)
				strsqlJ = strsqlJ & " AND B.ESTADO = 'J' " & chr(13)
				strsqlJ = strsqlJ & " GROUP BY B.PRODUCTO " & CHR(13)
				strsqlJ = strsqlJ & " ORDER BY B.PRODUCTO " & CHR(13)

			END IF


			IF (REQUEST.QueryString("LISTADO") = "5") THEN
				TITULO = "<TR><TD COLSPAN='3'><H5>INFORME POR CARACTER DE EVENTO<H5></TD></TR>"

				strsql = " SELECT B.CARACTER, ISNULL(C.DESCRI,'') AS NIVEL1, SUM(B.MONTO) AS TOTAL, " & CHR(13)
				strsql = strsql & " COUNT(B.LINEANEGOCIO) AS CANTIDAD " & CHR(13)
				strsql = strsql & ",SUM( CASE CARACTER WHEN '06' THEN B.MONTO * -1 WHEN '07' THEN B.MONTO * -1 ELSE B.MONTO END ) AS TOTALBCRA " & chr(13)
				strsql = strsql & " FROM ORREGINFOBCRA A inner JOIN ORREGINFOBCRADET B 	ON (A.CODLOTE=B.CODLOTE) " & CHR(13)
				strsql = strsql & " LEFT JOIN ORTIPOCARACTER C ON C.CODBCRA = B.CARACTER " & CHR(13)
				strsql = strsql & " WHERE A.OBJID = " & REQUEST.QueryString("TRIMESTRE") & CHR(13)
				strsql = strsql & " AND B.ESTADO = 'N' " & chr(13)
				strsql = strsql & " GROUP BY B.CARACTER, C.DESCRI " & CHR(13)
				strsql = strsql & " ORDER BY C.DESCRI " & CHR(13)

				strsqlR = " SELECT B.CARACTER, ISNULL(C.DESCRI,'') AS NIVEL1, SUM(B.MONTO) AS TOTAL, " & CHR(13)
				strsqlR = strsqlR & " COUNT(B.LINEANEGOCIO) AS CANTIDAD " & CHR(13)
				strsqlR = strsqlR & ",SUM( CASE CARACTER WHEN '06' THEN B.MONTO * -1 WHEN '07' THEN B.MONTO * -1 ELSE B.MONTO END ) AS TOTALBCRA " & chr(13)
				strsqlR = strsqlR & " FROM ORREGINFOBCRA A inner JOIN ORREGINFOBCRADET B 	ON (A.CODLOTE=B.CODLOTE) " & CHR(13)
				strsqlR = strsqlR & " LEFT JOIN ORTIPOCARACTER C ON C.CODBCRA = B.CARACTER " & CHR(13)
				strsqlR = strsqlR & " WHERE A.OBJID = " & REQUEST.QueryString("TRIMESTRE") & CHR(13)
				strsqlR = strsqlR & " AND B.ESTADO = 'R' " & chr(13)
				strsqlR = strsqlR & " GROUP BY B.CARACTER, C.DESCRI " & CHR(13)
				strsqlR = strsqlR & " ORDER BY C.DESCRI " & CHR(13)

				strsqlJ = " SELECT B.CARACTER, ISNULL(C.DESCRI,'') AS NIVEL1, SUM(B.MONTO) AS TOTAL, " & CHR(13)
				strsqlJ = strsqlJ & " COUNT(B.LINEANEGOCIO) AS CANTIDAD " & CHR(13)
				strsqlJ = strsqlJ & ",SUM( CASE CARACTER WHEN '06' THEN B.MONTO * -1 WHEN '07' THEN B.MONTO * -1 ELSE B.MONTO END ) AS TOTALBCRA " & chr(13)
				strsqlJ = strsqlJ & " FROM ORREGINFOBCRA A inner JOIN ORREGINFOBCRADET B 	ON (A.CODLOTE=B.CODLOTE) " & CHR(13)
				strsqlJ = strsqlJ & " LEFT JOIN ORTIPOCARACTER C ON C.CODBCRA = B.CARACTER " & CHR(13)
				strsqlJ = strsqlJ & " WHERE A.OBJID = " & REQUEST.QueryString("TRIMESTRE") & CHR(13)
				strsqlJ = strsqlJ & " AND B.ESTADO = 'J' " & chr(13)
				strsqlJ = strsqlJ & " GROUP BY B.CARACTER, C.DESCRI " & CHR(13)
				strsqlJ = strsqlJ & " ORDER BY C.DESCRI " & CHR(13)

			END IF


			set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			set RSR = ofv.crearconsultaEx(StrSqlR,cn,1,parametros)
			set RSJ = ofv.crearconsultaEx(StrSqlJ,cn,1,parametros)

			FormularioD = ""
			FormularioD = FormularioD & TITULO

			TTOTALPESOS = 0
			TTOTALPESOSBCRA = 0
			TTOTALCANTIDAD = 0


			IF(NOT RSR.EOF) THEN
				TSUBTOTALPESOS = 0
				TSUBTOTALPESOSBCRA = 0
				TSUBTOTALCANTIDAD = 0
				sinput="<hr width='100%'>"
				Celda = ofv.gencelda("","vti colspan=3 ","","","","","",sInput,"si")
				Fila = ofv.genrow("","","","","25","","",Celda,"si")
				FormularioD = FormularioD & fila
				Fila = ""
				Celda = ""
				TITULON = "<TR><TD COLSPAN='3'><H5>ESTADO RECTIFICATIVO<H5></TD></TR>"
				FormularioD = FormularioD & TITULON
				FormularioD = FormularioD & wwwheader
				Do While (Not RSR.EOF)
					Celda =""

					sinput= RSR("NIVEL1")
					Celda = ofv.gencelda("","","","","","","",sInput,"si")

					TSUBTOTALPESOS = TSUBTOTALPESOS + CDBL(RSR("TOTAL"))
					TTOTALPESOS = TTOTALPESOS + CDBL(RSR("TOTAL"))
					sinput= FormatNumber(RSR("TOTAL"),2)
					Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

					TSUBTOTALPESOSBCRA = TSUBTOTALPESOSBCRA + CDBL(RSR("TOTALBCRA"))
					TTOTALPESOSBCRA = TTOTALPESOSBCRA + CDBL(RSR("TOTALBCRA"))
					sinput= FormatNumber(RSR("TOTALBCRA"),2)
					Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

					TSUBTOTALCANTIDAD = TSUBTOTALCANTIDAD + CDBL(RSR("CANTIDAD"))
					TTOTALCANTIDAD = TTOTALCANTIDAD + CDBL(RSR("CANTIDAD"))
					sinput= RSR("CANTIDAD")
					Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")
					Fila = ofv.genrow("","","","","25","","",Celda,"si")

					FormularioD = FormularioD & fila

					RSR.MOVENEXT()
				LOOP
				sinput="<hr width='100%'>"
				Celda = ofv.gencelda("","vti colspan=3 ","","","","","",sInput,"si")
				Fila = ofv.genrow("","","","","25","","",Celda,"si")
				FormularioD = FormularioD & fila
				Fila = ""
				Celda = ""

				sinput="<pre style='display:inline'>&#09;&#09;</pre> TOTAL RECTIFICATIVO: "
				Celda = ofv.gencelda("","","","","","","",sInput,"si")

				sinput= FormatNumber(TSUBTOTALPESOS,2)
				Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

				sinput= FormatNumber(TSUBTOTALPESOSBCRA,2)
				Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

				sinput= TSUBTOTALCANTIDAD
				Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")
				Fila = ofv.genrow("","","","","25","","",Celda,"si")
				FormularioD = FormularioD & fila
				Fila = ""
				Celda = ""
			END IF

			IF(NOT RS.EOF) THEN
				TSUBTOTALPESOS = 0
				TSUBTOTALPESOSBCRA = 0
				TSUBTOTALCANTIDAD = 0
				sinput="<hr width='100%'>"
				Celda = ofv.gencelda("","vti colspan=3 ","","","","","",sInput,"si")
				Fila = ofv.genrow("","","","","25","","",Celda,"si")
				FormularioD = FormularioD & fila
				Fila = ""
				Celda = ""
				TITULON = "<TR><TD COLSPAN='3'><H5>ESTADO NORMAL<H5></TD></TR>"
				FormularioD = FormularioD & TITULON
				FormularioD = FormularioD & wwwheader
				Do While (Not rs.EOF)
					Celda =""

					sinput= RS("NIVEL1")
					Celda = ofv.gencelda("","","","","","","",sInput,"si")

					TSUBTOTALPESOS = TSUBTOTALPESOS + CDBL(RS("TOTAL"))
					TTOTALPESOS = TTOTALPESOS + CDBL(RS("TOTAL"))
					sinput= FormatNumber(RS("TOTAL"),2)
					Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

					TSUBTOTALPESOSBCRA = TSUBTOTALPESOSBCRA + CDBL(RS("TOTALBCRA"))
					TTOTALPESOSBCRA = TTOTALPESOSBCRA + CDBL(RS("TOTALBCRA"))
					sinput= FormatNumber(RS("TOTALBCRA"),2)
					Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

					TSUBTOTALCANTIDAD = TSUBTOTALCANTIDAD + CDBL(RS("CANTIDAD"))
					TTOTALCANTIDAD = TTOTALCANTIDAD + CDBL(RS("CANTIDAD"))
					sinput= RS("CANTIDAD")
					Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")
					Fila = ofv.genrow("","","","","25","","",Celda,"si")

					FormularioD = FormularioD & fila

					RS.MOVENEXT()
				LOOP
				sinput="<hr width='100%'>"
				Celda = ofv.gencelda("","vti colspan=3 ","","","","","",sInput,"si")
				Fila = ofv.genrow("","","","","25","","",Celda,"si")
				FormularioD = FormularioD & fila
				Fila = ""
				Celda = ""

				sinput="<pre style='display:inline'>&#09;&#09;</pre> TOTAL NORMAL: "
				Celda = ofv.gencelda("","","","","","","",sInput,"si")

				sinput=  FormatNumber(TSUBTOTALPESOS,2)
				Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

				sinput=  FormatNumber(TSUBTOTALPESOSBCRA,2)
				Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

				sinput=  TSUBTOTALCANTIDAD
				Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")
				Fila = ofv.genrow("","","","","25","","",Celda,"si")
				FormularioD = FormularioD & fila
				Fila = ""
				Celda = ""
			END IF

			IF(NOT RSJ.EOF) THEN
				TSUBTOTALPESOS = 0
				TSUBTOTALPESOSBCRA = 0
				TSUBTOTALCANTIDAD = 0
				sinput="<hr width='100%'>"
				Celda = ofv.gencelda("","vti colspan=3 ","","","","","",sInput,"si")
				Fila = ofv.genrow("","","","","25","","",Celda,"si")
				FormularioD = FormularioD & fila
				Fila = ""
				Celda = ""
				TITULON = "<TR><TD COLSPAN='3'><H5>ESTADO AJUSTE<H5></TD></TR>"
				FormularioD = FormularioD & TITULON
				FormularioD = FormularioD & wwwheader
				Do While (Not RSJ.EOF)
					Celda =""

					sinput= RSJ("NIVEL1")
					Celda = ofv.gencelda("","","","","","","",sInput,"si")

					TSUBTOTALPESOS = TSUBTOTALPESOS + CDBL(RSJ("TOTAL"))
					TTOTALPESOS = TTOTALPESOS + (CDBL(RSJ("TOTAL")))
					sinput= FormatNumber(RSJ("TOTAL"),2)
					Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

					TSUBTOTALPESOSBCRA = TSUBTOTALPESOSBCRA + CDBL(RSJ("TOTALBCRA"))
					TTOTALPESOSBCRA = TTOTALPESOSBCRA + (CDBL(RSJ("TOTALBCRA")))
					sinput= FormatNumber(RSJ("TOTALBCRA"),2)
					Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

					TSUBTOTALCANTIDAD = TSUBTOTALCANTIDAD + CDBL(RSJ("CANTIDAD"))
					TTOTALCANTIDAD = TTOTALCANTIDAD + CDBL(RSJ("CANTIDAD"))
					sinput= RSJ("CANTIDAD")
					Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")
					Fila = ofv.genrow("","","","","25","","",Celda,"si")

					FormularioD = FormularioD & fila

					RSJ.MOVENEXT()
				LOOP
				sinput="<hr width='100%'>"
				Celda = ofv.gencelda("","vti colspan=3 ","","","","","",sInput,"si")
				Fila = ofv.genrow("","","","","25","","",Celda,"si")
				FormularioD = FormularioD & fila
				Fila = ""
				Celda = ""

				sinput="<pre style='display:inline'>&#09;&#09;</pre> TOTAL AJUSTE: "
				Celda = ofv.gencelda("","","","","","","",sInput,"si")

				sinput=  FormatNumber(TSUBTOTALPESOS,2)
				Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

				sinput=  FormatNumber(TSUBTOTALPESOSBCRA,2)
				Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

				sinput= TSUBTOTALCANTIDAD
				Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")
				Fila = ofv.genrow("","","","","25","","",Celda,"si")
				FormularioD = FormularioD & fila
				Fila = ""
				Celda = ""
			END IF


				Fila = ""
				Celda = ""

				sinput="<hr width='100%'>"
				Celda = ofv.gencelda("","vti colspan=3 ","","","","","",sInput,"si")
				Fila = ofv.genrow("","","","","25","","",Celda,"si")
				FormularioD = FormularioD & fila
				FormularioD = FormularioD & fila
				Fila = ""
				Celda = ""

				sinput="<pre style='display:inline'>&#09;&#09;</pre> TOTAL GENERAL: "
				Celda = ofv.gencelda("","","","","","","",sInput,"si")

				sinput= FormatNumber(TTOTALPESOS,2)
				Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

				sinput= FormatNumber(TTOTALPESOSBCRA,2)
				Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")

				sinput= TTOTALCANTIDAD
				Celda = Celda & ofv.gencelda("","zz style='text-align=right;'","","","","","",sInput,"si")
				Fila = ofv.genrow("","","","","25","","",Celda,"si")
				FormularioD = FormularioD & fila
				Fila = ""
				Celda = ""

				Formulario = Formulario & ofv.GenTabla("","","","80%","","0","","0","1","","",FormularioD,"si")

		END IF

		IF (CSTR(REQUEST.QueryString("LISTADO")) = "6") THEN
			'wwwheader = chr(13) & " <TR ><TD  class=tt width=60% >DESCRIPCION</TD><TD  class=tt width=20% >IMPORTE</TD><TD  class=tt width=20% >CANTIDAD</TD></tr>" & chr(13)
			CANTREGISTROS = REQUEST.QueryString("CANTREGISTROS")

			IF (REQUEST.QueryString("LISTADO") = "6") THEN
				TITULO = "<TR><TD><H5>RANKING POR IMPORTE<H5></TD></TR>"

				IF (CANTREGISTROS = 0) THEN CANTREGISTROS = 5 END IF

				strsql = " SELECT TOP " & CANTREGISTROS & CHR(13)
				strsql = strsql & " B.OBJIDORTRACER AS RAIZ " & CHR(13)
				strsql = strsql & " ,OT.TRCOBSERV AS DESCRIPCION " & CHR(13)
				strsql = strsql & " ,B.MONTO " & CHR(13)
				strsql = strsql & " ,B.SUCURSAL AS SECTOR " & CHR(13)
				strsql = strsql & " ,B.EVENTORAIZ " & CHR(13)
				strsql = strsql & " ,ISNULL(L1.DESCRIPCION, '') + ' // ' + ISNULL(L2.DESCRIPCION, '') AS LN_DESCRIPCION " & CHR(13)
				strsql = strsql & " ,ISNULL(TE1.DESCRI,'') + ' // ' + ISNULL(TE2.DESCRI, '') AS E_DESCRIPCION " & CHR(13)
				strsql = strsql & " ,C.DESCRI AS CARACTER " & CHR(13)
				strsql = strsql & " FROM  " & CHR(13)
				strsql = strsql & " ORREGINFOBCRA A  " & CHR(13)
				strsql = strsql & " INNER JOIN ORREGINFOBCRADET B ON A.CODLOTE = B.CODLOTE " & CHR(13)
				strsql = strsql & " LEFT JOIN ORLINNEGO2 L2 ON L2.CODBCRA = B.LINEANEGOCIO " & CHR(13)
				strsql = strsql & " LEFT JOIN ORLINNEGO1 L1 ON L2.NIVEL1 = L1.OBJID " & CHR(13)
				strsql = strsql & " LEFT JOIN ORTIPOCATEGO TE2 ON TE2.CODBCRA = B.TIPOEVENTO " & CHR(13)
				strsql = strsql & " LEFT JOIN ORTIPOEVENTO TE1 ON TE1.OBJID = TE2.TIPOEVENTO " & CHR(13)
				strsql = strsql & " LEFT JOIN ORTIPOCARACTER C ON C.CODBCRA = B.CARACTER " & CHR(13)
				strsql = strsql & " LEFT JOIN ORTRACER OT ON OT.OBJID = B.OBJIDORTRACER " & CHR(13)
				strsql = strsql & " WHERE  " & CHR(13)
				strsql = strsql & " A.OBJID = " & REQUEST.QueryString("TRIMESTRE") & CHR(13)
				strsql = strsql & " AND B.ESTADO = 'N' " & CHR(13)
				'strsql = strsql & " AND B.SECUENCIA = 0 " & CHR(13)
				strsql = strsql & " AND B.CARACTER IN ('01', '03', '05') " & CHR(13)
				strsql = strsql & " ORDER BY " & CHR(13)
				strsql = strsql & " B.MONTO DESC "

			END IF
			
			'response.Write(chr(13) & chr(13))
			'response.Write(strsql)
			'response.Write(chr(13) & chr(13))

			set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


			Formulario = Formulario & TITULO

			posi = 0
			TTOTALPESOS = 0
			TSUBTOTALPESOS = 0
			TTOTALCANTIDAD = 0
			TSUBTOTALCANTIDAD = 0

			IF(NOT RS.EOF) THEN
				Do While (Not rs.EOF)

					sinput= "EVENTO NRO: " & RS("RAIZ")
					Celda = ofv.gencelda("","XX WIDTH='50%'","","","","","",sInput,"si")
					sinput= "IMPORTE: " & FormatNumber(CDBL(RS("MONTO")),2)
					Celda = Celda & ofv.gencelda("","XX WIDTH='50%'","","","","","",sInput,"si")
					Fila = ofv.genrow("","","","","25","","",Celda,"si")
					FormularioD = fila
					TTOTALPESOS = TTOTALPESOS + CDBL(RS("MONTO"))



					sinput="CARACTER: " & RS("CARACTER")
					Celda = ofv.gencelda("","XX colspan=2","","","","","",sInput,"si")
					Fila = ofv.genrow("","","","","25","","",Celda,"si")
					FormularioD = FormularioD & fila


					sinput="AREA / C. COSTO: " & RS("SECTOR")
					Celda = ofv.gencelda("","XX colspan=2","","","","","",sInput,"si")
					Fila = ofv.genrow("","","","","25","","",Celda,"si")
					FormularioD = FormularioD & fila

					sinput="LINEA NEGOCIO: " & RS("LN_DESCRIPCION")
					Celda = ofv.gencelda("","XX colspan=2","","","","","",sInput,"si")
					Fila = ofv.genrow("","","","","25","","",Celda,"si")
					FormularioD = FormularioD & fila

					sinput="EVENTO: " & RS("E_DESCRIPCION")
					Celda = ofv.gencelda("","XX colspan=2","","","","","",sInput,"si")
					Fila = ofv.genrow("","","","","25","","",Celda,"si")
					FormularioD = FormularioD & fila

					sinput="DESCRIPCION: " & RS("DESCRIPCION")
					Celda = ofv.gencelda("","XX colspan=2","","","","","",sInput,"si")
					Fila = ofv.genrow("","","","","25","","",Celda,"si")
					FormularioD = FormularioD & fila

					Formulario = Formulario & ofv.GenTabla("","XX style='border-style:solid;' WIDTH='600'","","","","0","","0","1","","",FormularioD,"si")
					Formulario = Formulario & "<BR>"
					POSI = POSI + 1
					RS.MOVENEXT()
				LOOP
				sinput= ""
				Celda = ofv.gencelda("","XX WIDTH='25%'","","","","","",sInput,"si")
				sinput= "IMPORTE"
				Celda = Celda & ofv.gencelda("","XX WIDTH='50%' ALIGN='RIGHT'","","","","","",sInput,"si")
				sinput= "#CASOS"
				Celda = Celda & ofv.gencelda("","XX WIDTH='25%' ALIGN='RIGHT'","","","","","",sInput,"si")
				Fila = ofv.genrow("","","","","25","","",Celda,"si")
				FormularioD = Fila

				sinput= "RANKING"
				Celda = ofv.gencelda("","XX WIDTH='25%' ALIGN='RIGHT'","","","","","",sInput,"si")
				sinput= FormatNumber(TTOTALPESOS,2)
				Celda = Celda & ofv.gencelda("","XX WIDTH='50%' ALIGN='RIGHT'","","","","","",sInput,"si")
				sinput= POSI
				Celda = Celda & ofv.gencelda("","XX WIDTH='25%' ALIGN='RIGHT'","","","","","",sInput,"si")
				Fila = ofv.genrow("","","","","25","","",Celda,"si")
				FormularioD = FormularioD & Fila

				DIVISION = "<TABLE WIDTH='600' ><TR><TD ALIGN='RIGHT'>"
				DIVISION = DIVISION & ofv.GenTabla("","XX style='border-style:solid;' WIDTH='300'","","","","0","","0","1","","",FormularioD,"si")
				DIVISION = DIVISION & "</TD></TR></TABLE>"

				Formulario = Formulario & DIVISION
			END IF
		END IF

		sHTML = sHTML & Formulario

		call ofv.cerrarconsulta(rs)
		call ofv.cerrarconn(cn)

		sHTML = sHTML & "</body></html>"
	end if
Response.Write sHTML
%>