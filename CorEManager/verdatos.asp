<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0
	Response.AddHeader "Content-Type", "text/html; charset=iso-8859-1"
	'RESPONSE.WRITE REQUEST.FORM

	tipoG=request.form("miSelect")
	fdesde=request.form("fechaDesde")
	fhasta=request.form("fechaHasta")
    TIPOGRAF=request.form("TIPOGRAF")

	fdesde=cdbl(fechaConvert(fdesde))  - CLng(CDate("01-01-1900"))
	fhasta=cdbl(fechaConvert(fhasta)) - CLng(CDate("01-01-1900"))


	If tipoG="" Then tipoG="DOCMASLEIDOS"


	 IF tipoG="WKF" or tipoG="WKFPASO30" THEN
		Set cn = ofv.conectar(ofv.strconn5)
	 ELSE
    	Set cn = ofv.conectar(ofv.strconn0)
    END IF

	Select Case tipoG
		Case "DOCMASLEIDOS"
        tipo=TIPOGRAF
		tit="LOS 10 TIPOS DE DOCUMENTOS MAS CONSULTADOS"
		TOTAL="CONSULTADOS"
		sql="select top 10 sum(b.cant) as cant ,max(b.tipobjeto) as tipobjeto,max(c.descri) as descriT from"
		sql= sql & "("
		sql= sql & "  select  CODOBJETO, TIPOBJETO, AVG(HFINAL - HINICIO) AS PROM, SUM(1) AS CANT "
		sql= sql & " from objalogdet where tacc = 'CONSEN' "
		sql= sql & " AND (FINICIO > " & fdesde & " AND FINICIO < " & fhasta &")"

		sql= sql & " GROUP BY CODOBJETO,TIPOBJETO"

		sql= sql & ") b"
		sql= sql & " inner join TIPOBJETO c on (c.tipobjeto=b.TIPOBJETO)"
		sql= sql & " group by b.TIPOBJETO"
		sql= sql & " order by cant desc"

		Case "DOCPORESTADOS"
		tipo=TIPOGRAF
		tit="DOCUMENTOS POR ESTADO"
		TOTAL="TOTAL DE DOCUMENTOS"
		sql="  SELECT MAX(C.ESTADO) AS DESCRIT,SUM(1) AS CANT  "

		Sql= sql & "	FROM OBJPRTY B"
		Sql= sql & "	INNER JOIN ESTADOS C ON (C.OBJID=B.ESTADO)"
		Sql= sql & " where  (Fecemis > " & fdesde & "  AND fecemis < " & fhasta & ")"
		Sql= sql & "GROUP BY B.ESTADO"

    CASE "DOCPORAUTOR"
     tipo=TIPOGRAF
	 tit="DOCUMENTOS POR AUTOR"
	 TOTAL="TOTAL DE DOCUMENTOS"
      SQL= "SELECT AUTOR AS DESCRIT,SUM(1) AS CANT FROM OBJETOS A "
	  Sql= sql & "INNER JOIN OBJPRTY B ON (B.codobjeto=A.codobjeto)"
	  Sql= sql & " where  (objun=1 and Fecemis > " & fdesde & "  AND fecemis < " & fhasta & ")"
	  Sql= sql & "GROUP BY AUTOR "
      Sql= sql & " ORDER BY CANT DESC "


    CASE "DOCPORSUBTEMA"
      tipo=TIPOGRAF
      tit="DOCUMENTOS POR SUBTEMA"
	  TOTAL="TOTAL DE DOCUMENTOS"
      SQL= "select TOP 10 max(TEMADESCRI) +' /' + max(STEMADESCRI) as descrit,sum(1) as cant "
	  Sql= sql & " from "
	  Sql= sql & " CSRELPATHDESCRIex  "
	  Sql= sql & " WHERE OBJUN=1 "
	  Sql= sql & " group by stema "
	  Sql= sql & " order by cant desc "


	CASE "WKF"
	      tipo=TIPOGRAF
	      tit="WKF POR ESTADOS"
		  TOTAL="TOTALES POR ESTADOS"

	      SQL= "       select COUNT(*) AS CANT , MAX('CANCELADOS') AS DESCRIT from CKTVIEWEX where  fin like '%cancelado%'"
	      Sql= sql &  " and DATEDIFF(dd,convert(date, '01/01/1900'),DBO.ARMAFECHA(COMIENZO,'',1)) > " & fdesde
		  Sql= sql &  " AND "
		  Sql= sql & "DATEDIFF(dd,convert(date, '01/01/1900'),DBO.ARMAFECHA(COMIENZO,'',1)) < " & fhasta



		  Sql= sql & " UNION "

		  Sql= sql & " select COUNT(*) AS CANT , MAX('EN CURSO') AS DESCRIT from CKTVIEWEX where  fin='NO' AND PSCFIN <> 'NO' "
		  Sql= sql &  " and DATEDIFF(dd,convert(date, '01/01/1900'),DBO.ARMAFECHA(COMIENZO,'',1)) > " & fdesde
		  Sql= sql &  " AND "
		  Sql= sql & "DATEDIFF(dd,convert(date, '01/01/1900'),DBO.ARMAFECHA(COMIENZO,'',1)) < " & fhasta
		  Sql= sql & " UNION "

		  Sql= sql & "  select COUNT(*) AS CANT , MAX('FINALIZADOS') AS DESCRIT  from CKTVIEWEX where  fin not like  '%cancelado%' and fin <>'NO'"
		  Sql= sql &  " and DATEDIFF(dd,convert(date, '01/01/1900'),DBO.ARMAFECHA(COMIENZO,'',1)) > " & fdesde
		  Sql= sql &  " AND "
		  Sql= sql & " DATEDIFF(dd,convert(date, '01/01/1900'),DBO.ARMAFECHA(COMIENZO,'',1)) < " & fhasta


   CASE "WKFPASO30"
   	      tipo=TIPOGRAF
   	      tit="WKF PASO 30"
   		  TOTAL="TOTALES"

   	      sql= "selecT count(a1.total) as cant,isnull(max('RECHAZADO'),'RECHAZADO') + ' : ' + cast(count(a1.total)as nvarchar) as descrit from ("

		 Sql= sql & " select count(*) as total from CKTVIEWEX a "
		 Sql= sql & " inner join wkfpasos b on( a.pscmodel=b.WKFPASCOD and  a.model=b.wkfcod and b.orden=30)"
		 Sql= sql &  " where  A.FIN='NO' AND DATEDIFF(dd,convert(date, '01/01/1900'),DBO.ARMAFECHA(COMIENZO,'',1)) >= " & fdesde
		 Sql= sql &  " AND "
		 Sql= sql & " DATEDIFF(dd,convert(date, '01/01/1900'),DBO.ARMAFECHA(COMIENZO,'',1)) <= " & fhasta
		 Sql= sql & " group by SEGCKT) a1 "

		  Sql= sql & "UNION"

		 Sql= sql & " selecT count(a1.total) as cant,isnull(max('APROBADO'),'APROBADO')  + ' : ' + cast(count(a1.total)as nvarchar) as descrit  from ("

		  Sql= sql & "select count(*) as total from CKTVIEWEX a "
		 Sql= sql & " inner join wkfpasos b on( a.pscmodel=b.WKFPASCOD and  a.model=b.wkfcod and b.orden <> 30)"
		 Sql= sql &  " where A.FIN <> 'NO' AND DATEDIFF(dd,convert(date, '01/01/1900'),DBO.ARMAFECHA(COMIENZO,'',1)) >= " & fdesde
		 Sql= sql &  " AND "
		 Sql= sql & " DATEDIFF(dd,convert(date, '01/01/1900'),DBO.ARMAFECHA(COMIENZO,'',1))  <= " & fhasta
		 Sql= sql & " group by SEGCKT) a1 "

	End Select


'RESPONSE.WRITE SQL
	Set rsx1 = OFV.crearconsultaEx(sql, cn, 1, "")

	shtml="{"
	etiqueta=chr(34) & "etiquetas" & chr(34) & ":["
	dato=chr(34) & "datos" & chr(34) & ":["
	TIT=chr(34) & "titulo" & chr(34) & ":" & chr(34) & TIT & chr(34)
	total=chr(34) & "total" & chr(34) & ":" & chr(34) & total & chr(34)
	tipo=chr(34) & "tipo" & chr(34) & ":" & chr(34) & tipo & chr(34)
	delim=""
	Do While Not rsx1.eof
	    IF tipoG="DOCPORAUTOR" THEN
	    XXX=GETINFOUSER(rsx1("descriT"))
	     etiqueta=etiqueta & delim & chr(34) & session("SGU_nombre") &  chr(34)
	    ELSE
		 etiqueta=etiqueta & delim & chr(34) & rsx1("descriT") &  chr(34)
		END IF
		' "etiquetas": ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio"],
		' "datos": [65, 59, 80, 81, 56, 55]
		dato=dato & delim & rsx1("cant")
		delim=","
		rsx1.movenext
	Loop
	dato=dato & "]"
	etiqueta=etiqueta & "]"
	shtml=shtml & etiqueta & "," & dato & "," & TIT & "," & total & "," & tipo & "}"
End If
response.write shtml



Function fechaConvert(wkfecha)
	Nfecha=CDate(wkfecha)
	Nfecha=DateSerial(Year(wkfecha),Month(wkfecha),Day(wkfecha))
	fechaConvert=Nfecha
End Function

%>