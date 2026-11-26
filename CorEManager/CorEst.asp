<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0

	' XXXX = GENLOGEXT("EDIDET")



	nCodigo = request.querystring("Codobjeto")
	sTipo = request.querystring("Tipobjeto")
	vuelta = ofv.vueltaaspEx("../","0","0")'  &"_Codobjeto=" & nCodigo & "_Tipobjeto=" & sTipo
	ofv.ObtenerAtributos Session("Form2"),""
	bModificar = ofv.AModificar
	bSuprimir = ofv.AEliminar

	Set cn = ofv.conectar(ofv.strconn2)
	Set cnx = ofv.conectar(ofv.strconn5)



	sHTML = ofv.FormHeader(session("FormN"))

	'----------------------------------------------

	shtml = shtml & CorPnlFnc("N","")

	sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"

	'-------------------------------------------------

	TotalVesiones=0
	sidz=  BuscarIds(nCodigo)
	strsql ="select sum(1) from objprty where codobjeto in " & sidz & "  and version='V' group by version"
	' RESPONSE.WRITE STRSQL
	Set rs2 = OFV.crearconsultaEx(strsql, cn, 1, "")
	If Not rs2.eof Then TotalVesiones=rs2(0)
	Call ofv.cerrarconsulta(rs2)
    celda =  ofv.GenCelda("","ttt colspan=3","","","","","","TOTAL DE VERSIONES","si")
	fila =  ofv.GenRow("","","","","","","",celda,"si")
	celda =  ofv.GenCelda("","ut colspan=3","","","","","",TotalVesiones,"si")
	fila =  FILA & ofv.GenRow("","","","","","","",CELDA,"si")

	strsql ="select distinct autor  from objprty where codobjeto in " & sidz & " "

	Set rs2 = OFV.crearconsultaEx(strsql, cn, 1, "")
	Sowner=""
	 celda =  ofv.GenCelda("","ttt colspan=3","","","","","","HISTORIAL DE AUTORES","si")
	 fila =   FILA & ofv.GenRow("","","","","","","",celda,"si")

	Do While Not rs2.eof
		xx=getinfouser(rs2(0))
		Sowner=Sowner & session("SGU_nombre")
		 celda = ofv.GenCelda("","ut   colspan=3","","1%","","","",Sowner,"si")
		 fila =  FILA &  ofv.GenRow("","","","","","","",celda,"si")

		rs2.movenext
	Loop
	Call ofv.cerrarconsulta(rs2)


	strsql =   "SELECT A.CODOBJETO, A.TIPOBJETO, A.FULTACC, A.ACON, AVG(ISNULL(PROM,0)) AS PROMF, SUM(ISNULL(CANT,0)) AS CANTF, "
	strsql = strsql & " MAX(ESTADO) AS ESTADO , MAX(VERSION) AS VERSION, "
	strsql = strsql & " MAX(NROVER) AS NROVER, MAX(FECVIGEN) AS FECVIGEN, "
	strsql = strsql & " MAX(CODIDEN) AS CODIDEN, MAX(C.DESCRI) AS DESCRI  "
	strsql = strsql & " FROM OBJALOG A LEFT JOIN "
	If session("UN") = 1 Then
		strsql = strsql & " (select CODOBJETO, TIPOBJETO, FINICIO, AVG(HFINAL - HINICIO) AS PROM, SUM(1) AS CANT "
		strsql = strsql & " from objalogdet "
		strsql = strsql & " where tacc = 'CONSEN' "
		strsql = strsql & " GROUP BY CODOBJETO,TIPOBJETO,FINICIO) B "
	Else
		strsql = strsql & " (select objalogdet.CODOBJETO, objalogdet.TIPOBJETO, objalogdet.FINICIO, AVG(HFINAL - HINICIO) AS PROM, SUM(1) AS CANT "
		strsql = strsql & " from objalogdet "
		strsql = strsql & " inner join objalog c on c.tipobjeto='SKRLOG' and c.usuario=objalogdet.usuario and c.codemp=" & session("UN")
		strsql = strsql & " where tacc = 'CONSEN' "
		strsql = strsql & " GROUP BY objalogdet.CODOBJETO,objalogdet.TIPOBJETO,objalogdet.FINICIO) B "
	End If

	strsql = strsql & " ON A.CODOBJETO = B.CODOBJETO AND A.TIPOBJETO = B.TIPOBJETO "
	strsql = strsql & " INNER JOIN OBJETOS C ON A.CODOBJETO = C.CODOBJETO  "
	strsql = strsql & " INNER JOIN OBJPRTY D ON A.CODOBJETO = D.CODOBJETO  "
	If SESSION("BASE")="K" Then
		strsql = strsql & " INNER JOIN tipobjeto Df ON C.tipobjeto = Df.tipobjeto and df.TIPOCPT='K'  "
	Else
		strsql = strsql & " INNER JOIN tipobjeto Df ON C.tipobjeto = Df.tipobjeto and df.TIPOCPT <> 'K'  "
	End If

	strsql = strsql & "WHERE A.CODOBJETO > 0  AND ESTADO in (3)  "

	strsql=  strsql & " and A.CODOBJETO in " & sidz
	strsql = strsql & " GROUP BY A.CODOBJETO,A.TIPOBJETO,FULTACC, ACON "

	'RESPONSE.WRITE STRSQL & "<BR>" & CHR(13)
	Set rs2 = OFV.crearconsultaEx(strsql, cn, 1, "")
	totalConsultas=0
	If Not rs2.eof Then
		totalConsultas=rs2("ACON")
	End If
	Call ofv.cerrarconsulta(rs2)

	celda =  ofv.GenCelda("","ttt colspan=3","","","","","","TOTAL DE CONSULTAS","si")
    Fila =  FILA & ofv.GenRow("","","","","","","",celda,"si")
	celda =  ofv.GenCelda("","ut colspan=3","","","","","",totalConsultas,"si")
	fila =  FILA & ofv.GenRow("","","","","","","",CELDA,"si")
    celda =  ofv.GenCelda("","ttt ","","20%","","","","HISTORIAL DE AUTORIZADORES","si")
    celda =  CELDA & ofv.GenCelda("","ttt ","","50%","","","","DOCUMENTO","si")
    celda =  CELDA & ofv.GenCelda("","ttt ","","30%","","","","FECHA DE VIGENCIA","si")
    Fila =  FILA & ofv.GenRow("","","","","","","",celda,"si")
	strsql = "select distinct pscresponsable,objcod,objtipo from cktviewex where objcod in "  & sidz &" order  by objcod"
	Set rs2 = OFV.crearconsultaEx(strsql, cnx, 1, "")
	xTrColor = "#FFFFFF"
	idant=0
	fi = CLng(CDate("01-01-1900"))
	'dVigencia = CDate(CLng(rs(6)) + fi)
	'fechadorCor(dVigencia, "")
	Do While Not rs2.eof
        idant=rs2("objcod")
        idI=rs2("objcod")
        sESQ=EsqnomPx(RS2(2), RS2(1), ofv)

        sql="select fecvigen,fecbaja from objprty where codobjeto=" & idant
        Set rs3 = OFV.crearconsultaEx(sql, cn, 1, "")
        if not rs3.eof then
          IF rs3("fecbaja")=0 THEN
           dFECHAS="<I><B>VIGENCIA :<B/>" & fechadorCor(CDATE(rs3("fecvigen")+ FI), "")  &"</I>"
          ELSE
          dFECHAS="<I><B>VIGENCIA :</B>" & fechadorCor(CDATE(rs3("fecvigen")+ FI), "") &  "<B> FECHA DE BAJA :</B>" & fechadorCor(CDATE(rs3("fecbaja")+FI), "") &"</I>"
          END IF

        end if
        Call ofv.cerrarconsulta(rs3)
       do while idant=idi
		xx=getinfouser(rs2(0))
		UsrPaso= session("SGU_nombre")


		celda =  ofv.GenCelda("","ut","","","","","",UsrPaso,"si")
		celda =  celda & ofv.GenCelda("","ut","","","","","",sESQ,"si")

         CELDA =  celda & ofv.GenCelda("","ut","","","","","",dFECHAS,"si")

	     fila =  FILA & ofv.GenRow("","aaa bgcolor=" & xTrColor ,"","","","","",CELDA,"si")

		rs2.movenext
		if not rs2.eof then
		 idI=rs2("objcod")
		else
		 idI=0
		end if
		sESQ=""
	   loop
	   If xTrColor = "#FFFFFF" Then
	   		  xTrColor = "#F0F0F0"
	   		  Else
	   		  xTrColor = "#FFFFFF"
		End If
	Loop



	Call ofv.cerrarconsulta(rs2)





 sHTML = sHTML  & ofv.GenTabla("","aaa","","98%","","0","","0","0","","",FILA,"si")
'
'--------------------------------------------------------------------------------
	  sHTML = sHTML & "</div>"

	  Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
	  Fila = ofv.GenRow("","","","","","","",Celda,"si")
	  sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

	  'Estructura del Indice

	  sHTML = sHTML & CorPnlIndx("N","")

'--------------------------------------------------------------------------------
	sHTML = sHTML & ofv.botoneraESSX("0","",CantRegAMover,CantRegAMostrar,"no","no",posi,"","","no",Agregar,Listar,"")


	Call ofv.cerrarconn(cn)
	Call ofv.cerrarconn(cnx)


	sHTML = sHTML & "</body></html>"
End If



response.write sHTML


Function BuscarIds(id)
	Dim strsql,rsid

	Dim sids


	sids=id
	delim=","
	strsql="select objrem from objetos where codobjeto=0" &id
	'response.write strsql
	Set rsid = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If Not rsid.eof Then
		Orem= rsid(0)
		sids =sids & delim & Orem
	End If
	Call ofv.cerrarconsulta(rsid)

	strsql="select objrem  from objetos a inner join objprty b "
	strsql= strsql & " on ( a.codobjeto=b.codobjeto )"
	strsql= strsql & " where a.codobjeto=0" & Orem

	' response.write "<br>" & strsql &"<BR>"
	Set rsid = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If  Not rsid.eof Then
		Do While Not rsid.eof
			Orem=rsid(0)


			If (Orem=0) Then Exit Do
			sids =sids & delim & Orem
			strsql="select objrem  from objetos a inner join objprty b "
			strsql= strsql & " on ( a.codobjeto=b.codobjeto )"
			strsql= strsql & " where a.codobjeto=0" & Orem
			Call ofv.cerrarconsulta(rsid)
			Set rsid = ofv.crearconsultaEx(StrSql,cn,1,parametros)

		Loop
	End If
	BuscarIds="(" & sids & ")"
End Function
%>