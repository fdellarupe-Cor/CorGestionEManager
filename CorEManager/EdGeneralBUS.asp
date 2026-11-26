<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

If ok <> "Y" Then
	sHTML = ok
Else
	response.expires=0


	sHTML = ofv.FormHeader(session("FormN"))

	'----------------------------------------------

	shtml = shtml & CorPnlFnc("N","")

	sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

	'-------------------------------------------------


	CantRegAMover = CDbl(request.querystring("CantRegAMover"))
	CantRegMover  = CDbl(request.querystring("CantRegMover"))
	If CantRegMover > 0 Then CantRegAMover = CantRegMover
	CantRegAMostrar = 10
	vuelta = ofv.vueltaaspEx("../",CantRegAMostrar,CantRegAMover)
	vuelta2 = ofv.vueltaaspEx("",CantRegAMostrar,CantRegAMover)
	OpcAgregar = request.querystring("OpcAgregar")

	Set cn = ofv.conectar(ofv.strconn2)

	ofv.ObtenerAtributos Session("Form2"),""
	bModificar = ofv.AModificar
	If session("Nomodificar") = "S" Then bModificar = False

	nCodigo = request.querystring("Codobjeto")
	sTipo = request.querystring("Tipobjeto")

	Sopcion = request.querystring("Sopcion")
	VASP = request.querystring("VASP")

	If VASP <> "" Then
		SESSION("VASP") = VASP
	Else
		VASP = SESSION("VASP")
	End If

	Select Case VASP


		Case "RP"

		VASPN = "CORREQPERFILPROP.ASP"

		Case Else

		'''''''''''''VASPN = "EDgeneral.ASP"
		VASPN = "EDgeneralBSSS.ASP"

	End Select

	If Sopcion = "A" Then
		SBuscar = request.querystring("SLetra")
		Session("Sletra") = SBuscar
	ElseIf Sopcion = "B" Then
		SBuscar = request.querystring("SLetra")
		Session("Sletra") = SBuscar
	ElseIf Sopcion = "I" Then
		SBuscar = request.form("alfa")
		Session("Sletra") = SBuscar
	Else
		SBuscar = Session("Sletra")
	End If

	If sbuscar <> "" Then sbuscar = Replace(sbuscar,"'","")

	If CantRegAMover > 0 Then
		NPagina =  CInt((CantRegAMover / CantRegAMostrar)) + 1
	Else
		NPagina = 1
	End If

	ReDim nivel(27)
	nivel(0) = "A"
	nivel(1) = "B"
	nivel(2) = "C"
	nivel(3) = "D"
	nivel(4) = "E"
	nivel(5) = "F"
	nivel(6) = "G"
	nivel(7) = "H"
	nivel(8) = "I"
	nivel(9) = "J"
	nivel(10) = "K"
	nivel(11) = "L"
	nivel(12) = "M"
	nivel(13) = "N"
	nivel(14) = "Ñ"
	nivel(15) = "O"
	nivel(16) = "P"
	nivel(17) = "Q"
	nivel(18) = "R"
	nivel(19) = "S"
	nivel(20) = "T"
	nivel(21) = "U"
	nivel(22) = "V"
	nivel(23) = "W"
	nivel(24) = "X"
	nivel(25) = "Y"
	nivel(26) = "Z"


	sAccion = VASPN & "?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&ESQ=0&OPCA=SI"
	sInput  = "<input  type=SUBMIT  size=10%  value='No Reemplaza' class=bt >"
	sInput =  ofv.genform("tab3","a style=margin:0",sAccion,"","post","",sInput,"si")
	Celda   =  ofv.gencelda("","","center","","","","",sInput,"si")
	Formulario4 =  ofv.genrow("","","","","","","",Celda,"si")

	sAccion  = VASPN & "?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&SLetra=A&Sopcion=A&VASP=CM" & "&Objrem=" & nObjrem & "&nXDescri=" & nXDescri
	sInput = "<input  type=SUBMIT  size=10%  value=' Volver ' class=bt >"
	sInput =  ofv.genform("tab3","a style=margin:0",sAccion,"","post","",sInput,"si")
	Celda  =  ofv.gencelda("","","center","","","","",sInput,"si")
	Formulario3   = ofv.genrow("","","","","","","",Celda,"si")


	sAccion    = "EDgeneralBus.asp?Sopcion=I&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
	'response.write(chr(13) & saccion & chr(13))

	SCOMSEP = ""
	scomK   = ""

	For i=0 To 26
		scomK = scomK & SCOMSEP & nivel(i)
		SCOMSEP = Chr(9)
	Next

	scomD = scomK

	sLbus = UCase(SBuscar)
	If  Len(sLbus) > 1 Then sLbus = Left(sLbus,1)

	sInput = ofv.gencomboFX("tab1","alfa",sLbus,scomK,scomD,"")
	sInput2 = "<input  type=SUBMIT  value='Buscar' class=bt >"
	sInput =  ofv.genform("tab1","a style=margin:0;display:flex;gap:5px;",sAccion,"","post","",sInput & sInput2,"si")
	Celda  = ofv.gencelda("","ut","center","","","","",sInput,"si")
	Formulario   = ofv.genrow("","","","","","","",Celda,"si")

	sAccion     = "EDgeneralBus.asp?Sopcion=I&Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
	sInput = ofv.GenerarInput("alfa",sbuscar,"agregar","","70","dt","")
	sInput2 = "<input  type=SUBMIT  value='Buscar' class=bt >"
	sInput =  ofv.genform("tab2","a style=margin:0;display:flex;gap:5px;",sAccion,"","post","",sInput & sInput2,"si")
	Celda  = ofv.gencelda("","ut","center","","","","",sInput,"si")
	Formulario2   = ofv.genrow("","","","","","","",Celda,"si")




	TABLA0 = ofv.GenTabla("","aaa","","","","0","","0","0","","",Formulario4,"si")
	TABLA1 = ofv.GenTabla("","aaa","","","","0","","0","0","","",Formulario2,"si")
	TABLA2 = ofv.GenTabla("","aaa","","","","0","","0","0","","",Formulario3,"si")
	TABLA3 = ofv.GenTabla("","aaa","","","","0","","0","0","","",Formulario,"si")

	celdas = ofv.GenCelda("","ut","center","25%","","","",TABLA0,"si")
	celdas = celdas & ofv.GenCelda("","ut","center","25","","","",TABLA3,"si")
	celdas = celdas & ofv.GenCelda("","ut","center","25","","","",TABLA1,"si")
	celdas = celdas & ofv.GenCelda("","ut","center","25","","","",TABLA2,"si")
	filas  =  ofv.GenRow("","","","","10","","",celdas,"si")

	ftab1  = ofv.GenTabla("","aaa style='border-color:" & Session("CMFPNTIT") & "' ","","100%","","0","","0","0","","",filas,"si")


	celdas = ofv.GenCelda("","ttt colspan=3","CENTER","","","","","LISTA DE DOCUMENTOS","si")
	celdas = celdas & ofv.GenCelda("","ttt","right","","","","","PAGINA&nbsp;" & NPagina,"si")
	filas  =  ofv.GenRow("","","","","20","","",celdas,"si")


	celdas = ofv.GenCelda("","ttr","","","","","","Descripción","si")
	celdas = celdas & ofv.GenCelda("","ttr","","","","","","Identificación","si")
	celdas = celdas & ofv.GenCelda("","ttr","","","","","","Tipo Esquema","si")
	celdas = celdas & ofv.GenCelda("","ttr","","","","","","Versión","si")
	filas  =  filas & ofv.GenRow("","","","","25","","",celdas,"si")


	'StrSql = "Select codobjeto,objDESCRI,descri,codiden,version,nrover From csobjetosRDOAPL  where estado=3 and aplic in('EM', 'PL', 'GLB') and objDESCRI like '" & SBuscar & "%' Order By DESCRI"
	If session("base")="K" Then
		StrSql = "Select codobjeto,objDESCRI,csobjetosRDOAPL.descri,codiden,version,nrover From csobjetosRDOAPL  inner join tipobjeto on (csobjetosRDOAPL.tipobjeto=tipobjeto.tipobjeto and tipocpt='K' )   where csobjetosRDOAPL.estado=3 and  csobjetosRDOAPL.aplic in('EM', 'PL','GLB') and objDESCRI like '" & SBuscar & "%' Order By csobjetosRDOAPL.DESCRI"
	Else
		StrSql = "Select codobjeto,objDESCRI,csobjetosRDOAPL.descri,codiden,version,nrover From csobjetosRDOAPL  inner join tipobjeto on (csobjetosRDOAPL.tipobjeto=tipobjeto.tipobjeto and tipocpt <> 'K' )   where csobjetosRDOAPL.estado=3 and  csobjetosRDOAPL.aplic in('EM', 'PL','GLB') and objDESCRI like '" & SBuscar & "%' Order By csobjetosRDOAPL.DESCRI"
	End If

	''response.write strsql
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If Not rs.eof Then rs.move(CantRegAMover)
	posI = 0
	xTrColor = "#FFFFFF"
	Do While (Not rs.EOF) And (posi < CantRegAMostrar)
		nID = rs(0)

		''response.write VASPN & " -197- " & chr(13)
		sHref  = VASPN & "?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo & "&ESQ=" & rs(0) & "&OPCA=SI"
		sLink  = ofv.GenLink("","nt",sHref,"","","",rs(1),"si")
		celdas = ofv.GenCelda("","ut","","60%","","","",sLink,"si")

		sLink  = ofv.GenLink("","nt",sHref,"","","",rs(3),"si")
		celdas = celdas & ofv.GenCelda("","ut","","15%","","","",sLink,"si")

		sLink  = ofv.GenLink("","nt",sHref,"","","",rs(2),"si")
		celdas = celdas & ofv.GenCelda("","ut","","20%","","","",sLink,"si")

		sLink  = ofv.GenLink("","cc",sHref,"","","",rs(4) & "&nbsp;" & rs(5),"si")
		celdas = celdas & ofv.GenCelda("","nt","","5%","","","",sLink,"si")

		filas  =  filas & ofv.GenRow("","aaa bgcolor=" & xTrColor ,"","","15","","",celdas,"si")
        If xTrColor = "#FFFFFF" Then
				xTrColor = "#F0F0F0"
			Else
				xTrColor = "#FFFFFF"
			End If
		posi = posi + 1
		rs.movenext
	Loop


	ftab2 = ofv.GenTabla("","aaa style='border-color:" & Session("CMFPNTIT") & "' ","","100%","","0","","-1","-1","","",filas,"si")




	Celda = ofv.GenCelda("","","center","","","","",ftab1,"si")
	FilaS = ofv.GenRow("","","","","","","",Celda,"si")

	Celda = ofv.GenCelda("","","center","","","","",ftab2,"si")
	FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

	sHTML = sHTML & ofv.GenTabla("","aaa  ","","100%","","0","","0","0","","",filas,"si")


	sAccion = VASPN & "?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
	Celda = ofv.BotonVolver(sAccion)



	'--------------------------------------------------------------------------------
	sHTML = sHTML & "</div>"

	Celda = ofv.GenCelda("","vt","","60%","","","","&nbsp;&nbsp;","si")
	Fila = ofv.GenRow("","","","","","","",Celda,"si")
	sHTML = sHTML &  ofv.GenTabla("TCEX1","onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")

	'Estructura del Indice

	sHTML = sHTML & CorPnlIndx("N","")

	'--------------------------------------------------------------------------------

	sHTML   = sHTML & ofv.botoneraESSX("a",rs,CantRegAMover,CantRegAMostrar,"","","",nCodigo,sTipo,"no","no",Celda,"")

	Call ofv.cerrarconsulta(rs)
	Call ofv.cerrarconn(cn)



	sHTML = sHTML & "</body></html>"
End If

response.write sHTML

%>