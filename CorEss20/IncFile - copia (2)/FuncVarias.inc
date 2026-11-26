<%
Class funcvarias

	Public FBotTab(20, 6)
	Public FTitTab(20, 6)
	Public ATXCOM(200, 4)

	Private conn
	Private m_oServerVar
	Private m_oRequestVar
	Private m_oResponseVar
	Private m_oSessionVar
	Private m_sCorLogo
	Private m_sCorFormHeader
	Private m_sCorSetMHeader
	Private m_sCorSetFHeader
	Private m_sCorSetCombo
	Private m_sTitCls
	Private m_Browser
	Private m_sModTitCls
	Private m_sModBot

	Private m_sExplorador, m_iwspacing, m_iwfont, m_iwsize
	Private m_iwresto, m_swbody
	Public strconn2,strconn0,strconn1,strconn3,strconn4,strconn5,strconn6,strconn7
	Public m_sConn0, m_sConn1, m_sConn2, m_sConn3, m_sConn4
	Public m_sConn5, m_sConn6, m_sConn7, m_sUV
	Public m_sConn8, m_sConn9, m_sConn10, m_sConn11, m_sConn12, m_sConn13, m_sConn14
	Private m_bEstado , m_bAgregar
	Private m_bEliminar , m_bConsultar
	Private m_bModificar
	Private m_bTrusted
	Private m_bDbUid
	Private m_bDbPwd
	Private m_bDbUidL
	Private m_bDbPwdL



	Private m_sErrorDesc, m_iStyler

	Private oEH 'As ErrorHandler.AppError
	Private oErr 'As ErrorHandler.ObjErr
	Public wbody
	Public wstyles

	Public FTitDat1
	Public FTitDat2
	Public FTitDat3
	Public FTitDat4
	Public FTitDat5
	Public FTitDat6
	Public FTitDat7

	Public FBotDat1
	Public FBotDat2
	Public FBotDat3
	Public FBotDat4
	Public FBotDat5
	Public FBotDat6
	Public FBotDat7


	Public FBbtnt
	Public FBbtnf

	Public FBbodt
	Public FBbodf

	Public FBmodo
	Public FBBarr
	Public FBFcnt


	Public GLBCOMBO
	Public GLBCACTU

	Public CMBFiller
	Public CMBFiller2
	Public CMBFillOPC

	Public Property Let Styler(ByVal iStyler)

	m_iStyler = iStyler

	End Property

	Public Property Get Styler()

	Styler = m_iStyler

	End Property



	Public Property Let TitCls(ByVal sTitCls)

	m_sTitCls = sTitCls

	End Property


	Public Property Let SetBrowser(ByVal sTitCls)

	m_Browser = sTitCls

	End Property


	Public Property Get TitCls()

	TitCls = m_sTitCls

	End Property

	Public Property Let ModTitCls(ByVal sModTitCls)

	m_sModTitCls = sModTitCls

	End Property


	Public Property Get ModTitCls()

	ModTitCls = m_sModTitCls

	End Property

	Public Property Let ModBot(ByVal sModBot)

	m_sModBot = sModBot

	End Property


	Public Property Get ModBot()

	ModBot = m_sModBot

	End Property


	Public Property Let ErrorDesc(ByVal sErrorDesc)

	m_sErrorDesc = sErrorDesc

	End Property


	Public Property Get ErrorDesc()

	ErrorDesc = m_sErrorDesc

	End Property


	Public Property Let CorLogo(ByVal sCorLogo)

	m_sCorLogo = sCorLogo

	End Property


	Public Property Get CorLogo()

	CorLogo = m_sCorLogo

	End Property

	Public Property Let CorFormHeader(ByVal sCorFormHeader)

	m_sCorFormHeader = sCorFormHeader

	End Property


	Public Property Get CorFormHeader()

	CorFormHeader = m_sCorFormHeader

	End Property

	Public Property Let CorSetMHeader(ByVal sCorSetMHeader)

	m_sCorSetMHeader = sCorSetMHeader

	End Property


	Public Property Get CorSetMHeader()

	CorSetMHeader = m_sCorSetMHeader

	End Property

	Public Property Let CorSetFHeader(ByVal sCorSetFHeader)

	m_sCorSetFHeader = sCorSetFHeader

	End Property




	Public Property Get CorSetCombo()

	CorSetCombo = m_sCorSetCombo

	End Property


	Public Property Let CorSetCombo(ByVal sCorSetcombo)

	m_sCorSetCombo = sCorSetcombo

	End Property




	Public Property Get CorSetFHeader()

	CorSetFHeader = m_sCorSetFHeader

	End Property


	Public Property Let ATrusted(ByVal bTrusted )

	m_bTrusted = bTrusted

	End Property

	Public Property Get ATrusted()

	ATrusted = m_bTrusted

	End Property

	Public Property Let ADbUid(ByVal bDbUid )

	m_bDbUid = bDbUid

	End Property

	Public Property Get ADbUid()

	ADbUid = m_bDbUid

	End Property
	Public Property Let ADbPwd(ByVal bDbPwd )

	m_bDbPwd = bDbPwd

	End Property

	Public Property Get ADbPwd()

	ADbPwd = m_bDbPwd

	End Property
	Public Property Let ADbUidl(ByVal bDbUidl )

	m_bDbUidL = bDbUidl

	End Property

	Public Property Get ADbUidl()

	ADbUidl = m_bDbUidL

	End Property
	Public Property Let ADbPwdl(ByVal bDbPwdl )

	m_bDbPwdL = bDbPwdl

	End Property

	Public Property Get ADbPwdl()

	ADbPwdl = m_bDbPwdL

	End Property



	Public Property Let AModificar(ByVal bModificar )

	m_bModificar = bModificar

	End Property

	Public Property Get AModificar()

	AModificar = m_bModificar

	End Property

	Public Property Let AConsultar(ByVal bConsultar )

	m_bConsultar = bConsultar

	End Property

	Public Property Get AConsultar()

	AConsultar = m_bConsultar

	End Property

	Public Property Let AEliminar(ByVal bEliminar )

	m_bEliminar = bEliminar

	End Property

	Public Property Get AEliminar()

	AEliminar = m_bEliminar

	End Property

	Public Property Let AAgregar(ByVal bAgregar )

	m_bAgregar = bAgregar

	End Property

	Public Property Get AAgregar()

	AAgregar = m_bAgregar

	End Property

	Public Property Let AEstado(ByVal bEstado )

	m_bEstado = bEstado

	End Property

	Public Property Get AEstado()

	AEstado = m_bEstado

	End Property

	Public Property Get UV()

	UV = m_sUV

	End Property

	Public Property Let UV(ByVal sUV)

	m_sUV = sUV

	End Property

	Public Property Get xstrconn0()

	strconn0 = m_sConn0

	End Property

	Public Property Let xstrconn0( sConn0)

	m_sConn0 = sConn0

	End Property

	Public Property Get xstrconn1()

	strconn1 = m_sConn1

	End Property

	Public Property Let xstrconn1(ByVal sConn1)

	m_sConn1 = sConn1

	End Property

	Public Property Get xstrconn2()

	strconn= m_sConn2

	End Property

	Public Property Let xstrconn2(ByVal sConn2)

	m_sConn= sConn2

	End Property

	Public Property Get xstrconn3()

	strconn3 = m_sConn3

	End Property

	Public Property Let xstrconn3(ByVal sConn3)

	m_sConn3 = sConn3

	End Property

	Public Property Get xstrconn4()

	strconn4 = m_sConn4

	End Property

	Public Property Let xstrconn4(ByVal sConn4)

	m_sConn4 = sConn4

	End Property

	Public Property Get xstrconn5()

	strconn5 = m_sConn5

	End Property

	Public Property Let xstrconn5(ByVal sConn5)

	m_sConn5 = sConn5

	End Property

	Public Property Get xstrconn6()

	strconn6 = m_sConn6

	End Property

	Public Property Let xstrconn6(ByVal sConn6)

	m_sConn6 = sConn6

	End Property

	Public Property Get xstrconn7()

	strconn7 = m_sConn7

	End Property

	Public Property Let xstrconn7(ByVal sConn7)

	m_sConn7 = sConn7

	End Property

	Public Property Get strconn8()

	strconn8 = m_sConn8

	End Property

	Public Property Let strconn8(ByVal sConn8)

	m_sConn8 = sConn8

	End Property


	Public Property Get strconn9()

	strconn9 = m_sConn9

	End Property

	Public Property Let strconn9(ByVal sConn9)

	m_sConn9 = sConn9

	End Property


	Public Property Get strconn10()

	strconn10 = m_sConn10

	End Property

	Public Property Let strconn10(ByVal sConn10)

	m_sConn10 = sConn10

	End Property

	Public Property Get strconn11()

	strconn11 = m_sConn11

	End Property

	Public Property Let strconn11(ByVal sConn11)

	m_sConn11 = sConn11

	End Property

	Public Property Get strconn12()

	strconn1= m_sConn12

	End Property

	Public Property Let strconn12(ByVal sConn12)

	m_sConn1= sConn12

	End Property

	Public Property Get strconn13()

	strconn13 = m_sConn13

	End Property

	Public Property Let strconn13(ByVal sConn13)

	m_sConn13 = sConn13

	End Property

	Public Property Get strconn14()

	strconn14 = m_sConn14

	End Property

	Public Property Let strconn14(ByVal sConn14)

	m_sConn14 = sConn14

	End Property



	Public Property Get Explorador()

	Explorador = m_sExplorador

	End Property

	Private Property Let Explorador(ByVal sExplorador)

	m_sExplorador = sExplorador

	End Property

	Public Property Get wspacing()

	wspacing = m_iwspacing

	End Property

	Private Property Let wspacing(ByVal iwspacing)

	m_iwspacing = iwspacing

	End Property

	Public Property Get wfont()

	wfont = m_iwfont

	End Property

	Private Property Let wfont(ByVal iwfont)

	m_iwfont = iwfont

	End Property

	Public Property Get wsize()

	wsize = m_iwsize

	End Property

	Private Property Let wsize(ByVal iwsize)

	m_iwsize = iwsize

	End Property

	Public Property Get wresto()

	wresto = m_iwresto

	End Property

	Private Property Let wresto(ByVal iwresto)

	m_iwresto = iwresto

	End Property

	'Public Property Get wbody()
	'
	'5 wbody = m_swbody
	'
	'End Property
	'
	'Private Property Let wbody(ByVal swbody)
	'
	'5    m_swbody = swbody
	'
	'End Property

	'Public Property Set Server(ByVal oServer )

	'    Set m_oServerVar = oServer

	'  Set oEH.Server = oServer

	'End Property

	'Public Property Get Server()

	'    Set Server = m_oServerVar

	'End Property

	'Public Property Set Request(ByVal oRequest )

	'    Set m_oRequestVar = oRequest

	'Set oEH.Request = oRequest

	'End Property

	'Public Property Get Request()

	'    Set Request = m_oRequestVar
	'
	'End Property
	'
	'Public Property Set Response(ByVal oResponse )
	'
	'    Set m_oResponseVar = oResponse

	'End Property

	'Public Property Get Response()

	'    Set Response = m_oResponseVar

	'End Property

	'Public Property Set Session(ByVal oSession )

	'    Set m_oSessionVar = oSession

	'End Property

	'Public Property Get Session()

	'    Set Session = m_oSessionVar

	'End Property

	Public Function Version()

		Version = App.EXEName & " Version: " & App.Major & "." & App.Minor & "." & App.Revision

	End Function

	Public Function AddBot(FBindex , dato1 , dato, dato3 , dato4 , dato5 , dato6 , dato7 )


		If IsNumeric(FBindex) Then

			If FBindex < 21 Then

				FBotTab(FBindex, 0) = dato1
				FBotTab(FBindex, 1) = dato2
				FBotTab(FBindex, 2) = dato3
				FBotTab(FBindex, 3) = dato4
				FBotTab(FBindex, 4) = dato5
				FBotTab(FBindex, 5) = dato6
				FBotTab(FBindex, 6) = dato7

			End If

		End If

	End Function

	Public Function TitBot(FBindex , dato1 , dato, dato3 , dato4 , dato5 , dato6 , dato7 )


		If IsNumeric(FBindex) Then

			If FBindex < 21 Then

				FTitTab(FBindex, 0) = dato1
				FTitTab(FBindex, 1) = dato2
				FTitTab(FBindex, 2) = dato3
				FTitTab(FBindex, 3) = dato4
				FTitTab(FBindex, 4) = dato5
				FTitTab(FBindex, 5) = dato6
				FTitTab(FBindex, 6) = dato7

			End If

		End If

	End Function

	Public Function GetTit(FBindex )

		If IsNumeric(FBindex) Then

			If FBindex < 21 Then


				FTitDat1 = FTitTab(FBindex, 0)
				FTitDat= FTitTab(FBindex, 1)
				FTitDat3 = FTitTab(FBindex, 2)
				FTitDat4 = FTitTab(FBindex, 3)
				FTitDat5 = FTitTab(FBindex, 4)
				FTitDat6 = FTitTab(FBindex, 5)
				FTitDat7 = FTitTab(FBindex, 6)

			End If

		End If

	End Function


	Public Function GetBot(FBindex )

		If IsNumeric(FBindex) Then

			If FBindex < 21 Then


				FBotDat1 = FBotTab(FBindex, 0)
				FBotDat= FBotTab(FBindex, 1)
				FBotDat3 = FBotTab(FBindex, 2)
				FBotDat4 = FBotTab(FBindex, 3)
				FBotDat5 = FBotTab(FBindex, 4)
				FBotDat6 = FBotTab(FBindex, 5)
				FBotDat7 = FBotTab(FBindex, 6)

			End If

		End If

	End Function


	Public Function GetCript(sentido , dato )

		Dim i, masc, sDato, y
		Dim atabla(255)
		Dim atablaz(255)

		For i = 0 To 255
			atabla(i) = Chr(i)
			atablaz(i) = 999 - i
		Next
		sDato = ""
		If sentido = "E" Then

			For i = 1 To Len(dato)
				For y = 0 To 255
					If Mid(dato, i, 1) = atabla(y) Then
						masc = CInt(i * 7) + CInt(i * 31 / 17) + 100
						sDato = sDato & (CInt(atablaz(y)) - masc)
						Exit For
					End If
				Next
			Next
			GetCript = sDato
		Else
			y = 1
			For i = 1 To Len(dato) Step 3
				masc = 999 - (CInt(Mid(dato, i, 3)) + (CInt(y * 7) + CInt(y * 31 / 17) + 100))
				sDato = sDato & atabla(masc)
				y = y + 1
			Next
			GetCript = sDato

		End If

	End Function


	Public Sub Browser()

		''On Error GoTo BrowserErr

		Dim bc

		bc = Request.servervariables("HTTP_USER_AGENT")
		If InStr(bc, "MSIE") > 0 Then
			Explorador = "MSIE"
			wfont = 7
			wspacing = 1
			wsize = 1
			wresto = 0
			wbody = "bgcolor=#cccccc"
		Else
			Explorador = "NETSCAPE"
			wfont = 8
			wspacing = 0
			wsize = 2
			wresto = 3
			wbody = "bgcolor=#ccbbaa"
		End If

		Exit Sub

		BrowserErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		bc = oEH.RegError(oErr)
		Response.WRITE bc

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Sub

	Public Function GenTag(sTag, sName, sClass, sAlign, sWidth, sHeight, sMarginwidth, sMarginheight, sCellspacing, sCellpadding, sBackground, sImage, sAccion, sTarget, sMetodo, sHref, sTitulo, sSrc, sBorde, sFunc, sContenido, sCierra)

		'On Error GoTo GenTagErr

		Dim xHTML

		sTag = UCase(sTag)
		'    sName = UCase(sName)
		sAlign = UCase(sAlign)
		sBackground = UCase(sBackground)
		sImage = UCase(sImage)
		sCierra = UCase(sCierra)
		sAccion = UCase(sAccion)

		xHTML = "<" & sTag & ""
		If sName <> "" And sName <> "NO" Then xHTML = xHTML & " name=" & sName & " id=" & sName
		If sClass <> "" And UCase(sClass) <> "NO" Then xHTML = xHTML & " class=" & sClass
		If sAlign <> "" And sAlign <> "NO" Then xHTML = xHTML & " align=" & sAlign
		If sWidth <> "" And UCase(sWidth) <> "NO" Then xHTML = xHTML & " width=" & sWidth & " "
		If sHeight <> "" And UCase(sHeight) <> "NO" Then xHTML = xHTML & " height=" & sHeight & " "
		If sMarginwidth <> "" And UCase(sMarginwidth) <> "NO" Then xHTML = xHTML & " marginwidth=" & sMarginwidth & " "
		If sMarginheight <> "" And UCase(sMarginheight) <> "NO" Then xHTML = xHTML & " marginheight=" & sMarginheight & " "
		If sCellspacing <> "" And UCase(sCellspacing) <> "NO" Then xHTML = xHTML & " cellspacing=" & wspacing + sCellspacing & " "
		If sCellpadding <> "" And UCase(sCellpadding) <> "NO" Then xHTML = xHTML & " cellpadding=" & 1 + wspacing + sCellpadding & " "
		If sBackground <> "" And sBackground <> "NO" Then xHTML = xHTML & " bgcolor=" & sBackground & " "
		If sImage <> "" And sImage <> "NO" Then xHTML = xHTML & " background=" & sImage & " "
		If sAccion <> "" And sAccion <> "NO" Then xHTML = xHTML & " action=" & Chr(34) & sAccion & Chr(34) & " "
		If sTarget <> "" And UCase(sTarget) <> "NO" Then xHTML = xHTML & " target=" & sTarget & " "
		If sFunc <> "" And UCase(sFunc) <> "NO" Then xHTML = xHTML & sFunc & " "
		If sMetodo <> "" And UCase(sMetodo) <> "NO" Then xHTML = xHTML & " method=" & sMetodo & " "

		xHTML = xHTML & "> "

		If sContenido <> "" Then xHTML = xHTML & sContenido
		If sCierra <> "NO" Then xHTML = xHTML & "</" & sTag & "> "

		If UV = "SI" Then GenTag = xHTML

		Exit Function

		GenTagErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		GenTag = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function GenTabla(sName, sClass, sAlign, sWidth, sHeight, sMarginwidth, sMarginheight, sCellspacing, sCellpadding, sBackground, sImage, sContenido, sCierra)

		'On Error GoTo GenTablaErr

		Dim xHTML

		'    xHTML = GenTag("Table", sName, sClass, sAlign, sWidth, sHeight, sMarginwidth, sMarginheight, sCellspacing, sCellpadding, sBackground, sImage, "", "", "", "", "", "", "", "", sContenido, sCierra)

		sAlign = UCase(sAlign)
		sBackground = UCase(sBackground)
		sImage = UCase(sImage)
		sCierra = UCase(sCierra)

		xHTML = "<TABLE "
		If sName <> "" And sName <> "NO" Then xHTML = xHTML & " name=" & sName & " id=" & sName
		If sClass <> "" And UCase(sClass) <> "NO" Then xHTML = xHTML & " class=" & sClass
		If sAlign <> "" And sAlign <> "NO" Then xHTML = xHTML & " align=" & sAlign
		If sWidth <> "" And UCase(sWidth) <> "NO" Then xHTML = xHTML & " width=" & sWidth & " "
		If sHeight <> "" And UCase(sHeight) <> "NO" Then xHTML = xHTML & " height=" & sHeight & " "
		If sMarginwidth <> "" And UCase(sMarginwidth) <> "NO" Then xHTML = xHTML & " marginwidth=" & sMarginwidth & " "
		If sMarginheight <> "" And UCase(sMarginheight) <> "NO" Then xHTML = xHTML & " marginheight=" & sMarginheight & " "
		If sCellspacing <> "" And UCase(sCellspacing) <> "NO" Then xHTML = xHTML & " cellspacing=" & wspacing + sCellspacing & " "
		If sCellpadding <> "" And UCase(sCellpadding) <> "NO" Then xHTML = xHTML & " cellpadding=" & 1 + wspacing + sCellpadding & " "
		If sBackground <> "" And sBackground <> "NO" Then xHTML = xHTML & " bgcolor=" & sBackground & " "
		If sImage <> "" And sImage <> "NO" Then xHTML = xHTML & " background=" & sImage & " "
		xHTML = xHTML & "> "

		If sContenido <> "" Then xHTML = xHTML & sContenido
		If sCierra <> "NO" Then xHTML = xHTML & "</TABLE>"

		If UV = "SI" Then GenTabla = xHTML

		Exit Function

		GenTablaErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		GenTabla = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function GenFTabla()

		If UV = "SI" Then GenFTabla = "</TABLE> "

	End Function

	Public Function GenRow(GRnom, GRclase, GRalign, GRw, GRh, GRbg, GRbi, GRcont, GRc)

		'On Error GoTo GenRowErr

		Dim xHTML

		'    xHTML = GenTag("Tr", GRnom, GRclase, GRalign, GRw, GRh, "", "", "", "", GRbg, GRbi, "", "", "", "", "", "", "", "", GRcont, GRc)

		xHTML = "<TR "
		If GRnom <> "" And UCase(GRnom) <> "NO" Then xHTML = xHTML & " name=" & GRnom & " ID=" & GRnom
		If GRclase <> "" And UCase(GRclase) <> "NO" Then xHTML = xHTML & " class=" & GRclase
		If GRalign <> "" And UCase(GRalign) <> "NO" Then xHTML = xHTML & " align=" & GRalign
		If GRw <> "" And UCase(GRw) <> "NO" Then xHTML = xHTML & " width=" & GRw & " "
		If GRh <> "" And UCase(GRh) <> "NO" Then xHTML = xHTML & " height=" & GRh & " "
		If GRbg <> "" And UCase(GRbg) <> "NO" Then xHTML = xHTML & " bgcolor=" & GRbg & " "
		If GRbi <> "" And UCase(GRbi) <> "NO" Then xHTML = xHTML & " background=" & GRbi & " "
		xHTML = xHTML & "> "
		If GRcont <> "" Then xHTML = xHTML & GRcont
		If UCase(GRc) <> "NO" Then xHTML = xHTML & "</TR> "

		If UV = "SI" Then GenRow = xHTML

		Exit Function

		GenRowErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		GenRow = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function GenFRow()

		If UV = "SI" Then GenFRow = "</TR> "

	End Function

	Public Function GenCelda(GCnom, GCclase, GCalign, GCw, GCh, GCbg, GCbi, GCcont, GCc)

		'On Error GoTo GenCeldaErr

		Dim xHTML

		'    xHTML = GenTag("Td", GCnom, GCclase, GCalign, GCw, GCh, "", "", "", "", GCbg, GCbi, "", "", "", "", "", "", "", "", GCcont, GCc)

		xHTML = "<TD "
		If GCnom <> "" And UCase(GCnom) <> "NO" Then xHTML = xHTML & " name=" & GCnom & " ID=" & GCnom
		If GCclase <> "" And UCase(GCclase) <> "NO" Then xHTML = xHTML & " class=" & GCclase
		If GCalign <> "" And UCase(GCalign) <> "NO" Then xHTML = xHTML & " align=" & GCalign
		If GCw <> "" And UCase(GCw) <> "NO" Then xHTML = xHTML & " width=" & GCw & " "
		If GCh <> "" And UCase(GCh) <> "NO" Then xHTML = xHTML & " height=" & GCh & " "
		If GCbg <> "" And UCase(GCbg) <> "NO" Then xHTML = xHTML & " bgcolor=" & GCbg & " "
		If GCbi <> "" And UCase(GCbi) <> "NO" Then xHTML = xHTML & " background=" & GCbi & " "
		xHTML = xHTML & ">"

		If GCcont <> "" Then xHTML = xHTML & GCcont
		If UCase(GCc) <> "NO" Then xHTML = xHTML & "</TD>"

		If UV = "SI" Then GenCelda = xHTML

		Exit Function

		GenCeldaErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		GenCelda = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function GenFCelda()

		If UV = "SI" Then GenFCelda = "</TD>"

	End Function

	Public Function GenForm(GFnom, GFclase, GFaction, GFtarget, GFmethod, GFfunc, GFcont, GFc)

		'On Error GoTo GenFormErr

		Dim xHTML

		'    xHTML = GenTag("Form", GFnom, GFclase, "", "", "", "", "", "", "", "", "", GFaction, GFtarget, GFmethod, "", "", "", "", GFfunc, GFcont, GFc)

		xHTML = "<FORM "
		If GFnom <> "" And UCase(GFnom) <> "NO" Then xHTML = xHTML & " name=" & GFnom & " ID=" & GFnom
		If GFclase <> "" And UCase(GFclase) <> "NO" Then xHTML = xHTML & " class=" & GFclase
		If GFaction <> "" And UCase(GFaction) <> "NO" Then xHTML = xHTML & " action=" & Chr(34) & GFaction & Chr(34) & " "
		If GFtarget <> "" And UCase(GFtarget) <> "NO" Then xHTML = xHTML & " target=" & GFtarget & " "
		If GFfunc <> "" And UCase(GFfunc) <> "NO" Then xHTML = xHTML & GFfunc & " "
		If GFmethod <> "" And UCase(GFmethod) <> "NO" Then xHTML = xHTML & " method=" & GFmethod & " "
		xHTML = xHTML & ">"

		If GFcont <> "" Then xHTML = xHTML & GFcont
		If UCase(GFc) <> "NO" Then xHTML = xHTML & "</FORM>"

		If UV = "SI" Then GenForm = xHTML

		Exit Function

		GenFormErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		GenForm = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function GenFForm()

		If UV = "SI" Then GenFForm = "</FORM>"

	End Function

	Public Function GenLink(GLnom, GLclase, GLhref, GLtarget, GFtitle, GLfunc, GLcont, GLc)

		'On Error GoTo GenLinkErr

		Dim xHTML

		'    xHTML = GenTag("A", GLnom, GLclase, "", "", "", "", "", "", "", "", "", "", GLtarget, "", GLhref, GFtitle, "", "", GLfunc, GLcont, GLc)

		xHTML = "<A "
		If GLnom <> "" And UCase(GLnom) <> "NO" Then xHTML = xHTML & " name=" & GLnom & " ID=" & GLnom
		If GLclase <> "" And UCase(GLclase) <> "NO" Then xHTML = xHTML & " class=" & GLclase
		If GLhref <> "" And UCase(GLhref) <> "NO" Then xHTML = xHTML & " href=" & Chr(34) & GLhref & Chr(34) & " "
		If GLtarget <> "" And UCase(GLtarget) <> "NO" Then xHTML = xHTML & " target=" & GLtarget & " "
		If GLfunc <> "" And UCase(GLfunc) <> "NO" Then xHTML = xHTML & GLfunc & " "
		If GFtitle <> "" And UCase(GFtitle) <> "NO" Then xHTML = xHTML & " title=" & GFtitle & " "
		xHTML = xHTML & ">"

		If GLcont <> "" Then xHTML = xHTML & GLcont
		If UCase(GLc) <> "NO" Then xHTML = xHTML & "</A>"

		If UV = "SI" Then GenLink = xHTML

		Exit Function

		GenLinkErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		GenLink = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function GenFLink()

		If UV = "SI" Then GenFLink = "</A>"

	End Function

	Public Function GenImage(GInom, GIclase, GIw, GIh, GIsrc, GIborde, GIfunc)

		'On Error GoTo GenImageErr

		Dim xHTML

		'    xHTML = GenTag("Img", GInom, GIclase, "", GIw, GIh, "", "", "", "", "", "", "", "", "", "", "", GIsrc, GIborde, GIfunc, "", "no")

		xHTML = "<IMG "
		If GInom <> "" And UCase(GInom) <> "NO" Then xHTML = xHTML & " name=" & GInom & " ID=" & GInom
		If GIclase <> "" And UCase(GIclase) <> "NO" Then xHTML = xHTML & " class=" & GIclase
		If GIw <> "" And UCase(GIw) <> "NO" Then xHTML = xHTML & " width=" & GIw & " "
		If GIh <> "" And UCase(GIh) <> "NO" Then xHTML = xHTML & " height=" & GIh & " "
		If GIsrc <> "" And UCase(GIsrc) <> "NO" Then xHTML = xHTML & " src=" & Chr(34) & GIsrc & Chr(34) & " "
		If GIborde <> "" And UCase(GIborde) <> "NO" Then xHTML = xHTML & " border=" & GIborde & " "
		If GIfunc <> "" And UCase(GIfunc) <> "NO" Then xHTML = xHTML & GIfunc & " "
		xHTML = xHTML & ">"

		If UV = "SI" Then GenImage = xHTML

		Exit Function

		GenImageErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		GenImage = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function GenerarInput(Gnombre, Gvalor, Gtipo, Gform, Gsize, Gclase, Gnumb)

		'On Error GoTo GenerarInputErr

		Dim xHTML, xGtipo, xTag

		xHTML = "<input "
		If UCase(Gtipo) = UCase("readonly") Or UCase(Gtipo) = UCase("agregar") Or UCase(Gtipo) = UCase("textsb") Then
			xGtipo = "TEXT"
		Else
			xGtipo = UCase(Gtipo)
		End If
		If UCase(Gtipo) = UCase("readonly") Then
			If Explorador = "MSIE" Then
				xTag = " readonly "
			Else
				xTag = " "
			End If
		Else
			xTag = " "
		End If

		xHTML = xHTML & " type=" & xGtipo & xTag

		If UCase(Gtipo) = UCase("text") Then
			xHTML = xHTML & ValidarActualizacion(Gnumb, Gform, Gnombre)
		Else
			If UCase(Gtipo) = UCase("textsb") Then
				xHTML = xHTML & " onblur=" & Chr(34) & "document." & Gform & ".submit(); return true;" & Chr(34) & " "
			ElseIf UCase(Gtipo) = UCase("button") Then
				If Gnumb <> "" Then
					xHTML = xHTML & " onclick=" & Chr(34) & Gnumb & Chr(34) & " "
				End If
			End If
		End If

		If UCase(Gtipo) <> UCase("hidden") Then
			xHTML = xHTML & " size=" & (Gsize / wsize) - wresto & "% "
		End If
		xHTML = xHTML & " value='" & Gvalor & "'"
		If UCase(Gtipo) <> UCase("readonly") And UCase(Gtipo) <> UCase("button") And UCase(Gtipo) <> UCase("submit") Then
			xHTML = xHTML & " name=" & Gnombre & " id='" & Gnombre & "'"
		End If
		If UCase(Gtipo) <> UCase("hidden") Then
			xHTML = xHTML & " class=" & Gclase
		End If
		xHTML = xHTML & " >"

		If UV = "SI" Then GenerarInput = xHTML

		Exit Function

		GenerarInputErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		GenerarInput = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number
	End Function
	Public Function GenComboFX48(fn )
		GenComboFX48 = fn
	End Function




	Function GenComboFX(sForm, NombreCampox, ValorCampo, Indice, Descripcion, ID)
		GenComboFX = GenComboM("0", sForm, NombreCampox, ValorCampo, "", "", "", "", "", "", "", "", "", Indice, Descripcion, ID)
	End Function

	Function GenerarCombo(ASPx, NombreCampox, ValorCampo, SQL, Conexion, Indice, Descripcion, ID)
		GenerarCombo = GenComboM("1", ASPx, NombreCampox, ValorCampo, "", "", "", "", "", "", "", SQL, Conexion, Indice, Descripcion, ID)
	End Function

	Function GenerarComboSP(ASP, NombreCampo, ValorCampo, SQL, Conexion, Indice, Descripcion, ID)
		GenerarComboSP = GenComboM("2", ASP, NombreCampo, ValorCampo, "", "", "", "", "", "", "", SQL, Conexion, Indice, Descripcion, ID)
	End Function

	Function GenerarLista(SQL, Conexion, Campo)
		GenerarLista = GenComboM("3", "", Campo, "", "", "", "", "", "", "", "", SQL, Conexion, "", "", "")
	End Function

	Function GenerarCombox(SQL, Conexion, Campo, Nombre)
		GenerarCombox = GenComboM("4", "", Nombre, "", Campo, "", "", "", "", "", "", SQL, Conexion, "", "", "")
	End Function

	Function GenerarCombox2(SQL, Conexion, Campo1, Campo2, Nombre)
		GenerarCombox2 = GenComboM("5", "", Nombre, "", Campo1, "", "", "", "", "", "", SQL, Conexion, "", Campo2, "")
	End Function

	Function GenerarCombox3(SQL, Conexion, Campo1, Campo2, Nombre)
		GenerarCombox3 = GenComboM("6", "", Nombre, "", Campo1, "", "", "", "", "", "", SQL, Conexion, "", Campo2, "")
	End Function

	Function GenerarCombo2(NombreCampo, ValorCampo, SQL, Conexion, Indice, Descripcion, ID)
		GenerarCombo2 = GenComboM("7", "", NombreCampo, ValorCampo, "", "", "", "", "", "", "", SQL, Conexion, Indice, Descripcion, ID)
	End Function

	Function GenerarCombo3(ASP, xArray, NombreCampo, ValorCampo)
		GenerarCombo3 = GenComboM("8", ASP, NombreCampo, ValorCampo, "", "", "", "", xArray, "", "", "", "", "", "", "")
	End Function

	Function GenerarCombo333(xArray, NombreCampo, ValorCampo, direc, imagen)
		GenerarCombo333 = GenComboM("9", "", NombreCampo, ValorCampo, "", "", direc, imagen, xArray, "", "", "", "", "", "", "")
	End Function

	Function GenerarCombo33(xArray, NombreCampo, ValorCampo)
		GenerarCombo33 = GenComboM("a", "", NombreCampo, ValorCampo, "", "", "", "", xArray, "", "", "", "", "", "", "")
	End Function

	Function comboSiNO(Form, Valor, Nombre)
		''100     comboSiNO = GenComboM("b", Form, Nombre, Valor, "", "", "", "", "", "", "", "", "", "", "", "")
		comboSiNO = ComboSiNoNEW(Form, Valor, Nombre, "", "", "b")
	End Function

	Function comboSiNoEx(Form, Valor, Nombre, s, n)
		''100     comboSiNoEx = GenComboM("c", Form, Nombre, Valor, "", "", "", "", "", s, n, "", "", "", "", "")
		comboSiNoEx = ComboSiNoNEW(Form, Valor, Nombre, s, n, "c")
	End Function

	Function comboSiNOmin(Nombre)
		''100     comboSiNOmin = GenComboM("d", "", Nombre, "", "", "", "", "", "", "", "", "", "", "", "", "")
		comboSiNOmin = ComboSiNoNEW("", "", Nombre, "", "", "d")
	End Function







	Public Function GenComboM(BTipo, Bform, BNCampo, BVCampo, BNCampo2, BVCampo2, Bdir, Bima, BArray, Bs, Bn, BSQL, BCon, BInd, BDesc, BOpc)

		'On Error GoTo GenComboMErr

		Dim sID, bfnom, bfnom2, bfval, bfval2, array1, array2, arrayx
		Dim arrayx1, arrayx2
		Dim scx, sx, nx, n

		Dim XFFunC, XFNEWC, XFCOMB, XFCFLD, XFNOMBCM, XFY

		Dim XFORM, XFDATO, XFNOMB, XFTCOM, XFSQL, XFCONN, XFKEY, XFDSC, XFTIP

		bfnom = BNCampo
		bfnom= BNCampo2
		bfval = BVCampo
		bfval= BVCampo2


		If IsNull(BVCampo) Then bfval = ""
		If IsNull(BVCampo2) Then bfval= ""




		XFTIP = ""




		'' if GLBCOMBO    = 0 then
		''    XFFunC      = gencomboXS()
		''    FBBarr      = FBBarr & XFFunC
		'' end if





		Select Case BTipo

			Case "0"

			XFORM = Bform
			XFDATO = bfval
			XFNOMB = bfnom

			XFSQL = "N"
			XFCONN = ""
			XFKEY = BInd
			XFDSC = BDesc

			Select Case BOpc

				Case "no"
				XFTCOM = "0"
				Case "no0"
				XFTCOM = "10"
				Case "sb"
				XFTCOM = "1"
				Case "sb0"
				XFTCOM = "11"
				Case Else
				XFTCOM = "2"
			End Select

			Case "1"

			XFORM = Bform
			XFDATO = bfval
			XFNOMB = bfnom

			XFSQL = BSQL
			Set XFCONN = BCon
			XFKEY = BInd
			XFDSC = BDesc


			Select Case BOpc

				Case "no"
				XFTCOM = "0"
				Case "no0"
				XFTCOM = "10"
				Case "sb"
				XFTCOM = "1"
				Case "sb0"
				XFTCOM = "11"
				Case Else
				XFTCOM = "2"
			End Select



			Case "2"

			XFORM = Bform
			XFDATO = bfval
			XFNOMB = bfnom

			XFSQL = BSQL
			Set XFCONN = BCon
			XFKEY = BInd
			XFDSC = BDesc

			XFTCOM = "2"



			Case "3"

			XFORM = Bform
			XFDATO = ""
			XFNOMB = "lista"

			XFSQL = BSQL
			Set XFCONN = BCon
			XFKEY = bfnom
			XFDSC = bfnom

			XFTCOM = "20"



			Case "4"

			XFORM = Bform
			XFDATO = ""
			XFNOMB = bfnom

			XFSQL = BSQL
			Set XFCONN = BCon
			XFKEY = bfnom2
			XFDSC = bfnom2

			XFTCOM = "20"


			Case "5"

			XFORM = Bform
			XFDATO = ""
			XFNOMB = bfnom

			XFSQL = BSQL
			Set XFCONN = BCon
			XFKEY = bfnom2
			XFDSC = BDesc

			XFTCOM = "20"



			Case "6"

			XFORM = Bform
			XFDATO = ""
			XFNOMB = bfnom

			XFSQL = BSQL
			Set XFCONN = BCon
			XFKEY = bfnom2
			XFDSC = BDesc

			XFTCOM = "0"


			Case "7"

			XFORM = Bform
			XFDATO = bfval
			XFNOMB = bfnom

			XFSQL = BSQL
			Set XFCONN = BCon
			XFKEY = BInd
			XFDSC = BDesc

			XFTCOM = "0"


			Case "8"

			XFORM = Bform
			XFDATO = bfval
			XFNOMB = bfnom

			XFSQL = "E"
			XFCONN = Join(BArray, Chr(9))
			XFKEY = ""
			XFDSC = ""

			XFTCOM = "2"


			Case "9"

			XFORM = Bform
			XFDATO = bfval
			XFNOMB = bfnom

			XFSQL = "E"
			XFCONN = Join(BArray, Chr(9))
			XFKEY = ""
			XFDSC = ""

			XFTCOM = "0"



			Case "a"

			XFORM = Bform
			XFDATO = bfval
			XFNOMB = bfnom

			XFSQL = "E"
			XFCONN = Join(BArray, Chr(9))
			XFKEY = ""
			XFDSC = ""

			XFTCOM = "0"


		End Select

		''''''''''  buscar si ya esta el combo

		XFNEWC = "S"
		For XFY = 1 To GLBCOMBO
			If (XFNOMB & XFTCOM) = ATXCOM(XFY, 0) Then
				GLBCACTU = XFY
				XFNEWC = "N"
				Exit For
			End If
		Next

		'''''''''' si es nuevo

		If XFNEWC = "S" Then

			GLBCOMBO = GLBCOMBO + 1
			ATXCOM(GLBCOMBO, 0) = XFNOMB & XFTCOM
			GLBCACTU = GLBCOMBO
			XFCOMB = GenComboINT(XFNOMB, XFTCOM, XFSQL, XFCONN, XFKEY, XFDSC, XFTIP)
			FBBarr = FBBarr & XFCOMB
		End If

		'''''''''''''''''''''''''''''''''''''

		XFCFLD = GenComboFLD(XFORM, XFDATO, XFNOMB, XFTCOM, XFSQL, XFCONN, XFKEY, XFDSC, XFTIP)





		If UV = "SI" Then GenComboM = XFCFLD

		Exit Function

		GenComboMErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision & " - GenComboM"
		End With

		'On Error Resume Next

		GenComboM = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function
	''''''''''''''''''''''''''
	Function ComboSiNoNEW(XFORM, XFDATO, XFNOMB, XFBS, XFBN, XFTIP)

		'On Error GoTo ComboSiNoNEWMErr


		Dim sID1, sID2, scx, scx1, scx2, XFDATOSHW, XFS1, XFS2

		XFS1 = "<B><FONT class=bt1 title='SI' >a</FONT></B>"
		XFS= "<B><FONT class=bttitle='NO' >r</FONT></B>"

		sID1 = "<A CLASS=NT HREF='javascript:void(0);' onclick=""document.getElementById('" & XFNOMB & "').value='XXX2';XXX3 return false; "" >"
		sID= "</A><INPUT TYPE='HIDDEN' VALUE='XXX1' NAME='" & XFNOMB & "'  id='" & XFNOMB & "'>"


		Select Case XFTIP

			Case "b"

			scx = Replace(sID1, "XXX3", "document." & XFORM & ".submit();")

			If XFDATO Then
				scx = Replace(scx, "XXX2", "0") & XFS1 & Replace(sID2, "XXX1", "1")
			Else
				scx = Replace(scx, "XXX2", "1") & XFS& Replace(sID2, "XXX1", "0")
			End If

			Case "c"

			scx = Replace(sID1, "XXX3", "document." & XFORM & ".submit();")

			If XFDATO Then
				scx = Replace(scx, "XXX2", XFBN) & XFS1 & Replace(sID2, "XXX1", XFBS)
			Else
				scx = Replace(scx, "XXX2", XFBS) & XFS& Replace(sID2, "XXX1", XFBN)
			End If

			Case Else

			scx = Replace(sID1, "XXX3", "")

			If XFDATO Then
				scx = Replace(scx, "XXX2", "0") & XFS1 & Replace(sID2, "XXX1", "1")
			Else
				scx = Replace(scx, "XXX2", "1") & XFS& Replace(sID2, "XXX1", "0")
			End If


		End Select






		ComboSiNoNEW = scx

		Exit Function

		ComboSiNoNEWMErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision & " - ComboSiNoNEW"
		End With

		'On Error Resume Next

		ComboSiNoNEW = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number


	End Function

	''''''''''''''''''''''''''
	Function GenComboINT(XFNOMB, XFTCOM, XFSQL, XFCONN, XFKEY, XFDSC, XFTIP)

		'On Error GoTo GenComboINTErr


		Dim sID1, sID2, scx, scx1, scx2, XFDATOSHW, XFS1, XFS2, sCOPT, nx


		XFS1 = "<DIV class='DINDM' id='DF@" & XFNOMB & XFTCOM & "' "
		XFS1 = XFS1 & " STYLE='display:none;position:absolute;left:300;top:300;width:300;' >"

		XFS1 = XFS1 & "<table id='DF1CMB' style='z-index:81;' width=100% class='TWIND'  cellspacing=0 cellpadding=>"
		XFS1 = XFS1 & "<tr>"
		XFS1 = XFS1 & "<td class=mbt   >&nbsp;"

		XFS1 = XFS1 & "SELECCIONE DE LA LISTA..."

		XFS1 = XFS1 & "</td>"
		XFS1 = XFS1 & "<td CLASS=mbt width=onclick=""window.top.xclcomS();"" TITLE='OCULTAR' >"
		' If m_Browser = "IE" Then
		XFS1 = XFS1 & "<font  onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & Session("CMFMNBOT") & "';"" face=Corflaticon size=>r</font>"
		' Else
		' XFS1 = XFS1 & "<font  onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & Session("CMFMNBOT") & "';"" face=Wingdings size=>&#57369;</font>"
		' End If
		XFS1 = XFS1 & "</td>"
		XFS1 = XFS1 & "</tr>"
		XFS1 = XFS1 & "</table>"

		XFS1 = XFS1 & "<TABLE id='DF2CMB' width=100% style='z-index:81;' class='TWINE2' >"
		XFS1 = XFS1 & "<TR HEIGHT=10 ><TD COLSPAN=3>&nbsp;</TD></TR>"
		XFS1 = XFS1 & "<TR><TD >&nbsp;</TD><TD ALIGN=CENTER >XXX1</TD><TD >&nbsp;</TD></TR>"
		XFS1 = XFS1 & "<TR HEIGHT=10 ><TD COLSPAN=3>&nbsp;</TD></TR>"
		XFS1 = XFS1 & "</TABLE>"

		XFS1 = XFS1 & "</DIV>"


		sCOPT = GenComboOPT(XFNOMB, XFTCOM, XFSQL, XFCONN, XFKEY, XFDSC, XFTIP)


		scx = "<FORM METHOD=POST><FONT CLASS=AT  >"
		scx = scx & "<SELECT onchange='window.top.xclcom();' NAME='" & XFNOMB & XFTCOM & "' ID='" & XFNOMB & XFTCOM & "' >"
		scx = scx & sCOPT & "</SELECT></FONT></FORM>"



		GenComboINT = Replace(XFS1, "XXX1", scx)

		Exit Function

		GenComboINTErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision & " - GenComboINT"
		End With

		'On Error Resume Next

		GenComboINT = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number



	End Function


	Function GenComboOPT(XFNOMB, XFTCOM, XFSQL, XFCONN, XFKEY, XFDSC, XFTIP)

		'On Error GoTo GenComboOPTErr


		Dim XFarray1, XFarray2
		Dim sID1, sID2, scx, scx1, scx2, XFDATOSHW, XFS1, XFS2, nx, XFTCOMGB


		'''' XFTCOM '''''''''''''''''''''''''''''''''''''' XFTCOM ''''''''''''
		''''    0 = sin submit con espacio
		''''   10 = sin submit con cero
		''''   20 = sin submit                                 1
		'''''''''''''''''''''''''''''''''''''''''''''''''
		''''    1 = con submit con espacio
		''''   11 = con submit con cero
		''''   21 = con submit                                 2
		'''''''''''''''''''''''''''''''''''''''''''''''''
		''''    = con submit y valid con espacio
		''''   1= con submit y valid con cero
		''''   2= con submit y valid                         3
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


		XFS1 = "<OPTION CLASS=OCMB VALUE='XXX0' >XXX1</OPTION>"
		scx = ""

		XFTCOMGB = "N"


		If XFTCOM = "0" Or XFTCOM = "1" Or XFTCOM = "2" Then

			scx = scx & Replace(XFS1, "XXX0", "nsnc")
			XFTCOMGB = "S1"

		ElseIf XFTCOM = "10" Or XFTCOM = "11" Or XFTCOM = "12" Then

			scx = scx & Replace(XFS1, "XXX0", "0")
			XFTCOMGB = "S2"

		Else

			scx = " "


		End If

		'''CMBFiller

		If CMBFillOPC = "2" Then
			scx = Replace(scx, "XXX1", CMBFiller2)
		Else
			scx = Replace(scx, "XXX1", CMBFiller)
		End If


		Select Case UCase(XFSQL)

			Case "N"
			'''''''''''''''''''''''''' arrays de claves y datos (valores fijos)

			If XFTCOMGB <> "N" Then

				If XFTCOMGB = "S1" Then
					XFKEY = "nsnc" & Chr(9) & XFKEY
				Else
					XFKEY = "0" & Chr(9) & XFKEY
				End If



				If CMBFillOPC = "2" Then
					XFDSC = CMBFiller& Chr(9) & XFDSC
				Else
					XFDSC = CMBFiller & Chr(9) & XFDSC
				End If

				scx = ""

			End If


			XFarray1 = Split(XFKEY, Chr(9))
			XFarray2= Split(XFDSC, Chr(9))




			ATXCOM(GLBCACTU, 2) = "N"
			ATXCOM(GLBCACTU, 3) = XFKEY
			ATXCOM(GLBCACTU, 4) = XFDSC

			For nx = 0 To UBound(XFarray1)
				If XFarray1(nx) <> "" Then
					scx = scx & Replace(Replace(XFS1, "XXX1", XFarray2(nx)), "XXX0", ConvertirCar(XFarray1(nx), " ", "_"))
				End If
			Next

			Case "E"
			'''''''''''''''''''''''''' arrays a partir de un array externo

			XFarray1 = Split(XFCONN, Chr(9))

			ATXCOM(GLBCACTU, 2) = "E"
			ATXCOM(GLBCACTU, 3) = XFCONN
			ATXCOM(GLBCACTU, 4) = XFCONN

			For nx = 0 To UBound(XFarray1)
				If XFarray1(nx) <> "" Then
					scx = scx & Replace(Replace(XFS1, "XXX1", XFarray1(nx)), "XXX0", UCase(ConvertirCar(XFarray1(nx), " ", "_")))
				End If
			Next

			Case "A"
			'''''''''''''''''''''''''' arrays a partir de un directorio (archivos/imagenes)

			XFarray1 = GetFiles(XFCONN)

			ATXCOM(GLBCACTU, 2) = "A"
			ATXCOM(GLBCACTU, 3) = Join(XFarray1, Chr(9))
			ATXCOM(GLBCACTU, 4) = ATXCOM(GLBCACTU, 3)

			For nx = 0 To UBound(XFarray1)
				If XFarray1(nx) <> "" Then
					scx = scx & Replace(Replace(XFS1, "XXX1", XFarray1(nx)), "XXX0", UCase(ConvertirCar(XFarray1(nx), " ", "_")))
				End If
			Next

			Case Else
			'''''''''''''''''''''''''' arrays a partir de una consulta (normal)

			XFarray1 = GenerarArray(XFSQL, XFCONN, XFKEY)
			XFarray2= GenerarArray(XFSQL, XFCONN, XFDSC)



			If XFTCOMGB <> "N" Then



				If XFTCOMGB = "S1" Then
					XFarray1(0) = "nsnc"
				Else
					XFarray1(0) = "0"
				End If


				'   response.write CMBFiller2
				If CMBFillOPC = "2" Then
					XFarray2(0) = CMBFiller2
				Else
					XFarray2(0) = CMBFiller
				End If

				scx = ""


			End If



			ATXCOM(GLBCACTU, 2) = "G"
			ATXCOM(GLBCACTU, 3) = Join(XFarray1, Chr(9))
			ATXCOM(GLBCACTU, 4) = Join(XFarray2, Chr(9))



			For nx = 1 To UBound(XFarray1)
				If XFarray1(nx) <> "" Then
					scx = scx & Replace(Replace(XFS1, "XXX1", XFarray2(nx)), "XXX0", ConvertirCar(XFarray1(nx), " ", "_"))
				End If
			Next

		End Select



		GenComboOPT = scx

		Exit Function

		GenComboOPTErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision & " - GenComboOPT"
		End With

		'On Error Resume Next

		GenComboOPT = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number


	End Function


	Function GenComboFLD(XFORM, XFDATO, XFNOMB, XFTCOM, XFSQL, XFCONN, XFKEY, XFDSC, XFTIP)

		'On Error GoTo GenComboFLDErr

		Dim XFarray1, XFarray2
		Dim sID1, sID2, scx, scx1, scx2, XFDATOSHW, XFS1, XFS2, XFTCOMFNC, XFY

		'''' XFTCOM '''''''''''''''''''''''''''''''''''''' XFTCOM ''''''''''''
		''''    0 = sin submit con espacio
		''''   10 = sin submit con cero
		''''   20 = sin submit                                 1
		'''''''''''''''''''''''''''''''''''''''''''''''''
		''''    1 = con submit con espacio
		''''   11 = con submit con cero
		''''   21 = con submit                                 2
		'''''''''''''''''''''''''''''''''''''''''''''''''
		''''    = con submit y valid con espacio
		''''   1= con submit y valid con cero
		''''   2= con submit y valid                         3
		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

		Select Case XFTCOM

			Case "0"
			XFTCOMFNC = "1"

			Case "10"
			XFTCOMFNC = "1"

			Case "20"
			XFTCOMFNC = "1"

			Case "1"
			XFTCOMFNC = "2"

			Case "11"
			XFTCOMFNC = "2"

			Case "21"
			XFTCOMFNC = "2"

			Case "2"
			XFTCOMFNC = "3"

			Case "12"
			XFTCOMFNC = "3"

			Case "22"
			XFTCOMFNC = "3"


		End Select


		sID1 = "window.top.xopcom('" & XFNOMB & "','" & XFNOMB & XFTCOM & "','" & XFORM & "','" & XFTCOMFNC & "'); return false;"


		XFS1 = "<TABLE CLASS=TBCM cellspacing=0 cellpadding=0 >"
		XFS1 = XFS1 & "<TD CLASS=TDCM1 ID='TD@" & XFNOMB & XFTCOM & "' onclick=" & Chr(34) & sID1 & Chr(34) & " title='Seleccione de la lista' >"
		XFS1 = XFS1 & "XXX0"
		XFS1 = XFS1 & "</TD>"
		XFS1 = XFS1 & "<TD CLASS=TDCMWIDTH=onclick=" & Chr(34) & sID1 & Chr(34) & " title='Seleccione de la lista' >"
		XFS1 = XFS1 & "XXX1"
		XFS1 = XFS1 & "</TD>"
		XFS1 = XFS1 & "</TABLE>"
		' If m_Browser = "IE" Then
		XFS= "<B><FONT class=btc>6</FONT></B>"
		' Else
		' XFS= "<B><FONT class=btc>&#8679;</FONT></B>"
		' End If


		XFarray1 = Split(ATXCOM(GLBCACTU, 3), Chr(9))
		XFarray2= Split(ATXCOM(GLBCACTU, 4), Chr(9))

		XFDATOSHW = "&nbsp;&nbsp;"

		For XFY = 0 To UBound(XFarray1)
			If XFarray1(XFY) = XFDATO Or UCase(XFarray1(XFY) & "") = XFDATO & "" Then
				XFDATOSHW = XFarray2(XFY) & "&nbsp;&nbsp;"
				Exit For
			End If
		Next


		''''      sID1 = "<A CLASS=CMB HREF='javascript:void(0);' onclick=""window.top.xopcom('" & XFNOMB & "','" & XFNOMB & XFTCOM & "','" & XFORM & "','" & XFTCOMFNC & "'); return false; "" title='Seleccione de la lista' >XXX2</A>"

		sID1 = "XXX2"
		sID= "<INPUT TYPE='HIDDEN' VALUE='" & XFDATO & "' id='" & XFNOMB & "' NAME='" & XFNOMB & "' >"

		scx = XFS1

		scx1 = Replace(sID1, "XXX2", XFDATOSHW)
		scx= Replace(sID1, "XXX2", XFS2)

		scx = Replace(scx, "XXX0", scx1)
		scx = Replace(scx, "XXX1", scx2)
		scx = scx & sID2

		GenComboFLD = scx

		Exit Function

		GenComboFLDErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision & " - GenComboFLD"
		End With

		'On Error Resume Next

		GenComboFLD = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number


	End Function

	''''''''''''''''''''''''''

	''''''''''''''''''''''
	Public Function GenComboX(BTipo, Bform, BNCampo, BVCampo, BNCampo2, BVCampo2, Bdir, Bima, BArray, Bs, Bn, BSQL, BCon, BInd, BDesc, BOpc)

		'On Error GoTo GenComboXErr

		Dim sID, bfnom, bfnom2, bfval, bfval2, array1, array2, arrayx
		Dim arrayx1, arrayx2
		Dim scx, sx, nx, n

		bfnom = BNCampo
		bfnom= BNCampo2
		bfval = BVCampo
		bfval= BVCampo2


		If IsNull(BVCampo) Then bfval = ""
		If IsNull(BVCampo2) Then bfval= ""

		Select Case BTipo

			Case "1"



			array1 = GenerarArray(BSQL, BCon, BInd)
			Array= GenerarArray(BSQL, BCon, BDesc)

			If BOpc = "no" Or BOpc = "no0" Then
				sID = ""
			Else
				If BOpc = "sb" Or BOpc = "sb0" Then
					sID = " onchange=" & Chr(34) & Bform & ".submit(); return true;" & Chr(34) & " "
				Else
					sID = ValidarActualizacionCombo(Bform, bfnom)
				End If
			End If
			If BOpc = "no0" Or BOpc = "sb0" Then
				scx = "<option value=0>&nbsp;</option>"
			Else
				scx = "<option value=nsnc>&nbsp;</option>"
			End If

			For nx = 1 To UBound(array1)

				If array1(nx) <> "" Then

					If CStr(array1(nx)) = CStr(bfval) Then
						scx = scx & "<option value=" & ConvertirCar(array1(nx), " ", "_") & " selected>" & array2(nx) & "</option>"
					Else
						scx = scx & "<option value=" & ConvertirCar(array1(nx), " ", "_") & ">" & array2(nx) & "</option>"
					End If
				End If
			Next




			Case "2"

			array1 = GenerarArray(BSQL, BCon, BInd)
			Array= GenerarArray(BSQL, BCon, BDesc)

			sID = ValidarActualizacionCombo(Bform, bfnom)

			scx = "<option value=nsnc>&nbsp;</option>"

			For n = 1 To UBound(array1)
				If array1(n) <> "" Then
					If CStr(array1(n)) = CStr(bfval) Then
						scx = scx & "<option value=" & ConvertirCar(array1(n), " ", "_") & " selected>" & array2(n) & "</option>"
					Else
						scx = scx & "<option value=" & ConvertirCar(array1(n), " ", "_") & ">" & array2(n) & "</option>"
					End If
				End If
			Next


			Case "3"

			arrayx = GenerarArray(BSQL, BCon, bfnom)

			sID = ""

			bfnom = "lista"
			scx = ""
			If UBound(arrayx) > 0 Then
				For n = 1 To UBound(arrayx)
					scx = scx & "<option value=" & ConvertirCar(arrayx(n), " ", "_") & ">" & arrayx(n) & "</option>"
				Next
			End If


			Case "4"

			arrayx = GenerarArray(BSQL, BCon, bfnom2)

			sID = ""
			scx = ""
			For n = 1 To UBound(arrayx)
				scx = scx & "<option value=" & Chr(34) & ConvertirCar(arrayx(n), " ", "_")
				scx = scx & Chr(34) & ">" & arrayx(n) & "</option>"
			Next

			Case "5"

			arrayx1 = GenerarArray(BSQL, BCon, bfnom2)
			arrayx= GenerarArray(BSQL, BCon, BDesc)

			sID = ""

			scx = " "

			For n = 1 To UBound(arrayx1)
				scx = scx & "<option value=" & Chr(34) & ConvertirCar(arrayx1(n), " ", "_")
				scx = scx & Chr(34) & ">" & arrayx2(n) & "</option>"
			Next

			Case "6"

			arrayx1 = GenerarArray(BSQL, BCon, bfnom2)
			arrayx= GenerarArray(BSQL, BCon, BDesc)

			sID = ""

			scx = "<option value=nsnc>&nbsp;</option>"

			For n = 1 To UBound(arrayx1)
				scx = scx & "<option value=" & Chr(34) & ConvertirCar(arrayx1(n), " ", "_")
				scx = scx & Chr(34) & ">" & arrayx2(n) & "</option>"
			Next

			Case "7"

			array1 = GenerarArray(BSQL, BCon, BInd)
			Array= GenerarArray(BSQL, BCon, BDesc)

			sID = ""

			scx = "<option value=nsnc>&nbsp;</option>"

			For n = 1 To UBound(array1)
				If array1(n) <> "" Then
					If CStr(array1(n)) = CStr(bfval) Then
						scx = scx & "<option value=" & array1(n) & " selected>" & array2(n) & "</option>"
					Else
						scx = scx & "<option value=" & array1(n) & ">" & array2(n) & "</option>"
					End If
				End If
			Next

			Case "8"

			sID = ValidarActualizacionCombo(Bform, bfnom)

			scx = "<option value=nsnc>&nbsp;</option>"

			For n = 1 To UBound(BArray)
				If BArray(n) <> "" Then
					If UCase(CStr(BArray(n))) = UCase(CStr(bfval)) Then
						scx = scx & "<option value=" & Chr(34) & BArray(n) & Chr(34) & " selected >" & Replace(BArray(n), " ", "_") & "</option>"
					Else
						scx = scx & "<option value=" & Chr(34) & BArray(n) & Chr(34) & " >" & Replace(BArray(n), " ", "_") & "</option>"
					End If
				End If
			Next

			Case "9"

			sID = "document." & Bima & ".src='Iconos/' + this.value ;"

			scx = "<option value=nsnc>&nbsp;</option>"

			For n = 1 To UBound(BArray)
				If BArray(n) <> "" Then
					If UCase(CStr(BArray(n))) = UCase(CStr(bfval)) Then
						scx = scx & "<option value=" & Chr(34) & BArray(n) & Chr(34) & " selected >" & Replace(BArray(n), " ", "_") & "</option>"
					Else
						scx = scx & "<option value=" & Chr(34) & BArray(n) & Chr(34) & " >" & Replace(BArray(n), " ", "_") & "</option>"
					End If
				End If
			Next

			Case "a"

			sID = ""

			scx = "<option value=nsnc>&nbsp;</option>"

			For n = 1 To UBound(BArray)
				If BArray(n) <> "" Then
					If UCase(CStr(BArray(n))) = UCase(CStr(bfval)) Then
						scx = scx & "<option value=" & Chr(34) & BArray(n) & Chr(34) & " selected >" & Replace(BArray(n), " ", "_") & "</option>"
					Else
						scx = scx & "<option value=" & Chr(34) & BArray(n) & Chr(34) & " >" & Replace(BArray(n), " ", "_") & "</option>"
					End If
				End If
			Next

			Case "b"

			sID = " onchange=document." & Bform & ".submit() "

			If bfval Then
				scx = " <option selected value=true>Si</option>"
				scx = scx & "<option value=false>No</option>"
			Else
				scx = " <option  value=true>Si</option>"
				scx = scx & " <option selected value=false>No</option>"
			End If

			Case "c"

			sID = " onchange=document." & Bform & ".submit() "

			If bfval Then
				scx = " <option selected value=" & Bs & ">Si</option>"
				scx = scx & " <option value=" & Bn & ">No</option>"
			Else
				scx = " <option  value=" & Bs & ">Si</option>"
				scx = scx & " <option selected value=" & Bn & ">No</option>"
			End If

			Case "d"

			sID = ""
			scx = "<option value=True>Si</option>" & Chr(13)
			scx = scx & "<option value=False>No</option>" & Chr(13)


		End Select

		sx = " <font class=at><select " & sID & " name=" & bfnom & " id=" & bfnom & ">"
		sx = sx & scx
		sx = sx & "</select></font>"

		If UV = "SI" Then GenComboX = sx

		Exit Function

		GenComboXErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		GenComboX = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function


	''''''''''''''''''''''''








	Public Function DevParametros(parametros)

		'On Error GoTo DevParametrosErr

		Dim str, i

		For i = 0 To UBound(parametros)
			str = str & "&P" & CStr(i) & "=" & parametros(i)
		Next
		DevParametros = str

		Exit Function

		DevParametrosErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		DevParametros = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function ErrorTipo(volver, dato, columna)

		'On Error GoTo ErrorTipoErr

		Dim s, sMsg

		If dato = "" Then
			If columna = "" Then
				sMsg = "Hubo un error desconocido al grabar los datos</font></p>"
			Else
				sMsg = "</font><font color=#770000 face=verdana size=3><b>" & columna & "</font></b><font color=#000077 face=verdana size=2> no puede ser nulo</font></p>"
			End If
		Else
			sMsg = "El tipo de dato " & dato & "</font></b><font color=#770000 face=verdana size=3> no es válido para </font><b><font color=#000077 face=verdana size=2>" & columna & "</font></b></p>"
		End If

		s = "<html>" & FormStyler() & "<body><br><br><p align=center><font color=#000077 face=verdana size=2>Error</p><br><br><br><br><p align=center>"
		s = s & sMsg & "<p align=center><form action=" & volver & " method=post><input type=submit class=bt value=Volver></body></html>"

		ErrorTipo = s

		Exit Function

		ErrorTipoErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		ErrorTipo = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function Mensajes(sMsg, sVolver, nBotones)

		'On Error GoTo MensajesErr

		Dim s, sInput, Celda, Fila, Tabla, Formulario

		nBotones = CInt(nBotones)

		CheckUsuario

		s = FormHeader("Atencion")

		sInput = "<font color=#336699 face=verdana size=1><b>" & sMsg & "</b></font>"
		Celda = GenCelda("", "", "center", "100%", "", "", "", sInput, "si")
		Fila = GenRow("", "", "", "", "", "", "", Celda, "si")
		Tabla = s & GenTabla("", "", "", "100%", "", "0", "", "0", "0", "", "", Fila, "si")
		Tabla = Tabla & "<hr size=1 color=#000099>"



		If nBotones = 1 Then
			'  sInput = GenerarInput("", "Volver", "submit", "", "10", "bt", "")
			'  Celda = GenCelda("", "", "center", "50%", "", "", "", sInput, "si")
			'  Formulario = GenForm("Mensajes", "", sVolver, "", "post", "", Celda, "si")

			Formulario = GenBoton(7, "MenBot", "BT", "S", "", "S", "S", "N", "MenFrm", sVolver & "")

		Else
			' sInput = GenerarInput("", "Eliminar", "submit", "", "10", "bt", "")
			' Celda = GenCelda("", "", "center", "50%", "", "", "", sInput, "si")
			' Formulario = GenForm("Mensajes", "", sVolver & "&Eliminar=SI", "", "post", "", Celda, "si")

			Formulario = GenBoton(15, "MensBot", "BT", "S", "", "S", "S", "N", "MensFrm", sVolver & "&Eliminar=SI")

			' sInput = GenerarInput("", "Cancelar", "submit", "", "10", "bt", "")
			' Celda = GenCelda("", "", "center", "50%", "", "", "", sInput, "si")
			' Formulario = Formulario & GenForm("Mensajesb", "", sVolver & "&Eliminar=NO", "", "post", "", Celda, "si")

			Formulario = Formulario & GenBoton(14, "MennBot", "BT", "S", "", "S", "S", "N", "MennFrm", sVolver & "&Eliminar=NO")

		End If
		Fila = GenRow("", "", "", "", "", "", "", Formulario, "si")

		Tabla = Tabla & "<br><br><center>" & GenTabla("", "", "", "", "", "0", "", "0", "0", "", "", Fila, "si") & "</center></body></html>"

		Mensajes = Tabla

		Exit Function

		MensajesErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		Mensajes = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function VariableVacia(dato, Parametro)

		If Len(Trim(CStr(dato))) = 0 Then
			VariableVacia = Parametro
		Else
			VariableVacia = dato
		End If

	End Function

	Public Function ConvertirCar(cadena, car1, car2)

		Dim i, s

		For i = 1 To Len(cadena)
			If Mid(cadena, i, 1) = car1 Then
				s = s & car2
			Else
				s = s & Mid(cadena, i, 1)
			End If
		Next

		ConvertirCar = s

	End Function

	Public Sub ConvertirArregloParam(ByRef p, car1, car2)

		Dim n

		If IsArray(p) Then
			For n = 0 To UBound(p)
				p(n) = ConvertirCar(p(n), car1, car2)
			Next
		End If

	End Sub

	Public Function MenuHeader(sColor, sImagen)

		Dim s

		s = CorSetMHeader

		s = Replace(s, " XSFNK ", sImagen)



		MenuHeader = s

	End Function

	Public Function FormHeader(sTitulo)

		'On Error GoTo FormHeaderErr

		Dim n, s, s1

		s = CorSetFHeader

		s = Replace(s, " XSFNK ", "")

		s = s & CorFormHeader

		Styler = 1

		FormHeader = s

		Exit Function

		FormHeaderErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		FormHeader = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function ReportHeader(sTitulo)

		'On Error GoTo ReportHeaderErr

		Dim n, s

		n = 0
		sTitulo = UCase(sTitulo)
		n = InStr(sTitulo, "&")
		If n > 0 Then sTitulo = Left(sTitulo, n - 1) & Mid(sTitulo, n + 1, 1) & Mid(sTitulo, n + 7)
		s = "<html><Title>Normas y Procedimientos</title>"
		s = s & "<META HTTP-EQUIV=" & Chr(34) & "Content-Type" & Chr(34) & " CONTENT=" & Chr(34) & "text/html; charset=windows-1252" & Chr(34) & "><style>"
		s = s & "a{font-family:tahoma;color:#990000;font-size:7pt;text-decoration:none;}"
		s = s & "a.link{font-family:tahoma;color:#990000;font-size:7pt;text-decoration:none;}"
		s = s & "input{font-family:tahoma;color:#000077;font-size:7pt;background-color:#f7eedd;}"
		s = s & "input.bt{font-family:tahoma;color:#ffffff;font-size:7pt;background-color:#336699;}"
		s = s & "select{font-family:tahoma;color:#000077;font-size:7pt;background-color:#f7eedd;}"
		s = s & "img{border-style:solid,border-width:thin;}"
		s = s & "td.td1{background-color:#f0e68c;border-style:solid;border-width:thin;text-align:center;color=#696969;font-family:verdana;font-size:9pt;font-weight:bold}"
		s = s & "td.td2{text-align=center;background-color:#f7eedd;border-style:solid;border-width:thin;color=#696969;font-family:verdana;font-size:7pt;font-weight:normal}"
		s = s & "a.a2{background-color:#f7eedd;color=#696969;font-family:verdana;font-size:7pt;font-weight:normal}"
		s = s & "td.td3{text-align=center;background-color:#d0d099;border-style:solid;border-width:thin;color=#696969;font-family:verdana;font-size:7pt;font-weight:normal}"
		s = s & "</style>"
		s = s & "<body bgcolor=#ffffff><table width=100% marginwidth=0 cellspacing=" & wspacing
		s = s & " cellpadding=" & 1 + wspacing & " align=center>"
		s = s & "<tr align=top>"
		If Session("CliLogo") <> "" Then s = s & "<td width=4% ><img src=" & Session("CliLogo") & " border=0 height=22></td>"
		s = s & "<td align=left ><p><b><font face=arial,"
		s = s & "tahoma color=#880000 size=1>" & sTitulo
		s = s & "</font></b></p></td></tr></table><hr width=100% color=#000088 size=1><br><br>"

		ReportHeader = s

		Exit Function

		ReportHeaderErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		ReportHeader = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function Repbody(sTitulo)

		Dim s

		s = "<html><Title>COR EManager - " & sTitulo & "</title>"
		s = s & "<meta HTTP-EQUIV=" & Chr(34) & "expires" & Chr(34) & " CONTENT=" & Chr(34) & "60" & Chr(34) & " >" & FormStylerCon()
		s = s & "<body style={display:" & Chr(34) & "table-header-group" & Chr(34) & "} bgcolor=#ffffff>"

		Repbody = s

	End Function

	Public Function RepHeader2(sTitulo, nAncho)

		'On Error GoTo RepHeader2Err

		Dim s, Fecha

		If nAncho = "" Then nAncho = "100%"

		s = "<table width=" & nAncho & " marginwidth=0 cellspacing=0 cellpadding=0>"
		s = s & "<tr align=top>"
		If Session("CliLogo") <> "" Then s = s & "<td width=4% ><img src=" & Session("CliLogo") & " border=0 height=22></td>"
		s = s & "<td class=ut align=left width=55% >" & Session("Cliente") & "</td></tr></table>"
		s = s & "<br><table width=" & nAncho & " marginwidth=0 cellspacing=0 cellpadding=0>"
		s = s & "<tr align=top><td class=qt align=left width=60% >" & sTitulo & "</td>"
		Fecha = Fechador()
		s = s & "<td class=ut align=right>" & Fecha & "</td></tr></table>"
		s = s & "<hr color=#000088 size=1 width=" & nAncho & " align=left>"

		RepHeader= s

		Exit Function

		RepHeader2Err:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		RepHeader= oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function GetscriptNSC()

		'On Error GoTo GetscriptNSCErr

		Dim Ascript

		Ascript = "<script language=javascript >" & Chr(13)
		Ascript = Ascript & " window.captureEvents(Event.MOUSEDOWN);"
		Ascript = Ascript & Chr(13) & " window.onMouseDown=Wedita;"
		Ascript = Ascript & Chr(13) & " " & Chr(13)
		Ascript = Ascript & " function Wedita(e)" & Chr(13)
		Ascript = Ascript & " { " & Chr(13) & " if (e.which == 3) { "
		Ascript = Ascript & Chr(13) & " if (e.target.name != null) { "
		Ascript = Ascript & Chr(13) & " myWin= open(" & Chr(34)
		Ascript = Ascript & "Edicion.asp?ID=" & Chr(34)
		Ascript = Ascript & " + e.target.name , " & Chr(34) & "Edicion"
		Ascript = Ascript & Chr(34) & " , " & Chr(34)
		Ascript = Ascript & "width=280,height=320,status=no,toolbar=no,"
		Ascript = Ascript & "menubar=no,resizable=yes,screenX="
		Ascript = Ascript & Chr(34) & " + (e.pageX + 400) + " & Chr(34)
		Ascript = Ascript & ",screenY=" & Chr(34) & " + e.pageY  "
		Ascript = Ascript & " ); " & Chr(13) & " return false}"
		Ascript = Ascript & Chr(13) & " } " & Chr(13) & " } " & Chr(13)
		Ascript = Ascript & "  " & Chr(13) & " </script> " & Chr(13)

		GetscriptNSC = Ascript

		Exit Function

		GetscriptNSCErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		GetscriptNSC = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function FormHeaderX(sTitulo, volver)

		'On Error GoTo FormHeaderXErr

		Dim WScript, s

		volver = Chr(34) & volver & Chr(34)
		sTitulo = UCase(sTitulo)
		If Explorador <> "MSIE" Then WScript = GetscriptNSC()

		s = "<html><Title>Cor E-Solution Suite</title>" & FormStyler() & WScript

		s = s & "<SCRIPT language=" & Chr(34) & "VBScript" & Chr(34) & ">" & Chr(13)
		s = s & "function EditarMenu(Elemento)" & Chr(13)
		s = s & "if elemento = " & Chr(34) & Chr(34) & " Then " & Chr(13)
		s = s & "exit function" & Chr(13)
		s = s & "end if" & Chr(13)
		s = s & "dato = " & Chr(34) & "MenuAtributo.asp?ID=" & Chr(34) & " & elemento " & Chr(13)
		s = s & "window.showModalDialog dato"
		s = s & ",0," & Chr(34) & "dialogWidth:420px;dialogHeight:140px;help:no;"
		s = s & "resizable:yes;dialogTop:240px;dialogLeft:300px;status:no" & Chr(34) & Chr(13)
		s = s & "end function" & Chr(13)

		s = s & "function EditarPantalla(Elemento)" & Chr(13)
		s = s & "if elemento = " & Chr(34) & Chr(34) & " Then " & Chr(13)
		s = s & "exit function" & Chr(13)
		s = s & "end if" & Chr(13)
		s = s & "dato = " & Chr(34) & "PantallaAtributo.asp?ID=" & Chr(34) & " & elemento " & Chr(13)
		s = s & "window.showModalDialog dato"
		s = s & ",0," & Chr(34) & "dialogWidth:420px;dialogHeight:140px;help:no;"
		s = s & "resizable:yes;dialogTop:240px;dialogLeft:300px;status:no" & Chr(34) & Chr(13)
		s = s & "end function" & Chr(13)

		s = s & "function EliminarMenu(Elemento)" & Chr(13)
		s = s & "if elemento = " & Chr(34) & Chr(34) & " Then " & Chr(13)
		s = s & "exit function" & Chr(13)
		s = s & "end if" & Chr(13)
		s = s & "dato = " & Chr(34) & "asp/EliminarAtributo.asp?ID=" & Chr(34) & " & elemento " & Chr(13)
		s = s & "window.showModalDialog dato"
		s = s & ",0," & Chr(34) & "dialogWidth:420px;dialogHeight:140px;help:no;"
		s = s & "resizable:yes;dialogTop:240px;dialogLeft:300px;status:no" & Chr(34) & Chr(13)
		s = s & "end function" & Chr(13)

		s = s & "function EliminarPantalla(Elemento)" & Chr(13)
		s = s & "if elemento = " & Chr(34) & Chr(34) & " Then " & Chr(13)
		s = s & "exit function" & Chr(13)
		s = s & "end if" & Chr(13)
		s = s & "dato = " & Chr(34) & "asp/EliminarAtributo.asp?ID=" & Chr(34) & " & elemento " & Chr(13)
		s = s & "window.showModalDialog dato"
		s = s & ",0," & Chr(34) & "dialogWidth:420px;dialogHeight:140px;help:no;"
		s = s & "resizable:yes;dialogTop:240px;dialogLeft:300px;status:no" & Chr(34) & Chr(13)
		s = s & "end function" & Chr(13) & "</SCRIPT>" & Chr(13)

		s = s & "<body " & wbody & " >"

		'''2 s = s & "<table width=100% marginwidth=0 cellspacing=" & wspacing & " cellpadding="
		'''2s = s & 1 + wspacing & "><tr align=top>"
		'''2   s = s & GenCelda("", "cc", "center", "35%", "", "", "", "<b>" & sTitulo & "</b>", "si")
		'''s = s & GenCelda("", "cc", "right", "", "", "", "", "Usuario: " & Session("SegUsuario") & " ", "si")
		'''2s = s & "</tr></table>"
		'''
		'''3050    If Explorador = "MSIE" Then s = s & "<hr color=#000088 size=1>"

		s = s & CorFormHeader


		FormHeaderX = s

		Exit Function

		FormHeaderXErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		FormHeaderX = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function FormHeader3(sTitulo)

		'On Error GoTo FormHeader3Err

		Dim s

		s = FormHeader(sTitulo)

		s = s & "<SCRIPT language=" & Chr(34) & "VBScript" & Chr(34)
		s = s & ">" & Chr(13) & "function ActualizarCombo(ID,ID2)"
		s = s & Chr(13) & "ID2.selectedindex=ID.selectedindex" & Chr(13)
		s = s & "end function" & Chr(13) & "</SCRIPT>" & Chr(13)

		FormHeader3 = s

		Exit Function

		FormHeader3Err:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		FormHeader3 = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function FormHeader2(sTitulo, volver)

		'On Error GoTo FormHeader2Err

		Dim WScript, s, s1, YPOS

		volver = Chr(34) & volver & Chr(34)
		sTitulo = UCase(sTitulo)
		If Explorador <> "MSIE" Then WScript = GetscriptNSC()

		s = "<html><Title>Cor E-Solution Suite</title>" & FormStyler() & WScript

		s = s & "<SCRIPT language=" & Chr(34) & "VBScript" & Chr(34) & ">" & Chr(13)
		s = s & "function Editar(Elemento,YPos,Volver)" & Chr(13)
		s = s & "if elemento = " & Chr(34) & Chr(34) & " Then " & Chr(13)
		s = s & "exit function" & Chr(13)
		s = s & "end if" & Chr(13)
		s = s & "dato = " & Chr(34) & "Edicion.asp?ID=" & Chr(34) & " & elemento "
		s = s & Chr(13)
		s = s & "window.showModalDialog dato"
		s = s & ",0," & Chr(34) & "dialogWidth:280px;dialogHeight:320px;help:no;"
		s = s & "resizable:yes;dialogTop:" & YPOS + 250 & "px;dialogLeft:300px;status:no"
		s = s & Chr(34) & Chr(13)
		s = s & "end function" & Chr(13) & "</SCRIPT>" & Chr(13)

		s = s & "<body " & wbody & " onmousedown = " & Chr(34) & "if window.event.button=then "
		s = s & "Editar window.event.srcElement.ID,window.event.clienty,volver end if"
		s = s & Chr(34) & ">"

		''' s = s & GenTabla("headm", "", "", "100%", "", 0, 0, 0, 0, "", "", "", "no")
		'''
		''' s1 = GenCelda("", "cc", "center", "35%", "", "", "", "<b>" & sTitulo & "</b>", "si")
		'''s1 = s1 & GenCelda("", "cc", "right", "", "", "", "", "Usuario: " & Session("SegUsuario") & " ", "si")
		'''
		'''
		'''
		'''1   s = s & GenRow("", "", "top", "", "", "", "", s1, "si")
		'''1 s = s & GenFTabla()
		'''
		'''  If Explorador = "MSIE" Then s = s & "<hr color=#000088 size=1>"

		s = s & CorFormHeader


		FormHeader= s

		Exit Function

		FormHeader2Err:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		FormHeader= oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function FormBody()

		Dim s

		s = "<html><Title>Cor E-Solution Suite</title>"
		s = s & FormStyler() & "<body " & wbody & " >"

		FormBody = s

	End Function

	Public Function FormBodyx()

		'On Error GoTo FormBodyxErr

		Dim s

		s = "<html><Title>Cor E-Solution Suite</title>"
		s = s & Chr(13) & "<script language=" & Chr(34) & "VBScript"
		s = s & Chr(34) & ">" & Chr(13) & "function Ayuda(ID)" & Chr(13)
		s = s & "if ID=63 then" & Chr(13) & "sURL = " & Chr(34)
		s = s & "Ayuda/Info.htm" & Chr(34) & Chr(13)
		s = s & "window.open sURL" & Chr(13) & "end if" & Chr(13)
		s = s & "end function" & Chr(13) & "</script>" & Chr(13)
		s = s & FormStyler() & "<body " & wbody
		s = s & " onkeypress=Ayuda(window.event.keycode)>"

		FormBodyx = s

		Exit Function

		FormBodyxErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		FormBodyx = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function FormBody2()

		'On Error GoTo FormBody2Err

		Dim s

		s = FormBody()
		s = s & "<SCRIPT language=" & Chr(34) & "JavaScript" & Chr(34) & ">" & Chr(13)
		s = s & "function AccionesAlt(ID)" & Chr(13)
		s = s & "{ sURL = 'AccionesAlt.asp?' + ID; " & Chr(13)
		s = s & "myWin= open(sURL"
		s = s & "," & Chr(34) & "Principal" & Chr(34) & "," & Chr(34) & "height=200,width=480,location=no,resizable=no,titlebar=0"
		s = s & Chr(34) & ")"
		s = s & "; }" & Chr(13)

		s = s & "function AccionesCon(ID)" & Chr(13)
		s = s & "{ sURL = 'CoAccionesAlt.asp?' + ID; " & Chr(13)
		s = s & "myWin= open(sURL"
		s = s & "," & Chr(34) & "Principal" & Chr(34) & "," & Chr(34) & "height=200,width=480,location=no,resizable=no,titlebar=0"
		s = s & Chr(34) & ")"
		s = s & "; }" & Chr(13)

		s = s & "function TextoEsquema(ID)" & Chr(13)
		s = s & "{ sURL = 'TextoEsquema.asp?' + ID; " & Chr(13)
		s = s & " myWin2= open(sURL"
		s = s & "," & Chr(34) & "Principal" & Chr(34) & "," & Chr(34) & "height=200,width=480,location=no,resizable=no,titlebar=0"
		s = s & Chr(34) & ")"
		s = s & "; }" & Chr(13) & "</SCRIPT>" & Chr(13)

		FormBody= s

		Exit Function

		FormBody2Err:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		FormBody= oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function QueESx(Valor)

		Dim st

		st = ""
		Select Case Valor
			Case 8, 135, 202
			st = "'"
		End Select

		QueESx = st

	End Function

	Public Function GetValor(Conexion, SQL)

		'  Set rst = CrearConsultaEx(SQL, cn, 1, "")
		'
		'  If Not rst.EOF Then GetValor = rst(0)
		'
		'  Call CerrarConsulta(rst)

	End Function

	Public Function ValidarTipo(Campo, Valor)

		'On Error GoTo ValidarTipoErr

		Dim tipo

		ValidarTipo = True

		tipo = VarType(Campo)
		Select Case tipo
			Case 2, 3, 4, 5
			If Not IsNumeric(Valor) Then ValidarTipo = False
			Case 7
			If Not IsDate(Valor) Then ValidarTipo = False
			Case 11
			If UCase(Valor) <> "TRUE" And UCase(Valor) <> "FALSE" Then ValidarTipo = False
		End Select
		If Valor = "nsnc" Then ValidarTipo = True

		Exit Function

		ValidarTipoErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		ValidarTipo = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function EsNulo(Campo, Valor)

		'On Error GoTo EsNuloErr

		Dim tipo

		EsNulo = False

		tipo = VarType(Campo)
		If tipo = 8 Then
			If Valor = "" Then EsNulo = True
		ElseIf tipo = 3 Then
			If Valor = "" Then EsNulo = True
		End If

		Exit Function

		EsNuloErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		EsNulo = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function GenerarArray(SQL, Conexion, Campo)

		' On Error resume next

		Dim rsArray
		Dim n

		Set rsArray = CrearConsultaEx(SQL, Conexion, 1, "")

		n = 1
		ReDim s(0)
		Do Until rsArray.EOF
			ReDim Preserve s(n)
			s(n) = rsArray(Campo) & ""
			rsArray.movenext
			n = n + 1
		Loop
		Call CerrarConsulta(rsArray)

		GenerarArray = s

		Exit Function

		GenerarArrayErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		GenerarArray = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function GetFiles(Directorio)

		'On Error GoTo GetFilesErr

		Dim fso , Folder , Files
		Dim F
		Dim sPath, n

		Set fso = server.CreateObject("Scripting.FileSystemObject")

		sPath = Server.mappath("1")
		sPath = Left(sPath, Len(sPath) - 1) & Directorio

		Set Folder = fso.getfolder(sPath)
		Set Files = Folder.Files

		ReDim s(0)
		n = 1
		For Each F In Files
			ReDim Preserve s(n)
			s(n) = F.Name
			n = n + 1
		Next
		GetFiles = s

		Set Files = Nothing
		Set Folder = Nothing
		Set fso = Nothing

		Exit Function

		GetFilesErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		GetFiles = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function AgregarBoton(Form, Accion, sTarget)

		'On Error GoTo AgregarBotonErr

		Dim s, s1, s2

		If sTarget = "" Then sTarget = "info"



		s = GenFForm()

		s = s & "<br><br><br><center>"

		s = s & GenTabla("agreg", "", "", "", "", "0", "0", "0", "0", "", "", "", "no")


		''''   s1 = GenBoton(13, "AcepBot", "BT", "S", "", "N", "S", "N", Form & "", "")



		If sTarget = "no" Then

			s1 = GenBoton(13, "AcepBot", "BT", "N", " onclick='document." & Form & ".submit();' ", "N", "S", "N", "", "")
			s= GenBoton(14, "CancBot", "BT", "N", " onclick='window.close();' ", "N", "S", "N", "", "")

		Else

			s1 = GenBoton(13, "AcepBot", "BT", "N", " onclick='document." & Form & ".submit();' ", "S", "S", "N", "AcepFRM", "")
			s= GenBoton(14, "CancBot", "BT", "S", "", "S", "S", "N", "CancFrm", Accion & "")

		End If

		s = s & GenRow("", "", "top", "", "", "", "", s1 & s2, "si") & GenFTabla() & "</center>"


		AgregarBoton = s

		Exit Function

		AgregarBotonErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		AgregarBoton = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function BotonAceptarCancelar(Form, Accion)

		'On Error GoTo BotonAceptarCancelarErr

		Dim s1, s2

		s1 = GenCelda("", "", "center", "", "", "", "", GenerarInput("", "Aceptar", "submit", "", "10", "bt", ""), "si") & GenFForm()

		s= GenCelda("", "", "center", "", "", "", "", GenerarInput("", "Cancelar", "submit", "", "10", "bt", ""), "si")
		s= GenForm(Form, "", "", "", "post", "", s2, "si")

		BotonAceptarCancelar = s1 & s2

		Exit Function

		BotonAceptarCancelarErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		BotonAceptarCancelar = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function BotonEliminar(sNombre, sHref)

		'On Error GoTo BotonEliminarErr

		Dim sImagen, sLink, Celda

		FBFcnt = FBFcnt + 1

		sNombre = ConvertirCar(sNombre, " ", "&nbsp;")
		sHref = sHref & "&Nombre=" & sNombre

		'  sImagen = GenImage("", "", "15", "15", "imagenes/botdel.jpg", "0", "")
		'  sLink = GenLink("", "", sHref, "", "Eliminar&nbsp;" & sNombre, "", sImagen, "si")
		'  Celda = GenCelda("", "", "", "", "", "", "", sLink, "si")

		sImagen = FBotTab(8, 4)

		FBotTab(8, 4) = FBotTab(8, 4) & "&nbsp;" & sNombre

		Celda = GenBoton(8, "ElimBot", "BT", "S", "", "S", "A", "N", "ElimFrm" & CStr(FBFcnt), sHref & "")

		FBotTab(8, 4) = sImagen

		BotonEliminar = Celda

		Exit Function

		BotonEliminarErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		BotonEliminar = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function BotonEliminar2(sNombre, sHref, sTarget)

		'On Error GoTo BotonEliminar2Err

		Dim sImagen, sLink, Celda

		FBFcnt = FBFcnt + 1

		sNombre = ConvertirCar(sNombre, " ", "&nbsp;")
		sHref = sHref & "&Nombre=" & sNombre

		'   sImagen = GenImage("", "", "15", "15", "imagenes/botdel.jpg", "0", "")
		'   sLink = GenLink("", "", sHref, sTarget, "Eliminar&nbsp;" & sNombre, "", sImagen, "si")
		'   Celda = GenCelda("", "", "", "", "", "", "", sLink, "si")

		sImagen = FBotTab(8, 4)
		sLink = FBotTab(8, 5)

		FBotTab(8, 4) = FBotTab(8, 4) & "&nbsp;" & sNombre
		FBotTab(8, 5) = sTarget

		Celda = GenBoton(8, "ElimBot", "BT", "S", "", "S", "A", "N", "ElimFrm" & CStr(FBFcnt), sHref & "")

		FBotTab(8, 4) = sImagen
		FBotTab(8, 5) = sLink


		BotonEliminar= Celda

		Exit Function

		BotonEliminar2Err:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		BotonEliminar= oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function BotonAgregar(sAccion)

		'On Error GoTo BotonAgregarErr

		Dim sInput, Celda, x1, x2


		FBmodo = GenBoton(3, "AgrBot", "BT", "S", "", "S", "A", "N", "AgrFrm", sAccion & "")


		BotonAgregar = FBmodo


		Exit Function

		BotonAgregarErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		BotonAgregar = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function


	Public Function GenBoton(sBInd , sBNom , sBClas , sBFun , sBAdFun , sBFrm , sBTd , sBAn , sBFname , sBAccion )

		'On Error GoTo GenBotonErr

		Dim sInput, Celda, x1, x2, xFBmodo, xFFun, xAnch, xAnch2

		xFBmodo = ""

		x1 = ""
		x= ""

		xFFun = ""
		xAnch = ""
		xAnch= ""

		If sBFun = "S" Then
			xFFun = " onclick=""document.getElementById('" & sBFname & "').submit();"""
		End If



		If sBTd = "S" Or sBTd = "A" Then

			If ModBot = "2" Then
				x1 = GenCelda("", "", "", "", "", "", "", "", "no")
			Else
				x1 = GenCelda(sBNom & "Td", FBotTab(sBInd, 3) & " style='cursor:hand;' " & xFFun & sBAdFun & " title='" & FBotTab(sBInd, 4) & "' ", "", "", "", "", "", "", "no")
				''    x   = x1 & " title='" & FBotTab(sBInd, 4) & "' ", "", "", "", "", "", "", "no")
			End If

			x= GenFCelda()

		End If

		If sBAn = "S" Or sBAn = "A" Then
			xAnch = "<a class=" & FBotTab(sBInd, 3) & " name='" & sBNom & "An' id='" & sBNom & "An' "
			xAnch = xAnch & " href='javascript:void(0);' "

			If ModBot = "2" Then
				xAnch = xAnch & " "
			Else
				xAnch = xAnch & xFFun & sBAdFun
			End If

			xAnch = xAnch & " title='" & FBotTab(sBInd, 4) & "' >"
			xAnch= "</a>"

		End If


		If ModBot = "1" Or ModBot = "5" Then

			xFBmodo = FBotTab(sBInd, 0)

		ElseIf ModBot = "2" Then

			If sBAn = "A" Or sBTd = "A" Or sBFun <> "S" Then
				Celda = "button"
				sInput = FBotTab(sBInd, 3) & " " & xFFun & sBAdFun & " "
			Else
				Celda = "submit"
				sInput = FBotTab(sBInd, 3) & " " & sBAdFun & " "
			End If

			xFBmodo = GenerarInput(sBNom, FBotTab(sBInd, 6), Celda, "", "10", sInput, "")

		Else

			xFBmodo = FBotTab(sBInd, 6)

		End If


		xFBmodo = Replace(xFBmodo, "XXXFONTXXX", FBotTab(sBInd, 2))
		xFBmodo = Replace(xFBmodo, "XXXDATOXXX", FBotTab(sBInd, 1))

		If sBFrm = "S" Then

			If ModBot = "2" Then

				If sBAn = "A" Or sBTd = "A" Then
					FBBarr = FBBarr & GenForm(sBFname, "BBB style='display:none;' ", sBAccion, FBotTab(sBInd, 5), "post", "", GenerarInput(sBNom & "Bb", "no", "hidden", "", "10", FBotTab(sBInd, 3), ""), "si") & Chr(13)
				Else
					xFBmodo = GenForm(sBFname, "", sBAccion, FBotTab(sBInd, 5), "post", "", xFBmodo, "si") & Chr(13)

				End If


			Else
				If sBAn = "A" Or sBTd = "A" Then
					FBBarr = FBBarr & GenForm(sBFname, "BBB style='display:none;' ", sBAccion, FBotTab(sBInd, 5), "post", "", GenerarInput(sBNom & "Bb", "no", "hidden", "", "10", FBotTab(sBInd, 3), ""), "si") & Chr(13)
				Else
					xFBmodo = xFBmodo & GenForm(sBFname, "BBB style='display:none;' ", sBAccion, FBotTab(sBInd, 5), "post", "", GenerarInput(sBNom & "Bb", "no", "hidden", "", "10", FBotTab(sBInd, 3), ""), "si") & Chr(13)
				End If
			End If

		End If



		GenBoton = x1 & xAnch & xFBmodo & xAnch& x2


		Exit Function

		GenBotonErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		GenBoton = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function BotonTest(sAccion)

		'On Error GoTo BotonTestErr


		FBmodo = ""

		FBmodo = GenBoton(6, "TestBot", "BT", "S", "", "S", "A", "N", "TestFrm", sAccion & "")



		BotonTest = FBmodo


		Exit Function

		BotonTestErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		BotonTest = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function BotonConsultar(sAccion)

		'On Error GoTo BotonConsultarErr


		FBmodo = ""

		FBmodo = GenBoton(5, "ConsBot", "BT", "S", "", "S", "A", "N", "ConsFrm", sAccion & "")






		BotonConsultar = FBmodo


		Exit Function

		BotonConsultarErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		BotonConsultar = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function BotonVolver(sAccion)

		'On Error GoTo BotonVolverErr


		FBmodo = ""


		FBmodo = GenBoton(7, "VolvBot", "BT", "S", "", "S", "A", "N", "VolvFrm", sAccion & "")






		BotonVolver = FBmodo


		Exit Function

		BotonVolverErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		BotonVolver = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function BotonEsp(sAccion, sInd, sDato, sTitulo, sTarget)

		'On Error GoTo BotonEspErr

		Dim xsDato, xsTitulo, xsTarget

		FBmodo = ""


		If sDato = "" Then sDato = "N"
		If sTitulo = "" Then sTitulo = "N"
		If sTarget = "" Then sTarget = "N"


		If sDato <> "N" Then
			xsDato = FBotTab(CLng(sInd), 6)
			FBotTab(CLng(sInd), 6) = sDato
		End If
		If sTitulo <> "N" Then
			xsTitulo = FBotTab(CLng(sInd), 4)
			FBotTab(CLng(sInd), 4) = sTitulo
		End If
		If sTarget <> "N" Then
			xsTarget = FBotTab(CLng(sInd), 5)
			FBotTab(CLng(sInd), 5) = sTarget
		End If

		FBmodo = GenBoton(CLng(sInd), "EspBot" & CStr(sInd), "BT", "S", "", "S", "S", "N", "EspFrm" & CStr(sInd), sAccion & "")


		If sDato <> "N" Then
			FBotTab(CLng(sInd), 6) = xsDato
		End If
		If sTitulo <> "N" Then
			FBotTab(CLng(sInd), 4) = xsTitulo
		End If
		If sTarget <> "N" Then
			FBotTab(CLng(sInd), 5) = xsTarget
		End If



		BotonEsp = FBmodo


		Exit Function

		BotonEspErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		BotonEsp = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

Public Function Encabezados(  vParametros() )

       ' On Error GoTo EncabezadosErr

     Dim v, i,  iMax,  Celda,  xTitCls,  xGenNom,  xGenCont,  xGenWidth,  sHtml

     If TitCls = "" Then TitCls = "tt"

     xTitCls = TitCls

     xGenCont = "*"
     xGenNom = "xta"

     If ModTitCls = "1" Or ModTitCls = "5" Or ModTitCls = "0" Then
        xGenCont = FTitTab(CLng(ModTitCls), 0)
        If ModTitCls = "1" Or ModTitCls = "0" Then
           xGenWidth = "10px"
           xGenCont = Replace(xGenCont, "XXXDATOXXX", "*")
           xGenCont = Replace(xGenCont, "XXXWIDTHXXX", xGenWidth)
           xGenCont = Replace(xGenCont, "XXXIDXXX", "r" & xGenNom)
        End If
     End If



     Celda = GenCelda(xGenNom, xTitCls, "", "2%", "", "", "", xGenCont, "si")

     i = 0
     iMax = UBound(vParametros)
     For Each v In vParametros
         xGenNom = "xt" & CStr(i)
         v = ConvertirCar(v, "_", " ")

         If Not IsNumeric(v) Then
            xGenCont = v

            If ModTitCls = "1" Or ModTitCls = "5" Or ModTitCls = "0" Then

               xGenCont = FTitTab(CLng(ModTitCls), 0)

               If ModTitCls = "1" Or ModTitCls = "0" Then

                  xGenWidth = CStr(Len(v) * 10) & "px"
                  xGenCont = Replace(xGenCont, "XXXDATOXXX", v)
                  xGenCont = Replace(xGenCont, "XXXWIDTHXXX", xGenWidth)
                  xGenCont = Replace(xGenCont, "XXXIDXXX", "r" & xGenNom)

               End If

            End If

         End If




         If (i + 1) > iMax Then

             If Not IsNumeric(v) Then

                Celda = Celda & GenCelda(xGenNom, xTitCls, "", "2%", "", "", "", xGenCont, "si")
             End If

             Exit For
         End If

         If IsNumeric(vParametros(i + 1)) Then

             If Not IsNumeric(v) Then

                Celda = Celda & GenCelda(xGenNom, xTitCls & " colspan=" & CInt(vParametros(i + 1)), "", "2%", "", "", "", xGenCont, "si")
             End If

        Else

            If Not IsNumeric(v) Then

               Celda = Celda & GenCelda(xGenNom, xTitCls, "", "2%", "", "", "", xGenCont, "si")
            End If

        End If

        i = i + 1
    Next

    Celda = Celda & GenCelda("xtfin", "", "", "100%", "", "", "", "", "si")


    If iMax > 0 Then

       If ModTitCls = "1" Then

          sHtml = ""
          sHtml = sHtml & "<script  language=javascript>" & Chr(13)
          sHtml = sHtml & ""
          sHtml = sHtml & " var fbcnt = 0; " & Chr(13)
          sHtml = sHtml & " fbrsz(); " & Chr(13)
          sHtml = sHtml & " function fbrsz() {  " & Chr(13)
          sHtml = sHtml & ""
          sHtml = sHtml & " if (document.all.r" & xGenNom & ") { fbcalc('this'); " & Chr(13)
          sHtml = sHtml & " } else {  if (fbcnt > 10) {} else { fbcnt++; " & Chr(13)
          sHtml = sHtml & " setTimeout (" & Chr(34) & "fbrsz()" & Chr(34) & ", 350); }}" & Chr(13)
          sHtml = sHtml & " } " & Chr(13)
          sHtml = sHtml & ""
          sHtml = sHtml & " function fbcalc(fbcell) {"
          sHtml = sHtml & "    var count=0;"
          sHtml = sHtml & "    var NRR='no';"
          sHtml = sHtml & "        for (j=0; j < document.all.XFKR.cells.length-1; j++) {"
          sHtml = sHtml & "            NRR='r' + document.all.XFKR.cells(j).id;"
          sHtml = sHtml & "            document.all[NRR].style.width = document.all.XFKR.cells(j).clientWidth;"
          sHtml = sHtml & "            count++;"
          sHtml = sHtml & "        }"
          sHtml = sHtml & " }"
          sHtml = sHtml & " </script>" & Chr(13)

          FBBarr = FBBarr & sHtml

       End If
    End If



    Encabezados = GenRow("XFKR", "XFK", "", "", "", "", "", Celda, "si")

    Exit Function

 EncabezadosErr:
    With oErr
        .Number = Err.Number
        .Desc = Err.Description
        .Source = Err.Source
        .Line = Erl
        .Version = App.Major & "." & App.Minor & "." & App.Revision
    End With

    On Error Resume Next

    Encabezados = oEH.RegError(oErr)

    If Err.Number <> 0 Then Response.WRITE Err.Number

End Function

	Public Function Encabezadows(  )

		'On Error GoTo EncabezadosErr

		Dim v, i, iMax, Celda, xTitCls, xGenNom, xGenCont, xGenWidth, sHtml

		If TitCls = "" Then TitCls = "tt"

		xTitCls = TitCls

		xGenCont = "*"
		xGenNom = "xta"
		response.write CStr(vParametros)
		If ModTitCls = "1" Or ModTitCls = "5" Or ModTitCls = "0" Then
			xGenCont = FTitTab(CLng(ModTitCls), 0)
			If ModTitCls = "1" Or ModTitCls = "0" Then
				xGenWidth = "10px"
				xGenCont = Replace(xGenCont, "XXXDATOXXX", "*")
				xGenCont = Replace(xGenCont, "XXXWIDTHXXX", xGenWidth)
				xGenCont = Replace(xGenCont, "XXXIDXXX", "r" & xGenNom)
			End If
		End If



		Celda = GenCelda(xGenNom, xTitCls, "", "2%", "", "", "", xGenCont, "si")

		i = 0
		iMax = UBound(vParametros)
		For Each v In vParametros
			xGenNom = "xt" & CStr(i)
			v = ConvertirCar(v, "_", " ")

			If Not IsNumeric(v) Then
				xGenCont = v

				If ModTitCls = "1" Or ModTitCls = "5" Or ModTitCls = "0" Then

					xGenCont = FTitTab(CLng(ModTitCls), 0)

					If ModTitCls = "1" Or ModTitCls = "0" Then

						xGenWidth = CStr(Len(v) * 10) & "px"
						xGenCont = Replace(xGenCont, "XXXDATOXXX", v)
						xGenCont = Replace(xGenCont, "XXXWIDTHXXX", xGenWidth)
						xGenCont = Replace(xGenCont, "XXXIDXXX", "r" & xGenNom)

					End If

				End If

			End If




			If (i + 1) > iMax Then

				If Not IsNumeric(v) Then

					Celda = Celda & GenCelda(xGenNom, xTitCls, "", "2%", "", "", "", xGenCont, "si")
				End If

				Exit For
			End If

			If IsNumeric(vParametros(i + 1)) Then

				If Not IsNumeric(v) Then

					Celda = Celda & GenCelda(xGenNom, xTitCls & " colspan=" & CInt(vParametros(i + 1)), "", "2%", "", "", "", xGenCont, "si")
				End If

			Else

				If Not IsNumeric(v) Then

					Celda = Celda & GenCelda(xGenNom, xTitCls, "", "2%", "", "", "", xGenCont, "si")
				End If

			End If

			i = i + 1
		Next

		Celda = Celda & GenCelda("xtfin", "", "", "100%", "", "", "", "", "si")


		If iMax > 0 Then

			If ModTitCls = "1" Then

				sHtml = ""
				sHtml = sHtml & "<script  language=javascript>" & Chr(13)
				sHtml = sHtml & ""
				sHtml = sHtml & " var fbcnt = 0; " & Chr(13)
				sHtml = sHtml & " fbrsz(); " & Chr(13)
				sHtml = sHtml & " function fbrsz() {  " & Chr(13)
				sHtml = sHtml & ""
				sHtml = sHtml & " if (document.getElementById('.r" & xGenNom & "') { fbcalc('this'); " & Chr(13)
				sHtml = sHtml & " } else {  if (fbcnt > 10) {} else { fbcnt++; " & Chr(13)
				sHtml = sHtml & " setTimeout (" & Chr(34) & "fbrsz()" & Chr(34) & ", 350); }}" & Chr(13)
				sHtml = sHtml & " } " & Chr(13)
				sHtml = sHtml & ""
				sHtml = sHtml & " function fbcalc(fbcell) {"
				sHtml = sHtml & "    var count=0;"
				sHtml = sHtml & "    var NRR='no';"
				sHtml = sHtml & "        for (j=0; j < document.getElementById('XFKR').cells.length-1; j++) {"
				sHtml = sHtml & "            NRR='r' + document.getElementById('XFKR').cells(j).id;"
				sHtml = sHtml & "            document.getElementById(NRR).style.width = document.getElementById('XFKR').cells(j).clientWidth;"
				sHtml = sHtml & "            count++;"
				sHtml = sHtml & "        }"
				sHtml = sHtml & " }"
				sHtml = sHtml & " </script>" & Chr(13)

				FBBarr = FBBarr & sHtml

			End If

		End If



		Encabezados = GenRow("XFKR", "XFK", "", "", "", "", "", Celda, "si")

		Exit Function

		EncabezadosErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		Encabezados = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function EncabezadosReportes( vParametros)

		'On Error GoTo EncabezadosReportesErr

		Dim Celda, Fila, i, j

		j = 0
		For i = 0 To UBound(vParametros)
			If j = 0 Or j = 2 Then
				Celda = Celda & GenCelda("", "tx", "10", "10%", "", "", "", vParametros(i), "si")
				j = j + 1
			ElseIf j = 1 Or j = 3 Then
				Celda = Celda & GenCelda("", "cc", "10", "40%", "", "", "", vParametros(i), "si")
				j = j + 1
			End If
			If j = 4 Then
				Fila = Fila & GenRow("", "", "", "", "", "", "", Celda, "si")
				Celda = ""
				j = 0
			End If
		Next
		If j = 2 Then Fila = Fila & GenRow("", "", "", "", "", "", "", Celda, "si")

		EncabezadosReportes = GenTabla("", "", "", "100%", "", "0", "0", "0", "0", "", "", Fila, "si")

		Exit Function

		EncabezadosReportesErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		EncabezadosReportes = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function Armarvector(cadena, caracter)

		'On Error GoTo ArmarvectorErr

		Dim i, h, car

		i = 1
		h = 0
		ReDim Vector(h)
		Do Until i > Len(cadena)
			car = ""
			Do Until Mid(cadena, i, 1) = caracter
				car = car & Mid(cadena, i, 1)
				i = i + 1
			Loop
			ReDim Preserve Vector(h)
			Vector(h) = car
			i = i + 1
			h = h + 1
		Loop
		Armarvector = Vector

		Exit Function

		ArmarvectorErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		Armarvector = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function FormStyler()

		FormStyler = "<meta HTTP-EQUIV=" & Chr(34) & "expires" & Chr(34) & " CONTENT=" & Chr(34) & "0" & Chr(34) & " >" & FormStylerCon()

	End Function

	Public Function FormStylerCon()

		'On Error GoTo FormStylerConErr

		Dim s

		If wstyles = "" Then

			s = "<META HTTP-EQUIV=" & Chr(34) & "Content-Type" & Chr(34) & " CONTENT=" & Chr(34) & "text/html; charset=windows-1252" & Chr(34) & "><style>"
			'124    If Styler = 0 Or Styler = 1 Then
			s = s & "td.mt{font-family:tahoma;color:#000077;font-size:" & wfont & "pt;text-transform:uppercase;font-weight:600;border-style:outset;border-width:1pt;background-color:#eeeeee;}"
			s = s & "td.ut{font-family:tahoma;color:#000077;font-size:8pt;font-weight:500;}"
			s = s & "td.tt{font-family:tahoma;color:#dddddd;font-size:" & wfont & "pt;text-transform:uppercase;font-weight:600;background-color:#336699;}"
			s = s & "input.dt{font-family:tahoma;color:#000077;font-size:" & wfont & "pt;background-color:#f7eedd;}"
			s = s & "input.dtn{font-family:tahoma;color:#000077;font-size:" & wfont & "pt;background-color:#f7eedd;text-align=right;}"
			s = s & "input.bt{font-family:tahoma;color:#eeeeee;font-size:" & wfont & "pt;background-color:#336699;}"
			s = s & "select.at{font-family:tahoma;color:#000077;font-size:" & wfont & "pt;background-color:#f7eedd;}"
			s = s & "font.at{font-family:tahoma;color:#000077;font-size:" & wfont & "pt;}"
			'125 End If
			'125    If Styler = 0 Then
			s = s & "a{font-size:" & wfont & "pt;font-family:tahoma;color:#ffffff;text-transform:capitalize;font-weight:500;text-decoration:none;}"
			s = s & "a.link{color:#ffffff;text-decoration:none;}"
			s = s & "a.nn{font-size:" & wfont & "pt;font-family:tahoma;color:#ffffff;text-transform:capitalize;font-weight:500;text-decoration:none;}"
			s = s & "a.tt{font-size:" & wfont & "pt;font-family:tahoma;color:#ffffff;text-transform:uppercase;font-weight:800;text-decoration:none;}"
			s = s & "a.nt{font-size:" & wfont & "pt;font-family:tahoma;color:#000055;font-weight:500;text-decoration:none;}"
			s = s & "input{font-family:tahoma;color:#000077;font-size:" & wfont & "pt;background-color:#f7eedd;}"
			s = s & "input.at{font-family:tahoma;color:#000055;font-size:" & wfont & "pt;background-color:#ffffff;}"
			s = s & "input.dtb{font-family:tahoma;color:#000077;font-size:" & wfont & "pt;background-color:#f7eedd;font-weight:bold;}"
			s = s & "input.rd{font-family:tahoma;color:#000077;font-size:" & wfont & "pt;background-color:#cccccc;}"
			s = s & "input.it{font-family:tahoma;color:#000077;font-size:" & wfont & "pt;background-color:#cccccc;}"
			s = s & "select{font-family:tahoma;color:#000077;font-size:" & wfont & "pt;background-color:#f7eedd;}"
			s = s & "span.at{font-family:tahoma;color:#000077;font-size:" & wfont & "pt;background-color:#f7eedd;}"
			s = s & "span.mt{font-family:tahoma;color:#000077;font-size:" & wfont & "pt;text-transform:uppercase;font-weight:600;border-style:outset;border-width:1pt;background-color:#eeeeee;width:80%;}"
			s = s & "option{font-family:tahoma;color:#000077;font-size:" & wfont & "pt;background-color:#f7eedd;}"
			s = s & "th{font-family:tahoma;color:#dddddd;font-size:" & wfont & "pt;text-transform:uppercase;font-weight:600;background-color:#336699;}"
			s = s & "tr.st{background-color:#336699;}"
			s = s & "tr.cc{background-color:#eeeeee;border-style:ridge;border-width:1pt;}td.aa{text-transform: capitalize;}"
			s = s & "td{font-family:tahoma;font-size:" & wfont & "pt;}"
			s = s & "td.at{font-family:tahoma;color:#eeeeee;font-size:" & wfont & "pt;font-weight:500;}"
			s = s & "td.cc{font-family:tahoma;color:#000077;font-size:" & wfont & "pt;background-color:#f7eedd;border-style:ridge;border-width:1pt;}"
			s = s & "td.ee{font-size:" & wfont & "pt;font-family:tahoma;border-width:1pt;border-color:#4466bb;color:#770000;font-weight:500;}"
			s = s & "td.lt{font-family:verdana;color:#550000;font-size:8pt;font-weight:600;text-indent:10;}"
			s = s & "td.dt{font-family:tahoma;color:#000077;font-size:" & wfont & "pt;}"
			s = s & "td.qt{font-family:tahoma;color:#550000;font-size:8pt;font-weight:600;}"
			s = s & "td.rr{font-size:" & wfont & "pt;font-family:tahoma;border-width:1pt;border-color:#4466bb;color:#ffffff;font-weight:500;}"
			s = s & "td.st{font-family:verdana;color:#dddddd;font-size:" & wfont & "pt;text-transform:uppercase;font-weight:600;background-color:#336699;}"
			s = s & "td.tu{font-family:verdana;color:#336699;font-size:10px;font-weight:bold;}"
			s = s & "td.tx{font-family:arial narrow;color:#000077;font-size:" & wfont & "pt;text-transform:uppercase;border-style:outset;border-width:1pt;background-color:#aabbdd;}"
			s = s & "td.ut2{font-family:tahoma;color:#000077;font-size:" & wfont + 1 & "pt;font-weight:none;}"
			s = s & "td.vt{font-family:tahoma;color:#000055;font-size:8pt;font-weight:600;}"
			s = s & "td.wt{font-family:tahoma;color:#550000;font-size:" & wfont & "pt;font-weight:600;}"
			s = s & "td.xt{font-family:arial narrow;color:#dddddd;font-size:" & wfont + 1 & "pt;text-transform:uppercase;font-weight:600;background-color:#336699;}"
			s = s & "td.za{font-family:tahoma;color:#ff0000;font-size:" & wfont + 3 & "pt;text-transform:uppercase;}"
			s = s & "td.zb{font-family:tahoma;color:#0000ff;font-size:" & wfont + 3 & "pt;text-transform:uppercase;}"
			s = s & "td.zc{font-family:tahoma;color:#000000;font-size:" & wfont & "pt;text-transform:uppercase;}"
			s = s & "td.zd{font-family:tahoma;color:#000099;font-size:" & wfont & "pt;text-transform:uppercase;}"
			s = s & "td.ze{font-family:tahoma;color:#000033;font-size:" & wfont & "pt;text-transform:uppercase;}"
			s = s & "div{page-break-after:always;}"
			s = s & "textarea.dt{font-family:tahoma;color:#000077;font-size:" & wfont & "pt;background-color:#f7eedd;}"
			'129    End If
			s = s & "</style>"

		Else
			s = wstyles
		End If

		FormStylerCon = s

		Exit Function

		FormStylerConErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		FormStylerCon = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function CheckUsuario()

		'On Error GoTo CheckUsuarioErr

		Dim s, t1, sPath, Fecha

		If ValidarUsuario = "OK" Then
			If Session("Usuario") = "" Then
				'            s = "<html><Title>Cor E-Solution Suite</title>" & FormStyler()
				'            s = s & "<body bgcolor=" & Session("CMLOGCHK") & ">"

				s = MenuHeader("", "")
				s = s & GenTabla("headm1", "", "", "50%", "", 0, 0, 0, 0, "", "", "", "no")
				t1 = GenImage("", "", "22", "22", "Imagenes/Corlogox1.jpg", "0", "")
				t1 = GenCelda("", "", "left", "4%", "", "", "", t1, "si")
				t1 = t1 & GenCelda("", "ut", "left", "55%", "", "", "", "COR ESolution Suite", "si")
				s = s & GenRow("", "", "top", "", "", "", "", t1, "si")
				s = s & GenFTabla()
				s = s & "<br>"
				s = s & GenTabla("headm2", "", "", "100%", "", 0, 0, 0, 0, "", "", "", "no")
				Fecha = Fechador()
				t1 = GenCelda("", "lt", "left", "60%", "", "", "", "Logon Check", "si")
				t1 = t1 & GenCelda("", "ut", "right", "", "", "", "", Fecha, "si")
				s = s & GenRow("", "", "top", "", "", "", "", t1, "si")
				s = s & GenFTabla()
				s = s & "<hr color=#000088 size=1 width=100% align=left>"
				s = s & GenTabla("headm3", "", "", "100%", "", 0, 0, 0, 0, "", "", "", "no")
				t1 = "<font size=3 face=tahoma color=#993333>La Aplicacion no esta disponible sin Logon,<br>"
				t1 = t1 & " el login fue invalido o expiro por Tiempo de espera excedido!! </font> "
				t1 = GenCelda("", "vt", "left", "4%", "", "", "", t1, "si")
				s = s & GenRow("", "", "top", "", "", "", "", t1, "si")

				'         sPath = "../CoreSS20/"
				'       t1 = GenForm("logoff", "", sPath & "inicio.asp?Opcion=I", "Master", "post", "", GenerarInput("blo1", "Logon", "submit", "", "10", "bt", ""), "si")
				'         t1 = GenCelda("", "", "center", "4%", "", "", "", "<br><br>" & t1, "si")
				'        s = s & GenRow("", "", "top", "", "", "", "", t1, "si")

				s = s & GenFTabla()
				s = s & "</body></html>"
			Else
				s = "Y"
			End If
		Else
			s = "<html><Title>Cor E-Solution Suite</title>" & FormStyler()
			s = s & "<body bgcolor=" & Session("CMLOGCHK") & " >"
			s = s & GenTabla("headm1", "", "", "50%", "", 0, 0, 0, 0, "", "", "", "no")
			t1 = GenImage("", "", "22", "22", "Imagenes/Corlogox1.jpg", "0", "")
			t1 = GenCelda("", "", "left", "4%", "", "", "", t1, "si")
			t1 = t1 & GenCelda("", "ut", "left", "55%", "", "", "", "COR EManager", "si")
			s = s & GenRow("", "", "top", "", "", "", "", t1, "si")
			s = s & GenFTabla()
			s = s & "<br>"
			s = s & GenTabla("headm2", "", "", "100%", "", 0, 0, 0, 0, "", "", "", "no")
			Fecha = Fechador()
			t1 = GenCelda("", "lt", "left", "60%", "", "", "", "Logon Check", "si")
			t1 = t1 & GenCelda("", "ut", "right", "", "", "", "", Fecha, "si")
			s = s & GenRow("", "", "top", "", "", "", "", t1, "si")
			s = s & GenFTabla()
			s = s & "<hr color=#000088 size=1 width=100% align=left>"
			s = s & GenTabla("headm3", "", "", "100%", "", 0, 0, 0, 0, "", "", "", "no")
			t1 = "La Aplicacion ha expirado o no esta procesando en el servidor registrado,"
			t1 = t1 & " no podra continuar. Contactese con el proveedor"
			t1 = GenCelda("", "cc", "left", "4%", "", "", "", t1, "si")
			s = s & GenRow("", "", "top", "", "", "", "", t1, "si")
			s = s & GenFTabla()
			s = s & "</body></html>"
		End If

		CheckUsuario = s

		Exit Function

		CheckUsuarioErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		CheckUsuario = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function FechadorX(Fecha)

		'On Error GoTo FechadorXErr

		Dim ndia, nMes, nAno

		ndia = Mid(Fecha, 4, 2)
		nMes = Left(Fecha, 2)
		nAno = Right(Fecha, 2)

		FechadorX = ndia & "-" & nMes & "-" & nAno

		Exit Function

		FechadorXErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		FechadorX = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function Fechador()

		'On Error GoTo FechadorErr

		Dim ndia, nMes, fec

		fec = Date
		ndia = cdia(Weekday(fec))
		nMes = cmes(Month(fec))

		Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)

		Exit Function

		FechadorErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		Fechador = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function cmes(mes)

		'On Error GoTo cmesErr

		Dim xmes

		xmes = Array("nulo", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")

		cmes = xmes(mes)

		Exit Function

		cmesErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		cmes = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function cdia(dia)

		'On Error GoTo cdiaErr

		Dim xdia

		xdia = Array("nulo", "Domingo", "Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado")

		cdia = xdia(dia)

		Exit Function

		cdiaErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		cdia = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Sub ErrorDB(cn, volver)

		Dim Errores
		Dim sDescripcion, nNumero

		Set Errores = cn.Errors.Item(0)
		sDescripcion = ConvertirCar(Errores.Description, " ", "_")
		nNumero = Errores.Number
		Set Errores = Nothing

		cn.Errors.Clear
		Response.redirect "ErrordbEx.asp?volver=" & volver & "&Descripcion=" & sDescripcion & "&Numero=" & nNumero

	End Sub

	Public Function Editar(sASP)

		'On Error GoTo EditarErr

		Dim s

		s = GenCelda("", "", "center", "33%", "", "", "", GenerarInput("bed1", "Editar", "button", "", "10", "bt", "AbrirVentana()") & "<br><br>", "si")
		s = GenForm("", "", "", "", "post", "", s, "si")

		s = s & "<script language=JavaScript>" & Chr(13)
		s = s & "function AbrirVentana() {" & Chr(13)
		s = s & " myWin= open(" & Chr(34) & sASP & Chr(34)
		s = s & "," & Chr(34) & "edi" & Chr(34) & "," & Chr(34)
		s = s & "menubar=no,toolbar=no,height=570,width=790" & Chr(34)
		s = s & " ); " & Chr(13) & "window.close(); " & Chr(13)
		s = s & " } " & Chr(13) & "</script>" & Chr(13)

		Editar = s

		Exit Function

		EditarErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		Editar = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function Eliminar(sASP)

		'On Error GoTo EliminarErr

		Dim s

		s = GenCelda("", "", "center", "33%", "", "", "", GenerarInput("bel1", "Eliminar", "button", "", "10", "bt", "AbrirEliminar()") & "<br><br>", "si")
		s = GenForm("", "", "", "", "post", "", s, "si")

		s = s & "<script language=JavaScript>" & Chr(13)
		s = s & "function AbrirEliminar() {" & Chr(13)
		s = s & " myWin= open(" & Chr(34) & sASP & Chr(34)
		s = s & "," & Chr(34) & "edi" & Chr(34) & "," & Chr(34)
		s = s & "menubar=no,toolbar=no" & Chr(34)
		s = s & " ); " & Chr(13) & "window.close(); " & Chr(13)
		s = s & " } " & Chr(13) & "</script>" & Chr(13)

		Eliminar = s

		Exit Function

		EliminarErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		Eliminar = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function Copiar(sASP)

		'On Error GoTo CopiarErr

		Dim s

		s = GenCelda("", "", "center", "33%", "", "", "", GenerarInput("bco1", "Copiar", "button", "", "10", "bt", "AbrirCopiar()") & "<br><br>", "si")
		s = GenForm("", "", "", "", "post", "", s, "si")

		s = s & "<script language=JavaScript>" & Chr(13)
		s = s & "function AbrirCopiar() {" & Chr(13)
		s = s & " myWin= open(" & Chr(34) & sASP & Chr(34)
		s = s & "," & Chr(34) & "edi" & Chr(34) & "," & Chr(34)
		s = s & "menubar=no,toolbar=no,height=120,width=400" & Chr(34)
		s = s & " ); " & Chr(13) & "window.close(); " & Chr(13)
		s = s & " } " & Chr(13) & "</script>" & Chr(13)

		Copiar = s

		Exit Function

		CopiarErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		Copiar = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function ArmarRuta()

		'On Error GoTo ArmarRutaErr

		Dim str, NameServer, Pos

		NameServer = "//" & Request.servervariables("SERVER_NAME") & "/"
		str = NameServer & ConvertirCar(Server.mappath(Request.servervariables("PATH_INFO")), "\", "/")
		str = ConvertirCar(str, ":", "")
		Pos = InStrRev(str, "/", -1)

		ArmarRuta = Mid(str, 1, Pos)

		Exit Function

		ArmarRutaErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		ArmarRuta = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function VueltaAsp(mdir, m1, m2)

		'On Error GoTo VueltaAspErr

		Dim nn

		nn = StrReverse(Request.servervariables("SCRIPT_NAME"))
		nn = StrReverse(Left(nn, InStr(nn, "/") - 1))
		nn = mdir & nn & "?cantregmostrar=" & CStr(m1) & "_CantRegMover=" & CStr(m2)

		VueltaAsp = nn

		Exit Function

		VueltaAspErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		VueltaAsp = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function VueltaAspEx(mdir, m1, m2)

		'On Error GoTo VueltaAspExErr

		Dim nn

		nn = StrReverse(Request.servervariables("SCRIPT_NAME"))
		nn = StrReverse(Left(nn, InStr(nn, "/") - 1))
		nn = mdir & nn & "?CantRegAMostrar=" & CStr(m1) & "_CantRegAMover=" & CStr(m2)

		VueltaAspEx = nn

		Exit Function

		VueltaAspExErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		VueltaAspEx = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function VueltaAspExC(mdirx, m1x, m2x)

		'On Error GoTo VueltaAspExCErr

		Dim nnx

		nnx = StrReverse(Request.servervariables("SCRIPT_NAME"))
		nnx = StrReverse(Left(nnx, InStr(nnx, "/") - 1))
		nnx = mdirx & nnx & "?CantRegAMover=" & CStr(m1x) & "_CantRegAMover2=" & CStr(m2x)

		VueltaAspExC = nnx

		Exit Function

		VueltaAspExCErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		VueltaAspExC = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function VueltaAspEx2(mdir, m1, m2)

		'On Error GoTo VueltaAspEx2Err

		Dim nn

		nn = StrReverse(Request.servervariables("SCRIPT_NAME"))
		nn = StrReverse(Left(nn, InStr(nn, "/") - 1))
		nn = mdir & nn & "?CantRegAMostrar2=" & CStr(m1) & "_CantRegAMover2=" & CStr(m2)

		VueltaAspEx= nn

		Exit Function

		VueltaAspEx2Err:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		VueltaAspEx= oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function VueltaAsp2(mdir)

		'On Error GoTo VueltaAsp2Err

		Dim nn

		nn = StrReverse(Request.servervariables("SCRIPT_NAME"))
		nn = StrReverse(Left(nn, InStr(nn, "/") - 1))
		nn = mdir & nn
		'response.write nn
		VueltaAsp2 = CStr(nn)

		Exit Function

		VueltaAsp2Err:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		VueltaAsp= oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function xVueltaAsp(mdir, m1, m2, m3)

		'On Error GoTo xVueltaAspErr

		Dim nn

		nn = StrReverse(Request.servervariables("SCRIPT_NAME"))
		nn = StrReverse(Left(nn, InStr(nn, "/") - 1))
		nn = mdir & nn
		If m1 <> "" Then nn = nn & "?CantRegMover=" & CStr(m1)
		If m<> "" Then nn = nn & "_CantRegMover2=" & CStr(m2)
		If m3 <> "" Then nn = nn & "_CantRegMostrar=" & CStr(m3)

		xVueltaAsp = nn

		Exit Function

		xVueltaAspErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		xVueltaAsp = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function botoneraESSX(Bclase, Brs, BMover, BMostrar, BMover2, BMostrar2, Bposi, BcodeT, BTipo, BActualizar, bAgregar, BVolver, BRetornar)

		'On Error GoTo botoneraESSXErr

		Dim l, CTabla, Bretro, Bavan, x, x1, x2, botonbuscar, xbInd

		l = ""
		CTabla = "no"
		Bretro = ""
		Bavan = ""
		xbInd = 0

		l = StrReverse(Request.servervariables("SCRIPT_NAME"))
		l = StrReverse(Left(l, InStr(l, "/") - 1))

		Select Case Bclase

			Case "0"

			CTabla = "si"

			Bretro = "NO"
			Bavan = "NO"



			Case "1"

			CTabla = "no"

			If BMover <> 0 Then
				If Brs.EOF Then
					Bretro = l & "?cantregmostrar=" & CStr(BMostrar) & "&CantRegMover=" & CStr((BMover - BMostrar))
				Else
					Bretro = l & "?cantregmostrar=" & CStr(BMostrar) & "&CantRegMover=" & CStr((BMover - Bposi))
				End If
			End If


			If Not Brs.EOF Then
				Bavan = l & "?cantregmostrar=" & CStr(BMostrar) & "&CantRegMover=" & CStr(BMover + Bposi)
			End If


			Case "2"

			CTabla = "si"


			If BMover <> 0 Then
				If Brs.EOF Then
					Bretro = l & "?cantregmostrar=" & CStr(BMostrar) & "&CantRegMover=" & CStr((BMover - BMostrar))
				Else
					Bretro = l & "?cantregmostrar=" & CStr(BMostrar) & "&CantRegMover=" & CStr((BMover - Bposi))
				End If
			End If


			If Not Brs.EOF Then
				Bavan = l & "?cantregmostrar=" & CStr(BMostrar) & "&CantRegMover=" & CStr(BMover + Bposi)
			End If



			Case "3"

			CTabla = "si"


			If BMover <> 0 Then
				If Brs.EOF Then
					Bretro = l & "?cantregamostrar=" & CStr(BMostrar) & "&CantRegaMover=" & CStr((BMover - BMostrar))
				Else
					Bretro = l & "?cantregamostrar=" & CStr(BMostrar) & "&CantRegaMover=" & CStr((BMover - BMostrar))
				End If
				If BRetornar <> "" Then Bretro = Bretro & BRetornar
			End If


			If Not Brs.EOF Then
				Bavan = l & "?cantregamostrar=" & CStr(BMostrar) & "&CantRegaMover=" & CStr(BMover + BMostrar)
				If BRetornar <> "" Then Bavan = Bavan & BRetornar
			End If

			Case "4"

			CTabla = "si"


			If BMover <> 0 Then
				If Brs.EOF Then
					Bretro = l & "?cantregamostrar=" & CStr(BMostrar) & "&CantRegaMover=" & CStr((BMover - BMostrar))
				Else
					Bretro = l & "?cantregamostrar=" & CStr(BMostrar) & "&CantRegaMover=" & CStr((BMover - BMostrar))
				End If
				If BRetornar <> "" Then Bretro = Bretro & BRetornar
			End If


			If Not Brs.EOF Then
				Bavan = l & "?cantregamostrar=" & CStr(BMostrar) & "&CantRegaMover=" & CStr(BMover + BMostrar)
				If BRetornar <> "" Then Bavan = Bavan & BRetornar
			End If


			Case "5"

			CTabla = "si"


			If BMover<> 0 Then
				If Brs.EOF Then
					Bretro = l & "?cantregamostrar2=" & CStr(BMostrar2) & "&CantRegaMover=" & CStr(BMover) & "&CantRegaMover2=" & CStr((BMover- BMostrar2))
				Else
					Bretro = l & "?cantregamostrar2=" & CStr(BMostrar2) & "&CantRegaMover=" & CStr(BMover) & "&CantRegaMover2=" & CStr((BMover- BMostrar2))
				End If
				If BRetornar <> "" Then Bretro = Bretro & BRetornar
			End If


			If Not Brs.EOF Then
				Bavan = l & "?cantregamostrar2=" & CStr(BMostrar2) & "&CantRegaMover=" & CStr(BMover) & "&CantRegaMover2=" & CStr((BMover+ BMostrar2))
				If BRetornar <> "" Then Bavan = Bavan & BRetornar
			End If


			Case "6"

			CTabla = "no"


			If BMover <> 0 Then
				If Brs.EOF Then
					Bretro = l & "?cantregamostrar=" & CStr(BMostrar) & "&CantRegaMover=" & CStr((BMover - BMostrar))
				Else
					Bretro = l & "?cantregamostrar=" & CStr(BMostrar) & "&CantRegaMover=" & CStr((BMover - BMostrar))
				End If
				Bretro = Bretro & "&Codobjeto=" & BcodeT & "&Tipobjeto=" & BTipo

			End If


			If Not Brs.EOF Then
				Bavan = l & "?cantregamostrar=" & CStr(BMostrar) & "&CantRegaMover=" & CStr(BMover + BMostrar)
				Bavan = Bavan & "&Codobjeto=" & BcodeT & "&Tipobjeto=" & BTipo

			End If


			Case "7"

			CTabla = "no"


			If BMover<> 0 Then
				If Brs.EOF Then
					Bretro = l & "?cantregamostrar2=" & CStr(BMostrar2) & "&CantRegaMover2=" & CStr((BMover- BMostrar2))
				Else
					Bretro = l & "?cantregamostrar2=" & CStr(BMostrar2) & "&CantRegaMover2=" & CStr((BMover- BMostrar2))
				End If
				Bretro = Bretro & "&Codobjeto=" & BcodeT & "&Tipobjeto=" & BTipo

			End If


			If Not Brs.EOF Then
				Bavan = l & "?cantregamostrar2=" & CStr(BMostrar2) & "&CantRegaMover2=" & CStr(BMover+ BMostrar2)
				Bavan = Bavan & "&Codobjeto=" & BcodeT & "&Tipobjeto=" & BTipo

			End If


			Case "8"

			CTabla = "si"


			If BMover <> 0 Then
				If Brs.EOF Then
					Bretro = l & "?cantregamostrar=" & CStr(BMostrar) & "&CantRegaMover=" & CStr((BMover - BMostrar))
				Else
					Bretro = l & "?cantregamostrar=" & CStr(BMostrar) & "&CantRegaMover=" & CStr((BMover - BMostrar))
				End If
				Bretro = Bretro & "&Codcomp=" & BcodeT & "&Tipocomp=" & BTipo

			End If


			If Not Brs.EOF Then
				Bavan = l & "?cantregamostrar=" & CStr(BMostrar) & "&CantRegaMover=" & CStr(BMover + BMostrar)
				Bavan = Bavan & "&Codcomp=" & BcodeT & "&Tipocomp=" & BTipo

			End If


			Case "9"

			CTabla = "si"


			If BMover<> 0 Then
				If Brs.EOF Then
					Bretro = l & "?cantregamostrar2=" & CStr(BMostrar2) & "&CantRegaMover2=" & CStr((BMover- BMostrar2))
				Else
					Bretro = l & "?cantregamostrar2=" & CStr(BMostrar2) & "&CantRegaMover2=" & CStr((BMover- BMostrar2))
				End If
				Bretro = Bretro & "&Codcomp=" & BcodeT & "&Tipocomp=" & BTipo

			End If


			If Not Brs.EOF Then
				Bavan = l & "?cantregamostrar2=" & CStr(BMostrar2) & "&CantRegaMover2=" & CStr(BMover+ BMostrar2)
				Bavan = Bavan & "&Codcomp=" & BcodeT & "&Tipocomp=" & BTipo

			End If


			Case "a"

			CTabla = "si"


			If BMover <> 0 Then
				If Brs.EOF Then
					Bretro = l & "?cantregamostrar=" & CStr(BMostrar) & "&CantRegaMover=" & CStr((BMover - BMostrar))
				Else
					Bretro = l & "?cantregamostrar=" & CStr(BMostrar) & "&CantRegaMover=" & CStr((BMover - BMostrar))
				End If
				Bretro = Bretro & "&Codobjeto=" & BcodeT & "&Tipobjeto=" & BTipo

			End If


			If Not Brs.EOF Then
				Bavan = l & "?cantregamostrar=" & CStr(BMostrar) & "&CantRegaMover=" & CStr(BMover + BMostrar)
				Bavan = Bavan & "&Codobjeto=" & BcodeT & "&Tipobjeto=" & BTipo

			End If


			Case "b"

			CTabla = "no"


			If BMover<> 0 Then
				If Brs.EOF Then
					Bretro = l & "?cantregamostrar2=" & CStr(BMostrar2) & "&CantRegaMover2=" & CStr((BMover- BMostrar2))
				Else
					Bretro = l & "?cantregamostrar2=" & CStr(BMostrar2) & "&CantRegaMover2=" & CStr((BMover- BMostrar2))
				End If
				Bretro = Bretro & "&Codobjeto=" & BcodeT & "&Tipobjeto=" & BTipo

			End If


			If Not Brs.EOF Then
				Bavan = l & "?cantregamostrar2=" & CStr(BMostrar2) & "&CantRegaMover2=" & CStr(BMover+ BMostrar2)
				Bavan = Bavan & "&Codobjeto=" & BcodeT & "&Tipobjeto=" & BTipo

			End If

			Case "c"

			CTabla = "si"


			If BMover <> 0 Then
				If Brs.EOF Then
					Bretro = l & "?cantregmostrar=" & CStr(BMostrar) & "&CantRegMover=" & CStr((BMover - BMostrar))
				Else
					Bretro = l & "?cantregmostrar=" & CStr(BMostrar) & "&CantRegMover=" & CStr((BMover - Bposi))
				End If
				Bretro = Bretro & BcodeT
			End If


			If Not Brs.EOF Then
				Bavan = l & "?cantregmostrar=" & CStr(BMostrar) & "&CantRegMover=" & CStr(BMover + Bposi)
				Bavan = Bavan & BcodeT
			End If


		End Select

		x = ""

		x1 = GenCelda("", "", "", "", "", "", "", "", "no")
		x= GenFCelda()

		If CTabla <> "no" Then
			x = x & FBbtnt
			x = x & GenTabla("btnas", "tbb", "", "2%", "", 0, 0, 1, 1, "", "", "", "no")
			x = x & GenRow("", "", "", "", "", "", "", "", "no") & Chr(13)

		End If


		' ModBot ''''' modal botones
		'''''''''''''' 1- VML
		'''''''''''''' 2- Botones
		'''''''''''''' 3- TD
		'''''''''''''' 4- Texto

		FBmodo = ""

		botonbuscar = ""

		If Bretro <> "NO" Then

			xbInd = xbInd + 1

			If Bretro <> "" Then

				FBmodo = GenBoton(1, "BretroBot", "BT", "S", "", "S", "A", "N", "BretroFrm", Bretro)

				x = x & FBmodo

				botonbuscar = "si"

			Else

				FBmodo = GenBoton(11, "BretroBot", "BT", "N", "", "S", "A", "N", "", "")

				x = x & FBmodo

				botonbuscar = "si"

			End If

		End If

		FBmodo = ""

		If Bavan <> "NO" Then

			xbInd = xbInd + 1

			If Bavan <> "" Then

				FBmodo = GenBoton(2, "BavanBot", "BT", "S", "", "S", "A", "N", "BavanFrm", Bavan)

				x = x & FBmodo

				botonbuscar = "si"

			Else

				FBmodo = GenBoton(12, "BavanBot", "BT", "N", "", "S", "A", "N", "", "")

				x = x & FBmodo



				botonbuscar = "si"

			End If

		End If


		If BActualizar <> "no" And BActualizar <> "" Then
			x = x & BActualizar & Chr(13)
			xbInd = xbInd + 1
		End If

		If bAgregar <> "no" And bAgregar <> "" Then
			x = x & bAgregar & Chr(13)
			xbInd = xbInd + 1
		End If

		If BVolver <> "no" And BVolver <> "" Then
			x = x & BVolver & Chr(13)
			xbInd = xbInd + 1
		End If

		If Bclase < 3 Or Bclase > 5 Then
			If BRetornar <> "no" And BRetornar <> "" Then
				x = x & BRetornar & Chr(13)
				xbInd = xbInd + 1
			End If
		End If

		If xbInd > 0 And ModBot = "1" Then

			FBmodo = BotonHelp("S")
			x = x & FBmodo

		End If



		If CTabla <> "no" Then x = x & GenFRow() & GenFTabla() & FBbtnf & Chr(13)

		botoneraESSX = x & FBBarr & Chr(13)

		Exit Function

		botoneraESSXErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		botoneraESSX = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function


	Public Function BotonHelp(sOpcion)

		'On Error GoTo BotonHelpErr

		Dim i, j, YY, Buscar, sx, xsfunc, sDiver, sinput2, sImagenx, xGenCont, xGenConti


		FBmodo = ""


		FBmodo = GenBoton(0, "HelpBot", "BT", "N", " onclick='ArmarHelp();' ", "N", "S", "N", "HelpFrm", "")


		xsfunc = "" & Chr(13)


		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

		sDiver = "<div id='DBHL' class='DINDM' style='display:none;"
		sDiver = sDiver & "z-index:80;margin-top:1;margin-left:1;position:absolute;"
		sDiver = sDiver & "left:300px;top:200px;'"
		sDiver = sDiver & " >"


		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

		sDiver = sDiver & "<table id='DBHL1' style='z-index:81;' class='TWIND' width=cellspacing=0 cellpadding=>"
		sDiver = sDiver & "<tr>"
		sDiver = sDiver & "<td class=mbt width=99%  >&nbsp;"

		sDiver = sDiver & UCase(FBotTab(0, 6))

		sDiver = sDiver & "</td>"
		sDiver = sDiver & "<td CLASS=mbt width=1% onclick=""document.getElementById('DBHL').style.display='none';"" TITLE='OCULTAR' >"
		'If m_Browser = "IE" Then
		sDiver = sDiver & "<font  onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & Session("CMFMNBOT") & "';"" face=CorFlaticon size=>r</font>"
		'Else
		'  sDiver = sDiver & "<font  onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & Session("CMFMNBOT") & "';"" face=Wingdings size=>&#8999;</font>"
		'End If

		sDiver = sDiver & "</td>"
		sDiver = sDiver & "</tr>"
		sDiver = sDiver & "</table>"

		sDiver = sDiver & "<table id=""DBHL2"" style='z-index:81;' class='TWINE2' width=cellspacing=4 cellpadding="
		sDiver = sDiver & " >"
		sDiver = sDiver & "</table>"

		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


		sDiver = sDiver & "</div>"


		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


		xsfunc = xsfunc & sDiver & Chr(13)
		xsfunc = xsfunc & "<script language=javascript>" & Chr(13)
		xsfunc = xsfunc & " var xktab = 0;" & Chr(13)
		xsfunc = xsfunc & "function ArmarHelp() {" & Chr(13)
		xsfunc = xsfunc & "if (xktab==0){CargarHelp();}" & Chr(13)
		xsfunc = xsfunc & "if (document.getElementById('DBSC')){document.getElementById('DBSC').style.display='none';}" & Chr(13)
		xsfunc = xsfunc & "document.getElementById('DBHL').style.display='block';" & Chr(13)
		xsfunc = xsfunc & "}" & Chr(13)
		xsfunc = xsfunc & "function CargarHelp() {" & Chr(13)

		xsfunc = xsfunc & "    var NRR;"
		xsfunc = xsfunc & "    var NRO;"
		xsfunc = xsfunc & "    var NRP;"
		xsfunc = xsfunc & "    var NRQ;"
		xsfunc = xsfunc & "        for (j=0; j < document.getElementById('btnas').rows(0).cells.length; j++) {"
		xsfunc = xsfunc & "            NRQ = document.getElementById('DBHL2').insertRow();"
		xsfunc = xsfunc & "            NRR = NRQ.insertCell();"
		xsfunc = xsfunc & "            NRP = NRQ.insertCell();"
		xsfunc = xsfunc & "            NRR.className  = 'vt';"
		xsfunc = xsfunc & "            NRP.className  = 'vt';"
		xsfunc = xsfunc & "            NRR.width  = '60px';"
		xsfunc = xsfunc & "            NRR.innerHTML = document.getElementById('btnas').rows(0).cells(j).innerHTML;"
		xsfunc = xsfunc & "            NRP.innerHTML = document.getElementById('btnas').rows(0).cells(j).title;"
		xsfunc = xsfunc & "         }" & Chr(13)
		xsfunc = xsfunc & " xktab=1;" & Chr(13)
		xsfunc = xsfunc & " }" & Chr(13)
		xsfunc = xsfunc & "</script>" & Chr(13)

		FBBarr = FBBarr & xsfunc


		BotonHelp = FBmodo

		Exit Function

		BotonHelpErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		BotonHelp = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function



	Public Function ObtenerNom(Vector, VectorD, rs)

		Dim i

		ObtenerNom = ""

		If IsArray(Vector) And IsArray(VectorD) Then
			For i = 0 To UBound(Vector)
				If rs = Vector(i) Then
					ObtenerNom = VectorD(i)
					Exit For
				End If
			Next
		End If

	End Function

	Public Function ComboBuscar(strsql1, strsql2, cn, CantRegMostrarx, CantRegMoverx, vuelta)

		'On Error GoTo ComboBuscarErr

		Dim i, j, YY, Buscar, sx, x1, x2, sAccion, sDiver, sinput2, sImagenx, rx, xGenCont, xGenConti
		Dim Perfiles(), Perfilesd()

		sx = "?"
		i = InStr(vuelta, "?")
		If i > 0 Then sx = "&"

		i = 0
		j = 0



		FBmodo = ""


		FBmodo = GenBoton(4, "BusBot", "BT", "N", " onclick='AbreBusc();' ", "S", "A", "N", "BusFrm", "")


		Set rx = CrearConsultaEx(strsql1, cn, 1, "")
		Do Until rx.EOF
			For YY = 1 To CantRegMostrarx
				If rx.EOF Then Exit For
				ReDim Preserve Perfiles(j)
				ReDim Preserve Perfilesd(j)
				Perfiles(j) = i
				Perfilesd(j) = rx(0)
				rx.movenext
				j = j + 1
			Next
			i = i + CantRegMostrarx
		Loop
		Call CerrarConsulta(rx)

		Set rx = CrearConsultaEx(strsql2, cn, 1, "")
		Buscar = "<form name=formbuscar id=formbuscar  method=post>"

		Buscar = Buscar & "<font class=at>"
		Buscar = Buscar & "<select SIZE=10 class=at "
		Buscar = Buscar & " onchange=" & Chr(34) & "document.formbuscar.action = '" & vuelta & sx & " CantRegMostrar=" & CantRegMostrarx
		Buscar = Buscar & "&CantRegMover=' + document.formbuscar.buscarlis[this.selectedIndex].value; document.formbuscar.submit();"
		Buscar = Buscar & Chr(34) & " name=buscarlis id=buscarlis >"
		Buscar = Buscar & "<option class='OPCM1' value=" & CantRegMoverx & " selected >------- Seleccione busqueda -------</option>"
		Do Until rx.EOF
			For YY = 0 To UBound(Perfiles)
				If Perfilesd(YY) = rx(0) Then
					Buscar = Buscar & "<option value=" & Perfiles(YY) & " >" & Perfilesd(YY) & "</option>"
					Perfilesd(YY) = "***"
					Exit For
				End If
			Next
			rx.movenext
		Loop
		Call CerrarConsulta(rx)
		Buscar = Buscar & "</select></font></form>"
		'  If m_Browser = "IE" Then
		sImagenx = "<B><FONT FACE=CorFlaticon COLOR=#6699cc SIZE=1>r</FONT></B>"
		'  Else
		'  sImagenx = "<B><FONT FACE=Wingdings COLOR=#6699cc SIZE=1>&#8999;</FONT></B>"
		'  End If

		''   sinput= "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & Chr(34)
		''   sinput= sinput& "underline" & Chr(34) & ";' onmouseout='this.style.textDecoration=" & Chr(34)
		''    sinput= sinput& "none" & Chr(34) & ";' href='javascript:void(0);' ONCLICK=" & Chr(34) & "document.all.DBSC.style.display='none';" & Chr(34)
		''  sinput= sinput& " TITLE='OCULTAR' ><b>" & sImagenx & "</b></a>"

		sDiver = "<div id='DBSC' class='DINDM' style='display:none;"
		sDiver = sDiver & "z-index:80;margin-top:1;margin-left:1;position:absolute;"
		sDiver = sDiver & "left:300px;top:200px;'"
		sDiver = sDiver & " >"


		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

		sDiver = sDiver & "<table id='DBSC1'  class='TWIND' width=cellspacing=0 cellpadding=>"
		sDiver = sDiver & "<tr>"
		sDiver = sDiver & "<td class=mbt width=99%  >&nbsp;"

		sDiver = sDiver & UCase(FBotTab(4, 6))

		sDiver = sDiver & "</td>"
		sDiver = sDiver & "<td CLASS=mbt width=1% onclick=""document.getElementById('DBSC').style.display='none';"" title='Ocultar.' >"
		'If m_Browser = "IE" Then

		sDiver = sDiver & "<font  onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & Session("CMFMNBOT") & "';"" face=CorFlaticon size=>r</font>"
		'Else
		'sDiver = sDiver & "<font  onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & Session("CMFMNBOT") & "';"" face=Wingdings size=>&#8999;</font>"

		'"<font  onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & Session("CMFMNBOT") & "';"" face=Wingdings size=>&#8999;</font>"
		'End If
		sDiver = sDiver & "</td>"
		sDiver = sDiver & "</tr>"
		sDiver = sDiver & "</table>"

		sDiver = sDiver & "<table id=""DBSC2""  class='TWINE2' width=cellspacing=4 cellpadding="
		sDiver = sDiver & " >"
		sDiver = sDiver & "<tr>"
		sDiver = sDiver & "<td class=ut align=center  >"
		sDiver = sDiver & Buscar
		sDiver = sDiver & "</td>"
		sDiver = sDiver & "</tr>"
		sDiver = sDiver & "</table>"

		'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


		sDiver = sDiver & "</div>"


		sDiver = sDiver & "<script language=javascript>" & Chr(13)
		sDiver = sDiver & "function AbreBusc() {" & Chr(13)
		sDiver = sDiver & "  if (document.getElementById('DBHL')){document.getElementById('DBHL').style.display='none';}" & Chr(13)
		sDiver = sDiver & "  document.getElementById('DBSC').style.display='block';" & Chr(13)
		sDiver = sDiver & "}" & Chr(13)
		sDiver = sDiver & "</script>" & Chr(13)

		FBBarr = FBBarr & sDiver

		ComboBuscar = FBmodo

		Exit Function

		ComboBuscarErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		ComboBuscar = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function ComboFiltrar(CantRegMostrarx, CantRegMoverx, vuelta, filtro)

		'On Error GoTo ComboFiltrarErr

		Dim Filtrar

		Filtrar = "<form name=formfiltrar  id=formfiltrar method=post>"
		Filtrar = Filtrar & "<td height=20>"
		Filtrar = Filtrar & "<input class=bt type=button size=10% value='Mostrar' ></td>"
		Filtrar = Filtrar & "<td><font class=at><select class=at onchange=" & Chr(34) & "document.formfiltrar.action = '" & ConvertirCar(vuelta, "_", "&") & "&Filtrar=' + document.formfiltrar.filtrarlis[this.selectedIndex].value; document.formfiltrar.submit();" & Chr(34) & " name=filtrarlis >"
		Filtrar = Filtrar & "<option value=no>TODOS"
		If filtro = "cri" Then
			Filtrar = Filtrar & "<option value=cri selected>CRITICOS"
		Else
			Filtrar = Filtrar & "<option value=cri>CRITICOS"
		End If
		If filtro = "nocri" Then
			Filtrar = Filtrar & "<option value=nocri selected>NO CRITICOS"
		Else
			Filtrar = Filtrar & "<option value=nocri>NO CRITICOS"
		End If
		Filtrar = Filtrar & "</select></font></td></form>"

		ComboFiltrar = Filtrar

		Exit Function

		ComboFiltrarErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		ComboFiltrar = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function ValClaveRep(Clave)

		'On Error GoTo ValClaveRepErr

		'testea si un valor en un imput esta contenido en un vector
		Dim Repetido, j

		Repetido = "<script language=JavaScript >"
		Repetido = Repetido & " function vectord(Clave) "
		Repetido = Repetido & " { var vectclave = new Array("
		For j = 0 To UBound(Clave)
			Repetido = Repetido & Chr(34) & Clave(j) & Chr(34) & ","
		Next
		Repetido = Mid(Repetido, 1, Len(Repetido) - 1) & "); "
		Repetido = Repetido & " return vectclave[Clave]; } "
		Repetido = Repetido & " function valida(dato) { sale=0; "
		Repetido = Repetido & " for(i=0; i < " & UBound(Clave) + 1 & "; i++) "
		Repetido = Repetido & "{ if ( vectord(i) == dato.value) { "
		Repetido = Repetido & " alert ('El Codigo ' + dato.value + ' "
		Repetido = Repetido & "ya fue utilizado.'); "
		Repetido = Repetido & "dato.value = dato.defaultValue; "
		Repetido = Repetido & "dato.focus(); sale=1; i=" & UBound(Clave)
		Repetido = Repetido & " + 2; } } if (sale==1) {return false;} }"
		Repetido = Repetido & "</script>"

		ValClaveRep = Repetido

		Exit Function

		ValClaveRepErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		ValClaveRep = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function VectorClave(str, rspos)

		'On Error GoTo VectorClaveErr

		Dim n, Clave()
		Dim cn , rs

		Set cn = Conectar(strconn4)

		Set rs = CrearConsultaEx(str, cn, 1, "")
		If rs.EOF Then
			ReDim Preserve Clave(0)
			Clave(0) = ""
		End If
		n = 0
		Do Until rs.EOF
			ReDim Preserve Clave(n)
			Clave(n) = rs(rspos)
			rs.movenext
			n = n + 1
		Loop
		Call CerrarConsulta(rs)

		VectorClave = Clave

		Exit Function

		VectorClaveErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		VectorClave = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function Nombre(str, Valor, cn)

		'On Error GoTo NombreErr

		Dim rnom

		Nombre = ""

		Set rnom = CrearConsultaEx(str, cn, 1, "")
		If Not rnom.EOF Then Nombre = rnom(0)

		Call CerrarConsulta(rnom)

		Exit Function

		NombreErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		Nombre = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function ValidarActualizacion(bNumber, Gform, Gcampo)

		'On Error GoTo ValidarActualizacionErr

		Dim s, Gnombre

		bNumber = UCase(bNumber)
		Gnombre = DescCampo(Gcampo)

		s = s & "onblur=" & Chr(34) & "if (this.value.length == 0) "
		s = s & "{alert('" & Gnombre & " no debe ser nulo'); this.focus(); "
		s = s & "this.value=this.defaultValue;this.select(); return false;} "
		s = s & "else { if (this.value != this.defaultValue) {"
		If bNumber = "SI" Or bNumber = "SIR" Then
			s = s & "if (isNaN(this.value) || this.value < 0) {alert('" & Gnombre & " debe ser numerico');"
			s = s & " this.focus();this.value=this.defaultValue;this.select();"
			s = s & " return false;} else {"
		End If
		If bNumber = "SI" Or bNumber = "" Then
			s = s & " document." & Gform & ".submit(); return true;} "
		End If
		If bNumber = "SIR" Or bNumber = "NOR" Then
			s = s & " valida(this);if (sale==0) {document." & Gform & ".submit();} } "
		End If
		If bNumber = "SI" Or bNumber = "SIR" Then s = s & "}"
		s = s & " else {return false;} };" & Chr(34)

		ValidarActualizacion = s

		Exit Function

		ValidarActualizacionErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		ValidarActualizacion = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function ValidarActualizacionCombo(Gform, Gcampo)

		'On Error GoTo ValidarActualizacionComboErr

		Dim s, Gnombre

		Gnombre = DescCampo(Gcampo)

		s = s & " onchange=" & Chr(34) & "if (this.options[this.selectedIndex].value=='nsnc') { "
		s = s & "alert('La seleccion es incorrecta para " & Gnombre & "');"
		s = s & "for (i = 0; i < this.options.length; i++) {if (this.options[i].defaultSelected== true){"
		s = s & "this.options[i].selected=true;}};"

		s = s & " return false;} else {" & Gform & ".submit(); return true;};" & Chr(34)

		ValidarActualizacionCombo = s

		Exit Function

		ValidarActualizacionComboErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		ValidarActualizacionCombo = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function DescCampo(Gcampo)

		'On Error GoTo DescCampoErr

		Dim IDCol(21), DescCol(21), yn

		IDCol(0) = "descri"
		DescCol(0) = "Descripcion"
		IDCol(1) = "tipoperfil"
		DescCol(1) = "Perfil"
		IDCol(2) = "codinter"
		DescCol(2) = "Codigo interno"
		IDCol(3) = "ecomaster"
		DescCol(3) = "Dependencia"
		IDCol(4) = "codperfil"
		DescCol(4) = "Perfil"
		IDCol(5) = "sisgen"
		DescCol(5) = "Descripcion"
		IDCol(6) = "tipogen"
		DescCol(6) = "Descripcion"
		IDCol(8) = "tipobjeto"
		DescCol(8) = "Tipo"
		IDCol(9) = "tipomodo"
		DescCol(9) = "Tipo"
		IDCol(10) = "codiden"
		DescCol(10) = "Identificacion"
		IDCol(11) = "codsector"
		DescCol(11) = "Sector"
		IDCol(12) = "secuen"
		DescCol(12) = "Secuencia"
		IDCol(13) = "fecemis"
		DescCol(13) = "Fecha de emision"
		IDCol(14) = "instrumen"
		DescCol(14) = "Instrumento"
		IDCol(15) = "instid"
		DescCol(15) = "Identificacion del instrumento"
		IDCol(16) = "inda"
		DescCol(16) = "Orden"
		IDCol(17) = "codasoc"
		DescCol(17) = "Asociacion"
		IDCol(18) = "tipocontin"
		DescCol(18) = "Codigo"
		IDCol(19) = "codnivel"
		DescCol(19) = "Nivel"
		IDCol(20) = "plazomax"
		DescCol(20) = "Plazo"
		IDCol(21) = "indb"
		DescCol(21) = "Orden"

		DescCampo = Gcampo
		For yn = 0 To UBound(IDCol)
			If Gcampo = IDCol(yn) Then
				DescCampo = DescCol(yn)
				Exit For
			End If
		Next

		Exit Function

		DescCampoErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		DescCampo = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function Conectar(strconn)

		Dim STCON
		If strconn <> "" Then
			'On Error GoTo ConectarErr
			STCON = UCase(Session("TIPOCON") & "")

			Select Case STCON

				Case ""

				If ATrusted <> "S" Then
					strconn = strconn & ";UID=" & ADbUidl
					strconn = strconn & ";PWD=" & GetCript("D", ADbPwdl)
				End If

				ErrorDesc = ""
				Set conn = Server.CreateObject("adodb.connection")
				conn.Open strconn
				'Set Conectar = conn

				Case "OLEDBE"

				If ATrusted <> "S" Then
					'strconn = strconn & ";UID=" & ADbUidl
					'strconn = strconn & ";PWD=" & GetCript("D", ADbPwdl)
					Set conn = Server.CreateObject("adodb.connection")
					'response.write strconn & "222"
					conn.Open strconn, ADbUidl, GetCript("D", ADbPwdl)
					'   Set Conectar = conn
				Else
					Set conn = Server.CreateObject("adodb.connection")
					conn.Open strconn
				End If

				ErrorDesc = ""
				Case "OLEDB"
				If ATrusted <> "S" Then
					'strconn = strconn & ";UID=" & ADbUidl
					'strconn = strconn & ";PWD=" & GetCript("D", ADbPwdl)
					Set conn = Server.CreateObject("adodb.connection")
					conn.Open strconn, ADbUidl, ADbPwdl
					'Set Conectar = conn
				Else
					Set conn = Server.CreateObject("adodb.connection")
					conn.Open strconn
				End If

				ErrorDesc = ""


				Case "CNFREE"

				Set conn = Server.CreateObject("adodb.connection")
				conn.Open strconn

				ErrorDesc = ""


			End Select
			Set Conectar = conn
			Exit Function

			ConectarErr:
			With oErr
				.Number = Err.Number
				.Desc = Err.Description
				.Source = Err.Source
				.Line = Erl
				.Version = App.Major & "." & App.Minor & "." & App.Revision
			End With

			'On Error Resume Next

			Conectar = oEH.RegError(oErr)

			If Err.Number <> 0 Then Response.WRITE Err.Number
		End If
	End Function

	Public Function CrearConsultaCn(strsql)

		'On Error GoTo CrearConsultaCnErr

		Set CrearConsultaCn = conn.Execute(strsql)

		Exit Function

		CrearConsultaCnErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		CrearConsultaCn = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function CrearConsulta(strsql, strconn)

		'On Error GoTo CrearConsultaErr

		Dim rs
		Const adUseClient = 3

		Set rs = Server.CreateObject("ador.recordset")

		rs.CursorLocation = adUseClient
		rs.Open strsql, strconn
		Set rs.ActiveConnection = Nothing
		Set CrearConsulta = rs
		Set rs = Nothing

		Exit Function

		CrearConsultaErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		CrearConsulta = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function CrearConsultaEx(ByRef strsql, ByRef conn, TipoConsulta, ByRef a)

		'On Error GoTo CrearConsultaExErr

		Dim rs

		ErrorDesc = ""
		Set rs = Server.CreateObject("adodb.command")
		rs.ActiveConnection = conn
		rs.CommandText = strsql
		rs.CommandType = TipoConsulta

		Select Case TipoConsulta
			Case 1
			On Error Resume Next
			' response.write strsql
			Set CrearConsultaEx = conn.Execute(strsql) 'rs.Execute()
			Case 4
			'            ConvertirArregloParam a, carblanco, " "
			Set CrearConsultaEx = rs.Execute(, a)
		End Select
		Set rs = Nothing

		Exit Function

		CrearConsultaExErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		CrearConsultaEx = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function CrearConsulta2(ByRef strsql, ByRef conn)

		'On Error GoTo CrearConsulta2Err

		Dim rs

		'Abre un cursor estatico (con desplazamiento hacia atras y adelante)
		Set rs = Server.CreateObject("adodb.recordset")

		rs.ActiveConnection = conn
		rs.CursorLocation = 3
		rs.CursorType = 3
		rs.Open strsql
		Set rs.ActiveConnection = Nothing

		Set CrearConsulta= rs
		Set rs = Nothing

		Exit Function

		CrearConsulta2Err:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		CrearConsulta= oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Function ConsultaVacia(ByRef rs)

		'On Error GoTo ConsultaVaciaErr

		ConsultaVacia = (rs.EOF And rs.BOF)

		Exit Function

		ConsultaVaciaErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		ConsultaVacia = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	Public Sub CerrarConsulta(ByRef consulta)

		On Error Resume Next

		consulta.Close
		Set consulta = Nothing

	End Sub

	Public Sub Cerrarconn(ByRef conn)

		On Error Resume Next

		conn.Close
		Set conn = Nothing

	End Sub

	Public Function ValidarUsuario()
		Dim strconnx

		'On Error GoTo ValidarUsuarioErr

		Dim oControl
		Dim STCON
		'   Dim oControl

		'Set oControl = CreateObject("ControlarAplicacion.Controlar")

		Set oControl = Server.CreateObject("controlamultiex.controlar")

		Dim iRet
		strconnx = strconn1
		STCON = UCase(Session("TIPOCON") & "")
		Select Case STCON
			Case "OLEDBE"
			If ATrusted <> "S" Then
				strconnx = strconnx & ";user id =" & ADbUid
				strconnx = strconnx & ";password=" & GetCript("D", ADbPwd)
			End If
			Case Else
			If ATrusted <> "S" Then
				strconnx = strconnx & ";uid=" & ADbUid
				strconnx = strconnx & ";pwd=" & GetCript("D", ADbPwd)
			End If
		End Select

		Set oControl.Server = Server
		oControl.strconn0 = strconnx

		iRet = oControl.Controlar
		If iRet = 0 Then
			ValidarUsuario = "OK"
			UV = "SI"
		Else
			ValidarUsuario = oControl.ErrorControl
			UV = "NO"
		End If

		Set oControl = Nothing

		Exit Function

		ValidarUsuarioErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		ValidarUsuario = oEH.RegError(oErr)

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Function

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'' Seguridad


	Public Sub ObtenerAtributos(sForm,  sNodo)

		Call ObtenerAtributos2(sForm,  sNodo,  "",  "")
	End Sub


	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Public Sub ObtenerAtributos2(sForm,  sNodo,  sCampo,  sValor)

		'On Error GoTo ObtenerAtributosErr

		Dim strsql, atrok, cn, rs, sSQL, xcount , XMFI , XMFSEP, xsql
		Dim XMFTGRP
		' Dim a, b, c, d, e
		AEstado = True
		AConsultar = True
		AEliminar = True
		AModificar = True
		AAgregar = True

		Dim XTabPer(4)

		xcount = 0


		Set cn = Conectar(strconn0)

		atrok = "N"
		strsql = "Select Alias,usrEstado,usragregar,usrmodificar,usreliminar,usrconsultar "
		strsql = strsql & " From atrmXusuario Where "
		If sNodo <> "" Then
			strsql = strsql & "Tipo = 'N'  And alias = '" & sNodo & "'  And depen = '" & sForm & "'  "
		Else
			strsql = strsql & "Tipo = 'F'  And depen = '" & sForm & "'  "
		End If
		strsql = strsql & " And usuario = '" & Session("segperfil") & "' And "
		strsql = strsql & "  Aplicacion = '" & Session("CorApli") & "'"
		'  xsql = strsql & Chr(13)
		Set rs = CrearConsultaEx(strsql, cn, 1, "")
		If rs.EOF Then
			atrok = "O"
		Else
			atrok = "S"
			AEstado = rs(1)
			AConsultar = rs(5)
			AEliminar = rs(4)
			AModificar = rs(3)
			AAgregar = rs(2)
		End If
		Call CerrarConsulta(rs)

		If atrok = "O" Then
			strsql = "Select Alias,grpEstado,grpagregar,grpmodificar,grpeliminar,grpconsultar "
			strsql = strsql & " From atrmXgrupo Where  "
			If sNodo <> "" Then
				strsql = strsql & "Tipo = 'N'  And alias = '" & sNodo & "'  And depen = '" & sForm & "'  "
			Else
				strsql = strsql & "Tipo = 'F'  And depen = '" & sForm & "'  "
			End If
			strsql = strsql & " And Aplicacion = '" & Session("CorApli") & "'"

			If Session("DELLOG") = "S" Then

				XMFTGRP = Split(Session("SegGrpLst"), Chr(9))
				xcount = UBound(XMFTGRP) + 1
				XMFSEP = ""

				For XMFI = 0 To 4
					XTabPer(XMFI) = xcount
				Next

				strsql = strsql & " and grupo in ("

				For XMFI = 0 To UBound(XMFTGRP)
					strsql = strsql & XMFSEP & "'" & XMFTGRP(XMFI) & "'"
					XMFSEP = ","
				Next

				strsql = strsql & ") "
				' xsql = xsql & strsql & Chr(13)
				Set rs = CrearConsultaEx(strsql, cn, 1, "")
				If rs.EOF Then
					atrok = "O"
				Else
					atrok = "S"
					Do While Not rs.EOF


						If Not rs(5) Then XTabPer(1) = XTabPer(1) - 1 ''' consultar
						If Not rs(4) Then XTabPer(2) = XTabPer(2) - 1 ''' eliminar
						If Not rs(3) Then XTabPer(3) = XTabPer(3) - 1 ''' modificar
						If Not rs(2) Then XTabPer(4) = XTabPer(4) - 1 ''' agregar

						rs.movenext
					Loop

				End If

				If atrok = "S" Then


					If XTabPer(1) = 0 Or XTabPer(1) = False Then AConsultar = False
					If XTabPer(2) = 0 Or XTabPer(2) = False Then AEliminar = False
					If XTabPer(3) = 0 Or XTabPer(3) = False Then AModificar = False
					If XTabPer(4) = 0 Or XTabPer(4) = False Then AAgregar = False


				End If


			Else

				strsql = strsql & " And grupo = '" & Session("seggrp") & "'  "
				' xsql = xsql & strsql & Chr(13)
				Set rs = CrearConsultaEx(strsql, cn, 1, "")
				If rs.EOF Then
					atrok = "O"
				Else
					atrok = "S"

					AEstado = CBool(rs(1))
					AConsultar = CBool(rs(5))
					AEliminar = CBool(rs(4))
					AModificar = CBool(rs(3))
					AAgregar = CBool(rs(2))
				End If

			End If

			Call CerrarConsulta(rs)
		End If

		Call Cerrarconn(cn)
		' Session("xsql") = xsql
		Exit Sub

		ObtenerAtributosErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		sSQL = oEH.RegError(oErr)
		Response.WRITE sSQL

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Sub

	Public Sub ObtenerAtributosNodo(sForm,  sNodo,  sCampo,  sValor)

		'On Error GoTo ObtenerAtributosErr

		Dim strsql, atrok, cn, rs, sSQL, xcount , XMFI , XMFSEP
		Dim XMFTGRP

		AEstado = True
		AConsultar = True
		AEliminar = True
		AModificar = True
		AAgregar = True

		Dim XTabPer(4)

		xcount = 0



		Set cn = Conectar(strconn0)

		atrok = "N"
		strsql = "Select Alias,usrEstado,usragregar,usrmodificar,usreliminar,usrconsultar "
		strsql = strsql & " From atrmXusuario Where "
		If sCampo <> "" Then
			strsql = strsql & "Tipo = 'N'  And nodo = '" & sCampo & "'  And valor = '" & sValor & "' And depen = '" & sForm & "'  "
		ElseIf sNodo <> "" Then
			strsql = strsql & "Tipo = 'N'  And alias = '" & sNodo & "'  And depen = '" & sForm & "'  "
		Else
			strsql = strsql & "Tipo = 'F'  And alias = '" & sForm & "'  "
		End If
		strsql = strsql & " And usuario = '" & Session("segperfil") & "' And "
		strsql = strsql & "  Aplicacion = '" & Session("CorApli") & "'"
		Set rs = CrearConsultaEx(strsql, cn, 1, "")
		If rs.EOF Then
			atrok = "O"
		Else
			atrok = "S"
			AEstado = rs(1)
			AConsultar = rs(5)
			AEliminar = rs(4)
			AModificar = rs(3)
			AAgregar = rs(2)
		End If
		Call CerrarConsulta(rs)

		If atrok = "O" Then
			strsql = "Select Alias,grpEstado,grpagregar,grpmodificar,grpeliminar,grpconsultar "
			strsql = strsql & " From atrmXgrupo Where  "
			If sCampo <> "" Then
				strsql = strsql & "Tipo = 'N'  And nodo = '" & sCampo & "'  And valor = '" & sValor & "' And depen = '" & sForm & "'  "
			ElseIf sNodo <> "" Then
				strsql = strsql & "Tipo = 'N'  And alias = '" & sNodo & "'  And depen = '" & sForm & "'  "
			Else
				strsql = strsql & "Tipo = 'F'  And alias = '" & sForm & "'  "
			End If
			strsql = strsql & " And Aplicacion = '" & Session("CorApli") & "'"

			If Session("DELLOG") = "S" Then

				XMFTGRP = Split(Session("SegGrpLst"), Chr(9))
				xcount = UBound(XMFTGRP) + 1
				XMFSEP = ""

				For XMFI = 0 To 4
					XTabPer(XMFI) = xcount
				Next

				strsql = strsql & " and grupo in ("

				For XMFI = 0 To UBound(XMFTGRP)
					strsql = strsql & XMFSEP & "'" & XMFTGRP(XMFI) & "'"
					XMFSEP = ","
				Next

				strsql = strsql & ") "
				Set rs = CrearConsultaEx(strsql, cn, 1, "")
				If rs.EOF Then
					atrok = "O"
				Else
					atrok = "S"

					Do While Not rs.EOF

						If Not rs(5) Then XTabPer(1) = XTabPer(1) - 1 ''' consultar
						If Not rs(4) Then XTabPer(2) = XTabPer(2) - 1 ''' eliminar
						If Not rs(3) Then XTabPer(3) = XTabPer(3) - 1 ''' modificar
						If Not rs(2) Then XTabPer(4) = XTabPer(4) - 1 ''' agregar

						rs.movenext
					Loop

				End If

				If atrok = "S" Then

					If XTabPer(1) = 0 Or XTabPer(1) = False Then AConsultar = False
					If XTabPer(2) = 0 Or XTabPer(2) = False Then AEliminar = False
					If XTabPer(3) = 0 Or XTabPer(3) = False Then AModificar = False
					If XTabPer(4) = 0 Or XTabPer(4) = False Then AAgregar = False

				End If


			Else

				strsql = strsql & " And grupo = '" & Session("seggrp") & "'  "
				Set rs = CrearConsultaEx(strsql, cn, 1, "")
				If rs.EOF Then
					atrok = "O"
				Else
					atrok = "S"
					AEstado = rs(1)
					AConsultar = rs(5)
					AEliminar = rs(4)
					AModificar = rs(3)
					AAgregar = rs(2)
				End If

			End If

			Call CerrarConsulta(rs)
		End If

		Call Cerrarconn(cn)

		Exit Sub

		ObtenerAtributosErr:
		With oErr
			.Number = Err.Number
			.Desc = Err.Description
			.Source = Err.Source
			.Line = Erl
			.Version = App.Major & "." & App.Minor & "." & App.Revision
		End With

		'On Error Resume Next

		sSQL = oEH.RegError(oErr)
		Response.WRITE sSQL

		If Err.Number <> 0 Then Response.WRITE Err.Number

	End Sub

	Private Sub Class_Initializeold()
		'  response.write "222222222"
		ATrusted = "N"

		ADbUid = "sadmin"
		ADbPwd = "781784763767752745723"
		strconn0="Provider=SQLOLEDB;Data Source=server-40;Initial Catalog=Dbcordatv5;user id=sadmin;password=manager;"

		strconn1 = strconn0
		strconn2 = strconn0 '"Provider=SQLOLEDB;Data Source=server-40;Initial Catalog=Dbcordatv5;user id=sadmin;password=manager;"
		strconn3 = ostrconn0
		strconn4 = "dsn=Dbcordat3sql"

		strconn5 ="Provider=SQLOLEDB;Data Source=server-40;Initial Catalog=Dbcorrqtv5"
		strconn6=strconn5


		strconn7 =strconn2
		'strconn8 ="PROVIDER=SQLOLEDB;Data Source=server-40;Database=Dbsearchv5;User Id=sadmin;Password=manager;"
		'ofv.strconn6=strconn8
		'Set oEH = New ErrorHandler

	End Sub

	Private Sub Class_Terminate()

		Set oEH = Nothing

	End Sub

End Class
%>