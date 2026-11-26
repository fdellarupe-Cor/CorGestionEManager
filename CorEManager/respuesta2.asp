<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	
	Generar = ""
	
	Addqry = ""
	Response.expires = 0
	
	Server.scripttimeout = 600
	Response.buffer = True
	
	'Response.WRITE Request.querystring & "<BR>"
	'Response.WRITE Request.FORM & "<br><br>"
	'
	
	XXXX    = NEWEMPRE("NO")
	smenu   = Request.querystring("MENU")
	sorden  = Request.querystring("sorden")
	
	
	SMENU= Trim(sMenu)
	
	base=request.querystring("BASE")
	
	If base <> "" Then
		SESSION("BASE")=base
	End If
	
	
	If Session("CorApli") = "COREM20" Then
		wkaplic = "EM"
	ElseIf Session("CorApli") = "COREP20" Then
		wkaplic = "PL"
	Else
		wkaplic = "OR"
	End If
	
	
	If Len(sorden) = 0 Then
		sorden = "N"
	End If
	
	Sfiltro = Request.form("sfiltro")
	
	If Len(Sfiltro) = 0 Then
		Sfiltro = Request.querystring("sfiltro")
	End If
	
	If Len(Sfiltro) = 0 Then
		Sfiltro = "TODOS"
	End If
	
	Sfiltro = Replace(Sfiltro, "_", " ")
	CantRegMostrar = 10
	CantRegMover = CDbl("0" &Request.querystring("CantRegMover"))
	ReDim Preserve bconsultar(4)
	OFV.ObtenerAtributosNodo Session("Form"), "", "ESTADO", "4"
	bconsultar(4) = OFV.AConsultar
	OFV.ObtenerAtributosNodo Session("Form"), "", "ESTADO", "2"
	bconsultar(2) = OFV.AConsultar
	OFV.ObtenerAtributosNodo Session("Form"), "", "ESTADO", "1"
	bconsultar(1) = OFV.AConsultar
	OFV.ObtenerAtributosNodo Session("Form"), "", "ESTADO", "3"
	bconsultar(3) = OFV.AConsultar
	
	'response.write Session("Form") & bconsultar(3)
	For i = 1 To 4
		If bconsultar(i) Then
			sverestado = sverestado & i & ","
			If i = 1 Then sverestado = sverestado & "6" & ","
			If i = 2 Then sverestado = sverestado & "5" & ","
			
		End If
	Next
	sverestado = "(" & Left(sverestado, Len(sverestado) - 1) & ") "
	' response.write sverestado
	''''''     FILTRO POR EMPRESA
	
	addfrom =  ""
	'response.write request.form & "</BR></BR></BR>"
	'response.write SMENU & " ZZZZ</BR></BR></BR>"
	
	If smenu = "CONSGRAL" Then
		
		
		
		
		
		FESTADO         = CStr(request.form("estado") & "")
		
		FBANCA          = CStr(request.form("banca")& "")
		FGUARDA         = CStr(request.form("TGUARDA")& "")
		FPRODSERV       = CStr(request.form("PRODU-SERVI")& "")
		FCATEGORIA      = CStr(request.form("CATEGORIAF")& "")
		FCLASE          = CStr(request.form("CLASEF")& "")
		FCODIGO         = CStr(request.form("CODIGOF")& "")
		FFIRMAS         = CStr(request.form("FIRMASF")& "")
		IMPRIME         = CStr(request.form("IMPRESION")& "")
		TFECHA          = CStr(request.form("FECHA")& "")
		FDESDE          = CStr(request.form("DESDE2")& "")
		FHASTA          = CStr(request.form("HASTA2")& "")
		FPROCESO          = CStr(request.form("PROCESO")& "")
		
		
		
		If FESTADO <> ""  Then
			session("FESTADO")=FESTADO
		Else
			FESTADO=session("FESTADO")
		End If
		
		
		If TFECHA <> ""  Then
			session("TFECHA")=TFECHA
		Else
			TFECHA=session("TFECHA")
		End If
		
		If FDESDE <> ""  Then
			session("FDESDE")=FDESDE
		Else
			FDESDE=session("FDESDE")
		End If
		If FHASTA <> ""  Then
			session("FHASTA")=FHASTA
		Else
			FHASTA=session("FHASTA")
		End If
		If FBANCA <> ""  Then
			session("FBANCA")=FBANCA
		Else
			FBANCA=session("FBANCA")
		End If
		
		
		If FGUARDA <> ""  Then
			session("FGUARDA")=FGUARDA
		Else
			FGUARDA=session("FGUARDA")
		End If
		
		
		If FPROCESO <> ""  Then
			session("FPROCESO")=FPROCESO
		Else
			FPROCESO=session("FPROCESO")
		End If
		
		
		If FPRODSERV <> ""  Then
			session("FPRODSERV")=FPRODSERV
		Else
			FPRODSERV=session("FPRODSERV")
		End If
		
		
		If FCATEGORIA <> ""  Then
			session("FCATEGORIA")=FCATEGORIA
		Else
			FCATEGORIA=session("FCATEGORIA")
		End If
		
		If FCLASE <> ""  Then
			session("FCLASE")=FCLASE
		Else
			FCLASE=session("FCLASE")
		End If
		
		If FCODIGO <> ""  Then
			session("FCODIGO")=FCODIGO
		Else
			FCODIGO=session("FCODIGO")
		End If
		
		If FFIRMAS <> ""  Then
			session("FFIRMAS")=FFIRMAS
		Else
			FFIRMAS=session("FFIRMAS")
		End If
		
		
		If IMPRIME <> ""  Then
			session("IMPRIME")=IMPRIME
		Else
			IMPRIME=session("IMPRIME")
		End If
		
		
		
		
		
		'response.write  request.form
		
	End If
	scal=request.querystring("scal")
	If smenu = "PORACTA" Then
		
		
		
	ElseIf smenu = "PORIDENTIFICACION" Then
		
		
	Else
		
		If  Robj = "F" Then
			
			SPerfil = Request.Form("perfil") & ""
			sID     = Request.Form("lista") & ""
			sID2    = Request.Form("lista") & ""
			sID3    = Request.Form("buscar") & ""
			
			Addqry  = "&Lista=" & sID2 & "&buscar=" & sID3 & "&perfil=" & SPerfil
			
			SPerfil = Replace(SPerfil, "_", " ")
			
		Else
			
			SPerfil = Request.querystring("perfil") & ""
			sID     = Request.querystring("Lista") & ""
			sID2    = Request.querystring("Lista") & ""
			sID3    = Request.querystring("buscar") & ""
			Addqry  = "&Lista=" & sID2 & "&buscar=" & sID3 & "&perfil=" & SPerfil
			SPerfil = Replace(SPerfil, "_", " ")
			
		End If
		
		sID         = OFV.convertircar(sID, "_", " ")
		If sID = "" Then sID = sID3
		
		sMenuAdd    = ""
		sMenuMas    = "&smenu=" & smenu & "&lista=" & sID2 & "&buscar="
		sMenuMas    = sMenuMas & sID3 & "&perfil=" & SPerfil
		
	End If
	
	
	
	Addqry = "&Menu=" & smenu & Addqry & "&Nombre=" & sNombre
	
	starget = "RESPUESTA2.asp?cantregmover=0&cantregmostrar=" & CantRegMostrar & Addqry &"&DESDE=" & FDESDE & "&HASTA=" & FHASTA
	sVolver = "RESPUESTA2.asp?Menu=" & smenu & sMenuAdd
	
	Addqry = Addqry & "&sorden=" & sorden & "&sfiltro=" & Replace(Sfiltro, " ", "_")
	
	'  response.write sorden &" - "
	
	'response.write Addqry &
	
	Select Case sorden
		Case "N"
		sTitOrd = sTitOrd & "Nombre"
		strsql  = strsql & " order by a.descri,csporfechas.fecvigen "
		Case "B"
		sTitOrd = sTitOrd & "Tipo"
		strsql  = strsql & " order by C.descri,csporfechas.fecvigen "
		Case "P"
		sTitOrd = sTitOrd & "Producto"
		strsql  = strsql & " order by  tp.descri,d.fecvigen desc "
		Case "E"
		sTitOrd = sTitOrd & "Estado"
		strsql  = strsql & " order by  EST.estado, A.DESCRI "
		Case "C"
		sTitOrd = sTitOrd & "Codigo"
		strsql  = strsql & " order by  b.codigoform "
		Case "V"
		sTitOrd = sTitOrd & "Vigencia"
		strsql  = strsql & " order by  csporfechas.fecvigen desc "
		Case "A"
		sTitOrd = sTitOrd & "Autor"
		strsql  = strsql & " order by  d.autor "
	End Select
	
	
	contar=""
	
	Set cn = OFV.conectar(OFV.strconn2)
	cn.CommandTimeout = 0
	
	contarsql=" SELECT count(distinct a.codobjeto) "
	
	newstrsql=""
	
	SQL= " SELECT distinct A.DESCRI,TP.DESCRI,c.descri,b.codigoform,EST.estado,csporfechas.FECVIGEN,D.autor,A.codobjeto,A.tipobjeto,D.FECPUBLIC,C.perfil ,cstr.temadescri,cstr.stemadescri,D.fecvigen,d.fecemis	 "
	
	newstrsql=newstrsql & "  FROM OBJETOS A inner join csrelpathdescriex cstr on (cstr.cobj=a.codobjeto) "
	NEWSTRSQL= NEWSTRSQL & " INNER JOIN OBJPROPIEDADES B ON ("
	
	
	If FGUARDA ="" Or FGUARDA="0" Then
		NEWSTRSQL= NEWSTRSQL & " A.CODOBJETO=B.CODOBJETO  "
	Else
		NEWSTRSQL= NEWSTRSQL & " A.CODOBJETO=B.CODOBJETO AND tiempoguarda='"& FGUARDA & "'"
	End If
	
	If FCATEGORIA ="" Or FCATEGORIA="0" Then
		
		'NEWSTRSQL= NEWSTRSQL & " AND catForm >'0' "
	Else
		NEWSTRSQL= NEWSTRSQL & " AND catForm='" & FCATEGORIA & "' "
	End If
	
	If FPROCESO ="" Or FPROCESO="0" Then
		
		'NEWSTRSQL= NEWSTRSQL & " AND catForm >'0' "
	Else
		NEWSTRSQL= NEWSTRSQL & " AND PROCESO='" & FPROCESO & "' "
	End If
	
	If FCLASE ="" Or FCLASE ="0" Then
		'  newstrsql = newstrsql & " AND claseform > '0' "
	Else
		
		NEWSTRSQL= NEWSTRSQL & " AND claseform='"& FCLASE &"' "
	End If
	
	If FCODIGO ="" Or FCODIGO ="0" Then
		' NEWSTRSQL= NEWSTRSQL & " and CODIGOFORM> '0' "
	Else
		NEWSTRSQL= NEWSTRSQL & " and CODIGOFORM='"& FCODIGO &"' "
	End If
	
	If FFIRMAS ="" Or FFIRMAS ="0" Then
		' NEWSTRSQL= NEWSTRSQL & " AND FIRMASFORM> '0'"
	Else
		NEWSTRSQL= NEWSTRSQL & " AND FIRMASFORM='"& FFIRMAS &"' "
	End If
	
	If IMPRIME ="" Or IMPRIME ="0" Then
	Else
		NEWSTRSQL= NEWSTRSQL & " AND modoimpresion='"& IMPRIME &"' "
	End If
	
	
	If FBANCA ="" Or FBANCA ="0" Then
	Else
		NEWSTRSQL= NEWSTRSQL & " AND BANCA='"& FBANCA &"' "
	End If
	
	
	NEWSTRSQL=NEWSTRSQL &")"
	
	NEWSTRSQL= NEWSTRSQL & " INNER JOIN TIPOBJETO C ON (A.TIPOBJETO=C.TIPOBJETO AND C.TIPOCPT='K') "
	NEWSTRSQL= NEWSTRSQL & " inner join csporfechas  on (a.codobjeto=csporfechas.codobjeto) "
	NEWSTRSQL= NEWSTRSQL & " INNER JOIN OBJPRTY D ON (D.CODOBJETO=A.CODOBJETO and d.estado in " & sverestado
	
	
	If TFECHA = "Emision" Then
		
		NEWSTRSQL= NEWSTRSQL & "  AND D.FECEMIS > = '"& FDESDE &"' AND csporfechas.fecemis  <= '"& FHASTA &"' "
		
	ElseIf TFECHA = "Vigencia" Then
		
		NEWSTRSQL= NEWSTRSQL & "  AND D.FECVIGEN  > = '"& FDESDE &"' AND csporfechas.FECVIGEN <= '"& FHASTA &"' "
		
	ElseIf TFECHA = "Eliminacion" Then
		
		NEWSTRSQL= NEWSTRSQL & "  AND D.FECBAJA > = '"& FDESDE &"' AND csporfechas.fecbaja  <= '"& FHASTA &"' "
		
	ElseIf TFECHA = "Vencimiento" Then
		
		NEWSTRSQL= NEWSTRSQL & " AND  D.FECVTO  > = '"& FDESDE &"' AND csporfechas.FECVTO <= '"& FHASTA &"' "
		
	ElseIf TFECHA = "Publicacion" Then
		
		NEWSTRSQL= NEWSTRSQL & " AND  D.FECPUBLIC > = '"& FDESDE &"' AND csporfechas.FECPUBLIC <= '"& FHASTA &"' "
		
	End If
	
	If FESTADO ="" Or FESTADO ="0" Then
		NEWSTRSQL= NEWSTRSQL & "  ) "
	Else
		NEWSTRSQL= NEWSTRSQL & "  and d.estado='"& FESTADO &"') "
	End If
	
	
	'   response.write FBANCA
	'        if FBANCA ="" or FBANCA ="0" then
	'  NEWSTRSQL= NEWSTRSQL & " INNER JOIN OBJTSEGAGR F ON (F.CODOBJETO=A.codobjeto AND F.TSEGMENTO >'0') "
	'      else
	'		    NEWSTRSQL= NEWSTRSQL & " INNER JOIN OBJTSEGAGR F ON (F.CODOBJETO=A.codobjeto AND F.TSEGMENTO='"& FBANCA &"') "
	'	    NEWSTRSQL= NEWSTRSQL & " INNER JOIN TIPOSEGMENTO TS ON (TS.OBJID=F.Tagrupa) "
	'
	'     end if
	
	
	
	
	NEWSTRSQL= NEWSTRSQL & " INNER JOIN ESTADOS EST ON (EST.objid=D.estado) "
	
	
	If FPRODSERV ="" Or FPRODSERV ="0" Then
		
		NEWSTRSQL= NEWSTRSQL & " left JOIN TIPOPRODUCTO TP ON (TP.OBJID=A.TSEGMENTO) "
	Else
		NEWSTRSQL= NEWSTRSQL & " INNER JOIN TIPOPRODUCTO TP ON (TP.OBJID=A.TSEGMENTO AND tp.objid='"& FPRODSERV &"') "
		' NEWSTRSQL= NEWSTRSQL & " WHERE tp.objid='"& FPRODSERV &"' "
	End If
	
	contarsql=contarsql & newstrsql
	
	NEWSTRSQL= NEWSTRSQL &  strsql
	
	sql=sql & newstrsql
	
	'sTitOrd  = ""
	'' sTitOrd2 = ""
	
	
	'response.write "<br><br><br>" & sql & "<br>"
	
	If request.querystring("scal")="S" Then
		sql=session("sqlcons")
	Else
		session("sqlcons")=sql
	End If
	'response.write "<br><br><br>" & request.querystring("scal") & "      " &  session("sqlcons") & "<br>"
	
	'response.write sql
	
	
	
	''      set rscont = ofv.crearconsultaEx(contarsql, cn, 1, "")
	Set rs = OFV.crearconsultaEx(contarsql, cn, 1, "")
	
	totalpag=0
	If Not rs.eof Then
		
		
		totalpag = rs(0)
		
	End If
	
	
	
	
	
	Set rs = OFV.crearconsultaEx(sql, cn, 1, "")
	
	
	
	sHTML = ofv.FormHeader(session("FormN")) & Chr(13)
	xs = "<script language=JavaScript>" & Chr(13)
	xs = xs & "function AbrirCons(Datos) {" & Chr(13)
	xs = xs & " var h=570;w=790 ;"
	If  session("APLMDCLI")="BC" Then
		xs = xs &  "if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;} "
	End If
	
	xs = xs & " myWin= open(Datos"
	xs = xs & "," & Chr(34) & "Cons" & Chr(34) & ","' & Chr(34)
	'xs = xs & "menubar=no,toolbar=no,height=570,width=790,resizable=yes,status=yes" & Chr(34)
	xs = xs &"  'menubar=no,toolbar=no,width='+w+',height='+h+',top=0,left=0,resizable=yes,status=yes'" '& CHR(34)
	xs = xs & " ); " & Chr(13) & " } " & Chr(13) & "</script>" & Chr(13)
	
	sHTML = sHTML & xs
	
	
	'   response.write Request.querystring & "<br>"
	'   response.write Request.form & "<br>"
	
	sImagenM = "<B><FONT FACE=""WEBDINGS"" COLOR=#005500 SIZE=3>2</FONT></B>"
	sImagenI2 = "<B><FONT FACE=""WEBDINGS"" COLOR=#f7aa22 SIZE=3>&#0128;</FONT></B>"
	sImagenI = "<B><FONT FACE=""WINGDINGS"" COLOR=#f7aa22 SIZE=2>u</FONT></B>"
	sImagenI3 = "<B><FONT FACE=""WEBDINGS"" COLOR=#6699CC SIZE=2>=</FONT></B>"
	sImagenEM = "<B><FONT FACE=WEBDINGS COLOR=#F7AA22 SIZE=3>&#0252;</FONT></B>"
	sImagen4si = "<B><FONT FACE=WEBDINGS COLOR=#993333 SIZE=3>4</FONT></B>"
	sImagenx = "<B><FONT FACE=WingDINGS COLOR=#aa0000 SIZE=4>x</FONT></B>"
	'sImagena = "<B><FONT FACE=WEBDINGS COLOR=#f7aa22 SIZE=2>4</FONT></B>"
	sImagena = "<B><i class='las la-angle-right' STYLE=COLOR:#f7aa22;></i></B>"

	
	
	Fila    = ""
	Celda   = ""
	sHref   = ""
	
	If  Session("EMPMUN") = "S" Then
		
		sLink   = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(0,5),"0","")
		Celda   = Celda & ofv.GenCelda("","tt ","","8%","","","","EMPRESA:&nbsp;&nbsp;","si")
		Celda   = Celda & ofv.GenCelda("","ttt ","center","5%","","","",sLink,"si")
		Celda   = Celda & ofv.GenCelda("","ttt  ","","","","","","&nbsp;<BIG>" & ATREMP(0,4) & "</BIG>","si")
		
		Fila    = fila & ofv.GenRow("","","top","","","","",Celda,"si")
		
		Tabla   = ofv.GenTabla("","aaa  ","","100% ","","0","","-1","-2","","",Fila,"si")
		sHTML   = sHTML & Tabla
	End If
	
	Fila = ""
	
	'***************** LINK DEL CAMBIAR VISTA
	sNombre=Replace(sNombre," ","_")
	Select Case smenu
		
		Case "CONSGRAL"
		sAccion = "ConsdocumentoswebScons.asp?xx=" & sql & Addqry &"&DESDE=" & FDESDE & "&HASTA=" & FHASTA ' "&id=" & sID & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N&Nombre=" & sNombre & "&Desde="& Request.querystring("Desde")  &"&Hasta="& Request.querystring("hasta") &"&Fecha="& si3
		sAccion = sAccion &"&nacta="& nacta
	End Select
	
	
	FDESDE= DateSerial(Year(CDate(FDESDE)),Month(CDate(FDESDE)),Day(CDate(FDESDE))+2)
	FHASTA= DateSerial(Year(CDate(FHASTA)),Month(CDate(FHASTA)),Day(CDate(FHASTA))+2)
	
	'response.write "<br> starget:"& starget
	
	sTitulo= "Desde: " & FDESDE & " - Hasta: " & FHASTA
	sTitulo=""
	
	'*****************************************
	
	sInput   = ofv.GenerarInput("","Vista Web","submit","","10","bt","")
	
	Formvista = ofv.GenForm("vistar","",sAccion,"","post"," style='display:none;' ",sInput,"si") & "Cambiar a Vista Web"
	
	Celda  = OFV.GenCelda("", "vt style='border-style:solid;border-width:1;border-color:#dddddd;'  ", "", "", "", "", "", "Total de Resultados: " & "&nbsp;" & totalpag & "", "si")
	Celda  = celda & OFV.GenCelda("", "vt style='border-style:solid;border-width:1;border-color:#dddddd;'  ", "", "", "", "", "", sImagena & "&nbsp;" & sTitulo & "", "si")
	
	If smenu <> "PORACTA" Then
		Celda  = Celda & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#dddddd;cursor:hand;' ","center onclick='document.all.vistar.submit();' ","20%","","","",Formvista,"si")
	Else
		Celda  = Celda & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#dddddd;cursor:hand;' ","center onclicka='document.all.vistar.submit();' ","20%","","","","&nbsp;","si")
	End If
	Celda  = Celda & ofv.GenCelda("","vt style='border-style:solid;border-width:1;border-color:#dddddd;' ","right","20%","","","","&nbsp;" & sTitOrd2 & "Orden:&nbsp;" & sTitOrd & "&nbsp;&nbsp;","si")
	
	
	Fila   = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
	
	tablaz = OFV.GenTabla("", "aaa bgcolor='#eeeeee'", "", "100% ", "", "0", "", "1", "2", "", "", Fila, "si")
	
	sHTML = sHTML & Tablaz
	
	Fila = ""
	
	'' AGREGADO 26052014
	
	If smenu ="PORESTADO" Then
		If ssEstado = "VIGENTE" Then                     'pregunto si es estado vigente para obtener los totales por tipo'
			strsql = " SELECT isnull(max(TIPOBJETO.descri),'SIN DATOS') as tipo, isnull(sum(1),0) as total "
			strsql = strsql & " FROM OBJETOS, TIPOBJETO, OBJPRTY, Estados , csrelpath a, tema t, prodserv  st "
			strsql = strsql & " where  OBJETOS.tipobjeto = TIPOBJETO.tipobjeto and OBJPRTY.estado = Estados.objid and OBJETOS.codobjeto = OBJPRTY.codobjeto "
			If SESSION("BASE")="K" Then
				strsql = strsql & " and OBJPRTY.estado = 3 and Objetos.objnivel >= 3 And Estados.objid In (9,1,6,2,5,3,4)  and Objetos.OBJUN = 1  and TIPOBJETO.TIPOCPT = 'K' "
			Else
				strsql = strsql & " and OBJPRTY.estado = 3 and Objetos.objnivel >= 3 And Estados.objid In (9,1,6,2,5,3,4)  and Objetos.OBJUN = 1  and TIPOBJETO.TIPOCPT = 'D' "
			End If
			
			strsql = strsql & " and (OBJETOS.APLIC = 'EM' or OBJETOS.APLIC = 'GLB')  and (a.cobj = OBJETOS.codobjeto)  and (a.tema = t.codtema) "
			strsql = strsql & " and (a.stema = st.codprodserv) "
			strsql = strsql & " group by  TIPOBJETO.descri "
			
			'response.write strsql
			Set rsx1 = OFV.crearconsultaEx(strsql, cn, 1, "")
			
			If Not rsx1.eof  Then
				
				'' Titulo de Totales por Tipo -unica vez-
				
				Celda = ""
				Celda  = OFV.GenCelda("", "vt xstyle='border-style:solid;border-width:5;border:none;'  ", "", "20%", "", "", "", "&nbsp;" & "TOTALES POR TIPO" & "", "si")
				Fila   = OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
				tablac = OFV.GenTabla("", "aaa bgcolor='#eeeeee' border=1 xstyle='border-style:solid;border-width:1;border:none;' ", "", "60% ", "", "0", "", "1", "2", "", "", Fila, "si")
				sHTML = sHTML & Tablac
				Fila = ""
				
				'' Rutina para mostrar los totales
				'' response.write " count "    & rsx1.fields.count   &"<br>"
				
				iii = 0
				iii3 = 0
				For iii = 1 To 1         'to 9 step 3
					
					'        response.write " &rsx1 " & rsx1(0) &"<br>"
					'        response.write " &iii "  & iii     &"<br>"
					'        response.write " Xrsx1X " & Hxcod & " " & Hxtip & " " & iii &"<br>"
					
					Fila  = ""
					Do While Not rsx1.eof
						Celda = ""
						For iii3=1 To 3
							If  rsx1.eof Then
								Exit For 'RSX1.MOVENEXT
								
							End If
							Hxcod = rsx1(0) & ""  'descripcion
							Hxtip = rsx1(1) & ""  'cantidad
							
							''                response.write " *rsx1 " & rsx1(0) &"<br>"
							
							Celda  = Celda & OFV.GenCelda("", "vt xstyle='border-style:solid;border-width:5;border:none;'  ", "", "20%", "", "", "", "&nbsp;" & Hxcod & ": " & Hxtip & "", "si")
							
							RSX1.MOVENEXT
						Next
						
						Fila   = fila & OFV.GenRow("", "", "", "", "", "", "", Celda, "si")
						
					Loop
					
				Next
				
			End If
		End If
		
		tablac = OFV.GenTabla("", "aaa bgcolor='#eeeeee' border=1 xstyle='border-style:solid;border-width:1;border:none;' ", "", "60% ", "", "0", "", "1", "2", "", "", Fila, "si")
		sHTML = sHTML & Tablac
		Fila = ""
		
	End If
	
	'' AGREGADO 26052014 FIN
	
	'' sHTML = OFV.MENUHeader("","")
	
	'  response.write "<br><br><br><br>" & Request.form & "<br>"
	If SESSION("BASE")="K" Then
		Celda = OFV.GenCelda("", "ttx colspan=1", "", "40%", "", "", "", "<a class=tt href=" & starget & "&sorden=N target=_self ><B><u>&nbsp;FORMULARIO&nbsp;</u></B></A>", "si")
		Celda = CELDA & OFV.GenCelda("", "ttx", "", "10%", "", "", "", "<a class=tt href=" & starget & "&sorden=B target=_self ><B><u>&nbsp;TIPO&nbsp;</u></B></A>", "si")
		Celda = CELDA &  OFV.GenCelda("", "ttx", "", "10%", "", "", "", "<a class=tt href=" & starget & "&sorden=P target=_self ><B><u>&nbsp;PRODUCTO&nbsp;</u></B></A>", "si")
		Celda =  CELDA &  OFV.GenCelda("", "ttx", "", "10%", "", "", "", "<a class=tt href=" & starget & "&sorden=C target=_self ><B><u>&nbsp;CODIGO&nbsp;</u></B></A>", "si")
		
		Celda = CELDA & OFV.GenCelda("", "ttx", "", "10%", "", "", "", "<a class=tt href=" & starget & "&sorden=E target=_self ><B><u>&nbsp;ESTADO&nbsp;</u></B></A>", "si")
		Celda = CELDA &  OFV.GenCelda("", "ttx", "", "10%", "", "", "", "<a class=tt href=" & starget & "&sorden=V target=_self ><B><u>&nbsp;VIGENCIA&nbsp;</u></B></A>", "si")
		Celda =  CELDA &  OFV.GenCelda("", "ttx", "", "10%", "", "", "", "<a class=tt href=" & starget & "&sorden=A target=_self ><B><u>&nbsp;AUTOR&nbsp;</u></B></A>", "si")
		
	End If
	
	
	
	
	
	Fila = fila & OFV.GenRow("", "", "", "", "20", "", "", Celda, "si")
	
	Response.write sHTML
	sHTML = ""
	Response.flush
	
	
	'''''''''''''''''''' FIN ENCABEZADOS ''''''''''''''''''''''''''''''''''''''
	
	nLinea = 0
	
	fi = CLng(CDate("01-01-1900"))
	PosReg = 0
	
	xTrColor = "#FFFFFF"
	
	On Error Resume Next
	If Not rs.EOF Then
		rs.Move (CantRegMover)
		Do While (Not rs.EOF) And (PosReg < CantRegMostrar)
			dEmision  = "..."
			dVigencia = "..."
			sArchivo = "Iconos/" & rs(11)
			
			FPRODUC="..."
			If Not IsNull(rs(1)) Then FPRODUC = rs(1)
			
			
			If Not IsNull(rs(0)) Then sNOM = rs(0)
			If Not IsNull(rs(2)) Then stipo = rs(2)
			If Not IsNull(rs(3)) Then CODPRO = rs(3)
			If Not IsNull(rs(4)) Then FESTAD = rs(4)
			If Not IsNull(rs(4)) Then FAUTOR= rs(6)
			If Not IsNull(rs(5)) Then dVigencia= fechatostr(CLng(rs(5)) + fi)
			
			
			dEmi = InStr(dEmision, "1900")
			If dEmi > 0 Then dEmision = "..."
			
			dpublic = "..."
			
			dVign = InStr(dVigencia, "1900")
			If dVign > 0 Then dVigencia = "..."
			
			If smenu <> "PORESTADO" Then
				
			Else
				sAutor   = "&nbsp;"
				sAprobo  = "&nbsp;"
				sElimino = "&nbsp;"
				If Not IsNull(rs(12)) Then sAutor   = rs(12)
				If Not IsNull(rs(13)) Then sAprobo  = rs(13)
				If Not IsNull(rs(14)) Then sElimino = rs(14)
			End If
			
			
			'
			If Not IsNull(rs(9)) Then nIden = rs(9)
			If Not IsNull(rs(10)) Then sobjeto = rs(10)
			
			
			If Session("USRES") <> "B" Then
				sImagen = OFV.GenImage("", "", "", "", sArchivo, "0", "")
			Else
				sImagen = "<font face=wingdings size=4 color=#993333>2</font>"
			End If
			
			'sImagen = "<font face='webdings' size=3 color='#993333'>4</font>"
			sImagen = "<i class='las la-file-alt' style=margin-right:10px;color:" & Session("CMFPNTIT") & " ></i>"

			sActcon = "Consultar.asp?Codigo=" & rs("codobjeto") & "&Tipo=" & rs("tipobjeto")
			sActcon = sActcon & "&Perfil=" & rs("perfil") & "&Datos=" & OFV.convertircar(sMenuMas, "&", "*")
			sActcon = sActcon & "&Volver=" & sVolver
			
			sHref = "&nbsp;"
			sFuncion = " onclick=" & Chr(34) & " AbrirCons('" & sActcon & "'); return false;" & Chr(34)
			
			sLink = OFV.GenLink("", "ut", sHref, "", "", sFuncion,sImagen & EsqnomPx(rs("tipobjeto"), rs("codobjeto"), ofv), "si")
			
			Celda = OFV.GenCelda("", "ut", "", "40%", "", "", "", sLink, "si")
			
			smasc = "&nbsp;"
			
			If session("APLMDCLI")="BC" Then
				sLink = OFV.GenLink("", "nt", sHref, "", "", sFuncion, "<b>" & rs("codiden") & "&nbsp;-&nbsp;</b>" & sobjeto, "si")
			Else
				sLink = OFV.GenLink("", "nt", sHref, "", "", sFuncion, EsqnomPx(rs("tipobjeto"), rs("codobjeto"), ofv), "si")
			End If
			sTipoDsc = stipo
			' sTipoDsc = "<a class=nt title='Tipo:  " & stipo & "' href=" & starget & "&sorden=" & sorden & "&sfiltro=" & Replace(stipo, " ", "_") & " target=_self ><font color='" & Session("CMFPNIPFT") & "' >" & sTipoDsc & "</font></A>"
			
			' 'Celda = Celda & OFV.GenCelda("", "ut ", "", "40%", "", "", "", sLink, "si")
			'Celda = Celda & OFV.GenCelda("", "ut ", "", "10%", "", "", "","" &  , "si")
			Celda = Celda & OFV.GenCelda("", "ut ", "", "10%", "", "", "","" &  sTipoDsc  , "si")
			Celda = Celda & OFV.GenCelda("", "ut ", "", "10%", "", "", "","" &  FPRODUC , "si")
			Celda = Celda & OFV.GenCelda("", "ut ", "", "10%", "", "", "","" &  CODPRO, "si")
			Celda = Celda & OFV.GenCelda("", "ut ", "", "10%", "", "", "","" &  FESTAD, "si")
			Celda = Celda & OFV.GenCelda("", "ut ", "", "10%", "", "", "","" &   dVigencia , "si")
			Celda = Celda & OFV.GenCelda("", "ut ", "", "10%", "", "", "","" &  FAUTOR, "si")
			
			Fila = Fila & OFV.GenRow("", "aaa bgcolor=" & xTrColor & " ", "", "", "22", "", "", Celda, "si")
			
			
			If xTrColor = "#FFFFFF" Then
				xTrColor = "#F0F0F0"
			Else
				xTrColor = "#FFFFFF"
			End If
			
			PosReg = PosReg + 1
			rs.movenext
			nLinea = nLinea + 1
			Response.write sHTML
			sHTML = ""
			If nLinea > 20 Then
				nLinea = 0
				Response.flush
			End If
			
		Loop
		
	Else
		
		'XFRASE = "<BR><BR>&nbsp;<B><font color=#993333 SIZE=3 FACE='WEBDINGS' >i</font></B>"
		XFRASE = "<BR><BR><i class='las la-info-circle' style=color:#993333;font-size:10pt; ></i>"

		XFRASE = XFRASE &  "&nbsp;NO SE ENCONTRARON DOCUMENTOS PARA LA BUSQUEDA SOLICITADA<BR><BR>"
		
		Celda = OFV.GenCelda("", "ut colspan=2 ", "", "", "", "", "", XFRASE, "si")
		
		
		
		
		Fila = Fila & OFV.GenRow("", "aaa bgcolor=" & xTrColor & " ", "", "", "22", "", "", Celda, "si")
		
		
	End If
	
	
	sHTML = sHTML & OFV.GenTabla("", "aaa", "", "100%", "", "0", "", "-1", "0", "", "", Fila, "si")
	'response.write Addqry
	
	If SMENU = "PORREVISION" Then
		sAccion = "COFECREVISION.asp?Menu=" & sMenu & "&Nombre=" & sNombre
	Else
		sAccion = "SCONSUL.asp?Menu=" & sMenu & "&Nombre=" & sNombre
	End If
	
	Celda   = ofv.BotonVolver(sAccion)
	
	
	sHTML = sHTML & OFV.botoneraESSX("c", rs, CantRegMover, CantRegMostrar, "no", "no", PosReg, Addqry, "", "no", "", "", Celda)
	
	
	sHTML = sHTML  & "</body></html>"
	
	If smenu = "PORSECTOR" Then
		Session("sperfil") = Replace(SPerfil, " ", "_")
	Else
		Session("sperfil") = ""
	End If
	'Call OFV.cerrarconsulta(rscont)
	Call OFV.cerrarconsulta(rs)
	Call OFV.cerrarconn(cn)
	
	
	Generar = sHTML
	
	'''''''''''''''''''
	
	sHTML = Generar
	
	'' sHTML = oClc.Generar
	
	
	
End If

response.write sHTML
response.flush

Function vertexto(nwtexto)
	
	nwtextoF = nwtexto & ""
	
	If  nwtextoF <> "" Then
		
		nwtextoA = Split(nwtextoF," ")
		
		nwsep    = ""
		nwtextoF = ""
		
		For iii2 = 0 To UBound(nwtextoA)
			
			If Len(nwtextoA(iii2)) > 2 Then
				
				nwtextoF =  nwtextoF & nwsep & nwtextoA(iii2)
				
				nwsep = "%"
			End If
		Next
		
		
	End If
	
	
	
	vertexto = nwtextoF
	
End Function

Function fechaConvert(wkfecha)
	Nfecha=CDate(wkfecha)
	Nfecha=DateSerial(Year(wkfecha),Month(wkfecha),Day(wkfecha))
	fechaConvert=Nfecha
End Function

Function fechatostr(wkfecha)
	fechatostr=CDate(wkfecha)
	fechatostr=Day(wkfecha) & "/" & Month(wkfecha) & "/" & Year(wkfecha)
	
End Function



%>