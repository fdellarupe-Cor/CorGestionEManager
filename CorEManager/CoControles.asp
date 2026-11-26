<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/FMTCONTROLES.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0

	nCodigo  = request.querystring("Codobjeto")
	sTipo    = request.querystring("Tipobjeto")
	sSectorx = request.querystring("Sector")

	sCambios = request.querystring("sCambios")


	If  sCambios = "" Then sCambios = "NO"

	sNIV1   = request.querystring("NIV1")
	sNIV2   = request.querystring("NIV2")
	sNIV3   = request.querystring("NIV3")
	sNIV4   = request.querystring("NIV4")
	sNIV5   = request.querystring("NIV5")

	nCantRegAMover = request.querystring("CantRegAMover")


	If nCantRegAMover = "x" Then
		xsproced = request.form("selcom")
	ElseIf nCantRegAMover = "A" Then
		xsproced = request.querystring("control")
	End If

	'----------------------------------------------
	'''''''''''''''''''''''''''''''''''''''''''''''
	'' ACTIVACION PUNTERO DEL LINK
	''
	''''''''''''''''''''''''''''''''''''''''''''''''

	If SNIV1 = "7" And SNIV2 <> "0" Then

		xsproced =  SNIV2

		nCantRegAMover = "A"


	Else

		SPLFUNC = ""

	End If

	SPLFUNC = SPLFUNC & "<script language='javascript' > var NN1 = '" & SNIV1 & "'; var NN2 = '" & SNIV2 & "';var NN3 = '" & SNIV3 & "';</script>"



	sMenux = request.querystring("Menu")
	''response.write sMenux & " por sMenux " & "<br>"
	If sMenux = "PORSECTOR" Then
		sSelecx = " and descri= '" & ofv.convertircar(ofv.convertircar(sSectorx,"@"," "),"_"," ") & "' "
	Else
		sSelecx = ""
	End If
	If nCantRegAMover = "x" Then
		Strsql = "SELECT * from cscontrolsel "
		Strsql = Strsql & "Where codobjeto = " & nCodigo
		Strsql = Strsql & " and objid =0" & xsproced & sSelecx
		'response.write " por aca -70- " & "<br>"
	Else
		'''response.write nCantRegAMover & " nCantRegAMover " & "<br>"
		If nCantRegAMover = "A" Then
			Strsql = "SELECT * from cscontrolsel "
			Strsql = Strsql & "Where codobjeto = " & nCodigo
			Strsql = Strsql & " and objid =0" & xsproced & sSelecx
			'response.write " por aca -76- " & "<br>"
			'response.write Strsql & " strsql " & "<br>"
		Else
			Strsql = "SELECT * from cscontrolsel "
			Strsql = Strsql & "Where codobjeto = " & nCodigo & sSelecx
			'Strsql = Strsql & " and seldes is NOT NULL and descri is NOT NULL "
			CantRegAMover = "0" & request.querystring("CantRegAMover")
			CantRegAMover = cdbl(CantRegAMover)
			nCantRegAMover = CantRegAMover
			'response.write " por aca -83- " & "<br>"
			'response.write Strsql & " strsql " & "<br>"
			'response.write CantRegAMover & " CantRegAMover " & "<br>"

		End If
	End If
	'response.write Strsql & " strsql " & "<br>"

	CantRegAMover2  = cdbl(request.querystring("CantRegAMover2"))
	CantRegAMostrar = 1
	If BrowserName() = "IE" Or BrowserName() = "CHROME" Then
		CantRegAMostrar2 = 7
	Else
		CantRegAMostrar2 = 5
	End If

	vuelta  = ofv.vueltaaspEx("../",CantRegAMostrar,nCantRegAMover) & "_control=" & xsproced
	vuelta2 = ofv.vueltaaspEx("../",CantRegAMostrar2,CantRegAMover2)
	vueltaC = ofv.vueltaaspExC("../",CantRegAMover,CantRegAMover2)





	sMenu = Session("menu")
	bModificar = session("Modificar")

	scontdesc = "NO"

	sHTML = ofv.FormHeader(session("FormN"))
	'  shtml= shtml &  "<script src='incfile/searchhi.js' type='text/javascript' language='JavaScript'></script>" & chr(13)
	shtml = shtml & CorPnlFnc("N","")

	'----------------------------------------------


	Fila = ""

	''''''''''''''''''''''''''''''''''

	Set cn = ofv.conectar(ofv.strconn2)
    if Cobjun=1 and Session("EMPMUN") = "S" then
       sql="exec SpSectoresCorp " & nCodigo &",'" & sTipo & "'," & session("un")
'     response.write sql
      Set rsSecCorp = ofv.crearconsultaEx(sql,cn,1,parametros)

	 if not rsSecCorp.eof then

		 ASecCorp =rsSecCorp.getrows
	  end if
	Call ofv.cerrarconsulta(rsSecCorp)
     end if


	strsqlx = "Select Version,estado,objetos.objun From Objprty "
	strsqlx = strsqlx & " inner join objetos on(objetos.codobjeto=Objprty.codobjeto) "
	strsqlx= strsqlx & "Where Objprty.codobjeto = " & nCodigo
	Set rs = ofv.crearconsultaEx(StrSqlx,cn,1,parametros)
	'response.write strsqlx & " strsqlx " & "<br>"
	bVersion = "NO"
	Cobjun=1

	If Not rs.eof Then
	    Cobjun=rs(2)
		If rs(0) = "V" Then bVersion = "SI"
		Nestado=rs(1)
	End If
	Call ofv.cerrarconsulta(rs)
   if Cobjun=1 and Session("EMPMUN") = "S" then
         	  sql="exec SpSectoresCorp " & nCodigo &",'" & sTipo & "'," & session("un")
         '     response.write sql
         	  Set rsSecCorp = ofv.crearconsultaEx(sql,cn,1,parametros)

                 if not rsSecCorp.eof then

                     ASecCorp =rsSecCorp.getrows
                  end if
                Call ofv.cerrarconsulta(rsSecCorp)
     end if
'response.write Cobjun
	''sHTML = sHTML & "<DIV ID='DCEX1' style='display:block;' >"

	' response.write strsql & "<br>" & chr(13)

	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

	do while Not rs.eof
		nID       = rs(0)
		scontdesc = rs(4)
	if isarray(ASecCorp) and Cobjun=1 and Session("EMPMUN") = "S" then
		SECD  = rs(1)
		SECD = FindInArray(ASecCorp,6,rs("eco"),1,SECD)
        bDescri=SECD & "-" & rs("perfil") & "-" & rs("frecuencia")
	else
		bDescri = rs(1)
    end if

		'bDescri   = rs(1)
		If len(scontdesc) = 0 Then scontdesc = "NO"


		''''''''''''''''''
		stexto1 = "Seleccionar&nbsp;Control"
		stexto1 = "&nbsp;&nbsp;" & stexto1

		Celdat  = stexto1

		sAccion = "javascript:void(0);"
		sOnClik = " onclick=""document.getElementById('TIND2').style.display='';document.getElementById('TIND1').style.display='none';"" "


		stexto1 = "<A ID='ASCO' NAME='ASCO' class=utg href='" & sAccion & "' " & sOnClik & " "
		stexto1 = stexto1 & "   >" & Celdat & "</a>"

		Celdat  = "<td class=uts style='cursor:pointer;' " & sOnClik & " >" & sfidk0 & "</td>"
		Celdat  = Celdat & "<td class=utg  >" & stexto1 & "</td>"
		'   Filat   = filat & ofv.GenRow("","","","","","","",Celdat,"si")

		stexto1 = "Controles"

		Celdat  = "<td class=uts >" & sfidk0 & "</td>"
		Celdat  = Celdat & "<td class=utg  >" & stexto1 & "</td>"
		'   Filatt   = filatt & ofv.GenRow("","","","","","","",Celdat,"si")



		''''''''''''''''''''''

		Strsql = "SELECT * from cscontrolsel "
		Strsql = Strsql & "Where codobjeto = " & nCodigo & sSelecx
		'Strsql = Strsql & " and seldes is NOT NULL and descri is NOT NULL "
		'  response.write strsql & " strsql " & "<br>"
		Set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		'contar = 0
		Do While Not rsx.eof

			' contar = contar + 1
			' response.write contar & "contar " & "<br>"

			If clng(rsx(0))   = clng(rs(0).value) Then
				sfidkD  = sfidkC

				if isarray(ASecCorp) and Cobjun=1 and Session("EMPMUN") = "S" then
				SECD  = rsx(2)
	            SECD = FindInArray(ASecCorp,6,rsx("eco"),1,SECD)

				stexto1=SECD & "-" & rsx("perfil") & "-" & rsx("frecuencia")
				else
				stexto1 = rsx(1)
				end if
				stextoact = stexto1
				'response.write stexto1 & " stexto1 por si " & "<br>"
			Else
				sfidkD  = sfidkA
			'	stexto1 = rsx(1)
				if isarray(ASecCorp) and Cobjun=1 and Session("EMPMUN") = "S" then
				 SECD  = rsx(2)
			     SECD = FindInArray(ASecCorp,6,rsx("eco"),1,SECD)
				 stexto1=SECD & "-" & rsx("perfil") & "-" & rsx("frecuencia")
				else
				 stexto1 = rsx(1)
				end if

				'response.write stexto1 & " stexto1 por else " & "<br>"
				sAccion  = "CoControles.asp?CantRegAMover=A&Codobjeto=" & nCodigo & "&Menu=" & sMenux
				sAccion  = sAccion & "&Sector=" & sSectorx & "&Tipobjeto=" & sTipo & "&control=" & rsx(0)
				sOnClik1 = ""
				sOnClik2 = " onclick=""document.getElementById('ASC" & rsx(0) & "').click();"" "

				Celdat  = stexto1

				stexto1 = "<A ID='ASC" & rsx(0) & "' NAME='ASC" & rsx(0) & "' class=ut href='" & sAccion & "' " & sOnClik1 & " "
				stexto1 = stexto1 & "   >" & Celdat & "</a>"

				Celdat  = "<td class=uts style='cursor:pointer;' " & sOnClik2 & " >" & sfidkD & "</td>"
				Celdat  = Celdat & "<td class=ut  >" & stexto1 & "</td>"

				''response.write Celdat & " Celdat " & "<br>"

				Filat  = filat & ofv.GenRow("","","","","","","",Celdat,"si")

				''response.write Filat & " Filat " & "<br>"

			End If

			rsx.movenext
		Loop

		''''''''''''''''''''''''''''''''


	'	Call ofv.cerrarconsulta(rs)

	'Else

	'	nID = "no"

	'End If

	stitu = bDescri

		If scambios = "SI" Then
			stitu = stitu & "&nbsp;(Cambios)"
		End If

     if stitu & "" <> "" then


		Celda = ofv.GenCelda("","ttt","","60%","","","",stitu ,"si")
		Fila = fila & ofv.GenRow("","","","","25","","",Celda,"si")


		' Referencia
		If scontdesc <> "NO" Then
			Celda = ofv.GenCelda("","utg style='text-align:justify;' ","","100%","","","",replace(scontdesc,chr(13),"<br>"),"si")
			Fila = fila & ofv.GenRow("","","","","25","","",Celda,"si")
	     End If
       Celda = ofv.GenCelda("","utg style='text-align:left;' ","","100%","","","",GENFMTCONTROLES(nCodigo,STIPO),"si")
	   Fila = fila & ofv.GenRow("","","","","25","","",Celda,"si")
	 END IF
	 rs.movenext
	loop
	'' response.write vuelta & "<br>" & chr(13)

	'Fila = ""




	'''''''''''''''''''''''''''''''''''''''''''''''''


	''sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' ><span id=111>"
     sHTML = sHTML &  "<div id='mt2'  style='display:none;height:350px;overflow:auto;" & styl & "' >"

 sHTML = sHTML & "<div id=zglos class='norma-card' >"
	sHTML = sHTML &  ofv.GenTabla("","","","100%","","0","","4","0","","",Fila,"si") & "<br>"


	'sHTML = sHTML & GENFMTCONTROLES(nCodigo,STIPO)


	sHTML = sHTML & "</span></div>"

	sHTML = sHTML & "</div>"


	Celda = ofv.GenCelda("","ut","","100%","","","","&nbsp;","si")
	Fila  = ofv.GenRow("","","","","","","",Celda,"si")
	'sHTML = sHTML &  ofv.GenTabla("TCEX1","aaa onresize='DCEXRCLC();' ","","100%","","0","","4","0","","",Fila,"si")


	''''''' selector del glosario

	GLhtml = VerGlosario(shtml)

	sHTML = GLhtml

	''' call ofv.cerrarconn(cn)

	''''''''''''''''''''''''''''''

	'Estructura del Indice

	sHTML = sHTML & CorPnlIndx("N","")


	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	'Botonera
	Celt1 = ""
	''' Celt1 = ofv.GenBoton(17, "indBot", "BT", "N", " onclick='clIndice();' ", "S", "A", "N", "indFrm", "")
	Celt2 = ofv.GenBoton(18, "glosBot", "BT", "N", " onclick='versup();' ", "S", "A", "N", "glosFrm", "")


	'cambios
	Listar = ""
	'response.write origen & "222"
	If origen <> "ED" Then

		If bVersion = "SI" Then
			If SESSION("SHWMOD") <> "N" Then

				If  sCambios = "NO" Then
					xcambios = "SI"
				Else
					xcambios = "NO"
				End If

				' xxxx    = CorGenBotonXp(3, "NO", "&#0113;","NO", "NO", "Lista de cambios", "NO", "Cambios")
				sAccion = "CoControles.asp?CantRegAMover=A&Codobjeto=" & nCodigo & "&Menu=" & sMenux
				sAccion = sAccion & "&Sector=" & sSectorx & "&Tipobjeto=" & sTipo & "&control=" & nid & "&scambios=" & xcambios
				'Listar  = ofv.GenBoton(3, "cambBot", "BT", "S", "", "S", "A", "N", "cambFrm", sAccion & "")
				Listar  = genbotoncambios(sAccion)

			ElseIf SESSION("VIEWMOD") <> "N" Then

				' xxxx    = CorGenBotonXp(3, "NO", "&#0113;","NO", "NO", "Lista de cambios", "NO", "Cambios")
				Listar  = genbotoncambios("") 'ofv.GenBoton(3, "cambBot", "BT", "N", " onclick='AbrirCambios();' ", "N", "S", "N", "", "")

				sasp = "CoGenCambios.asp?Codobjeto=" & nCodigo & "&Tipobjeto=" & sTipo
				sHTML = sHTML & ConsCambios(sasp)


			End If




		End If
	End If




	Agregar = Celt2
	Buscar  = Celt1
	Celda   = Listar

	sHTML   = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)




	If nID <> "no" Then Call ofv.cerrarconsulta(rs2)
	Call ofv.cerrarconn(cn)

	sHTML = sHTML & GTCOBS()

	sHTML = sHTML & "</body></html>"
End If

response.write sHTML

%>