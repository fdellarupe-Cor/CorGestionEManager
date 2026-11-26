<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=0

	CantRegMostrarx = 8
	CantRegMoverx = CDbl(request.querystring("CantRegMover"))
	vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)

	Set cn = ofv.conectar(ofv.strconn2)

	sHTML = ofv.FormHeader(session("FormN"))

	base=request.querystring("BASE")
	If base <>"" Then
		SESSION("BASE")=base
	End If

	'response.write "STRING: " & request.querystring & "<br>" & "FORM: " & request.form
	XXXX = NEWEMPRE("NO")

	session("ASP")="SUBCONSULTA.ASP"

	sEMPRE = request.querystring("EMPRE")
	'response.write "sEMPRE: " & sEMPRE
	If SEMPRE = ""  And SESSION("UN") = "" Then

		SEMPRE = ATREMP(0, 0)
		SESSION("UN") = SEMPRE & ""
		If (SESSION("UN") & "") = "" Then SESSION("UN") = 1
	ElseIf sEMPRE <> "" Then
		XXXX = NEWEMPRE(sEMPRE)
		SESSION("UN")=SEMPRE
	Else
		XXXX = NEWEMPRE(SESSION("UN"))
		sempre=SESSION("UN")
	End If
	'CeldaH = ofv.GenCelda("","tt","","","","","",selecEmpre(),"si")
	'FilaH = ofv.genrow("","","","","","","",CeldaH,"si")
	TablaH = ofv.GenTabla("","","","20%","","0","0","0","25","","",selecEmpre(),"si")


	shtml = shtml & tablah '& "<br>"

	'Encabezados
	If Session("EMPMUN") = "S" Then
		Fila = ofv.Encabezados("Descripcion","Orden","Alcance","Modo","Redir.","Subtemas","Prop.")
	Else
		Fila = ofv.Encabezados("Descripcion","Orden","Alcance","Modo","Redir.","Subtemas")
	End If

	ofv.ObtenerAtributos Session("Form"),""
	bAgregar = ofv.AAgregar
	bModificar = ofv.AModificar
	bSuprimir = ofv.AEliminar
	bConsultar = ofv.AConsultar

	If Session("CorApli") = "COREM20" Then
		wkaplic = "EM"
	ElseIf Session("CorApli") = "COREP20" Then
		wkaplic = "PL"
	Else
		wkaplic = "OR"
	End If






	strsql = "select distinct a.codtema,a.descri,a.inda,a.pstipo,a.psurl,a.aplic from ("
	strsql = strsql & " Select  isnull(cast (b.atrmverst as int),1) as st,b.atrmcodes ,a.* From Tema  a "
	If session("UN") <> 1 Then
		strsql = strsql & " left join  atrmtemast  b on (a.codtema=b.atrmcodst and b.atrmtipost='T'  and b.atrmcodes=" & session("UN") & " ) "
	Else
		strsql = strsql & " left join  atrmtemast  b on (a.codtema=b.atrmcodst and b.atrmtipost='T') "
	End If

	If SESSION("BASE")="K" Then
		strsql = strsql & " where a.aplic = 'EM' or a.aplic = 'GLB' ) a where a.st > 0 AND A.PSTIPO = 'K' order by a.descri "
	Else
		strsql = strsql & " where a.aplic = 'EM' or a.aplic = 'GLB' ) a where a.st > 0 AND A.PSTIPO <> 'K' order by a.descri "
	End If

	'	response.write strsql

	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If Not rs.EOF Then   rs.move(CantRegMoverx)
	posI = 0
	Do While (Not rs.EOF) And (posi< CantRegMostrarx)
		Formulario = ""
		sForm = "form" & posi

		'Eliminar
		If bSuprimir Then
			sHref = "asp/EliminarTema.asp?ID=" & ofv.convertircar(rs(0)," ","_")
			sHref = sHref & "&volver=" & vuelta & "&Eliminar=NSNC"
			Celda = ofv.BotonEliminar(rs(1),sHref)
		Else
			Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
		End If

		wkstylo = ""

		If rs(5) = "GLB" Then wkstylo = " style='border-radius:20px;border-style:solid;border-width:1;border-color:" & Session("CMFPNBOT") & ";background-color:#eeeeee;' "


		'Descripcion
		sInput = ofv.GenerarInput(rs(1).name,rs(1),"text",sForm,"150","dt","")
		Celda = Celda & ofv.GenCelda("","ut " & wkstylo,"","","","","",sInput,"si")

		'orden
		sInput = ofv.GenerarInput(rs(2).name,rs(2),"text",sForm,"5","dtn","")
		Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

		'APLICACION
		scomK = wkaplic & Chr(9) & "GLB"
		scomD = "Aplicacion" & Chr(9) & "GLOBAL"


		If bModificar Then
			sInput = ofv.gencomboFX(sForm,rs(5).name,rs(5),scomK,scomD," ")
		Else
			sInput = ofv.GenerarInput("",rs(5),"readonly","","10","dt","")
		End If





		Celda = Celda & ofv.GenCelda("","" ,"","","","","",sInput,"si")



		'tipo
		Strsql = "Select * From Prodserv where codtema = " & rs(0)
		Set rsx = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		If Not rsx.eof Then

			If rsx("pstipo") ="K" Then
				carpeta="K"
			Else
				carpeta="Carpeta"
			End If

			sInput = ofv.GenerarInput("",carpeta,"readonly","","20","dt","")

		Else

			scomK = "C" & Chr(9) & "R"
			scomD = "Carpeta" & Chr(9) & "Redireccion"


			If bModificar And RS("PSTIPO") <> "K" Then
				sInput = ofv.gencomboFX(sForm,rs(3).name,rs(3),scomK,scomD," ")
			Else
				sInput = ofv.GenerarInput("",rs(3),"readonly","","10","dt","")
			End If



		End If

		Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")


		If rs(3) = "R" Then

			'Descripcion
			sInput = ofv.GenerarInput(rs(4).name,rs(4),"text",sForm,"20","dt","")
			Celda = Celda & ofv.GenCelda("","","","","","","",sInput,"si")

			sInput = "----"
			Celda = Celda & ofv.GenCelda("","ut","center","","","","",sInput,"si")


		Else

			sInput = "----"
			Celda = Celda & ofv.GenCelda("","ut","center","","","","",sInput,"si")


			'sImagenP ="&#0232;"
			'sFuncAnc = " onmouseover=""this.className='ft1o';"" onmouseout=""this.className='ft1n';"" " & sFuncAnc &
			sinput = "<a href='subtemas.asp?Retorno=" & CantRegMoverx & "&tema=" & rs(0)
			sinput = sinput  & "' title='Propiedades de " & rs(1) & "' ><i class='las la-arrow-right'></i></a>"
			Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")


		End If


		If Session("EMPMUN") = "S" Then
			sImagenP = "&#0232;"
			sFuncAnc = " onmouseover=""this.className='ft2o';"" onmouseout=""this.className='ft2n';"" "
			sinput = "<a class='ft2n' " & sFuncAnc & " href='ATRMTEMAST.asp?CantRegAMover=" & CantRegMoverx & "&OPCION=" & rs(0)
			sinput = sinput  & "&T=T&NOMBRE=" & Replace(RS(1)," ","_")
			sinput = sinput  & "' title='Propiedades de " & rs(1) & "' >" & sImagenP & "</a>"
			Celda = Celda & ofv.gencelda("","ft","center","","15%","","",sInput,"si")

		End If

		sAccion = ""
		If bModificar Then
			sAccion = "asp/actualizar.asp?fname=" & ofv.convertircar(rs(0)," ","*")
			sAccion = sAccion & "&tabla=Tema&volver=" & vuelta
		End If
		Formulario = Formulario & ofv.GenForm(sForm,"",sAccion,"","post","",Celda,"si")


		Fila = Fila & ofv.genrow("","","","","","","",Formulario,"si")

		posi=posi+1
		rs.Movenext
	Loop

  ftab = ofv.GenTabla("","","","","","0","","0","0","","",Fila,"si")

    Celda = ofv.GenCelda("","","center","","","","",ftab,"si")
    FilaS = filas & ofv.GenRow("","","","","","","",Celda,"si")

    sHTML = sHTML & ofv.GenTabla("","aaa ","","","","0","","0","0","","",filas,"si")

	Agregar = ""
	If bAgregar Then
		sAccion = "AgregarTemaSUBTEMA.asp?Regisvuelta=" & CantRegMoverx & "&Opcion=si&tabla=TEMA"
		Agregar = ofv.BotonAgregar(sAccion)
	End If



	strsql = "select distinct a.descri,a.codtema,a.inda,a.pstipo,a.psurl,a.aplic from ("
	strsql = strsql & " Select  isnull(cast (b.atrmverst as int),1) as st,b.atrmcodes ,a.* From Tema  a "
	If session("UN") <> 1 Then
		strsql = strsql & " left join  atrmtemast  b on (a.codtema=b.atrmcodst and b.atrmtipost='T'  and b.atrmcodes=" & session("UN") & " ) "
	Else
		strsql = strsql & " left join  atrmtemast  b on (a.codtema=b.atrmcodst and b.atrmtipost='T') "
	End If
	If SESSION("BASE")="K" Then
		strsql = strsql & " where a.aplic = 'EM' or a.aplic = 'GLB' ) a where a.st > 0 AND A.PSTIPO = 'K' order by a.descri "
	Else
		strsql = strsql & " where a.aplic = 'EM' or a.aplic = 'GLB' ) a where a.st > 0 AND A.PSTIPO  <>'K' order by a.descri "
	End If
	'response.write strsql
	'str1 = "Select descri From Tema where aplic = '" & wkaplic & "' or aplic = 'GLB' Order By descri"
	'str2 = "Select descri From Tema where aplic = '" & wkaplic & "' or aplic = 'GLB' Order By descri"

	Buscar = ofv.Combobuscar(strsql,strsql,cn,cantregmostrarx,CantRegMoverx,ofv.vueltaasp2(""))

	sHTML = sHTML & ofv.botoneraESSX("2",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","",Agregar,Buscar,"")

	Call ofv.cerrarconsulta(rs)
	Call ofv.cerrarconn(cn)

	sHTML = sHTML & "</body></html>"
End If

Function selecEmpre()

	sImagenx = "<B><FONT FACE=CorFlatIcon COLOR=#aa0000 SIZE=4>x</FONT></B>"

	If  Session("EMPMUN") = "S" Then

		Fila    = ""

		SINPUT3 = "&nbsp;<A class=tt style='color:#000000;' title='SELECCIONE EMPRESA' HREF='javascript:void(0);' onclick='shwDem();' >EMPRESA:&nbsp;&nbsp;</a>"

		Fila    = ""
		Celda   = ""
		sHref   = ""

		sImage  = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(0,5),"0","")
		sLink   = ofv.GenLink("","ttt",sHref,"ftemas","","",sImage,"si")
		Celda   = Celda & ofv.GenCelda("","tt ","","8%","","","",SINPUT3,"si")
		Celda   = Celda & ofv.GenCelda("","ttt ","center","5%","","","",sLink,"si")
		Celda   = Celda & ofv.GenCelda("","ttt  ","","","","","","&nbsp;<BIG>" & ATREMP(0,4) & "</BIG>","si")

		Fila    = fila & ofv.GenRow("","","top","","","","",Celda,"si")

		Tabla   = ofv.GenTabla("","aaa style='"  & Session("CBRDABL") & "' ","","","","0","","-1","-2","","",Fila,"si")





		shtml=shtml & "<script language=javascript>" & Chr(13)
		shtml=shtml & "function shwDem(){" & Chr(13)
		shtml=shtml & "document.all.DEM1.style.display='block';" & Chr(13)
		shtml=shtml & "}" & Chr(13)
		shtml=shtml & "</script>" & Chr(13)
		'----------------------------------------------
		sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & Chr(34)
		sinput2 = sinput2  & "underline" & Chr(34) & ";' onmouseout='this.style.textDecoration=" & Chr(34)
		sinput2 = sinput2  & "none" & Chr(34) & ";' href='javascript:void(0);' ONCLICK=" & Chr(34) & "document.all.DEM1.style.display='none';" & Chr(34)
		sinput2 = sinput2  & " TITLE='OCULTAR' ><b>" & sImagenx & "</b></a>"
		TABLA = Tabla & "<DIV id='DEM1' STYLE='display:none;z-index:1;margin-top:1;margin-left:1;background-color:#ffffff;position:absolute;left:1px;top:240px;left:20px;height:250px;width:350px;border-style:inset;border-color:#f7aa22;border-width:2;overflow:auto;'>"


		TABLAX = TablaX & OFV.gentabla("", "aaa style='filter:shadow(color=#dddddd,direction=125);' ", "", "100%", "", "0", "", "-2", "-1", "", "", "", "no")
		TABLAX = TablaX & "<tr><td  "
		TABLAX = TablaX & "class=vT align=left colspan=3 STYLE='border-style:outset;border-width:2;background-color:#eeeeee;' >" & sinput2 & "&nbsp;SELECCIONE EMPRESA</td></tr>"
		TABLAX = TablaX & "<tr><td colspan=3 "
		TABLAX = TablaX & " >&nbsp;</td></tr>"

		For IAEMP=1 To TAEMP

			'''  RESPONSE.WRITE  ATREMP(IAEMP,2)

			If ATREMP(IAEMP,2) = 1 Then

				If ATREMP(IAEMP,0) = ATREMP(0,0) Then
					TABLAX = TablaX & "<tr BGCOLOR=#EEEEEE >"
				Else
					TABLAX = TablaX & "<tr>"
				End If

				TABLAX = TablaX & "<td width=5 "
				If ATREMP(IAEMP,0) = ATREMP(0,0) Then
					TABLAX = TablaX & "class=vt  >" & sImagen4si & "</td>"
				Else
					TABLAX = TablaX & "class=vt  >" & sImagenI3 & "</td>"
				End If

				sImage = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(IAEMP,5),"0","")
				TABLAX = TablaX & "<td width=5 "
				TABLAX = TablaX & "class=vt  >" & sImage & "</td>"

				TABLAX = TablaX & "<td class=vt  >"

				'sHref = "subconsulta.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0&EMPRE=" & ATREMP(IAEMP,0)
				sHref = "temassubtemas.asp?EMPRE=" & ATREMP(IAEMP,0)
				'''  sHref = "OBJtreeTema.asp?Tipo=" & sTipo & "&Nombre=" & LNombre & "&Menu=" & SESSION("MENU") & "&EMPRE=" & ATREMP(IAEMP,0)
				sLink = ofv.GenLink("","cc",sHref,"","","","<b>" & ATREMP(IAEMP,4) & "</b>","si")
				TABLAX = TablaX & sLink
				TABLAX = TablaX & "</td>"
				TABLAX = TablaX & "</tr>"
			End If

		Next

		TABLAX = TablaX & OFV.genFtabla()

		TABLA = Tabla & TABLAX

		TABLA = Tabla & "</DIV>"

		sHTML = sHTML & TABLA

	End If

End Function

Response.Write sHTML

%>