<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else

	response.expires = 0

	sModal = request.querystring("modal")
	sTipo = request.querystring("Tipo")
	sNombre = request.querystring("Nombre")
	sNombre = ofv.convertircar(sNombre,"_"," ")
	sMenu = request.querystring("Menu")
	base=request.querystring("BASE")

	If base <> "" Then
		SESSION("BASE")=base
	End If

	If sModal = "" Then sModal = "NORMAL"

	session("ModMain") = "ED"

	If Session("CorApli") = "COREM20" Then
		wkaplic = "EM"
	ElseIf Session("CorApli") = "COREP20" Then
		wkaplic = "PL"
	Else
		wkaplic = "OR"
	End If


	xapli = 1

	On Error Resume Next
	Set cn = ofv.conectar(ofv.strconn1)
	Strsql = "Select ID From CorAplic where codapli = '" & Trim(Session("CorApli")) & "' "
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If Not rs.eof Then
		xapli = rs(0)
	End If

	Strsql = "Select menu From menux where menuatr = '" & snombre & "' and aplicacion = " & xapli
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If Not rs.eof Then
		sNombre = rs(0)
	End If
	Call ofv.cerrarconsulta(rs)
	Call ofv.cerrarconn(cn)

	''sNombre = ofv.convertircar(sNombre," ","_")



	''sHTML = ofv.MENUheader("",FUNCION)

	''''''''''''''''''''''''''''

	session("xorigen") = "tema"

	XXXX = NEWEMPRE("NO")

	sEMPRE = request.querystring("EMPRE")

	If SEMPRE = "" Then
		SEMPRE = ATREMP(0,0)
	Else
		recarga = NEWEMPRE(SEMPRE)
	End If

	'' sTipo = request.querystring("Tipo")


	sUN = SEMPRE

	If sUN = "" Then sUN = 1
	session("UN") = sUN

	''  LNombre = request.querystring("Nombre")
	''  sNombre = request.querystring("Nombre")
	''  sNombre = ofv.convertircar(sNombre,"_"," ")

	LNombre = ofv.convertircar(sNombre," ","_")

	''  if request.querystring("Menu") <> "" then Session("Menu") = request.querystring("Menu")

	If sMenu <> "" Then Session("Menu") = sMenu

	sID = request.querystring("ID")
	If sNombre = "" Then sNombre = UCase(sTipo)
	sitem = ofv.convertircar(request.querystring("sitem"),"_"," ")
	stitem = ofv.convertircar(request.querystring("stitem"),"_"," ")
	sticono = request.querystring("sticono")
	stnivel = request.querystring("stnivel")

	''response.write sTipo & " sTipo " & "<br>"

	ssTipo = sTipo
	session("sEstadon") = "1"
	session("sNivelCero") = sTipo
	If sTipo="Temas" Then
		session("sEstado") = "APROBADA"
		session("sEstadon") = "3"
	ElseIf sTipo="Desarrollo" Then
		session("sEstado") = "DESARROLLO"
		session("sEstadon") = "1"
	ElseIf sTipo="ParaAprobar" Then
		session("sEstado") = "PARA APROBAR"
		session("sEstadon") = "2"
	ElseIf sTipo="Eliminadas" Then
		session("sEstado") = "ELIMINADA"
		session("sEstadon") = "4"
	ElseIf sTipo="Consulta" Then
		session("sEstado") = "APROBADA"
		session("sEstadon") = "3"
	End If
	Set cn = ofv.conectar(ofv.strconn2)
	Strsql = "Select estado from estados where objid = " & session("sEstadon")
	'response.write strsql & "<br>"
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If Not rs.eof Then
		ssTipo = rs(0)
	End If
	Call ofv.cerrarconsulta(rs)




	'  sImagenI = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=2>1u</FONT></B>"
	'  sImagenI2 = "<B><FONT FACE=""corflaticon"" COLOR=#f7aa22 SIZE=3>1q</FONT></B>"
	'  sImagenI3 = "<B><FONT FACE=""corflaticon"" COLOR=#6699CC SIZE=2>1=</FONT></B>"
	'  sImagenEM = "<B><FONT FACE=CorFlaticon COLOR=#F7AA22 SIZE=3>&#48661;</FONT></B>"
	'  sImagen4si = "<B><FONT FACE=CorFlaticon COLOR=#993333 SIZE=3>&#57406;</FONT></B>"
	'  sImagenx = "<B><FONT FACE=CorFlaticon COLOR=#aa0000 SIZE=4>&#57367;</FONT></B>"


	sImagenI = "<B><FONT FACE=""CorFlaticon"" COLOR=#f7aa22 SIZE=2>u</FONT></B>"
	sImagenI2 = "<B><FONT FACE=""CorFlaticon"" COLOR=#f7aa22 SIZE=3>q</FONT></B>"
	sImagenI3 = "<B><FONT FACE=""CorFlaticon"" COLOR=#6699CC SIZE=2>=</FONT></B>"
	sImagenEM = "<B><FONT FACE=CorFlaticon COLOR=#F7AA22 SIZE=3>&#0252;</FONT></B>"
	sImagen4si = "<B><FONT FACE=CorFlaticon COLOR=#993333 SIZE=3>=</FONT></B>"
	sImagenx = "<B><FONT FACE=CorFlaticon COLOR=#aa0000 SIZE=4>x</FONT></B>"


	FUNCION = "  onload="  & Chr(34)

	If  recarga = "S" And sTipo <> "Consulta" Then
		FUNCION = FUNCION & "document.recarga.submit();"
	End If

	FUNCION = FUNCION & "rszdx4();"  & Chr(34)
	FUNCION = FUNCION & " onmousedown = " & Chr(34) & "Abre();" & Chr(34) & " "
	FUNCION = FUNCION & " onresize = " & Chr(34) & "rszdx4();" & Chr(34) & " "

	sHTML = ofv.MENUheader("",FUNCION)

	sHTML = sHTML & "<SCRIPT language=" & Chr(34) & "VBScript" & Chr(34) & ">" & Chr(13)

	sHTML = sHTML & " function Abred() " & Chr(13)
	sHTML = sHTML & " if window.event.button=2  then  " & Chr(13)
	sHTML = sHTML & "    if window.event.srcElement.tagName = ""A"" or window.event.srcElement.tagName = ""IMG""  or window.event.srcElement.tagName = ""FONT"" then  " & Chr(13)

	sHTML = sHTML & "       Editar(window.event.srcElement.ID) " & Chr(13)
	sHTML = sHTML & "    end if " & Chr(13)
	sHTML = sHTML & " end if " & Chr(13)
	sHTML = sHTML & " end function " & Chr(13) & Chr(13)

	sHTML = sHTML & " function Editard(Elemento) " & Chr(13)
	sHTML = sHTML & " if elemento = " & Chr(34) & Chr(34) & " Then " & Chr(13)
	sHTML = sHTML & " exit function " & Chr(13)
	sHTML = sHTML & " end if " & Chr(13)
	sHTML = sHTML & " dato = " & Chr(34) & "Edicion.asp?"

	If sModal = "PROCESOS" Then
		sHTML = sHTML & "MODAL=PROCESOS&"
	Else
		sHTML = sHTML & "MODAL=NORMAL&"
	End If


	sHTML = sHTML & "ID=" & Chr(34) & " & elemento " & Chr(13)
	sHTML = sHTML & " window.showModalDialog "
	sHTML = sHTML & "dato,window," & Chr(34) & "dialogWidth:280px;dialogHeight:320px;help:no;resizable:yes;dialogTop:250px;dialogLeft:300px;status:no" & Chr(34) & Chr(13)
	sHTML = sHTML & " end function " & Chr(13)

	sHTML = sHTML & "</SCRIPT> " & Chr(13)

	'------------------------------------------------------

	sHTML = sHTML & "<SCRIPT language='javascript' >" & Chr(13)

	sHTML = sHTML & "  	function Abre(){  " & Chr(13)
	sHTML = sHTML & " if (window.event.button==2){ " & Chr(13)
	sHTML = sHTML & "    if (window.event.srcElement.tagName == 'A' || window.event.srcElement.tagName == 'IMG'  || window.event.srcElement.tagName == 'FONT') {  " & Chr(13)

	sHTML = sHTML & "       Editar(window.event.srcElement.id); " & Chr(13)
	sHTML = sHTML & "    } " & Chr(13)
	sHTML = sHTML & " } " & Chr(13)
	sHTML = sHTML & " } " & Chr(13) & Chr(13)
	sHTML = sHTML & "</SCRIPT> " & Chr(13)


	sHTML = sHTML & "<SCRIPT language='javascript' >" & Chr(13)

	sHTML = sHTML & "  function Editar(Elemento) {" & Chr(13)
	sHTML = sHTML & " Elemento=Elemento + ''; " & Chr(13)
	sHTML = sHTML & " if (Elemento == ''){return; " & Chr(13)
	sHTML = sHTML & " } else {" & Chr(13)
	sHTML = sHTML & "  var dato = 'Edicion.asp?"

	If sModal = "PROCESOS" Then
		sHTML = sHTML & "MODAL=PROCESOS&"
	Else
		sHTML = sHTML & "MODAL=NORMAL&"
	End If


	sHTML = sHTML & "ID=' + Elemento ;" & Chr(13)
	sHTML = sHTML & " window.showModalDialog("
	sHTML = sHTML & "dato,window,'dialogWidth:280px;dialogHeight:320px;help:no;resizable:yes;dialogTop:250px;dialogLeft:300px;status:no');" & Chr(13)
	sHTML = sHTML & " }} " & Chr(13)

	sHTML = sHTML & "</SCRIPT> " & Chr(13)
	'------------------------------------------------------

	shtml = shtml & "<script language=javascript>" & Chr(13)
	sHTML = sHTML & "   function Editarx(Elemento) { " & Chr(13)
	sHTML = sHTML & " if (Elemento=='') { " & Chr(13)
	sHTML = sHTML & " } " & Chr(13)
	sHTML = sHTML & " else { " & Chr(13)
	sHTML = sHTML & " var dato = 'Edicion.asp?"

	If sModal = "PROCESOS" Then
		sHTML = sHTML & "MODAL=PROCESOS&"
	Else
		sHTML = sHTML & "MODAL=NORMAL&"
	End If


	sHTML = sHTML & "ID=' + Elemento; " & Chr(13)

	If BrowserName() = "IE" Then

		sHTML = sHTML & " if (window.showModalDialog) {window.showModalDialog("
		sHTML = sHTML & "dato,window,'dialogWidth:280px;dialogHeight:320px;help:no;resizable:yes;dialogTop:250px;dialogLeft:300px;status:no');} "  & Chr(13)
		sHTML = sHTML & " else { window.open(dato,window,'Width=280,Height=320,help:no;resizable:yes;dialogTop:250px;dialogLeft:300px;status:no');}"

	Else

		sHTML = sHTML & " window.open("
		sHTML = sHTML & "dato,window,'Width=280,Height=320,help:no;resizable:yes;dialogTop:250px;dialogLeft:300px;status:no'); "  & Chr(13)

	End If

	sHTML = sHTML & " } } " & Chr(13)

	sHTML = sHTML & "</SCRIPT> " & Chr(13)

	shtml=shtml & "<script language=javascript>" & Chr(13)
	shtml=shtml & "function mostrarmenuD(){" & Chr(13)
	shtml=shtml & "document.getElementById('Dx1').style.display='';" & Chr(13)
	shtml=shtml & "}" & Chr(13)
	shtml=shtml & "function rszdx4(){" & Chr(13)
	shtml=shtml & "document.getElementById('DX2').style.display='';" & Chr(13)
	shtml=shtml & "document.getElementById('DX4').style.display='';" & Chr(13)
	shtml=shtml & "var spread = document.getElementById('DX4').offsetTop + 50;" & Chr(13)
	shtml=shtml & "document.getElementById('DX4').style.left=document.getElementById('DX2').offsetLeft + document.getElementById('DX2').clientWidth + 1;" & Chr(13)
	shtml=shtml & "document.getElementById('DX2').style.height=document.body.clientHeight - spread;" & Chr(13)
	shtml=shtml & "document.getElementById('DX4').style.height=document.body.clientHeight - spread;" & Chr(13)
	'  shtml=shtml & "alert(document.body.clientHeight + ' ' + document.all.DX4.offsetTop + ' ' + window.screen.availHeight + ' ' + document.all.DX4.style.height);" & chr(13)
	shtml=shtml & "}" & Chr(13)
	shtml=shtml & "</script>" & Chr(13)

 ' sHTML = sHTML & ofv.CorFormHeader
 sHTML = sHTML & ofv.FormHeader(session("FormN"))



	sNombre = ofv.convertircar(sNombre," ","_")

	HrfAsp = "Objetos.asp?"
	HrfTrg = "fesquemas"

	If sTipo="Consulta" Then
		sHTML = sHTML & "<center>"
		sstipo = session("FormN")
		HrfAsp = "ObjTreeSubTemaCons.asp?"
		HrfTrg = ""
	End If

	Fila    = ""
	Celda   = ""
	sHref   = ""

	If  Session("EMPMUN") = "S" Then

		Fila    = ""

		SINPUT3 = "&nbsp;<A class=tt title='SELECCIONE EMPRESA' HREF='javascript:void(0);' onclick='shwDem();' >EMPRESA:&nbsp;&nbsp;</a>"

		Fila    = ""
		Celda   = ""
		sHref   = ""

		sImage  = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(0,5),"0","")
		sLink   = ofv.GenLink("","ttt",sHref,"ftemas","","",sImage,"si")
		Celda   = Celda & ofv.GenCelda("","tt ","","8%","","","",SINPUT3,"si")
		Celda   = Celda & ofv.GenCelda("","ttt ","center","5%","","","",sLink,"si")
		Celda   = Celda & ofv.GenCelda("","ttt  ","","","","","","&nbsp;<BIG>" & ATREMP(0,4) & "</BIG>","si")

		Fila    = fila & ofv.GenRow("","","top","","","","",Celda,"si")

		Tabla   = ofv.GenTabla("","aaa style='"  & Session("CBRDABL") & "' ","","100% ","","0","","-1","-2","","",Fila,"si")





		shtml=shtml & "<script language=javascript>" & Chr(13)
		shtml=shtml & "function shwDem(){" & Chr(13)
		shtml=shtml & "document.all.DEM1.style.display='';" & Chr(13)
		shtml=shtml & "}" & Chr(13)
		shtml=shtml & "</script>" & Chr(13)
		'----------------------------------------------
		sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & Chr(34)
		sinput2 = sinput2  & "underline" & Chr(34) & ";' onmouseout='this.style.textDecoration=" & Chr(34)
		sinput2 = sinput2  & "none" & Chr(34) & ";' href='javascript:void(0);' ONCLICK=" & Chr(34) & "document.getElementById('DEM1').style.display='none';" & Chr(34)
		sinput2 = sinput2  & " TITLE='OCULTAR' ><b>" & sImagenx & "</b></a>"
		TABLA = Tabla & "<DIV id='DEM1' STYLE='display:none;z-index:1;margin-top:1;margin-left:1;background-color:#ffffff;position:absolute;left:1px;top:220px;left:20px;height:280px;width:380px;border-style:outset;border-color:#888888;border-width:2;border-radius: 15px 15px 15px 15px;overflow:auto;'>"


		TABLAX = TablaX & OFV.gentabla("", "aaa style='border-radius: 15px 15px 15px 15px;' ", "", "100%", "", "0", "", "-2", "-1", "", "", "", "no")
		TABLAX = TablaX & "<tr><td  "
		TABLAX = TablaX & "class=ttt align=left colspan=3 STYLE='border-style:outset;border-width:1;background-color:" & session("CMFMNBOT") & ";' >" & sinput2 & "&nbsp;SELECCIONE EMPRESA</td></tr>"
		TABLAX = TablaX & "<tr><td colspan=3 "
		TABLAX = TablaX & " >&nbsp;</td></tr>"

		For IAEMP=1 To TAEMP

			'''  RESPONSE.WRITE  ATREMP(IAEMP,2)

			If ATREMP(IAEMP,2) = 1 Then

				If ATREMP(IAEMP,0) = ATREMP(0,0) Then
					TABLAX = TablaX & "<tr  >"
				Else
					TABLAX = TablaX & "<tr>"
				End If

				TABLAX = TablaX & "<td width=10% "
				If ATREMP(IAEMP,0) = ATREMP(0,0) Then
					'   TABLAX = TablaX & "class=vt  >" & sImagen4si & "</td>"
					xbord = " style='border-style:solid;border-width:1;' "
					ximg = sImagen4si & "&nbsp;"
				Else
					'    TABLAX = TablaX & "class=vt  >" & sImagenI3 & "</td>"
					xbord = "  "
					ximg = sImagenI3 & "&nbsp;"
				End If




				TABLAX = TablaX & "<td class=vti width=60% " & xbord & " >"
				sHref = "OBJtree.asp?Tipo=" & sTipo & "&Nombre=" & LNombre & "&Menu=" & SESSION("MENU") & "&EMPRE=" & ATREMP(IAEMP,0)
				sLink = ofv.GenLink("","cc",sHref,"","","",ximg & "<b>" & Replace(ATREMP(IAEMP,4) & ""," ","&nbsp;") & "</b>","si")
				TABLAX = TablaX & sLink
				TABLAX = TablaX & "</td>"

				sImage = "<img src='../COREss20/Empresa/" & ATREMP(IAEMP,5) & "' width='100' >"
				'''  sImage = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(IAEMP,5),"0","")
				TABLAX = TablaX & "<td width=40% "
				TABLAX = TablaX & "class=vt  >" & sImage & "</td>"

				TABLAX = TablaX & "</tr>"
			End If

		Next

		TABLAX = TablaX & OFV.genFtabla()

		TABLA = Tabla & TABLAX

		TABLA = Tabla & "</DIV>"

		sHTML = sHTML & TABLA

	End If

	Fila = ""


	Tabla = ""



	'  TABLA = Tabla & "<DIV id='x2' STYLE='display:BLOCK;z-index:1;margin-top:1;margin-left:1;overflow:auto;'>"
	Fila = ""

	sHref = ""
	sImage = ofv.GenImage("","","","","Iconos\ICorMMenu.jpg","0","")

	sImage = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(0,5),"0","")

	sLink = ofv.GenLink("","nt",sHref,"ftemas","","",sImage,"si")
	Celda = ofv.GenCelda("","aaa colspan=3 rowspan=2  style='" & Session("CBRDICO") & "'","center","2%","","","",sLink,"si")



	xxx = "<b><u>" & UCase(sstipo) & "</u></b>"

	sLink = ofv.GenLink("","cc",sHref,"ftemas","","","<b><u>" & ATREMP(0,4) & "</u></b>","si")
	Celda = Celda & ofv.GenCelda("","aaa rowspan=2","","","","","","&nbsp;","si")
	Celda = Celda & ofv.GenCelda("","aaa colspan=3 rowspan=2","","","","","",sLink,"si")
	Celda = Celda & ofv.GenCelda("","","","","","","",SINPUT3,"si")
	'  Fila = fila & ofv.GenRow("","","top","","","","",Celda,"si")


	' end if

	stitem = "Tema"
	sIcono = "ICorTema.jpg"
	sNuevoTipo = "SUBTEMAS"
	nNivel = 0

	sID = 0
	nv = 0
	cborderB = "border-bottom-style:solid;border-bottom-width:1;border-bottom-color:" & Session("APLSHW") & ";"
	cborderl = "border-left-style:solid;border-left-width:1;border-left-color:" & Session("APLSHW") & ";"
	borderB = ""
	borderl = ""
	If SESSION("BASE")="K" Then
		Strsql = "Select * From Tema where PSTIPO = 'K' AND (aplic = '" & wkaplic & "' or aplic = 'GLB')  order by inda, descri"
	Else
		Strsql = "Select * From Tema where PSTIPO <> 'K' AND (aplic = '" & wkaplic & "' or aplic = 'GLB')  order by inda, descri"
	End If
	'response.write strsql & "<br>"
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If rs.eof Then
		Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
		Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
		Celda = ""

		Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
		Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")


		Celda = Celda & ofv.GenCelda("","aaa colspan=3 rowspan=2","","2%","","","","&nbsp;","si")
		Celda = Celda & ofv.GenCelda("","aaa rowspan=2","","","","","","&nbsp;","si")
		Celda = Celda & ofv.GenCelda("","vt  rowspan=2","","","","","","No hay elementos asociados","si")
		Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
		Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
	Else

		xTrColor = "#FFFFFF"

		'    Fila = ""
		Do Until rs.eof


			WTVER = 1
' response.write session("UN")
			Strsql = "Select * From ATRMTEMAST WHERE ATRMCODST = " & rs(0) & " AND ATRMTIPOST = 'T' AND ATRMCODES = " & session("UN")
			''response.write strsql & "<br>"
			Set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not rs1.eof Then
				If Not RS1(4) Then  WTVER = 0
				'' response.write strsql & " " & WTVER & "<br>"
			End If

			If WTVER = 1 Then

				If nv = 0 Then
					Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
					Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
				Else
					borderl = cborderl
					borderb = ""
					Celda = ""
					Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
					Celda = Celda & ofv.GenCelda("","aaa style='" & borderl & borderb & "'","","","","","","&nbsp;","si")
					Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
					Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

				End If



				borderl = cborderl
				borderb = cborderb
				Celda = ""
				Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
				Celda = Celda & ofv.GenCelda("","aaa style='" & borderl & borderb & "'","","","","","","&nbsp;","si")




				If rs(3) = "R" Then

					sHref = "'javascript:void(0);'   "



					If rs(5) = "GLB" Then
						sImage = "<font face='CorFlaticon' size=2 color='#800000'>&#0252;</font>"
					Else
						sImage = "<font face='CorFlaticon' size=2 color='#800000'>&#060;</font>"
					End If



				Else



					sHref = HrfAsp & "Tipo=" & sNuevoTipo & "&ID=" & rs(0) & "&sitem=TEMA:_"
					sHref = sHref & ofv.convertircar(rs(1)," ","_") & "&stitem="
					sHref = sHref & ofv.convertircar(stitem," ","_") & "&sticono=" & sIcono
					sHref = sHref & "&stnivel=" & nNivel

					If rs(5) = "GLB" Then
						sImage = "<font face='CorFlaticon' size=2 color='#f7aa22'>&#0252;</font>"
					Else
						sImage = "<font face='CorFlaticon' size=2 color='#f7aa22'>&#060;</font>"
					End If


				End If



				sLink = ofv.GenLink("","nt",sHref,HrfTrg,"","",sImage,"si")
				Celda = Celda & ofv.GenCelda("","ut style='" & XspSShw("#dddddd","125") & ";' colspan=3 rowspan=2 ","center","2%","","","",sLink,"si")
				sLink = ofv.GenLink("","nt",sHref,HrfTrg,"","",rs(1),"si")
				Celda = Celda & ofv.GenCelda("","aaa rowspan=2","","","","","","&nbsp;","si")
				Celda = Celda & ofv.GenCelda("","aaa  rowspan=2 bgcolor=" & xTrColor & " ","","","","","",sLink,"si")
				Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")


				Fila = Fila & ofv.GenRow("","aaa ","","","","","",Celda,"si")


				If xTrColor = "#FFFFFF" Then
					xTrColor = "#F0F0F0"
				Else
					xTrColor = "#FFFFFF"
				End If

				nv = 1

			End If

			rs.movenext
		Loop

	End If
	If session("sEstadon") = "4" Then
		If nv = 0 Then
			Celda = ofv.GenCelda("","","","","","","","&nbsp;","si")
			Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")
		Else
			borderl = cborderl
			borderb = ""
			Celda = ""
			Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
			Celda = Celda & ofv.GenCelda("","aaa style='" & borderl & borderb & "'","","","","","","&nbsp;","si")
			Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
			Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

		End If


		borderl = cborderl
		borderb = cborderb
		Celda = ""
		Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
		Celda = Celda & ofv.GenCelda("","aaa style='" & borderl & borderb & "'","","","","","","&nbsp;","si")


		sHref = HrfAsp & "Tipo=" & sNuevoTipo & "&ID=xxx&sitem=RECICLAJE"
		sHref = sHref & "&stitem="
		sHref = sHref & ofv.convertircar(stitem," ","_") & "&sticono=ICorRecycle.jpg"
		sHref = sHref & "&stnivel=" & nNivel
		sImage = ofv.GenImage("xxx-RYC","","","","Iconos\ICorRecycle.jpg","0","")
		sLink = ofv.GenLink("","nt",sHref,HrfTrg,"","",sImage,"si")
		Celda = Celda & ofv.GenCelda("","aaa colspan=3 rowspan=2 style='" & Session("CBRDICO") & "'","center","2%","","","",sLink,"si")
		sLink = ofv.GenLink("","nt",sHref,HrfTrg,"","","RECICLAJE","si")
		Celda = Celda & ofv.GenCelda("","aaa rowspan=2","","","","","","&nbsp;","si")
		Celda = Celda & ofv.GenCelda("","aaa rowspan=2 bgcolor=" & xTrColor & " ","","","","","",sLink,"si")
		Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
		Fila = Fila & ofv.GenRow("","aaa ","","","","","",Celda,"si")
	End If
	Celda = ""
	Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
	Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
	Celda = Celda & ofv.GenCelda("","","","","","","","&nbsp;","si")
	Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

	Celda = ""
	Celda = Celda & ofv.GenCelda("","","","2%","","","","&nbsp;","si")
	Celda = Celda & ofv.GenCelda("","","","2%","","","","&nbsp;","si")
	Celda = Celda & ofv.GenCelda("","","","2%","","","","&nbsp;","si")
	Celda = Celda & ofv.GenCelda("","","","2%","","","","&nbsp;","si")
	Celda = Celda & ofv.GenCelda("","","","1%","","","","&nbsp;","si")
	Celda = Celda & ofv.GenCelda("","","","1%","","","","&nbsp;","si")
	Celda = Celda & ofv.GenCelda("","","","89%","","","","&nbsp;","si")
	Celda = Celda & ofv.GenCelda("","","","1%","","","","&nbsp;","si")
	Fila = fila & ofv.GenRow("","","","","","","",Celda,"si")

	If sTipo="Consulta" Then
		Tabla = Tabla & ofv.GenTabla("","aaa  topmargin=4 leftmargin=4 style='border-style=solid;border-width:1;' ","","95% ","","0","","-1","-2","","",Fila,"si")
	Else
		Tabla = Tabla & ofv.GenTabla("","aaa style='margin-top:4;margin-left:4;' ","","90% ","","0","","-1","-2","","",Fila,"si")

		If recarga = "S" Then
			Formulario = OFV.GenForm("recarga", "", "nulo2.asp", "fesquemas", "post", "","<input type=hidden value='si' >", "si")
			Tabla = Tabla & Formulario
		End If


	End If

	'  TABLA = Tabla & "</DIV>"

	Call ofv.cerrarconsulta(rs)
	Call ofv.cerrarconn(cn)










	'''''''''''''''''''''''''''

	Celda = ""

	Celda = Celda & "<div id='DX2' style='margin-top:13;margin-left:1;position:absolute;left:10px;display:none;"
	Celda = Celda & "width:35%;height:100%;overflow:auto;border-style:solid;border-width:1;border-color:#993333;"
	Celda = Celda & "background-color:" & session("APLBGD") & ";"
	Celda = Celda & "' >"

	Celda = Celda & TABLA

	Celda = Celda & "</div>"



	Celda = Celda & "<div id='DX4' style='margin-top:13;margin-left:5;position:absolute;left:300px;display:none;"
    Celda = Celda & "  width:63%;height:100%;overflow:auto;border-style:solid;border-width:1;border-color:#993333;"
	Celda = Celda & "background-color:" & session("APLBGD") & ";"
	Celda = Celda & "' >"

	Celda = Celda & "&nbsp;"

	Celda = Celda & "</div>"



	sHTML = sHTML & Celda
    sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","",Agregar,Buscar,"no",Reportes)
    '  shtml=shtml & ARMAM()
    '<script>document.getElementById('THM').style.display='none';alert('1');</script>
	sHTML = sHTML & "</body></html>"

End If

response.write  sHTML

%>

