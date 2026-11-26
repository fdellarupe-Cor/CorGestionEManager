<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%
ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires = 0

	XXXX = NEWEMPRE("NO")
	base=request.querystring("BASE")

        IF base <> "" THEN
          SESSION("BASE")=base
      END IF
	session("ASP")="LisVigentes.ASP"
	'response.write "STRING: " & request.querystring & "<br>" & "FORM: " & request.form & "<br>" & "estados: " & request.form("estados")
	sUN = request.querystring("EMPRE")
	'response.write "sEMPRE: " & sEMPRE
	If sUN = ""  And SESSION("UN") = "" Then

		sUN = ATREMP(0, 0)
		SESSION("UN") = sUN & ""
		If (SESSION("UN") & "") = "" Then SESSION("UN") = 1
	ElseIf sUN <> "" Then
		XXXX = NEWEMPRE(sUN)
		SESSION("UN")=sUN
	Else
		XXXX = NEWEMPRE(SESSION("UN"))
		sUN=SESSION("UN")
	End If


	shtml = ofv.FormHeader(session("FormN"))

	TablaH = ofv.GenTabla("","","","20%","","0","0","0","25","","",selecEmpre(),"si")

	shtml = shtml & tablah & ""

	XFNAME   = request.querystring("XFNAME")
	XFSRH    = request.querystring("XFSRH")
	XFSRD    = request.querystring("XFSRD")



	If XFNAME = "" Then XFNAME = "PV"
	If XFSRD  = "" Then XFSRD = 0
	If XFSRH  = "" Then XFSRH = 0

	' PRIMERA VEZ
	If XFNAME="PV"  Then
		'XFSRD = clng(DATE()) - clng(cdate("01-01-1900"))
		XFSRD = 0
		XFSRH = XFSRD '+ 15

	End If

	If XFNAME = "FD" Then
		XFSRD = request.form("DESDE")
		XFSRH = XFSRH

	End If

	If XFNAME = "FH" Then
		XFSRD = XFSRD
		XFSRH =request.form("HASTA")

	End If


	sImagenE = "<B><i CLASS='las la-calendar' STYLE=font-size:14pt;COLOR:" & Session("CMFPNBOT") & "></i></B>"
	sAccion = "'javascript:void(0);'" & GTCSETCALR("frmCalendarioD","Desde",GTCSMT,GTCFCH,GTCFCHDAT,XFSRD,15,2,0,0,0)
	sinputE = "&nbsp;&nbsp;<a  class=UT href=" & sAccion
	sinputE = sinputE  & " TITLE='MODIFICAR FECHA' >" & sImagenE & "</a>"


	sAccionHasta = "'javascript:void(0);'" & GTCSETCALR("frmCalendarioH","Hasta",GTCSMT,GTCFCH,GTCFCHDAT,XFSRH,4,2,0,0,0)
	sinputHasta = "&nbsp;&nbsp;<a  class=UT href=" & sAccionHasta
	sinputHasta = sinputHasta  & " TITLE='MODIFICAR FECHA' >" & sImagenE & "</a>"



	' CELDAS DE DESDE
	If  CLng(XFSRD) = 0 Then
		sInput2 = "TODOS"
	Else
		sInput2 = CDate(CLng(XFSRD) + CLng(CDate("01-01-1900")) )
		swd = DatePart("w",sInput2)
		sInput2 = Fechador(XFSRD)
	End If




	Celda  = ofv.GenCelda("","tc","","","","","","<font>Fecha vigencia desde</font>","si")
	sInput = ofv.GenerarInput("Desde2",sInput2,"agregar","","80","dt","") & ofv.GenerarInput("Desde3",CStr(XFSRD + CLng(CDate("01-01-1900"))),"hidden","","20","dt","")
	Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput&sinputE,"si")



	Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")


	' CELDAS DE HASTA
	If CLng(XFSRH) = 0 Then
		sInput2 = "TODOS"
	Else
		sInput2 = CDate(CLng(XFSRH) + CLng(CDate("01-01-1900")) )
		swd = DatePart("w",sInput2)
		sInput2 = Fechador(XFSRH)
	End If



	If XFSRD >XFSRH Then
		msg=" onhover='validarfechas();'"
	Else
		msg=""
	End If

	Celda  = ofv.GenCelda("","tc","","","","","","<font>Fecha vigencia hasta</font>","si")
	sInput = ofv.GenerarInput("Hasta2",sInput2,"agregar","","80","dt","") & ofv.GenerarInput("Hasta3",CStr(XFSRH+ CLng(CDate("01-01-1900"))),"hidden","","20","dt","")
	Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput&sinputHasta,"si")

	Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

	Celda  = ofv.GenCelda("","tc","","","","","","<font>Tipo</font>","si")

	Set cn = ofv.conectar(ofv.strconn2)

	if session("base") = "K" THEN
		 STPIPOF= " IN ('K') "
		ELSE
		 STPIPOF= " NOT IN ('K','PD') "
	END IF

	sql = "select tipobjeto, descri from tipobjeto where TIPOCPT " & STPIPOF & " order by descri"
	Set rs = ofv.crearconsultaEx(sql,cn,1,parametros)

	If Not rs.eof Then tipo = "TODOS"

	scomK = "TODOS"
	scomD = "TODOS"
	scomSEP = Chr(9)

	Do While Not rs.eof

		scomK = scomK & scomSEP & rs(0)
		scomD = scomD & scomSEP & rs(1)

		rs.movenext
	Loop

	sInput = ofv.gencomboFX(formu,"tipo",tipo,scomK,scomD,"no")
	Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
	Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")


	sInputZ = "<input class=bt type=submit xonclick=document.getElementById('formu').submit value=Buscar />"

	Celda  = ofv.GenCelda("","ut","Center colspan=2","","","","",sInputZ,"si")

	Fila = Fila & ofv.GenRow("","","","","","","",Celda,"si")





	sTARGET  = ""

	' ****** CODIGO PARA CALENDARIO
	sInputDesde=ofv.GenerarInput("Desde",XFSRD+ CLng(CDate("01-01-1900")),"hidden","","20","dt","")
	sAccionRetornoD="LisVigentes.asp?nombre="& sNombren &"&XFNAME=FD&Desde=" & XFSRD & "&Hasta=" & XFSRH & "&Fecha=" & fecha &"&sorden=" & sorden & "&CantRegMostrar=10&CantRegMover=0&XFSRD="& XFSRD & "&XFSRH=" & XFSRH
	formRetornoD= ofv.GenForm("frmCalendarioD","",sAccionRetornoD,STARGET,"post","",sInputDesde,"si")


	sInputHasta=ofv.GenerarInput("Hasta",XFSRH+ CLng(CDate("01-01-1900")),"hidden","","20","dt","")
	sAccionRetornoH="LisVigentes.asp?nombre="& sNombren &"&XFNAME=FH&Desde=" & XFSRD & "&Hasta=" & XFSRH & "&Fecha=" & fecha & "&sorden=" & sorden & "&CantRegMostrar=10&CantRegMover=0&XFSRD="& XFSRD & "&XFSRH=" & XFSRH
	formRetornoH= ofv.GenForm("frmCalendarioH","",sAccionRetornoH,STARGET,"post","",sInputHasta,"si")

	' ****** CODIGO PARA CALENDARIO
	sAccion = "LisVigentesFiltro.asp"
	tabla= ofv.GenTabla("","aaa bgcolor='#eeeeee' ","","80% ","","0","","5","3","","",Fila,"si")
	Formulario =  formulario & ofv.GenForm("Subconsulta","",sAccion,STARGET,"post" & funxt,"",TABLA,"si")  & formRetornoD & formRetornoH




	Agregar = ""
	Buscar  = ""
	Celda   = ""


	response.write  ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

	sHTML = sHTML & GTCCALENDAR("") & Formulario

	response.write "</body></html>"

End If

response.write sHTML

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

		Tabla   = ofv.GenTabla("","aaa style='"  & Session("CBRDABL") & "' ","","100% ","","0","","-1","-2","","",Fila,"si")





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
				sHref = "LisVigentes.asp?EMPRE=" & ATREMP(IAEMP,0)
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
Function Fechador(wkfecha)


	fec = CDate(wkfecha + CLng(CDate("01-01-1900")))
	ndia = ofv.cdia(Weekday(fec))
	nMes = ofv.cmes(Month(fec))

	Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)


End Function



%>
