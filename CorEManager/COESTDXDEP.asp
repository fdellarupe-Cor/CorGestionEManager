<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	
	'' session("ModMain") = "CL"
	base=request.querystring("BASE")
	
	If base <> "" Then
		SESSION("BASE")=base
	End If
	
	response.expires=-1441
	Response.AddHeader "PRAGMA", "NO-CACHE"
	
	
	Set cn = ofv.conectar(ofv.strconn2)
	
	sHTML = ofv.FormHeader( session("FormN"))
	sHTML = sHTML & "<center>"
	
	tt="&nbsp;"
	
	
	XXXX = NEWEMPRE("NO")
	
	session("ASP")="SUBCONSULTA.ASP"
	
	sEMPRE = request.querystring("EMPRE")
	
	If sEMPRE = ""  And SESSION("UN") = "" Then
		
		sEMPRE = ATREMP(0, 0)
		SESSION("UN") = sEMPRE & ""
		If (SESSION("UN") & "") = "" Then SESSION("UN") = 1
	ElseIf sEMPRE <> "" Then
		XXXX = NEWEMPRE(sEMPRE)
		SESSION("UN")=sEMPRE
	Else
		XXXX = NEWEMPRE(SESSION("UN"))
		sEMPRE=SESSION("UN")
	End If
	
	' response.write session("UN")
	TablaH = ofv.GenTabla("","","","20%","","0","0","0","25","","",selecEmpre(),"si")
	
	shtml = shtml & tablah
	
	
	XFNAME   = request.querystring("XFNAME")
	XFSRH    = request.querystring("XFSRH")
	XFSRD    = request.querystring("XFSRD")
	
	
	
	If XFNAME = "" Then XFNAME = "PV"
	If XFSRD  = "" Then XFSRD = 0
	If XFSRH  = "" Then XFSRH = 0
	
	' PRIMERA VEZ
	If XFNAME="PV"  Then
		XFSRD = CLng(Date()) - CLng(CDate("01-01-1900"))
		'''''''''''''''XFSRH = XFSRD + 15
		XFSRH = XFSRD
		
	End If
	
	If XFNAME = "FD" Then
		XFSRD = request.form("DESDE")
		XFSRH = XFSRH
		
	End If
	
	If XFNAME = "FH" Then
		XFSRD = XFSRD
		XFSRH =request.form("HASTA")
		
	End If
	
	
	'Descripcion
	'sInput = "<h3>Panel de filtros para la consulta</h3>"
	'Celda  = ofv.GenCelda("","ut colspan=2 bgcolor='#f4f7fc' style='border-style:solid;border-width:1;border-color:#dddddd;' ","","","","","",sInput,"si")
	'Fila   = ofv.genrow("","","","","30","","",Celda,"si")
	Celda  = ofv.GenCelda("","vt","","30%","","","","&nbsp;","si")
	Celda  = Celda & ofv.GenCelda("","vti","","70%","","","","&nbsp;","si")
	'Fila   = Fila & ofv.GenRow("","","","","10","","",Celda,"si")
	
	sImagenE = "<B><i class='las la-calendar'STYLE=font-size:14pt;color:" & Session("CMFPNBOT") &"></i></B>"
	sAccion = "'javascript:void(0);'" & GTCSETCALR("frmCalendarioD","Desde",GTCSMT,GTCFCH,GTCFCHDAT,XFSRD,4,2,0,0,0)
	sinputE = "&nbsp;&nbsp;<a  class=UT href=" & sAccion
	sinputE = sinputE  & " TITLE='MODIFICAR FECHA' >" & sImagenE & "</a>"
	
	
	sAccionHasta = "'javascript:void(0);'" & GTCSETCALR("frmCalendarioH","Hasta",GTCSMT,GTCFCH,GTCFCHDAT,XFSRH,4,2,0,0,0)
	sinputHasta = "&nbsp;&nbsp;<a  class=UT href=" & sAccionHasta
	sinputHasta = sinputHasta  & " TITLE='MODIFICAR FECHA' >" & sImagenE & "</a>"
	
	
	
	' CELDAS DE DESDE
	
	If  CLng(XFSRD) = 0 Then
		sInput2 = "------------"
	Else
		sInput2 = CDate(CLng(XFSRD) + CLng(CDate("01-01-1900")) )
		swd = DatePart("w",sInput2)
		sInput2 = Fechador(XFSRD)
	End If
	
	
	
	
	Celda  = ofv.GenCelda("","tc","","","","","","<font>Desde</font>","si")
	sInput = ofv.GenerarInput("Desde2",sInput2,"agregar","","80","dt","") & ofv.GenerarInput("Desde3",XFSRD,"hidden","","20","dt","")
	Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput&sinputE,"si")
	Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")
	
	
	' CELDAS DE HASTA
	If CLng(XFSRH) = 0 Then
		sInput2 = "------------"
	Else
		sInput2 = CDate(CLng(XFSRH) + CLng(CDate("01-01-1900")) )
		swd = DatePart("w",sInput2)
		sInput2 = Fechador(XFSRH)
	End If
	
	
	Celda  = ofv.GenCelda("","tc","","","","","","<font>Hasta</font>","si")
	sInput = ofv.GenerarInput("Hasta2",sInput2,"agregar","","80","dt","") & ofv.GenerarInput("Hasta3",XFSRH,"hidden","","20","dt","")
	Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput&sinputHasta,"si")
	Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")
	
	
	If  session("APLMDCLI")="CFA" Then
		scomK ="NOL"
		scomD = "SIN ACCESOS"
		sInput = ofv.gencomboFX("Subconsulta","ESEL","NOL",scomK,scomD,"no")
		
	Else
		
		scomK = "MAS" & Chr(9) & "MEN" & Chr(9) & "NOL"
		scomD = "MAS ACCESOS" & Chr(9) & "MENOS ACCESOS" & Chr(9) & "SIN ACCESOS"
		sInput = ofv.gencomboFX("Subconsulta","ESEL","MAS",scomK,scomD,"no")
		
	End If
	
	Celda  = ""
	Celda  = ofv.GenCelda("","tc","","","","","","<font>Orden:</font>","si")
	Celda  = celda& ofv.GenCelda("","uti","","","","","",sInput,"si")
	Fila   = Fila & ofv.GenRow("","","","","25","","",Celda,"si")
	
	
	Celda    = ""
	Celda    = Celda & ofv.GenCelda("","tc","","","","","","<font> Departamento:</font>","si")
	
	
	Strsql = "     select ECO,DESCrI FROM ECO where eco.estado=3 and ecoUN=" & session("UN") & " ORDER BY DESCRI "
	Set rsr = ofv.crearconsultaEx(Strsql,cn,1,parametros)
	
	If Not rsr.eof Then
		sSectort=rsr(0)
	End If
	
	Call ofv.cerrarconsulta(rsr)
	
	
	sInput = ofv.GenerarCombo2("lista",sSectort,Strsql,cn,0,1,"sb")
	Celda  = Celda & ofv.GenCelda("","uti","","","","","",sInput,"si")
	Fila   = Fila & ofv.GenRow("","","","","25","","",Celda,"si")
	
	
	
	'IF session("APLMDCLI")="BLP" THEN
	
	Celda  = ofv.GenCelda("","tc","","","","","","<font>Formato Exportacion :</font>","si")
	SCOMD= "XLS" &  Chr(9) & "NO"
	scomK=SCOMD
	sInput = ofv.gencomboFX("Subconsulta","EXPORT","NO",scomK,scomD,"no")
	Celda = Celda & ofv.gencelda("","uti","","","","","",sInput,"si")
	Fila = Fila & ofv.genrow("","","","","25","","",Celda,"si")
	'END IF
	
	
	
	
	
	' BOTON BUSCAR
	sTARGET  = ""
	sAccion    = "ConsdocumentoswebScons.asp?Menu=PORREVISION&CantRegMostrar=10&CantRegMover=0&SORDEN=N&Nombre=POR REVISION"
	
	Celda    = ofv.GenCelda("","ut","","","","","","&nbsp;","si")
	sInput   = ofv.GenerarInput("","Buscar","submit","","10","bt","")
	
	sInput   = sinput & ofv.GenerarInput("Tcons",sMenu,"hidden","","1","dt","")
	Celda    = Celda & ofv.GenCelda("","vti","Center","","","","",sInput,"si")
	formx    = fila
	formx    = formx  & ofv.GenRow("","","","","40","","",Celda,"si")
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	sInput = ofv.GenerarInput("","Buscar","submit","","10","bt","")
	Celda  = ofv.GenCelda("","","Center colspan=2","","","","",sInput,"si")
	Fila   = Fila & ofv.GenRow("","","","","25","","",Celda,"si")
	
	formx = fila
	formx = formx  & ofv.GenRow("","","","","40","","",Celda,"si")
	
	
	sAccion = "COESTULis.asp?CantRegMostrar=10&CantRegMover=0&tipoC=ACCXESQDEP"
	
	If session("APLMDCLI")="BLP" Then
		Formulario = ofv.GenForm("Subconsulta","",sAccion,"","post Onsubmit="&Chr(34) &"if(document.Subconsulta.EXPORT.value=='XLS'){this.target='_NEW'; } else{this.target='_self'};" &Chr(34),"",Fila,"si")
		
	Else
		
		Formulario = ofv.GenForm("Subconsulta","",sAccion,"","post","",Fila,"si")
		
		
	End If
	
	
	
	
	
	
	' ****** CODIGO PARA CALENDARIO
	
	sInputDesde=ofv.GenerarInput("Desde",CDate(XFSRD+ CLng(CDate("01-01-1900"))),"hidden","","20","dt","")
	sAccionRetornoD="COESTDXDEP.asp?nombre="& sNombren &"&XFNAME=FD&Desde=" & XFSRD & "&Hasta=" & XFSRH & "&Fecha=" & fecha &"&sorden=" & sorden & "&CantRegMostrar=10&CantRegMover=0&XFSRD="& XFSRD & "&XFSRH=" & XFSRH
	formRetornoD= ofv.GenForm("frmCalendarioD","",sAccionRetornoD,STARGET,"post","",sInputDesde,"si")
	
	
	sInputHasta=ofv.GenerarInput("Hasta",CDate(XFSRH+ CLng(CDate("01-01-1900"))),"hidden","","20","dt","")
	sAccionRetornoH="COESTDXDEP.asp?nombre="& sNombren &"&XFNAME=FH&Desde=" & XFSRD & "&Hasta=" & XFSRH & "&Fecha=" & fecha & "&sorden=" & sorden & "&CantRegMostrar=10&CantRegMover=0&XFSRD="& XFSRD & "&XFSRH=" & XFSRH
	formRetornoH= ofv.GenForm("frmCalendarioH","",sAccionRetornoH,STARGET,"post","",sInputHasta,"si")
	
	' ****** CODIGO PARA CALENDARIO
	
	Formulario = Formulario &formRetornoD & formRetornoH
	
	
	
	
	
End If



filas2 = "" & filas2 & "" & ofv.GenTabla("","aaa","","80%","","0","","5","3","","",Formulario,"si")

sHTML = sHTML & filas2




'Botonera
Agregar = ""
Buscar  = ""
Celda   = ""

sHTML = sHTML & "</center>"

sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

sHTML = sHTML & "</script>"& GTCCALENDAR("")

sHTML = sHTML & "</body></html>"

Call ofv.cerrarconn(cn)


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
				sHref = "COESTDXDEP.asp?EMPRE=" & ATREMP(IAEMP,0)
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