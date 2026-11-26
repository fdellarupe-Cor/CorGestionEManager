<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	
	'' session("ModMain") = "CL"
	
	response.expires=-1441
	Response.AddHeader "PRAGMA", "NO-CACHE"
	
	base=request.querystring("BASE")
	
	If base <> "" Then
		SESSION("BASE")=base
	End If
	
	
	sHTML = ofv.FormHeader( session("FormN"))& Chr(13)
	
	sHTML = sHTML & "<center>"
	
	tt="&nbsp;"
	
	
	
	
	XFNAME   = request.querystring("XFNAME")
	XFSRH    = request.querystring("XFSRH")
	XFSRD    = request.querystring("XFSRD")
	
	
	
	If XFNAME = "" Then XFNAME = "PV"
	If XFSRD  = "" Then XFSRD = 0
	If XFSRH  = "" Then XFSRH = 0
	
	' PRIMERA VEZ
	If XFNAME="PV"  Then
		XFSRD = CLng(Date()) - CLng(CDate("01-01-1900"))
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
	
	
	
	'Descripcion
	'sInput = "<h3>Panel de filtros para la consulta</h3>"
	'Celda  = ofv.GenCelda("","ut colspan=2 bgcolor='#f7eedd' style='border-style:solid;border-width:1;border-color:#dddddd;' ","","","","","",sInput,"si")
	'Fila   = ofv.genrow("","","","","30","","",Celda,"si")
	Celda  = Celda
	Celda  = Celda
	Fila   = Fila
	
	
	
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
	
	
	' BOTON BUSCAR
	sTARGET  = ""
	sAccion    = "listadoconsultas.asp?Menu=PORREVISION&CantRegMostrar=10&CantRegMover=0&SORDEN=N&Nombre=POR REVISION"
	
	Celda    = ofv.GenCelda("","ut","","","","","","&nbsp;","si")
	sInput   = ofv.GenerarInput("","Buscar","submit","","10","bt","")
	
	sInput   = sinput & ofv.GenerarInput("Tcons",sMenu,"hidden","","1","dt","")
	Celda    = Celda & ofv.GenCelda("","vti","Center","","","","",sInput,"si")
	formx = fila
	formx    = formx  & ofv.GenRow("","","","","40","","",Celda,"si")
	
	
	
	
	
	
	
	' ****** CODIGO PARA CALENDARIO
	
	sInputDesde=ofv.GenerarInput("Desde",CDate(XFSRD+ CLng(CDate("01-01-1900"))),"hidden","","20","dt","")
	sAccionRetornoD="COfecrevision.asp?nombre="& sNombren &"&XFNAME=FD&Desde=" & XFSRD & "&Hasta=" & XFSRH & "&Fecha=" & fecha &"&sorden=" & sorden & "&CantRegMostrar=10&CantRegMover=0&XFSRD="& XFSRD & "&XFSRH=" & XFSRH
	formRetornoD= ofv.GenForm("frmCalendarioD","",sAccionRetornoD,STARGET,"post","",sInputDesde,"si")
	
	
	sInputHasta=ofv.GenerarInput("Hasta",CDate(XFSRH+ CLng(CDate("01-01-1900"))),"hidden","","20","dt","")
	sAccionRetornoH="COfecrevision.asp?nombre="& sNombren &"&XFNAME=FH&Desde=" & XFSRD & "&Hasta=" & XFSRH & "&Fecha=" & fecha & "&sorden=" & sorden & "&CantRegMostrar=10&CantRegMover=0&XFSRD="& XFSRD & "&XFSRH=" & XFSRH
	formRetornoH= ofv.GenForm("frmCalendarioH","",sAccionRetornoH,STARGET,"post","",sInputHasta,"si")
	
	' ****** CODIGO PARA CALENDARIO
	
	Formulario = ofv.GenForm("Subconsulta","",sAccion,STARGET,"post","",formx,"si") &formRetornoD & formRetornoH
	
	
	
	
	
	'sHTML = sHTML & "<center><BR>"
	'sHTML = sHTML
	'sHTML = sHTML & "</center>"
	
	
	
	
	
	
	
	filas2 = "" & filas2 & "" & ofv.GenTabla("","aaa bgcolor='#eeeeee'","","60% ","","0","","5","3","","",Formulario,"si")
	'sHTML = sHTML &                     ofv.GenTabla("","aaa bgcolor='#eeeeee' style='border-style:solid;border-width:1;border-color:#dddddd;' ","","60% ","","0","","5","3","","",Formulario,"si")
	'filas2 = "<br>" & filas2 & "<br>" & ofv.GenTabla("","AAA style='border-style:solid;border-width:1;border-color:#dddddd;'","","30%","","0","","0","0","","",FILA,"si")
	
	sHTML = sHTML & filas2
	
End If

'Botonera
Agregar = ""
Buscar  = ""
Celda   = ""

sHTML = sHTML & "</center>"

sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)

sHTML = sHTML & "</script>"& GTCCALENDAR("")

sHTML = sHTML & "</body></html>"

' call ofv.cerrarconn(cn)


response.write sHTML

Function Fechador(wkfecha)
	
	
	
	fec = CDate(wkfecha + CLng(CDate("01-01-1900")))
	ndia = ofv.cdia(Weekday(fec))
	nMes = ofv.cmes(Month(fec))
	
	Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)
	
	
End Function

%>

