<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires=-1441
	Response.AddHeader "PRAGMA", "NO-CACHE"
	
	CantRegMostrarx = 10
	CantRegMoverx = CDbl(request.querystring("CantRegMover"))
	vuelta = ofv.vueltaasp("../",CantRegMostrarx,Cantregmoverx)
	
	ofv.ObtenerAtributos Session("Form"),""
	bAgregar = ofv.AAgregar
	bModificar = ofv.AModificar
	bSuprimir = ofv.AEliminar
	bConsultar = ofv.AConsultar
	
	APLCHECK()
	
	nFecpub = CLng(CDate("01-01-1900"))
	Set cn = ofv.conectar(ofv.strconn0)
	'set rsx1 = ofv.conectar(ofv.strconn0)
	
	sForm = "formlis"
	Filax = ""
	
	'RESPONSE.WRITE REQUEST.QUERYSTRING("OPCION")
	sopcion=REQUEST.QUERYSTRING("OPCION")
	' RESPONSE.WRITE SOPCION & "222"
	
	DEST="C:\exportarBpn\"
	XORDLISD    = "" & request.FORM("XORDLIS")
	XTESQAD     = "" & request.FORM("XTESQA")
	XTESAD      = "" & request.FORM("XTESA")
	Xdesde     = "" & request.FORM("desde")
	Xhasta     = "" & request.FORM("hasta")
	
	
	If XORDLISD  = ""  Then xordlisd  = "1"
	If XORDLISD  = "0" Then xordlisd  = "1"
	If xtesqad   = "" Then xtesqad   = "0"
	If xtesad    = "" Then xtesad    = "0"
	If Xdesde    = "" Then xdesde  = CLng(Date) - CLng(CDate("01-01-1900"))
	If Xhasta    = "" Then xhasta  = CLng(Date) - CLng(CDate("01-01-1900"))
	
	
	'RESPONSE.WRITE  ".1." & request.FORM("XORDLIS")'
	'RESPONSE.WRITE  ".2." & request.FORM("XTESQA")'
	'RESPONSE.WRITE  ".3." & request.FORM("XTESA"')
	'RESPONSE.WRITE  ".4." & request.FORM("desde")
	'RESPONSE.WRITE  ".5." & request.FORM("hasta")
	'RESPONSE.WRITE  ".6." & request.FORM("desdex")
	'RESPONSE.WRITE  ".7." & request.FORM("hastax")
	
	
	
	XFNAME   = request.querystring("XFNAME")
	
	
	
	
	If XFNAME = "FD" Then
		xdesde = request.form("DESDEx")
		
	End If
	
	If XFNAME = "FH" Then
		Xhasta = request.form("HASTAx")
		
	End If
	
	
	
	
	If sopcion ="SETLAUTO"  Then
		SCOFECHAPUBLICD    		= UCase(request.FORM("COFECHAPUBLICD"))
	Else
		
		SCOFECHAPUBLICD    		= "TODOS" 'UCASE(request.QUERYSTRING("SCORFECHAPUBLICD"))
	End If
	'RESPONSE.WRITE  SORFECHAPUBLICD & "222"
	
	
	If sopcion ="SETLAUTO"  Then
		SCOFECHAPUBLICH    		= UCase(request.FORM("COFECHAPUBLICH"))
	Else
		SCOFECHAPUBLICH    		= "TODOS"' UCASE(request.QUERYSTRING("SCORFECHAPUBLICH"))
	End If
	'RESPONSE.WRITE  SORFECHAPUBLICH & "222"
	
	If sopcion ="SETLAUTO"  Then
		SCORDLISTADO   	        = UCase(request.FORM("CORDLISTADO"))
	Else
		SCORDLISTADO        	= "TODOS"  'UCASE(request.QUERYSTRING("SCORDLISTADO"))
	End If
	
	If sopcion ="SETLAUTO"  Then
		SCOTIPOESQUEMA    	= UCase(request.FORM("COTIPOESQUEMA"))
	Else
		SCOTIPOESQUEMA		= "TODOS"  'UCASE(request.QUERYSTRING("SCOTIPOESQUEMA"))
	End If
	
	If sopcion ="SETLAUTO"  Then
		SCOESTADOS    		= UCase(request.FORM("COESTADOS"))
	Else
		SCOESTADOS	        = "TODOS"  'UCASE(request.QUERYSTRING("SCOESTADOS"))
	End If
	
	If sopcion ="SETLAUTO"   Then
		SCOARCHIVO    		= request.FORM("COARCHIVO")
	Else
		SCOARCHIVO     		= dest ' EliminarTexto(UCASE(request.QUERYSTRING("SCORARCHIVO")))
	End If
	
	
	If request.FORM("COARCHIVO") <> ""   Then
		session("SCOARCHIVO")= (request.FORM("COARCHIVO"))
		
	Else
		
		SCOARCHIVO     		= session("SCOARCHIVO")
	End If
	
	If SCOARCHIVO = "" Then SCOARCHIVO = dest
	
	sHTML = ofv.FormHeader(session("FormN")) & Chr(13)
	sHTML = sHTML & "<CENTER>"
	
	'---------------------------------------------------------------------------------------------------------------------
	''  response.write REQUEST.QUERYSTRING("OPCIONINS") & chr(13)
	If REQUEST.QUERYSTRING("OPCIONINS") = "INSERTAR" Then
		
		
		EROR_FECHA = ""
		
		
		''      response.write request.querystring
		STRSQL = "INSERT INTO "
		STRSQL = STRSQL & "	COPARAEXPORTDOCTM (USUARIO, TIPOESQUEMA, ESTADO, "
		STRSQL = STRSQL & "	F_PEDIDO, F_GENERADO, F_BORRADO, "
		STRSQL = STRSQL & "	NOMBRE_ARCHIVO, FDESDE, FHASTA, PROGRAMA "
		STRSQL = STRSQL & "	,FLG_TERMINADO, ORDENLISTADO)" & Chr(13)
		STRSQL = STRSQL & " VALUES (" & Chr(13)
		
		STRSQL = STRSQL & "	'" & SESSION("USUARIO") & "',"
		''RESPONSE.WRITE CHR(13) & SESSION("USUARIO") & CHR(13)
		
		
		If XTESQAD = "TODOS" Then
			STRSQL = STRSQL & "		0," & Chr(13)
		Else
			STRSQL = STRSQL & "		'" & XTESQAD & "',"
		End If
		''''RESPONSE.WRITE CHR(13) & "esquema " & XTESQAD & "<br>"
		
		If XTESAD  = "TODOS" Then
			STRSQL = STRSQL & "		0," & Chr(13)
		Else
			STRSQL = STRSQL & "		'" & XTESAD & "',"
		End If
		''''RESPONSE.WRITE CHR(13) & "estado " & XTESAD & "<br>"
		
		dFechadia = CLng(Date) - CLng(CDate("01-01-1900"))
		
		strsql = strsql & "	'" & dFechadia & "','0','0',"
		
		'  nombreArchivo = SCOARCHIVO
		nombreArchivo = Trim(sCOARCHIVO& "")
		If Right(nombreArchivo,1) <> "\"  Then
			nombreArchivo=nombreArchivo & "\"
		End If
		STRSQL = STRSQL & "	'" & nombreArchivo & "', "
		
		STRSQL = STRSQL & " '" & xdesde & "',"
		
		
		STRSQL = STRSQL & "	'" & xhasta & "',"
		
		STRSQL = STRSQL & "	'" & "AUTOR" & "',1,"
		
		STRSQL = STRSQL & "	'" & XORDLISD & "'"
		
		
		STRSQL = STRSQL & "	)"
		
		'''''''' RESPONSE.WRITE  EROR_FECHA &  "22222222"
		'''''''' RESPONSE.WRITE CHR(13) & STRSQL & CHR(13) & CHR(13)
		'''''''' RESPONSE.WRITE CHR(13) & REQUEST.QUERYSTRING() & CHR(13) & CHR(13)
		
		
		
		'' RESPONSE.WRITE CHR(13) & STRSQL & CHR(13)
		Set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		'sHTML = sHTML & "<script type='text/javascript'> alert('entro aca tambien'); </script>"
		''RESPONSE.WRITE CHR(13) & Request.QueryString() & "<br>"
		STRSQL = ""
		
		SCOTIPOESQUEMA = "TODOS"
		SCOESTADO      = "TODOS"
		F_PEDIDO       = "0"
		F_GENERADO     = "0"
		F_BORRADO      = "0"
		OBSERVACIONES  = " "
		ORDENLISTADO   = "0"
		PERIODO        = "0"
		SCOARCHIVO 	=dest
		SESSION("CARGAINFO") = "RECARGAR"
	End If
	
	'---------------------------------------------------------------------------------------------------------------------
	WKACCION1 = "LISTADO AUTORIZADORES"
	Celda = Celda & ofv.GenCelda("","uti  ","","90%","","","",WKACCION1,"si")
	Celda = Celda & ofv.GenCelda("","ut   ","","2%","","","",wkproccod,"si")
	Fila  = ofv.GenRow("","","","","","","",Celda,"si")
	sHTML = sHTML & ofv.GenTabla("","aaa style='border-radius: 10px 10px 0px 0px;border-bottom:none;'","","80%","30","0","","-1","1","","",Fila,"si")
	
	'---------------------------------------------------------------------------------------------------------------------
	'''AGREGADO PARA LAS FECHAS DESDE - HASTA
	'---------------------------------------------------------------------------------------------------------------------
	
	
	sImagenE = "<B><i class='las la-calendar'STYLE=font-size:14pt;color:" & Session("CMFPNBOT") &"></i></B>"
	
	sAccion = "'javascript:void(0);'" & GTCSETCALR("frmCalendarioD","Desdex",GTCSMT,GTCFCH,GTCFCHDAT,xdesde,15,2,0,0,0)
	sinputE = "&nbsp;&nbsp;<a  class=UT href=" & sAccion
	sinputE = sinputE  & " TITLE='MODIFICAR FECHA' >" & sImagenE & "</a>"
	
	
	sAccionHasta = "'javascript:void(0);'" & GTCSETCALR("frmCalendarioH","Hastax",GTCSMT,GTCFCH,GTCFCHDAT,xhasta,4,2,0,0,0)
	sinputHasta = "&nbsp;&nbsp;<a  class=UT href=" & sAccionHasta
	sinputHasta = sinputHasta  & " TITLE='MODIFICAR FECHA' >" & sImagenE & "</a>"
	
	XFSRD = "0"
	XFSRH = "0"
	
	' CELDAS DE DESDE
	
	If  XFSRD = "0" Then
		sInput2 = "------------"
	Else
		sInput2 = CDate(CLng(XFSRD) + CLng(CDate("01-01-1900")) )
		swd = DatePart("w",sInput2)
		sInput2 = Fechador(XFSRD)
	End If
	
	Celda = ""
	Celda = Celda & ofv.gencelda("","tc","","","","","","FECHA DESDE","si")
	sInput = ofv.GenerarInput("Desde",xdesde,"hidden","","8","dt","") & ofv.GenerarInput("Desde3",fechatostr(xdesde + CLng(CDate("01-01-1900"))),"text","","40","dt","")
	Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput&sinputE,"si")
	
	Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
	
	
	' CELDAS DE HASTA
	If XFSRH = "0" Then
		sInput2 = "------------"
	Else
		sInput2 = CDate(CLng(XFSRH) + CLng(CDate("01-01-1900")) )
		swd = DatePart("w",sInput2)
		sInput2 = Fechador(XFSRH)
	End If
	
	Celda = ""
	Celda = Celda & ofv.gencelda("","tc","","","","","","FECHA HASTA","si")
	sInput = ofv.GenerarInput("Hasta",xhasta,"hidden","","8","dt","") & ofv.GenerarInput("Hasta3",fechatostr(xhasta+ CLng(CDate("01-01-1900"))),"text","","40","dt","")
	Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput&sinputHasta,"si")
	Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
	
	
	formx    = Fila
	sTARGET  = ""
	''sAccion1 = sAccion1 = "COEXPORLISAUTO.ASP?OPCION=SETLAUTO"
	'---------------------------------------------------------------------------------------------------------------------
	'ORDEN DEL LISTADO
	Celda = ""
	Celda = Celda & ofv.gencelda("","tc","","","","","","ORDEN LISTADO","si")
	
	OP1 = "A (TEMA-SUBTEMA-COD.IDENTIFICACION-ESTADO-VERSION) "
	OP2 = "B (PRODUCTO-AGRUPAMIENTO-IDENTIFICACION-ESTADO-VERSION)  "
	
	
	scomK   = "1" & Chr(9) & "2"
	scomD   = OP1 & Chr(9) & OP2
	
	sInput = ofv.gencomboFX(sForm,"XORDLIS",XORDLISD,scomK,scomD,"")
	Celda = Celda & ofv.gencelda("","VTI colspan=2","","","","","",SINPUT,"si")
	Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
	
	'---------------------------------------------------------------------------------------------------------------------
	'TIPO ESQUEMA
	Celda = ""
	Celda = Celda & ofv.gencelda("","tc","","","","","","TIPO ESQUEMA","si")
	
	Strsql = "Select * From Tipobjeto ORDER BY DESCRI "
	Set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	
	scomK = "0"
	scomD = "TODOS"
	SCOMSEP = Chr(9)
	
	Do While Not rsx1.eof
		scomK = scomK & SCOMSEP & rsx1(0)
		scomD = scomD & SCOMSEP & rsx1(1)
		SCOMSEP = Chr(9)
		rsx1.movenext
	Loop
	
	sInput = ofv.gencomboFX(sForm,"XTESQA",XTESQAD,scomK,scomD,"")
	Celda = Celda & ofv.gencelda("","VTI colspan=2","","","","","",SINPUT,"si")
	Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
	
	'---------------------------------------------------------------------------------------------------------------------
	'ESTADO
	Celda = ""
	Celda = Celda & ofv.gencelda("","tc","","","","","","ESTADOS" ,"si")
	Strsql = "Select * From ESTADOS ORDER BY estado "
	Set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	
	scomK   = "0"
	scomD   = "TODOS"
	SCOMSEP = Chr(9)
	
	Do While Not rsx1.eof
		scomK = scomK & SCOMSEP & rsx1(0)
		scomD = scomD & SCOMSEP & rsx1(1)
		SCOMSEP = Chr(9)
		rsx1.movenext
	Loop
	
	sInput = ofv.gencomboFX(sForm,"XTESA",XTESAD,scomK,scomD,"")
	Celda = Celda & ofv.gencelda("","VTI colspan=2'","","","","","",SINPUT,"si")
	Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
	'----------------------------------------------------------------------------------------
	'NOMBRE DEL ARCHIVO
	
	Celda = ""
	Celda = Celda & ofv.gencelda("","tc","","","","","","RUTA DE SALIDA","si")
	'SINPUT = ofv.GenerarInput("ORARCHIVO",SORARCHIVO,"text",sForm,"70","dt","")
	'SINPUT = ofv.GenerarInput("ORARCHIVO",SORARCHIVO,"agregar",sForm,"70","dt","no")
	SINPUT="<input name=COARCHIVO align=right size=90% value=" & SCOARCHIVO  & " type=text class=dt >"
	
	
	Celda = Celda & ofv.gencelda("","VTI colspan=2","","40%","","","",SINPUT,"si")
	Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
	
	
	'----------------------------------------------------------------------------------------
	
	Celda = ""
	FilaInput = ""
	SINPUT = "<INPUT TYPE='SUBMIT' VALUE='AGREGAR' CLASS='BT' onclick="&Chr(34)&"disp_confirm('SE AGREGARAN ESTOS PARAMETROS PARA LA EXTRACCION DE INFORMACION','" & sform & "');"&Chr(34)&">&nbsp;&nbsp;"
	Celda = Celda & ofv.gencelda("","ut colspan=3 ","center","","","","",SINPUT,"si")
	'FilaInput  = FilaInput & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
	Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
	
	sAccion1  = "COEXPORLISAUTO.ASP?OPCION=SETLAUTO"
	
	Formulario =  ofv.GenForm(sform,"",sAccion1,"","post onsubmit='return false;'","",Filax,"si")
	
	' ****** CODIGO PARA CALENDARIO
	
	sInputDesde=ofv.GenerarInput("Desdex",Xdesde+ CLng(CDate("01-01-1900")),"hidden","","20","dt","")
	SINPUT = ofv.GenerarInput("XORDLIS",XORDLISD,"hidden",sForm,"70","dt","")
	SINPUT = ofv.GenerarInput("XTESQA",XTESQAD,"hidden",sForm,"70","dt","")
	SINPUT = SINPUT & ofv.GenerarInput("XTESA",XTESAD,"hidden",sForm,"70","dt","")
	SINPUT = SINPUT & ofv.GenerarInput("hasta",Xhasta,"hidden",sForm,"70","dt","")
	sAccionRetornoD="COEXPORLISAUTO.asp?XFNAME=FD"
	formRetornoD= ofv.GenForm("frmCalendarioD","",sAccionRetornoD,STARGET,"post","",sInputDesde & SINPUT,"si")
	
	
	sInputHasta=ofv.GenerarInput("Hastax",Xhasta+ CLng(CDate("01-01-1900")),"hidden","","20","dt","")
	SINPUT = ofv.GenerarInput("XORDLIS",XORDLISD,"hidden",sForm,"70","dt","")
	SINPUT = ofv.GenerarInput("XTESQA",XTESQAD,"hidden",sForm,"70","dt","")
	SINPUT = SINPUT & ofv.GenerarInput("XTESA",XTESAD,"hidden",sForm,"70","dt","")
	SINPUT = SINPUT & ofv.GenerarInput("desde",Xdesde,"hidden",sForm,"70","dt","")
	sAccionRetornoH="COEXPORLISAUTO.asp?XFNAME=FH"
	formRetornoH= ofv.GenForm("frmCalendarioH","",sAccionRetornoH,STARGET,"post","",sInputHasta & SINPUT,"si")
	
	' ****** CODIGO PARA CALENDARIO
	
	
	' tabla= ofv.GenTabla("","aaa bgcolor='#eeeeee'  style='border-style:solid;border-width:1;border-color:#dddddd;' ","","80% ","","0","","5","3","","",filax,"si")
	'''Formulario = formulario & ofv.GenForm("Subconsulta","",sAccion,STARGET,"post" & filax,"",TABLA,"si")  & formRetornoD & formRetornoH
	Formulario = formulario &  formRetornoD & formRetornoH
	
	
	
	sHTML = sHTML & ofv.GenTabla("","aaa style='border-radius:0px;border-bottom:none;'","","80%","","0","","5","0","","",Formulario,"si")
	
	sHTML = sHTML & ""
	
	STRSQL = ""
	STRSQL = "SELECT " & Chr(13)
	STRSQL = STRSQL & "		  ID, USUARIO, " & Chr(13)
	STRSQL = STRSQL & "		  TIPOESQUEMA, " & Chr(13)
	STRSQL = STRSQL & "		  COPARAEXPORTDOCTM.ESTADO, " & Chr(13)
	STRSQL = STRSQL & "		  ESTADOS.estado, " & Chr(13)
	STRSQL = STRSQL & "		  NOMBRE_ARCHIVO,  " & Chr(13)
	STRSQL = STRSQL & "		  FDESDE, " & Chr(13)
	STRSQL = STRSQL & "		  FHASTA,  " & Chr(13)
	STRSQL = STRSQL & "		  ORDENLISTADO  " & Chr(13)
	
	STRSQL = STRSQL & "	      FROM " & Chr(13)
	STRSQL = STRSQL & "	      COPARAEXPORTDOCTM " & Chr(13)
	STRSQL = STRSQL & "	      left join ESTADOS    on (ESTADOS.OBJID = COPARAEXPORTDOCTM.ESTADO) "
	
	STRSQL = STRSQL & "	WHERE FLG_TERMINADO <> '3' and PROGRAMA = 'AUTOR'"
	''response.write strsql
	
	Set RS = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	
	If Not RS.EOF Then
		
		Celda = ""
		FilaListado = ""
		MiTabla = "<TABLE class=aaa style='border-radius: 0px 0px 10px 10px;' width=80% ><TR>"
		MiTabla = MiTabla & "<TD class=ttx align=center>*</TD>" & "<TD class=ttx>ID</TD>" & "<TD class=ttx>USUARIO</TD>"
		MiTabla = MiTabla & "<TD class=ttx>ORDEN LISTADO</TD>" & "<TD class=ttx>F PUBLIC DESDE</TD>" & "<TD class=ttx>F PUBLIC HASTA</TD>"
		MiTabla = MiTabla & "<TD class=ttx>TIPO ESQUEMA</TD>" & "<TD class=ttx>ESTADO</TD>"
		MiTabla = MiTabla & "<TD class=ttx>RUTA DE SALIDA</TD>"
		MiTabla = MiTabla & "</TR>"
		Do While Not RS.EOF
			
			'Eliminar
			If bSuprimir And rs(0) > "0" Then
				sHref = "asp/EliminarRegistro.asp?ID=" & ofv.convertircar(rs(0)," ","_")
				sHref = sHref & "&Tabla=COPARAEXPORTDOCTM&volver=" & vuelta & "&Eliminar=NSNC"
				Celda = "<TR>" & ofv.BotonEliminar(rs(1),sHref)
			Else
				Celda = "<TR>" & ofv.GenCelda("","","","","","","","&nbsp;","si")
			End If
			
			sInput = ""
			sInput = sInput & ofv.gencelda("","VTI  STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",RS("ID"),"si")
			'	sInput = sInput & "</TD>"
			Celda = Celda & sInput
			
			sInput = ""
			sInput = sInput & ofv.gencelda("","VTI  STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",RS("USUARIO"),"si")
			'	sInput = sInput & "</TD>"
			Celda = Celda & sInput
			
			sInput = ""
			''response.write ORDENLISTADO & " ORDENLISTADO " & "<br>"
			If  RS("ORDENLISTADO") = "1" Then
				xordlist= "1 (TEMA-SUBTEMA-COD.IDENTIFICACION-ESTADO-VERSION) "
			Else
				If  RS("ORDENLISTADO") = "2" Then
					xordlist= "2 (PRODUCTO-AGRUPAMIENTO-IDENTIFICACION-ESTADO-VERSION)  "
				Else
					xordlist= "1 (TEMA-SUBTEMA-COD.IDENTIFICACION-ESTADO-VERSION) "
				End If
			End If
			''''''''''''sInput = sInput & ofv.gencelda("","VTI  STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",RS("ORDENLISTADO"),"si")
			sInput = sInput & ofv.gencelda("","VTI  STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",xordlist,"si")
			
			'	sInput = sInput & "</TD>"
			Celda = Celda & sInput
			
			sInput = ""
			''''xfecha= RS("FDESDE")
			xfecha = fechatostr(CLng(rs("FDESDE")) + CLng(CDate("01-01-1900")))
			sInput = sInput & ofv.gencelda("","VTI  STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",xfecha,"si")
			'	sInput = sInput & "</TD>"
			Celda = Celda & sInput
			
			sInput =  ""
			''''fecha= RS("FHASTA")
			xfecha = fechatostr(CLng(rs("FHASTA")) + CLng(CDate("01-01-1900")))
			sInput = sInput & ofv.gencelda("","VTI  STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",xfecha,"si")
			'sInput = sInput & "</TD>"
			Celda = Celda & sInput
			
			sInput = ""
			''response.write TIPOESQUEMA & " TIPOESQUEMA " & "<br>"
			If RS("TIPOESQUEMA") < "1" Then
				sInput = sInput & ofv.gencelda("","VTI  STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","","TODOS","si")
			Else
				sInput = sInput & ofv.gencelda("","VTI  STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",RS("TIPOESQUEMA"),"si")
			End If
			'	sInput = sInput & "</TD>"
			Celda = Celda & sInput
			
			sInput = ""
			''response.write ESTADO & " ESTADO " & "<br>"
			If Len(RS("ESTADO") & "") < 1  Then
				sInput = sInput & ofv.gencelda("","VTI  STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","","TODOS","si")
			Else
				'''''''sInput = sInput & ofv.gencelda("","VTI  STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",RS("ESTADO"),"si")
				sInput = sInput & ofv.gencelda("","VTI  STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",RS(4),"si")
				
			End If
			'	sInput = sInput & "</TD>"
			
			Celda = Celda & sInput
			sInput = ""
			sInput = sInput & ofv.gencelda("","VTI  STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",RS("NOMBRE_ARCHIVO"),"si")
			'	sInput = sInput & "</TD>"
			Celda = Celda & sInput
			
			Celda = Celda &  "</TR>"
			
			FilaListado  = FilaListado & Celda
			RS.MOVENEXT
		Loop
	End If
	'' RESPONSE.WRITE CHR(13) & FilaListado & CHR(13)
	
	FormularioListado =  FormularioListado & MiTabla & FilaListado & "</TABLE>"
	
	sHTML = sHTML & ofv.GenTabla("","    ","","100% ","","0","","5","0","","",FormularioListado,"si")
	
	'Botonera
	Agregar = ""
	Buscar  = ""
	Celda   = ""
	
	sHTML = sHTML & ofv.botoneraESSX("0","",cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,"")
	
	
	
	'----------------------------------------------------------------------------------------------
	'AGREGADO DE LOS ARCHIVOS YA INGRESADOS PARA LA BÚSQUEDA
	'----------------------------------------------------------------------------------------------
	
	sHTML = sHTML & Chr(13) & "<script type='text/javascript'>" & Chr(13)
	sHTML = sHTML & "function disp_confirm(disp_msg,disp_frm)" & Chr(13)
	sHTML = sHTML & "  {" & Chr(13)
	sHTML = sHTML & "  var r=confirm(disp_msg);var fr= document.getElementById(disp_frm);" & Chr(13)
	sHTML = sHTML & "  if (r==true)" & Chr(13)
	sHTML = sHTML & "    {" & Chr(13)
	sHTML = sHTML & "		fr.action='COEXPORLISAUTO.ASP?OPCION=SETLAUTO&OPCIONINS=INSERTAR'; "
	sHTML = sHTML & "    	fr.submit();" & Chr(13)
	sHTML = sHTML & "    }" & Chr(13)
	sHTML = sHTML & "  else" & Chr(13)
	sHTML = sHTML & "    {" & Chr(13)
	sHTML = sHTML & "    alert('Se cancelo la accion!');" & Chr(13)
	sHTML = sHTML & "    return false;" & Chr(13)
	sHTML = sHTML & "    }" & Chr(13)
	sHTML = sHTML & "  }" & Chr(13)
	
	''    sHTML = sHTML & "  }"
	sHTML = sHTML & "</script>" & Chr(13)
	
	
	
	
	
	sHTML = sHTML & "</script>" & GTCCALENDAR("")
	
	sHTML = sHTML & "</body></html>"
	
	Call ofv.cerrarconn(cn)
	Call ofv.cerrarconn(rsx1)
	
End If

response.write sHTML

'------------------------------------------------------------

Function reemplazaTexto(texto, patrn, replStr)
	Dim regEx
	Set regEx = New RegExp            ' Crea una expresión regular.
	regEx.Pattern = patrn            ' Establece el modelo.
	regEx.IgnoreCase = True            ' Hace distinción entre mayúsculas y minúsculas.
	regEx.global = True
	reemplazaTexto = regEx.Replace(texto, replStr)   ' Realiza el reemplazo.
	Set regEx = Nothing
End Function

'------------------------------------------------------------

Function EliminarTexto(TEXTO)
	
	If Len(texto) <> 0 Then
		texto = reemplazaTexto(texto,"[ /:*?<>|\\]","_")
	End If
	EliminarTexto = texto
End Function

'------------------------------------------------------------

Function Fechador(wkfecha)
	
	
	fec = CDate(wkfecha + CLng(CDate("01-01-1900")))
	ndia = ofv.cdia(Weekday(fec))
	nMes = ofv.cmes(Month(fec))
	
	Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)
	
	
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

