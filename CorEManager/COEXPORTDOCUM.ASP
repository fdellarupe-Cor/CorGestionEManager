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
	nFecini = CLng(CDate("01-01-1900"))
	Set cn = ofv.conectar(ofv.strconn0)
	'set rsx1 = ofv.conectar(ofv.strconn0)
	
	sForm = "formlis"
	Filax = ""
	
	'RESPONSE.WRITE REQUEST.QUERYSTRING("OPCION")
	sopcion=REQUEST.QUERYSTRING("OPCION")
	' RESPONSE.WRITE SOPCION & "222"
	
	
	xfiltrod    = "" & request.FORM("xfiltro")
	xtemad      = "" & request.FORM("xtema")
	xsubtemad   = "" & request.FORM("xsubtema")
	XTESQAD     = "" & request.FORM("XTESQA")
	XTESAD      = "" & request.FORM("XTESA")
	XPRDSRVD    = "" & request.FORM("XPRDSRV")
	XAGRUPD     = "" & request.FORM("XAGRUP")
	XSEGMD      = "" & request.FORM("XSEGM")
	
	
	If xfiltrod  = "" Then xfiltrod  = "1"
	If xfiltrod  = "0" Then xfiltrod  = "1"
	If xtemad    = "" Then xtemad    = "0"
	If xsubtemad = "" Then xsubtemad = "0"
	If xtesqad   = "" Then xtesqad   = "0"
	If xtesad    = "" Then xtesad    = "0"
	If xprdsrvd  = "" Then xprdsrvd  = "0"
	If xagrupd   = "" Then xagrupd   = "0"
	If xsegmd    = "" Then xsegmd    = "0"
	DEST="C:\exportarBpn\"
	'' Por Tema-Subtema
	If sopcion ="SETTEMA"  Then
		SCOTEMASUBTEMA    	= UCase(request.FORM("COTEMASUBTEMA"))
	Else
		
		SCOTEMASUBTEMA    	= "TODOS"  'UCASE(request.QUERYSTRING("SCOTEMASUBTEMA"))
	End If
	
	If sopcion ="SETTEMA"  Then
		SCOSUBTEMA    		= UCase(request.FORM("COSUBTEMA"))
	Else
		SCOSUBTEMA         	= "TODOS"  'UCASE(request.QUERYSTRING("SCOSUBTEMA"))
	End If
	
	If sopcion ="SETTEMA"  Then
		SCOTIPOESQUEMA    	= UCase(request.FORM("COTIPOESQUEMA"))
	Else
		SCOTIPOESQUEMA		= "TODOS"  'UCASE(request.QUERYSTRING("SCOTIPOESQUEMA"))
	End If
	
	If sopcion ="SETTEMA"  Then
		SCOESTADOS    		= UCase(request.FORM("COESTADOS"))
	Else
		SCOESTADOS	        = "TODOS"  'UCASE(request.QUERYSTRING("SCOESTADOS"))
	End If
	
	If sopcion ="SETTEMA"   Then
		SCOARCHIVO    		= (request.FORM("COARCHIVO"))
	Else
		SCOARCHIVO     		= DEST 'UCASE(request.QUERYSTRING("SCORARCHIVO"))
	End If
	
	'' Por Producto-Servicio
	
	If sopcion ="SETPRDS"  Then
		SCOPRODSERV   		= UCase(request.FORM("COPRODSERV"))
	Else
		SCOPRODSERV			= "TODOS"  'UCASE(request.QUERYSTRING("SCOPRODSERV"))
	End If
	
	If sopcion ="SETPRDS"  Then
		SCOAGRUPAM    		= UCase(request.FORM("COAGRUPAM"))
	Else
		SCOAGRUPAM		    = "TODOS"  'UCASE(request.QUERYSTRING("SCOAGRUPAM"))
	End If
	
	If sopcion ="SETPRDS"  Then
		SCOSEGMENTO    		= UCase(request.FORM("COSEGMENTO"))
	Else
		SCOSEGMENTO     	= "TODOS"  'UCASE(request.QUERYSTRING("SCOSEGMENTO"))
	End If
	
	If sopcion ="SETPRDS"  Then
		SCOTIPOESQUEMA    	= UCase(request.FORM("COTIPOESQUEMA"))
	Else
		SCOTIPOESQUEMA 		= "TODOS"  'UCASE(request.QUERYSTRING("SCOTIPOESQUEMA"))
	End If
	
	If sopcion ="SETPRDS"  Then
		SCOESTADOS    		= UCase(request.FORM("COESTADOS"))
	Else
		SCOESTADOS      	= "TODOS"  'UCASE(request.QUERYSTRING("SCOESTADOS"))
	End If
	
	If request.FORM("COARCHIVO") <> ""   Then
		session("SCOARCHIVO")= (request.FORM("COARCHIVO"))
		
	Else
		
		SCOARCHIVO     		= session("SCOARCHIVO")
	End If
	
	If SCOARCHIVO= "" Then SCOARCHIVO=dest
	
	'RESPONSE.WRITE UCASE(request.FORM("COARCHIVO"))
	
	
	sHTML = ofv.FormHeader(session("FormN")) & Chr(13)
	sHTML = sHTML & "<CENTER>"
	
	'---------------------------------------------------------------------------------------------------------------------
	
	If REQUEST.QUERYSTRING("OPCIONINS") = "INSERTAR" Then
		
		
		EROR_FECHA = ""
		
		''      response.write request.querystring
		STRSQL = "INSERT INTO " & Chr(13)
		STRSQL = STRSQL & "	COPARAEXPORTDOCTM (USUARIO, TEMASUBTEMA, SUBTEMA, TIPOESQUEMA, ESTADO, " & Chr(13)
		STRSQL = STRSQL & "	PRODSERV, AGRUPAM, SEGMENTO, " & Chr(13)
		STRSQL = STRSQL & "	F_PEDIDO, F_GENERADO, F_BORRADO, OBSERVACIONES, PERIODO," & Chr(13)
		STRSQL = STRSQL & "	NOMBRE_ARCHIVO, FDESDE, FHASTA, PROGRAMA ,FLG_TERMINADO" & Chr(13)
		STRSQL = STRSQL & "	)" & Chr(13)
		STRSQL = STRSQL & " VALUES (" & Chr(13)
		
		STRSQL = STRSQL & "	'" & SESSION("USUARIO") & "'," & Chr(13)
		'''RESPONSE.WRITE CHR(13) & SESSION("USUARIO") & CHR(13)
		
		If XTEMAD = "TODOS" Then
			STRSQL = STRSQL & "		0," & Chr(13)
		Else
			STRSQL = STRSQL & "		'" & XTEMAD & "'," & Chr(13)
		End If
		'''RESPONSE.WRITE CHR(13) & "Tema " & XTEMAD & CHR(13)
		
		If XSUBTEMAD = "TODOS" Then
			STRSQL = STRSQL & "		0," & Chr(13)
		Else
			STRSQL = STRSQL & "		'" & XSUBTEMAD & "'," & Chr(13)
		End If
		'''RESPONSE.WRITE CHR(13) & "Subtema " & XSUBTEMAD & CHR(13)
		
		If XTESQAD = "TODOS" Then
			STRSQL = STRSQL & "		0," & Chr(13)
		Else
			STRSQL = STRSQL & "		'" & XTESQAD & "'," & Chr(13)
		End If
		'''RESPONSE.WRITE CHR(13) & "esquema " & XTESQAD & CHR(13)
		
		If XTESAD  = "TODOS" Then
			STRSQL = STRSQL & "		0," & Chr(13)
		Else
			STRSQL = STRSQL & "		'" & XTESAD & "'," & Chr(13)
		End If
		'''RESPONSE.WRITE CHR(13) & "estado " & XTESAD & CHR(13)
		
		If XPRDSRVD = "TODOS" Then
			STRSQL = STRSQL & "		0," & Chr(13)
		Else
			STRSQL = STRSQL & "		'" & XPRDSRVD & "'," & Chr(13)
		End If
		'''RESPONSE.WRITE CHR(13) & XPRDSRVD & CHR(13)
		
		If XAGRUPD = "TODOS" Then
			STRSQL = STRSQL & "		0," & Chr(13)
		Else
			STRSQL = STRSQL & "		'" & XAGRUPD & "'," & Chr(13)
		End If
		'''RESPONSE.WRITE CHR(13) & XAGRUPD & CHR(13)
		
		If XSEGMD = "TODOS" Then
			STRSQL = STRSQL & "		0," & Chr(13)
		Else
			STRSQL = STRSQL & "		'" & XSEGMD & "'," & Chr(13)
		End If
		'''RESPONSE.WRITE CHR(13) & XSEGMD & CHR(13)
		
		'''RESPONSE.WRITE CHR(13) & xfiltrod & CHR(13)
		
		dFechadia = CLng(Date) - CLng(CDate("01-01-1900"))
		
		strsql = strsql & dFechadia & "," &  "'0'," & "'0'," & "' '," & xfiltrod &  ","   & Chr(13)
		
		nombreArchivo = Trim(sCOARCHIVO& "")
		If Right(nombreArchivo,1) <> "\"  Then
			nombreArchivo=nombreArchivo & "\"
		End If
		
		STRSQL = STRSQL & "		'" & nombreArchivo & "', " & Chr(13)
		
		strsql = strsql & "'0'," & "'0'," & Chr(13)
		
		STRSQL = STRSQL & "	'" & "EXPORT" & "'"
		
		
		STRSQL = STRSQL & ",1	)" & Chr(13)
		
		'''''''' RESPONSE.WRITE  EROR_FECHA &  "22222222"
		'''''''' RESPONSE.WRITE CHR(13) & STRSQL & CHR(13) & CHR(13)
		'''''''' RESPONSE.WRITE CHR(13) & REQUEST.QUERYSTRING() & CHR(13) & CHR(13)
		
		
		If EROR_FECHA = "ERROR" Then
			STRSQL = ""
			sHTML = sHTML & "<script type='text/javascript'> alert('POR FAVOR VERIFIQUE LAS FECHAS'); </script>"
			
		Else
			''''RESPONSE.WRITE CHR(13) & STRSQL & CHR(13)
			Set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			'sHTML = sHTML & "<script type='text/javascript'> alert('entro aca tambien'); </script>"
			''RESPONSE.WRITE CHR(13) & Request.QueryString() & CHR(13)
			STRSQL = ""
			
			SCOTEMASUBTEMA = "TODOS"
			SCOSUBTEMA     = "TODOS"
			SCOTIPOESQUEMA = "TODOS"
			SCOESTADO      = "TODOS"
			SCOPRODSERV    = "TODOS"
			SCOAGRUPAM     = "TODOS"
			SCOSEGMENTO    = "TODOS"
			F_PEDIDO       = "0"
			F_GENERADO     = "0"
			F_BORRADO      = "0"
			OBSERVACIONES  = " "
			PERIODO        = "0"
			SCOARCHIVO 	= "C:\exportarBpn\"
			SESSION("CARGAINFO") = "RECARGAR"
		End If
	End If
	
	'---------------------------------------------------------------------------------------------------------------------
	WKACCION1 = "EXPORTACION DE DOCUMENTOS"
	Celda = Celda & ofv.GenCelda("","ttt  ","","100%","","","",WKACCION1,"si")
	Celda = Celda & ofv.GenCelda("","ut   ","","2%","","","",wkproccod,"si")
	Fila  = ofv.GenRow("","","","","","","",Celda,"si")
	sHTML = sHTML & ofv.GenTabla("","aaa style='border-radius: 10px 10px 0px 0px;border-bottom:none;'","","80% ","30","0","","-1","1","","",Fila,"si")
	sHTML = sHTML & ""
	'---------------------------------------------------------------------------------------------------------------------
	'TIPO EXPORTACION
	Celda = ""
	Celda = Celda & ofv.gencelda("","tc","","","","","","TIPO EXPORTACION","si")
	
	OP1 = "Por Tema - Subtema "
	OP2 = "Por Agrupamiento "
	
	
	scomK   = "1" & Chr(9) & "2"
	scomD   = OP1 & Chr(9) & OP2
	
	sInput = ofv.gencomboFX(sForm,"xfiltro",xfiltrod,scomK,scomD,"")
	Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")
	Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
	'---------------------------------------------------------------------------------------------------------------------
	If xfiltrod = "1" Then
		'TEMA
		Celda = ""
		Celda = Celda & ofv.gencelda("","tc","","","","","","TEMA","si")
		Strsql = "Select * From TEMA ORDER BY DESCRI "
		'''response.write Strsql
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
		
		sInput = ofv.gencomboFX(sForm,"XTEMA",XTEMAD,scomK,scomD,"")
		Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")
		Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
	End If
	'---------------------------------------------------------------------------------------------------------------------
	
	'''response.write XTEMAD & " XTEMAD " & "<br>"
	
	If xfiltrod = "1" Then
		'SUBTEMA
		Celda = ""
		Celda =  ofv.gencelda("","tc","","","","","","SUBTEMA","si")
		If XTEMAD  = "0" Then
			scomK   = "0"
			scomD   = "TODOS"
			ScOMSEP = Chr(9)
		Else
			
			''''''Strsql = "Select * From SUBTEMA where tipobjeto = '" & XTEMAD & "' ORDER BY DESCRI "
			Strsql = "Select * From PRODSERV where codtema = '" & XTEMAD & "' ORDER BY DESCRI "
			Set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			''''''response.write Strsql & "<br>"
			
			scomK   = "0"
			scomD   = "TODOS"
			SCOMSEP = Chr(9)
			
			Do While Not rsx1.eof
				scomK = scomK & SCOMSEP & rsx1(0)   ' codprodserv
				scomD = scomD & SCOMSEP & rsx1(1)   ' descripcion
				SCOMSEP = Chr(9)
				rsx1.movenext
			Loop
		End If
		
		sInput = ofv.gencomboFX(sForm,"XSUBTEMA",XSUBTEMAD,scomK,scomD,"")
		Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","40%","","","",SINPUT,"si")
		Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
	End If
	
	''' response.write XSUBTEMAD & " XSUBTEMAD " & "<br>"
	
	'---------------------------------------------------------------------------------------------------------------------
	If xfiltrod = "2" Then
		'PRODUCTO / SERVICIO
		Celda = ""
		Celda = Celda & ofv.gencelda("","tc","","","","","","PRODUCTO/SERVICIO","si")
		Strsql = "Select * From TIPOPRODUCTO ORDER BY DESCRI "
		
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
		
		sInput = ofv.gencomboFX(sForm,"XPRDSRV",XPRDSRVD,scomK,scomD,"")
		Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")
		Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
	End If
	'---------------------------------------------------------------------------------------------------------------------
	If xfiltrod = "2" Then
		'AGRUPAMIENTO
		Celda = ""
		Celda = Celda & ofv.gencelda("","tc","","","","","","AGRUPAMIENTO","si")
		Strsql = "Select * From TipoAgrupa Order By descri "
		
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
		
		sInput = ofv.gencomboFX(sForm,"XAGRUP",XAGRUPD,scomK,scomD,"")
		Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")
		Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
	End If
	'----------------------------------------------------------------------------------------
	If xfiltrod = "2" Then
		'SEGMENTO
		Celda = ""
		Celda = Celda & ofv.gencelda("","tc","","","","","","SEGMENTO","si")
		
		Strsql = "Select * From TIPOSEGMENTO Order By descri"
		
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
		
		sInput = ofv.gencomboFX(sForm,"XSEGM",XSEGMD,scomK,scomD,"")
		Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")
		Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
	End If
	'---------------------------------------------------------------------------------------------------------------------
	'TIPO ESQUEMA
	Celda = ""
	Celda = Celda & ofv.gencelda("","tc","","","","","","TIPO ESQUEMA","si")
	
	Strsql = "Select * From Tipobjeto where tipocpt ='D' and aplic in ('GLB','EM')  ORDER BY DESCRI "
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
	Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")
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
	Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",SINPUT,"si")
	Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
	'----------------------------------------------------------------------------------------
	'NOMBRE DEL ARCHIVO
	
	Celda = ""
	Celda = Celda & ofv.gencelda("","tc","","","","","","RUTA DE SALIDA","si")
	'SINPUT = ofv.GenerarInput("ORARCHIVO",SORARCHIVO,"text",sForm,"70","dt","")
	'SINPUT = ofv.GenerarInput("ORARCHIVO",SORARCHIVO,"agregar",sForm,"70","dt","no")
	SINPUT="<input name=COARCHIVO id=COARCHIVO align=right size=90% value='" & SCOARCHIVO  & "' type=text class=dt >"
	
	
	Celda = Celda & ofv.gencelda("","VTI colspan=2 STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","40%","","","",SINPUT,"si")
	Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
	
	
	'----------------------------------------------------------------------------------------
	
	Celda = ""
	FilaInput = ""
	SINPUT = "<INPUT TYPE='SUBMIT' VALUE='AGREGAR' CLASS='BT' onclick="&Chr(34)&"disp_confirm('SE AGREGARAN ESTOS PARAMETROS PARA LA EXTRACCION DE INFORMACION','" & sform & "');"&Chr(34)&">&nbsp;&nbsp;"
	Celda = Celda & ofv.gencelda("","ut colspan=3'","center","","","","",SINPUT,"si")
	'FilaInput  = FilaInput & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
	Filax  = Filax & ofv.genrow("","AAA bgcolor=" & xTrColor & " ","","","","","",Celda,"si")
	
	sAccion1 = "COEXPORTDOCUM.ASP?OPCION=SETTEMA"
	
	Formulario =  ofv.GenForm(sform,"",sAccion1,"","post onsubmit='return false;'","",Filax,"si")
	
	sHTML = sHTML & ofv.GenTabla("","aaa style='border-radius:0px;border-bottom:none;'","","80%","","0","","5","0","","",Formulario,"si")
	
	sHTML = sHTML & ""
	
	STRSQL = ""
	STRSQL = "SELECT " & Chr(13)
	STRSQL = STRSQL & "		  ID, USUARIO, " & Chr(13)
	STRSQL = STRSQL & "		  TEMASUBTEMA, " & Chr(13)
	STRSQL = STRSQL & "		  TEMA.descri, " & Chr(13)
	STRSQL = STRSQL & "		  SUBTEMA, " & Chr(13)
	STRSQL = STRSQL & "		  A.descri, " & Chr(13)
	STRSQL = STRSQL & "		  TIPOESQUEMA, " & Chr(13)
	STRSQL = STRSQL & "		  TIPOBJETO.descri, " & Chr(13)
	STRSQL = STRSQL & "		  COPARAEXPORTDOCTM.ESTADO, " & Chr(13)
	STRSQL = STRSQL & "		  ESTADOS.estado, " & Chr(13)
	STRSQL = STRSQL & "		  PERIODO, " & Chr(13)
	STRSQL = STRSQL & "	      PRODSERV, " & Chr(13)
	STRSQL = STRSQL & "	      B.descri, " & Chr(13)
	STRSQL = STRSQL & "	      AGRUPAM, " & Chr(13)
	STRSQL = STRSQL & "	      TIPOAGRUPA.descri, " & Chr(13)
	STRSQL = STRSQL & "	      SEGMENTO, " & Chr(13)
	STRSQL = STRSQL & "	      TIPOSEGMENTO.descri, " & Chr(13)
	STRSQL = STRSQL & "		  NOMBRE_ARCHIVO  " & Chr(13)
	
	STRSQL = STRSQL & "	      FROM " & Chr(13)
	STRSQL = STRSQL & "	      COPARAEXPORTDOCTM " & Chr(13)
	
	STRSQL = STRSQL & "	      left join TEMA        on (TEMA.CODTEMA=COPARAEXPORTDOCTM.TEMASUBTEMA) "
	STRSQL = STRSQL & "	      left join TIPOBJETO   on (TIPOBJETO.TIPOBJETO = COPARAEXPORTDOCTM.TIPOESQUEMA) "
	STRSQL = STRSQL & "	      left join PRODSERV A  on (A.codprodserv= COPARAEXPORTDOCTM.SUBTEMA) "
	STRSQL = STRSQL & "	      left join TIPOPRODUCTO B on (B.OBJID = COPARAEXPORTDOCTM.PRODSERV) "
	STRSQL = STRSQL & "	      left join ESTADOS     on (ESTADOS.OBJID = COPARAEXPORTDOCTM.ESTADO) "
	STRSQL = STRSQL & "	      left join TIPOAGRUPA  on (TIPOAGRUPA.OBJID = COPARAEXPORTDOCTM.AGRUPAM) "
	STRSQL = STRSQL & "	      left join TIPOSEGMENTO on (TIPOSEGMENTO.OBJID = COPARAEXPORTDOCTM.SEGMENTO) "
	
	STRSQL = STRSQL & "	WHERE FLG_TERMINADO <> '3' and PROGRAMA = 'EXPORT'"
	''''response.write strsql
	
	Set RS = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	
	If Not RS.EOF Then
		
		Celda = ""
		FilaListado = ""
		MiTabla = "<TABLE class=aaa width=80% style='border-radius: 0px 0px 10px 10px;'><TR>"
		MiTabla = MiTabla & "<TD class=ttx align=center>*</TD>" & "<TD class=ttx>ID</TD>" & "<TD class=ttx>USUARIO</TD>"
		MiTabla = MiTabla & "<TD class=ttx>TEMA</TD>" & "<TD class=ttx>SUBTEMA</TD>" & "<TD class=ttx>TIPO ESQUEMA</TD>" & "<TD class=ttx>ESTADO</TD>"
		MiTabla = MiTabla & "<TD class=ttx>TIPO EXP</TD>" & "<TD class=ttx>PRODSERV</TD>"
		MiTabla = MiTabla & "<TD class=ttx>AGRUPAM</TD>" & "<TD class=ttx>SEGMENTO</TD>" & "<TD class=ttx>RUTA DE SALIDA</TD>"
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
			
			sInput = "<TD><TABLE><TR><TD>"
			sInput = sInput & ofv.gencelda("","VTI","","","","","",RS("ID"),"si")
			sInput = sInput & "</TD></TR></TABLE></TD>"
			Celda = Celda & sInput
			
			sInput = "<TD><TABLE><TR><TD>"
			sInput = sInput & ofv.gencelda("","VTI","","","","","",RS("USUARIO"),"si")
			sInput = sInput & "</TD></TR></TABLE></TD>"
			Celda = Celda & sInput
			
			sInput = "<TD><TABLE><TR>"
			
			Set rsx1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			
			'''''''response.write TEMASUBTEMA & " TEMASUBTEMA " & "<br>"
			If RS("TEMASUBTEMA") < 1 Then
				If  RS("PERIODO") = 2 Then
					sInput = sInput & ofv.gencelda("","VTI","","","","","","0","si")
				Else
					sInput = sInput & ofv.gencelda("","VTI","","","","","","TODOS","si")
				End If
			Else
				''''sInput = sInput & ofv.gencelda("","VTI  STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",RS("TEMASUBTEMA"),"si")
				sInput = sInput & ofv.gencelda("","VTI","","","","","",RS(3),"si")
			End If
			
			sInput = sInput & "</TR></TABLE></TD>"
			Celda = Celda & sInput
			
			sInput = "<TD><TABLE><TR>"
			
			If RS("SUBTEMA") <  "1" Then
				If  RS("PERIODO") = 2 Then
					sInput = sInput & ofv.gencelda("","VTI","","","","","","0","si")
				Else
					sInput = sInput & ofv.gencelda("","VTI","","","","","","TODOS","si")
				End If
			Else
				'''''  sInput = sInput & ofv.gencelda("","VTI  STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",RS("SUBTEMA"),"si")
				sInput = sInput & ofv.gencelda("","VTI","","","","","",RS(5),"si")
			End If
			
			sInput = sInput & "</TR></TABLE></TD>"
			Celda = Celda & sInput
			
			sInput = "<TD><TABLE><TR>"
			'''response.write TIPOESQUEMA & " TIPOESQUEMA " & "<br>"
			If RS("TIPOESQUEMA") < "1" Then
				sInput = sInput & ofv.gencelda("","VTI","","","","","","TODOS","si")
			Else
				''''''''''''''''sInput = sInput & ofv.gencelda("","VTI  STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",RS("TIPOESQUEMA"),"si")
				sInput = sInput & ofv.gencelda("","VTI","","","","","",RS(7),"si")
				
			End If
			sInput = sInput & "</TR></TABLE></TD>"
			Celda = Celda & sInput
			
			sInput = "<TD><TABLE><TR>"
			''''response.write ESTADO & " ESTADO " & "<br>"
			If Len(RS("ESTADO") & "") < "1"  Then
				sInput = sInput & ofv.gencelda("","VTI","","","","","","TODOS","si")
			Else
				sInput = sInput & ofv.gencelda("","VTI","","","","","",RS("ESTADO"),"si")
			End If
			sInput = sInput & "</TR></TABLE></TD>"
			Celda = Celda & sInput
			
			sInput = "<TD><TABLE><TR>"
			
			''''response.write PERIODO & " PERIODO " & "<br>"
			If RS("PERIODO") = 1 Then
				xPERIODO= "Por Tema-Subtema"
			Else
				If RS("PERIODO") = 2 Then
					xPERIODO= "Por Agrupamiento"
				Else
					xPERIODO= "Por Tema-Subtema"
				End If
			End If
			sInput = sInput & ofv.gencelda("","VTI","","","","","",xperiodo,"si")
			
			sInput = sInput & "</TR></TABLE></TD>"
			
			Celda = Celda & sInput
			sInput = "<TD><TABLE><TR>"
			If  RS("PRODSERV") < 1 Then
				If  RS("PERIODO") = 1 Then
					sInput = sInput & ofv.gencelda("","VTI","","","","","","0","si")
				Else
					sInput = sInput & ofv.gencelda("","VTI","","","","","","TODOS","si")
				End If
			Else
				'''''sInput = sInput & ofv.gencelda("","VTI  STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",RS("PRODSERV"),"si")
				sInput = sInput & ofv.gencelda("","VTI","","","","","",RS(12),"si")
				
			End If
			sInput = sInput & "</TR></TABLE></TD>"
			
			Celda = Celda & sInput
			sInput = "<TD><TABLE><TR>"
			If  RS("AGRUPAM") < 1 Then
				If  RS("PERIODO") = 1 Then
					sInput = sInput & ofv.gencelda("","VTI","","","","","","0","si")
				Else
					sInput = sInput & ofv.gencelda("","VTI","","","","","","TODOS","si")
				End If
			Else
				'''''''''''sInput = sInput & ofv.gencelda("","VTI  STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",RS("AGRUPAM"),"si")
				sInput = sInput & ofv.gencelda("","VTI","","","","","",RS(14),"si")
			End If
			sInput = sInput & "</TR></TABLE></TD>"
			
			Celda = Celda & sInput
			sInput = "<TD><TABLE><TR>"
			If RS("SEGMENTO") < 1 Then
				If  RS("PERIODO") = 1 Then
					sInput = sInput & ofv.gencelda("","VTI","","","","","","0","si")
				Else
					sInput = sInput & ofv.gencelda("","VTI","","","","","","TODOS","si")
				End If
			Else
				'''''''''''sInput = sInput & ofv.gencelda("","VTI  STYLE='border-style:SOLID;border-width:1;border-color:#888888;'","","","","","",RS("SEGMENTO"),"si")
				sInput = sInput & ofv.gencelda("","VTI","","","","","",RS(16),"si")
			End If
			sInput = sInput & "</TR></TABLE></TD>"
			
			Celda = Celda & sInput
			sInput = "<TD><TABLE><TR>"
			sInput = sInput & ofv.gencelda("","VTI","","","","","",RS("NOMBRE_ARCHIVO"),"si")
			sInput = sInput & "</TR></TABLE></TD>"
			Celda = Celda & sInput
			
			
			FilaListado  = FilaListado & Celda
			RS.MOVENEXT
		Loop
	End If
	'' RESPONSE.WRITE CHR(13) & FilaListado & CHR(13)
	
	FormularioListado =  FormularioListado & MiTabla & FilaListado & "</TR></TABLE>"
	
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
	sHTML = sHTML & "  if (r==true){" & Chr(13)
	
	sHTML = sHTML & "		fr.action='COEXPORTDOCUM.ASP?OPCION=SETTEMA&OPCIONINS=INSERTAR'; "
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
	
	sHTML = sHTML & "</CENTER>"
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

%>