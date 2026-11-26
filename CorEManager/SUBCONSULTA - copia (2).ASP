<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
if ok <> "Y" then
	uv = ofv.uv
	if uv = "NO" then ok = ofv.ValidarUsuario
	sHTML = ok
else



	session("sqlcons")=""
	session("ModMain") = "CL"

	response.expires=-1441
	Response.AddHeader "PRAGMA", "NO-CACHE"

	sMenu = request.querystring("Menu")
	sNombre = request.querystring("Nombre")


	'response.write request.querystring
	''response.write sNombre & " sNombre " & chr(13)

	sTema = request.form("codtema")
	if sTema = "" or sTema = "nsnc" then sTema = 0
	sSubtema = request.form("codprodserv")
	if sSubtema = "" or sSubtema = "nsnc" then sSubtema = 0
	sEsquema = request.form("codobjeto")
	if sEsquema = "" or sEsquema = "nsnc" then sEsquema = 0
	xComp = ""
	xComp = request.form("XComplem")


	base=request.querystring("BASE")

    IF base <> "" THEN
        SESSION("BASE")=base
    END IF





	XXXX = NEWEMPRE("NO")

	session("ASP")="SUBCONSULTA.ASP"

	sEMPRE = request.querystring("EMPRE")

	IF SEMPRE = ""  AND SESSION("UN") = "" THEN

		SEMPRE = ATREMP(0, 0)
		SESSION("UN") = SEMPRE & ""
		If (SESSION("UN") & "") = "" Then SESSION("UN") = 1
	elseIF sEMPRE <> "" THEN
		XXXX = NEWEMPRE(sEMPRE)
		SESSION("UN")=SEMPRE
	ELSE
		XXXX = NEWEMPRE(SESSION("UN"))
		sempre=SESSION("UN")
	end if

	set cn = ofv.conectar(ofv.strconn2)


	sNombren = sNombre

	''response.write sNombre & " sNombre linea 63 " & chr(13)


	sNombre = "Consulta por " & sNombre
	snombre = ucase(snombre)

	sHTML = ofv.FormHeader(session("FormN")) & chr(13)

	sImagenM = "<B><FONT FACE=""CorFlatIcon"" COLOR=#005500 SIZE=3>2</FONT></B>"
	sImagenI2 = "<B><FONT FACE=""CorFlatIcon"" COLOR=#f7aa22 SIZE=3>&#0128;</FONT></B>"
	sImagenI = "<B><FONT FACE=""CorFlatIcon"" COLOR=#f7aa22 SIZE=2>u</FONT></B>"
	sImagenI3 = "<B><FONT FACE=""CorFlatIcon"" COLOR=#6699CC SIZE=2>=</FONT></B>"
	sImagenEM = "<B><FONT FACE=CorFlatIcon COLOR=#F7AA22 SIZE=3>&#0252;</FONT></B>"
	sImagen4si = "<B><FONT FACE=CorFlatIcon COLOR=#993333 SIZE=3>4</FONT></B>"
	sImagenx = "<B><FONT FACE=CorFlatIcon COLOR=#aa0000 SIZE=4>x</FONT></B>"


	Fila    = ""
	Celda   = ""
	sHref   = ""



	IF  Session("EMPMUN") = "S" THEN

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





		shtml=shtml & "<script language=javascript>" & chr(13)
		shtml=shtml & "function shwDem(){" & chr(13)
		shtml=shtml & "document.all.DEM1.style.display='block';" & chr(13)
		shtml=shtml & "}" & chr(13)
		shtml=shtml & "</script>" & chr(13)
		'----------------------------------------------
		sinput2 = "<a  style='text-decoration:none;' onmouseover='this.style.textDecoration=" & chr(34)
		sinput2 = sinput2  & "underline" & chr(34) & ";' onmouseout='this.style.textDecoration=" & chr(34)
		sinput2 = sinput2  & "none" & chr(34) & ";' href='javascript:void(0);' ONCLICK=" & chr(34) & "document.all.DEM1.style.display='none';" & chr(34)
		sinput2 = sinput2  & " TITLE='OCULTAR' ><b>" & sImagenx & "</b></a>"
		TABLA = Tabla & "<DIV id='DEM1' STYLE='display:none;z-index:1;margin-top:1;margin-left:1;background-color:#ffffff;position:absolute;left:1px;top:240px;left:20px;height:250px;width:350px;border-style:inset;border-color:#f7aa22;border-width:2;overflow:auto;'>"


		TABLAX = TablaX & OFV.gentabla("", "aaa style='filter:shadow(color=#dddddd,direction=125);' ", "", "100%", "", "0", "", "-2", "-1", "", "", "", "no")
		TABLAX = TablaX & "<tr><td  "
		TABLAX = TablaX & "class=vT align=left colspan=3 STYLE='border-style:outset;border-width:2;background-color:#eeeeee;' >" & sinput2 & "&nbsp;SELECCIONE EMPRESA</td></tr>"
		TABLAX = TablaX & "<tr><td colspan=3 "
		TABLAX = TablaX & " >&nbsp;</td></tr>"

		FOR IAEMP=1 TO TAEMP

			'''  RESPONSE.WRITE  ATREMP(IAEMP,2)

			IF ATREMP(IAEMP,2) = 1 THEN

				IF ATREMP(IAEMP,0) = ATREMP(0,0) THEN
					TABLAX = TablaX & "<tr BGCOLOR=#EEEEEE >"
				ELSE
					TABLAX = TablaX & "<tr>"
				END IF

				TABLAX = TablaX & "<td width=5 "
				IF ATREMP(IAEMP,0) = ATREMP(0,0) THEN
					TABLAX = TablaX & "class=vt  >" & sImagen4si & "</td>"
				ELSE
					TABLAX = TablaX & "class=vt  >" & sImagenI3 & "</td>"
				END IF

				sImage = ofv.GenImage("","","","","..\CORESS20\EMPRESA\" & ATREMP(IAEMP,5),"0","")
				TABLAX = TablaX & "<td width=5 "
				TABLAX = TablaX & "class=vt  >" & sImage & "</td>"

				TABLAX = TablaX & "<td class=vt  >"

				sHref = "subconsulta.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0&EMPRE=" & ATREMP(IAEMP,0)

				'''  sHref = "OBJtreeTema.asp?Tipo=" & sTipo & "&Nombre=" & LNombre & "&Menu=" & SESSION("MENU") & "&EMPRE=" & ATREMP(IAEMP,0)
				sLink = ofv.GenLink("","cc",sHref,"","","","<b>" & ATREMP(IAEMP,4) & "</b>","si")
				TABLAX = TablaX & sLink
				TABLAX = TablaX & "</td>"
				TABLAX = TablaX & "</tr>"
			END IF

		NEXT

		TABLAX = TablaX & OFV.genFtabla()

		TABLA = Tabla & TABLAX

		TABLA = Tabla & "</DIV>"

		sHTML = sHTML & TABLA

	end if

	Fila = ""




	tt="&nbsp;"

	''response.write Ucase(sNombreN) & " Ucase " & chr(13)

	select case UCase(sNombreN)
		case "ESTADO"
		Strsql = "Select * From Estados order by estado"

		case "SECTOR"

		case "FECHA"
		Strsql = "SELECT DISTINCT Format(fecemis,'Short Date'), fecemis AS Fecha "
		Strsql = Strsql & "FROM CSEMISION WHERE Str(fecemis) Is Not Null "
		Strsql = Strsql & " order by fecemis "
		case "INSTRUMENTO"
		Strsql = "SELECT * from instrumen order by descri "

		case "PRODUCTO"

		Strsql = "SELECT distinct objtsegagr.tsegmento,TIPOSEGMENTO.descri  FROM objtsegagr, TIPOSEGMENTO,OBJETOS"

		if SESSION("EMPMUN") = "S" THEN
			if SESSION("GLBMUN") = 2 THEN
				Strsql = Strsql & " ,OBJMEMPALC "
			END IF
		END IF

		Strsql = Strsql & " where objtsegagr.tsegmento = TIPOSEGMENTO.objid "
		Strsql = Strsql & " AND OBJETOS.CODOBJETO = OBJTSEGAGR.CODOBJETO AND "

		if SESSION("EMPMUN") <> "S"  THEN
			Strsql = Strsql & " OBJETOS.OBJUN = 1"
		ELSE

			if SESSION("EMPMUN") = "S" AND SESSION("GLBMUN") = 1 THEN
				Strsql = Strsql & " OBJETOS.OBJUN = " & SEMPRE
			ELSE
				Strsql = Strsql & " ( (OBJETOS.OBJUN = " & SEMPRE & ") OR "
				Strsql = Strsql & "   ( OBJETOS.OBJUN = 1 AND "
				Strsql = Strsql & "     OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO AND "
				Strsql = Strsql & "     OBJMEMPALC.CODEMPRE = " & SEMPRE & ") "
				Strsql = Strsql & " ) "
			END IF

		end if

		Strsql = Strsql & " order by TIPOSEGMENTO.descri "


		SQL_SELECT="SELECT distinct objtsegagr.tsegmento,TIPOSEGMENTO.descri   "
		SQL_JOIN= " FROM OBJETOS "
		SQL_JOIN= SQL_JOIN  & " INNER JOIN objtsegagr on  OBJETOS.codobjeto = objtsegagr.codobjeto  "
		'SQL_JOIN= SQL_JOIN  & " INNER JOIN objtsegagr  ON  objtsegagr.codobjeto = OBJETOS.codobjeto  "
		SQL_JOIN= SQL_JOIN  & " INNER JOIN TIPOSEGMENTO ON TIPOSEGMENTO.objid = objtsegagr.tsegmento "


		If (Session("EMPMUN") = "S") Then
			SQL_JOIN= SQL_JOIN  & " LEFT join OBJMEMPALC  on OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO  "
		END IF
		If (Session("EMPMUN") = "S") Then
			SQL_WHERE= " WHERE  ( (OBJETOS.OBJUN = " & Session("UN") &") OR    (   OBJMEMPALC.CODEMPRE = " & Session("UN") & ")  )"
		END IF
		SQL_ORDEN =  " order by TIPOSEGMENTO.descri "

		STRSQL=SQL_SELECT  &  SQL_JOIN &  SQL_WHERE & SQL_ORDEN
		' RESPONSE.WRITE "<FONT COLOR=#000000>" & STRSQL & "<BR>"
		'  RESPONSE.WRITE SESSION("EMPMUN") & "<BR>"
		'  RESPONSE.WRITE SESSION("GLBMUN") & "<BR>"

		case "NUMEROCIRCULAR"

		Strsql = "SELECT distinct objtsegagr.tagrupa,TIPOagrupa.descri  FROM objtsegagr, TIPOagrupa, OBJETOS"

		if SESSION("EMPMUN") = "S" THEN
			if SESSION("GLBMUN") = 2 THEN
				Strsql = Strsql & " ,OBJMEMPALC "
			END IF
		END IF

		Strsql = Strsql & " where objtsegagr.tagrupa = TIPOagrupa.objid "
		Strsql = Strsql & " AND OBJETOS.CODOBJETO = OBJTSEGAGR.CODOBJETO AND "

		if SESSION("EMPMUN") <> "S"  THEN
			Strsql = Strsql & " OBJETOS.OBJUN = 1"
		ELSE

			if SESSION("EMPMUN") = "S" AND SESSION("GLBMUN") = 1 THEN
				Strsql = Strsql & " OBJETOS.OBJUN = " & SEMPRE
			ELSE
				Strsql = Strsql & " ( (OBJETOS.OBJUN = " & SEMPRE & ") OR "
				Strsql = Strsql & "   ( OBJETOS.OBJUN = 1 AND "
				Strsql = Strsql & "     OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO AND "
				Strsql = Strsql & "     OBJMEMPALC.CODEMPRE = " & SEMPRE & ") "
				Strsql = Strsql & " ) "
			END IF

		END IF

		Strsql = Strsql & " order by TIPOagrupa.descri "
		case "ACTAS"
		''''''''''''''''''''
		'DEL MENU DE ACTAS SE DEBE DE GENERAR LA CONSULTA PARA LLENAR LOS COMBOS
		'''''''''''''''''


	end select

	Fila       = ""
	formx      = ""
	Formulario = ""

	'Descripcion
	if session("APLMDCLI")="BC" THEN
		sInput = "<h5>Panel de filtros para la consulta</h5>"
	else
		sInput = "<h4>Panel de filtros para la consulta</h4>"
	end if
	'sInput = "<h3>Panel de filtros para la consulta Basica</h3>"
	Celda  = ofv.GenCelda("","ut colspan=2 bgcolor='#f7eedd' style='border-style:solid;border-width:1;border-color:#dddddd;' ","","","","","",sInput,"si")
	Fila   = ofv.genrow("","","","","30","","",Celda,"si")

	Celda  = ofv.GenCelda("","vt","","30%","","","","&nbsp;","si")
	Celda  = Celda & ofv.GenCelda("","vti","","70%","","","","&nbsp;","si")
	Fila   = Fila & ofv.GenRow("","","","","10","","",Celda,"si")
    Fila=""
    Celda=""
	''response.write sNombre & " sNombre " & " linea 286 "

	if sNombre = "CONSULTA POR PENDIENTES" then

		scodmaster = "000"
		stipmaster = "PROCESO"


		'' sMenu = request.querystring("Menu")
		XTPROC = request.form("PROC")
		XTTEVEN = request.form("TEVEN")
        XAPL = request.form("APL")
        XAPLE=request.QUERYSTRING("APLE")


  ' response.write  request.form & "   " & XAPLE
		IF XTPROC = "" THEN XTPROC = "0"
		IF XTTEVEN = "" THEN XTTEVEN = "0"
        IF XAPL = "" THEN XAPL = "0"
  'RESPONSE.WRITE "XAPLE=" & XAPLE &"       XAPL="& XAPL & "   XTPROC="  & XTPROC & "   XTTEVEN=" &XTTEVEN & "  XAPLE <> XAPL =" & cstr(XAPLE <> XAPL)
  IF XAPLE <> XAPL THEN
   XTTEVEN=0
   XTPROC=0
  END IF



		'--------------------
        Celda  = ofv.GenCelda("","tc","","","","","","<font>TIPO DE CIRCUITO</font>","si")
          SCOMSEP = ""
  		 scomK   = ""
           scomD   = ""
           scomK = scomK & SCOMSEP & "0"
           scomD = scomD & SCOMSEP & "TODOS"
           SCOMSEP = chr(9)
           scomK = scomK & SCOMSEP & "C"
           scomD = scomD & SCOMSEP & "DOCUMENTOS"
           scomK = scomK & SCOMSEP & "K"
           scomD = scomD & SCOMSEP & "FORMULARIOS"
           scomK = scomK & SCOMSEP & "PD"
           scomD = scomD & SCOMSEP & "PEDIDOS"


         sInput = ofv.gencomboFX("Subconsulta1","APL",XAPL,scomK,scomD,"")


  	    Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
  	    Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")


		Celda  = ofv.GenCelda("","tc","","","","","","<font>TIPO DE DOCUMENTO</font>","si")

		SCOMSEP = ""
		scomK   = ""
		scomD   = ""
		sseld   = "0"
		ssels   = ","

		scomK = scomK & SCOMSEP & "0"
		scomD = scomD & SCOMSEP & "TODOS"
		SCOMSEP = chr(9)
		set cn = ofv.conectar(ofv.strconn2)
		set cnx = ofv.conectar(ofv.strconn5)
            IF XAPL="C" THEN
            SQL= "SELECT * FROM TIPOBJETO  WHERE APLIC IN ('EM','GLB') AND TIPOCPT IN('D','C')"
            ELSEIF XAPL="PD" THEN
              SQL= "SELECT * FROM TIPOBJETO  WHERE APLIC IN ('EM','GLB') AND TIPOCPT IN('PD')"
            ELSEIF XAPL="K" THEN
              SQL= "SELECT * FROM TIPOBJETO  WHERE APLIC IN ('EM','GLB') AND TIPOCPT IN('K')"
            ELSE
               SQL= "SELECT * FROM TIPOBJETO WHERE APLIC IN ('EM','GLB') "
            END IF
		DELIM=""
		STIPOS=""
           'RESPONSE.WRITE SQL
  	      set rs = ofv.crearconsultaEx(SQL,cn,1,parametros)
                 SCOMSEP=""
                scomK   = ""
                scomD   = ""
                scomK = scomK & SCOMSEP & "0"
  			    scomD = scomD & SCOMSEP & "TODOS"
                     SCOMSEP = chr(9)
                     STIPOS=""
		DO WHILE NOT RS.EOF

  			        scomK = scomK & SCOMSEP & rs(0)
    				    scomD = scomD & SCOMSEP & rs(1)
			STIPOS=STIPOS & DELIM & "'" & RS(0) & "'"
			DELIM=","
  			       SCOMSEP=CHR(9)
			RS.MOVENEXT
		LOOP
		if STIPOS <> "" then
			STIPOS= "and TIPOBJETO IN(" & STIPOS & ")"
		end if

		call ofv.cerrarconsulta(rs)

		STRSQL = "SELECT TIPOBJETO,descri  FROM TIPOBJETO WHERE APLIC IN ('EM','GLB')  " & stipos
		set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		DELIM=""
		STIPOS=""
		do until rsX1.eof
			scomK = scomK & SCOMSEP & rsx1(0)
			scomD = scomD & SCOMSEP & rsx1(1)
			STIPOS=STIPOS & DELIM & "'" & rsx1(0) & "'"
			DELIM=","

			RSX1.MOVENEXT
		LOOP
		call ofv.cerrarconsulta(rsX1)
		' call ofv.cerrarconn(cnx)

		Strsql = ""

		'Strsql = Strsql & " SELECT A.CODOBJETO, (A.DESCRI + ': ' + A.OBJDESCRI) AS NOMBRE "
		'Strsql = Strsql & " FROM CSOBJETOSRDO A "
		'Strsql = Strsql & " where a.codobjeto in (" & sseld & ") "
		'Strsql = Strsql & " ORDER BY NOMBRE "
		'set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		'do until rsX1.eof

		'   scomK = scomK & SCOMSEP & RSX1(0)
		'   scomD = scomD & SCOMSEP & RSX1(1)
		'   SCOMSEP = chr(9)
		'

		'RSX1.MOVENEXT
		'LOOP
		'call ofv.cerrarconsulta(rsX1)

		sInput = ofv.gencomboFX("Subconsulta1","PROC",XTPROC,scomK,scomD,"")


		Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
		Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")
       ' response.write XAPL & "  "  & XTPROC & "  "  & XTTEVEN
        sAccion = "subconsulta.asp?APLE="&XAPL&"&Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
		Formulario = Formulario & ofv.GenForm("Subconsulta1","",sAccion,"","post","",Fila,"si")

		Fila    = ""
		'---------------------------------------------------------------------------------------------------------



		Celda = ofv.GenCelda("","tc","","","","","","<font>USUARIO RESPONSABLE</font>","si")

		SCOMSEP = ""
		scomK   = ""
		scomD   = ""
		scomK = scomK & SCOMSEP & "0"
		scomD = scomD & SCOMSEP & "TODOS"
		SCOMSEP = chr(9)

		'scomK = "0" & SCOMSEP & "1" & SCOMSEP & "2" & SCOMSEP & "3" & SCOMSEP & "4" & SCOMSEP & "5" & SCOMSEP & "10" & SCOMSEP & "99"
		'scomD = "TODOS" & SCOMSEP & "1 DIA" & SCOMSEP & "2 DIAS" & SCOMSEP & "3 DIAS" & SCOMSEP & "4 DIAS" & SCOMSEP & "5 DIAS" & SCOMSEP & "10 DIAS" & SCOMSEP & "MAS DE 10 DIAS"
		Strsql = Strsql & " SELECT distinct PSCRESPONSABLE  FROM CKTVIEW where "
		Strsql = Strsql & " FIN = 'NO' AND PSCFIN = 'NO' "

  	   IF STIPOS <> "" AND XTPROC & "" <> "0" THEN
			STRSQL=STRSQL & " AND objtipo IN ('" &  XTPROC &"')"
  	   ELSEIF  XTPROC & "" = "0" THEN
   	   STRSQL=STRSQL & " AND objtipo IN (" &  STIPOS &")"
		END IF

		'RESPONSE.WRITE STRSQL

		set rs = ofv.crearconsultaEx(StrSql,cnX,1,parametros)

		do until rs.eof
			scomK = scomK & SCOMSEP & rs(0)
			xxxx = getInfoUser(RS(0))
			sInput=RS(0)
			if session("SGU_nombre") <> "" then    sInput =  session("SGU_nombre")

			scomD = scomD & SCOMSEP & sInput
			RS.MOVENEXT
		LOOP



		sInput = ofv.gencomboFX("Subconsulta2","TEVEN",XTTEVEN,scomK,scomD,"")
        sInput = sinput & ofv.GenerarInput("PROC",XTPROC,"hidden","","1","dt","")
        sInput = sinput & ofv.GenerarInput("APL",xAPL,"hidden","","1","dt","")

		Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
		Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

        sAccion = "subconsulta.asp?APLE="&XAPL&"&Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
		Formulario = Formulario & ofv.GenForm("Subconsulta2","",sAccion,"","post","",Fila,"si")

		Fila    = ""



		'  if XTPROC  = 0 then XTPROC = "ALL"
		'  if XTTEVEN = 0 then XTTEVEN = "ALL"


		Celda  = ""
		Celda  = Celda & ofv.GenCelda("","vti COLSPAN=2","","","","","","&nbsp;","si")
		Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

		formx  = Fila

		'-----------------------------

		'  Celda = ofv.GenCelda("","tc","","","","","","<font>DIAS DE DEMORA</font>","si")


		'  SCOMSEP = chr(9)
		'  scomK   = ""
		'  scomD   = ""

		'  scomK = "0" & SCOMSEP & "1" & SCOMSEP & "2" & SCOMSEP & "3" & SCOMSEP & "4" & SCOMSEP & "5" & SCOMSEP & "10" & SCOMSEP & "99"
		'  scomD = "TODOS" & SCOMSEP & "1 DIA" & SCOMSEP & "2 DIAS" & SCOMSEP & "3 DIAS" & SCOMSEP & "4 DIAS" & SCOMSEP & "5 DIAS" & SCOMSEP & "10 DIAS" & SCOMSEP & "MAS DE 10 DIAS"


		'  sInput = ofv.gencomboFX("Subconsulta2","TEVEN",XTTEVEN,scomK,scomD,"")
		'  sInput = sinput & ofv.GenerarInput("PROC",XTPROC,"hidden","Subconsulta2","1","dt","")

		'  Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
		'  Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

		'  sAccion = "subconsulta.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
		'  Formulario = Formulario & ofv.GenForm("Subconsulta2","",sAccion,"","post","",Fila,"si")

		'  Fila    = ""



		'  if XTPROC  = 0 then XTPROC = "ALL"
		'  if XTTEVEN = 0 then XTTEVEN = "ALL"


		'  Celda  = ""
		'  Celda  = Celda & ofv.GenCelda("","vti COLSPAN=2","","","","","","&nbsp;","si")
		'  Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

		'  formx  = Fila






        sAccion  = "LISTAPENDIENTES.asp?FILSEL=CAT&PROCESO=" & XTPROC & "&SUBFILTRO=" & XTTEVEN & "&SUBFILTRO2=" & XTCATE & "&SUBFILTRO3=" & XTRSK & "&Nombre=" & sNombren & "&Menu=" & sMenu & "&APL=" & XAPL
		sTARGET  = ""


		'' AGREGADO 10102014 PARA CONSULTA DE DOCUMENTOS A VENCER

	elseif sNombre = "CONSULTA POR AVENCER" then

		scodmaster = "000"
		stipmaster = "PROCESO"


		'' sMenu = request.querystring("Menu")
		XTPROC = request.form("PROC")
		XTTEVEN = request.form("TEVEN")

		IF XTPROC = "" THEN XTPROC = "0"
		IF XTTEVEN = "" THEN XTTEVEN = "0"
		'RESPONSE.WRITE XTPROC & "-----------" &XTTEVEN
		'--------------------

		Celda  = ofv.GenCelda("","tc","","","","","","<font>TIPO DE DOCUMENTO</font>","si")

		SCOMSEP = ""
		scomK   = ""
		scomD   = ""
		sseld   = "0"
		ssels   = ","

		scomK = scomK & SCOMSEP & "0"
		scomD = scomD & SCOMSEP & "TODOS"
		SCOMSEP = chr(9)
		set cn = ofv.conectar(ofv.strconn2)
		'   set cnx = ofv.conectar(ofv.strconn5)


		' Strsql = Strsql & " SELECT distinct objtipo  FROM CKTVIEW where "
		' Strsql = Strsql & " FIN = 'NO' AND PSCFIN = 'NO' "
		'response.write strsql
		'     set rs = ofv.crearconsultaEx(StrSql,cnx,1,parametros)

		'     DELIM=""
		'     STIPOS=""
		'     DO WHILE NOT RS.EOF
		'
		'     STIPOS=STIPOS & DELIM & "'" & RS(0) & "'"
		'     DELIM=","
		'     RS.MOVENEXT
		'     LOOP
		'  if STIPOS <> "" then
		'    STIPOS= "and TIPOBJETO IN(" & STIPOS & ")"
		'  end if
		'
		'   call ofv.cerrarconsulta(rs)

		STRSQL = "SELECT TIPOBJETO,descri  FROM TIPOBJETO WHERE APLIC IN ('EM','GLB')  " '& stipos
		set rsX1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		DELIM=""
		STIPOS=""
		do until rsX1.eof
			scomK = scomK & SCOMSEP & rsx1(0)
			scomD = scomD & SCOMSEP & rsx1(1)
			STIPOS=STIPOS & DELIM & "'" & rsx1(0) & "'"
			DELIM=","

			RSX1.MOVENEXT
		LOOP
		call ofv.cerrarconsulta(rsX1)
		' call ofv.cerrarconn(cnx)

		Strsql = ""

		sInput = ofv.gencomboFX("Subconsulta1","PROC",XTPROC,scomK,scomD,"")


		Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
		Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

		sAccion = "subconsulta.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
		Formulario = Formulario & ofv.GenForm("Subconsulta1","",sAccion,"","post","",Fila,"si")

		Fila    = ""
		'---------------------------------------------------------------------------------------------------------



		Celda = ofv.GenCelda("","tc","","","","","","<font>USUARIO RESPONSABLE</font>","si")

      dFechadia = clng(fechaConvert(date)) - clng(cdate("01-01-1900"))

		SCOMSEP = ""
		scomK   = ""
		scomD   = ""
		scomK = scomK & SCOMSEP & "0"
		scomD = scomD & SCOMSEP & "TODOS"
		SCOMSEP = chr(9)

		'scomK = "0" & SCOMSEP & "1" & SCOMSEP & "2" & SCOMSEP & "3" & SCOMSEP & "4" & SCOMSEP & "5" & SCOMSEP & "10" & SCOMSEP & "99"
		'scomD = "TODOS" & SCOMSEP & "1 DIA" & SCOMSEP & "2 DIAS" & SCOMSEP & "3 DIAS" & SCOMSEP & "4 DIAS" & SCOMSEP & "5 DIAS" & SCOMSEP & "10 DIAS" & SCOMSEP & "MAS DE 10 DIAS"
		Strsql = Strsql & " SELECT distinct autor,OBJUN  FROM objprty "
		Strsql = Strsql & " INNER JOIN OBJETOS ON(objprty.CODOBJETO=OBJETOS.CODOBJETO) WHERE "

		Strsql = Strsql & " fecvto > '" & dFechadia & "' "

		''''''RESPONSE.WRITE xtproc & " xtproc "
		if XTPROC <> "" then
			Strsql = Strsql & " and objprty.tipobjeto ='" & XTPROC & "' "
		end if
		''RESPONSE.WRITE xtproc
		' RESPONSE.WRITE STRSQL

		set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

		do until rs.eof
			scomK = scomK & SCOMSEP & rs(0)
			xxxx = getInfoUserEX(RS(0),RS("OBJUN"))
			sInput=RS(0)
			if session("SGU_nombre") <> "" then    sInput =  session("SGU_nombre")

			scomD = scomD & SCOMSEP & sInput
			RS.MOVENEXT
		LOOP



		sInput = ofv.gencomboFX("Subconsulta2","TEVEN",XTTEVEN,scomK,scomD,"")
      sInput = sinput & ofv.GenerarInput("PROC",XTPROC,"hidden","","1","dt","")
	   sInput = sinput & ofv.GenerarInput("APL",xAPL,"hidden","","1","dt","")

     ' sInput = sinput & ofv.GenerarInput("PROC",XTPROC,"hidden","Subconsulta2","1","dt","")

		Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
		Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

		sAccion = "subconsulta.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
		Formulario = Formulario & ofv.GenForm("Subconsulta2","",sAccion,"","post","",Fila,"si")

		Fila    = ""



		'  if XTPROC  = 0 then XTPROC = "ALL"
		'  if XTTEVEN = 0 then XTTEVEN = "ALL"


		Celda  = ""
		Celda  = Celda & ofv.GenCelda("","vti COLSPAN=2","","","","","","&nbsp;","si")
		Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

		formx  = Fila

		'-----------------------------

		'  Celda = ofv.GenCelda("","tc","","","","","","<font>DIAS DE DEMORA</font>","si")


		'  SCOMSEP = chr(9)
		'  scomK   = ""
		'  scomD   = ""

		'  scomK = "0" & SCOMSEP & "1" & SCOMSEP & "2" & SCOMSEP & "3" & SCOMSEP & "4" & SCOMSEP & "5" & SCOMSEP & "10" & SCOMSEP & "99"
		'  scomD = "TODOS" & SCOMSEP & "1 DIA" & SCOMSEP & "2 DIAS" & SCOMSEP & "3 DIAS" & SCOMSEP & "4 DIAS" & SCOMSEP & "5 DIAS" & SCOMSEP & "10 DIAS" & SCOMSEP & "MAS DE 10 DIAS"


		'  sInput = ofv.gencomboFX("Subconsulta2","TEVEN",XTTEVEN,scomK,scomD,"")
		'  sInput = sinput & ofv.GenerarInput("PROC",XTPROC,"hidden","Subconsulta2","1","dt","")

		'  Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
		'  Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

		'  sAccion = "subconsulta.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
		'  Formulario = Formulario & ofv.GenForm("Subconsulta2","",sAccion,"","post","",Fila,"si")

		'  Fila    = ""



		'  if XTPROC  = 0 then XTPROC = "ALL"
		'  if XTTEVEN = 0 then XTTEVEN = "ALL"


		'  Celda  = ""
		'  Celda  = Celda & ofv.GenCelda("","vti COLSPAN=2","","","","","","&nbsp;","si")
		'  Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

		'  formx  = Fila

		''response.write XTTEVEN & " XTTEVEN "

		sAccion  = "LISTAPENDAVENCER.asp?FILSEL=CAT&PROCESO=" & XTPROC & "&SUBFILTRO=" & XTTEVEN & "&SUBFILTRO2=" & XTCATE & "&SUBFILTRO3=" & XTRSK & "&Nombre=" & sNombren & "&Menu=" & sMenu
		sTARGET  = ""

		''
		'' FIN AGREGADO 10102014 PARA CONSULTA DE DOCUMENTOS A VENCER
		''

		'''''''''''''''''''''''''''''''''' nueva a vencer desde 31-01-2018 '''''''''''''''''''''''''''''''''''''''''''''''
	elseif sNombre = "CONSULTA POR AVENCERnooo" then


		'Tipo de esquema

		scomK = ""
		scomD = ""

		scomK = scomK & SCOMSEP & "0"
		scomD = scomD & SCOMSEP & "TODOS"
		SCOMSEP = chr(9)

		Strsql = "SELEct * FROM TIPOBJETO WHERE TIPOCPT = 'D' and FILTRO = 'S' and aplic='"& WKAPLIC &"' order by ICONO,descri "
		''  response.write strsql & "<br>" & chr(13)
		set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		DO UNTIL rsniv.eof
			scomK = scomK & rsniv(0) & chr(9)
			scomD = scomD & rsniv(1) & chr(9)
			RSNIV.MOVENEXT
		LOOP
		call ofv.cerrarconsulta(rsniv)

		sInput = ofv.gencomboFX("Subconsulta","segm","0",scomK,scomD,"")



		celda  = ofv.GenCelda("","tc","","","","","","<font>Tipo de documento</font>","si")
		Celda  = celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
		Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")


		formx    = Fila

		sTARGET  = ""

		sAccion    = "ListadoConsultas.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren


		'Fin preparacion


		''''''''''''''''''''''''''''''''''''' fin nueva a vencer ''''''''''''''''''''''''''''''''''''

	elseif sNombre = "CONSULTA POR PRODUCTO" then


		'' sMenu = request.querystring("Menu")
		xtseg = clng("0" & request.form("segm"))
		xtagr = clng("0" & request.form("agrm"))
		xtPRO = clng("0" & request.form("PROm"))

		'response.write segm

		'--------------------

		Celda  = ofv.GenCelda("","tc","","","","","","<font>Segmento</font>","si")

		scomKK = ""
		scomDD = ""
		''  response.write strsql & "<br>" & chr(13)
		set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		DO UNTIL rsniv.eof
			scomKK = scomKK & rsniv(0) & chr(9)
			scomDD = scomDD & rsniv(1) & chr(9)
			RSNIV.MOVENEXT
		LOOP
		call ofv.cerrarconsulta(rsniv)

		sInput = ofv.gencomboFX("Subconsulta1","segm",xtseg,scomKK,scomDD,"")

		'' sInput = ofv.GenerarCombo("Flform","Fltipo",Fltipo,Strsql,cn,0,1,"si")
		'sInput = ofv.GenerarCombo("Subconsulta1","segm",xtseg,Strsql,cn,0,1,"")
		Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
		Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

		sAccion = "subconsulta.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
		Formulario = Formulario & ofv.GenForm("Subconsulta1","",sAccion,"","post","",Fila,"si")

		Fila    = ""
		'-----------------------------

		Celda = ofv.GenCelda("","tc","","","","","","<font>Grupo</font>","si")

		Strsql = "SELECT distinct objtsegagr.tagrupa,tipoagrupa.descri  FROM objtsegagr, tipoagrupa, OBJETOS"
		if SESSION("EMPMUN") = "S" THEN
			if SESSION("GLBMUN") = 2 THEN
				Strsql = Strsql & " ,OBJMEMPALC "
			END IF
		END IF
		Strsql = Strsql & " where objtsegagr.tagrupa = tipoagrupa.objid and "
		Strsql = Strsql & " objtsegagr.tsegmento = " & xtseg  & " AND "


		if SESSION("EMPMUN") <> "S"  THEN
			Strsql = Strsql & " OBJETOS.OBJUN = 1"
		ELSE

			if SESSION("EMPMUN") = "S" AND SESSION("GLBMUN") = 1 THEN
				Strsql = Strsql & " OBJETOS.OBJUN = " & SEMPRE
			ELSE
				Strsql = Strsql & " ( (OBJETOS.OBJUN = " & SEMPRE & ") OR "
				Strsql = Strsql & "   ( OBJETOS.OBJUN = 1 AND "
				Strsql = Strsql & "     OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO AND "
				Strsql = Strsql & "     OBJMEMPALC.CODEMPRE = " & SEMPRE & ") "
				Strsql = Strsql & " ) "
			END IF

		end if

		Strsql = Strsql & " order by tipoagrupa.descri "

		SQL_SELECT="SELECT distinct objtsegagr.tagrupa,tipoagrupa.descri   "
		SQL_JOIN= " FROM OBJETOS "
		SQL_JOIN= SQL_JOIN  & " INNER JOIN objtsegagr on  OBJETOS.codobjeto = objtsegagr.codobjeto  "

		SQL_JOIN= SQL_JOIN  & " INNER JOIN tipoagrupa ON tipoagrupa.objid = objtsegagr.tagrupa "


		If (Session("EMPMUN") = "S") Then
			SQL_JOIN= SQL_JOIN  & " LEFT join OBJMEMPALC  on OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO  "
		END IF
		SQL_WHERE= " WHERE "
		If (Session("EMPMUN") = "S") Then
			SQL_WHERE= SQL_WHERE & "   ( (OBJETOS.OBJUN = " & Session("UN") &") OR    (   OBJMEMPALC.CODEMPRE = " & Session("UN") & ")  ) AND "
		END IF

		SQL_WHERE=SQL_WHERE &  " objtsegagr.tsegmento = " & xtseg

		SQL_ORDEN =  " order by tipoagrupa.descri"

		STRSQL=SQL_SELECT  &  SQL_JOIN  & SQL_WHERE &  SQL_ORDEN


		scomKK = ""
		scomDD = ""
		' response.write strsql & "<br>" & chr(13)
		set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		DO UNTIL rsniv.eof
			scomKK = scomKK & rsniv(0) & chr(9)
			scomDD = scomDD & rsniv(1) & chr(9)
			RSNIV.MOVENEXT
		LOOP
		call ofv.cerrarconsulta(rsniv)

		sInput = ofv.gencomboFX("Subconsulta2","agrm",xtagr,scomKK,scomDD,"")
		'sInput = ofv.GenerarCombo("Subconsulta2","agrm",xtagr,Strsql,cn,0,1,"") & "Grupo"
		sInput = sinput & ofv.GenerarInput("segm",xtseg,"hidden","Subconsulta2","1","dt","")

		Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
		Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

		sAccion = "subconsulta.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
		Formulario = Formulario & ofv.GenForm("Subconsulta2","",sAccion,"","post","",Fila,"si")

		Fila    = ""

		'-----------------------------

		Celda = ofv.GenCelda("","tc","","","","","","<font>Producto</font>","si")

		Strsql = "SELECT distinct objetos.tsegmento,tipoproducto.descri "
		Strsql = Strsql & " FROM objtsegagr, objetos, tipoproducto "

		if SESSION("EMPMUN") = "S" THEN
			if SESSION("GLBMUN") = 2 THEN
				Strsql = Strsql & " ,OBJMEMPALC "
			END IF
		END IF

		Strsql = Strsql & " where objtsegagr.tagrupa = " & xtagr & " and "
		Strsql = Strsql & " objtsegagr.tsegmento = " & xtseg & " and "
		Strsql = Strsql & " objtsegagr.codobjeto = objetos.codobjeto and "
		Strsql = Strsql & " objetos.tsegmento = tipoproducto.objid AND "


		if SESSION("EMPMUN") <> "S"  THEN
			Strsql = Strsql & " OBJETOS.OBJUN = 1"
		ELSE

			if SESSION("EMPMUN") = "S" AND SESSION("GLBMUN") = 1 THEN
				Strsql = Strsql & " OBJETOS.OBJUN = " & SEMPRE
			ELSE
				Strsql = Strsql & " ( (OBJETOS.OBJUN = " & SEMPRE & ") OR "
				Strsql = Strsql & "   ( OBJETOS.OBJUN = 1 AND "
				Strsql = Strsql & "     OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO AND "
				Strsql = Strsql & "     OBJMEMPALC.CODEMPRE = " & SEMPRE & ") "
				Strsql = Strsql & " ) "
			END IF

		end if


		Strsql = Strsql & "  order by tipoproducto.descri "
		scomKK = ""
		scomDD = ""
		' response.write strsql & "<br>" & chr(13)
		set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		DO UNTIL rsniv.eof
			scomKK = scomKK & rsniv(0) & chr(9)
			scomDD = scomDD & rsniv(1) & chr(9)
			RSNIV.MOVENEXT
		LOOP
		call ofv.cerrarconsulta(rsniv)

		sInput = ofv.gencomboFX("Subconsulta3","PROm",xtPRO,scomKK,scomDD,"")
		'sInput = ofv.GenerarCombo("Subconsulta3","PROm",xtPRO,Strsql,cn,0,1,"")& "producto"
		sInput = sinput & ofv.GenerarInput("SEGM",xtseg,"hidden","Subconsulta3","1","dt","")
		sInput = sinput & ofv.GenerarInput("AGRM",xtagr,"hidden","Subconsulta3","1","dt","")
		Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
		Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

		sAccion = "subconsulta.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
		Formulario = Formulario & ofv.GenForm("Subconsulta3","",sAccion,"","post","",Fila,"si")

		Fila    = ""

		'-------------------------

		Celda = ofv.GenCelda("","tc","","","","","","<font>Tipo</font>","si")

		Strsql = "SELECT distinct TIPOBJETO.DESCRI  FROM objtsegagr, objetos, tipoproducto,TIPOBJETO "

		if SESSION("EMPMUN") = "S" THEN
			if SESSION("GLBMUN") = 2 THEN
				Strsql = Strsql & " ,OBJMEMPALC "
			END IF
		END IF

		Strsql = Strsql & " where TIPOBJETO.TIPOBJETO = OBJETOS.TIPOBJETO AND "
		Strsql = Strsql & " objtsegagr.tagrupa = " & xtagr & " and objtsegagr.tsegmento = " & xtseg
		Strsql = Strsql & " and objtsegagr.codobjeto = objetos.codobjeto and "
		Strsql = Strsql & " objetos.tsegmento = tipoproducto.objid  AND tipoproducto.objid = " & XTPRO & " AND "


		if SESSION("EMPMUN") <> "S"  THEN
			Strsql = Strsql & " OBJETOS.OBJUN = 1"
		ELSE

			if SESSION("EMPMUN") = "S" AND SESSION("GLBMUN") = 1 THEN
				Strsql = Strsql & " OBJETOS.OBJUN = " & SEMPRE
			ELSE
				Strsql = Strsql & " ( (OBJETOS.OBJUN = " & SEMPRE & ") OR "
				Strsql = Strsql & "   ( OBJETOS.OBJUN = 1 AND "
				Strsql = Strsql & "     OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO AND "
				Strsql = Strsql & "     OBJMEMPALC.CODEMPRE = " & SEMPRE & ") "
				Strsql = Strsql & " ) "
			END IF

		end if



		sInput = ofv.GenerarCombox3(strSQL, cn, 0,0, "SFILTRO")

		if xtseg = 0 then xtseg = "nsnc"
		if xtagr = 0 then xtagr = "nsnc"
		if xtPRO = 0 then xtPRO = "nsnc"

		sInput = sinput & ofv.GenerarInput("ListaS",xtseg,"hidden","Subconsulta","1","dt","")
		sInput = sinput & ofv.GenerarInput("ListaG",xtagr,"hidden","Subconsulta","1","dt","")
		sInput = sinput & ofv.GenerarInput("ListaP",xtPRO,"hidden","Subconsulta","1","dt","")

		Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
		Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

		formx  = Fila


		sTARGET  = ""
		sAccion    = "ListadoConsultas.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren



	elseif sNombre = "CONSULTA POR NUMEROCIRCULAR" then

		xtagr   = clng("0" & request.form("agrm"))

		Celda   = ofv.GenCelda("","tc","","","","","","<font>Grupo</font>","si")
		scomKK = ""
		scomDD = ""
		''  response.write strsql & "<br>" & chr(13)
		set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		DO UNTIL rsniv.eof
			scomKK = scomKK & rsniv(0) & chr(9)
			scomDD = scomDD & rsniv(1) & chr(9)
			RSNIV.MOVENEXT
		LOOP
		call ofv.cerrarconsulta(rsniv)

		sInput = ofv.gencomboFX("Subconsulta1","agrm",xtagr,scomKK,scomDD,"")
		'sInput  = ofv.GenerarCombo("Subconsulta1","agrm",xtagr,Strsql,cn,0,1,"")
		Celda   = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
		Fila    = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

		sAccion = "subconsulta.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
		Formulario = Formulario & ofv.GenForm("Subconsulta1","",sAccion,"","post","",Fila,"si")

		Fila    = ""

		Celda = ofv.GenCelda("","tc","","","","","","<font>Numero</font>","si")

		Strsql = "SELECT distinct objetos.codiden, "
		strsql = strsql & " INDI = CASE "
		strsql = strsql & " WHEN ISNUMERIC(objetos.codiden)= 1 "
		strsql = strsql & " THEN REPLACE(CONVERT(NVARCHAR(5), SPACE(5 - (DATALENGTH(OBJETOS.CODIDEN)/2)) + objetos.codiden ),' ','0') "
		strsql = strsql & " ELSE CONVERT(NVARCHAR(8), objetos.codiden ) "
		strsql = strsql & " END "
		strsql = strsql & " FROM objtsegagr, objetos "

		if SESSION("EMPMUN") = "S" THEN
			if SESSION("GLBMUN") = 2 THEN
				Strsql = Strsql & " ,OBJMEMPALC "
			END IF
		END IF

		strsql = strsql & " where objtsegagr.tagrupa = " & xtagr & " and "
		strsql = strsql & " objtsegagr.codobjeto = objetos.codobjeto AND "


		if SESSION("EMPMUN") <> "S"  THEN
			Strsql = Strsql & " OBJETOS.OBJUN = 1"
		ELSE

			if SESSION("EMPMUN") = "S" AND SESSION("GLBMUN") = 1 THEN
				Strsql = Strsql & " OBJETOS.OBJUN = " & SEMPRE
			ELSE
				Strsql = Strsql & " ( (OBJETOS.OBJUN = " & SEMPRE & ") OR "
				Strsql = Strsql & "   ( "
				Strsql = Strsql & "     OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO AND "
				Strsql = Strsql & "     OBJMEMPALC.CODEMPRE = " & SEMPRE & ") "
				Strsql = Strsql & " ) "
			END IF

		end if

		strsql = strsql & " order by INDI, objetos.codiden "
		scomKK = ""
		scomDD = ""
		'  response.write strsql & "<br>" & chr(13)
		set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
		DO UNTIL rsniv.eof
			scomKK = scomKK & rsniv(0) & chr(9)
			scomDD = scomDD & rsniv(1) & chr(9)
			RSNIV.MOVENEXT
		LOOP
		call ofv.cerrarconsulta(rsniv)

		sInput = ofv.gencomboFX("Subconsulta","ListaN","",scomKK,scomDD,"no0")
		'' sInput   = ofv.GenerarCombox3(strSQL, cn, 0,1, "ListaN")
		if xtagr = 0 then xtagr = "nsnc"
		sInput   = sinput & ofv.GenerarInput("ListaG",xtagr,"hidden","Subconsulta","1","dt","")
		Celda    = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
		Fila     = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

		formx    = Fila

		sTARGET  = ""
		sAccion    = "ListadoConsultas.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren


	elseif sNombre = "CONSULTA POR INSTRUMENTO" then

		Celda      = ofv.GenCelda("","tc","","","","","","<font>Instrumento</font>","si")

		Set rsPerfil = ofv.crearconsultaEx(Strsql, cn, 1, parametros)
		scomK = ""'"null" & Chr(9)

		scomD = "" '"" & Chr(9)
		SCOMSEP =""
		tttt=""
		if  not rsperfil.eof then
			tttt=rsperfil(0)
		end if

		do while not rsperfil.eof
			scomK=scomK & SCOMSEP & rsperfil(0)
			scomd=scomd & SCOMSEP & rsperfil(1)
			SCOMSEP = chr(9)
			rsPerfil.movenext
		loop

		sInput =  ofv.gencomboFX("Subconsulta", "Lista", tttt, scomK, scomD, "no0")

		'' sInput    = ofv.GenerarCombox2(strSQL, cn, 0,1, "Lista")
		Celda      = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
		Fila       = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

		Celda      = ofv.GenCelda("","tc","","","","","","<font>Identificacion</font>","si")
		sInput     = ofv.GenerarInput("buscar","","agregar","","60","dt","")
		Celda      = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
		Fila       = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

		formx    = Fila

		sTARGET  = ""
		sAccion  = "ListadoConsultas.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren





	elseif sNombre = "CONSULTA POR IDENTIFICACION" then


		Celda      = ofv.GenCelda("","tc","","","","","","<font>Identificacion</font>","si")
		sInput     = ofv.GenerarInput("buscar","","agregar","","60","dt","")
		Celda      = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
		Fila       = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

		formx    = Fila

		sTARGET  = ""
		sAccion    = "ListadoConsultas.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren

	elseif sNombre = "CONSULTA POR TEXTO" then



		sInput  = ""
		sInput2  = ""
		sInput3  = ""
		sInput4  = ""
		sInput5  = ""
		sInput6  = ""
		sInput7  = ""

		xstdoc   = "0"
		xstsel   = "0"
		xstsec   = "0"



		''response.write "<br><br><br>---------------------------"'
		''for each x in request.form
		'   response.write "-- " & x & " : " & request.form(x) & " --<br>"
		'next
		''response.write "---------------------------"


		xx1 = request.form("Ntipodoc") & ""
		xx2 = request.form("Ntcod") & ""
		xx3 = request.form("Nlista") & ""
		xx4 = request.form("textot") & ""
		xx5 = request.form("tsel") & ""

		ttcc = request.form("ttcc") & ""


		IF TTCC = "" then TTCC = SESSION("TTCC")
		IF TTCC = "" then TTCC = "BASIC"


		stipodoc    = request.form("Ntipodoc") & ""
		if stipodoc = "" then
			'   stipodoc = "99999"
			xstdoc   = "0"
		else
			xstdoc   = "1"
		end if



		tsel     = request.form("tsel") & ""

		if tsel  = "" then
			tsel  = "C"
			tgrp  = "T"
			tcod  = "99999"
			ttip  = "TODOS"
			tnam  = "TEMA"
			xstsel   = "0"
		else
			xstsel   = "1"

			if tsel  = "" then
				tgrp  = "T"
			end if

		end if

		frmsen = "N"

		if tsel  = "T" then
			tgrp  = "M"
			tcod  = "99999"
			ttip  = "TODOS"
			tnam  = "TODOS"

		end if

		sSectort = request.form("Nlista") & ""

		sSectort    = ofv.convertircar(SsectorT,"_"," ")
		if ssectort = "" then
			'   ssectort = "99999"
			xstsec   = "0"
		else
			xstsec   = "1"
		end if

		textot = request.form("textot") & ""




		formul = request.querystring("frm")

		if len(formul) = 0 then

			tgrp  = request.form("tgrp") & ""
			tcod  = request.form("Ntcod") & ""
			ttip  = request.form("ttip") & ""
			tnam  = request.form("tnam") & ""


		elseif formul = 4 then

			if tsel  <> "T" then
				tgrp   = request.form("ptgrp") & ""
				tcod   = request.form("ptcod") & ""
				ttip   = request.form("pttip") & ""
				tnam   = request.form("ptnam") & ""
				textot = request.form("ptextot") & ""
			end if

		else

			if tsel  <> "T" then
				tgrp  = request.form("tgrp") & ""
				tcod  = request.form("Ntcod") & ""
				ttip  = request.form("ttip") & ""
				tnam  = request.form("tnam") & ""

			end if


		end if

		if  tgrp <> "M" and formul=4 then frmsen = "S"

		''''''''' introduccion '''''''''''''''''''''''''''''''''''''

		if ttcc = "BASIC"  then
			tcod     = "99999"
			sSectort = "99999"
			stipodoc = "99999"
			tsel     = "C"
			tgrp     = "T"
			ttip     = "TODOS"
			tnam     = "TODOS"
			xstdoc   = "1"
			xstsel   = "1"
			xstsec   = "1"

		end if

		sInputxx1 = ofv.GenerarInput("Ntipodoc","99999","hidden","","1","dt","")
		sInputxx2 = ofv.GenerarInput("Ntcod","99999","hidden","","1","dt","")
		sInputxx3 = ofv.GenerarInput("Nlista","99999","hidden","","1","dt","")
		sInputxx5 = ofv.GenerarInput("tsel","C","hidden","","1","dt","")
		sInputxx6  = ofv.GenerarInput("tgrp","T","hidden","","1","dt","")
		sInputxx7  = ofv.GenerarInput("ttip","TODOS","hidden","","1","dt","")
		sInputxx8  = ofv.GenerarInput("tnam","TODOS","hidden","","1","dt","")

		sInputxx9b  = ofv.GenerarInput("ttcc","BASIC","hidden","","1","dt","")

		sInputxx9a  = ofv.GenerarInput("ttcc","AVANZADA","hidden","","1","dt","")


		IF TTCC="BASIC" OR TTCC="" THEN
			sInputxb1 = ofv.GenerarInput("bx1","BUSQUEDA AVANZADA","BUTTON","","1","dt style=""background-color:color-mix(in srgb, " & Session("CMFPNBOT") & ", transparent 80%);border-color:color-mix(in srgb, " & Session("CMFPNBOT") & ", transparent 20%)""","document.all.Subconsultax1.submit();")
			sInputxb2=""
		ELSE
			sInputxb1=""
			sInputxb2 = ofv.GenerarInput("bx2","BUSQUEDA BASICA","BUTTON","","1","dt style=""background-color:color-mix(in srgb, " & Session("CMFPNBOT") & ", transparent 80%);border-color:color-mix(in srgb, " & Session("CMFPNBOT") & ", transparent 20%)""","document.all.Subconsultax2.submit();")
		END IF


		sAccion = "Subconsulta.asp?Menu=" & sMenu & "&Nombre=" & request.querystring("Nombre")
		Formuln2 = ofv.GenForm("Subconsultax2","aa style='margin:0;'",sAccion,"","post","",sInputxx1 & sInputxx2 & sInputxx3 & sInputxx5 & sInputxx6 & sInputxx7 & sInputxx8 & sInputxx9B,"si")


		sAccion = "Subconsulta.asp?Menu=" & sMenu & "&Nombre=" & request.querystring("Nombre")
		Formuln1 = ofv.GenForm("Subconsultax1","aa style='margin:0;'",sAccion,"","post","",sInputxx5 & sInputxx6 & sInputxx7 & sInputxx8 & sInputxx9A,"si")



		sHTML = sHTML & Formuln2 & Formuln1



		celdan = ""

		celdan = celdan & ofv.GenCelda("","ut ","left","","","","",sInputxb1,"si") ' avanzada
		celdan = celdan & ofv.GenCelda("","ut   ","","","","","",sInputxb2,"si") 'basica
		Filan  = ofv.GenRow("","","","","","","",celdan,"si")
		Filan  = ofv.GenTabla(""," style='background-color:color-mix(in srgb, " & Session("CMFPNTIT") & ", transparent 90%);border-style:solid;border-width:2px;border-color:color-mix(in srgb, " & Session("CMFPNTIT") & ", transparent 50%);' ","","","","0","","0","0","","",Filan,"si")'''aaa style
		'Filan  = ofv.GenTabla("","aaa ' ","","20% ","","0","","0","0","","",Filan,"si")

		celda = ofv.GenCelda("","ut COLSPAN=2  ","left","","","","",Filan,"si")
		Fila  = fila & ofv.GenRow("","","","","30","","",Celda,"si")



		'sglmar = "<font id='fnt01' size=4 color='#f7aa22' face='corflaticon' onclick='alert(""Seleccione los parametros para la busqueda...."");' >&#0125;</font>&nbsp;"
		'sglpnt = "<font size=2 color='#00aa00' face='corflaticon' >4</font>&nbsp;"
		'sinput = sglmar & "Se recomienda acotar la búsqueda mediante la utilización de los filtros, a efecto de hacerla mas rápida, precisa y eficiente."


		sintab = "<script language='javascript' > "
		sintab = sintab & " function opsel(dtfld) {document.getElementById(dtfld + 'tab1').style.display='none';document.getElementById(dtfld + 'sl').style.display='block';} "
		sintab = sintab & " function setsel(dtfld) {"
		sintab = sintab & "   document.all.Subconsultax[dtfld].value=document.all.Subconsultax[dtfld + 'sl'].value;"
		sintab = sintab & "   document.getElementById(dtfld +'td1').innerHTML=document.all.Subconsultax[dtfld + 'sl'].options[document.all.Subconsultax[dtfld + 'sl'].selectedIndex].text;"
		sintab = sintab & "   document.getElementById(dtfld + 'tab1').style.display='block';document.getElementById(dtfld + 'sl').style.display='none';"
		sintab = sintab & "   document.all.Subconsultax.submit();"
		sintab = sintab & "                    } "
		sintab = sintab & "</script>"

		sinput = sinput & sintab

		celda = ofv.GenCelda("","utg COLSPAN=2 style='text-align:justify;' ","","","","","",sinput,"si")
		Fila  = fila & ofv.GenRow("","","","","30","","",Celda,"si")

		if ttcc="xxxxxx"  then
			TTCC     = SESSION("TTCC")
			tcod     = "99999"
			sSectort = "99999"
			stipodoc = "99999"
			tsel="C"
			tgrp="T"
			ttip="TODOS"
			tnam="TODOS"
			xstdoc   = "1"
			xstsel   = "1"
			xstsec   = "1"


		end if

		'response.write ttcc
		if ttcc <> "BASIC" THEN



			''''''''' tipo de documento '''''''''''''''''''''''''''''''''''''
			'sInput = "<b>Usted est&aacute en la consulta avanzada</b>"
			'Celda = ofv.GenCelda("","utg  style='border-style:solid;border-width:1;border-color:#000099;'","center","50%","","","",sInput,"si")
			'Fila  = fila & ofv.GenRow("","","center","","30","","",Celda,"si")

			sinputtt = "<font>¿Qu&eacute; tipo de objeto est&aacute; Ud. buscando ? </font>"

			celda = ofv.GenCelda("","tc","","","","","",sinputtt,"si")


			Strsql = "SELEct * FROM TIPOBJETO WHERE TIPOCPT = 'D' and FILTRO = 'S' and aplic='"& WKAPLIC &"' order by ICONO,descri "

			scomK = ""
			scomD = "SELECCIONE DOCUMENTO"

			sinputdtK = ""
			sinputdtd = "SELECCIONE DOCUMENTO"
			sintab2   = ""
			SCOMSEP = chr(9)
			set rsr = ofv.crearconsultaEx(Strsql,cn,1,parametros)
			do while not rsr.eof

				'scomK = scomK & chr(9) & rsr(0)
				'scomD = scomD & chr(9) & rsr(1)
				'SCOMSEP =""
				'tttt=""

				'  sintab2 = sintab2 & "<option value='" & rsr(0) & "' "


				'  if stipodoc  = cstr(rsr(0)) then
				'     sinputdtd = rsr(1)
				'     sintab2    = sintab2 & " selected "
				'  end if

				'  sintab2 = sintab2 & " >" & rsr(1) & "</option>"
				if stipodoc  = cstr(rsr(0)) then
					sinput = rsr(1)
				end if
				'if not rsperfil.eof then wwww = rsperfil(0)

				scomK=scomK & SCOMSEP & rsr(0)
				scomd=scomd & SCOMSEP & rsr(1)
				SCOMSEP = chr(9)

				rsr.movenext
			loop
			sInput =  ofv.gencomboFX("Subconsultax", "Ntipodoc",stipodoc, scomK, scomD, "no1")
			call ofv.cerrarconsulta(rsr)

			' scomK = scomK & chr(9) & "99999"
			' scomD = scomD & chr(9) & "* CUALQUIER TIPO DE OBJETO"

			' sintab2 = sintab2 & "<option value='99999' "

			' if stipodoc  = "99999" then
			'    sinputdtd = "* CUALQUIER TIPO DE OBJETO"
			'    sintab2    = sintab2 & " selected "
			' end if

			' sintab2 = sintab2 & " >* CUALQUIER TIPO DE OBJETO</option>"


			' sintab = "<table id='Ntipodoctab1' width='70%' cellspacing=0 cellpadding=0 style='display:block;' >"
			' sintab = sintab & "<td id='Ntipodoctd1' width='100%' class='tdcm1' onclick=""opsel('Ntipodoc');"" >" & sinputdtd & "</td>"
			' sintab = sintab & "<td id='Ntipodoctd2' width='1'    class='tdcm2' onclick=""opsel('Ntipodoc');"" >"
			' sintab = sintab & "<font class='btic'  >6<font></td>"
			' sintab = sintab & "</table>"


			' sintab3 = "<select name='Ntipodocsl' size='5' onchange=""setsel('Ntipodoc');"" "
			' sintab3 = sintab3 & " style='display:none;border-style:solid;border-width:1;border-color:#888888;' >" & sintab2
			' sintab3 = sintab3 & "</select>"


			' sinputdtK = ofv.GenerarInput("Ntipodoc",stipodoc,"hidden","","1","dt","")

			' sInput   = sintab & sintab3 & sinputdtK


			'  sInputxxx  = stipodoc & "-" & tcod & "-" & ofv.convertircar(SsectorT," ","_") & "<hr>"


			'  Celdaxxx = ofv.GenCelda("","ut colspan=2 ","center","","","","",sInputxxx,"si")
			'  Fila  = fila & ofv.GenRow("","","","","30","","",Celdaxxx,"si")




			'  sInputxxx  = request.form("Ntipodoc") & "" & "-" & request.form("Ntcod") & "" & "-" & request.form("Nlista") & "" & "<hr>"


			'  Celdaxxx = ofv.GenCelda("","ut colspan=2 ","center","","","","",sInputxxx,"si")
			'  Fila  = fila & ofv.GenRow("","","","","30","","",Celdaxxx,"si")




			Celda = Celda & ofv.GenCelda("","ut","","","","","",sInput,"si")
			Fila  = fila & ofv.GenRow("","","","","30","","",Celda,"si")





			'''''''''''' carpetas ''''''''''''''''''''''''''''''''''''''''''''''''

			if  xstdoc   = "1" then

				sinputtt = "<font>¿En qu&eacute; carpeta desea buscarlo ?</font>"

				celda = ofv.GenCelda("","tc","","","","","",sinputtt,"si")

				' Strsql = "SELEct * FROM tema where PSTIPO = 'C' order by inda, descri "

				strsql = "select distinct a.codtema,a.descri from ("
				strsql = strsql & " Select  isnull(cast (b.atrmverst as int),1) as st,b.atrmcodes ,a.* From Tema  a "
				if session("UN") <> 1   then
					strsql = strsql & " left join  atrmtemast  b on (a.codtema=b.atrmcodst and b.atrmtipost='T'  and b.atrmcodes=" & session("UN") & " ) "
				else
					strsql = strsql & " left join  atrmtemast  b on (a.codtema=b.atrmcodst and b.atrmtipost='T') "
				end if
				strsql = strsql & " where a.aplic = 'EM' or a.aplic = 'GLB' ) a where a.st > 0 order by a.descri "

               'response.write strsql
				scomK = ""
				scomD = "SELECCIONE CARPETA"

				sinputdtK = ""
				sinputdtd = "SELECCIONE CARPETA"
				sintab2   = ""

				set rsr = ofv.crearconsultaEx(Strsql,cn,1,parametros)

				do while not rsr.eof
					'' scomK = scomK & chr(9) & rsr(0)
					'' scomD = scomD & chr(9) & rsr(1)

					'  sintab2 = sintab2 & "<option value='" & rsr(0) & "' "


					'  if tcod  = cstr(rsr(0)) then
					'     tnam       = replace(rsr(1)," ","_")
					'     sinputdtd  = rsr(1)
					'     sintab2    = sintab2 & " selected "
					'  end if

					'  sintab2 = sintab2 & " >" & rsr(1) & "</option>"
					'scomK = scomK & chr(9) & rsr(0)
					' scomD = scomD & chr(9) & rsr(1)
					' SCOMSEP =""
					' tttt=""

					if tcod  = cstr(rsr(0)) then
						tnam       = replace(rsr(1)," ","_")
						sinput = rsr(1)
					end if

					scomK=scomK & SCOMSEP & rsr(0)
					scomd=scomd & SCOMSEP & rsr(1)
					SCOMSEP = chr(9)
					rsr.movenext
				loop
				'id del form, nombre del dato, el dato,'
				sInput =  ofv.gencomboFX("Subconsultax", "Ntcod",tcod, scomK, scomD, "no1")
				call ofv.cerrarconsulta(rsr)

				' rsr.movenext
				' loop
				' call ofv.cerrarconsulta(rsr)

				'    scomK = scomK & chr(9) & "99999"
				'    scomD = scomD & chr(9) & "* TODAS LAS CARPETAS"

				'    sintab2 = sintab2 & "<option value='99999' "

				'    if tcod  = "99999" then
				'       tnam       =  "TODAS_LAS_CARPETAS"
				'       sinputdtd  = "* TODAS LAS CARPETAS"
				'       sintab2    = sintab2 & " selected "
				'    end if

				'    sintab2 = sintab2 & " >* TODAS LAS CARPETAS</option>"


				'    sintab = "<table id='Ntcodtab1' width='70%' cellspacing=0 cellpadding=0 style='display:block;' >"
				'    sintab = sintab & "<td id='Ntcodtd1' width='100%' class='tdcm1' onclick=""opsel('Ntcod');"" >" & sinputdtd & "</td>"
				'    sintab = sintab & "<td id='Ntcodtd2' width='1'    class='tdcm2' onclick=""opsel('Ntcod');"" >"
				'    sintab = sintab & "<font class='btic' >6<font></td>"
				'    sintab = sintab & "</table>"


				'    sintab3 = "<select name='Ntcodsl' size='5' onchange=""setsel('Ntcod');"" "
				'    sintab3 = sintab3 & " style='display:none;border-style:solid;border-width:1;border-color:#888888;' >" & sintab2
				'    sintab3 = sintab3 & "</select>"


				'sinputdtK = ofv.GenerarInput("Ntcod",tcod,"hidden","","1","dt","")

				' sInput   = sintab & sintab3 & sinputdtK


				sInput1  = ofv.GenerarInput("tsel",tsel,"hidden","","1","dt","")
				sInput2  = ofv.GenerarInput("tgrp",tgrp,"hidden","","1","dt","")
				sInput3  = ofv.GenerarInput("ttip",ttip,"hidden","","1","dt","")
				sInput4  = ofv.GenerarInput("tnam",tnam,"hidden","","1","dt","")

				Celda = Celda & ofv.GenCelda("","ut  ","","","","","",sInput & sInput1 & sInput2 & sInput3 & sInput4,"si")
				'Celda = Celda & ofv.GenCelda("","ut  ","center","","","","",sInput1,"si")
				Fila  = fila & ofv.GenRow("","","","","30","","",Celda,"si")

			end if



			'''''''''' sector '''''''''''''''''''''''''''''''''''''''''''''''''


			if  xstdoc   = "1" and   xstsel   = "1" then

				if  stipodoc = "99999" or stipodoc = "PCD" then


					Strsql = "SELECT DISTINCT eco.eco,ECO.descri FROM OBJSECTOR, ECO,  OBJETOS "


					if SESSION("EMPMUN") = "S" THEN
						if SESSION("GLBMUN") = 2 THEN
							Strsql = Strsql & " ,OBJMEMPALC "
						END IF
					END IF

					Strsql = Strsql & " where OBJSECTOR.codsector = ECO.eco AND ECO.ECOSITE = " & SEMPRE
					Strsql = Strsql & " AND OBJSECTOR.CODOBJETO = OBJETOS.CODOBJETO  "

					if SESSION("EMPMUN") = "S" then
						Strsql = Strsql & " AND "
						if SESSION("GLBMUN") = 1 THEN
							Strsql = Strsql & " OBJETOS.OBJUN = " & SEMPRE
						ELSE
							Strsql = Strsql & " ( (OBJETOS.OBJUN = " & SEMPRE & ") OR "
							Strsql = Strsql & "   ( OBJETOS.OBJUN = 1 AND "
							Strsql = Strsql & "     OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO AND "
							Strsql = Strsql & "     OBJMEMPALC.CODEMPRE = " & SEMPRE & ") "
							Strsql = Strsql & " ) "
						END IF
					END IF


					Strsql  = Strsql & " ORDER BY ECO.descri"

              '     response.write strsql
					scomK = ""
					scomD = "SELECCIONE SECTOR"

					sinputdtK = ""
					sinputdtd = "SELECCIONE SECTOR"
					sintab2   = ""


					scomK = scomK & chr(9) & "99999"
					scomD = scomD & chr(9) & "* TODOS LOS SECTORES"

					sintab2 = sintab2 & "<option value='99999' "

					if ofv.convertircar(SsectorT," ","_")  = "99999" then
						sinputdtd = "* TODOS LOS SECTORES"
						sintab2    = sintab2 & " selected "
					end if

					sintab2 = sintab2 & " >* TODOS LOS SECTORES</option>"


					set rsr = ofv.crearconsultaEx(Strsql,cn,1,parametros)
					do until rsr.eof

						scomK = scomK & chr(9) & ofv.convertircar(rsr(0)," ","_")
						scomD = scomD & chr(9) & rsr(1)
						SCOMSEP =""
						tttt=""

						if ofv.convertircar(SsectorT," ","_")  = ofv.convertircar(rsr(0)," ","_") then
							sinput = rs(1)
						end if

						scomK=scomK & SCOMSEP & rsr(0)
						scomd=scomd & SCOMSEP & rsr(1)
						SCOMSEP = chr(9)
						rsr.movenext
					loop
					'id del form, nombre del dato, el dato,'
					sInput =  ofv.gencomboFX("Subconsultax", "Nlista",ofv.convertircar(SsectorT," ","_"), scomK, scomD, "no1")
					call ofv.cerrarconsulta(rsr)

					'' scomK = scomK & chr(9) & ofv.convertircar(rsr(0)," ","_")
					'' scomD = scomD & chr(9) & rsr(1)


					' sintab2 = sintab2 & "<option value='" & ofv.convertircar(rsr(0)," ","_") & "' "


					' if ofv.convertircar(SsectorT," ","_")  = ofv.convertircar(rsr(0)," ","_") then
					'    sinputdtd = rsr(1)
					'    sintab2    = sintab2 & " selected "
					' end if

					' sintab2 = sintab2 & " >" & rsr(1) & "</option>"


					'rsr.movenext
					'loop
					'call ofv.cerrarconsulta(rsr)


					'     sintab = "<table id='Nlistatab1' width='70%' cellspacing=0 cellpadding=0 style='display:block;' >"
					'     sintab = sintab & "<td id='Nlistatd1' width='100%' class='tdcm1' onclick=""opsel('Nlista');"" >" & sinputdtd & "</td>"
					'     sintab = sintab & "<td id='Nlistatd2' width='1'    class='tdcm2' onclick=""opsel('Nlista');"" >"
					'     sintab = sintab & "<font class='btic' >6<font></td>"
					'     sintab = sintab & "</table>"


					'     sintab3 = "<select name='Nlistasl' size='5' onchange=""setsel('Nlista');"" "
					'     sintab3 = sintab3 & " style='display:none;border-style:solid;border-width:1;border-color:#888888;' >" & sintab2
					'     sintab3 = sintab3 & "</select>"


					' sinputdtK = ofv.GenerarInput("Nlista",ofv.convertircar(SsectorT," ","_"),"hidden","","1","dt","")

					'  sInput   = sintab & sintab3 & sinputdtK

					sinputtt = "<font>¿Desea acotar la b&uacute;squeda a alg&uacute;n sector en particular ?</font>"

					celda = ofv.GenCelda("","tc","","","","","",sinputtt,"si")
					Celda = Celda & ofv.GenCelda("","ut","","","","","",sInput,"si")
					Fila  = fila & ofv.GenRow("","","","","30","","",Celda,"si")

				else

					SsectorT = "99999"

					xstsec   = "1"

				end if

			else

				SsectorT = "99999"

			end if

		END IF





		'''''''''''' texto a buscar '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		xentro="N"
		if  xstdoc   = "1" and   xstsel   = "1" and   xstsec   = "1"   then
			if TTCC="BASIC" then
				' sInput = "<b>Usted est&aacute en la consulta b&aacutesica</b>"
				'  Celda = ofv.GenCelda("","utg  style='border-style:solid;border-width:1;border-color:#000099;'","center","50%","","","",sInput,"si")
				'Fila  = fila & ofv.GenRow("","","","","30","","",Celda,"si")
			end if

			if textot = "" then
				textot =""'"INGRESE EL TEXTO A BUSCAR"
			end if
			xentro="S"
			if session("APLMDCLI")="BC" THEN
				sinputtt = "<font>Ingrese el texto a buscar</font>"
			else
				sinputtt = "<font>Ingrese el texto a buscar</font>"
				sInputxb1 = ofv.GenerarInput("bx1","BUSQUEDA AVANZADA","BUTTON","","1","dt style=""background-color:color-mix(in srgb, " & Session("CMFPNBOT") & ", transparent 80%);border-color:color-mix(in srgb, " & Session("CMFPNBOT") & ", transparent 20%)""","document.all.Subconsultax1.submit();")
			end if

			celda = ofv.GenCelda("","tc","","","","","",sinputtt,"si")


		end if

		sInputTTCC  = ofv.GenerarInput("TTCC",ttcc,"hidden","","1","dt","")


		sAccion = "Subconsulta.asp?frm=5&Menu=" & sMenu & "&Nombre=" & request.querystring("Nombre")
		Formulario = Formulario  & ofv.GenForm("Subconsultax","",sAccion,"","post","",Fila & sInputTTCC ,"si")



		Fila    = ""

		if tgrp  = "" then
			tgrp  = "T"
		end if

		sInput0  = ofv.GenerarInput("tdoc",stipodoc,"hidden","","1","dt","")
		sInput1  = ofv.GenerarInput("lista",sSectort,"hidden","","1","dt","")
		sInput2  = ofv.GenerarInput("tsel",tsel,"hidden","","1","dt","")
		sInput4  = ofv.GenerarInput("tgrp",tgrp,"hidden","","1","dt","")
		sInput5  = ofv.GenerarInput("tcod",tcod,"hidden","","1","dt","")
		sInput6  = ofv.GenerarInput("ttip",ttip,"hidden","","1","dt","")
		sInput7  = ofv.GenerarInput("tnam",tnam,"hidden","","1","dt","")



		if session("APLMDCLI")="BC" THEN
			sInput3     = ofv.GenerarInput("buscar",textot,"agregar","","50","dt","")
		ELSE
			sInput3     = ofv.GenerarInput("buscar",textot,"agregar","","50","dt ","")
		END IF






		'  if  textot <> "" then
		if xentro="S" then

			sInput     = "&nbsp;"
			Celda      = Celda & ofv.GenCelda("","UT","","","","","",sInput0 & sInput1 & sInput2 & sInput3 & sInput4 & sInput5 & sInput6 & sInput7,"si")

			Fila       = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

		end if





		formx    = Fila

		sTARGET  = ""
		sAccion    = "ListadoConsultas.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren




	elseif sNombre = "CONSULTA POR BUSCADOR" then










		if textot = "" then
			textot = ""' "TEXTO A BUSCAR2"
		end if

		sinputtt = "<font>Ingrese el texto a buscar</font>"

		sInput3     = ofv.GenerarInput("buscar",textot,"agregar","","50","dt","")

		Celda  = ofv.GenCelda("","tc","","","","","",sinputtt,"si")
		Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput3,"si")
		Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

		sInputTTCC   = ofv.GenerarInput("","Buscar","submit","","30","bt","")

		sImage  = ""'ofv.GenImage("","","","","IMAGENES\GOOGLE.jpg","0","")

		Celda  = ofv.GenCelda("","vt","","","","","",sImage,"si")
		Celda  = ofv.GenCelda("","ut","Center colspan=2","","","","",sInputTTCC,"si")
		Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

		'sAccion = "SubconsEXTERN.asp?frm=5&Menu=" & sMenu & "&Nombre=" & request.querystring("Nombre")
		If  SESSION("APLMDCLI") = "BPN" Or SESSION("APLMDCLI") = "BC" Or SESSION("APLMDCLI") = "BGAL" THEN
			sAccion = "SubconsEXTERN.asp?frm=5&Menu=PORBUSCADOR&Nombre=BUSCADOREX"
		ELSE
			sAccion = "consBuscadorSet.asp"
		END IF
		Formulario = Formulario  & ofv.GenForm("Subconsultax","",sAccion,"","post","",Fila ,"si")

		Fila  = ""
		sAccion = ""
		formx  = ""




	elseif sNombre = "CONSULTA POR CATEGORIAS" then

		sInput  = ""
		sInput2  = ""
		sInput3  = ""
		sInput4  = ""
		sInput5  = ""
		sInput6  = ""
		sInput7  = ""

		xstdoc   = "0"
		xstsel   = "0"
		xstsec   = "0"

		''response.write "<br><br><br>---------------------------"
		'for each x in request.form
		'   response.write "-- " & x & " : " & request.form(x) & " --<br>"
		'next
		''response.write "---------------------------"


		xx1 = request.form("Ntipodoc") & ""
		xx2 = request.form("Ntcod") & ""
		xx3 = request.form("Nlista") & ""
		xx4 = request.form("textot") & ""
		xx5 = request.form("tsel") & ""

		ttcc = request.form("ttcc") & ""





		stipodoc    = request.form("Ntipodoc") & ""
		if stipodoc = "" then
			'   stipodoc = "99999"
			xstdoc   = "0"
		else
			xstdoc   = "1"
		end if



		tsel     = request.form("tsel") & ""

		if tsel  = "" then
			tsel  = "C"
			tgrp  = "T"
			tcod  = "99999"
			ttip  = "TODOS"
			tnam  = "TEMA"
			xstsel   = "0"
		else
			xstsel   = "1"

			if tsel  = "" then
				tgrp  = "T"
			end if

		end if

		frmsen = "N"

		if tsel  = "T" then
			tgrp  = "M"
			tcod  = "99999"
			ttip  = "TODOS"
			tnam  = "TODOS"

		end if

		sSectort = request.form("Nlista") & ""

		sSectort    = ofv.convertircar(SsectorT,"_"," ")
		if ssectort = "" then
			'   ssectort = "99999"
			xstsec   = "0"
		else
			xstsec   = "1"
		end if

		textot = request.form("textot") & ""








		formul = request.querystring("frm")

		if len(formul) = 0 then

			tgrp  = request.form("tgrp") & ""
			tcod  = request.form("Ntcod") & ""
			ttip  = request.form("ttip") & ""
			tnam  = request.form("tnam") & ""


		elseif formul = 4 then

			if tsel  <> "T" then
				tgrp   = request.form("ptgrp") & ""
				tcod   = request.form("ptcod") & ""
				ttip   = request.form("pttip") & ""
				tnam   = request.form("ptnam") & ""
				textot = request.form("ptextot") & ""
			end if

		else

			if tsel  <> "T" then
				tgrp  = request.form("tgrp") & ""
				tcod  = request.form("Ntcod") & ""
				ttip  = request.form("ttip") & ""
				tnam  = request.form("tnam") & ""

			end if


		end if




		if  tgrp <> "M" and formul=4 then frmsen = "S"

		''''''''' introduccion '''''''''''''''''''''''''''''''''''''


		sInputxx1 = ofv.GenerarInput("Ntipodoc","99999","hidden","","1","dt","")
		sInputxx2 = ofv.GenerarInput("Ntcod","99999","hidden","","1","dt","")
		sInputxx3 = ofv.GenerarInput("Nlista","99999","hidden","","1","dt","")
		sInputxx5 = ofv.GenerarInput("tsel","C","hidden","","1","dt","")
		sInputxx6  = ofv.GenerarInput("tgrp","T","hidden","","1","dt","")
		sInputxx7  = ofv.GenerarInput("ttip","TODOS","hidden","","1","dt","")
		sInputxx8  = ofv.GenerarInput("tnam","TODOS","hidden","","1","dt","")




		sInputxx9b  = ofv.GenerarInput("ttcc","BASIC","hidden","","1","dt","")

		sInputxx9a  = ofv.GenerarInput("ttcc","AVANZADA","hidden","","1","dt","")


		sInputxb1 = ofv.GenerarInput("bx1","AVANZADA","BUTTON","","1","dt","document.all.Subconsultax1.submit();")
		sInputxb2 = ofv.GenerarInput("bx2","BASICA","BUTTON","","1","dt","document.all.Subconsultax2.submit();")






		sAccion = "Subconsulta.asp?Menu=" & sMenu & "&Nombre=" & request.querystring("Nombre")
		Formuln2 = ofv.GenForm("Subconsultax2","",sAccion,"","post","",sInputxx1 & sInputxx2 & sInputxx3 & sInputxx5 & sInputxx6 & sInputxx7 & sInputxx8 & sInputxx9B,"si")


		sAccion = "Subconsulta.asp?Menu=" & sMenu & "&Nombre=" & request.querystring("Nombre")
		Formuln1 = ofv.GenForm("Subconsultax1","",sAccion,"","post","",sInputxx5 & sInputxx6 & sInputxx7 & sInputxx8 & sInputxx9A,"si")


		sHTML = sHTML & Formuln2 & Formuln1


		celdan = ""
		celdan = celdan & ofv.GenCelda("","ut   ","","","","","",sInputxb1,"si")
		celdan = celdan & ofv.GenCelda("","ut   ","","","","","",sInputxb2,"si")
		Filan  = ofv.GenRow("","","","","","","",celdan,"si")
		Filan  = ofv.GenTabla("","aaa style='border-style:solid;border-width:1;border-color:#dddddd;display:none;' ","","20% ","","0","","0","0","","",Filan,"si")


		celda = ofv.GenCelda("","ut COLSPAN=2  ","RIGHT","","","","",Filan,"si")
		Fila  = fila & ofv.GenRow("","","","","30","","",Celda,"si")



		sglmar = "<font id='fnt01' size=4 color='#f7aa22' face='CorFlatIcon' onclick='alert(""Seleccione los parametros para la busqueda...."");' >&#0125;</font>&nbsp;"

		sglpnt = "<font size=2 color='#00aa00' face='CorFlatIcon' >4</font>&nbsp;"

		sinput = ""


		sintab = "<script language='javascript' > "
		sintab = sintab & " function opsel(dtfld) {document.all[dtfld + 'tab1'].style.display='none';document.all[dtfld + 'sl'].style.display='block';} "
		sintab = sintab & " function setsel(dtfld) {"
		sintab = sintab & "   document.all.Subconsultax[dtfld].value=document.all.Subconsultax[dtfld + 'sl'].value;"
		sintab = sintab & "   document.all[dtfld +'td1'].innerHTML=document.all.Subconsultax[dtfld + 'sl'].options[document.all.Subconsultax[dtfld + 'sl'].selectedIndex].text;"
		sintab = sintab & "   document.all[dtfld + 'tab1'].style.display='block';document.all[dtfld + 'sl'].style.display='none';"
		sintab = sintab & "   document.all.Subconsultax.submit();"
		sintab = sintab & "                    } "
		sintab = sintab & "</script>"

		sinput = sinput & sintab

		celda = ofv.GenCelda("","utg COLSPAN=2 style='text-align:justify;' ","","","","","",sinput,"si")
		Fila  = fila & ofv.GenRow("","","","","30","","",Celda,"si")





		if ttcc <> "BASIC" THEN



			''''''''' tipo de RIESGO (SITUACION) '''''''''''''''''''''''''''''''''''''

			sinputtt = "<font>SITUACION</font>"

			celda = ofv.GenCelda("","uT ","","50%","","","",sinputtt,"si")


			Strsql = "SELEct * FROM RSKSITUACION  order by descri "

			scomK = ""
			scomD = "SELECCIONE SITUACION"

			sinputdtK = ""
			sinputdtd = "SELECCIONE SITUACION"
			sintab2   = ""

			set rsr = ofv.crearconsultaEx(Strsql,cn,1,parametros)
			do until rsr.eof

				scomK = scomK & chr(9) & rsr(0)
				scomD = scomD & chr(9) & rsr(1)

				sintab2 = sintab2 & "<option value='" & rsr(0) & "' "


				if stipodoc  = cstr(rsr(0)) then
					sinputdtd = rsr(1)
					sintab2    = sintab2 & " selected "
				end if

				sintab2 = sintab2 & " >" & rsr(1) & "</option>"

				rsr.movenext
			loop
			call ofv.cerrarconsulta(rsr)

			'      scomK = scomK & chr(9) & "99999"
			'      scomD = scomD & chr(9) & "* CUALQUIER TIPO DE OBJETO"

			'      sintab2 = sintab2 & "<option value='99999' "

			'      if stipodoc  = "99999" then
			'         sinputdtd = "* CUALQUIER TIPO DE OBJETO"
			'         sintab2    = sintab2 & " selected "
			'      end if

			'      sintab2 = sintab2 & " >* CUALQUIER TIPO DE OBJETO</option>"


			sintab = "<table id='Ntipodoctab1' width='70%' cellspacing=0 cellpadding=0 style='display:block;' >"
			sintab = sintab & "<td id='Ntipodoctd1' width='100%' class='tdcm1' onclick=""opsel('Ntipodoc');"" >" & sinputdtd & "</td>"
			sintab = sintab & "<td id='Ntipodoctd2' width='1'    class='tdcm2' onclick=""opsel('Ntipodoc');"" >"
			sintab = sintab & "<font class='btic'  >6<font></td>"
			sintab = sintab & "</table>"


			sintab3 = "<select name='Ntipodocsl' size='5' onchange=""setsel('Ntipodoc');"" "
			sintab3 = sintab3 & " style='display:none;border-style:solid;border-width:1;border-color:#888888;' >" & sintab2
			sintab3 = sintab3 & "</select>"


			sinputdtK = ofv.GenerarInput("Ntipodoc",stipodoc,"hidden","","1","dt","")

			sInput   = sintab & sintab3 & sinputdtK

			Celda = Celda & ofv.GenCelda("","ut  ","center","50%","","","",sInput,"si")
			Fila  = fila & ofv.GenRow("","","","","30","","",Celda,"si")





			'''''''''''' CATEGORIAS ''''''''''''''''''''''''''''''''''''''''''''''''

			if  xstdoc   = "1" then

				sinputtt = "<font>CATEGORIA</font>"

				celda = ofv.GenCelda("","uT ","","","","","",sinputtt,"si")

				Strsql = "SELEct * FROM RSKCATEGO WHERE SITUACION = 0" & stipodoc & " order by descri "

				scomK = ""
				scomD = "SELECCIONE CATEGORIA"

				sinputdtK = ""
				sinputdtd = "SELECCIONE CATEGORIA"
				sintab2   = ""

				set rsr = ofv.crearconsultaEx(Strsql,cn,1,parametros)
				do until rsr.eof

					scomK = scomK & chr(9) & rsr(0)
					scomD = scomD & chr(9) & rsr(2)

					sintab2 = sintab2 & "<option value='" & rsr(0) & "' "


					if tcod  = cstr(rsr(0)) then
						tnam       = replace(rsr(2)," ","_")
						sinputdtd  = rsr(2)
						sintab2    = sintab2 & " selected "
					end if

					sintab2 = sintab2 & " >" & rsr(2) & "</option>"


					rsr.movenext
				loop
				call ofv.cerrarconsulta(rsr)

				'    scomK = scomK & chr(9) & "99999"
				'    scomD = scomD & chr(9) & "* TODAS LAS CARPETAS"

				'    sintab2 = sintab2 & "<option value='99999' "

				'    if tcod  = "99999" then
				'       tnam       =  "TODAS_LAS_CARPETAS"
				'       sinputdtd  = "* TODAS LAS CARPETAS"
				'       sintab2    = sintab2 & " selected "
				'    end if

				'    sintab2 = sintab2 & " >* TODAS LAS CARPETAS</option>"


				sintab = "<table id='Ntcodtab1' width='70%' cellspacing=0 cellpadding=0 style='display:block;' >"
				sintab = sintab & "<td id='Ntcodtd1' width='100%' class='tdcm1' onclick=""opsel('Ntcod');"" >" & sinputdtd & "</td>"
				sintab = sintab & "<td id='Ntcodtd2' width='1'    class='tdcm2' onclick=""opsel('Ntcod');"" >"
				sintab = sintab & "<font class='btic' >6<font></td>"
				sintab = sintab & "</table>"


				sintab3 = "<select name='Ntcodsl' size='5' onchange=""setsel('Ntcod');"" "
				sintab3 = sintab3 & " style='display:none;border-style:solid;border-width:1;border-color:#888888;' >" & sintab2
				sintab3 = sintab3 & "</select>"


				sinputdtK = ofv.GenerarInput("Ntcod",tcod,"hidden","","1","dt","")

				sInput   = sintab & sintab3 & sinputdtK


				sInput1  = ofv.GenerarInput("tsel",tsel,"hidden","","1","dt","")
				sInput2  = ofv.GenerarInput("tgrp",tgrp,"hidden","","1","dt","")
				sInput3  = ofv.GenerarInput("ttip",ttip,"hidden","","1","dt","")
				sInput4  = ofv.GenerarInput("tnam",tnam,"hidden","","1","dt","")

				Celda = Celda & ofv.GenCelda("","ut  ","center","","","","",sInput & sInput1 & sInput2 & sInput3 & sInput4,"si")
				Fila  = fila & ofv.GenRow("","","","","30","","",Celda,"si")

			end if



			'''''''''' sector '''''''''''''''''''''''''''''''''''''''''''''''''


			if  xstdoc   = "1" and   xstsel   = "1" then

				' if  stipodoc = "99999" or stipodoc = "PCD" then

				Strsql = "SELEct * FROM RSKEVENTO WHERE CATEGO = 0" & tcod & " order by descri "

				scomK = ""
				scomD = "SELECCIONE EVENTO"

				sinputdtK = ""
				sinputdtd = "SELECCIONE EVENTO"
				sintab2   = ""





				set rsr = ofv.crearconsultaEx(Strsql,cn,1,parametros)
				do until rsr.eof

					scomK = scomK & chr(9) & ofv.convertircar(rsr(0)," ","_")
					scomD = scomD & chr(9) & rsr(3)


					sintab2 = sintab2 & "<option value='" & ofv.convertircar(rsr(0)," ","_") & "' "


					if ofv.convertircar(SsectorT," ","_")  = ofv.convertircar(rsr(0)," ","_") then
						sinputdtd = rsr(3)
						sintab2    = sintab2 & " selected "
					end if

					sintab2 = sintab2 & " >" & rsr(3) & "</option>"


					rsr.movenext
				loop
				call ofv.cerrarconsulta(rsr)


				sintab = "<table id='Nlistatab1' width='70%' cellspacing=0 cellpadding=0 style='display:block;' >"
				sintab = sintab & "<td id='Nlistatd1' width='100%' class='tdcm1' onclick=""opsel('Nlista');"" >" & sinputdtd & "</td>"
				sintab = sintab & "<td id='Nlistatd2' width='1'    class='tdcm2' onclick=""opsel('Nlista');"" >"
				sintab = sintab & "<font class='btic' >6<font></td>"
				sintab = sintab & "</table>"


				sintab3 = "<select name='Nlistasl' size='5' onchange=""setsel('Nlista');"" "
				sintab3 = sintab3 & " style='display:none;border-style:solid;border-width:1;border-color:#888888;' >" & sintab2
				sintab3 = sintab3 & "</select>"


				sinputdtK = ofv.GenerarInput("Nlista",ofv.convertircar(SsectorT," ","_"),"hidden","","1","dt","")

				sInput   = sintab & sintab3 & sinputdtK

				sinputtt = "<font>EVENTO DE PERDIDAS</font>"

				celda = ofv.GenCelda("","uT ","","","","","",sinputtt,"si")
				Celda = Celda & ofv.GenCelda("","ut  ","center","","","","",sInput,"si")
				Fila  = fila & ofv.GenRow("","","","","30","","",Celda,"si")



			else

				SsectorT = "99999"

			end if

		END IF



		'''''''''''' texto a buscar '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

		if  xstdoc   = "1" and   xstsel   = "1" and   xstsec   = "1"  then

			if textot = "" then
				textot = " "
			end if

			sinputtt = "<font>CONFIRMA SELECCION DE BUSQUEDA?</font>"

			celda = ofv.GenCelda("","uT ","","","","","",sinputtt,"si")


		end if


		sAccion = "Subconsulta.asp?frm=5&Menu=" & sMenu & "&Nombre=" & request.querystring("Nombre")
		Formulario = Formulario  & ofv.GenForm("Subconsultax","",sAccion,"","post","",Fila,"si")

		'' END IF

		Fila    = ""

		if tgrp  = "" then
			tgrp  = "T"
		end if

		sInput0  = ofv.GenerarInput("tdoc",stipodoc,"hidden","","1","dt","")
		sInput1  = ofv.GenerarInput("lista",sSectort,"hidden","","1","dt","")
		sInput2  = ofv.GenerarInput("tsel",tsel,"hidden","","1","dt","")
		sInput4  = ofv.GenerarInput("tgrp",tgrp,"hidden","","1","dt","")
		sInput5  = ofv.GenerarInput("tcod",tcod,"hidden","","1","dt","")
		sInput6  = ofv.GenerarInput("ttip",ttip,"hidden","","1","dt","")
		sInput7  = ofv.GenerarInput("tnam",tnam,"hidden","","1","dt","")




		sInput3     = ofv.GenerarInput("buscar",textot,"HIDDEN","","50","dt","")





		if  textot <> "" then

			sInput     = "&nbsp;"
			Celda      = Celda & ofv.GenCelda("","UT","center","","","","",sInput0 & sInput1 & sInput2 & sInput3 & sInput4 & sInput5 & sInput6 & sInput7,"si")
			Fila       = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

		end if



		formx    = Fila



		sTARGET  = ""
		sAccion    = "ListadoConsultas.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren


		'******************CONSULTA FOR FECHA *********************************
	elseif sNombre = "CONSULTA POR FECHA" then


		XFNAME   = request.querystring("XFNAME")
		XFSRH    = request.querystring("XFSRH")
		XFSRD    = request.querystring("XFSRD")



		IF XFNAME = "" THEN XFNAME = "PV"
		IF XFSRD  = "" THEN XFSRD = 0
		IF XFSRH  = "" THEN XFSRH = 0

		' PRIMERA VEZ
		IF XFNAME="PV"  THEN
	     XFSRD = clng(fechaconvert(DATE())) - clng(cdate("01-01-1900"))
			XFSRH = XFSRD '+ 15

		END IF

		IF XFNAME = "FD" THEN
			XFSRD = request.form("DESDE")
			XFSRH = XFSRH

		END IF

		IF XFNAME = "FH" THEN
			XFSRD = XFSRD
			XFSRH =request.form("HASTA")

		END IF






		sImagenE = "<B><i CLASS='las la-calendar' STYLE=font-size:14pt;COLOR:" & Session("CMFPNBOT") & "></i></B>"
		sAccion = "'javascript:void(0);'" & GTCSETCALR("frmCalendarioD","Desde",GTCSMT,GTCFCH,GTCFCHDAT,XFSRD,15,2,0,0,0)
		sinputE = "&nbsp;&nbsp;<a  class=UT href=" & sAccion
		sinputE = sinputE  & " TITLE='MODIFICAR FECHA' >" & sImagenE & "</a>"


		sAccionHasta = "'javascript:void(0);'" & GTCSETCALR("frmCalendarioH","Hasta",GTCSMT,GTCFCH,GTCFCHDAT,XFSRH,4,2,0,0,0)
		sinputHasta = "&nbsp;&nbsp;<a  class=UT href=" & sAccionHasta
		sinputHasta = sinputHasta  & " TITLE='MODIFICAR FECHA' >" & sImagenE & "</a>"



		' CELDAS DE DESDE

		if  clng(XFSRD) = 0 then
			sInput2 = "------------"
		else
			sInput2 = cdate(clng(XFSRD) + clng(cdate("01-01-1900")) )
			swd = datepart("w",sInput2)
			sInput2 = Fechador(XFSRD)
		END IF




		Celda  = ofv.GenCelda("","tc","","","","","","<font>Desde</font>","si")
		sInput = ofv.GenerarInput("Desde2",sInput2,"agregar","","80","dt","") & ofv.GenerarInput("Desde3",cstr(cdate(XFSRD + clng(cdate("01-01-1900")))),"hidden","","20","dt","")
		Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput&sinputE,"si")



		Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")


		' CELDAS DE HASTA
		if clng(XFSRH) = 0 then
			sInput2 = "------------"
		else
	           sInput2 = fechaConvert(clng(XFSRH) + clng(cdate("01-01-1900")) )
			swd = datepart("w",sInput2)
			sInput2 = Fechador(XFSRH)
		END IF



		IF XFSRD >XFSRH THEN
			msg=" onhover='validarfechas();'"
		else
			msg=""
		END IF

		Celda  = ofv.GenCelda("","tc","","","","","","<font>Hasta</font>","si")
		sInput = ofv.GenerarInput("Hasta2",sInput2,"agregar","","80","dt","") & ofv.GenerarInput("Hasta3",cdate(XFSRH+ clng(cdate("01-01-1900"))),"hidden","","20","dt","")
		Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput&sinputHasta,"si")

		Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

		Celda  = ofv.GenCelda("","tc","","","","","","<font>Tipo</font>","si")
		scomK = "Emision" & chr(9) & "Vigencia"  & chr(9) & "Publicacion"  & chr(9) & "Vencimiento" & chr(9) & "Eliminacion"
		scomD = "Emision" & chr(9) & "Vigencia"  & chr(9) & "Publicacion"  & chr(9) & "Revision" & chr(9) & "Eliminacion"
		sInput = ofv.gencomboFX("subc","fecha","Emision",scomK,scomD,"no")
		Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
		Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")






		formx    = Fila
		sTARGET  = ""
		sAccion    = "ListadoConsultas.asp?Menu=" & "PORFECHA" & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren








		'************************************

	ELSEIF sNombre = "CONSULTA POR ACTA" then
		'RESPONSE.WRITE (CHR(13) & "ENTRO" & CHR(13))



		XFNAME   = request.querystring("XFNAME")
		XFSRH    = request.querystring("XFSRH")
		XFSRD    = request.querystring("XFSRD")



		IF XFNAME = "" THEN XFNAME = "PV"
		IF XFSRD  = "" THEN XFSRD = 0
		IF XFSRH  = "" THEN XFSRH = 0

		' PRIMERA VEZ
		IF XFNAME="PV"  THEN
			XFSRD = 0 'clng(DATE()) - clng(cdate("01-01-1900"))
			XFSRH = XFSRD '+ 15
			'response.write XFSRD
		END IF

		IF XFNAME = "FD" THEN
			XFSRD = request.form("DESDE")
			XFSRH = XFSRH
			if XFSRH < XFSRD then XFSRH=XFSRD
		END IF

		IF XFNAME = "FH" THEN

			XFSRD = XFSRD
			XFSRH =request.form("HASTA")


		END IF



		' CELDA NRO DE ACTA




		sImagenE = "<B><i CLASS='las la-calendar' STYLE=font-size:14pt;COLOR:" & Session("CMFPNBOT") & "></i></B>"
		sAccion = "'javascript:void(0);'" & GTCSETCALR("frmCalendarioD","Desde",GTCSMT,GTCFCH,GTCFCHDAT,XFSRD,4,2,0,0,0)
		sinputE = "&nbsp;&nbsp;<a  class=UT href=" & sAccion
		sinputE = sinputE  & " TITLE='MODIFICAR FECHA' >" & sImagenE & "</a>"


		sAccionHasta = "'javascript:void(0);'" & GTCSETCALR("frmCalendarioH","Hasta",GTCSMT,GTCFCH,GTCFCHDAT,XFSRH,4,2,0,0,0)
		sinputHasta = "&nbsp;&nbsp;<a  class=UT href=" & sAccionHasta
		sinputHasta = sinputHasta  & " TITLE='MODIFICAR FECHA' >" & sImagenE & "</a>"


		' CELDAS DE DESDE

		if  clng(XFSRD) = 0 then
			sInput2 = "------------"
		else
		   sInput2 = fechaconvert(clng(XFSRD) + clng(cdate("01-01-1900")) )
			swd = datepart("w",sInput2)
			sInput2 = Fechador(XFSRD)
		END IF




		Celda  =  ofv.GenCelda("","tc","","","","","","<font>Desde</font>","si")
		sInput = ofv.GenerarInput("Desde2",sInput2,"agregar","","80","dt","") & ofv.GenerarInput("Desde3",cstr(cdate(XFSRD + clng(cdate("01-01-1900")))),"hidden","","20","dt","")
		'' sInput = ofv.GenerarInput("Desde2",sInput2,"agregar","","40","dt","") & ofv.GenerarInput("Desde3",XFSRH,"hidden","","20","dt","")
		Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput&sinputE,"si")



		Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")


		' CELDAS DE HASTA
		if clng(XFSRH) = 0 then
			sInput2 = "------------"
		else
			sInput2 = cdate(clng(XFSRH) + clng(cdate("01-01-1900")) )
			swd = datepart("w",sInput2)
			sInput2 = Fechador(XFSRH)
		END IF

		Celda  = ofv.GenCelda("","tc","","","","","","<font>Hasta</font>","si")
		sInput = ofv.GenerarInput("Hasta2",sInput2,"agregar","","80","dt","") & ofv.GenerarInput("Hasta3",cdate(XFSRH+ clng(cdate("01-01-1900"))),"hidden","","20","dt","")
		'sInput = ofv.GenerarInput("Hasta2",sInput2,"agregar","","40","dt","") & ofv.GenerarInput("Hasta3",XFSRH,"hidden","","20","dt","")
		Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput&sinputHasta,"si")

		'Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

		Fila       = Fila & ofv.GenRow("","","","","30","","",Celda,"si")
		Celda      = ofv.GenCelda("","tc","","","","","","<font>Nro.Acta</font>","si")
		sInput     = ofv.GenerarInput("nacta","","agregar","","20","dt","")
		Celda      = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
		Fila       = Fila & ofv.GenRow("","","","","30","","",Celda,"si")


		'Celda      = ofv.GenCelda("","tc","","","","","","<font>Desde</font>","si")
		'sInput     = ofv.GenerarInput("desde","","agregar","","20","dt","")
		'Celda      = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
		'Fila       = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

		'Celda      = ofv.GenCelda("","tc","","","","","","<font>Hasta</font>","si")
		'sInput     = ofv.GenerarInput("hasta","","agregar","","20","dt","")
		'Celda      = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
		'Fila       = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

		formx    = Fila

		sTARGET  = ""
		sAccion    = "ListadoConsultas.asp?Menu=PORACTA" & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren
		'sAccion    = "ListadoConsultas.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren




		'************************************

	ELSEIF sNombre = "CONSULTA POR RESUMEN" then
		'RESPONSE.WRITE (CHR(13) & "ENTRO" & CHR(13))



		XFNAME   = request.querystring("XFNAME")
		XFSRH    = request.querystring("XFSRH")
		XFSRD    = request.querystring("XFSRD")



		IF XFNAME = "" THEN XFNAME = "PV"
		IF XFSRD  = "" THEN XFSRD = 0
		IF XFSRH  = "" THEN XFSRH = 0

		' PRIMERA VEZ
		IF XFNAME="PV"  THEN
			XFSRD = clng(DATE()) - clng(cdate("01-01-1900"))
			XFSRH = XFSRD '+ 15

		END IF

		IF XFNAME = "FD" THEN
			XFSRD = request.form("DESDE")
			XFSRH = XFSRH

		END IF

		IF XFNAME = "FH" THEN
			XFSRD = XFSRD
			XFSRH =request.form("HASTA")

		END IF



		' ////////////////////// FECHA PUBLICACION




		sImagenE = "<B><i CLASS='las la-calendar' STYLE=font-size:14pt;COLOR:" & Session("CMFPNBOT") & "></i></B>"
		sAccion = "'javascript:void(0);'" & GTCSETCALR("frmCalendarioD","Desde",GTCSMT,GTCFCH,GTCFCHDAT,XFSRD,4,2,0,0,0)
		sinputE = "&nbsp;&nbsp;<a  class=UT href=" & sAccion
		sinputE = sinputE  & " TITLE='MODIFICAR FECHA' >" & sImagenE & "</a>"


		sAccionHasta = "'javascript:void(0);'" & GTCSETCALR("frmCalendarioH","Hasta",GTCSMT,GTCFCH,GTCFCHDAT,XFSRH,4,2,0,0,0)
		sinputHasta = "&nbsp;&nbsp;<a  class=UT href=" & sAccionHasta
		sinputHasta = sinputHasta  & " TITLE='MODIFICAR FECHA' >" & sImagenE & "</a>"


		' CELDAS DE DESDE

		if  clng(XFSRD) = 0 then
			sInput2 = "------------"
		else
			sInput2 = cdate(clng(XFSRD) + clng(cdate("01-01-1900")) )
			swd = datepart("w",sInput2)
			sInput2 = Fechador(XFSRD)
		END IF




		Celda  =  ofv.GenCelda("","tc","","","","","","<font>Publicados desde</font>","si")
		sInput = ofv.GenerarInput("Desde2",sInput2,"agregar","","80","dt","") & ofv.GenerarInput("Desde3",cstr(cdate(XFSRD + clng(cdate("01-01-1900")))),"hidden","","20","dt","")
		'' sInput = ofv.GenerarInput("Desde2",sInput2,"agregar","","40","dt","") & ofv.GenerarInput("Desde3",XFSRH,"hidden","","20","dt","")
		Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput&sinputE,"si")



		Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")


		' CELDAS DE HASTA
		if clng(XFSRH) = 0 then
			sInput2 = "------------"
		else
			sInput2 = cdate(clng(XFSRH) + clng(cdate("01-01-1900")) )
			swd = datepart("w",sInput2)
			sInput2 = Fechador(XFSRH)
		END IF

		Celda  = ofv.GenCelda("","tc","","","","","","<font>Publicados Hasta</font>","si")
		sInput = ofv.GenerarInput("Hasta2",sInput2,"agregar","","80","dt","") & ofv.GenerarInput("Hasta3",cdate(XFSRH+ clng(cdate("01-01-1900"))),"hidden","","20","dt","")
		'sInput = ofv.GenerarInput("Hasta2",sInput2,"agregar","","40","dt","") & ofv.GenerarInput("Hasta3",XFSRH,"hidden","","20","dt","")
		Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput&sinputHasta,"si")

		'Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

		Fila       = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

		Celda      = ofv.GenCelda("","ut","","","","","","&nbsp;","si")
		sInput     = ofv.GenerarInput("nacta","","hidden","","20","dt","")
		Celda      = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
		Fila       = Fila & ofv.GenRow("","","","","30","","",Celda,"si")



		formx    = Fila

		sTARGET  = ""
		sAccion    = "ListadoConsultas.asp?Menu=PORRESUMEN" & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren
		'sAccion    = "ListadoConsultas.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren





	elseif sNombre = "CONSULTA POR TEMA" then

		response.redirect "OBJtreeTema.asp?Tipo=Consulta&Nombre=TEMAS&Menu=Consulta&EMPRE=" & SEMPRE

		'Listas de Complementos
	elseif sNombre = "CONSULTA POR COMPONENTE" then


		Celda      = ofv.GenCelda("","tc","","","","","","<font>Descripcion</font>","si")
		sInput     = ofv.GenerarInput("buscar","","agregar","","60","dt","")
		Celda      = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
		Fila       = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

		Celda    = ofv.GenCelda("","ut","","","","","","&nbsp;","si")
		sInput   = ofv.GenerarInput("","Consultar","submit","","10","bt","")
		sInput   = sinput & ofv.GenerarInput("Tcons",sMenu,"hidden","","1","dt","")
		Celda    = ofv.GenCelda("","ut","Center colspan=2","","","","",sInput,"si")
		Fila     = Fila & ofv.GenRow("","","","","40","","",Celda,"si")

		Celda    = ofv.GenCelda("","ut colspan=2","","","","","","<hr size=1 width=100% >","si")
		Fila     = Fila & ofv.GenRow("","","","","10","","",Celda,"si")

		sAccion    = "ListadoConsultas.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0" & "&Nombre=" & sNombren
		Formulario = Formulario & ofv.GenForm("buscar1","",sAccion,"","post","",Fila,"si")

		Fila    = ""


		SQL      = "Select descri,tipo From Tipocomp Order By descri"
		Celda    = ofv.GenCelda("","tc","","","","","","<font>Tipo</font>","si")
		''   sInput   = ofv.GenerarComboSP("radio","XComplem",xComp,SQL,cn,1,0,"XComplem")
		Set rsPerfil = ofv.crearconsultaEx(SQL, cn, 1, parametros)
		'sInput = ofv.GenerarCombo("edProcedimientos",rs(6).name,rs(6).value,Strsql,cn,0,1,"sb")
		scomK = ""'"null" & Chr(9)

		scomD = "" '"" & Chr(9)
		SCOMSEP =""
		do while not rsperfil.eof
			scomK=scomK & SCOMSEP & rsperfil(1)
			scomd=scomd & SCOMSEP & rsperfil(0)
			SCOMSEP = chr(9)
			rsPerfil.movenext
		loop

		sInput =  ofv.gencomboFX("radio", "XComplem", xComp, scomK, scomD, "")
		'sInput   = replace(sInput,"onchage=","onclick=")
		Celda    = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
		Fila     = Fila & ofv.GenRow("","","","","30","","",Celda,"si")


		sAccion = "Subconsulta.asp?Menu=" & sMenu & "&Nombre=" & request.querystring("Nombre")
		Formulario = Formulario  & ofv.GenForm("radio","",sAccion,"","post","",Fila,"si")

		Fila    = ""



		Celda    = ofv.GenCelda("","tc","","","","","","<font>Componente</font>","si")

		Strsql   = "SELECT DISTINCT COMPLEMENTOS.codcomp, COMPLEMENTOS.codiden, "
		Strsql   = Strsql & "COMPLEMENTOS.descri FROM COMPLEMENTOS, "
		Strsql   = Strsql & " objlnk where COMPLEMENTOS.codcomp = objlnk.codobjetol and objlnk.ltipo = 'KL' "
		Strsql   = Strsql & "and Complementos.tipocomp = '" & xComp & "' "
		Strsql   = Strsql & "ORDER BY COMPLEMENTOS.descri"

		sInput   = ofv.GenerarLista(Strsql,cn,2)
		Celda    = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
		Fila     = Fila & ofv.GenRow("","","","","30","","",Celda,"si")
		'response.write strsql

		formx    = Fila

		sTARGET  = ""
		sAccion    = "ListadoConsultas.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren


		'Listas de Sectores
	elseif sNombre = "CONSULTA POR SECTOR" then


		'response.write request.querystring & "<br>" & request.form

		sSectort = request.form("lista") & ""
		sSectort = ofv.convertircar(SsectorT,"_"," ")

		Strsql = "SELECT DISTINCT eco.eco,ECO.descri FROM OBJSECTOR "
		Strsql = Strsql & " inner join ECO     on (OBJSECTOR.codsector = ECO.eco) "
		Strsql = Strsql & " inner join OBJETOS on (OBJSECTOR.CODOBJETO = OBJETOS.CODOBJETO) "


		if SESSION("EMPMUN") = "S" THEN
			if SESSION("GLBMUN") = 2 THEN
				Strsql = Strsql & " left join OBJMEMPALC on (OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO) "
			END IF
		END IF

		Strsql = Strsql & " WHERE eco.eco > 0 "
		''''''      Strsql = Strsql & " where  ECO.ECOSITE = " & SEMPRE
		''''''      Strsql = Strsql & " AND OBJSECTOR.CODOBJETO = OBJETOS.CODOBJETO  "




		if SESSION("EMPMUN") = "S" then
			Strsql = Strsql & " AND "
			if SESSION("GLBMUN") = 1 THEN
				Strsql = Strsql & " OBJETOS.OBJUN = " & SEMPRE
			ELSE
				Strsql = Strsql & " ( (OBJETOS.OBJUN = " & SEMPRE & ") OR "
				Strsql = Strsql & "   ( OBJETOS.OBJUN = 1 AND "
				'''''    Strsql = Strsql & "     OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO AND "
				Strsql = Strsql & "     OBJMEMPALC.CODEMPRE = " & SEMPRE & ") "
				Strsql = Strsql & " ) "
			END IF
		END IF

		Strsql2 = Strsql & " AND ECO.eco = '" &  sSectort & "' "
		Strsql  = Strsql & " ORDER BY ECO.descri"
		'response.write strsql

		set rsr = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
		if not rsr.eof then
			sSectort=rsr(0)
		end if
		call ofv.cerrarconsulta(rsr)

		if session("un") <> 1 or SESSION("EMPMUN") = "N" then
			strsql = "select eco,descri from eco where ecoUN=" & session("UN") & " order by descri"
		else
			strsql = "select a.eco,b.cliente + ' - ' + a.descri from eco a "
			strsql = strsql & " inner join clienteplan b on (a.ecoUN=b.id) "
			strsql = strsql & " order by b.cliente,a.descri"
		end if
		sInput   = ofv.GenerarComboSP("Subconsulta1","lista",sSectort,Strsql,cn,0,1," ")


		celda = ofv.GenCelda("","tc","","","","","","<font>Sector</font>","si")
		Celda = celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
		Fila  = fila & ofv.GenRow("","","","","30","","",Celda,"si")

		sAccion = "Subconsulta.asp?Menu=" & sMenu & "&Nombre=" & request.querystring("Nombre") & "&OPCION=A"
		Formulario = Formulario  & ofv.GenForm("Subconsulta1","",sAccion,"","post","",Fila,"si")

		Fila    = ""


		x2sql="SELECT OBJSECTOR.objid FROM OBJSECTOR, ECO, OBJETOS "


		if SESSION("EMPMUN") = "S" THEN
			if SESSION("GLBMUN") = 2 THEN
				x2sql=x2sql & " ,OBJMEMPALC "
			END IF
		END IF

		x2sql=x2sql & " where OBJSECTOR.codsector = ECO.eco "
		x2sql=x2sql & " AND OBJSECTOR.CODOBJETO = OBJETOS.CODOBJETO "

		if SESSION("EMPMUN") = "S" then
			x2sql=x2sql & " AND "

			if SESSION("GLBMUN") = 1 THEN
				x2sql=x2sql & " OBJETOS.OBJUN = " & SEMPRE
			ELSE
				if SEMPRE = 1 and sSectort <> "" then
					sql = "select ecoUN from eco where eco=" & sSectort
					'response.write sql
					set rx = ofv.crearconsultaEx(sql,cn,1,parametros)
					if not rx.eof then SEMPRE=rx(0)
					call ofv.cerrarconsulta(rx)
				end if
				x2sql=x2sql & " ( (OBJETOS.OBJUN = " & SEMPRE & ") OR "
				x2sql=x2sql & "   ( OBJETOS.OBJUN = 1 AND "
				x2sql=x2sql & "     OBJETOS.CODOBJETO = OBJMEMPALC.CODOBJETO AND "
				x2sql=x2sql & "     OBJMEMPALC.CODEMPRE = " & SEMPRE & ") "
				x2sql=x2sql & " ) "
			END IF
		END IF
		x2sql=x2sql & " and ECO.eco='" & ofv.convertircar(SsectorT,"_"," ") & "'"
		''   response.write x2sql & "<BR>"
		set rs = ofv.crearconsultaEx(x2sql,cn,1,parametros)
		if not rs.eof then
			sSector=rs(0)

			xsql="SELECT tipoperfil,TIPOPERFIL.descri "
			xsql =xsql & " FROM Objsector, ECO, ECOPERFIL, TIPOPERFIL "
			xsql=xsql & " where Eco.eco = Objsector.codsector and "
			xsql=xsql &  "ECO.eco = ECOPERFIL.codsector and "
			xsql=xsql & "ECOPERFIL.codperfil = TIPOPERFIL.tipoperfil "
			xsql=xsql & "and objsector.objid =" & rs(0)
			call ofv.cerrarconsulta(rs)

			sInput = ofv.GenerarCombo("Subconsulta","perfil","nsnc",Xsql,cn,0,1,"no")

		else

			sInput = "&nbsp;" & ofv.GenerarInput("perfil","nsnc","hidden","","1","dt","")

		end if

		sInput   = sinput & ofv.GenerarInput("lista",ofv.convertircar(SsectorT," ","_"),"hidden","","1","dt","")
		celda    = ofv.GenCelda("","tc","","","","","","<font>Perfil</font>","si")
		Celda    = celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
		Fila     = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

		formx    = Fila

		sTARGET  = ""
		sAccion    = "ListadoConsultas.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren & "&empre=" & sempre

ElseIf sNombre = "CONSULTA POR LINEA" Then

		If request.querystring("id") <> "" Then
			id = request.querystring("id")
		Else
			id = CLng("0" & request.form("id"))
		End If
		proc = CLng("0" & request.form("proc"))
		subp = CLng("0" & request.form("subp"))

		opc = request.querystring("opc")
		fil = 0
		If opc = "vuel" Then

			id = request.querystring("id")
			proc = request.querystring("proc")
			subp = request.querystring("subp")
			fil = request.querystring("sfiltro")

			If fil = "TODOS" Then fil = 0

		End If

		'--------------------
		'response.write request.querystring & "<BR>" & request.form

		'if request.querystring("id") = "" then

		Celda  = ofv.GenCelda("","tc","","","","","","<font>Linea de negocio</font>","si")

		scomKK = ""
		scomDD = ""
		'strsql = "select distinct a.* from tipoproducto a inner join objetos b on (a.objid = b.tsegmento) where b.tsegmento <> 0 order by a.descri"
		strsql = "select distinct a.objid,a.descri from tipoproducto a "
		strsql= strsql & " inner join objetos b on (a.objid = b.tsegmento) "
		strsql= strsql & " inner join objprty c on (b.codobjeto = c.codobjeto and b.tipobjeto=c.tipobjeto and c.estado=3)"
		strsql= strsql & " order by a.descri"


		Set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)

		Do Until rsniv.eof
			scomKK = scomKK & rsniv(0) & Chr(9)
			scomDD = scomDD & rsniv(1) & Chr(9)
			RSNIV.MOVENEXT
		Loop
		Call ofv.cerrarconsulta(rsniv)

		sInput = ofv.gencomboFX("Subconsulta1","id",id,scomKK,scomDD,"")


		Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
		Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

		sAccion = "subconsulta.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
		Formulario = Formulario & ofv.GenForm("Subconsulta1","",sAccion,"","post","",Fila,"si")

		'end if

		Fila    = ""

		'-----------------------------
		If request.querystring("id") <> "" Or id <> "" Then

			Celda = ofv.GenCelda("","tc","","","","","","<font>Proceso</font>","si")

			Strsql = "Select distinct c.objid,c.descri "
			Strsql = Strsql & " From objtsegagr a "
			Strsql = Strsql & " inner join objetos b on (a.codobjeto = b.codobjeto and a.tipobjeto = b.tipobjeto and b.tsegmento=" & id & ") "
			Strsql = Strsql & " inner join objprty d on (a.codobjeto = d.codobjeto and a.tipobjeto = d.tipobjeto and d.estado=3) "
			Strsql = Strsql & " inner join tipoagrupa c on (a.tagrupa = c.objid and c.estado = 3) "
			Strsql = Strsql & " order by c.descri "

			scomKK = "0"
			scomDD = "TODOS"
			DELIM=Chr(9)
			'response.write strsql
			Set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			Do Until rsniv.eof
				scomKK = scomKK & DELIM & rsniv(0) '& chr(9)
				scomDD = scomDD & DELIM & rsniv(1) '& chr(9)
				DELIM=Chr(9)
				RSNIV.MOVENEXT
			Loop
			Call ofv.cerrarconsulta(rsniv)

			sInput = ofv.gencomboFX("Subconsulta2","proc",proc,scomKK,scomDD,"")
			sInput = sinput & ofv.GenerarInput("id",id,"hidden","Subconsulta2","1","dt","")

			Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
			Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

			sAccion = "subconsulta.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0&id=" & id
			Formulario = Formulario & ofv.GenForm("Subconsulta2","",sAccion,"","post","",Fila,"si")

			Fila    = ""

			'-----------------------------

			Celda = ofv.GenCelda("","tc","","","","","","<font>Subproceso</font>","si")

			Strsql = "Select distinct a.objid,a.descri From TipoSegmento a "
			Strsql = Strsql & " inner join objtsegagr b on (a.objid=b.tsegmento and  b.tagrupa = " & proc & ") "
			Strsql = Strsql & " inner join objetos c on (b.codobjeto = c.codobjeto and b.tipobjeto=c.tipobjeto and c.tsegmento = " & id & ") "
			Strsql = Strsql & " where a.estado = 3"
			Strsql = Strsql & " order by a.descri"

			Strsql = "Select distinct b.objid,b.descri "
			Strsql = Strsql & " From objtsegagr a "
			Strsql = Strsql & " inner join TipoSegmento b on (a.tsegmento=b.objid and  b.estado = 3) "
			Strsql = Strsql & " inner join objetos c on (a.codobjeto = c.codobjeto and a.tipobjeto=c.tipobjeto and c.tsegmento = " & id & ") "
			Strsql = Strsql & " inner join objprty d on (a.codobjeto = d.codobjeto and a.tipobjeto=d.tipobjeto and d.estado = 3) "
			Strsql = Strsql & " where a.tagrupa= " & proc
			Strsql = Strsql & " order by b.objid,b.descri "

			scomKK = "0"
			scomDD = "TODOS"
			DELIM= Chr(9)
			' RESPONSE.WRITE subp
			'response.write strsql & "<br>" & chr(13)
			Set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			Do Until rsniv.eof
				scomKK = scomKK & DELIM & rsniv(0)
				scomDD = scomDD & DELIM & rsniv(1)
				RSNIV.MOVENEXT
			Loop
			Call ofv.cerrarconsulta(rsniv)

			sInput = ofv.gencomboFX("Subconsulta3","subp",subp,scomKK,scomDD,"")
			'sInput = ofv.GenerarCombo("Subconsulta3","PROm",xtPRO,Strsql,cn,0,1,"")& "producto"
			sInput = sinput & ofv.GenerarInput("id",id,"hidden","Subconsulta3","1","dt","")
			sInput = sinput & ofv.GenerarInput("proc",proc,"hidden","Subconsulta3","1","dt","")
			Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
			Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

			sAccion = "subconsulta.asp?Menu=" & sMenu & "&Nombre=" & sNombren & "&CantRegMostrar=10&CantRegMover=0"
			Formulario = Formulario & ofv.GenForm("Subconsulta3","",sAccion,"","post","",Fila,"si")

			Fila    = ""

			'-------------------------

			Celda = ofv.GenCelda("","tc","","","","","","<font>Tipo de esquema</font>","si")

			strsql = "select distinct a.tipobjeto,c.descri "
			strsql = strsql & " from objetos a "
			strsql = strsql & " inner join tipobjeto c on (a.tipobjeto = c.tipobjeto ) "
			strsql = strsql & " inner join objprty d on (a.codobjeto=d.codobjeto and a.tipobjeto=d.tipobjeto and d.estado=3 ) "

			If id <> 0 And proc = 0 And subp = 0 Then

				strsql = strsql & " where a.tsegmento = " & id

			ElseIf id <> 0 And proc <> 0 And subp = 0 Then

				strsql = strsql & " inner join objtsegagr b on (a.codobjeto=b.codobjeto and a.tipobjeto=b.tipobjeto and b.tagrupa = " & proc & " ) "
				strsql = strsql & " where a.tsegmento = " & id

			ElseIf id <> 0 And proc <> 0 And subp <> 0 Then

				strsql = strsql & " inner join objtsegagr b on (a.codobjeto=b.codobjeto and a.tipobjeto=b.tipobjeto and b.tsegmento = " & subp & " and b.tagrupa = " & proc & " ) "
				strsql = strsql & " where a.tsegmento = " & id

			End If

			Set rsniv = ofv.crearconsultaEx(StrSql,cn,1,parametros)

			scomKK = "0"
			scomDD = "TODOS"
			DELIM= Chr(9)
			Do Until rsniv.eof
				scomKK = scomKK & DELIM & rsniv(0)
				scomDD = scomDD & DELIM & rsniv(1)
				RSNIV.MOVENEXT
			Loop
			Call ofv.cerrarconsulta(rsniv)


			'sInput = ofv.GenerarCombox3(strSQL, cn, 0,1, "SFILTRO")
			sInput = ofv.gencomboFX("Subconsulta","SFILTRO",fil,scomKK,scomDD,"no")

			If id = 0 Then id = "nsnc"
			If proc = 0 Then proc = "nsnc"
			If subp = 0 Then subp = "nsnc"

			sInput = sinput & ofv.GenerarInput("ListaS",id,"hidden","Subconsulta","1","dt","")
			sInput = sinput & ofv.GenerarInput("ListaG",proc,"hidden","Subconsulta","1","dt","")
			sInput = sinput & ofv.GenerarInput("ListaP",subp,"hidden","Subconsulta","1","dt","")


			Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
			Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

			formx  =  Fila


			sTARGET  = ""
			sAccion    = "ListadoConsultas.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren
		End If

		'Listas de Estados
	else

		sNombre = "CONSULTA POR ESTADO"

		'' NUEVO Agregado para version 4.6
		'response.write request.querystring


		XFNAME   = request.querystring("XFNAME")
		XFSRH    = request.querystring("XFSRH")
		XFSRD    = request.querystring("XFSRD")

		IF XFNAME = "" THEN XFNAME = "PV"
		IF XFSRD  = "" THEN XFSRD = 0
		IF XFSRH  = "" THEN XFSRH = 0

		' PRIMERA VEZ
		IF XFNAME="PV"  THEN
			XFSRD = clng(DATE()) - clng(cdate("01-01-1900"))
			XFSRH = XFSRD '+ 15

		END IF
		'   RESPONSE.WRITE request.form
		IF XFNAME = "FD" THEN
			' RESPONSE.WRITE request.FORM
			XFSRD = request.form("DESDE")
			XFSRH = XFSRH

		END IF

		IF XFNAME = "FH" THEN
			XFSRD = XFSRD
			XFSRH =request.form("HASTA")

		END IF
		' response.write request.querystring
		sImagenE = "<B><i CLASS='las la-calendar' STYLE=font-size:14pt;COLOR:" & Session("CMFPNBOT") & "></i></B>"
		sAccion = "'javascript:void(0);'" & GTCSETCALR("frmCalendarioD","Desde",GTCSMT,GTCFCH,GTCFCHDAT,XFSRD,4,2,0,0,0)
		sinputE = "&nbsp;&nbsp;<a  class=UT href=" & sAccion
		sinputE = sinputE  & " TITLE='MODIFICAR FECHA' >" & sImagenE & "</a>"

		sAccionHasta = "'javascript:void(0);'" & GTCSETCALR("frmCalendarioH","Hasta",GTCSMT,GTCFCH,GTCFCHDAT,XFSRH,4,2,0,0,0)
		sinputHasta = "&nbsp;&nbsp;<a  class=UT href=" & sAccionHasta
		sinputHasta = sinputHasta  & " TITLE='MODIFICAR FECHA' >" & sImagenE & "</a>"


		' FECHA DESDE
		'' response.write "   " &  XFSRD & "iii"
		if  clng(XFSRD) = 0 then
			sInput2 = "------------"
		else
			sInput2 = cdate(clng(XFSRD) + clng(cdate("01-01-1900")) )
			swd = datepart("w",sInput2)
			sInput2 = Fechador(XFSRD)
		END IF
		'aca
		'Celda  =  ofv.GenCelda("","tc","","","","","","<font>Desde</font>","si")
		''sInput = ofv.GenerarInput("Desde2",sInput2,"agregar","","40","dt","") & ofv.GenerarInput("Desde",cstr(cdate(XFSRD + clng(cdate("01-01-1900")))),"hidden","","20","dt","")
		'sInput = ofv.GenerarInput("Desde2",sInput2,"agregar","","80","dt","") & ofv.GenerarInput("Desde3",XFSRD + clng(cdate("01-01-1900")) ,"hidden","","20","dt","")

		'Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput&sinputE,"si")

		'Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

		' FECHA HASTA
		if clng(XFSRH) = 0 then
			sInput2 = "------------"
		else
			sInput2 = cdate(clng(XFSRH) + clng(cdate("01-01-1900")) )
			swd = datepart("w",sInput2)
			sInput2 = Fechador(XFSRH)
		END IF

		'	Celda  = ofv.GenCelda("","tc","","","","","","<font>Hasta</font>","si")
		''sInput = ofv.GenerarInput("Hasta2",sInput2,"agregar","","40","dt","") & ofv.GenerarInput("Hasta",cdate(XFSRH+ clng(cdate("01-01-1900"))),"hidden","","20","dt","")
		'aca
		''  sInput = ofv.GenerarInput("Hasta2",sInput2,"agregar","","80","dt","") & ofv.GenerarInput("Hasta3",XFSRH + clng(cdate("01-01-1900")),"hidden","","20","dt","")
		'	Celda  = Celda & ofv.GenCelda("","vti","","","","","",sInput&sinputHasta,"si")

		' 	Fila       = Fila & ofv.GenRow("","","","","30","","",Celda,"si")

		'' NUEVO Agregado para version 4.6 FIN

		'SECTOR


		Strsql = "Select * From Estados order by estado"
		sInput = ofv.GenerarCombo2("lista","1",Strsql,cn,0,1,"sb")
		celda  = ofv.GenCelda("","tc","","","","","","<span>Estado</span>","si")
		Celda  = celda & ofv.GenCelda("","vti","","","","","",sInput,"si")
		Fila   = Fila & ofv.GenRow("","","","","30","","",Celda,"si")


		formx    = Fila

		sTARGET  = ""

		'' sAccion    = "ListadoConsultas.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren
		sAccion    = "ListadoConsultas.asp?Menu=PORESTADO" & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren

		'Fin preparacion
	end if




	'''' Armado del Form de despacho '''''''''''''''''''''''

	'   if formx <> "22" then
	'     Celda    = ofv.GenCelda("","tc","","","","","","&nbsp;","si")
	'     sInput   = ofv.GenerarInput("","Buscar","submit","","10","bt","")
	'     sInput   = sinput & ofv.GenerarInput("Tcons",sMenu,"hidden","","1","dt","")
	'     Celda    = Celda & ofv.GenCelda("","vti","Center","","","","",sInput,"si")
	'     formx    = formx  & ofv.GenRow("","","","","40","","",Celda,"si")
	'   sAccion    = "ListadoConsultas.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren
	'   Formulario = Formulario & ofv.GenForm("Subconsulta","",sAccion,"","post","",formx,"si")
	'   end if

	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''


	if formx <> "" then

		IF sAccion = "" THEN
			sAccion    = "ListadoConsultas.asp?Menu=" & sMenu & "&CantRegMostrar=10&CantRegMover=0&SORDEN=N" & "&Nombre=" & sNombren
		END IF


		Celda    = ofv.GenCelda("","ut","","","","","","&nbsp;","si")
		if session("APLMDCLI")="BC" THEN
			sInput   = ofv.GenerarInput("","&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Buscar&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;","submit","","10","bt"& msg,"")
		ELSE
			sInput   = ofv.GenerarInput("","Buscar","submit","","10","bt"& msg,"")
		END IF
		sInput   = sinput & ofv.GenerarInput("Tcons",sMenu,"hidden","","1","dt","")
		Celda    = ofv.GenCelda("","ut","Center colspan=2","","","","",sInput,"si")
		formx    = formx  & ofv.GenRow("","","","","40","","",Celda,"si")

		funxt=""
		if sNombre = "CONSULTA POR INSTRUMENTO" then
			funxt = " onsubmit=""return ValBus();"""
		end if

		%>

		<script type="text/javascript">

		function ValBus(){

			if(document.getElementById('Lista').value=="0"){

				alert('Seleccione una opcion');
				//alert(document.getElementsByTagName('input'));
				//alert(document.getElementsByClassName("dt"));
				//alert(document.getElementById('Lista').value);

				return false;
				}
			else{
				// alert('it Works!');
				return true;
				}
				}
				</script>

				<%



				' ****** CODIGO PARA CALENDARIO
	 			sInputDesde=ofv.GenerarInput("Desde",fechaconvert(XFSRD+ clng(cdate("01-01-1900"))),"hidden","","20","dt","")
				sAccionRetornoD="Subconsulta.asp?nombre="& sNombren &"&XFNAME=FD&Desde=" & XFSRD & "&Hasta=" & XFSRH & "&Fecha=" & fecha &"&sorden=" & sorden & "&CantRegMostrar=10&CantRegMover=0&XFSRD="& XFSRD & "&XFSRH=" & XFSRH
				formRetornoD= ofv.GenForm("frmCalendarioD","",sAccionRetornoD,STARGET,"post","",sInputDesde,"si")

	 			sInputHasta=ofv.GenerarInput("Hasta",fechaconvert(XFSRH+ clng(cdate("01-01-1900"))),"hidden","","20","dt","")
				sAccionRetornoH="Subconsulta.asp?nombre="& sNombren &"&XFNAME=FH&Desde=" & XFSRD & "&Hasta=" & XFSRH & "&Fecha=" & fecha & "&sorden=" & sorden & "&CantRegMostrar=10&CantRegMover=0&XFSRD="& XFSRD & "&XFSRH=" & XFSRH
				formRetornoH= ofv.GenForm("frmCalendarioH","",sAccionRetornoH,STARGET,"post","",sInputHasta,"si")
'RESPONSE.WRITE sMenu
				' ****** CODIGO PARA CALENDARIO
				SELECT CASE sMenu
				CASE "PORFECHA"
					 tabla= ofv.GenTabla("","aaa Xstyle='border-radius: 0px 0px 10px 10px;' ","","80% ","","0","","5","3","","",formx,"si")
			    CASE "PORESTADO"
			         tabla= ofv.GenTabla("","aaa Xstyle='border-radius: 0px 0px 10px 10px;' ","","80% ","","0","","5","3","","",formx,"si")
			    CASE "PORIDENTIFICACION"
			         tabla= ofv.GenTabla("","aaa Xstyle='border-radius: 0px 0px 10px 10px;' ","","80% ","","0","","5","3","","",formx,"si")
				CASE "PORINSTRUMENTO"
			         tabla= ofv.GenTabla("","aaa Xstyle='border-radius: 0px 0px 10px 10px;' ","","80% ","","0","","5","3","","",formx,"si")
				CASE "PORBUSCADOR"
			         tabla= ofv.GenTabla("","aaa Xstyle='border-radius: 0px 0px 10px 10px;' ","","80% ","","0","","5","3","","",formx,"si")

				CASE ELSE
				 tabla= ofv.GenTabla("","aaa style='border-radius: 0px 0px 10px 10px;' ","","80% ","","0","","5","3","","",formx,"si")
				END SELECT
				Formulario =  formulario & ofv.GenForm("Subconsulta","",sAccion,STARGET,"post" & funxt,"",TABLA,"si")  & formRetornoD & formRetornoH

				'onclick=""; document.getElementById('').submit();return false;""'
				'function validacionbuscar()
				'if(document.getElementById('form')="")
				'{
				'alert('El campo debe tener un valor distinto de nulo');
				'return false;
				'}
				'return true;



			end if



			'''''''''''''''''''''''''''''''''''''''''

			'sHTML = sHTML & "<center><BR>"

			if sNombre = "CONSULTA POR TEXTO" then

				sHTML = sHTML & ofv.GenTabla("","aaa style='border-radius: 10px 10px 0px 0px;border-bottom:none;'","","80%","","0","","5","3","","",Formulario,"si")

			elseif sNombre = "CONSULTA POR ESTADO" then
				'RESPONSE.WRITE REQUEST.FORM
				sHTML=SHTML & FORMULARIO
			elseif sNombre = "CONSULTA POR FECHA" then
			 sHTML = sHTML & ofv.GenTabla("","","","","","","","","","","",Formulario,"si")
			elseif sNombre = "CONSULTA POR INSTRUMENTO" then
			 sHTML = sHTML & ofv.GenTabla("","","","","","","","","","","",Formulario,"si")
            elseif sNombre = "CONSULTA POR IDENTIFICACION" then
			 sHTML = sHTML & ofv.GenTabla("","","","","","","","","","","",Formulario,"si")

			ELSE
                IF SMENU="PORBUSCADOR" THEN
                 	 sHTML = sHTML & ofv.GenTabla("","aaa Xstyle='border-radius: 10px 10px 0px 0px;border-bottom:none;'","","80% ","","0","","5","3","","",Formulario,"si")

                ELSE
				 sHTML = sHTML & ofv.GenTabla("","aaa style='border-radius: 10px 10px 0px 0px;border-bottom:none;'","","80% ","","0","","5","3","","",Formulario,"si")
                END IF
			end if

			'sHTML = sHTML & "</center>"

			sHTML = sHTML & HTMgenlista

			'Botonera
			Agregar = ""
			Buscar  = ""
			Celda   = ""

			if session("SMPLMENU") = "S" then
				sAccion = "novedmenu.asp"
				Celda   = ofv.BotonVolver(sAccion)
			end if

			sHTML = sHTML & ofv.botoneraESSX("0",rs,cantregmoverx,cantregmostrarx,"no","no",posi,"","","no",Agregar,Buscar,celda)


			sHTML = sHTML & "<script language=JavaScript>"
			sHTML = sHTML & " function Init() "
			'sHTML = sHTML & "   { ff(); setTimeout (" & chr(34) & "Init()" & chr(34) & ", 350); }"
			sHTML = sHTML & "   { setTimeout (" & chr(34) & "Init()" & chr(34) & ", 350); }"
			sHTML = sHTML & " function ff() "
			sHTML = sHTML & "  {   "
			sHTML = sHTML & "  if (document.all.fnt01.color=='#f7aa22') { document.all.fnt01.color='#aa0000'; } "
			sHTML = sHTML & "  else { if (document.all.fnt01.color=='#aa0000') { document.all.fnt01.color='#f75500'; } "
			sHTML = sHTML & "  else {if (document.all.fnt01.color=='#f75500') { document.all.fnt01.color='#f7aa22'; } } } "
			sHTML = sHTML & "   } "

			if sNombre = "CONSULTA POR TEXTO" then

				sHTML = sHTML & " Init();"

			end if



			sHTML = sHTML & "</script>"& GTCCALENDAR("")


			sHTML = sHTML & "</body></html>"

			call ofv.cerrarconn(cn)
end if

''RESPONSE.write sNombre & " snombre" & chr(13)

response.write sHTML

function GenListaDoc(HTgrp,HTcod,HTtip,HTnam)

htdata = ""




'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' SELECCION ACTUAL '''''''''''''''''''''''''''''''''''''''''''''''


       if HTgrp = "M" then
          descrit = "TODOS LOS DOCUMENTOS"
       else
          descrit = "SIN DESCRIPCION CONOCIDA"
          if HTgrp = "T" then

              Strsql2 = " select descri from tema where codtema = 0" & HTcod
              set rsr = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
              if not rsr.eof then
                 descrit=rsr(0)
              end if
              call ofv.cerrarconsulta(rsr)


          elseif HTgrp = "S" then

              Strsql2 = " select descri from PRODSERV where codPRODSERV = 0" & HTcod
              set rsr = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
              if not rsr.eof then
                 descrit=rsr(0)
              end if
              call ofv.cerrarconsulta(rsr)

          else

              Strsql2 = " select descri from OBJETOS where codOBJETO = 0" & HTcod
              set rsr = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
              if not rsr.eof then
                 descrit=rsr(0)
              end if
              call ofv.cerrarconsulta(rsr)

          end if

       end if

      sInput1  = ofv.GenerarInput("lista",sSectort,"hidden","","1","dt","")
      sInput2  = ofv.GenerarInput("tsel",tsel,"hidden","","1","dt","")
      sInput3  = ofv.GenerarInput("ptextot",textot,"hidden","","1","dt","")
      sInput4  = ofv.GenerarInput("ptgrp",HTgrp,"hidden","","1","dt","")
      sInput5  = ofv.GenerarInput("ptcod",HTcod,"hidden","","1","dt","")
      sInput6  = ofv.GenerarInput("pttip",HTtip,"hidden","","1","dt","")
      sInput7  = ofv.GenerarInput("ptnam",HTnam,"hidden","","1","dt","")

      sInputb  = ofv.GenerarInput("bot","OK","button","","1","bt onclick='clselecc();' ","")

      HTdatos = sInput1 & sINPUT2 & sINPUT3 & sINPUT4 & sINPUT5 & sINPUT6  & sINPUT7 & sInputb


      sImagen2 = "<B><FONT FACE=CorFlatIcon COLOR=#00aa00 SIZE=3>&#097;</FONT></B>"
      sImagen  = "<B><FONT FACE=CorFlatIcon COLOR=#f7aa22 SIZE=2>4</FONT></B>"
      sImagen3 = "<B><FONT FACE=CorFlatIcon COLOR=#99ccee SIZE=1>&#060;</FONT></B>"

      celdasl = ""
      Celdasl = celdasl & ofv.GenCelda("","uts style='border-bottom-style:solid;border-bottom-width:1;border-bottom-color:#dddddd;' ","","1%","","","",sImagen2,"si")
      Celdasl = celdasl & ofv.GenCelda("","ttt","","","","","","&nbsp;" & descrit,"si")
      Celdasl = celdasl & ofv.GenCelda("","tc","","2%","","","",HTdatos,"si")
      Filasl  = ofv.GenRow("","","","","","","",Celdasl,"si")

     sAccion = "Subconsulta.asp?frm=4&Menu=" & sMenu & "&Nombre=" & request.querystring("Nombre")

  ''response.write smenu & " smenu " & chr(13)

     Filasl  = ofv.GenForm("Subconsulta4","",sAccion,"","post","",Filasl,"si")


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

   STYLOTE = " style='border-style:solid;border-width:1;border-color:#336699;background-color:#ffffff;' "
   STYLOT  = ""

   dim ATABLVL(20,2)
   FilaslT  = ""

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' LISTA '''''''''''''''''''''''''''''''''''''''''''''''


'' 1ERA '''''''''''''''''''''''''''''''''''''''''''''''

       descrit = "TODOS LOS DOCUMENTOS"

       HTLEVEL = 1

       if HTgrp = "M" then

          STYLOT  = STYLOTE
       end if

      sINPUTT = "<A class=nt HREF='javascript:void(0);' ONCLICK=""clactfrm('M','99999','TODOS','TODOS');document.Subconsulta4.submit();return false;"" >  " & descrit & "</A>"

      celdasl = ofv.GenCelda("","UT","","2%","","","",sImagen,"si")
      Celdasl = celdasl & ofv.GenCelda("","UT COLSPAN=15 " & STYLOT & " ","","","","","",sINPUTT,"si")
      FilaslT  = FilaslT & ofv.GenRow("","","","","","","",Celdasl,"si")


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'' 2DA '''''''''''''''''''''''''''''''''''''''''''''''


   STYLOT  = ""

       if HTgrp = "T" OR HTgrp = "S" OR HTgrp = "D" then

          HTLEVEL = 2

          descrit = "TEMAS"


          descrit = "SIN DESCRIPCION CONOCIDA"
          if HTgrp = "T" then

              Hxcod = HTcod

              Strsql2 = " select descri from tema where codtema = 0" & HTcod
              set rsr = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
              if not rsr.eof then
                 descrit=rsr(0)
              end if
              call ofv.cerrarconsulta(rsr)

              STYLOT  = STYLOTE


          elseif HTgrp = "S" then

              Strsql2 = " select codtema from PRODSERV where codPRODSERV = 0" & HTcod
              set rsr = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
              if not rsr.eof then

                 Hxcod = rsr(0)

                 Strsql2 = " select descri from tema where codtema = 0" & rsr(0)
                 set rsr = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
                 if not rsr.eof then
                    descrit=rsr(0)
                 end if

              end if
              call ofv.cerrarconsulta(rsr)

          else



              Strsql2 = " select tema from CsRelPath where cobj = 0" & HTcod & " and tobj = '" & HTtip & "' "
              set rsr = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
              if not rsr.eof then

                 Hxcod = rsr(0)


                 Strsql2 = " select descri from tema where codtema = 0" & rsr(0)
                 set rsr = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
                 if not rsr.eof then
                    descrit=rsr(0)
                 end if

              end if
              call ofv.cerrarconsulta(rsr)






          end if

        ATABLVL(2,1) = Hxcod
        ATABLVL(2,2) = "TEMA"

         sINPUTT = "<A class=nt HREF='javascript:void(0);' ONCLICK=""clactfrm('T','" & Hxcod & "','TEMA','" & replace(descrit," ","_") & "');document.Subconsulta4.submit();return false;"" >  " & descrit & "</A>"

         celdasl = ofv.GenCelda("","UT","","1%","","","","&nbsp;","si")
         celdasl = celdasl & ofv.GenCelda("","UT","","2%","","","",sImagen,"si")
         Celdasl = celdasl & ofv.GenCelda("","UT COLSPAN=14 " & STYLOT & " ","","","","","",sINPUTT,"si")
         FilaslT  = FilaslT & ofv.GenRow("","","","","","","",Celdasl,"si")

       end if


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'' 3erA '''''''''''''''''''''''''''''''''''''''''''''''




   STYLOT  = ""



       if HTgrp = "S" OR HTgrp = "D" then

          descrit = "SUBTEMAS"

          HTLEVEL = 3


          descrit = "SIN DESCRIPCION CONOCIDA"

         if HTgrp = "S" then

              Hxcod = HTcod
              Hxtip = HTtip

              Strsql2 = " select descri from PRODSERV where codPRODSERV = 0" & HTcod
              set rsr = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
              if not rsr.eof then
                 descrit=rsr(0)
              end if
              call ofv.cerrarconsulta(rsr)

              STYLOT  = STYLOTE


          else



              Strsql2 = " select stema,stematipo from CsRelPath where cobj = 0" & HTcod & " and tobj = '" & HTtip & "' "
              set rsr = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
              if not rsr.eof then

                 Hxcod = rsr(0)
                 Hxtip = rsr(1)


                 Strsql2 = " select descri from PRODSERV where codPRODSERV = 0" & rsr(0)
                 set rsr = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
                 if not rsr.eof then
                    descrit=rsr(0)
                 end if

              end if
              call ofv.cerrarconsulta(rsr)






          end if

        ATABLVL(3,1) = Hxcod
        ATABLVL(3,2) = Hxtip


         sINPUTT = "<A class=nt HREF='javascript:void(0);' ONCLICK=""clactfrm('S','" & Hxcod & "','" & Hxtip & "','" & replace(descrit," ","_") & "');document.Subconsulta4.submit();return false;"" >  " & descrit & "</A>"

         celdasl = ofv.GenCelda("","UT","","1%","","","","&nbsp;","si")
         celdasl = celdasl & ofv.GenCelda("","UT","","1%","","","","&nbsp;","si")
         celdasl = celdasl & ofv.GenCelda("","UT","","2%","","","",sImagen,"si")
         Celdasl = celdasl & ofv.GenCelda("","UT COLSPAN=13 " & STYLOT & " ","","","","","",sINPUTT,"si")
         FilaslT  = FilaslT & ofv.GenRow("","","","","","","",Celdasl,"si")

       end if


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'' 4rtA '''''''''''''''''''''''''''''''''''''''''''''''

   STYLOT  = ""

       descrit = "DOCUMENTOS"

       if HTgrp = "D" then


          descrit = "SIN DESCRIPCION CONOCIDA"


              Strsql2 = " select * from CsRelPath where cobj = 0" & HTcod & " and tobj = '" & HTtip & "' "
              set rsr = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
              if not rsr.eof then

                 iii2 = 0
                 iii4 = 0
                 for iii = 1 to 13 step 2

                     httemp = rsr(3 + iii) & ""

                     if  httemp  <> "" then
                         iii2 = iii2 + 1
                         if iii2 > 1 then
                            Hxcod = rsr(3 + iii) & ""
                            Hxtip = rsr(4 + iii) & ""

                            HTLEVEL = HTLEVEL + 1

                            ATABLVL(HTLEVEL,1) = Hxcod
                            ATABLVL(HTLEVEL,2) = Hxtip

                            Strsql2 = " select descri from OBJETOS where codOBJETO = 0" & Hxcod
                            set rsr2 = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
                            if not rsr2.eof then
                               descrit=rsr2(0)
                            end if
                            call ofv.cerrarconsulta(rsr2)



         sINPUTT = "<A class=nt HREF='javascript:void(0);' ONCLICK=""clactfrm('D','" & Hxcod & "','" & Hxtip & "','" & replace(descrit," ","_") & "');document.Subconsulta4.submit();return false;"" >  " & descrit & "</A>"

         celdasl = ""
         for iii3=1 to (HTLEVEL-1)
             celdasl = celdasl & ofv.GenCelda("","UT","","1%","","","","&nbsp;","si")
         next
         celdasl = celdasl & ofv.GenCelda("","UT","","2%","","","",sImagen,"si")
         Celdasl = celdasl & ofv.GenCelda("","UT COLSPAN=" & (16 - HTLEVEL) & " " & STYLOT & " ","","","","","",sINPUTT,"si")
         FilaslT  = FilaslT & ofv.GenRow("","","","","","","",Celdasl,"si")


                         end if
                     end if


                 next



              end if
              call ofv.cerrarconsulta(rsr)



        Hxcod = HTcod
        Hxtip = HTtip

        HTLEVEL = HTLEVEL + 1

        ATABLVL(HTLEVEL,1) = Hxcod
        ATABLVL(HTLEVEL,2) = Hxtip

        Strsql2 = " select descri from OBJETOS where codOBJETO = 0" & HTcod
        set rsr2 = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
        if not rsr2.eof then
           descrit=rsr2(0)
        end if
        call ofv.cerrarconsulta(rsr2)

         STYLOT  = STYLOTE


         sINPUTT = "<A class=nt HREF='javascript:void(0);' ONCLICK=""clactfrm('D','" & Hxcod & "','" & Hxtip & "','" & replace(descrit," ","_") & "');document.Subconsulta4.submit();return false;"" >  " & descrit & "</A>"

         celdasl = ""
         for iii3=1 to (HTLEVEL-1)
             celdasl = celdasl & ofv.GenCelda("","UT","","1%","","","","&nbsp;","si")
         next
         celdasl = celdasl & ofv.GenCelda("","UT","","2%","","","",sImagen,"si")
         Celdasl = celdasl & ofv.GenCelda("","UT COLSPAN=" & (16 - HTLEVEL) & " " & STYLOT & " ","","","","","",sINPUTT,"si")
         FilaslT  = FilaslT & ofv.GenRow("","","","","","","",Celdasl,"si")







       end if


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'' linea de datos '''''''''''''''''''''''''''''''''''''''''''''''

   STYLOT  = ""
       descrit = "DATOS"

       if HTLEVEL = 1 then



          Strsql2 = " select codtema,descri from tema ORDER BY INDA, DESCRI "
          set rsr = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
          DO UNTIL rsr.eof
             Hxcod   = rsr(0)
             HxTIP   = "TEMA"
             descrit = rsr(1)


         sINPUTT = "<A class=nt HREF='javascript:void(0);' ONCLICK=""clactfrm('T','" & Hxcod & "','" & Hxtip & "','" & replace(descrit," ","_") & "');document.Subconsulta4.submit();return false;"" >  " & descrit & "</A>"

         celdasl = ""
         for iii3=1 to (HTLEVEL)
             celdasl = celdasl & ofv.GenCelda("","UT","","1%","","","","&nbsp;","si")
         next
         celdasl = celdasl & ofv.GenCelda("","UT","","2%","","","",sImagen3,"si")
         Celdasl = celdasl & ofv.GenCelda("","UT COLSPAN=" & (15 - HTLEVEL) & " " & STYLOT & " ","","","","","",sINPUTT,"si")
         FilaslT  = FilaslT & ofv.GenRow("","","","","","","",Celdasl,"si")

          RSR.MOVENEXT
          LOOP
          call ofv.cerrarconsulta(rsr)



       ELSEif HTLEVEL = 2 then

          Strsql2 = " select codPRODSERV,MARCAPS,DESCRI from PRODSERV where codtema = 0" & HTcod
          set rsr = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
          DO UNTIL rsr.eof
             Hxcod   = rsr(0)
             HxTIP   = rsr(1)
             descrit = rsr(2)


         sINPUTT = "<A class=nt HREF='javascript:void(0);' ONCLICK=""clactfrm('S','" & Hxcod & "','" & Hxtip & "','" & replace(descrit," ","_") & "');document.Subconsulta4.submit();return false;"" >  " & descrit & "</A>"

         celdasl = ""
         for iii3=1 to (HTLEVEL)
             celdasl = celdasl & ofv.GenCelda("","UT","","1%","","","","&nbsp;","si")
         next
         celdasl = celdasl & ofv.GenCelda("","UT","","2%","","","",sImagen3,"si")
         Celdasl = celdasl & ofv.GenCelda("","UT COLSPAN=" & (15 - HTLEVEL) & " " & STYLOT & " ","","","","","",sINPUTT,"si")
         FilaslT  = FilaslT & ofv.GenRow("","","","","","","",Celdasl,"si")

          RSR.MOVENEXT
          LOOP
          call ofv.cerrarconsulta(rsr)


     '''  ELSEif HTLEVEL = 3 then

       ELSE

         '' response.write "aaa" & HTLEVEL

          HTARMAWHERE = ""

          HTLEVELW = HTLEVEL

          IF HTLEVELW > 10 THEN HTLEVELW = 10

          III4 = 8

          FOR III3=3 TO HTLEVELW

              III2 = HTLEVELW - III3 + 3
              III4 = III4 - 1

              HTARMAWHERE = HTARMAWHERE & " AND MC" & III4 & " = " & ATABLVL(III2,1)
              HTARMAWHERE = HTARMAWHERE & " AND MT" & III4 & " = '" & ATABLVL(III2,2) & "' "


          NEXT



          descrit = "SIN DESCRIPCION CONOCIDA"


              Strsql2 = " select COBJ,TOBJ from CsRelPath where cobj > 0 " & HTARMAWHERE
              set rsr = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
              DO UNTIL rsr.eof

                            Hxcod = rsr(0) & ""
                            Hxtip = rsr(1) & ""

                            Strsql2 = " select descri from OBJETOS where codOBJETO = 0" & Hxcod
                            set rsr2 = ofv.crearconsultaEx(Strsql2,cn,1,parametros)
                            if not rsr2.eof then
                               descrit=rsr2(0)
                            end if
                            call ofv.cerrarconsulta(rsr2)


         sINPUTT = "<A class=nt HREF='javascript:void(0);' ONCLICK=""clactfrm('D','" & Hxcod & "','" & Hxtip & "','" & replace(descrit," ","_") & "');document.Subconsulta4.submit();return false;"" >  " & descrit & "</A>"

         celdasl = ""
         for iii3=1 to (HTLEVEL)
             celdasl = celdasl & ofv.GenCelda("","UT","","1%","","","","&nbsp;","si")
         next
         celdasl = celdasl & ofv.GenCelda("","UT","","2%","","","",sImagen3,"si")
         Celdasl = celdasl & ofv.GenCelda("","UT COLSPAN=" & (15 - HTLEVEL) & " " & STYLOT & " ","","","","","",sINPUTT,"si")
         FilaslT  = FilaslT & ofv.GenRow("","","","","","","",Celdasl,"si")





              RSR.MOVENEXT
              LOOP
              call ofv.cerrarconsulta(rsr)











       end if


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''












''''''''''''''''''''''
 htdata = htdata &  "<DIV class='DINDM' ID='HTX1'  NAME='HTX1'  style='display:none;position:absolute;top:140;left:220;HEIGHT=298px;WIDTH=558px;' >"

 htdata = htdata &  "<DIV class='DINDC' ID='HTX1C' NAME='HTX1C' style='position:relative;top:+2px;left:+2;HEIGHT=290px;WIDTH=550px;' >"
 htdata = htdata &  "<CENTER>"

 htdata = htdata &  "<table id=""MTINDX"" class='TWIND' width=100% cellspacing=0 cellpadding=2 >"
 htdata = htdata &  "<tr>"
 htdata = htdata &  "<td CLASS=mbt width=99%  >&nbsp;SELECCION DE CARPETAS</td>"
 htdata = htdata &  "<td CLASS=mbt width=1% onclick=""clselecc();"" title='Ocultar seleccion.' >"
 htdata = htdata &  "<b><font style=cursor:pointer;  onmouseover=""this.color='#bb0000';"" onmouseout=""this.color='" & session("CMFMNBOT") & "';"" face=CorFlatIcon size=1   >&#0114;</font></b>"
 htdata = htdata &  "</td>"
 htdata = htdata &  "</tr>"
 htdata = htdata &  "</table>"

 htdata = htdata &  ofv.GenTabla("TSEL1","aaa style='display:block;' ","","95%","","0","0","2","0","","",FilaSL,"si")

 htdata = htdata &  "<DIV class='DINDC' ID='HTX1d' NAME='HTX1d' style='position:relative;HEIGHT=220px;WIDTH=546px;overflow:auto;' >"
 htdata = htdata &  "<CENTER>"
 htdata = htdata &  ofv.GenTabla("TSEL2","aaa style='display:block;' ","","95%","","0","0","2","0","","",FilaSLT,"si")
 htdata = htdata &  "</CENTER>"
 htdata = htdata &  "</div>"

 htdata = htdata &  "</CENTER>"
 htdata = htdata &  "</div>"
 htdata = htdata &  "</div>"

 htdata = htdata & "<script language=javascript >"
 htdata = htdata & " function clselecc() { " & chr(13)
 htdata = htdata & "  if  (document.all.HTX1.style.display=='none')   " & chr(13)
 htdata = htdata & "      { document.all.HTX1.style.display='block';} " & chr(13)
 htdata = htdata & "  else                                             " & chr(13)
 htdata = htdata & "      { document.all.HTX1.style.display='none';}  " & chr(13)
 htdata = htdata & " } " & chr(13)
 htdata = htdata & " function clactfrm(fngrp,fncod,fntip,fnnam) { " & chr(13)
 htdata = htdata & "    document.all.ptgrp.value=fngrp; " & chr(13)
 htdata = htdata & "    document.all.ptcod.value=fncod; " & chr(13)
 htdata = htdata & "    document.all.pttip.value=fntip; " & chr(13)
 htdata = htdata & "    document.all.ptnam.value=fnnam; " & chr(13)
 htdata = htdata & "    document.all.ptextot.value=document.all.textot.value; " & chr(13)
' htdata = htdata & "    document.Subconsulta4.submit(); " & chr(13)
' htdata = htdata & "    return false; " & chr(13)
 htdata = htdata & " } " & chr(13)



 if  frmsen = "S" then
     htdata = htdata & " clselecc(); " & chr(13)

 end if


 htdata = htdata & "</script>"

' Whtml = Whtml & "<script language=javascript >"
' Whtml = Whtml & "document.all.mt1.style.display='none';"
' Whtml = Whtml & "document.all.mt2.style.display='block';"
' Whtml = Whtml & " DCEXRCLC();"
' Whtml = Whtml & "</script>"

''''''''''''''''''''''

   GenListaDoc = htdata

end function





Function Fechador(wkfecha)


      fec = cdate(wkfecha + clng(cdate("01-01-1900")))
      ndia = ofv.cdia(Weekday(fec))
      nMes = ofv.cmes(Month(fec))

      Fechador = "  " & ndia & ", " & Day(fec) & " de " & nMes & " de " & Year(fec)


End Function
Function fechaConvert(wkfecha)
  Nfecha=cdate(wkfecha)
  Nfecha=dateserial(year(wkfecha),month(wkfecha),day(wkfecha))
  fechaConvert=Nfecha
end function

Function fechatostr(wkfecha)
  fechatostr=cdate(wkfecha)
  fechatostr=day(wkfecha) & "/" & month(wkfecha) & "/" & year(wkfecha)

end function





%>



<SCRIPT LANGUAGE="JavaScript">


function validarfechas(){



	alert("La fecha Desde debe ser menor a la fecha Hasta .");
	return false;

}



//-->
</SCRIPT>