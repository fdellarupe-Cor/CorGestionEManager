<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else
	response.expires = 0
	server.scripttimeout = 600
	response.buffer = True

topeDef = 0
wIndDef = 0


	Dim carpetaVirtual
    carpetaVirtual = "ModeloIconos\imagenesV5"
    Dim carpetaFisica
  '  carpetaFisica = Server.MapPath("/")
  '  response.write carpetaFisica

    ruta=server.mappath(request.servervariables("script_name"))

    pos=InStrRev(ruta,"\")
	carpetaFisica=Mid(ruta,1,pos) & carpetaVirtual
	'response.write  carpetaFisica  & carpetaVirtual

    ' 2. Especificar el filtro con múltiples extensiones, separadas por coma
    Dim misFiltros
    misFiltros = "jpg, png, gif, jpeg" ' <-- Pasamos todas las extensiones en un string

    ' 3. Llamar a nuestra función con la ruta y los filtros
    Dim misArchivos
    rutaw="ModeloIconos/imagenesV5/"
  '  aDefCnt = ListarArchivos(carpetaFisica, misFiltros,rutaw)

'''''''''''''''''''' carga imagenes default


if IsArray(aDefCnt) then  wtopeDef = ubound(aDefCnt)-1

wIndDef  = 0

	''''''''''''''
	filtro=""
	'RESPONSE.WRITE request.form
if request.querystring("salidaD")="S" then ' solo si se usa filtro
 For Each fieldName In Request.Form

     Dim fieldValues
   '
     Set fieldValues = Request.Form(fieldName)
     fieldName=UCASE(fieldName)
     'multiples con el mismo nombre
     If fieldValues.Count > 1 Then
         '
         filtro="TIPOBJETO.tipobjeto"  & " in ("
         delim=""
         For i = 1 To fieldValues.Count
             filtro=filtro &  delim & "'" & fieldValues(i) & "'"
             delim=","
         Next

         filtro=filtro & ")"
         filtro= " AND " & filtro

     Else
         ' If there's only one value, access it directly
        SELECT CASE fieldName
        CASE "TEXTO"
          sbuscarNew= fieldValues
        CASE "ORDEN"
           IF fieldValues="NO" THEN
            ORDENDOC="order by objetos.indobj, OBJETOS.tipobjeto, OBJETOS.descri, Estados.estado "
           ELSE
           ORDENDOC= "order by " &  fieldValues & ",OBJETOS.tipobjeto, OBJETOS.descri"
           END IF
        CASE else
         filtro=" and tipobjeto.tipobjeto='" & fieldValues  & "'"

        END SELECT
     End If


	Next
ELSE
  ORDENDOC= "order by objetos.indobj, OBJETOS.tipobjeto, OBJETOS.descri, Estados.estado "
end if
'response.write filtro
'filtro=""
	XXXX = NEWEMPRE("NO")

	volver = OFV.vueltaasp2("../")
	volver2 = OFV.vueltaasp2("")
	CantregMostrar=14
	sNombre = "Tema - Subtema"

	sID = Request.querystring("ID")
	TIPOVISTA    = Request.querystring("TIPOVISTA")
	sitem = OFV.convertircar(Request.querystring("sitem"), "_", " ")
	stitem = OFV.convertircar(Request.querystring("stitem"), "_", " ")
	Stnivel = Request.querystring("stnivel")
	CantregMover= Request.querystring("CantregMover")

	cantregmostar=14
	Snivel = Request.querystring("NIVEL")
	sformato=request.querystring("FORMATO")
	acc=request.querystring("acc")

	sbuscar=request.form("buscar")
	xbt=Request.querystring("xbt")
	'  response.write "xbt=" & xbt
	If xbt = "S" Then session("sbuscar")=""
	If sbuscar <> "" Then
		session("sbuscar")=sbuscar
	Else
		sbuscar=session("sbuscar")
	End If

	'response.write  sbuscar
	If CantregMover="" Then CantregMover=0

	Session(Stnivel & "CantregMover")   = CantregMover

	CantregMover=Session(Stnivel & "CantregMover")
	'  response.write CantregMover & "  nivel"  & Stnivel

	If Request.querystring("stesq") = "" Then
		stesq = "TEMA"
	Else
		stesq = Request.querystring("stesq")
	End If
	If acc="UPD" Then
		RESPONSE.WRITE  "<script>"
		RESPONSE.WRITE "var frmcor=window.top.getIFD('info').getElementById('Coridformupd').value;"
		RESPONSE.WRITE "var Corididinput=window.top.getIFD('info').getElementById('Corididinput').value;"
		RESPONSE.WRITE "var CorContenedor=window.top.getIFD('info').getElementById('CorContenedor').value;"
		RESPONSE.WRITE  "window.top.getIFD('info').getElementById(CorContenedor).style.display='none';window.top.getIFD('info').getElementById(Corididinput).value='" & SNIVEL & "-" & sID & "'; window.top.getIFD('info').getElementById(frmcor).submit();</script>"
	End If

	volver2 = volver2 & "?ID=" & sID
	volver2 = volver2 & "&sitem=" & OFV.convertircar(sitem, " ", "_")
	volver2 = volver2 & "&stitem=" & OFV.convertircar(stitem, " ", "_")
	volver2 = volver2 & "&stnivel=" & Stnivel & "&stesq=" & stesq
	volver2 = OFV.convertircar(volver2, "&", "*")

	ReDim Preserve bConsultar(4)
	OFV.ObtenerAtributosNOdo Session("Form"), "", "ESTADO", "4"
	bConsultar(4) = OFV.AConsultar
	OFV.ObtenerAtributosNOdo Session("Form"), "", "ESTADO", "2"
	bConsultar(2) = OFV.AConsultar
	OFV.ObtenerAtributosNOdo Session("Form"), "", "ESTADO", "1"
	bConsultar(1) = OFV.AConsultar
	OFV.ObtenerAtributosNOdo Session("Form"), "", "ESTADO", "3"
	bConsultar(3) = OFV.AConsultar

	sVerEstado = "  and Objetos.objnivel >= " & Session("SegNivel") & " "



	' response.write session("FormN")

	If TIPOVISTA="CONS" Then
		sverestado = sverestado & " and Objprty.Estado In (3,1)"
	Else
		sverestado = sverestado & " and Objprty.Estado In (9,"
		' sverestado = sverestado & " and Objprty.Estado In (3)"
		For i = 1 To 4
			If bConsultar(i) Then
				sVerEstado = sVerEstado & i & ","
				If i = 1 Then sVerEstado = sVerEstado & "6,"
				If i = 2 Then sVerEstado = sVerEstado & "5,"
			End If
		Next
		sVerEstado = Left(sVerEstado, Len(sVerEstado) - 1) & ") "
	End If
	'sHTML      = ofv.FormHeader(session("FormN"))

	If sformato="ABR" Then
		WSPDIVN = "DTTEMAINT"       ' ID DEL DIV DE VENTANA PARA MOSTRAR
		sHTML = ""
		sHTML = sHTML & "<HTML>"
		sHTML = sHTML & "<BODY onload=""window.top.LDVENTANA('" & WSPDIVN & "');"" >"
	Else

		sHTML      = ofv.FormHeader(session("FormN"))
	End If


	'     sActcon  = "Consultar.asp?Codigo=" & rs(0) & "&Tipo="
	'     sActcon  = sActcon & stitem & "&Perfil=" & sPerfil
	'     sActcon  = sActcon & "&Datos=" & "&Volver=" & volver2


	xs = "<script language=JavaScript>" & Chr(13)
	xs = xs & "function AbrirCons(Datos) {" & Chr(13)
	xs = xs & " var h=570;w=790 ;"
	If  session("APLMDCLI")="BaC" Then
		xs = xs &  "if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;} "
	End If

	xs = xs & " myWin= open(Datos"
	xs = xs & "," & Chr(34) & "Cons" & Chr(34) & "," '& Chr(34)
	'xs = xs & "menubar=no,toolbar=no,status=yes,height=570,width=790,resizable=yes" & Chr(34)
	xs = xs &"  'menubar=no,toolbar=no,width='+w+',height='+h+',top=0,left=0,resizable=yes,status=yes'" '& CHR(34
	xs = xs & " ); " & Chr(13) & " } " & Chr(13)

	xs = xs & "function Axn(c1,c2,c3) {" & Chr(13)

	xs = xs & " var mydat= 'Consultar.asp?Codigo=' + c1 + '&Tipo=' + c2 + '&Perfil=' + c3 + '&Datos=&Volver=" & volver2 & "'; "
	'xs = xs & " var h=570;w=790 ;"
	'If  session("APLMDCLI")="BaC" Then
		xs = xs &  "if (window.screen) {w = window.screen.availWidth;h = window.screen.availHeight-60;} "
	'End If
	xs = xs & " myWin= open(mydat"
	xs = xs & "," & Chr(34) & "Cons" & Chr(34) & "," '& Chr(34)
	'xs = xs & "menubar=no,toolbar=no,status=yes,height=570,width=790,resizable=yes" & Chr(34)
	xs = xs &"  'menubar=no,toolbar=no,width='+w+',height='+h+',top=0,left=0,resizable=yes,status=yes'" '& CHR(34
	xs = xs & " ); " & Chr(13) & " } " & Chr(13)


	xs = xs & "</script>" & Chr(13)

	sHTML = sHTML & xs

	sImagenM   = "<B><FONT FACE=""CorFlatIcon"" COLOR=#005500 SIZE=3>2</FONT></B>"
	sImagenI2  = "<B><FONT FACE=""CorFlatIcon"" COLOR=#f7aa22 SIZE=3>&#0128;</FONT></B>"
	sImagenI   = "<B><FONT FACE=""CorFlatIcon"" COLOR=#f7aa22 SIZE=2>u</FONT></B>"
	sImagenI3  = "<B><FONT FACE=""CorFlatIcon"" COLOR=#6699CC SIZE=2>=</FONT></B>"
	sImagenEM  = "<B><FONT FACE=CorFlatIcon COLOR=#F7AA22 SIZE=3>&#0252;</FONT></B>"
	sImagen4si = "<B><FONT FACE=CorFlatIcon COLOR=#993333 SIZE=3>4</FONT></B>"
	sImagenx   = "<B><FONT FACE=CorFlatIcon COLOR=#aa0000 SIZE=4>x</FONT></B>"



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


	sHTML = sHTML & ""

    if sformato <> "N" then
	Response.write sHTML
	sHTML = ""
	Response.flush
    end if
	Set cn = OFV.conectar(OFV.strconn2)

	xTrColor = "#FFFFFF"

	If Session("EMPMUN") = "S" Then

		sxlink2 = ""
		sxlink2 = sxlink2 & "<A class=nt "
		sxlink2 = sxlink2 & " HREF="
		sxlink2 = sxlink2 & "OBJtreeTema.asp?Tipo=Consulta&Nombre=TEMAS&Menu=Consulta&UN=" & aempre0
		sxlink2 = sxlink2 & " >"

	End If



	If Stnivel = 1 Then

		Session("XstemaSid")     = sID
		Session("XstemaSitem")   = sitem
		Session("XstemaStitem")  = stitem
		Session("XstemaStnivel") = Stnivel
		Session("XstemaStesq")   = stesq

	Else

		sesn                      = "X" & Stnivel
		Session(sesn & "Sid")     = sID
		Session(sesn & "Sitem")   = sitem
		Session(sesn & "Stitem")  = stitem
		Session(sesn & "Stnivel") = Stnivel
		Session(sesn & "Stesq")   = stesq

	End If



	stitemt     = "Tema"
	sIconot     = "ICorTema.jpg"
	sNuevoTipot = "SUBTEMAS"
	nNivelt     = 0

	StrSql = "Select codtema from prodserv where codprodserv = 0" & Session("XstemaSid")
	Set rsx = OFV.crearconsultaEx(StrSql, cn, 1, "")
	If Not rsx.EOF Then
		Strsql = "Select descri, codtema From Tema where codtema = " & rsx(0)
		Set rsx = OFV.crearconsultaEx(StrSql, cn, 1, "")
		If Not rsx.EOF Then
			stemaid   = rsx(1)
			stemades  = rsx(0)
		End If
	End If
	Call OFV.cerrarconsulta(rsx)



	sHref = "ObjTreeSubTemaCons.asp?Tipo=" & sNuevoTipot & "&ID=" & stemaid & "&sitem=TEMA:_"
	sHref = sHref & ofv.convertircar(stemades," ","_") & "&stitem="
	sHref = sHref & ofv.convertircar(stitemt," ","_") & "&sticono=" & sIconot
	sHref = sHref & "&stnivel=" & nNivelt &"&TIPOVISTA="&TIPOVISTA  &"&xbt=S" & "&formato=" & sformato


	sHref2 = "ConsObjetos.asp?ID=" & Session("XstemaSid")
	sHref2 = sHref2 & "&sitem=" & OFV.convertircar(Session("XstemaSitem"), " ", "_") & "&stitem="
	sHref2 = sHref2 & OFV.convertircar(Session("XstemaStitem"), " ", "_")
	sHref2 = sHref2 & "&stnivel=" & Session("XstemaStnivel") & "&stesq=" & Session("XstemaStesq")&"&TIPOVISTA="&TIPOVISTA &"&xbt=S"



	TABLA   = ofv.GenTabla("","aaa  style='border-style=solid;border-width:1;border-color:#eeeeee;' ","","90% ","","0","","-1","3","","","","no")
	sHTML   = sHTML & Tabla
	'''''''''''''''''''''''''''''

	''''''''''''''''' TEMA ''''''''''''
	Celda   = ""
	sxicon  = "ICorAbierto.jpg"

	''   sImage  = ofv.GenImage("","","","","Iconos\" & sxicon,"0","")

	'sImage  = "<font face='CorFlatIcon' size=3 color='#f7aa22'>4</font>"
	sImage  ="<font face='CorFlatIcon' size=3 color='#f7aa22'>&#052;</font>"
	sLink   = ofv.GenLink("","ut",sHref,"","","",sImage,"si")
	Celda   = Celda & ofv.GenCelda("","uts","center","2%","","","",sLink,"si")
	If sformato="ABR" Then
		sLink   = ofv.GenLink("","ut",sHref,"fesquemas","","","TEMA:&nbsp;" & stemades ,"si")
	Else

		sLink   = ofv.GenLink("","ut",sHref,"","","","TEMA:&nbsp;" & stemades,"si")
	End If
	Celda   = Celda & ofv.GenCelda("","ut bgcolor='" & xTrColor & "' ","","90%","","","",sLink,"si")

	Fila    = Fila & ofv.GenRow("","aaa  ","","","22","","",Celda,"si")

	sHTML   = sHTML & Fila
	sHTML   = sHTML & "<tr>"

	sLink   = "&nbsp;"
	sHTML   = sHTML & ofv.GenCelda("","ut","center","2%","","","",sLink,"si")
	sHTML   = sHTML & ofv.GenCelda("","ut","","2%","","","","","no")

	Fila    = ""
	sHTML   = sHTML & "<table cellpadding=2 cellspacing=0 width=98%   >"


	''''''''''''''''''''''''''''''


	''''''''''''''''' SUBTEMA ''''''''''''
	Celda   = ""
	sHref   = sHref2

	'' sImage  = ofv.GenImage("","","","","Iconos\" & sxicon,"0","")

	' sImage  = "<font face='CorFlatIcon' size=3 color='#f7aa22'>4</font>"
	sImage  ="<font face='CorFlatIcon' size=3 color='#f7aa22'>&#052;</font>"
	sLink   = ofv.GenLink("","ut",sHref,"","","",sImage,"si")
	Celda   = Celda & ofv.GenCelda("","uts","center","2%","","","",sLink,"si")
	If sformato="ABR" Then
		sLink   = ofv.GenLink("","ut",sHref,"fesquemas","","",Session("XstemaSitem"),"si")
	Else

		sLink   = ofv.GenLink("","ut",sHref,"","","",Session("XstemaSitem"),"si")
	End If
	Celda   = Celda & ofv.GenCelda("","ut bgcolor='" & xTrColor & "' ","","90%","","","",sLink,"si")

	Fila    = Fila & ofv.GenRow("","aaa  ","","","22","","",Celda,"si")

	sHTML   = sHTML & Fila
	sHTML   = sHTML & "<tr>"

	sLink   = "&nbsp;"
	sHTML   = sHTML & ofv.GenCelda("","ut","center","2%","","","",sLink,"si")
	sHTML   = sHTML & ofv.GenCelda("","ut","","2%","","","","","no")
	''''''''''''''''''''''''''''''





	If Stnivel > 1 Then
		For xxin = 2 To Stnivel

			sesn    = "X" & xxin
			sxlink  = ""
			sxlink  = sxlink & "<A class=nt HREF=ConsObjetos.asp?ID="
			sxlink  = sxlink & Session(sesn & "Sid") & "&sitem="
			sxlink  = sxlink & OFV.convertircar(Session(sesn & "Sitem"), " ", "_") & "&stitem="
			sxlink  = sxlink & OFV.convertircar(Session(sesn & "Stitem"), " ", "_")
			sxlink  = sxlink & "&stnivel=" & Session(sesn & "Stnivel") & "&stesq=" & Session(sesn & "Stesq")  &"&TIPOVISTA="&TIPOVISTA
			sxlink  = sxlink & "  >"

			sxlink  = ""
			sxlink  = sxlink & "ConsObjetos.asp?ID="
			sxlink  = sxlink & Session(sesn & "Sid") & "&sitem="
			sxlink  = sxlink & OFV.convertircar(Session(sesn & "Sitem"), " ", "_") & "&stitem="
			sxlink  = sxlink & OFV.convertircar(Session(sesn & "Stitem"), " ", "_")
			sxlink  = sxlink & "&stnivel=" & Session(sesn & "Stnivel") & "&stesq=" & Session(sesn & "Stesq") &"&TIPOVISTA="&TIPOVISTA


			Fila    = ""
			sHTML   = sHTML & "<table cellpadding=2 cellspacing=2 width=98%   >"

			If Session("USRES") <> "B" Then
				''        sHTML = sHTML & "<img border=0 src=Iconos\ICorAbierto.jpg>"
			Else
				''        sHTML = sHTML & "<b><font face=corflaticon size=4 color=#993333>1</font></b>"
			End If


			''''''''''''''''' ESQUEMAS SUPERIORES ''''''''''''
			Celda   = ""
			sHref   = sxlink

			''   sImage  = ofv.GenImage("","","","","Iconos\" & sxicon,"0","")

			'sImage  = "<font face='CorFlatIcon' size=3 color='#f7aa22'>4</font>"
			sImage  ="<font face='CorFlatIcon' size=3 color='#f7aa22'>&#052;</font>"
			sLink   = ofv.GenLink("","ut",sHref,"","","",sImage,"si")
			Celda   = Celda & ofv.GenCelda("","uts","center","2%","","","",sLink,"si")

			sLink   = ofv.GenLink("","ut",sHref,"","","",Session(sesn & "Sitem"),"si")
			Celda   = Celda & ofv.GenCelda("","ut bgcolor='" & xTrColor & "' ","","90%","","","",sLink,"si")

			Fila    = Fila & ofv.GenRow("","aaa  ","","","22","","",Celda,"si")

			sHTML   = sHTML & Fila
			sHTML   = sHTML & "<tr>"

			sLink   = "&nbsp;"
			sHTML   = sHTML & ofv.GenCelda("","ut","center","2%","","","",sLink,"si")
			sHTML   = sHTML & ofv.GenCelda("","ut","","2%","","","","","no")
			''''''''''''''''''''''''''''''

		Next
	End If



	Fila    = ""
	sHTML   = sHTML & "<table cellpadding=3 cellspacing=0 width=95%   >"
	If SFORMATO="ABR" Then
		StrSql = "SELECT DISTINCT OBJETOS.codobjeto, OBJETOS.descri, OBJETOS.tipobjeto, Estados.estado, "
		StrSql = StrSql & "Relaobj.tipobjasoc, OBJPRTY.fecemis, OBJPRTY.fecvigen, OBJPRTY.fecbaja,"
		StrSql = StrSql & " OBJPRTY.autor,OBJPRTY.aprobo, OBJPRTY.elimino, OBJPRTY.fecvto, OBJETOS.OBJUN, objetos.indobj,TIPOBJETO.TIPOCPT,OBJPRTY.fecpublic"
		StrSql = StrSql & " FROM RELAOBJ, OBJETOS, OBJPRTY, Estados,TIPOBJETO "
		StrSql = StrSql & " where "
		StrSql = StrSql & " RELAOBJ.codojasoc = OBJETOS.codobjeto and (TIPOBJETO.TIPOBJETO=OBJETOS.TIPOBJETO AND TIPOBJETO.TIPOCPT='C') AND "
		StrSql = StrSql & " OBJPRTY.estado    = Estados.objid and "
		StrSql = StrSql & " OBJETOS.codobjeto = OBJPRTY.codobjeto  and "
		StrSql = StrSql & " (objetos.objun =  " & session("UN") & " or "
		StrSql = StrSql & "  objetos.objun =  1 )  and "
		StrSql = StrSql & " RELAOBJ.codobjmaster = " & sID
		StrSql = StrSql & " and relaobj.tipobjmaster = '" & stitem & "' " & sVerEstado
		StrSql = StrSql & " order by objetos.indobj, OBJETOS.tipobjeto, OBJETOS.descri, Estados.estado "

	Else



		StrSql = "SELECT DISTINCT OBJETOS.codobjeto, OBJETOS.descri, OBJETOS.tipobjeto, Estados.estado, "
		StrSql = StrSql & "Relaobj.tipobjasoc, OBJPRTY.fecemis, OBJPRTY.fecvigen, OBJPRTY.fecbaja,"
		StrSql = StrSql & " OBJPRTY.autor,OBJPRTY.aprobo, OBJPRTY.elimino, OBJPRTY.fecvto, OBJETOS.OBJUN, objetos.indobj,tipobjeto.TIPOCPT,OBJPRTY.fecpublic"
		StrSql = StrSql & " FROM RELAOBJ, OBJETOS, OBJPRTY, Estados,TIPOBJETO "
		StrSql = StrSql & " where "
		StrSql = StrSql & " RELAOBJ.codojasoc = OBJETOS.codobjeto and "
		StrSql = StrSql & " OBJPRTY.estado    = Estados.objid and "
		StrSql = StrSql & " OBJETOS.codobjeto = OBJPRTY.codobjeto  and "
		StrSql = StrSql & " OBJETOS.tipobjeto = tipobjeto.tipobjeto  and "
		StrSql = StrSql & " (objetos.objun =  " & session("UN") & " or "
		StrSql = StrSql & "  objetos.objun =  1 )  and "
		StrSql = StrSql & " RELAOBJ.codobjmaster = " & sID
		StrSql = StrSql & " and relaobj.tipobjmaster = '" & stitem & "' " & sVerEstado
		StrSql = StrSql & "  and(OBJETOS.descri COLLATE SQL_Latin1_General_CP1_CI_AI  like '%" & sbuscarnew & "%' or OBJETOS.codiden COLLATE SQL_Latin1_General_CP1_CI_AI like '%" & sbuscarnew & "%')  " & filtro
		'StrSql = StrSql & " order by objetos.indobj, OBJETOS.tipobjeto, OBJETOS.descri, Estados.estado "
		StrSql = StrSql & ORDENDOC
	End If

	'StrSql = StrSql & "  order by objetos.indobj, OBJETOS.tipobjeto, OBJETOS.descri, Estados.estado"
	' response.write strsql
	sNuevoTipo = "ESQUEMAS"
	nNivel = Stnivel + 1
	Session("sID2") = sID
	'response.write strsql & "<br>"
	sVino = "N"
	Set rs = OFV.crearconsultaEx(StrSql, cn, 1, "")



	If rs.EOF Then

		sVino = "N"


	Else

		sVino = "N"

		nlinea = 0
		'aca
		'  response.write CantregMover
		rs.move CantregMover
		Do Until rs.EOF   'Or nlinea >= 14
			sVinook = "N"
			If  CStr(rs(12) & "") = CStr(session("UN") & "") Then
				'' response.write "<br><br>entro"
				sVino   = "S"
				sVinook = "S"
			End If

			If SESSION("EMPMUN") = "S" Then
				StrSql = "Select OBJMEMPALC.CODEMPRE  from OBJMEMPALC "
				StrSql = StrSql & " where "
				StrSql = StrSql & " OBJMEMPALC.codobjeto = " & rs(0) & " and "
				StrSql = StrSql & " OBJMEMPALC.CODEMPRE = " & Session("UN") & " "
				Set rsmas2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
				If  Not rsmas2.EOF Then
					sVino   = "S"
					sVinook = "S"
				End If
			End If

			sNIVE   = "D"

			If  sVinook = "S" Then

				stitem = rs(4)
				StrSql = "Select archivo, descri, perfil, TIPOCPT From Tipobjeto "
				StrSql = StrSql & "Where Tipobjeto.tipobjeto = '" & rs(4) & "'"
				Set rsIcono = OFV.crearconsultaEx(StrSql, cn, 1, "")
				If Not rsIcono.EOF Then
					sIcono  = rsIcono(0)
					sTipoE  = rsIcono(1)
					sPerfil = rsIcono(2)
					sNIVE   = rsIcono(3)
					'''   sitem   = sTipoE & ": " & rs(1) & " - (" & rs(3) & ")"

					'sitem = sTipoE & ": " & rs(1)
					sitem = EsqnomPx (rs(2),rs(0),ofv)

					If sNIVE = "D" Or sNIVE = "K" Then  sitem = sitem & " - (" & rs(3) & ")"


				End If
				Call OFV.cerrarconsulta(rsIcono)



				If sNIVE = "C" Then


					Strsql = "Select * From OBJSDK Where tipobjeto='" & rs(4) & "' and codobjeto = " & rs(0)
					Set rsIcono = ofv.crearconsultaEx(StrSql,cn,1,parametros)
					If Not rsIcono.eof Then


						If Session("Usuario") <> rsIcono(7) And rsIcono(5) Then


							Call ofv.cerrarconsulta(rsIcono)


							Strsql = "Select * From OBJSDKDET Where tipobjeto='" & rs(4)
							Strsql = Strsql & "' and codobjeto = "  & rs(0)
							Strsql = Strsql & " and ("
							Strsql = Strsql & "(grpusr = 'USR-" & Session("Usuario") & "') "

							'''''''''''''''''''''''''''''''''''''''''''
							If  Session("DELLOG") = "S" Then

								XMFTGRP = Split(session("SegGrpLst"),Chr(9))
								xcount  = UBound(XMFTGRP) + 1
								XMFSEP  = ""
								StrSql = StrSql & " or ( grpusr in ("

								For XMFI = 0 To UBound(XMFTGRP)
									StrSql = StrSql  & XMFSEP & "'GRP-" & XMFTGRP(XMFI) & "'"
									XMFSEP  = ","
								Next

								StrSql = StrSql & ")) "

							Else
								Strsql = Strsql & " or (grpusr = 'GRP-" & Session("Seggrp") & "')"
							End If


							''''''''''''''''''''''''''''''''''''''''''


							Strsql = Strsql & ") "

							'response.write strsql & "<br>"
							Set rsIcono = ofv.crearconsultaEx(StrSql,cn,1,parametros)
							If rsIcono.eof Then sVinook = "N"

						End If
					End If
					Call ofv.cerrarconsulta(rsIcono)



				End If

				If  sVinook = "S" Then



					smas = "&nbsp;"

					StrSql = "Select relaobj.codojasoc, objetos.objun,tipobjeto.TIPOCPT  from relaobj, OBJPRTY, Estados, objetos ,tipobjeto "

					StrSql = StrSql & " where "
					If sformato="ABR" Then
						StrSql = StrSql & " relaobj.codojasoc = OBJPRTY.codobjeto AND (objetos.TIPOBJETO=tipobjeto.TIPOBJETO AND tipobjeto.TIPOCPT='C' ) AND "
					Else

						StrSql = StrSql & " relaobj.codojasoc = OBJPRTY.codobjeto and "
					End If
					StrSql = StrSql & " objetos.codobjeto = relaobj.codojasoc and "
					StrSql = StrSql & " (objetos.objun =  " & session("UN") & " or "
					StrSql = StrSql & "  objetos.objun =  1 ) and "
					StrSql = StrSql & " OBJPRTY.estado = Estados.objid and "
					StrSql = StrSql & " relaobj.tipobjmaster = '" & rs(2) & "'"
					StrSql = StrSql & " and relaobj.codobjmaster = " & rs(0) & sVerEstado
					' response.write "<br><br><br><br>" & strsql & "<br>"
					'nuevo sql

					'   StrSql = " Select relaobj.codojasoc, objetos.objun from objetos "

					'	 StrSql = StrSql &  " inner join  relaobj on ( objetos.codobjeto = relaobj.codojasoc and relaobj.tipobjmaster = '" & rs(2) & "' and relaobj.codobjmaster = " &  rs(0) &"  )"
					'	 StrSql = StrSql &  " inner join  objprty  on (relaobj.codojasoc=objprty.codobjeto)"
					'	 StrSql = StrSql &  " inner join  Estados on (Estados.objid=objprty.estado)"
					'	 StrSql = StrSql &  " where  "
					'	 StrSql = StrSql &  " (objetos.objun = " & session("UN") & " or objetos.objun = 1 ) "
					' StrSql = StrSql &  " and  "

					'     StrSql = StrSql &  sVerEstado ' " Objetos.objnivel >= 3 and Objprty.Estado In (9,1,6,2,5,3,4) "

					'response.write "<br><br><br><br>" & strsql & "<br>"

					Set rsmas = OFV.crearconsultaEx(StrSql, cn, 1, "")
					If  Not rsmas.EOF Then
						'' response.write "<br><br>entro " & session("UN")
						Do Until rsmas.eof
							'' response.write "<br><br>entro 2 -" & session("UN") & " " & rsmas(1) & "<br>"
							If  CStr(rsmas(1) & "") = CStr(session("UN") & "") Then
								'' response.write "<br><br>entro"
								smas = "<i class='las la-folder-plus la-3x'></i>"

								Exit Do
							End If

							If SESSION("EMPMUN") = "S" Then
								StrSql = "Select OBJMEMPALC.CODEMPRE  from OBJMEMPALC "
								StrSql = StrSql & " where "
								StrSql = StrSql & " OBJMEMPALC.codobjeto = " & rsmas(0) & " and "
								StrSql = StrSql & " OBJMEMPALC.CODEMPRE = " & Session("UN") & " "
								Set rsmas2 = OFV.crearconsultaEx(StrSql, cn, 1, "")
								If  Not rsmas2.EOF Then
									smas = "<i class='las la-folder-plus la-3x'></i>"
									Exit Do
								End If
							End If

							rsmas.movenext
						Loop

					End If
					Call OFV.cerrarconsulta(rsmas)

					sxlink = ""
					sxlink  = sxlink & "ConsObjetos.asp?TIPOVISTA=" & TIPOVISTA &"&"
					sxlink  = sxlink & "ID=" & rs(0) & "&sitem="
					sxlink  = sxlink & OFV.convertircar(sitem, " ", "_") & "&stitem="
					sxlink  = sxlink & OFV.convertircar(stitem, " ", "_") & "&stnivel="
					sxlink  = sxlink & nNivel
					sxlink  = sxlink & "&stesq=" & OFV.convertircar(rs(2), " ", "_") & "&xbt=S&cantregmover=" & 0 & "&formato=" & sformato

					sHref   = sxlink


					If sNIVE   = "D" Or SNIVE   = "K"  Then


						If  smas <> "&nbsp;" Then


							'   IF session("APLMDCLI")="BC" then
							sFuncion = " onclick=" & Chr(34) & " Axn('" & rs(0) & "','" & stitem & "','" & sPerfil & "'); return false;" & Chr(34)
							sHref="'javacript:void(0);'	 "
							scon     = "<A class=nt href=" & sHref & sFuncion  & "' title='Ver...'  >"
							'    ELSE
							'  scon     = "<A class=nt href=" & sHref   & "' title='Expandir Estructura...'  >"
							'    END IF
							'  scon     = "<A class=nt href='" & sHrefw & "' title='ver...'  >"
							scon     = scon & "<font face='CorFlatIcon' size=2 color='#aa0000' style='border-style=solid;"
							scon     = scon & "border-width:1;border-color:#f9f9f9;' >&#050;</font>"
							scon     = scon & "</a>"

						Else
							scon     = "&nbsp;"




							sHref   = sHref & "&rsiguiente=" & rsiguiente & "&Ranterior=" & ranterior
						End If

					Else
						scon     = "&nbsp;"
					End If



					If CLng("0" & rs(12)) = 1 And  SESSION("UN") <> 1 And  SESSION("EMPMUN") = "S" Then
						smasc = "&nbsp;&nbsp;&nbsp;<font color=#993333 ><font class=btic title='Corporativo' >G</font></font>"
					Else
						smasc = "&nbsp;"
					End If


					Celda   = ""
					Celda   = Celda & ofv.GenCelda("","ut","center","1%","","","",smasxxx,"si")
					Celda   = Celda & ofv.GenCelda("","uts","center","2%","","","",scon,"si")



					If  sNIVE   = "D" Or SNIVE="K" Then

						sImage  = "<font face='CorFlatIcon' size=4 color='#005500' >&#0165;</font>"
						sHref    = "&nbsp;"
						sFuncion = " onclick=" & Chr(34) & " Axn('" & rs(0) & "','" & stitem & "','" & sPerfil & "'); return false;" & Chr(34)

                        If  smas <> "&nbsp;" Then
                        scon     =  "<A class=nt href='" & sxlink & "' " & wsFuncion & " title='Expandir Estructura...'  >" & smas
											'	scon     =  scon & masc & "nbsp;&nbsp;&nbsp;&nbsp" & smasc & smas
											    scon     =  scon & masc & "&nbsp;&nbsp;&nbsp" & smasc
						scon     = scon & "</a>"
                        end if
						scon     = scon & "<A class=nt href='" & sHref & "' " & sFuncion & " title='ver...'  >"
						scon     = scon & sImage
						scon     = scon & "</a>"

						sLink    = scon
						sLink2=sLink
						Celda   = Celda & ofv.GenCelda("","uts","center","2%","","","",sLink,"si")
                        'sli= sli & "<div style='background-color: lightgray; border: 1px solid black;padding: 10px;'>"  & sLink & "</div>"
						'aca

						scon     = "<A class=nt href='" & shref & "' " & sFuncion & " title='Ver...'  >"
						scon     =  scon & sitem '& smasc & "&nbsp;&nbsp;&nbsp;&nbsp" & smas
						scon     = scon & "</a>"




						sLink    = scon
						Celda   = Celda & ofv.GenCelda("","ut bgcolor='" & xTrColor & "' ","","95%","","","",sLink,"si")
						'sli= sli & "<div class=Dvindice >"  & sLink2 & "  " & sLink & "</div>"
						sli= sli & "<div class=Dvindice><img class=Imgmini src='"& ImgDefault("")& "'><p style='font-family: Arial;font-size: 13px;line-height: 1.3em;'>"  & sLink2 & "  " & sLink & "</P></div>"


					Else
						sImage   = "<font face='CorFlatIcon' size=4 color='#336699' >&#048;</font>"
						sLink    = ofv.GenLink("","ut",sHref,"","","",sImage,"si")

						Celda   = Celda & ofv.GenCelda("","uts","center","2%","","","",sLink,"si")
						If SFORMATO="ABR" Then
							sLink   = ofv.GenLink("","ut",sHref,"fesquemas","","",sitem & smasc & "&nbsp;&nbsp;&nbsp;&nbsp;" & smas,"si")
							sHrefEX=sHref & "&acc=UPD&NIVEL=" & SNIVE
							sdato="<font title=Confimar class=bt1>ü</font>"
							sLinkEX2   = ofv.GenLink("","ut",sHrefEX,"fesquemas","","",sdato,"si")
						Else
                            sLink   = ofv.GenLink("","ut",sHref,"","","", smas &"&nbsp;" &sitem & smasc  ,"si")
							'sLink   = ofv.GenLink("","ut",sHref,"","","",sitem & smasc & "&nbsp;&nbsp;&nbsp;&nbsp;" & smas,"si")
							sLinkEX2=""
						End If
						Celda   = Celda & ofv.GenCelda("","ut bgcolor='" & xTrColor & "' ","","95%","","","",sLink & sLinkEX2 ,"si")
                        if snive="C" then
                         sli= sli & "<div class=Dvindice>"   & sLink & sLinkEX2 & "</div>"
                        else
						 sli= sli & "<div class=Dvindice><img class=Imgmini src='"& ImgDefault("")& "'> "   & sLink & sLinkEX2 & "</div>"
						end if

					End If




					Fila    = Fila & ofv.GenRow("","aaa  ","","","25","","",Celda,"si")

					If  xTrColor = "#FFFFFF" Then
						xTrColor = "#F9F9F9"
					Else
						xTrColor = "#FFFFFF"
					End If

				End If

			End If


			rs.movenext
			nlinea = nlinea + 1
			if sformato <> "N" then
			 Response.write sHTML
			 sHTML = ""
            end if

			If nlinea > 20 and sformato<> "N" Then
				nlinea = 0
				Response.flush
			End If
		Loop

		'   tablab =  "<DIV id='DTBUSCAR' STYLE='float: left; z-index:1;overflow:auto;width:98%;height:50px;border-style:solid;border-width:2;border-color:#ffffff;' >"



		sInputT = "<input type='text' class='dt' name='BUSCAR' id='textot' value='"& sbuscar & "' size='40' style='border-style:solid;border-width:1;border-color:#666666;' >"

		CELDAx = ""
		CELDAx = CELDAx & ofv.GenCelda("", "UT  ", "right", "30%", "", "", "", sInputT, "si")

		sInput = ofv.GenerarInput("", "Buscar", "submit", "", "5", "at STYLE='color:#ffffff;font-weight:bold;background-color:#990000;border-style:solid;border-width:1;border-color:#dddddd;height:20px;cursor:pointer;' ", "")
		CELDAx = CELDAx & ofv.GenCelda("", "vt", "", "70%", "", "", "", sInput, "si")

		formx = ofv.GenRow("", "", "", "", "", "", "", CELDAx, "si")

		sTarget = "_self"
		saccion="consobjetos.asp?id=" & Request.querystring("ID") &"&tipovista=" & TIPOVISTA &"&sitem=" & Request.querystring("sitem") & "&stitem=" & Request.querystring("stitem") & "&stnivel=" &  Request.querystring("stnivel") & "&cantregmover=0"


		xform =  ofv.GenTabla("", "aaa xSTYLE='border-style:solid;border-width:1;border-color:#dddddd;' ", "", "50%", "", "0", "", "-2", "0", "", "", formx, "si")
		'Filax = ofv.genform("Subconsulta", "", sAccion, sTarget, "post", "", xform, "si")
        sAccionf=sAccion
		tablab = filax
		sxlinka = ""
		sxlinka = sxlinka & "<A class=nt HREF=consobjetos.asp?id=" & Request.querystring("ID") &"&tipovista=" & TIPOVISTA &"&sitem=" & Request.querystring("sitem") & "&stitem=" & Request.querystring("stitem") & "&stnivel=" &  Request.querystring("stnivel") & "&cantregmover=" & cantregmover + nlinea &"> <font size=4 color=red face='Corflaticon'>&#052;</font></a>"

		If   rs.eof Then sxlinka=""
		sxlinkd = ""
		If  cantregmover - nlinea < 0 Then
			cantregmover =0
			sxlinkd=""
		Else
			cantregmover =cantregmover - 14
			sxlinkd =   "<A class=nt HREF=consobjetos.asp?id=" & Request.querystring("ID") &"&tipovista=" & TIPOVISTA &"&sitem=" & Request.querystring("sitem") & "&stitem=" & Request.querystring("stitem") & "&stnivel=" &  Request.querystring("stnivel") & "&cantregmover=" & cantregmover &"><font size=4 color=red face='Corflaticon'>&#051;</font></a>"
		End If

		Celda   = ""
		If sxlinkd="" And sxlinka="" Then  tablab=""
		sLink   = "&nbsp;"
		Celda   = Celda & ofv.GenCelda("","ut","left colspan=5","2%","","","",sxlinkd & "&nbsp;" & sxlinka & tablab,"si")
		Fila    = Fila & ofv.GenRow("","aaa bgcolor='" & xTrColor & "' ","","","25","","",Celda,"si")
		'response.write sxlinkd & "&nbsp;" & sxlinka & wtablab

	End If

	If  sVino = "N" Then

		Celda   = ""

		sLink   = "&nbsp;"
		Celda   = Celda & ofv.GenCelda("","ut","center","2%","","","",sLink,"si")

		sLink   = "&nbsp;"
		Celda   = Celda & ofv.GenCelda("","ut","center","2%","","","",sLink,"si")

		sLink   = "No hay Esquemas asociados"
		Celda   = Celda & ofv.GenCelda("","ttt","","90%","","","",sLink,"si")

		Fila    = Fila & ofv.GenRow("","aaa bgcolor='" & xTrColor & "' ","","","25","","",Celda,"si")

	End If

	sHTML = sHTML & Fila
	sHTML = sHTML & "</table>"

	If  Stnivel > 1 Then
		For xxin = 2 To Stnivel

			sHTML = sHTML & "</td>"
			sHTML = sHTML & "</tr>"
			sHTML = sHTML & "</table>"

		Next
	End If

	sHTML = sHTML & "</td>"
	sHTML = sHTML & "</tr>"
	sHTML = sHTML & "</table>"

	sHTML = sHTML & "</td>"
	sHTML = sHTML & "</tr>"
	sHTML = sHTML & "</table>"


	Call OFV.cerrarconsulta(rs)



	'Botonera
	Agregar = "" 'sxlinkd & "&nbsp;" & sxlinka & tablab
	Buscar  = ""
	Celda   = ""

	If session("SMPLMENU") = "S" Then

		sAccion = "novedmenu.asp"
		xxxx    = CorGenBotonXp(16, "NO", "home", "tahoma", "NO", "Home", "_self", "Home")
		Celda   = ofv.GenBoton(16, "esp1Bot", "BT", "S", "", "S", "A", "N", "esp1Frm", sAccion & "")
	End If


 if sformato="N" then
   sHTML      = ofv.FormHeader(session("FormN")) & xs
   xconteledor= "<div class='contenedor'  >"


   titulo="<a href='subconsulta.asp?MENU=PORTEMA&NOMBRE=TEMA&NDAT=ESTADO&NVAL=2&seg=CLIMS154172171124031051096140&MNN=POR_INDICE&base=D'>" & ATREMP(0,4) &" / </a> "


   saccion="ObjTreeSubTemaCons.asp?formato=" & sformato &"&ID=" & stemaid & "&sitem=" & stemades & "&stitem=Tema&stnivel=0&tipo=SUBTEMAS"
   titulo=titulo & "<a href='" & saccion & "' >TEMA : " & stemades  & " / </a>"
   saccion="consobjetos.asp?id=" & Session("XstemaSid") &"&tipovista=" & TIPOVISTA &"&sitem=" & Session("XstemaSitem") & "&stitem=" & Session("XstemaStitem") & "&stnivel=" &  Session("XstemaStnivel") & "&cantregmover=0&formato=" & sformato
   titulo=titulo & "<a href='" & saccion & "' >" & Session("XstemaSitem")  & " </a>"
   delim=""




   for i =1 to  Stnivel
    sesn                      = "X" & i
    saccion="consobjetos.asp?id=" & Session(sesn & "Sid") &"&tipovista=" & TIPOVISTA &"&sitem=" & Session(sesn & "Sitem") & "&stitem=" & Session(sesn & "Stitem") & "&stnivel=" &  Session(sesn & "Stnivel") & "&cantregmover=0&formato=" & sformato
    titulo=titulo & delim & "<a href='" & saccion & "' >" & Session(sesn & "Sitem")  & " </a>"

    delim=" / "
   next
   	xconteledor=xconteledor &     "<div class='header'><button id='abrir-dialog'>Filtrar</button>" & titulo  &"</div>"
    xconteledor=xconteledor &     "<div class='contenido'  >"

   	xconteledor=xconteledor & "<div id=resultados>  "
    xconteledor=xconteledor &  sli &"</div>"

    xconteledor=xconteledor &     "</div>"
    xconteledor=xconteledor &"  </div>"

    sHTML=sHTML & xconteledor
   end if
	sHTML = sHTML & ofv.botoneraESSX("0","","","","","","","","","no",Agregar,Buscar,celda)
sdialog="<div id='dialog-form' title='Realizar Búsqueda'>"
sdialog= sdialog & "       <form method=post id=cc ACTION='" &sAccionf&"&formato=N' target=enviaf >"
sdialog= sdialog & "            <fieldset>"
sdialog= sdialog & "                <label for='tipobjeto'>Tipos de Documentos</label>"
sdialog= sdialog & "                <select id='tipobjeto' name='tipobjeto' multiple='multiple' style='width:100%;'>"
sql="select TIPOBJETO,DESCRI from tipobjeto where tipocpt in ('D') ORDER BY DESCRI"
Set rsT = OFV.crearconsultaEx(SQL, cn, 1, "")

DO WHILE NOT RST.EOF
sdialog= sdialog & "<option value='" & RST(0) & "'>" & RST(1) &"</option>"
RST.MOVENEXT
LOOP
Call OFV.cerrarconsulta(rsT)

sdialog= sdialog & "                </select>"

sdialog= sdialog & "                <br><br>"
sdialog= sdialog & "  <label for='ORDEN'>Ordenar Por</label>"
sdialog= sdialog & "            <select name='ORDEN' id='ORDEN' class='ui-widget-content ui-corner-all full-width-select'>"
sdialog= sdialog & "                <option value='NO'>SIN ORDEN</option>"
sdialog= sdialog & "                <option value='FECEMIS'>FECHA DE CREACION </option>"
sdialog= sdialog & "                <option value='FECEMIS DESC'>FECHA DE CREACION DESC</option>"
sdialog= sdialog & "                <option value='FECPUBLIC'>FECHA DE PUBLICACION </option>"
sdialog= sdialog & "                <option value='FECPUBLIC DESC'>FECHA DE PUBLICACION DESC</option>"
sdialog= sdialog & "                <option value='FECVIGEN'>FECHA DE VIGENCIA</option>"
sdialog= sdialog & "                <option value='FECVIGEN DESC'>FECHA DE VIGENCIA DESC</option>"
sdialog= sdialog & "            </select>"
sdialog= sdialog & "                <br><br>"
sdialog= sdialog & "                <label for='termino-busqueda'>Nombre del documento a buscar</label>"
sdialog= sdialog & "                <input type='termino-busqueda' id='termino-busqueda' name='texto'  class='ui-widget-content ui-corner-all' style='width: 98%; padding: 5px;'>"
sdialog= sdialog & ""
sdialog= sdialog & "            </fieldset>"
sdialog= sdialog & "        </form>"
sdialog= sdialog & "    </div>"
sdialog= sdialog & "<div id='resultados' style='display:none'>"
sdialog= sdialog & "        <p>Los resultados de la búsqueda aparecerán aquí.</p>"
sdialog= sdialog & "    </div>"


sdialog= sdialog & "<div id='loader-wrapper' sSTYLE='display:none'>"
sdialog= sdialog & "        <div class='loader'></div>"
sdialog= sdialog & "        <p>Cargando</p>"
sdialog= sdialog & "    </div>"
sHTML = sHTML & sdialog & "<iframe WIDTH=100% HEIGHT=400PX  style='display:none' id=enviaf name=enviaf></iframe></body></html>"




end if
	'''''''''''''

if request.querystring("salidad")="S" then
  response.write sli
else
  response.write sHTML
end if

function sqlejec(sql,cn,byref t)

STRSQLS= "PRC_EXECSQL '" & Replace(SQL,"'","''") & "'"

Set VV=ofv.crearconsultaEx(StrSqlS,cn,1,volver)
	If cn.errors.count = 0 Then
		If Not IsNull(VV(0)) Then
	     sqlejec=vv(1)
	     t=VV(0)
		End If
   end if
end function

Call OFV.cerrarconn(cn)
If ok <> "Y" Then response.flush



%>
