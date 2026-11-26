<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()

response.expires=0

smodal = request.querystring("modal")

sElemento = request.querystring("ID")
'''response.write sElemento & "<br>"
n = InStr(sElemento,"-")
nCodobjeto = Left(sElemento,n-1)
n2 = InStr(n+1,sElemento,"-")
sTipobjeto = Mid(sElemento,n+1,n2-(n+1))
sEstado = Mid(sElemento,n2+1)
sEsqSup = Trim(sTipobjeto)
sEsqSup = Mid(sEsqSup,Len(sEsqSup))
sTipobjeto = Left(sTipobjeto,Len(sTipobjeto)-1)

ofv.ObtenerAtributos Session("Form2"),""
bAgregar   = ofv.AAgregar
bModificar = ofv.AModificar
bSuprimir  = ofv.AEliminar
bConsultar = ofv.AConsultar

Set cn = ofv.conectar(ofv.strconn2)

scodigo = nCodobjeto
stipo   = sTipobjeto
strsql = "SELECT tipocpt FROM tipobjeto  where tipobjeto ='" & stipo & "'"


Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
If Not rs.eof Then
	tipoPedido=rs(0)
Else
	tipoPedido=""
End If
Call ofv.cerrarconsulta(rs)

srestric = "N"

'xxxxxxxxxxxxxxxxxxxxxxxx

If sEsqSup <> "Y" Then
	ReDim Preserve sPath(0)
	i = 1
	Do

		strsql = "select * from csbarr where codojasoc = " & sCodigo

		strsql = strsql & " and tipobjasoc = '" & sTipo & "'"
		Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

		If Not rs.eof Then
			ReDim Preserve sPath(i)
			sCodigo = rs(1)
			sTipo = rs(2)
			If Not IsNull(IsNull(rs(0))) Then
				sPath(i) = rs(0)
			Else
				i = i + 1
				Exit Do
			End If
			i = i + 1
		Else
			Exit Do
		End If
		Call ofv.cerrarconsulta(rs)
	Loop

End If

strsql = "SELECT * FROM PRODSERV where codprodserv = " & sCodigo

Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)

If Not rs.eof Then
	If rs("segaceso") <> 0 Or rs("segactual") <> 0 Then srestric = "S"

End If
'   response.write rs("segaceso") & "   "   &  rs("segactual") & srestric
Call ofv.cerrarconsulta(rs)

If srestric = "S" Then
	StrSql = "Select * From subtemaskr "
	StrSql = StrSql & "Where codprodserv = " & sCodigo & " and tiposeg = 'T' "
	StrSql = StrSql & " and (grpusr = 'USR-" & UCase(session("usuario")) & "' "

	'''''''''''''''''''''''''''''''''''''''''''
	If  Session("DELLOG") = "S" Then

		XMFTGRP = Split(session("SegGrpLst"),Chr(9))
		xcount  = UBound(XMFTGRP) + 1
		XMFSEP  = ""
		StrSql = StrSql & " or grpusr in ("

		For XMFI = 0 To UBound(XMFTGRP)
			StrSql = StrSql  & XMFSEP & "'GRP-" & XMFTGRP(XMFI) & "'"
			XMFSEP  = ","
		Next

		StrSql = StrSql & ") "

	Else
		StrSql = StrSql & " or grpusr = 'GRP-" & UCase(Session("seggrp")) & "' "
	End If


	''''''''''''''''''''''''''''''''''''''''''

	StrSql = StrSql & ")"

	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If Not rs.eof Then
		srestric = "N"
	Else
		srestric = "S"
	End If
	Call ofv.cerrarconsulta(rs)
End If



If srestric   = "S" Then
	bAgregar   = 0
	bModificar = 0
	bSuprimir  = 0
End If


'xxxxxxxxxxxxxxxxxxxxxxxx

'v TEMAS         CLIMS2751557214332898409701  CLIMS28915558161334435117553 FORMULARIOS VIGENTES CLIMS46712740044434860617908
'd DESARROLLO    CLIMS2761557214333698409701  CLIMS29015558161341435117553 FORMULARIOS DESARROLLO CLIMS4681274004444386061790
'a PARA APROBAR  CLIMS2771557214334498409701  CLIMS29115558161348435117553 FORMULARIOS PARA APROBAR CLIMS48112739033140999275609
'n ELIMINADAS    CLIMS2781557214335198409701  CLIMS29215558161356435117553 FORMULARIOS ELIMINADAS CLIMS48212739033350999275609



'response.write smodal

session("Nomodificar") = "N"

If smodal <> "PROCESOS" Then

	'response.write Session("Form2")
	If IsNumeric(sEstado) Then
		'and  Session("Form2") <> "DESARROLLO"
		If Session("Form2") <> "CLIMS46712740044434860617908" And Session("Form2") <> "TEMAS" And Session("Form2") <> "CLIMS2751557214332898409701" And Session("Form2") <> "CLIMS28915558161334435117553" And  sEstado = 3 Then
			bSuprimir = False
			session("Nomodificar") = "S"
		End If
		If  Session("Form2") <> "CLIMS4681274004444386061790" And Session("Form2") <> "DESARROLLO"    And Session("Form2") <> "CLIMS2761557214333698409701" And Session("Form2") <> "CLIMS29015558161341435117553" And sEstado = 1 Then
			bSuprimir = False
			session("Nomodificar") = "S"
		End If
		If Session("Form2") <> "CLIMS48112739033140999275609" And Session("Form2") <> "PARA APROBAR"  And Session("Form2") <> "CLIMS2771557214334498409701" And Session("Form2") <> "CLIMS29115558161348435117553" And sEstado = 2 Then
			bSuprimir = False
			session("Nomodificar") = "S"
		End If
	End If

	If Session("Form2") = "CLIMS48212739033350999275609" Or Session("Form2") = "TEMAS" Or Session("Form2") = "ELIMINADAS" Or Session("Form2") = "CLIMS2751557214332898409701" Or Session("Form2") = "CLIMS2781557214335198409701" Or Session("Form2") = "CLIMS28915558161334435117553"  Or Session("Form2") = "CLIMS29215558161356435117553" Then
		bAgregar = False
		'session("Nomodificar") = "S"
	End If

Else

	bSuprimir = False
	bAgregar  = False

End If

'session("Nomodificar")="S"


session("EsqSup") = sEsqSup



Session("Segct") = "S"
Session("Segmo") = "S"
Session("Segco") = "S"
Session("Segce") = "S"

If sEsqSup = "Y" Then
	xtipo = "SUBTEMAS"
	strsql = "SELECT DESCRI FROM PRODSERV where codprodserv = " & NCODOBJETO
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If Not rs.eof Then sNombredOC = rs(0)


	Strsql = "SELECT SUBTEMA.* FROM PRODSERV, SUBTEMA where "
	Strsql = Strsql & "PRODSERV.marcaps = SUBTEMA.tipobjeto and "
	Strsql = Strsql & " prodserv.marcaps = '" & sTipobjeto & "'"
Else
	xtipo = "ESQUEMAS"
	strsql = "SELECT DESCRI FROM OBJETOS where CODOBJETO = " & NCODOBJETO
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If Not rs.eof Then sNombredOC = rs(0)

	Strsql = "Select * From Tipobjeto Where tipobjeto='" & sTipobjeto & "'"
End If
'ss=strsql

'response.write strsql
Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
If sEsqSup <> "Y" Then
	TIPOCPT=rs("TIPOCPT")
Else
	TIPOCPT=""
End If

x = 0
ReDim Niveles(1)
If Not rs.eof Then
	sNombre  = rs(1)
	sTIPOdOC = rs(1)
	sPerfil = rs(3)
	For z = 5 To 14
		If rs(z) Then
			x = x + 1
			ReDim Preserve Niveles(x)
			Niveles(x) = Right(rs(z).name,1)
			'    RESPONSE.WRITE right(rs(z).name,1) & "<BR>"
		End If
	Next
End If

Call ofv.cerrarconsulta(rs)



sHTML = "<html xmlns:v='urn:schemas-microsoft-com:vml'  ><title>" & sNombre & "</title>"
sHTML = sHTML & "<link rel=""stylesheet"" type=""text/css"" href=""../fonts/Corfont/Corflaticon.css"">"

'sHTML = sHTML & "<STYLE>v\:* {behavior:url(#default#VML);}</STYLE>"
'sHTML = sHTML & "<STYLE>"
'sHTML = sHTML & " input.bt{font-family:" & Session("CMFPNBOTFTFC") & ";color:" & Session("CMFPNBOTFT") & ";font-size:" & Session("CMFPNBOTFTSZ") & "pt;font-weight:bold;" & XspEx("#dddddd",session("CMFPNBOT"),0) & ";border-style:solid;border-width:1;border-color:#dddddd;height:20px;cursor:pointer;}"
'sHTML = sHTML & " td.vti{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:600;padding-left:20px;}"
'sHTML = sHTML & " td.ttr{font-family:" & Session("CMFPNTITFTFC") & ";color:" & Session("CMFPNTITFT") & ";font-size:" & Session("CMFPNTITFTSZ") & "pt;text-transform:uppercase;font-weight:600;line-height:10pt;text-indent:4pt;padding-right:4pt;border-style:none;border-width:1;border-color:#eeeeee;" & XspEx("#EEEEEE",Session("CMFPNTIT"),0) & ";}"
'sHTML = sHTML & " td.uti{font-family:" & Session("CMFPNGNFTFC") & ";color:" & Session("CMFPNGNFT") & ";font-size:" & Session("CMFPNGNFTSZ") & "pt;font-weight:500;padding-left:20px;}"
'sHTML = sHTML & " td.cc{font-family:" & Session("CMFPNITFTFC") & ";color:" & Session("CMFPNITFT") & ";text-indent:4px;font-size:" & Session("CMFPNITFTSZ") & "pt;background-color:" & Session("CMFPNIT") & ";border-style:solid;border-width:1;border-color:#eeeeee;}"
'sHTML = sHTML & "</STYLE>"


	    wk1 = ""

        wk1 = wk1 & "<STYLE>"

        wk1 = wk1 & ".mini {"
        wk1 = wk1 & "  border-radius: 10%;"
        wk1 = wk1 & "  box-shadow: 4px 4px 5px grey;"
        wk1 = wk1 & "  width: 90px;"
        wk1 = wk1 & "  height: 60px;"
        wk1 = wk1 & " }"

        wk1 = wk1 & "html,"
        wk1 = wk1 & "body { height: 100%; }"

        wk1 = wk1 & "body {"
        wk1 = wk1 & "  margin: 0;"
        wk1 = wk1 & "  background: linear-gradient(#eeeeee, #cccccc); "
        wk1 = wk1 & "  overflow: hidden;"
        wk1 = wk1 & "}"

        wk1 = wk1 & ".selector {"
        wk1 = wk1 & "  position: absolute;"
        wk1 = wk1 & "  left: 50%;"
        wk1 = wk1 & "  top: 50%;"
        wk1 = wk1 & "  width: 140px;"
        wk1 = wk1 & "  height: 140px;"
        wk1 = wk1 & "  margin-top: -70px;"
        wk1 = wk1 & "  margin-left: -70px;"
        wk1 = wk1 & "}"

        wk1 = wk1 & ".selector,"
        wk1 = wk1 & ".selector button {"
        wk1 = wk1 & "  font-family: 'Oswald', sans-serif;"
        wk1 = wk1 & "  font-weight: 300;"
        wk1 = wk1 & "}"

        wk1 = wk1 & ".selector button {"
        wk1 = wk1 & "  position: relative;"
        wk1 = wk1 & "  width: 100%;"
        wk1 = wk1 & "  height: 100%;"
        wk1 = wk1 & "  padding: 10px;"
        wk1 = wk1 & "  background: #428bca;"
        wk1 = wk1 & "  border-radius: 50%;"
        wk1 = wk1 & "  border: 0;"
        wk1 = wk1 & "  color: white;"
        wk1 = wk1 & "  font-size: 20px;"
        wk1 = wk1 & "  cursor: pointer;"
        wk1 = wk1 & "  box-shadow: 0 3px 3px rgba(0, 0, 0, 0.1);"
        wk1 = wk1 & "  transition: all .1s;"
        wk1 = wk1 & "}"

        wk1 = wk1 & ".selector button:hover { background: #3071a9; }"

        wk1 = wk1 & ".selector button:focus { outline: none; }"

        wk1 = wk1 & ".selector ul {"
        wk1 = wk1 & "  position: absolute;"
        wk1 = wk1 & "  list-style: none;"
        wk1 = wk1 & "  padding: 0;"
        wk1 = wk1 & "  margin: 0;"
        wk1 = wk1 & "  top: -20px;"
        wk1 = wk1 & "  right: -20px;"
        wk1 = wk1 & "  bottom: -20px;"
        wk1 = wk1 & "  left: -20px;"
        wk1 = wk1 & "}"

        wk1 = wk1 & ".selector li {"
        wk1 = wk1 & "  position: absolute;"
        wk1 = wk1 & "  width: 0;"
        wk1 = wk1 & "  height: 100%;"
        wk1 = wk1 & "  margin: 0 50%;"
        wk1 = wk1 & "  -webkit-transform: rotate(-360deg);"
        wk1 = wk1 & "  transition: all 0.8s ease-in-out;"
        wk1 = wk1 & "}"

        wk1 = wk1 & ".selector li input { display: none; }"

        wk1 = wk1 & ".selector li input + label {"
        wk1 = wk1 & "  position: absolute;"
        wk1 = wk1 & "  left: 50%;"
        wk1 = wk1 & "  bottom: 100%;"
        wk1 = wk1 & "  width: 0;"
        wk1 = wk1 & "  height: 0;"
        wk1 = wk1 & "  line-height: 1px;"
        wk1 = wk1 & "  margin-left: 0;"
        wk1 = wk1 & "  background: #fff;"
        wk1 = wk1 & "  border-radius: 50%;"
        wk1 = wk1 & "  text-align: center;"
        wk1 = wk1 & "  font-size: 1px;"
        wk1 = wk1 & "  overflow: hidden;"
        wk1 = wk1 & "  cursor: pointer;"
        wk1 = wk1 & "  box-shadow: none;"
        wk1 = wk1 & "  transition: all 0.8s ease-in-out, color 0.1s, background 0.1s;"
        wk1 = wk1 & "}"

        wk1 = wk1 & ".selector li input + label:hover { background: #f0f0f0; }"

        wk1 = wk1 & ".selector li input:checked + label {"
        wk1 = wk1 & "  background: #5cb85c;"
        wk1 = wk1 & "  color: white;"
        wk1 = wk1 & "}"

        wk1 = wk1 & ".selector li input:checked + label:hover { background: #449d44; }"

        wk1 = wk1 & ".selector.open li input + label {"
        wk1 = wk1 & "  width: 80px;"
        wk1 = wk1 & "  height: 80px;"
        wk1 = wk1 & "  line-height: 80px;"
        wk1 = wk1 & "  margin-left: -40px;"
        wk1 = wk1 & "  box-shadow: 0 3px 3px rgba(0, 0, 0, 0.1);"
        wk1 = wk1 & "  font-size: 14px;"
        wk1 = wk1 & "}"

        wk1 = wk1 & ".botO1 {"
        wk1 = wk1 & "  width:60px;height:60px;border-radius: 60%;color:white;background: linear-gradient(#000088, #99ccee);box-shadow: 0 3px 3px rgba(0, 0, 0, 0.8);"
        wk1 = wk1 & "}"

        wk1 = wk1 & ".botO2 {"
        wk1 = wk1 & "  width:65px;height:40px;border-radius: 15px;color:white;background: linear-gradient(#000088, #99ccee);box-shadow: 0 5px 5px rgba(0, 0, 0, 0.8);"
        wk1 = wk1 & "  -webkit-transform: rotate(-360deg);"
        wk1 = wk1 & "  transition: all 0.8s ease-in-out;"
        wk1 = wk1 & "}"

        wk1 = wk1 & ".botA2 {"
        wk1 = wk1 & "  width:200px;height:30px;border-radius: 15px;color:white;background: linear-gradient(#f7aa22, #f76644);box-shadow: 0 5px 5px rgba(0, 0, 0, 0.8);"
        wk1 = wk1 & "  border-style:solid;border-width:1px;border-color:grey;"
        wk1 = wk1 & "  -webkit-transform: rotate(-360deg);"
        wk1 = wk1 & "  transition: all 0.8s ease-in-out;"
        wk1 = wk1 & "}"



        wk1 = wk1 & "</STYLE>"


        wk1 = wk1 & "<script src=" & chr(34) & "http://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js" & chr(34) & " ></script>"
        wk1 = wk1 & "<script>"
        wk1 = wk1 & "var nbOptions = 8;"
        wk1 = wk1 & "var angleStart = -360;"

        wk1 = wk1 & "// jquery rotate animation"
        wk1 = wk1 & "function rotate(li,d) {"
        wk1 = wk1 & "    $({d:angleStart}).animate({d:d}, {"
        wk1 = wk1 & "        step: function(now) {"
        wk1 = wk1 & "            $(li)"
        wk1 = wk1 & "               .css({ transform: 'rotate('+now+'deg)' })"
        wk1 = wk1 & "               .find('label')"
        wk1 = wk1 & "                  .css({ transform: 'rotate('+(-now)+'deg)' });"
        wk1 = wk1 & "        }, duration: 0"
        wk1 = wk1 & "    });"
        wk1 = wk1 & "}"

        wk1 = wk1 & "// show / hide the options"
        wk1 = wk1 & "function toggleOptions(s) {"
        wk1 = wk1 & "    $(s).toggleClass('open');"
        wk1 = wk1 & "    var li = $(s).find('li');"
        wk1 = wk1 & "    var deg = $(s).hasClass('half') ? 180/(li.length-1) : 360/li.length;"
        wk1 = wk1 & "    for(var i=0; i<li.length; i++) {"
        wk1 = wk1 & "        var d = $(s).hasClass('half') ? (i*deg)-90 : i*deg;"
        wk1 = wk1 & "        $(s).hasClass('open') ? rotate(li[i],d) : rotate(li[i],angleStart);"
        wk1 = wk1 & "    }"
        wk1 = wk1 & "}"

        wk1 = wk1 & "$('.selector button').click(function(e) {"
        wk1 = wk1 & "    toggleOptions($(this).parent());"
        wk1 = wk1 & "});"

        wk1 = wk1 & "setTimeout(function() { toggleOptions('.selector'); }, 100);//@ sourceURL=pen.js"
        wk1 = wk1 & "</script>"
        wk1 = wk1 & "<script type='text/javascript'>"

        wk1 = wk1 & "  var _gaq = _gaq || [];"
        wk1 = wk1 & "  _gaq.push(['_setAccount', 'UA-36251023-1']);"
        wk1 = wk1 & "  _gaq.push(['_setDomainName', 'jqueryscript.net']);"
        wk1 = wk1 & "  _gaq.push(['_trackPageview']);"
        wk1 = wk1 & ""
        wk1 = wk1 & "  (function() {"
        wk1 = wk1 & "    var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;"
        wk1 = wk1 & "    ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';"
        wk1 = wk1 & "    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);"
        wk1 = wk1 & "  })();"

        wk1 = wk1 & "</script>"







	sHTML = sHTML & "<body  onload=""window.top.LDVENTANA('DXM');"" >"
	sHTML = sHTML & wk1

xs = "<script language=JavaScript>" & Chr(13)
xs = xs & "function AbrirAgregar(Datos) {" & Chr(13)
If BrowserName()="IE" Then
	xs = xs & " window.dialogArguments.open(Datos"
	xs = xs & "," & Chr(34) & "edi" & Chr(34) & "," & Chr(34)
	xs = xs & "resizable=yes,width=500,height=500,menubar=no,toolbar=no" & Chr(34)
	xs = xs & " ); " & Chr(13)
	'wxs = wxs & " window.open(" & Chr(34) & sASP & Chr(34)
Else
	xs = xs & " window.opener.open(Datos"
	xs = xs & "," & Chr(34) & "edi" & Chr(34) & "," & Chr(34)
	xs = xs & "resizable=yes,width=500,height=500,menubar=no,toolbar=no" & Chr(34)
	xs = xs & " ); " & Chr(13)
End If

' xs = xs & " window.dialogArguments.open(Datos"
' xs = xs & "," & chr(34) & "edi" & chr(34) & "," & chr(34)
'xs = xs & "resizable=yes,width=500,height=500,menubar=no,toolbar=no" & chr(34)
'xs = xs & " ); " & chr(13)
xs = xs & "window.close(); " & Chr(13)
xs = xs & " } " & Chr(13)

			tabla1f = ""
			tabla1f = tabla1f & "function MenuInOut(va) {"
			tabla1f = tabla1f & "var i;"
			tabla1f = tabla1f & "var x = document.getElementsByClassName('mySlides'); "
			tabla1f = tabla1f & "for (i = 0; i < x.length; i++) {"
			tabla1f = tabla1f & "      x[i].style.display = 'none'; "
			tabla1f = tabla1f & "    }"
			tabla1f = tabla1f & "    slideIndex++;"
			tabla1f = tabla1f & "    if (slideIndex > x.length) {slideIndex = 1}"
			tabla1f = tabla1f & "    x[slideIndex-1].style.display = 'block'; "
			tabla1f = tabla1f & "    setTimeout(MenuInOut(va), 4000); "
			tabla1f = tabla1f & "}"

xs = xs & tabla1f

xs = xs & "</script>" & Chr(13)

sHTML = sHTML & xs

inder = 0

Session("SegRst") = "N"

If sEsqSup = "N" Then

	Set cn = ofv.conectar(ofv.strconn2)

	Strsql = "Select * From OBJSDK Where tipobjeto='" & sTipobjeto & "' and codobjeto = "  & nCodobjeto
	Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
	If Not rs.eof Then


		If Session("Usuario") <> rs(7) Then

			If rs(3) Then Session("Segct") = "N"
			If rs(4) Then Session("Segmo") = "N"
			If rs(5) Then Session("Segco") = "N"
			If rs(6) Then Session("Segce") = "N"

			If rs(3) Or rs(4) Or rs(6)   Then Session("SegRst") = "S"


			Call ofv.cerrarconsulta(rs)
			Strsql = "Select * From OBJSDKDET Where tipobjeto='" & sTipobjeto
			Strsql = Strsql & "' and codobjeto = "  & nCodobjeto
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

			Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			Do Until rs.eof

				Select Case rs(3)
					Case "1"
					Session("Segct") = "S"
					Session("Segmo") = "S"
					Session("Segco") = "S"
					Session("Segce") = "S"

					Case "2"
					Session("Segmo") = "S"
					Session("Segco") = "S"

					Case "3"
					Session("Segco") = "S"

					Case "4"
					Session("Segco") = "S"
					Session("Segce") = "S"

				End Select
				rs.movenext
			Loop
		End If
	End If
	Call ofv.cerrarconsulta(rs)

	If srestric = "S" Then
		Session("Segct") = "N"
		Session("Segmo") = "N"
		Session("Segce") = "N"
	End If


	If Session("Segce") = "S" And session("Nomodificar") <> "S" Then
		session("Nomodificarce") = "N"
	Else
		session("Nomodificarce") = "S"
	End If

	If Session("Segct") <> "S" And Session("Segmo") <> "S" Then
		session("Nomodificar") = "S"
	End If

	BCOPYSPC = "N"
	'RESPONSE.WRITE SESSION("WMATRMEMP")  & "    "  & SESSION("WMATRMEMPACT")
	' SESSION("WMATRMEMPACT") = "S"
	If   SESSION("WMATRMEMP") = "S" And  SESSION("WMATRMEMPACT") = "N" Then


		bAgregar = 0
		bSuprimir = 0
		session("Nomodificar") = "S"
		session("Nomodificarce") = "S"

		If bModificar And Session("Segct") = "S" Then
			BCOPYSPC = "S"
		End If

		bModificar = 0

		Session("Segct") = "N"
		Session("Segmo") = "N"
		Session("Segce") = "N"

	End If
       ' sHTML = sHTML & "<center>"



        sHTML = sHTML & "<div style='float:left;margin:30px;width:350px;height:350px;background-color: rgba(140,170,200,0.6);border-radius: 60%;box-shadow: 0 6px 10px rgba(80, 80, 80, 0.8);' >"
		sHTML = sHTML & "<table  width=70% height=70% valign='middle' cellspacing=0 cellpadding=1 style='xdisplay:none;position:relative;top:30px;'  >"

	    sHTML = sHTML & "<tr >" & "<td rowspan=8 >"
	    sHTML = sHTML & "<button style='width:150px;height:150px;border-radius: 60%;color:white;background: linear-gradient(#cccccc, #f78822);box-shadow: 0 3px 3px rgba(0, 0, 0, 0.8);border-style:solid;border-width:1px;border-color:grey;'  >" & sNombredOC & "</button>"
	    sHTML = sHTML & "</td >"
	    sHTML = sHTML & "<td class=vt onclick=""document.getElementById('DXM').style.display= 'none'; "" >" & "X"  & "</td >" & "</tr >"

	If smodal <> "PROCESOS" Then


		sHTMLa = sHTMLa & "<div class='selector' style='border-style:solid;border-width:2;border-color:#000000;' >"
		sHTMLa = sHTMLa & "  <ul>"



		aspeditar   = CStr("Editar.asp?Codigo=" & nCodobjeto & "&Tipo=" & sTipobjeto & "&Perfil=" & sPerfil  & "&tipocpt=" &   TIPOCPT)
		aspeliminar = CStr("asp/EliminarEsquema.asp?Codobjeto=" & nCodobjeto & "&Tipobjeto=" & sTipobjeto & "&Eliminar=NSNC")
		aspcopiar   = CStr("Copiar.asp?Codigo=" & nCodobjeto & "&Tipo=" & sTipobjeto & "&Perfil=" & sPerfil)

''''alert('va');document.getElementById('DXMARCO').style.display= 'block';

		If bConsultar And Session("Segco") = "S" Then
                    sHTML = sHTML & "<tr ><td>" & "<button class='botO2' onclick=""document.SubFormED.submit();"" >Editar</button>"  & "</td></tr >" '& Editar(CStr(aspeditar))

		    FormFF = ofv.GenForm("SubFormED", "", "Editar.asp?Codigo=" & nCodobjeto & "&Tipo=" & sTipobjeto & "&Perfil=" & sPerfil  & "&tipocpt=" &   TIPOCPT, "_new", "post", "", "", "si")
                    inder = inder + 1

	            sHTMLa = sHTMLa & "    <li>"
		    sHTMLa = sHTMLa & "      <input id='c" & inder & "' type='checkbox'>"
		    sHTMLa = sHTMLa & "      <label for='c" & inder & "'>Editar</label>"
		    sHTMLa = sHTMLa & "    </li>"
                end if

		If bSuprimir And Session("Segct") = "S" And TIPOPEDIDO <> "PD" And TIPOPEDIDO <> "SN" Then
                  ' sHTML = sHTML & "<tr >" & Eliminar(CStr(aspeliminar)) & "</tr >"
                    sHTML = sHTML & "<tr ><td>" & "<button class='botO2'  >Eliminar</button>"  & "</td></tr >"

                    inder = inder + 1

	            sHTMLa = sHTMLa & "    <li>"
		    sHTMLa = sHTMLa & "      <input id='c" & inder & "' type='checkbox'>"
		    sHTMLa = sHTMLa & "      <label for='c" & inder & "'>Eliminar</label>"
		    sHTMLa = sHTMLa & "    </li>"
                end if


		If ((bModificar And Session("Segct") = "S") Or (BCOPYSPC = "S")) And TIPOPEDIDO <> "PD" And TIPOPEDIDO <> "SN" Then

		'''	sHTML = sHTML & "<tr >" & Copiar(CStr(aspcopiar)) & "</tr >"
            sHTML = sHTML & "<tr ><td>" & "<button class='botO2'  >Copiar</button>"  & "</td></tr >"

                    inder = inder + 1

	            sHTMLa = sHTMLa & "    <li>"
		    sHTMLa = sHTMLa & "      <input id='c" & inder & "' type='checkbox'>"
		    sHTMLa = sHTMLa & "      <label for='c" & inder & "'>Copiar</label>"
		    sHTMLa = sHTMLa & "    </li>"

		End If

          sHTML = sHTML & "<tr ><td>" & "<button id='BTADD' class='botO2' style='display:none;' onclick=""document.getElementById('DXTADD').style.display= 'inline-block';"" >Agregar</button>"  & "</td></tr >"

	'	sHTML = sHTML & "</tr>"


'		sHTML = sHTML & "</tr></table>"

	End If

End If

	'	sHTML = sHTML & "<tr>"

'sHTML = sHTML & "<table width=100% cellspacing=0 cellpadding=4  style='xdisplay:none;' ><tr >"

If smodal <> "PROCESOS" Then

	If bModificar And TIPOPEDIDO <> "PD" And TIPOPEDIDO <> "SN"  Then
		xtipo = "ESQUEMAS"
		SNOMDOC = STIPODOC & ":" & SNOMBREDOC
		SNOMDOC = Replace(SNOMDOC," ","_")
		asporden    = CStr("edordenar.asp?Codigo=" & nCodobjeto & "&Tipo=" & sTipobjeto & "&XTipo=" & xtipo  & "&Perfil=" & sPerfil & "&NOMDOC=" & SNOMDOC)
	'	sHTML = sHTML & "<tr >" & Ordenar(CStr(asporden)) & "</tr >"
        sHTML = sHTML & "<tr ><td>" & "<button class='botO2'  >Ordenar</button>"  & "</td></tr >"
                    inder = inder + 1

	            sHTMLa = sHTMLa & "    <li>"
		    sHTMLa = sHTMLa & "      <input id='c" & inder & "' type='checkbox'>"
		    sHTMLa = sHTMLa & "      <label for='c" & inder & "'>Ordenar</label>"
		    sHTMLa = sHTMLa & "    </li>"

	End If

	If ((bModificar And Session("Segct") = "S") Or (BCOPYSPC = "S")) And TIPOPEDIDO <> "PD" And TIPOPEDIDO <> "SN" Then
		If IsNumeric(sEstado) Then
			If sEstado <> 4 Or session("APLMDCLI")="CFA" Or  session("APLMDCLI")="BC" Then
				aspmover  =   CStr("Copiar.asp?EsqSup="&sEsqSup&"&MOVERESQUEMA=SI&Codigo=" & nCodobjeto & "&Tipo=" & sTipobjeto & "&Perfil=" & sPerfil)
				'sHTML = sHTML & "<tr >" & MOVER(CStr(aspmover)) & "</tr >"
                sHTML = sHTML & "<tr ><td>" & "<button class='botO2'  >Mover</button>"  & "</td></tr >"
                    inder = inder + 1

	            sHTMLa = sHTMLa & "    <li>"
		    sHTMLa = sHTMLa & "      <input id='c" & inder & "' type='checkbox'>"
		    sHTMLa = sHTMLa & "      <label for='c" & inder & "'>MOVER</label>"
		    sHTMLa = sHTMLa & "    </li>"

			End If
		End If
	End If


Else

	If bModificar Then
		xtipo = "ESQUEMAS"
		SNOMDOC = STIPODOC & ":" & SNOMBREDOC
		SNOMDOC = Replace(SNOMDOC," ","_")
		aspprocesos    = CStr("edprocesos.asp?codobjeto=" & nCodobjeto & "&Tipobjeto=" & sTipobjeto)
	'''	sHTML = sHTML & "<tr >" & Aprocesos(CStr(aspprocesos)) & "</tr >"
        sHTML = sHTML & "<tr ><td>" & "<button class='botO2'  >Procesar</button>"  & "</td></tr >"
                    inder = inder + 1

	            sHTMLa = sHTMLa & "    <li>"
		    sHTMLa = sHTMLa & "      <input id='c" & inder & "' type='checkbox'>"
		    sHTMLa = sHTMLa & "      <label for='c" & inder & "'>Procesos</label>"
		    sHTMLa = sHTMLa & "    </li>"

	End If







End If




sHTML = sHTML & "</table></div>" & FormFF





'sHTML = sHTML & "</center>"

'response.write "<br>niveles=" & Niveles(1)
If Niveles(1) <> ""  And TIPOPEDIDO <> "PD" And TIPOPEDIDO <> "SN"  Then
	If bAgregar Then
		If SESSION("BASE")="K" Then
			Strsql = "Select * From Tipobjeto Where TIPOCPT='K' AND (aplic = 'GLB' or aplic = '" & wkaplic & "' "
		Else
			Strsql = "Select * From Tipobjeto Where TIPOCPT <> 'K' AND (aplic = 'GLB' or aplic = '" & wkaplic & "' "
		End If
		Strsql = Strsql & ") and nivel IN("
		For x = 1 To UBound(Niveles)
			Strsql = Strsql & "'" & Niveles(x) & "',"
		Next
		Strsql = Left(Strsql,Len(Strsql)-1) & ")"

		Set rs = ofv.crearconsultaEx(StrSql,cn,1,parametros)


		If ofv.ModTitCls = "1" Or ofv.ModTitCls = "0" Then
			xGenConta = ofv.GetTit(CLng(ofv.ModTitCls) + 5)
			xGenConta = ofv.FTitDat1

			xGenCont = xGenConta

			xGenConti = "agregar"

			xGenCont = Replace(xGenCont, "XXXDATOXXX", xGenConti)
			xGenCont = Replace(xGenCont, "XXXWIDTHXXX", "98%")
			xGenCont = Replace(xGenCont, "XXXIDXXX", "DBML1R")

			sInput = xGenCont

			xsclase = "xts"

		Else

			sInput = "agregar"

			xsclase = "ttr"

		End If



		sHTML = sHTML & "<table  width=100% style='display:none;' ><tr>"
		sHTML = sHTML & OFV.gencelda("", xsclase, "", "", "", "", "", sInput, "si")
		sHTML = sHTML & "</tr></table>"

		sHTML = sHTML & "<script language='javascript'> document.getElementById('DXTADD').style.display= 'block'; </script>"
		sHTML = sHTML & "<div id='DXTADD'  style='display:none;margin:30px;width:350px;height:250px;background-color: rgba(140,170,200,0.6);border-radius: 50px 50px 50px 50px ;box-shadow: 0 6px 10px rgba(80, 80, 80, 0.8);position:relative;left:-100px;top:50px;' ><span id=111>"

		sHTML = sHTML & "<div   style='overflow:auto;margin:30px;width:300px;height:200px;' >"


		sHTML = sHTML & "<table  valign='middle' cellspacing=2 cellpadding=2 style='xdisplay:none;' >"

		Do Until rs.eof


			If  RS(21) > "C" Then
				sxiconv = "<font  face='CorFlatIcon' size=4 color='#993333'>&#0158;</font>"
			Else
				sxiconv = "<font  face='CorFlatIcon' size=4 color='#336699'>&#0204;</font>"
			End If


			sHTML = sHTML & "<tr>"
			sHTML = sHTML & "<td id='TD" & rs(1) & "'   "
			sHTML = sHTML & "  "
			sHTML = sHTML & " onclick=" & Chr(34)
			sHTML = sHTML & "AbrirAgregar('" & "Agregar.asp?Tipo=" & rs(0)
			sHTML = sHTML & "&Nombre=" & ofv.convertircar(rs(1)," ","_") & "&Elemento="
			sHTML = sHTML & sElemento & "');" & Chr(34)
			sHTML = sHTML & " >" & "<button class='botA2'  >" & rs(1) & "</button>"  &  Chr(13)
			sHTML = sHTML & "</td></tr>"


			rs.movenext
		Loop
		sHTML = sHTML & "</table>"
		sHTML = sHTML & "</span></div>"
		sHTML = sHTML & "<script language='javascript'> document.getElementById('BTADD').style.display= 'block'; </script>"
		sHTML = sHTML & "<hr size=1 color=#000099 style='display:none;' >"
		Call ofv.cerrarconsulta(rs)
	End If
End If


	            sHTMLa = sHTMLa & "  </ul>"
	            sHTMLa = sHTMLa & "  <button>" & sNombredOC & "</button>"
	            sHTMLa = sHTMLa & "</div>"


'sHTML = sHTML & sHTMLa

Call ofv.cerrarconn(cn)

        wk1 = ""

        wk1 = wk1 & "<script src=" & chr(34) & "http://ajax.googleapis.com/ajax/libs/jquery/1.11.1/jquery.min.js" & chr(34) & " ></script>"
        wk1 = wk1 & "<script>"
        wk1 = wk1 & "var nbOptions = " & inder & ";"
        wk1 = wk1 & "var angleStart = -360;"

        wk1 = wk1 & "// jquery rotate animation"
        wk1 = wk1 & "function rotate(li,d) {"
        wk1 = wk1 & "    $({d:angleStart}).animate({d:d}, {"
        wk1 = wk1 & "        step: function(now) {"
        wk1 = wk1 & "            $(li)"
        wk1 = wk1 & "               .css({ transform: 'rotate('+now+'deg)' })"
        wk1 = wk1 & "               .find('label')"
        wk1 = wk1 & "                  .css({ transform: 'rotate('+(-now)+'deg)' });"
        wk1 = wk1 & "        }, duration: 0"
        wk1 = wk1 & "    });"
        wk1 = wk1 & "}"

        wk1 = wk1 & "// show / hide the options"
        wk1 = wk1 & "function toggleOptions(s) {"
        wk1 = wk1 & "    $(s).toggleClass('open');"
        wk1 = wk1 & "    var li = $(s).find('li');"
        wk1 = wk1 & "    var deg = $(s).hasClass('half') ? 180/(li.length-1) : 360/li.length;"
        wk1 = wk1 & "    for(var i=0; i<li.length; i++) {"
        wk1 = wk1 & "        var d = $(s).hasClass('half') ? (i*deg)-90 : i*deg;"
        wk1 = wk1 & "        $(s).hasClass('open') ? rotate(li[i],d) : rotate(li[i],angleStart);"
        wk1 = wk1 & "    }"
        wk1 = wk1 & "}"

        wk1 = wk1 & "$('.selector button').click(function(e) {"
        wk1 = wk1 & "    toggleOptions($(this).parent());"
        wk1 = wk1 & "});"

        wk1 = wk1 & "setTimeout(function() { toggleOptions('.selector'); }, 100);//@ sourceURL=pen.js"
        wk1 = wk1 & "</script>"
        wk1 = wk1 & "<script type='text/javascript'>"

        wk1 = wk1 & "  var _gaq = _gaq || [];"
        wk1 = wk1 & "  _gaq.push(['_setAccount', 'UA-36251023-1']);"
        wk1 = wk1 & "  _gaq.push(['_setDomainName', 'jqueryscript.net']);"
        wk1 = wk1 & "  _gaq.push(['_trackPageview']);"
        wk1 = wk1 & ""
        wk1 = wk1 & "  (function() {"
        wk1 = wk1 & "    var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;"
        wk1 = wk1 & "    ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';"
        wk1 = wk1 & "    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);"
        wk1 = wk1 & "  })();"

        wk1 = wk1 & "</script>"



	sHTML = sHTML & wk1

sHTML = sHTML & "</BODY></html>"

response.write sHTML

%>