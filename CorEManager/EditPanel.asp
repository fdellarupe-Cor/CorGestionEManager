<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<%

ok=ofv.CheckUsuario()
If ok <> "Y" Then
	uv = ofv.uv
	If uv = "NO" Then ok = ofv.ValidarUsuario
	sHTML = ok
Else

	Response.ContentType = "text/html"
	Response.AddHeader "Content-Type", "text/html;charset=ISO-8859-1"
Response.CharSet = "ISO-8859-1"

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


'	FUNCION = "  onload="  & Chr(34)

	If  recarga = "S" And sTipo <> "Consulta" Then
'		FUNCION = FUNCION & "document.recarga.submit();"
	End If

'	FUNCION = FUNCION & "rszdx4();"  & Chr(34)
	FUNCION = FUNCION & " onmousedown = " & Chr(34) & "Abre();" & Chr(34) & " "
'	FUNCION = FUNCION & " onresize = " & Chr(34) & "rszdx4();" & Chr(34) & " "

	'''sHTML = ofv.MENUheader("",FUNCION)

	sHTML = ""



        wkpanel = ofv.FormHeader("Edicion")



		wkpanel = wkpanel & "<html><head>"

		wkpanel = wkpanel & "<script type='text/javascript'  src='../coress20/incfile/func.js'></script>"


		wkpanel = wkpanel & "</head>"

        wkpanel = replace(wkpanel,"<body ", "<body " & FUNCION )

	sHTML = sHTML & wkpanel

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
        wk1 = wk1 & "  background: linear-gradient(#eeeeee, #cccccc);"
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

        wk1 = wk1 & "</STYLE>"

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
        wk1 = wk1 & " (function() {"
        wk1 = wk1 & "    var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;"
        wk1 = wk1 & "    ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';"
        wk1 = wk1 & "    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);"
        wk1 = wk1 & "  })();"


		'JQUERY
        wk1 = wk1 & "$(function(){ "

		'VER FILTROS
		wk1 = wk1 & "$('#verFilt, #verFilt2').click(function(){"
		wk1 = wk1 & "var idinv = $(this).attr('data-idinv');"
		wk1 = wk1 & "var tipo = $(this).attr('data-tipo');"
		wk1 = wk1 & "var caso = $(this).attr('data-caso');"
		wk1 = wk1 & "var pos = $(this).attr('data-pos');"
		wk1 = wk1 & "levantaPanelFiltros(idinv, tipo, caso, pos);"
		wk1 = wk1 & "});"

		wk1 = wk1 & "$('.cargados').sortable({"
		wk1 = wk1 & "connectWith: '#eliminar',"
		wk1 = wk1 & "containment: '#tabCons'"
		wk1 = wk1 & "});"

		wk1 = wk1 & "$('#busc').click(function(){"
		wk1 = wk1 & "consultar();"
		wk1 = wk1 & "});"

		wk1 = wk1 & "$('#selecinv').click(function(){"
		wk1 = wk1 & "selecInventario();"
		wk1 = wk1 & "});"

        wk1 = wk1 & "});"

		wk1 = wk1 & "</script>"

	sHTML = sHTML & wk1

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
	sHTML = sHTML & "  var dato = 'EditPanelDiaL.asp?"

	If sModal = "PROCESOS" Then
		sHTML = sHTML & "MODAL=PROCESOS&"
	Else
		sHTML = sHTML & "MODAL=NORMAL&"
	End If


	sHTML = sHTML & "ID=' + Elemento ;" & Chr(13)
	sHTML = sHTML & " window.showModalDialog("
	sHTML = sHTML & "dato,window,'dialogWidth:280px;dialogHeight:500px;help:no;resizable:yes;dialogTop:250px;dialogLeft:300px;status:no');" & Chr(13)
	sHTML = sHTML & " }} " & Chr(13)

	sHTML = sHTML & "</SCRIPT> " & Chr(13)
	'------------------------------------------------------

	shtml = shtml & "<script language=javascript>" & Chr(13)
	sHTML = sHTML & "   function Editarx(Elemento) { " & Chr(13)
	'sHTML = sHTML & " alert(Elemento);" & Chr(13)
	sHTML = sHTML & " if (Elemento=='') { " & Chr(13)
	sHTML = sHTML & " } " & Chr(13)
	sHTML = sHTML & " else { " & Chr(13)
	sHTML = sHTML & " var dato = 'EditPanelDiaL.asp?"

	If sModal = "PROCESOS" Then
		sHTML = sHTML & "MODAL=PROCESOS&"
	Else
		sHTML = sHTML & "MODAL=NORMAL&"
	End If


	sHTML = sHTML & "ID=' + Elemento; " & Chr(13)

	''sHTML = sHTML & " alert( dato);" & Chr(13)
'DXM
'document.SubForm.submit();

	sHTML = sHTML & " document.getElementById('SubForm').action= dato; "
	sHTML = sHTML & " document.getElementById('SubForm').submit(); "
	sHTML = sHTML & " document.getElementById('DXM').style.display= 'block'; "

	sHTML = sHTML & " } } " & Chr(13)

	sHTML = sHTML & "   function EditarxZ(Elemento) { " & Chr(13)
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

		sHTML = sHTML & " window.showModalDialog("
		sHTML = sHTML & "dato,window,'dialogWidth:280px;dialogHeight:500px;help:no;resizable:yes;dialogTop:250px;dialogLeft:300px;status:no'); "  & Chr(13)

	Else

		sHTML = sHTML & " window.open("
		sHTML = sHTML & "dato,window,'Width=280,Height=320,help:no;resizable:yes;dialogTop:250px;dialogLeft:300px;status:no'); "  & Chr(13)

	End If

	sHTML = sHTML & " } } " & Chr(13)

	sHTML = sHTML & "</SCRIPT> " & Chr(13)

	'''''sHTML = sHTML & Editar("edit")


	shtml=shtml & "<script language=javascript>" & Chr(13)

	shtml=shtml & "function mostrarmenuD(){" & Chr(13)
	shtml=shtml & "document.getElementById('Dx1').style.display='';" & Chr(13)
	shtml=shtml & "}" & Chr(13)

	shtml=shtml & "function rszdx4(){" & Chr(13)
	shtml=shtml & "document.getElementById('DX2').style.display='';" & Chr(13)
	shtml=shtml & "document.getElementById('DX4').style.display='';" & Chr(13)
	shtml=shtml & "var spread = document.getElementById('DX4').offsetTop + 150;" & Chr(13)
	shtml=shtml & "document.getElementById('DX4').style.left=document.getElementById('DX2').offsetLeft + document.getElementById('DX2').clientWidth + 1;" & Chr(13)
	shtml=shtml & "document.getElementById('DX2').style.height=document.body.clientHeight - spread;" & Chr(13)
	shtml=shtml & "document.getElementById('DX4').style.height=document.body.clientHeight - spread;" & Chr(13)
	'  shtml=shtml & "alert(document.body.clientHeight + ' ' + document.all.DX4.offsetTop + ' ' + window.screen.availHeight + ' ' + document.all.DX4.style.height);" & chr(13)
	shtml=shtml & "}" & Chr(13)
	shtml=shtml & "</script>" & Chr(13)

	''''sHTML = sHTML & ofv.CorFormHeader



	sNombre = ofv.convertircar(sNombre," ","_")

	HrfAsp = "EditPanelDocs.asp?"
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
	''response.write strsql & "<br>"
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

			Strsql = "Select * From ATRMTEMAST WHERE ATRMCODST = " & rs(0) & " AND ATRMTIPOST = 'T' AND ATRMCODES = " & session("UN")
			''response.write strsql & "<br>"
			Set rs1 = ofv.crearconsultaEx(StrSql,cn,1,parametros)
			If Not rs1.eof Then
				If Not RS1(4) Then  WTVER = 0
				' response.write strsql & " " & WTVER & "<br>"
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

		sxlink = ""
		sxlink = sxlink & "EditPanelDocs.asp?Tipo=TEMAS&ID=0"
		sxlink = sxlink & "&sitem=0&CANTREGMOVER=0&stitem=0&sticono=0&stnivel=0&stesq=0"

                TARGET = "fesquemas"


		FormFF = ofv.GenForm("SubForm", "", sxlink, TARGET, "post", "", "", "si")
                scp = "<script language='javascript'> document.SubForm.submit(); </script>"
	Celda = Celda & "<div id='verFilt2' data-idinv='100' data-tipo='F' data-caso='1' data-pos='0' data-base='D' data-pantalla='editdoc' class='butt' style='border:none;height:18px;margin-right:5px;width:100px;float:left;margin-left:2px;'>FILTRAR&nbsp;&nbsp;<span class='la la-external-link-square'></span></div>"
	Celda = Celda & "<table id='tabCons' style='width:90%;display:none;' title='Panel para filtrar documentos' xborder>"
	Celda = Celda & "<tr>"
	Celda = Celda & "<td colspan=3 valign=middle>"
	Celda = Celda & "<button id='verFilt' data-idinv='100' data-tipo='F' data-caso='1' data-pos='0' data-base='D' data-pantalla='editdoc' class='butt' style='border:none;height:30px;margin-right:5px;'>DOCUMENTOS</button>"
	Celda = Celda & "<button id='selecinv' class='butt' style='height:30px;border:none;margin-right:5px;' data-idpanel='0' data-caso='0' data-tipo='I'><span class='la la-refresh' style='font-size:15px;'></span></button>"
	Celda = Celda & "<button id='busc' class='butt pCons' style='border:none;height:30px;display:none;'>FILTRAR</button>"
	Celda = Celda & "</td>"
	Celda = Celda & "</tr>"
	Celda = Celda & "<tr>"
	Celda = Celda & "<td id='eliminar' style='width:10%;border-radius:15px 15px 15px 15px;' align=center>"
	Celda = Celda & "<span id='imgElim' class='la la-trash' style='font-size:30pt;color:#FF3636;display:none;' title='Eliminar filtros'></span>"
	Celda = Celda & "</td>"
	Celda = Celda & "<td style='width:35%;' valign=top>"
	Celda = Celda & "<div id='ulcont' style='width:100%;'>"
	Celda = Celda & "<ul class='cargados' data-idinv='104'></ul></td>"
	Celda = Celda & "</div>"
	Celda = Celda & "<td id='contPan' style='width:45%;' valign=top></td>"
	Celda = Celda & "</tr>"
	Celda = Celda & "</table>"


	Celda = Celda & "<div id='DX4' style='margin-top:1px;margin-left:2px;margin-bottom:5px;xdisplay:none;"
	Celda = Celda & "width:100%;height:600px;overflow:auto;border-style:none;border-width:1px;border-color:#dddddd;"
	'Celda = Celda & "background-color:" & session("APLBGD") & ";"
	Celda = Celda & "' >"

	Celda = Celda & "&nbsp;"

	Celda = Celda & "</div><br>"

	Celda = Celda & "<div id='DX2' style='margin-top:1;margin-left:1;position:absolute;left:250px;top:200px;display:none;"
	Celda = Celda & "width:300px;height:300px;overflow:auto;border-style:solid;border-width:1px;border-color:#dddddd;"
	Celda = Celda & "background-color:" & session("APLBGD") & ";"
	Celda = Celda & "' >"

	Celda = Celda & TABLA

	Celda = Celda & "</div>"


	Celda = Celda & "<div id='DXM' style='margin-top:1;margin-left:1;position:absolute;left:350px;top:300px;display:none;"
	Celda = Celda & "width:900px;height:500px;overflow:auto;border-style:none;border-width:1px;border-color:#dddddd;"
	'''Celda = Celda & "background-color:" & session("APLBGD") & ";"
	Celda = Celda & "' >"

	Celda = Celda & "&nbsp;"

	Celda = Celda & "</div>"

	sHTML = sHTML & Celda & FormFF & scp




	Agregar = ""
	Buscar = ""
	CELDA = ""

	shtml = shtml & ofv.botoneraESSX("0", "", "", "", "", "", "", "", "", "no", Agregar, Buscar, CELDA)

	   		    	   tabla1F = ""
	   		    	   tabla1F = tabla1F & "<iframe id='DXMARCO' name='DXMARCO' frameborder='0' scrolling='auto'  "
	   		    	   tabla1F = tabla1F & "allowtransparency='true'  "
	   		    	   tabla1F = tabla1F & "src='' "
	   		    	   tabla1F = tabla1F & " style='position:absolute;left:350px;top:300px;display:none;"
	   		    	   tabla1F = tabla1F & "width:900px;height:500px;overflow:auto;border-style:none;border-width:1px;border-color:#dddddd;"
	   		    	   tabla1F = tabla1F & "'  ></iframe>"

	sHTML = sHTML & tabla1F

	sHTML = sHTML & "<form id='envia' method='post' style='display:none'>"
	sHTML = sHTML & "<input name=datoe id=datoe type=text value='' size=100>"
	sHTML = sHTML & "</form>"

	sHTML = sHTML & "<div id='resultado' style='display:none;overflow:hidden;width:550px;height:350px;' title='Consultado'></div>"
	sHTML = sHTML & "</body></html>"

End If

response.write  sHTML

%>

