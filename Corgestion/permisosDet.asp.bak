<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<%
Response.Write headhtml()
html = html & "<body>"

Set Cnn = Server.CreateObject("ADODB.Connection")
On Error Resume Next
Cnn.open scon1
If Err.Number <> 0 Then
	Response.Write "<div style='color:red;font-family:monospace;padding:10px;border:1px solid red;margin:10px;'>"
	Response.Write "<b>ERROR DE CONEXION A BASE DE DATOS</b><br>"
	Response.Write "Codigo: " & Err.Number & "<br>"
	Response.Write "Detalle: " & Server.HTMLEncode(Err.Description) & "</div>"
	Response.End
End If
On Error Goto 0

xmenu = request.querystring("xmenu")
xtmenu = request.querystring("xtmenu")
idusr = request.querystring("idusr")
tipousr = request.querystring("tipousr")
opc = request.querystring("opc")
idmenu = request.querystring("idmenu")
'sscr = request.querystring("ssc")
sTipoMenu = request.querystring("tipomenu")

if opc = "actEst" then

	sEstado = request.form("estado")
	if sEstado = 1 then
		sEstado = 0
	elseif sEstado = 0 then
		sEstado = 1
	end if

	ExecNonQuery Cnn, "sp_cg_MenuUsuario @Modo='actualizar_flag', @Flag='estado', @Valor=" & sEstado & ", @Id=" & idmenu


	html = html & "<input type='button' style='display:none;' onclick='teste()' id='scr'>"

	html = html & "<script type='text/javascript'>"
	html = html & " function teste(){ "
	html = html & " scroll(0,250); "
	html = html & " } "
	html = html & " window.onload = function() { "
	html = html & " document.getElementById('scr').click(); "
	html = html & " } "
	html = html & "</script>"

end if

if opc = "actAgre" then

	sAgregar = request.form("agregar")
	if sAgregar = 1 then
		sAgregar = 0
	elseif sAgregar = 0 then
		sAgregar = 1
	end if

	ExecNonQuery Cnn, "sp_cg_MenuUsuario @Modo='actualizar_flag', @Flag='agregar', @Valor=" & sAgregar & ", @Id=" & idmenu



	html = html & "<input type='button' style='display:none;' onclick='teste()' id='scr'>"

	html = html & "<script type='text/javascript'>"
	html = html & " function teste(){ "
	html = html & " scroll(0,250); "
	html = html & " } "
	html = html & " window.onload = function() { "
	html = html & " document.getElementById('scr').click(); "
	html = html & " } "
	html = html & "</script>"

end if

if opc = "actModi" then

	sModificar = request.form("modificar")
	if sModificar = 1 then
		sModificar = 0
	elseif sModificar = 0 then
		sModificar = 1
	end if

	ExecNonQuery Cnn, "sp_cg_MenuUsuario @Modo='actualizar_flag', @Flag='modificar', @Valor=" & sModificar & ", @Id=" & idmenu

	html = html & "<input type='button' style='display:none;' onclick='teste()' id='scr'>"

	html = html & "<script type='text/javascript'>"
	html = html & " function teste(){ "
	html = html & " scroll(0,250); "
	html = html & " } "
	html = html & " window.onload = function() { "
	html = html & " document.getElementById('scr').click(); "
	html = html & " } "
	html = html & "</script>"

end if

if opc = "actElim" then

	sEliminar = request.form("eliminar")
	if sEliminar = 1 then
		sEliminar = 0
	elseif sEliminar = 0 then
		sEliminar = 1
	end if


	ExecNonQuery Cnn, "sp_cg_MenuUsuario @Modo='actualizar_flag', @Flag='eliminar', @Valor=" & sEliminar & ", @Id=" & idmenu


	html = html & "<input type='button' style='display:none;' onclick='teste()' id='scr'>"

	html = html & "<script type='text/javascript'>"
	html = html & " function teste(){ "
	html = html & " scroll(0,250); "
	html = html & " } "
	html = html & " window.onload = function() { "
	html = html & " document.getElementById('scr').click(); "
	html = html & " } "
	html = html & "</script>"

end if


if opc = "actCons" then

	sConsultar = request.form("consultar")
	if sConsultar = 1 then
		sConsultar = 0
	elseif sConsultar = 0 then
		sConsultar = 1
	end if

	ExecNonQuery Cnn, "sp_cg_MenuUsuario @Modo='actualizar_flag', @Flag='consultar', @Valor=" & sConsultar & ", @Id=" & idmenu


	html = html & "<input type='button' style='display:none;' onclick='teste()' id='scr'>"

	html = html & "<script type='text/javascript'>"
	html = html & " function teste(){ "
	html = html & " scroll(0,250); "
	html = html & " } "
	html = html & " window.onload = function() { "
	html = html & " document.getElementById('scr').click(); "
	html = html & " } "
	html = html & "</script>"

end if



html = html & "<div>"

html = html & "<header id='header' style='background:#0e1949;padding:1rem 0;'>"

html = html & "<div >"

html = html & "<table style='width:100%;'><tr style='display:flex;padding: 0 1rem;'>"

html = html & "<td style='width:75%;' valign=center><a href='index.asp' class='logo' >"
html = html & "<img src='images/iso_cor.png' style='width:6rem;' >"
html = html & "<font style='font-size:15pt;color:white;'><b>Cor Gestion</b> | Sistema Administrativo</font></a></td>"
html = html & "<td style='width:25%;display:flex;align-items:center;flex-direction:row-reverse;' align='right'>"
html = html & "<nav>"
html = html & "<ul>"
'html = html & "<li><a href='#menu'></a></li>"
html = html & "</ul>"
html = html & "</nav>"
html = html & "</td>"

html = html & "</tr></table>"

html = html & "</div>"

html = html & "</header>"

html = html & "<table class='titulo-seccion' >"
html = html & "<tr style='display:flex;align-items:center;justify-content:space-between;'>"

html = html & "<td><h2>SEGURIDAD</h2></td>"

form5 = "<form id='form5' style='height:70%;' method='post' action='permisos.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
boton5 = "<input class='bot' type='button' value='VOLVER' onclick=""document.getElementById('form5').submit();"">"

html = html & "<td align=right>" & boton5 & "</td>"

html = html & "</tr>"

html = html & "</table>"

html = html & "</div>"

html = html & "<div style='height:5%;width:100%;'>"

html = html & "</div>"

html = html & "<div style='height:57%;width:100%;float:left;'>"

html = html & "<table class='tabl' style='height:80%;width:55%;color:#000000;' cellspacing=3 cellpadding=3 align=center>"

sUsr = getName(cnn,"","cor_responsables",idusr)

html = html & "<tr>"
html = html & "<td colspan=6 style='font-size:15pt;' align=center><b>PERFIL&nbsp;" & sUsr & "</b></td>"
html = html & "</tr>"

sTit = ""

if sTipoMenu = "INICIO" then
	sTit = "MENU PRINCIPAL"
elseif sTipoMenu = "PARAM" then
	sTit = "PARAMETROS"
end if

html = html & "<tr>"
html = html & "<td colspan=6><u>" & sTit & "</u></td>"
html = html & "</tr>"

html = html & "<tr>"
html = html & "<td colspan=6>&nbsp;</td>"
html = html & "</tr>"

html = html & "<tr style='background-color:#A4A5B1;'>"
html = html & "<td style='width:15%;' align=center>*</td>"
html = html & "<td style='width:8%;' align=center>ESTADO</td>"
html = html & "<td style='width:8%;' align=center>AGREGAR</td>"
html = html & "<td style='width:8%;' align=center>MODIFICAR</td>"
html = html & "<td style='width:8%;' align=center>ELIMINAR</td>"
html = html & "<td style='width:8%;' align=center>CONSULTAR</td>"
html = html & "</tr>"


Set rs = Exec(Cnn, "sp_cg_MenuUsuario @Modo='listar', @IdUsr=" & SqlStr(idusr) & ", @FuncUsr='" & tipousr & "', @TipoMenu='" & sTipoMenu & "'")


sColor = "#EAEBF2;"

Do While Not rs.eof

	'estado
	if rs(1) = 1 then
		sImgEst = "<img src='images/circuloV.jpg' alt='estado' width='23' height='23' onclick=""document.getElementById('xFormEst" & rs(6) & "').submit();"" />"
	elseif rs(1) = 0 then
		sImgEst = "<img src='images/circuloR.jpg' alt='estado' width='20' height='20' onclick=""document.getElementById('xFormEst" & rs(6) & "').submit();"" />"
	else
		sImgEst = ""
	end if
	xLinkEst = "<form id='xFormEst" & rs(6) & "' name='xFormEst" & rs(6) & "' method='post' action='permisosDet.asp?opc=actEst&idmenu=" & rs(6) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&idusr=" & Server.URLEncode(idusr & "") & "&tipousr=" & tipousr & "&ssc=par&tipomenu=" & sTipoMenu & "'>"
	xLinkEst = xLinkEst & "<input type='hidden' id='estado' name='estado' value='" & rs(1) & "' >" & sImgEst
	xLinkEst = xLinkEst & "</form>"

	'agregar
	if rs(2) = 1 then
		sImgAgr = "<img src='images/circuloV.jpg' alt='agregar' width='23' height='23' onclick=""document.getElementById('xFormAgre" & rs(6) & "').submit();"" />"
	elseif rs(2) = 0 then
		sImgAgr = "<img src='images/circuloR.jpg' alt='agregar' width='20' height='20' onclick=""document.getElementById('xFormAgre" & rs(6) & "').submit();"" />"
	else
		sImgAgr = ""
	end if
	xLinkAgr = "<form id='xFormAgre" & rs(6) & "' name='xFormAgre" & rs(6) & "' method='post' action='permisosDet.asp?opc=actAgre&idmenu=" & rs(6) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&idusr=" & Server.URLEncode(idusr & "") & "&tipousr=" & tipousr & "&ssc=par&tipomenu=" & sTipoMenu & "'>"
	xLinkAgr = xLinkAgr & "<input type='hidden' id='agregar' name='agregar' value='" & rs(2) & "' >" & sImgAgr
	xLinkAgr = xLinkAgr & "</form>"

	'modificar
	if rs(3) = 1 then
		sImgMod = "<img src='images/circuloV.jpg' alt='modificar' width='23' height='23' onclick=""document.getElementById('xFormModi" & rs(6) & "').submit();"" />"
	elseif rs(3) = 0 then
		sImgMod = "<img src='images/circuloR.jpg' alt='modificar' width='20' height='20' onclick=""document.getElementById('xFormModi" & rs(6) & "').submit();"" />"
	else
		sImgMod = ""
	end if
	xLinkMod = "<form id='xFormModi" & rs(6) & "' name='xFormModi" & rs(6) & "' method='post' action='permisosDet.asp?opc=actModi&idmenu=" & rs(6) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&idusr=" & Server.URLEncode(idusr & "") & "&tipousr=" & tipousr & "&ssc=par&tipomenu=" & sTipoMenu & "'>"
	xLinkMod = xLinkMod & "<input type='hidden' id='modificar' name='modificar' value='" & rs(3) & "' >" & sImgMod
	xLinkMod = xLinkMod & "</form>"

	'eliminar
	if rs(4) = 1 then
		sImgElim = "<img src='images/circuloV.jpg' alt='eliminar' width='23' height='23' onclick=""document.getElementById('xFormElim" & rs(6) & "').submit();"" />"
	elseif rs(4) = 0 then
		sImgElim = "<img src='images/circuloR.jpg' alt='eliminar' width='20' height='20' onclick=""document.getElementById('xFormElim" & rs(6) & "').submit();"" />"
	else
		sImgElim = ""
	end if
	xLinkElim = "<form id='xFormElim" & rs(6) & "' name='xFormElim" & rs(6) & "' method='post' action='permisosDet.asp?opc=actElim&idmenu=" & rs(6) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&idusr=" & Server.URLEncode(idusr & "") & "&tipousr=" & tipousr & "&ssc=par&tipomenu=" & sTipoMenu & "'>"
	xLinkElim = xLinkElim & "<input type='hidden' id='eliminar' name='eliminar' value='" & rs(4) & "' >" & sImgElim
	xLinkElim = xLinkElim & "</form>"

	'consultar
	if rs(5) = 1 then
		sImgCons = "<img src='images/circuloV.jpg' alt='consultar' width='23' height='23' onclick=""document.getElementById('xFormCons" & rs(6) & "').submit();"" />"
	elseif rs(5) = 0 then
		sImgCons = "<img src='images/circuloR.jpg' alt='consultar' width='20' height='20' onclick=""document.getElementById('xFormCons" & rs(6) & "').submit();"" />"
	else
		sImgCons = ""
	end if
	xLinkCons = "<form id='xFormCons" & rs(6) & "' name='xFormCons" & rs(6) & "' method='post' action='permisosDet.asp?opc=actCons&idmenu=" & rs(6) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&idusr=" & Server.URLEncode(idusr & "") & "&tipousr=" & tipousr & "&ssc=par&tipomenu=" & sTipoMenu & "'>"
	xLinkCons = xLinkCons & "<input type='hidden' id='consultar' name='consultar' value='" & rs(5) & "' >" & sImgCons
	xLinkCons = xLinkCons & "</form>"

	html = html & "<tr style='background-color:" & sColor & "'>"
	html = html & "<td align=center >" & UCASE(rs(0)) & "</td>"
	html = html & "<td align=center style='padding-top:30px;'>" & xLinkEst & "</td>"
	html = html & "<td align=center >" & xLinkAgr & "</td>"
	html = html & "<td align=center >" & xLinkMod & "</td>"
	html = html & "<td align=center >" & xLinkElim & "</td>"
	html = html & "<td align=center >" & xLinkCons & "</td>"
	html = html & "</tr>"

	if sColor = "#EAEBF2;" then
		sColor = "#D9D9E3;"
	elseif sColor = "#D9D9E3;" then
		sColor = "#EAEBF2;"
	end if

rs.movenext
Loop

html = html & "<tr>"
html = html & "<td colspan=6>&nbsp;</td>"
html = html & "</tr>"

html = html & "</table>"

html = html & "</div>"

html = html & "<script src='assets/js/jquery.min.js'></script>"
html = html & "<script src='assets/js/skel.min.js'></script>"
html = html & "<script src='assets/js/util.js'></script>"
html = html & "<script src='assets/js/main.js'></script>"

html = html & form5 & "</body>"
html = html & "</html>"

response.write html
%>