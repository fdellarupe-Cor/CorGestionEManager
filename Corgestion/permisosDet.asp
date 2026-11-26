<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/ui_componentes.inc" -->
<%
'[UI-ANT] ' ===== Cuerpo anterior completo: reemplazado por el bloque [UI] que sigue. Se borra entero buscando [UI-ANT] =====
'[UI-ANT] Response.Write headhtml()
'[UI-ANT] html = html & "<body>"

'[UI-ANT] Set Cnn = Server.CreateObject("ADODB.Connection")
'[UI-ANT] On Error Resume Next
'[UI-ANT] Cnn.open scon1
'[UI-ANT] If Err.Number <> 0 Then
	'[UI-ANT] Response.Write "<div style='color:red;font-family:monospace;padding:10px;border:1px solid red;margin:10px;'>"
	'[UI-ANT] Response.Write "<b>ERROR DE CONEXION A BASE DE DATOS</b><br>"
	'[UI-ANT] Response.Write "Codigo: " & Err.Number & "<br>"
	'[UI-ANT] Response.Write "Detalle: " & Server.HTMLEncode(Err.Description) & "</div>"
	'[UI-ANT] Response.End
'[UI-ANT] End If
'[UI-ANT] On Error Goto 0

'[UI-ANT] xmenu = request.querystring("xmenu")
'[UI-ANT] xtmenu = request.querystring("xtmenu")
'[UI-ANT] idusr = request.querystring("idusr")
'[UI-ANT] tipousr = request.querystring("tipousr")
'[UI-ANT] opc = request.querystring("opc")
'[UI-ANT] idmenu = request.querystring("idmenu")
'[UI-ANT] 'sscr = request.querystring("ssc")
'[UI-ANT] sTipoMenu = request.querystring("tipomenu")

'[UI-ANT] if opc = "actEst" then

	'[UI-ANT] sEstado = request.form("estado")
	'[UI-ANT] if sEstado = 1 then
		'[UI-ANT] sEstado = 0
	'[UI-ANT] elseif sEstado = 0 then
		'[UI-ANT] sEstado = 1
	'[UI-ANT] end if

	'[UI-ANT] ExecNonQuery Cnn, "sp_cg_MenuUsuario @Modo='actualizar_flag', @Flag='estado', @Valor=" & sEstado & ", @Id=" & idmenu


	'[UI-ANT] html = html & "<input type='button' style='display:none;' onclick='teste()' id='scr'>"

	'[UI-ANT] html = html & "<script type='text/javascript'>"
	'[UI-ANT] html = html & " function teste(){ "
	'[UI-ANT] html = html & " scroll(0,250); "
	'[UI-ANT] html = html & " } "
	'[UI-ANT] html = html & " window.onload = function() { "
	'[UI-ANT] html = html & " document.getElementById('scr').click(); "
	'[UI-ANT] html = html & " } "
	'[UI-ANT] html = html & "</script>"

'[UI-ANT] end if

'[UI-ANT] if opc = "actAgre" then

	'[UI-ANT] sAgregar = request.form("agregar")
	'[UI-ANT] if sAgregar = 1 then
		'[UI-ANT] sAgregar = 0
	'[UI-ANT] elseif sAgregar = 0 then
		'[UI-ANT] sAgregar = 1
	'[UI-ANT] end if

	'[UI-ANT] ExecNonQuery Cnn, "sp_cg_MenuUsuario @Modo='actualizar_flag', @Flag='agregar', @Valor=" & sAgregar & ", @Id=" & idmenu



	'[UI-ANT] html = html & "<input type='button' style='display:none;' onclick='teste()' id='scr'>"

	'[UI-ANT] html = html & "<script type='text/javascript'>"
	'[UI-ANT] html = html & " function teste(){ "
	'[UI-ANT] html = html & " scroll(0,250); "
	'[UI-ANT] html = html & " } "
	'[UI-ANT] html = html & " window.onload = function() { "
	'[UI-ANT] html = html & " document.getElementById('scr').click(); "
	'[UI-ANT] html = html & " } "
	'[UI-ANT] html = html & "</script>"

'[UI-ANT] end if

'[UI-ANT] if opc = "actModi" then

	'[UI-ANT] sModificar = request.form("modificar")
	'[UI-ANT] if sModificar = 1 then
		'[UI-ANT] sModificar = 0
	'[UI-ANT] elseif sModificar = 0 then
		'[UI-ANT] sModificar = 1
	'[UI-ANT] end if

	'[UI-ANT] ExecNonQuery Cnn, "sp_cg_MenuUsuario @Modo='actualizar_flag', @Flag='modificar', @Valor=" & sModificar & ", @Id=" & idmenu

	'[UI-ANT] html = html & "<input type='button' style='display:none;' onclick='teste()' id='scr'>"

	'[UI-ANT] html = html & "<script type='text/javascript'>"
	'[UI-ANT] html = html & " function teste(){ "
	'[UI-ANT] html = html & " scroll(0,250); "
	'[UI-ANT] html = html & " } "
	'[UI-ANT] html = html & " window.onload = function() { "
	'[UI-ANT] html = html & " document.getElementById('scr').click(); "
	'[UI-ANT] html = html & " } "
	'[UI-ANT] html = html & "</script>"

'[UI-ANT] end if

'[UI-ANT] if opc = "actElim" then

	'[UI-ANT] sEliminar = request.form("eliminar")
	'[UI-ANT] if sEliminar = 1 then
		'[UI-ANT] sEliminar = 0
	'[UI-ANT] elseif sEliminar = 0 then
		'[UI-ANT] sEliminar = 1
	'[UI-ANT] end if


	'[UI-ANT] ExecNonQuery Cnn, "sp_cg_MenuUsuario @Modo='actualizar_flag', @Flag='eliminar', @Valor=" & sEliminar & ", @Id=" & idmenu


	'[UI-ANT] html = html & "<input type='button' style='display:none;' onclick='teste()' id='scr'>"

	'[UI-ANT] html = html & "<script type='text/javascript'>"
	'[UI-ANT] html = html & " function teste(){ "
	'[UI-ANT] html = html & " scroll(0,250); "
	'[UI-ANT] html = html & " } "
	'[UI-ANT] html = html & " window.onload = function() { "
	'[UI-ANT] html = html & " document.getElementById('scr').click(); "
	'[UI-ANT] html = html & " } "
	'[UI-ANT] html = html & "</script>"

'[UI-ANT] end if


'[UI-ANT] if opc = "actCons" then

	'[UI-ANT] sConsultar = request.form("consultar")
	'[UI-ANT] if sConsultar = 1 then
		'[UI-ANT] sConsultar = 0
	'[UI-ANT] elseif sConsultar = 0 then
		'[UI-ANT] sConsultar = 1
	'[UI-ANT] end if

	'[UI-ANT] ExecNonQuery Cnn, "sp_cg_MenuUsuario @Modo='actualizar_flag', @Flag='consultar', @Valor=" & sConsultar & ", @Id=" & idmenu


	'[UI-ANT] html = html & "<input type='button' style='display:none;' onclick='teste()' id='scr'>"

	'[UI-ANT] html = html & "<script type='text/javascript'>"
	'[UI-ANT] html = html & " function teste(){ "
	'[UI-ANT] html = html & " scroll(0,250); "
	'[UI-ANT] html = html & " } "
	'[UI-ANT] html = html & " window.onload = function() { "
	'[UI-ANT] html = html & " document.getElementById('scr').click(); "
	'[UI-ANT] html = html & " } "
	'[UI-ANT] html = html & "</script>"

'[UI-ANT] end if



'[UI-ANT] html = html & "<div>"

'[UI-ANT] html = html & "<header id='header' style='background:#0e1949;padding:1rem 0;'>"

'[UI-ANT] html = html & "<div >"

'[UI-ANT] html = html & "<table style='width:100%;'><tr style='display:flex;padding: 0 1rem;'>"

'[UI-ANT] html = html & "<td style='width:75%;' valign=center><a href='index.asp' class='logo' >"
'[UI-ANT] html = html & "<img src='images/iso_cor.png' style='width:6rem;' >"
'[UI-ANT] html = html & "<font style='font-size:15pt;color:white;'><b>Cor Gestion</b> | Sistema Administrativo</font></a></td>"
'[UI-ANT] html = html & "<td style='width:25%;display:flex;align-items:center;flex-direction:row-reverse;' align='right'>"
'[UI-ANT] html = html & "<nav>"
'[UI-ANT] html = html & "<ul>"
'[UI-ANT] 'html = html & "<li><a href='#menu'></a></li>"
'[UI-ANT] html = html & "</ul>"
'[UI-ANT] html = html & "</nav>"
'[UI-ANT] html = html & "</td>"

'[UI-ANT] html = html & "</tr></table>"

'[UI-ANT] html = html & "</div>"

'[UI-ANT] html = html & "</header>"

'[UI-ANT] html = html & "<table class='titulo-seccion' >"
'[UI-ANT] html = html & "<tr style='display:flex;align-items:center;justify-content:space-between;'>"

'[UI-ANT] html = html & "<td><h2>SEGURIDAD</h2></td>"

'[UI-ANT] form5 = "<form id='form5' style='height:70%;' method='post' action='permisos.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
'[UI-ANT] boton5 = "<input class='bot' type='button' value='VOLVER' onclick=""document.getElementById('form5').submit();"">"

'[UI-ANT] html = html & "<td align=right>" & boton5 & "</td>"

'[UI-ANT] html = html & "</tr>"

'[UI-ANT] html = html & "</table>"

'[UI-ANT] html = html & "</div>"

'[UI-ANT] html = html & "<div style='height:5%;width:100%;'>"

'[UI-ANT] html = html & "</div>"

'[UI-ANT] html = html & "<div style='height:57%;width:100%;float:left;'>"

'[UI-ANT] html = html & "<table class='tabl' style='height:80%;width:55%;color:#000000;' cellspacing=3 cellpadding=3 align=center>"

'[UI-ANT] sUsr = getName(cnn,"","cor_responsables",idusr)

'[UI-ANT] html = html & "<tr>"
'[UI-ANT] html = html & "<td colspan=6 style='font-size:15pt;' align=center><b>PERFIL&nbsp;" & sUsr & "</b></td>"
'[UI-ANT] html = html & "</tr>"

'[UI-ANT] sTit = ""

'[UI-ANT] if sTipoMenu = "INICIO" then
	'[UI-ANT] sTit = "MENU PRINCIPAL"
'[UI-ANT] elseif sTipoMenu = "PARAM" then
	'[UI-ANT] sTit = "PARAMETROS"
'[UI-ANT] end if

'[UI-ANT] html = html & "<tr>"
'[UI-ANT] html = html & "<td colspan=6><u>" & sTit & "</u></td>"
'[UI-ANT] html = html & "</tr>"

'[UI-ANT] html = html & "<tr>"
'[UI-ANT] html = html & "<td colspan=6>&nbsp;</td>"
'[UI-ANT] html = html & "</tr>"

'[UI-ANT] html = html & "<tr style='background-color:#A4A5B1;'>"
'[UI-ANT] html = html & "<td style='width:15%;' align=center>*</td>"
'[UI-ANT] html = html & "<td style='width:8%;' align=center>ESTADO</td>"
'[UI-ANT] html = html & "<td style='width:8%;' align=center>AGREGAR</td>"
'[UI-ANT] html = html & "<td style='width:8%;' align=center>MODIFICAR</td>"
'[UI-ANT] html = html & "<td style='width:8%;' align=center>ELIMINAR</td>"
'[UI-ANT] html = html & "<td style='width:8%;' align=center>CONSULTAR</td>"
'[UI-ANT] html = html & "</tr>"


'[UI-ANT] Set rs = Exec(Cnn, "sp_cg_MenuUsuario @Modo='listar', @IdUsr=" & SqlStr(idusr) & ", @FuncUsr='" & tipousr & "', @TipoMenu='" & sTipoMenu & "'")


'[UI-ANT] sColor = "#EAEBF2;"

'[UI-ANT] Do While Not rs.eof

	'[UI-ANT] 'estado
	'[UI-ANT] if rs(1) = 1 then
		'[UI-ANT] sImgEst = "<img src='images/circuloV.jpg' alt='estado' width='23' height='23' onclick=""document.getElementById('xFormEst" & rs(6) & "').submit();"" />"
	'[UI-ANT] elseif rs(1) = 0 then
		'[UI-ANT] sImgEst = "<img src='images/circuloR.jpg' alt='estado' width='20' height='20' onclick=""document.getElementById('xFormEst" & rs(6) & "').submit();"" />"
	'[UI-ANT] else
		'[UI-ANT] sImgEst = ""
	'[UI-ANT] end if
	'[UI-ANT] xLinkEst = "<form id='xFormEst" & rs(6) & "' name='xFormEst" & rs(6) & "' method='post' action='permisosDet.asp?opc=actEst&idmenu=" & rs(6) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&idusr=" & Server.URLEncode(idusr & "") & "&tipousr=" & tipousr & "&ssc=par&tipomenu=" & sTipoMenu & "'>"
	'[UI-ANT] xLinkEst = xLinkEst & "<input type='hidden' id='estado' name='estado' value='" & rs(1) & "' >" & sImgEst
	'[UI-ANT] xLinkEst = xLinkEst & "</form>"

	'[UI-ANT] 'agregar
	'[UI-ANT] if rs(2) = 1 then
		'[UI-ANT] sImgAgr = "<img src='images/circuloV.jpg' alt='agregar' width='23' height='23' onclick=""document.getElementById('xFormAgre" & rs(6) & "').submit();"" />"
	'[UI-ANT] elseif rs(2) = 0 then
		'[UI-ANT] sImgAgr = "<img src='images/circuloR.jpg' alt='agregar' width='20' height='20' onclick=""document.getElementById('xFormAgre" & rs(6) & "').submit();"" />"
	'[UI-ANT] else
		'[UI-ANT] sImgAgr = ""
	'[UI-ANT] end if
	'[UI-ANT] xLinkAgr = "<form id='xFormAgre" & rs(6) & "' name='xFormAgre" & rs(6) & "' method='post' action='permisosDet.asp?opc=actAgre&idmenu=" & rs(6) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&idusr=" & Server.URLEncode(idusr & "") & "&tipousr=" & tipousr & "&ssc=par&tipomenu=" & sTipoMenu & "'>"
	'[UI-ANT] xLinkAgr = xLinkAgr & "<input type='hidden' id='agregar' name='agregar' value='" & rs(2) & "' >" & sImgAgr
	'[UI-ANT] xLinkAgr = xLinkAgr & "</form>"

	'[UI-ANT] 'modificar
	'[UI-ANT] if rs(3) = 1 then
		'[UI-ANT] sImgMod = "<img src='images/circuloV.jpg' alt='modificar' width='23' height='23' onclick=""document.getElementById('xFormModi" & rs(6) & "').submit();"" />"
	'[UI-ANT] elseif rs(3) = 0 then
		'[UI-ANT] sImgMod = "<img src='images/circuloR.jpg' alt='modificar' width='20' height='20' onclick=""document.getElementById('xFormModi" & rs(6) & "').submit();"" />"
	'[UI-ANT] else
		'[UI-ANT] sImgMod = ""
	'[UI-ANT] end if
	'[UI-ANT] xLinkMod = "<form id='xFormModi" & rs(6) & "' name='xFormModi" & rs(6) & "' method='post' action='permisosDet.asp?opc=actModi&idmenu=" & rs(6) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&idusr=" & Server.URLEncode(idusr & "") & "&tipousr=" & tipousr & "&ssc=par&tipomenu=" & sTipoMenu & "'>"
	'[UI-ANT] xLinkMod = xLinkMod & "<input type='hidden' id='modificar' name='modificar' value='" & rs(3) & "' >" & sImgMod
	'[UI-ANT] xLinkMod = xLinkMod & "</form>"

	'[UI-ANT] 'eliminar
	'[UI-ANT] if rs(4) = 1 then
		'[UI-ANT] sImgElim = "<img src='images/circuloV.jpg' alt='eliminar' width='23' height='23' onclick=""document.getElementById('xFormElim" & rs(6) & "').submit();"" />"
	'[UI-ANT] elseif rs(4) = 0 then
		'[UI-ANT] sImgElim = "<img src='images/circuloR.jpg' alt='eliminar' width='20' height='20' onclick=""document.getElementById('xFormElim" & rs(6) & "').submit();"" />"
	'[UI-ANT] else
		'[UI-ANT] sImgElim = ""
	'[UI-ANT] end if
	'[UI-ANT] xLinkElim = "<form id='xFormElim" & rs(6) & "' name='xFormElim" & rs(6) & "' method='post' action='permisosDet.asp?opc=actElim&idmenu=" & rs(6) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&idusr=" & Server.URLEncode(idusr & "") & "&tipousr=" & tipousr & "&ssc=par&tipomenu=" & sTipoMenu & "'>"
	'[UI-ANT] xLinkElim = xLinkElim & "<input type='hidden' id='eliminar' name='eliminar' value='" & rs(4) & "' >" & sImgElim
	'[UI-ANT] xLinkElim = xLinkElim & "</form>"

	'[UI-ANT] 'consultar
	'[UI-ANT] if rs(5) = 1 then
		'[UI-ANT] sImgCons = "<img src='images/circuloV.jpg' alt='consultar' width='23' height='23' onclick=""document.getElementById('xFormCons" & rs(6) & "').submit();"" />"
	'[UI-ANT] elseif rs(5) = 0 then
		'[UI-ANT] sImgCons = "<img src='images/circuloR.jpg' alt='consultar' width='20' height='20' onclick=""document.getElementById('xFormCons" & rs(6) & "').submit();"" />"
	'[UI-ANT] else
		'[UI-ANT] sImgCons = ""
	'[UI-ANT] end if
	'[UI-ANT] xLinkCons = "<form id='xFormCons" & rs(6) & "' name='xFormCons" & rs(6) & "' method='post' action='permisosDet.asp?opc=actCons&idmenu=" & rs(6) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&idusr=" & Server.URLEncode(idusr & "") & "&tipousr=" & tipousr & "&ssc=par&tipomenu=" & sTipoMenu & "'>"
	'[UI-ANT] xLinkCons = xLinkCons & "<input type='hidden' id='consultar' name='consultar' value='" & rs(5) & "' >" & sImgCons
	'[UI-ANT] xLinkCons = xLinkCons & "</form>"

	'[UI-ANT] html = html & "<tr style='background-color:" & sColor & "'>"
	'[UI-ANT] html = html & "<td align=center >" & UCASE(rs(0)) & "</td>"
	'[UI-ANT] html = html & "<td align=center style='padding-top:30px;'>" & xLinkEst & "</td>"
	'[UI-ANT] html = html & "<td align=center >" & xLinkAgr & "</td>"
	'[UI-ANT] html = html & "<td align=center >" & xLinkMod & "</td>"
	'[UI-ANT] html = html & "<td align=center >" & xLinkElim & "</td>"
	'[UI-ANT] html = html & "<td align=center >" & xLinkCons & "</td>"
	'[UI-ANT] html = html & "</tr>"

	'[UI-ANT] if sColor = "#EAEBF2;" then
		'[UI-ANT] sColor = "#D9D9E3;"
	'[UI-ANT] elseif sColor = "#D9D9E3;" then
		'[UI-ANT] sColor = "#EAEBF2;"
	'[UI-ANT] end if

'[UI-ANT] rs.movenext
'[UI-ANT] Loop

'[UI-ANT] html = html & "<tr>"
'[UI-ANT] html = html & "<td colspan=6>&nbsp;</td>"
'[UI-ANT] html = html & "</tr>"

'[UI-ANT] html = html & "</table>"

'[UI-ANT] html = html & "</div>"

'[UI-ANT] html = html & "<script src='assets/js/jquery.min.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/skel.min.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/util.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/main.js'></script>"

'[UI-ANT] html = html & form5 & "</body>"
'[UI-ANT] html = html & "</html>"

'[UI-ANT] response.write html

' =========================================================================
' [UI] PANTALLA con el sistema visual (IncFile/ui_componentes.inc)
' Seguridad: permisos de un usuario sobre cada opcion del menu principal o de parametros.
' El cuerpo anterior completo queda arriba, comentado con '[UI-ANT].
' Mismo mecanismo de antes: cada permiso envia su valor actual y el servidor lo invierte
' (sp_cg_MenuUsuario @Modo='actualizar_flag'). Cambios: un solo bloque para los cinco permisos,
' valores e ids validados, permiso de modificar chequeado y confirmacion despues de guardar.
' =========================================================================
Dim puedeModif, accTop, aviso, flagOpc, campoOpc, nombreFlag, valorAct, valorNuevo, guardado, opcionGuardada
Dim filas, cantOpc, urlBase, sUsr, sTit

html = headhtml()
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
sTipoMenu = request.querystring("tipomenu")

' [UI] Solo dos tipos de menu posibles (se usa en la consulta)
If sTipoMenu <> "INICIO" And sTipoMenu <> "PARAM" Then sTipoMenu = "INICIO"
puedeModif = (getPermisosAtrm("CG_PERMISOS", "MODIF") = 1)

' --- Cambio de un permiso: opc -> (campo del formulario, flag de la base) ---
flagOpc = ""
Select Case opc
	Case "actEst"
		campoOpc = "estado"
		flagOpc = "estado"
		nombreFlag = "Acceso"
	Case "actAgre"
		campoOpc = "agregar"
		flagOpc = "agregar"
		nombreFlag = "Agregar"
	Case "actModi"
		campoOpc = "modificar"
		flagOpc = "modificar"
		nombreFlag = "Modificar"
	Case "actElim"
		campoOpc = "eliminar"
		flagOpc = "eliminar"
		nombreFlag = "Eliminar"
	Case "actCons"
		campoOpc = "consultar"
		flagOpc = "consultar"
		nombreFlag = "Consultar"
End Select

aviso = ""
guardado = False
If flagOpc <> "" Then
	valorAct = request.form(campoOpc) & ""
	If Not puedeModif Then
		' [UI] Antes cualquiera con la URL podia cambiar permisos
		aviso = "<div class='ui-alert ui-alert--danger' role='alert'>No ten&eacute;s permiso para modificar permisos.</div>"
	ElseIf (valorAct = "0" Or valorAct = "1") And IsNumeric(idmenu & "") And Trim(idmenu & "") <> "" Then
		' [UI] Valor e id validados: antes un valor distinto de 0/1 se pegaba tal cual en la consulta
		If valorAct = "1" Then valorNuevo = 0 Else valorNuevo = 1
		ExecNonQuery Cnn, "sp_cg_MenuUsuario @Modo='actualizar_flag', @Flag='" & flagOpc & "', @Valor=" & valorNuevo & ", @Id=" & CLng(idmenu)
		guardado = True
	End If
End If

html = html & "<div id='wrapper' role='banner'>" & cabecera() & "</div>"

form5 = "<form id='form5' class='form-oculto' method='post' action='permisos.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
accTop = UiBoton("Volver", "secondary", "document.getElementById('form5').submit();", True, "")

sUsr = getName(cnn, "", "cor_responsables", idusr)
If sTipoMenu = "PARAM" Then
	sTit = "Par&aacute;metros"
Else
	sTit = "Men&uacute; principal"
End If

' --- Matriz de permisos ---
urlBase = "permisosDet.asp?idmenu={ID}&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&idusr=" & Server.URLEncode(idusr & "") & "&tipousr=" & Server.URLEncode(tipousr & "") & "&ssc=par&tipomenu=" & sTipoMenu

' [UI] Tipo de usuario y menu escapados en la consulta (antes se pegaban tal cual)
Set rs = Exec(Cnn, "sp_cg_MenuUsuario @Modo='listar', @IdUsr=" & SqlStr(idusr) & ", @FuncUsr=" & SqlStr(tipousr) & ", @TipoMenu=" & SqlStr(sTipoMenu))

filas = ""
cantOpc = 0
opcionGuardada = ""
Do While Not rs.eof
	If guardado And CStr(rs(6)) = CStr(idmenu) Then opcionGuardada = UiCapitalizar(rs(0))
	filas = filas & "<tr id='m" & rs(6) & "'>" & _
		"<td><span class='ui-table__titulo'>" & UiTexto(UiCapitalizar(rs(0))) & "</span></td>" & _
		"<td class='ui-table__cell--center'>" & PermInterruptor("actEst", "estado", "Acceso", rs(1), rs(6), rs(0)) & "</td>" & _
		"<td class='ui-table__cell--center'>" & PermInterruptor("actAgre", "agregar", "Agregar", rs(2), rs(6), rs(0)) & "</td>" & _
		"<td class='ui-table__cell--center'>" & PermInterruptor("actModi", "modificar", "Modificar", rs(3), rs(6), rs(0)) & "</td>" & _
		"<td class='ui-table__cell--center'>" & PermInterruptor("actElim", "eliminar", "Eliminar", rs(4), rs(6), rs(0)) & "</td>" & _
		"<td class='ui-table__cell--center'>" & PermInterruptor("actCons", "consultar", "Consultar", rs(5), rs(6), rs(0)) & "</td>" & _
		"</tr>"
	cantOpc = cantOpc + 1
	rs.movenext
Loop

If guardado Then
	aviso = "<div class='ui-alert ui-alert--success' role='status'>Permiso actualizado: " & nombreFlag
	If opcionGuardada <> "" Then aviso = aviso & " en " & UiTexto(opcionGuardada)
	aviso = aviso & ".</div>"
End If

If Trim(sUsr & "") = "" Then
	html = html & UiPantallaSeguridadAbre("Permisos", accTop, "")
Else
	html = html & UiPantallaSeguridadAbre("Permisos de " & UiTexto(sUsr), accTop, "")
End If
html = html & aviso
html = html & UiCardAbre(sTit, "")
html = html & UiTablaAbre("Permisos del usuario en cada opci&oacute;n del " & LCase(sTit))
html = html & "<thead><tr>" & UiTh("Opci&oacute;n", "") & UiTh("Acceso", "ui-table__cell--center") & UiTh("Agregar", "ui-table__cell--center") & _
	UiTh("Modificar", "ui-table__cell--center") & UiTh("Eliminar", "ui-table__cell--center") & UiTh("Consultar", "ui-table__cell--center") & "</tr></thead><tbody>"
If cantOpc = 0 Then
	html = html & UiFilaVacia(6, "No hay opciones de men&uacute; para este usuario.")
Else
	html = html & filas
End If
html = html & UiTablaCierra()
If puedeModif Then
	html = html & "<p class='ui-table__nota'>Toc&aacute; S&iacute; o No para cambiar el permiso. Se guarda al instante. Un gui&oacute;n indica que el permiso no aplica a esa opci&oacute;n.</p>"
Else
	html = html & "<p class='ui-table__nota'>Solo lectura: no ten&eacute;s permiso para modificar permisos. Un gui&oacute;n indica que el permiso no aplica a esa opci&oacute;n.</p>"
End If
html = html & UiCardCierra()
html = html & UiPantallaCierra()

html = html & "<script src='assets/js/jquery.min.js'></script>"
html = html & "<script src='assets/js/skel.min.js'></script>"
html = html & "<script src='assets/js/util.js'></script>"
html = html & "<script src='assets/js/main.js'></script>"

html = html & form5 & "</body>"
html = html & "</html>"

response.write html


' =============================================================================
' FUNCIONES DE ESTA PANTALLA
' =============================================================================

' Interruptor Si / No de un permiso. Envia el valor actual (como antes) y vuelve a la misma fila.
' Valor distinto de 0 o 1: el permiso no aplica a esa opcion (guion, sin formulario)
Function PermInterruptor(opcion, campo, nombre, valor, idMenuFila, nombreOpcion)
	Dim activo, texto, etiqueta, accion, presionado, clase
	If CStr(valor & "") <> "1" And CStr(valor & "") <> "0" Then
		PermInterruptor = "<span class='ui-status-empty' aria-label='No aplica'>&mdash;</span>"
		Exit Function
	End If
	activo = (CStr(valor) = "1")
	' Texto fijo para aria-pressed: CStr(True) da "Verdadero" en un servidor en espanol
	If activo Then
		texto = "S&iacute;"
		presionado = "true"
		clase = "ui-permiso ui-permiso--si"
	Else
		texto = "No"
		presionado = "false"
		clase = "ui-permiso ui-permiso--no"
	End If
	etiqueta = nombre & " en " & UiTexto(UiCapitalizar(nombreOpcion)) & ": " & texto
	If puedeModif Then
		accion = Replace(urlBase, "{ID}", idMenuFila) & "&opc=" & opcion & "#m" & idMenuFila
		PermInterruptor = "<form class='form-linea' method='post' action='" & accion & "'>" & _
			"<input type='hidden' name='" & campo & "' value='" & valor & "'>" & _
			"<button type='submit' class='" & clase & "' aria-pressed='" & presionado & "' aria-label='" & etiqueta & ". Cambiar'>" & texto & "</button></form>"
	Else
		PermInterruptor = "<span class='" & clase & " ui-permiso--fijo' role='img' aria-label='" & etiqueta & "'>" & texto & "</span>"
	End If
End Function
%>