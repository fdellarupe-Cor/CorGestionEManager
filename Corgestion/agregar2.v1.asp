<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/datosAgregar.inc" -->
<%

html = headhtml()
html = html & "<body>"
ruta=server.mappath(request.servervariables("script_name"))

pos=InStrRev(ruta,"\")
ruta=Mid(ruta,1,pos)


opc = request.querystring("opc")
tipo = request.querystring("tipo")
xmenu = request.querystring("xmenu")
xtmenu = request.querystring("xtmenu")

'Dim st0,st1,sel,hs

'st0 = "height:22.5px;border-radius:2.5px;width:175px;font-family:arial;font-size:7.75pt;border-color:#dedede;background-color:fff;border-style:none;"
'st1 = "font-family:verdana;font-size:7.75pt;padding-left:10px;"
'sel = "&nbsp;&nbsp;&nbsp;------&nbsp;&nbsp;&nbsp; SELECCIONE &nbsp;&nbsp;&nbsp;------ "
'hs = "width:40px;text-align:right;padding-right:5px;"

'response.write request.querystring
Set Cnn = Server.CreateObject("ADODB.Connection")
'Cnn.open "Provider=SQLOLEDB;Data Source=server-2025;Initial Catalog=dbcorgestionv5; user id=sa;password=bb4515$;"
On Error Resume Next
Cnn.Open scon1
If Err.Number <> 0 Then
	Response.Write "<div style='color:red;font-family:monospace;padding:10px;border:1px solid red;margin:10px;'>"
	Response.Write "<b>ERROR DE CONEXION A BASE DE DATOS</b><br>"
	Response.Write "Codigo: " & Err.Number & "<br>"
	Response.Write "Detalle: " & Server.HTMLEncode(Err.Description) & "</div>"
	Response.End
End If
On Error Goto 0

If opc <> "ins" Then

	Select Case tipo

		Case "INC"

		tabla = "INCIDENTES"
		tit = "INCIDENTE"
		volver = "incidentes.asp?opc=listar"

		Case "PROY"

		tabla = "PROYECTOS"
		tit = "PROYECTO"
		volver = "proyectos.asp"

		Case "LIC"

		tabla = "LICENCIAS"
		tit = "LICENCIA"
		volver = "licencias.asp"

	End Select

End If

If opc = "ins" Then

	nFecini = 0

	If nFecini < CLng(CDate("01-01-1900")) Then
		nFecini = CLng(CDate("01-01-1900"))
	Else
		nFecini = 0
	End If

	sFec = CLng(Date) - nFecini

	tabla = request.querystring("tabla")

	If tabla = "INCIDENTES" Then

		' El estado inicial (estado_desa 0/2 segun haya respon_tec) ahora se resuelve
		' dentro del SP; se preserva la misma logica.
		InsertarIncidente Cnn, request.form("cli"), request.form("con"), request.form("apli"), _
			request.form("det"), request.form("prio"), request.form("rg"), request.form("rt"), _
			request.form("rp"), sFec, request.form("br"), request.form("vp"), request.form("tit"), RolResponsableActual()

		XXX = SYSLOGON(cnn,"AGREGA_INC")

		xid = dameUltId("I",cnn)

		vvv = ObjAcc(xid,"I",cnn,"C")
		 set fso = SERVER.CreateObject("Scripting.FileSystemObject")
		  ruta=ruta & "incidentes\"  &   xid
		  ' response.write ruta
	       If not fso.FolderExists(ruta ) Then
		     fso.CreateFolder (ruta)
	       end if
          set fso=nothing
	ElseIf tabla = "LICENCIAS" Then

		sFecDsd = clng(CDATE(request.form("fd"))) - CLng(CDate("01-01-1900"))
		sFecHas = clng(CDATE(request.form("fh"))) - CLng(CDate("01-01-1900"))

		InsertarLicencia Cnn, session("usuario"), sFecDsd, sFecHas, request.form("mot")

		XXX = SYSLOGON(cnn,"AGREGA_LIC")

	Else 'PROYECTOS
	
		response.write "Guardar Proyecto"
		sfecentrega = 0
		if request.form("fe") & "" <> "" then
			sfecentrega = clng(cdate(request.form("fe"))) - clng(cdate("01-01-1900"))

		end if
		insertarproyecto cnn, request.form("sp"), request.form("tp"), request.form("det"), request.form("con"), _
			request.form("cli"), request.form("apli"), request.form("rg"), request.form("rt"), sfec, sfecentrega, _
			request.form("pri"), request.form("rd"), request.form("di"), request.form("de"), request.form("te"), _
			request.form("br"), request.form("vp"), request.form("rq"), rolresponsableactual() , "1"
			
			
		xxx = syslogon(cnn,"agrega_proy")

		xid = dameultid("P",cnn)

		 vvv = objacc(xid,"P",cnn,"C")

		set fso = server.createobject("scripting.filesystemobject")
			  ruta=ruta & "proyectos\"  &   xid
			  response.write ruta
			   if not fso.folderexists(ruta ) then
				 fso.createfolder (ruta)
			   end if
	   set fso=nothing

	End If

End If

html = html & "<div>"

html = html & cabecera()

If tipo = "INC" Then

	'html = html & "<div style='height:8%;width:100%;border-top:solid 5px #FF3D3D;border-bottom:double 6px #001BFF;'>"

	'html = html & "<table style='height:100%;width:100%;'>"

	'html = html & "<tr>"
	html = html & "<table class='titulo-seccion ui-page-header' >"
	html = html & "<tr style='display:flex;align-items:center;justify-content:space-between;'>"

	html = html & "<td><h2>NUEVO " & tit & "</h2></td>"

	form5 = "<form id='form5' style='height:70%;' method='post' action='incidentes.asp?opc=listar&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&tipo=INC'></form>"
	boton5 = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('form5').submit();"">"

	html = html & "<td>" & boton5 & "</td>"

	html = html & "</tr>"

	html = html & "</table>"

	html = html & "</div>"


	' Usamos tu contenedor estÃ¡ndar para mantener el padding global del sistema
	html = html & "<div class='inner' style='padding:2rem'>"
	
	'html = html & "<div style='height:82%;width:100%;'><br>"
	' Le damos un max-width y margin: 0 auto para que el formulario quede centrado y no se estire al infinito en monitores grandes
	html = html & "<div style='max-width: 1000px; margin: 0 auto;'>"

	html = html & "<form id='agre' method='post' action='agregar2.asp?opc=ins&tabla=" & tabla & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&tipo=INC&volver=" & volver & "'>"

	'html = html & "<table class='tabl' align=center cellspacing=0 cellpadding=3>"
	html = html & "<div class='bloque-datos ui-panel'>"

	' --- HEADER DEL FORMULARIO ---
	'html = html & "<tr>"
	'html = html & "<td class='titu' colspan=4><u>COMPLETE LOS SIGUIENTES DATOS</u></td>"
	'html = html & "</tr>"
	html = html & "<div class='bloque-header'>COMPLETE LOS SIGUIENTES DATOS</div>"

	'html = html & "<tr>"
	'html = html & "<td class='fil' colspan=4 align=left>&nbsp;</td>"
	'html = html & "</tr>"
	
	' --- BODY DEL FORMULARIO Y GRILLA DOBLE ---
	html = html & "<div class='bloque-body ui-panel__body'>"
	html = html & "<div class='grid-datos-doble ui-field-grid--dual'>"

	' --- FILA 1 ---
	'html = html & "<tr>"
	'html = html & "<td class='fil'>CLIENTE:</td>"
	'html = html & "<td class='fil'><select class='selec' id='cli' name='cli'>"
	html = html & "<div class='dato-label'>CLIENTE</div>"
	html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='cli' name='cli'>"

	Set rs = ObtenerClientes(Cnn)

	Do While Not rs.eof
		html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
		rs.movenext
	Loop
	html = html & "</select></div>"

	'html = html & "<td class='fil'>TITULO:</td>"
	'html = html & "<td class='fil'><input class='inpu' id='tit' name='tit' type='text'></td>"
	'html = html & "</tr>"
	html = html & "<div class='dato-label'>TITULO</div>"
	html = html & "<div class='dato-value'><input class='input-moderno ui-input' id='tit' name='tit' type='text'></div>"


	' --- FILA 2 ---
	'html = html & "<tr>"
	'html = html & "<td class='fil'>APLICACION:</td>"
	'html = html & "<td class='fil'><select class='selec' id='apli' name='apli'>"
	html = html & "<div class='dato-label'>APLICACION</div>"
	html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='apli' name='apli'>"

	Set rs = ObtenerAplicativos(Cnn)

	Do While Not rs.eof
		html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
		rs.movenext
	Loop
	html = html & "</select></div>"

	'html = html & "<td class='fil'>PRIORIDAD:</font></td>"
	'html = html & "<td class='fil'><select class='selec' id='prio' name='prio'>"
	html = html & "<div class='dato-label'>PRIORIDAD</div>"
	html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='prio' name='prio'>"

	Set rs = ObtenerPrioridades(Cnn)

	Do While Not rs.eof
		if rs(1) = "MEDIA" then
			html = html & "<option value='" & rs(0) & "' selected>" & Server.HTMLEncode(rs(1) & "") & "</option>"
		else
			html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
		end if
		rs.movenext
	Loop
	html = html & "</select></div>"


	' --- FILA 3 ---
	'html = html & "<tr>"
	'html = html & "<td class='fil'>RESP. GESTION:</td>"
	'html = html & "<td class='fil'><select class='selec' id='rg' name='rg'>"
	html = html & "<div class='dato-label'>RESP. GESTION</div>"
	html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='rg' name='rg'>"

	Set rs = ObtenerResponsables(Cnn, "FUNC")

	Do While Not rs.eof
		html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
		rs.movenext
	Loop
	html = html & "</select></div>"

	'html = html & "<td class='fil'>RESP. TECNICO:</td>"
	'html = html & "<td class='fil'><select class='selec' id='rt' name='rt'>"
	html = html & "<div class='dato-label'>RESP. TECNICO</div>"
	html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='rt' name='rt'>"

	Set rs = ObtenerResponsables(Cnn, "TEC")

	html = html & "<option value='0'>NINGUNO (CONSULTA)</option>"

	Do While Not rs.eof
		html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
		rs.movenext
	Loop
	html = html & "</select></div>"


	' --- FILA 4 ---
	'html = html & "<tr>"
	'html = html & "<td class='fil'>REQ. PRESUP:</td>"
	'html = html & "<td class='fil'><select class='selec' id='rp' name='rp'>"
	'html = html & "<option value='SI'>SI</option>"
	'html = html & "<option value='NO'>NO</option>"
	'html = html & "</select></td>"
	html = html & "<div class='dato-label'>REQ. PRESUP</div>"
	html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='rp' name='rp'>"
	html = html & "<option value='SI'>SI</option>"
	html = html & "<option value='NO'>NO</option>"
	html = html & "</select></div>"

	'html = html & "<td class='fil'>SOLIC. PROYECTO:</td>"
	'html = html & "<td class='fil'><input class='inpu' id='sp' name='sp' type='text'></td>"
	'html = html & "</tr>"
	html = html & "<div class='dato-label'>SOLIC. PROYECTO</div>"
	html = html & "<div class='dato-value'><input class='input-moderno ui-input' id='sp' name='sp' type='text'></div>"


	' --- FILA 5 ---
	'html = html & "<tr>"
	'html = html & "<td class='fil'>VIA PEDIDO:</td>"
	'html = html & "<td class='fil'><select class='selec' id='vp' name='vp'>"
	html = html & "<div class='dato-label'>VIA PEDIDO</div>"
	html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='vp' name='vp'>"

	Set rs = ObtenerViaPedido(Cnn)

	Do While Not rs.eof
		html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
		rs.movenext
	Loop
	html = html & "</select></div>"

	'html = html & "<td class='fil'>CONTACTO:</td>"
	'html = html & "<td class='fil'><select class='selec' id='con' name='con'>"
	html = html & "<div class='dato-label'>CONTACTO</div>"
	html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='con' name='con'>"

	Set rs = ObtenerContactosParaIncidente(Cnn)

	Do While Not rs.eof
		html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
	rs.movenext
	Loop
	html = html & "</select></div>"


	' --- FILA 6 ---
	'html = html & "<tr>"
	'html = html & "<td class='fil'>BROWSER:</td>"
	'html = html & "<td class='fil'><select class='selec' id='br' name='br'>"
	html = html & "<div class='dato-label'>BROWSER</div>"
	html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='br' name='br'>"

	Set rs = ObtenerBrowsers(Cnn)

	Do While Not rs.eof
		html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
		rs.movenext
	Loop
	html = html & "</select></div>"

	' Completamos la grilla para que no se desfase (antes eran <td> vacÃ­os)
	'html = html & "<td></td>"
	'html = html & "<td></td>"
	'html = html & "</tr>"
	html = html & "<div></div><div></div>"
	
	' Â¡CERRAMOS LA GRILLA DE 4 COLUMNAS ACÃ�!
	html = html & "</div>"


	' --- SECCIÃ“N DETALLES (FUERA DE LA GRILLA) ---
	'html = html & "<tr>"
	'html = html & "<td class='fil' colspan=4 align=left>&nbsp;</td>"
	'html = html & "</tr>"
	'html = html & "<tr>"
	'html = html & "<td class='fil' colspan=4 >DETALLES:</td>"
	'html = html & "</tr>"
	'html = html & "<tr>"
	'html = html & "<td style='height:150px;' colspan=4><textarea class='txtar' id='det' name='det'></textarea></td>"
	'html = html & "</tr>"
	
	' Agregamos un div contenedor con margen superior para separar "DETALLES" de la grilla de arriba
	html = html & "<div style='margin-top: 1.5rem;'>"
	html = html & "<div class='dato-label' style='margin-bottom: 0.5rem;'>DETALLES</div>"
	html = html & "<textarea class='input-moderno ui-input' id='det' name='det' style='height:150px; resize:vertical; font-family: inherit;'></textarea>"
	html = html & "</div>"


	' --- SECCIÃ“N BOTÃ“N ---
	'html = html & "<tr>"
	'html = html & "<td colspan=4>&nbsp;</td>"
	'html = html & "</tr>"
	'html = html & "<tr>"
	'html = html & "<td colspan=4 align=center><input type='button' class='bot' style='height:25px;' value='Agregar' onclick='agregar();' ></td>"
	'html = html & "</tr>"
	
	' Contenedor del botÃ³n alineado al centro y con estilo moderno
	html = html & "<div style='margin-top: 2rem; text-align: center; border-top: 1px solid #e2e8f0; padding-top: 1.5rem;'>"
	html = html & "<input class='ui-button ui-button--primary' type='button' style='padding: 0.6rem 2rem; background-color: #3b82f6; color: white; border: none; border-radius: 6px; font-weight: 600; cursor: pointer; transition: all 0.2s ease;' value='Agregar' onclick='agregar();' onmouseover=""this.style.backgroundColor='#2563eb'"" onmouseout=""this.style.backgroundColor='#3b82f6'"" onmousedown=""this.style.transform='scale(0.95)'"" onmouseup=""this.style.transform='scale(1)'"">"
	html = html & "</div>"

	'html = html & "</table>"

	' Cerramos los contenedores de la tarjeta moderna
	html = html & "</div>" ' Cierra bloque-body
	html = html & "</div>" ' Cierra bloque-datos

	html = html & "</form>"
	html = html & "</div>" ' Cierra el div wrapper principal
	html = html & "</div>"

ElseIf tipo = "PROY" Then



	'html = html & "<div style='height:8%;width:70%;border-top:solid 5px #FF3D3D;border-bottom:double 6px #001BFF;margin:0 auto;'>"
	'html = html & "<table style='height:100%;width:100%;'>"
	'html = html & "<tr>"
	html = html & "<table class='titulo-seccion ui-page-header' >"
	html = html & "<tr style='display:flex;align-items:center;justify-content:space-between;'>"

	html = html & "<td><h2>NUEVO " & tit & "</h2></td>"

	form6 = "<form id='form6' style='height:70%;' method='post' action='proyectos.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
	boton6 = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('form6').submit();"">"

	html = html & "<td>" & boton6 & "</td>"

	html = html & "</tr>"
	html = html & "</table>"
	html = html & "</div>"
	
	' ==========================================
	' FORMULARIO DE ALTA DE PROYECTOS
	' ==========================================
	
	' Usamos el contenedor estÃ¡ndar con padding
	html = html & "<div class='inner' style='padding:2rem'>"
	
	' Limitamos el ancho para no deformar en monitores grandes
	html = html & "<div style='max-width: 1000px; margin: 0 auto;'>"

	html = html & "<form id='agre' method='post' action='agregar2.asp?opc=ins&tabla=" & tabla & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&volver=" & volver & "'>"
	html = html & "<INPUT TYPE=HIDDEN NAME=ID_CLI>"

	' --- TARJETA PRINCIPAL ---
	html = html & "<div class='bloque-datos ui-panel'>"

	' --- HEADER ---
	'html = html & "<td class='titu' colspan=4 align=center style='font-family:helvetica;font-size:11pt;font-weight:bold;'><br>COMPLETE LOS SIGUIENTES DATOS</td>"
	html = html & "<div class='bloque-header'>COMPLETE LOS SIGUIENTES DATOS</div>"

	' --- BODY Y GRILLA DE CAMPOS ---
	html = html & "<div class='bloque-body ui-panel__body'>"
	html = html & "<div class='grid-datos-doble ui-field-grid--dual'>"

	' --- FILA 1 ---
	'html = html & "<td class='fil' style='" & st1 & "' >SOLIC. PROYECTO NRO:</td>"
	'html = html & "<td class='fil'><input class='inpu' style='" & st0 & "' id='sp' name='sp' type='text'></td>"
	html = html & "<div class='dato-label'>SOLIC. PROYECTO NRO</div>"
	html = html & "<div class='dato-value'><input class='input-moderno ui-input' id='sp' name='sp' type='text'></div>"

	'html = html & "<td class='fil' style='" & st1 & "' >TITULO PROYECTO:</td>"
	'html = html & "<td class='fil'><input class='inpu' style='" & st0 & "' id='tp' name='tp' type='text' autocomplete='off'></td>"
	html = html & "<div class='dato-label'>TITULO PROYECTO</div>"
	html = html & "<div class='dato-value'><input class='input-moderno ui-input' id='tp' name='tp' type='text' autocomplete='off'></div>"

	' --- FILA 2 ---
	'html = html & "<td class='fil' style='" & st1 & "' >CLIENTE:</td>"
	'html = html & "<td class='fil'><select class='selec' style='" & st0 & "'  id='cli' name='cli' ONCHANGE='VerSel()'>"
	html = html & "<div class='dato-label'>CLIENTE</div>"
	html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='cli' name='cli' onchange='VerSel()'>"

	' Mantenemos el SCRIPT intacto
	html = html & "<SCRIPT>"
	html = html & " function VerSel() { "
	html = html & " var a; "
	html = html & " a = document.getElementById('cli'); "
	html = html & " var v; "
	html = html & " v = a.value; "
	html = html & " console.log(v); "
	html = html & " return v; }"
	html = html & "</SCRIPT>"

	Set rs = ObtenerClientes(Cnn)
		html = html & "<option value='0' >" & sel & "</option>"
	Do While Not rs.eof
		html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
		rs.movenext
	Loop
	html = html & "</select></div>"

	'html = html & "<td class='fil' style='" & st1 & "' >CONTACTO:</td>"
	'html = html & "<select class='selec' style='" & st0 & "' id='con' name='con'>"
	html = html & "<div class='dato-label'>CONTACTO</div>"
	html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='con' name='con'>"

	Set rs = ObtenerContactosParaProyecto(Cnn)
		html = html & "<option value='0' >" & sel & "</option>"
	Do While Not rs.eof
		html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
		rs.movenext
	Loop
	html = html & "</select></div>"

	' --- FILA 3 ---
	'html = html & "<td class='fil' style='" & st1 & "' >APLICACION:</td>"
	html = html & "<div class='dato-label'>APLICACION</div>"
	html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='apli' name='apli'>"

	Set rs = ObtenerAplicativos(Cnn)
		html = html & "<option value='0' >" & sel & "</option>"
	Do While Not rs.eof
		html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
		rs.movenext
	Loop
	html = html & "</select></div>"

	'html = html & "<td class='fil' style='" & st1 & "' >PRIORIDAD:</td>"
	html = html & "<div class='dato-label'>PRIORIDAD</div>"
	html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='pri' name='pri'>"

	Set rs = ObtenerPrioridades(Cnn)
		html = html & "<option value='0' >" & sel & "</option>"
	Do While Not rs.eof
		html = html & "<option value='" & rs(0) & "' >" & Server.HTMLEncode(rs(1) & "") & "</option>"
		rs.movenext
	Loop
	html = html & "</select></div>"

	' --- FILA 4 ---
	'html = html & "<td class='fil' style='" & st1 & "' >RESP. GESTION:</td>"
	html = html & "<div class='dato-label'>RESP. GESTION</div>"
	html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='rg' name='rg'>"

	Set rs = ObtenerResponsables(Cnn, "FUNC")
		html = html & "<option value='0' >" & sel & "</option>"
	Do While Not rs.eof
		html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & " " & rs(2) & "") & "</option>"
		rs.movenext
	Loop
	html = html & "</select></div>"

	'html = html & "<td class='fil' style='" & st1 & "' >RESP. TECNICO:</td>"
	html = html & "<div class='dato-label'>RESP. TECNICO</div>"
	html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='rt' name='rt'>"

	Set rs = ObtenerResponsables(Cnn, "TEC")
		html = html & "<option value='0' >" & sel & "</option>"
	Do While Not rs.eof
		html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & " " & rs(2) & "") & "</option>"
		rs.movenext
	Loop
	html = html & "</select></div>"

	' --- FILA 5 ---
	'html = html & "<td class='fil' style='" & st1 & "' >REQ. PRESUP:</td>"
	html = html & "<div class='dato-label'>REQ. PRESUP</div>"
	html = html & "<div class='dato-value'><select class='input-moderno ui-input' style='width: 100px;' id='rq' name='rq'>"
	html = html & "<option value='SI'>SI</option>"
	html = html & "<option value='NO'>NO</option>"
	html = html & "</select></div>"

	'html = html & "<td class='fil' style='" & st1 & "' >BROWSER:</td>"
	html = html & "<div class='dato-label'>BROWSER</div>"
	html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='br' name='br'>"

	Set rs = ObtenerBrowsers(Cnn)
		html = html & "<option value='0' >" & sel & "</option>"
	Do While Not rs.eof
		html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
	rs.movenext
	Loop
	html = html & "</select></div>"

	' --- FILA 6 ---
	'html = html & "<td class='fil' style='" & st1 & "' >VIA PEDIDO:</td>"
	html = html & "<div class='dato-label'>VIA PEDIDO</div>"
	html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='vp' name='vp'>"

	Set rs = ObtenerViaPedido(Cnn)
		html = html & "<option value='0' >" & sel & "</option>"
	Do While Not rs.eof
		html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
		rs.movenext
	Loop
	html = html & "</select></div>"

	'html = html & "<td class='fil' style='" & st1 & "' >FECHA ENTREGA:</td>"
	html = html & "<div class='dato-label'>FECHA ENTREGA</div>"
	html = html & "<div class='dato-value'><input class='input-moderno ui-input' id='fe' name='fe' type='text'></div>"

	' --- FILA 7 ---
	'html = html & "<td class='fil' style='" & st1 & "' >RELEV Y DEFINICION (hs):</td>"
	html = html & "<div class='dato-label'>RELEV Y DEFINICION (hs)</div>"
	' A los campos de horas les limitamos un poco el ancho para que no queden como una barra gigante
	html = html & "<div class='dato-value'><input class='input-moderno ui-input' style='max-width: 120px;' id='rd' name='rd' type='text'></div>"

	'html = html & "<td class='fil' style='" & st1 & "' >DESARROLLO (hs):</td>"
	html = html & "<div class='dato-label'>DESARROLLO (hs)</div>"
	html = html & "<div class='dato-value'><input class='input-moderno ui-input' style='max-width: 120px;' id='de' name='de' type='text'></div>"

	' --- FILA 8 ---
	'html = html & "<td class='fil' style='" & st1 & "' >DISE&Ntilde;O (hs):</td>"
	html = html & "<div class='dato-label'>DISE&Ntilde;O (hs)</div>"
	html = html & "<div class='dato-value'><input class='input-moderno ui-input' style='max-width: 120px;' id='di' name='di' type='text'></div>"

	'html = html & "<td class='fil' style='" & st1 & "' >TEST (hs):</td>"
	html = html & "<div class='dato-label'>TEST (hs)</div>"
	html = html & "<div class='dato-value'><input class='input-moderno ui-input' style='max-width: 120px;' id='te' name='te' type='text'></div>"

	' Cerramos la grilla
	html = html & "</div>"

	' --- TEXTAREA: DESCRIPCION DEL PROYECTO (Fuera de la grilla) ---
	html = html & "<div style='margin-top: 1.5rem;'>"
	html = html & "<div class='dato-label' style='margin-bottom: 0.5rem;'>DESCRIPCION DEL PROYECTO</div>"
	html = html & "<textarea class='input-moderno ui-input' id='det' name='det' style='height: 300px; resize: vertical; font-family: inherit;'></textarea>"
	html = html & "</div>"

	' --- BOTON AGREGAR ---
	html = html & "<div style='margin-top: 2rem; text-align: center; border-top: 1px solid #e2e8f0; padding-top: 1.5rem;'>"
	html = html & "<input class='ui-button ui-button--primary' type='button' style='padding: 0.6rem 2rem; background-color: #3b82f6; color: white; border: none; border-radius: 6px; font-weight: 600; cursor: pointer; transition: all 0.2s ease;' value='Agregar' onclick='agregar2();' onmouseover=""this.style.backgroundColor='#2563eb'"" onmouseout=""this.style.backgroundColor='#3b82f6'"" onmousedown=""this.style.transform='scale(0.95)'"" onmouseup=""this.style.transform='scale(1)'"">"
	html = html & "</div>"

	' Cerramos todos los contenedores
	html = html & "</div>" ' Cierra bloque-body
	html = html & "</div>" ' Cierra bloque-datos
	html = html & "</form>"
	html = html & "</div>" ' Cierra div del form centrado
	html = html & "</div>" ' Cierra div inner

Elseif tipo = "LIC" then

	html = html & "<table class='titulo-seccion ui-page-header' >"
	html = html & "<tr style='display:flex;align-items:center;justify-content:space-between;'>"

	html = html & "<td><h2>NUEVA " & tit & "</h2></td>"

	form5 = "<form id='form5' style='height:70%;' method='post' action='licencias.asp?" & request.querystring & "'></form>"
	boton5 = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('form5').submit();"">"

	html = html & "<td>" & boton5 & "</td>"

	html = html & "</tr>"

	html = html & "</table>"

	html = html & "</div>"

	' --- INICIO NORMALIZACION: FORMULARIO ALTA DE LICENCIAS ---
	
	'html = html & "<div style='height:120%;width:100%;'><br>"
	html = html & "<div class='inner' style='padding: 2rem;'>" ' Usamos un contenedor con aire

	'html = html & "<form id='agre' method='post' action='agregar2.asp?opc=ins&tabla=" & tabla & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&volver=" & volver & "'>"
	html = html & "<form id='agre' method='post' action='agregar2.asp?opc=ins&tabla=" & tabla & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&volver=" & volver & "'>"

	'html = html & "<table class='tabl' style='width:40%;' cellspacing=0 cellpadding=3 align=center >"
	html = html & "<div class='bloque-datos ui-panel' style='max-width: 600px; margin: 0 auto;'>"

	'html = html & "<tr>"
	'html = html & "<td class='fil' style='font-size:12pt;' colspan=2 align=center><b><u>COMPLETE LOS SIGUIENTES DATOS</u></b></td>"
	'html = html & "</tr>"
	html = html & "<div class='bloque-header' style='text-align: center; text-transform: uppercase;'>COMPLETE LOS SIGUIENTES DATOS</div>"

	'html = html & "<tr>"
	'html = html & "<td colspan=2>&nbsp;</td>"
	'html = html & "</tr>"
	html = html & "<div class='bloque-body ui-panel__body'>"
	
	' Abrimos la grilla moderna para el formulario
	html = html & "<div class='grid-datos ui-field-grid'>"

	'html = html & "<tr>"
	'html = html & "<td class='fil'>RESPONSBALE</td>"
	html = html & "<div class='dato-label'>RESPONSABLE</div>"

	'html = html & "<td class='fil'>"
	'html = html & getName(cnn,"","cor_responsables",session("usuario"))
	'html = html & "</td>"
	'html = html & "</tr>"
	html = html & "<div class='dato-value'>" & getName(cnn,"","cor_responsables",session("usuario")) & "</div>"

	'html = html & "<tr>"
	'html = html & "<td class='fil'>FECHA DESDE</td>"
	html = html & "<div class='dato-label'>FECHA DESDE</div>"

	'html = html & "<td class='fil'><input class='inpu' style='width:80%;' id='fd' name='fd' type='text'></td>"
	'html = html & "</tr>"
	html = html & "<div><input class='input-moderno ui-input' id='fd' name='fd' type='text'></div>"

	'html = html & "<tr>"
	'html = html & "<td class='fil'>FECHA HASTA</td>"
	html = html & "<div class='dato-label'>FECHA HASTA</div>"

	'html = html & "<td class='fil'><input class='inpu' style='width:80%;' id='fh' name='fh' type='text'></td>"
	'html = html & "</tr>"
	html = html & "<div><input class='input-moderno ui-input' id='fh' name='fh' type='text'></div>"

	'html = html & "<tr>"
	'html = html & "<td class='fil'>MOTIVO</td>"
	html = html & "<div class='dato-label'>MOTIVO</div>"

	'html = html & "<td class='fil'><input class='inpu' style='width:80%;' id='mot' name='mot' type='text'></td>"
	'html = html & "</tr>"
	'html = html & "<div><input class='input-moderno ui-input' id='mot' name='mot' type='text'></div>"
	html = html & "<div><textarea class='input-moderno ui-input' id='mot' name='mot' rows='4' style='resize: vertical;'></textarea>"

	html = html & "</div>" ' Cierre de .grid-datos

	'html = html & "<tr>"
	'html = html & "<td colspan=2>&nbsp;</td>"
	'html = html & "</tr>"
	'html = html & "<hr style='border:none;'>" ' Separador visual en reemplazo del espacio vacÃ­o

	'html = html & "<tr>"
	'html = html & "<td colspan=4 align=center><input class='bot' type='button' value='Agregar' onclick='agregar3();' ></td>"
	'html = html & "</tr>"
	html = html & "<div style='grid-column: 1 / -1; display: flex; justify-content: center; margin-top: 1.5rem;'><input class='btn-principal ui-button ui-button--primary' type='button' value='Agregar' onclick='agregar3();'></div>"

	'html = html & "</table>"
	html = html & "</div>" ' Cierre de .bloque-body
	html = html & "</div>" ' Cierre de .bloque-datos

	'html = html & "</form>"
	html = html & "</form>"

	'html = html & "</div>"
	html = html & "</div>" ' Cierre de .inner

End If


html = html & "<script type='text/javascript'>"
html = html & " function agregar(){ "
html = html & " var fi = document.getElementById('tit').value;"
html = html & " if (fi == ''){ "
html = html & " alert('Titulo no debe ser nulo'); return false; "
html = html & " } "
html = html & " document.getElementById('agre').submit();"
html = html & " } "
html = html & "</script>"

html = html & "<script type='text/javascript'>"
html = html & " function agregar2(){ "
html = html & " var sp = document.getElementById('sp').value;"
html = html & " var tp = document.getElementById('tp').value;"
html = html & " if (sp == '' || tp == ''){ "
html = html & " alert('Solicitud y titulo no deben ser nulo'); return false; "
html = html & " } "
html = html & " document.getElementById('agre').submit();"
html = html & " } "
html = html & "</script>"

html = html & "<script type='text/javascript'>"
html = html & " function agregar3(){ "
html = html & " var fd = document.getElementById('fd').value;"
html = html & " var fh = document.getElementById('fh').value;"
html = html & " if (fd == '' || fh == ''){ "
html = html & " alert('Las fechas no pueden ser nulas'); return false; "
html = html & " } "
html = html & " document.getElementById('agre').submit();"
html = html & " } "
html = html & "</script>"

html = html & "<script type='text/javascript'>"
html = html & " $( function() {"
html = html & "$.datepicker.regional['es'] = {"
html = html & "closeText: 'Cerrar',"
html = html & "prevText: '< Ant',"
html = html & "nextText: 'Sig >',"
html = html & "currentText: 'Hoy',"
html = html & "monthNames: ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'],"
html = html & "monthNamesShort: ['Ene','Feb','Mar','Abr', 'May','Jun','Jul','Ago','Sep', 'Oct','Nov','Dic'],"
html = html & "dayNames: ['Domingo', 'Lunes', 'Martes', 'Miercoles', 'Jueves', 'Viernes', 'Sabado'],"
html = html & "dayNamesShort: ['Dom','Lun','Mar','Miï¿½','Juv','Vie','Sï¿½b'],"
html = html & "dayNamesMin: ['Do','Lu','Ma','Mi','Ju','Vi','Sa'],"
html = html & "weekHeader: 'Sm',"
html = html & "dateFormat: 'dd/mm/yy',"
html = html & "firstDay: 1,"
html = html & "isRTL: false,"
html = html & "showMonthAfterYear: false,"
html = html & "yearSuffix: ''"
html = html & "};"	
html = html & "$.datepicker.setDefaults($.datepicker.regional['es']);"
html = html & "});"

html = html & "$( function() { "
html = html & "$( '#fd' ).datepicker(); "
html = html & "$( '#fh' ).datepicker(); "
html = html & "$( '#fe' ).datepicker(); "
html = html & "} ); "
html = html & "</script>"


If opc = "ins" Then

	vuelta = request.querystring("volver")
	if tipo <> "INC" then
		vuelta = vuelta & "?xmenu=" & xmenu & "&xtmenu=" & xtmenu
	else
		vuelta = vuelta & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu
	end if

	html = html & "<FORM name='init' method='post'  ACTION='" & vuelta & "' >"
	html = html & "<INPUT TYPE='SUBMIT' VALUE='INICIAR' style='display:none;' >"
	html = html & "</FORM>"

	html = html & "<script type='text/javascript'>"
	html = html & " alert('Agregado exitosamente.');"
	html = html & " document.all.init.submit();"
	html = html & "</script>"

End If

html = html & form6 & form5 & "</body>"
html = html & "</html>"

response.write html
%>