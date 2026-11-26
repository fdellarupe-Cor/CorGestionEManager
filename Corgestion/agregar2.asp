<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/ui_componentes.inc" -->
<!-- #INCLUDE FILE="IncFile/datosAgregar.inc" -->
<%
' [CG] Control de usuario E-Manager (igual que CorEManager/Acciones.asp)
cgOk = ofv.CheckUsuario()
If cgOk <> "Y" Then
	cgUv = ofv.uv
	If cgUv = "NO" Then cgOk = ofv.ValidarUsuario
	Response.Write cgOk
	Response.End
End If

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
' [UI] (linea comentada con la cadena de conexion y su contrasena: eliminada por seguridad)
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
			request.form("rp"), sFec, request.form("br"), request.form("vp"), request.form("tit"), RolResponsableActual(), request.form("amb"), request.form("proy")

		XXX = SYSLOGON(cnn,"AGREGA_INC")

		xid = dameUltId("I",cnn)
		CgLog "INCIDENTE", xid, "ALTA"

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
		CgLogUltimo "LICENCIA", "CG_LICENCIAS", "ALTA"

	Else 'PROYECTOS
	
		' [UI] Resto de depuracion: escribia este texto arriba de la pagina al cargar un proyecto
		'[UI-ANT] response.write "Guardar Proyecto"
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
		CgLog "PROYECTO", xid, "ALTA"

		 vvv = objacc(xid,"P",cnn,"C")

		set fso = server.createobject("scripting.filesystemobject")
			  ruta=ruta & "proyectos\"  &   xid
			  ' [UI] Resto de depuracion: mostraba la ruta interna del servidor al cargar un proyecto
			  '[UI-ANT] response.write ruta
			   if not fso.folderexists(ruta ) then
				 fso.createfolder (ruta)
			   end if
	   set fso=nothing

	End If

End If

'[UI-ANT] ' ===== Vista anterior completa: reemplazada por el bloque [UI] que sigue. Se borra entera buscando [UI-ANT] =====
'[UI-ANT] html = html & "<div>"

'[UI-ANT] html = html & cabecera()

'[UI-ANT] If tipo = "INC" Then

	'[UI-ANT] 'html = html & "<div style='height:8%;width:100%;border-top:solid 5px #FF3D3D;border-bottom:double 6px #001BFF;'>"

	'[UI-ANT] 'html = html & "<table style='height:100%;width:100%;'>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] html = html & "<table class='titulo-seccion ui-page-header' >"
	'[UI-ANT] html = html & "<tr style='display:flex;align-items:center;justify-content:space-between;'>"

	'[UI-ANT] html = html & "<td><h2>NUEVO " & tit & "</h2></td>"

	'[UI-ANT] form5 = "<form id='form5' style='height:70%;' method='post' action='incidentes.asp?opc=listar&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&tipo=INC'></form>"
	'[UI-ANT] boton5 = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('form5').submit();"">"

	'[UI-ANT] html = html & "<td>" & boton5 & "</td>"

	'[UI-ANT] html = html & "</tr>"

	'[UI-ANT] html = html & "</table>"

	'[UI-ANT] html = html & "</div>"


	'[UI-ANT] ' Usamos tu contenedor estandar para mantener el padding global del sistema
	'[UI-ANT] html = html & "<div class='inner' style='padding:2rem'>"
	
	'[UI-ANT] 'html = html & "<div style='height:82%;width:100%;'><br>"
	'[UI-ANT] ' Le damos un max-width y margin: 0 auto para que el formulario quede centrado y no se estire al infinito en monitores grandes
	'[UI-ANT] html = html & "<div style='max-width: 1000px; margin: 0 auto;'>"

	'[UI-ANT] html = html & "<form id='agre' method='post' action='agregar2.asp?opc=ins&tabla=" & tabla & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&tipo=INC&volver=" & volver & "'>"

	'[UI-ANT] 'html = html & "<table class='tabl' align=center cellspacing=0 cellpadding=3>"
	'[UI-ANT] html = html & "<div class='bloque-datos ui-panel'>"

	'[UI-ANT] ' --- HEADER DEL FORMULARIO ---
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='titu' colspan=4><u>COMPLETE LOS SIGUIENTES DATOS</u></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='bloque-header'>COMPLETE LOS SIGUIENTES DATOS</div>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' colspan=4 align=left>&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"
	
	'[UI-ANT] ' --- BODY DEL FORMULARIO Y GRILLA DOBLE ---
	'[UI-ANT] html = html & "<div class='bloque-body ui-panel__body'>"
	'[UI-ANT] html = html & "<div class='grid-datos-doble ui-field-grid--dual'>"

	'[UI-ANT] ' --- FILA 1 ---
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil'>CLIENTE:</td>"
	'[UI-ANT] 'html = html & "<td class='fil'><select class='selec' id='cli' name='cli'>"
	'[UI-ANT] html = html & "<div class='dato-label'>CLIENTE</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='cli' name='cli'>"

	'[UI-ANT] Set rs = ObtenerClientes(Cnn)

	'[UI-ANT] Do While Not rs.eof
		'[UI-ANT] html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
		'[UI-ANT] rs.movenext
	'[UI-ANT] Loop
	'[UI-ANT] html = html & "</select></div>"

	'[UI-ANT] 'html = html & "<td class='fil'>TITULO:</td>"
	'[UI-ANT] 'html = html & "<td class='fil'><input class='inpu' id='tit' name='tit' type='text'></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>TITULO</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><input class='input-moderno ui-input' id='tit' name='tit' type='text'></div>"


	'[UI-ANT] ' --- FILA 2 ---
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil'>APLICACION:</td>"
	'[UI-ANT] 'html = html & "<td class='fil'><select class='selec' id='apli' name='apli'>"
	'[UI-ANT] html = html & "<div class='dato-label'>APLICACION</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='apli' name='apli'>"

	'[UI-ANT] Set rs = ObtenerAplicativos(Cnn)

	'[UI-ANT] Do While Not rs.eof
		'[UI-ANT] html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
		'[UI-ANT] rs.movenext
	'[UI-ANT] Loop
	'[UI-ANT] html = html & "</select></div>"

	'[UI-ANT] 'html = html & "<td class='fil'>PRIORIDAD:</font></td>"
	'[UI-ANT] 'html = html & "<td class='fil'><select class='selec' id='prio' name='prio'>"
	'[UI-ANT] html = html & "<div class='dato-label'>PRIORIDAD</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='prio' name='prio'>"

	'[UI-ANT] Set rs = ObtenerPrioridades(Cnn)

	'[UI-ANT] Do While Not rs.eof
		'[UI-ANT] if rs(1) = "MEDIA" then
			'[UI-ANT] html = html & "<option value='" & rs(0) & "' selected>" & Server.HTMLEncode(rs(1) & "") & "</option>"
		'[UI-ANT] else
			'[UI-ANT] html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
		'[UI-ANT] end if
		'[UI-ANT] rs.movenext
	'[UI-ANT] Loop
	'[UI-ANT] html = html & "</select></div>"


	'[UI-ANT] ' --- FILA 3 ---
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil'>RESP. GESTION:</td>"
	'[UI-ANT] 'html = html & "<td class='fil'><select class='selec' id='rg' name='rg'>"
	'[UI-ANT] html = html & "<div class='dato-label'>RESP. GESTION</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='rg' name='rg'>"

	'[UI-ANT] Set rs = ObtenerResponsables(Cnn, "FUNC")

	'[UI-ANT] Do While Not rs.eof
		'[UI-ANT] html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
		'[UI-ANT] rs.movenext
	'[UI-ANT] Loop
	'[UI-ANT] html = html & "</select></div>"

	'[UI-ANT] 'html = html & "<td class='fil'>RESP. TECNICO:</td>"
	'[UI-ANT] 'html = html & "<td class='fil'><select class='selec' id='rt' name='rt'>"
	'[UI-ANT] html = html & "<div class='dato-label'>RESP. TECNICO</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='rt' name='rt'>"

	'[UI-ANT] Set rs = ObtenerResponsables(Cnn, "TEC")

	'[UI-ANT] html = html & "<option value='0'>NINGUNO (CONSULTA)</option>"

	'[UI-ANT] Do While Not rs.eof
		'[UI-ANT] html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
		'[UI-ANT] rs.movenext
	'[UI-ANT] Loop
	'[UI-ANT] html = html & "</select></div>"


	'[UI-ANT] ' --- FILA 4 ---
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil'>REQ. PRESUP:</td>"
	'[UI-ANT] 'html = html & "<td class='fil'><select class='selec' id='rp' name='rp'>"
	'[UI-ANT] 'html = html & "<option value='SI'>SI</option>"
	'[UI-ANT] 'html = html & "<option value='NO'>NO</option>"
	'[UI-ANT] 'html = html & "</select></td>"
	'[UI-ANT] html = html & "<div class='dato-label'>REQ. PRESUP</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='rp' name='rp'>"
	'[UI-ANT] html = html & "<option value='SI'>SI</option>"
	'[UI-ANT] html = html & "<option value='NO'>NO</option>"
	'[UI-ANT] html = html & "</select></div>"

	'[UI-ANT] 'html = html & "<td class='fil'>SOLIC. PROYECTO:</td>"
	'[UI-ANT] 'html = html & "<td class='fil'><input class='inpu' id='sp' name='sp' type='text'></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-label'>SOLIC. PROYECTO</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><input class='input-moderno ui-input' id='sp' name='sp' type='text'></div>"


	'[UI-ANT] ' --- FILA 5 ---
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil'>VIA PEDIDO:</td>"
	'[UI-ANT] 'html = html & "<td class='fil'><select class='selec' id='vp' name='vp'>"
	'[UI-ANT] html = html & "<div class='dato-label'>VIA PEDIDO</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='vp' name='vp'>"

	'[UI-ANT] Set rs = ObtenerViaPedido(Cnn)

	'[UI-ANT] Do While Not rs.eof
		'[UI-ANT] html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
		'[UI-ANT] rs.movenext
	'[UI-ANT] Loop
	'[UI-ANT] html = html & "</select></div>"

	'[UI-ANT] 'html = html & "<td class='fil'>CONTACTO:</td>"
	'[UI-ANT] 'html = html & "<td class='fil'><select class='selec' id='con' name='con'>"
	'[UI-ANT] html = html & "<div class='dato-label'>CONTACTO</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='con' name='con'>"

	'[UI-ANT] Set rs = ObtenerContactosParaIncidente(Cnn)

	'[UI-ANT] Do While Not rs.eof
		'[UI-ANT] html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
	'[UI-ANT] rs.movenext
	'[UI-ANT] Loop
	'[UI-ANT] html = html & "</select></div>"


	'[UI-ANT] ' --- FILA 6 ---
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil'>BROWSER:</td>"
	'[UI-ANT] 'html = html & "<td class='fil'><select class='selec' id='br' name='br'>"
	'[UI-ANT] html = html & "<div class='dato-label'>BROWSER</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='br' name='br'>"

	'[UI-ANT] Set rs = ObtenerBrowsers(Cnn)

	'[UI-ANT] Do While Not rs.eof
		'[UI-ANT] html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
		'[UI-ANT] rs.movenext
	'[UI-ANT] Loop
	'[UI-ANT] html = html & "</select></div>"

	'[UI-ANT] ' Completamos la grilla para que no se desfase (antes eran <td> vacios)
	'[UI-ANT] 'html = html & "<td></td>"
	'[UI-ANT] 'html = html & "<td></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div></div><div></div>"
	
	'[UI-ANT] ' A!CERRAMOS LA GRILLA DE 4 COLUMNAS ACA?!
	'[UI-ANT] html = html & "</div>"


	'[UI-ANT] ' --- SECCION DETALLES (FUERA DE LA GRILLA) ---
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' colspan=4 align=left>&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' colspan=4 >DETALLES:</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td style='height:150px;' colspan=4><textarea class='txtar' id='det' name='det'></textarea></td>"
	'[UI-ANT] 'html = html & "</tr>"
	
	'[UI-ANT] ' Agregamos un div contenedor con margen superior para separar "DETALLES" de la grilla de arriba
	'[UI-ANT] html = html & "<div style='margin-top: 1.5rem;'>"
	'[UI-ANT] html = html & "<div class='dato-label' style='margin-bottom: 0.5rem;'>DETALLES</div>"
	'[UI-ANT] html = html & "<textarea class='input-moderno ui-input' id='det' name='det' style='height:150px; resize:vertical; font-family: inherit;'></textarea>"
	'[UI-ANT] html = html & "</div>"


	'[UI-ANT] ' --- SECCION BOTON ---
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td colspan=4>&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td colspan=4 align=center><input type='button' class='bot' style='height:25px;' value='Agregar' onclick='agregar();' ></td>"
	'[UI-ANT] 'html = html & "</tr>"
	
	'[UI-ANT] ' Contenedor del boton alineado al centro y con estilo moderno
	'[UI-ANT] html = html & "<div style='margin-top: 2rem; text-align: center; border-top: 1px solid #e2e8f0; padding-top: 1.5rem;'>"
	'[UI-ANT] html = html & "<input class='ui-button ui-button--primary' type='button' style='padding: 0.6rem 2rem; background-color: #3b82f6; color: white; border: none; border-radius: 6px; font-weight: 600; cursor: pointer; transition: all 0.2s ease;' value='Agregar' onclick='agregar();' onmouseover=""this.style.backgroundColor='#2563eb'"" onmouseout=""this.style.backgroundColor='#3b82f6'"" onmousedown=""this.style.transform='scale(0.95)'"" onmouseup=""this.style.transform='scale(1)'"">"
	'[UI-ANT] html = html & "</div>"

	'[UI-ANT] 'html = html & "</table>"

	'[UI-ANT] ' Cerramos los contenedores de la tarjeta moderna
	'[UI-ANT] html = html & "</div>" ' Cierra bloque-body
	'[UI-ANT] html = html & "</div>" ' Cierra bloque-datos

	'[UI-ANT] html = html & "</form>"
	'[UI-ANT] html = html & "</div>" ' Cierra el div wrapper principal
	'[UI-ANT] html = html & "</div>"

'[UI-ANT] ElseIf tipo = "PROY" Then



	'[UI-ANT] 'html = html & "<div style='height:8%;width:70%;border-top:solid 5px #FF3D3D;border-bottom:double 6px #001BFF;margin:0 auto;'>"
	'[UI-ANT] 'html = html & "<table style='height:100%;width:100%;'>"
	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] html = html & "<table class='titulo-seccion ui-page-header' >"
	'[UI-ANT] html = html & "<tr style='display:flex;align-items:center;justify-content:space-between;'>"

	'[UI-ANT] html = html & "<td><h2>NUEVO " & tit & "</h2></td>"

	'[UI-ANT] form6 = "<form id='form6' style='height:70%;' method='post' action='proyectos.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
	'[UI-ANT] boton6 = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('form6').submit();"">"

	'[UI-ANT] html = html & "<td>" & boton6 & "</td>"

	'[UI-ANT] html = html & "</tr>"
	'[UI-ANT] html = html & "</table>"
	'[UI-ANT] html = html & "</div>"
	
	'[UI-ANT] ' ==========================================
	'[UI-ANT] ' FORMULARIO DE ALTA DE PROYECTOS
	'[UI-ANT] ' ==========================================
	
	'[UI-ANT] ' Usamos el contenedor estandar con padding
	'[UI-ANT] html = html & "<div class='inner' style='padding:2rem'>"
	
	'[UI-ANT] ' Limitamos el ancho para no deformar en monitores grandes
	'[UI-ANT] html = html & "<div style='max-width: 1000px; margin: 0 auto;'>"

	'[UI-ANT] html = html & "<form id='agre' method='post' action='agregar2.asp?opc=ins&tabla=" & tabla & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&volver=" & volver & "'>"
	'[UI-ANT] html = html & "<INPUT TYPE=HIDDEN NAME=ID_CLI>"

	'[UI-ANT] ' --- TARJETA PRINCIPAL ---
	'[UI-ANT] html = html & "<div class='bloque-datos ui-panel'>"

	'[UI-ANT] ' --- HEADER ---
	'[UI-ANT] 'html = html & "<td class='titu' colspan=4 align=center style='font-family:helvetica;font-size:11pt;font-weight:bold;'><br>COMPLETE LOS SIGUIENTES DATOS</td>"
	'[UI-ANT] html = html & "<div class='bloque-header'>COMPLETE LOS SIGUIENTES DATOS</div>"

	'[UI-ANT] ' --- BODY Y GRILLA DE CAMPOS ---
	'[UI-ANT] html = html & "<div class='bloque-body ui-panel__body'>"
	'[UI-ANT] html = html & "<div class='grid-datos-doble ui-field-grid--dual'>"

	'[UI-ANT] ' --- FILA 1 ---
	'[UI-ANT] 'html = html & "<td class='fil' style='" & st1 & "' >SOLIC. PROYECTO NRO:</td>"
	'[UI-ANT] 'html = html & "<td class='fil'><input class='inpu' style='" & st0 & "' id='sp' name='sp' type='text'></td>"
	'[UI-ANT] html = html & "<div class='dato-label'>SOLIC. PROYECTO NRO</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><input class='input-moderno ui-input' id='sp' name='sp' type='text'></div>"

	'[UI-ANT] 'html = html & "<td class='fil' style='" & st1 & "' >TITULO PROYECTO:</td>"
	'[UI-ANT] 'html = html & "<td class='fil'><input class='inpu' style='" & st0 & "' id='tp' name='tp' type='text' autocomplete='off'></td>"
	'[UI-ANT] html = html & "<div class='dato-label'>TITULO PROYECTO</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><input class='input-moderno ui-input' id='tp' name='tp' type='text' autocomplete='off'></div>"

	'[UI-ANT] ' --- FILA 2 ---
	'[UI-ANT] 'html = html & "<td class='fil' style='" & st1 & "' >CLIENTE:</td>"
	'[UI-ANT] 'html = html & "<td class='fil'><select class='selec' style='" & st0 & "'  id='cli' name='cli' ONCHANGE='VerSel()'>"
	'[UI-ANT] html = html & "<div class='dato-label'>CLIENTE</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='cli' name='cli' onchange='VerSel()'>"

	'[UI-ANT] ' Mantenemos el SCRIPT intacto
	'[UI-ANT] html = html & "<SCRIPT>"
	'[UI-ANT] html = html & " function VerSel() { "
	'[UI-ANT] html = html & " var a; "
	'[UI-ANT] html = html & " a = document.getElementById('cli'); "
	'[UI-ANT] html = html & " var v; "
	'[UI-ANT] html = html & " v = a.value; "
	'[UI-ANT] html = html & " console.log(v); "
	'[UI-ANT] html = html & " return v; }"
	'[UI-ANT] html = html & "</SCRIPT>"

	'[UI-ANT] Set rs = ObtenerClientes(Cnn)
		'[UI-ANT] html = html & "<option value='0' >" & sel & "</option>"
	'[UI-ANT] Do While Not rs.eof
		'[UI-ANT] html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
		'[UI-ANT] rs.movenext
	'[UI-ANT] Loop
	'[UI-ANT] html = html & "</select></div>"

	'[UI-ANT] 'html = html & "<td class='fil' style='" & st1 & "' >CONTACTO:</td>"
	'[UI-ANT] 'html = html & "<select class='selec' style='" & st0 & "' id='con' name='con'>"
	'[UI-ANT] html = html & "<div class='dato-label'>CONTACTO</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='con' name='con'>"

	'[UI-ANT] Set rs = ObtenerContactosParaProyecto(Cnn)
		'[UI-ANT] html = html & "<option value='0' >" & sel & "</option>"
	'[UI-ANT] Do While Not rs.eof
		'[UI-ANT] html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
		'[UI-ANT] rs.movenext
	'[UI-ANT] Loop
	'[UI-ANT] html = html & "</select></div>"

	'[UI-ANT] ' --- FILA 3 ---
	'[UI-ANT] 'html = html & "<td class='fil' style='" & st1 & "' >APLICACION:</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>APLICACION</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='apli' name='apli'>"

	'[UI-ANT] Set rs = ObtenerAplicativos(Cnn)
		'[UI-ANT] html = html & "<option value='0' >" & sel & "</option>"
	'[UI-ANT] Do While Not rs.eof
		'[UI-ANT] html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
		'[UI-ANT] rs.movenext
	'[UI-ANT] Loop
	'[UI-ANT] html = html & "</select></div>"

	'[UI-ANT] 'html = html & "<td class='fil' style='" & st1 & "' >PRIORIDAD:</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>PRIORIDAD</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='pri' name='pri'>"

	'[UI-ANT] Set rs = ObtenerPrioridades(Cnn)
		'[UI-ANT] html = html & "<option value='0' >" & sel & "</option>"
	'[UI-ANT] Do While Not rs.eof
		'[UI-ANT] html = html & "<option value='" & rs(0) & "' >" & Server.HTMLEncode(rs(1) & "") & "</option>"
		'[UI-ANT] rs.movenext
	'[UI-ANT] Loop
	'[UI-ANT] html = html & "</select></div>"

	'[UI-ANT] ' --- FILA 4 ---
	'[UI-ANT] 'html = html & "<td class='fil' style='" & st1 & "' >RESP. GESTION:</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>RESP. GESTION</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='rg' name='rg'>"

	'[UI-ANT] Set rs = ObtenerResponsables(Cnn, "FUNC")
		'[UI-ANT] html = html & "<option value='0' >" & sel & "</option>"
	'[UI-ANT] Do While Not rs.eof
		'[UI-ANT] html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & " " & rs(2) & "") & "</option>"
		'[UI-ANT] rs.movenext
	'[UI-ANT] Loop
	'[UI-ANT] html = html & "</select></div>"

	'[UI-ANT] 'html = html & "<td class='fil' style='" & st1 & "' >RESP. TECNICO:</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>RESP. TECNICO</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='rt' name='rt'>"

	'[UI-ANT] Set rs = ObtenerResponsables(Cnn, "TEC")
		'[UI-ANT] html = html & "<option value='0' >" & sel & "</option>"
	'[UI-ANT] Do While Not rs.eof
		'[UI-ANT] html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & " " & rs(2) & "") & "</option>"
		'[UI-ANT] rs.movenext
	'[UI-ANT] Loop
	'[UI-ANT] html = html & "</select></div>"

	'[UI-ANT] ' --- FILA 5 ---
	'[UI-ANT] 'html = html & "<td class='fil' style='" & st1 & "' >REQ. PRESUP:</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>REQ. PRESUP</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><select class='input-moderno ui-input' style='width: 100px;' id='rq' name='rq'>"
	'[UI-ANT] html = html & "<option value='SI'>SI</option>"
	'[UI-ANT] html = html & "<option value='NO'>NO</option>"
	'[UI-ANT] html = html & "</select></div>"

	'[UI-ANT] 'html = html & "<td class='fil' style='" & st1 & "' >BROWSER:</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>BROWSER</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='br' name='br'>"

	'[UI-ANT] Set rs = ObtenerBrowsers(Cnn)
		'[UI-ANT] html = html & "<option value='0' >" & sel & "</option>"
	'[UI-ANT] Do While Not rs.eof
		'[UI-ANT] html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
	'[UI-ANT] rs.movenext
	'[UI-ANT] Loop
	'[UI-ANT] html = html & "</select></div>"

	'[UI-ANT] ' --- FILA 6 ---
	'[UI-ANT] 'html = html & "<td class='fil' style='" & st1 & "' >VIA PEDIDO:</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>VIA PEDIDO</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><select class='input-moderno ui-input' id='vp' name='vp'>"

	'[UI-ANT] Set rs = ObtenerViaPedido(Cnn)
		'[UI-ANT] html = html & "<option value='0' >" & sel & "</option>"
	'[UI-ANT] Do While Not rs.eof
		'[UI-ANT] html = html & "<option value='" & rs(0) & "'>" & Server.HTMLEncode(rs(1) & "") & "</option>"
		'[UI-ANT] rs.movenext
	'[UI-ANT] Loop
	'[UI-ANT] html = html & "</select></div>"

	'[UI-ANT] 'html = html & "<td class='fil' style='" & st1 & "' >FECHA ENTREGA:</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>FECHA ENTREGA</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><input class='input-moderno ui-input' id='fe' name='fe' type='text'></div>"

	'[UI-ANT] ' --- FILA 7 ---
	'[UI-ANT] 'html = html & "<td class='fil' style='" & st1 & "' >RELEV Y DEFINICION (hs):</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>RELEV Y DEFINICION (hs)</div>"
	'[UI-ANT] ' A los campos de horas les limitamos un poco el ancho para que no queden como una barra gigante
	'[UI-ANT] html = html & "<div class='dato-value'><input class='input-moderno ui-input' style='max-width: 120px;' id='rd' name='rd' type='text'></div>"

	'[UI-ANT] 'html = html & "<td class='fil' style='" & st1 & "' >DESARROLLO (hs):</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>DESARROLLO (hs)</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><input class='input-moderno ui-input' style='max-width: 120px;' id='de' name='de' type='text'></div>"

	'[UI-ANT] ' --- FILA 8 ---
	'[UI-ANT] 'html = html & "<td class='fil' style='" & st1 & "' >DISE&Ntilde;O (hs):</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>DISE&Ntilde;O (hs)</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><input class='input-moderno ui-input' style='max-width: 120px;' id='di' name='di' type='text'></div>"

	'[UI-ANT] 'html = html & "<td class='fil' style='" & st1 & "' >TEST (hs):</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>TEST (hs)</div>"
	'[UI-ANT] html = html & "<div class='dato-value'><input class='input-moderno ui-input' style='max-width: 120px;' id='te' name='te' type='text'></div>"

	'[UI-ANT] ' Cerramos la grilla
	'[UI-ANT] html = html & "</div>"

	'[UI-ANT] ' --- TEXTAREA: DESCRIPCION DEL PROYECTO (Fuera de la grilla) ---
	'[UI-ANT] html = html & "<div style='margin-top: 1.5rem;'>"
	'[UI-ANT] html = html & "<div class='dato-label' style='margin-bottom: 0.5rem;'>DESCRIPCION DEL PROYECTO</div>"
	'[UI-ANT] html = html & "<textarea class='input-moderno ui-input' id='det' name='det' style='height: 300px; resize: vertical; font-family: inherit;'></textarea>"
	'[UI-ANT] html = html & "</div>"

	'[UI-ANT] ' --- BOTON AGREGAR ---
	'[UI-ANT] html = html & "<div style='margin-top: 2rem; text-align: center; border-top: 1px solid #e2e8f0; padding-top: 1.5rem;'>"
	'[UI-ANT] html = html & "<input class='ui-button ui-button--primary' type='button' style='padding: 0.6rem 2rem; background-color: #3b82f6; color: white; border: none; border-radius: 6px; font-weight: 600; cursor: pointer; transition: all 0.2s ease;' value='Agregar' onclick='agregar2();' onmouseover=""this.style.backgroundColor='#2563eb'"" onmouseout=""this.style.backgroundColor='#3b82f6'"" onmousedown=""this.style.transform='scale(0.95)'"" onmouseup=""this.style.transform='scale(1)'"">"
	'[UI-ANT] html = html & "</div>"

	'[UI-ANT] ' Cerramos todos los contenedores
	'[UI-ANT] html = html & "</div>" ' Cierra bloque-body
	'[UI-ANT] html = html & "</div>" ' Cierra bloque-datos
	'[UI-ANT] html = html & "</form>"
	'[UI-ANT] html = html & "</div>" ' Cierra div del form centrado
	'[UI-ANT] html = html & "</div>" ' Cierra div inner

'[UI-ANT] Elseif tipo = "LIC" then

	'[UI-ANT] html = html & "<table class='titulo-seccion ui-page-header' >"
	'[UI-ANT] html = html & "<tr style='display:flex;align-items:center;justify-content:space-between;'>"

	'[UI-ANT] html = html & "<td><h2>NUEVA " & tit & "</h2></td>"

	'[UI-ANT] form5 = "<form id='form5' style='height:70%;' method='post' action='licencias.asp?" & request.querystring & "'></form>"
	'[UI-ANT] boton5 = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('form5').submit();"">"

	'[UI-ANT] html = html & "<td>" & boton5 & "</td>"

	'[UI-ANT] html = html & "</tr>"

	'[UI-ANT] html = html & "</table>"

	'[UI-ANT] html = html & "</div>"

	'[UI-ANT] ' --- INICIO NORMALIZACION: FORMULARIO ALTA DE LICENCIAS ---
	
	'[UI-ANT] 'html = html & "<div style='height:120%;width:100%;'><br>"
	'[UI-ANT] html = html & "<div class='inner' style='padding: 2rem;'>" ' Usamos un contenedor con aire

	'[UI-ANT] 'html = html & "<form id='agre' method='post' action='agregar2.asp?opc=ins&tabla=" & tabla & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&volver=" & volver & "'>"
	'[UI-ANT] html = html & "<form id='agre' method='post' action='agregar2.asp?opc=ins&tabla=" & tabla & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&volver=" & volver & "'>"

	'[UI-ANT] 'html = html & "<table class='tabl' style='width:40%;' cellspacing=0 cellpadding=3 align=center >"
	'[UI-ANT] html = html & "<div class='bloque-datos ui-panel' style='max-width: 600px; margin: 0 auto;'>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil' style='font-size:12pt;' colspan=2 align=center><b><u>COMPLETE LOS SIGUIENTES DATOS</u></b></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='bloque-header' style='text-align: center; text-transform: uppercase;'>COMPLETE LOS SIGUIENTES DATOS</div>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td colspan=2>&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='bloque-body ui-panel__body'>"
	
	'[UI-ANT] ' Abrimos la grilla moderna para el formulario
	'[UI-ANT] html = html & "<div class='grid-datos ui-field-grid'>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil'>RESPONSBALE</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>RESPONSABLE</div>"

	'[UI-ANT] 'html = html & "<td class='fil'>"
	'[UI-ANT] 'html = html & getName(cnn,"","cor_responsables",session("usuario"))
	'[UI-ANT] 'html = html & "</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div class='dato-value'>" & getName(cnn,"","cor_responsables",session("usuario")) & "</div>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil'>FECHA DESDE</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>FECHA DESDE</div>"

	'[UI-ANT] 'html = html & "<td class='fil'><input class='inpu' style='width:80%;' id='fd' name='fd' type='text'></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div><input class='input-moderno ui-input' id='fd' name='fd' type='text'></div>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil'>FECHA HASTA</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>FECHA HASTA</div>"

	'[UI-ANT] 'html = html & "<td class='fil'><input class='inpu' style='width:80%;' id='fh' name='fh' type='text'></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div><input class='input-moderno ui-input' id='fh' name='fh' type='text'></div>"

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td class='fil'>MOTIVO</td>"
	'[UI-ANT] html = html & "<div class='dato-label'>MOTIVO</div>"

	'[UI-ANT] 'html = html & "<td class='fil'><input class='inpu' style='width:80%;' id='mot' name='mot' type='text'></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] 'html = html & "<div><input class='input-moderno ui-input' id='mot' name='mot' type='text'></div>"
	'[UI-ANT] html = html & "<div><textarea class='input-moderno ui-input' id='mot' name='mot' rows='4' style='resize: vertical;'></textarea>"

	'[UI-ANT] html = html & "</div>" ' Cierre de .grid-datos

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td colspan=2>&nbsp;</td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] 'html = html & "<hr style='border:none;'>" ' Separador visual en reemplazo del espacio vacio

	'[UI-ANT] 'html = html & "<tr>"
	'[UI-ANT] 'html = html & "<td colspan=4 align=center><input class='bot' type='button' value='Agregar' onclick='agregar3();' ></td>"
	'[UI-ANT] 'html = html & "</tr>"
	'[UI-ANT] html = html & "<div style='grid-column: 1 / -1; display: flex; justify-content: center; margin-top: 1.5rem;'><input class='btn-principal ui-button ui-button--primary' type='button' value='Agregar' onclick='agregar3();'></div>"

	'[UI-ANT] 'html = html & "</table>"
	'[UI-ANT] html = html & "</div>" ' Cierre de .bloque-body
	'[UI-ANT] html = html & "</div>" ' Cierre de .bloque-datos

	'[UI-ANT] 'html = html & "</form>"
	'[UI-ANT] html = html & "</form>"

	'[UI-ANT] 'html = html & "</div>"
	'[UI-ANT] html = html & "</div>" ' Cierre de .inner

'[UI-ANT] End If


'[UI-ANT] html = html & "<script type='text/javascript'>"
'[UI-ANT] html = html & " function agregar(){ "
'[UI-ANT] html = html & " var fi = document.getElementById('tit').value;"
'[UI-ANT] html = html & " if (fi == ''){ "
'[UI-ANT] html = html & " alert('Titulo no debe ser nulo'); return false; "
'[UI-ANT] html = html & " } "
'[UI-ANT] html = html & " document.getElementById('agre').submit();"
'[UI-ANT] html = html & " } "
'[UI-ANT] html = html & "</script>"

'[UI-ANT] html = html & "<script type='text/javascript'>"
'[UI-ANT] html = html & " function agregar2(){ "
'[UI-ANT] html = html & " var sp = document.getElementById('sp').value;"
'[UI-ANT] html = html & " var tp = document.getElementById('tp').value;"
'[UI-ANT] html = html & " if (sp == '' || tp == ''){ "
'[UI-ANT] html = html & " alert('Solicitud y titulo no deben ser nulo'); return false; "
'[UI-ANT] html = html & " } "
'[UI-ANT] html = html & " document.getElementById('agre').submit();"
'[UI-ANT] html = html & " } "
'[UI-ANT] html = html & "</script>"

'[UI-ANT] html = html & "<script type='text/javascript'>"
'[UI-ANT] html = html & " function agregar3(){ "
'[UI-ANT] html = html & " var fd = document.getElementById('fd').value;"
'[UI-ANT] html = html & " var fh = document.getElementById('fh').value;"
'[UI-ANT] html = html & " if (fd == '' || fh == ''){ "
'[UI-ANT] html = html & " alert('Las fechas no pueden ser nulas'); return false; "
'[UI-ANT] html = html & " } "
'[UI-ANT] html = html & " document.getElementById('agre').submit();"
'[UI-ANT] html = html & " } "
'[UI-ANT] html = html & "</script>"

'[UI-ANT] html = html & "<script type='text/javascript'>"
'[UI-ANT] html = html & " $( function() {"
'[UI-ANT] html = html & "$.datepicker.regional['es'] = {"
'[UI-ANT] html = html & "closeText: 'Cerrar',"
'[UI-ANT] html = html & "prevText: '< Ant',"
'[UI-ANT] html = html & "nextText: 'Sig >',"
'[UI-ANT] html = html & "currentText: 'Hoy',"
'[UI-ANT] html = html & "monthNames: ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'],"
'[UI-ANT] html = html & "monthNamesShort: ['Ene','Feb','Mar','Abr', 'May','Jun','Jul','Ago','Sep', 'Oct','Nov','Dic'],"
'[UI-ANT] html = html & "dayNames: ['Domingo', 'Lunes', 'Martes', 'Miercoles', 'Jueves', 'Viernes', 'Sabado'],"
'[UI-ANT] html = html & "dayNamesShort: ['Dom','Lun','Mar','Mi?','Juv','Vie','S?b'],"
'[UI-ANT] html = html & "dayNamesMin: ['Do','Lu','Ma','Mi','Ju','Vi','Sa'],"
'[UI-ANT] html = html & "weekHeader: 'Sm',"
'[UI-ANT] html = html & "dateFormat: 'dd/mm/yy',"
'[UI-ANT] html = html & "firstDay: 1,"
'[UI-ANT] html = html & "isRTL: false,"
'[UI-ANT] html = html & "showMonthAfterYear: false,"
'[UI-ANT] html = html & "yearSuffix: ''"
'[UI-ANT] html = html & "};"
'[UI-ANT] html = html & "$.datepicker.setDefaults($.datepicker.regional['es']);"
'[UI-ANT] html = html & "});"

'[UI-ANT] html = html & "$( function() { "
'[UI-ANT] html = html & "$( '#fd' ).datepicker(); "
'[UI-ANT] html = html & "$( '#fh' ).datepicker(); "
'[UI-ANT] html = html & "$( '#fe' ).datepicker(); "
'[UI-ANT] html = html & "} ); "
'[UI-ANT] html = html & "</script>"

' =========================================================================
' [UI] VISTA con el sistema visual (IncFile/ui_componentes.inc)
' El bloque anterior completo queda arriba, comentado con '[UI-ANT].
' Ids y nombres de los campos, formularios y validaciones: los mismos de antes.
' =========================================================================
Dim accTop, rsSel, htmlSel

html = html & "<div>"
html = html & cabecera()
html = html & "</div>"

If tipo = "INC" Then

	form5 = "<form id='form5' class='form-oculto' method='post' action='incidentes.asp?opc=listar&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&tipo=INC'></form>"
	accTop = UiBoton("Volver", "secondary", "document.getElementById('form5').submit();", True, "")
	html = html & UiPantallaAbre("Nuevo incidente", accTop, "")

	html = html & "<form id='agre' method='post' action='agregar2.asp?opc=ins&tabla=" & tabla & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&tipo=INC&volver=" & volver & "'>"
	html = html & UiCardAbre("Datos del incidente", "") & UiNotaObligatorios() & UiFormGridAbre()

	html = html & UiCampo("tit", "T&iacute;tulo", UiInputTexto("tit", True, "autocomplete='off'"), True, True)
	html = html & UiCampo("cli", "Cliente", UiSelectRs("cli", ObtenerClientes(Cnn), "", "", "", ""), False, False)
	html = html & UiCampo("apli", "Aplicaci&oacute;n", UiSelectRs("apli", ObtenerAplicativos(Cnn), "", "", "", ""), False, False)
	html = html & UiCampo("amb", "Ambiente", AmbienteSelectHtml("amb", "", True), False, False)
	' Proyecto relacionado: opcional; los proyectos se filtran segun el cliente elegido
	html = html & UiCampo("proy", "Proyecto relacionado", ProyectoComboHtml(Cnn), False, False)
	html = html & UiCampo("prio", "Prioridad", UiSelectRs("prio", ObtenerPrioridades(Cnn), "MEDIA", "", "", ""), False, False)
	html = html & UiCampo("con", "Contacto", UiSelectRs("con", ObtenerContactosParaIncidente(Cnn), "", "", "", ""), False, False)
	html = html & UiCampo("rg", "Responsable de gesti&oacute;n", UiSelectRs("rg", ObtenerResponsables(Cnn, "FUNC"), "", "", "", ""), False, False)
	html = html & UiCampo("rt", "Responsable t&eacute;cnico", UiSelectRs("rt", ObtenerResponsables(Cnn, "TEC"), "", "Ninguno (consulta)", "0", ""), False, False)
	html = html & UiCampo("vp", "V&iacute;a de pedido", UiSelectRs("vp", ObtenerViaPedido(Cnn), "", "", "", ""), False, False)
	html = html & UiCampo("br", "Browser", UiSelectRs("br", ObtenerBrowsers(Cnn), "", "", "", ""), False, False)
	html = html & UiCampo("rp", "Requiere presupuesto", UiSelectOpciones("rp", "SI|S&iacute;||NO|No", "SI", ""), False, False)
	' "Solicitud de proyecto" se quito de este formulario: InsertarIncidente no la recibe y el dato se perdia
	html = html & UiCampo("det", "Detalle", UiTextarea("det", 7, False), False, True)

	html = html & UiFormGridCierra() & UiCardCierra()
	html = html & UiFormAcciones(UiBotonPrincipal("Agregar incidente", "agregar();"))
	html = html & "</form>"
	html = html & UiPantallaCierra()

ElseIf tipo = "PROY" Then

	form6 = "<form id='form6' class='form-oculto' method='post' action='proyectos.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
	accTop = UiBoton("Volver", "secondary", "document.getElementById('form6').submit();", True, "")
	html = html & UiPantallaAbre("Nuevo proyecto", accTop, "")

	html = html & "<form id='agre' method='post' action='agregar2.asp?opc=ins&tabla=" & tabla & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&volver=" & volver & "'>"
	html = html & "<INPUT TYPE=HIDDEN NAME=ID_CLI>"
	html = html & UiCardAbre("Datos del proyecto", "") & UiNotaObligatorios() & UiFormGridAbre()

	html = html & UiCampo("tp", "T&iacute;tulo del proyecto", UiInputTexto("tp", True, "autocomplete='off'"), True, True)
	html = html & UiCampo("sp", "N&uacute;mero de solicitud", UiInputTexto("sp", True, "autocomplete='off'"), True, False)
	html = html & UiCampo("cli", "Cliente", UiSelectRs("cli", ObtenerClientes(Cnn), "", "Seleccionar", "0", "onchange='VerSel()'"), False, False)
	html = html & UiCampo("con", "Contacto", UiSelectRs("con", ObtenerContactosParaProyecto(Cnn), "", "Seleccionar", "0", ""), False, False)
	html = html & UiCampo("apli", "Aplicaci&oacute;n", UiSelectRs("apli", ObtenerAplicativos(Cnn), "", "Seleccionar", "0", ""), False, False)
	html = html & UiCampo("pri", "Prioridad", UiSelectRs("pri", ObtenerPrioridades(Cnn), "", "Seleccionar", "0", ""), False, False)

	' Responsables: nombre y apellido (rs(1) y rs(2)), por eso no usan UiSelectRs
	Set rsSel = ObtenerResponsables(Cnn, "FUNC")
	htmlSel = "<select class='ui-select' id='rg' name='rg'><option value='0'>Seleccionar</option>"
	Do While Not rsSel.eof
		htmlSel = htmlSel & "<option value='" & rsSel(0) & "'>" & Server.HTMLEncode(rsSel(1) & " " & rsSel(2) & "") & "</option>"
		rsSel.movenext
	Loop
	html = html & UiCampo("rg", "Responsable de gesti&oacute;n", htmlSel & "</select>", False, False)

	Set rsSel = ObtenerResponsables(Cnn, "TEC")
	htmlSel = "<select class='ui-select' id='rt' name='rt'><option value='0'>Seleccionar</option>"
	Do While Not rsSel.eof
		htmlSel = htmlSel & "<option value='" & rsSel(0) & "'>" & Server.HTMLEncode(rsSel(1) & " " & rsSel(2) & "") & "</option>"
		rsSel.movenext
	Loop
	html = html & UiCampo("rt", "Responsable t&eacute;cnico", htmlSel & "</select>", False, False)

	html = html & UiCampo("vp", "V&iacute;a de pedido", UiSelectRs("vp", ObtenerViaPedido(Cnn), "", "Seleccionar", "0", ""), False, False)
	html = html & UiCampo("br", "Browser", UiSelectRs("br", ObtenerBrowsers(Cnn), "", "Seleccionar", "0", ""), False, False)
	html = html & UiCampo("rq", "Requiere presupuesto", UiSelectOpciones("rq", "SI|S&iacute;||NO|No", "SI", ""), False, False)
	html = html & UiCampo("fe", "Fecha de entrega", UiInputTexto("fe", False, "autocomplete='off' placeholder='dd/mm/aaaa'"), False, False)

	html = html & UiGrupoAbre("Horas estimadas")
	html = html & UiCampo("rd", "Relevamiento y definici&oacute;n", UiInputTexto("rd", False, "inputmode='numeric'"), False, False)
	html = html & UiCampo("di", "Dise&ntilde;o", UiInputTexto("di", False, "inputmode='numeric'"), False, False)
	html = html & UiCampo("de", "Desarrollo", UiInputTexto("de", False, "inputmode='numeric'"), False, False)
	html = html & UiCampo("te", "Testeo", UiInputTexto("te", False, "inputmode='numeric'"), False, False)
	html = html & UiGrupoCierra()

	html = html & UiCampo("det", "Descripci&oacute;n del proyecto", UiTextarea("det", 10, False), False, True)

	html = html & UiFormGridCierra() & UiCardCierra()
	html = html & UiFormAcciones(UiBotonPrincipal("Agregar proyecto", "agregar2();"))
	html = html & "</form>"
	html = html & UiPantallaCierra()

ElseIf tipo = "LIC" Then

	form5 = "<form id='form5' class='form-oculto' method='post' action='licencias.asp?" & request.querystring & "'></form>"
	accTop = UiBoton("Volver", "secondary", "document.getElementById('form5').submit();", True, "")
	html = html & UiPantallaAbre("Nueva licencia", accTop, "narrow")

	html = html & "<form id='agre' method='post' action='agregar2.asp?opc=ins&tabla=" & tabla & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&volver=" & volver & "'>"
	html = html & UiCardAbre("Datos de la licencia", "") & UiNotaObligatorios() & UiFormGridAbre()

	html = html & UiCampoLectura("Responsable", getName(cnn,"","cor_responsables",session("usuario")), True)
	html = html & UiCampo("fd", "Desde", UiInputTexto("fd", True, "autocomplete='off' placeholder='dd/mm/aaaa'"), True, False)
	html = html & UiCampo("fh", "Hasta", UiInputTexto("fh", True, "autocomplete='off' placeholder='dd/mm/aaaa'"), True, False)
	html = html & UiCampo("mot", "Motivo", UiTextarea("mot", 4, False), False, True)

	html = html & UiFormGridCierra() & UiCardCierra()
	html = html & UiFormAcciones(UiBotonPrincipal("Agregar licencia", "agregar3();"))
	html = html & "</form>"
	html = html & UiPantallaCierra()

End If

' Antes estaba dentro del select de cliente; mismo comportamiento
html = html & "<script type='text/javascript'>"
html = html & " function VerSel() { var a = document.getElementById('cli'); var v = a.value; console.log(v); return v; }"
html = html & "</script>"

html = html & "<script type='text/javascript'>"
html = html & " function agregar(){ "
html = html & " var fi = document.getElementById('tit');"
html = html & " if (fi.value == ''){ "
html = html & " alert('Complet\u00e1 el t\u00edtulo.'); fi.focus(); return false; "
html = html & " } "
html = html & " document.getElementById('agre').submit();"
html = html & " } "
html = html & "</script>"

html = html & "<script type='text/javascript'>"
html = html & " function agregar2(){ "
html = html & " var sp = document.getElementById('sp');"
html = html & " var tp = document.getElementById('tp');"
html = html & " if (sp.value == '' || tp.value == ''){ "
html = html & " alert('Complet\u00e1 el n\u00famero de solicitud y el t\u00edtulo.'); (tp.value == '' ? tp : sp).focus(); return false; "
html = html & " } "
html = html & " document.getElementById('agre').submit();"
html = html & " } "
html = html & "</script>"

html = html & "<script type='text/javascript'>"
html = html & " function agregar3(){ "
html = html & " var fd = document.getElementById('fd');"
html = html & " var fh = document.getElementById('fh');"
html = html & " if (fd.value == '' || fh.value == ''){ "
html = html & " alert('Complet\u00e1 las fechas desde y hasta.'); (fd.value == '' ? fd : fh).focus(); return false; "
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
html = html & "dayNames: ['Domingo', 'Lunes', 'Martes', 'Mi\u00e9rcoles', 'Jueves', 'Viernes', 'S\u00e1bado'],"
html = html & "dayNamesShort: ['Dom','Lun','Mar','Mi\u00e9','Jue','Vie','S\u00e1b'],"
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