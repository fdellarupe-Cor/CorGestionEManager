<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/agenda_ext.inc" -->
<!-- #INCLUDE FILE="IncFile/agenda_invitacion.inc" -->
<!-- #INCLUDE FILE="IncFile/ui_componentes.inc" -->
<%
'[UI-ANT] ' ===== Cuerpo anterior completo: reemplazado por el bloque [UI] que sigue. Se borra entero buscando [UI-ANT] =====

'[UI-ANT] xhtml = headhtml()
'[UI-ANT] xhtml = xhtml & "<body>"

'[UI-ANT] opc = request.querystring("opc")
'[UI-ANT] xmenu = request.querystring("xmenu")
'[UI-ANT] xtmenu = request.querystring("xtmenu")

'[UI-ANT] CantRegMostrar = 15
'[UI-ANT] CantRegMover = CDbl(Request.querystring("CantRegMover"))
'[UI-ANT] if CantRegMover < 0 then CantRegMover = 0

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

'[UI-ANT] if request.querystring("opc2") & "" <> "" then opc = request.querystring("opc2")

'[UI-ANT] if opc = "edicion" then opc = "det"
'[UI-ANT] if AgendaExtDebeManejar(opc) then
	'[UI-ANT] AgendaExtRender Cnn, opc, xmenu, xtmenu
	'[UI-ANT] Response.End
'[UI-ANT] end if

'[UI-ANT] xhtml = xhtml & cabecera()

'xhtml = xhtml & "<div style='height:8%;width:100%;border-top:solid 5px #FF3D3D;border-bottom:double 6px #001BFF;'>"

'[UI-ANT] xhtml = xhtml & "<div class='titulo-seccion ui-page-header' style='padding: 1rem 2rem;'>"
'[UI-ANT] xhtml = xhtml & "<div class='ui-app ui-topbar'><div class='ui-topbar__inner'><div class='ui-topbar__top'>"
' DEPURADO [UI]: la fila de la barra ya ordena titulo y acciones
'[UI-ANT] xhtml = xhtml & "<div class='ui-toolbar' style='display:flex;align-items:center;justify-content:space-between;'>"

'[UI-ANT] if opc = "" then
	'[UI-ANT] xhtml = xhtml & "<div><h2>RECORDATORIOS</h2></div>"
	'[UI-ANT] xhtml = xhtml & "<div class='ui-topbar__ident'><h2 class='ui-topbar__title'>Recordatorios</h2></div>"
'[UI-ANT] else
	' [UI] Titulo segun la pantalla, con los datos que ya llegan por querystring
	'[UI-ANT] tituloAgenda = "Agenda" ' [UI]
	'[UI-ANT] if opc = "cargar" then tituloAgenda = "Nuevo recordatorio" ' [UI]
	'[UI-ANT] if opc = "det" then tituloAgenda = "Recordatorio #" & Server.HTMLEncode(request.querystring("id") & "") ' [UI]
	'[UI-ANT] xhtml = xhtml & "<div><h2>AGENDA</h2></div>"
	'[UI-ANT] xhtml = xhtml & "<div class='ui-topbar__ident'><h2 class='ui-topbar__title'>" & tituloAgenda & "</h2></div>"
'[UI-ANT] end if

'[UI-ANT] if opc = "" then
	'[UI-ANT] xform = "<form id='xform' class='form-oculto' method='post' action='index.asp'></form>"
'[UI-ANT] else
	'[UI-ANT] xform = "<form id='xform' class='form-oculto' method='post' action='agenda.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
'[UI-ANT] end if

'[UI-ANT] xboton = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('xform').submit();"">"
'[UI-ANT] xboton = "<input class='ui-button ui-button--secondary ui-button--sm' type='button' value='Volver' onclick=""document.getElementById('xform').submit();"">"

'[UI-ANT] xhtml = xhtml & "<div class='ui-toolbar__actions' style='display:flex; gap: 1rem;'>"
'[UI-ANT] xhtml = xhtml & "<div class='ui-topbar__actions'>"

'[UI-ANT] if opc <> "cargar" and opc <> "det" then

	'[UI-ANT] xforma = "<form id='xformA' class='form-oculto' method='post' action='agenda.asp?opc=cargar&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
	'[UI-ANT] if getPermisos(xmenu,xtmenu,"ALTA") = 1 then
		'[UI-ANT] xbotona = "<input class='bot ui-button ui-button--primary' type='button' value='CARGAR RECORDATORIO' onclick=""document.getElementById('xformA').submit();"">"
		'[UI-ANT] xbotona = "<input class='ui-button ui-button--light ui-button--sm' type='button' value='Cargar recordatorio' onclick=""document.getElementById('xformA').submit();"">"
	'[UI-ANT] else
		'[UI-ANT] xbotona = "<input class='botsp ui-button' type='button' value='CARGAR RECORDATORIO' onclick=""alert('Sin permiso para agregar.');"">"
		'[UI-ANT] xbotona = "<input class='ui-button ui-button--light ui-button--sm ui-button--disabled' type='button' value='Cargar recordatorio' onclick=""alert('Sin permiso para agregar.');"">"
	'[UI-ANT] end if
	'[UI-ANT] xhtml = xhtml & xbotona
	'[UI-ANT] xbotonSemana = "<input class='bot ui-button' type='button' value='VISTA SEMANAL' onclick=""window.open('agenda_semana.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "','agenda_semana','width=1520,height=900,resizable=yes,scrollbars=yes');"">"
	'[UI-ANT] xbotonSemana = "<input class='ui-button ui-button--secondary ui-button--sm' type='button' value='Vista semanal' onclick=""window.open('agenda_semana.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "','agenda_semana','width=1520,height=900,resizable=yes,scrollbars=yes');"">"
	'[UI-ANT] xhtml = xhtml & xbotonSemana
	' DEPURADO [UI]: la paginacion pasa al pie de la tabla. Aca aparecia como texto en las pantallas que no son el listado
	'[UI-ANT] xhtml = xhtml & "[BOTONERA_PAGINACION]"
	'[UI-ANT] xhtml = xhtml & xboton

'[UI-ANT] end if

'[UI-ANT] if opc = "cargar" or opc = "det" then

	'[UI-ANT] xformb = "<form id='xformB' class='form-oculto' method='post' action='agenda.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
	'[UI-ANT] xbotonb = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('xformB').submit();"">"
	'[UI-ANT] xbotonb = "<input class='ui-button ui-button--secondary ui-button--sm' type='button' value='Volver' onclick=""document.getElementById('xformB').submit();"">"
	'[UI-ANT] xhtml = xhtml & xbotonb

'[UI-ANT] end if

'[UI-ANT] xhtml = xhtml & "</div>" ' Fin ui-toolbar__actions

'[UI-ANT] xhtml = xhtml & "</div>" ' Fin ui-toolbar
'[UI-ANT] xhtml = xhtml & "</div>" ' Fin ui-topbar__top
'[UI-ANT] xhtml = xhtml & "</div>" ' Fin titulo-seccion
'[UI-ANT] xhtml = xhtml & "</div></div>" ' Fin ui-topbar__inner y ui-topbar

'[UI-ANT] if opc = "cargar" then

'[UI-ANT] resp = request.form("resp")
'[UI-ANT] if resp & "" = "" then resp = 0

'[UI-ANT] xhtml = xhtml & "<form id='xformu' method='post' action='agenda.asp?opc=cargar&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"

' --- INICIO NORMALIZACION: NUEVO RECORDATORIO ---

'xhtml = xhtml & "<table class='tabl' style='width:35%;' align=center cellspacing=0 cellpadding=3 >"
'[UI-ANT] xhtml = xhtml & "<div class='inner ui-main' style='padding: 2rem;'>"
'[UI-ANT] xhtml = xhtml & "<div class='ui-app ui-page'><div class='ui-container ui-container--narrow'>"
'[UI-ANT] xhtml = xhtml & "<div class='bloque-datos ui-panel' style='max-width: 600px; margin: 0 auto;'>"
'[UI-ANT] xhtml = xhtml & "<div class='ui-card'>"

'xhtml = xhtml & "<tr>"
'xhtml = xhtml & "<td class='fil' style='font-size:12pt;' colspan=2 align=center><u><b>NUEVO RECORDATORIO</b></u></td>"
'xhtml = xhtml & "</tr>"
'[UI-ANT] xhtml = xhtml & "<div class='bloque-header ui-panel__header' style='text-align: center; text-transform: uppercase;'>NUEVO RECORDATORIO</div>"
'[UI-ANT] xhtml = xhtml & "<div class='ui-card__title'>Datos del recordatorio</div>"

'xhtml = xhtml & "<tr>"
'xhtml = xhtml & "<td colspan=2 >&nbsp;</td>"
'xhtml = xhtml & "</tr>"
'[UI-ANT] xhtml = xhtml & "<div class='bloque-body ui-panel__body'>"
'[UI-ANT] xhtml = xhtml & "<div>"
'[UI-ANT] xhtml = xhtml & "<div class='grid-datos ui-field-grid'>"
'[UI-ANT] xhtml = xhtml & "<div class='ui-fields'>"

''RESPON
'xhtml = xhtml & "<tr>"
'xhtml = xhtml & "<td class='fil' >RESPONSABLE</td>"
'[UI-ANT] xhtml = xhtml & "<div class='dato-label'>RESPONSABLE</div>"
'[UI-ANT] xhtml = xhtml & "<div><div class='ui-field__label'>Responsable</div>"
	
'xhtml = xhtml & "<td class='fil' >"
'xhtml = xhtml & getName(cnn,"","cor_responsables",session("usuario"))
'xhtml = xhtml & "</td>"
'xhtml = xhtml & "</tr>"
'[UI-ANT] xhtml = xhtml & "<div class='dato-value'>" & getName(cnn,"","cor_responsables",session("usuario")) & "</div>"
'[UI-ANT] xhtml = xhtml & "<div class='ui-field__value'>" & getName(cnn,"","cor_responsables",session("usuario")) & "</div></div>"

''ASUNTO
'xhtml = xhtml & "<tr>"
'xhtml = xhtml & "<td class='fil' >ASUNTO</font></td>"
'[UI-ANT] xhtml = xhtml & "<div class='dato-label'>ASUNTO</div>"
'[UI-ANT] xhtml = xhtml & "<div class='ui-grid__col--wide'><div class='ui-field__label'>Asunto</div>"
	
'xhtml = xhtml & "<td class='fil' ><input class='inpu' style='width:80%;' id='asun' name='asun' type='text' /></td>"
'xhtml = xhtml & "</tr>"
'[UI-ANT] xhtml = xhtml & "<div><input class='input-moderno ui-input' id='asun' name='asun' type='text' /></div>"
'[UI-ANT] xhtml = xhtml & "<input class='ui-input' id='asun' name='asun' type='text' /></div>"

''DIA
'xhtml = xhtml & "<tr>"
'xhtml = xhtml & "<td class='fil' >AVISAR DIA</td>"
'[UI-ANT] xhtml = xhtml & "<div class='dato-label'>AVISAR DIA</div>"
'[UI-ANT] xhtml = xhtml & "<div><div class='ui-field__label'>Avisar el d&iacute;a</div>"
	
'xhtml = xhtml & "<td class='fil' ><input class='inpu' style='width:80%;' id='datepicker' name='datepicker' type='text'></td>"
'xhtml = xhtml & "</tr>"
'[UI-ANT] xhtml = xhtml & "<div><input class='input-moderno ui-input' id='datepicker' name='datepicker' type='text'></div>"
'[UI-ANT] xhtml = xhtml & "<input class='ui-input ui-date' id='datepicker' name='datepicker' type='text' autocomplete='off'></div>"

''ESPACIO (Separador Visual)
'xhtml = xhtml & "<tr>"
'xhtml = xhtml & "<td colspan=2>&nbsp;</td>"
'xhtml = xhtml & "</tr>"
'[UI-ANT] xhtml = xhtml & "<div style='grid-column: 1 / -1;'><hr style='border: 0; border-top: 1px solid #e2e8f0; margin: 1.5rem 0;'></div>"
'[UI-ANT] xhtml = xhtml & "</div></div></div><div class='ui-card'><div class='ui-card__title'>Enviar copia a</div><div><div class='ui-fields'>"

''AVISAR A
'xhtml = xhtml & "<tr>"
'xhtml = xhtml & "<td class='fil' colspan=2>ENVIAR COPIA A:</td>"
'xhtml = xhtml & "</tr>"
' DEPURADO [UI]: el titulo ahora es el de la tarjeta
'[UI-ANT] xhtml = xhtml & "<div class='dato-label' style='grid-column: 1 / -1; margin-bottom: 0.5rem;'>ENVIAR COPIA A:</div>"

'[UI-ANT] Set rx = Exec(cnn, "sp_cg_Responsables @Modo='listar_activos_sin', @ExcluirId=" & SqlStr(session("usuario")))

	'[UI-ANT] do while not rx.eof

		'[UI-ANT] chkbx = ""
		'xhtml = xhtml & "<td class='fil' >" & UCASE(rx(1)) & "</td>"
		' Usamos un padding-left para jerarquizar visualmente los nombres bajo el titulo "ENVIAR COPIA A:"
		'[UI-ANT] xhtml = xhtml & "<div class='dato-label' style='padding-left: 1rem; color: #475569;'>" & UCASE(rx(1)) & "</div>"
		'[UI-ANT] xhtml = xhtml & "<div><div class='ui-field__label'>" & Server.HTMLEncode(rx(1) & "") & "</div>"
		'[UI-ANT] sel = "<select class='selec ui-select' style='width:30%;' name='u" & rx(0) & "'  >"
		'[UI-ANT] sel = "<select class='ui-select' name='u" & rx(0) & "'>"
		'[UI-ANT] sel = sel & "<option value='0'>NO</option>"
		'[UI-ANT] sel = sel & "<option value='0'>No</option>"
		'[UI-ANT] sel = sel & "<option value='1'>SI</option>"
		'[UI-ANT] sel = sel & "<option value='1'>S&iacute;</option>"
		'[UI-ANT] sel = sel & "</select>"
		'xhtml = xhtml & "<td class='fil'>" & sel & "</td>"
		' Envolver en .dato-value activa el parche de tu CSS para selects
		'[UI-ANT] xhtml = xhtml & "<div class='dato-value'>" & sel & "</div>"
		'[UI-ANT] xhtml = xhtml & sel & "</div>"
		'xhtml = xhtml & "</tr>"

	'[UI-ANT] rx.movenext
	'[UI-ANT] loop

'xhtml = xhtml & "<tr>"
'xhtml = xhtml & "<td  colspan=2>&nbsp;</td>"
'xhtml = xhtml & "</tr>"

'xhtml = xhtml & "<tr>"
'xboton = "<input class='bot' type='button' value='Cargar' onclick=""ejecutar('xformu','Insertar','ins');"" />"
'[UI-ANT] xboton = "<input class='btn-principal ui-button ui-button--primary' type='button' value='Cargar' onclick=""ejecutar('xformu','Insertar','ins');"" />"
'[UI-ANT] xboton = "<input class='ui-button ui-button--primary ui-button--lg' type='button' value='Cargar recordatorio' onclick=""ejecutar('xformu','\u00bfCargar el recordatorio?','ins');"" />"

'xhtml = xhtml & "<td class='fil' colspan=2 align=center>" & xboton & "</td>"
'xhtml = xhtml & "</tr>"
' Contenedor Flexbox para centrar el boton rompiendo la grilla
' DEPURADO [UI]: el boton va en la barra de guardado, fuera de las tarjetas
'[UI-ANT] xhtml = xhtml & "<div style='grid-column: 1 / -1; display: flex; justify-content: center; margin-top: 1.5rem;'>"
'[UI-ANT] xhtml = xhtml & xboton
'[UI-ANT] xhtml = xhtml & "</div>"

'xhtml = xhtml & "</table>"
'[UI-ANT] xhtml = xhtml & "</div>" ' Cierre de .grid-datos
'[UI-ANT] xhtml = xhtml & "</div>" ' Cierre de .bloque-body
'[UI-ANT] xhtml = xhtml & "</div>" ' Cierre de .bloque-datos
'[UI-ANT] xhtml = xhtml & "<div class='ui-savebar'>" & xboton & "</div>" ' [UI]
'[UI-ANT] xhtml = xhtml & "</div>" ' Cierre de .inner
'[UI-ANT] xhtml = xhtml & "</div></div>" ' Cierre de ui-container y ui-page

'[UI-ANT] xhtml = xhtml & "</form>"

'[UI-ANT] elseif opc = "modif" then

'[UI-ANT] id = request.querystring("id")
'[UI-ANT] idresp = request.querystring("idresp")
'[UI-ANT] ccs = request.querystring("ccs")

'[UI-ANT] ExecNonQuery Cnn, "sp_cg_Agenda @Modo='actualizar_asunto', @Id=" & id & ", @Asunto='" & request.form("asun") & "'"

'[UI-ANT] xhtml = xhtml & "<FORM name='initx' method='post'  ACTION='agenda.asp?opc=det&id=" & id & "&idresp=" & Server.URLEncode(idresp & "") & "&ccs=" & Server.URLEncode(ccs & "") & "' >"
'[UI-ANT] xhtml = xhtml & "<INPUT TYPE='SUBMIT' VALUE='INICIAR' style='display:none;' >"
'[UI-ANT] xhtml = xhtml & "</FORM>"

'[UI-ANT] xhtml = xhtml & "<script type='text/javascript'>"
'[UI-ANT] xhtml = xhtml & " alert('Actualizado exitosamente');"
'[UI-ANT] xhtml = xhtml & " document.all.initx.submit();"
'[UI-ANT] xhtml = xhtml & "</script>"

'[UI-ANT] elseif opc = "insertar" then

'[UI-ANT] set rz = Cnn.Execute("sp_cg_Responsables @Modo='listar_ids_activos_sin', @ExcluirId=" & SqlStr(session("usuario")))

'[UI-ANT] copias = ""
'[UI-ANT] sep = ""

'[UI-ANT] do while not rz.eof

	'[UI-ANT] num = rz(0)

	'[UI-ANT] if request.form("u" & num) = 1 then
		'[UI-ANT] copias = copias & sep & rz(0)
		'[UI-ANT] sep = ","
	'[UI-ANT] end if

'[UI-ANT] rz.movenext
'[UI-ANT] loop

'[UI-ANT] sDiaAviso = clng(CDATE(request.form("datepicker"))) - CLng(CDate("01-01-1900"))
'[UI-ANT] sFecHoy = clng(date) - CLng(CDate("01-01-1900"))


'[UI-ANT] ExecNonQuery Cnn, "sp_cg_Agenda @Modo='insertar', @Asunto='" & request.form("asun") & "', @DiaAviso=" & sDiaAviso & ", @IdRespon=" & SqlStr(session("usuario")) & ", @Copias='" & copias & "', @FecCarga=" & sFecHoy

'[UI-ANT] XXX = SYSLOGON(cnn,"AGREGA_REC")

'[UI-ANT] xhtml = xhtml & "<FORM name='init' method='post'  ACTION='agenda.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "' >"
'[UI-ANT] xhtml = xhtml & "<INPUT TYPE='SUBMIT' VALUE='INICIAR' style='display:none;' >"
'[UI-ANT] xhtml = xhtml & "</FORM>"

'[UI-ANT] xhtml = xhtml & "<script type='text/javascript'>"
'[UI-ANT] xhtml = xhtml & " alert('Agregado exitosamente');"
'[UI-ANT] xhtml = xhtml & " document.all.init.submit();"
'[UI-ANT] xhtml = xhtml & "</script>"

'[UI-ANT] elseif opc = "det" then

'[UI-ANT] id = request.querystring("id")
'[UI-ANT] idresp = request.querystring("idresp")
'[UI-ANT] idagre = request.querystring("idagre")
'[UI-ANT] ccs = request.querystring("ccs")
'[UI-ANT] xopt = request.querystring("xopt")

'[UI-ANT] set rp = Cnn.Execute("sp_cg_Agenda @Modo='obtener', @Id=" & id)

'[UI-ANT] puede = 1

'[UI-ANT] if not rp.eof then
	'[UI-ANT] if rp(0) = "SI" then puede = 0
'[UI-ANT] end if

'[UI-ANT] if xopt = "agre" then

'[UI-ANT] if ccs & "" <> "" then

'[UI-ANT] ExecNonQuery Cnn, "sp_cg_Agenda @Modo='actualizar_copias', @Id=" & id & ", @Copias='" & ccs & "," & idagre & "'"

'[UI-ANT] ccs = ccs & "," & idagre

'[UI-ANT] else

'[UI-ANT] ExecNonQuery Cnn, "sp_cg_Agenda @Modo='actualizar_copias', @Id=" & id & ", @Copias='" & idagre & "'"

'[UI-ANT] ccs = idagre

'[UI-ANT] end if

'[UI-ANT] end if

'[UI-ANT] xhtml = xhtml & "<div class='inner ui-main' style='padding: 2rem;'>"
'[UI-ANT] xhtml = xhtml & "<div class='ui-app ui-page'><div class='ui-container ui-container--narrow'>"

' [UI] Tarjeta Asunto. Antes iba al final y solo si podia modificarse.
' [UI] Ahora se muestra siempre: editable para el responsable, lectura para el resto.
'[UI-ANT] xhtml = xhtml & "<div class='ui-card'><div class='ui-card__title'>Asunto</div>" ' [UI]
'[UI-ANT] if not rp.eof then ' [UI]
	'[UI-ANT] if puede = 1 and cstr(idresp) = cstr(session("USUARIO")) then ' [UI]
		'[UI-ANT] xhtml = xhtml & "<form id='act' method='post' action='agenda.asp?opc=modif&id=" & id & "&ccs=" & Server.URLEncode(ccs & "") & "&idresp=" & Server.URLEncode(idresp & "") & "'>" ' [UI]
		'[UI-ANT] xhtml = xhtml & "<input id='asun' name='asun' class='ui-input' type='text' value='" & Server.HTMLEncode(rp(1) & "") & "' onblur=""document.getElementById('act').submit()"">" ' [UI]
		'[UI-ANT] xhtml = xhtml & "<div class='ui-field__hint'>Se guarda al salir del campo.</div>" ' [UI]
		'[UI-ANT] xhtml = xhtml & "</form>" ' [UI]
	'[UI-ANT] else ' [UI]
		'[UI-ANT] xhtml = xhtml & "<div class='ui-prose'>" & Server.HTMLEncode(rp(1) & "") & "</div>" ' [UI]
	'[UI-ANT] end if ' [UI]
'[UI-ANT] else ' [UI]
	'[UI-ANT] xhtml = xhtml & "<div class='ui-empty'>No se encontr&oacute; el recordatorio.</div>" ' [UI]
'[UI-ANT] end if ' [UI]
'[UI-ANT] xhtml = xhtml & "</div>" ' [UI] Fin tarjeta Asunto


'[UI-ANT] xhtml = xhtml & "<div class='bloque-datos ui-panel' style='max-width: 480px; margin: 0 auto;'>"
'[UI-ANT] xhtml = xhtml & "<div class='ui-card'>"
'[UI-ANT] xhtml = xhtml & "<div class='bloque-header ui-panel__header' style='text-align: center;'>RECORDATORIO " & id & "</div>"
'[UI-ANT] xhtml = xhtml & "<div class='ui-card__title'>Copias enviadas a</div>"
'[UI-ANT] xhtml = xhtml & "<div class='bloque-body ui-panel__body'>"
'[UI-ANT] xhtml = xhtml & "<div>"

'[UI-ANT] if ccs & "" <> "" then

'[UI-ANT] set rf = Cnn.Execute("sp_cg_Responsables @Modo='nombre_id_in', @Ids='" & ccs & "'")

'[UI-ANT] xfrm = ""

'[UI-ANT] xhtml = xhtml & "<div style='display: flex; flex-direction: column; gap: 0.5rem;'>"
'[UI-ANT] xhtml = xhtml & "<div class='ui-list'>"

'[UI-ANT] do while not rf.eof

'[UI-ANT] if puede = 1 then
	'[UI-ANT] xfrm = xfrm & "<form id='elim" & DomId(rf(1)) & "' method='post' action='agenda.asp?opc=elim&idag=" & id & "&ccs=" & Server.URLEncode(ccs & "") & "&idresp=" & Server.URLEncode(idresp & "") & "&idelim=" & Server.URLEncode(rf(1) & "") & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "' ></form>"
'[UI-ANT] end if

'[UI-ANT] if getPermisos(xmenu,xtmenu,"ELIM") = 1 then
	'[UI-ANT] xinput = "<input class='bot ui-button ui-button--icon ui-button--danger' type='button' value='X' onclick=""document.getElementById('elim" & DomId(rf(1)) & "').submit();"">"
	'[UI-ANT] xinput = "<input class='ui-button ui-button--icon ui-button--danger' type='button' value='&times;' title='Quitar copia' onclick=""document.getElementById('elim" & DomId(rf(1)) & "').submit();"">"
'[UI-ANT] else
	'[UI-ANT] xinput = "<input class='botsp ui-button ui-button--icon' type='button' value='X' onclick=""alert('Sin permiso para eliminar.');"">"
	'[UI-ANT] xinput = "<input class='ui-button ui-button--icon ui-button--disabled' type='button' value='&times;' title='Sin permiso para eliminar' onclick=""alert('Sin permiso para eliminar.');"">"
'[UI-ANT] end if

'[UI-ANT] xhtml = xhtml & "<div style='display: flex; align-items: center; gap: 0.5rem;'>"
'[UI-ANT] xhtml = xhtml & "<div class='ui-list__item ui-list__item--action'>"

'[UI-ANT] if cstr(idresp) = cstr(session("USUARIO")) then
	'[UI-ANT] xhtml = xhtml & xinput
'[UI-ANT] end if
'[UI-ANT] xhtml = xhtml & "<span>-" & rf(0) & "</span>"
'[UI-ANT] xhtml = xhtml & "<span>" & Server.HTMLEncode(rf(0) & "") & "</span>"

'[UI-ANT] xhtml = xhtml & "</div>"

'[UI-ANT] rf.movenext
'[UI-ANT] loop

'[UI-ANT] xhtml = xhtml & "</div>"

'[UI-ANT] else

'[UI-ANT] xhtml = xhtml & "<div class='dato-value'>NO REGISTRA COPIAS</div>"
'[UI-ANT] xhtml = xhtml & "<div class='ui-empty'>No registra copias.</div>"

'[UI-ANT] end if

'[UI-ANT] xhtml = xhtml & "</div>" ' Fin bloque-body
'[UI-ANT] xhtml = xhtml & "</div>" ' Fin bloque-datos (recordatorio)

'[UI-ANT] if ccs & "" <> "" then

'[UI-ANT] set rm = Cnn.Execute("sp_cg_Responsables @Modo='listar_not_in', @Ids='" & ccs & "', @ExcluirId=" & SqlStr(idresp))

'[UI-ANT] else

'[UI-ANT] set rm = Cnn.Execute("sp_cg_Responsables @Modo='id_nombre_sin', @ExcluirId=" & SqlStr(idresp))

'[UI-ANT] end if

'[UI-ANT] xhtml = xhtml & "<div class='bloque-datos ui-panel' style='max-width: 480px; margin: 1.5rem auto 0;'>"
'[UI-ANT] xhtml = xhtml & "<div class='ui-card'>"
'[UI-ANT] xhtml = xhtml & "<div class='bloque-header ui-panel__header' style='text-align: center;'>AGREGAR USUARIO PARA ENVIAR COPIA</div>"
'[UI-ANT] xhtml = xhtml & "<div class='ui-card__title'>Agregar copia a</div>"
'[UI-ANT] xhtml = xhtml & "<div class='bloque-body ui-panel__body'>"
'[UI-ANT] xhtml = xhtml & "<div>"

'[UI-ANT] if not rm.eof then

'[UI-ANT] xhtml = xhtml & "<div style='display: flex; flex-direction: column; gap: 0.5rem;'>"
'[UI-ANT] xhtml = xhtml & "<div class='ui-list'>"

'[UI-ANT] do while not rm.eof

'[UI-ANT] if puede = 1 then

	'[UI-ANT] if cstr(idresp) = cstr(session("USUARIO")) then
		'[UI-ANT] ref = "<a class='red' href='agenda.asp?opc=det&xopt=agre&id=" & id & "&ccs=" & Server.URLEncode(ccs & "") & "&idresp=" & Server.URLEncode(idresp & "") & "&idagre=" & Server.URLEncode(rm(0) & "") & "'>" & rm(1) & "</a>"
		'[UI-ANT] ref = "<a class='ui-link' href='agenda.asp?opc=det&xopt=agre&id=" & id & "&ccs=" & Server.URLEncode(ccs & "") & "&idresp=" & Server.URLEncode(idresp & "") & "&idagre=" & Server.URLEncode(rm(0) & "") & "'>" & Server.HTMLEncode(rm(1) & "") & "</a>"
	'[UI-ANT] else
		'[UI-ANT] ref = rm(1)
		'[UI-ANT] ref = Server.HTMLEncode(rm(1) & "")
	'[UI-ANT] end if

'[UI-ANT] else
	'[UI-ANT] ref = rm(1)
	'[UI-ANT] ref = Server.HTMLEncode(rm(1) & "")
'[UI-ANT] end if

'[UI-ANT] xhtml = xhtml & "<div>" & ref & "</div>"
'[UI-ANT] xhtml = xhtml & "<div class='ui-list__item'>" & ref & "</div>"

'[UI-ANT] rm.movenext
'[UI-ANT] loop

'[UI-ANT] xhtml = xhtml & "</div>"
'[UI-ANT] if puede <> 1 or cstr(idresp) <> cstr(session("USUARIO")) then xhtml = xhtml & "<div class='ui-field__hint'>Solo el responsable del recordatorio puede agregar copias.</div>" ' [UI]

'[UI-ANT] else

'[UI-ANT] xhtml = xhtml & "<div class='dato-value'>SIN USUARIOS DISPONIBLES</div>"
'[UI-ANT] xhtml = xhtml & "<div class='ui-empty'>No hay usuarios disponibles.</div>"

'[UI-ANT] end if

'[UI-ANT] xhtml = xhtml & "</div>" ' Fin bloque-body
'[UI-ANT] xhtml = xhtml & "</div>" ' Fin bloque-datos (agregar usuario)

' DEPURADO [UI]: la tarjeta Asunto ahora va primero (ver arriba)
'[UI-ANT] if puede = 1 then

'[UI-ANT] xhtml = xhtml & "<form id='act' method='post' action='agenda.asp?opc=modif&id=" & id & "&ccs=" & Server.URLEncode(ccs & "") & "&idresp=" & Server.URLEncode(idresp & "") & "'>"

'[UI-ANT] xhtml = xhtml & "<div class='bloque-datos ui-panel' style='max-width: 480px; margin: 1.5rem auto 0;'>"
'[UI-ANT] xhtml = xhtml & "<div class='bloque-header ui-panel__header' style='text-align: center;'>MODIFICAR ASUNTO</div>"
'[UI-ANT] xhtml = xhtml & "<div class='bloque-body ui-panel__body'>"

'[UI-ANT] if cstr(idresp) = cstr(session("USUARIO")) then
	'[UI-ANT] xhtml = xhtml & "<input id='asun' name='asun' class='inpu ui-input' style='width:100%;' type='text' value='" & rp(1) & "' onblur=""document.getElementById('act').submit()"">"
'[UI-ANT] else
	'[UI-ANT] xhtml = xhtml & "<div class='dato-value'>" & rp(1) & "</div>"
'[UI-ANT] end if

'[UI-ANT] xhtml = xhtml & "</div>" ' Fin bloque-body
'[UI-ANT] xhtml = xhtml & "</div>" ' Fin bloque-datos (modificar asunto)

'[UI-ANT] xhtml = xhtml & "</form>"

'[UI-ANT] end if

'[UI-ANT] xhtml = xhtml & "</div>" ' Fin inner
'[UI-ANT] xhtml = xhtml & "</div></div>" ' Fin ui-container y ui-page

'[UI-ANT] elseif opc = "elim" then

'[UI-ANT] idresp = request.querystring("idresp")
'[UI-ANT] idelim = request.querystring("idelim")
'[UI-ANT] idag = request.querystring("idag")

'[UI-ANT] set rt = Cnn.Execute("sp_cg_Agenda @Modo='obtener_copias', @Id=" & idag)

'[UI-ANT] if not rt.eof then

'[UI-ANT] xcopy = split(rt(0),",")
'[UI-ANT] xcopyx = ""
'[UI-ANT] sep = ""

'[UI-ANT] for i=0 to ubound(xcopy)

	'[UI-ANT] if xcopy(i) <> idelim then
		'[UI-ANT] xcopyx = xcopyx & sep & xcopy(i)
		'[UI-ANT] sep = ","
	'[UI-ANT] end if

'[UI-ANT] next

'[UI-ANT] rt.close

'[UI-ANT] ExecNonQuery Cnn, "sp_cg_Agenda @Modo='actualizar_copias', @Id=" & idag & ", @Copias='" & xcopyx & "'"

'[UI-ANT] set rt = Cnn.Execute("sp_cg_Agenda @Modo='obtener_copias', @Id=" & idag)

'[UI-ANT] ccs = ""

'[UI-ANT] if not rt.eof then ccs = rt(0)

'[UI-ANT] xhtml = xhtml & "<FORM name='init' method='post'  ACTION='agenda.asp?opc=det&id=" & idag & "&ccs=" & Server.URLEncode(ccs & "") & "&idresp=" & Server.URLEncode(idresp & "") & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "' >"
'[UI-ANT] xhtml = xhtml & "<INPUT TYPE='SUBMIT' VALUE='INICIAR' style='display:none;' >"
'[UI-ANT] xhtml = xhtml & "</FORM>"

'[UI-ANT] xhtml = xhtml & "<script type='text/javascript'>"
'[UI-ANT] xhtml = xhtml & " document.all.init.submit();"
'[UI-ANT] xhtml = xhtml & "</script>"

'[UI-ANT] end if

'[UI-ANT] else 'listado de recordatorios

	' --- INICIO NORMALIZACION: AGENDA (RECORDATORIOS) ---
	'[UI-ANT] XXX = SYSLOGON(cnn,"MENU_AGENDA")

	' Contenedor con padding para que la tabla no pegue a los bordes
	'[UI-ANT] xhtml = xhtml & "<div class='inner ui-main' style='padding: 2rem;'>"
	'[UI-ANT] xhtml = xhtml & "<div class='ui-app ui-page'><div class='ui-container'><div class='ui-card'><div class='ui-card__title'>Listado de recordatorios</div>"

	' Tarjeta contenedora con scroll horizontal si es necesario
	'[UI-ANT] xhtml = xhtml & "<div class='bloque-datos table-responsive ui-table-wrap'>"
	'[UI-ANT] xhtml = xhtml & "<div class='ui-table-wrap'>"

	'xhtml = xhtml & "<table style='width:62%;color:#000000;font-family:Serif;text-align:center;' cellspacing=0 align=center >"
	'[UI-ANT] xhtml = xhtml & "<table class='tabla-moderna ui-table'>"
	'[UI-ANT] xhtml = xhtml & "<table class='ui-table ui-table--middle'>"

	' Encabezados unificados en thead
	'[UI-ANT] xhtml = xhtml & "<thead>"
	'[UI-ANT] xhtml = xhtml & "<tr>"
	'[UI-ANT] xhtml = xhtml & "<th style='width: 50px;'>&nbsp;</th>" ' Columna de la X
	'[UI-ANT] xhtml = xhtml & "<th class='ui-table__col--action'></th>"
	'[UI-ANT] xhtml = xhtml & "<th>ID</th>"
	'[UI-ANT] xhtml = xhtml & "<th>ID</th>"
	'[UI-ANT] xhtml = xhtml & "<th>RESPONSABLE</th>"
	'[UI-ANT] xhtml = xhtml & "<th>Responsable</th>"
	'[UI-ANT] xhtml = xhtml & "<th>ASUNTO</th>"
	'[UI-ANT] xhtml = xhtml & "<th>Asunto</th>"
	'[UI-ANT] xhtml = xhtml & "<th>DIA</th>"
	'[UI-ANT] xhtml = xhtml & "<th>D&iacute;a de aviso</th>"
	'[UI-ANT] xhtml = xhtml & "<th style='text-align: center;'>AVISADO</th>"
	'[UI-ANT] xhtml = xhtml & "<th class='ui-table__cell--center'>Avisado</th>"
	'[UI-ANT] xhtml = xhtml & "</tr>"
	'[UI-ANT] xhtml = xhtml & "</thead>"
	
	'[UI-ANT] xhtml = xhtml & "<tbody>"

'[UI-ANT] set rs = Cnn.Execute("sp_cg_Agenda @Modo='listar'")

'[UI-ANT] formx1 = ""

'[UI-ANT] if not rs.eof then rs.Move (CantRegMover)

'[UI-ANT] PosReg = 0
'[UI-ANT] sColor = "#FFFFFF"

'[UI-ANT] Do While (Not rs.eof) And (PosReg < CantRegMostrar)

'[UI-ANT] formx1 = formx1 & "<form id='formx1" & rs(0) & "' style='height:1px;' method='post' action='eliminar.asp?opc=rec&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=agenda.asp'></form>"
'[UI-ANT] if getPermisos(xmenu,xtmenu,"ELIM") = 1 then
	'botonx1 = "<input class='bot' type='button' value='X' onclick=""ejecutar('formx1" & rs(0) & "','Borrar','borra');"">"
	'[UI-ANT] botonx1 = "<input type='button' class='btn-accion btn-accion-peligro ui-button ui-button--icon ui-button--danger' value='X' onclick=""ejecutar('formx1" & rs(0) & "','Borrar','borra');"">"
	'[UI-ANT] botonx1 = "<input type='button' class='ui-button ui-button--icon ui-button--danger' value='&times;' title='Eliminar recordatorio' onclick=""ejecutar('formx1" & rs(0) & "','\u00bfEliminar el recordatorio?','borra');"">"
'[UI-ANT] else
	'botonx1 = "<input class='botsp' type='button' value='X' onclick=""alert('Sin permiso para eliminar.');"">"
	'[UI-ANT] botonx1 = "<input type='button' class='btn-accion btn-accion-deshabilitado ui-button ui-button--icon' value='X' onclick=""alert('Sin permiso para eliminar.');"">"
	'[UI-ANT] botonx1 = "<input type='button' class='ui-button ui-button--icon ui-button--disabled' value='&times;' title='Sin permiso para eliminar' onclick=""alert('Sin permiso para eliminar.');"">"
'[UI-ANT] end if

'[UI-ANT] resp = getName(Cnn,"","cor_responsables",rs(3))

'[UI-ANT] if rs(4) = "SI" then
	'aviso = "<font style='color:#088A08;'>SI</font>"
	'[UI-ANT] aviso = "<span class='badge badge-ok ui-badge ui-badge--ok'>SI</span>"
	'[UI-ANT] aviso = "<span class='ui-badge ui-badge--ok'>S&iacute;</span>"
'[UI-ANT] elseif rs(4) = "NO" then
	'aviso = "<font style='color:#FF0000;'>NO</font>"
	'[UI-ANT] aviso = "<span class='badge badge-peligro ui-badge ui-badge--danger'>NO</span>"
	'[UI-ANT] aviso = "<span class='ui-badge ui-badge--danger'>No</span>"
'[UI-ANT] else
    '[UI-ANT] aviso = "S/D"
    '[UI-ANT] aviso = "<span class='ui-field__value--muted'>S/D</span>"
'[UI-ANT] end if

'[UI-ANT] SFECHA = ""
'[UI-ANT] IF RS(2) <> 0 THEN
	'[UI-ANT] SFECHA = CDATE(rs(2)) + CLng(CDate("01-01-1900"))
'[UI-ANT] END IF

' [UI] Asunto recortado solo si es largo, con puntos suspensivos reales
'[UI-ANT] sAsunto = Server.HTMLEncode(RTrim(left(rs(1) & "",30))) ' [UI]
'[UI-ANT] if len(rs(1) & "") > 30 then sAsunto = sAsunto & "&hellip;" ' [UI]
'[UI-ANT] UrlDetalle = "agenda.asp?opc=det&id=" & rs(0) & "&ccs=" & Server.URLEncode(rs(6) & "") & "&idresp=" & Server.URLEncode(rs(3) & "") & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu

'[UI-ANT] if getPermisos(xmenu,xtmenu,"ELIM") = 1 then
	'[UI-ANT] JsClickFila = "onclick=""document.location='" & UrlDetalle & "'"""
'[UI-ANT] else
	'[UI-ANT] JsClickFila = "onclick=""alert('Sin permiso para modificar.');"""
'[UI-ANT] end if

'[UI-ANT] xhtml = xhtml & "<tr>"

'[UI-ANT] xhtml = xhtml & "<td>" & botonx1 & "</td>"
'[UI-ANT] xhtml = xhtml & "<td class='ui-table__col--action'>" & botonx1 & "</td>"
'xhtml = xhtml & "<td class='fil' style='background-color:" & sColor & ";'>" & rs(0) & "</td>"
'[UI-ANT] xhtml = xhtml & "<td style='cursor:pointer;' " & JsClickFila & ">" & rs(0) & "</td>"
'[UI-ANT] xhtml = xhtml & "<td class='ui-table__cell--link' " & JsClickFila & ">" & rs(0) & "</td>"
'[UI-ANT] xhtml = xhtml & "<td style='cursor:pointer;' " & JsClickFila & ">" & resp & "</td>"
'[UI-ANT] xhtml = xhtml & "<td class='ui-table__cell--link' " & JsClickFila & ">" & Server.HTMLEncode(resp & "") & "</td>"
'[UI-ANT] xhtml = xhtml & "<td style='cursor:pointer;' " & JsClickFila & " title='" & rs(1) & "'>" & left(rs(1),30) & ".." & "</td>"
'[UI-ANT] xhtml = xhtml & "<td class='ui-table__cell--link' " & JsClickFila & " title='" & Server.HTMLEncode(rs(1) & "") & "'>" & sAsunto & "</td>"
'[UI-ANT] xhtml = xhtml & "<td style='cursor:pointer;' " & JsClickFila & ">" & sFECHA & "</td>"
'[UI-ANT] xhtml = xhtml & "<td class='ui-table__cell--link' " & JsClickFila & ">" & sFECHA & "</td>"
'[UI-ANT] xhtml = xhtml & "<td style='text-align:center;cursor:pointer;' " & JsClickFila & ">" & aviso & "</td>"
'[UI-ANT] xhtml = xhtml & "<td class='ui-table__cell--link ui-table__cell--center' " & JsClickFila & ">" & aviso & "</td>"

'[UI-ANT] xhtml = xhtml & "</tr>"

'if sColor = "#FFFFFF" then
	'sColor = "#ECECEC"
'elseif sColor = "#ECECEC" then
	'sColor = "#FFFFFF"
'end if

'[UI-ANT] PosReg = PosReg + 1

'[UI-ANT] rs.movenext
'[UI-ANT] loop

' [UI] Estado vacio cuando no hay registros para mostrar
'[UI-ANT] if PosReg = 0 then xhtml = xhtml & "<tr><td colspan='6'><div class='ui-empty'>No hay recordatorios cargados.</div></td></tr>" ' [UI]
'[UI-ANT] xhtml = xhtml & "</table>"
'[UI-ANT] xhtml = xhtml & "</tbody></table>"

'[UI-ANT] xhtml = xhtml & "</div>"
' [UI] Paginacion al pie de la tarjeta; se completa mas abajo con Replace()
'[UI-ANT] xhtml = xhtml & "[BOTONERA_PAGINACION]" ' [UI]
'[UI-ANT] xhtml = xhtml & "</div></div></div>" ' [UI] Fin ui-card, ui-container y ui-page

' --- ARMADO DE LA BOTONERA DE AGENDA ---

'shtml =  "<div id=btn style='top:85px;left:750px;position:absolute;'><table align=right><tr>"
'[UI-ANT] misBotones = ""

'[UI-ANT] if cantregmover <> 0 then
	'formx22 = "<form id='formx22' style='height:1px;' method='post' action='agenda.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&cantregmover=" & (cantregmover - cantregmostrar) & "'></form>"
	'[UI-ANT] formx22 = "<form id='formx22' class='form-oculto' method='post' action='agenda.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&cantregmover=" & (cantregmover - cantregmostrar) & "'></form>"
	
	'botonx22 = "<input class='bot' type='button' value='ATRAS' onclick=""document.getElementById('formx22').submit();"">"
	'[UI-ANT] botonx22 = "<input class='bot ui-button' type='button' value='ATRAS' onclick=""document.getElementById('formx22').submit();"">"
	'[UI-ANT] botonx22 = "<input class='ui-button ui-button--secondary ui-button--sm' type='button' value='Anterior' onclick=""document.getElementById('formx22').submit();"">"
'[UI-ANT] else
	'botonx22 = "<input class='botsp' type='button' value='ATRAS'>"
	'[UI-ANT] formx22 = ""
	'[UI-ANT] botonx22 = "<input class='botsp ui-button' type='button' value='ATRAS' disabled>"
	'[UI-ANT] botonx22 = "<input class='ui-button ui-button--secondary ui-button--sm ui-button--disabled' type='button' value='Anterior' disabled>"
'[UI-ANT] end if

'[UI-ANT] if not rs.eof then
	'formx23 = "<form id='formx23' style='height:1px;' method='post' action='agenda.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&cantregmover=" & (cantregmover + cantregmostrar) & "'></form>"
	'[UI-ANT] formx23 = "<form id='formx23' class='form-oculto' method='post' action='agenda.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&cantregmover=" & (cantregmover + cantregmostrar) & "'></form>"

	'botonx23 = "<input class='bot' type='button' value='SIGUIENTE' onclick=""document.getElementById('formx23').submit();"">"
	'[UI-ANT] botonx23 = "<input class='bot ui-button' type='button' value='SIGUIENTE' onclick=""document.getElementById('formx23').submit();"">"
	'[UI-ANT] botonx23 = "<input class='ui-button ui-button--secondary ui-button--sm' type='button' value='Siguiente' onclick=""document.getElementById('formx23').submit();"">"
'[UI-ANT] else
	'botonx23 = "<input class='botsp' type='button' value='SIGUIENTE'>"
	'[UI-ANT] formx23 = ""
	'[UI-ANT] botonx23 = "<input class='botsp ui-button' type='button' value='SIGUIENTE' disabled>"
	'[UI-ANT] botonx23 = "<input class='ui-button ui-button--secondary ui-button--sm ui-button--disabled' type='button' value='Siguiente' disabled>"
'[UI-ANT] end if

'shtml = shtml & "<td>" & botonx22 & "</td>"
'shtml = shtml & "<td>" & botonx23 & "</td>"
'shtml = shtml & "</tr></table></div>"

' Agrupamos los formularios y botones (con flexbox para alinear)
'[UI-ANT] misBotones = "<div class='ui-pagination' style='display: flex; gap: 0.5rem;'>" & formx22 & botonx22 & formx23 & botonx23 & "</div>"
'[UI-ANT] misBotones = "<div class='ui-card__footer'>" & formx22 & botonx22 & formx23 & botonx23 & "</div>"

'[UI-ANT] xhtml = Replace(xhtml, "[BOTONERA_PAGINACION]", misBotones)

'[UI-ANT] end if

'[UI-ANT] xhtml = xhtml & "<script type='text/javascript'>"
'[UI-ANT] xhtml = xhtml & " function ejecutar(formu,dsc,acc){ "
'[UI-ANT] xhtml = xhtml & " var r = confirm(dsc); "
'[UI-ANT] xhtml = xhtml & " if (r){ "
'[UI-ANT] xhtml = xhtml & " if (acc == 'ins') {document.getElementById(formu).action=document.getElementById(formu).action + '&opc2=insertar';}"
'[UI-ANT] xhtml = xhtml & " document.getElementById(formu).submit(); "
'[UI-ANT] xhtml = xhtml & " }else{ "
'[UI-ANT] xhtml = xhtml & " return false;"
'[UI-ANT] xhtml = xhtml & " }} "
'[UI-ANT] xhtml = xhtml & "</script>"

'[UI-ANT] xhtml = xhtml & "<script type='text/javascript'>"
'[UI-ANT] xhtml = xhtml & " $( function() {"
'[UI-ANT] xhtml = xhtml & "$.datepicker.regional['es'] = {"
'[UI-ANT] xhtml = xhtml & "closeText: 'Cerrar',"
'[UI-ANT] xhtml = xhtml & "prevText: '< Ant',"
'[UI-ANT] xhtml = xhtml & "nextText: 'Sig >',"
'[UI-ANT] xhtml = xhtml & "currentText: 'Hoy',"
'[UI-ANT] xhtml = xhtml & "monthNames: ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'],"
'[UI-ANT] xhtml = xhtml & "monthNamesShort: ['Ene','Feb','Mar','Abr', 'May','Jun','Jul','Ago','Sep', 'Oct','Nov','Dic'],"
'[UI-ANT] xhtml = xhtml & "dayNames: ['Domingo', 'Lunes', 'Martes', 'Miercoles', 'Jueves', 'Viernes', 'Sabado'],"
'[UI-ANT] xhtml = xhtml & "dayNames: ['Domingo', 'Lunes', 'Martes', 'Mi\u00e9rcoles', 'Jueves', 'Viernes', 'S\u00e1bado'],"
'[UI-ANT] xhtml = xhtml & "dayNamesShort: ['Dom','Lun','Mar','Mi?','Juv','Vie','S?b'],"
'[UI-ANT] xhtml = xhtml & "dayNamesShort: ['Dom','Lun','Mar','Mi\u00e9','Jue','Vie','S\u00e1b'],"
'[UI-ANT] xhtml = xhtml & "dayNamesMin: ['Do','Lu','Ma','Mi','Ju','Vi','Sa'],"
'[UI-ANT] xhtml = xhtml & "weekHeader: 'Sm',"
'[UI-ANT] xhtml = xhtml & "dateFormat: 'dd/mm/yy',"
'[UI-ANT] xhtml = xhtml & "firstDay: 1,"
'[UI-ANT] xhtml = xhtml & "isRTL: false,"
'[UI-ANT] xhtml = xhtml & "showMonthAfterYear: false,"
'[UI-ANT] xhtml = xhtml & "yearSuffix: ''"
'[UI-ANT] xhtml = xhtml & "};"
'[UI-ANT] xhtml = xhtml & "$.datepicker.setDefaults($.datepicker.regional['es']);"
'[UI-ANT] xhtml = xhtml & "});"

'[UI-ANT] xhtml = xhtml & "$( function() { "
'[UI-ANT] xhtml = xhtml & "$( '#datepicker' ).datepicker(); "
'[UI-ANT] xhtml = xhtml & "} ); "
'[UI-ANT] xhtml = xhtml & "</script>"




'[UI-ANT] xhtml = xhtml & xform & xforma & xformu & formx1 & formx2 & xfrm & xagregar & shtml & formx22 & formx23 & xformb & "</body></html>"

'[UI-ANT] response.write xhtml

'[UI-ANT] Function armaCombo(sql,cn,clave,valor,texto,comboid)

	'[UI-ANT] Set nrs = Exec(cn, sql)

	'[UI-ANT] if comboid = "resp" then
		'[UI-ANT] html = html & "<select style='font-size:12pt;height:80%;color:#000000;border:solid 1px #000000;' id='" & comboid & "' name='" & comboid & "' onchange=""document.getElementById('xformu').submit();"" disabled>"
	'[UI-ANT] else
		'[UI-ANT] html = html & "<select style='font-size:12pt;height:80%;color:#000000;border:solid 1px #000000;' id='" & comboid & "' name='" & comboid & "' onchange=""document.getElementById('xformu').submit();"">"
	'[UI-ANT] end if

	'[UI-ANT] if comboid = "resp" then

	'[UI-ANT] Set rf = Exec(cn, "sp_cg_Responsables @Modo='id_nombre', @IdSeg=" & SqlStr(session("usuario")))

		'[UI-ANT] if not rf.eof then html = html & "<option value='" & rf(0) & "'>" & rf(1) & "</option>"

	'[UI-ANT] else

		'[UI-ANT] html = html & "<option value='0'>S / D</option>"

		'[UI-ANT] Do While Not nrs.eof

			'[UI-ANT] If cstr(clave) = cstr(nrs(valor)) Then

				'[UI-ANT] html = html & "<option value='" & nrs(valor) & "' selected>" & nrs(texto) & "</option>"

			'[UI-ANT] Else

				'[UI-ANT] html = html & "<option value='" & nrs(valor) & "'>" & nrs(texto) & "</option>"

			'[UI-ANT] End If

			'[UI-ANT] nrs.movenext
		'[UI-ANT] Loop

	'[UI-ANT] end if

	'[UI-ANT] armaCombo = html & "</select>"

'[UI-ANT] End Function


' =========================================================================
' [UI] PANTALLA con el sistema visual (IncFile/ui_componentes.inc)
' Agenda: recordatorios (listado, alta, detalle con copias). Las reuniones (opc det/edicion)
' las sigue resolviendo agenda_ext.inc, igual que antes y antes que nada.
' El cuerpo anterior completo queda arriba, comentado con '[UI-ANT].
' Cambios: consultas seguras (asunto, ids y copias se pegaban tal cual), fecha validada,
' copias con casillas, fecha de aviso en dd/mm/aaaa y Vista semanal en la misma ventana.
' =========================================================================
Dim usuario, accTop, puedeAlta, puedeElim, hoySerial, cantRec, filas, diaAviso, estadoAviso, urlDet, clickFila
Dim asunto, nombreResp, cantCopias, copiasHtml, agregarHtml, esDueno, enviado, rp, rf, rm, rz, rx, rt
Dim pagAnt, pagSig, cantDisp

xhtml = headhtml()
xhtml = xhtml & "<body>"

opc = request.querystring("opc")
xmenu = request.querystring("xmenu")
xtmenu = request.querystring("xtmenu")

CantRegMostrar = 15
CantRegMover = 0
If IsNumeric(Request.querystring("CantRegMover") & "") Then CantRegMover = CDbl(Request.querystring("CantRegMover"))
if CantRegMover < 0 then CantRegMover = 0

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

if request.querystring("opc2") & "" <> "" then opc = request.querystring("opc2")

' Reuniones: las resuelve agenda_ext.inc (sin cambios)
if opc = "edicion" then opc = "det"
if AgendaExtDebeManejar(opc) then
	AgendaExtRender Cnn, opc, xmenu, xtmenu
	Response.End
end if

' Usuario: la sesion puede venir vacia; si falta, UserLogon()
usuario = Trim(Session("usuario") & "")
If usuario = "" Then usuario = Trim(UserLogon() & "")
puedeAlta = (getPermisosAtrm("CG_AGENDA", "ALTA") = 1)
puedeElim = (getPermisosAtrm("CG_AGENDA", "ELIM") = 1)
hoySerial = CLng(Date) - CLng(CDate("01-01-1900"))

xhtml = xhtml & "<div id='wrapper' role='banner'>" & cabecera() & "</div>"

xform = "<form id='xform' class='form-oculto' method='post' action='agenda.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"

' =============================================================================
' ALTA DE RECORDATORIO
' NOTA: hoy no se ejecuta: agenda_ext.inc resuelve opc=cargar y opc=det antes de llegar aca
' (el alta y el detalle son los de las reuniones). Se conserva por si se vuelve a separar.
' =============================================================================
if opc = "cargar" then

	accTop = UiBoton("Volver", "secondary", "document.getElementById('xform').submit();", True, "")
	xhtml = xhtml & UiPantallaAbre("Nuevo recordatorio", accTop, "narrow")

	xhtml = xhtml & "<form id='xformu' method='post' action='agenda.asp?opc=cargar&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "' onsubmit='return false;'>"
	xhtml = xhtml & UiCardAbre("Datos del recordatorio", "") & UiNotaObligatorios() & UiFormGridAbre()
	xhtml = xhtml & UiCampoLectura("Responsable", UiDato(getName(cnn, "", "cor_responsables", usuario)), True)
	xhtml = xhtml & UiCampo("asun", "Asunto", UiInputTexto("asun", True, "maxlength='1000' autocomplete='off'"), True, True)
	xhtml = xhtml & UiCampo("datepicker", "Avisar el d&iacute;a", UiInputTexto("datepicker", True, "autocomplete='off' placeholder='dd/mm/aaaa'"), True, False)
	xhtml = xhtml & UiFormGridCierra() & UiCardCierra()

	' Copias: una casilla por persona (antes un desplegable Si/No por cada una). Mismo nombre de campo: u<login> = 1
	xhtml = xhtml & "<section class='ui-card' aria-labelledby='tituloCopias'><h2 class='ui-card__title' id='tituloCopias'>Enviar copia a</h2>"
	Set rx = Exec(cnn, "sp_cg_Responsables @Modo='listar_activos_sin', @ExcluirId=" & SqlStr(usuario))
	If rx.eof Then
		xhtml = xhtml & "<div class='ui-empty'>No hay otras personas activas.</div>"
	Else
		xhtml = xhtml & "<fieldset class='ui-checks'><legend class='ui-sr-only'>Personas que reciben copia</legend>"
		Do While Not rx.eof
			xhtml = xhtml & "<label class='ui-opcion'><input class='ui-check' type='checkbox' name='u" & Server.HTMLEncode(rx(0) & "") & "' value='1'><span>" & UiTexto(Trim(rx(1) & " " & rx(2))) & "</span></label>"
			rx.movenext
		Loop
		xhtml = xhtml & "</fieldset>"
	End If
	xhtml = xhtml & "</section>"

	xhtml = xhtml & UiFormAcciones(UiBotonPrincipal("Cargar recordatorio", "cargarRecordatorio();"))
	xhtml = xhtml & "</form>"
	xhtml = xhtml & UiPantallaCierra()

' =============================================================================
' MODIFICAR ASUNTO
' =============================================================================
elseif opc = "modif" then

	id = request.querystring("id")
	idresp = request.querystring("idresp")
	ccs = AgLimpiarIds(request.querystring("ccs"))
	If IsNumeric(id & "") And Trim(id & "") <> "" Then
		' [UI] Asunto escapado e id validado (antes se pegaban tal cual)
		ExecNonQuery Cnn, "sp_cg_Agenda @Modo='actualizar_asunto', @Id=" & CLng(id) & ", @Asunto=" & SqlStr(request.form("asun"))
	End If
	xhtml = xhtml & "<form name='initx' method='post' action='agenda.asp?opc=det&id=" & Server.URLEncode(id & "") & "&idresp=" & Server.URLEncode(idresp & "") & "&ccs=" & Server.URLEncode(ccs) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&ok=1'></form>"
	xhtml = xhtml & "<script type='text/javascript'>document.forms['initx'].submit();</script>"

' =============================================================================
' INSERTAR RECORDATORIO
' =============================================================================
elseif opc = "insertar" then

	Set rz = Cnn.Execute("sp_cg_Responsables @Modo='listar_ids_activos_sin', @ExcluirId=" & SqlStr(usuario))
	copias = ""
	sep = ""
	Do While Not rz.eof
		num = rz(0)
		If request.form("u" & num) = "1" Then
			copias = copias & sep & rz(0)
			sep = ","
		End If
		rz.movenext
	Loop

	' [UI] Fecha y asunto validados: una fecha mal escrita rompia la pagina
	If Trim(request.form("asun") & "") = "" Or Not IsDate(request.form("datepicker") & "") Then
		xhtml = xhtml & "<script type='text/javascript'>alert('Complet\u00e1 el asunto y una fecha v\u00e1lida.'); window.history.back();</script>"
	Else
		sDiaAviso = CLng(CDate(request.form("datepicker"))) - CLng(CDate("01-01-1900"))
		sFecHoy = hoySerial
		ExecNonQuery Cnn, "sp_cg_Agenda @Modo='insertar', @Asunto=" & SqlStr(request.form("asun")) & ", @DiaAviso=" & sDiaAviso & ", @IdRespon=" & SqlStr(usuario) & ", @Copias=" & SqlStr(copias) & ", @FecCarga=" & sFecHoy
		XXX = SYSLOGON(cnn, "AGREGA_REC")
		xhtml = xhtml & "<form name='init' method='post' action='agenda.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&ok=alta'></form>"
		xhtml = xhtml & "<script type='text/javascript'>document.forms['init'].submit();</script>"
	End If

' =============================================================================
' DETALLE: asunto y copias
' =============================================================================
elseif opc = "det" then

	id = request.querystring("id")
	idresp = request.querystring("idresp")
	idagre = AgLimpiarIds(request.querystring("idagre"))
	ccs = AgLimpiarIds(request.querystring("ccs"))
	xopt = request.querystring("xopt")
	If IsNumeric(id & "") And Trim(id & "") <> "" Then id = CLng(id) Else id = 0

	Set rp = Cnn.Execute("sp_cg_Agenda @Modo='obtener', @Id=" & id)
	enviado = False
	If Not rp.eof Then
		If rp(0) = "SI" Then enviado = True
	End If
	esDueno = (CStr(idresp & "") = CStr(usuario))

	' Agregar una copia (solo el responsable, si el aviso no se envio)
	If xopt = "agre" And idagre <> "" And esDueno And Not enviado Then
		If ccs <> "" Then ccs = ccs & "," & idagre Else ccs = idagre
		ExecNonQuery Cnn, "sp_cg_Agenda @Modo='actualizar_copias', @Id=" & id & ", @Copias=" & SqlStr(ccs)
	End If

	accTop = UiBoton("Volver", "secondary", "document.getElementById('xform').submit();", True, "")
	xhtml = xhtml & UiPantallaAbre("Recordatorio #" & id, accTop, "narrow")

	If request.querystring("ok") = "1" Then xhtml = xhtml & "<div class='ui-alert ui-alert--success' role='status'>Asunto actualizado.</div>"

	If rp.eof Then
		xhtml = xhtml & UiCardAbre("Recordatorio", "") & "<div class='ui-empty'>No se encontr&oacute; el recordatorio.</div>" & UiCardCierra()
	Else
		asunto = rp(1) & ""
		nombreResp = getName(cnn, "", "cor_responsables", idresp)

		' --- Datos ---
		xhtml = xhtml & UiCardAbre("Recordatorio", "")
		xhtml = xhtml & "<div class='ui-fields'>"
		xhtml = xhtml & UiCampoLectura("Responsable", UiDato(nombreResp), False)
		If enviado Then
			xhtml = xhtml & UiCampoLectura("Aviso", UiBadge("Enviado", "green"), False)
		Else
			xhtml = xhtml & UiCampoLectura("Aviso", UiBadge("Pendiente", "amber"), False)
		End If
		xhtml = xhtml & "</div>"
		If esDueno And Not enviado Then
			xhtml = xhtml & "<form id='act' class='ui-agenda-asunto' method='post' action='agenda.asp?opc=modif&id=" & id & "&ccs=" & Server.URLEncode(ccs) & "&idresp=" & Server.URLEncode(idresp & "") & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"
			xhtml = xhtml & UiCampo("asun", "Asunto", UiInputTexto("asun", True, "maxlength='1000' value='" & Server.HTMLEncode(asunto) & "'"), True, True)
			xhtml = xhtml & "<div class='ui-card__acciones'>" & UiBotonEnviar("Guardar asunto", True, "") & "</div></form>"
		Else
			xhtml = xhtml & UiCampoLectura("Asunto", "<div class='ui-prose'>" & UiTexto(asunto) & "</div>", True)
			If enviado Then
				xhtml = xhtml & "<p class='ui-table__nota'>El aviso ya se envi&oacute;: el recordatorio no se puede modificar.</p>"
			Else
				xhtml = xhtml & "<p class='ui-table__nota'>Solo el responsable puede modificar el recordatorio.</p>"
			End If
		End If
		xhtml = xhtml & UiCardCierra()

		' --- Copias ---
		xfrm = ""
		copiasHtml = ""
		cantCopias = 0
		If ccs <> "" Then
			Set rf = Cnn.Execute("sp_cg_Responsables @Modo='nombre_id_in', @Ids=" & SqlStr(ccs))
			Do While Not rf.eof
				copiasHtml = copiasHtml & "<li class='ui-copia'><span>" & UiTexto(rf(0)) & "</span>"
				If esDueno And Not enviado Then
					xfrm = xfrm & "<form id='elim" & DomId(rf(1)) & "' class='form-oculto' method='post' action='agenda.asp?opc=elim&idag=" & id & "&ccs=" & Server.URLEncode(ccs) & "&idresp=" & Server.URLEncode(idresp & "") & "&idelim=" & Server.URLEncode(rf(1) & "") & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
					copiasHtml = copiasHtml & UiBotonIcono("&times;", "Quitar copia a " & UiTexto(rf(0)), "danger", "document.getElementById('elim" & DomId(rf(1)) & "').submit();", puedeElim, "Sin permiso para eliminar")
				End If
				copiasHtml = copiasHtml & "</li>"
				cantCopias = cantCopias + 1
				rf.movenext
			Loop
		End If
		xhtml = xhtml & UiCardAbre("Copias", "")
		If cantCopias = 0 Then
			xhtml = xhtml & "<div class='ui-empty'>No registra copias.</div>"
		Else
			xhtml = xhtml & "<ul class='ui-copias'>" & copiasHtml & "</ul>"
		End If

		' --- Agregar copia (solo el responsable, si el aviso no se envio) ---
		If esDueno And Not enviado Then
			If ccs <> "" Then
				Set rm = Cnn.Execute("sp_cg_Responsables @Modo='listar_not_in', @Ids=" & SqlStr(ccs) & ", @ExcluirId=" & SqlStr(idresp))
			Else
				Set rm = Cnn.Execute("sp_cg_Responsables @Modo='id_nombre_sin', @ExcluirId=" & SqlStr(idresp))
			End If
			agregarHtml = ""
			cantDisp = 0
			Do While Not rm.eof
				agregarHtml = agregarHtml & "<li><a class='ui-copia-agregar' href='agenda.asp?opc=det&xopt=agre&id=" & id & "&ccs=" & Server.URLEncode(ccs) & "&idresp=" & Server.URLEncode(idresp & "") & "&idagre=" & Server.URLEncode(rm(0) & "") & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "' aria-label='Agregar copia a " & UiTexto(rm(1)) & "'>+ " & UiTexto(rm(1)) & "</a></li>"
				cantDisp = cantDisp + 1
				rm.movenext
			Loop
			If cantDisp > 0 Then xhtml = xhtml & "<p class='ui-copias__titulo'>Agregar copia a</p><ul class='ui-copias-agregar'>" & agregarHtml & "</ul>"
		End If
		xhtml = xhtml & UiCardCierra()
	End If
	xhtml = xhtml & UiPantallaCierra()

' =============================================================================
' QUITAR UNA COPIA
' =============================================================================
elseif opc = "elim" then

	idresp = request.querystring("idresp")
	idelim = request.querystring("idelim") & ""
	idag = request.querystring("idag")
	If IsNumeric(idag & "") And Trim(idag & "") <> "" Then idag = CLng(idag) Else idag = 0

	If puedeElim And CStr(idresp & "") = CStr(usuario) Then
		Set rt = Cnn.Execute("sp_cg_Agenda @Modo='obtener_copias', @Id=" & idag)
		If Not rt.eof Then
			xcopy = Split(rt(0) & "", ",")
			xcopyx = ""
			sep = ""
			For i = 0 To UBound(xcopy)
				If xcopy(i) <> idelim Then
					xcopyx = xcopyx & sep & xcopy(i)
					sep = ","
				End If
			Next
			rt.Close
			ExecNonQuery Cnn, "sp_cg_Agenda @Modo='actualizar_copias', @Id=" & idag & ", @Copias=" & SqlStr(xcopyx)
		End If
	End If
	ccs = ""
	Set rt = Cnn.Execute("sp_cg_Agenda @Modo='obtener_copias', @Id=" & idag)
	If Not rt.eof Then ccs = rt(0) & ""
	xhtml = xhtml & "<form name='init' method='post' action='agenda.asp?opc=det&id=" & idag & "&ccs=" & Server.URLEncode(ccs) & "&idresp=" & Server.URLEncode(idresp & "") & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
	xhtml = xhtml & "<script type='text/javascript'>document.forms['init'].submit();</script>"

' =============================================================================
' LISTADO DE RECORDATORIOS
' =============================================================================
else

	XXX = SYSLOGON(cnn, "MENU_AGENDA")

	xformA = "<form id='xformA' class='form-oculto' method='post' action='agenda.asp?opc=cargar&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
	xform = "<form id='xform' class='form-oculto' method='post' action='inicio.asp'></form>"
	' Vista semanal: en la misma ventana (antes una ventana emergente)
	' opc=cargar lo resuelve agenda_ext.inc: abre el alta de reunion (de ahi el nombre del boton)
	accTop = UiBoton("+ Nueva reuni&oacute;n", "light", "document.getElementById('xformA').submit();", puedeAlta, "Sin permiso para agregar.") & _
		"<a class='ui-button ui-button--secondary ui-button--sm' href='agenda_semana.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>Vista semanal</a>" & _
		UiBoton("Volver", "secondary", "document.getElementById('xform').submit();", True, "")
	xhtml = xhtml & UiPantallaAbre("Agenda", accTop, "wide")
	If request.querystring("ok") = "alta" Then xhtml = xhtml & "<div class='ui-alert ui-alert--success' role='status'>Recordatorio cargado.</div>"


	' =========================================================================
	' [AGENDA] Filtros de la lista: por defecto "Mis recordatorios" (el usuario es responsable, esta en
	' COPIAS o es corresponsable); "Todos" muestra todos. Busqueda de texto en asunto o mensaje.
	' Requiere los parametros @Respon / @Texto de sp_cg_Agenda (DB/Cambios_Agenda_Listar_Filtros.sql).
	' =========================================================================
	Dim agVista, agTexto, agQs, agSp, agAviso, agFiltros, agHayFiltro
	agVista = LCase(Trim(Request.QueryString("vista") & ""))
	If agVista <> "todos" Then agVista = "mis"
	agTexto = Trim(Request.QueryString("q") & "")
	If Len(agTexto) > 200 Then agTexto = Left(agTexto, 200)
	agQs = "&vista=" & agVista & "&q=" & Server.URLEncode(agTexto)
	agSp = "sp_cg_Agenda @Modo='listar'"
	If agVista = "mis" And usuario <> "" Then agSp = agSp & ", @Respon=" & SqlStr(usuario)
	If agTexto <> "" Then agSp = agSp & ", @Texto=" & SqlStr(agTexto)
	agAviso = ""
	On Error Resume Next
	Set rs = Cnn.Execute(agSp)
	If Err.Number <> 0 Then
		Err.Clear
		Set rs = Cnn.Execute("sp_cg_Agenda @Modo='listar'")
		agAviso = "<div class='ui-alert ui-alert--warning' role='alert'>Los filtros todav&iacute;a no est&aacute;n disponibles en la base (falta aplicar DB/Cambios_Agenda_Listar_Filtros.sql). Se muestra la lista completa.</div>"
	End If
	On Error Goto 0
	agHayFiltro = (agVista = "todos" Or agTexto <> "")
	agFiltros = "<form id='xFiltroAg' class='ui-card__filtros' method='get' action='agenda.asp'>" & _
		"<input type='hidden' name='xmenu' value='" & Server.HTMLEncode(xmenu & "") & "'><input type='hidden' name='xtmenu' value='" & Server.HTMLEncode(xtmenu & "") & "'>" & _
		UiFiltroCompacto("vista", "Mostrar", "<select class='ui-select' id='vista' name='vista' onchange=""document.getElementById('xFiltroAg').submit();"">" & _
			"<option value='mis'" & IIf(agVista = "mis", " selected", "") & ">Mis recordatorios (responsable o participante)</option>" & _
			"<option value='todos'" & IIf(agVista = "todos", " selected", "") & ">Todos</option></select>") & _
		UiFiltroCompacto("q", "Buscar en asunto o mensaje", "<input class='ui-input' type='search' id='q' name='q' maxlength='200' placeholder='Buscar en asunto o mensaje' value='" & Server.HTMLEncode(agTexto) & "'>") & _
		"<button type='submit' class='ui-button ui-button--secondary ui-button--sm'>Buscar</button>"
	If agHayFiltro Then agFiltros = agFiltros & "<a class='ui-button ui-button--quiet ui-button--sm' href='agenda.asp?xmenu=" & Server.URLEncode(xmenu & "") & "&xtmenu=" & Server.URLEncode(xtmenu & "") & "'>Limpiar</a>"
	agFiltros = agFiltros & "</form>"
	xhtml = xhtml & agAviso

	xhtml = xhtml & UiCardAbre("Recordatorios", agFiltros)
	xhtml = xhtml & UiTablaAbre("Recordatorios: asunto, responsable, d&iacute;a de aviso y si ya se envi&oacute;")
	xhtml = xhtml & "<thead><tr>" & UiTh("<span class='ui-sr-only'>Acciones</span>", "ui-table__col--action") & UiTh("Asunto", "") & UiTh("Responsable", "") & _
		UiTh("Aviso", "") & UiTh("Estado", "") & "</tr></thead><tbody>"

	' (la consulta se hizo arriba, con los filtros)
	formx1 = ""
	If Not rs.eof Then rs.Move (CantRegMover)
	PosReg = 0
	filas = ""
	Do While (Not rs.eof) And (PosReg < CantRegMostrar)
		formx1 = formx1 & "<form id='formx1" & rs(0) & "' class='form-oculto' method='post' action='eliminar.asp?opc=rec&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=agenda.asp'></form>"

		diaAviso = CLng("0" & rs(2))
		If rs(4) & "" = "SI" Then
			estadoAviso = UiBadge("Enviado", "green")
		ElseIf diaAviso > 0 And diaAviso < hoySerial Then
			estadoAviso = UiBadge("Vencido", "red")
		ElseIf diaAviso = hoySerial Then
			estadoAviso = UiBadge("Hoy", "blue")
		Else
			estadoAviso = UiBadge("Pendiente", "amber")
		End If

		urlDet = "agenda.asp?opc=det&id=" & rs(0) & "&ccs=" & Server.URLEncode(rs(6) & "") & "&idresp=" & Server.URLEncode(rs(3) & "") & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu
		' [PERMISOS] La fila abre el detalle con cualquier permiso sobre la agenda (consultar, modificar o eliminar); antes solo con eliminar
		If puedeElim Or getPermisosAtrm("CG_AGENDA", "CONS") = 1 Or getPermisosAtrm("CG_AGENDA", "MODIF") = 1 Then
			clickFila = " class='ui-table__cell--link' onclick=""document.location='" & urlDet & "'"""
		Else
			clickFila = ""
		End If

		filas = filas & "<tr>" & _
			"<td>" & UiBotonIcono("&times;", "Eliminar recordatorio " & rs(0), "danger", "borrar('formx1" & rs(0) & "');", puedeElim, "Sin permiso para eliminar") & "</td>" & _
			"<td" & clickFila & "><span class='ui-table__titulo' title='" & Server.HTMLEncode(rs(1) & "") & "'>" & UiTexto(AgCorto(rs(1), 70)) & "</span></td>" & _
			"<td" & clickFila & ">" & UiDato(getName(Cnn, "", "cor_responsables", rs(3))) & "</td>" & _
			"<td" & clickFila & "><span class='ui-table__nowrap'>" & UiFecha(rs(2)) & "</span></td>" & _
			"<td" & clickFila & ">" & estadoAviso & "</td>" & _
			"</tr>"
		PosReg = PosReg + 1
		rs.movenext
	Loop
	If PosReg = 0 Then
		If agTexto <> "" Then
			xhtml = xhtml & UiFilaVacia(5, "No hay recordatorios que coincidan con la b&uacute;squeda.")
		ElseIf agVista = "mis" Then
			xhtml = xhtml & UiFilaVacia(5, "No ten&eacute;s recordatorios propios (como responsable o participante). Eleg&iacute; ""Todos"" para ver los dem&aacute;s.")
		Else
			xhtml = xhtml & UiFilaVacia(5, "No hay recordatorios cargados.")
		End If
	Else
		xhtml = xhtml & filas
	End If
	xhtml = xhtml & UiTablaCierra()

	' --- Paginacion (mismos formularios de antes) ---
	If CantRegMover <> 0 Then
		formx22 = "<form id='formx22' class='form-oculto' method='post' action='agenda.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&cantregmover=" & (CantRegMover - CantRegMostrar) & agQs & "'></form>"
		pagAnt = UiBoton("Anterior", "secondary", "document.getElementById('formx22').submit();", True, "")
	Else
		formx22 = ""
		pagAnt = UiBoton("Anterior", "secondary", "", False, "Primera p&aacute;gina")
	End If
	If Not rs.eof Then
		formx23 = "<form id='formx23' class='form-oculto' method='post' action='agenda.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&cantregmover=" & (CantRegMover + CantRegMostrar) & agQs & "'></form>"
		pagSig = UiBoton("Siguiente", "secondary", "document.getElementById('formx23').submit();", True, "")
	Else
		formx23 = ""
		pagSig = UiBoton("Siguiente", "secondary", "", False, "&Uacute;ltima p&aacute;gina")
	End If
	xhtml = xhtml & UiPaginacion(CantRegMover + 1, CantRegMover + PosReg, pagAnt, pagSig)
	xhtml = xhtml & UiCardCierra()
	xhtml = xhtml & UiPantallaCierra()

end if

' --- Scripts ---
xhtml = xhtml & "<script type='text/javascript'>"
xhtml = xhtml & " function borrar(formu){ if (confirm('\u00bfEliminar el recordatorio?')) { document.getElementById(formu).submit(); } }"
xhtml = xhtml & " function cargarRecordatorio(){"
xhtml = xhtml & "   var a = document.getElementById('asun'), f = document.getElementById('datepicker');"
xhtml = xhtml & "   if (!a.value.trim()) { alert('Complet\u00e1 el asunto.'); a.focus(); return; }"
xhtml = xhtml & "   if (!/^\d{2}\/\d{2}\/\d{4}$/.test(f.value.trim())) { alert('Eleg\u00ed el d\u00eda de aviso (dd/mm/aaaa).'); f.focus(); return; }"
xhtml = xhtml & "   if (!confirm('\u00bfCargar el recordatorio?')) return;"
xhtml = xhtml & "   var fm = document.getElementById('xformu'); fm.action = fm.action + '&opc2=insertar'; fm.submit();"
xhtml = xhtml & " }"
xhtml = xhtml & "</script>"

xhtml = xhtml & "<script type='text/javascript'>"
xhtml = xhtml & " $( function() {"
xhtml = xhtml & "$.datepicker.regional['es'] = {"
xhtml = xhtml & "closeText: 'Cerrar',"
xhtml = xhtml & "prevText: '< Ant',"
xhtml = xhtml & "nextText: 'Sig >',"
xhtml = xhtml & "currentText: 'Hoy',"
xhtml = xhtml & "monthNames: ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio', 'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre'],"
xhtml = xhtml & "monthNamesShort: ['Ene','Feb','Mar','Abr', 'May','Jun','Jul','Ago','Sep', 'Oct','Nov','Dic'],"
xhtml = xhtml & "dayNames: ['Domingo', 'Lunes', 'Martes', 'Mi\u00e9rcoles', 'Jueves', 'Viernes', 'S\u00e1bado'],"
xhtml = xhtml & "dayNamesShort: ['Dom','Lun','Mar','Mi\u00e9','Jue','Vie','S\u00e1b'],"
xhtml = xhtml & "dayNamesMin: ['Do','Lu','Ma','Mi','Ju','Vi','Sa'],"
xhtml = xhtml & "weekHeader: 'Sm',"
xhtml = xhtml & "dateFormat: 'dd/mm/yy',"
xhtml = xhtml & "firstDay: 1,"
xhtml = xhtml & "isRTL: false,"
xhtml = xhtml & "showMonthAfterYear: false,"
xhtml = xhtml & "yearSuffix: ''"
xhtml = xhtml & "};"
xhtml = xhtml & "$.datepicker.setDefaults($.datepicker.regional['es']);"
xhtml = xhtml & "$( '#datepicker' ).datepicker({ minDate: 0 });"
xhtml = xhtml & "});"
xhtml = xhtml & "</script>"

xhtml = xhtml & xform & xformA & formx1 & xfrm & formx22 & formx23 & "</body></html>"

response.write xhtml


' =============================================================================
' FUNCIONES DE ESTA PANTALLA
' =============================================================================

' Lista de logins (copias) sin caracteres peligrosos: letras, numeros, coma, punto, guion y guion bajo
Function AgLimpiarIds(v)
	Dim s, i, c
	s = ""
	For i = 1 To Len(v & "")
		c = Mid(v & "", i, 1)
		If (c >= "0" And c <= "9") Or (LCase(c) >= "a" And LCase(c) <= "z") Or c = "," Or c = "." Or c = "-" Or c = "_" Then s = s & c
	Next
	AgLimpiarIds = s
End Function

' Texto acortado con puntos suspensivos (el completo queda en el title)
Function AgCorto(v, n)
	Dim t
	t = Trim(v & "")
	If Len(t) > n Then t = RTrim(Left(t, n)) & "..."
	AgCorto = t
End Function
%>