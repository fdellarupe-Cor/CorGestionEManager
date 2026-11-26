<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/agenda_ext.inc" -->
<%

xhtml = headhtml()
xhtml = xhtml & "<body>"

opc = request.querystring("opc")
xmenu = request.querystring("xmenu")
xtmenu = request.querystring("xtmenu")

CantRegMostrar = 15
CantRegMover = CDbl(Request.querystring("CantRegMover"))
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

if opc = "edicion" then opc = "det"
if AgendaExtDebeManejar(opc) then
	AgendaExtRender Cnn, opc, xmenu, xtmenu
	Response.End
end if

xhtml = xhtml & cabecera()

'xhtml = xhtml & "<div style='height:8%;width:100%;border-top:solid 5px #FF3D3D;border-bottom:double 6px #001BFF;'>"

xhtml = xhtml & "<div class='titulo-seccion ui-page-header' style='padding: 1rem 2rem;'>"
xhtml = xhtml & "<div class='ui-toolbar' style='display:flex;align-items:center;justify-content:space-between;'>"

if opc = "" then
	xhtml = xhtml & "<div><h2>RECORDATORIOS</h2></div>"
else
	xhtml = xhtml & "<div><h2>AGENDA</h2></div>"
end if

if opc = "" then
	xform = "<form id='xform' class='form-oculto' method='post' action='index.asp'></form>"
else
	xform = "<form id='xform' class='form-oculto' method='post' action='agenda.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
end if

xboton = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('xform').submit();"">"

xhtml = xhtml & "<div class='ui-toolbar__actions' style='display:flex; gap: 1rem;'>"

if opc <> "cargar" and opc <> "det" then

	xforma = "<form id='xformA' class='form-oculto' method='post' action='agenda.asp?opc=cargar&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
	if getPermisos(xmenu,xtmenu,"ALTA") = 1 then
		xbotona = "<input class='bot ui-button ui-button--primary' type='button' value='CARGAR RECORDATORIO' onclick=""document.getElementById('xformA').submit();"">"
	else
		xbotona = "<input class='botsp ui-button' type='button' value='CARGAR RECORDATORIO' onclick=""alert('Sin permiso para agregar.');"">"
	end if
	xhtml = xhtml & xbotona
	xbotonSemana = "<input class='bot ui-button' type='button' value='VISTA SEMANAL' onclick=""window.open('agenda_semana.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "','agenda_semana','width=1520,height=900,resizable=yes,scrollbars=yes');"">"
	xhtml = xhtml & xbotonSemana
	xhtml = xhtml & "[BOTONERA_PAGINACION]"
	xhtml = xhtml & xboton

end if

if opc = "cargar" or opc = "det" then

	xformb = "<form id='xformB' class='form-oculto' method='post' action='agenda.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
	xbotonb = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('xformB').submit();"">"
	xhtml = xhtml & xbotonb

end if

xhtml = xhtml & "</div>" ' Fin ui-toolbar__actions

xhtml = xhtml & "</div>" ' Fin ui-toolbar
xhtml = xhtml & "</div>" ' Fin titulo-seccion

if opc = "cargar" then

resp = request.form("resp")
if resp & "" = "" then resp = 0

xhtml = xhtml & "<form id='xformu' method='post' action='agenda.asp?opc=cargar&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'>"

' --- INICIO NORMALIZACION: NUEVO RECORDATORIO ---

'xhtml = xhtml & "<table class='tabl' style='width:35%;' align=center cellspacing=0 cellpadding=3 >"
xhtml = xhtml & "<div class='inner ui-main' style='padding: 2rem;'>"
xhtml = xhtml & "<div class='bloque-datos ui-panel' style='max-width: 600px; margin: 0 auto;'>"

'xhtml = xhtml & "<tr>"
'xhtml = xhtml & "<td class='fil' style='font-size:12pt;' colspan=2 align=center><u><b>NUEVO RECORDATORIO</b></u></td>"
'xhtml = xhtml & "</tr>"
xhtml = xhtml & "<div class='bloque-header ui-panel__header' style='text-align: center; text-transform: uppercase;'>NUEVO RECORDATORIO</div>"

'xhtml = xhtml & "<tr>"
'xhtml = xhtml & "<td colspan=2 >&nbsp;</td>"
'xhtml = xhtml & "</tr>"
xhtml = xhtml & "<div class='bloque-body ui-panel__body'>"
xhtml = xhtml & "<div class='grid-datos ui-field-grid'>"

''RESPON
'xhtml = xhtml & "<tr>"
'xhtml = xhtml & "<td class='fil' >RESPONSABLE</td>"
xhtml = xhtml & "<div class='dato-label'>RESPONSABLE</div>"
	
'xhtml = xhtml & "<td class='fil' >"
'xhtml = xhtml & getName(cnn,"","cor_responsables",session("usuario"))
'xhtml = xhtml & "</td>"
'xhtml = xhtml & "</tr>"
xhtml = xhtml & "<div class='dato-value'>" & getName(cnn,"","cor_responsables",session("usuario")) & "</div>"

''ASUNTO
'xhtml = xhtml & "<tr>"
'xhtml = xhtml & "<td class='fil' >ASUNTO</font></td>"
xhtml = xhtml & "<div class='dato-label'>ASUNTO</div>"
	
'xhtml = xhtml & "<td class='fil' ><input class='inpu' style='width:80%;' id='asun' name='asun' type='text' /></td>"
'xhtml = xhtml & "</tr>"
xhtml = xhtml & "<div><input class='input-moderno ui-input' id='asun' name='asun' type='text' /></div>"

''DIA
'xhtml = xhtml & "<tr>"
'xhtml = xhtml & "<td class='fil' >AVISAR DIA</td>"
xhtml = xhtml & "<div class='dato-label'>AVISAR DIA</div>"
	
'xhtml = xhtml & "<td class='fil' ><input class='inpu' style='width:80%;' id='datepicker' name='datepicker' type='text'></td>"
'xhtml = xhtml & "</tr>"
xhtml = xhtml & "<div><input class='input-moderno ui-input' id='datepicker' name='datepicker' type='text'></div>"

''ESPACIO (Separador Visual)
'xhtml = xhtml & "<tr>"
'xhtml = xhtml & "<td colspan=2>&nbsp;</td>"
'xhtml = xhtml & "</tr>"
xhtml = xhtml & "<div style='grid-column: 1 / -1;'><hr style='border: 0; border-top: 1px solid #e2e8f0; margin: 1.5rem 0;'></div>"

''AVISAR A
'xhtml = xhtml & "<tr>"
'xhtml = xhtml & "<td class='fil' colspan=2>ENVIAR COPIA A:</td>"
'xhtml = xhtml & "</tr>"
xhtml = xhtml & "<div class='dato-label' style='grid-column: 1 / -1; margin-bottom: 0.5rem;'>ENVIAR COPIA A:</div>"

Set rx = Exec(cnn, "sp_cg_Responsables @Modo='listar_activos_sin', @ExcluirId=" & SqlStr(session("usuario")))

	do while not rx.eof

		chkbx = ""
		'xhtml = xhtml & "<td class='fil' >" & UCASE(rx(1)) & "</td>"
		' Usamos un padding-left para jerarquizar visualmente los nombres bajo el título "ENVIAR COPIA A:"
		xhtml = xhtml & "<div class='dato-label' style='padding-left: 1rem; color: #475569;'>" & UCASE(rx(1)) & "</div>"
		sel = "<select class='selec ui-select' style='width:30%;' name='u" & rx(0) & "'  >"
		sel = sel & "<option value='0'>NO</option>"
		sel = sel & "<option value='1'>SI</option>"
		sel = sel & "</select>"
		'xhtml = xhtml & "<td class='fil'>" & sel & "</td>"
		' Envolver en .dato-value activa el parche de tu CSS para selects
		xhtml = xhtml & "<div class='dato-value'>" & sel & "</div>"
		'xhtml = xhtml & "</tr>"

	rx.movenext
	loop

'xhtml = xhtml & "<tr>"
'xhtml = xhtml & "<td  colspan=2>&nbsp;</td>"
'xhtml = xhtml & "</tr>"

'xhtml = xhtml & "<tr>"
'xboton = "<input class='bot' type='button' value='Cargar' onclick=""ejecutar('xformu','Insertar','ins');"" />"
xboton = "<input class='btn-principal ui-button ui-button--primary' type='button' value='Cargar' onclick=""ejecutar('xformu','Insertar','ins');"" />"

'xhtml = xhtml & "<td class='fil' colspan=2 align=center>" & xboton & "</td>"
'xhtml = xhtml & "</tr>"
' Contenedor Flexbox para centrar el botón rompiendo la grilla
xhtml = xhtml & "<div style='grid-column: 1 / -1; display: flex; justify-content: center; margin-top: 1.5rem;'>"
xhtml = xhtml & xboton
xhtml = xhtml & "</div>"

'xhtml = xhtml & "</table>"
xhtml = xhtml & "</div>" ' Cierre de .grid-datos
xhtml = xhtml & "</div>" ' Cierre de .bloque-body
xhtml = xhtml & "</div>" ' Cierre de .bloque-datos
xhtml = xhtml & "</div>" ' Cierre de .inner

xhtml = xhtml & "</form>"

elseif opc = "modif" then

id = request.querystring("id")
idresp = request.querystring("idresp")
ccs = request.querystring("ccs")

ExecNonQuery Cnn, "sp_cg_Agenda @Modo='actualizar_asunto', @Id=" & id & ", @Asunto='" & request.form("asun") & "'"

xhtml = xhtml & "<FORM name='initx' method='post'  ACTION='agenda.asp?opc=det&id=" & id & "&idresp=" & Server.URLEncode(idresp & "") & "&ccs=" & Server.URLEncode(ccs & "") & "' >"
xhtml = xhtml & "<INPUT TYPE='SUBMIT' VALUE='INICIAR' style='display:none;' >"
xhtml = xhtml & "</FORM>"

xhtml = xhtml & "<script type='text/javascript'>"
xhtml = xhtml & " alert('Actualizado exitosamente');"
xhtml = xhtml & " document.all.initx.submit();"
xhtml = xhtml & "</script>"

elseif opc = "insertar" then

set rz = Cnn.Execute("sp_cg_Responsables @Modo='listar_ids_activos_sin', @ExcluirId=" & SqlStr(session("usuario")))

copias = ""
sep = ""

do while not rz.eof

	num = rz(0)

	if request.form("u" & num) = 1 then
		copias = copias & sep & rz(0)
		sep = ","
	end if

rz.movenext
loop

sDiaAviso = clng(CDATE(request.form("datepicker"))) - CLng(CDate("01-01-1900"))
sFecHoy = clng(date) - CLng(CDate("01-01-1900"))


ExecNonQuery Cnn, "sp_cg_Agenda @Modo='insertar', @Asunto='" & request.form("asun") & "', @DiaAviso=" & sDiaAviso & ", @IdRespon=" & SqlStr(session("usuario")) & ", @Copias='" & copias & "', @FecCarga=" & sFecHoy

XXX = SYSLOGON(cnn,"AGREGA_REC")

xhtml = xhtml & "<FORM name='init' method='post'  ACTION='agenda.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "' >"
xhtml = xhtml & "<INPUT TYPE='SUBMIT' VALUE='INICIAR' style='display:none;' >"
xhtml = xhtml & "</FORM>"

xhtml = xhtml & "<script type='text/javascript'>"
xhtml = xhtml & " alert('Agregado exitosamente');"
xhtml = xhtml & " document.all.init.submit();"
xhtml = xhtml & "</script>"

elseif opc = "det" then

id = request.querystring("id")
idresp = request.querystring("idresp")
idagre = request.querystring("idagre")
ccs = request.querystring("ccs")
xopt = request.querystring("xopt")

set rp = Cnn.Execute("sp_cg_Agenda @Modo='obtener', @Id=" & id)

puede = 1

if not rp.eof then
	if rp(0) = "SI" then puede = 0
end if

if xopt = "agre" then

if ccs & "" <> "" then

ExecNonQuery Cnn, "sp_cg_Agenda @Modo='actualizar_copias', @Id=" & id & ", @Copias='" & ccs & "," & idagre & "'"

ccs = ccs & "," & idagre

else

ExecNonQuery Cnn, "sp_cg_Agenda @Modo='actualizar_copias', @Id=" & id & ", @Copias='" & idagre & "'"

ccs = idagre

end if

end if

xhtml = xhtml & "<div class='inner ui-main' style='padding: 2rem;'>"

xhtml = xhtml & "<div class='bloque-datos ui-panel' style='max-width: 480px; margin: 0 auto;'>"
xhtml = xhtml & "<div class='bloque-header ui-panel__header' style='text-align: center;'>RECORDATORIO " & id & "</div>"
xhtml = xhtml & "<div class='bloque-body ui-panel__body'>"

if ccs & "" <> "" then

set rf = Cnn.Execute("sp_cg_Responsables @Modo='nombre_id_in', @Ids='" & ccs & "'")

xfrm = ""

xhtml = xhtml & "<div style='display: flex; flex-direction: column; gap: 0.5rem;'>"

do while not rf.eof

if puede = 1 then
	xfrm = xfrm & "<form id='elim" & DomId(rf(1)) & "' method='post' action='agenda.asp?opc=elim&idag=" & id & "&ccs=" & Server.URLEncode(ccs & "") & "&idresp=" & Server.URLEncode(idresp & "") & "&idelim=" & Server.URLEncode(rf(1) & "") & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "' ></form>"
end if

if getPermisos(xmenu,xtmenu,"ELIM") = 1 then
	xinput = "<input class='bot ui-button ui-button--icon ui-button--danger' type='button' value='X' onclick=""document.getElementById('elim" & DomId(rf(1)) & "').submit();"">"
else
	xinput = "<input class='botsp ui-button ui-button--icon' type='button' value='X' onclick=""alert('Sin permiso para eliminar.');"">"
end if

xhtml = xhtml & "<div style='display: flex; align-items: center; gap: 0.5rem;'>"

if cstr(idresp) = cstr(session("USUARIO")) then
	xhtml = xhtml & xinput
end if
xhtml = xhtml & "<span>-" & rf(0) & "</span>"

xhtml = xhtml & "</div>"

rf.movenext
loop

xhtml = xhtml & "</div>"

else

xhtml = xhtml & "<div class='dato-value'>NO REGISTRA COPIAS</div>"

end if

xhtml = xhtml & "</div>" ' Fin bloque-body
xhtml = xhtml & "</div>" ' Fin bloque-datos (recordatorio)

if ccs & "" <> "" then

set rm = Cnn.Execute("sp_cg_Responsables @Modo='listar_not_in', @Ids='" & ccs & "', @ExcluirId=" & SqlStr(idresp))

else

set rm = Cnn.Execute("sp_cg_Responsables @Modo='id_nombre_sin', @ExcluirId=" & SqlStr(idresp))

end if

xhtml = xhtml & "<div class='bloque-datos ui-panel' style='max-width: 480px; margin: 1.5rem auto 0;'>"
xhtml = xhtml & "<div class='bloque-header ui-panel__header' style='text-align: center;'>AGREGAR USUARIO PARA ENVIAR COPIA</div>"
xhtml = xhtml & "<div class='bloque-body ui-panel__body'>"

if not rm.eof then

xhtml = xhtml & "<div style='display: flex; flex-direction: column; gap: 0.5rem;'>"

do while not rm.eof

if puede = 1 then

	if cstr(idresp) = cstr(session("USUARIO")) then
		ref = "<a class='red' href='agenda.asp?opc=det&xopt=agre&id=" & id & "&ccs=" & Server.URLEncode(ccs & "") & "&idresp=" & Server.URLEncode(idresp & "") & "&idagre=" & Server.URLEncode(rm(0) & "") & "'>" & rm(1) & "</a>"
	else
		ref = rm(1)
	end if

else
	ref = rm(1)
end if

xhtml = xhtml & "<div>" & ref & "</div>"

rm.movenext
loop

xhtml = xhtml & "</div>"

else

xhtml = xhtml & "<div class='dato-value'>SIN USUARIOS DISPONIBLES</div>"

end if

xhtml = xhtml & "</div>" ' Fin bloque-body
xhtml = xhtml & "</div>" ' Fin bloque-datos (agregar usuario)

if puede = 1 then

xhtml = xhtml & "<form id='act' method='post' action='agenda.asp?opc=modif&id=" & id & "&ccs=" & Server.URLEncode(ccs & "") & "&idresp=" & Server.URLEncode(idresp & "") & "'>"

xhtml = xhtml & "<div class='bloque-datos ui-panel' style='max-width: 480px; margin: 1.5rem auto 0;'>"
xhtml = xhtml & "<div class='bloque-header ui-panel__header' style='text-align: center;'>MODIFICAR ASUNTO</div>"
xhtml = xhtml & "<div class='bloque-body ui-panel__body'>"

if cstr(idresp) = cstr(session("USUARIO")) then
	xhtml = xhtml & "<input id='asun' name='asun' class='inpu ui-input' style='width:100%;' type='text' value='" & rp(1) & "' onblur=""document.getElementById('act').submit()"">"
else
	xhtml = xhtml & "<div class='dato-value'>" & rp(1) & "</div>"
end if

xhtml = xhtml & "</div>" ' Fin bloque-body
xhtml = xhtml & "</div>" ' Fin bloque-datos (modificar asunto)

xhtml = xhtml & "</form>"

end if

xhtml = xhtml & "</div>" ' Fin inner

elseif opc = "elim" then

idresp = request.querystring("idresp")
idelim = request.querystring("idelim")
idag = request.querystring("idag")

set rt = Cnn.Execute("sp_cg_Agenda @Modo='obtener_copias', @Id=" & idag)

if not rt.eof then

xcopy = split(rt(0),",")
xcopyx = ""
sep = ""

for i=0 to ubound(xcopy)

	if xcopy(i) <> idelim then
		xcopyx = xcopyx & sep & xcopy(i)
		sep = ","
	end if

next

rt.close

ExecNonQuery Cnn, "sp_cg_Agenda @Modo='actualizar_copias', @Id=" & idag & ", @Copias='" & xcopyx & "'"

set rt = Cnn.Execute("sp_cg_Agenda @Modo='obtener_copias', @Id=" & idag)

ccs = ""

if not rt.eof then ccs = rt(0)

xhtml = xhtml & "<FORM name='init' method='post'  ACTION='agenda.asp?opc=det&id=" & idag & "&ccs=" & Server.URLEncode(ccs & "") & "&idresp=" & Server.URLEncode(idresp & "") & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "' >"
xhtml = xhtml & "<INPUT TYPE='SUBMIT' VALUE='INICIAR' style='display:none;' >"
xhtml = xhtml & "</FORM>"

xhtml = xhtml & "<script type='text/javascript'>"
xhtml = xhtml & " document.all.init.submit();"
xhtml = xhtml & "</script>"

end if

else 'listado de recordatorios

	' --- INICIO NORMALIZACION: AGENDA (RECORDATORIOS) ---
	XXX = SYSLOGON(cnn,"MENU_AGENDA")

	' Contenedor con padding para que la tabla no pegue a los bordes
	xhtml = xhtml & "<div class='inner ui-main' style='padding: 2rem;'>"

	' Tarjeta contenedora con scroll horizontal si es necesario
	xhtml = xhtml & "<div class='bloque-datos table-responsive ui-table-wrap'>"

	'xhtml = xhtml & "<table style='width:62%;color:#000000;font-family:Serif;text-align:center;' cellspacing=0 align=center >"
	xhtml = xhtml & "<table class='tabla-moderna ui-table'>"

	' Encabezados unificados en thead
	xhtml = xhtml & "<thead>"
	xhtml = xhtml & "<tr>"
	xhtml = xhtml & "<th style='width: 50px;'>&nbsp;</th>" ' Columna de la X
	xhtml = xhtml & "<th>ID</th>"
	xhtml = xhtml & "<th>RESPONSABLE</th>"
	xhtml = xhtml & "<th>ASUNTO</th>"
	xhtml = xhtml & "<th>DIA</th>"
	xhtml = xhtml & "<th style='text-align: center;'>AVISADO</th>"
	xhtml = xhtml & "</tr>"
	xhtml = xhtml & "</thead>"
	
	xhtml = xhtml & "<tbody>"

set rs = Cnn.Execute("sp_cg_Agenda @Modo='listar'")

formx1 = ""

if not rs.eof then rs.Move (CantRegMover)

PosReg = 0
sColor = "#FFFFFF"

Do While (Not rs.eof) And (PosReg < CantRegMostrar)

formx1 = formx1 & "<form id='formx1" & rs(0) & "' style='height:1px;' method='post' action='eliminar.asp?opc=rec&id=" & rs(0) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=agenda.asp'></form>"
if getPermisos(xmenu,xtmenu,"ELIM") = 1 then
	'botonx1 = "<input class='bot' type='button' value='X' onclick=""ejecutar('formx1" & rs(0) & "','Borrar','borra');"">"
	botonx1 = "<input type='button' class='btn-accion btn-accion-peligro ui-button ui-button--icon ui-button--danger' value='X' onclick=""ejecutar('formx1" & rs(0) & "','Borrar','borra');"">"
else
	'botonx1 = "<input class='botsp' type='button' value='X' onclick=""alert('Sin permiso para eliminar.');"">"
	botonx1 = "<input type='button' class='btn-accion btn-accion-deshabilitado ui-button ui-button--icon' value='X' onclick=""alert('Sin permiso para eliminar.');"">"
end if

resp = getName(Cnn,"","cor_responsables",rs(3))

if rs(4) = "SI" then
	'aviso = "<font style='color:#088A08;'>SI</font>"
	aviso = "<span class='badge badge-ok ui-badge ui-badge--ok'>SI</span>"
elseif rs(4) = "NO" then
	'aviso = "<font style='color:#FF0000;'>NO</font>"
	aviso = "<span class='badge badge-peligro ui-badge ui-badge--danger'>NO</span>"
else
    aviso = "S/D"
end if

SFECHA = ""
IF RS(2) <> 0 THEN
	SFECHA = CDATE(rs(2)) + CLng(CDate("01-01-1900"))
END IF

UrlDetalle = "agenda.asp?opc=det&id=" & rs(0) & "&ccs=" & Server.URLEncode(rs(6) & "") & "&idresp=" & Server.URLEncode(rs(3) & "") & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu

if getPermisos(xmenu,xtmenu,"ELIM") = 1 then
	JsClickFila = "onclick=""document.location='" & UrlDetalle & "'"""
else
	JsClickFila = "onclick=""alert('Sin permiso para modificar.');"""
end if

xhtml = xhtml & "<tr>"

xhtml = xhtml & "<td>" & botonx1 & "</td>"
'xhtml = xhtml & "<td class='fil' style='background-color:" & sColor & ";'>" & rs(0) & "</td>"
xhtml = xhtml & "<td style='cursor:pointer;' " & JsClickFila & ">" & rs(0) & "</td>"
xhtml = xhtml & "<td style='cursor:pointer;' " & JsClickFila & ">" & resp & "</td>"
xhtml = xhtml & "<td style='cursor:pointer;' " & JsClickFila & " title='" & rs(1) & "'>" & left(rs(1),30) & ".." & "</td>"
xhtml = xhtml & "<td style='cursor:pointer;' " & JsClickFila & ">" & sFECHA & "</td>"
xhtml = xhtml & "<td style='text-align:center;cursor:pointer;' " & JsClickFila & ">" & aviso & "</td>"

xhtml = xhtml & "</tr>"

'if sColor = "#FFFFFF" then
	'sColor = "#ECECEC"
'elseif sColor = "#ECECEC" then
	'sColor = "#FFFFFF"
'end if

PosReg = PosReg + 1

rs.movenext
loop

xhtml = xhtml & "</table>"

xhtml = xhtml & "</div>"

' --- ARMADO DE LA BOTONERA DE AGENDA ---

'shtml =  "<div id=btn style='top:85px;left:750px;position:absolute;'><table align=right><tr>"
misBotones = ""

if cantregmover <> 0 then
	'formx22 = "<form id='formx22' style='height:1px;' method='post' action='agenda.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&cantregmover=" & (cantregmover - cantregmostrar) & "'></form>"
	formx22 = "<form id='formx22' class='form-oculto' method='post' action='agenda.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&cantregmover=" & (cantregmover - cantregmostrar) & "'></form>"
	
	'botonx22 = "<input class='bot' type='button' value='ATRAS' onclick=""document.getElementById('formx22').submit();"">"
	botonx22 = "<input class='bot ui-button' type='button' value='ATRAS' onclick=""document.getElementById('formx22').submit();"">"
else
	'botonx22 = "<input class='botsp' type='button' value='ATRAS'>"
	formx22 = ""
	botonx22 = "<input class='botsp ui-button' type='button' value='ATRAS' disabled>"
end if

if not rs.eof then
	'formx23 = "<form id='formx23' style='height:1px;' method='post' action='agenda.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&cantregmover=" & (cantregmover + cantregmostrar) & "'></form>"
	formx23 = "<form id='formx23' class='form-oculto' method='post' action='agenda.asp?xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&cantregmover=" & (cantregmover + cantregmostrar) & "'></form>"

	'botonx23 = "<input class='bot' type='button' value='SIGUIENTE' onclick=""document.getElementById('formx23').submit();"">"
	botonx23 = "<input class='bot ui-button' type='button' value='SIGUIENTE' onclick=""document.getElementById('formx23').submit();"">"
else
	'botonx23 = "<input class='botsp' type='button' value='SIGUIENTE'>"
	formx23 = ""
	botonx23 = "<input class='botsp ui-button' type='button' value='SIGUIENTE' disabled>"
end if

'shtml = shtml & "<td>" & botonx22 & "</td>"
'shtml = shtml & "<td>" & botonx23 & "</td>"
'shtml = shtml & "</tr></table></div>"

' Agrupamos los formularios y botones (con flexbox para alinear)
misBotones = "<div class='ui-pagination' style='display: flex; gap: 0.5rem;'>" & formx22 & botonx22 & formx23 & botonx23 & "</div>"

xhtml = Replace(xhtml, "[BOTONERA_PAGINACION]", misBotones)

end if

xhtml = xhtml & "<script type='text/javascript'>"
xhtml = xhtml & " function ejecutar(formu,dsc,acc){ "
xhtml = xhtml & " var r = confirm(dsc); "
xhtml = xhtml & " if (r){ "
xhtml = xhtml & " if (acc == 'ins') {document.getElementById(formu).action=document.getElementById(formu).action + '&opc2=insertar';}"
xhtml = xhtml & " document.getElementById(formu).submit(); "
xhtml = xhtml & " }else{ "
xhtml = xhtml & " return false;"
xhtml = xhtml & " }} "
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
xhtml = xhtml & "dayNames: ['Domingo', 'Lunes', 'Martes', 'Miercoles', 'Jueves', 'Viernes', 'Sabado'],"
xhtml = xhtml & "dayNamesShort: ['Dom','Lun','Mar','Mi?','Juv','Vie','S?b'],"
xhtml = xhtml & "dayNamesMin: ['Do','Lu','Ma','Mi','Ju','Vi','Sa'],"
xhtml = xhtml & "weekHeader: 'Sm',"
xhtml = xhtml & "dateFormat: 'dd/mm/yy',"
xhtml = xhtml & "firstDay: 1,"
xhtml = xhtml & "isRTL: false,"
xhtml = xhtml & "showMonthAfterYear: false,"
xhtml = xhtml & "yearSuffix: ''"
xhtml = xhtml & "};"
xhtml = xhtml & "$.datepicker.setDefaults($.datepicker.regional['es']);"
xhtml = xhtml & "});"

xhtml = xhtml & "$( function() { "
xhtml = xhtml & "$( '#datepicker' ).datepicker(); "
xhtml = xhtml & "} ); "
xhtml = xhtml & "</script>"




xhtml = xhtml & xform & xforma & xformu & formx1 & formx2 & xfrm & xagregar & shtml & formx22 & formx23 & xformb & "</body></html>"

response.write xhtml

Function armaCombo(sql,cn,clave,valor,texto,comboid)

	Set nrs = Exec(cn, sql)

	if comboid = "resp" then
		html = html & "<select style='font-size:12pt;height:80%;color:#000000;border:solid 1px #000000;' id='" & comboid & "' name='" & comboid & "' onchange=""document.getElementById('xformu').submit();"" disabled>"
	else
		html = html & "<select style='font-size:12pt;height:80%;color:#000000;border:solid 1px #000000;' id='" & comboid & "' name='" & comboid & "' onchange=""document.getElementById('xformu').submit();"">"
	end if

	if comboid = "resp" then

	Set rf = Exec(cn, "sp_cg_Responsables @Modo='id_nombre', @IdSeg=" & SqlStr(session("usuario")))

		if not rf.eof then html = html & "<option value='" & rf(0) & "'>" & rf(1) & "</option>"

	else

		html = html & "<option value='0'>S / D</option>"

		Do While Not nrs.eof

			If cstr(clave) = cstr(nrs(valor)) Then

				html = html & "<option value='" & nrs(valor) & "' selected>" & nrs(texto) & "</option>"

			Else

				html = html & "<option value='" & nrs(valor) & "'>" & nrs(texto) & "</option>"

			End If

			nrs.movenext
		Loop

	end if

	armaCombo = html & "</select>"

End Function

%>