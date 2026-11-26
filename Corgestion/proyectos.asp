<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/ui_componentes.inc" -->
<!-- #INCLUDE FILE="IncFile/datosDetalles.inc" --> 
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

' [UI] Desactivado: declaraba UTF-8 pero la pagina se envia en ANSI (como todo el sistema),
' [UI] y los textos con tilde que vienen de la base se veian rotos solo en esta pantalla.
'[UI-ANT] Response.CharSet = "utf-8" ' Buena practica

' --- INICIALIZACION DE CONEXION (Requerido por backend) ---
Set Cnn = Server.CreateObject("ADODB.Connection")
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

' --- CAPTURA DE PARAMETROS (Logica de Control) ---
xmenu = request.querystring("xmenu")
xtmenu = request.querystring("xtmenu")
opc = request.querystring("opc")
IsConsulta = (opc = "consi") ' Booleano para saber que vista renderizar

' Paginacion
CantRegMostrar = 12
If IsConsulta Then CantRegMostrar = 10   ' [CONSULTA] las consultas paginan de a 10 (como incidentes y APL)
CantRegMover = Request.querystring("CantRegMover")
if not IsNumeric(CantRegMover) or CantRegMover < 0 then CantRegMover = 0

' Filtros Comunes
cliente = request.form("cli")
if cliente & "" = "" then cliente = request.querystring("cli")
if not IsNumeric(cliente) then cliente = 0

aplicativos = request.form("apli")
if aplicativos & "" = "" then aplicativos = request.querystring("apli")
if not IsNumeric(aplicativos) then aplicativos = 0

tit = request.form("tit") '-- En consulta se usa 'desc', unificamos a 'tit' en el SP
if tit & "" = "" then tit = request.querystring("tit")
if tit & "" = "" then tit = request.form("desc") '-- compatibilidad con form consulta

ord = request.querystring("ord")

' ' Filtros especificos de Consulta
estadoFiltro = request.form("est")
if not IsNumeric(estadoFiltro) then estadoFiltro = 0

fechaFiltro = request.form("fec")' -- Asumimos formato fecha dd/mm/aaaa

' ' --- SEGURIDAD Y LOGS (Mismo comportamiento original) ---
If IsConsulta Then
    XXX = SYSLOGON(cnn,"CONSULTA_PROYECTOS")
Else
    XXX = SYSLOGON(cnn,"MENU_PROYECTOS")
End If
usuario = userlogon()

' --- CARGA DE DATOS (Llamadas al Backend) ---
Dim rsClientes, rsApli, rsGrid

' 1. Cargar combos (siempre necesarios)
'Set rsClientes = GetListaCombo(Cnn,"CLIENTES")
' Set rsApli = GetListaCombo(Cnn,"APLICATIVOS")
Set rsClientes = ObtenerClientes(Cnn)
Set rsApli = ObtenerAplicativos(Cnn)

' ' 2. Cargar Grilla principal llamando al SP mediante el backend
' ' Pasamos los parametros capturados. El SP unifica la logica de filtros de ambas vistas.
' [CG] Filtro por responsable: por defecto los registros del usuario de la sesion; "*" = todos. En consulta (consi) no se filtra.
If IsConsulta Then cgResp = "*" Else cgResp = CgRespSeleccionado()
 Set rsGrid = GetProyectosListado(scon1 , cliente, aplicativos, tit, estadoFiltro, fechaFiltro, ord, CantRegMover, CantRegMostrar, CgRespParam(cgResp))

' ' Determinar si hay registros y si hay pagina siguiente (usando COUNT(*) OVER() del SP)
Dim TotalRegistrosTotal, HayRegistros, TieneSiguiente
HayRegistros = Not rsGrid.EOF
If HayRegistros Then
    TotalRegistrosTotal = rsGrid("TotalRegistros")
    ' Es un Recordset desconectado de solo CantRegMostrar, no podemos usar rsGrid.EOF para 'Siguiente'
    ' Calculamos matematicamente
    TieneSiguiente = (CDbl(CantRegMover) + CantRegMostrar) < TotalRegistrosTotal
Else
    TotalRegistrosTotal = 0
    TieneSiguiente = False
End If


'[UI-ANT] ' ===== Renderizado anterior completo: reemplazado por el bloque [UI] que sigue. Se borra entero buscando [UI-ANT] =====
'[UI-ANT] ' ==============================================================================
'[UI-ANT] ' INICIO DE RENDERIZADO (HTML)
'[UI-ANT] ' ==============================================================================
'[UI-ANT] html = headhtml()
'[UI-ANT] html = html & "<body>"
'[UI-ANT] html = html & "<div>"

'[UI-ANT] ' --- ENCABEZADO (Comun) ---
'[UI-ANT] html = html & cabecera()

'[UI-ANT] ' --- TITULO SECCION Y BOTONERA SUPERIOR ---
'[UI-ANT] html = html & "<table class='titulo-seccion ui-page-header'>"
'[UI-ANT] html = html & "<tr style='display:flex;align-items:center;justify-content:space-between;'>"

'[UI-ANT] If IsConsulta Then
    '[UI-ANT] html = html & "<td><h2>CONSULTA PROYECTOS</h2></td>"
    '[UI-ANT] ' Boton Volver especifico de consulta
    '[UI-ANT] formx10 = "<form id='formx10' class='form-oculto' method='post' action='consultas.asp?sel=" & request.querystring("sel") & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
    '[UI-ANT] botonVolver = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('formx10').submit();"">"
    '[UI-ANT] html = html & "<td>" & formx10 & botonVolver & "</td>"
'[UI-ANT] Else
    '[UI-ANT] html = html & "<td><h2>PROYECTOS</h2></td>"

    '[UI-ANT] ' Form Filtro Cliente
    '[UI-ANT] html = html & "<td><form id='xForm' method='post' action='proyectos.asp?ord=" & ord & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&apli=" & aplicativos & "'>"
    '[UI-ANT] html = html & "<select class='ui-select' onchange=""document.getElementById('xForm').submit();"" id='cli' name='cli'>"
    '[UI-ANT] html = html & "<option value='0' >CLIENTE: TODOS</option>"
    '[UI-ANT] Do While Not rsClientes.EOF
        '[UI-ANT] selected = ""
        '[UI-ANT] if cstr(rsClientes(0)) = cstr(cliente) then selected = "selected"
        '[UI-ANT] html = html & "<option value='" & rsClientes(0) & "' " & selected & " >" & Server.HTMLEncode(rsClientes(1) & "") & "</option>"
        '[UI-ANT] rsClientes.movenext
    '[UI-ANT] Loop
    '[UI-ANT] html = html & "</select></form></td>"

    '[UI-ANT] ' Form Filtro Aplicativo
    '[UI-ANT] html = html & "<td><form id='xFormApli' method='post' action='proyectos.asp?ord=" & ord & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&cli=" & cliente & "'>"
    '[UI-ANT] html = html & "<select class='ui-select' onchange=""document.getElementById('xFormApli').submit();"" id='apli' name='apli'>"
    '[UI-ANT] html = html & "<option value='0' >APLICATIVO: TODOS</option>"
    '[UI-ANT] Do While Not rsApli.EOF
        '[UI-ANT] selected = ""
        '[UI-ANT] If cstr(rsApli(0)) = cstr(aplicativos) Then selected = "selected"
        '[UI-ANT] html = html & "<option value='" & rsApli(0) & "' " & selected & " >" & Server.HTMLEncode(rsApli(1) & "") & "</option>"
        '[UI-ANT] rsApli.MoveNext
    '[UI-ANT] Loop
    '[UI-ANT] html = html & "</select></form></td>"

    '[UI-ANT] ' Boton Nuevo (Manejo de Permisos original)
    '[UI-ANT] if session("FUNC") = "FUNC" or session("FUNC") = "TEC"  then
        '[UI-ANT] formx19 = "<form id='formx1' class='form-oculto' method='post' action='agregar2.asp?tipo=PROY&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
        '[UI-ANT] if CStr(getPermisos(xmenu, xtmenu, "ALTA")) = "1" then
            '[UI-ANT] botonx1 = "<input class='bot ui-button ui-button--primary' type='button' value='NUEVO PROYECTO' onclick=""document.getElementById('formx1').submit();"">"
        '[UI-ANT] else
            '[UI-ANT] botonx1 = "<input class='botsp ui-button' type='button' value='CARGAR NUEVO PROYECTO' disabled title='Sin permiso para agregar.'>"
        '[UI-ANT] end if
        '[UI-ANT] html = html & "<td>" & formx19 & botonx1 & "</td>"
    '[UI-ANT] else
        '[UI-ANT] html = html & "<td>&nbsp;</td>"
    '[UI-ANT] end if

    '[UI-ANT] ' Placeholder Paginacion
    '[UI-ANT] html = html & "<td>[BOTONERA_PAGINACION]</td>"

    '[UI-ANT] ' Boton Volver estandar
    '[UI-ANT] formx29 = "<form id='formx2' class='form-oculto' method='post' action='index.asp'></form>"
    '[UI-ANT] botonx2 = "<input class='bot ui-button' type='button' value='VOLVER' onclick=""document.getElementById('formx2').submit();"">"
    '[UI-ANT] html = html & "<td>" & formx29 & botonx2 & "</td>"
'[UI-ANT] End If

'[UI-ANT] html = html & "</tr>"
'[UI-ANT] html = html & "</table>"
'[UI-ANT] html = html & "<div style='visibility:hidden;height:5px;'>&nbsp;</div>"


'[UI-ANT] ' ==============================================================================
'[UI-ANT] ' GRILLA DE DATOS (UNIFICADA)
'[UI-ANT] ' ==============================================================================
'[UI-ANT] html = html & "<div class='inner ui-main' style='padding:2rem'>"
'[UI-ANT] html = html & "<div class='bloque-datos table-responsive ui-table-wrap'>"
'[UI-ANT] html = html & "<table class='tabla-moderna ui-table'>"

'[UI-ANT] ' --- ENCABEZADOS ---
'[UI-ANT] html = html & "<thead><tr>"

'[UI-ANT] ' Columnas de accion (solo vista normal)
'[UI-ANT] If Not IsConsulta Then
    '[UI-ANT] html = html & "<th></th>" ' Eliminar
    '[UI-ANT] html = html & "<th></th>" ' Actualizar
'[UI-ANT] End If
'[UI-ANT] html = html & "<th></th>"     ' Detalle

'[UI-ANT] ' Columnas de datos con ordenamiento
'[UI-ANT] UrlBaseParams = "&cli=" & cliente & "&apli=" & aplicativos & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu
'[UI-ANT] if IsConsulta then UrlBaseParams = UrlBaseParams & "&opc=consi"

'[UI-ANT] html = html & "<th>ID</th>"
'[UI-ANT] html = html & "<th>SOLICITUD</th>"
'[UI-ANT] html = html & "<th><a class='col-link' href='proyectos.asp?ord=titulo" & UrlBaseParams & "'>TITULO</a></th>"
'[UI-ANT] html = html & "<th><a class='col-link' href='proyectos.asp?ord=cliente" & UrlBaseParams & "'>CLIENTE</a></th>"

'[UI-ANT] ' Columnas extra especificas
'[UI-ANT] If Not IsConsulta Then
    '[UI-ANT] html = html & "<th>PRIORIDAD</th>"
'[UI-ANT] End If

'[UI-ANT] html = html & "<th>DESARROLLO</th>"
'[UI-ANT] html = html & "<th>FUNCIONAL</th>"
'[UI-ANT] html = html & "<th>FEC INGRESO</th>"
'[UI-ANT] html = html & "<th><a class='col-link' href='proyectos.asp?ord=aplicativos" & UrlBaseParams & "'>APLICATIVO</a></th>"

'[UI-ANT] html = html & "</tr></thead>"
'[UI-ANT] html = html & "<tbody>"

'[UI-ANT] ' --- DETALLE (BUCLE) ---
'[UI-ANT] formsAccion = "" ' Acumulador de formularios ocultos

'[UI-ANT] If Not HayRegistros Then
    '[UI-ANT] cols = 10
    '[UI-ANT] if IsConsulta then cols = 7
    '[UI-ANT] html = html & "<tr><td colspan='" & cols & "' align='center'>No se encontraron proyectos con los filtros seleccionados.</td></tr>"
'[UI-ANT] Else
    '[UI-ANT] ' Reutilizamos Recordset rsGrid cargado al principio
    '[UI-ANT] Do While Not rsGrid.EOF
        
        '[UI-ANT] ' 1. Preparar Badges y Datos (usando funciones del backend)
        '[UI-ANT] prioBadge = ""
        '[UI-ANT] If Not IsConsulta Then prioBadge = GetBadgePrioridad(rsGrid("prioridad_desc"))
        
        '[UI-ANT] estDesaBadge = GetBadgeEstado(rsGrid("estado_desa_desc"), rsGrid("respon_tec_nombre"))
        '[UI-ANT] estFuncBadge = GetBadgeEstado(rsGrid("estado_func_desc"), rsGrid("respon_gestion_nombre"))
        
        '[UI-ANT] ' Fecha (asumimos que el SP devuelve DateTime, formateamos en vista si es necesario)
        '[UI-ANT] sFec = rsGrid("fecha_ingreso") -- CDate(rsGrid("fecha_ingreso"))

	'[UI-ANT] 'xformx = xformx & "<form id='xformx" & rs(0) & "' style='margin: 0; padding: 0;' method='post' action='detalles.asp?id=" & rs(0) & "&opc=verp&tipo=cons&" & request.form & "'></form>"
	
        '[UI-ANT] ' URL Detalle comun
        '[UI-ANT] UrlDetalle = "detalles.asp?id=" & rsGrid("id") & "&sol=" & rsGrid("solic_nro") & _
                     '[UI-ANT] "&cli=" & cliente & "&opc=verp" & "&tit=" & server.URLEncode(tit)  & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu
        '[UI-ANT] If IsConsulta Then UrlDetalle = UrlDetalle & "&opc=verp&tipo=cons&" & request.form

        '[UI-ANT] ' Evento click en fila
        '[UI-ANT] JsClickFila = "onclick=""document.location='" & UrlDetalle & "'"" style='cursor:pointer;'"

        '[UI-ANT] html = html & "<tr>"
        
        '[UI-ANT] ' 2. Renderizado de Columnas de Accion (Vista Normal)
        '[UI-ANT] If Not IsConsulta Then
            '[UI-ANT] ' Boton Eliminar
            '[UI-ANT] formsAccion = formsAccion & "<form id='formElim" & rsGrid("id") & "' class='form-oculto' method='post' action='eliminar.asp?opc=proy&id=" & rsGrid("id") & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=proyectos.asp'></form>"
            '[UI-ANT] if CStr(getPermisos(xmenu,xtmenu,"ELIM")) = "1" then
                '[UI-ANT] botonx1 = "<input class='btn-accion btn-accion-peligro ui-button ui-button--icon ui-button--danger' type='button' value='X' onclick=""borrar('formElim" & rsGrid("id") & "');"">"
            '[UI-ANT] else
                '[UI-ANT] botonx1 = "<input class='btn-accion btn-accion-deshabilitado ui-button ui-button--icon' type='button' value='X' disabled title='Sin permiso para eliminar.'>"
            '[UI-ANT] end if
            '[UI-ANT] html = html & "<td style='text-align:center;'>" & botonx1 & "</td>"

            '[UI-ANT] ' Boton Actualizar
            '[UI-ANT] formsAccion = formsAccion & "<form id='formAct" & rsGrid("id") & "' class='form-oculto' method='post' action='detalles.asp?id=" & rsGrid("id") & "&sol=" & rsGrid("solic_nro") & "&opc=actp&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
            '[UI-ANT] ' Nota: Validacion USUARIO original comentada, se mantiene igual.
            '[UI-ANT] botonx2 = "<input class='btn-accion btn-accion-primario ui-button ui-button--icon ui-button--primary' type='button' value='A' onclick=""document.getElementById('formAct" & rsGrid("id") & "').submit();"">"
            '[UI-ANT] html = html & "<td style='text-align:center;'>" & botonx2 & "</td>"

            '[UI-ANT] ' Boton Detalle (D) - En vista normal es redundante con click en fila pero se mantiene
            '[UI-ANT] botonx3 = "<input class='btn-accionbtn-accion-secundario ui-button ui-button--icon' style='color:#008520;' type='button' value='D' onclick=""document.location='" & UrlDetalle & "' "">"
            '[UI-ANT] html = html & "<td style='text-align:center;'>" & botonx3 & "</td>"

        '[UI-ANT] Else
            '[UI-ANT] ' Boton Detalle (D) - Vista Consulta (con permisos fijos 10,INICIO original)
            '[UI-ANT] formsAccion = formsAccion & "<form id='formGridCons" & rsGrid("id") & "' class='form-oculto' method='post' action='" & UrlDetalle & "'></form>"
            '[UI-ANT] if getPermisos(10,"INICIO","CONS") = 1 then
                '[UI-ANT] xbotonx = "<input class='btn-accion ui-button ui-button--icon' style='background-color: #e0e7ff; color: #4338ca;' type='button' value='D' onclick=""document.getElementById('formGridCons" & rsGrid("id") & "').submit();"">"
            '[UI-ANT] else
                '[UI-ANT] xbotonx = "<input class='btn-accion btn-accion-deshabilitado ui-button ui-button--icon' type='button' value='D' disabled title='Sin permiso para consultar.'>"
            '[UI-ANT] end if
            '[UI-ANT] html = html & "<td style='text-align: center;'>" & xbotonx & "</td>"
        '[UI-ANT] End If


        '[UI-ANT] ' 3. Renderizado de Datos (Comun)
        '[UI-ANT] html = html & "<td " & JsClickFila & ">" & rsGrid("id") & "</td>"
        '[UI-ANT] html = html & "<td " & JsClickFila & " style='text-align:center;'>" & rsGrid("solic_nro") & "</td>"
        
        '[UI-ANT] ' Titulo (Trunco original)
        '[UI-ANT] DescTrunca = rsGrid("titulo") & ""
        '[UI-ANT] if Len(DescTrunca) > 30 then DescTrunca = left(UCASE(DescTrunca),30) & "..."
        '[UI-ANT] html = html & "<td " & JsClickFila & " style='text-align:left;'>" & Server.HTMLEncode(DescTrunca) & "</td>"

        '[UI-ANT] html = html & "<td " & JsClickFila & ">" & Server.HTMLEncode(rsGrid("cliente_desc") & "") & "</td>"

        '[UI-ANT] If Not IsConsulta Then
            '[UI-ANT] html = html & "<td " & JsClickFila & ">" & prioBadge & "</td>"
        '[UI-ANT] End If

        '[UI-ANT] html = html & "<td " & JsClickFila & ">" & estDesaBadge & "</td>"
        '[UI-ANT] html = html & "<td " & JsClickFila & ">" & estFuncBadge & "</td>"
        '[UI-ANT] html = html & "<td " & JsClickFila & ">" & sFec & "</td>"
        '[UI-ANT] html = html & "<td " & JsClickFila & ">" & Server.HTMLEncode(rsGrid("aplicativo_desc") & "") & "</td>"

        '[UI-ANT] html = html & "</tr>"

        '[UI-ANT] rsGrid.movenext
    '[UI-ANT] Loop
'[UI-ANT] End If

'[UI-ANT] html = html & "</tbody></table></div></div>" ' Cierre grilla


'[UI-ANT] ' ==============================================================================
'[UI-ANT] ' PAGINACION (Solo vista normal) Y SCRIPTS
'[UI-ANT] ' ==============================================================================
'[UI-ANT] If Not IsConsulta Then
    '[UI-ANT] misBotones = "<div class='ui-pagination' style='display: flex; gap: 0.5rem;'>"
    '[UI-ANT] UrlPagiBase = "proyectos.asp?ord=" & ord & "&cli=" & cliente & "&apli=" & aplicativos & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu

    '[UI-ANT] ' Boton ATRAS
    '[UI-ANT] if cantregmover <> 0 then
        '[UI-ANT] formsAccion = formsAccion & "<form id='formAtras' class='form-oculto' method='post' action='" & UrlPagiBase & "&cantregmover=" & (cantregmover - cantregmostrar) & "'></form>"
        '[UI-ANT] botonx22 = "<input class='bot ui-button' type='button' value='ATRAS' onclick=""document.getElementById('formAtras').submit();"">"
    '[UI-ANT] else
        '[UI-ANT] botonx22 = "<input class='botsp ui-button' type='button' value='ATRAS' disabled>"
    '[UI-ANT] end if

    '[UI-ANT] ' Boton SIGUIENTE
    '[UI-ANT] if TieneSiguiente then
        '[UI-ANT] formsAccion = formsAccion & "<form id='formSigue' class='form-oculto' method='post' action='" & UrlPagiBase & "&cantregmover=" & (cantregmover + cantregmostrar) & "'></form>"
        '[UI-ANT] botonx23 = "<input class='bot ui-button' type='button' value='SIGUIENTE' onclick=""document.getElementById('formSigue').submit();"">"
    '[UI-ANT] else
        '[UI-ANT] botonx23 = "<input class='botsp ui-button' type='button' value='SIGUIENTE' disabled>"
    '[UI-ANT] end if

    '[UI-ANT] misBotones = misBotones & botonx22 & botonx23 & "</div>"
    '[UI-ANT] ' Reemplazar placeholder
    '[UI-ANT] html = Replace(html, "[BOTONERA_PAGINACION]", misBotones)
'[UI-ANT] End If

'[UI-ANT] html = html & "</div>" ' Cierre div principal

'[UI-ANT] ' --- SCRIPTS JS ---
'[UI-ANT] html = html & "<script type='text/javascript'>"
'[UI-ANT] html = html & " function borrar(formu){ "
'[UI-ANT] html = html & " var r = confirm('Desea eliminar proyecto?'); "
'[UI-ANT] html = html & " if (r){ "
'[UI-ANT] html = html & " document.getElementById(formu).submit(); "
'[UI-ANT] html = html & " }else{ "
'[UI-ANT] html = html & " return false;"
'[UI-ANT] html = html & " }} "
'[UI-ANT] html = html & "</script>"

'[UI-ANT] html = html & "<script src='assets/js/jquery.min.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/skel.min.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/util.js'></script>"
'[UI-ANT] html = html & "<script src='assets/js/main.js'></script>"

'[UI-ANT] ' Inyeccion de formularios ocultos acumulados
'[UI-ANT] html = html & formsAccion
'[UI-ANT] html = html & "</body></html>"

' ==============================================================================
' [UI] RENDERIZADO con el sistema visual (IncFile/ui_componentes.inc)
' El bloque anterior completo queda arriba, comentado con '[UI-ANT].
' La carga de datos (arriba) no cambia. Formularios, URLs y permisos son los mismos.
' ==============================================================================
html = headhtml()
html = html & "<body>"
html = html & "<div>"

' --- ENCABEZADO (Comun) ---
html = html & cabecera()

' Parametros de ordenamiento (los mismos de antes)
UrlBaseParams = "&cli=" & cliente & "&apli=" & aplicativos & "&resp=" & Server.URLEncode(cgResp) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu
if IsConsulta then UrlBaseParams = UrlBaseParams & "&opc=consi&sel=" & Server.URLEncode(request.querystring("sel") & "")

formsAccion = "" ' Acumulador de formularios ocultos (van fuera de la tabla)

' Consulta: los filtros de estado, fecha y descripcion solo se leen del formulario. Para paginar y ordenar
' sin perderlos, se reenvian con este formulario oculto (irConsulta cambia el destino y lo envia).
Dim UrlConsulta
If IsConsulta Then
    UrlConsulta = "proyectos.asp?opc=consi&sel=" & Server.URLEncode(request.querystring("sel") & "") & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu
    formsAccion = formsAccion & "<form id='formFiltros' class='form-oculto' method='post' action='" & UrlConsulta & "'>" & _
        "<input type='hidden' name='cli' value='" & Server.HTMLEncode(cliente & "") & "'>" & _
        "<input type='hidden' name='apli' value='" & Server.HTMLEncode(aplicativos & "") & "'>" & _
        "<input type='hidden' name='desc' value='" & Server.HTMLEncode(tit & "") & "'>" & _
        "<input type='hidden' name='est' value='" & Server.HTMLEncode(request.form("est") & "") & "'>" & _
        "<input type='hidden' name='fec' value='" & Server.HTMLEncode(request.form("fec") & "") & "'>" & _
        "</form>"
End If

' --- BARRA SUPERIOR Y FILTROS ---
Dim accTop, filtrosHtml, tituloPantalla, tituloTarjeta, selHtml
accTop = ""
filtrosHtml = ""
If IsConsulta Then
    tituloPantalla = "Consulta de proyectos"
    tituloTarjeta = "Resultados"
    formsAccion = formsAccion & "<form id='formx10' class='form-oculto' method='post' action='consultas.asp?sel=" & request.querystring("sel") & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
    accTop = UiBoton("Volver", "secondary", "document.getElementById('formx10').submit();", True, "")
Else
    tituloPantalla = "Proyectos"
    tituloTarjeta = "Listado de proyectos"

    ' Boton Cargar (manejo de permisos original)
    if session("FUNC") = "FUNC" or session("FUNC") = "TEC" then
        formsAccion = formsAccion & "<form id='formx1' class='form-oculto' method='post' action='agregar2.asp?tipo=PROY&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
        accTop = accTop & UiBoton("Cargar proyecto", "light", "document.getElementById('formx1').submit();", CStr(getPermisosAtrm("CG_PROYECTOS", "ALTA")) = "1", "Sin permiso para agregar.")
    end if
    formsAccion = formsAccion & "<form id='formx2' class='form-oculto' method='post' action='index.asp'></form>"
    accTop = accTop & UiBoton("Volver", "secondary", "document.getElementById('formx2').submit();", True, "")

    ' Filtros en el encabezado de la tarjeta. Dos formularios como antes: cada uno conserva el otro filtro.
    filtrosHtml = "<div class='ui-card__filtros'>"

    filtrosHtml = filtrosHtml & "<form id='xForm' method='post' action='proyectos.asp?ord=" & ord & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&apli=" & aplicativos & "&resp=" & Server.URLEncode(cgResp) & "'>"
    selHtml = "<select class='ui-select' onchange=""document.getElementById('xForm').submit();"" id='cli' name='cli'><option value='0'>Todos los clientes</option>"
    Do While Not rsClientes.EOF
        selected = ""
        if cstr(rsClientes(0)) = cstr(cliente) then selected = "selected"
        selHtml = selHtml & "<option value='" & rsClientes(0) & "' " & selected & ">" & Server.HTMLEncode(rsClientes(1) & "") & "</option>"
        rsClientes.movenext
    Loop
    filtrosHtml = filtrosHtml & UiFiltroCompacto("cli", "Filtrar por cliente", selHtml & "</select>") & "</form>"

    filtrosHtml = filtrosHtml & "<form id='xFormApli' method='post' action='proyectos.asp?ord=" & ord & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&cli=" & cliente & "&resp=" & Server.URLEncode(cgResp) & "'>"
    selHtml = "<select class='ui-select' onchange=""document.getElementById('xFormApli').submit();"" id='apli' name='apli'><option value='0'>Todos los aplicativos</option>"
    Do While Not rsApli.EOF
        selected = ""
        If cstr(rsApli(0)) = cstr(aplicativos) Then selected = "selected"
        selHtml = selHtml & "<option value='" & rsApli(0) & "' " & selected & ">" & Server.HTMLEncode(rsApli(1) & "") & "</option>"
        rsApli.MoveNext
    Loop
    filtrosHtml = filtrosHtml & UiFiltroCompacto("apli", "Filtrar por aplicativo", selHtml & "</select>") & "</form>"

    ' [CG] Filtro Responsable (por defecto: usuario de la sesion)
    filtrosHtml = filtrosHtml & "<form id='xFormResp' method='post' action='proyectos.asp?ord=" & ord & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&cli=" & cliente & "&apli=" & aplicativos & "'>"
    filtrosHtml = filtrosHtml & UiFiltroCompacto("resp", "Filtrar por responsable", CgComboResponsable(Cnn, cgResp, "xFormResp")) & "</form>"

    filtrosHtml = filtrosHtml & "</div>"
End If

html = html & UiPantallaAbre(tituloPantalla, accTop, "wide")
html = html & UiCardAbre(tituloTarjeta, filtrosHtml)

' --- TABLA (mismo orden de columnas que Incidentes) ---
html = html & UiTablaAbre("Listado de proyectos. Se puede ordenar por t&iacute;tulo, cliente o aplicativo.")
html = html & "<thead><tr>"
If Not IsConsulta Then html = html & UiTh("<span class='ui-sr-only'>Acciones</span>", "ui-table__col--actions")
html = html & UiTh("ID", "") & UiTh("Solicitud", "")
html = html & UiThOrden("T&iacute;tulo", "proyectos.asp?ord=titulo" & UrlBaseParams, ord = "titulo")
html = html & UiThOrden("Cliente", "proyectos.asp?ord=cliente" & UrlBaseParams, ord = "cliente")
If Not IsConsulta Then html = html & UiTh("Prioridad", "")
html = html & UiThOrden("Aplicativo", "proyectos.asp?ord=aplicativos" & UrlBaseParams, ord = "aplicativos")
html = html & UiTh("Desarrollo", "") & UiTh("Funcional", "") & UiTh("Ingreso", "")
If IsConsulta Then html = html & UiTh("<span class='ui-sr-only'>Acciones</span>", "ui-table__col--actions")
html = html & "</tr></thead><tbody>"
If IsConsulta Then html = Replace(html, "class='ui-table__sort' href=", "class='ui-table__sort' onclick='return irConsulta(this.href);' href=")

Dim cantFilas
cantFilas = 0
If Not HayRegistros Then
    cols = 10
    if IsConsulta then cols = 9
    html = html & UiFilaVacia(cols, "No se encontraron proyectos con los filtros seleccionados.")
Else
    Do While Not rsGrid.EOF
        ' URL Detalle comun (la misma de antes)
        UrlDetalle = "detalles.asp?id=" & rsGrid("id") & "&sol=" & rsGrid("solic_nro") & _
                     "&cli=" & cliente & "&opc=verp" & "&tit=" & server.URLEncode(tit)  & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu
        '[UI-ANT] If IsConsulta Then UrlDetalle = UrlDetalle & "&opc=verp&tipo=cons&" & request.form
        ' [UI] En la consulta, titulo, fila y Ver detalle abren la vista de consulta (solo lectura);
        ' los filtros viajan para que Volver regrese a estos resultados
        If IsConsulta Then UrlDetalle = "vistaconsulta.asp?tipo=P&id=" & rsGrid("id") & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&" & request.form

        ' Click en la celda: abre el detalle, como antes
        JsClickFila = " class='ui-table__cell--link' onclick=""document.location='" & UrlDetalle & "'"""

        ' Fecha legible (la linea anterior sumaba la fecha consigo misma: "--" no es comentario en VBScript)
        sFec = UiFecha(rsGrid("fecha_ingreso"))

        html = html & "<tr>"

        If Not IsConsulta Then
            formsAccion = formsAccion & "<form id='formElim" & rsGrid("id") & "' class='form-oculto' method='post' action='eliminar.asp?opc=proy&id=" & rsGrid("id") & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=proyectos.asp'></form>"
            formsAccion = formsAccion & "<form id='formAct" & rsGrid("id") & "' class='form-oculto' method='post' action='detalles.asp?id=" & rsGrid("id") & "&sol=" & rsGrid("solic_nro") & "&opc=actp&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "'></form>"
            ' El boton "D" (detalle) se quita: el titulo ya es el enlace al detalle, y la fila tambien lo abre
            html = html & "<td><div class='ui-table__actions'>" & _
                UiBotonIcono("&times;", "Eliminar proyecto " & rsGrid("id"), "danger", "borrar('formElim" & rsGrid("id") & "');", CStr(getPermisosAtrm("CG_PROYECTOS", "ELIM")) = "1", "Sin permiso para eliminar") & _
                UiBotonIcono("&#9998;", "Editar proyecto " & rsGrid("id"), "secondary", "document.getElementById('formAct" & rsGrid("id") & "').submit();", True, "") & _
                "</div></td>"
        End If

        html = html & "<td" & JsClickFila & ">" & rsGrid("id") & "</td>"
        html = html & "<td" & JsClickFila & "><span class='ui-table__nowrap'>" & UiDato(rsGrid("solic_nro")) & "</span></td>"
        html = html & "<td>" & UiCeldaEnlace(UiTexto(rsGrid("titulo")), UrlDetalle, UiTexto(rsGrid("titulo"))) & "</td>"
        html = html & "<td" & JsClickFila & ">" & UiTexto(rsGrid("cliente_desc")) & "</td>"
        If Not IsConsulta Then html = html & "<td" & JsClickFila & ">" & UiBadgePrioridad(rsGrid("prioridad_desc")) & "</td>"
        html = html & "<td" & JsClickFila & ">" & UiTexto(rsGrid("aplicativo_desc")) & "</td>"
        html = html & "<td" & JsClickFila & ">" & UiBadgeEstado(rsGrid("estado_desa_desc")) & UiSub(UiTexto(rsGrid("respon_tec_nombre"))) & "</td>"
        html = html & "<td" & JsClickFila & ">" & UiBadgeEstado(rsGrid("estado_func_desc")) & UiSub(UiTexto(rsGrid("respon_gestion_nombre"))) & "</td>"
        html = html & "<td" & JsClickFila & "><span class='ui-table__nowrap'>" & sFec & "</span></td>"

        If IsConsulta Then
            ' Detalle en consulta (con permisos fijos 10, INICIO, como antes)
            formsAccion = formsAccion & "<form id='formGridCons" & rsGrid("id") & "' class='form-oculto' method='post' action='" & UrlDetalle & "'></form>"
            html = html & "<td>" & UiBoton("Ver detalle", "secondary", "document.getElementById('formGridCons" & rsGrid("id") & "').submit();", getPermisosAtrm("CG_PROYECTOS", "CONS") = 1, "Sin permiso para consultar.") & "</td>"
        End If

        html = html & "</tr>"
        cantFilas = cantFilas + 1
        rsGrid.movenext
    Loop
End If

html = html & UiTablaCierra()

' --- PIE: rango mostrado y paginacion ---
Dim botonAnt, botonSig
If Not IsConsulta Then
    UrlPagiBase = "proyectos.asp?ord=" & ord & "&cli=" & cliente & "&apli=" & aplicativos & "&resp=" & Server.URLEncode(cgResp) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu

    if cantregmover <> 0 then
        formsAccion = formsAccion & "<form id='formAtras' class='form-oculto' method='post' action='" & UrlPagiBase & "&cantregmover=" & (cantregmover - cantregmostrar) & "'></form>"
        botonAnt = UiBoton("Anterior", "secondary", "document.getElementById('formAtras').submit();", True, "")
    else
        botonAnt = UiBoton("Anterior", "secondary", "", False, "Primera p&aacute;gina")
    end if

    if TieneSiguiente then
        formsAccion = formsAccion & "<form id='formSigue' class='form-oculto' method='post' action='" & UrlPagiBase & "&cantregmover=" & (cantregmover + cantregmostrar) & "'></form>"
        botonSig = UiBoton("Siguiente", "secondary", "document.getElementById('formSigue').submit();", True, "")
    else
        botonSig = UiBoton("Siguiente", "secondary", "", False, "&Uacute;ltima p&aacute;gina")
    end if

    html = html & UiPaginacionDe(CDbl(CantRegMover) + 1, CDbl(CantRegMover) + cantFilas, TotalRegistrosTotal, botonAnt, botonSig)
Else
    ' Consulta: misma paginacion, reenviando los filtros (antes mostraba solo los primeros resultados)
    if cantregmover <> 0 then
        botonAnt = UiBoton("Anterior", "secondary", "irConsulta('" & UrlConsulta & "&ord=" & ord & "&cantregmover=" & (cantregmover - cantregmostrar) & "');", True, "")
    else
        botonAnt = UiBoton("Anterior", "secondary", "", False, "Primera p&aacute;gina")
    end if
    if TieneSiguiente then
        botonSig = UiBoton("Siguiente", "secondary", "irConsulta('" & UrlConsulta & "&ord=" & ord & "&cantregmover=" & (cantregmover + cantregmostrar) & "');", True, "")
    else
        botonSig = UiBoton("Siguiente", "secondary", "", False, "&Uacute;ltima p&aacute;gina")
    end if
    html = html & UiPaginacionDe(CDbl(CantRegMover) + 1, CDbl(CantRegMover) + cantFilas, TotalRegistrosTotal, botonAnt, botonSig)
End If

html = html & UiCardCierra()
html = html & UiPantallaCierra()
html = html & "</div>" ' Cierre div principal

' --- SCRIPTS JS ---
html = html & "<script type='text/javascript'>"
html = html & " function irConsulta(url){ var f = document.getElementById('formFiltros'); f.action = url; f.submit(); return false; } "
html = html & " function borrar(formu){ "
html = html & " var r = confirm('\u00bfDesea eliminar el proyecto?'); "
html = html & " if (r){ "
html = html & " document.getElementById(formu).submit(); "
html = html & " }else{ "
html = html & " return false;"
html = html & " }} "
html = html & "</script>"

html = html & "<script src='assets/js/jquery.min.js'></script>"
html = html & "<script src='assets/js/skel.min.js'></script>"
html = html & "<script src='assets/js/util.js'></script>"
html = html & "<script src='assets/js/main.js'></script>"

' Inyeccion de formularios ocultos acumulados
html = html & formsAccion
html = html & "</body></html>"

' --- CIERRE DE CONEXION ---
Cnn.Close
Set Cnn = Nothing
' If HayRegistros Then rsGrid.Close : 
Set rsGrid = Nothing
Set rsClientes = Nothing
Set rsApli = Nothing

Response.Write html
%>