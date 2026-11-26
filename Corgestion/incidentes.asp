<!-- #INCLUDE FILE="IncFile/CorConect.inc" -->
<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/ui_componentes.inc" -->
<%
' [CG] Control de usuario E-Manager (igual que CorEManager/Acciones.asp)
cgOk = ofv.CheckUsuario()
If cgOk <> "Y" Then
	cgUv = ofv.uv
	If cgUv = "NO" Then cgOk = ofv.ValidarUsuario
	Response.Write cgOk
	Response.End
End If

'Option Explicit

' =========================================================================
' TODO EL CODIGO VIVE EN UN UNICO BLOQUE ASP. TODA SALIDA SE HACE
' MEDIANTE Response.Write / concatenacion de strings.
' =========================================================================

' --- CONEXION A BASE DE DATOS ---
Dim Cnn
Set Cnn = Server.CreateObject("ADODB.Connection")
On Error Resume Next
Cnn.Open scon1
If Err.Number <> 0 Then
	Dim sErrConn
	sErrConn = "<div style='color:red;font-family:monospace;padding:10px;border:1px solid red;margin:10px;'>"
	sErrConn = sErrConn & "<b>ERROR DE CONEXION A BASE DE DATOS</b><br>"
	sErrConn = sErrConn & "Codigo: " & Err.Number & "<br>"
	sErrConn = sErrConn & "Detalle: " & Server.HTMLEncode(Err.Description) & "</div>"
	Response.Write sErrConn
	Response.End
End If
On Error Goto 0

' --- DECLARACION OBLIGATORIA DE TODAS LAS VARIABLES ---
Dim opc, xmenu, xtmenu, ord, cliente, tit, aplicativo, aplic, CantRegMover, CantRegMostrar
Dim sel, descri, estado, fecha, XXX, nFecini
Dim sHtml, sPage

opc = Request.QueryString("opc")
xmenu = Request.QueryString("xmenu")
xtmenu = Request.QueryString("xtmenu")
ord = Request.QueryString("ord")

nFecini = 0
If nFecini < CLng(CDate("01-01-1900")) Then
    nFecini = CLng(CDate("01-01-1900"))
End If

' =========================================================================
' CAPA DE DATOS Y NEGOCIO (DAL/BLL)
' =========================================================================

Function ObtenerListadoIncidentes(cn, idCli, idApli, titulo, orden, respon)
    Dim cmd
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandText = "sp_cg_Incidentes_Listar"
    cmd.CommandType = 4 ' adCmdStoredProc

    cmd.Parameters.Append cmd.CreateParameter("@IdCliente", 3, 1, , CLng(idCli))
    cmd.Parameters.Append cmd.CreateParameter("@IdAplicativo", 3, 1, , CLng(idApli))
    cmd.Parameters.Append cmd.CreateParameter("@Titulo", 200, 1, 255, CStr(titulo & ""))
    cmd.Parameters.Append cmd.CreateParameter("@Orden", 200, 1, 20, CStr(orden & ""))
    ' [CG] filtro por responsable (usuario de red); "" = todos
    If Trim(respon & "") = "" Then
        cmd.Parameters.Append cmd.CreateParameter("@Respon", 200, 1, 100, Null)
    Else
        cmd.Parameters.Append cmd.CreateParameter("@Respon", 200, 1, 100, CStr(respon))
    End If

    Set ObtenerListadoIncidentes = cmd.Execute()
End Function

Function ConsultarIncidentes(cn, idCli, titulo, est, idApli, fec)
    Dim cmd
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandText = "sp_cg_Incidentes_Consultar"
    cmd.CommandType = 4

    cmd.Parameters.Append cmd.CreateParameter("@IdCliente", 3, 1, , CLng(idCli))
    cmd.Parameters.Append cmd.CreateParameter("@Titulo", 200, 1, 255, CStr(titulo & ""))
    cmd.Parameters.Append cmd.CreateParameter("@Estado", 3, 1, , CLng(est))
    cmd.Parameters.Append cmd.CreateParameter("@IdAplicativo", 3, 1, , CLng(idApli))

    If IsNull(fec) Or fec = "" Then
        cmd.Parameters.Append cmd.CreateParameter("@Fecha", 3, 1, , Null)
    Else
        cmd.Parameters.Append cmd.CreateParameter("@Fecha", 3, 1, , CLng(fec))
    End If

    Set ConsultarIncidentes = cmd.Execute()
End Function

' Catalogos de combo (clientes, aplicativos, etc.) se cargan con
' GetListaComboExtended(), definida en IncFile/funciones.inc -- no duplicar
' aca una funcion local equivalente.

'[UI-ANT] ' ===== Helpers anteriores: reemplazados por UiBadgeEstado / UiBadgePrioridad (IncFile/ui_componentes.inc) =====
'[UI-ANT] ' =========================================================================
'[UI-ANT] ' HELPERS DE PRESENTACION (arman fragmentos HTML como string)
'[UI-ANT] ' =========================================================================

'[UI-ANT] Function ArmarBadgeEstado(estadoVal, nombreResp)
    '[UI-ANT] Dim s, respSeguro
    '[UI-ANT] respSeguro = Server.HTMLEncode(nombreResp & "")
    '[UI-ANT] Select Case estadoVal
        '[UI-ANT] Case "PENDIENTE": s = "<span class='badge badge-media ui-badge ui-badge--warning' aria-label='Pendiente'>PEND</span> (" & respSeguro & ")"
        '[UI-ANT] Case "OK": s = "<span class='badge badge-ok ui-badge ui-badge--ok' aria-label='OK'>OK</span> (" & respSeguro & ")"
        '[UI-ANT] Case "BAJA": s = "<span class='badge badge-peligro ui-badge ui-badge--danger' aria-label='Baja'>BAJA</span> (" & respSeguro & ")"
        '[UI-ANT] Case "TESTING EN CLIENTE": s = "<span class='badge badge-media ui-badge ui-badge--warning' aria-label='Testing en cliente'>TEST</span> (" & respSeguro & ")"
        '[UI-ANT] Case Else: s = "-------------"
    '[UI-ANT] End Select
    '[UI-ANT] ArmarBadgeEstado = s
'[UI-ANT] End Function

'[UI-ANT] Function ArmarBadgeEstadoConsulta(estadoVal)
    '[UI-ANT] Dim s
    '[UI-ANT] Select Case estadoVal
        '[UI-ANT] Case "PENDIENTE": s = "<span class='badge badge-media ui-badge ui-badge--warning' aria-label='Pendiente'>PENDIENTE</span>"
        '[UI-ANT] Case "OK": s = "<span class='badge badge-ok ui-badge ui-badge--ok' aria-label='OK'>OK</span>"
        '[UI-ANT] Case "BAJA": s = "<span class='badge badge-peligro ui-badge ui-badge--danger' aria-label='Baja'>BAJA</span>"
        '[UI-ANT] Case Else: s = ""
    '[UI-ANT] End Select
    '[UI-ANT] ArmarBadgeEstadoConsulta = s
'[UI-ANT] End Function

'[UI-ANT] Function ArmarBadgePrioridad(prioridadVal)
    '[UI-ANT] Dim s, prioSeguro
    '[UI-ANT] prioSeguro = Server.HTMLEncode(prioridadVal & "")
    '[UI-ANT] If prioridadVal = "ALTA" Then
        '[UI-ANT] s = "<span class='badge badge-peligro ui-badge ui-badge--danger' aria-label='Prioridad alta'>ALTA</span>"
    '[UI-ANT] ElseIf prioridadVal = "MEDIA" Then
        '[UI-ANT] s = "<span class='badge badge-media ui-badge ui-badge--warning' aria-label='Prioridad media'>MEDIA</span>"
    '[UI-ANT] Else
        '[UI-ANT] s = "<span class='badge badge-pendiente ui-badge ui-badge--pending' aria-label='Prioridad " & prioSeguro & "'>" & prioSeguro & "</span>"
    '[UI-ANT] End If
    '[UI-ANT] ArmarBadgePrioridad = s
'[UI-ANT] End Function

'[UI-ANT] ' ===== Vista anterior completa: reemplazada por el bloque [UI] que sigue. Se borra entera buscando [UI-ANT] =====
'[UI-ANT] ' =========================================================================
'[UI-ANT] ' VISTA / PRESENTACION (todo armado en strings, un solo Response.Write al final)
'[UI-ANT] ' =========================================================================

'[UI-ANT] sPage = ""
'[UI-ANT] sPage = sPage & headhtml()

'[UI-ANT] sPage = sPage & "<body id=""body"" name=""body"" style=""height:100%;width:100%;"">" & vbCrLf
'[UI-ANT] sPage = sPage & "<div id=""wrapper"">" & vbCrLf
'[UI-ANT] sPage = sPage & cabecera()
'[UI-ANT] sPage = sPage & "</div>" & vbCrLf
'[UI-ANT] sPage = sPage & getMenuParam()

'[UI-ANT] If opc <> "consi" Then
    '[UI-ANT] ' --- VISTA PRINCIPAL: LISTADO DE INCIDENTES ---
    '[UI-ANT] XXX = SYSLOGON(Cnn, "MENU_INCIDENTES")

    '[UI-ANT] cliente = Request.Form("cli")
    '[UI-ANT] If cliente & "" = "" Then cliente = Request.QueryString("cli")
    '[UI-ANT] If Not IsNumeric(cliente) Then cliente = 0

    '[UI-ANT] tit = Request.Form("tit")
    '[UI-ANT] If tit & "" = "" Then tit = Request.QueryString("tit")

    '[UI-ANT] aplicativo = Request.Form("apli")
    '[UI-ANT] If aplicativo & "" = "" Then aplicativo = Request.QueryString("apli")
    '[UI-ANT] If Not IsNumeric(aplicativo) Then aplicativo = 0

    '[UI-ANT] CantRegMostrar = 10
    '[UI-ANT] CantRegMover = Request.QueryString("CantRegMover")
    '[UI-ANT] If Not IsNumeric(CantRegMover) Or CantRegMover & "" = "" Then
        '[UI-ANT] CantRegMover = 0
    '[UI-ANT] Else
        '[UI-ANT] CantRegMover = CDbl(CantRegMover)
    '[UI-ANT] End If
    '[UI-ANT] If CantRegMover < 0 Then CantRegMover = 0

    '[UI-ANT] ' Consulta de datos mediante SP
    '[UI-ANT] Dim rsListado
    '[UI-ANT] Set rsListado = ObtenerListadoIncidentes(Cnn, cliente, aplicativo, tit, ord)

    '[UI-ANT] If Not rsListado.EOF Then rsListado.Move(CantRegMover)

    '[UI-ANT] sPage = sPage & "<div class=""inner cg-page ui-main"">" & vbCrLf
    '[UI-ANT] sPage = sPage & "    <div class=""titulo-seccion cg-titlebar ui-page-header"">" & vbCrLf
    '[UI-ANT] sPage = sPage & "        <div class=""page-toolbar cg-titlebar-row ui-toolbar"">" & vbCrLf
    '[UI-ANT] sPage = sPage & "                <div class=""page-toolbar-title cg-titlebar-title""><h2>INCIDENTES</h2></div>" & vbCrLf
    '[UI-ANT] sPage = sPage & "                <div class=""page-toolbar-filters cg-titlebar-filters ui-toolbar__filters"">" & vbCrLf
    '[UI-ANT] sPage = sPage & "                    <form id=""xForm"" method=""post"" action=""incidentes.asp?ord=" & ord & "&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & """>" & vbCrLf

    '[UI-ANT] ' Combo Clientes
    '[UI-ANT] sPage = sPage & "                        <select class=""ui-select"" onchange=""document.getElementById('xForm').submit();"" id=""cli"" name=""cli"">" & vbCrLf
    '[UI-ANT] sPage = sPage & "                            <option value=""0"">CLIENTE: TODOS</option>" & vbCrLf
    '[UI-ANT] Dim rsCli, selCli
    '[UI-ANT] Set rsCli = GetListaComboExtended(Cnn, "clientes")
    '[UI-ANT] Do While Not rsCli.EOF
        '[UI-ANT] If CStr(rsCli("id")) = CStr(cliente) Then selCli = "selected" Else selCli = ""
        '[UI-ANT] sPage = sPage & "                            <option value=""" & rsCli("id") & """ " & selCli & ">" & Server.HTMLEncode(rsCli("descripcion") & "") & "</option>" & vbCrLf
        '[UI-ANT] rsCli.MoveNext
    '[UI-ANT] Loop
    '[UI-ANT] rsCli.Close
    '[UI-ANT] sPage = sPage & "                        </select>" & vbCrLf

    '[UI-ANT] ' Combo Aplicativos
    '[UI-ANT] sPage = sPage & "                        <select class=""ui-select"" onchange=""document.getElementById('xForm').submit();"" id=""apli"" name=""apli"">" & vbCrLf
    '[UI-ANT] sPage = sPage & "                            <option value=""0"">APLICATIVO: TODOS</option>" & vbCrLf
    '[UI-ANT] Dim rsApli, selApli
    '[UI-ANT] Set rsApli = GetListaComboExtended(Cnn, "aplicativos")
    '[UI-ANT] Do While Not rsApli.EOF
        '[UI-ANT] If CStr(rsApli("id")) = CStr(aplicativo) Then selApli = "selected" Else selApli = ""
        '[UI-ANT] sPage = sPage & "                            <option value=""" & rsApli("id") & """ " & selApli & ">" & Server.HTMLEncode(rsApli("descripcion") & "") & "</option>" & vbCrLf
        '[UI-ANT] rsApli.MoveNext
    '[UI-ANT] Loop
    '[UI-ANT] rsApli.Close
    '[UI-ANT] sPage = sPage & "                        </select>" & vbCrLf
    '[UI-ANT] sPage = sPage & "                    </form>" & vbCrLf
    '[UI-ANT] sPage = sPage & "                </div>" & vbCrLf

    '[UI-ANT] ' Acciones (alta)
    '[UI-ANT] If Session("FUNC") = "FUNC" Or Session("FUNC") = "TEC" Then
        '[UI-ANT] sPage = sPage & "                <div class=""page-toolbar-action cg-titlebar-actions ui-toolbar__actions"">" & vbCrLf
        '[UI-ANT] sPage = sPage & "                    <form id=""formx4"" class=""form-oculto"" method=""post"" action=""agregar2.asp?tipo=INC&xmenu=" & xmenu & "&xtmenu=" & xtmenu & """></form>" & vbCrLf
        '[UI-ANT] If CStr(getPermisos(xmenu, xtmenu, "ALTA")) = "1" Then
            '[UI-ANT] sPage = sPage & "                    <input class=""bot ui-button ui-button--primary"" type=""button"" value=""CARGAR NUEVO INCIDENTE"" onclick=""document.getElementById('formx4').submit();"">" & vbCrLf
        '[UI-ANT] Else
            '[UI-ANT] sPage = sPage & "                    <input class=""botsp ui-button"" type=""button"" value=""CARGAR NUEVO INCIDENTE"" disabled title=""Sin permiso para agregar."">" & vbCrLf
        '[UI-ANT] End If
        '[UI-ANT] sPage = sPage & "                </div>" & vbCrLf
    '[UI-ANT] End If

    '[UI-ANT] ' Paginacion: ATRAS
    '[UI-ANT] sPage = sPage & "                <div class=""page-toolbar-action cg-titlebar-actions ui-toolbar__actions ui-pagination"">" & vbCrLf
    '[UI-ANT] If CantRegMover <> 0 Then
        '[UI-ANT] sPage = sPage & "                    <form id=""formx22"" class=""form-oculto"" method=""post"" action=""incidentes.asp?opc=listar&cantregmover=" & (CantRegMover - CantRegMostrar) & "&ord=" & ord & "&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & """></form>" & vbCrLf
        '[UI-ANT] sPage = sPage & "                    <input class=""bot ui-button"" type=""button"" value=""ATRAS"" onclick=""document.getElementById('formx22').submit();"">" & vbCrLf
    '[UI-ANT] Else
        '[UI-ANT] sPage = sPage & "                    <input class=""botsp ui-button"" type=""button"" value=""ATRAS"" disabled>" & vbCrLf
    '[UI-ANT] End If
    '[UI-ANT] sPage = sPage & "                </div>" & vbCrLf

    '[UI-ANT] ' Paginacion: SIGUIENTE
    '[UI-ANT] sPage = sPage & "                <div class=""page-toolbar-action cg-titlebar-actions ui-toolbar__actions ui-pagination"">" & vbCrLf
    '[UI-ANT] If Not rsListado.EOF Then
        '[UI-ANT] sPage = sPage & "                    <form id=""formx23"" class=""form-oculto"" method=""post"" action=""incidentes.asp?opc=listar&cantregmover=" & (CantRegMover + CantRegMostrar) & "&ord=" & ord & "&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & """></form>" & vbCrLf
        '[UI-ANT] sPage = sPage & "                    <input class=""bot ui-button"" type=""button"" value=""SIGUIENTE"" onclick=""document.getElementById('formx23').submit();"">" & vbCrLf
    '[UI-ANT] Else
        '[UI-ANT] sPage = sPage & "                    <input class=""botsp ui-button"" type=""button"" value=""SIGUIENTE"" disabled>" & vbCrLf
    '[UI-ANT] End If
    '[UI-ANT] sPage = sPage & "                </div>" & vbCrLf

    '[UI-ANT] ' Volver
    '[UI-ANT] sPage = sPage & "                <div class=""page-toolbar-action cg-titlebar-actions ui-toolbar__actions"">" & vbCrLf
    '[UI-ANT] sPage = sPage & "                    <form id=""formx5"" class=""form-oculto"" method=""post"" action=""index.asp""></form>" & vbCrLf
    '[UI-ANT] sPage = sPage & "                    <input class=""bot ui-button"" type=""button"" value=""VOLVER"" onclick=""document.getElementById('formx5').submit();"">" & vbCrLf
    '[UI-ANT] sPage = sPage & "                </div>" & vbCrLf
    '[UI-ANT] sPage = sPage & "        </div>" & vbCrLf
    '[UI-ANT] sPage = sPage & "    </div>" & vbCrLf

    '[UI-ANT] ' --- GRILLA PRINCIPAL ---
    '[UI-ANT] sPage = sPage & "    <div class=""bloque-datos table-responsive ui-table-wrap"">" & vbCrLf
    '[UI-ANT] sPage = sPage & "        <table class=""tabla-moderna ui-table"">" & vbCrLf
    '[UI-ANT] sPage = sPage & "            <thead>" & vbCrLf
    '[UI-ANT] sPage = sPage & "                <tr>" & vbCrLf
    '[UI-ANT] sPage = sPage & "                    <th></th>" & vbCrLf
    '[UI-ANT] sPage = sPage & "                    <th></th>" & vbCrLf
    '[UI-ANT] sPage = sPage & "                    <th>ID</th>" & vbCrLf
    '[UI-ANT] sPage = sPage & "                    <th><a class=""col-link"" style=""text-decoration: none;"" href=""incidentes.asp?ord=titulo&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & """>TITULO</a></th>" & vbCrLf
    '[UI-ANT] sPage = sPage & "                    <th><a class=""col-link"" style=""text-decoration: none;"" href=""incidentes.asp?ord=cliente&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & """>CLIENTE</a></th>" & vbCrLf
    '[UI-ANT] sPage = sPage & "                    <th>PRIORIDAD</th>" & vbCrLf
    '[UI-ANT] sPage = sPage & "                    <th>APLICATIVO</th>" & vbCrLf
    '[UI-ANT] sPage = sPage & "                    <th>DESARROLLO</th>" & vbCrLf
    '[UI-ANT] sPage = sPage & "                    <th>FUNCIONAL</th>" & vbCrLf
    '[UI-ANT] sPage = sPage & "                    <th>FEC INGRESO</th>" & vbCrLf
    '[UI-ANT] sPage = sPage & "                </tr>" & vbCrLf
    '[UI-ANT] sPage = sPage & "            </thead>" & vbCrLf
    '[UI-ANT] sPage = sPage & "            <tbody>" & vbCrLf

    '[UI-ANT] Dim PosReg, est1, est2, prio, sFec, idReg, urlDetalle, xBtnElim
    '[UI-ANT] PosReg = 0

    '[UI-ANT] Dim tituloSeguro, tituloTrunc, clienteSeguro, aplicativoSeguro
    '[UI-ANT] Do While (Not rsListado.EOF) And (PosReg < CantRegMostrar)
        '[UI-ANT] idReg = rsListado("id")
        '[UI-ANT] est1 = ArmarBadgeEstado(rsListado("estado_desa"), rsListado("respon_tec_nombre"))
        '[UI-ANT] est2 = ArmarBadgeEstado(rsListado("estado_func"), rsListado("respon_gestion_nombre"))
        '[UI-ANT] prio = ArmarBadgePrioridad(rsListado("prioridad"))
        '[UI-ANT] sFec = Fechador(rsListado("fecha_ingreso") + nFecini)
        '[UI-ANT] urlDetalle = "detalles.asp?id=" & idReg & "&opc=veri&cli=" & cliente & "&tit=" & tit & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu
        '[UI-ANT] tituloSeguro = Server.HTMLEncode(rsListado("titulo") & "")
        '[UI-ANT] tituloTrunc = Server.HTMLEncode(UCase(Left(rsListado("titulo") & "", 40)))
        '[UI-ANT] clienteSeguro = Server.HTMLEncode(rsListado("cliente") & "")
        '[UI-ANT] aplicativoSeguro = Server.HTMLEncode(rsListado("aplicativo") & "")

        '[UI-ANT] sPage = sPage & "                <form id=""formx" & idReg & """ class=""form-oculto"" method=""post"" action=""eliminar.asp?opc=inc&id=" & idReg & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=incidentes.asp""></form>" & vbCrLf
        '[UI-ANT] sPage = sPage & "                <form id=""formx2" & idReg & """ class=""form-oculto"" method=""post"" action=""detalles.asp?id=" & idReg & "&opc=acti&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & """></form>" & vbCrLf

        '[UI-ANT] sPage = sPage & "                <tr>" & vbCrLf
        '[UI-ANT] sPage = sPage & "                    <td style=""text-align: center;"">" & vbCrLf
        '[UI-ANT] If CStr(getPermisos(xmenu, xtmenu, "ELIM")) = "1" Then
            '[UI-ANT] sPage = sPage & "                        <input class=""btn-accion btn-accion-peligro ui-button ui-button--icon ui-button--danger"" type=""button"" value=""X"" onclick=""borrar('formx" & idReg & "');"">" & vbCrLf
        '[UI-ANT] Else
            '[UI-ANT] sPage = sPage & "                        <input class=""btn-accion btn-accion-deshabilitado ui-button ui-button--icon"" type=""button"" value=""X"" disabled title=""Sin permiso para eliminar."">" & vbCrLf
        '[UI-ANT] End If
        '[UI-ANT] sPage = sPage & "                    </td>" & vbCrLf
        '[UI-ANT] sPage = sPage & "                    <td style=""text-align: center;"">" & vbCrLf
        '[UI-ANT] sPage = sPage & "                        <input class=""btn-accion btn-accion-primario ui-button ui-button--icon ui-button--primary"" type=""button"" value=""A"" onclick=""document.getElementById('formx2" & idReg & "').submit();"">" & vbCrLf
        '[UI-ANT] sPage = sPage & "                    </td>" & vbCrLf
        '[UI-ANT] sPage = sPage & "                    <td onclick=""document.location='" & urlDetalle & "'"">" & idReg & "</td>" & vbCrLf
        '[UI-ANT] sPage = sPage & "                    <td title=""" & tituloSeguro & """ onclick=""document.location='" & urlDetalle & "'"">" & vbCrLf
        '[UI-ANT] sPage = sPage & "                        " & tituloTrunc & vbCrLf
        '[UI-ANT] If Trim(rsListado("respon_tec") & "") = "0" Then
            '[UI-ANT] sPage = sPage & "                        <span class=""ui-badge ui-badge--pending"" aria-label=""Consulta""><b>(C)</b></span>" & vbCrLf
        '[UI-ANT] End If
        '[UI-ANT] sPage = sPage & "                    </td>" & vbCrLf
        '[UI-ANT] sPage = sPage & "                    <td onclick=""document.location='" & urlDetalle & "'"">" & clienteSeguro & "</td>" & vbCrLf
        '[UI-ANT] sPage = sPage & "                    <td onclick=""document.location='" & urlDetalle & "'"">" & prio & "</td>" & vbCrLf
        '[UI-ANT] sPage = sPage & "                    <td onclick=""document.location='" & urlDetalle & "'"">" & aplicativoSeguro & "</td>" & vbCrLf
        '[UI-ANT] sPage = sPage & "                    <td onclick=""document.location='" & urlDetalle & "'"">" & est1 & "</td>" & vbCrLf
        '[UI-ANT] sPage = sPage & "                    <td onclick=""document.location='" & urlDetalle & "'"">" & est2 & "</td>" & vbCrLf
        '[UI-ANT] sPage = sPage & "                    <td onclick=""document.location='" & urlDetalle & "'"">" & sFec & "</td>" & vbCrLf
        '[UI-ANT] sPage = sPage & "                </tr>" & vbCrLf

        '[UI-ANT] PosReg = PosReg + 1
        '[UI-ANT] rsListado.MoveNext
    '[UI-ANT] Loop
    '[UI-ANT] rsListado.Close

    '[UI-ANT] sPage = sPage & "            </tbody>" & vbCrLf
    '[UI-ANT] sPage = sPage & "        </table>" & vbCrLf
    '[UI-ANT] sPage = sPage & "    </div>" & vbCrLf

    '[UI-ANT] sPage = sPage & "    <div class=""cg-legend table-legend ui-legend"">" & vbCrLf
    '[UI-ANT] sPage = sPage & "        <span><span class=""table-legend-mark""><b>*</b></span><span class=""table-legend-consult""><b>(C)</b></span></span>" & vbCrLf
    '[UI-ANT] sPage = sPage & "        <span><span class=""table-legend-mark"">&nbsp;&nbsp;(Consulta) resuelta sin incumbencia tecnica.</span></span>" & vbCrLf
    '[UI-ANT] sPage = sPage & "    </div>" & vbCrLf
    '[UI-ANT] sPage = sPage & "</div>" & vbCrLf

'[UI-ANT] ElseIf opc = "consi" Then
    '[UI-ANT] ' --- VISTA 2: CONSULTA DE INCIDENTES ---
    '[UI-ANT] sel = Request.QueryString("sel")
    '[UI-ANT] XXX = SYSLOGON(Cnn, "CONSULTA_INCIDENTES")

    '[UI-ANT] cliente = Request.Form("cli")
    '[UI-ANT] If cliente & "" = "" Then cliente = Request.QueryString("cli")
    '[UI-ANT] If Not IsNumeric(cliente) Then cliente = 0

    '[UI-ANT] descri = Request.Form("desc")
    '[UI-ANT] If descri & "" = "" Then descri = Request.QueryString("desc")

    '[UI-ANT] estado = Request.Form("est")
    '[UI-ANT] If estado & "" = "" Then estado = Request.QueryString("est")
    '[UI-ANT] If Not IsNumeric(estado) Then estado = 0

    '[UI-ANT] aplic = Request.Form("apli")
    '[UI-ANT] If aplic & "" = "" Then aplic = Request.QueryString("apli")
    '[UI-ANT] If Not IsNumeric(aplic) Then aplic = 0

    '[UI-ANT] fecha = Request.Form("fec")
    '[UI-ANT] If fecha & "" = "" Then fecha = Request.QueryString("fec")
    '[UI-ANT] If IsDate(fecha) Then
        '[UI-ANT] fecha = CLng(CDate(fecha)) - nFecini
    '[UI-ANT] Else
        '[UI-ANT] fecha = Null
    '[UI-ANT] End If

    '[UI-ANT] ' Consulta de datos mediante SP
    '[UI-ANT] Dim rsConsulta
    '[UI-ANT] Set rsConsulta = ConsultarIncidentes(Cnn, cliente, descri, estado, aplic, fecha)

    '[UI-ANT] sPage = sPage & "<div class=""inner cg-page ui-main"">" & vbCrLf
    '[UI-ANT] sPage = sPage & "    <div class=""titulo-seccion cg-titlebar ui-page-header"">" & vbCrLf
    '[UI-ANT] sPage = sPage & "        <div class=""page-toolbar cg-titlebar-row ui-toolbar"">" & vbCrLf
    '[UI-ANT] sPage = sPage & "            <div class=""page-toolbar-title cg-titlebar-title""><h2>CONSULTA INCIDENTES</h2></div>" & vbCrLf
    '[UI-ANT] sPage = sPage & "            <div class=""page-toolbar-action cg-titlebar-actions ui-toolbar__actions"">" & vbCrLf
    '[UI-ANT] sPage = sPage & "                <form id=""formx10"" class=""form-oculto"" method=""post"" action=""consultas.asp?sel=" & sel & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & """></form>" & vbCrLf
    '[UI-ANT] sPage = sPage & "                <input class=""bot ui-button"" type=""button"" value=""VOLVER"" onclick=""document.getElementById('formx10').submit();"">" & vbCrLf
    '[UI-ANT] sPage = sPage & "            </div>" & vbCrLf
    '[UI-ANT] sPage = sPage & "        </div>" & vbCrLf
    '[UI-ANT] sPage = sPage & "    </div>" & vbCrLf

    '[UI-ANT] sPage = sPage & "    <div class=""bloque-datos table-responsive ui-table-wrap"">" & vbCrLf
    '[UI-ANT] sPage = sPage & "        <table class=""tabla-moderna ui-table"" style=""width: 100%;"">" & vbCrLf
    '[UI-ANT] sPage = sPage & "            <thead>" & vbCrLf
    '[UI-ANT] sPage = sPage & "                <tr>" & vbCrLf
    '[UI-ANT] sPage = sPage & "                    <th style=""width: 5%; text-align: center;"">ID</th>" & vbCrLf
    '[UI-ANT] sPage = sPage & "                    <th style=""width: 40%;"">TITULO</th>" & vbCrLf
    '[UI-ANT] sPage = sPage & "                    <th style=""width: 40px; text-align: center;"">&nbsp;</th>" & vbCrLf
    '[UI-ANT] sPage = sPage & "                    <th style=""width: 18%;"">CLIENTE</th>" & vbCrLf
    '[UI-ANT] sPage = sPage & "                    <th style=""width: 10.5%; text-align: center;"">DESARROLLO</th>" & vbCrLf
    '[UI-ANT] sPage = sPage & "                    <th style=""width: 10.5%; text-align: center;"">FUNCIONAL</th>" & vbCrLf
    '[UI-ANT] sPage = sPage & "                    <th style=""width: 15%; text-align: center;"">FEC INGRESO</th>" & vbCrLf
    '[UI-ANT] sPage = sPage & "                </tr>" & vbCrLf
    '[UI-ANT] sPage = sPage & "            </thead>" & vbCrLf
    '[UI-ANT] sPage = sPage & "            <tbody>" & vbCrLf

    '[UI-ANT] Dim estD, estF, idCons, tituloConsSeguro, clienteConsSeguro
    '[UI-ANT] Do While Not rsConsulta.EOF
        '[UI-ANT] idCons = rsConsulta("id")
        '[UI-ANT] estD = ArmarBadgeEstadoConsulta(rsConsulta("estado_desa"))
        '[UI-ANT] estF = ArmarBadgeEstadoConsulta(rsConsulta("estado_func"))
        '[UI-ANT] tituloConsSeguro = Server.HTMLEncode(UCase(Mid(rsConsulta("titulo") & "", 1, 58)))
        '[UI-ANT] clienteConsSeguro = Server.HTMLEncode(rsConsulta("cliente") & "")

        '[UI-ANT] sPage = sPage & "                <form id=""xform" & idCons & """ class=""form-oculto"" method=""post"" action=""detalles.asp?id=" & idCons & "&opc=veri&tipo=cons""></form>" & vbCrLf
        '[UI-ANT] sPage = sPage & "                <tr>" & vbCrLf
        '[UI-ANT] sPage = sPage & "                    <td style=""text-align: center; color: #475569; font-weight: 500;"">" & idCons & "</td>" & vbCrLf
        '[UI-ANT] sPage = sPage & "                    <td>" & tituloConsSeguro & "</td>" & vbCrLf
        '[UI-ANT] sPage = sPage & "                    <td style=""text-align: center; vertical-align: middle;"">" & vbCrLf
        '[UI-ANT] If getPermisos(9, "INICIO", "CONS") = 1 Then
            '[UI-ANT] sPage = sPage & "                        <input class=""btn-accion btn-accion-primario ui-button ui-button--icon ui-button--primary"" type=""button"" value=""D"" onclick=""document.getElementById('xform" & idCons & "').submit();"">" & vbCrLf
        '[UI-ANT] Else
            '[UI-ANT] sPage = sPage & "                        <input class=""btn-accion btn-accion-deshabilitado ui-button ui-button--icon"" type=""button"" value=""D"" disabled title=""Sin permiso para consultar."">" & vbCrLf
        '[UI-ANT] End If
        '[UI-ANT] sPage = sPage & "                    </td>" & vbCrLf
        '[UI-ANT] sPage = sPage & "                    <td style=""padding-left: 1rem;"">" & clienteConsSeguro & "</td>" & vbCrLf
        '[UI-ANT] sPage = sPage & "                    <td style=""text-align: center;"">" & estD & "</td>" & vbCrLf
        '[UI-ANT] sPage = sPage & "                    <td style=""text-align: center;"">" & estF & "</td>" & vbCrLf
        '[UI-ANT] sPage = sPage & "                    <td style=""text-align: center;"">" & Fechador(rsConsulta("fecha_ingreso") + nFecini) & "</td>" & vbCrLf
        '[UI-ANT] sPage = sPage & "                </tr>" & vbCrLf

        '[UI-ANT] rsConsulta.MoveNext
    '[UI-ANT] Loop
    '[UI-ANT] rsConsulta.Close

    '[UI-ANT] sPage = sPage & "            </tbody>" & vbCrLf
    '[UI-ANT] sPage = sPage & "        </table>" & vbCrLf
    '[UI-ANT] sPage = sPage & "    </div>" & vbCrLf
    '[UI-ANT] sPage = sPage & "</div>" & vbCrLf
'[UI-ANT] End If

'[UI-ANT] ' --- SCRIPT Y CIERRE DE DOCUMENTO ---
'[UI-ANT] sPage = sPage & "<script type=""text/javascript"">" & vbCrLf
'[UI-ANT] sPage = sPage & "    function borrar(formu){ " & vbCrLf
'[UI-ANT] sPage = sPage & "        var r = confirm('?Desea eliminar el incidente?'); " & vbCrLf
'[UI-ANT] sPage = sPage & "        if (r){ " & vbCrLf
'[UI-ANT] sPage = sPage & "            document.getElementById(formu).submit(); " & vbCrLf
'[UI-ANT] sPage = sPage & "        } else { " & vbCrLf
'[UI-ANT] sPage = sPage & "            return false;" & vbCrLf
'[UI-ANT] sPage = sPage & "        }" & vbCrLf
'[UI-ANT] sPage = sPage & "    }" & vbCrLf
'[UI-ANT] sPage = sPage & "</script>" & vbCrLf
'[UI-ANT] sPage = sPage & "</body>" & vbCrLf
'[UI-ANT] sPage = sPage & "</html>"

' =========================================================================
' [UI] VISTA / PRESENTACION con el sistema visual (IncFile/ui_componentes.inc)
' El bloque anterior completo queda arriba, comentado con '[UI-ANT].
' Las lecturas de datos, permisos y procedimientos son las mismas de antes.
' =========================================================================

sPage = ""
sPage = sPage & headhtml()

sPage = sPage & "<body id=""body"" name=""body"" style=""height:100%;width:100%;"">" & vbCrLf
' La cabecera del sistema es el encabezado principal de la pagina (region banner)
sPage = sPage & "<div id=""wrapper"" role=""banner"">" & vbCrLf
sPage = sPage & cabecera()
sPage = sPage & "</div>" & vbCrLf
sPage = sPage & getMenuParam()

' Formularios ocultos de las acciones: van fuera de la tabla (dentro de un tbody no son HTML valido)
Dim sForms
sForms = ""

If opc <> "consi" Then
    ' --- VISTA PRINCIPAL: LISTADO DE INCIDENTES ---
    XXX = SYSLOGON(Cnn, "MENU_INCIDENTES")

    cliente = Request.Form("cli")
    If cliente & "" = "" Then cliente = Request.QueryString("cli")
    If Not IsNumeric(cliente) Then cliente = 0

    tit = Request.Form("tit")
    If tit & "" = "" Then tit = Request.QueryString("tit")

    aplicativo = Request.Form("apli")
    If aplicativo & "" = "" Then aplicativo = Request.QueryString("apli")
    If Not IsNumeric(aplicativo) Then aplicativo = 0

    CantRegMostrar = 10
    CantRegMover = Request.QueryString("CantRegMover")
    If Not IsNumeric(CantRegMover) Or CantRegMover & "" = "" Then
        CantRegMover = 0
    Else
        CantRegMover = CDbl(CantRegMover)
    End If
    If CantRegMover < 0 Then CantRegMover = 0

    ' Consulta de datos mediante SP
    Dim rsListado
    Dim cgResp
    cgResp = CgRespSeleccionado() ' [CG] por defecto: registros del usuario de la sesion; "*" = todos
    Set rsListado = ObtenerListadoIncidentes(Cnn, cliente, aplicativo, tit, ord, CgRespParam(cgResp))

    If Not rsListado.EOF Then rsListado.Move(CantRegMover)

    ' Filtros que se conservan al ordenar y al paginar (antes se perdian el aplicativo y el titulo)
    Dim qsFiltros
    qsFiltros = "&cli=" & cliente & "&apli=" & aplicativo & "&resp=" & Server.URLEncode(cgResp) & "&tit=" & Server.URLEncode(tit & "") & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu

    ' --- Barra superior: titulo y acciones ---
    Dim accTop
    accTop = ""
    If Session("FUNC") = "FUNC" Or Session("FUNC") = "TEC" Then
        accTop = accTop & "<form id=""formx4"" class=""form-oculto"" method=""post"" action=""agregar2.asp?tipo=INC&xmenu=" & xmenu & "&xtmenu=" & xtmenu & """></form>"
        accTop = accTop & UiBoton("Cargar incidente", "light", "document.getElementById('formx4').submit();", CStr(getPermisosAtrm("CG_INCIDENTES", "ALTA")) = "1", "Sin permiso para agregar.")
    End If
    accTop = accTop & "<form id=""formx5"" class=""form-oculto"" method=""post"" action=""index.asp""></form>"
    accTop = accTop & UiBoton("Volver", "secondary", "document.getElementById('formx5').submit();", True, "")
    sPage = sPage & UiPantallaAbre("Incidentes", accTop, "wide") & vbCrLf

    ' --- Filtros en el encabezado de la tarjeta: se envian solos al cambiar, como antes ---
    Dim filtrosHtml
    filtrosHtml = "<form id=""xForm"" class=""ui-card__filtros"" method=""post"" action=""incidentes.asp?ord=" & ord & "&cli=" & cliente & "&resp=" & Server.URLEncode(cgResp) & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & """>"

    Dim rsCli, selCli, htmlCombo
    htmlCombo = "<select class=""ui-select"" onchange=""document.getElementById('xForm').submit();"" id=""cli"" name=""cli""><option value=""0"">Todos los clientes</option>"
    Set rsCli = GetListaComboExtended(Cnn, "clientes")
    Do While Not rsCli.EOF
        If CStr(rsCli("id")) = CStr(cliente) Then selCli = "selected" Else selCli = ""
        htmlCombo = htmlCombo & "<option value=""" & rsCli("id") & """ " & selCli & ">" & Server.HTMLEncode(rsCli("descripcion") & "") & "</option>"
        rsCli.MoveNext
    Loop
    rsCli.Close
    filtrosHtml = filtrosHtml & UiFiltroCompacto("cli", "Filtrar por cliente", htmlCombo & "</select>")

    Dim rsApli, selApli
    htmlCombo = "<select class=""ui-select"" onchange=""document.getElementById('xForm').submit();"" id=""apli"" name=""apli""><option value=""0"">Todos los aplicativos</option>"
    Set rsApli = GetListaComboExtended(Cnn, "aplicativos")
    Do While Not rsApli.EOF
        If CStr(rsApli("id")) = CStr(aplicativo) Then selApli = "selected" Else selApli = ""
        htmlCombo = htmlCombo & "<option value=""" & rsApli("id") & """ " & selApli & ">" & Server.HTMLEncode(rsApli("descripcion") & "") & "</option>"
        rsApli.MoveNext
    Loop
    rsApli.Close
    filtrosHtml = filtrosHtml & UiFiltroCompacto("apli", "Filtrar por aplicativo", htmlCombo & "</select>")
    ' [CG] Combo Responsable (por defecto: usuario de la sesion)
    filtrosHtml = filtrosHtml & UiFiltroCompacto("resp", "Filtrar por responsable", CgComboResponsable(Cnn, cgResp, "xForm")) & "</form>"
    sPage = sPage & UiCardAbre("Listado de incidentes", filtrosHtml) & vbCrLf

    ' --- Tabla ---
    sPage = sPage & UiTablaAbre("Listado de incidentes. Se puede ordenar por t&iacute;tulo o por cliente.") & vbCrLf
    sPage = sPage & "<thead><tr>" & _
        UiTh("<span class='ui-sr-only'>Acciones</span>", "ui-table__col--actions") & _
        UiTh("ID", "") & _
        UiThOrden("T&iacute;tulo", "incidentes.asp?ord=titulo" & qsFiltros, ord = "titulo") & _
        UiThOrden("Cliente", "incidentes.asp?ord=cliente" & qsFiltros, ord = "cliente") & _
        UiTh("Prioridad", "") & _
        UiTh("Aplicativo", "") & _
        UiTh("Desarrollo", "") & _
        UiTh("Funcional", "") & _
        UiTh("Ingreso", "") & _
        "</tr></thead><tbody>" & vbCrLf

    Dim PosReg, sFec, idReg, urlDetalle, accFila, irDetalle, marcaConsulta
    PosReg = 0

    Do While (Not rsListado.EOF) And (PosReg < CantRegMostrar)
        idReg = rsListado("id")
        sFec = UiFecha(rsListado("fecha_ingreso"))   ' Fechador(fecha + nFecini) mostraba el numero crudo (ej. 45931)
        urlDetalle = "detalles.asp?id=" & idReg & "&opc=veri&cli=" & cliente & "&tit=" & tit & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu
        irDetalle = " class='ui-table__cell--link' onclick=""document.location='" & urlDetalle & "'"""

        sForms = sForms & "<form id=""formx" & idReg & """ class=""form-oculto"" method=""post"" action=""eliminar.asp?opc=inc&id=" & idReg & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=incidentes.asp""></form>" & vbCrLf
        sForms = sForms & "<form id=""formx2" & idReg & """ class=""form-oculto"" method=""post"" action=""detalles.asp?id=" & idReg & "&opc=acti&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & """></form>" & vbCrLf

        accFila = UiBotonIcono("&times;", "Eliminar incidente " & idReg, "danger", "borrar('formx" & idReg & "');", CStr(getPermisosAtrm("CG_INCIDENTES", "ELIM")) = "1", "Sin permiso para eliminar")
        accFila = accFila & UiBotonIcono("&#9998;", "Editar incidente " & idReg, "secondary", "document.getElementById('formx2" & idReg & "').submit();", True, "")

        ' Consulta: resuelta sin responsable tecnico
        marcaConsulta = ""
        If Trim(rsListado("respon_tec") & "") = "0" Then marcaConsulta = " " & "<span class='ui-badge ui-badge--blue' title='Resuelta sin incumbencia t&eacute;cnica'>Consulta</span>"

        sPage = sPage & "<tr>" & _
            "<td><div class='ui-table__actions'>" & accFila & "</div></td>" & _
            "<td" & irDetalle & ">" & idReg & "</td>" & _
            "<td>" & UiCeldaEnlace(UiTexto(rsListado("titulo")), urlDetalle, UiTexto(rsListado("titulo"))) & marcaConsulta & "</td>" & _
            "<td" & irDetalle & ">" & UiTexto(rsListado("cliente")) & "</td>" & _
            "<td" & irDetalle & ">" & UiBadgePrioridad(rsListado("prioridad")) & "</td>" & _
            "<td" & irDetalle & ">" & UiTexto(rsListado("aplicativo")) & "</td>" & _
            "<td" & irDetalle & ">" & UiBadgeEstado(rsListado("estado_desa")) & UiSub(UiTexto(rsListado("respon_tec_nombre"))) & "</td>" & _
            "<td" & irDetalle & ">" & UiBadgeEstado(rsListado("estado_func")) & UiSub(UiTexto(rsListado("respon_gestion_nombre"))) & "</td>" & _
            "<td" & irDetalle & "><span class='ui-table__nowrap'>" & sFec & "</span></td>" & _
            "</tr>" & vbCrLf

        PosReg = PosReg + 1
        rsListado.MoveNext
    Loop
    If PosReg = 0 Then sPage = sPage & UiFilaVacia(9, "No hay incidentes con estos filtros.") & vbCrLf

    ' "Siguiente" solo si quedan registros despues de esta pagina (antes se evaluaba antes de recorrerla)
    Dim hayMas, pagAnt, pagSig
    hayMas = Not rsListado.EOF
    rsListado.Close

    If CantRegMover <> 0 Then
        sForms = sForms & "<form id=""formx22"" class=""form-oculto"" method=""post"" action=""incidentes.asp?opc=listar&cantregmover=" & (CantRegMover - CantRegMostrar) & "&ord=" & ord & qsFiltros & """></form>" & vbCrLf
        pagAnt = UiBoton("Anterior", "secondary", "document.getElementById('formx22').submit();", True, "")
    Else
        pagAnt = UiBoton("Anterior", "secondary", "", False, "Primera p&aacute;gina")
    End If
    If hayMas Then
        sForms = sForms & "<form id=""formx23"" class=""form-oculto"" method=""post"" action=""incidentes.asp?opc=listar&cantregmover=" & (CantRegMover + CantRegMostrar) & "&ord=" & ord & qsFiltros & """></form>" & vbCrLf
        pagSig = UiBoton("Siguiente", "secondary", "document.getElementById('formx23').submit();", True, "")
    Else
        pagSig = UiBoton("Siguiente", "secondary", "", False, "&Uacute;ltima p&aacute;gina")
    End If

    sPage = sPage & UiTablaCierra() & vbCrLf
    sPage = sPage & "<p class='ui-table__nota'><span class='ui-badge ui-badge--blue'>Consulta</span> Resuelta sin incumbencia t&eacute;cnica.</p>" & vbCrLf
    sPage = sPage & UiPaginacion(CantRegMover + 1, CantRegMover + PosReg, pagAnt, pagSig) & vbCrLf
    sPage = sPage & UiCardCierra() & vbCrLf
    sPage = sPage & sForms
    sPage = sPage & UiPantallaCierra() & vbCrLf

ElseIf opc = "consi" Then
    ' --- VISTA 2: CONSULTA DE INCIDENTES ---
    sel = Request.QueryString("sel")
    XXX = SYSLOGON(Cnn, "CONSULTA_INCIDENTES")

    cliente = Request.Form("cli")
    If cliente & "" = "" Then cliente = Request.QueryString("cli")
    If Not IsNumeric(cliente) Then cliente = 0

    descri = Request.Form("desc")
    If descri & "" = "" Then descri = Request.QueryString("desc")

    estado = Request.Form("est")
    If estado & "" = "" Then estado = Request.QueryString("est")
    If Not IsNumeric(estado) Then estado = 0

    aplic = Request.Form("apli")
    If aplic & "" = "" Then aplic = Request.QueryString("apli")
    If Not IsNumeric(aplic) Then aplic = 0

    fecha = Request.Form("fec")
    If fecha & "" = "" Then fecha = Request.QueryString("fec")
    Dim fechaTxtC
    fechaTxtC = fecha & ""   ' fecha tal como se escribio (para reenviar los filtros al paginar)
    If IsDate(fecha) Then
        fecha = CLng(CDate(fecha)) - nFecini
    Else
        fecha = Null
    End If

    ' Consulta de datos mediante SP
    Dim rsConsulta
    Set rsConsulta = ConsultarIncidentes(Cnn, cliente, descri, estado, aplic, fecha)

    ' --- Barra superior ---
    Dim accCons
    accCons = "<form id=""formx10"" class=""form-oculto"" method=""post"" action=""consultas.asp?sel=" & sel & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & """></form>"
    accCons = accCons & UiBoton("Volver", "secondary", "document.getElementById('formx10').submit();", True, "")
    sPage = sPage & UiPantallaAbre("Consulta de incidentes", accCons, "wide") & vbCrLf
    sPage = sPage & UiCardAbre("Resultados", "") & vbCrLf
    sPage = sPage & UiTablaAbre("Resultados de la consulta de incidentes") & vbCrLf
    sPage = sPage & "<thead><tr>" & _
        UiTh("ID", "") & _
        UiTh("T&iacute;tulo", "") & _
        UiTh("Cliente", "") & _
        UiTh("Desarrollo", "") & _
        UiTh("Funcional", "") & _
        UiTh("Ingreso", "") & _
        UiTh("<span class='ui-sr-only'>Acciones</span>", "ui-table__col--actions") & _
        "</tr></thead><tbody>" & vbCrLf

    ' [CONSULTA] Paginado de a 10 (igual que proyectos y APL). La pagina corta las filas; los filtros se reenvian con un formulario oculto.
    Dim idCons, cantCons, puedeConsultar, cantRegMoverC, cantRegMostrarC, idxPag, hastaC, botonAntC, botonSigC, urlConsC
    cantRegMostrarC = 10
    cantRegMoverC = Request.QueryString("cantregmover")
    If Not IsNumeric(cantRegMoverC & "") Or Trim(cantRegMoverC & "") = "" Then cantRegMoverC = 0
    cantRegMoverC = CLng(cantRegMoverC)
    If cantRegMoverC < 0 Then cantRegMoverC = 0
    idxPag = 0
    urlConsC = "incidentes.asp?opc=consi&sel=" & sel & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu
    sForms = sForms & "<form id=""formFiltros"" class=""form-oculto"" method=""post"" action=""" & urlConsC & """>" & _
        "<input type=""hidden"" name=""cli"" value=""" & Server.HTMLEncode(cliente & "") & """>" & _
        "<input type=""hidden"" name=""desc"" value=""" & Server.HTMLEncode(descri & "") & """>" & _
        "<input type=""hidden"" name=""est"" value=""" & Server.HTMLEncode(estado & "") & """>" & _
        "<input type=""hidden"" name=""apli"" value=""" & Server.HTMLEncode(aplic & "") & """>" & _
        "<input type=""hidden"" name=""fec"" value=""" & Server.HTMLEncode(fechaTxtC) & """></form>" & vbCrLf
    cantCons = 0
    puedeConsultar = (getPermisosAtrm("CG_INCIDENTES", "CONS") = 1)
    Do While Not rsConsulta.EOF
      If cantCons >= cantRegMoverC And idxPag < cantRegMostrarC Then
        idCons = rsConsulta("id")
        ' [UI] Ver detalle abre la vista de consulta (solo lectura); los filtros viajan para que Volver regrese a estos resultados
        '[UI-ANT] sForms = sForms & "<form id=""xform" & idCons & """ class=""form-oculto"" method=""post"" action=""detalles.asp?id=" & idCons & "&opc=veri&tipo=cons""></form>" & vbCrLf
        sForms = sForms & "<form id=""xform" & idCons & """ class=""form-oculto"" method=""post"" action=""vistaconsulta.asp?tipo=I&id=" & idCons & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&" & Request.Form & """></form>" & vbCrLf

        sPage = sPage & "<tr>" & _
            "<td>" & idCons & "</td>" & _
            "<td><span class='ui-table__titulo'>" & UiTexto(rsConsulta("titulo")) & "</span></td>" & _
            "<td>" & UiTexto(rsConsulta("cliente")) & "</td>" & _
            "<td>" & UiBadgeEstado(rsConsulta("estado_desa")) & "</td>" & _
            "<td>" & UiBadgeEstado(rsConsulta("estado_func")) & "</td>" & _
            "<td><span class='ui-table__nowrap'>" & UiFecha(rsConsulta("fecha_ingreso")) & "</span></td>" & _
            "<td>" & UiBoton("Ver detalle", "secondary", "document.getElementById('xform" & idCons & "').submit();", puedeConsultar, "Sin permiso para consultar.") & "</td>" & _
            "</tr>" & vbCrLf

        idxPag = idxPag + 1
      End If
        cantCons = cantCons + 1
        rsConsulta.MoveNext
    Loop
    rsConsulta.Close
    If cantCons = 0 Then sPage = sPage & UiFilaVacia(7, "No hay incidentes que coincidan con la consulta.") & vbCrLf

    sPage = sPage & UiTablaCierra() & vbCrLf
    hastaC = cantRegMoverC + idxPag
    If cantRegMoverC > 0 Then
        botonAntC = UiBoton("Anterior", "secondary", "irConsulta('" & urlConsC & "&cantregmover=" & (cantRegMoverC - cantRegMostrarC) & "');", True, "")
    Else
        botonAntC = UiBoton("Anterior", "secondary", "", False, "Primera p&aacute;gina")
    End If
    If hastaC < cantCons Then
        botonSigC = UiBoton("Siguiente", "secondary", "irConsulta('" & urlConsC & "&cantregmover=" & (cantRegMoverC + cantRegMostrarC) & "');", True, "")
    Else
        botonSigC = UiBoton("Siguiente", "secondary", "", False, "&Uacute;ltima p&aacute;gina")
    End If
    sPage = sPage & UiPaginacionDe(cantRegMoverC + 1, hastaC, cantCons, botonAntC, botonSigC) & vbCrLf
    sPage = sPage & UiCardCierra() & vbCrLf
    sPage = sPage & sForms
    sPage = sPage & UiPantallaCierra() & vbCrLf
End If

' --- SCRIPT Y CIERRE DE DOCUMENTO ---
sPage = sPage & "<script type=""text/javascript"">" & vbCrLf
sPage = sPage & "    function irConsulta(url){ var f = document.getElementById('formFiltros'); if (!f) return false; f.action = url; f.submit(); return false; }" & vbCrLf
sPage = sPage & "    function borrar(formu){ " & vbCrLf
sPage = sPage & "        var r = confirm('\u00bfDesea eliminar el incidente?'); " & vbCrLf
sPage = sPage & "        if (r){ " & vbCrLf
sPage = sPage & "            document.getElementById(formu).submit(); " & vbCrLf
sPage = sPage & "        } else { " & vbCrLf
sPage = sPage & "            return false;" & vbCrLf
sPage = sPage & "        }" & vbCrLf
sPage = sPage & "    }" & vbCrLf
sPage = sPage & "</script>" & vbCrLf
sPage = sPage & "</body>" & vbCrLf
sPage = sPage & "</html>"

Response.Write sPage

Cnn.Close
Set Cnn = Nothing
%>
