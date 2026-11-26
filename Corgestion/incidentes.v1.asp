<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<%
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

Function ObtenerListadoIncidentes(cn, idCli, idApli, titulo, orden)
    Dim cmd
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandText = "sp_cg_Incidentes_Listar"
    cmd.CommandType = 4 ' adCmdStoredProc

    cmd.Parameters.Append cmd.CreateParameter("@IdCliente", 3, 1, , CLng(idCli))
    cmd.Parameters.Append cmd.CreateParameter("@IdAplicativo", 3, 1, , CLng(idApli))
    cmd.Parameters.Append cmd.CreateParameter("@Titulo", 200, 1, 255, CStr(titulo & ""))
    cmd.Parameters.Append cmd.CreateParameter("@Orden", 200, 1, 20, CStr(orden & ""))

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

' =========================================================================
' HELPERS DE PRESENTACION (arman fragmentos HTML como string)
' =========================================================================

Function ArmarBadgeEstado(estadoVal, nombreResp)
    Dim s, respSeguro
    respSeguro = Server.HTMLEncode(nombreResp & "")
    Select Case estadoVal
        Case "PENDIENTE": s = "<span class='badge badge-media ui-badge ui-badge--warning' aria-label='Pendiente'>PEND</span> (" & respSeguro & ")"
        Case "OK": s = "<span class='badge badge-ok ui-badge ui-badge--ok' aria-label='OK'>OK</span> (" & respSeguro & ")"
        Case "BAJA": s = "<span class='badge badge-peligro ui-badge ui-badge--danger' aria-label='Baja'>BAJA</span> (" & respSeguro & ")"
        Case "TESTING EN CLIENTE": s = "<span class='badge badge-media ui-badge ui-badge--warning' aria-label='Testing en cliente'>TEST</span> (" & respSeguro & ")"
        Case Else: s = "-------------"
    End Select
    ArmarBadgeEstado = s
End Function

Function ArmarBadgeEstadoConsulta(estadoVal)
    Dim s
    Select Case estadoVal
        Case "PENDIENTE": s = "<span class='badge badge-media ui-badge ui-badge--warning' aria-label='Pendiente'>PENDIENTE</span>"
        Case "OK": s = "<span class='badge badge-ok ui-badge ui-badge--ok' aria-label='OK'>OK</span>"
        Case "BAJA": s = "<span class='badge badge-peligro ui-badge ui-badge--danger' aria-label='Baja'>BAJA</span>"
        Case Else: s = ""
    End Select
    ArmarBadgeEstadoConsulta = s
End Function

Function ArmarBadgePrioridad(prioridadVal)
    Dim s, prioSeguro
    prioSeguro = Server.HTMLEncode(prioridadVal & "")
    If prioridadVal = "ALTA" Then
        s = "<span class='badge badge-peligro ui-badge ui-badge--danger' aria-label='Prioridad alta'>ALTA</span>"
    ElseIf prioridadVal = "MEDIA" Then
        s = "<span class='badge badge-media ui-badge ui-badge--warning' aria-label='Prioridad media'>MEDIA</span>"
    Else
        s = "<span class='badge badge-pendiente ui-badge ui-badge--pending' aria-label='Prioridad " & prioSeguro & "'>" & prioSeguro & "</span>"
    End If
    ArmarBadgePrioridad = s
End Function

' =========================================================================
' VISTA / PRESENTACION (todo armado en strings, un solo Response.Write al final)
' =========================================================================

sPage = ""
sPage = sPage & headhtml()

sPage = sPage & "<body id=""body"" name=""body"" style=""height:100%;width:100%;"">" & vbCrLf
sPage = sPage & "<div id=""wrapper"">" & vbCrLf
sPage = sPage & cabecera()
sPage = sPage & "</div>" & vbCrLf
sPage = sPage & getMenuParam()

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
    Set rsListado = ObtenerListadoIncidentes(Cnn, cliente, aplicativo, tit, ord)

    If Not rsListado.EOF Then rsListado.Move(CantRegMover)

    sPage = sPage & "<div class=""inner cg-page ui-main"">" & vbCrLf
    sPage = sPage & "    <div class=""titulo-seccion cg-titlebar ui-page-header"">" & vbCrLf
    sPage = sPage & "        <div class=""page-toolbar cg-titlebar-row ui-toolbar"">" & vbCrLf
    sPage = sPage & "                <div class=""page-toolbar-title cg-titlebar-title""><h2>INCIDENTES</h2></div>" & vbCrLf
    sPage = sPage & "                <div class=""page-toolbar-filters cg-titlebar-filters ui-toolbar__filters"">" & vbCrLf
    sPage = sPage & "                    <form id=""xForm"" method=""post"" action=""incidentes.asp?ord=" & ord & "&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & """>" & vbCrLf

    ' Combo Clientes
    sPage = sPage & "                        <select class=""ui-select"" onchange=""document.getElementById('xForm').submit();"" id=""cli"" name=""cli"">" & vbCrLf
    sPage = sPage & "                            <option value=""0"">CLIENTE: TODOS</option>" & vbCrLf
    Dim rsCli, selCli
    Set rsCli = GetListaComboExtended(Cnn, "clientes")
    Do While Not rsCli.EOF
        If CStr(rsCli("id")) = CStr(cliente) Then selCli = "selected" Else selCli = ""
        sPage = sPage & "                            <option value=""" & rsCli("id") & """ " & selCli & ">" & Server.HTMLEncode(rsCli("descripcion") & "") & "</option>" & vbCrLf
        rsCli.MoveNext
    Loop
    rsCli.Close
    sPage = sPage & "                        </select>" & vbCrLf

    ' Combo Aplicativos
    sPage = sPage & "                        <select class=""ui-select"" onchange=""document.getElementById('xForm').submit();"" id=""apli"" name=""apli"">" & vbCrLf
    sPage = sPage & "                            <option value=""0"">APLICATIVO: TODOS</option>" & vbCrLf
    Dim rsApli, selApli
    Set rsApli = GetListaComboExtended(Cnn, "aplicativos")
    Do While Not rsApli.EOF
        If CStr(rsApli("id")) = CStr(aplicativo) Then selApli = "selected" Else selApli = ""
        sPage = sPage & "                            <option value=""" & rsApli("id") & """ " & selApli & ">" & Server.HTMLEncode(rsApli("descripcion") & "") & "</option>" & vbCrLf
        rsApli.MoveNext
    Loop
    rsApli.Close
    sPage = sPage & "                        </select>" & vbCrLf
    sPage = sPage & "                    </form>" & vbCrLf
    sPage = sPage & "                </div>" & vbCrLf

    ' Acciones (alta)
    If Session("FUNC") = "FUNC" Or Session("FUNC") = "TEC" Then
        sPage = sPage & "                <div class=""page-toolbar-action cg-titlebar-actions ui-toolbar__actions"">" & vbCrLf
        sPage = sPage & "                    <form id=""formx4"" class=""form-oculto"" method=""post"" action=""agregar2.asp?tipo=INC&xmenu=" & xmenu & "&xtmenu=" & xtmenu & """></form>" & vbCrLf
        If CStr(getPermisos(xmenu, xtmenu, "ALTA")) = "1" Then
            sPage = sPage & "                    <input class=""bot ui-button ui-button--primary"" type=""button"" value=""CARGAR NUEVO INCIDENTE"" onclick=""document.getElementById('formx4').submit();"">" & vbCrLf
        Else
            sPage = sPage & "                    <input class=""botsp ui-button"" type=""button"" value=""CARGAR NUEVO INCIDENTE"" disabled title=""Sin permiso para agregar."">" & vbCrLf
        End If
        sPage = sPage & "                </div>" & vbCrLf
    End If

    ' Paginacion: ATRAS
    sPage = sPage & "                <div class=""page-toolbar-action cg-titlebar-actions ui-toolbar__actions ui-pagination"">" & vbCrLf
    If CantRegMover <> 0 Then
        sPage = sPage & "                    <form id=""formx22"" class=""form-oculto"" method=""post"" action=""incidentes.asp?opc=listar&cantregmover=" & (CantRegMover - CantRegMostrar) & "&ord=" & ord & "&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & """></form>" & vbCrLf
        sPage = sPage & "                    <input class=""bot ui-button"" type=""button"" value=""ATRAS"" onclick=""document.getElementById('formx22').submit();"">" & vbCrLf
    Else
        sPage = sPage & "                    <input class=""botsp ui-button"" type=""button"" value=""ATRAS"" disabled>" & vbCrLf
    End If
    sPage = sPage & "                </div>" & vbCrLf

    ' Paginacion: SIGUIENTE
    sPage = sPage & "                <div class=""page-toolbar-action cg-titlebar-actions ui-toolbar__actions ui-pagination"">" & vbCrLf
    If Not rsListado.EOF Then
        sPage = sPage & "                    <form id=""formx23"" class=""form-oculto"" method=""post"" action=""incidentes.asp?opc=listar&cantregmover=" & (CantRegMover + CantRegMostrar) & "&ord=" & ord & "&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & """></form>" & vbCrLf
        sPage = sPage & "                    <input class=""bot ui-button"" type=""button"" value=""SIGUIENTE"" onclick=""document.getElementById('formx23').submit();"">" & vbCrLf
    Else
        sPage = sPage & "                    <input class=""botsp ui-button"" type=""button"" value=""SIGUIENTE"" disabled>" & vbCrLf
    End If
    sPage = sPage & "                </div>" & vbCrLf

    ' Volver
    sPage = sPage & "                <div class=""page-toolbar-action cg-titlebar-actions ui-toolbar__actions"">" & vbCrLf
    sPage = sPage & "                    <form id=""formx5"" class=""form-oculto"" method=""post"" action=""index.asp""></form>" & vbCrLf
    sPage = sPage & "                    <input class=""bot ui-button"" type=""button"" value=""VOLVER"" onclick=""document.getElementById('formx5').submit();"">" & vbCrLf
    sPage = sPage & "                </div>" & vbCrLf
    sPage = sPage & "        </div>" & vbCrLf
    sPage = sPage & "    </div>" & vbCrLf

    ' --- GRILLA PRINCIPAL ---
    sPage = sPage & "    <div class=""bloque-datos table-responsive ui-table-wrap"">" & vbCrLf
    sPage = sPage & "        <table class=""tabla-moderna ui-table"">" & vbCrLf
    sPage = sPage & "            <thead>" & vbCrLf
    sPage = sPage & "                <tr>" & vbCrLf
    sPage = sPage & "                    <th></th>" & vbCrLf
    sPage = sPage & "                    <th></th>" & vbCrLf
    sPage = sPage & "                    <th>ID</th>" & vbCrLf
    sPage = sPage & "                    <th><a class=""col-link"" style=""text-decoration: none;"" href=""incidentes.asp?ord=titulo&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & """>TITULO</a></th>" & vbCrLf
    sPage = sPage & "                    <th><a class=""col-link"" style=""text-decoration: none;"" href=""incidentes.asp?ord=cliente&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & """>CLIENTE</a></th>" & vbCrLf
    sPage = sPage & "                    <th>PRIORIDAD</th>" & vbCrLf
    sPage = sPage & "                    <th>APLICATIVO</th>" & vbCrLf
    sPage = sPage & "                    <th>DESARROLLO</th>" & vbCrLf
    sPage = sPage & "                    <th>FUNCIONAL</th>" & vbCrLf
    sPage = sPage & "                    <th>FEC INGRESO</th>" & vbCrLf
    sPage = sPage & "                </tr>" & vbCrLf
    sPage = sPage & "            </thead>" & vbCrLf
    sPage = sPage & "            <tbody>" & vbCrLf

    Dim PosReg, est1, est2, prio, sFec, idReg, urlDetalle, xBtnElim
    PosReg = 0

    Dim tituloSeguro, tituloTrunc, clienteSeguro, aplicativoSeguro
    Do While (Not rsListado.EOF) And (PosReg < CantRegMostrar)
        idReg = rsListado("id")
        est1 = ArmarBadgeEstado(rsListado("estado_desa"), rsListado("respon_tec_nombre"))
        est2 = ArmarBadgeEstado(rsListado("estado_func"), rsListado("respon_gestion_nombre"))
        prio = ArmarBadgePrioridad(rsListado("prioridad"))
        sFec = Fechador(rsListado("fecha_ingreso") + nFecini)
        urlDetalle = "detalles.asp?id=" & idReg & "&opc=veri&cli=" & cliente & "&tit=" & tit & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu
        tituloSeguro = Server.HTMLEncode(rsListado("titulo") & "")
        tituloTrunc = Server.HTMLEncode(UCase(Left(rsListado("titulo") & "", 40)))
        clienteSeguro = Server.HTMLEncode(rsListado("cliente") & "")
        aplicativoSeguro = Server.HTMLEncode(rsListado("aplicativo") & "")

        sPage = sPage & "                <form id=""formx" & idReg & """ class=""form-oculto"" method=""post"" action=""eliminar.asp?opc=inc&id=" & idReg & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & "&vuelta=incidentes.asp""></form>" & vbCrLf
        sPage = sPage & "                <form id=""formx2" & idReg & """ class=""form-oculto"" method=""post"" action=""detalles.asp?id=" & idReg & "&opc=acti&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & """></form>" & vbCrLf

        sPage = sPage & "                <tr>" & vbCrLf
        sPage = sPage & "                    <td style=""text-align: center;"">" & vbCrLf
        If CStr(getPermisos(xmenu, xtmenu, "ELIM")) = "1" Then
            sPage = sPage & "                        <input class=""btn-accion btn-accion-peligro ui-button ui-button--icon ui-button--danger"" type=""button"" value=""X"" onclick=""borrar('formx" & idReg & "');"">" & vbCrLf
        Else
            sPage = sPage & "                        <input class=""btn-accion btn-accion-deshabilitado ui-button ui-button--icon"" type=""button"" value=""X"" disabled title=""Sin permiso para eliminar."">" & vbCrLf
        End If
        sPage = sPage & "                    </td>" & vbCrLf
        sPage = sPage & "                    <td style=""text-align: center;"">" & vbCrLf
        sPage = sPage & "                        <input class=""btn-accion btn-accion-primario ui-button ui-button--icon ui-button--primary"" type=""button"" value=""A"" onclick=""document.getElementById('formx2" & idReg & "').submit();"">" & vbCrLf
        sPage = sPage & "                    </td>" & vbCrLf
        sPage = sPage & "                    <td onclick=""document.location='" & urlDetalle & "'"">" & idReg & "</td>" & vbCrLf
        sPage = sPage & "                    <td title=""" & tituloSeguro & """ onclick=""document.location='" & urlDetalle & "'"">" & vbCrLf
        sPage = sPage & "                        " & tituloTrunc & vbCrLf
        If Trim(rsListado("respon_tec") & "") = "0" Then
            sPage = sPage & "                        <span class=""ui-badge ui-badge--pending"" aria-label=""Consulta""><b>(C)</b></span>" & vbCrLf
        End If
        sPage = sPage & "                    </td>" & vbCrLf
        sPage = sPage & "                    <td onclick=""document.location='" & urlDetalle & "'"">" & clienteSeguro & "</td>" & vbCrLf
        sPage = sPage & "                    <td onclick=""document.location='" & urlDetalle & "'"">" & prio & "</td>" & vbCrLf
        sPage = sPage & "                    <td onclick=""document.location='" & urlDetalle & "'"">" & aplicativoSeguro & "</td>" & vbCrLf
        sPage = sPage & "                    <td onclick=""document.location='" & urlDetalle & "'"">" & est1 & "</td>" & vbCrLf
        sPage = sPage & "                    <td onclick=""document.location='" & urlDetalle & "'"">" & est2 & "</td>" & vbCrLf
        sPage = sPage & "                    <td onclick=""document.location='" & urlDetalle & "'"">" & sFec & "</td>" & vbCrLf
        sPage = sPage & "                </tr>" & vbCrLf

        PosReg = PosReg + 1
        rsListado.MoveNext
    Loop
    rsListado.Close

    sPage = sPage & "            </tbody>" & vbCrLf
    sPage = sPage & "        </table>" & vbCrLf
    sPage = sPage & "    </div>" & vbCrLf

    sPage = sPage & "    <div class=""cg-legend table-legend ui-legend"">" & vbCrLf
    sPage = sPage & "        <span><span class=""table-legend-mark""><b>*</b></span><span class=""table-legend-consult""><b>(C)</b></span></span>" & vbCrLf
    sPage = sPage & "        <span><span class=""table-legend-mark"">&nbsp;&nbsp;(Consulta) resuelta sin incumbencia tecnica.</span></span>" & vbCrLf
    sPage = sPage & "    </div>" & vbCrLf
    sPage = sPage & "</div>" & vbCrLf

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
    If IsDate(fecha) Then
        fecha = CLng(CDate(fecha)) - nFecini
    Else
        fecha = Null
    End If

    ' Consulta de datos mediante SP
    Dim rsConsulta
    Set rsConsulta = ConsultarIncidentes(Cnn, cliente, descri, estado, aplic, fecha)

    sPage = sPage & "<div class=""inner cg-page ui-main"">" & vbCrLf
    sPage = sPage & "    <div class=""titulo-seccion cg-titlebar ui-page-header"">" & vbCrLf
    sPage = sPage & "        <div class=""page-toolbar cg-titlebar-row ui-toolbar"">" & vbCrLf
    sPage = sPage & "            <div class=""page-toolbar-title cg-titlebar-title""><h2>CONSULTA INCIDENTES</h2></div>" & vbCrLf
    sPage = sPage & "            <div class=""page-toolbar-action cg-titlebar-actions ui-toolbar__actions"">" & vbCrLf
    sPage = sPage & "                <form id=""formx10"" class=""form-oculto"" method=""post"" action=""consultas.asp?sel=" & sel & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu & """></form>" & vbCrLf
    sPage = sPage & "                <input class=""bot ui-button"" type=""button"" value=""VOLVER"" onclick=""document.getElementById('formx10').submit();"">" & vbCrLf
    sPage = sPage & "            </div>" & vbCrLf
    sPage = sPage & "        </div>" & vbCrLf
    sPage = sPage & "    </div>" & vbCrLf

    sPage = sPage & "    <div class=""bloque-datos table-responsive ui-table-wrap"">" & vbCrLf
    sPage = sPage & "        <table class=""tabla-moderna ui-table"" style=""width: 100%;"">" & vbCrLf
    sPage = sPage & "            <thead>" & vbCrLf
    sPage = sPage & "                <tr>" & vbCrLf
    sPage = sPage & "                    <th style=""width: 5%; text-align: center;"">ID</th>" & vbCrLf
    sPage = sPage & "                    <th style=""width: 40%;"">TITULO</th>" & vbCrLf
    sPage = sPage & "                    <th style=""width: 40px; text-align: center;"">&nbsp;</th>" & vbCrLf
    sPage = sPage & "                    <th style=""width: 18%;"">CLIENTE</th>" & vbCrLf
    sPage = sPage & "                    <th style=""width: 10.5%; text-align: center;"">DESARROLLO</th>" & vbCrLf
    sPage = sPage & "                    <th style=""width: 10.5%; text-align: center;"">FUNCIONAL</th>" & vbCrLf
    sPage = sPage & "                    <th style=""width: 15%; text-align: center;"">FEC INGRESO</th>" & vbCrLf
    sPage = sPage & "                </tr>" & vbCrLf
    sPage = sPage & "            </thead>" & vbCrLf
    sPage = sPage & "            <tbody>" & vbCrLf

    Dim estD, estF, idCons, tituloConsSeguro, clienteConsSeguro
    Do While Not rsConsulta.EOF
        idCons = rsConsulta("id")
        estD = ArmarBadgeEstadoConsulta(rsConsulta("estado_desa"))
        estF = ArmarBadgeEstadoConsulta(rsConsulta("estado_func"))
        tituloConsSeguro = Server.HTMLEncode(UCase(Mid(rsConsulta("titulo") & "", 1, 58)))
        clienteConsSeguro = Server.HTMLEncode(rsConsulta("cliente") & "")

        sPage = sPage & "                <form id=""xform" & idCons & """ class=""form-oculto"" method=""post"" action=""detalles.asp?id=" & idCons & "&opc=veri&tipo=cons""></form>" & vbCrLf
        sPage = sPage & "                <tr>" & vbCrLf
        sPage = sPage & "                    <td style=""text-align: center; color: #475569; font-weight: 500;"">" & idCons & "</td>" & vbCrLf
        sPage = sPage & "                    <td>" & tituloConsSeguro & "</td>" & vbCrLf
        sPage = sPage & "                    <td style=""text-align: center; vertical-align: middle;"">" & vbCrLf
        If getPermisos(9, "INICIO", "CONS") = 1 Then
            sPage = sPage & "                        <input class=""btn-accion btn-accion-primario ui-button ui-button--icon ui-button--primary"" type=""button"" value=""D"" onclick=""document.getElementById('xform" & idCons & "').submit();"">" & vbCrLf
        Else
            sPage = sPage & "                        <input class=""btn-accion btn-accion-deshabilitado ui-button ui-button--icon"" type=""button"" value=""D"" disabled title=""Sin permiso para consultar."">" & vbCrLf
        End If
        sPage = sPage & "                    </td>" & vbCrLf
        sPage = sPage & "                    <td style=""padding-left: 1rem;"">" & clienteConsSeguro & "</td>" & vbCrLf
        sPage = sPage & "                    <td style=""text-align: center;"">" & estD & "</td>" & vbCrLf
        sPage = sPage & "                    <td style=""text-align: center;"">" & estF & "</td>" & vbCrLf
        sPage = sPage & "                    <td style=""text-align: center;"">" & Fechador(rsConsulta("fecha_ingreso") + nFecini) & "</td>" & vbCrLf
        sPage = sPage & "                </tr>" & vbCrLf

        rsConsulta.MoveNext
    Loop
    rsConsulta.Close

    sPage = sPage & "            </tbody>" & vbCrLf
    sPage = sPage & "        </table>" & vbCrLf
    sPage = sPage & "    </div>" & vbCrLf
    sPage = sPage & "</div>" & vbCrLf
End If

' --- SCRIPT Y CIERRE DE DOCUMENTO ---
sPage = sPage & "<script type=""text/javascript"">" & vbCrLf
sPage = sPage & "    function borrar(formu){ " & vbCrLf
sPage = sPage & "        var r = confirm('�Desea eliminar el incidente?'); " & vbCrLf
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
