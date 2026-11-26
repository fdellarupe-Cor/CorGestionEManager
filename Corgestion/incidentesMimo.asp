<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<%
'Option Explicit

' --- CONEXION A BASE DE DATOS ---
Dim Cnn
Set Cnn = Server.CreateObject("ADODB.Connection")
On Error Resume Next
Cnn.Open scon1
If Err.Number <> 0 Then
	Response.Write "<div class='error-panel'>"
	Response.Write "<b>ERROR DE CONEXION A BASE DE DATOS</b><br>"
	Response.Write "Codigo: " & Err.Number & "<br>"
	Response.Write "Detalle: " & Server.HTMLEncode(Err.Description) & "</div>"
	Response.End
End If
On Error Goto 0

' --- DECLARACION OBLIGATORIA DE TODAS LAS VARIABLES ---
Dim opc, xmenu, xtmenu, ord, cliente, tit, aplicativo, aplic, CantRegMover, CantRegMostrar
Dim sel, descri, estado, fecha, XXX, nFecini

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
    cmd.CommandText = "sp_Incidentes_Listar"
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
    cmd.CommandText = "sp_Incidentes_Consultar"
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

Function ObtenerCatalogo(cn, nombreTabla)
    Dim cmd
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandText = "sp_Obtener_Catalogo"
    cmd.CommandType = 4
    
    cmd.Parameters.Append cmd.CreateParameter("@Tabla", 200, 1, 30, CStr(nombreTabla))
    Set ObtenerCatalogo = cmd.Execute()
End Function

' =========================================================================
' VISTA / PRESENTACION (HTML) - CSS OPTIMIZADO
' =========================================================================

Response.Write headhtml()
%>
<body id="body" name="body">

<!-- HEADER - Usa div.flex en vez de table -->
<header class="app-header">
    <div class="app-header__inner">
        <a href="index.asp" class="app-brand">
            <img src="images/iso_cor.png" class="app-brand__logo" alt="Cor Gestion" style="width:6rem;">
            <span class="app-brand__text">Cor Gestion | Sistema Administrativo</span>
        </a>
        <nav class="app-nav">
            <a href="#menu" class="btn-hamburguesa" aria-label="Abrir menu">
                <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2">
                    <path stroke-linecap="round" stroke-linejoin="round" d="M4 6h16M4 12h16M4 18h16" />
                </svg>
            </a>
        </nav>
    </div>
</header>

<%= getMenuParam() %>

<%
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
%>

    <main class="page-shell">

        <!-- TITULO Y FILTROS - Usa .toolbar en vez de table -->
        <section class="page-header">
            <h1 class="page-title">INCIDENTES</h1>

            <form id="xForm" class="toolbar__filters" method="post" action="incidentes.asp?ord=<%=ord%>&cli=<%=cliente%>&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>">
                <label class="field">
                    <span class="field__label">Cliente</span>
                    <select class="field__control" onchange="document.getElementById('xForm').submit();" id="cli" name="cli">
                        <option value="0">TODOS</option>
                        <%
                        Dim rsCli, selCli
                        Set rsCli = ObtenerCatalogo(Cnn, "clientes")
                        Do While Not rsCli.EOF
                            If CStr(rsCli("id")) = CStr(cliente) Then selCli = "selected" Else selCli = ""
                        %>
                            <option value="<%=rsCli("id")%>" <%=selCli%>><%=rsCli("descripcion")%></option>
                        <%
                            rsCli.MoveNext
                        Loop
                        rsCli.Close
                        %>
                    </select>
                </label>

                <label class="field">
                    <span class="field__label">Aplicativo</span>
                    <select class="field__control" onchange="document.getElementById('xForm').submit();" id="apli" name="apli">
                        <option value="0">TODOS</option>
                        <%
                        Dim rsApli, selApli
                        Set rsApli = ObtenerCatalogo(Cnn, "aplicativos")
                        Do While Not rsApli.EOF
                            If CStr(rsApli("id")) = CStr(aplicativo) Then selApli = "selected" Else selApli = ""
                        %>
                            <option value="<%=rsApli("id")%>" <%=selApli%>><%=rsApli("descripcion")%></option>
                        <%
                            rsApli.MoveNext
                        Loop
                        rsApli.Close
                        %>
                    </select>
                </label>
            </form>

            <div class="toolbar__actions">
                <!-- Boton nuevo incidente -->
                <% If Session("FUNC") = "FUNC" Or Session("FUNC") = "TEC" Then %>
                    <form id="formx4" class="form-oculto" method="post" action="agregar2.asp?tipo=INC&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>"></form>
                    <% If CStr(getPermisos(xmenu, xtmenu, "ALTA")) = "1" Then %>
                        <input class="bot" type="button" value="CARGAR NUEVO INCIDENTE" onclick="document.getElementById('formx4').submit();">
                    <% Else %>
                        <input class="botsp" type="button" value="CARGAR NUEVO INCIDENTE" disabled>
                    <% End If %>
                <% End If %>

                <!-- Paginacion -->
                <% If CantRegMover <> 0 Then %>
                    <form id="formx22" class="form-oculto" method="post" action="incidentes.asp?opc=listar&cantregmover=<%=(CantRegMover - CantRegMostrar)%>&ord=<%=ord%>&cli=<%=cliente%>&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>"></form>
                    <input class="bot" type="button" value="ATRAS" onclick="document.getElementById('formx22').submit();">
                <% Else %>
                    <input class="botsp" type="button" value="ATRAS" disabled>
                <% End If %>

                <% If Not rsListado.EOF Then %>
                    <form id="formx23" class="form-oculto" method="post" action="incidentes.asp?opc=listar&cantregmover=<%=(CantRegMover + CantRegMostrar)%>&ord=<%=ord%>&cli=<%=cliente%>&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>"></form>
                    <input class="bot" type="button" value="SIGUIENTE" onclick="document.getElementById('formx23').submit();">
                <% Else %>
                    <input class="botsp" type="button" value="SIGUIENTE" disabled>
                <% End If %>

                <!-- Volver -->
                <form id="formx5" class="form-oculto" method="post" action="index.asp"></form>
                <input class="bot" type="button" value="VOLVER" onclick="document.getElementById('formx5').submit();">
            </div>
        </section>

        <!-- GRILLA PRINCIPAL -->
        <section class="data-panel">
            <div class="data-table-wrap">
                <table class="data-table">
                    <thead>
                        <tr>
                            <th class="col-action"></th>
                            <th class="col-action"></th>
                            <th class="col-id">ID</th>
                            <th class="col-link"><a class="col-ordenable" href="incidentes.asp?ord=titulo&cli=<%=cliente%>&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>">TITULO</a></th>
                            <th><a class="col-ordenable" href="incidentes.asp?ord=cliente&cli=<%=cliente%>&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>">CLIENTE</a></th>
                            <th>PRIORIDAD</th>
                            <th>APLICATIVO</th>
                            <th>DESARROLLO</th>
                            <th>FUNCIONAL</th>
                            <th>FEC INGRESO</th>
                        </tr>
                    </thead>
                    <tbody>
                    <%
                    Dim PosReg, est1, est2, prio, sFec
                    PosReg = 0

                    Do While (Not rsListado.EOF) And (PosReg < CantRegMostrar)
                        Select Case rsListado("estado_desa")
                            Case "PENDIENTE": est1 = "<span class='status-badge status-badge--pending'>PEND</span> (" & rsListado("respon_tec_nombre") & ")"
                            Case "OK": est1 = "<span class='status-badge status-badge--ok'>OK</span> (" & rsListado("respon_tec_nombre") & ")"
                            Case "BAJA": est1 = "<span class='status-badge status-badge--danger'>BAJA</span> (" & rsListado("respon_tec_nombre") & ")"
                            Case "TESTING EN CLIENTE": est1 = "<span class='status-badge status-badge--pending'>TEST</span> (" & rsListado("respon_tec_nombre") & ")"
                            Case Else: est1 = "-------------"
                        End Select

                        Select Case rsListado("estado_func")
                            Case "PENDIENTE": est2 = "<span class='status-badge status-badge--pending'>PEND</span> (" & rsListado("respon_gestion_nombre") & ")"
                            Case "OK": est2 = "<span class='status-badge status-badge--ok'>OK</span> (" & rsListado("respon_gestion_nombre") & ")"
                            Case "BAJA": est2 = "<span class='status-badge status-badge--danger'>BAJA</span> (" & rsListado("respon_gestion_nombre") & ")"
                            Case "TESTING EN CLIENTE": est2 = "<span class='status-badge status-badge--pending'>TEST</span> (" & rsListado("respon_gestion_nombre") & ")"
                            Case Else: est2 = "-------------"
                        End Select

                        If rsListado("prioridad") = "ALTA" Then
                            prio = "<span class='status-badge status-badge--danger'>ALTA</span>"
                        ElseIf rsListado("prioridad") = "MEDIA" Then
                            prio = "<span class='status-badge status-badge--pending'>MEDIA</span>"
                        Else
                            prio = "<span class='status-badge status-badge--muted'>" & rsListado("prioridad") & "</span>"
                        End If

                        sFec = Fechador(rsListado("fecha_ingreso") + nFecini)
                    %>
                        <form id="formx<%=rsListado("id")%>" class="form-oculto" method="post" action="eliminar.asp?opc=inc&id=<%=rsListado("id")%>&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>&vuelta=incidentes.asp"></form>
                        <form id="formx2<%=rsListado("id")%>" class="form-oculto" method="post" action="detalles.asp?id=<%=rsListado("id")%>&opc=acti&cli=<%=cliente%>&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>"></form>

                        <tr>
                            <td class="col-action">
                                <% If CStr(getPermisos(xmenu, xtmenu, "ELIM")) = "1" Then %>
                                    <a class="action-button action-button--danger" href="javascript:borrar('formx<%=rsListado("id")%>');" title="Eliminar">X</a>
                                <% Else %>
                                    <span class="action-button action-button--muted" title="Sin permiso para eliminar">X</span>
                                <% End If %>
                            </td>
                            <td class="col-action">
                                <a class="action-button action-button--primary" href="javascript:document.getElementById('formx2<%=rsListado("id")%>').submit();" title="Ver actividad">A</a>
                            </td>
                            <td class="col-id td-clickable" onclick="document.location='detalles.asp?id=<%=rsListado("id")%>&opc=veri&cli=<%=cliente%>&tit=<%=tit%>&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>'"><%=rsListado("id")%></td>
                            <td class="td-clickable" title="<%=Server.HTMLEncode(rsListado("titulo"))%>" onclick="document.location='detalles.asp?id=<%=rsListado("id")%>&opc=veri&cli=<%=cliente%>&tit=<%=tit%>&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>'">
                                <%=Left(UCase(rsListado("titulo")), 40)%>
                                <% If rsListado("respon_tec") = 0 Then %> <span class="tag-consulta">(C)</span><% End If %>
                            </td>
                            <td class="td-clickable" onclick="document.location='detalles.asp?id=<%=rsListado("id")%>&opc=veri&cli=<%=cliente%>&tit=<%=tit%>&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>'"><%=rsListado("cliente")%></td>
                            <td class="td-clickable" onclick="document.location='detalles.asp?id=<%=rsListado("id")%>&opc=veri&cli=<%=cliente%>&tit=<%=tit%>&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>'"><%=prio%></td>
                            <td class="td-clickable" onclick="document.location='detalles.asp?id=<%=rsListado("id")%>&opc=veri&cli=<%=cliente%>&tit=<%=tit%>&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>'"><%=rsListado("aplicativo")%></td>
                            <td class="td-clickable" onclick="document.location='detalles.asp?id=<%=rsListado("id")%>&opc=veri&cli=<%=cliente%>&tit=<%=tit%>&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>'"><%=est1%></td>
                            <td class="td-clickable" onclick="document.location='detalles.asp?id=<%=rsListado("id")%>&opc=veri&cli=<%=cliente%>&tit=<%=tit%>&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>'"><%=est2%></td>
                            <td class="td-clickable" onclick="document.location='detalles.asp?id=<%=rsListado("id")%>&opc=veri&cli=<%=cliente%>&tit=<%=tit%>&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>'"><%=sFec%></td>
                        </tr>
                    <%
                        PosReg = PosReg + 1
                        rsListado.MoveNext
                    Loop
                    rsListado.Close
                    %>
                    </tbody>
                </table>
            </div>
        </section>

        <!-- LEYENDA - Usa div en vez de table -->
        <div class="legend">
            <span class="legend__symbol">*</span>
            <span class="legend__symbol legend__symbol--blue">(C)</span>
            <span class="legend__text">(Consulta) resuelta sin incumbencia tecnica.</span>
        </div>

    </main>

<% ElseIf opc = "consi" Then 
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
%>
    <main class="page-shell">

        <section class="page-header">
            <h1 class="page-title">CONSULTA INCIDENTES</h1>
            <div class="toolbar__actions">
                <form id="formx10" class="form-oculto" method="post" action="consultas.asp?sel=<%=sel%>&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>"></form>
                <input class="bot" type="button" value="VOLVER" onclick="document.getElementById('formx10').submit();">
            </div>
        </section>

        <section class="data-panel">
            <div class="data-table-wrap">
                <table class="data-table">
                    <thead>
                        <tr>
                            <th class="col-id">ID</th>
                            <th>TITULO</th>
                            <th class="col-action">&nbsp;</th>
                            <th>CLIENTE</th>
                            <th>DESARROLLO</th>
                            <th>FUNCIONAL</th>
                            <th>FEC INGRESO</th>
                        </tr>
                    </thead>
                    <tbody>
                    <%
                    Dim estD, estF
                    Do While Not rsConsulta.EOF
                        Select Case rsConsulta("estado_desa")
                            Case "PENDIENTE": estD = "<span class='status-badge status-badge--pending'>PENDIENTE</span>"
                            Case "OK": estD = "<span class='status-badge status-badge--ok'>OK</span>"
                            Case "BAJA": estD = "<span class='status-badge status-badge--danger'>BAJA</span>"
                            Case Else: estD = ""
                        End Select

                        Select Case rsConsulta("estado_func")
                            Case "PENDIENTE": estF = "<span class='status-badge status-badge--pending'>PENDIENTE</span>"
                            Case "OK": estF = "<span class='status-badge status-badge--ok'>OK</span>"
                            Case "BAJA": estF = "<span class='status-badge status-badge--danger'>BAJA</span>"
                            Case Else: estF = ""
                        End Select
                    %>
                        <form id="xform<%=rsConsulta("id")%>" class="form-oculto" method="post" action="detalles.asp?id=<%=rsConsulta("id")%>&opc=veri&tipo=cons"></form>
                        <tr>
                            <td class="col-id"><%=rsConsulta("id")%></td>
                            <td><%=Mid(UCase(rsConsulta("titulo")), 1, 58)%></td>
                            <td class="col-action">
                                <% If getPermisos(9, "INICIO", "CONS") = 1 Then %>
                                    <a class="action-button action-button--primary" href="javascript:document.getElementById('xform<%=rsConsulta("id")%>').submit();" title="Ver detalle">D</a>
                                <% Else %>
                                    <span class="action-button action-button--muted" title="Sin permiso para consultar">D</span>
                                <% End If %>
                            </td>
                            <td><%=rsConsulta("cliente")%></td>
                            <td class="text-center"><%=estD%></td>
                            <td class="text-center"><%=estF%></td>
                            <td class="text-center"><%=Fechador(rsConsulta("fecha_ingreso") + nFecini)%></td>
                        </tr>
                    <%
                        rsConsulta.MoveNext
                    Loop
                    rsConsulta.Close
                    %>
                    </tbody>
                </table>
            </div>
        </section>

    </main>
<% End If %>

<script type="text/javascript">
    function borrar(formu){ 
        var r = confirm('Desea eliminar el incidente?'); 
        if (r){ 
            document.getElementById(formu).submit(); 
        } else { 
            return false;
        }
    }
</script>
</body>
</html>
<%
Cnn.Close
Set Cnn = Nothing
%>