<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<%
'Option Explicit

' --- CONEXI�N A BASE DE DATOS ---
Dim Cnn
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

' --- DECLARACI�N OBLIGATORIA DE TODAS LAS VARIABLES ---
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
' VISTA / PRESENTACI�N (HTML)
' =========================================================================

Response.Write Replace(headhtml(), "</head>", "<link rel='stylesheet' href='assets/css/copilot-ui.css' />" & "</head>")
%>
<body id="body" name="body" class="cg-copilot-page">
<div id="wrapper">
    <header id="header" class="app-header">
        <a href="index.asp" class="logo app-header__brand">
            <img src="images/iso_cor.png" alt="Cor Gestion" class="app-header__logo">
            <span><b>Cor Gestion</b> | Sistema Administrativo</span>
        </a>
        <nav class="app-header__nav">
            <a href="#menu" class="btn-hamburguesa" aria-label="Abrir menú">
                <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor" stroke-width="2">
                    <path stroke-linecap="round" stroke-linejoin="round" d="M4 6h16M4 12h16M4 18h16" />
                </svg>
            </a>
        </nav>
    </header>
</div>
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

    <div class="inner cg-page-content">
        <!-- CABECERA Y FILTROS -->
        <div class="titulo-seccion page-header">
            <div class="page-header__title"><h2>INCIDENTES</h2></div>
            <div class="toolbar__filters">
                        <form id="xForm" method="post" action="incidentes.asp?ord=<%=ord%>&cli=<%=cliente%>&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>">
                            <!-- Combo Clientes -->
                            <select onchange="document.getElementById('xForm').submit();" id="cli" name="cli">
                                <option value="0">CLIENTE: TODOS</option>
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

                            <!-- Combo Aplicativos -->
                            <select onchange="document.getElementById('xForm').submit();" id="apli" name="apli">
                                <option value="0">APLICATIVO: TODOS</option>
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
                        </form>
            </div>

                    <!-- ACCIONES -->
            <div class="toolbar__actions">
                    <% If Session("FUNC") = "FUNC" Or Session("FUNC") = "TEC" Then %>
                            <form id="formx4" class="form-oculto" method="post" action="agregar2.asp?tipo=INC&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>"></form>
                            <% If CStr(getPermisos(xmenu, xtmenu, "ALTA")) = "1" Then %>
                                <input class="bot" type="button" value="CARGAR NUEVO INCIDENTE" onclick="document.getElementById('formx4').submit();">
                            <% Else %>
                                <input class="botsp" type="button" value="CARGAR NUEVO INCIDENTE" disabled>
                            <% End If %>
                    <% End If %>

                    <!-- PAGINACI�N Y NAVEGACI�N -->
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
                        <form id="formx5" style="height:70%;" method="post" action="index.asp"></form>
                        <input class="bot" type="button" value="VOLVER" onclick="document.getElementById('formx5').submit();">
            </div>
        </div>

        <!-- GRILLA PRINCIPAL -->
        <div class="bloque-datos table-responsive">
            <table class="tabla-moderna">
                <thead>
                    <tr>
                        <th></th>
                        <th></th>
                        <th>ID</th>
                        <th><a class="col-link" style="text-decoration: none;" href="incidentes.asp?ord=titulo&cli=<%=cliente%>&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>">TITULO</a></th>
                        <th><a class="col-link" style="text-decoration: none;" href="incidentes.asp?ord=cliente&cli=<%=cliente%>&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>">CLIENTE</a></th>
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
                        Case "PENDIENTE": est1 = "<span class='badge badge-media'>PEND</span> (" & rsListado("respon_tec_nombre") & ")"
                        Case "OK": est1 = "<span class='badge badge-ok'>OK</span> (" & rsListado("respon_tec_nombre") & ")"
                        Case "BAJA": est1 = "<span class='badge badge-peligro'>BAJA</span> (" & rsListado("respon_tec_nombre") & ")"
                        Case "TESTING EN CLIENTE": est1 = "<span class='badge badge-media'>TEST</span> (" & rsListado("respon_tec_nombre") & ")"
                        Case Else: est1 = "-------------"
                    End Select

                    Select Case rsListado("estado_func")
                        Case "PENDIENTE": est2 = "<span class='badge badge-media'>PEND</span> (" & rsListado("respon_gestion_nombre") & ")"
                        Case "OK": est2 = "<span class='badge badge-ok'>OK</span> (" & rsListado("respon_gestion_nombre") & ")"
                        Case "BAJA": est2 = "<span class='badge badge-peligro'>BAJA</span> (" & rsListado("respon_gestion_nombre") & ")"
                        Case "TESTING EN CLIENTE": est2 = "<span class='badge badge-media'>TEST</span> (" & rsListado("respon_gestion_nombre") & ")"
                        Case Else: est2 = "-------------"
                    End Select

                    If rsListado("prioridad") = "ALTA" Then
                        prio = "<span class='badge badge-peligro'>ALTA</span>"
                    ElseIf rsListado("prioridad") = "MEDIA" Then
                        prio = "<span class='badge badge-media'>MEDIA</span>"
                    Else
                        prio = "<span class='badge badge-pendiente'>" & rsListado("prioridad") & "</span>"
                    End If

                    sFec = Fechador(rsListado("fecha_ingreso") + nFecini)
                %>
                    <form id="formx<%=rsListado("id")%>" class="form-oculto" method="post" action="eliminar.asp?opc=inc&id=<%=rsListado("id")%>&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>&vuelta=incidentes.asp"></form>
                    <form id="formx2<%=rsListado("id")%>" class="form-oculto" method="post" action="detalles.asp?id=<%=rsListado("id")%>&opc=acti&cli=<%=cliente%>&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>"></form>

                    <tr>
                        <td>
                            <% If CStr(getPermisos(xmenu, xtmenu, "ELIM")) = "1" Then %>
                                <input class="btn-accion btn-accion-peligro" type="button" value="X" onclick="borrar('formx<%=rsListado("id")%>');">
                            <% Else %>
                                <input class="btn-accion btn-accion-deshabilitado" type="button" value="X" onclick="alert('Sin permiso para eliminar.');">
                            <% End If %>
                        </td>
                        <td>
                            <input class="btn-accion btn-accion-primario" type="button" value="A" onclick="document.getElementById('formx2<%=rsListado("id")%>').submit();">
                        </td>
                        <td onclick="document.location='detalles.asp?id=<%=rsListado("id")%>&opc=veri&cli=<%=cliente%>&tit=<%=tit%>&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>'"><%=rsListado("id")%></td>
                        <td title="<%=rsListado("titulo")%>" onclick="document.location='detalles.asp?id=<%=rsListado("id")%>&opc=veri&cli=<%=cliente%>&tit=<%=tit%>&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>'">
                            <%=Left(UCase(rsListado("titulo")), 40)%>
                            <% If rsListado("respon_tec") = 0 Then %> <font style="color:#000FFF;"><b>(C)</b></font><% End If %>
                        </td>
                        <td onclick="document.location='detalles.asp?id=<%=rsListado("id")%>&opc=veri&cli=<%=cliente%>&tit=<%=tit%>&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>'"><%=rsListado("cliente")%></td>
                        <td onclick="document.location='detalles.asp?id=<%=rsListado("id")%>&opc=veri&cli=<%=cliente%>&tit=<%=tit%>&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>'"><%=prio%></td>
                        <td onclick="document.location='detalles.asp?id=<%=rsListado("id")%>&opc=veri&cli=<%=cliente%>&tit=<%=tit%>&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>'"><%=rsListado("aplicativo")%></td>
                        <td onclick="document.location='detalles.asp?id=<%=rsListado("id")%>&opc=veri&cli=<%=cliente%>&tit=<%=tit%>&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>'"><%=est1%></td>
                        <td onclick="document.location='detalles.asp?id=<%=rsListado("id")%>&opc=veri&cli=<%=cliente%>&tit=<%=tit%>&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>'"><%=est2%></td>
                        <td onclick="document.location='detalles.asp?id=<%=rsListado("id")%>&opc=veri&cli=<%=cliente%>&tit=<%=tit%>&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>'"><%=sFec%></td>
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

        <div class="legend cg-legend">
            <span class="legend__key"><b>*</b><b>(C)</b></span>
            <span class="legend__text">(Consulta) resuelta sin incumbencia tecnica.</span>
        </div>
    </div>

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
    <div class="inner cg-page-content">
        <div class="titulo-seccion page-header">
            <div class="page-header__title"><h2>CONSULTA INCIDENTES</h2></div>
            <div class="toolbar__actions">
                <form id="formx10" class="form-oculto" method="post" action="consultas.asp?sel=<%=sel%>&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>"></form>
                <input class="bot" type="button" value="VOLVER" onclick="document.getElementById('formx10').submit();">
            </div>
        </div>

        <div class="bloque-datos table-responsive">
            <table class="tabla-moderna" style="width: 100%;">
                <thead>
                    <tr>
                        <th style="width: 5%; text-align: center;">ID</th>
                        <th style="width: 40%;">TITULO</th>
                        <th style="width: 40px; text-align: center;">&nbsp;</th>
                        <th style="width: 18%;">CLIENTE</th>
                        <th style="width: 10.5%; text-align: center;">DESARROLLO</th>
                        <th style="width: 10.5%; text-align: center;">FUNCIONAL</th>
                        <th style="width: 15%; text-align: center;">FEC INGRESO</th>
                    </tr>
                </thead>
                <tbody>
                <%
                Dim estD, estF
                Do While Not rsConsulta.EOF
                    Select Case rsConsulta("estado_desa")
                        Case "PENDIENTE": estD = "<span class='badge badge-media'>PENDIENTE</span>"
                        Case "OK": estD = "<span class='badge badge-ok'>OK</span>"
                        Case "BAJA": estD = "<span class='badge badge-peligro'>BAJA</span>"
                        Case Else: estD = ""
                    End Select

                    Select Case rsConsulta("estado_func")
                        Case "PENDIENTE": estF = "<span class='badge badge-media'>PENDIENTE</span>"
                        Case "OK": estF = "<span class='badge badge-ok'>OK</span>"
                        Case "BAJA": estF = "<span class='badge badge-peligro'>BAJA</span>"
                        Case Else: estF = ""
                    End Select
                %>
                    <form id="xform<%=rsConsulta("id")%>" style="margin: 0; padding: 0;" method="post" action="detalles.asp?id=<%=rsConsulta("id")%>&opc=veri&tipo=cons"></form>
                    <tr>
                        <td style="text-align: center; color: #475569; font-weight: 500;"><%=rsConsulta("id")%></td>
                        <td><%=Mid(UCase(rsConsulta("titulo")), 1, 58)%></td>
                        <td style="text-align: center; vertical-align: middle;">
                            <% If getPermisos(9, "INICIO", "CONS") = 1 Then %>
                                <input class="btn-accion btn-accion-primario" type="button" value="D" onclick="document.getElementById('xform<%=rsConsulta("id")%>').submit();">
                            <% Else %>
                                <input class="btn-accion btn-accion-deshabilitado" type="button" value="D" onclick="alert('Sin permiso para consultar.');">
                            <% End If %>
                        </td>
                        <td style="padding-left: 1rem;"><%=rsConsulta("cliente")%></td>
                        <td style="text-align: center;"><%=estD%></td>
                        <td style="text-align: center;"><%=estF%></td>
                        <td style="text-align: center;"><%=Fechador(rsConsulta("fecha_ingreso") + nFecini)%></td>
                    </tr>
                <%
                    rsConsulta.MoveNext
                Loop
                rsConsulta.Close
                %>
                </tbody>
            </table>
        </div>
    </div>
<% End If %>

<script type="text/javascript">
    function borrar(formu){ 
        var r = confirm('�Desea eliminar el incidente?'); 
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