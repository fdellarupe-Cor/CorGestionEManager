<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<%
Dim Cnn, cliente, aplicativo, tit, ord, xmenu, xtmenu, mover, mostrar, rsListado
Set Cnn = Server.CreateObject("ADODB.Connection")
Cnn.Open scon1
cliente = Request("cli")
If Not IsNumeric(cliente) Then cliente = 0
aplicativo = Request("apli")
If Not IsNumeric(aplicativo) Then aplicativo = 0
tit = Request("tit")
ord = Request.QueryString("ord")
xmenu = Request.QueryString("xmenu")
xtmenu = Request.QueryString("xtmenu")
mover = Request.QueryString("cantregmover")
If Not IsNumeric(mover) Then mover = 0
mostrar = 10

Function QwenListado(cn, idCli, idApli, titulo, orden)
    Dim cmd
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = cn
    cmd.CommandText = "sp_Incidentes_Listar"
    cmd.CommandType = 4
    cmd.Parameters.Append cmd.CreateParameter("@IdCliente", 3, 1, , CLng(idCli))
    cmd.Parameters.Append cmd.CreateParameter("@IdAplicativo", 3, 1, , CLng(idApli))
    cmd.Parameters.Append cmd.CreateParameter("@Titulo", 200, 1, 255, CStr(titulo & ""))
    cmd.Parameters.Append cmd.CreateParameter("@Orden", 200, 1, 20, CStr(orden & ""))
    Set QwenListado = cmd.Execute()
End Function

Response.Write headhtml()
%>
<link rel="stylesheet" href="assets/css/qwen-ui.css">
<body>
<header class="qw-header"><a class="qw-brand" href="index.asp"><img src="images/iso_cor.png" alt="COR Gestion"><span>Cor Gestion<small>Sistema Administrativo</small></span></a><a class="btn-hamburguesa" href="#menu" aria-label="Abrir menu">&#9776;</a></header>
<%=getMenuParam()%>
<main class="qw-page">
<section class="qw-toolbar" aria-labelledby="qw-title"><h1 id="qw-title">Incidentes</h1><form class="qw-filters" method="post" action="incidentesqwen.asp"><div class="qw-field"><label for="cli">Cliente</label><select id="cli" name="cli" onchange="this.form.submit()"><option value="0">Todos</option><%
Dim rsCli, rsApli, selected
Set rsCli = Exec(Cnn, "sp_Obtener_Catalogo @Tabla='clientes'")
Do While Not rsCli.EOF
    selected = ""
    If CStr(rsCli("id")) = CStr(cliente) Then selected = " selected"
%><option value="<%=Server.HTMLEncode(rsCli("id"))%>"<%=selected%>><%=Server.HTMLEncode(rsCli("descripcion"))%></option><% rsCli.MoveNext: Loop
rsCli.Close
%></select></div><div class="qw-field"><label for="apli">Aplicativo</label><select id="apli" name="apli" onchange="this.form.submit()"><option value="0">Todos</option><%
Set rsApli = Exec(Cnn, "sp_Obtener_Catalogo @Tabla='aplicativos'")
Do While Not rsApli.EOF
    selected = ""
    If CStr(rsApli("id")) = CStr(aplicativo) Then selected = " selected"
%><option value="<%=Server.HTMLEncode(rsApli("id"))%>"<%=selected%>><%=Server.HTMLEncode(rsApli("descripcion"))%></option><% rsApli.MoveNext: Loop
rsApli.Close
%></select></div><div class="qw-field"><label for="tit">Titulo</label><input id="tit" name="tit" value="<%=Server.HTMLEncode(tit & "")%>"></div></form><div class="qw-actions"><a class="bot" href="agregarqwen.asp?tipo=INC&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>">Nuevo incidente</a><a class="bot" href="incidentesqwen.asp?cantregmover=<%=mover-mostrar%>&cli=<%=cliente%>&apli=<%=aplicativo%>"<% If CLng(mover)=0 Then %> aria-disabled="true"<% End If %>>Anterior</a><a class="bot" href="incidentesqwen.asp?cantregmover=<%=mover+mostrar%>&cli=<%=cliente%>&apli=<%=aplicativo%>">Siguiente</a><a class="bot" href="index.asp">Volver</a></div></section>
<section class="bloque-datos qw-table-wrap"><table class="qw-table"><caption class="sr-only">Listado de incidentes</caption><thead><tr><th>Acciones</th><th>ID</th><th>Titulo</th><th>Cliente</th><th>Prioridad</th><th>Aplicativo</th><th>Desarrollo</th><th>Funcional</th><th>Ingreso</th></tr></thead><tbody><%
Set rsListado = QwenListado(Cnn, cliente, aplicativo, tit, ord)
If Not rsListado.EOF Then rsListado.Move CLng(mover)
Dim i, detalleUrl, editarUrl, estadoD, estadoF, prioridad
For i = 1 To mostrar
    If rsListado.EOF Then Exit For
    detalleUrl = "detallesqwen.asp?opc=veri&id=" & rsListado("id") & "&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu
    editarUrl = "detallesqwen.asp?opc=acti&id=" & rsListado("id") & "&cli=" & cliente & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu
    estadoD = Server.HTMLEncode(rsListado("estado_desa") & "")
    estadoF = Server.HTMLEncode(rsListado("estado_func") & "")
    prioridad = Server.HTMLEncode(rsListado("prioridad") & "")
%><tr><td class="qw-actions-cell"><a class="btn-accion btn-accion-primario" href="<%=editarUrl%>" aria-label="Editar incidente <%=rsListado("id")%>">A</a><form method="post" action="eliminar.asp?opc=inc&id=<%=rsListado("id")%>&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>&vuelta=incidentesqwen.asp" style="display:inline"><button class="btn-accion btn-accion-peligro" type="submit" aria-label="Eliminar incidente <%=rsListado("id")%>">X</button></form></td><td><a class="qw-link" href="<%=detalleUrl%>"><%=rsListado("id")%></a></td><td><a class="qw-link" href="<%=detalleUrl%>"><%=Server.HTMLEncode(Left(UCase(rsListado("titulo") & ""),60))%></a></td><td><%=Server.HTMLEncode(rsListado("cliente") & "")%></td><td><span class="badge"><%=prioridad%></span></td><td><%=Server.HTMLEncode(rsListado("aplicativo") & "")%></td><td><span class="badge"><%=estadoD%></span> <%=Server.HTMLEncode(rsListado("respon_tec_nombre") & "")%></td><td><span class="badge"><%=estadoF%></span> <%=Server.HTMLEncode(rsListado("respon_gestion_nombre") & "")%></td><td><%=Fechador(rsListado("fecha_ingreso") + 0)%></td></tr><%
    rsListado.MoveNext
Next
rsListado.Close
%></tbody></table></section><p><strong>(C)</strong> Consulta resuelta sin incumbencia tecnica.</p></main>
</body>
