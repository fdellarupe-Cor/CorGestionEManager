<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<!-- #INCLUDE FILE="IncFile/datosAgregar.inc" -->
<%
Dim Cnn, xmenu, xtmenu, tipo, tabla, volver
Set Cnn = Server.CreateObject("ADODB.Connection")
Cnn.Open scon1
xmenu = Request.QueryString("xmenu")
xtmenu = Request.QueryString("xtmenu")
tipo = UCase(Request.QueryString("tipo"))
If tipo = "" Then tipo = "INC"
If tipo = "INC" Then
    tabla = "INCIDENTES"
    volver = "incidentesqwen.asp"
Else
    Response.Redirect "agregar2.asp?tipo=" & tipo & "&xmenu=" & xmenu & "&xtmenu=" & xtmenu
End If
Response.Write headhtml()
%>
<link rel="stylesheet" href="assets/css/qwen-ui.css">
<body><main class="qw-add"><header class="qw-add-head"><div><p>Alta de incidente</p><h1>Nuevo incidente</h1></div><div class="qw-add-actions"><a class="bot" href="<%=volver%>">Volver</a></div></header><form id="agre" class="qw-add-card" method="post" action="agregar2.asp?opc=ins&tabla=<%=tabla%>&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>&tipo=INC&volver=<%=volver%>"><div class="qw-add-grid"><div class="qw-field"><label for="cli">Cliente</label><select id="cli" name="cli"><% Dim rs: Set rs = ObtenerClientes(Cnn): Do While Not rs.EOF %><option value="<%=rs(0)%>"><%=Server.HTMLEncode(rs(1) & "")%></option><% rs.MoveNext: Loop %></select></div><div class="qw-field"><label for="apli">Aplicacion</label><select id="apli" name="apli"><% Set rs = ObtenerAplicativos(Cnn): Do While Not rs.EOF %><option value="<%=rs(0)%>"><%=Server.HTMLEncode(rs(1) & "")%></option><% rs.MoveNext: Loop %></select></div><div class="qw-field"><label for="prio">Prioridad</label><select id="prio" name="prio"><% Set rs = ObtenerPrioridades(Cnn): Do While Not rs.EOF %><option value="<%=rs(0)%>"><%=Server.HTMLEncode(rs(1) & "")%></option><% rs.MoveNext: Loop %></select></div><div class="qw-field"><label for="vp">Via de pedido</label><select id="vp" name="vp"><% Set rs = ObtenerViaPedido(Cnn): Do While Not rs.EOF %><option value="<%=rs(0)%>"><%=Server.HTMLEncode(rs(1) & "")%></option><% rs.MoveNext: Loop %></select></div><div class="qw-field qw-wide"><label for="tit">Titulo</label><input id="tit" name="tit" type="text" required maxlength="255"></div><div class="qw-field"><label for="con">Contacto</label><select id="con" name="con"><% Set rs = ObtenerContactosParaIncidente(Cnn): Do While Not rs.EOF %><option value="<%=rs(0)%>"><%=Server.HTMLEncode(rs(1) & "")%></option><% rs.MoveNext: Loop %></select></div><div class="qw-field"><label for="rg">Responsable gestion</label><select id="rg" name="rg"><% Set rs = ObtenerResponsables(Cnn,"FUNC"): Do While Not rs.EOF %><option value="<%=rs(0)%>"><%=Server.HTMLEncode((rs(1) & "") & " " & (rs(2) & ""))%></option><% rs.MoveNext: Loop %></select></div><div class="qw-field"><label for="rt">Responsable tecnico</label><select id="rt" name="rt"><option value="0">Ninguno</option><% Set rs = ObtenerResponsables(Cnn,"TEC"): Do While Not rs.EOF %><option value="<%=rs(0)%>"><%=Server.HTMLEncode((rs(1) & "") & " " & (rs(2) & ""))%></option><% rs.MoveNext: Loop %></select></div><div class="qw-field"><label for="rp">Requiere presupuesto</label><select id="rp" name="rp"><option value="NO">NO</option><option value="SI">SI</option></select></div><div class="qw-field"><label for="br">Browser</label><select id="br" name="br"><% Set rs = ObtenerBrowsers(Cnn): Do While Not rs.EOF %><option value="<%=rs(0)%>"><%=Server.HTMLEncode(rs(1) & "")%></option><% rs.MoveNext: Loop %></select></div><div class="qw-field qw-wide"><label for="det">Detalles</label><textarea id="det" name="det"></textarea></div></div><div class="qw-submit"><button class="bot" type="reset">Limpiar</button><button class="bot" type="submit">Agregar incidente</button></div></form></main></body>
