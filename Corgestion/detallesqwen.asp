<!-- #INCLUDE FILE="IncFile/datosDetalles.inc" -->
<!-- #INCLUDE FILE="IncFile/funciones.inc" -->
<%
Dim Cnn, id, opc, xmenu, xtmenu, rx, puedeEditar, cliente
Set Cnn = Server.CreateObject("ADODB.Connection")
Cnn.Open scon1
id = Request.QueryString("id")
opc = Request.QueryString("opc")
xmenu = Request.QueryString("xmenu")
xtmenu = Request.QueryString("xtmenu")
If Not IsNumeric(id) Then id = 0
puedeEditar = (CStr(getPermisos(xmenu, xtmenu, "MODIF")) = "1")

Response.Write headhtml()
%>
<link rel="stylesheet" href="assets/css/qwen-ui.css">
<body><main class="qw-detail"><header class="qw-detail-head"><div><p>Incidente #<%=id%></p><h1><%=Server.HTMLEncode(getName(Cnn,"titulo","incidentes",id) & "")%></h1></div><div class="qw-detail-actions"><a class="bot" href="incidentesqwen.asp?xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>">Volver</a><% If puedeEditar And LCase(opc) <> "acti" Then %><a class="bot" href="detallesqwen.asp?opc=acti&id=<%=id%>&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>">Editar</a><% End If %></div></header>
<% If LCase(opc) = "acti" Then
    Set rx = ObtenerIncidenteEdicion(Cnn, id)
    If rx.EOF Then
        Response.Write "<p class='bloque-datos'>No se encontro el incidente.</p>"
    Else
        Dim editaFunc, editaTec, editaFuncEstado
        editaFunc = puedeEditar And Session("FUNC") = "FUNC"
        editaTec = puedeEditar And Session("FUNC") = "TEC"
        editaFuncEstado = editaFunc
%><form id="sForm" class="qw-edit-form" method="post" action="actualizar_v2.asp?tabla=incidentes&relainc=relainc&id=<%=id%>&opt=actinc&xmenu=<%=xmenu%>&xtmenu=<%=xtmenu%>&vuelta=detallesqwen.asp?opc=acti"><section class="qw-edit-card"><h2>Identificacion</h2><div class="qw-edit-grid"><div class="qw-field"><label>Numero</label><input value="<%=id%>" disabled></div><div class="qw-field qw-field--wide"><label for="titulo">Titulo</label><%=HtmlInputCampo("titulo", rx(28), editaFunc)%></div><div class="qw-field"><label for="id_cliente">Cliente</label><%=ComboCampo(Cnn, rx(1), "id_cliente", "id_cliente", editaFunc)%></div><div class="qw-field"><label for="contacto">Contacto</label><%=ComboCampo(Cnn, rx(3), "contacto", "contacto", editaFunc)%></div><div class="qw-field"><label for="id_aplic">Aplicacion</label><%=ComboCampo(Cnn, rx(5), "id_aplic", "id_aplic", editaFunc)%></div><div class="qw-field"><label>Proyecto relacionado</label><%=ComboProyectoCampo(Cnn, rx(1), rx(1), editaFunc)%></div></div></section><section class="qw-edit-card"><h2>Clasificacion y estados</h2><div class="qw-edit-grid"><div class="qw-field"><label>Requiere presupuesto</label><%=HtmlSelectSiNo("req_presup", rx(14), editaFunc)%></div><div class="qw-field"><label>Browser</label><%=ComboCampo(Cnn, rx(20), "browser", "browser", editaFunc)%></div><div class="qw-field"><label>Via de pedido</label><%=ComboCampo(Cnn, rx(22), "via_pedido", "via_pedido", editaFunc)%></div><div class="qw-field"><label>Estado desarrollo</label><%=ComboCampo(Cnn, rx(15), "estado_desa", "estado_desa", editaTec)%></div><div class="qw-field"><label>Estado funcional</label><%=ComboCampo(Cnn, rx(17), "estado_func", "estado_func", editaFuncEstado)%></div><div class="qw-field"><label>Fecha OK cliente</label><%=HtmlInputCampo("ok_cliente", rx(25), editaFunc)%></div><div class="qw-field"><label>Fecha ingreso</label><input value="<%=Fechador(rx(19) + 0)%>" disabled></div></div></section><section class="qw-edit-card"><h2>Detalle y documentacion</h2><div class="qw-edit-grid"><div class="qw-field qw-field--wide"><label for="detalles">Descripcion</label><%=HtmlTextAreaCampo("detalles", rx(8), "220px", editaFunc)%></div><div class="qw-field qw-field--wide"><label for="ruta">Link documento</label><%=HtmlInputCampo("ruta", rx("ruta"), editaFunc)%></div></div></section><section class="qw-edit-card"><h2>Avance</h2><div class="qw-edit-grid"><div class="qw-field qw-field--wide"><label for="avances">Nuevo comentario</label><textarea id="avances" name="avances"></textarea></div></div></section><div class="qw-sticky"><a class="bot" href="incidentesqwen.asp">Cancelar</a><button class="bot" type="submit">Guardar cambios</button></div></form><%
    End If
Else
    Set rx = ObtenerDetalleIncidente(Cnn, id)
    If rx.EOF Then
        Response.Write "<p class='bloque-datos'>No se encontro el incidente.</p>"
    Else
%><section class="qw-edit-card"><h2>Datos del incidente</h2><div class="qw-read-grid"><div class="qw-label">Numero</div><div class="qw-value"><%=rx(0)%></div><div class="qw-label">Titulo</div><div class="qw-value"><%=Server.HTMLEncode(rx(22) & "")%></div><div class="qw-label">Cliente</div><div class="qw-value"><%=Server.HTMLEncode(rx(1) & "")%></div><div class="qw-label">Contacto</div><div class="qw-value"><%=Server.HTMLEncode(rx(2) & "")%></div><div class="qw-label">Aplicacion</div><div class="qw-value"><%=Server.HTMLEncode(rx(3) & "")%></div><div class="qw-label">Prioridad</div><div class="qw-value"><%=Server.HTMLEncode(rx(5) & "")%></div><div class="qw-label">Estado desarrollo</div><div class="qw-value"><%=Server.HTMLEncode(rx(9) & "")%></div><div class="qw-label">Estado funcional</div><div class="qw-value"><%=Server.HTMLEncode(rx(10) & "")%></div></div></section><section class="qw-edit-card"><h2>Detalle</h2><div class="qw-value"><%=Replace(Server.HTMLEncode(rx(4) & ""), vbCrLf, "<br>")%></div></section><%
    End If
End If
%></main></body>
