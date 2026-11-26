<!DOCTYPE html>
<html>
<head runat="server">
    <title>Sistema de Tickets</title>
    <style>
        .comment-box { border: 1px solid #ccc; padding: 10px; margin-top: 10px; }
        .comment-user { font-weight: bold; color: #005a9c; }
        .comment-date { font-size: 0.8em; color: #666; }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            <h2>Ticket ID: 12345</h2>
            <p>Este es el detalle del problema que reportó el usuario...</p>

            <hr />

            <h3>Comentarios</h3>
            
            <asp:Repeater ID="rptComentarios" runat="server">
                <ItemTemplate>
                    <div class="comment-box">
                        <span class="comment-user"><%# Eval("NombreUsuario") %></span>
                        (<span class="comment-date"><%# Eval("Fecha", "{0:dd/MM/yyyy HH:mm}") %></span>)
                        <p><%# Eval("Texto") %></p>
                    </div>
                </ItemTemplate>
            </asp:Repeater>

            <hr />

            <h3>Agregar un Comentario</h3>
            <asp:Label ID="lblError" runat="server" ForeColor="Red"></asp:Label>
            
            <asp:TextBox ID="txtNuevoComentario" runat="server" 
                         TextMode="MultiLine" 
                         Rows="4" 
                         Width="100%"></asp:TextBox>
            <br />
            
            <asp:Button ID="btnGuardar" runat="server" 
                        Text="Guardar Comentario" 
                        OnClick="btnGuardar_Click" />
        </div>
    </form>
</body>
</html>