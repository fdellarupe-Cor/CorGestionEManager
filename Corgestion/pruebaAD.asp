<%
    ' 1. OBTENER AL USUARIO LOGUEADO
    Dim usuarioActual
    usuarioActual = Request.ServerVariables("LOGON_USER")
    
    ' 2. OBTENER LOS DATOS DEL FORMULARIO
    Dim comentarioTexto
    Dim idTicket
    comentarioTexto = Request.Form("txtComentario")
    idTicket = Request.Form("hiddenTicketID")

    ' 3. (Opcional) Limpiar el nombre de usuario
    Dim nombreSimple
    Dim partesUsuario
    partesUsuario = Split(usuarioActual, "\")
    If UBound(partesUsuario) = 1 Then
        nombreSimple = partesUsuario(1)
    Else
        nombreSimple = usuarioActual
    End If

    ' 4. GUARDAR LOS DATOS
    ' Aquí va tu lógica de conexión a la base de datos (ADODB)
    ' Ejemplo:
    ' Call GuardarEnBaseDeDatos(idTicket, nombreSimple, comentarioTexto)
    
    ' 5. Redirigir de vuelta al ticket
    Response.Write "Comentario guardado por: " & nombreSimple
    Response.Write "<br><a href='ticket.asp?ID=" & idTicket & "'>Volver al Ticket</a>"
%>