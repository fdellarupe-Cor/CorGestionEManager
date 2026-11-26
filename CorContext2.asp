<%@ Language=VBScript CodePage=65001 %>
<%
' 🚀 Forzamos al motor IIS a procesar la sesión y las respuestas en UTF-8
Response.CharSet = "UTF-8"
Response.CodePage = 65001
Session.CodePage = 65001
%>


<head>
    <!-- 🚀 Le dice al navegador el juego de caracteres correcto -->
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta charset="utf-8" />
    <title>Buscador Semántico de Auditoría</title>
</head>

<form action="http://server-2025:8025/Buscar?handler=Buscar" method="post" >
    <div class="input-group">
        <!-- 'name' debe coincidir exactamente con el BindProperty [FraseUsuario] de BuscarFraseModel -->
        <input type="text" name="FraseUsuario" class="form-control form-control-lg" placeholder="Escriba su consulta legal o bancaria aquí..." required />
		 <input type="text" name="fuente" id="fuente" value='externo' />
		
        <button type="submit" class="btn btn-primary btn-lg px-4" id="btnBuscar">Buscar Normas</button>
    </div>
</form>


<!-- Incluimos jQuery desde una CDN clásica compatible con ASP Viejo -->
<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.7.1/jquery.min.js"></script>


