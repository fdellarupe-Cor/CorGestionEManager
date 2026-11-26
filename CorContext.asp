<%@ Language=VBScript CodePage=65001 %>
<%
' 🚀 Forzamos al motor IIS a procesar la sesión y las respuestas en UTF-8 nativo
Response.CharSet = "UTF-8"
Response.CodePage = 65001
Session.CodePage = 65001
%>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <meta charset="utf-8" />
    <title>Buscador Semántico de Auditoría</title>
    

    
    
</head>
<body class="bg-light">

<!-- Formulario clásico de auditoría semántica -->
<div style="margin: 20px; font-family: Arial, sans-serif;">
    <h2>🕵️‍♂️ Buscador Semántico desde ASP Classic</h2>
    <div style="margin-bottom: 15px;" class="input-group w-50">
        <input type="text" id="txtFraseUsuario" class="form-control" style="padding: 10px; font-size: 14px;" placeholder="Escriba su consulta legal o bancaria aquí..." />
        <button type="button" id="btnBuscar" class="btn btn-primary" style="font-size: 14px; cursor: pointer;" onclick="buscarNormasPorAjax()">Buscar Normas</button>
    </div>

    <!-- 🎯 CONTENEDOR DESTINO: Aquí se inyectará el bloque parcial HTML unificado si da successful -->
    <div id="divResultadosAgrupados" class="mt-4"></div>
</div>

<!-- 📦 POPUP CONTROLADO 100% POR JQUERY Y JQUERY UI -->
<div id="popupResumen" style="display: none; position: fixed; z-index: 9999; left: 0; top: 0; width: 100%; height: 100%; background-color: rgba(0,0,0,0.5);">
    
    <!-- Contenedor del cuadro Draggable -->
    <div id="modalDialogDraggable">
        
        <!-- Encabezado / Manija para arrastrar configurada para jQuery UI -->
        <div id="modalHeaderHandle" style="background: #17a2b8; color: #ffffff; padding: 15px; font-weight: bold; cursor: move; display: flex; justify-content: space-between; align-items: center; font-family: Arial, sans-serif;">
            <span id="lblTituloModal">✨ Resumen Ejecutivo</span>
            <button type="button" id="btnCerrarPopup" style="background: none; border: none; color: white; font-size: 20px; cursor: pointer; font-weight: bold;">&times;</button>
        </div>
        
        <!-- Cuerpo del texto devuelto por Llama3 -->
        <div style="background: #f8f9fa; padding: 20px; max-height: 400px; overflow-y: auto;">
            <div id="cuerpoTextoResumen" style="background: #ffffff; padding: 15px; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); color: #333333; white-space: pre-line; font-size: 15px; line-height: 1.6; font-family: Arial, sans-serif; text-align: justify;">
                <!-- Aquí se inyectará el JSON de Llama3 -->
            </div>
        </div>
        
        <!-- Pie del Popup -->
        <div style="background: #ffffff; padding: 15px; border-top: 1px solid #dee2e6; display: flex; justify-content: space-between; align-items: center; font-family: Arial, sans-serif;">
            <small style="color: #6c757d;">Procesado con Llama3 localmente</small>
            <button type="button" id="btnEntendidoPopup" class="btn btn-secondary btn-sm px-4">Entendido</button>
        </div>
    </div>
</div>

<!-- 📦 INCLUSIÓN DE LIBRERÍAS DE ORDEN ESTRICTO -->
<!-- 1. jQuery Core Nativo -->
<script src="https://ajax.googleapis.com/ajax/libs/jquery/3.7.1/jquery.min.js"></script>
<script src="https://code.jquery.com/ui/1.14.2/jquery-ui.min.js"> </script>
<link rel="stylesheet" href="https://code.jquery.com/ui/1.14.2/themes/base/jquery-ui.css">


<!-- 2. jQuery UI (Requerido para la función .draggable) -->


<script>
$(document).ready(function () {
    
    // 🚀 INICIALIZACIÓN DRAGGABLE DE JQUERY UI:
    // Hacemos que todo el cuadro se mueva usando únicamente el header azul como manija de arrastre
    $("#modalDialogDraggable").draggable({
        handle: "#modalHeaderHandle",
        containment: "window", // Evita que el usuario arrastre el popup fuera de los límites de la pantalla
        scroll: false
    });

    // Cierre controlado del popup con jQuery fadeOut
    $("#btnCerrarPopup, #btnEntendidoPopup").on("click", function () {
        $("#popupResumen").fadeOut(100, function() {
            $("#cuerpoTextoResumen").empty(); // Limpieza del búfer de texto residual
        });
    });
});

// ⚡ FUNCIÓN AJAX 1: BÚSQUEDA SEMÁNTICA GENERAL DE NORMAS
function buscarNormasPorAjax() {
    var fraseIngresada = $("#txtFraseUsuario").val();
    
    if (!fraseIngresada.trim()) {
        alert("Por favor, escriba una frase para buscar.");
        return;
    }

    $("#btnBuscar").prop("disabled", true).text("⏳ Buscando en SQL 2025...");

    var fData = new FormData();
    fData.append("FraseUsuario", fraseIngresada);
    fData.append("fuente", "externo"); // Flag interceptado por el BindProperty de C#

    $.ajax({
        url: "http://server-2025:8025/buscar?handler=Buscar", 
        type: "POST",
        data: fData,
        processData: false, 
        contentType: false, 
        crossDomain: true,
        xhrFields: { withCredentials: true },
        success: function (htmlParcialDevuelto) {
            // Inyectamos el tramo HTML de las tarjetas agrupadas directamente en tu div
            $("#divResultadosAgrupados").html(htmlParcialDevuelto); 
        },
        error: function (xhr, status, error) {
            $("#divResultadosAgrupados").html(
                '<div class="alert alert-danger">❌ Error al conectar con el servidor de auditoría semántica (.NET Core).</div>'
            );
            console.error("Detalle técnico del error AJAX:", error);
        },
        complete: function () {
            $("#btnBuscar").prop("disabled", false).text("Buscar Normas");
        }
    });
}

// ⚡ FUNCIÓN AJAX 2: SOLICITUD DE RESUMEN EJECUTIVO MEDIANTE LLAMA3
function solicitarResumenPorAjax(boton, nombreArchivoReal) {
    var fraseActual = $("#txtFraseUsuario").val();
    var textoOriginalBoton = $(boton).html();

    $(boton).css("pointer-events", "none").html("⏳ Procesando resumen mediante Llama3... Por favor, espere.");

    var fData = new FormData();
    fData.append("FraseUsuario", fraseActual);
    fData.append("DocumentoSeleccionado", nombreArchivoReal);
    fData.append("fuente", "externo"); 

    $.ajax({
        url: "http://server-2025:8025/buscar?handler=Resumir",
        type: "POST",
        data: fData,
        processData: false,
        contentType: false,
        crossDomain: true,
        xhrFields: { withCredentials: true },
        success: function (dataJson) {
            if (dataJson.success) {
                // Inyectamos los textos dinámicos devueltos por el JSON de .NET Core
                $("#lblTituloModal").html("✨ Resumen Ejecutivo: " + dataJson.archivo);
                $("#cuerpoTextoResumen").html(dataJson.resumen);

                // Centrado exacto del cuadro Draggable respecto al monitor actual
                var cuadroDraggable = $("#modalDialogDraggable");
               

                // Desplegamos el popup usando transiciones nativas suaves de jQuery
                $("#popupResumen").show();
            } else {
                alert("⚠️ No se pudo generar el resumen: " + dataJson.mensaje);
            }
        },
        error: function () {
            alert("❌ Error crítico en la comunicación remota con Llama3.");
        },
        complete: function () {
            $(boton).css("pointer-events", "auto").html(textoOriginalBoton);
        }
    });
}
</script>

</body>
</html>
