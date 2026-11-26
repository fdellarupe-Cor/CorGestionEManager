<%
' 1. Configurar los encabezados de respuesta para que se descargue como archivo de Word
Response.ContentType = "application/msword"
Response.AddHeader "Content-Disposition", "attachment; filename=InformeCompleto.doc"

' Opcional: Declarar el Programmatic ID para ayudar a Word a reconocer el formato
Response.Write "<meta name=ProgId content=Word.Document>"
%>
<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">
<head>
    <title>Informe con Numeración de Página Robusta</title>

    <style>
        /* 1. Definir la sección del documento */
        @page Section1 {
            size: A4;
            margin: 1in 1in 1in 1in;

            /* Declaraciones MSO para Encabezados/Pies */
            mso-header: h1;          /* Encabezado páginas impares */
            mso-even-header: eh1;    /* Encabezado páginas pares */
            mso-first-header: fh1;   /* Encabezado primera página */
            mso-footer: f1;          /* Pie de páginas impares */
            mso-even-footer: ef1;    /* Pie de páginas pares */
            mso-first-footer: ff1;   /* Pie de primera página */
        }
        div.Section1 { page: Section1; }

        /* 2. Ocultar los bloques de encabezado/pie de página para el navegador */
        .HeaderFooter {
            display: none;
            /* Directivas CRÍTICAS para que Word los reconozca como plantilla y los oculte en el cuerpo */
            mso-element: header;
            mso-hide: all;
        }

        /* Estilo para el número de página (Word lo usa internamente) */
        .MsoPageNumber { }

        /* 3. Estilos de contenido */
        body { font-family: 'Times New Roman', Times, serif; }
        h2 { page-break-before: always; }

        .data-table {
            border-collapse: collapse;
            width: 100%;
        }
        .data-table th, .data-table td {
            border: 1px solid #ccc;
            padding: 8px;
        }
    </style>
</head>
<body>

<div class="HeaderFooter" id="fh1">
    <div style="text-align: center; border-bottom: 2pt solid #000080; padding-bottom: 5px;">
        <span style="font-size: 14pt; font-weight: bold; color: #000080;">INFORME PRINCIPAL - Confidencial</span>
    </div>
</div>

<div class="HeaderFooter" id="h1">
    <div style="text-align: right; border-bottom: 1pt solid #808080; font-size: 10pt;">
        Informe - Sección A
    </div>
</div>

<div class="HeaderFooter" id="eh1">
    <div style="text-align: left; border-bottom: 1pt solid #808080; font-size: 10pt;">
        Generado el: <%= Date() %>
    </div>
</div>

<div class="HeaderFooter" id="ff1">
    <div style="text-align: center; padding-top: 10px; font-size: 9pt;">
        Página
        <span class=MsoPageNumber><o:p></o:p></span>
        de
        <span class=MsoPageNumber><o:p></o:p></span>
        | <span style="font-weight: bold;">Inicio del Documento</span>
    </div>
</div>

<div class="HeaderFooter" id="f1">
    <div style="text-align: right; font-size: 9pt;">
        Página
        <span class=MsoPageNumber><o:p></o:p></span>
        de
        <span class=MsoPageNumber><o:p></o:p></span>
    </div>
</div>

<div class="HeaderFooter" id="ef1">
    <div style="text-align: left; font-size: 9pt;">
        Página
        <span class=MsoPageNumber><o:p></o:p></span>
        de
        <span class=MsoPageNumber><o:p></o:p></span>
    </div>
</div>

<div class="Section1">

    <h1>Título Principal del Informe</h1>
    <p>Este documento está estructurado con una sección que define encabezados y pies de página específicos para la primera, páginas pares e impares, además de la numeración automática de Word.</p>

    <table class="data-table">
        <thead>
            <tr>
                <th>Producto</th>
                <th>Ventas Q1</th>
                <th>Ventas Q2</th>
            </tr>
        </thead>
        <tbody>
            <%
            ' Ejemplo de datos
            Datos = Array(_
                Array("Laptop Pro", 1200, 1550), _
                Array("Monitor Ultra", 800, 950), _
                Array("Teclado Mecánico", 2500, 3100) _
            )

            For Each item In Datos
                Response.Write "<tr>"
                Response.Write "<td>" & item(0) & "</td>"
                Response.Write "<td style='text-align: center;'>" & item(1) & "</td>"
                Response.Write "<td style='text-align: center;'>" & item(2) & "</td>"
                Response.Write "</tr>"
            Next
            %>
        </tbody>
    </table>

    <br><br>

    <h2>Contenido Extenso (Paginación)</h2>
    <%
    ' Bucle para generar contenido y forzar el salto de página (mínimo 5 páginas)
    For i = 1 To 150
        If i = 50 Then Response.Write "<h2>Nuevo Tema (Inicia Nueva Página)</h2>"
        If i = 100 Then Response.Write "<h2>Otro Tema (Inicia Nueva Página)</h2>"
        Response.Write "<p>Párrafo de relleno N° " & i & ". Esto garantiza que se generen páginas pares e impares para probar la funcionalidad de los encabezados/pies.</p>"
    Next
    %>

    <p>*** Fin del Contenido Principal ***</p>

</div> ---

<p style='margin-top:0pt; margin-bottom:0pt;'>
    <span style='mso-special-character:line-break;mso-tab-count:1'></span>
</p>
<div style='mso-element:section-delimiter'></div>
</body>
</html>