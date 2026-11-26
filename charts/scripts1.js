$(document).ready(function() {
    // Usamos el método ajax de jQuery para obtener los datos
    $.ajax({
        url: 'datos.json', // Ruta a nuestro archivo JSON
        type: 'GET',
        dataType: 'json',
        success: function(data) {
            // Si la solicitud es exitosa, llamamos a la función para crear el gráfico
            crearGrafico(data);
        },
        error: function(error) {
            // Si hay un error, lo mostramos en la consola
            console.error('Error al cargar los datos:', error);
        }
    });

    function crearGrafico(data) {
        // Obtenemos el contexto del canvas
        const ctx = document.getElementById('miGrafico').getContext('2d');

        // Creamos el gráfico
        const miGrafico = new Chart(ctx, {
            type: 'bar', // Tipo de gráfico: 'bar', 'line', 'pie', etc.
            data: {
                labels: data.etiquetas, // Usamos las etiquetas del JSON
                datasets: [{
                    label: 'Ventas Mensuales', // Leyenda del dataset
                    data: data.datos, // Usamos los datos del JSON
                    backgroundColor: [
                        'rgba(255, 99, 132, 0.5)',
                        'rgba(54, 162, 235, 0.5)',
                        'rgba(255, 206, 86, 0.5)',
                        'rgba(75, 192, 192, 0.5)',
                        'rgba(153, 102, 255, 0.5)',
                        'rgba(255, 159, 64, 0.5)'
                    ],
                    borderColor: [
                        'rgba(255, 99, 132, 1)',
                        'rgba(54, 162, 235, 1)',
                        'rgba(255, 206, 86, 1)',
                        'rgba(75, 192, 192, 1)',
                        'rgba(153, 102, 255, 1)',
                        'rgba(255, 159, 64, 1)'
                    ],
                    borderWidth: 1
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: true,
                scales: {
                    y: {
                        beginAtZero: true
                    }
                },
                plugins: {
                    title: {
                        display: true,
                        text: 'Ejemplo de Gráfico con AJAX y Chart.js'
                    }
                }
            }
        });
    }
});