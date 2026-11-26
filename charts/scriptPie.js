$(document).ready(function() {
    $.ajax({
        url: 'datos.json',
        type: 'GET',
        dataType: 'json',
        success: function(data) {
            crearGraficoPastel(data); // Llamamos a la función para crear el gráfico de pastel
        },
        error: function(error) {
            console.error('Error al cargar los datos:', error);
        }
    });

    function crearGraficoPastel(data) {
        const ctx = document.getElementById('miGrafico').getContext('2d');

        const miGraficoPastel = new Chart(ctx, {
            type: 'pie', // ¡Cambiamos el tipo a 'pie'!
            data: {
                labels: data.etiquetas, // Las etiquetas serán las secciones del pastel
                datasets: [{
                    label: 'Ventas Mensuales',
                    data: data.datos, // Los datos serán el tamaño de cada sección
                    backgroundColor: [ // Colores para cada sección del pastel
                        'rgba(255, 99, 132, 0.7)', // Rojo
                        'rgba(54, 162, 235, 0.7)', // Azul
                        'rgba(255, 206, 86, 0.7)', // Amarillo
                        'rgba(75, 192, 192, 0.7)', // Verde azulado
                        'rgba(153, 102, 255, 0.7)', // Púrpura
                        'rgba(255, 159, 64, 0.7)'  // Naranja
                    ],
                    borderColor: [ // Bordes para cada sección
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
                plugins: {
                    title: {
                        display: true,
                        text: 'Ventas Mensuales (Gráfico de Pastel)'
                    },
                    tooltip: {
                        callbacks: {
                            label: function(context) {
                                let label = context.label || '';
                                if (label) {
                                    label += ': ';
                                }
                                if (context.parsed !== null) {
                                    label += new Intl.NumberFormat('es-ES', { style: 'currency', currency: 'USD' }).format(context.parsed); // Formato de moneda, ajusta si es necesario
                                }
                                return label;
                            }
                        }
                    },
                    legend: {
                        display: true,
                        position: 'right' // Muestra la leyenda a la derecha
                    }
                },
                 onClick: (event, elements) => {
				                    // Verificamos si el usuario hizo clic en algún elemento del gráfico
				                    if (elements.length > 0) {
				                        const primerElemento = elements[0]; // Obtenemos el primer elemento clickeado
				                        const indice = primerElemento.index; // Obtenemos el índice del dato en el array

				                        // Usamos el 'indice' para obtener la etiqueta y el valor correspondiente
				                        const etiqueta = miGraficoPastel.data.labels[indice];
				                        const valor = miGraficoPastel.data.datasets[0].data[indice];

				                        // Mostramos una alerta con la información
				                        alert(`Hiciste clic en '${etiqueta}' que tiene un valor de ${valor}.`);

				                        // Aquí podrías hacer cualquier otra cosa, como redirigir a una página:
				                        // window.location.href = '/detalles/' + etiqueta;
				                    }
                }
                // Los gráficos de pastel no suelen tener escalas 'y' o 'x'
                // por lo que las removemos de las opciones.
            }
        });
    }
});