// assets/js/relaciones.js

function capturarYProcesar() {
    // El popup (armaCargaParam.asp?param=IOP) solo tiene estos 2 campos:
    //   valor1 = ID del incidente/proyecto a relacionar (idRelac)
    //   valor2 = tipo, 'I' o 'P' (tipoRelac)
    // El ID del APL NO viene del popup: ya está en la página en <div id='nid'>
    // (se renderiza en el modo "Actualizar APL" de detalles2.asp).
    const idRelac = document.getElementById('valor1').value;
    const tipoRelac = document.getElementById('valor2').value;
    const nidEl = document.getElementById('nid');
    const idApl = nidEl ? nidEl.textContent.trim() : '';

    if (!idApl || !idRelac) {
        alert('Por favor complete todos los campos.');
        return;
    }

    // El popup pide el ID del incidente/proyecto a relacionar: el preview
    // debe confirmar los datos de ESA entidad, no los del APL actual.
    procesarGuardadoAPL(idApl, idRelac, tipoRelac, false);
}

// Usado desde detalles.asp (pantalla de Incidente/Proyecto, boton
// "RELACIONAR APL"). Aca es al reves: el popup (armaCargaParam.asp?param=APL)
// solo tiene un campo, valor1 = ID del APL a relacionar (idApl). El ID del
// incidente/proyecto actual ya esta en la pagina en <div id='nid'>, y el
// tipo ('I' o 'P') lo resuelve el ASP segun si se esta editando un incidente
// o un proyecto (no es un campo del formulario).
function capturarYProcesarApl(tipoRelac) {
    const idApl = document.getElementById('valor1').value;
    const nidEl = document.getElementById('nid');
    const idRelac = nidEl ? nidEl.textContent.trim() : '';

    if (!idApl || !idRelac) {
        alert('Por favor complete todos los campos.');
        return;
    }

    // Aca es al reves: el usuario escribio el ID del APL a relacionar. El
    // preview debe confirmar los datos del APL (Cliente/Aplicativo/Nombre),
    // no los del incidente/proyecto que ya esta abierto en la pagina.
    procesarGuardadoAPL(idApl, idRelac, tipoRelac, true);
}

function procesarGuardadoAPL(idApl, idRelac, tipoRelac, confirmarApl) {
    // 1. Pre-consulta de datos para el popup
    //    - confirmarApl=true  (desde incidente/proyecto): el dato a confirmar
    //      es el APL -> dato2=idApl, dato3='A' (modo APL del SP).
    //    - confirmarApl=false (desde APL): el dato a confirmar es la entidad
    //      -> dato2=idRelac, dato3=tipoRelac.
    const previewDato2 = confirmarApl ? idApl : idRelac;
    const previewDato3 = confirmarApl ? 'A' : tipoRelac;

    fetch(`agreParam.asp?param=APL&dato1=${idApl}&dato2=${previewDato2}&dato3=${previewDato3}&accion=preview`)
        .then(response => response.text())
        .then(data => {
            if (data.startsWith('OK|')) {
                const res = data.split('|');
                const cliente = res[1];
                const aplicativo = res[2];
                const detalle = res[3];

                const mensaje = `¿Confirma los siguientes datos?\n\n` +
								`- DT: ${idApl} -  ${idRelac} - ${tipoRelac} \n` +
                                `- Cliente: ${cliente}\n` +
                                `- Aplicativo: ${aplicativo}\n` +
                                `- Detalle: ${detalle}`;

                if (confirm(mensaje)) {
                    ejecutarInsercion(idApl, idRelac, tipoRelac);
                }
            } else {
                alert(data.replace('ERROR: ', ''));
            }
        })
        .catch(error => alert('Error en la comunicación con el servidor: ' + error));
}

function ejecutarInsercion(idApl, idRelac, tipoRelac) {
    // 2. Confirmación aceptada: Inserción en DB
    fetch(`agreParam.asp?param=APL&dato1=${idApl}&dato2=${idRelac}&dato3=${tipoRelac}&accion=guardar`)
        .then(response => response.text())
        .then(res => {
            if (res.trim() === 'SUCCESS') {
                alert('Registro guardado correctamente.');
                location.reload();
            } else {
                alert(res.replace('ERROR: ', ''));
            }
        })
        .catch(error => alert('Error al guardar el registro: ' + error));
}