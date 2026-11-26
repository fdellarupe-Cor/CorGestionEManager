# Despliegue a produccion - Corv48UNIFV5

Base de produccion: `dbcordatv485Cor`.  Lista de pasos y de archivos. Todo lo que dice "manual" lo hace el usuario;
nada de esto se ejecuta ni se copia automaticamente.

## 0. Antes de empezar
- Backup completo de `dbcordatv485Cor` (o confirmar que se trabaja sobre la copia restaurada).
- Copia de la carpeta web de produccion (`...\webintranet\Corv48UNIFV5`) antes de pisar archivos.
- Probar primero todo el circuito (puntos 1 a 4) sobre una copia restaurada de la base.

## 1. Scripts SQL (carpeta `DB`), en este orden
| # | Script | Que hace | Notas |
|---|--------|----------|-------|
| 1 | `Migracion_00_Diagnostico.sql` | Solo consultas | Revisar cada bloque antes de seguir |
| 2 | `Migracion_01_Estructura.sql` | Columnas nuevas, `CG_FERIADOS`, `CG_INCIDENTES.AVANCES` a `nvarchar(max)`, 4 objetos nuevos y 10 procedimientos modificados | Repetible |
| 2b | `Cambios_Avances_Estado.sql`, `Cambios_Inicio_IdAvance.sql`, `Cambios_Avances_TipoReclamo.sql` | Posteriores al export con que se genero el 01: estado del avance ("Pendiente" al agregar y cambio de estado), id del avance en inicio.asp y tipo "Reclamo" | Ejecutarlos DESPUES del 01 (reemplazan `sp_cg_Registro_Avance`, `sp_cg_ConsultarEntidad` y `sp_cg_Inicio`); `Cambios_Avances_Insistir.sql` ya esta incluido en el 01 |
| 2c | `Permisos_Alta_Formularios.sql` | Da de alta en E-Manager (Atrm/AtrmUsuarios/AtrmGrupos) las pantallas que ahora usan `getPermisosAtrm` (CG_APL, CG_CLIENTES, ...), copiando los permisos de CG_INCIDENTES | Ejecutar ANTES de subir los .asp: `sp_cg_Permisos` da acceso total a una pantalla sin filas. Termina en `ROLLBACK` |
| 2d | `Cambios_Agenda_Invitacion.sql` | Agenda: columnas ENLACE_REUNION, FEC_ULT_ENVIO, USU_ULT_ENVIO, SEQ_INVITACION y modos nuevos de `sp_cg_AgendaExt` | Ejecutar ANTES de subir `agenda_ext.inc` (si no, guardar una reunion falla por el parametro @Enlace) |
| 2d | `Cambios_Log_Acciones.sql` | Tabla `CG_LOG` (log de acciones: un registro por usuario + sesion + elemento, con la accion de mayor nivel) y `sp_cg_Log_Registrar` / `sp_cg_Log_Purgar` | Ejecutar ANTES de subir los .asp (sin la tabla, `CgLog` no registra pero no rompe nada). Termina en `ROLLBACK`. Purga: `EXEC sp_cg_Log_Purgar @Meses = 12` (la ejecuta el usuario) |
| 3 | `Migracion_02_Datos.sql` | Vacia las `CG_` de prueba y carga los datos de las 17 tablas originales | Termina en `ROLLBACK`: revisar los resultados y cambiar a `COMMIT TRAN` |
| 4 | `Migracion_04_Validacion.sql` | Solo consultas: conteos, referencias, responsables, contenido | Todos los controles del bloque A deben dar OK (o un REVISAR explicado) |
| 5 | `Migracion_03_Cierre.sql` | Renombra las 17 originales a `*_OLD` | Solo cuando todo este validado y la aplicacion probada |

Scripts de cambios individuales (`Cambios_*.sql`): sus cambios ya estan incluidos en `Migracion_01_Estructura.sql` (texto vigente en
desarrollo al 23/10), salvo los tres de la fila 2b, que son posteriores y se ejecutan despues del 01.

## 2. Aplicacion (carpeta `Corgestion`)
**Antes de copiar: cambiar la conexion (MANUAL, no se toca en el proyecto).**
`IncFile/funciones.inc`, linea 11: `scon1` tiene `Initial Catalog=Dbcordatv48_UNIF485`; en produccion debe apuntar a `dbcordatv485Cor`.
Todas las paginas abren la base con `Cnn.open scon1`.

Archivos modificados desde el 19/10/2026 12:00 (comparar con los de produccion y subir los que difieran):
- `actualizar_v2.asp` (guarda el estado del avance y el avance de origen), `detalles2.asp`
- `incidentes.asp`, `proyectos.asp`, `proyRelac.asp`, `consultas.asp` (las consultas de Consultas paginan de a 10: incidentes, proyectos, APL, proyectos con incidentes y por responsable)
- `vistaconsulta.asp` (vista de consulta de incidentes, proyectos y APL: llamada desde Consultas; ahora con seguridad E-Manager y permiso de consulta)
- `agenda.asp` (filtros de la lista, "Volver" a `inicio.asp`)
- `agenda_semana.asp` (vista mensual, licencias con color propio, descripcion de la reunion)
- `detalles.asp` (sin "Estado cliente", prioridad en edicion, "Relacionado a proyecto" opcional, estado del avance, "Insistir", apertura del avance desde inicio)
- `agregar2.asp` y `IncFile/datosAgregar.inc` (ambiente y proyecto relacionado en el alta de incidentes)
- `inicio.asp` (testing y solicitudes pendientes con acceso directo al avance, recordatorios del usuario)
- `agenda.asp` (incluye `agenda_invitacion.inc`), `IncFile/agenda_invitacion.inc` (NUEVO: invitacion por mail con .ics y enlaces a Google Calendar / Outlook / Teams; requiere `Session("SMTPServer")`)
- `IncFile/agenda_ext.inc` (descripcion breve, participantes sin duplicados, "Alerta" de cruce)
- `IncFile/datosActualizar.inc` (estado del avance y avance de origen de "Insistir")
- Log de acciones (`CgLog`): `IncFile/funciones.inc`, `IncFile/datosEliminar.inc`, `IncFile/agenda_ext.inc`, `actualizar_v2.asp`, `detalles.asp`, `detalles2.asp`, `vistaconsulta.asp`, `agregar2.asp`, `agregar3.asp`, y para los maestros: `agregar1.asp`, `clientes`, `aplicativos`, `browsers`, `estados`, `prioridades`, `tipoavance`, `viapedido`, `contactos`, `corresp`, `clientesDet`, `detalleCliente`, `detalleAplic`
- `IncFile/funciones.inc` (ayuda de combos, pop-up de lectura; cambiar `scon1` como se indica arriba)
- `assets/css/corgestion-ui.css` (con el `?v=20261028` ya referenciado desde `funciones.inc`)

No subir: `feriados_test.asp` (pagina de diagnostico con opcion de reiniciar feriados), ni los `*.v1.*` (respaldos), ni las variantes
(`*Copilot*`, `*Mimo*`, `*-ORIG*`, etc.).

Todas las pantallas (apl, detalles2, clientes, clientesDet, detalleCliente, aplicativos, detalleAplic, browsers, contactos, corresp, estados, prioridades, tipoavance, viapedido, licencias, agenda, consultas, estadisticas, estadisticasDet, reportes, permisosDet, todo, actualizar_v2) pasaron de `getPermisos` a `getPermisosAtrm`: subirlas junto con el script 2c.

## 3. Primer arranque en produccion
- `CG_FERIADOS` esta vacia: se llena sola la primera vez que se abre `agenda_semana.asp` (consulta la API de feriados).
- E-Manager: confirmar el menu de la agenda semanal (`agenda_semana.asp`) y los permisos de las pantallas `CG_INCIDENTES` y `CG_PROYECTOS`.
- Prueba funcional con un usuario de cada perfil:
  - Inicio: casos asignados, testing pendientes, solicitudes pendientes, recordatorios.
  - Incidentes: listado, alta (con ambiente y proyecto), detalle, edicion, avances.
  - Proyectos: listado, alta, detalle, edicion.
  - Agenda: semanal y mensual, licencias, reunion con participantes, aviso de cruce.
  - Maestros: clientes, aplicativos, contactos, APL, estadisticas.

## 4. Cierre
- Ejecutar `Migracion_03_Cierre.sql` (renombra a `*_OLD`, sin borrar).
- Pasado un tiempo prudente sin incidencias, eliminar las `*_OLD`.
- Despues del renombre, `Migracion_02_Datos.sql` ya no se puede repetir sin la reversa que trae `Migracion_03_Cierre.sql`.
