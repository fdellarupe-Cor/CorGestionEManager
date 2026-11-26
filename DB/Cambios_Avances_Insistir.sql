/* =====================================================================
   Cambios_Avances_Insistir.sql  -  trazabilidad de avances (CG_AVANCES.ID_ANT) para "Insistir"

   ESTADO: SCRIPT PARA REVISAR. NO SE EJECUTO.
   Base destino: Dbcordatv48_UNIF485     Compatibilidad: nivel 100
   Requisito   : la columna ya existe (ALTER TABLE CG_AVANCES ADD ID_ANT INT NULL).
   Base de partida: texto de los procedimientos en DB/Corv48UNIFV5_DB.sql del 22/10/2026 11:28
                    (sp_cg_ConsultarEntidad parte de la version que ya incluye el ambiente).
   OJO: cada ALTER reemplaza el procedimiento completo con ese texto. Si los modificaste despues,
        volve a exportarlos antes de aplicar este script.

   IDEA: "Insistir" es un acceso directo para crear un avance NUEVO indicando que se relaciona con uno
   anterior (ej.: insistir sobre una solicitud), para tener trazabilidad. Solo se preparan los procedimientos;
   las pantallas NO se modifican todavia.

   1) sp_cg_Registro_Avance : parametro nuevo @IdAvanceAnt INT = NULL (al final, opcional). Si viene, se valida
      que el avance exista y sea del mismo caso (mismo CODOBJETO y TIPOOBJETO); error 50004 si no. Se guarda en
      CG_AVANCES.ID_ANT. Sin el parametro, todo sigue igual.
   2) sp_cg_ConsultarEntidad (rama 'comentarios', historial de avances): dos columnas nuevas AL FINAL,
      ID e ID_ANT. En las filas iniciales (la del incidente o proyecto) son NULL; en los avances, los de
      CG_AVANCES. Las columnas existentes no cambian de lugar. Es el ID que P1 espera en el boton "Insistir".
      El resto del procedimiento queda identico.
   PENDIENTE PARA CUANDO SE ACTIVE LA PANTALLA (no se hace ahora):
      - detalles.asp: boton "Insistir" y campo oculto avance_origen (ya existen en P1).
      - actualizar_v2.asp y IncFile/datosActualizar.inc (RegistrarAvance): pasar request.form("avance_origen")
        como @IdAvanceAnt.
   SI NO SE EJECUTA : las pantallas siguen funcionando igual.
   REVERSA    : volver a ejecutar la definicion original de ambos procedimientos (DB/Corv48UNIFV5_DB.sql).
   ===================================================================== */

USE [Dbcordatv48_UNIF485]
GO

ALTER PROCEDURE [dbo].[sp_cg_Registro_Avance]
    @tabla         VARCHAR(20),
    @id            INT,
    @descripcion   NVARCHAR(4000),
    @fecha         INT,
    @hora          INT,
    @responsable   NVARCHAR(20),
    @idusuario     NVARCHAR(100),
    @IdTipoAvance  INT = NULL,           -- CG_TIPO_AVANCE.ID (opcional)
    @IdAvanceAnt   INT = NULL            -- [INSISTIR] CG_AVANCES.ID del avance anterior al que se relaciona (opcional, trazabilidad)
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;

    IF @tabla NOT IN ('incidente', 'proyecto')
        THROW 50001, 'Tabla de avance no valida.', 1;

    IF UPPER(@responsable) NOT IN ('FUNCIONAL', 'TECNICO', 'CLIENTE')
        THROW 50002, 'Responsable de avance no valido.', 1;

    IF @IdTipoAvance IS NOT NULL AND NOT EXISTS (SELECT 1 FROM dbo.CG_TIPO_AVANCE WHERE ID = @IdTipoAvance)
        THROW 50003, 'El tipo de avance indicado no existe.', 1;

    -- [INSISTIR] El avance anterior debe existir y ser del mismo caso (mismo registro y mismo tipo de objeto)
    IF @IdAvanceAnt IS NOT NULL AND NOT EXISTS (SELECT 1 FROM dbo.CG_AVANCES
                                                 WHERE ID = @IdAvanceAnt AND codobjeto = @id
                                                   AND tipoobjeto = CASE WHEN @tabla = 'incidente' THEN 'inc' ELSE 'proy' END)
        THROW 50004, 'El avance anterior indicado no existe o no pertenece a este caso.', 1;

    BEGIN TRANSACTION;

    INSERT INTO dbo.CG_AVANCES
        (codobjeto, tipoobjeto, orden, fecha, hora, descripcion, idusuario_seg, responsable, tipo, FechaRegistro, ID_ANT)
    SELECT @id, CASE WHEN @tabla = 'incidente' THEN 'inc' ELSE 'proy' END,
           ISNULL(MAX(orden), 0) + 1, @fecha, @hora, @descripcion, @idusuario,
           UPPER(@responsable), CAST(@IdTipoAvance AS NVARCHAR(50)),
           DATEDIFF(DAY, CONVERT(date, '19000101'), CONVERT(date, GETDATE())),
           @IdAvanceAnt
    FROM dbo.CG_AVANCES
    WHERE codobjeto = @id
      AND tipoobjeto = CASE WHEN @tabla = 'incidente' THEN 'inc' ELSE 'proy' END;

    IF @tabla = 'incidente'
        UPDATE dbo.CG_INCIDENTES SET responsable = UPPER(@responsable) WHERE id = @id;
    ELSE
        UPDATE dbo.CG_PROYECTOS SET responsable = UPPER(@responsable) WHERE id = @id;

    COMMIT TRANSACTION;
END;

GO

ALTER PROCEDURE [dbo].[sp_cg_ConsultarEntidad]
    @Entidad         VARCHAR(15)  = 'incidente',   -- 'incidente' | 'proyecto' | 'apl' | 'comentarios' | 'puedeact'
    @Modo            VARCHAR(15)  = 'lista',       -- 'lista' | 'consulta' | 'detalle' | 'edicion' | 'porproyecto'
    @Id              INT          = NULL,          -- dato unico (detalle/edicion/porproyecto/puedeact)
    @IdCliente       INT          = 0,             -- filtro lista/consulta
    @IdAplicativo    INT          = 0,             -- filtro lista/consulta
    @Titulo          NVARCHAR(255)= NULL,          -- filtro lista/consulta (busca titulo y detalles en incidentes)
    @Estado          INT          = 0,             -- filtro consulta: 0=all, 1=OK, 2=PEND, 10=BAJA
    @Fecha           INT          = NULL,          -- filtro consulta: fecha_ingreso
    @Orden           NVARCHAR(50) = NULL,          -- 'titulo'|'cliente'|'aplicativos'|'nom'|'sec'
    @TextoBusqueda   NVARCHAR(100)= NULL,          -- filtro proyectos lista (titulo/descripcion)
    @EstadoFiltro    INT          = 0,             -- filtro proyectos lista: 0=all, 1=OK, 2=PEND, 10=BAJA
    @FechaIngreso    INT          = NULL,          -- filtro proyectos lista
    @RegistroInicial INT          = 0,             -- paginacion proyectos lista
    @CantidadMostrar INT          = 12,            -- paginacion proyectos lista
    @CodObjeto       INT          = NULL,          -- comentarios: codigo del registro
    @TipoObjeto      VARCHAR(10)  = NULL,          -- comentarios: 'inc'|'cmd'
    @Cons            VARCHAR(20)  = '',            -- apl consulta: 'cli'|'apli'|filtro libre
    @IdTip           INT          = 0,             -- apl: filtro tipo_apl
    @Asp             VARCHAR(100) = '',            -- apl: filtro motivo_apl
    @Fec             INT          = 0,             -- apl: filtro fecha_creacion
    @Respon          NVARCHAR(100)= NULL           -- incidente lista: filtro por responsable (gestion o tec)
AS
BEGIN
    SET NOCOUNT ON;

    /* =====================================================================
       INCIDENTES
       ===================================================================== */
    IF @Entidad = 'incidente' AND @Modo = 'lista'
    BEGIN
        SELECT a.id, b.descripcion AS cliente, c.descripcion AS prioridad,
               f.descripcion AS estado_desa, g.descripcion AS estado_func,
               a.fecha_ingreso, a.respon_gestion_seg AS respon_gestion, a.respon_tec_seg AS respon_tec,
               a.respon_gestion_seg AS respon_gestion_nombre, a.respon_tec_seg AS respon_tec_nombre,
               a.CREACION, a.titulo, i.descripcion AS aplicativo, a.AMBIENTE
        FROM CG_INCIDENTES a
        INNER JOIN CG_CLIENTES b ON (a.id_cliente = b.id)
        INNER JOIN CG_PRIORIDADES c ON (a.prioridad = c.id)
        INNER JOIN CG_ESTADOS f ON (a.estado_desa = f.id)
        INNER JOIN CG_ESTADOS g ON (a.estado_func = g.id)
        --INNER JOIN CG_COR_RESPONSABLES h ON (a.respon_gestion_seg = h.id_seg)
        INNER JOIN CG_APLICATIVOS i ON (a.ID_APLIC = i.ID)
        --LEFT JOIN CG_COR_RESPONSABLES j ON (a.respon_tec_seg = j.id_seg)
        WHERE a.ESTADO = '1'
          AND (@Id IS NULL OR a.id = @Id)
          AND (@IdCliente = 0 OR (@IdCliente = 4 AND a.id_cliente IN (4,38,39,40,41)) OR (@IdCliente <> 4 AND a.id_cliente = @IdCliente))
          AND (@IdAplicativo = 0 OR a.id_aplic = @IdAplicativo)
          AND (@Titulo IS NULL OR @Titulo = '' OR (a.titulo LIKE '%' + @Titulo + '%' OR a.detalles LIKE '%' + @Titulo + '%'))
          AND (@Estado = 0 OR
              (@Estado = 1 AND a.estado_desa = 1 AND a.estado_func = 1) OR
              (@Estado = 2 AND (a.estado_desa = 2 OR a.estado_func = 2)) OR
              (@Estado = 10 AND (a.estado_desa = 10 OR a.estado_func = 10)) OR
              (@Estado NOT IN (1,2,10) AND (a.estado_desa = @Estado OR a.estado_func = @Estado)))
          AND (@Respon IS NULL OR a.respon_gestion_seg = @Respon OR a.respon_tec_seg = @Respon)
        ORDER BY
            CASE WHEN @Orden = 'titulo' THEN a.titulo END ASC,
            CASE WHEN @Orden = 'cliente' THEN a.id_cliente END ASC,
            CASE WHEN @Orden = 'aplicativos' THEN a.id_aplic END ASC,
            CASE WHEN @Orden = 'aplicativo' THEN a.id END ASC,
            CASE WHEN @Orden = 'id' THEN a.id END DESC,
            CASE WHEN @Orden IS NULL OR @Orden = '' THEN a.FECHA_INGRESO END DESC,
            a.PRIORIDAD ASC;
    END

    ELSE IF @Entidad = 'incidente' AND @Modo = 'consulta'
    BEGIN
        SELECT a.id, a.titulo, b.descripcion AS cliente, f.descripcion AS estado_desa,
               g.descripcion AS estado_func, a.fecha_ingreso, i.descripcion AS aplicativo, 
               j.descripcion AS estado_cliente, TIEMPO_FUNC , TIEMPO_DESA, a.AMBIENTE
        FROM CG_INCIDENTES a
        INNER JOIN CG_CLIENTES b ON (a.id_cliente = b.id)
        INNER JOIN CG_PRIORIDADES c ON (a.prioridad = c.id)
        LEFT JOIN CG_ESTADOS f ON (a.estado_desa = f.id)
        LEFT JOIN CG_ESTADOS g ON (a.estado_func = g.id)
       -- LEFT JOIN vw_UserGrupos h on USRNET = a.respon_gestion_seg
        --INNER JOIN CG_COR_RESPONSABLES h ON (a.respon_gestion_seg = h.id_seg)
        INNER JOIN CG_APLICATIVOS i ON (a.ID_APLIC = i.ID)
        LEFT JOIN CG_ESTADOS j ON (a.ESTADO_CLIENTE = j.id)
        WHERE a.ESTADO = 1
          AND (@IdCliente = 0 OR a.id_cliente = @IdCliente)
          AND (@Titulo IS NULL OR @Titulo = '' OR a.titulo LIKE '%' + @Titulo + '%')
          AND (@Estado = 0 OR
              (@Estado = 1 AND a.estado_desa = 1 AND a.estado_func = 1) OR
              (@Estado = 2 AND (a.estado_desa = 2 OR a.estado_func = 2)) OR
              (@Estado = 10 AND (a.estado_desa = 10 OR a.estado_func = 10)) OR
              (@Estado NOT IN (1,2,10) AND (a.estado_desa = @Estado OR a.estado_func = @Estado)))
          AND (@IdAplicativo = 0 OR i.id = @IdAplicativo)
          AND (@Fecha IS NULL OR a.fecha_ingreso = @Fecha)
        ORDER BY a.id DESC;
    END

    ELSE IF @Entidad = 'incidente' AND @Modo = 'todo'
    BEGIN
        SELECT a.id, b.descripcion, c.descripcion, f.descripcion, g.descripcion,
               a.fecha_ingreso, a.respon_gestion_seg AS respon_gestion, a.respon_tec_seg AS respon_tec, a.respon_gestion_seg, a.respon_tec_seg,
               a.CREACION, a.titulo, (f.orden + g.orden) AS suma, a.AMBIENTE
        FROM CG_INCIDENTES a
        INNER JOIN CG_CLIENTES b ON (a.id_cliente = b.id)
        INNER JOIN CG_PRIORIDADES c ON (a.prioridad = c.id)
        LEFT JOIN CG_ESTADOS f ON (a.estado_desa = f.id)
        LEFT JOIN CG_ESTADOS g ON (a.estado_func = g.id)
        --INNER JOIN CG_COR_RESPONSABLES h ON (a.respon_gestion_seg = h.id_seg)
        --LEFT JOIN CG_COR_RESPONSABLES j ON (a.respon_tec_seg = j.id_seg)
        WHERE a.ESTADO = '1'
          AND (f.id = 2 OR g.id = 2)
          AND (@IdCliente = 0 OR (@IdCliente = 4 AND a.id_cliente IN (4,38,39,40,41)) OR (@IdCliente <> 4 AND a.id_cliente = @IdCliente))
          AND (@Titulo IS NULL OR @Titulo = '' OR (a.titulo LIKE '%' + @Titulo + '%' OR a.detalles LIKE '%' + @Titulo + '%'))
        ORDER BY
            CASE WHEN @Orden = 'titulo' THEN a.titulo END ASC,
            CASE WHEN @Orden = 'cliente' THEN a.id_cliente END ASC,
            CASE WHEN @Orden IS NULL OR @Orden = '' THEN a.CREACION END ASC,
            CASE WHEN @Orden IS NULL OR @Orden = '' THEN a.prioridad END ASC,
            CASE WHEN @Orden IS NULL OR @Orden = '' THEN (f.orden + g.orden) END DESC;
    END

    ELSE IF @Entidad = 'incidente' AND @Modo = 'detalle'
    BEGIN
        SELECT a.id, b.descripcion, c.nombre + ' ' + c.apellido AS contacto_nombre, d.descripcion, a.detalles,
               e.descripcion, f.USRNAME AS respon_func, g.USRNAME AS respon_tec,
               a.req_presup, h.descripcion, i.descripcion, a.fecha_ingreso, j.descripcion,
               k.descripcion, a.id_proyecto, c.id AS id_contacto, a.ok_desarrollo, a.ok_funcional,
               a.ok_cliente, a.ruta, a.creacion, a.avances, a.titulo, l.nombre AS nombre_apl, a.id_apl,
               m.solic_nro, m.titulo AS titulo_proy, n.descripcion AS estado_cliente, a.TIEMPO_FUNC , a.TIEMPO_DESA, a.AMBIENTE
        FROM CG_INCIDENTES a
        INNER JOIN CG_CLIENTES b ON (a.id_cliente = b.id)
        INNER JOIN CG_CONTACTOS c ON (a.contacto = c.id)
        INNER JOIN CG_APLICATIVOS d ON (a.id_aplic = d.id)
        INNER JOIN CG_PRIORIDADES e ON (a.prioridad = e.id)
        LEFT JOIN vw_UserGrupos f on f.USRNET = a.respon_gestion_seg and f.grupo = 'FUNC'
        LEFT JOIN vw_UserGrupos g on g.USRNET = a.respon_tec_seg and g.grupo = 'TEC'
        --INNER JOIN CG_COR_RESPONSABLES f ON (a.respon_gestion_seg = f.id_seg AND f.funcion = 'FUNC')
        --LEFT JOIN CG_COR_RESPONSABLES g ON (a.respon_tec_seg = g.id_seg AND g.funcion = 'TEC')
        LEFT JOIN CG_ESTADOS h ON (a.estado_desa = h.id)
        LEFT JOIN CG_ESTADOS i ON (a.estado_func = i.id)
        INNER JOIN CG_BROWSERS j ON (a.browser = j.id)
        INNER JOIN CG_VIA_PEDIDO k ON (a.via_pedido = k.id)
        LEFT JOIN CG_APL l ON (a.id_apl = l.id)
        LEFT JOIN CG_PROYECTOS m ON (a.id_proyecto = m.id)   -- [P1->P2] titulo y solicitud del proyecto
        LEFT JOIN CG_ESTADOS n ON (a.ESTADO_CLIENTE = n.id)
        WHERE a.id = @Id;
    END

    ELSE IF @Entidad = 'incidente' AND @Modo = 'edicion'
    BEGIN
        SELECT a.id, b.id, b.descripcion, c.id, c.nombre + ' ' + c.apellido, d.id, d.descripcion,
               e.id, a.detalles, e.descripcion, f.USRNET, f.USRNAME,
               ISNULL(g.USRNET, N' '), ISNULL(g.USRNAME, ' '), a.req_presup,
               h.id, h.descripcion, i.id, i.descripcion, a.fecha_ingreso, j.id, j.descripcion,
               k.id, k.descripcion, a.id_proyecto, a.ok_cliente, a.ruta, a.avances, a.titulo,
               a.id_apl, m.solic_nro, m.titulo, a.creacion,  isnull(n.id,0) , n.descripcion  , a.TIEMPO_FUNC , a.TIEMPO_DESA, a.AMBIENTE
        FROM CG_INCIDENTES a
        INNER JOIN CG_CLIENTES b ON (a.id_cliente = b.id)
        LEFT JOIN CG_CONTACTOS c ON (a.contacto = c.id)
        INNER JOIN CG_APLICATIVOS d ON (a.id_aplic = d.id)
        INNER JOIN CG_PRIORIDADES e ON (a.prioridad = e.id)
        LEFT JOIN vw_UserGrupos f on f.USRNET = a.respon_gestion_seg AND f.GRUPO = 'FUNC'
        LEFT JOIN vw_UserGrupos g on g.USRNET = a.respon_tec_seg AND g.GRUPO = 'TEC'
        --INNER JOIN CG_COR_RESPONSABLES f ON (a.respon_gestion_seg = f.id_seg AND f.funcion = 'FUNC')
        --LEFT JOIN CG_COR_RESPONSABLES g ON (a.respon_tec_seg = g.id_seg AND g.funcion = 'TEC')
        LEFT JOIN CG_ESTADOS h ON (a.estado_desa = h.id)
        LEFT JOIN CG_ESTADOS i ON (a.estado_func = i.id)
        INNER JOIN CG_BROWSERS j ON (a.browser = j.id)
        INNER JOIN CG_VIA_PEDIDO k ON (a.via_pedido = k.id)
        LEFT JOIN CG_PROYECTOS m ON (a.id_proyecto = m.id)
        LEFT JOIN CG_ESTADOS n ON (a.ESTADO_CLIENTE = n.id)
        WHERE a.id = @Id;
    END

    ELSE IF @Entidad = 'incidente' AND @Modo = 'porproyecto'
    BEGIN
        SELECT id, titulo, estado_desa, estado_func
        FROM CG_INCIDENTES
        WHERE id_proyecto = @Id;
    END

    /* =====================================================================
       PROYECTOS
       ===================================================================== */
    ELSE IF @Entidad = 'proyecto' AND @Modo = 'lista'
    BEGIN
        WITH ProyectosBase AS (
            SELECT a.id, a.solic_nro, a.titulo, a.descripcion,
                   b.id AS id_cliente, b.descripcion AS cliente_desc,
                   c.descripcion AS prioridad_desc, a.prioridad AS prioridad_id,
                   f.descripcion AS estado_desa_desc, f.id AS estado_desa_id,
                   g.descripcion AS estado_func_desc, g.id AS estado_func_id,
                   a.fecha_ingreso, a.respon_gestion_seg AS respon_gestion_nombre,
                   a.respon_tec_seg AS respon_tec_nombre, a.respon_tec_seg AS respon_tec, a.respon_gestion_seg AS respon_gestion,
                   i.descripcion AS aplicativo_desc, i.id AS aplicativo_id, a.tipo_proy
            FROM CG_PROYECTOS a
            INNER JOIN CG_CLIENTES b ON a.id_cliente = b.id
            INNER JOIN CG_PRIORIDADES c ON a.prioridad = c.id
            LEFT JOIN CG_ESTADOS f ON a.estado_desa = f.id
            LEFT JOIN CG_ESTADOS g ON a.estado_func = g.id
            --INNER JOIN CG_COR_RESPONSABLES h ON a.respon_gestion_seg = h.id_seg
            --INNER JOIN CG_COR_RESPONSABLES j ON a.respon_tec_seg = j.id_seg
            LEFT JOIN CG_APLICATIVOS i ON a.id_aplic = i.id
            WHERE a.ESTADO = '1'
        ),
        ProyectosFiltrados AS (
            SELECT *,
                   COUNT(*) OVER() AS TotalRegistros
            FROM ProyectosBase
            WHERE (@IdCliente = 0 OR id_cliente = @IdCliente)
              AND (@IdAplicativo = 0 OR aplicativo_id = @IdAplicativo)
              AND (@TextoBusqueda IS NULL OR titulo LIKE '%' + @TextoBusqueda + '%' OR descripcion LIKE '%' + @TextoBusqueda + '%')
              AND (@EstadoFiltro = 0 OR
                  (@EstadoFiltro = 1 AND estado_desa_id = 1 AND estado_func_id = 1) OR
                  (@EstadoFiltro = 2 AND (estado_desa_id = 2 OR estado_func_id = 2)) OR
                  (@EstadoFiltro = 10 AND (estado_desa_id = 10 OR estado_func_id = 10)))
              AND (@FechaIngreso IS NULL OR fecha_ingreso = @FechaIngreso)
              AND (@Respon IS NULL OR respon_gestion = @Respon OR respon_tec = @Respon)   -- [P1->P2] filtro por responsable
                ),
        ProyectosNumerados AS (
            SELECT *,
                   ROW_NUMBER() OVER (
                       ORDER BY
                           CASE WHEN @Orden = 'titulo'      THEN titulo END ASC,
                           CASE WHEN @Orden = 'cliente'     THEN cliente_desc END ASC,
                           CASE WHEN @Orden = 'aplicativos' THEN aplicativo_desc END ASC,
                           CASE WHEN @Orden IS NULL OR @Orden = '' THEN id END DESC,
                           id DESC
                   ) AS RN
            FROM ProyectosFiltrados
        )
        SELECT id, solic_nro, titulo, descripcion, id_cliente, cliente_desc,
               prioridad_desc, prioridad_id, estado_desa_desc, estado_desa_id,
               estado_func_desc, estado_func_id, fecha_ingreso,
               respon_gestion_nombre, respon_tec_nombre, respon_tec, respon_gestion,
               aplicativo_desc, aplicativo_id, tipo_proy, TotalRegistros
        FROM ProyectosNumerados
        WHERE RN > @RegistroInicial
          AND RN <= @RegistroInicial + @CantidadMostrar
        ORDER BY RN;

    END

    ELSE IF @Entidad = 'proyecto' AND @Modo = 'detalle'
    BEGIN
        SELECT a.id, a.solic_nro, a.titulo, a.descripcion, b.descripcion, c.nombre + ' ' + c.apellido,
               d.descripcion, f.USRNAME, g.USRNAME, a.fecha_ingreso,
               a.fecha_entrega, h.descripcion, i.descripcion, e.descripcion, a.tiempo_relevamiento,
               a.tiempo_diseno, a.tiempo_desarr, a.tiempo_test, j.descripcion, k.descripcion,
               a.req_presup, SUM(a.tiempo_diseno + a.tiempo_desarr + a.tiempo_relevamiento + a.tiempo_test),
               c.id, a.ok_desarrollo, a.ok_funcional, a.ok_cliente, a.ruta, a.avances, l.nombre, a.id_apl, a.tipo_proy
        FROM CG_PROYECTOS a
        INNER JOIN CG_CLIENTES b ON (a.id_cliente = b.id)
        INNER JOIN CG_CONTACTOS c ON (a.contacto = c.id)
        INNER JOIN CG_APLICATIVOS d ON (a.id_aplic = d.id)
        INNER JOIN CG_PRIORIDADES e ON (a.prioridad = e.id)
        LEFT JOIN vw_UserGrupos f on f.USRNET = a.respon_gestion_seg AND f.GRUPO = 'FUNC'
        LEFT JOIN vw_UserGrupos g on g.USRNET = a.respon_tec_seg AND g.GRUPO = 'TEC'
        --INNER JOIN CG_COR_RESPONSABLES f ON (a.respon_gestion_seg = f.id_seg AND f.funcion = 'FUNC')
        --INNER JOIN CG_COR_RESPONSABLES g ON (a.respon_tec_seg = g.id_seg AND g.funcion = 'TEC')
        LEFT JOIN CG_ESTADOS h ON (a.estado_desa = h.id)
        LEFT JOIN CG_ESTADOS i ON (a.estado_func = i.id)
        INNER JOIN CG_BROWSERS j ON (a.browser = j.id)
        INNER JOIN CG_VIA_PEDIDO k ON (a.via_pedido = k.id)
        LEFT JOIN CG_APL l ON (a.id_apl = l.id)
        WHERE a.id = @Id
        GROUP BY a.id, a.solic_nro, a.titulo, a.descripcion, b.descripcion, c.nombre + ' ' + c.apellido,
                 d.descripcion, f.USRNAME, g.USRNAME, a.fecha_ingreso,
                 a.fecha_entrega, h.descripcion, i.descripcion, e.descripcion, a.tiempo_relevamiento,
                 a.tiempo_diseno, a.tiempo_desarr, a.tiempo_test, j.descripcion, k.descripcion,
                 a.req_presup, a.id_cliente, a.contacto, a.id_aplic, a.respon_gestion_seg, a.respon_tec_seg,
                 a.estado_func, a.estado_desa, a.prioridad, a.browser, a.via_pedido, c.id,
                 a.ok_desarrollo, a.ok_funcional, a.ok_cliente, a.ruta, a.avances, l.nombre, a.id_apl,a.tipo_proy;
    END

    ELSE IF @Entidad = 'proyecto' AND @Modo = 'edicion'
    BEGIN
        SELECT a.id, a.solic_nro, a.titulo, a.descripcion, b.descripcion, c.nombre + ' ' + c.apellido,
               d.descripcion, f.USRNAME, g.USRNAME, a.fecha_ingreso,
               a.fecha_entrega, h.descripcion, i.descripcion, e.descripcion, a.tiempo_relevamiento,
               a.tiempo_diseno, a.tiempo_desarr, a.tiempo_test, j.descripcion, k.descripcion,
               a.req_presup, a.id_cliente, a.contacto, id_aplic, a.respon_gestion_seg, a.respon_tec_seg,
               a.estado_func, a.estado_desa, a.prioridad, a.browser, a.via_pedido, a.req_presup,
               SUM(a.tiempo_diseno + a.tiempo_desarr + a.tiempo_relevamiento + a.tiempo_test),
               a.ok_cliente, a.ruta, a.avances, a.id_apl, a.tipo_proy
        FROM CG_PROYECTOS a
        INNER JOIN CG_CLIENTES b ON (a.id_cliente = b.id)
        INNER JOIN CG_CONTACTOS c ON (a.contacto = c.id)
        INNER JOIN CG_APLICATIVOS d ON (a.id_aplic = d.id)
        INNER JOIN CG_PRIORIDADES e ON (a.prioridad = e.id)
        --INNER JOIN CG_COR_RESPONSABLES f ON (a.respon_gestion_seg = f.id_seg AND f.funcion = 'FUNC')
        --INNER JOIN CG_COR_RESPONSABLES g ON (a.respon_tec_seg = g.id_seg AND g.funcion = 'TEC')
        LEFT JOIN vw_UserGrupos f on f.USRNET = a.respon_gestion_seg AND f.GRUPO = 'FUNC'
        LEFT JOIN vw_UserGrupos g on g.USRNET = a.respon_tec_seg AND g.GRUPO = 'TEC'
        LEFT JOIN CG_ESTADOS h ON (a.estado_desa = h.id)
        LEFT JOIN CG_ESTADOS i ON (a.estado_func = i.id)
        INNER JOIN CG_BROWSERS j ON (a.browser = j.id)
        INNER JOIN CG_VIA_PEDIDO k ON (a.via_pedido = k.id)
        WHERE a.id = @Id
        GROUP BY a.id, a.solic_nro, a.titulo, a.descripcion, b.descripcion, c.nombre + ' ' + c.apellido,
                 d.descripcion, f.USRNAME, g.USRNAME, a.fecha_ingreso,
                 a.fecha_entrega, h.descripcion, i.descripcion, e.descripcion, a.tiempo_relevamiento,
                 a.tiempo_diseno, a.tiempo_desarr, a.tiempo_test, j.descripcion, k.descripcion,
                 a.req_presup, a.id_cliente, a.contacto, id_aplic, a.respon_gestion_seg, a.respon_tec_seg,
                 a.estado_func, a.estado_desa, a.prioridad, a.browser, a.via_pedido, a.ok_cliente,
                 a.ruta, a.avances, a.id_apl, a.tipo_proy;
    END

    /* =====================================================================
       APL (vista principal / consultas)
       ===================================================================== */
    ELSE IF @Entidad = 'apl' AND (@Modo = 'lista' OR @Modo = 'consulta')
    BEGIN
        IF @Cons = '' OR @Cons IS NULL
        BEGIN
            SELECT a.id, a.nombre, b.descripcion AS cliente, c.descripcion AS aplicativo,
                   a.fecha_envio, a.sec, d.descripcion AS tipo_apl
            FROM CG_APL a
            INNER JOIN CG_CLIENTES b ON a.id_cliente = b.id
            INNER JOIN CG_APLICATIVOS c ON a.id_aplic = c.id
            LEFT JOIN CG_TIPO_APL d ON a.tipo_apl = d.id
            WHERE a.ESTADO = 1
              AND (@IdCliente = 0 OR a.id_cliente = @IdCliente)
              AND (@IdAplicativo = 0 OR a.id_aplic = @IdAplicativo)
              AND (@IdTip = 0 OR a.tipo_apl = @IdTip)
            ORDER BY
                CASE WHEN @Orden = 'nom' THEN a.nombre END DESC,
                CASE WHEN @Orden <> 'nom' THEN a.sec END DESC,
                a.id DESC;
        END
        ELSE IF @Cons = 'cli' OR @Cons = 'apli'
        BEGIN
            SELECT a.id, a.nombre, b.descripcion AS cliente, c.descripcion AS aplicativo,
                   a.fecha_envio, a.sec
            FROM CG_APL a
            INNER JOIN CG_CLIENTES b ON a.id_cliente = b.id
            INNER JOIN CG_APLICATIVOS c ON a.id_aplic = c.id
            WHERE (@Cons = 'cli' AND a.id_cliente = @IdCliente)
               OR (@Cons = 'apli' AND a.id_aplic = @IdAplicativo)
            ORDER BY a.id DESC;
        END
        ELSE
        BEGIN
            SELECT a.id, a.nombre, b.descripcion AS cliente, c.descripcion AS aplicativo,
                   a.fecha_envio, a.sec
            FROM CG_APL a
            INNER JOIN CG_CLIENTES b ON a.id_cliente = b.id
            INNER JOIN CG_APLICATIVOS c ON a.id_aplic = c.id
            WHERE a.ESTADO = 1
              AND (@IdCliente = 0 OR a.id_cliente = @IdCliente)
              AND (@IdAplicativo = 0 OR a.id_aplic = @IdAplicativo)
              AND (@Asp = '' OR a.motivo_apl LIKE '%' + @Asp + '%')
              AND (@Fec = 0 OR a.fecha_creacion = @Fec)
            ORDER BY a.id DESC;
        END
    END

    /* =====================================================================
       APL (detalle lectura / edicion / existe) - datosDetalles.inc
       ===================================================================== */
    ELSE IF @Entidad = 'apl' AND @Modo = 'detalle'
    BEGIN
        SELECT a.id, a.nombre, b.descripcion, c.descripcion, d.USRNAME,
               a.fecha_envio, a.fecha_instalacion_test, a.motivo_apl, a.fecha_instalacion_prod,
               a.fecha_creacion, a.incidente, a.proyecto, a.sec
        FROM CG_APL a
        INNER JOIN CG_CLIENTES b ON (a.id_cliente = b.id)
        INNER JOIN CG_APLICATIVOS c ON (a.id_aplic = c.id)
        --INNER JOIN CG_COR_RESPONSABLES d ON (a.respon_envio_seg = d.id_seg AND d.funcion = 'FUNC')
        LEFT JOIN vw_UserGrupos d on d.USRNET = a.respon_envio_seg --AND d.GRUPO = 'FUNC'
        WHERE a.id = @Id;
    END

    ELSE IF @Entidad = 'apl' AND @Modo = 'edicion'
    BEGIN
        SELECT a.id, a.nombre, b.id, c.id, a.respon_envio_seg, a.fecha_envio, a.fecha_instalacion_test,
               a.motivo_apl, a.fecha_instalacion_prod, a.incidente, a.proyecto, a.sec,
               a.tipo_apl, e.descripcion
        FROM CG_APL a
        INNER JOIN CG_CLIENTES b ON (a.id_cliente = b.id)
        INNER JOIN CG_APLICATIVOS c ON (a.id_aplic = c.id)
        --INNER JOIN CG_COR_RESPONSABLES d ON (a.respon_envio_seg = d.id_seg AND d.funcion = 'FUNC')
        --LEFT JOIN vw_UserGrupos d on d.USRNET = a.respon_envio_seg AND d.GRUPO = 'FUNC'
        LEFT JOIN CG_TIPO_APL e ON (a.tipo_apl = e.id)
        WHERE a.id = @Id;
    END

    ELSE IF @Entidad = 'apl' AND @Modo = 'existe'
    BEGIN
        SELECT id FROM CG_APL WHERE id = @Id;
    END

    /* =====================================================================
       CARGA DE INCIDENTE (cargaInc.asp)
       ===================================================================== */
    ELSE IF @Entidad = 'carga' AND @Modo = 'incidente'
    BEGIN
        SELECT a.titulo, a.id_proyecto, a.detalles, b.descripcion, c.descripcion,
               d.USRNAME, a.req_presup, e.descripcion, f.descripcion,
               g.descripcion, h.USRNAME, i.nombre + ' ' + i.apellido
        FROM CG_INCIDENTES a
        LEFT JOIN CG_CLIENTES b ON (a.id_cliente = b.id)
        LEFT JOIN CG_APLICATIVOS c ON (a.id_aplic = c.id)
        --LEFT JOIN CG_COR_RESPONSABLES d ON (a.respon_gestion_seg = d.id_seg AND d.funcion = 'FUNC')
        LEFT JOIN vw_UserGrupos d on d.USRNET = a.respon_gestion_seg AND d.GRUPO = 'FUNC'
        LEFT JOIN CG_VIA_PEDIDO e ON (a.via_pedido = e.id)
        LEFT JOIN CG_BROWSERS f ON (a.browser = f.id)
        LEFT JOIN CG_PRIORIDADES g ON (a.prioridad = g.id)
        --LEFT JOIN CG_COR_RESPONSABLES h ON (a.respon_tec_seg = h.id_seg AND h.funcion = 'TEC')
        LEFT JOIN vw_UserGrupos h on h.USRNET = a.respon_tec_seg AND h.GRUPO = 'TEC'
        LEFT JOIN CG_CONTACTOS i ON (a.contacto = i.id)
        WHERE a.id = @Id;
    END

    /* =====================================================================
       COMENTARIOS (historial de avances)
       ===================================================================== */
    ELSE IF @Entidad = 'comentarios'
    BEGIN

      SELECT * FROM (
        SELECT AVANCES as descripcion , FECHA_INGRESO as fecha , 0 as hora ,
        ISNULL(a.RESPON_GESTION_seg, '') AS usuario,
                'FUNCIONAL' AS responsable, 'INICIAL' AV_ESTADO, 'FUNCIONAL' TAV_DESCRIPCION, '-' VIA_PEDIDO,
        CAST(NULL AS INT) AS ID, CAST(NULL AS INT) AS ID_ANT
        from CG_INCIDENTES a
        --LEFT JOIN CG_COR_RESPONSABLES b ON (b.id_seg = a.RESPON_GESTION_seg)
        where a.ID = @CodObjeto AND 'inc' = @TipoObjeto
        UNION ALL
        SELECT AVANCES, FECHA_INGRESO, 0,
        ISNULL(a.RESPON_GESTION_seg, '') AS usuario,
                'FUNCIONAL' AS responsable, 'INICIAL' AV_ESTADO, 'FUNCIONAL' TAV_DESCRIPCION, '-' VIA_PEDIDO,
        CAST(NULL AS INT), CAST(NULL AS INT)
        FROM CG_PROYECTOS a
        --LEFT JOIN CG_COR_RESPONSABLES b ON (b.id_seg = a.RESPON_GESTION_seg)
        WHERE a.ID = @CodObjeto AND 'proy' = @TipoObjeto
        UNION ALL
        SELECT a.descripcion, a.fecha, a.hora,
               ISNULL(a.idusuario_seg, '') AS usuario,
               ISNULL(a.responsable, '') AS responsable, c.DESCRIPCION, d.DESCRIPCION, e.DESCRIPCION,
               a.ID, a.ID_ANT
               FROM CG_AVANCES a
        --LEFT JOIN CG_COR_RESPONSABLES b ON (b.id_seg = a.idusuario_seg)
        LEFT JOIN CG_ESTADOS c ON (a.IDESTADO = c.id)
        LEFT JOIN CG_TIPO_AVANCE d ON (TRY_CAST(a.TIPO AS INT) = d.id)   -- TIPO es varchar(10): puede ser NULL o no numerico
        LEFT JOIN CG_VIA_PEDIDO e ON (a.VIA_AVANCE = e.id)
        WHERE a.codobjeto = @CodObjeto AND a.tipoobjeto = @TipoObjeto ) A1
         ORDER BY A1.fecha DESC, A1.hora DESC;


    END

   

    /* =====================================================================
       PUEDE ACTUALIZAR (validacion de incidentes)
       ===================================================================== */
    ELSE IF @Entidad = 'puedeact'
    BEGIN
        SELECT estado_desa, estado_func, respon_gestion_seg AS respon_gestion, respon_tec_seg AS respon_tec
        FROM CG_INCIDENTES
        WHERE id = @Id;
    END

    -- ADDED (migracion detalles-ORIG.asp / detallesnEW.asp): la funcion local
    -- puedeAct() de esas paginas lee el recordset POSICIONALMENTE como
    -- fcs(0)=respon_gestion, fcs(1)=respon_tec. Devuelve solo esos dos campos.
    ELSE IF @Entidad = 'puede_resp'
    BEGIN
        SELECT respon_gestion_seg AS respon_gestion, respon_tec_seg AS respon_tec
        FROM CG_INCIDENTES
        WHERE id = @Id;
    END
END;
GO
