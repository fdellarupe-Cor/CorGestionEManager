/* =====================================================================
   Cambios_Incidentes_Ambiente.sql  -  campo AMBIENTE de los incidentes

   ESTADO: SCRIPT PARA REVISAR. NO SE EJECUTO.
   Base destino: Dbcordatv48_UNIF485     Compatibilidad: nivel 100
   Requisito: ya ejecutado  ALTER TABLE CG_INCIDENTES ADD AMBIENTE varchar(1) NULL

   CAMPO      : CG_INCIDENTES.AMBIENTE = ''  Unico | 'T' Testing | 'P' Produccion | 'D' Desarrollo.
                Se asocia con CG_CLIEN_X_APLIC por ID_CLIENTE, ID_APLIC y TIPO_AMB (Unico = TIPO_AMB vacio;
                se comparan sin espacios y se ignoran las filas con ESTADO distinto de 1).
   MARCA      : al guardar (alta o modificacion) se pone al PRINCIPIO del campo DETALLES una primera linea
                "- <CG_CLIEN_X_APLIC.AMBIENTE> -". Si no existe esa configuracion, la linea dice
                "- Sin configuracion conocida -". Si la primera linea de DETALLES ya es una marca
                ("- algo -"), se REEMPLAZA. El resto del detalle queda igual.
                Si no se informa ambiente (NULL, por ejemplo incidentes viejos), DETALLES no se toca.
                DETALLES es nvarchar(4000): si con la marca no entra, se recorta el final del texto.

   QUE CREA / CAMBIA (cada ALTER reemplaza el procedimiento completo):
     1. dbo.fn_cg_DetallesConAmbiente   NUEVA funcion que arma el texto con la marca.
     2. sp_cg_Registro_Insertar         parametro nuevo @AmbienteInc (al final); guarda AMBIENTE y la marca.
     3. sp_cg_Incidentes_Insertar       parametro nuevo @ambiente (al final); lo pasa al anterior.
     4. sp_cg_Registro_ActualizarCampos el campo 'ambiente' se guarda como cualquier columna; despues, si
                                        cambio ambiente, cliente, aplicativo o detalles de un incidente, se actualiza la marca.
                                        TAMBIEN CORRIGE el valor recibido: el ASP lo manda escapado (JSONEscape) y el
                                        procedimiento lo guardaba tal cual, asi que los saltos de linea quedaban como el
                                        texto "
" y las barras duplicadas. Ahora se deshace el escape (afecta a todos los
                                        campos de incidentes y proyectos que se editan por este camino).
     5. sp_cg_ConsultarEntidad          AMBIENTE agregado AL FINAL de las ramas de incidente lista, consulta,
                                        todo, detalle y edicion (las pantallas leen por posicion).
   BASE DE PARTIDA: Corv48UNIFV5_DB.sql del 15/10/2026 15:39 para los puntos 2 a 4; para el punto 5, el
                    texto del script 03_ConsultarEntidad_Respon_y_Proyecto.sql (ya aplicado).
                    OJO: si modificaste esos procedimientos despues, volve a exportarlos antes de aplicar.
   ORDEN      : ejecutar el script completo (la funcion se crea antes de los procedimientos que la usan).
   SI NO SE EJECUTA : agregar2.asp y detalles.asp envian el campo nuevo y el alta falla por parametro desconocido;
                      desplegar el script junto con los ASP.
   REVERSA    : volver a ejecutar la definicion original de cada procedimiento (DB/Corv48UNIFV5_DB.sql y el
                script 03) y  DROP FUNCTION dbo.fn_cg_DetallesConAmbiente;
   ===================================================================== */

USE [Dbcordatv48_UNIF485]
GO

-- La funcion se crea vacia si no existe y luego se define con ALTER (funciona en cualquier version)
IF OBJECT_ID(N'dbo.fn_cg_DetallesConAmbiente', N'FN') IS NULL
    EXEC (N'CREATE FUNCTION dbo.fn_cg_DetallesConAmbiente() RETURNS NVARCHAR(MAX) AS BEGIN RETURN NULL; END');
GO

ALTER FUNCTION dbo.fn_cg_DetallesConAmbiente
(
    @IdCliente INT,
    @IdAplic   INT,
    @Ambiente  VARCHAR(1),        -- '' Unico | 'T' Testing | 'P' Produccion | 'D' Desarrollo | NULL = no tocar
    @Detalles  NVARCHAR(MAX)
)
RETURNS NVARCHAR(MAX)
AS
BEGIN
    -- Sin ambiente informado (por ejemplo incidentes viejos): el texto queda como esta.
    IF @Ambiente IS NULL
        RETURN @Detalles;

    DECLARE @tipo  VARCHAR(1);
    DECLARE @amb   NVARCHAR(100);
    DECLARE @marca NVARCHAR(200);
    DECLARE @det   NVARCHAR(MAX);
    DECLARE @pos   INT;
    DECLARE @pos2  INT;
    DECLARE @salto INT;
    DECLARE @linea NVARCHAR(MAX);

    SET @tipo = LTRIM(RTRIM(@Ambiente));     -- Unico = vacio

    -- Texto del ambiente configurado para ese cliente + aplicativo + tipo de ambiente
    SELECT TOP 1 @amb = NULLIF(LTRIM(RTRIM(x.AMBIENTE)), '')
    FROM dbo.CG_CLIEN_X_APLIC x
    WHERE x.ID_CLIENTE = @IdCliente
      AND x.ID_APLIC   = @IdAplic
      AND LTRIM(RTRIM(ISNULL(x.TIPO_AMB, ''))) = @tipo
      AND ISNULL(x.ESTADO, 1) = 1
    ORDER BY x.ID DESC;

    SET @marca = N'- ' + ISNULL(@amb, N'Sin configuracion conocida') + N' -';

    -- Si la primera linea ya es una marca ("- algo -"), se quita para reemplazarla
    SET @det = ISNULL(@Detalles, N'');
    -- El fin de la primera linea es un salto real (CHAR(10)) o, en textos guardados por versiones anteriores
    -- de la edicion, el texto "\n" (barra + n); se toma el que aparezca primero.
    SET @pos  = CHARINDEX(CHAR(10), @det);
    SET @pos2 = CHARINDEX(N'\n', @det);
    SET @salto = 1;
    IF @pos2 > 0 AND (@pos = 0 OR @pos2 < @pos)
    BEGIN
        SET @pos = @pos2;
        SET @salto = 2;
    END
    IF @pos = 0
        SET @linea = @det;
    ELSE
        SET @linea = LEFT(@det, @pos - 1);
    SET @linea = LTRIM(RTRIM(REPLACE(@linea, CHAR(13), N'')));

    IF LEN(@linea) >= 4 AND LEFT(@linea, 2) = N'- ' AND RIGHT(@linea, 2) = N' -'
    BEGIN
        IF @pos = 0
            SET @det = N'';
        ELSE
            SET @det = SUBSTRING(@det, @pos + @salto, LEN(@det));
    END

    IF LEN(@det) = 0
        RETURN LEFT(@marca, 4000);

    -- La columna DETALLES es nvarchar(4000): si no entra, se recorta el final del texto
    RETURN LEFT(@marca + CHAR(13) + CHAR(10) + @det, 4000);
END;
GO

ALTER PROCEDURE [dbo].[sp_cg_Registro_Insertar]
    @Entidad           VARCHAR(20)  = 'incidente',  -- 'incidente' | 'proyecto' | 'licencia' | 'auxiliar'
    @Param             VARCHAR(30)  = NULL,         -- solo para @Entidad='auxiliar'
    @Dato1             VARCHAR(255) = NULL,
    @Dato2             VARCHAR(255) = NULL,
    @Dato3             VARCHAR(255) = NULL,
    -- Campos incidente
    @IdCliente         INT          = NULL,
    @Contacto          INT          = NULL,
    @IdAplic           INT          = NULL,
    @Detalles          VARCHAR(MAX) = NULL,
    @Prioridad         INT          = NULL,
    @ResponGestion     NVARCHAR(100)          = NULL,
    @ResponTec         NVARCHAR(100)          = NULL,
    @ReqPresup         VARCHAR(10)  = NULL,
    @FechaIngreso      INT          = NULL,
    @Browser           INT          = NULL,
    @ViaPedido         INT          = NULL,
    @Titulo            VARCHAR(200) = NULL,
    @Responsable       NVARCHAR(20) = NULL,
    -- Campos proyecto
    @SolicNro          VARCHAR(50)  = NULL,
    @Descripcion       VARCHAR(MAX) = NULL,
    @FechaEntrega      INT          = NULL,
    @TiempoRelevamiento DECIMAL(9,2) = NULL,
    @TiempoDiseno      DECIMAL(9,2) = NULL,
    @TiempoDesarr      DECIMAL(9,2) = NULL,
    @TiempoTest        DECIMAL(9,2) = NULL,
    @TipoProyecto      INT          = 1,
    -- Campos licencia
    @IdRespon          NVARCHAR(100)  = NULL,
    @FInicio           INT          = NULL,
    @FFin              INT          = NULL,
    @Motivo            VARCHAR(MAX) = NULL,
    -- Campos apl
    @Nombre            VARCHAR(200) = NULL,
    @ResponEnvio       NVARCHAR(100)          = NULL,
    @FechaEnvio        INT          = NULL,
    @MotivoApl         VARCHAR(MAX)= NULL,
    @Sec               INT          = NULL,
    @TipoApl           INT          = NULL,
    @FechaIngresoApl   INT          = NULL,
    -- Campos Clientes_det
    @servidor          VARCHAR(20)  = NULL,
    @ambiente          VARCHAR(10)  = NULL,
    @db_rqt            VARCHAR(200) = NULL,
    @licencia          NVARCHAR(100) = NULL,
    -- Ambiente del incidente: '' Unico | 'T' Testing | 'P' Produccion | 'D' Desarrollo (NULL = sin marca en DETALLES)
    @AmbienteInc       VARCHAR(1)   = NULL
AS
BEGIN
    SET NOCOUNT ON;

    /* ------------------------------------------------------------------
       ALTA DE INCIDENTE
       ------------------------------------------------------------------ */
    IF @Entidad = 'incidente'
    BEGIN
        DECLARE @estado_desa INT = CASE WHEN ISNULL(NULLIF(@ResponTec, N''), N'0') <> N'0' THEN 2 ELSE 0 END;
        DECLARE @estado_func INT = 2;
       
        INSERT INTO dbo.CG_INCIDENTES
            (id_cliente, contacto, id_aplic, detalles, prioridad, respon_gestion_seg, respon_tec_seg,
             req_presup, estado_desa, estado_func, fecha_ingreso, browser, via_pedido,
             id_proyecto, estado, titulo, id_apl, ok_desarrollo, ok_funcional, responsable, ambiente)
        VALUES
            (@IdCliente, @Contacto, @IdAplic, dbo.fn_cg_DetallesConAmbiente(@IdCliente, @IdAplic, @AmbienteInc, @Detalles), @Prioridad, @ResponGestion, @ResponTec,
             @ReqPresup, @estado_desa, @estado_func, @FechaIngreso, @Browser, @ViaPedido,
             0, 1, @Titulo, 0, 0, 0, UPPER(ISNULL(@Responsable, 'FUNCIONAL')), @AmbienteInc);
    END

    /* ------------------------------------------------------------------
       ALTA DE PROYECTO
       ------------------------------------------------------------------ */
    ELSE IF @Entidad = 'proyecto'
    BEGIN
        INSERT INTO dbo.CG_PROYECTOS
            (solic_nro, titulo, descripcion, contacto, id_cliente, id_aplic, respon_gestion_seg, respon_tec_seg,
             fecha_ingreso, fecha_entrega, estado_func, estado_desa, prioridad, tiempo_relevamiento,
             tiempo_diseno, tiempo_desarr, tiempo_test, browser, via_pedido, req_presup, estado, id_apl,
             ok_desarrollo, ok_funcional, responsable, tipo_proy)
        VALUES
            (@SolicNro, @Titulo, @Descripcion, @Contacto, @IdCliente, @IdAplic, @ResponGestion, @ResponTec,
             @FechaIngreso, @FechaEntrega, 2, 2, @Prioridad, @TiempoRelevamiento,
             @TiempoDiseno, @TiempoDesarr, @TiempoTest, @Browser, @ViaPedido, @ReqPresup, 1, 0,
             0, 0, UPPER(ISNULL(@Responsable, 'FUNCIONAL')),@TipoProyecto);
    END

    /* ------------------------------------------------------------------
       ALTA DE LICENCIA
       ------------------------------------------------------------------ */
    ELSE IF @Entidad = 'licencia'
    BEGIN
        INSERT INTO dbo.CG_LICENCIAS (id_respon_seg, f_inicio, f_fin, motivo, estado)
        VALUES (@IdRespon, @FInicio, @FFin, @Motivo, 1);
        
    END

    /* ------------------------------------------------------------------
       ALTA DE APL (agregar3.asp)
       ------------------------------------------------------------------ */
    ELSE IF @Entidad = 'apl'
    BEGIN
        INSERT INTO dbo.CG_APL
            (nombre, id_cliente, id_aplic, respon_envio_seg, fecha_envio, motivo_apl,
             estado, fecha_creacion, fecha_instalacion_prod, fecha_instalacion_test,
             incidente, proyecto, sec, tipo_apl)
        VALUES
            (@Nombre, @IdCliente, @IdAplic, @ResponEnvio, 0, @MotivoApl,
             1, @FechaIngresoApl, 0, 0, 0, 0, @Sec, @TipoApl);
    END

    /* ------------------------------------------------------------------
       ALTA DE clientes_det (clientes_det.asp)
       ------------------------------------------------------------------ */
    ELSE IF @Entidad = 'clientes_det'
    BEGIN
        INSERT INTO dbo.CG_CLIENTES_DET 
            (id_cliente, servidor, ambiente, db_rqt, licencia, estado) 
            VALUES 
            (@IdCliente, @servidor, @ambiente, @db_rqt, @licencia, 1 )
       
    END

    /* ------------------------------------------------------------------
       ALTA EN TABLAS AUXILIARES / RELACIONES APL
       ------------------------------------------------------------------ */
    ELSE IF @Entidad = 'auxiliar'
    BEGIN
        IF @Param = 'CLIENTE'
            INSERT INTO CG_CLIENTES (descripcion, estado) VALUES (@Dato1, 1 );
            
        ELSE IF @Param = 'APLICACION'
            INSERT INTO CG_APLICATIVOS (descripcion, orden, estado ) VALUES (@Dato1, ISNULL(TRY_CAST(@Dato2 AS INT), 0), 1 );

        ELSE IF @Param = 'VIA PEDIDO'
            INSERT INTO CG_VIA_PEDIDO (descripcion, estado ) VALUES (@Dato1, 1 );

        ELSE IF @Param = 'PRIORIDAD'
            INSERT INTO CG_PRIORIDADES (descripcion, estado ) VALUES (@Dato1, 1);

        ELSE IF @Param = 'BROWSER'
            INSERT INTO CG_BROWSERS (descripcion, estado ) VALUES (@Dato1, 1);

        ELSE IF @Param = 'ESTADO'
            INSERT INTO CG_ESTADOS (descripcion, estado ) VALUES (@Dato1, 1);
      
       ELSE IF @Param = 'TIPO_AVANCE'
            INSERT INTO CG_TIPO_AVANCE (descripcion, estado ) VALUES (@Dato1, 1);

   /*   ELSE IF @Param = 'RESPON GESTION'
            INSERT INTO CG_COR_RESPONSABLES (nombre, apellido, funcion, estado )
            VALUES (@Dato1, @Dato2, 'FUNC', 1);

        ELSE IF @Param = 'RESPON TECNICO'
            INSERT INTO CG_COR_RESPONSABLES (nombre, apellido, funcion, estado)
            VALUES (@Dato1, @Dato2, 'TEC', 1);
            */
        ELSE IF @Param = 'CONTACTO'
            INSERT INTO CG_CONTACTOS (nombre, apellido, estado)
            VALUES (@Dato1, @Dato2, 1);

        ELSE IF @Param = 'IOP'
            INSERT INTO CG_APL_RELAC (id_apl, id_relac, tipo_relac, estado)
            VALUES (@Dato3, @Dato1, @Dato2, 1 );
            

        ELSE IF @Param = 'APL'
        BEGIN
            DECLARE @TipoRelac VARCHAR(10);
            IF @Dato3 = 'acti' OR @Dato3 = 'I'
                SET @TipoRelac = 'I';
            ELSE
                SET @TipoRelac = 'P';

            IF NOT EXISTS (SELECT 1 FROM CG_APL WHERE id = TRY_CAST(@Dato1 AS INT))
            BEGIN
                RAISERROR('El ID de APL especificado no existe.', 16, 1);
                RETURN;
            END

            IF @TipoRelac = 'I' AND NOT EXISTS (SELECT 1 FROM CG_INCIDENTES WHERE ID = TRY_CAST(@Dato2 AS INT))
            BEGIN
                RAISERROR('El Incidente especificado no existe.', 16, 1);
                RETURN;
            END

            IF @TipoRelac = 'P' AND NOT EXISTS (SELECT 1 FROM CG_PROYECTOS WHERE ID = TRY_CAST(@Dato2 AS INT))
            BEGIN
                RAISERROR('El Proyecto especificado no existe.', 16, 1);
                RETURN;
            END

            INSERT INTO CG_APL_RELAC (id_apl, id_relac, tipo_relac, estado)
            VALUES (@Dato1, @Dato2, @TipoRelac, 1 );
        END
    END
END;
GO

ALTER PROCEDURE [dbo].[sp_cg_Incidentes_Insertar]
    @id_cliente     INT,
    @contacto       INT,
    @id_aplic       INT,
    @detalles       VARCHAR(MAX),
    @prioridad      INT,
    @respon_gestion NVARCHAR(100),
    @respon_tec     NVARCHAR(100),
    @req_presup     VARCHAR(10),
    @fecha_ingreso  INT,
    @browser        INT,
    @via_pedido     INT,
    @titulo         VARCHAR(200),
    @responsable    NVARCHAR(20) = NULL,
    @ambiente       VARCHAR(1) = NULL      -- '' Unico | 'T' Testing | 'P' Produccion | 'D' Desarrollo
AS
BEGIN
    SET NOCOUNT ON;
    EXEC dbo.sp_cg_Registro_Insertar @Entidad = 'incidente',
        @IdCliente = @id_cliente, @Contacto = @contacto, @IdAplic = @id_aplic,
        @Detalles = @detalles, @Prioridad = @prioridad, @ResponGestion = @respon_gestion,
        @ResponTec = @respon_tec, @ReqPresup = @req_presup, @FechaIngreso = @fecha_ingreso,
        @Browser = @browser, @ViaPedido = @via_pedido, @Titulo = @titulo,
        @Responsable = @responsable, @AmbienteInc = @ambiente;
END;
GO

ALTER PROCEDURE [dbo].[sp_cg_Registro_ActualizarCampos]
    @tabla      NVARCHAR(50),
    @id         INT,
    @camposJson NVARCHAR(MAX)
AS
BEGIN
    SET NOCOUNT ON;

    IF @tabla NOT IN ('INCIDENTES', 'PROYECTOS')
    BEGIN
        RAISERROR('Tabla no permitida: %s', 16, 1, @tabla);
        RETURN;
    END;

    SELECT @tabla = 'CG_'+ @tabla

    DECLARE @tmpFields TABLE (col NVARCHAR(128), val NVARCHAR(MAX));
    DECLARE @currentField NVARCHAR(MAX), @col NVARCHAR(128), @val NVARCHAR(MAX);
    DECLARE @startPos INT, @endPos INT, @jsonLen INT;

    SET @camposJson = LTRIM(RTRIM(@camposJson));
    SET @jsonLen = LEN(@camposJson);

    IF @camposJson = '[]' OR @camposJson IS NULL OR @camposJson = ''
        RETURN;

    SET @startPos = CHARINDEX('{', @camposJson, 2);

    WHILE @startPos > 0
    BEGIN
        SET @endPos = CHARINDEX('}', @camposJson, @startPos);
        IF @endPos = 0 BREAK;

        SET @currentField = SUBSTRING(@camposJson, @startPos, @endPos - @startPos + 1);

        SET @col = SUBSTRING(@currentField, CHARINDEX('"col":\"', @currentField) + 8, 128);
        SET @col = SUBSTRING(@col, 1, CHARINDEX('"', @col) - 1);

        SET @val = SUBSTRING(@currentField, CHARINDEX('"val":\"', @currentField) + 8, LEN(@currentField));
        SET @val = SUBSTRING(@val, 1, CHARINDEX('"', @val) - 1);
        -- [AMBIENTE] Deshace el escape del ASP: \n (barra + n) = salto de linea y \\ = una barra. Asi el texto se guarda igual que se escribio.
        SET @val = REPLACE(REPLACE(REPLACE(@val, N'\\', NCHAR(1)), N'\n', NCHAR(13) + NCHAR(10)), NCHAR(1), N'\');

        INSERT INTO @tmpFields (col, val) VALUES (@col, @val);

        SET @startPos = CHARINDEX('{', @camposJson, @endPos + 1);
    END;

    -- Los combos de responsable envian ID_SEG: respon_gestion/respon_tec se guardan en las columnas *_SEG.

    UPDATE @tmpFields SET col = col + N'_SEG' WHERE col IN (N'respon_gestion', N'respon_tec');


    IF EXISTS (
        SELECT 1
        FROM @tmpFields f
        WHERE NOT EXISTS (
            SELECT 1 FROM INFORMATION_SCHEMA.COLUMNS c
            WHERE c.TABLE_NAME = @tabla AND c.COLUMN_NAME = f.col
        )
    )
    BEGIN
        RAISERROR('Se recibio un campo que no es columna valida de %s', 16, 1, @tabla);
        RETURN;
    END;

    DECLARE @setClause NVARCHAR(MAX);

    SET @setClause = (
        SELECT ',' + QUOTENAME(col) + N' = ''' + REPLACE(ISNULL(val, ''), '''', '''''') + ''''
        FROM @tmpFields
        FOR XML PATH(''), TYPE
    ).value('.', 'NVARCHAR(MAX)');

    IF LEN(@setClause) > 0
        SET @setClause = STUFF(@setClause, 1, 1, '');

    DECLARE @sql NVARCHAR(MAX) =
        N'UPDATE ' +  QUOTENAME(@tabla) + N' SET ' + @setClause + N' WHERE id = @pId';

    EXEC sp_executesql @sql, N'@pId INT', @pId = @id;

    -- [AMBIENTE] En incidentes, si cambio el ambiente, el cliente, el aplicativo o el detalle, se actualiza la marca "- ambiente -" al principio de DETALLES
    IF @tabla = 'CG_INCIDENTES'
       AND EXISTS (SELECT 1 FROM @tmpFields WHERE col IN (N'ambiente', N'id_cliente', N'id_aplic', N'detalles'))
        UPDATE dbo.CG_INCIDENTES
           SET DETALLES = dbo.fn_cg_DetallesConAmbiente(ID_CLIENTE, ID_APLIC, AMBIENTE, DETALLES)
         WHERE ID = @id;
END
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
                'FUNCIONAL' AS responsable, 'INICIAL' AV_ESTADO, 'FUNCIONAL' TAV_DESCRIPCION, '-' VIA_PEDIDO
        from CG_INCIDENTES a
        --LEFT JOIN CG_COR_RESPONSABLES b ON (b.id_seg = a.RESPON_GESTION_seg)
        where a.ID = @CodObjeto AND 'inc' = @TipoObjeto
        UNION ALL
        SELECT AVANCES, FECHA_INGRESO, 0,
        ISNULL(a.RESPON_GESTION_seg, '') AS usuario,
                'FUNCIONAL' AS responsable, 'INICIAL' AV_ESTADO, 'FUNCIONAL' TAV_DESCRIPCION, '-' VIA_PEDIDO
        FROM CG_PROYECTOS a
        --LEFT JOIN CG_COR_RESPONSABLES b ON (b.id_seg = a.RESPON_GESTION_seg)
        WHERE a.ID = @CodObjeto AND 'proy' = @TipoObjeto
        UNION ALL
        SELECT a.descripcion, a.fecha, a.hora,
               ISNULL(a.idusuario_seg, '') AS usuario,
               ISNULL(a.responsable, '') AS responsable, c.DESCRIPCION, d.DESCRIPCION, e.DESCRIPCION
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
