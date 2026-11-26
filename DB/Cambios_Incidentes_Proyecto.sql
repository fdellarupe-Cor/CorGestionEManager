/* =====================================================================
   Cambios_Incidentes_Proyecto.sql  -  asociar un proyecto (opcional) al dar de alta un incidente

   ESTADO: SCRIPT PARA REVISAR. NO SE EJECUTO.
   Base destino: Dbcordatv48_UNIF485     Compatibilidad: nivel 100
   Base de partida: texto de los procedimientos en DB/Corv48UNIFV5_DB.sql del 22/10/2026 11:28
                    (ya incluye el ambiente de DB/Cambios_Incidentes_Ambiente.sql).
   OJO: cada ALTER reemplaza el procedimiento completo con ese texto. Si los modificaste despues,
        volve a exportarlos antes de aplicar este script.

   ANTES : el alta de incidentes guardaba CG_INCIDENTES.ID_PROYECTO fijo en 0 y no habia forma de indicar un proyecto.
   AHORA :
     1) sp_cg_Proyectos_ComboCliente (NUEVO): proyectos activos (ESTADO = 1) con ID, DESCRIPCION (solic_nro - titulo)
        e ID_CLIENTE. agregar2.asp lo usa para filtrar el combo "Proyecto relacionado" segun el cliente elegido.
     2) sp_cg_Registro_Insertar : parametro nuevo @IdProyecto INT = NULL (al final, opcional). NULL o 0 = sin proyecto
        (se guarda 0, como hasta ahora). Si viene un proyecto, se valida que exista y sea del cliente (error 50021).
     3) sp_cg_Incidentes_Insertar : parametro nuevo @id_proyecto INT = NULL, que se pasa al anterior.
   SI NO SE EJECUTA : el alta sigue igual; el combo de proyecto en agregar2.asp queda sin opciones (solo "Sin proyecto").
                      El ASP solo envia @id_proyecto cuando se elige un proyecto, asi que el alta sin proyecto no se afecta.
   REVERSA    : volver a ejecutar la definicion original de los dos procedimientos (DB/Corv48UNIFV5_DB.sql)
                y borrar sp_cg_Proyectos_ComboCliente.
   ===================================================================== */

USE [Dbcordatv48_UNIF485]
GO

IF OBJECT_ID(N'dbo.sp_cg_Proyectos_ComboCliente', N'P') IS NULL
    EXEC (N'CREATE PROCEDURE dbo.sp_cg_Proyectos_ComboCliente AS RETURN 0');
GO

ALTER PROCEDURE dbo.sp_cg_Proyectos_ComboCliente
AS
BEGIN
    SET NOCOUNT ON;
    SELECT p.id AS ID,
           ISNULL(p.solic_nro, '') + ' - ' + ISNULL(p.titulo, '') AS DESCRIPCION,
           p.id_cliente AS ID_CLIENTE
    FROM dbo.CG_PROYECTOS p
    WHERE p.ESTADO = 1
    ORDER BY p.id_cliente, p.solic_nro;
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
    @AmbienteInc       VARCHAR(1)   = NULL,
    -- Proyecto relacionado del incidente (CG_INCIDENTES.ID_PROYECTO): NULL o 0 = sin proyecto (es opcional)
    @IdProyecto        INT          = NULL
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

        -- [PROYECTO] Si se indica un proyecto, debe existir y ser del mismo cliente del incidente
        IF ISNULL(@IdProyecto, 0) <> 0
           AND NOT EXISTS (SELECT 1 FROM dbo.CG_PROYECTOS WHERE id = @IdProyecto AND id_cliente = @IdCliente)
            THROW 50021, 'El proyecto indicado no existe o no pertenece al cliente.', 1;
       
        INSERT INTO dbo.CG_INCIDENTES
            (id_cliente, contacto, id_aplic, detalles, prioridad, respon_gestion_seg, respon_tec_seg,
             req_presup, estado_desa, estado_func, fecha_ingreso, browser, via_pedido,
             id_proyecto, estado, titulo, id_apl, ok_desarrollo, ok_funcional, responsable, ambiente)
        VALUES
            (@IdCliente, @Contacto, @IdAplic, dbo.fn_cg_DetallesConAmbiente(@IdCliente, @IdAplic, @AmbienteInc, @Detalles), @Prioridad, @ResponGestion, @ResponTec,
             @ReqPresup, @estado_desa, @estado_func, @FechaIngreso, @Browser, @ViaPedido,
             ISNULL(@IdProyecto, 0), 1, @Titulo, 0, 0, 0, UPPER(ISNULL(@Responsable, 'FUNCIONAL')), @AmbienteInc);
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
    @ambiente       VARCHAR(1) = NULL,     -- '' Unico | 'T' Testing | 'P' Produccion | 'D' Desarrollo
    @id_proyecto    INT = NULL             -- proyecto relacionado (opcional; NULL o 0 = sin proyecto)
AS
BEGIN
    SET NOCOUNT ON;
    EXEC dbo.sp_cg_Registro_Insertar @Entidad = 'incidente',
        @IdCliente = @id_cliente, @Contacto = @contacto, @IdAplic = @id_aplic,
        @Detalles = @detalles, @Prioridad = @prioridad, @ResponGestion = @respon_gestion,
        @ResponTec = @respon_tec, @ReqPresup = @req_presup, @FechaIngreso = @fecha_ingreso,
        @Browser = @browser, @ViaPedido = @via_pedido, @Titulo = @titulo,
        @Responsable = @responsable, @AmbienteInc = @ambiente, @IdProyecto = @id_proyecto;
END;
GO
