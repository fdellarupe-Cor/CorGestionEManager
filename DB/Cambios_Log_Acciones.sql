/* =====================================================================================================
   Cambios_Log_Acciones.sql
   Log de acciones de usuario de Corgestion (CG_LOG): un registro por USUARIO + SESION + ELEMENTO, con la
   accion de MAYOR nivel. Pensado tambien para notificar al dueno del elemento.

   Niveles:  1 CONSULTA  <  3 ACTUALIZA  <  4 ALTA / BAJA
   Entidades: INCIDENTE, PROYECTO, APL, AGENDA (los avances se registran como ACTUALIZA del incidente/proyecto)
   Regla del SP: si la accion nueva tiene nivel MAYOR que la guardada, la reemplaza (y actualiza la fecha);
                 si es igual o menor, solo actualiza la ultima fecha y suma 1 al contador.
   Dueno del elemento (RESPONSABLE_SEG): se resuelve en el SP desde la tabla del elemento
                 (incidente/proyecto: RESPON_GESTION_SEG, APL: RESPON_ENVIO_SEG, agenda: ID_RESPON_SEG).

   Ejecutar por el usuario. Repetible (no pisa datos). Termina en ROLLBACK: revisar y cambiar a COMMIT TRAN.
   Ejecutar ANTES de subir los .asp (si la tabla no existe, CgLog falla en silencio y no registra nada).
   ===================================================================================================== */
SET NOCOUNT ON;
BEGIN TRAN;

IF OBJECT_ID(N'dbo.CG_LOG', N'U') IS NULL
BEGIN
    CREATE TABLE dbo.CG_LOG (
        ID                INT IDENTITY(1,1) NOT NULL CONSTRAINT PK_CG_LOG PRIMARY KEY,
        USUARIO           NVARCHAR(100) NOT NULL,
        SESION            NVARCHAR(60)  NOT NULL,
        ENTIDAD           VARCHAR(20)   NOT NULL,
        ID_ENTIDAD        INT           NOT NULL,
        ACCION            VARCHAR(20)   NOT NULL,
        NIVEL             TINYINT       NOT NULL,
        FECHA_PRIMERA     DATETIME      NOT NULL CONSTRAINT DF_CG_LOG_FP DEFAULT (GETDATE()),
        FECHA_ULTIMA      DATETIME      NOT NULL CONSTRAINT DF_CG_LOG_FU DEFAULT (GETDATE()),
        CONTADOR          INT           NOT NULL CONSTRAINT DF_CG_LOG_CNT DEFAULT (1),
        IP                VARCHAR(45)   NULL,
        PANTALLA          VARCHAR(60)   NULL,
        RESPONSABLE_SEG   NVARCHAR(100) NULL,
        LEIDO             BIT           NOT NULL CONSTRAINT DF_CG_LOG_LEIDO DEFAULT (0)   -- reservado para notificaciones
    );
    CREATE UNIQUE INDEX UX_CG_LOG_VISITA ON dbo.CG_LOG (USUARIO, SESION, ENTIDAD, ID_ENTIDAD);
    CREATE INDEX IX_CG_LOG_FECHA ON dbo.CG_LOG (FECHA_ULTIMA);
    CREATE INDEX IX_CG_LOG_RESP  ON dbo.CG_LOG (RESPONSABLE_SEG, FECHA_ULTIMA);
END
GO

CREATE OR ALTER PROCEDURE dbo.sp_cg_Log_Registrar
    @Usuario    NVARCHAR(100),
    @Sesion     NVARCHAR(60),
    @Entidad    VARCHAR(20),     -- INCIDENTE | PROYECTO | APL | AGENDA
    @IdEntidad  INT,
    @Accion     VARCHAR(20),     -- CONSULTA | ACTUALIZA | ALTA | BAJA
    @Ip         VARCHAR(45) = NULL,
    @Pantalla   VARCHAR(60) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRY
        DECLARE @Nivel TINYINT = CASE UPPER(@Accion) WHEN 'CONSULTA' THEN 1 WHEN 'ACTUALIZA' THEN 3 WHEN 'ALTA' THEN 4 WHEN 'BAJA' THEN 4 ELSE 0 END;
        IF @Nivel = 0 OR ISNULL(@IdEntidad, 0) <= 0 OR ISNULL(@Usuario, N'') = N'' RETURN;
        SET @Entidad = UPPER(@Entidad);
        SET @Accion  = UPPER(@Accion);

        UPDATE dbo.CG_LOG
           SET ACCION   = CASE WHEN @Nivel > NIVEL THEN @Accion ELSE ACCION END,
               NIVEL    = CASE WHEN @Nivel > NIVEL THEN @Nivel  ELSE NIVEL  END,
               PANTALLA = CASE WHEN @Nivel > NIVEL THEN @Pantalla ELSE PANTALLA END,
               FECHA_ULTIMA = GETDATE(),
               CONTADOR = CONTADOR + 1
         WHERE USUARIO = @Usuario AND SESION = @Sesion AND ENTIDAD = @Entidad AND ID_ENTIDAD = @IdEntidad;

        IF @@ROWCOUNT = 0
        BEGIN
            DECLARE @Resp NVARCHAR(100) = NULL;
            IF @Entidad = 'INCIDENTE' SELECT @Resp = RESPON_GESTION_SEG FROM dbo.CG_INCIDENTES WHERE ID = @IdEntidad;
            ELSE IF @Entidad = 'PROYECTO' SELECT @Resp = RESPON_GESTION_SEG FROM dbo.CG_PROYECTOS WHERE ID = @IdEntidad;
            ELSE IF @Entidad = 'APL' SELECT @Resp = RESPON_ENVIO_SEG FROM dbo.CG_APL WHERE ID = @IdEntidad;
            ELSE IF @Entidad = 'AGENDA' SELECT @Resp = ID_RESPON_SEG FROM dbo.CG_AGENDA WHERE ID = @IdEntidad;

            INSERT INTO dbo.CG_LOG (USUARIO, SESION, ENTIDAD, ID_ENTIDAD, ACCION, NIVEL, IP, PANTALLA, RESPONSABLE_SEG)
            VALUES (@Usuario, @Sesion, @Entidad, @IdEntidad, @Accion, @Nivel, @Ip, @Pantalla, @Resp);
        END
    END TRY
    BEGIN CATCH
        -- el log nunca debe romper la pantalla (carrera por indice unico: se descarta)
        RETURN;
    END CATCH
END
GO

CREATE OR ALTER PROCEDURE dbo.sp_cg_Log_Purgar
    @Meses INT = 12
AS
BEGIN
    SET NOCOUNT ON;
    DELETE FROM dbo.CG_LOG WHERE FECHA_ULTIMA < DATEADD(MONTH, -ABS(@Meses), GETDATE());
    SELECT @@ROWCOUNT AS Eliminados;
END
GO

/* Prueba rapida (se deshace con el ROLLBACK) */
EXEC dbo.sp_cg_Log_Registrar N'PRUEBA', N'S1', 'INCIDENTE', 1, 'CONSULTA', '127.0.0.1', 'prueba';
EXEC dbo.sp_cg_Log_Registrar N'PRUEBA', N'S1', 'INCIDENTE', 1, 'ACTUALIZA', '127.0.0.1', 'prueba';
EXEC dbo.sp_cg_Log_Registrar N'PRUEBA', N'S1', 'INCIDENTE', 1, 'CONSULTA', '127.0.0.1', 'prueba';
SELECT 'Debe quedar 1 fila: ACTUALIZA, contador 3' AS esperado, * FROM dbo.CG_LOG WHERE USUARIO = N'PRUEBA';

ROLLBACK TRAN;   -- Cambiar a COMMIT TRAN cuando el resultado sea el esperado
