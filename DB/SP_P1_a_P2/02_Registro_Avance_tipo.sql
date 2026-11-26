/* =====================================================================
   SP 02 - sp_cg_Registro_Avance: tipo de avance

   ESTADO: SCRIPT PARA REVISAR. NO SE EJECUTO.
   Base destino: Dbcordatv48_UNIF485     Compatibilidad: nivel 100
   Base de partida: texto del procedimiento en P2 (Corv48UNIFV5_DB.sql del 15/10/2026 15:39).
   Solo se AGREGA lo indicado; el resto del procedimiento queda igual (vista vw_UserGrupos,
   paginacion con ROW_NUMBER, @Respon, estado_cliente, etc.).

   QUE AGREGA : parametro opcional  @IdTipoAvance INT = NULL  (al final). Si viene, valida que exista en
                CG_TIPO_AVANCE y lo guarda en CG_AVANCES.TIPO. Si no viene, queda igual que hoy.
   PARA QUE   : poder registrar el tipo de cada avance de incidentes y proyectos (el catalogo se administra
                en tipoavance.asp). Los llamados actuales no cambian (parametro con valor por defecto).
   DE DONDE   : P1 (DB/migracion_cg/15_avance_tipo.sql). En P1 esta vigente.
   VERIFICADO : el procedimiento de P2 y el de P1 son iguales salvo lo del tipo de avance: SI (comparados linea por linea; las unicas diferencias son el parametro, la validacion y la columna TIPO).
   ASP        : para usarlo, datosActualizar.inc / actualizar_v2.asp deben enviar @IdTipoAvance (no incluido aqui).
   NOTA       : usa THROW (SQL Server 2012 o mas), igual que la version actual de P2.

   REVERSA: volver a ejecutar la definicion original del procedimiento (DB/Corv48UNIFV5_DB.sql).
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
    @IdTipoAvance  INT = NULL            -- CG_TIPO_AVANCE.ID (opcional)
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

    BEGIN TRANSACTION;

    INSERT INTO dbo.CG_AVANCES
        (codobjeto, tipoobjeto, orden, fecha, hora, descripcion, idusuario_seg, responsable, tipo, FechaRegistro)
    SELECT @id, CASE WHEN @tabla = 'incidente' THEN 'inc' ELSE 'proy' END,
           ISNULL(MAX(orden), 0) + 1, @fecha, @hora, @descripcion, @idusuario,
           UPPER(@responsable), CAST(@IdTipoAvance AS NVARCHAR(50)),
           DATEDIFF(DAY, CONVERT(date, '19000101'), CONVERT(date, GETDATE()))
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
