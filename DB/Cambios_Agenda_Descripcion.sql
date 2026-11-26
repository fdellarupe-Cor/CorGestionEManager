/* =====================================================================
   Cambios_Agenda_Descripcion.sql  -  descripcion breve de la reunion (CG_AGENDA.DESCRIPCION)

   ESTADO: SCRIPT PARA REVISAR. NO SE EJECUTO.
   Base destino: Dbcordatv48_UNIF485     Compatibilidad: nivel 100
   Requisito   : la columna ya existe (ALTER TABLE CG_AGENDA ADD DESCRIPCION VARCHAR(50) NULL).
   Base de partida: texto de los procedimientos en DB/Corv48UNIFV5_DB.sql del 22/10/2026 11:28.
   OJO: cada ALTER reemplaza el procedimiento completo con ese texto. Si los modificaste despues,
        volve a exportarlos antes de aplicar este script.

   1) sp_cg_Agenda_VistaSemanal : el primer resultado devuelve la columna DESCRIPCION (agenda_semana.asp
      la muestra en el panel de detalle, 3 renglones, y en el pop-up de la reunion completa).
   2) sp_cg_AgendaExt           : parametro nuevo @Descripcion varchar(50) = NULL (al final, opcional).
        - 'obtener'    devuelve DESCRIPCION.
        - 'insertar'   la guarda (vacia = NULL).
        - 'actualizar' la guarda; si el parametro no viene (NULL) la deja como estaba; si viene '' la borra.
   SI NO SE EJECUTA : las pantallas siguen funcionando; la descripcion no se muestra ni se guarda.
   REVERSA    : volver a ejecutar la definicion original de ambos procedimientos (DB/Corv48UNIFV5_DB.sql).
   ===================================================================== */

USE [Dbcordatv48_UNIF485]
GO

ALTER PROCEDURE [dbo].[sp_cg_Agenda_VistaSemanal]
    @FechaDesde INT,
    @FechaHasta INT
AS
BEGIN
    SET NOCOUNT ON;

    SELECT
        A.ID,
        A.ASUNTO,
        A.FECHA_REUNION,
        A.HORA_INICIO,
        A.HORA_FIN,
        ISNULL(A.ESTADO_AGENDA, 'PENDIENTE') AS ESTADO_AGENDA,
        A.ID_RESPONSABLE_SEG AS ID_RESPONSABLE,
        RTRIM(ISNULL((SELECT MAX(R.USRNAME) FROM dbo.vw_UserGrupos R WHERE R.USRNET = A.ID_RESPONSABLE_SEG), '')) AS RESPONSABLE,
        A.DESCRIPCION   -- [AGENDA] descripcion breve de la reunion (agenda_semana.asp la muestra en el panel)
    FROM dbo.CG_AGENDA A
    WHERE A.FEC_BAJA IS NULL
      AND A.FECHA_REUNION BETWEEN @FechaDesde AND @FechaHasta
    ORDER BY A.FECHA_REUNION, A.HORA_INICIO;

    SELECT
        C.ID_AGENDA,
        C.ID_PERSONA_SEG AS ID_PERSONA,
        RTRIM(ISNULL(
            NULLIF(LTRIM(RTRIM(C.NOMBRE_MANUAL)), ''),
            ISNULL((SELECT MAX(R.USRNAME) FROM dbo.vw_UserGrupos R WHERE R.USRNET = C.ID_PERSONA_SEG), '')
        )) AS NOMBRE
    FROM dbo.CG_AGENDA_CORRESPONSABLE C
    INNER JOIN dbo.CG_AGENDA A ON A.ID = C.ID_AGENDA AND A.FEC_BAJA IS NULL
    WHERE C.FEC_BAJA IS NULL
      AND A.FECHA_REUNION BETWEEN @FechaDesde AND @FechaHasta
    ORDER BY C.ID_AGENDA;

    -- [AGENDA] Tercer resultado: licencias vigentes que caen (total o parcialmente) en el rango.
    -- Si F_FIN esta vacio o en 0 se toma igual a F_INICIO. agenda_semana.asp las marca como un dia completo.
    SELECT
        L.ID,
        L.ID_RESPON_SEG AS ID_PERSONA,
        RTRIM(ISNULL((SELECT MAX(R.USRNAME) FROM dbo.vw_UserGrupos R WHERE R.USRNET = L.ID_RESPON_SEG), L.ID_RESPON_SEG)) AS PERSONA,
        L.F_INICIO,
        CASE WHEN ISNULL(L.F_FIN, 0) = 0 OR L.F_FIN < L.F_INICIO THEN L.F_INICIO ELSE L.F_FIN END AS F_FIN,
        L.MOTIVO
    FROM dbo.CG_LICENCIAS L
    WHERE L.ESTADO = 1
      AND L.FEC_BAJA IS NULL
      AND L.F_INICIO IS NOT NULL
      AND L.F_INICIO <= @FechaHasta
      AND (CASE WHEN ISNULL(L.F_FIN, 0) = 0 OR L.F_FIN < L.F_INICIO THEN L.F_INICIO ELSE L.F_FIN END) >= @FechaDesde
    ORDER BY L.F_INICIO, L.ID;
END;
GO

ALTER PROCEDURE [dbo].[sp_cg_AgendaExt]
    @Modo             varchar(30),
    @Id               int = NULL,
    @Asunto           nvarchar(1000) = NULL,
    @FechaReunion     INT = NULL,
    @HoraInicio       INT = NULL,
    @HoraFin          INT = NULL,
    @FechaAviso       INT = NULL,
    @IdRespon         nvarchar(100) = NULL,
    @IdHerramienta    int = NULL,
    @Mensaje          nvarchar(max) = NULL,
    @Conclusion       nvarchar(4000) = NULL,
    @TipoOrigen       char(1) = NULL,
    @Estado           varchar(15) = NULL,
    @IdPersona        nvarchar(100) = NULL,
    @NombreManual     nvarchar(100) = NULL,
    @EmailManual      nvarchar(150) = NULL,
    @EnviarCopia      bit = 0,
    @IdTipoRelacion   int = NULL,
    @IdProyecto       int = NULL,
    @IdIncidente      int = NULL,
    @Tema             nvarchar(500) = NULL,
    @TiempoUsado      int = 0,
    @Orden            int = 1,
    @Seguimiento      bit = 0,
    @IdAccion         int = NULL,
    @CantRegMover     int = 0,
    @CantRegMostrar   int = 15,
    @IdCorresponsable int = NULL,
    @Descripcion      varchar(50) = NULL    -- [AGENDA] descripcion breve (CG_AGENDA.DESCRIPCION). NULL en 'actualizar' = no cambiarla; '' = borrarla
AS
BEGIN
    SET NOCOUNT ON;
    SET XACT_ABORT ON;

    DECLARE @TipoRelacion varchar(20), @NuevoId int, @LimiteConsulta int;

    IF @Modo = 'estados_listar'
    BEGIN
        SELECT ID, DESCRIPCION
        FROM dbo.CG_ESTADOS
        WHERE ISNULL(ESTADO, 1) = 1 AND FEC_BAJA IS NULL
        ORDER BY ORDEN, DESCRIPCION;
        RETURN;
    END;

    IF @Modo = 'acciones_listar'
    BEGIN
        SELECT ID, DESCRIPCION
        FROM dbo.CG_TIPO_ACCION
        WHERE ACTIVO = 1 AND FEC_BAJA IS NULL
        ORDER BY DESCRIPCION;
        RETURN;
    END;

    IF @Modo = 'consulta'
    BEGIN
        SET @LimiteConsulta = ISNULL(@CantRegMostrar, 15) + 1;

        WITH AgendaBase AS (
            SELECT a.ID, a.ID_AGENDA_PADRE, a.ASUNTO, a.FECHA_REUNION,
                   a.HORA_INICIO, a.HORA_FIN, a.ESTADO_AGENDA, a.ID_RESPON_SEG,
                   ROW_NUMBER() OVER (
                       ORDER BY ISNULL(a.FECHA_REUNION, 0) DESC, a.ID DESC
                   ) AS RN
            FROM dbo.CG_AGENDA AS a
            WHERE a.FEC_BAJA IS NULL
        )
        SELECT b.ID, b.ID_AGENDA_PADRE, b.ASUNTO, b.FECHA_REUNION, b.HORA_INICIO, b.HORA_FIN,
               ISNULL(b.ESTADO_AGENDA, 'PENDIENTE') AS ESTADO,
               RTRIM(ISNULL((SELECT MAX(r.USRNAME)
                             FROM vw_UserGrupos r
                             WHERE r.USRNET = b.ID_RESPON_SEG), '')) AS RESPONSABLE,
               (SELECT COUNT(*) FROM dbo.CG_AGENDA_CORRESPONSABLE c
                 WHERE c.ID_AGENDA = b.ID AND c.FEC_BAJA IS NULL) AS CANT_CORRESP,
               (SELECT COUNT(*) FROM dbo.CG_AGENDA_SUBTEMA s
                 WHERE s.ID_AGENDA = b.ID AND s.FEC_BAJA IS NULL) AS CANT_SUBTEMAS
        FROM AgendaBase b
        WHERE b.RN >  ISNULL(@CantRegMover, 0)
          AND b.RN <= ISNULL(@CantRegMover, 0) + @LimiteConsulta
        ORDER BY b.RN;
        RETURN;
    END;

    IF @Modo = 'obtener'
    BEGIN
        SELECT
            a.ID,
            a.ASUNTO,
            a.FECHA_REUNION,
            a.HORA_INICIO,
            a.HORA_FIN,
            a.FECHA_AVISO,
            a.MENSAJE,
            a.CONCLUSION,
            a.ID_RESPONSABLE_SEG AS ID_RESPON,
            a.ID_HERRAMIENTA,
            ISNULL(a.ESTADO_AGENDA, CASE WHEN ISNULL(a.ESTADO, 0) = 1 THEN 'PENDIENTE' ELSE 'INACTIVA' END) AS ESTADO,
            a.TIPO_ORIGEN,
            a.ID_AGENDA_PADRE,
            a.DESCRIPCION
        FROM dbo.CG_AGENDA AS a
        WHERE a.ID = @Id AND a.FEC_BAJA IS NULL;
        RETURN;
    END;

    IF @Modo = 'insertar'
    BEGIN
        IF @Asunto IS NULL OR LTRIM(RTRIM(@Asunto)) = ''
            THROW 50001, 'El asunto de la agenda es obligatorio.', 1;
        IF @IdRespon IS NULL
            THROW 50002, 'El responsable de la agenda es obligatorio.', 1;
        IF @FechaReunion IS NULL OR @HoraInicio IS NULL OR @HoraFin IS NULL
            THROW 50003, 'La fecha y el horario de la reunion son obligatorios.', 1;
        IF @HoraFin <= @HoraInicio
            THROW 50004, 'La hora de fin debe ser posterior a la hora de inicio.', 1;

        INSERT INTO dbo.CG_AGENDA
        (
            ASUNTO, DIA_AVISO, ID_RESPON_SEG, ENVIADO, ESTADO, FECHADECARGA,
            MENSAJE, ID_VIA, FEC_REU, CONCLUSION, FECHA_REUNION, HORA_INICIO,
            HORA_FIN, FECHA_AVISO, ID_RESPONSABLE_SEG, ID_HERRAMIENTA,
            TIPO_ORIGEN, ESTADO_AGENDA, ENVIADO_BIT, FECHA_CREACION, DESCRIPCION
        )
        VALUES
        (
            @Asunto,
            @FechaAviso,
            @IdRespon, 'NO', 1, DATEDIFF(DAY, CONVERT(date, '19000101'), CONVERT(date, GETDATE())),
            @Mensaje, @IdHerramienta, @FechaReunion, @Conclusion,
            @FechaReunion, @HoraInicio, @HoraFin, @FechaAviso, @IdRespon, @IdHerramienta,
            ISNULL(@TipoOrigen, 'N'), ISNULL(@Estado, 'PENDIENTE'), 0, DATEDIFF(DAY, CONVERT(date, '19000101'), CONVERT(date, GETDATE())),
            NULLIF(LTRIM(RTRIM(@Descripcion)), '')
        );

        SET @NuevoId = CONVERT(int, SCOPE_IDENTITY());
        SELECT @NuevoId AS ID;
        RETURN;
    END;

    IF @Modo = 'actualizar'
    BEGIN
        IF NOT EXISTS (SELECT 1 FROM dbo.CG_AGENDA WHERE ID = @Id AND FEC_BAJA IS NULL)
            THROW 50005, 'La agenda indicada no existe o esta dada de baja.', 1;
        IF @HoraFin <= @HoraInicio
            THROW 50004, 'La hora de fin debe ser posterior a la hora de inicio.', 1;

        UPDATE dbo.CG_AGENDA
        SET ASUNTO = @Asunto,
            DIA_AVISO = @FechaAviso,
            ID_RESPON_SEG = @IdRespon,
            MENSAJE = @Mensaje,
            ID_VIA = @IdHerramienta,
            FEC_REU = @FechaReunion,
            CONCLUSION = @Conclusion,
            FECHA_REUNION = @FechaReunion,
            HORA_INICIO = @HoraInicio,
            HORA_FIN = @HoraFin,
            FECHA_AVISO = @FechaAviso,
            ID_RESPONSABLE_SEG = @IdRespon,
            ID_HERRAMIENTA = @IdHerramienta,
            ESTADO_AGENDA = ISNULL(@Estado, ESTADO_AGENDA),
            DESCRIPCION = CASE WHEN @Descripcion IS NULL THEN DESCRIPCION ELSE NULLIF(LTRIM(RTRIM(@Descripcion)), '') END
        WHERE ID = @Id;
        RETURN;
    END;

    IF @Modo = 'corresp_agregar'
    BEGIN
        IF @Id IS NULL OR NOT EXISTS (SELECT 1 FROM dbo.CG_AGENDA WHERE ID = @Id AND FEC_BAJA IS NULL)
            THROW 50006, 'La agenda indicada no existe.', 1;
        IF @IdPersona IS NULL AND NULLIF(LTRIM(RTRIM(@NombreManual)), '') IS NULL
            THROW 50007, 'Debe indicar una persona del sistema o un nombre manual.', 1;
        IF @EnviarCopia = 1 AND @IdPersona IS NULL AND NULLIF(LTRIM(RTRIM(@EmailManual)), '') IS NULL
            THROW 50008, 'El email es obligatorio para enviar copia a un contacto manual.', 1;

        INSERT INTO dbo.CG_AGENDA_CORRESPONSABLE
            (ID_AGENDA, ID_PERSONA_SEG, NOMBRE_MANUAL, EMAIL_MANUAL, ENVIAR_COPIA)
        VALUES
            (@Id, @IdPersona,
             CASE WHEN @IdPersona IS NULL THEN NULLIF(LTRIM(RTRIM(@NombreManual)), '') ELSE NULL END,
             CASE WHEN @IdPersona IS NULL THEN NULLIF(LTRIM(RTRIM(@EmailManual)), '') ELSE NULL END,
             ISNULL(@EnviarCopia, 0));
        RETURN;
    END;

    IF @Modo = 'corresp_listar'
    BEGIN
        SELECT c.ID_AGENDA_CORRESP,
               c.ID_PERSONA_SEG AS ID_PERSONA,
               c.NOMBRE_MANUAL,
               ISNULL(NULLIF(LTRIM(RTRIM(c.EMAIL_MANUAL)), ''), r.MAIL) AS EMAIL_MANUAL,
               c.ENVIAR_COPIA,
               NULLIF(LTRIM(RTRIM(ISNULL(r.USRNAME, ''))), '') AS NOMBRE_PERSONA
        FROM dbo.CG_AGENDA_CORRESPONSABLE AS c
       -- LEFT JOIN dbo.CG_COR_RESPONSABLES AS r ON r.ID_SEG = c.ID_PERSONA_SEG
        LEFT JOIN vw_UserGrupos r ON c.ID_PERSONA_SEG = r.USRNET 
        WHERE c.ID_AGENDA = @Id AND c.FEC_BAJA IS NULL
        ORDER BY c.ID_AGENDA_CORRESP;
        RETURN;
    END;

    IF @Modo = 'corresp_eliminar'
    BEGIN
        IF @Id IS NULL OR NOT EXISTS (SELECT 1 FROM dbo.CG_AGENDA WHERE ID = @Id AND FEC_BAJA IS NULL)
            THROW 50005, 'La agenda indicada no existe o esta dada de baja.', 1;
        IF ISNULL(@IdCorresponsable, TRY_CAST(@IdPersona AS INT)) IS NULL
            THROW 50016, 'Falta el ID del participante a eliminar.', 1;

        UPDATE dbo.CG_AGENDA_CORRESPONSABLE
        SET fec_baja = GETDATE()
        WHERE ID_AGENDA_CORRESP = ISNULL(@IdCorresponsable, TRY_CAST(@IdPersona AS INT))
          AND ID_AGENDA = @Id
          AND FEC_BAJA IS NULL;
        RETURN;
    END;

    IF @Modo = 'corresp_limpiar'
    BEGIN
        IF @Id IS NULL OR NOT EXISTS (SELECT 1 FROM dbo.CG_AGENDA WHERE ID = @Id AND FEC_BAJA IS NULL)
            THROW 50005, 'La agenda indicada no existe o esta dada de baja.', 1;

        UPDATE dbo.CG_AGENDA_CORRESPONSABLE
        SET FEC_BAJA = GETDATE()
        WHERE ID_AGENDA = @Id AND FEC_BAJA IS NULL;
        RETURN;
    END;

    IF @Modo = 'subtema_limpiar'
    BEGIN
        IF @Id IS NULL OR NOT EXISTS (SELECT 1 FROM dbo.CG_AGENDA WHERE ID = @Id AND FEC_BAJA IS NULL)
            THROW 50005, 'La agenda indicada no existe o esta dada de baja.', 1;

        UPDATE dbo.CG_AGENDA_SUBTEMA
        SET FEC_BAJA = GETDATE()
        WHERE ID_AGENDA = @Id AND FEC_BAJA IS NULL;
        RETURN;
    END;

    IF @Modo = 'subtema_agregar'
    BEGIN
        IF ISNULL(@IdTipoRelacion, 0) = 0 AND NULLIF(LTRIM(RTRIM(@Tema)), '') IS NULL
            RETURN;

        SET @TipoRelacion = CASE @IdTipoRelacion
            WHEN 1 THEN 'PROYECTO'
            WHEN 2 THEN 'INCIDENTE'
            WHEN 3 THEN 'CONSULTORIA'
            WHEN 4 THEN 'OTRO'
            ELSE NULL
        END;

        IF @TipoRelacion IS NULL
            THROW 50009, 'El tipo de relacion del subtema no es valido.', 1;
        IF NULLIF(LTRIM(RTRIM(@Tema)), '') IS NULL
            THROW 50010, 'El tema a tratar es obligatorio.', 1;
        IF @TiempoUsado < 0 OR @Orden < 1
            THROW 50011, 'El tiempo y el orden del subtema no pueden ser negativos.', 1;
        IF @TipoRelacion = 'PROYECTO' AND @IdProyecto IS NULL
            THROW 50012, 'Un subtema de proyecto requiere ID_PROYECTO.', 1;
        IF @TipoRelacion = 'INCIDENTE' AND @IdIncidente IS NULL
            THROW 50013, 'Un subtema de incidente requiere ID_INCIDENTE.', 1;

        INSERT INTO dbo.CG_AGENDA_SUBTEMA
            (ID_AGENDA, TIPO_RELACION, ID_PROYECTO, ID_INCIDENTE, TEMA_A_TRATAR,
             CONCLUSION, TIEMPO_USADO, ORDEN, SEGUIMIENTO, ID_ACCION)
        VALUES
            (@Id, @TipoRelacion,
             CASE WHEN @TipoRelacion = 'PROYECTO' THEN @IdProyecto ELSE NULL END,
             CASE WHEN @TipoRelacion = 'INCIDENTE' THEN @IdIncidente ELSE NULL END,
             @Tema, @Conclusion, @TiempoUsado, @Orden, ISNULL(@Seguimiento, 0), @IdAccion);
        RETURN;
    END;

    IF @Modo = 'subtema_listar'
    BEGIN
         SELECT s.ID_SUBTEMA, s.ID_AGENDA,
               CASE TIPO_RELACION WHEN 'PROYECTO' THEN 1 WHEN 'INCIDENTE' THEN 2 WHEN 'CONSULTORIA' THEN 3 WHEN 'OTRO' THEN 4 END AS ID_TIPO_RELACION,
             s.ID_PROYECTO, s.ID_INCIDENTE, s.TEMA_A_TRATAR, s.CONCLUSION,
               s.TIEMPO_USADO, s.ORDEN, s.SEGUIMIENTO, s.ID_ACCION, s.PATH_ADJUNTOS
         FROM dbo.CG_AGENDA_SUBTEMA AS s
         WHERE s.ID_AGENDA = @Id AND s.FEC_BAJA IS NULL
         ORDER BY s.ORDEN, s.ID_SUBTEMA;
        RETURN;
    END;

    IF @Modo IN ('postergar', 'seguimiento')
    BEGIN
        IF NOT EXISTS (SELECT 1 FROM dbo.CG_AGENDA WHERE ID = @Id AND FEC_BAJA IS NULL)
            THROW 50005, 'La agenda original no existe o esta dada de baja.', 1;
        IF @FechaReunion IS NULL OR @HoraInicio IS NULL OR @HoraFin IS NULL OR @HoraFin <= @HoraInicio
            THROW 50014, 'La nueva fecha y el horario no son validos.', 1;

        BEGIN TRY
            BEGIN TRANSACTION;

            INSERT INTO dbo.CG_AGENDA
            (
                ASUNTO, DIA_AVISO, ID_RESPON_SEG, ENVIADO, ESTADO, FECHADECARGA,
                MENSAJE, ID_VIA, FEC_REU, CONCLUSION, FECHA_REUNION, HORA_INICIO,
                HORA_FIN, FECHA_AVISO, ID_RESPONSABLE_SEG, ID_HERRAMIENTA,
                ID_AGENDA_PADRE, TIPO_ORIGEN, ESTADO_AGENDA, ENVIADO_BIT, FECHA_CREACION
            )
            SELECT ASUNTO,
                   DIA_AVISO, ID_RESPON_SEG, 'NO', 1, FECHADECARGA,
                   MENSAJE, ID_VIA, @FechaReunion,
                   CASE WHEN @Modo = 'seguimiento' THEN @Conclusion ELSE NULL END,
                   @FechaReunion, @HoraInicio, @HoraFin, FECHA_AVISO,
                   ID_RESPONSABLE_SEG, ID_HERRAMIENTA, ID, CASE WHEN @Modo = 'seguimiento' THEN 'S' ELSE 'P' END,
                   'PENDIENTE', 0, DATEDIFF(DAY, CONVERT(date, '19000101'), CONVERT(date, GETDATE()))
            FROM dbo.CG_AGENDA
            WHERE ID = @Id;

            SET @NuevoId = CONVERT(int, SCOPE_IDENTITY());
            IF @NuevoId IS NULL
                THROW 50015, 'No se pudo crear la nueva agenda.', 1;

            INSERT INTO dbo.CG_AGENDA_CORRESPONSABLE
                (ID_AGENDA, ID_PERSONA_SEG, NOMBRE_MANUAL, EMAIL_MANUAL, ENVIAR_COPIA)
            SELECT @NuevoId, ID_PERSONA_SEG, NOMBRE_MANUAL, EMAIL_MANUAL, ENVIAR_COPIA
            FROM dbo.CG_AGENDA_CORRESPONSABLE
            WHERE ID_AGENDA = @Id AND FEC_BAJA IS NULL;

            INSERT INTO dbo.CG_AGENDA_SUBTEMA
                (ID_AGENDA, TIPO_RELACION, ID_PROYECTO, ID_INCIDENTE, TEMA_A_TRATAR,
                 CONCLUSION, TIEMPO_USADO, ORDEN, SEGUIMIENTO, ID_ACCION, PATH_ADJUNTOS)
            SELECT @NuevoId, TIPO_RELACION, ID_PROYECTO, ID_INCIDENTE, TEMA_A_TRATAR,
                   NULL, 0, ORDEN,
                   CASE WHEN @Modo = 'seguimiento' THEN 0 ELSE SEGUIMIENTO END,
                   ID_ACCION, PATH_ADJUNTOS
            FROM dbo.CG_AGENDA_SUBTEMA
            WHERE ID_AGENDA = @Id AND FEC_BAJA IS NULL
              AND (@Modo = 'postergar' OR SEGUIMIENTO = 1);

            UPDATE dbo.CG_AGENDA
            SET ESTADO_AGENDA = CASE WHEN @Modo = 'seguimiento' THEN 'REALIZADA' ELSE 'POSTERGADA' END
            WHERE ID = @Id;

            COMMIT TRANSACTION;
            SELECT @NuevoId AS ID;
            RETURN;
        END TRY
        BEGIN CATCH
            IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
            DECLARE @MensajeError nvarchar(4000), @SeveridadError int, @EstadoError int;
            SELECT @MensajeError = ERROR_MESSAGE(), @SeveridadError = ERROR_SEVERITY(), @EstadoError = ERROR_STATE();
            RAISERROR(@MensajeError, @SeveridadError, @EstadoError);
            RETURN;
        END CATCH;
    END;

    RAISERROR('Modo de sp_cg_AgendaExt no reconocido.', 16, 1);
END;
GO
