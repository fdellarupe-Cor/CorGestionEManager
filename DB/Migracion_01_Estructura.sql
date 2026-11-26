/* =====================================================================
   Migracion_01_Estructura.sql  -  lleva a dbcordatv485Cor (prod) la estructura y el codigo de Dbcordatv48_UNIF485 (desarrollo)

   ESTADO: SCRIPT PARA REVISAR. NO SE EJECUTO.
   Base destino: dbcordatv485Cor
   Base de partida: DB/Corv48UNIFV5_DB.sql (desarrollo, 23/10/2026) contra DB/dbcordatv485Cor.sql (prod). Se compararon TODAS las
   tablas, funciones, procedimientos y vistas: estas son las unicas diferencias.

   1) Columnas: CG_AGENDA.DESCRIPCION (varchar 50), CG_AVANCES.ID_ANT (int), CG_INCIDENTES.AMBIENTE (varchar 1).
      Ademas CG_INCIDENTES.AVANCES pasa a nvarchar(max) (en prod hay avances de mas de 4000 caracteres; el original era max).
   2) Tabla nueva: CG_FERIADOS (feriados de la agenda semanal).
   3) Funciones y procedimientos (texto vigente en desarrollo; los nuevos se CREAN, los que ya existen se ALTERAN):
      Nuevos (4):    fn_cg_DetallesConAmbiente, sp_cg_Feriados, sp_cg_Inicio, sp_cg_Proyectos_ComboCliente
      Modificados (10): sp_cg_Agenda, sp_cg_Agenda_VistaSemanal, sp_cg_AgendaExt, sp_cg_Catalogo_Listar, sp_cg_ConsultarEntidad, sp_cg_Incidentes_Insertar, sp_cg_Licencias_Listar, sp_cg_Registro_ActualizarCampos, sp_cg_Registro_Avance, sp_cg_Registro_Insertar
   Ejecutar ANTES de Migracion_02_Datos.sql. Es repetible: las columnas y la tabla se crean solo si faltan, y los ALTER reemplazan el objeto.
   REVERSA: para los procedimientos, volver a ejecutar su definicion de DB/dbcordatv485Cor.sql; las columnas y la tabla nuevas pueden quedar.
   ===================================================================== */

USE [dbcordatv485Cor]
GO

-- 1) Columnas nuevas
IF COL_LENGTH('dbo.CG_AGENDA', 'DESCRIPCION') IS NULL
    ALTER TABLE dbo.CG_AGENDA ADD DESCRIPCION VARCHAR(50) NULL;
GO
IF COL_LENGTH('dbo.CG_AVANCES', 'ID_ANT') IS NULL
    ALTER TABLE dbo.CG_AVANCES ADD ID_ANT INT NULL;
GO
IF COL_LENGTH('dbo.CG_INCIDENTES', 'AMBIENTE') IS NULL
    ALTER TABLE dbo.CG_INCIDENTES ADD AMBIENTE VARCHAR(1) NULL;
GO

-- 1b) CG_INCIDENTES.AVANCES a nvarchar(max): en prod hay avances de mas de 4000 caracteres (el original era nvarchar(max)).
--     No hay indices ni vistas que dependan de la columna. Corre solo si todavia no es max.
IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'dbo.CG_INCIDENTES') AND name = N'AVANCES' AND max_length <> -1)
    ALTER TABLE dbo.CG_INCIDENTES ALTER COLUMN AVANCES NVARCHAR(MAX) NULL;
GO

-- 2) Tabla nueva
IF OBJECT_ID(N'dbo.CG_FERIADOS', N'U') IS NULL
    CREATE TABLE dbo.CG_FERIADOS (
        FECHA  INT NULL,
        TIPO   VARCHAR(5000) NULL,
        NOMBRE NVARCHAR(MAX) NULL
    );
GO

-- 3) Funciones y procedimientos

-- FUNCTION fn_cg_DetallesConAmbiente (nuevo)
CREATE FUNCTION [dbo].[fn_cg_DetallesConAmbiente]
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

-- PROCEDURE sp_cg_Agenda (modificado)
ALTER PROCEDURE [dbo].[sp_cg_Agenda]
    @Modo      VARCHAR(20),
    @Id        INT          = NULL,
    @Asunto    NVARCHAR(1000) = NULL,
    @DiaAviso  INT          = NULL,
    @IdRespon  NVARCHAR(100)          = NULL,
    @Copias    NVARCHAR(1000) = NULL,
    @FecCarga  INT          = NULL,
    @Mensaje   NVARCHAR(MAX) = NULL,
    @Respon    NVARCHAR(100) = NULL,   -- [AGENDA] 'listar': solo registros del usuario (login): responsable, en COPIAS o corresponsable
    @Texto     NVARCHAR(200) = NULL    -- [AGENDA] 'listar': texto a buscar en ASUNTO o MENSAJE
AS
BEGIN
    SET NOCOUNT ON;

    SET @Respon = NULLIF(LTRIM(RTRIM(@Respon)), '');
    SET @Texto  = NULLIF(LTRIM(RTRIM(@Texto)), '');

    IF @Modo = 'listar'
    BEGIN
        SELECT ID, ASUNTO, DIA_AVISO, ID_RESPON_SEG AS ID_RESPON, ENVIADO, ESTADO, COPIAS, FECHADECARGA, MENSAJE, FEC_BAJA, USU_BAJA, ID_VIA, FEC_REU, HORA_INI, ID_ANT, CONCLUSION, FECHA_REUNION, HORA_INICIO, FECHA_AVISO, ID_RESPONSABLE_SEG AS ID_RESPONSABLE, ID_HERRAMIENTA, ID_AGENDA_PADRE, TIPO_ORIGEN, ESTADO_AGENDA, ENVIADO_BIT, FECHA_CREACION, USUA_BAJA, HORA_FIN
        FROM CG_AGENDA
        WHERE estado = 1
          -- [AGENDA] Solo los del usuario: responsable, en la lista de copias (COPIAS: logins separados por coma) o corresponsable
          AND (@Respon IS NULL
               OR ID_RESPON_SEG = @Respon
               OR CHARINDEX(',' + @Respon + ',', ',' + REPLACE(ISNULL(COPIAS, ''), ' ', '') + ',') > 0
               OR EXISTS (SELECT 1 FROM dbo.CG_AGENDA_CORRESPONSABLE c
                           WHERE c.ID_AGENDA = CG_AGENDA.ID AND c.ID_PERSONA_SEG = @Respon AND c.FEC_BAJA IS NULL))
          -- [AGENDA] Busqueda de texto en asunto o mensaje
          AND (@Texto IS NULL
               OR CHARINDEX(@Texto, ISNULL(ASUNTO, '')) > 0
               OR CHARINDEX(@Texto, ISNULL(MENSAJE, '')) > 0)
        ORDER BY dia_aviso DESC;
    END

    ELSE IF @Modo = 'insertar'
    BEGIN
        INSERT INTO CG_AGENDA (asunto, dia_aviso, id_respon_seg, enviado, estado, copias, fechadecarga, mensaje)
        VALUES (@Asunto, @DiaAviso, @IdRespon, 'NO', 1, @Copias, @FecCarga, @Mensaje);
    END

    ELSE IF @Modo = 'obtener'
    BEGIN
        SELECT enviado, asunto
        FROM CG_AGENDA
        WHERE id = @Id;
    END

    ELSE IF @Modo = 'obtener_new'
    BEGIN
        SELECT enviado, asunto, MENSAJE
        FROM CG_AGENDA
        WHERE id = @Id;
    END

    ELSE IF @Modo = 'actualizar_asunto'
    BEGIN
        UPDATE CG_AGENDA SET asunto = @Asunto WHERE id = @Id;
    END

    ELSE IF @Modo = 'obtener_copias'
    BEGIN
        SELECT copias
        FROM CG_AGENDA
        WHERE id = @Id;
    END

    ELSE IF @Modo = 'actualizar_copias'
    BEGIN
        UPDATE CG_AGENDA SET copias = @Copias WHERE id = @Id;
    END

    ELSE IF @Modo = 'recordatorios'
    BEGIN
        SELECT a.ID, a.ASUNTO, a.DIA_AVISO, a.ID_RESPON_SEG AS ID_RESPON, a.ENVIADO, a.ESTADO, a.COPIAS, a.FECHADECARGA, a.MENSAJE, a.FEC_BAJA, a.USU_BAJA, a.ID_VIA, a.FEC_REU, a.HORA_INI, a.ID_ANT, a.CONCLUSION, a.FECHA_REUNION, a.HORA_INICIO, a.FECHA_AVISO, a.ID_RESPONSABLE_SEG AS ID_RESPONSABLE, a.ID_HERRAMIENTA, a.ID_AGENDA_PADRE, a.TIPO_ORIGEN, a.ESTADO_AGENDA, a.ENVIADO_BIT, a.FECHA_CREACION, a.USUA_BAJA, a.HORA_FIN, b.mail
        FROM CG_AGENDA a
        INNER JOIN CG_COR_RESPONSABLES b ON (a.id_respon_seg = b.id_seg)
        WHERE a.dia_aviso = @DiaAviso AND a.enviado = 'NO' AND a.estado = 1;
    END

    ELSE IF @Modo = 'marcar_enviado'
    BEGIN
        UPDATE CG_AGENDA SET enviado = 'SI' WHERE id = @Id;
    END
END;
GO

-- PROCEDURE sp_cg_Agenda_VistaSemanal (modificado)
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

-- PROCEDURE sp_cg_AgendaExt (modificado)
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

        -- [AGENDA] Sin duplicados: si la misma persona (o el mismo contacto manual) ya esta activa en la reunion, no se inserta de nuevo
        IF EXISTS (SELECT 1 FROM dbo.CG_AGENDA_CORRESPONSABLE x
                    WHERE x.ID_AGENDA = @Id AND x.FEC_BAJA IS NULL
                      AND ( (@IdPersona IS NOT NULL AND x.ID_PERSONA_SEG = @IdPersona)
                         OR (@IdPersona IS NULL AND x.ID_PERSONA_SEG IS NULL
                             AND ISNULL(x.NOMBRE_MANUAL, '') = ISNULL(NULLIF(LTRIM(RTRIM(@NombreManual)), ''), '')
                             AND ISNULL(x.EMAIL_MANUAL, '') = ISNULL(NULLIF(LTRIM(RTRIM(@EmailManual)), ''), '')) ))
            RETURN;

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
        OUTER APPLY (SELECT TOP 1 u.USRNAME, u.MAIL FROM vw_UserGrupos u WHERE u.USRNET = c.ID_PERSONA_SEG ORDER BY u.USRNAME) r   -- [AGENDA] una sola fila por persona (la vista tiene una por grupo) 
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

-- PROCEDURE sp_cg_Catalogo_Listar (modificado)
ALTER PROCEDURE [dbo].[sp_cg_Catalogo_Listar]
    @Catalogo VARCHAR(30),
    @Id       INT = NULL
AS
BEGIN
    SET NOCOUNT ON;

    IF @Catalogo = 'clientes'
        SELECT id, descripcion
        FROM CG_CLIENTES
        WHERE (@Id IS NULL OR id = @Id) AND ISNULL(ESTADO,1)=1
        ORDER BY descripcion;

    ELSE IF @Catalogo = 'aplicativos'
        SELECT id, descripcion
        FROM CG_APLICATIVOS
        WHERE (@Id IS NULL OR id = @Id) AND ISNULL(ESTADO,1)=1
        ORDER BY descripcion;

    ELSE IF @Catalogo = 'browsers'
        SELECT id, descripcion
        FROM CG_BROWSERS
        WHERE (@Id IS NULL OR id = @Id) AND ISNULL(ESTADO,1)=1
        ORDER BY descripcion;

    ELSE IF @Catalogo = 'estados'
        SELECT id, descripcion
        FROM CG_ESTADOS
        WHERE (@Id IS NULL OR id = @Id) AND ISNULL(ESTADO,1)=1
        ORDER BY id;

    ELSE IF @Catalogo = 'prioridades'
        SELECT id, descripcion
        FROM CG_PRIORIDADES
        WHERE (@Id IS NULL OR id = @Id) AND ISNULL(ESTADO,1)=1
        ORDER BY id;

    ELSE IF @Catalogo = 'via_pedido'
        SELECT id, descripcion
        FROM CG_VIA_PEDIDO
        WHERE (@Id IS NULL OR id = @Id) AND ISNULL(ESTADO,1)=1
        ORDER BY descripcion;

    ELSE IF @Catalogo = 'tipo_apl'
        SELECT id, descripcion
        FROM CG_TIPO_APL
        WHERE (@Id IS NULL OR id = @Id) --AND ISNULL(ESTADO,1)=1
        ORDER BY descripcion;

    ELSE IF @Catalogo = 'tipo_avance'
        SELECT id, descripcion
        FROM CG_TIPO_AVANCE
        WHERE (@Id IS NULL OR id = @Id) AND ISNULL(ESTADO,1)=1
        ORDER BY id;
END;
GO

-- PROCEDURE sp_cg_ConsultarEntidad (modificado)
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

-- PROCEDURE sp_cg_Feriados (nuevo)
CREATE PROCEDURE [dbo].[sp_cg_Feriados]
    @Modo   VARCHAR(15),                 -- 'semana' | 'sincronizar'
    @Desde  INT           = NULL,        -- dias desde 1900-01-01
    @Hasta  INT           = NULL,
    @Anio   INT           = NULL,        -- 'sincronizar'
    @Datos  XML           = NULL         -- 'sincronizar': <r><f d="2026-01-01" t="inamovible" n="Ano nuevo"/>...</r> (lo arma el ASP desde la API)
AS
BEGIN
    SET NOCOUNT ON;

    IF @Modo = 'semana'
    BEGIN
        SELECT FECHA, TIPO, NOMBRE
        FROM dbo.CG_FERIADOS
        WHERE FECHA BETWEEN @Desde AND @Hasta
        ORDER BY FECHA;
        RETURN;
    END

    IF @Modo = 'sincronizar'
    BEGIN
        IF @Anio IS NULL OR @Anio NOT BETWEEN 1990 AND 2200
            BEGIN RAISERROR('Anio no valido.', 16, 1); RETURN; END
        -- Sin datos: no se toca la tabla (evita borrar feriados por un error de la API)
        IF @Datos IS NULL
            BEGIN SELECT CAST(0 AS INT) AS CAMBIOS; RETURN; END

        DECLARE @ini INT = DATEDIFF(DAY, '19000101', DATEFROMPARTS(@Anio, 1, 1));
        DECLARE @fin INT = DATEDIFF(DAY, '19000101', DATEFROMPARTS(@Anio, 12, 31));

        DECLARE @api TABLE (FECHA INT, TIPO VARCHAR(5000), NOMBRE NVARCHAR(MAX));
        INSERT INTO @api (FECHA, TIPO, NOMBRE)
        SELECT DATEDIFF(DAY, '19000101', CONVERT(DATE, x.value('@d', 'VARCHAR(10)'), 23)),
               x.value('@t', 'VARCHAR(5000)'),
               x.value('@n', 'NVARCHAR(MAX)')
        FROM @Datos.nodes('/r/f') AS t(x)
        WHERE x.value('@d', 'VARCHAR(10)') LIKE '[0-9][0-9][0-9][0-9]-[0-9][0-9]-[0-9][0-9]'
          AND LEFT(x.value('@d', 'VARCHAR(10)'), 4) = CAST(@Anio AS VARCHAR(4));

        IF NOT EXISTS (SELECT 1 FROM @api)
            BEGIN SELECT CAST(0 AS INT) AS CAMBIOS; RETURN; END

        -- Sin diferencias: nada que hacer
        IF NOT EXISTS (SELECT FECHA, TIPO, NOMBRE FROM @api
                       EXCEPT SELECT FECHA, TIPO, NOMBRE FROM dbo.CG_FERIADOS WHERE FECHA BETWEEN @ini AND @fin)
           AND NOT EXISTS (SELECT FECHA, TIPO, NOMBRE FROM dbo.CG_FERIADOS WHERE FECHA BETWEEN @ini AND @fin
                           EXCEPT SELECT FECHA, TIPO, NOMBRE FROM @api)
            BEGIN SELECT CAST(0 AS INT) AS CAMBIOS; RETURN; END

        BEGIN TRY
            BEGIN TRANSACTION;
            DELETE FROM dbo.CG_FERIADOS WHERE FECHA BETWEEN @ini AND @fin;
            INSERT INTO dbo.CG_FERIADOS (FECHA, TIPO, NOMBRE) SELECT FECHA, TIPO, NOMBRE FROM @api;
            COMMIT TRANSACTION;
        END TRY
        BEGIN CATCH
            IF @@TRANCOUNT > 0 ROLLBACK TRANSACTION;
            THROW;
        END CATCH;

        SELECT CAST(1 AS INT) AS CAMBIOS;
        RETURN;
    END

    RAISERROR('Modo de sp_cg_Feriados no reconocido.', 16, 1);
END;

GO

-- PROCEDURE sp_cg_Incidentes_Insertar (modificado)
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

-- PROCEDURE sp_cg_Inicio (nuevo)
CREATE PROCEDURE [dbo].[sp_cg_Inicio]
    @Modo   VARCHAR(20),                 -- 'casos' | 'casos_cuenta' | 'testing' | 'solicitudes'
    @Respon NVARCHAR(100) = NULL         -- login de red del usuario (ID_SEG)
AS
BEGIN
    SET NOCOUNT ON;

    SET @Respon = NULLIF(LTRIM(RTRIM(@Respon)), '');

    IF @Modo = 'casos'
    BEGIN
        SELECT x.tipo, x.id, x.titulo, x.cliente, x.prioridad, x.prioridad_id, x.fecha_ingreso, x.fecha_compromiso
        FROM (
            SELECT CAST('I' AS CHAR(1)) AS tipo,
                   a.ID                 AS id,
                   a.TITULO             AS titulo,
                   c.DESCRIPCION        AS cliente,
                   pr.DESCRIPCION       AS prioridad,
                   a.PRIORIDAD          AS prioridad_id,
                   a.FECHA_INGRESO      AS fecha_ingreso,
                   CAST(NULL AS INT)    AS fecha_compromiso
            FROM dbo.CG_INCIDENTES a
            LEFT JOIN dbo.CG_CLIENTES c     ON a.ID_CLIENTE = c.ID
            LEFT JOIN dbo.CG_PRIORIDADES pr ON a.PRIORIDAD  = pr.ID
            WHERE a.ESTADO = 1
              AND @Respon IS NOT NULL
              AND (a.RESPON_GESTION_SEG = @Respon OR a.RESPON_TEC_SEG = @Respon)
              AND (NULLIF(LTRIM(RTRIM(a.OK_CLIENTE)), '') IS NULL OR LTRIM(RTRIM(a.OK_CLIENTE)) = '0')
            UNION ALL
            SELECT CAST('P' AS CHAR(1)),
                   p.ID,
                   p.TITULO,
                   c.DESCRIPCION,
                   pr.DESCRIPCION,
                   p.PRIORIDAD,
                   p.FECHA_INGRESO,
                   CAST(NULL AS INT)
            FROM dbo.CG_PROYECTOS p
            LEFT JOIN dbo.CG_CLIENTES c     ON p.ID_CLIENTE = c.ID
            LEFT JOIN dbo.CG_PRIORIDADES pr ON p.PRIORIDAD  = pr.ID
            WHERE p.ESTADO = 1
              AND @Respon IS NOT NULL
              AND (p.RESPON_GESTION_SEG = @Respon OR p.RESPON_TEC_SEG = @Respon)
              AND (NULLIF(LTRIM(RTRIM(p.OK_CLIENTE)), '') IS NULL OR LTRIM(RTRIM(p.OK_CLIENTE)) = '0')
        ) x
        ORDER BY ISNULL(x.prioridad_id, 999) ASC, ISNULL(x.fecha_ingreso, 0) ASC, x.tipo ASC, x.id ASC;
        RETURN;
    END

    IF @Modo = 'casos_cuenta'
    BEGIN
        IF @Respon IS NULL
        BEGIN
            SELECT CAST(0 AS INT) AS INCIDENTES, CAST(0 AS INT) AS PROYECTOS;
            RETURN;
        END

        SELECT
            (SELECT COUNT(*)
               FROM dbo.CG_INCIDENTES a
              WHERE a.ESTADO = 1
                AND (a.RESPON_GESTION_SEG = @Respon OR a.RESPON_TEC_SEG = @Respon)
                AND (NULLIF(LTRIM(RTRIM(a.OK_CLIENTE)), '') IS NULL OR LTRIM(RTRIM(a.OK_CLIENTE)) = '0')
            ) AS INCIDENTES,
            (SELECT COUNT(*)
               FROM dbo.CG_PROYECTOS p
              WHERE p.ESTADO = 1
                AND (p.RESPON_GESTION_SEG = @Respon OR p.RESPON_TEC_SEG = @Respon)
                AND (NULLIF(LTRIM(RTRIM(p.OK_CLIENTE)), '') IS NULL OR LTRIM(RTRIM(p.OK_CLIENTE)) = '0')
            ) AS PROYECTOS;
        RETURN;
    END

    -- [INICIO] Testing / Solicitudes pendientes: avances de ese tipo (CG_TIPO_AVANCE) de los incidentes y proyectos
    -- asociados al usuario (responsable de gestion o tecnico), activos y sin finalizar (OK_CLIENTE sin dato).
    -- Mas antiguos primero. insistencias = avances posteriores que apuntan a este por CG_AVANCES.ID_ANT; reintento = 0 (sin dato).
    IF @Modo IN ('testing', 'solicitudes')
    BEGIN
        DECLARE @TipoDesc VARCHAR(30);
        SET @TipoDesc = CASE WHEN @Modo = 'testing' THEN 'TESTING' ELSE 'SOLICITUD' END;

        SELECT v.tipo_entidad, v.id_entidad, v.texto, v.responsable, v.cliente, v.fecha_hora,
               CAST(0 AS INT) AS reintento,
               (SELECT COUNT(*) FROM dbo.CG_AVANCES s WHERE s.ID_ANT = v.id_avance AND s.FEC_BAJA IS NULL) AS insistencias
        FROM (
            SELECT CAST('I' AS CHAR(1)) AS tipo_entidad,
                   i.ID                AS id_entidad,
                   av.ID                AS id_avance,
                   CASE WHEN LEN(av.DESCRIPCION) > 300 THEN LEFT(av.DESCRIPCION, 300) + '...' ELSE av.DESCRIPCION END AS texto,
                   ISNULL((SELECT MAX(u.USRNAME) FROM dbo.vw_UserGrupos u WHERE u.USRNET = av.IDUSUARIO_SEG), av.IDUSUARIO_SEG) AS responsable,
                   c.DESCRIPCION        AS cliente,
                   av.FECHA             AS fecha_hora,
                   ISNULL(av.HORA, 0)   AS hora
            FROM dbo.CG_AVANCES av
            INNER JOIN dbo.CG_TIPO_AVANCE t ON CAST(t.ID AS NVARCHAR(50)) = LTRIM(RTRIM(av.TIPO))
            INNER JOIN dbo.CG_INCIDENTES i ON i.ID = av.CODOBJETO
            LEFT JOIN dbo.CG_CLIENTES c ON c.ID = i.ID_CLIENTE
            WHERE av.TIPOOBJETO = 'inc'
              AND av.FEC_BAJA IS NULL
              AND UPPER(LTRIM(RTRIM(t.DESCRIPCION))) = @TipoDesc
              AND i.ESTADO = 1
              AND (i.RESPON_GESTION_SEG = @Respon OR i.RESPON_TEC_SEG = @Respon)
              AND (NULLIF(LTRIM(RTRIM(i.OK_CLIENTE)), '') IS NULL OR LTRIM(RTRIM(i.OK_CLIENTE)) = '0')
            UNION ALL
            SELECT CAST('P' AS CHAR(1)) AS tipo_entidad,
                   p.ID                AS id_entidad,
                   av.ID                AS id_avance,
                   CASE WHEN LEN(av.DESCRIPCION) > 300 THEN LEFT(av.DESCRIPCION, 300) + '...' ELSE av.DESCRIPCION END AS texto,
                   ISNULL((SELECT MAX(u.USRNAME) FROM dbo.vw_UserGrupos u WHERE u.USRNET = av.IDUSUARIO_SEG), av.IDUSUARIO_SEG) AS responsable,
                   c.DESCRIPCION        AS cliente,
                   av.FECHA             AS fecha_hora,
                   ISNULL(av.HORA, 0)   AS hora
            FROM dbo.CG_AVANCES av
            INNER JOIN dbo.CG_TIPO_AVANCE t ON CAST(t.ID AS NVARCHAR(50)) = LTRIM(RTRIM(av.TIPO))
            INNER JOIN dbo.CG_PROYECTOS p ON p.ID = av.CODOBJETO
            LEFT JOIN dbo.CG_CLIENTES c ON c.ID = p.ID_CLIENTE
            WHERE av.TIPOOBJETO = 'proy'
              AND av.FEC_BAJA IS NULL
              AND UPPER(LTRIM(RTRIM(t.DESCRIPCION))) = @TipoDesc
              AND p.ESTADO = 1
              AND (p.RESPON_GESTION_SEG = @Respon OR p.RESPON_TEC_SEG = @Respon)
              AND (NULLIF(LTRIM(RTRIM(p.OK_CLIENTE)), '') IS NULL OR LTRIM(RTRIM(p.OK_CLIENTE)) = '0')
        ) v
        WHERE @Respon IS NOT NULL
        ORDER BY v.fecha_hora ASC, v.hora ASC, v.id_avance ASC;
        RETURN;
    END

    RAISERROR('Modo de sp_cg_Inicio no reconocido.', 16, 1);
END;
GO

-- PROCEDURE sp_cg_Licencias_Listar (modificado)
ALTER PROCEDURE [dbo].[sp_cg_Licencias_Listar]
AS
BEGIN
    SET NOCOUNT ON;

    SELECT a.id,
           ISNULL((SELECT MAX(r.USRNAME) FROM vw_UserGrupos r WHERE r.USRNET = a.id_respon_seg), a.id_respon_seg),
           ISNULL(a.f_inicio, 0),
           ISNULL(a.f_fin, 0), a.motivo
    FROM CG_LICENCIAS a
    WHERE A.ESTADO = 1
    ORDER BY a.id DESC;
END;
GO

-- PROCEDURE sp_cg_Proyectos_ComboCliente (nuevo)
CREATE PROCEDURE [dbo].[sp_cg_Proyectos_ComboCliente]
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

-- PROCEDURE sp_cg_Registro_ActualizarCampos (modificado)
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

-- PROCEDURE sp_cg_Registro_Avance (modificado)
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

-- PROCEDURE sp_cg_Registro_Insertar (modificado)
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
