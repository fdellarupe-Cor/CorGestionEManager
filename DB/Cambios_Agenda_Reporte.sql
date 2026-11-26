/* =====================================================================
   Cambios_Agenda_Reporte.sql  -  datos de reportesAgenda.asp (procedimiento NUEVO: no modifica ninguno existente)

   ESTADO: SCRIPT PARA REVISAR. NO SE EJECUTO.
   Base destino: Dbcordatv48_UNIF485     Compatibilidad: nivel 100 (sin TRY_CAST, STRING_AGG ni OPENJSON)

   sp_cg_AgendaReporte  @Modo = 'resumen' | 'detalle' | 'arbol'
     Filtros (todos opcionales):
       @Usuario  login (ID_SEG) ; NULL = todos los usuarios
       @Rol      'A' responsable o participante (indistinto) | 'R' solo responsable | 'P' solo participante
       @Estado   'PENDIENTE' | 'POSTERGADA' | 'REALIZADA' ; NULL = todos
       @Desde / @Hasta  dias desde 1900-01-01 (FECHA_REUNION) ; NULL = sin limite
     'resumen' : una fila: TOTAL, PENDIENTES, POSTERGADAS, REALIZADAS, CON_SEGUIMIENTO
                 (CON_SEGUIMIENTO = reuniones con al menos un tema marcado "Requiere seguimiento": cada reunion cuenta 1).
     'detalle' : una fila por reunion del filtro.
     'arbol'   : arbol de seguimientos. Raiz = tema con SEGUIMIENTO = 1 de una reunion del filtro que no nace de
                 un seguimiento de otro tema. Rama = tema de una reunion generada por "Generar seguimiento"
                 (TIPO_ORIGEN = 'S', ID_AGENDA_PADRE = reunion del tema anterior) con el MISMO TIPO_RELACION y
                 el MISMO TEXTO DE TEMA. Las ramas se muestran aunque la reunion quede fuera del periodo o del
                 usuario filtrado (EN_FILTRO = 0).
   REVERSA : DROP PROCEDURE dbo.sp_cg_AgendaReporte
   ===================================================================== */

USE [Dbcordatv48_UNIF485]
GO

IF OBJECT_ID(N'dbo.sp_cg_AgendaReporte', N'P') IS NULL
    EXEC('CREATE PROCEDURE dbo.sp_cg_AgendaReporte AS RETURN 0');
GO

ALTER PROCEDURE [dbo].[sp_cg_AgendaReporte]
    @Modo    varchar(15),
    @Usuario nvarchar(100) = NULL,
    @Rol     char(1)       = 'A',
    @Estado  varchar(15)   = NULL,
    @Desde   int           = NULL,
    @Hasta   int           = NULL
AS
BEGIN
    SET NOCOUNT ON;

    IF @Usuario IS NOT NULL AND LTRIM(RTRIM(@Usuario)) = '' SET @Usuario = NULL;
    IF @Estado  IS NOT NULL AND LTRIM(RTRIM(@Estado))  = '' SET @Estado  = NULL;
    IF @Rol IS NULL OR @Rol NOT IN ('A', 'R', 'P') SET @Rol = 'A';

    -- Reuniones del filtro
    CREATE TABLE #reu (ID int PRIMARY KEY, ES_RESP bit NOT NULL);
    INSERT INTO #reu (ID, ES_RESP)
    SELECT a.ID,
           CASE WHEN @Usuario IS NOT NULL AND a.ID_RESPON_SEG = @Usuario THEN 1 ELSE 0 END
    FROM dbo.CG_AGENDA AS a
    WHERE a.FEC_BAJA IS NULL
      AND (@Desde IS NULL OR a.FECHA_REUNION >= @Desde)
      AND (@Hasta IS NULL OR a.FECHA_REUNION <= @Hasta)
      AND (@Estado IS NULL OR ISNULL(a.ESTADO_AGENDA, 'PENDIENTE') = @Estado)
      AND (@Usuario IS NULL
           OR (@Rol IN ('A', 'R') AND a.ID_RESPON_SEG = @Usuario)
           OR (@Rol IN ('A', 'P') AND EXISTS (SELECT 1 FROM dbo.CG_AGENDA_CORRESPONSABLE c
                                              WHERE c.ID_AGENDA = a.ID AND c.FEC_BAJA IS NULL
                                                AND c.ID_PERSONA_SEG = @Usuario)));

    IF @Modo = 'resumen'
    BEGIN
        -- El agregado no puede contener una subconsulta: se calcula antes en una tabla derivada
        SELECT COUNT(*) AS TOTAL,
               ISNULL(SUM(CASE WHEN x.ESTADO = 'PENDIENTE' THEN 1 ELSE 0 END), 0) AS PENDIENTES,
               ISNULL(SUM(CASE WHEN x.ESTADO = 'POSTERGADA' THEN 1 ELSE 0 END), 0) AS POSTERGADAS,
               ISNULL(SUM(CASE WHEN x.ESTADO = 'REALIZADA' THEN 1 ELSE 0 END), 0) AS REALIZADAS,
               ISNULL(SUM(x.CON_SEG), 0) AS CON_SEGUIMIENTO
        FROM (SELECT ISNULL(a.ESTADO_AGENDA, 'PENDIENTE') AS ESTADO,
                     CASE WHEN EXISTS (SELECT 1 FROM dbo.CG_AGENDA_SUBTEMA s
                                       WHERE s.ID_AGENDA = a.ID AND s.FEC_BAJA IS NULL AND s.SEGUIMIENTO = 1)
                          THEN 1 ELSE 0 END AS CON_SEG
              FROM #reu r
              INNER JOIN dbo.CG_AGENDA a ON a.ID = r.ID) AS x;
        RETURN;
    END

    IF @Modo = 'detalle'
    BEGIN
        SELECT a.ID, a.FECHA_REUNION, a.HORA_INICIO, a.HORA_FIN, a.ASUNTO,
               RTRIM(ISNULL(p.NOMBRE, '') + ' ' + ISNULL(p.APELLIDO, '')) AS RESPONSABLE,
               a.ID_RESPON_SEG AS RESPONSABLE_LOGIN,
               CASE WHEN @Usuario IS NULL THEN '' WHEN r.ES_RESP = 1 THEN 'Responsable' ELSE 'Participante' END AS PARTICIPACION,
               ISNULL(a.ESTADO_AGENDA, 'PENDIENTE') AS ESTADO,
               a.ID_AGENDA_PADRE,
               (SELECT COUNT(*) FROM dbo.CG_AGENDA_SUBTEMA s
                WHERE s.ID_AGENDA = a.ID AND s.FEC_BAJA IS NULL AND s.SEGUIMIENTO = 1) AS TEMAS_SEG
        FROM #reu r
        INNER JOIN dbo.CG_AGENDA a ON a.ID = r.ID
        LEFT JOIN dbo.CG_COR_RESPONSABLES p ON p.ID_SEG = a.ID_RESPON_SEG
        ORDER BY a.FECHA_REUNION, a.HORA_INICIO, a.ID;
        RETURN;
    END

    IF @Modo = 'arbol'
    BEGIN
        -- Enlace hijo -> padre: tema de una reunion de seguimiento (TIPO_ORIGEN 'S') con el mismo tipo y texto
        -- que un tema marcado con seguimiento de la reunion de origen
        CREATE TABLE #enlace (HIJO int PRIMARY KEY, PADRE int NOT NULL);
        INSERT INTO #enlace (HIJO, PADRE)
        SELECT c.ID_SUBTEMA, MIN(p.ID_SUBTEMA)
        FROM dbo.CG_AGENDA_SUBTEMA c
        INNER JOIN dbo.CG_AGENDA ac ON ac.ID = c.ID_AGENDA AND ac.FEC_BAJA IS NULL AND ac.TIPO_ORIGEN = 'S'
        INNER JOIN dbo.CG_AGENDA_SUBTEMA p ON p.ID_AGENDA = ac.ID_AGENDA_PADRE
                                          AND p.FEC_BAJA IS NULL AND p.SEGUIMIENTO = 1
                                          AND p.TIPO_RELACION = c.TIPO_RELACION
                                          AND LTRIM(RTRIM(p.TEMA_A_TRATAR)) = LTRIM(RTRIM(c.TEMA_A_TRATAR))
        WHERE c.FEC_BAJA IS NULL
        GROUP BY c.ID_SUBTEMA;

        WITH T (ID_SUBTEMA, ID_RAIZ, NIVEL, RUTA) AS
        (
            SELECT s.ID_SUBTEMA, s.ID_SUBTEMA, 0,
                   CAST(RIGHT('0000000000' + CAST(s.ID_SUBTEMA AS varchar(10)), 10) AS varchar(900))
            FROM dbo.CG_AGENDA_SUBTEMA s
            INNER JOIN #reu r ON r.ID = s.ID_AGENDA
            WHERE s.FEC_BAJA IS NULL AND s.SEGUIMIENTO = 1
              AND NOT EXISTS (SELECT 1 FROM #enlace e WHERE e.HIJO = s.ID_SUBTEMA)
            UNION ALL
            SELECT e.HIJO, t.ID_RAIZ, t.NIVEL + 1,
                   CAST(t.RUTA + RIGHT('0000000000' + CAST(e.HIJO AS varchar(10)), 10) AS varchar(900))
            FROM T t
            INNER JOIN #enlace e ON e.PADRE = t.ID_SUBTEMA
            WHERE t.NIVEL < 30
        )
        SELECT t.ID_RAIZ, t.NIVEL, t.RUTA,
               s.ID_SUBTEMA, s.ID_AGENDA, a.ID_AGENDA_PADRE, a.ASUNTO,
               a.FECHA_REUNION, a.HORA_INICIO,
               ISNULL(a.ESTADO_AGENDA, 'PENDIENTE') AS ESTADO,
               RTRIM(ISNULL(p.NOMBRE, '') + ' ' + ISNULL(p.APELLIDO, '')) AS RESPONSABLE,
               s.TIPO_RELACION, s.ID_PROYECTO, s.ID_INCIDENTE, s.TEMA_A_TRATAR, s.CONCLUSION, s.SEGUIMIENTO,
               CASE WHEN r.ID IS NULL THEN 0 ELSE 1 END AS EN_FILTRO
        FROM T t
        INNER JOIN dbo.CG_AGENDA_SUBTEMA s ON s.ID_SUBTEMA = t.ID_SUBTEMA
        INNER JOIN dbo.CG_AGENDA a ON a.ID = s.ID_AGENDA
        LEFT JOIN dbo.CG_COR_RESPONSABLES p ON p.ID_SEG = a.ID_RESPON_SEG
        LEFT JOIN #reu r ON r.ID = a.ID
        ORDER BY t.ID_RAIZ, t.RUTA;
        RETURN;
    END

    RAISERROR('Modo de sp_cg_AgendaReporte no reconocido.', 16, 1);
END;
GO
