/* =====================================================================
   Cambios_Inicio_Testing_Solicitudes.sql  -  "Testing pendientes" y "Solicitudes pendientes" de inicio.asp

   ESTADO: SCRIPT PARA REVISAR. NO SE EJECUTO.
   Base destino: Dbcordatv48_UNIF485     Compatibilidad: nivel 100
   Base de partida: sp_cg_Inicio del export DB/Corv48UNIFV5_DB.sql del 22/10/2026 11:28 (modos 'casos' y 'casos_cuenta'
                    quedan IGUALES). OJO: el ALTER reemplaza el procedimiento completo; si lo modificaste despues, reexportalo.

   NUEVO en sp_cg_Inicio (@Modo, @Respon = login de red del usuario):
     'testing'     : avances cuyo tipo (CG_TIPO_AVANCE.DESCRIPCION) es "Testing".
     'solicitudes' : avances cuyo tipo es "Solicitud".
   Criterio comun:
     - De incidentes (TIPOOBJETO 'inc') y proyectos ('proy') asociados al usuario: responsable de gestion o tecnico.
     - Casos activos (ESTADO = 1) y sin finalizar (OK_CLIENTE sin dato), igual que "Casos asignados".
     - Avances no dados de baja (FEC_BAJA IS NULL). El tipo se compara sin distinguir mayusculas ni espacios.
     - Orden: mas antiguos primero.
   Columnas: tipo_entidad ('I'/'P'), id_entidad, texto (recortado a 300 caracteres), responsable (quien cargo el avance, de
     vw_UserGrupos), cliente, fecha_hora (dias desde 1900), reintento (0: todavia no hay dato) e insistencias (cantidad de avances
     cuyo ID_ANT apunta a este; 0 hasta que se active "Insistir").
   SI NO SE EJECUTA : los dos bloques siguen mostrando "Falta dato".
   REVERSA    : volver a ejecutar la definicion original de sp_cg_Inicio (DB/Corv48UNIFV5_DB.sql o Cambios_Inicio_CasosAsignados.sql).
   ===================================================================== */

USE [Dbcordatv48_UNIF485]
GO

ALTER PROCEDURE [dbo].[sp_cg_Inicio]
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
