/* =====================================================================
   Cambios_Inicio_CasosAsignados.sql  -  sp_cg_Inicio (pendientes del usuario para inicio.asp)

   ESTADO: SCRIPT PARA REVISAR. NO SE EJECUTO.
   Base destino: Dbcordatv48_UNIF485     Compatibilidad: nivel 100

   QUE CREA   : el procedimiento  sp_cg_Inicio  con dos modos:
                'casos'         lista los casos pendientes del usuario (incidentes + proyectos en una sola lista)
                'casos_cuenta'  cantidad de incidentes y de proyectos pendientes del usuario
   PARA QUE   : la pantalla inicio.asp muestra "Pendientes" (lista con pestanas Todos / Incidentes / Proyectos)
                y "Casos asignados" en la barra superior.
   CRITERIOS  :
     - Asignado   : el usuario (login de red) es el responsable de gestion o el tecnico del caso
                    (RESPON_GESTION_SEG o RESPON_TEC_SEG), igual que el filtro por responsable de los listados.
     - Pendiente  : OK_CLIENTE sin dato (NULL, vacio o '0'). Con dato = finalizado.
     - Solo casos activos (ESTADO = 1).
   MODO 'casos' - columnas:
     tipo ('I' incidente / 'P' proyecto), id, titulo, cliente, prioridad (texto), prioridad_id,
     fecha_ingreso (dias desde 01-01-1900), fecha_compromiso (por ahora NULL: queda reservada para el SLA,
     que se esta trabajando; con ella el semaforo de inicio.asp se calcula solo).
     Orden: prioridad (la mas urgente primero), luego fecha de ingreso (la mas antigua primero), luego id.
   MODO 'casos_cuenta': una fila con INCIDENTES y PROYECTOS (int).
   Si @Respon es NULL o vacio no devuelve casos (0 en la cuenta).
   SI NO SE EJECUTA : inicio.asp sigue funcionando y muestra "Falta dato" en "Pendientes".
   NO TOCA    : ninguna tabla ni dato. Solo crea (o reemplaza) este procedimiento.
   REVERSA    : DROP PROCEDURE dbo.sp_cg_Inicio;
   ===================================================================== */

USE [Dbcordatv48_UNIF485]
GO

-- Crear el procedimiento vacio si no existe (funciona en cualquier version; luego se define con ALTER)
IF OBJECT_ID(N'dbo.sp_cg_Inicio', N'P') IS NULL
    EXEC (N'CREATE PROCEDURE dbo.sp_cg_Inicio AS BEGIN SET NOCOUNT ON; END');
GO

ALTER PROCEDURE dbo.sp_cg_Inicio
    @Modo   VARCHAR(20),                 -- 'casos' | 'casos_cuenta'
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

    RAISERROR('Modo de sp_cg_Inicio no reconocido.', 16, 1);
END;
GO
