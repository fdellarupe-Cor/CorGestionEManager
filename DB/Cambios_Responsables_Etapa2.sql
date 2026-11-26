/* =====================================================================
   CAMBIOS_RESPONSABLES_ETAPA2.sql
   sp_cg_Responsables: pasa de leer CG_COR_RESPONSABLES a leer vw_UserGrupos

   ESTADO: SCRIPT PARA REVISAR. NO SE EJECUTO. Ejecutar solo despues de revisarlo.
   Base:   Dbcordatv48_UNIF485
   Compatibilidad: nivel 100 (sin STRING_SPLIT, sin THROW, sin TRY_*, sin CREATE OR ALTER)
   Requiere: vw_UserGrupos (ya existe en la base)

   Criterios aplicados (AJUSTAR si no coinciden con lo definido)
     1) Funcion (FUNC / TEC) = pertenencia al grupo del mismo nombre en vw_UserGrupos
        (vw_UserGrupos.GRUPO). Si una persona esta en FUNC y en TEC, la funcion
        "unica" que se devuelve es FUNC. [DECIDIR]
     2) "Activo" = existe en vw_UserGrupos. ACTIVO se devuelve siempre 'S'.
     3) Universo de personas de los listados = personas que pertenecen a FUNC o TEC
        (equivale al universo anterior de CG_COR_RESPONSABLES, donde todos tenian funcion).
        Las busquedas de UN login puntual (id_nombre, nombre_id, mail_id, id_seg)
        buscan en toda la vista. [DECIDIR]
     4) Nombre = USRNAME; si esta vacio se usa el usuario de red (que es descriptivo).
        APELLIDO y TEL se devuelven vacios (no existen en vw_UserGrupos).
     5) Misma cantidad y orden de columnas que antes en cada modo, porque los ASP leen
        por posicion (rs(0), rs(1), ...). Excepcion: listar_funcion devuelve
        ID, USRNAME, GRUPO, MAIL (los callers solo leen rs(0) y rs(1)).
     6) id_seg y seguridad devuelven como "ID numerico" un numero DERIVADO del usuario de
        red: ABS(CHECKSUM(usuario)). Sirve para que UsrAcc/Esta (CG_ESTADISTICAS, TIPO 'U')
        sigan funcionando sin la tabla. NO coincide con los ids anteriores de
        CG_COR_RESPONSABLES (las estadisticas de usuario previas no se vinculan). [DECIDIR]
     7) Modos de escritura (insertar, actualizar) -> error claro; los responsables se
        administran en AD / E-Manager. ultimo_id_funcion devuelve un resultado vacio.
     8) Los modos ...not_in / nombre_id_in reemplazan STRING_SPLIT con CHARINDEX sobre
        la lista de logins separada por comas.
   ===================================================================== */

USE [Dbcordatv48_UNIF485]
GO

ALTER PROCEDURE [dbo].[sp_cg_Responsables]
    @Modo      VARCHAR(25),
    @Id        INT           = NULL,
    @Funcion   VARCHAR(50)   = NULL,
    @IdSeg     NVARCHAR(100) = NULL,   -- usuario de red
    @ExcluirId NVARCHAR(100) = NULL,   -- usuario de red a excluir de los listados
    @Ids       NVARCHAR(2000)= NULL,   -- lista de usuarios de red separados por coma
    @Nombre    NVARCHAR(100) = NULL,
    @Apellido  NVARCHAR(100) = NULL,
    @Tel       NVARCHAR(100) = NULL,
    @Mail      NVARCHAR(300) = NULL,
    @Activo    NVARCHAR(10)  = NULL
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @Lista NVARCHAR(2100);
    SET @Lista = N',' + REPLACE(ISNULL(@Ids, N''), N' ', N'') + N',';

    -- Una fila por persona (cualquier grupo). FUNCION = FUNC / TEC / NULL
    DECLARE @P TABLE (
        USRNET  NVARCHAR(200) NOT NULL PRIMARY KEY,
        NOMBRE  NVARCHAR(200) NULL,
        MAIL    NVARCHAR(300) NULL,
        FUNCION NVARCHAR(50)  NULL
    );

    IF @Modo NOT IN ('listar', 'listar_funcion', 'combo_funcion', 'insertar', 'actualizar', 'ultimo_id_funcion')
    BEGIN
        INSERT INTO @P (USRNET, NOMBRE, MAIL, FUNCION)
        SELECT USRNET,
               ISNULL(NULLIF(MAX(USRNAME), N''), USRNET),
               MAX(MAIL),
               CASE WHEN MAX(CASE WHEN GRUPO = N'FUNC' THEN 1 ELSE 0 END) = 1 THEN N'FUNC'
                    WHEN MAX(CASE WHEN GRUPO = N'TEC'  THEN 1 ELSE 0 END) = 1 THEN N'TEC'
                    ELSE NULL END
        FROM dbo.vw_UserGrupos
        WHERE ISNULL(USRNET, N'') <> N''
        GROUP BY USRNET;
    END

    -- Una fila por persona y funcion (FUNC / TEC). 11 columnas, igual que la tabla anterior.
    IF @Modo = 'listar'
        SELECT USRNET AS ID,
               ISNULL(NULLIF(MAX(USRNAME), N''), USRNET) AS NOMBRE,
               CAST(N'' AS NVARCHAR(100))                AS APELLIDO,
               GRUPO                                     AS FUNCION,
               USRNET                                    AS ID_SEG,
               MAX(MAIL)                                 AS MAIL,
               CAST(N'' AS NVARCHAR(100))                AS TEL,
               CAST(N'S' AS NVARCHAR(10))                AS ACTIVO,
               CAST(1 AS BIT)                            AS ESTADO,
               CAST(NULL AS DATETIME)                    AS FEC_BAJA,
               CAST(NULL AS VARCHAR(100))                AS USU_BAJA
        FROM dbo.vw_UserGrupos
        WHERE ISNULL(USRNET, N'') <> N'' AND GRUPO IN (N'FUNC', N'TEC')
        GROUP BY USRNET, GRUPO
        ORDER BY USRNET, GRUPO;

    ELSE IF @Modo = 'listar_funcion'
        SELECT USRNET AS ID,
               ISNULL(NULLIF(MAX(USRNAME), N''), USRNET) AS USRNAME,
               GRUPO,
               MAX(MAIL) AS MAIL
        FROM dbo.vw_UserGrupos
        WHERE GRUPO = @Funcion AND ISNULL(USRNET, N'') <> N''
        GROUP BY USRNET, GRUPO
        ORDER BY USRNET;

    -- Mismas 11 columnas que 'listar', una fila por persona (FUNC / TEC)
    ELSE IF @Modo IN ('listar_sin', 'listar_activos_sin')
        SELECT USRNET                      AS ID,
               NOMBRE,
               CAST(N'' AS NVARCHAR(100))  AS APELLIDO,
               FUNCION,
               USRNET                      AS ID_SEG,
               MAIL,
               CAST(N'' AS NVARCHAR(100))  AS TEL,
               CAST(N'S' AS NVARCHAR(10))  AS ACTIVO,
               CAST(1 AS BIT)              AS ESTADO,
               CAST(NULL AS DATETIME)      AS FEC_BAJA,
               CAST(NULL AS VARCHAR(100))  AS USU_BAJA
        FROM @P
        WHERE FUNCION IS NOT NULL AND USRNET <> ISNULL(@ExcluirId, N'')
        ORDER BY USRNET;

    ELSE IF @Modo IN ('listar_ids_sin', 'listar_ids_activos_sin')
        SELECT USRNET AS ID_SEG
        FROM @P
        WHERE FUNCION IS NOT NULL AND USRNET <> ISNULL(@ExcluirId, N'')
        ORDER BY USRNET;

    ELSE IF @Modo = 'nombre_id_in'
        SELECT NOMBRE, USRNET
        FROM @P
        WHERE CHARINDEX(N',' + USRNET + N',', @Lista) > 0
        ORDER BY USRNET;

    ELSE IF @Modo IN ('listar_not_in', 'listar_not_in_act')
        SELECT USRNET, NOMBRE
        FROM @P
        WHERE FUNCION IS NOT NULL
          AND USRNET <> ISNULL(@ExcluirId, N'')
          AND CHARINDEX(N',' + USRNET + N',', @Lista) = 0
        ORDER BY USRNET;

    ELSE IF @Modo = 'id_nombre'
        SELECT USRNET, NOMBRE FROM @P WHERE USRNET = @IdSeg;

    ELSE IF @Modo = 'nombre_id'
        SELECT NOMBRE FROM @P WHERE USRNET = @IdSeg;

    ELSE IF @Modo IN ('id_nombre_sin', 'id_nombre_act_sin')
        SELECT USRNET, NOMBRE
        FROM @P
        WHERE FUNCION IS NOT NULL AND USRNET <> ISNULL(@ExcluirId, N'')
        ORDER BY USRNET;

    ELSE IF @Modo = 'mail_id'
        SELECT MAIL FROM @P WHERE USRNET = @IdSeg;

    ELSE IF @Modo = 'id_funcion'
        SELECT USRNET, FUNCION FROM @P WHERE FUNCION IS NOT NULL ORDER BY USRNET;

    ELSE IF @Modo = 'id_funcion_id'
        SELECT USRNET, FUNCION FROM @P WHERE FUNCION IS NOT NULL AND USRNET = @IdSeg;

    -- Ya no existe el "ultimo alta": resultado vacio con las mismas columnas
    ELSE IF @Modo = 'ultimo_id_funcion'
        SELECT CAST(N'' AS NVARCHAR(200)) AS ID_SEG, CAST(N'' AS NVARCHAR(50)) AS funcion WHERE 1 = 0;

    ELSE IF @Modo = 'combo_todos'
        SELECT USRNET, NOMBRE, FUNCION FROM @P WHERE FUNCION IS NOT NULL ORDER BY USRNET;

    ELSE IF @Modo = 'combo_funcion'
        SELECT USRNET AS ID_SEG,
               ISNULL(NULLIF(MAX(USRNAME), N''), USRNET) AS NOMBRE
        FROM dbo.vw_UserGrupos
        WHERE GRUPO = @Funcion AND ISNULL(USRNET, N'') <> N''
        GROUP BY USRNET
        ORDER BY USRNET;

    -- Solo quien pertenece a FUNC o TEC tiene sesion funcional en Corgestion
    ELSE IF @Modo = 'seguridad'
        SELECT FUNCION,
               USRNET AS ID_SEG,
               ABS(CHECKSUM(USRNET)) AS ID
        FROM @P
        WHERE FUNCION IS NOT NULL AND USRNET = @IdSeg;

    -- Numero derivado del usuario de red (para CG_ESTADISTICAS TIPO 'U')
    ELSE IF @Modo = 'id_seg'
        SELECT ABS(CHECKSUM(USRNET)) AS ID FROM @P WHERE USRNET = @IdSeg;

    ELSE IF @Modo IN ('insertar', 'actualizar')
        RAISERROR('Los responsables ya no se administran en Corgestion: se determinan por los grupos FUNC/TEC (AD / E-Manager).', 16, 1);
END
GO

/* ---------------------------------------------------------------------
   VERIFICACION (ejecutar a mano con un usuario real)
   --------------------------------------------------------------------- */
-- EXEC dbo.sp_cg_Responsables @Modo = 'listar_funcion', @Funcion = 'FUNC';
-- EXEC dbo.sp_cg_Responsables @Modo = 'listar';
-- EXEC dbo.sp_cg_Responsables @Modo = 'combo_todos';
-- EXEC dbo.sp_cg_Responsables @Modo = 'listar_activos_sin', @ExcluirId = 'USUARIO_DE_PRUEBA';
-- EXEC dbo.sp_cg_Responsables @Modo = 'nombre_id_in', @Ids = 'USUARIO1,USUARIO2';
-- EXEC dbo.sp_cg_Responsables @Modo = 'seguridad', @IdSeg = 'USUARIO_DE_PRUEBA';
-- EXEC dbo.sp_cg_Responsables @Modo = 'id_seg', @IdSeg = 'USUARIO_DE_PRUEBA';
/* REVERSA: volver a ejecutar el ALTER PROCEDURE original de sp_cg_Responsables
   (esta en DB/Corv48UNIFV5_DB.sql, version del 2026-10-09). */
