/* =====================================================================
   CAMBIOS_PERMISOS_ETAPA1.sql
   Corgestion: permisos al estilo E-Manager (tablas Atrm / AtrmUsuarios / AtrmGrupos)

   ESTADO: SCRIPT PARA REVISAR. NO SE EJECUTO. Ejecutar solo despues de revisarlo.
   Base:   Dbcordatv48_UNIF485
   Compatibilidad: nivel 100 (no usa STRING_SPLIT, OFFSET, THROW, TRY_*, CREATE OR ALTER)

   Contenido
     1) sp_cg_Permisos : replica la logica de FuncVarias.ObtenerAtributos2 (E-Manager)
                         sin usar el objeto COM ofv.
     2) Filas de Atrm (tipo 'F') para las pantallas de Corgestion de la etapa 1.
     3) Permisos de ejemplo para los grupos FUNC y TEC (AJUSTAR antes de ejecutar).
     4) Consultas de verificacion y de reversa (comentadas).

   Reglas que implementa sp_cg_Permisos (iguales a E-Manager)
     a) Sin ninguna fila de permiso -> TODO PERMITIDO (valor por defecto).
     b) Si hay fila para el USUARIO (AtrmXUsuario) -> se usa esa fila.
     c) Si no, se usan las filas de los GRUPOS del usuario (AtrmXGrupo):
        un permiso se niega solo si TODOS los grupos del usuario lo niegan.
        Basta un grupo que lo permita. Los grupos del usuario salen de vw_UserGrupos.
     d) Un permiso en NULL en la fila del usuario se toma como NEGADO (igual que en VBScript).
        Un permiso en NULL en la fila de un grupo se toma como "no niega".

   Diferencias respecto de E-Manager (a tener en cuenta)
     - Los grupos del usuario se leen SOLO de vw_UserGrupos (OBJALOG). Todavia no hay
       consulta a AD como segundo paso.
     - Si un grupo tiene filas duplicadas para el mismo formulario, aqui se niega cuando
       las negaciones son >= cantidad de grupos (en VBScript solo cuando son exactamente
       iguales). Solo cambia en ese caso extremo.
     - Estado (AEstado) queda en 1 en la rama de grupos, como en E-Manager.

   Parametros a definir
     - @Aplicacion: codigo de aplicacion de Atrm. Aqui se usa 'COREM20' (la de E-Manager),
       para que los permisos de Corgestion se administren en las mismas pantallas.
       Si se prefiere una aplicacion propia (por ejemplo 'CORGEST') hay que crearla en
       CorAplic/Atrm (tipo 'A') y cambiar el valor aqui y en Corgestion/IncFile/funciones.v1.inc
       (variable CG_ATRM_APLI).
   ===================================================================== */

USE [Dbcordatv48_UNIF485]
GO

/* ---------------------------------------------------------------------
   1) PROCEDIMIENTO
   --------------------------------------------------------------------- */
IF OBJECT_ID(N'dbo.sp_cg_Permisos', N'P') IS NOT NULL
    DROP PROCEDURE dbo.sp_cg_Permisos
GO

CREATE PROCEDURE dbo.sp_cg_Permisos
    @Form       NVARCHAR(200),               -- nombre de la pantalla (= Atrm.depen = Session("Form"))
    @Usuario    NVARCHAR(200),               -- usuario de red (= Usuarios.ID = OBJALOG.USUARIO)
    @Aplicacion NVARCHAR(200) = N'COREM20'   -- Atrm.aplicacion
AS
BEGIN
    SET NOCOUNT ON;

    DECLARE @Estado BIT, @Agregar BIT, @Modificar BIT, @Eliminar BIT, @Consultar BIT;
    DECLARE @Origen VARCHAR(10), @N INT;

    SET @Estado = 1;  SET @Agregar = 1;  SET @Modificar = 1;
    SET @Eliminar = 1;  SET @Consultar = 1;
    SET @Origen = 'DEFAULT';

    -- 1) Permisos propios del usuario
    SELECT TOP 1
           @Estado    = ISNULL(usrestado, 0),
           @Agregar   = ISNULL(usragregar, 0),
           @Modificar = ISNULL(usrmodificar, 0),
           @Eliminar  = ISNULL(usreliminar, 0),
           @Consultar = ISNULL(usrconsultar, 0),
           @Origen    = 'USUARIO'
    FROM dbo.AtrmXUsuario
    WHERE tipo = 'F'
      AND depen = @Form
      AND usuario = @Usuario
      AND aplicacion = @Aplicacion;

    -- 2) Si no hay fila de usuario: permisos por grupo
    IF @Origen = 'DEFAULT'
    BEGIN
        DECLARE @G TABLE (GRUPO NVARCHAR(200) NOT NULL PRIMARY KEY);

        INSERT INTO @G (GRUPO)
        SELECT DISTINCT GRUPO
        FROM dbo.vw_UserGrupos
        WHERE USRNET = @Usuario;

        SELECT @N = COUNT(*) FROM @G;

        IF EXISTS (SELECT 1
                   FROM dbo.AtrmXGrupo x
                   WHERE x.tipo = 'F'
                     AND x.depen = @Form
                     AND x.aplicacion = @Aplicacion
                     AND x.grupo IN (SELECT GRUPO FROM @G))
        BEGIN
            SELECT @Origen    = 'GRUPO',
                   @Consultar = CASE WHEN SUM(CASE WHEN x.grpconsultar = 0 THEN 1 ELSE 0 END) >= @N THEN 0 ELSE 1 END,
                   @Eliminar  = CASE WHEN SUM(CASE WHEN x.grpeliminar  = 0 THEN 1 ELSE 0 END) >= @N THEN 0 ELSE 1 END,
                   @Modificar = CASE WHEN SUM(CASE WHEN x.grpmodificar = 0 THEN 1 ELSE 0 END) >= @N THEN 0 ELSE 1 END,
                   @Agregar   = CASE WHEN SUM(CASE WHEN x.grpagregar   = 0 THEN 1 ELSE 0 END) >= @N THEN 0 ELSE 1 END
            FROM dbo.AtrmXGrupo x
            WHERE x.tipo = 'F'
              AND x.depen = @Form
              AND x.aplicacion = @Aplicacion
              AND x.grupo IN (SELECT GRUPO FROM @G);
        END
    END

    SELECT CAST(@Estado    AS INT) AS Estado,
           CAST(@Agregar   AS INT) AS Agregar,
           CAST(@Modificar AS INT) AS Modificar,
           CAST(@Eliminar  AS INT) AS Eliminar,
           CAST(@Consultar AS INT) AS Consultar,
           @Origen                 AS Origen;
END
GO

/* ---------------------------------------------------------------------
   2) FILAS DE Atrm (tipo 'F') PARA LAS PANTALLAS DE LA ETAPA 1
      atributo : identificador unico del permiso (lo usan AtrmUsuarios / AtrmGrupos)
      depen    : nombre de la pantalla (lo que Corgestion envia como @Form)
      [VERIFICAR] convencion de 'atributo' / 'depen' contra las pantallas de seguridad
      de E-Manager antes de ejecutar.
   --------------------------------------------------------------------- */
DECLARE @Apli NVARCHAR(200);
SET @Apli = N'COREM20';

IF NOT EXISTS (SELECT 1 FROM dbo.Atrm WHERE atributo = N'CG_INCIDENTES_F' AND aplicacion = @Apli)
    INSERT INTO dbo.Atrm (atributo, alias, observaciones, tipo, depen, comprobar, aplicacion)
    VALUES (N'CG_INCIDENTES_F', N'Cor Gestion - Incidentes', N'Corgestion etapa 1', N'F', N'CG_INCIDENTES', 1, @Apli);

IF NOT EXISTS (SELECT 1 FROM dbo.Atrm WHERE atributo = N'CG_PROYECTOS_F' AND aplicacion = @Apli)
    INSERT INTO dbo.Atrm (atributo, alias, observaciones, tipo, depen, comprobar, aplicacion)
    VALUES (N'CG_PROYECTOS_F', N'Cor Gestion - Proyectos', N'Corgestion etapa 1', N'F', N'CG_PROYECTOS', 1, @Apli);

/* ---------------------------------------------------------------------
   3) PERMISOS DE EJEMPLO PARA LOS GRUPOS FUNC Y TEC
      *** AJUSTAR LOS VALORES (1 = permite, 0 = niega) ANTES DE EJECUTAR ***
      Los valores de abajo son solo un punto de partida (todo permitido), no una
      definicion funcional.
      Los grupos deben existir en dbo.Grupos y coincidir con el texto de OBJALOG.GRUPOS.
   --------------------------------------------------------------------- */
DECLARE @Perm TABLE (Grupo NVARCHAR(200), Atributo NVARCHAR(400),
                     Estado BIT, Agregar BIT, Modificar BIT, Eliminar BIT, Consultar BIT);

--                  Grupo   Atributo            Estado Agregar Modificar Eliminar Consultar
INSERT INTO @Perm VALUES (N'FUNC', N'CG_INCIDENTES_F', 1, 1, 1, 1, 1);
INSERT INTO @Perm VALUES (N'TEC',  N'CG_INCIDENTES_F', 1, 1, 1, 1, 1);
INSERT INTO @Perm VALUES (N'FUNC', N'CG_PROYECTOS_F',  1, 1, 1, 1, 1);
INSERT INTO @Perm VALUES (N'TEC',  N'CG_PROYECTOS_F',  1, 1, 1, 1, 1);

INSERT INTO dbo.AtrmGrupos (Grupo, Atributo, Estado, Agregar, Modificar, Eliminar, Consultar)
SELECT p.Grupo, p.Atributo, p.Estado, p.Agregar, p.Modificar, p.Eliminar, p.Consultar
FROM @Perm p
WHERE NOT EXISTS (SELECT 1 FROM dbo.AtrmGrupos g
                  WHERE g.Grupo = p.Grupo AND g.Atributo = p.Atributo);
GO

/* ---------------------------------------------------------------------
   4) VERIFICACION (ejecutar a mano, con un usuario real)
   --------------------------------------------------------------------- */
-- SELECT * FROM dbo.Atrm WHERE atributo LIKE N'CG[_]%';
-- SELECT * FROM dbo.AtrmGrupos WHERE Atributo LIKE N'CG[_]%';
-- SELECT GRUPO FROM dbo.vw_UserGrupos WHERE USRNET = N'USUARIO_DE_PRUEBA';
-- EXEC dbo.sp_cg_Permisos @Form = N'CG_INCIDENTES', @Usuario = N'USUARIO_DE_PRUEBA';
-- Para probar una negacion: poner Agregar = 0 en AtrmGrupos para TODOS los grupos del usuario.

/* ---------------------------------------------------------------------
   REVERSA (comentada)
   --------------------------------------------------------------------- */
-- DELETE FROM dbo.AtrmGrupos WHERE Atributo IN (N'CG_INCIDENTES_F', N'CG_PROYECTOS_F');
-- DELETE FROM dbo.Atrm       WHERE atributo IN (N'CG_INCIDENTES_F', N'CG_PROYECTOS_F');
-- DROP PROCEDURE dbo.sp_cg_Permisos;
