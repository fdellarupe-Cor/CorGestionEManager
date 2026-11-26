/* =====================================================================================================
   Permisos_Alta_Formularios.sql
   Da de alta en E-Manager (Atrm / AtrmUsuarios / AtrmGrupos) las pantallas de Corgestion que ahora se
   controlan con getPermisosAtrm("CG_xxx", ...) -> sp_cg_Permisos (Aplicacion COREM20).

   Modelo: cada pantalla nueva COPIA los permisos de la pantalla modelo CG_INCIDENTES (mismo usuario/grupo,
   mismos Estado/Agregar/Modificar/Eliminar/Consultar). Asi nadie gana ni pierde acceso al pasar de
   getPermisos (CG_MENU_X_USUARIO) a getPermisosAtrm.

   IMPORTANTE: sp_cg_Permisos devuelve TODO EN 1 si la pantalla no tiene filas. Ejecutar este script ANTES
   de publicar los .asp, o las pantallas quedan abiertas a todos.

   Ejecutar por el usuario. Termina en ROLLBACK: revisar los resultados y cambiar a COMMIT TRAN.
   Repetible (no duplica lo ya cargado). Ejecutar en cada base (desarrollo/test y produccion).
   ===================================================================================================== */
SET NOCOUNT ON;

DECLARE @Modelo NVARCHAR(200) = N'CG_INCIDENTES';
DECLARE @Apli   NVARCHAR(200) = N'COREM20';

/* ---- Diagnostico previo: la pantalla modelo debe existir ---- */
SELECT 'MODELO Atrm' AS que, a.ID, a.atributo, a.tipo, a.depen, a.aplicacion
FROM dbo.Atrm a WHERE a.tipo = 'F' AND a.depen = @Modelo AND a.aplicacion = @Apli;

SELECT 'MODELO usuarios' AS que, COUNT(*) AS filas
FROM dbo.AtrmUsuarios u WHERE u.Atributo IN (SELECT atributo FROM dbo.Atrm WHERE tipo='F' AND depen=@Modelo AND aplicacion=@Apli);
SELECT 'MODELO grupos' AS que, COUNT(*) AS filas
FROM dbo.AtrmGrupos g WHERE g.Atributo IN (SELECT atributo FROM dbo.Atrm WHERE tipo='F' AND depen=@Modelo AND aplicacion=@Apli);

IF NOT EXISTS (SELECT 1 FROM dbo.Atrm WHERE tipo = 'F' AND depen = @Modelo AND aplicacion = @Apli)
BEGIN
    RAISERROR('No existe la pantalla modelo CG_INCIDENTES en Atrm (aplicacion COREM20). No se hace nada.', 16, 1);
    RETURN;
END

DECLARE @Forms TABLE (Form NVARCHAR(200) PRIMARY KEY);
INSERT INTO @Forms (Form) VALUES
 (N'CG_APL'), (N'CG_CLIENTES'), (N'CG_APLICATIVOS'), (N'CG_BROWSERS'), (N'CG_CONTACTOS'), (N'CG_RESPONSABLES'),
 (N'CG_ESTADOS'), (N'CG_PRIORIDADES'), (N'CG_TIPO_AVANCE'), (N'CG_VIA_PEDIDO'), (N'CG_LICENCIAS'), (N'CG_AGENDA'),
 (N'CG_CONSULTAS'), (N'CG_ESTADISTICAS'), (N'CG_REPORTES'), (N'CG_PERMISOS');
/* CG_PROYECTOS ya existe y no se toca. */

/* Tabla de equivalencias atributo modelo -> atributo nuevo, por pantalla */
DECLARE @Map TABLE (Form NVARCHAR(200), AtrModelo NVARCHAR(400), AtrNuevo NVARCHAR(400));
INSERT INTO @Map (Form, AtrModelo, AtrNuevo)
SELECT f.Form, a.atributo,
       CASE WHEN CHARINDEX(@Modelo, a.atributo) > 0 THEN REPLACE(a.atributo, @Modelo, f.Form)
            ELSE a.atributo + N'|' + f.Form END
FROM @Forms f
CROSS JOIN dbo.Atrm a
WHERE a.tipo = 'F' AND a.depen = @Modelo AND a.aplicacion = @Apli;

BEGIN TRAN;

/* 1) Atrm: la pantalla (solo si no existe) */
INSERT INTO dbo.Atrm (atributo, alias, observaciones, tipo, depen, comprobar, aplicacion, nodo, valor)
SELECT m.AtrNuevo, m.Form, a.observaciones, a.tipo, m.Form, a.comprobar, a.aplicacion, a.nodo, a.valor
FROM @Map m
JOIN dbo.Atrm a ON a.atributo = m.AtrModelo AND a.tipo = 'F' AND a.depen = @Modelo AND a.aplicacion = @Apli
WHERE NOT EXISTS (SELECT 1 FROM dbo.Atrm x WHERE x.tipo = 'F' AND x.depen = m.Form AND x.aplicacion = @Apli);
SELECT 'Atrm insertadas' AS que, @@ROWCOUNT AS filas;

/* 2) AtrmUsuarios: copia de los permisos de usuario del modelo */
INSERT INTO dbo.AtrmUsuarios (Usuario, Atributo, Estado, Agregar, Modificar, Eliminar, Consultar)
SELECT u.Usuario, m.AtrNuevo, u.Estado, u.Agregar, u.Modificar, u.Eliminar, u.Consultar
FROM @Map m
JOIN dbo.AtrmUsuarios u ON u.Atributo = m.AtrModelo
WHERE NOT EXISTS (SELECT 1 FROM dbo.AtrmUsuarios x WHERE x.Atributo = m.AtrNuevo AND x.Usuario = u.Usuario);
SELECT 'AtrmUsuarios insertadas' AS que, @@ROWCOUNT AS filas;

/* 3) AtrmGrupos: copia de los permisos de grupo del modelo */
INSERT INTO dbo.AtrmGrupos (Grupo, Atributo, Estado, Agregar, Modificar, Eliminar, Consultar)
SELECT g.Grupo, m.AtrNuevo, g.Estado, g.Agregar, g.Modificar, g.Eliminar, g.Consultar
FROM @Map m
JOIN dbo.AtrmGrupos g ON g.Atributo = m.AtrModelo
WHERE NOT EXISTS (SELECT 1 FROM dbo.AtrmGrupos x WHERE x.Atributo = m.AtrNuevo AND x.Grupo = g.Grupo);
SELECT 'AtrmGrupos insertadas' AS que, @@ROWCOUNT AS filas;

/* Resultado: permisos por pantalla (usuario y grupo) para revisar */
SELECT 'RESULTADO' AS que, a.depen AS pantalla, COUNT(DISTINCT u.Usuario) AS usuarios, COUNT(DISTINCT g.Grupo) AS grupos
FROM dbo.Atrm a
LEFT JOIN dbo.AtrmUsuarios u ON u.Atributo = a.atributo
LEFT JOIN dbo.AtrmGrupos g ON g.Atributo = a.atributo
WHERE a.tipo = 'F' AND a.aplicacion = @Apli AND (a.depen LIKE N'CG[_]%')
GROUP BY a.depen ORDER BY a.depen;

COMMIT TRAN;   -- Cambiar a COMMIT TRAN cuando el resultado sea el esperado
