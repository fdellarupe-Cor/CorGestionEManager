/* =====================================================================
   ESTRUCTURAS_P1_A_P2.sql
   Estructuras (claves, indices, defaults) que existen en el proyecto 1
   (Corgestionv5.v4 / dbcorgestionv5) y NO existen en el proyecto 2 (Dbcordatv48_UNIF485).

   ESTADO: SCRIPT PARA ELEGIR. NO SE EJECUTO.
   Base destino: Dbcordatv48_UNIF485
   Compatibilidad: nivel 100.
   Origen de las definiciones: DB/CorGestionDB.sql de Corgestionv5.v4 (export del 15/10/2026 15:39).

   COMO USARLO
     - Solo ESTRUCTURAS. No incluye procedimientos, vistas ni datos.
     - El bloque 0 (DIAGNOSTICO) es de solo lectura: se puede ejecutar siempre. Muestra que
       falta, si hay datos que impedirian crear una clave y cuantos registros afecta.
     - Cada bloque numerado (1 a 14) arranca con  DECLARE @EJECUTAR BIT = 0.
       Con 0 NO HACE NADA. Para elegirlo, cambiar a 1 en ESE bloque.
       Se puede ejecutar el archivo completo: solo actuan los bloques en 1.
     - Cada bloque se protege solo: si el objeto ya existe, o si hay datos que impiden crearlo,
       informa con PRINT y no hace nada.
     - ORDEN: las claves primarias (bloques 1 a 6) van ANTES que las claves foraneas
       (bloques 7 y 8), porque estas apuntan a CG_INCIDENTES y CG_PROYECTOS.

   RESUMEN
     Bloques 1-6  Claves primarias que faltan en 6 tablas CG_ de P2.
     Bloques 7-8  Claves foraneas AGENDA_SUBTEMA -> INCIDENTES / PROYECTOS.
     Bloques 9-12 Cuatro indices de la agenda.
     Bloques 13-14 Valores por defecto de CG_RELAOBJ.
     Nota final   CG_TIPO_INCIDENTE / ID_TIPO (no definidos en el script de P1).

   QUE NO SE TRASLADA (decision tomada en la comparacion)
     - USU_BAJA: P1 varchar(20), P2 varchar(100). Se mantiene P2 (mas ancho).
     - Tablas legadas de P1 sin uso: cli_con, comentarios, modulos, modulos_det, perfiles, TEST, test2.
     - Nombre de la PK de CG_RELAOBJ (solo cambia el nombre, no la definicion).
   ===================================================================== */

USE [Dbcordatv48_UNIF485]
GO

/* =====================================================================
   0) DIAGNOSTICO (solo lectura; se puede ejecutar siempre)
   ===================================================================== */
SELECT 'PK'  AS TIPO, 'CG_AVANCES.PK_RELA_INC'   AS ELEMENTO,
       CASE WHEN EXISTS (SELECT 1 FROM sys.key_constraints WHERE type='PK' AND parent_object_id = OBJECT_ID(N'dbo.CG_AVANCES'))   THEN 'YA EXISTE' ELSE 'FALTA' END AS ESTADO,
       (SELECT COUNT(*) FROM (SELECT ID FROM dbo.CG_AVANCES   GROUP BY ID HAVING COUNT(*) > 1) d) AS PROBLEMAS_DATOS
UNION ALL
SELECT 'PK', 'CG_ESTADOS.PK_ESTADOS',
       CASE WHEN EXISTS (SELECT 1 FROM sys.key_constraints WHERE type='PK' AND parent_object_id = OBJECT_ID(N'dbo.CG_ESTADOS'))    THEN 'YA EXISTE' ELSE 'FALTA' END,
       (SELECT COUNT(*) FROM (SELECT ID FROM dbo.CG_ESTADOS   GROUP BY ID HAVING COUNT(*) > 1) d)
UNION ALL
SELECT 'PK', 'CG_INCIDENTES.PK_INCIDENTES',
       CASE WHEN EXISTS (SELECT 1 FROM sys.key_constraints WHERE type='PK' AND parent_object_id = OBJECT_ID(N'dbo.CG_INCIDENTES')) THEN 'YA EXISTE' ELSE 'FALTA' END,
       (SELECT COUNT(*) FROM (SELECT ID FROM dbo.CG_INCIDENTES GROUP BY ID HAVING COUNT(*) > 1) d)
UNION ALL
SELECT 'PK', 'CG_LICENCIAS.PK_LICENCIAS',
       CASE WHEN EXISTS (SELECT 1 FROM sys.key_constraints WHERE type='PK' AND parent_object_id = OBJECT_ID(N'dbo.CG_LICENCIAS'))  THEN 'YA EXISTE' ELSE 'FALTA' END,
       (SELECT COUNT(*) FROM (SELECT ID FROM dbo.CG_LICENCIAS GROUP BY ID HAVING COUNT(*) > 1) d)
UNION ALL
SELECT 'PK', 'CG_PROYECTOS.PK_PROYECTOS',
       CASE WHEN EXISTS (SELECT 1 FROM sys.key_constraints WHERE type='PK' AND parent_object_id = OBJECT_ID(N'dbo.CG_PROYECTOS'))  THEN 'YA EXISTE' ELSE 'FALTA' END,
       (SELECT COUNT(*) FROM (SELECT ID FROM dbo.CG_PROYECTOS GROUP BY ID HAVING COUNT(*) > 1) d)
UNION ALL
SELECT 'PK', 'CG_TIPO_APL.PK_tipo_apl',
       CASE WHEN EXISTS (SELECT 1 FROM sys.key_constraints WHERE type='PK' AND parent_object_id = OBJECT_ID(N'dbo.CG_TIPO_APL'))   THEN 'YA EXISTE' ELSE 'FALTA' END,
       (SELECT COUNT(*) FROM (SELECT ID FROM dbo.CG_TIPO_APL  GROUP BY ID HAVING COUNT(*) > 1) d)
UNION ALL
SELECT 'FK', 'FK_AGENDA_SUBTEMA_INCIDENTE',
       CASE WHEN EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_AGENDA_SUBTEMA_INCIDENTE') THEN 'YA EXISTE' ELSE 'FALTA' END,
       (SELECT COUNT(*) FROM dbo.CG_AGENDA_SUBTEMA s
         WHERE s.ID_INCIDENTE IS NOT NULL AND NOT EXISTS (SELECT 1 FROM dbo.CG_INCIDENTES i WHERE i.ID = s.ID_INCIDENTE))
UNION ALL
SELECT 'FK', 'FK_AGENDA_SUBTEMA_PROYECTO',
       CASE WHEN EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_AGENDA_SUBTEMA_PROYECTO') THEN 'YA EXISTE' ELSE 'FALTA' END,
       (SELECT COUNT(*) FROM dbo.CG_AGENDA_SUBTEMA s
         WHERE s.ID_PROYECTO IS NOT NULL AND NOT EXISTS (SELECT 1 FROM dbo.CG_PROYECTOS p WHERE p.ID = s.ID_PROYECTO))
UNION ALL
SELECT 'INDICE', 'IX_AGENDA_FECHA_ESTADO',
       CASE WHEN EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_AGENDA_FECHA_ESTADO' AND object_id = OBJECT_ID(N'dbo.CG_AGENDA')) THEN 'YA EXISTE' ELSE 'FALTA' END, 0
UNION ALL
SELECT 'INDICE', 'IX_AGENDA_PADRE',
       CASE WHEN EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_AGENDA_PADRE' AND object_id = OBJECT_ID(N'dbo.CG_AGENDA')) THEN 'YA EXISTE' ELSE 'FALTA' END, 0
UNION ALL
SELECT 'INDICE', 'IX_AGENDA_CORRESP_ACTIVOS',
       CASE WHEN EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_AGENDA_CORRESP_ACTIVOS' AND object_id = OBJECT_ID(N'dbo.CG_AGENDA_CORRESPONSABLE')) THEN 'YA EXISTE' ELSE 'FALTA' END, 0
UNION ALL
SELECT 'INDICE', 'IX_AGENDA_SUBTEMA_AGENDA_ORDEN',
       CASE WHEN EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_AGENDA_SUBTEMA_AGENDA_ORDEN' AND object_id = OBJECT_ID(N'dbo.CG_AGENDA_SUBTEMA')) THEN 'YA EXISTE' ELSE 'FALTA' END, 0
UNION ALL
SELECT 'DEFAULT', 'CG_RELAOBJ.CODOBJMASTER = 0',
       CASE WHEN EXISTS (SELECT 1 FROM sys.default_constraints d JOIN sys.columns c ON c.object_id = d.parent_object_id AND c.column_id = d.parent_column_id
                         WHERE d.parent_object_id = OBJECT_ID(N'dbo.CG_RELAOBJ') AND c.name = N'CODOBJMASTER') THEN 'YA EXISTE' ELSE 'FALTA' END, 0
UNION ALL
SELECT 'DEFAULT', 'CG_RELAOBJ.CODOJASOC = 0',
       CASE WHEN EXISTS (SELECT 1 FROM sys.default_constraints d JOIN sys.columns c ON c.object_id = d.parent_object_id AND c.column_id = d.parent_column_id
                         WHERE d.parent_object_id = OBJECT_ID(N'dbo.CG_RELAOBJ') AND c.name = N'CODOJASOC') THEN 'YA EXISTE' ELSE 'FALTA' END, 0;
GO

/* =====================================================================
   1) PK_RELA_INC  -  CG_AVANCES (ID)
   QUE ES     : clave primaria clustered sobre CG_AVANCES.ID.
   PARA QUE   : CG_AVANCES (historial de avances) no tiene clave primaria en P2. Sin ella no hay
                garantia de unicidad de ID, y no se pueden crear claves foraneas hacia la tabla.
   RIESGO     : falla si hay IDs duplicados (el bloque 0 lo cuenta). Reorganiza la tabla
                (clustered): conviene hacerlo fuera del horario de uso.
   REVERSA    : ALTER TABLE dbo.CG_AVANCES DROP CONSTRAINT PK_RELA_INC;
   ===================================================================== */
DECLARE @EJECUTAR BIT = 0;   -- <== cambiar a 1 para ejecutar ESTE bloque
IF @EJECUTAR = 1
BEGIN
    IF EXISTS (SELECT 1 FROM sys.key_constraints WHERE type='PK' AND parent_object_id = OBJECT_ID(N'dbo.CG_AVANCES'))
        PRINT '1) CG_AVANCES ya tiene clave primaria: no se hace nada.';
    ELSE IF EXISTS (SELECT ID FROM dbo.CG_AVANCES GROUP BY ID HAVING COUNT(*) > 1)
        PRINT '1) CG_AVANCES tiene IDs duplicados: NO se crea la clave primaria.';
    ELSE
    BEGIN
        ALTER TABLE [dbo].[CG_AVANCES] ADD CONSTRAINT [PK_RELA_INC] PRIMARY KEY CLUSTERED ([ID] ASC)
            WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY];
        PRINT '1) PK_RELA_INC creada en CG_AVANCES.';
    END
END
GO

/* =====================================================================
   2) PK_ESTADOS  -  CG_ESTADOS (ID)
   QUE ES     : clave primaria clustered sobre CG_ESTADOS.ID.
   PARA QUE   : CG_ESTADOS es el catalogo de estados (OK, PENDIENTE, BAJA...) que usan incidentes y
                proyectos; hoy no tiene clave primaria en P2.
   RIESGO     : falla si hay IDs duplicados. Tabla chica: impacto bajo.
   REVERSA    : ALTER TABLE dbo.CG_ESTADOS DROP CONSTRAINT PK_ESTADOS;
   ===================================================================== */
DECLARE @EJECUTAR BIT = 0;   -- <== cambiar a 1 para ejecutar ESTE bloque
IF @EJECUTAR = 1
BEGIN
    IF EXISTS (SELECT 1 FROM sys.key_constraints WHERE type='PK' AND parent_object_id = OBJECT_ID(N'dbo.CG_ESTADOS'))
        PRINT '2) CG_ESTADOS ya tiene clave primaria: no se hace nada.';
    ELSE IF EXISTS (SELECT ID FROM dbo.CG_ESTADOS GROUP BY ID HAVING COUNT(*) > 1)
        PRINT '2) CG_ESTADOS tiene IDs duplicados: NO se crea la clave primaria.';
    ELSE
    BEGIN
        ALTER TABLE [dbo].[CG_ESTADOS] ADD CONSTRAINT [PK_ESTADOS] PRIMARY KEY CLUSTERED ([ID] ASC)
            WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY];
        PRINT '2) PK_ESTADOS creada en CG_ESTADOS.';
    END
END
GO

/* =====================================================================
   3) PK_INCIDENTES  -  CG_INCIDENTES (ID)
   QUE ES     : clave primaria clustered sobre CG_INCIDENTES.ID.
   PARA QUE   : la tabla principal de incidentes no tiene clave primaria en P2. Es requisito para la
                clave foranea del bloque 7 y garantiza que cada incidente tenga un ID unico.
   RIESGO     : falla si hay IDs duplicados. Reorganiza la tabla (clustered): conviene hacerlo fuera
                del horario de uso.
   REVERSA    : primero quitar la FK del bloque 7 (si se creo) y luego
                ALTER TABLE dbo.CG_INCIDENTES DROP CONSTRAINT PK_INCIDENTES;
   ===================================================================== */
DECLARE @EJECUTAR BIT = 0;   -- <== cambiar a 1 para ejecutar ESTE bloque
IF @EJECUTAR = 1
BEGIN
    IF EXISTS (SELECT 1 FROM sys.key_constraints WHERE type='PK' AND parent_object_id = OBJECT_ID(N'dbo.CG_INCIDENTES'))
        PRINT '3) CG_INCIDENTES ya tiene clave primaria: no se hace nada.';
    ELSE IF EXISTS (SELECT ID FROM dbo.CG_INCIDENTES GROUP BY ID HAVING COUNT(*) > 1)
        PRINT '3) CG_INCIDENTES tiene IDs duplicados: NO se crea la clave primaria.';
    ELSE
    BEGIN
        ALTER TABLE [dbo].[CG_INCIDENTES] ADD CONSTRAINT [PK_INCIDENTES] PRIMARY KEY CLUSTERED ([ID] ASC)
            WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY];
        PRINT '3) PK_INCIDENTES creada en CG_INCIDENTES.';
    END
END
GO

/* =====================================================================
   4) PK_LICENCIAS  -  CG_LICENCIAS (ID)
   QUE ES     : clave primaria clustered sobre CG_LICENCIAS.ID.
   PARA QUE   : las licencias (ausencias) que usa la agenda no tienen clave primaria en P2.
   RIESGO     : falla si hay IDs duplicados. Tabla chica: impacto bajo.
   REVERSA    : ALTER TABLE dbo.CG_LICENCIAS DROP CONSTRAINT PK_LICENCIAS;
   ===================================================================== */
DECLARE @EJECUTAR BIT = 0;   -- <== cambiar a 1 para ejecutar ESTE bloque
IF @EJECUTAR = 1
BEGIN
    IF EXISTS (SELECT 1 FROM sys.key_constraints WHERE type='PK' AND parent_object_id = OBJECT_ID(N'dbo.CG_LICENCIAS'))
        PRINT '4) CG_LICENCIAS ya tiene clave primaria: no se hace nada.';
    ELSE IF EXISTS (SELECT ID FROM dbo.CG_LICENCIAS GROUP BY ID HAVING COUNT(*) > 1)
        PRINT '4) CG_LICENCIAS tiene IDs duplicados: NO se crea la clave primaria.';
    ELSE
    BEGIN
        ALTER TABLE [dbo].[CG_LICENCIAS] ADD CONSTRAINT [PK_LICENCIAS] PRIMARY KEY CLUSTERED ([ID] ASC)
            WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY];
        PRINT '4) PK_LICENCIAS creada en CG_LICENCIAS.';
    END
END
GO

/* =====================================================================
   5) PK_PROYECTOS  -  CG_PROYECTOS (ID)
   QUE ES     : clave primaria clustered sobre CG_PROYECTOS.ID.
   PARA QUE   : la tabla principal de proyectos no tiene clave primaria en P2. Es requisito para la
                clave foranea del bloque 8.
   RIESGO     : falla si hay IDs duplicados. Reorganiza la tabla (clustered).
   REVERSA    : primero quitar la FK del bloque 8 (si se creo) y luego
                ALTER TABLE dbo.CG_PROYECTOS DROP CONSTRAINT PK_PROYECTOS;
   ===================================================================== */
DECLARE @EJECUTAR BIT = 0;   -- <== cambiar a 1 para ejecutar ESTE bloque
IF @EJECUTAR = 1
BEGIN
    IF EXISTS (SELECT 1 FROM sys.key_constraints WHERE type='PK' AND parent_object_id = OBJECT_ID(N'dbo.CG_PROYECTOS'))
        PRINT '5) CG_PROYECTOS ya tiene clave primaria: no se hace nada.';
    ELSE IF EXISTS (SELECT ID FROM dbo.CG_PROYECTOS GROUP BY ID HAVING COUNT(*) > 1)
        PRINT '5) CG_PROYECTOS tiene IDs duplicados: NO se crea la clave primaria.';
    ELSE
    BEGIN
        ALTER TABLE [dbo].[CG_PROYECTOS] ADD CONSTRAINT [PK_PROYECTOS] PRIMARY KEY CLUSTERED ([ID] ASC)
            WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY];
        PRINT '5) PK_PROYECTOS creada en CG_PROYECTOS.';
    END
END
GO

/* =====================================================================
   6) PK_tipo_apl  -  CG_TIPO_APL (ID)
   QUE ES     : clave primaria clustered sobre CG_TIPO_APL.ID.
   PARA QUE   : catalogo de tipos de APL; hoy sin clave primaria en P2.
   RIESGO     : falla si hay IDs duplicados. Tabla chica: impacto bajo.
   REVERSA    : ALTER TABLE dbo.CG_TIPO_APL DROP CONSTRAINT PK_tipo_apl;
   ===================================================================== */
DECLARE @EJECUTAR BIT = 0;   -- <== cambiar a 1 para ejecutar ESTE bloque
IF @EJECUTAR = 1
BEGIN
    IF EXISTS (SELECT 1 FROM sys.key_constraints WHERE type='PK' AND parent_object_id = OBJECT_ID(N'dbo.CG_TIPO_APL'))
        PRINT '6) CG_TIPO_APL ya tiene clave primaria: no se hace nada.';
    ELSE IF EXISTS (SELECT ID FROM dbo.CG_TIPO_APL GROUP BY ID HAVING COUNT(*) > 1)
        PRINT '6) CG_TIPO_APL tiene IDs duplicados: NO se crea la clave primaria.';
    ELSE
    BEGIN
        ALTER TABLE [dbo].[CG_TIPO_APL] ADD CONSTRAINT [PK_tipo_apl] PRIMARY KEY CLUSTERED ([ID] ASC)
            WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY];
        PRINT '6) PK_tipo_apl creada en CG_TIPO_APL.';
    END
END
GO

/* =====================================================================
   7) FK_AGENDA_SUBTEMA_INCIDENTE  -  CG_AGENDA_SUBTEMA(ID_INCIDENTE) -> CG_INCIDENTES(ID)
   QUE ES     : clave foranea.
   PARA QUE   : impide que un subtema de agenda apunte a un incidente que no existe.
   REQUISITO  : bloque 3 (PK de CG_INCIDENTES) ejecutado o ya existente.
   RIESGO     : falla si hay subtemas con ID_INCIDENTE inexistente (el bloque 0 los cuenta). Una vez
                creada, no se podra BORRAR FISICAMENTE un incidente que tenga subtemas; el borrado
                logico (ESTADO / FEC_BAJA) que usa Corgestion no se ve afectado.
   REVERSA    : ALTER TABLE dbo.CG_AGENDA_SUBTEMA DROP CONSTRAINT FK_AGENDA_SUBTEMA_INCIDENTE;
   ===================================================================== */
DECLARE @EJECUTAR BIT = 0;   -- <== cambiar a 1 para ejecutar ESTE bloque
IF @EJECUTAR = 1
BEGIN
    IF EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_AGENDA_SUBTEMA_INCIDENTE')
        PRINT '7) FK_AGENDA_SUBTEMA_INCIDENTE ya existe: no se hace nada.';
    ELSE IF NOT EXISTS (SELECT 1 FROM sys.key_constraints WHERE type='PK' AND parent_object_id = OBJECT_ID(N'dbo.CG_INCIDENTES'))
        PRINT '7) Falta la clave primaria de CG_INCIDENTES (bloque 3): NO se crea la clave foranea.';
    ELSE IF EXISTS (SELECT 1 FROM dbo.CG_AGENDA_SUBTEMA s WHERE s.ID_INCIDENTE IS NOT NULL
                    AND NOT EXISTS (SELECT 1 FROM dbo.CG_INCIDENTES i WHERE i.ID = s.ID_INCIDENTE))
        PRINT '7) Hay subtemas con ID_INCIDENTE inexistente: NO se crea la clave foranea (ver bloque 0).';
    ELSE
    BEGIN
        ALTER TABLE [dbo].[CG_AGENDA_SUBTEMA] WITH CHECK ADD CONSTRAINT [FK_AGENDA_SUBTEMA_INCIDENTE]
            FOREIGN KEY([ID_INCIDENTE]) REFERENCES [dbo].[CG_INCIDENTES] ([ID]);
        ALTER TABLE [dbo].[CG_AGENDA_SUBTEMA] CHECK CONSTRAINT [FK_AGENDA_SUBTEMA_INCIDENTE];
        PRINT '7) FK_AGENDA_SUBTEMA_INCIDENTE creada.';
    END
END
GO

/* =====================================================================
   8) FK_AGENDA_SUBTEMA_PROYECTO  -  CG_AGENDA_SUBTEMA(ID_PROYECTO) -> CG_PROYECTOS(ID)
   QUE ES     : clave foranea.
   PARA QUE   : impide que un subtema de agenda apunte a un proyecto que no existe.
   REQUISITO  : bloque 5 (PK de CG_PROYECTOS) ejecutado o ya existente.
   RIESGO     : igual que el bloque 7, con proyectos.
   REVERSA    : ALTER TABLE dbo.CG_AGENDA_SUBTEMA DROP CONSTRAINT FK_AGENDA_SUBTEMA_PROYECTO;
   ===================================================================== */
DECLARE @EJECUTAR BIT = 0;   -- <== cambiar a 1 para ejecutar ESTE bloque
IF @EJECUTAR = 1
BEGIN
    IF EXISTS (SELECT 1 FROM sys.foreign_keys WHERE name = N'FK_AGENDA_SUBTEMA_PROYECTO')
        PRINT '8) FK_AGENDA_SUBTEMA_PROYECTO ya existe: no se hace nada.';
    ELSE IF NOT EXISTS (SELECT 1 FROM sys.key_constraints WHERE type='PK' AND parent_object_id = OBJECT_ID(N'dbo.CG_PROYECTOS'))
        PRINT '8) Falta la clave primaria de CG_PROYECTOS (bloque 5): NO se crea la clave foranea.';
    ELSE IF EXISTS (SELECT 1 FROM dbo.CG_AGENDA_SUBTEMA s WHERE s.ID_PROYECTO IS NOT NULL
                    AND NOT EXISTS (SELECT 1 FROM dbo.CG_PROYECTOS p WHERE p.ID = s.ID_PROYECTO))
        PRINT '8) Hay subtemas con ID_PROYECTO inexistente: NO se crea la clave foranea (ver bloque 0).';
    ELSE
    BEGIN
        ALTER TABLE [dbo].[CG_AGENDA_SUBTEMA] WITH CHECK ADD CONSTRAINT [FK_AGENDA_SUBTEMA_PROYECTO]
            FOREIGN KEY([ID_PROYECTO]) REFERENCES [dbo].[CG_PROYECTOS] ([ID]);
        ALTER TABLE [dbo].[CG_AGENDA_SUBTEMA] CHECK CONSTRAINT [FK_AGENDA_SUBTEMA_PROYECTO];
        PRINT '8) FK_AGENDA_SUBTEMA_PROYECTO creada.';
    END
END
GO

/* =====================================================================
   9) IX_AGENDA_FECHA_ESTADO  -  CG_AGENDA (FECHA_REUNION, ESTADO_AGENDA, HORA_INICIO)
   QUE ES     : indice no clustered.
   PARA QUE   : acelera la vista semanal de la agenda y el chequeo de cruce de horarios, que filtran
                por fecha de reunion, estado y hora.
   RIESGO     : bajo. Ocupa algo de espacio y agrega un costo minimo a cada alta/modificacion de agenda.
   REVERSA    : DROP INDEX IX_AGENDA_FECHA_ESTADO ON dbo.CG_AGENDA;
   ===================================================================== */
DECLARE @EJECUTAR BIT = 0;   -- <== cambiar a 1 para ejecutar ESTE bloque
IF @EJECUTAR = 1
BEGIN
    IF EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_AGENDA_FECHA_ESTADO' AND object_id = OBJECT_ID(N'dbo.CG_AGENDA'))
        PRINT '9) IX_AGENDA_FECHA_ESTADO ya existe: no se hace nada.';
    ELSE
    BEGIN
        CREATE NONCLUSTERED INDEX [IX_AGENDA_FECHA_ESTADO] ON [dbo].[CG_AGENDA] ([FECHA_REUNION] ASC, [ESTADO_AGENDA] ASC, [HORA_INICIO] ASC)
            WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY];
        PRINT '9) IX_AGENDA_FECHA_ESTADO creado.';
    END
END
GO

/* =====================================================================
   10) IX_AGENDA_PADRE  -  CG_AGENDA (ID_AGENDA_PADRE)
   QUE ES     : indice no clustered.
   PARA QUE   : acelera la busqueda de agendas hijas (seguimientos y postergaciones) de una agenda.
   RIESGO     : bajo.
   REVERSA    : DROP INDEX IX_AGENDA_PADRE ON dbo.CG_AGENDA;
   ===================================================================== */
DECLARE @EJECUTAR BIT = 0;   -- <== cambiar a 1 para ejecutar ESTE bloque
IF @EJECUTAR = 1
BEGIN
    IF EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_AGENDA_PADRE' AND object_id = OBJECT_ID(N'dbo.CG_AGENDA'))
        PRINT '10) IX_AGENDA_PADRE ya existe: no se hace nada.';
    ELSE
    BEGIN
        CREATE NONCLUSTERED INDEX [IX_AGENDA_PADRE] ON [dbo].[CG_AGENDA] ([ID_AGENDA_PADRE] ASC)
            WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY];
        PRINT '10) IX_AGENDA_PADRE creado.';
    END
END
GO

/* =====================================================================
   11) IX_AGENDA_CORRESP_ACTIVOS  -  CG_AGENDA_CORRESPONSABLE (ID_AGENDA, ID_PERSONA, FEC_BAJA)
   QUE ES     : indice no clustered, tal como esta en P1.
   PARA QUE   : acelera la lectura de los participantes activos de una agenda.
   ATENCION   : usa la columna ID_PERSONA (numerica, heredada). En P2 los procedimientos buscan por
                ID_PERSONA_SEG (usuario de red). Para P2 puede ser mas util indexar ID_PERSONA_SEG:
                ver la variante comentada debajo. [DECIDIR cual de las dos]
   RIESGO     : bajo.
   REVERSA    : DROP INDEX IX_AGENDA_CORRESP_ACTIVOS ON dbo.CG_AGENDA_CORRESPONSABLE;
   ===================================================================== */
DECLARE @EJECUTAR BIT = 0;   -- <== cambiar a 1 para ejecutar ESTE bloque
IF @EJECUTAR = 1
BEGIN
    IF EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_AGENDA_CORRESP_ACTIVOS' AND object_id = OBJECT_ID(N'dbo.CG_AGENDA_CORRESPONSABLE'))
        PRINT '11) IX_AGENDA_CORRESP_ACTIVOS ya existe: no se hace nada.';
    ELSE
    BEGIN
        CREATE NONCLUSTERED INDEX [IX_AGENDA_CORRESP_ACTIVOS] ON [dbo].[CG_AGENDA_CORRESPONSABLE] ([ID_AGENDA] ASC, [ID_PERSONA] ASC, [FEC_BAJA] ASC)
            WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY];
        PRINT '11) IX_AGENDA_CORRESP_ACTIVOS creado.';
    END
    -- VARIANTE (no esta en P1) para P2: indexar por usuario de red en lugar de ID_PERSONA.
    -- CREATE NONCLUSTERED INDEX [IX_AGENDA_CORRESP_ACTIVOS_SEG] ON [dbo].[CG_AGENDA_CORRESPONSABLE] ([ID_AGENDA] ASC, [ID_PERSONA_SEG] ASC, [FEC_BAJA] ASC);
END
GO

/* =====================================================================
   12) IX_AGENDA_SUBTEMA_AGENDA_ORDEN  -  CG_AGENDA_SUBTEMA (ID_AGENDA, ORDEN, FEC_BAJA)
   QUE ES     : indice no clustered.
   PARA QUE   : acelera la lectura ordenada de los subtemas activos de una agenda.
   RIESGO     : bajo.
   REVERSA    : DROP INDEX IX_AGENDA_SUBTEMA_AGENDA_ORDEN ON dbo.CG_AGENDA_SUBTEMA;
   ===================================================================== */
DECLARE @EJECUTAR BIT = 0;   -- <== cambiar a 1 para ejecutar ESTE bloque
IF @EJECUTAR = 1
BEGIN
    IF EXISTS (SELECT 1 FROM sys.indexes WHERE name = N'IX_AGENDA_SUBTEMA_AGENDA_ORDEN' AND object_id = OBJECT_ID(N'dbo.CG_AGENDA_SUBTEMA'))
        PRINT '12) IX_AGENDA_SUBTEMA_AGENDA_ORDEN ya existe: no se hace nada.';
    ELSE
    BEGIN
        CREATE NONCLUSTERED INDEX [IX_AGENDA_SUBTEMA_AGENDA_ORDEN] ON [dbo].[CG_AGENDA_SUBTEMA] ([ID_AGENDA] ASC, [ORDEN] ASC, [FEC_BAJA] ASC)
            WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY];
        PRINT '12) IX_AGENDA_SUBTEMA_AGENDA_ORDEN creado.';
    END
END
GO

/* =====================================================================
   13) DEFAULT 0 en CG_RELAOBJ.CODOBJMASTER
   QUE ES     : valor por defecto ((0)).
   PARA QUE   : los inserts que no informan el objeto maestro de la relacion guardan 0 en lugar de NULL.
   RIESGO     : bajo. No modifica filas existentes (solo afecta altas futuras sin ese dato).
   REVERSA    : buscar el nombre en sys.default_constraints y hacer ALTER TABLE ... DROP CONSTRAINT.
   ===================================================================== */
DECLARE @EJECUTAR BIT = 0;   -- <== cambiar a 1 para ejecutar ESTE bloque
IF @EJECUTAR = 1
BEGIN
    IF EXISTS (SELECT 1 FROM sys.default_constraints d JOIN sys.columns c ON c.object_id = d.parent_object_id AND c.column_id = d.parent_column_id
               WHERE d.parent_object_id = OBJECT_ID(N'dbo.CG_RELAOBJ') AND c.name = N'CODOBJMASTER')
        PRINT '13) CG_RELAOBJ.CODOBJMASTER ya tiene un valor por defecto: no se hace nada.';
    ELSE
    BEGIN
        ALTER TABLE [dbo].[CG_RELAOBJ] ADD DEFAULT ((0)) FOR [CODOBJMASTER];
        PRINT '13) Default 0 creado en CG_RELAOBJ.CODOBJMASTER.';
    END
END
GO

/* =====================================================================
   14) DEFAULT 0 en CG_RELAOBJ.CODOJASOC
   QUE ES     : valor por defecto ((0)).
   PARA QUE   : igual que el bloque 13, para el objeto asociado.
   RIESGO     : bajo.
   REVERSA    : igual que el bloque 13.
   ===================================================================== */
DECLARE @EJECUTAR BIT = 0;   -- <== cambiar a 1 para ejecutar ESTE bloque
IF @EJECUTAR = 1
BEGIN
    IF EXISTS (SELECT 1 FROM sys.default_constraints d JOIN sys.columns c ON c.object_id = d.parent_object_id AND c.column_id = d.parent_column_id
               WHERE d.parent_object_id = OBJECT_ID(N'dbo.CG_RELAOBJ') AND c.name = N'CODOJASOC')
        PRINT '14) CG_RELAOBJ.CODOJASOC ya tiene un valor por defecto: no se hace nada.';
    ELSE
    BEGIN
        ALTER TABLE [dbo].[CG_RELAOBJ] ADD DEFAULT ((0)) FOR [CODOJASOC];
        PRINT '14) Default 0 creado en CG_RELAOBJ.CODOJASOC.';
    END
END
GO

/* =====================================================================
   NOTA FINAL (no ejecutable): TIPO DE INCIDENTE
   Los procedimientos de P1 (sp_cg_Catalogo_Listar 'tipo_incidente', sp_cg_Combo_Listar 'id_tipo',
   sp_cg_Incidentes_Insertar @id_tipo, sp_cg_Registro_ActualizarCampos) usan la tabla CG_TIPO_INCIDENTE
   y la columna CG_INCIDENTES.ID_TIPO, pero el script de P1 NO define ni la tabla ni la columna.
   No se copian aqui porque no se puede inventar su estructura. Para completarlo: confirmar en la base
   de P1 la definicion real de CG_TIPO_INCIDENTE (ID, DESCRIPCION, ESTADO, ...) y de CG_INCIDENTES.ID_TIPO,
   y agregarla como un bloque mas.
   ===================================================================== */
