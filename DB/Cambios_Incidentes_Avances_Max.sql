/* =====================================================================
   Cambios_Incidentes_Avances_Max.sql  -  CG_INCIDENTES.AVANCES a nvarchar(max)

   ESTADO: SCRIPT PARA REVISAR. NO SE EJECUTO.
   Base: Dbcordatv48_UNIF485 (desarrollo). En prod ya lo incluye Migracion_01_Estructura.sql.
   MOTIVO: en prod hay un incidente con un avance de 10552 caracteres; el original (INCIDENTES.AVANCES) era nvarchar(max) y
   CG_INCIDENTES.AVANCES quedo en nvarchar(4000). No hay indices ni vistas que dependan de la columna; los procedimientos solo la leen.
   Corre solo si la columna todavia no es max (repetible).
   REVERSA: ALTER TABLE dbo.CG_INCIDENTES ALTER COLUMN AVANCES NVARCHAR(4000) NULL;  (falla si ya hay textos mas largos).
   ===================================================================== */

USE [Dbcordatv48_UNIF485]
GO

IF EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID(N'dbo.CG_INCIDENTES') AND name = N'AVANCES' AND max_length <> -1)
    ALTER TABLE dbo.CG_INCIDENTES ALTER COLUMN AVANCES NVARCHAR(MAX) NULL;
GO
