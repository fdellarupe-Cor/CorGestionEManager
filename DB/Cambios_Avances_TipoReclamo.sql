/* =====================================================================
   Cambios_Avances_TipoReclamo.sql  -  tipo de avance "Reclamo" (valor por defecto de "Insistir")

   ESTADO: SCRIPT PARA REVISAR. NO SE EJECUTO.
   Base destino: Dbcordatv48_UNIF485 (desarrollo) y, cuando se decida, dbcordatv485Cor (prod).
   QUE HACE: crea el tipo "Reclamo" en CG_TIPO_AVANCE solo si todavia no existe un tipo con esa descripcion (sin distinguir
   mayusculas ni espacios). Repetible. El pop-up "Insistir" busca ese tipo por su descripcion y lo deja preseleccionado
   (editable); si no existe, el tipo queda sin seleccionar.
   REVERSA: DELETE FROM dbo.CG_TIPO_AVANCE WHERE DESCRIPCION = 'Reclamo' (si todavia no hay avances con ese tipo).
   ===================================================================== */

USE [Dbcordatv48_UNIF485]
GO

IF NOT EXISTS (SELECT 1 FROM dbo.CG_TIPO_AVANCE WHERE UPPER(LTRIM(RTRIM(DESCRIPCION))) = 'RECLAMO')
    INSERT INTO dbo.CG_TIPO_AVANCE (DESCRIPCION, ESTADO) VALUES (N'Reclamo', 1);
GO

SELECT ID, DESCRIPCION, ESTADO FROM dbo.CG_TIPO_AVANCE ORDER BY ID;
GO
