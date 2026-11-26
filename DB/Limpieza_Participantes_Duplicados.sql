/* =====================================================================
   Limpieza_Participantes_Duplicados.sql  -  da de baja los participantes repetidos de CG_AGENDA_CORRESPONSABLE

   ESTADO: SCRIPT PARA REVISAR. NO SE EJECUTO.
   Base destino: Dbcordatv48_UNIF485
   QUE HACE:
     1) Lista las reuniones con participantes activos repetidos (misma reunion + misma persona, o mismo contacto
        manual con el mismo nombre y mail) y cuantas filas sobran.
     2) Deja la fila MAS ANTIGUA de cada grupo y da de baja las demas (FEC_BAJA = GETDATE(), igual que
        corresp_eliminar: no se borra nada). Si alguna de las repetidas tenia ENVIAR_COPIA = 1, la que queda lo conserva.
     3) El UPDATE corre dentro de una transaccion que termina en ROLLBACK. Revisa los resultados y, si estan bien,
        cambia la ultima linea por COMMIT TRAN y vuelve a ejecutar.
   REVERSA (despues del COMMIT): UPDATE dbo.CG_AGENDA_CORRESPONSABLE SET FEC_BAJA = NULL WHERE <filas dadas de baja por la limpieza>.
   ===================================================================== */

USE [Dbcordatv48_UNIF485]
GO

-- 1) Diagnostico: grupos repetidos
SELECT c.ID_AGENDA, c.ID_PERSONA_SEG, c.NOMBRE_MANUAL, c.EMAIL_MANUAL,
       COUNT(*) AS FILAS, COUNT(*) - 1 AS SOBRAN
FROM dbo.CG_AGENDA_CORRESPONSABLE c
WHERE c.FEC_BAJA IS NULL
GROUP BY c.ID_AGENDA, c.ID_PERSONA_SEG, c.NOMBRE_MANUAL, c.EMAIL_MANUAL
HAVING COUNT(*) > 1
ORDER BY c.ID_AGENDA, c.ID_PERSONA_SEG;
GO

-- 2) Limpieza (con ROLLBACK por defecto)
BEGIN TRAN;

IF OBJECT_ID('tempdb..#dup') IS NOT NULL DROP TABLE #dup;

SELECT c.ID_AGENDA_CORRESP,
       ROW_NUMBER() OVER (PARTITION BY c.ID_AGENDA, ISNULL(c.ID_PERSONA_SEG, N''), ISNULL(c.NOMBRE_MANUAL, N''), ISNULL(c.EMAIL_MANUAL, N'')
                          ORDER BY c.ID_AGENDA_CORRESP) AS RN,
       MAX(CAST(c.ENVIAR_COPIA AS INT)) OVER (PARTITION BY c.ID_AGENDA, ISNULL(c.ID_PERSONA_SEG, N''), ISNULL(c.NOMBRE_MANUAL, N''), ISNULL(c.EMAIL_MANUAL, N'')) AS MAX_COPIA
INTO #dup
FROM dbo.CG_AGENDA_CORRESPONSABLE c
WHERE c.FEC_BAJA IS NULL;

-- la fila que queda conserva la copia si alguna repetida la tenia
UPDATE c SET ENVIAR_COPIA = 1
FROM dbo.CG_AGENDA_CORRESPONSABLE c
INNER JOIN #dup d ON d.ID_AGENDA_CORRESP = c.ID_AGENDA_CORRESP
WHERE d.RN = 1 AND d.MAX_COPIA = 1 AND c.ENVIAR_COPIA = 0;

-- baja de las repetidas
UPDATE c SET FEC_BAJA = GETDATE()
FROM dbo.CG_AGENDA_CORRESPONSABLE c
INNER JOIN #dup d ON d.ID_AGENDA_CORRESP = c.ID_AGENDA_CORRESP
WHERE d.RN > 1;

SELECT @@ROWCOUNT AS FILAS_DADAS_DE_BAJA;

DROP TABLE #dup;

ROLLBACK TRAN;   -- <<< cambiar por COMMIT TRAN despues de revisar el diagnostico
GO
