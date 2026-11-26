/* =====================================================================
   Cambios_Agenda_Licencias.sql  -  vista semanal con vw_UserGrupos y licencias

   ESTADO: SCRIPT PARA REVISAR. NO SE EJECUTO.
   Base destino: Dbcordatv48_UNIF485     Compatibilidad: nivel 100
   Base de partida: texto de los procedimientos en P2 (Corv48UNIFV5_DB.sql del 15/10/2026 15:39).
   OJO: cada ALTER reemplaza el procedimiento completo con ese texto. Si los modificaste despues,
        volve a exportarlos antes de aplicar este script.

   1) sp_cg_Agenda_VistaSemanal
      - El responsable de cada reunion y los participantes salen de vw_UserGrupos (USRNAME), una sola
        fila por persona; ya no se usa CG_COR_RESPONSABLES.
      - NUEVO tercer resultado: licencias vigentes (ESTADO = 1, sin FEC_BAJA) que se cruzan con el rango
        pedido. Columnas: ID, ID_PERSONA (login), PERSONA (nombre; si no esta en la vista, el login),
        F_INICIO, F_FIN (si esta vacio o en 0 se toma F_INICIO), MOTIVO.
      - Los dos primeros resultados mantienen sus columnas.
   2) sp_cg_Licencias_Listar
      - El nombre sale de vw_UserGrupos. Antes hacia INNER JOIN con CG_COR_RESPONSABLES y, con esa tabla
        vacia, el listado de licencias.asp quedaba sin filas. Ahora lista todas las licencias vigentes
        (si la persona ya no esta en la vista se muestra el login).
   SI NO SE EJECUTA : agenda_semana.asp sigue funcionando, sin licencias y sin nombres de la vista.
   REVERSA    : volver a ejecutar la definicion original de ambos procedimientos (DB/Corv48UNIFV5_DB.sql).
   ===================================================================== */

USE [Dbcordatv48_UNIF485]
GO

ALTER PROCEDURE [dbo].[sp_cg_Agenda_VistaSemanal]
    @FechaDesde INT,
    @FechaHasta INT
AS
BEGIN
    SET NOCOUNT ON;

    SELECT
        A.ID,
        A.ASUNTO,
        A.FECHA_REUNION,
        A.HORA_INICIO,
        A.HORA_FIN,
        ISNULL(A.ESTADO_AGENDA, 'PENDIENTE') AS ESTADO_AGENDA,
        A.ID_RESPONSABLE_SEG AS ID_RESPONSABLE,
        RTRIM(ISNULL((SELECT MAX(R.USRNAME) FROM dbo.vw_UserGrupos R WHERE R.USRNET = A.ID_RESPONSABLE_SEG), '')) AS RESPONSABLE
    FROM dbo.CG_AGENDA A
    WHERE A.FEC_BAJA IS NULL
      AND A.FECHA_REUNION BETWEEN @FechaDesde AND @FechaHasta
    ORDER BY A.FECHA_REUNION, A.HORA_INICIO;

    SELECT
        C.ID_AGENDA,
        C.ID_PERSONA_SEG AS ID_PERSONA,
        RTRIM(ISNULL(
            NULLIF(LTRIM(RTRIM(C.NOMBRE_MANUAL)), ''),
            ISNULL((SELECT MAX(R.USRNAME) FROM dbo.vw_UserGrupos R WHERE R.USRNET = C.ID_PERSONA_SEG), '')
        )) AS NOMBRE
    FROM dbo.CG_AGENDA_CORRESPONSABLE C
    INNER JOIN dbo.CG_AGENDA A ON A.ID = C.ID_AGENDA AND A.FEC_BAJA IS NULL
    WHERE C.FEC_BAJA IS NULL
      AND A.FECHA_REUNION BETWEEN @FechaDesde AND @FechaHasta
    ORDER BY C.ID_AGENDA;

    -- [AGENDA] Tercer resultado: licencias vigentes que caen (total o parcialmente) en el rango.
    -- Si F_FIN esta vacio o en 0 se toma igual a F_INICIO. agenda_semana.asp las marca como un dia completo.
    SELECT
        L.ID,
        L.ID_RESPON_SEG AS ID_PERSONA,
        RTRIM(ISNULL((SELECT MAX(R.USRNAME) FROM dbo.vw_UserGrupos R WHERE R.USRNET = L.ID_RESPON_SEG), L.ID_RESPON_SEG)) AS PERSONA,
        L.F_INICIO,
        CASE WHEN ISNULL(L.F_FIN, 0) = 0 OR L.F_FIN < L.F_INICIO THEN L.F_INICIO ELSE L.F_FIN END AS F_FIN,
        L.MOTIVO
    FROM dbo.CG_LICENCIAS L
    WHERE L.ESTADO = 1
      AND L.FEC_BAJA IS NULL
      AND L.F_INICIO IS NOT NULL
      AND L.F_INICIO <= @FechaHasta
      AND (CASE WHEN ISNULL(L.F_FIN, 0) = 0 OR L.F_FIN < L.F_INICIO THEN L.F_INICIO ELSE L.F_FIN END) >= @FechaDesde
    ORDER BY L.F_INICIO, L.ID;
END;
GO

ALTER PROCEDURE [dbo].[sp_cg_Licencias_Listar]
AS
BEGIN
    SET NOCOUNT ON;

    SELECT a.id,
           ISNULL((SELECT MAX(r.USRNAME) FROM vw_UserGrupos r WHERE r.USRNET = a.id_respon_seg), a.id_respon_seg),
           ISNULL(a.f_inicio, 0),
           ISNULL(a.f_fin, 0), a.motivo
    FROM CG_LICENCIAS a
    WHERE A.ESTADO = 1
    ORDER BY a.id DESC;
END;
GO
