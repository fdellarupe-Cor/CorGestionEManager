/* =====================================================================
   Cambios_Agenda_Listar_Filtros.sql  -  sp_cg_Agenda: filtros del listado de recordatorios

   ESTADO: SCRIPT PARA REVISAR. NO SE EJECUTO.
   Base destino: Dbcordatv48_UNIF485     Compatibilidad: nivel 100
   Base de partida: texto del procedimiento en P2 (Corv48UNIFV5_DB.sql del 15/10/2026 15:39).
   OJO: este ALTER reemplaza el procedimiento completo con ese texto. Si lo modificaste despues,
        volve a exportar el script antes de aplicar este.

   QUE CAMBIA (solo el modo 'listar'; los demas modos quedan igual):
     - Dos parametros nuevos, opcionales y al final:
         @Respon  login del usuario. Si viene, solo devuelve los registros donde el usuario es
                  - el responsable (ID_RESPON_SEG), o
                  - participante por copia (esta en COPIAS, lista de logins separados por coma), o
                  - corresponsable (CG_AGENDA_CORRESPONSABLE.ID_PERSONA_SEG, sin baja).
         @Texto   texto a buscar en ASUNTO o MENSAJE (contiene, sin distinguir mayusculas segun la
                  intercalacion de la base).
     - Con ambos en NULL (o vacios) se comporta igual que antes: devuelve todos los registros.
   PARA QUE   : agenda.asp (lista de Recordatorios) muestra por defecto los registros propios y permite
                buscar por texto. inicio.asp y los demas llamados a 'listar' no cambian.
   SI NO SE EJECUTA : agenda.asp sigue mostrando la lista completa y avisa que falta aplicar este script.
   REVERSA    : volver a ejecutar la definicion original del procedimiento (DB/Corv48UNIFV5_DB.sql).
   ===================================================================== */

USE [Dbcordatv48_UNIF485]
GO

ALTER PROCEDURE [dbo].[sp_cg_Agenda]
    @Modo      VARCHAR(20),
    @Id        INT          = NULL,
    @Asunto    NVARCHAR(1000) = NULL,
    @DiaAviso  INT          = NULL,
    @IdRespon  NVARCHAR(100)          = NULL,
    @Copias    NVARCHAR(1000) = NULL,
    @FecCarga  INT          = NULL,
    @Mensaje   NVARCHAR(MAX) = NULL,
    @Respon    NVARCHAR(100) = NULL,   -- [AGENDA] 'listar': solo registros del usuario (login): responsable, en COPIAS o corresponsable
    @Texto     NVARCHAR(200) = NULL    -- [AGENDA] 'listar': texto a buscar en ASUNTO o MENSAJE
AS
BEGIN
    SET NOCOUNT ON;

    SET @Respon = NULLIF(LTRIM(RTRIM(@Respon)), '');
    SET @Texto  = NULLIF(LTRIM(RTRIM(@Texto)), '');

    IF @Modo = 'listar'
    BEGIN
        SELECT ID, ASUNTO, DIA_AVISO, ID_RESPON_SEG AS ID_RESPON, ENVIADO, ESTADO, COPIAS, FECHADECARGA, MENSAJE, FEC_BAJA, USU_BAJA, ID_VIA, FEC_REU, HORA_INI, ID_ANT, CONCLUSION, FECHA_REUNION, HORA_INICIO, FECHA_AVISO, ID_RESPONSABLE_SEG AS ID_RESPONSABLE, ID_HERRAMIENTA, ID_AGENDA_PADRE, TIPO_ORIGEN, ESTADO_AGENDA, ENVIADO_BIT, FECHA_CREACION, USUA_BAJA, HORA_FIN
        FROM CG_AGENDA
        WHERE estado = 1
          -- [AGENDA] Solo los del usuario: responsable, en la lista de copias (COPIAS: logins separados por coma) o corresponsable
          AND (@Respon IS NULL
               OR ID_RESPON_SEG = @Respon
               OR CHARINDEX(',' + @Respon + ',', ',' + REPLACE(ISNULL(COPIAS, ''), ' ', '') + ',') > 0
               OR EXISTS (SELECT 1 FROM dbo.CG_AGENDA_CORRESPONSABLE c
                           WHERE c.ID_AGENDA = CG_AGENDA.ID AND c.ID_PERSONA_SEG = @Respon AND c.FEC_BAJA IS NULL))
          -- [AGENDA] Busqueda de texto en asunto o mensaje
          AND (@Texto IS NULL
               OR CHARINDEX(@Texto, ISNULL(ASUNTO, '')) > 0
               OR CHARINDEX(@Texto, ISNULL(MENSAJE, '')) > 0)
        ORDER BY dia_aviso DESC;
    END

    ELSE IF @Modo = 'insertar'
    BEGIN
        INSERT INTO CG_AGENDA (asunto, dia_aviso, id_respon_seg, enviado, estado, copias, fechadecarga, mensaje)
        VALUES (@Asunto, @DiaAviso, @IdRespon, 'NO', 1, @Copias, @FecCarga, @Mensaje);
    END

    ELSE IF @Modo = 'obtener'
    BEGIN
        SELECT enviado, asunto
        FROM CG_AGENDA
        WHERE id = @Id;
    END

    ELSE IF @Modo = 'obtener_new'
    BEGIN
        SELECT enviado, asunto, MENSAJE
        FROM CG_AGENDA
        WHERE id = @Id;
    END

    ELSE IF @Modo = 'actualizar_asunto'
    BEGIN
        UPDATE CG_AGENDA SET asunto = @Asunto WHERE id = @Id;
    END

    ELSE IF @Modo = 'obtener_copias'
    BEGIN
        SELECT copias
        FROM CG_AGENDA
        WHERE id = @Id;
    END

    ELSE IF @Modo = 'actualizar_copias'
    BEGIN
        UPDATE CG_AGENDA SET copias = @Copias WHERE id = @Id;
    END

    ELSE IF @Modo = 'recordatorios'
    BEGIN
        SELECT a.ID, a.ASUNTO, a.DIA_AVISO, a.ID_RESPON_SEG AS ID_RESPON, a.ENVIADO, a.ESTADO, a.COPIAS, a.FECHADECARGA, a.MENSAJE, a.FEC_BAJA, a.USU_BAJA, a.ID_VIA, a.FEC_REU, a.HORA_INI, a.ID_ANT, a.CONCLUSION, a.FECHA_REUNION, a.HORA_INICIO, a.FECHA_AVISO, a.ID_RESPONSABLE_SEG AS ID_RESPONSABLE, a.ID_HERRAMIENTA, a.ID_AGENDA_PADRE, a.TIPO_ORIGEN, a.ESTADO_AGENDA, a.ENVIADO_BIT, a.FECHA_CREACION, a.USUA_BAJA, a.HORA_FIN, b.mail
        FROM CG_AGENDA a
        INNER JOIN CG_COR_RESPONSABLES b ON (a.id_respon_seg = b.id_seg)
        WHERE a.dia_aviso = @DiaAviso AND a.enviado = 'NO' AND a.estado = 1;
    END

    ELSE IF @Modo = 'marcar_enviado'
    BEGIN
        UPDATE CG_AGENDA SET enviado = 'SI' WHERE id = @Id;
    END
END;
GO
