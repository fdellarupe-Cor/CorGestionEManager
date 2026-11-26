/* =====================================================================
   SP 01 - sp_cg_Catalogo_Listar: catalogo tipo_avance

   ESTADO: SCRIPT PARA REVISAR. NO SE EJECUTO.
   Base destino: Dbcordatv48_UNIF485     Compatibilidad: nivel 100
   Base de partida: texto del procedimiento en P2 (Corv48UNIFV5_DB.sql del 15/10/2026 15:39).
   Solo se AGREGA lo indicado; el resto del procedimiento queda igual (vista vw_UserGrupos,
   paginacion con ROW_NUMBER, @Respon, estado_cliente, etc.).

   QUE AGREGA : la rama  @Catalogo = 'tipo_avance'  (id, descripcion de CG_TIPO_AVANCE, solo activos).
   PARA QUE   : la pantalla tipoavance.asp lista el catalogo con sp_cg_Catalogo_Listar @Catalogo='tipo_avance';
                sin esta rama el listado no devuelve datos y da error.
   DE DONDE   : P1 (sp_cg_Catalogo_Listar). Se trae solo tipo_avance.
   NO SE TRAE : la rama 'tipo_incidente' de P1 (usa CG_TIPO_INCIDENTE, que no existe en P2; es parte del
                SLA de P1, que esta en pausa).
   REQUISITO  : tabla CG_TIPO_AVANCE (ya existe en P2).

   REVERSA: volver a ejecutar la definicion original del procedimiento (DB/Corv48UNIFV5_DB.sql).
   ===================================================================== */

USE [Dbcordatv48_UNIF485]
GO
ALTER PROCEDURE [dbo].[sp_cg_Catalogo_Listar]
    @Catalogo VARCHAR(30),
    @Id       INT = NULL
AS
BEGIN
    SET NOCOUNT ON;

    IF @Catalogo = 'clientes'
        SELECT id, descripcion
        FROM CG_CLIENTES
        WHERE (@Id IS NULL OR id = @Id) AND ISNULL(ESTADO,1)=1
        ORDER BY descripcion;

    ELSE IF @Catalogo = 'aplicativos'
        SELECT id, descripcion
        FROM CG_APLICATIVOS
        WHERE (@Id IS NULL OR id = @Id) AND ISNULL(ESTADO,1)=1
        ORDER BY descripcion;

    ELSE IF @Catalogo = 'browsers'
        SELECT id, descripcion
        FROM CG_BROWSERS
        WHERE (@Id IS NULL OR id = @Id) AND ISNULL(ESTADO,1)=1
        ORDER BY descripcion;

    ELSE IF @Catalogo = 'estados'
        SELECT id, descripcion
        FROM CG_ESTADOS
        WHERE (@Id IS NULL OR id = @Id) AND ISNULL(ESTADO,1)=1
        ORDER BY id;

    ELSE IF @Catalogo = 'prioridades'
        SELECT id, descripcion
        FROM CG_PRIORIDADES
        WHERE (@Id IS NULL OR id = @Id) AND ISNULL(ESTADO,1)=1
        ORDER BY id;

    ELSE IF @Catalogo = 'via_pedido'
        SELECT id, descripcion
        FROM CG_VIA_PEDIDO
        WHERE (@Id IS NULL OR id = @Id) AND ISNULL(ESTADO,1)=1
        ORDER BY descripcion;

    ELSE IF @Catalogo = 'tipo_apl'
        SELECT id, descripcion
        FROM CG_TIPO_APL
        WHERE (@Id IS NULL OR id = @Id) --AND ISNULL(ESTADO,1)=1
        ORDER BY descripcion;

    ELSE IF @Catalogo = 'tipo_avance'
        SELECT id, descripcion
        FROM CG_TIPO_AVANCE
        WHERE (@Id IS NULL OR id = @Id) AND ISNULL(ESTADO,1)=1
        ORDER BY id;
END;
GO
