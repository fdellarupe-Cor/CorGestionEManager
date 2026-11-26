/* =====================================================================
   CAMBIOS_FILTRO_RESPONSABLE_ETAPA3.sql
   Listados de incidentes y proyectos con filtro por responsable
   (por defecto: registros del usuario de la sesion; opcion "todos")

   ESTADO: SCRIPT PARA REVISAR. NO SE EJECUTO. Ejecutar solo despues de revisarlo.
   Base:   Dbcordatv48_UNIF485
   Compatibilidad: nivel 100.

   Contenido
     1) ALTER sp_cg_Incidentes_Listar : agrega @Respon (al final, opcional).
     2) ALTER sp_cg_Proyectos_Listar  : agrega @Respon (al final, opcional).
     3) PARCHE MANUAL de 1 linea en sp_cg_ConsultarEntidad (modo 'proyecto' / 'lista').
        No se genera un ALTER completo porque ese procedimiento ya fue modificado
        varias veces y la version vigente es la que esta en la base.

   Regla del filtro
     @Respon = usuario de red. Devuelve los registros donde el usuario es
     responsable de gestion O responsable tecnico (respon_gestion_seg / respon_tec_seg).
     @Respon = NULL -> todos los registros.

   Los parametros nuevos tienen valor por defecto NULL: los llamados anteriores siguen
   funcionando sin cambios.
   ===================================================================== */

USE [Dbcordatv48_UNIF485]
GO

/* ---------------------------------------------------------------------
   1) sp_cg_Incidentes_Listar
      sp_cg_ConsultarEntidad ya filtra incidentes por @Respon (modo 'lista').
   --------------------------------------------------------------------- */
ALTER PROCEDURE [dbo].[sp_cg_Incidentes_Listar]
    @IdCliente    INT = 0,
    @IdAplicativo INT = 0,
    @Titulo       NVARCHAR(255) = NULL,
    @Orden        VARCHAR(20) = NULL,
    @Respon       NVARCHAR(100) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    EXEC dbo.sp_cg_ConsultarEntidad @Entidad = 'incidente', @Modo = 'lista',
        @IdCliente = @IdCliente, @IdAplicativo = @IdAplicativo, @Titulo = @Titulo, @Orden = @Orden,
        @Respon = @Respon;
END;
GO

/* ---------------------------------------------------------------------
   2) sp_cg_Proyectos_Listar
   --------------------------------------------------------------------- */
ALTER PROCEDURE [dbo].[sp_cg_Proyectos_Listar]
    @IdCliente       INT = 0,
    @IdAplicativo    INT = 0,
    @TextoBusqueda   NVARCHAR(100) = NULL,
    @EstadoFiltro    INT = 0,
    @FechaIngreso    INT = NULL,
    @Orden           NVARCHAR(50) = NULL,
    @RegistroInicial INT = 0,
    @CantidadMostrar INT = 12,
    @Respon          NVARCHAR(100) = NULL
AS
BEGIN
    SET NOCOUNT ON;
    EXEC dbo.sp_cg_ConsultarEntidad @Entidad = 'proyecto', @Modo = 'lista',
        @IdCliente = @IdCliente, @IdAplicativo = @IdAplicativo, @TextoBusqueda = @TextoBusqueda,
        @EstadoFiltro = @EstadoFiltro, @FechaIngreso = @FechaIngreso, @Orden = @Orden,
        @RegistroInicial = @RegistroInicial, @CantidadMostrar = @CantidadMostrar,
        @Respon = @Respon;
END;
GO

/* ---------------------------------------------------------------------
   3) PARCHE MANUAL EN sp_cg_ConsultarEntidad  (HACER ANTES DE USAR LOS LISTADOS)

   El parametro @Respon ya existe en sp_cg_ConsultarEntidad. Solo falta aplicarlo en el
   listado de proyectos. En el modo  @Entidad = 'proyecto' AND @Modo = 'lista',
   dentro del CTE  ProyectosFiltrados, en su WHERE, AGREGAR esta linea al final:

              AND (@Respon IS NULL
                   OR respon_gestion = @Respon
                   OR respon_tec     = @Respon)

   Quedaria asi:

        ProyectosFiltrados AS (
            SELECT *,
                   COUNT(*) OVER() AS TotalRegistros
            FROM ProyectosBase
            WHERE (@IdCliente = 0 OR id_cliente = @IdCliente)
              AND (@IdAplicativo = 0 OR aplicativo_id = @IdAplicativo)
              AND (@TextoBusqueda IS NULL OR titulo LIKE '%' + @TextoBusqueda + '%' OR descripcion LIKE '%' + @TextoBusqueda + '%')
              AND (@EstadoFiltro = 0 OR ... )
              AND (@FechaIngreso IS NULL OR fecha_ingreso = @FechaIngreso)
              AND (@Respon IS NULL OR respon_gestion = @Respon OR respon_tec = @Respon)   -- <== NUEVA
        ),

   Notas
     - respon_gestion y respon_tec son los alias de a.respon_gestion_seg y a.respon_tec_seg
       en el CTE ProyectosBase.
     - El filtro va ANTES de numerar y paginar (esta dentro de ProyectosFiltrados), por lo que
       TotalRegistros y la paginacion cuentan solo los registros filtrados.
     - Si el parche no se aplica, el filtro de proyectos no tiene efecto (se ven todos).
   --------------------------------------------------------------------- */

/* REVERSA: volver a ejecutar los ALTER originales de sp_cg_Incidentes_Listar y
   sp_cg_Proyectos_Listar (version en DB/Corv48UNIFV5_DB.sql, 2026-10-09) y quitar la linea
   agregada en sp_cg_ConsultarEntidad. */
