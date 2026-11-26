USE [dbcordatv46_unif]
GO

/****** Object:  Table [dbo].[COPARAEXPORTDOCTM]    Script Date: 03/27/2018 17:35:28 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[COPARAEXPORTDOCTM](
	[id] [int] IDENTITY(1,1) NOT NULL,
	[USUARIO] [nvarchar](100) NULL,
	[TEMASUBTEMA] [nvarchar](50) NULL,
	[SUBTEMA] [nvarchar](50) NULL,
	[TIPOESQUEMA] [nvarchar](50) NULL,
	[ESTADO] [nvarchar](50) NULL,
	[NOMBRE_ARCHIVO] [nvarchar](100) NULL,
	[FLG_TERMINADO] [nvarchar](100) NULL,
	[FLG_BORRADO] [nvarchar](5) NULL,
	[F_PEDIDO] [nvarchar](50) NULL,
	[F_GENERADO] [nvarchar](50) NULL,
	[F_BORRADO] [nvarchar](50) NULL,
	[USUARIO_BORRADO] [nvarchar](100) NULL,
	[OBSERVACIONES] [nvarchar](100) NULL,
	[PERIODO] [nvarchar](100) NULL,
	[prodserv] [int] NULL,
	[agrupam] [nvarchar](20) NULL,
	[segmento] [int] NULL,
	[PROGRAMA] [nvarchar](100) NULL,
	[FDESDE] [int] NULL,
	[FHASTA] [int] NULL,
	[ORDENLISTADO] [nvarchar](50) NULL,
 CONSTRAINT [PK_COPARAEXPORTDOCTM] PRIMARY KEY CLUSTERED 
(
	[id] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

