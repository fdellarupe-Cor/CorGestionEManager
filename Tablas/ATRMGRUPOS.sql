USE [dbcordatv46_unif]
GO

/****** Object:  Table [dbo].[AtrmGrupos]    Script Date: 03/27/2018 17:16:11 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[AtrmGrupos](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[Grupo] [nvarchar](200) NULL,
	[Atributo] [nvarchar](200) NULL,
	[Estado] [bit] NOT NULL,
	[Agregar] [bit] NOT NULL,
	[Modificar] [bit] NOT NULL,
	[Eliminar] [bit] NOT NULL,
	[Consultar] [bit] NOT NULL,
 CONSTRAINT [PK_AtrmGrupos] PRIMARY KEY NONCLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY]

GO

