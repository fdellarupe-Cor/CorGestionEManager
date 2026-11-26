USE [dbcordatv46_unif]
GO

/****** Object:  Table [dbo].[AtrmUsuarios]    Script Date: 03/27/2018 17:20:00 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[AtrmUsuarios](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[Usuario] [nvarchar](200) NULL,
	[Atributo] [nvarchar](200) NULL,
	[Estado] [bit] NOT NULL,
	[Agregar] [bit] NOT NULL,
	[Modificar] [bit] NOT NULL,
	[Eliminar] [bit] NOT NULL,
	[Consultar] [bit] NOT NULL
) ON [PRIMARY]

GO

