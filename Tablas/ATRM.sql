USE [dbcordatv46_unif]
GO

/****** Object:  Table [dbo].[Atrm]    Script Date: 03/27/2018 17:14:11 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[Atrm](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[atributo] [nvarchar](400) NULL,
	[alias] [nvarchar](400) NULL,
	[observaciones] [nvarchar](400) NULL,
	[tipo] [nvarchar](2) NULL,
	[depen] [nvarchar](200) NULL,
	[comprobar] [bit] NOT NULL,
	[aplicacion] [nvarchar](200) NULL,
	[nodo] [nvarchar](200) NULL,
	[valor] [nvarchar](200) NULL,
 CONSTRAINT [PK_Atrm] PRIMARY KEY NONCLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY]

GO

ALTER TABLE [dbo].[Atrm] ADD  CONSTRAINT [DF_Atrm_comprobar]  DEFAULT ((1)) FOR [comprobar]
GO

