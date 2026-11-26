USE [dbcordatV46_UNIF]
GO

/****** Object:  Table [dbo].[CorConsulta]    Script Date: 03/27/2018 17:39:28 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[CorConsulta](
	[ID] [nvarchar](100) NOT NULL,
	[CodigoX] [nvarchar](100) NULL,
	[CodLicencia] [nvarchar](50) NULL,
	[CodEmpresa] [nvarchar](200) NULL,
	[objid] [int] IDENTITY(1,1) NOT NULL,
 CONSTRAINT [PK_CorConsulta] PRIMARY KEY CLUSTERED 
(
	[objid] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

