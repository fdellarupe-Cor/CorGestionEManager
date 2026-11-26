USE [dbcordatV46_UNIF]
GO

/****** Object:  Table [dbo].[CorAplic]    Script Date: 03/27/2018 17:37:42 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[CorAplic](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[CodApli] [nvarchar](100) NULL,
	[DescApli] [nvarchar](100) NULL,
	[Enable] [bit] NOT NULL,
	[VersionApli] [nvarchar](50) NULL,
	[ModoApli] [nvarchar](50) NULL,
	[IconoApli] [nvarchar](50) NULL,
	[DirApli] [nvarchar](50) NULL,
	[ImagenApli] [nvarchar](50) NULL,
	[Corpapli] [bit] NOT NULL,
	[TargApli] [nvarchar](50) NULL,
	[FormApli] [nvarchar](50) NULL,
	[TitApli] [nvarchar](50) NULL,
	[OrdenApli] [int] NULL,
	[TitAbrApli] [nvarchar](20) NULL,
 CONSTRAINT [aaaaaCorAplic_PK] PRIMARY KEY NONCLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY]

GO

ALTER TABLE [dbo].[CorAplic] ADD  CONSTRAINT [DF__Temporary__Enabl__08EA5793]  DEFAULT ((0)) FOR [Enable]
GO

ALTER TABLE [dbo].[CorAplic] ADD  CONSTRAINT [DF_CorAplic_Enable1]  DEFAULT ((0)) FOR [Corpapli]
GO

ALTER TABLE [dbo].[CorAplic] ADD  CONSTRAINT [DF_CorAplic_OrdenApli]  DEFAULT ((0)) FOR [OrdenApli]
GO

