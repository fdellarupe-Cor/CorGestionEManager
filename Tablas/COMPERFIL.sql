USE [dbcordatv46_unif]
GO

/****** Object:  Table [dbo].[COMPERFIL]    Script Date: 03/27/2018 17:29:00 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[COMPERFIL](
	[tipoperfil] [nvarchar](10) NOT NULL,
	[descri] [nvarchar](100) NULL,
	[compprty] [bit] NOT NULL,
	[compestr] [bit] NOT NULL,
	[compdet] [bit] NOT NULL,
	[compimag] [bit] NOT NULL,
	[comprel] [bit] NOT NULL,
	[cmp06] [bit] NULL,
	[cmp07] [bit] NULL,
	[cmp08] [bit] NULL,
	[cmp09] [bit] NULL,
	[cmp10] [bit] NULL,
	[cmp11] [bit] NULL,
	[cmp12] [bit] NULL,
	[cmp13] [bit] NULL,
	[cmp14] [bit] NULL,
	[cmp15] [bit] NULL,
	[cmpd01] [nvarchar](50) NULL,
	[cmpd02] [nvarchar](50) NULL,
	[cmpd03] [nvarchar](50) NULL,
	[cmpd04] [nvarchar](50) NULL,
	[cmpd05] [nvarchar](50) NULL,
	[cmpd06] [nvarchar](50) NULL,
	[cmpd07] [nvarchar](50) NULL,
	[cmpd08] [nvarchar](50) NULL,
	[cmpd09] [nvarchar](50) NULL,
	[cmpd10] [nvarchar](50) NULL,
	[cmpd11] [nvarchar](50) NULL,
	[cmpd12] [nvarchar](50) NULL,
	[cmpd13] [nvarchar](50) NULL,
	[cmpd14] [nvarchar](50) NULL,
	[cmpd15] [nvarchar](50) NULL,
	[cmpo01] [int] NULL,
	[cmpo02] [int] NULL,
	[cmpo03] [int] NULL,
	[cmpo04] [int] NULL,
	[cmpo05] [int] NULL,
	[cmpo06] [int] NULL,
	[cmpo07] [int] NULL,
	[cmpo08] [int] NULL,
	[cmpo09] [int] NULL,
	[cmpo10] [int] NULL,
	[cmpo11] [int] NULL,
	[cmpo12] [int] NULL,
	[cmpo13] [int] NULL,
	[cmpo14] [int] NULL,
	[cmpo15] [int] NULL,
 CONSTRAINT [aaaaaCOMPERFIL_PK] PRIMARY KEY NONCLUSTERED 
(
	[tipoperfil] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY]

GO

ALTER TABLE [dbo].[COMPERFIL] ADD  CONSTRAINT [DF__Temporary__compp__182C9B23]  DEFAULT ((1)) FOR [compprty]
GO

ALTER TABLE [dbo].[COMPERFIL] ADD  CONSTRAINT [DF__Temporary__compe__1920BF5C]  DEFAULT ((1)) FOR [compestr]
GO

ALTER TABLE [dbo].[COMPERFIL] ADD  CONSTRAINT [DF__Temporary__compd__1A14E395]  DEFAULT ((1)) FOR [compdet]
GO

ALTER TABLE [dbo].[COMPERFIL] ADD  CONSTRAINT [DF__Temporary__compi__1B0907CE]  DEFAULT ((1)) FOR [compimag]
GO

ALTER TABLE [dbo].[COMPERFIL] ADD  CONSTRAINT [DF__Temporary__compr__1BFD2C07]  DEFAULT ((1)) FOR [comprel]
GO

