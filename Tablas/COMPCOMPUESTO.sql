USE [dbcordatv46_unif]
GO

/****** Object:  Table [dbo].[Compcompuesto]    Script Date: 03/27/2018 17:26:46 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[Compcompuesto](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[CompMaster] [int] NULL,
	[TipoMaster] [nvarchar](20) NULL,
	[CompAsoc] [int] NULL,
	[TipoAsoc] [nvarchar](20) NULL,
 CONSTRAINT [aaaaaCompcompuesto_PK] PRIMARY KEY NONCLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY]

GO

ALTER TABLE [dbo].[Compcompuesto] ADD  CONSTRAINT [DF__Temporary__CompM__2D27B809]  DEFAULT ((0)) FOR [CompMaster]
GO

ALTER TABLE [dbo].[Compcompuesto] ADD  CONSTRAINT [DF__Temporary__CompA__2E1BDC42]  DEFAULT ((0)) FOR [CompAsoc]
GO

