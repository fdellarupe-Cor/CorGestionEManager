USE [dbcordatv46_unif]
GO

/****** Object:  Table [dbo].[COMPDET]    Script Date: 03/27/2018 17:27:58 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[COMPDET](
	[objid] [int] IDENTITY(1,1) NOT NULL,
	[codcomp] [int] NULL,
	[tipocomp] [nvarchar](20) NULL,
	[codestr] [int] NULL,
	[datodescri] [nvarchar](50) NULL,
	[datotipo] [nvarchar](30) NULL,
	[datoformul] [int] NULL,
 CONSTRAINT [aaaaaCOMPDET_PK] PRIMARY KEY NONCLUSTERED 
(
	[objid] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY]

GO

ALTER TABLE [dbo].[COMPDET] ADD  CONSTRAINT [DF__Temporary__codco__20C1E124]  DEFAULT ((0)) FOR [codcomp]
GO

ALTER TABLE [dbo].[COMPDET] ADD  CONSTRAINT [DF__Temporary__codes__21B6055D]  DEFAULT ((0)) FOR [codestr]
GO

ALTER TABLE [dbo].[COMPDET] ADD  CONSTRAINT [DF__Temporary__datof__22AA2996]  DEFAULT ((0)) FOR [datoformul]
GO

