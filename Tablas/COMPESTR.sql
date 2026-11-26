USE [dbcordatv46_unif]
GO

/****** Object:  Table [dbo].[COMPESTR]    Script Date: 03/27/2018 17:31:02 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[COMPESTR](
	[objid] [int] IDENTITY(1,1) NOT NULL,
	[codcomp] [int] NULL,
	[tipocomp] [nvarchar](20) NULL,
	[descri] [nvarchar](50) NULL,
 CONSTRAINT [aaaaaCOMPESTR_PK] PRIMARY KEY NONCLUSTERED 
(
	[objid] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON, FILLFACTOR = 80) ON [PRIMARY]
) ON [PRIMARY]

GO

ALTER TABLE [dbo].[COMPESTR] ADD  CONSTRAINT [DF__Temporary__codco__1367E606]  DEFAULT ((0)) FOR [codcomp]
GO

