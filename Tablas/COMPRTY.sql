USE [dbcordatv46_unif]
GO

/****** Object:  Table [dbo].[COMPPRTY]    Script Date: 03/27/2018 17:34:19 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[COMPPRTY](
	[objid] [int] IDENTITY(1,1) NOT NULL,
	[codcomp] [int] NULL,
	[tipocomp] [nvarchar](20) NULL,
	[tipogen] [nvarchar](50) NULL,
	[sistorigen] [nvarchar](50) NULL,
 CONSTRAINT [aaaaaCOMPPRTY_PK] PRIMARY KEY NONCLUSTERED 
(
	[objid] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY]

GO

ALTER TABLE [dbo].[COMPPRTY] ADD  CONSTRAINT [DF__Temporary__codco__32E0915F]  DEFAULT ((0)) FOR [codcomp]
GO

