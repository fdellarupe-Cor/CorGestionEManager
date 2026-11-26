USE [dbcordatv46_unif]
GO

/****** Object:  Table [dbo].[COMPIMAG]    Script Date: 03/27/2018 17:32:03 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[COMPIMAG](
	[objid] [int] IDENTITY(1,1) NOT NULL,
	[codcomp] [int] NULL,
	[tipocomp] [nvarchar](20) NULL,
	[archivo] [nvarchar](500) NULL,
	[descri] [nvarchar](200) NULL,
	[secuencia] [int] NULL,
 CONSTRAINT [aaaaaCOMPIMAG_PK] PRIMARY KEY NONCLUSTERED 
(
	[objid] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY]

GO

ALTER TABLE [dbo].[COMPIMAG] ADD  CONSTRAINT [DF__Temporary__codco__276EDEB3]  DEFAULT ((0)) FOR [codcomp]
GO

ALTER TABLE [dbo].[COMPIMAG] ADD  CONSTRAINT [DF__Temporary__secue__286302EC]  DEFAULT ((0)) FOR [secuencia]
GO

