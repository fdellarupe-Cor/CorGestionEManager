USE [dbcordatv46_unif]
GO

/****** Object:  Table [dbo].[COMParc]    Script Date: 03/27/2018 17:25:38 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[COMParc](
	[objid] [int] IDENTITY(1,1) NOT NULL,
	[codcomp] [int] NULL,
	[tipocomp] [nvarchar](20) NULL,
	[archivo] [nvarchar](200) NULL,
	[descri] [nvarchar](200) NULL,
	[secuencia] [int] NULL,
 CONSTRAINT [PK_COMParc] PRIMARY KEY NONCLUSTERED 
(
	[objid] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY]

GO

