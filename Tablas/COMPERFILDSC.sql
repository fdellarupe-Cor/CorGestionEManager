USE [dbcordatv46_unif]
GO

/****** Object:  Table [dbo].[comperfildsc]    Script Date: 03/27/2018 17:30:23 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [dbo].[comperfildsc](
	[objid] [int] IDENTITY(1,1) NOT NULL,
	[dato] [nvarchar](50) NULL,
	[descri] [nvarchar](200) NULL,
 CONSTRAINT [PK_comperfildsc] PRIMARY KEY CLUSTERED 
(
	[objid] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON, FILLFACTOR = 90) ON [PRIMARY]
) ON [PRIMARY]

GO

