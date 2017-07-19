/****** Object:  Table [Datafiles]    Script Date: 02/26/2013 14:39:52 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [Datafiles](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[Category] [varchar](100) NOT NULL,
	[Group] [varchar](100) NOT NULL,
	[Filename] [varchar](256) NOT NULL,
	[Extension] [varchar](32) NOT NULL,
	[Content] [varbinary](max) NOT NULL,
 CONSTRAINT [PK_Datafiles] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX  = OFF, STATISTICS_NORECOMPUTE  = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS  = ON, ALLOW_PAGE_LOCKS  = ON) ON [PRIMARY]
) ON [PRIMARY]

GO


