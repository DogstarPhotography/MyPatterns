
/****** Object:  Table [dbo].[SampleData]    Script Date: 10/08/2012 09:31:37 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [SampleData](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[Name] [varchar](50) NULL,
	[Value] [varchar](max) NULL
) ON [PRIMARY]
GO

INSERT INTO [EnterpriseLibrary_Patterns].[dbo].[SampleData]
           ([Name]
           ,[Value])
     VALUES
           ('Test'
           ,'One')
GO

INSERT INTO [EnterpriseLibrary_Patterns].[dbo].[SampleData]
           ([Name]
           ,[Value])
     VALUES
           ('Test'
           ,'Two')
GO

INSERT INTO [EnterpriseLibrary_Patterns].[dbo].[SampleData]
           ([Name]
           ,[Value])
     VALUES
           ('Test'
           ,'Three')
GO