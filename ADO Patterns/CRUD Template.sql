/****** Object:  Table [Settings]    Script Date: 10/10/2012 10:14:54 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE TABLE [Settings](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[Section] [nvarchar](50) NOT NULL,
	[Entry] [nvarchar](50) NOT NULL,
	[Value] [nvarchar](max) NULL
) ON [PRIMARY]

GO

/****** Object:  StoredProcedure [CreateSetting]    Script Date: 10/10/2012 10:10:58 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		DEVELOPER_NAME_HERE
-- Create date: DATE_HERE
-- Description:	DESCRIPTION_HERE
-- =============================================
CREATE PROCEDURE [CreateSetting]
	-- Add the parameters for the stored procedure here
	@Section nvarchar(50)
	,@Entry nvarchar(50)
	,@Value nvarchar(max)
	,@newID int OUTPUT
	
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

DECLARE @tmp AS TABLE ( id int )

    -- Insert statements for procedure here
INSERT INTO [MyPatterns].[dbo].[Settings]
           ([Section]
           ,[Entry]
           ,[Value])
     OUTPUT Inserted.ID into @tmp
     VALUES
           (@Section
           ,@Entry
           ,@Value)

	-- Read the top ID value from the temporary table
	--SELECT @newID = SCOPE_IDENTITY() -- This has bugs in parallel processing situations
	SELECT @newID = (Select TOP 1 ID from @tmp)
	
END

/****** Object:  StoredProcedure [RetrieveSetting]    Script Date: 10/10/2012 10:11:01 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		DEVELOPER_NAME_HERE
-- Create date: DATE_HERE
-- Description:	DESCRIPTION_HERE
-- =============================================
CREATE PROCEDURE [RetrieveSetting] 
	-- Add the parameters for the stored procedure here
	@ID int
	
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
SELECT [ID]
      ,[Section]
      ,[Entry]
      ,[Value]
  FROM [MyPatterns].[dbo].[Settings]
WHERE [ID] = @ID

END

/****** Object:  StoredProcedure [UpdateSetting]    Script Date: 10/10/2012 10:11:03 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		DEVELOPER_NAME_HERE
-- Create date: DATE_HERE
-- Description:	DESCRIPTION_HERE
-- =============================================
CREATE PROCEDURE [UpdateSetting] 
	-- Add the parameters for the stored procedure here
	@ID int
	,@Section nvarchar(50)
	,@Entry nvarchar(50)
	,@Value nvarchar(max)

AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
UPDATE [MyPatterns].[dbo].[Settings]
   SET [Section] = @Section
      ,[Entry] = @Entry
      ,[Value] = @Value
 WHERE [ID] = @ID

END

/****** Object:  StoredProcedure [DeleteSetting]    Script Date: 10/10/2012 10:11:00 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		DEVELOPER_NAME_HERE
-- Create date: DATE_HERE
-- Description:	DESCRIPTION_HERE
-- =============================================
CREATE PROCEDURE [DeleteSetting] 
	-- Add the parameters for the stored procedure here
	@ID int

AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
DELETE FROM [MyPatterns].[dbo].[Settings]
 WHERE [ID] = @ID

END

