/****** Object:  StoredProcedure [CreateDatafile]    Script Date: 02/26/2013 14:37:34 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [CreateDatafile]

	@category varchar(256)
	,@group varchar(256)
	,@filename varchar(256)
	,@extension varchar(256)
	,@content varbinary(max)
	,@newID int OUTPUT

AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

DECLARE @tmp AS TABLE ( id int )

INSERT INTO [Datafiles]
		   ([Category]
		   ,[Group]
		   ,[Filename]
		   ,[Extension]
		   ,[Content])
	 OUTPUT Inserted.ID into @tmp
	 VALUES
		   (@category
		   ,@group
		   ,@filename
		   ,@extension
		   ,@content)

	-- Read the top ID value from the temporary table
	SELECT @newID = (Select TOP 1 ID from @tmp)

END

GO

/****** Object:  StoredProcedure [DeleteDatafile]    Script Date: 02/26/2013 14:37:34 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [DeleteDatafile]

	@id int
	
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

DELETE FROM [Datafiles]
	  WHERE [ID] = @id
END

GO

/****** Object:  StoredProcedure [GetSampleData]    Script Date: 02/26/2013 14:37:34 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [GetSampleData] 

AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

	-- Insert statements for procedure here
SELECT [ID]
	  ,[Name]
	  ,[Value]
  FROM [SampleData]



END

GO

/****** Object:  StoredProcedure [GetSampleDataItem]    Script Date: 02/26/2013 14:37:34 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [GetSampleDataItem] 
	@id int
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

	-- Insert statements for procedure here
SELECT [ID]
	  ,[Name]
	  ,[Value]
  FROM [SampleData]
WHERE [ID] = @id


END


GO

/****** Object:  StoredProcedure [GetSampleDataItemNameValue]    Script Date: 02/26/2013 14:37:34 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [GetSampleDataItemNameValue] 
	@id int,
	@name varchar(50) OUTPUT,
	@value varchar(MAX) OUTPUT
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

	-- Insert statements for procedure here
SELECT @name = [Name],
	   @value = [Value]
  FROM [SampleData]
WHERE [ID] = @id


END


GO

/****** Object:  StoredProcedure [InsertSampleData]    Script Date: 02/26/2013 14:37:34 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [InsertSampleData] 
	-- Add the parameters for the stored procedure here
	@name varchar(50), 
	@value varchar(MAX)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

	-- Insert statements for procedure here
	INSERT INTO [SampleData]
		   ([Name]
		   ,[Value])
	 VALUES
		   (@name
		   ,@value)



END

GO

/****** Object:  StoredProcedure [RetrieveDatafileByID]    Script Date: 02/26/2013 14:37:34 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [RetrieveDatafileByID]

	@id int
	
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

SELECT [ID]
	  ,[Category]
	  ,[Group]
	  ,[Filename]
	  ,[Extension]
	  ,[Content]
  FROM [Datafiles]
WHERE [ID] = @id

END

GO

/****** Object:  StoredProcedure [RetrieveDatafiles]    Script Date: 02/26/2013 14:37:34 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO


-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [RetrieveDatafiles] 

	@category varchar(100),
	@group varchar(100)

AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

SELECT [ID]
	  ,[Category]
	  ,[Group]
	  ,[Filename]
	  ,[Extension]
	  ,[Content]
FROM [Datafiles]
WHERE 
	[Category] = @category
	and [Group] = @group

END


GO

/****** Object:  StoredProcedure [UpdateDatafile]    Script Date: 02/26/2013 14:37:34 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE [UpdateDatafile]

	@id int
	,@category varchar(100)
	,@group varchar(100)
	,@filename varchar(256)
	,@extension varchar(32)
	,@content varbinary(max)
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

UPDATE [Datafiles]
   SET [Category] = @category
	  ,[Group] = @group
	  ,[Filename] = @filename
	  ,[Extension] = @extension
	  ,[Content] = @content
 WHERE [ID] = @id


END

GO


