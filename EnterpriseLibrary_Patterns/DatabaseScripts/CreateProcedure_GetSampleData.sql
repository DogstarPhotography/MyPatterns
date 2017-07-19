/****** Object:  StoredProcedure [GetSampleData]    Script Date: 10/08/2012 10:42:32 ******/
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

/****** Object:  StoredProcedure [GetSampleData]    Script Date: 10/08/2012 14:23:10 ******/
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



/****** Object:  StoredProcedure [GetSampleData]    Script Date: 10/08/2012 14:23:10 ******/
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


