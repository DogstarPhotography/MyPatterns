/****** Object:  StoredProcedure [InsertSampleData]    Script Date: 10/08/2012 11:53:51 ******/
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


