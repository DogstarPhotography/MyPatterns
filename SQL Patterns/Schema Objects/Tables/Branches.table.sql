CREATE TABLE [dbo].[Branches] (
    [Index]     INT            IDENTITY (1, 1) NOT NULL,
    [Name]      NVARCHAR (MAX) NOT NULL,
    [RootIndex] INT            NOT NULL
);

