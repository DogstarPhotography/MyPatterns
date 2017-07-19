CREATE TABLE [dbo].[Insects] (
    [Index]       INT            IDENTITY (1, 1) NOT NULL,
    [Name]        NVARCHAR (MAX) NOT NULL,
    [BranchIndex] INT            NULL,
    [LeafIndex]   INT            NULL
);

