
-- --------------------------------------------------
-- Entity Designer DDL Script for SQL Server 2005, 2008, and Azure
-- --------------------------------------------------
-- Date Created: 06/13/2012 16:25:29
-- Generated from EDMX file: C:\Users\RobinBrown\Documents\MyPatterns\EntityFramework_Patterns\ModelFirst\DataModel.edmx
-- --------------------------------------------------

SET QUOTED_IDENTIFIER OFF;
GO
USE [MyPatterns];
GO
IF SCHEMA_ID(N'dbo') IS NULL EXECUTE(N'CREATE SCHEMA [dbo]');
GO

-- --------------------------------------------------
-- Dropping existing FOREIGN KEY constraints
-- --------------------------------------------------

IF OBJECT_ID(N'[dbo].[FK_RootBranch]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[Branches] DROP CONSTRAINT [FK_RootBranch];
GO
IF OBJECT_ID(N'[dbo].[FK_BranchLeaf]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[Leaves] DROP CONSTRAINT [FK_BranchLeaf];
GO
IF OBJECT_ID(N'[dbo].[FK_BranchInsect]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[Insects] DROP CONSTRAINT [FK_BranchInsect];
GO
IF OBJECT_ID(N'[dbo].[FK_LeafInsect]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[Insects] DROP CONSTRAINT [FK_LeafInsect];
GO

-- --------------------------------------------------
-- Dropping existing tables
-- --------------------------------------------------

IF OBJECT_ID(N'[dbo].[Roots]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Roots];
GO
IF OBJECT_ID(N'[dbo].[Branches]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Branches];
GO
IF OBJECT_ID(N'[dbo].[Leaves]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Leaves];
GO
IF OBJECT_ID(N'[dbo].[Insects]', 'U') IS NOT NULL
    DROP TABLE [dbo].[Insects];
GO

-- --------------------------------------------------
-- Creating all tables
-- --------------------------------------------------

-- Creating table 'Roots'
CREATE TABLE [dbo].[Roots] (
    [Index] int IDENTITY(1,1) NOT NULL,
    [Name] nvarchar(50)  NOT NULL
);
GO

-- Creating table 'Branches'
CREATE TABLE [dbo].[Branches] (
    [Index] int IDENTITY(1,1) NOT NULL,
    [Name] nvarchar(max)  NOT NULL,
    [RootIndex] int  NOT NULL
);
GO

-- Creating table 'Leaves'
CREATE TABLE [dbo].[Leaves] (
    [Index] int IDENTITY(1,1) NOT NULL,
    [Name] nvarchar(max)  NOT NULL,
    [BranchIndex] int  NOT NULL
);
GO

-- Creating table 'Insects'
CREATE TABLE [dbo].[Insects] (
    [Index] int IDENTITY(1,1) NOT NULL,
    [Name] nvarchar(max)  NOT NULL,
    [BranchIndex] int  NULL,
    [LeafIndex] int  NULL
);
GO

-- --------------------------------------------------
-- Creating all PRIMARY KEY constraints
-- --------------------------------------------------

-- Creating primary key on [Index] in table 'Roots'
ALTER TABLE [dbo].[Roots]
ADD CONSTRAINT [PK_Roots]
    PRIMARY KEY CLUSTERED ([Index] ASC);
GO

-- Creating primary key on [Index] in table 'Branches'
ALTER TABLE [dbo].[Branches]
ADD CONSTRAINT [PK_Branches]
    PRIMARY KEY CLUSTERED ([Index] ASC);
GO

-- Creating primary key on [Index] in table 'Leaves'
ALTER TABLE [dbo].[Leaves]
ADD CONSTRAINT [PK_Leaves]
    PRIMARY KEY CLUSTERED ([Index] ASC);
GO

-- Creating primary key on [Index] in table 'Insects'
ALTER TABLE [dbo].[Insects]
ADD CONSTRAINT [PK_Insects]
    PRIMARY KEY CLUSTERED ([Index] ASC);
GO

-- --------------------------------------------------
-- Creating all FOREIGN KEY constraints
-- --------------------------------------------------

-- Creating foreign key on [RootIndex] in table 'Branches'
ALTER TABLE [dbo].[Branches]
ADD CONSTRAINT [FK_RootBranch]
    FOREIGN KEY ([RootIndex])
    REFERENCES [dbo].[Roots]
        ([Index])
    ON DELETE NO ACTION ON UPDATE NO ACTION;

-- Creating non-clustered index for FOREIGN KEY 'FK_RootBranch'
CREATE INDEX [IX_FK_RootBranch]
ON [dbo].[Branches]
    ([RootIndex]);
GO

-- Creating foreign key on [BranchIndex] in table 'Leaves'
ALTER TABLE [dbo].[Leaves]
ADD CONSTRAINT [FK_BranchLeaf]
    FOREIGN KEY ([BranchIndex])
    REFERENCES [dbo].[Branches]
        ([Index])
    ON DELETE NO ACTION ON UPDATE NO ACTION;

-- Creating non-clustered index for FOREIGN KEY 'FK_BranchLeaf'
CREATE INDEX [IX_FK_BranchLeaf]
ON [dbo].[Leaves]
    ([BranchIndex]);
GO

-- Creating foreign key on [BranchIndex] in table 'Insects'
ALTER TABLE [dbo].[Insects]
ADD CONSTRAINT [FK_BranchInsect]
    FOREIGN KEY ([BranchIndex])
    REFERENCES [dbo].[Branches]
        ([Index])
    ON DELETE NO ACTION ON UPDATE NO ACTION;

-- Creating non-clustered index for FOREIGN KEY 'FK_BranchInsect'
CREATE INDEX [IX_FK_BranchInsect]
ON [dbo].[Insects]
    ([BranchIndex]);
GO

-- Creating foreign key on [LeafIndex] in table 'Insects'
ALTER TABLE [dbo].[Insects]
ADD CONSTRAINT [FK_LeafInsect]
    FOREIGN KEY ([LeafIndex])
    REFERENCES [dbo].[Leaves]
        ([Index])
    ON DELETE NO ACTION ON UPDATE NO ACTION;

-- Creating non-clustered index for FOREIGN KEY 'FK_LeafInsect'
CREATE INDEX [IX_FK_LeafInsect]
ON [dbo].[Insects]
    ([LeafIndex]);
GO

-- --------------------------------------------------
-- Script has ended
-- --------------------------------------------------