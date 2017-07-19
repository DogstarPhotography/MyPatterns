
-- --------------------------------------------------
-- Entity Designer DDL Script for SQL Server 2005, 2008, and Azure
-- --------------------------------------------------
-- Date Created: 06/12/2012 12:27:09
-- Generated from EDMX file: C:\Users\RobinBrown\documents\visual studio 2010\Projects\EntityFramework_Patterns\EntityFramework_Inheritance\InheritedModel.edmx
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

IF OBJECT_ID(N'[dbo].[FK_PrimaryDataSecondaryData]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[DataItems_SecondaryData] DROP CONSTRAINT [FK_PrimaryDataSecondaryData];
GO
IF OBJECT_ID(N'[dbo].[FK_PrimaryData_inherits_DataItem]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[DataItems_PrimaryData] DROP CONSTRAINT [FK_PrimaryData_inherits_DataItem];
GO
IF OBJECT_ID(N'[dbo].[FK_SecondaryData_inherits_DataItem]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[DataItems_SecondaryData] DROP CONSTRAINT [FK_SecondaryData_inherits_DataItem];
GO

-- --------------------------------------------------
-- Dropping existing tables
-- --------------------------------------------------

IF OBJECT_ID(N'[dbo].[DataItems]', 'U') IS NOT NULL
    DROP TABLE [dbo].[DataItems];
GO
IF OBJECT_ID(N'[dbo].[DataItems_PrimaryData]', 'U') IS NOT NULL
    DROP TABLE [dbo].[DataItems_PrimaryData];
GO
IF OBJECT_ID(N'[dbo].[DataItems_SecondaryData]', 'U') IS NOT NULL
    DROP TABLE [dbo].[DataItems_SecondaryData];
GO

-- --------------------------------------------------
-- Creating all tables
-- --------------------------------------------------

-- Creating table 'TPTDataItems'
CREATE TABLE [dbo].[TPTDataItems] (
    [Index] int IDENTITY(1,1) NOT NULL,
    [Name] nvarchar(max)  NOT NULL,
    [Value] nvarchar(max)  NOT NULL
);
GO

-- Creating table 'TPTDataItems_TPTPrimaryData'
CREATE TABLE [dbo].[TPTDataItems_TPTPrimaryData] (
    [Identifier] nvarchar(max)  NOT NULL,
    [Index] int  NOT NULL
);
GO

-- Creating table 'TPTDataItems_TPTSecondaryData'
CREATE TABLE [dbo].[TPTDataItems_TPTSecondaryData] (
    [Identifier] nvarchar(max)  NOT NULL,
    [PrimaryDataIndex] int  NOT NULL,
    [Index] int  NOT NULL
);
GO

-- --------------------------------------------------
-- Creating all PRIMARY KEY constraints
-- --------------------------------------------------

-- Creating primary key on [Index] in table 'TPTDataItems'
ALTER TABLE [dbo].[TPTDataItems]
ADD CONSTRAINT [PK_TPTDataItems]
    PRIMARY KEY CLUSTERED ([Index] ASC);
GO

-- Creating primary key on [Index] in table 'TPTDataItems_TPTPrimaryData'
ALTER TABLE [dbo].[TPTDataItems_TPTPrimaryData]
ADD CONSTRAINT [PK_TPTDataItems_TPTPrimaryData]
    PRIMARY KEY CLUSTERED ([Index] ASC);
GO

-- Creating primary key on [Index] in table 'TPTDataItems_TPTSecondaryData'
ALTER TABLE [dbo].[TPTDataItems_TPTSecondaryData]
ADD CONSTRAINT [PK_TPTDataItems_TPTSecondaryData]
    PRIMARY KEY CLUSTERED ([Index] ASC);
GO

-- --------------------------------------------------
-- Creating all FOREIGN KEY constraints
-- --------------------------------------------------

-- Creating foreign key on [PrimaryDataIndex] in table 'TPTDataItems_TPTSecondaryData'
ALTER TABLE [dbo].[TPTDataItems_TPTSecondaryData]
ADD CONSTRAINT [FK_PrimaryDataSecondaryData]
    FOREIGN KEY ([PrimaryDataIndex])
    REFERENCES [dbo].[TPTDataItems_TPTPrimaryData]
        ([Index])
    ON DELETE NO ACTION ON UPDATE NO ACTION;

-- Creating non-clustered index for FOREIGN KEY 'FK_PrimaryDataSecondaryData'
CREATE INDEX [IX_FK_PrimaryDataSecondaryData]
ON [dbo].[TPTDataItems_TPTSecondaryData]
    ([PrimaryDataIndex]);
GO

-- Creating foreign key on [Index] in table 'TPTDataItems_TPTPrimaryData'
ALTER TABLE [dbo].[TPTDataItems_TPTPrimaryData]
ADD CONSTRAINT [FK_TPTPrimaryData_inherits_TPTDataItem]
    FOREIGN KEY ([Index])
    REFERENCES [dbo].[TPTDataItems]
        ([Index])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- Creating foreign key on [Index] in table 'TPTDataItems_TPTSecondaryData'
ALTER TABLE [dbo].[TPTDataItems_TPTSecondaryData]
ADD CONSTRAINT [FK_TPTSecondaryData_inherits_TPTDataItem]
    FOREIGN KEY ([Index])
    REFERENCES [dbo].[TPTDataItems]
        ([Index])
    ON DELETE NO ACTION ON UPDATE NO ACTION;
GO

-- --------------------------------------------------
-- Script has ended
-- --------------------------------------------------