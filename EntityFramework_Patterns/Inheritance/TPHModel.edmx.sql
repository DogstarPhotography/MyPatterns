
-- --------------------------------------------------
-- Entity Designer DDL Script for SQL Server 2005, 2008, and Azure
-- --------------------------------------------------
-- Date Created: 06/12/2012 13:42:01
-- Generated from EDMX file: C:\Users\RobinBrown\documents\visual studio 2010\Projects\EntityFramework_Patterns\EntityFramework_Inheritance\TPHModel.edmx
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

IF OBJECT_ID(N'[dbo].[FK_TPHPrimaryDataTPHSecondaryData]', 'F') IS NOT NULL
    ALTER TABLE [dbo].[TPHDataItems] DROP CONSTRAINT [FK_TPHPrimaryDataTPHSecondaryData];
GO

-- --------------------------------------------------
-- Dropping existing tables
-- --------------------------------------------------

IF OBJECT_ID(N'[dbo].[TPHDataItems]', 'U') IS NOT NULL
    DROP TABLE [dbo].[TPHDataItems];
GO

-- --------------------------------------------------
-- Creating all tables
-- --------------------------------------------------

-- Creating table 'TPHDataItems'
CREATE TABLE [dbo].[TPHDataItems] (
    [Index] int IDENTITY(1,1) NOT NULL,
    [Name] nvarchar(max)  NOT NULL,
    [Value] nvarchar(max)  NOT NULL,
    [PrimaryIdentifier] nvarchar(max)  NULL,
    [SecondaryIdentifier] nvarchar(max)  NULL,
    [TPHPrimaryDataIndex] int  NOT NULL,
    [__Disc__] nvarchar(4000)  NOT NULL
);
GO

-- --------------------------------------------------
-- Creating all PRIMARY KEY constraints
-- --------------------------------------------------

-- Creating primary key on [Index] in table 'TPHDataItems'
ALTER TABLE [dbo].[TPHDataItems]
ADD CONSTRAINT [PK_TPHDataItems]
    PRIMARY KEY CLUSTERED ([Index] ASC);
GO

-- --------------------------------------------------
-- Creating all FOREIGN KEY constraints
-- --------------------------------------------------

-- Creating foreign key on [TPHPrimaryDataIndex] in table 'TPHDataItems'
ALTER TABLE [dbo].[TPHDataItems]
ADD CONSTRAINT [FK_TPHPrimaryDataTPHSecondaryData]
    FOREIGN KEY ([TPHPrimaryDataIndex])
    REFERENCES [dbo].[TPHDataItems]
        ([Index])
    ON DELETE NO ACTION ON UPDATE NO ACTION;

-- Creating non-clustered index for FOREIGN KEY 'FK_TPHPrimaryDataTPHSecondaryData'
CREATE INDEX [IX_FK_TPHPrimaryDataTPHSecondaryData]
ON [dbo].[TPHDataItems]
    ([TPHPrimaryDataIndex]);
GO

-- --------------------------------------------------
-- Script has ended
-- --------------------------------------------------