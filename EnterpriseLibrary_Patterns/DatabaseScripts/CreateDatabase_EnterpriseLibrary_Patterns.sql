/****** Object:  Database [EnterpriseLibrary_Patterns]    Script Date: 10/08/2012 09:54:35 ******/
CREATE DATABASE [EnterpriseLibrary_Patterns] ON  PRIMARY 
( NAME = N'EnterpriseLibrary_Patterns', FILENAME = N'c:\Program Files (x86)\Microsoft SQL Server\MSSQL.1\MSSQL\DATA\EnterpriseLibrary_Patterns.mdf' , SIZE = 3072KB , MAXSIZE = UNLIMITED, FILEGROWTH = 1024KB )
 LOG ON 
( NAME = N'EnterpriseLibrary_Patterns_log', FILENAME = N'c:\Program Files (x86)\Microsoft SQL Server\MSSQL.1\MSSQL\DATA\EnterpriseLibrary_Patterns_log.ldf' , SIZE = 1024KB , MAXSIZE = 2048GB , FILEGROWTH = 10%)
GO

ALTER DATABASE [EnterpriseLibrary_Patterns] SET COMPATIBILITY_LEVEL = 90
GO

IF (1 = FULLTEXTSERVICEPROPERTY('IsFullTextInstalled'))
begin
EXEC [EnterpriseLibrary_Patterns].[dbo].[sp_fulltext_database] @action = 'enable'
end
GO

ALTER DATABASE [EnterpriseLibrary_Patterns] SET ANSI_NULL_DEFAULT OFF 
GO

ALTER DATABASE [EnterpriseLibrary_Patterns] SET ANSI_NULLS OFF 
GO

ALTER DATABASE [EnterpriseLibrary_Patterns] SET ANSI_PADDING OFF 
GO

ALTER DATABASE [EnterpriseLibrary_Patterns] SET ANSI_WARNINGS OFF 
GO

ALTER DATABASE [EnterpriseLibrary_Patterns] SET ARITHABORT OFF 
GO

ALTER DATABASE [EnterpriseLibrary_Patterns] SET AUTO_CLOSE OFF 
GO

ALTER DATABASE [EnterpriseLibrary_Patterns] SET AUTO_CREATE_STATISTICS ON 
GO

ALTER DATABASE [EnterpriseLibrary_Patterns] SET AUTO_SHRINK OFF 
GO

ALTER DATABASE [EnterpriseLibrary_Patterns] SET AUTO_UPDATE_STATISTICS ON 
GO

ALTER DATABASE [EnterpriseLibrary_Patterns] SET CURSOR_CLOSE_ON_COMMIT OFF 
GO

ALTER DATABASE [EnterpriseLibrary_Patterns] SET CURSOR_DEFAULT  GLOBAL 
GO

ALTER DATABASE [EnterpriseLibrary_Patterns] SET CONCAT_NULL_YIELDS_NULL OFF 
GO

ALTER DATABASE [EnterpriseLibrary_Patterns] SET NUMERIC_ROUNDABORT OFF 
GO

ALTER DATABASE [EnterpriseLibrary_Patterns] SET QUOTED_IDENTIFIER OFF 
GO

ALTER DATABASE [EnterpriseLibrary_Patterns] SET RECURSIVE_TRIGGERS OFF 
GO

ALTER DATABASE [EnterpriseLibrary_Patterns] SET  DISABLE_BROKER 
GO

ALTER DATABASE [EnterpriseLibrary_Patterns] SET AUTO_UPDATE_STATISTICS_ASYNC OFF 
GO

ALTER DATABASE [EnterpriseLibrary_Patterns] SET DATE_CORRELATION_OPTIMIZATION OFF 
GO

ALTER DATABASE [EnterpriseLibrary_Patterns] SET TRUSTWORTHY OFF 
GO

ALTER DATABASE [EnterpriseLibrary_Patterns] SET ALLOW_SNAPSHOT_ISOLATION OFF 
GO

ALTER DATABASE [EnterpriseLibrary_Patterns] SET PARAMETERIZATION SIMPLE 
GO

ALTER DATABASE [EnterpriseLibrary_Patterns] SET READ_COMMITTED_SNAPSHOT OFF 
GO

ALTER DATABASE [EnterpriseLibrary_Patterns] SET  READ_WRITE 
GO

ALTER DATABASE [EnterpriseLibrary_Patterns] SET RECOVERY SIMPLE 
GO

ALTER DATABASE [EnterpriseLibrary_Patterns] SET  MULTI_USER 
GO

ALTER DATABASE [EnterpriseLibrary_Patterns] SET PAGE_VERIFY CHECKSUM  
GO

ALTER DATABASE [EnterpriseLibrary_Patterns] SET DB_CHAINING OFF 
GO


