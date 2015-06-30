--This SQL script to create Staging database and corresponding sku and result  tables--
--The data base name configure AplStockKeeping--
--The database table name 1.AplStaockKeeping, 2.Result--
-- You can configure AplStockKeeping database .mdf and ldf file location--

--Note : This APlStockKeeping database name has used our application for staging data process.If change the database name apllicaiton also needs to be change in DML process.

USE [Master]
GO

CREATE DATABASE [AplStockKeeping] ON  PRIMARY 
( NAME = N'AplStockKeeping', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL12.MSSQLSERVER\MSSQL\DATA\AplStockKeeping.mdf' , -- default location is configured.may change your desire location
  SIZE = 2GB , MAXSIZE = 8GB, FILEGROWTH = 1GB )
LOG ON 
( NAME = N'AplStockKeeping_log', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL12.MSSQLSERVER\MSSQL\DATA\AplStockKeeping_log.ldf' , -- default location is configured.may change your desire location
  SIZE = 1GB , MAXSIZE = 2GB , FILEGROWTH = 10%)
GO

USE [AplStockKeeping]
GO

SET ANSI_PADDING ON
GO
--Create table for staging sku's 
CREATE TABLE [dbo].[AplStockKeeping](
	[stockId] [int] IDENTITY(1,1) NOT NULL,
	[Sku] [nvarchar](30) NULL,
	[productName] [varchar](250) NULL,
	[companyName] [varchar](50) NULL,
	[productPrice] [decimal](18, 3) NULL,
	[shippingCost] [decimal](18, 3) NOT NULL,
	[inStock] [bit] NULL,
	[Crawl_Date] [datetime] NOT NULL,
	[createdDate] [datetime] NOT NULL,
	[lastModifiedDate] [datetime] NOT NULL,
 CONSTRAINT [PK_AplStockKeeping] PRIMARY KEY CLUSTERED 
(
	[stockId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

SET QUOTED_IDENTIFIER ON
GO
--Create table for result set updation.
CREATE TABLE [dbo].[Result](
	[ID] [int] IDENTITY(1,1) NOT NULL,
	[FileName] [nvarchar](100) NOT NULL,
	[Format] [nvarchar](25) NOT NULL,
	[RowsInFile] [int] NOT NULL,
	[RowsImported] [int] NULL,
	[RowsWithError] [int] NULL,
	[ErrorFilePath] [nvarchar](max) NULL,
	[ErrorFileName] [nvarchar](100) NULL,
	[StartTime] [datetime] NULL,
	[EndTime] [datetime] NULL,
	[Duration] [nvarchar](100) NULL,
	[SourceFilePath] [nvarchar](max) NULL,
 CONSTRAINT [PK_Result] PRIMARY KEY CLUSTERED 
(
	[ID] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

GO

--END--