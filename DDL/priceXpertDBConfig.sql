--This SQL Script to create Staging databse configuration info from global databse.
--This aplPriceExpertConfig will persist the staging db connection information and user id.
--Below condition will check the global database if exists or not.
--By testing purpose has used database name "APLPX".If change this name. we need to configure the newly changed name in WCF service wep.config file. 

    SET ANSI_PADDING ON
	CREATE TABLE [dbo].[aplPriceExpertConfig](
		[id] [int] IDENTITY(1,1) NOT NULL,
		[serverName] [nvarchar](max) NULL,
		[serverType] [nvarchar](max) NULL,
		[authentication] [varchar](max) NULL,
		[userName] [nvarchar](max) NULL,
		[password] [nvarchar](max) NULL,
		[databaseName] [nvarchar](max) NULL,
		[userId] [int] NULL
	) ON [PRIMARY] TEXTIMAGE_ON [PRIMARY]

SET ANSI_PADDING OFF
