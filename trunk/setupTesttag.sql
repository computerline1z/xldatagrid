USE [master]
GO
IF  EXISTS (SELECT name FROM sys.databases WHERE name = N'testtags')
DROP DATABASE [testtags]

USE [master]
GO
CREATE DATABASE [testtags] 
GO

USE [testtags]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[tags](
	[tag] [nvarchar](15) NOT NULL,
	[description] [nvarchar](50) NULL,
 CONSTRAINT [PK_tags] PRIMARY KEY CLUSTERED 
(
	[tag] ASC
)WITH (PAD_INDEX  = OFF, IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]

GO

INSERT dbo.tags (tag, description) VALUES ('1101', 'Cash')
INSERT dbo.tags (tag, description) VALUES ('1102', 'Cash in Bank')
INSERT dbo.tags (tag, description) VALUES ('41',   'Sales')
INSERT dbo.tags (tag, description) VALUES ('5101', 'Cost of Goods Sold')

