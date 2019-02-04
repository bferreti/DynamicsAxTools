--JUST FOR AX-Report script

USE [DynamicsAxTools]
GO

INSERT INTO [dbo].[AXTools_Servers]
           ([ENVIRONMENT]
           ,[ACTIVE]
           ,[SERVERNAME]
           ,[SERVERTYPE])
     VALUES
           ('ENV_Name'
           ,1 --Server is active 1/0
           ,'Server_Name'
           ,'Server_Type' --AOS, IIS, SQL, SRS
		   )
GO

INSERT INTO [dbo].[AXTools_Servers]
           ([ENVIRONMENT]
           ,[ACTIVE]
           ,[SERVERNAME]
           ,[SERVERTYPE])
     VALUES
           ('UATREG'
           ,1 --Server is active 1/0
           ,'UAPAOSR3-01'
           ,'AOS' --AOS, IIS, SQL, SRS
		   )
GO

INSERT INTO [dbo].[AXTools_Servers]
           ([ENVIRONMENT]
           ,[ACTIVE]
           ,[SERVERNAME]
           ,[SERVERTYPE])
     VALUES
           ('UATREG'
           ,1 --Server is active 1/0
           ,'UDBREGR3-01'
           ,'SQL' --AOS, IIS, SQL, SRS
		   )
GO


