USE [DynamicsAxTools]
GO

INSERT INTO [dbo].[AXTools_Environments]
           ([ENVIRONMENT]
           ,[DESCRIPTION]
           ,[DBSERVER]
           ,[DBNAME]
           ,[DBUSER]
           ,[CPUTHOLD]
           ,[BLOCKTHOLD]
           ,[WAITINGTHOLD]
           ,[RUNGRD]
           ,[RUNSTATS]
           ,[EMAILPROFILE]
           ,[LOCALADMINUSER])
     VALUES
           ('ENV_Name'
           ,'ENV_Description'
           ,'SQLServer_Name'
           ,'AX_Database'
           ,'UserAccount_AXDB' --If NULL runs as same user running the job
           ,75 --CPU threshold to trigger script
           ,15 --Number of blocking to trigger script
           ,1800000 --Waiting time from a single block to trigger script
           ,0 --If CPU above threshold script should update statistics
           ,2 --Run regular update statistics based on seetings. 0-Not run/1-Log statistics info but not run/2-Log and run.
           ,'EmailProfileID' --Configuration from AXTools_EmailProfile
           ,'LocalAdminUserAccount_AXServers' --Access to gather information from servers, if NULL runs as same user running the job
		   )
GO

INSERT INTO [dbo].[AXTools_Environments]
           ([ENVIRONMENT]
           ,[DESCRIPTION]
           ,[DBSERVER]
           ,[DBNAME]
           ,[DBUSER]
           ,[CPUTHOLD]
           ,[BLOCKTHOLD]
           ,[WAITINGTHOLD]
           ,[RUNGRD]
           ,[RUNSTATS]
           ,[EMAILPROFILE]
           ,[LOCALADMINUSER])
     VALUES
           ('UATREG'
           ,'UAT Regional Server 01'
           ,'UDBREGR3-01'
           ,'MFIStoreDB' MFIStoreDB
           ,''
           ,75 --CPU threshold to trigger script
           ,15 --Number of blocking to trigger script
           ,1800000 --Waiting time from a single block to trigger script
           ,0 --If CPU above threshold script should update statistics
           ,2 --Run regular update statistics based on seetings. 0-Not run/1-Log statistics info but not run/2-Log and run.
           ,'' --Configuration from AXTools_EmailProfile
           ,'' --Access to gather information from servers
		   )
GO

